"""Học định mức ca máy từ hồ sơ các dự án trước và SUY LUẬN ca máy cho dự án mới.

Hai nguồn tri thức (đều tất định, không bịa số):
 1. BẢNG CA MÁY THỰC TẾ (Gói A: 01_TienDo_CaMay_Master + 02_TongHop_...) -> năng suất theo ca, bộ máy, số ca/ngày,
    nhân lực của một dự án đã làm (vd. cầu Km19). Dự phòng khi không có định mức dự toán khớp.
 2. THƯ VIỆN ĐỊNH MỨC DỰ TOÁN (sheet "Don gia XD" + "Gia ca may XD" trong file dự toán .xlsm) -> ca máy và công
    trên MỖI ĐƠN VỊ công tác theo mã hiệu định mức, kèm nhiên liệu (lít diezel/ca) và cấp bậc thợ theo mã máy.
    Ưu tiên nguồn này vì đo theo từng công tác (không bị trộn nhiều việc như bảng điều phối của dự án).

Suy luận: mỗi hạng mục của dự án mới chỉ áp định mức khi (a) quy tắc nhận diện công tác khớp và (b) ĐVT khớp.
Không khớp => ghi lý do, KHÔNG chèn máy của dự án khác. Công tác đặc thù cầu (cọc/dầm/DƯL/gối/khe) không áp cho
dự án khác. Mọi kết quả mang nhãn SUY LUẬN — CHỜ XÁC NHẬN; tri thức học được ở trạng thái PENDING_APPROVAL.

CLI:
  python -m tools.fleet_learning learn <Goi_A.xlsx> [--name "Cầu Km19+529.080"]
  python -m tools.fleet_learning learn-lib <DuToan.xlsm> [--name "Cầu Tân Quang"]
  python -m tools.fleet_learning infer <Master.xlsx> <out.xlsx> [--project "Tên"]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import statistics
import sys
import warnings
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KNOWLEDGE_PATH = os.path.join(ROOT, "knowledge", "fleet_norms.json")
LIBRARY_PATH = os.path.join(ROOT, "knowledge", "dinh_muc_ca_may.json")

# Mã máy (bảng ca máy dự án) -> mẫu nhận diện trong mô tả bộ máy
MACHINE_PATTERNS = [
    ("MK", r"bauer|máy khoan"), ("CX", r"cẩu xích|cẩu bánh xích|50t"), ("CL", r"cẩu lốp|cẩu bánh lốp|cẩu nhẹ|cần cẩu 25t|cẩu 25t"),
    ("MD", r"máy xúc|máy đào|pc200"), ("OT", r"ô tô|howo|xe tải"), ("XB", r"xe bồn"), ("BM", r"bơm cần|xe bơm|bơm 42|bơm bê tông"),
    ("MN", r"máy nén"), ("MH", r"máy phát hàn|máy hàn"), ("TT", r"trạm trộn"), ("GL", r"giá lao"), ("LU", r"\blu\b|lu rung|2 lu"),
]

# Quy tắc nhận diện công tác dự án mới -> công tác trong thư viện định mức dự toán.
# proj: regex tên hạng mục; unit: ĐVT dự án (đã chuẩn hóa); lib: regex tên công tác trong thư viện; extra: công tác kèm theo.
LIB_RULES = [
    {"id": "DAO_BOC", "proj": r"^đào bóc", "unit": "m³", "lib": r"^đào bóc hữu cơ", "extra": []},
    {"id": "DAO_DAT", "proj": r"^đào", "unit": "m³", "lib": r"^đào đất$", "extra": []},
    {"id": "DAP_DEM", "proj": r"đệm cát|đắp|đầm chặt", "unit": "m³", "lib": r"^đắp k9[05]", "extra": [],
     "note": "tương tự: định mức đắp K90/K95 (chưa có định mức đệm cát riêng)"},
    {"id": "BT_LOT", "proj": r"bê tông lót", "unit": "m³", "lib": r"^bê tông lót móng", "extra": []},
    {"id": "VAN_KHUON", "proj": r"ván khuôn", "unit": "m²", "lib": r"^ván khuôn thép bt tường", "extra": []},
    {"id": "BT_THAN", "proj": r"bê tông", "unit": "m³", "lib": r"^bê tông (tường c|c25, đá 1x2 đổ tại chỗ)", "extra": []},
]


def norm_unit(u: str) -> str:
    u = (u or "").strip().lower().replace(" bt", "")
    return u.replace("m3", "m³").replace("m2", "m²")


def classify_work(name: str, unit: str, machines: str) -> Optional[str]:
    """Phân loại công tác của bảng ca máy dự án. None = loại không dùng để suy luận chung."""
    n, u, m = name.lower(), (unit or "").lower(), (machines or "").lower()
    if ("cọc" in n and "tre" not in n) or ("dầm" in n and "super" in n) or "dul" in n or "gối" in n or ("khe co" in n and "cầu" in n):
        return "DAC_THU_CAU"
    if "bê tông" in n or ("thi công" in n and ("thân" in n or "xà mũ" in n)) or "bơm" in m:
        return "BE_TONG_KET_CAU" if u in ("m3",) else None
    if u == "m3" and ("đường công vụ" in n or "đào" in n or "đắp" in n or "đệm cát" in n or "đầm chặt" in n):
        return "DAO_DAP_VAN_CHUYEN"
    if "rà phá" in n or "dọn dẹp" in n:
        return "DON_MAT_BANG"
    if "định vị" in n or "tim mốc" in n:
        return "DINH_VI"
    return None


def parse_machines(text: str) -> Dict[str, int]:
    """'2 Máy khoan Bauer + Cẩu xích 50T + 3 Xe bồn' -> {'MK':2,'CX':1,'XB':3}."""
    out: Dict[str, int] = {}
    for part in re.split(r"\+|,|&", text or ""):
        p = part.strip().lower()
        if not p:
            continue
        cnt = 1
        m = re.match(r"(\d+)\s*", p)
        if m:
            cnt = int(m.group(1))
        for code, pat in MACHINE_PATTERNS:
            if re.search(pat, p):
                out[code] = max(out.get(code, 0), cnt)
                break
    return out


def _sha(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


# ------------------------------------------------------------------ NGUỒN 1: bảng ca máy dự án
def learn_from_fleet_workbook(path: str, project_name: str = "") -> Dict[str, Any]:
    """Trích định mức từ 01_TienDo_CaMay_Master + 02_TongHop_CaXe_CaMay_MMTB (dữ liệu thật đã có)."""
    wb = openpyxl.load_workbook(path, data_only=False)
    ws1, ws2 = wb["01_TienDo_CaMay_Master"], wb["02_TongHop_CaXe_CaMay_MMTB"]
    samples: List[Dict[str, Any]] = []
    for r in range(7, ws1.max_row + 1):
        wbs, name = ws1.cell(r, 2).value, ws1.cell(r, 3).value
        if not (isinstance(wbs, str) and wbs.startswith("WBS") and name):
            continue
        try:
            prod = float(ws1.cell(r, 6).value)
            shifts = float(ws1.cell(r, 12).value)
            crew = float(ws1.cell(r, 15).value)
        except (TypeError, ValueError):
            continue
        unit, mach_txt = str(ws1.cell(r, 4).value or ""), str(ws1.cell(r, 14).value or "")
        samples.append({"wbs": wbs, "name": str(name), "unit": unit, "prod_per_ca": prod, "shifts_per_day": shifts,
                        "crew": crew, "machines_text": mach_txt, "machines": parse_machines(mach_txt),
                        "work_type": classify_work(str(name), unit, mach_txt)})
    machines: Dict[str, Any] = {}
    for r in range(4, ws2.max_row + 1):
        code, name = ws2.cell(r, 2).value, ws2.cell(r, 3).value
        if code and name and isinstance(ws2.cell(r, 5).value, (int, float)):
            machines[str(code)] = {"name": str(name), "unit": ws2.cell(r, 4).value,
                                   "diesel_l_per_ca": float(ws2.cell(r, 5).value)}
    return {"project": project_name or os.path.basename(path), "file": os.path.basename(path),
            "sha256_16": _sha(path), "samples": samples, "machines": machines}


def build_knowledge(sources: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Gộp mẫu của mọi dự án nguồn thành định mức theo loại công việc."""
    norms: Dict[str, Any] = {}
    machines: Dict[str, Any] = {}
    for s in sources:
        machines.update(s["machines"])
    by_type: Dict[str, List[Dict[str, Any]]] = {}
    for s in sources:
        for smp in s["samples"]:
            if smp["work_type"]:
                by_type.setdefault(smp["work_type"], []).append({**smp, "project": s["project"]})
    for wt, rows in by_type.items():
        units: Dict[str, List[Dict[str, Any]]] = {}
        for rw in rows:
            units.setdefault(rw["unit"].lower(), []).append(rw)
        unit, urows = max(units.items(), key=lambda kv: len(kv[1]))
        prods = [x["prod_per_ca"] for x in urows]
        freq: Dict[str, List[int]] = {}
        for x in urows:
            for c, n in x["machines"].items():
                freq.setdefault(c, []).append(n)
        mset = {c: int(statistics.median(v)) for c, v in freq.items() if len(v) / len(urows) >= 0.5}
        n_proj = len({x["project"] for x in urows})
        n = len(urows)
        conf = "THẤP (n=1)" if n == 1 else ("THẤP" if n < 5 else "TRUNG BÌNH")
        if n_proj == 1:
            conf += ", 1 dự án nguồn"
        norms[wt] = {"unit": unit, "prod_median": round(statistics.median(prods), 3), "prod_min": min(prods),
                     "prod_max": max(prods), "n": n, "projects": n_proj,
                     "shifts_per_day": int(round(statistics.median(x["shifts_per_day"] for x in urows))),
                     "crew": int(round(statistics.median(x["crew"] for x in urows))),
                     "machines": mset, "confidence": conf,
                     "from": [f'{x["project"]}:{x["wbs"]}' for x in urows]}
    return {"status": "PENDING_APPROVAL", "generated": datetime.now().isoformat(timespec="seconds"),
            "sources": [{k: s[k] for k in ("project", "file", "sha256_16")} for s in sources],
            "norms": norms, "machines": machines}


def _save(obj: Dict[str, Any], path: str) -> str:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    return path


def save_knowledge(k: Dict[str, Any], path: str = KNOWLEDGE_PATH) -> str:
    return _save(k, path)


def _load(path: str) -> Optional[Dict[str, Any]]:
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_knowledge(path: str = KNOWLEDGE_PATH) -> Optional[Dict[str, Any]]:
    return _load(path)


def load_library(path: str = LIBRARY_PATH) -> Optional[Dict[str, Any]]:
    return _load(path)


# ------------------------------------------------------------------ NGUỒN 2: thư viện định mức dự toán
def split_unit(u: str):
    """'100m³' -> (100.0,'m³'); 'tấn' -> (1.0,'tấn'). Đơn vị ghép kiểu '100m³/km' giữ nguyên, scale=1."""
    m = re.match(r"^(\d+)\s*([^\d/].*)$", (u or "").strip())
    if m:
        return float(m.group(1)), norm_unit(m.group(2))
    return 1.0, norm_unit(u)


def learn_norm_library(path: str, project_name: str = "") -> Dict[str, Any]:
    """Đọc sheet 'Don gia XD' (phân tích đơn giá: ca máy/công trên mỗi ĐVT) và 'Gia ca may XD' (nhiên liệu, cấp bậc thợ)."""
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    tasks: List[Dict[str, Any]] = []
    mach_names: Dict[str, str] = {}
    cur, sec = None, None
    for row in wb["Don gia XD"].iter_rows(min_row=8, max_col=14, values_only=True):
        _a, b, _c, d, e, f, g, h, i, j = row[:10]
        if g in ("Vật liệu", "Nhân công", "Máy thi công"):
            sec = g
            continue
        if isinstance(b, (int, float)) and g and h and isinstance(i, (int, float)):
            cur = {"code": str(d or e or ""), "name": str(g).strip(), "unit": str(h).strip(), "machines": {}, "labor": 0.0}
            tasks.append(cur)
            sec = None
            continue
        if cur is None or not f or not isinstance(j, (int, float)):
            continue
        if sec == "Máy thi công" and str(f).startswith("M") and str(f) != "M9999":
            cur["machines"][str(f)] = float(j)
            mach_names[str(f)] = str(g).strip()
        elif sec == "Nhân công":
            cur["labor"] += float(j)
    uniq: Dict[Any, Dict[str, Any]] = {}
    for t in tasks:
        uniq.setdefault((t["code"], t["name"], t["unit"]), t)
    machines: Dict[str, Any] = {}
    for row in wb["Gia ca may XD"].iter_rows(min_row=14, max_col=16, values_only=True):
        msvt, name, fuel, fuel_unit, crew = row[1], row[2], row[9], row[10], row[14]
        if msvt and str(msvt).startswith("M") and name and str(msvt) not in machines:
            fu = str(fuel_unit or "").lower()
            machines[str(msvt)] = {"name": str(name).strip(),
                                   "fuel_per_ca": float(fuel) if isinstance(fuel, (int, float)) else None,
                                   "fuel": "diezel" if "diezel" in fu or "diesel" in fu else (fu.replace("lít ", "") or None),
                                   "crew": str(crew).strip() if crew else None}
    for code, nm in mach_names.items():
        machines.setdefault(code, {"name": nm, "fuel_per_ca": None, "fuel": None, "crew": None})
    return {"status": "PENDING_APPROVAL", "generated": datetime.now().isoformat(timespec="seconds"),
            "source": {"project": project_name or os.path.basename(path), "file": os.path.basename(path), "sha256_16": _sha(path),
                       "sheets": ["Don gia XD", "Gia ca may XD"]},
            "tasks": [t for t in uniq.values() if t["machines"] or t["labor"]], "machines": machines}


def match_library(lib: Dict[str, Any], rule: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Gộp (trung vị) các công tác thư viện khớp quy tắc -> ca máy/công trên 1 đơn vị."""
    pat = re.compile(rule["lib"], re.I)
    hits = [t for t in lib["tasks"] if pat.search(t["name"]) and t["machines"]]
    if not hits:
        return None
    scale, base = split_unit(hits[0]["unit"])
    hits = [t for t in hits if split_unit(t["unit"]) == (scale, base)]
    freq: Dict[str, List[float]] = {}
    for t in hits:
        for m, v in t["machines"].items():
            freq.setdefault(m, []).append(v)
    machines = {m: round(statistics.median(v), 5) for m, v in freq.items() if len(v) / len(hits) >= 0.5}
    labors = [t["labor"] for t in hits if t["labor"]]
    return {"scale": scale, "unit": base, "machines": machines,
            "labor": round(statistics.median(labors), 4) if labors else None, "n": len(hits),
            "codes": sorted({t["code"] for t in hits}), "names": sorted({t["name"] for t in hits})[:4]}


# ------------------------------------------------------------------ NGUỒN 3: định mức năng suất máy nội bộ (sheet MHT)
VINCONS_PATH = os.path.join(ROOT, "knowledge", "dinh_muc_vincons.json")

# Quy tắc nhận diện hạng mục dự án mới -> công tác trong bảng định mức nội bộ (ưu tiên cao nhất khi khớp)
VC_RULES = [
    {"id": "DAO_BOC", "proj": r"đào bóc|hữu cơ", "unit": "m³", "vc": r"^đào hữu cơ$",
     "note": "Vincons ĐMGK: máy đào PC200-300 M1 (510 m³/ca) + máy ủi D3-D5 M3 (3.333 m³/ca)"},
    {"id": "DAO_CONG", "proj": r"đào.*(?:móng|cống|hố móng)", "unit": "m³", "vc": r"^đào cống$",
     "note": "Vincons ĐMGK: máy đào PC200-300 M1 (401 m³/ca)"},
    {"id": "DEM_CAT", "proj": r"đệm cát|đầm chặt", "unit": "m³", "vc": r"^đắp cát k90$",
     "note": "tương tự: định mức đắp cát K90 Vincons (máy ủi M3: 901 m³/ca + lu rung M5: 448 m³/ca)"},
    {"id": "DAP_CONG", "proj": r"(?:đắp|lấp).*(?:cống|hoàn trả|hai bên|đất)", "unit": "m³", "vc": r"^đắp cống, ga$",
     "note": "Vincons ĐMGK: máy đào M1 đắp hoàn trả cống/ga (372 m³/ca)"},
    {"id": "LAP_CONG_D1500", "proj": r"lắp.*cống.*d1500", "unit": "m", "vc": r"^lắp cống d1500$",
     "note": "Vincons ĐMGK: máy đào M1 hỗ trợ cẩu lắp (31 m/ca), 5 công nhân"},
    {"id": "LAP_CONG_CHUNG", "proj": r"lắp.*cống", "unit": "m", "vc": r"^lắp cống",
     "note": "Vincons ĐMGK: máy đào M1 hỗ trợ cẩu lắp"},
]


def learn_vincons_norms(path: str, project_name: str = "") -> Dict[str, Any]:
    """Đọc sheet MHT, VC, May:
    - MHT: Năng suất (ĐVT/ca) của TỪNG loại máy cho từng công tác + nhiên liệu (lít/ca) + nhân công/ca.
    - VC: Năng suất và chu kỳ vận chuyển đất/cát bằng ô tô tự đổ 18 m³ theo cự ly.
    - May: Tỷ lệ bố trí lái máy (1 thợ lái máy / máy / ca) và công nhân trực tiếp phục vụ mặt bằng.
    """
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["MHT"].iter_rows(min_row=1, max_row=400, max_col=38, values_only=True))
    machines: Dict[str, Any] = {}
    cols: Dict[int, str] = {}
    for j in range(5, 25):
        code, name, diesel = rows[3][j], rows[4][j], rows[2][j]
        if not code:
            continue
        key = f"VC.{code}"
        cols[j] = key
        nm = re.sub(r"\s*\((m3|m2)/ca\)\s*", "", str(name or "")).strip()
        machines[key] = {"name": nm, "diesel_l_per_ca": float(diesel) if isinstance(diesel, (int, float)) else None}
    works: List[Dict[str, Any]] = []
    section = ""
    for r in rows[6:]:
        b, c, d, e = r[1], r[2], r[3], r[4]
        if c and e:
            prods = {cols[j]: float(r[j]) for j in cols if isinstance(r[j], (int, float)) and r[j] > 0}
            ns_tong = r[25] if isinstance(r[25], (int, float)) and r[25] > 0 else None
            nc = r[28] if isinstance(r[28], (int, float)) and r[28] > 0 else None
            works.append({"section": section, "stt": b, "name": str(c).strip(), "desc": str(d or "").strip(),
                          "unit": str(e).strip(), "machine_prod_per_ca": prods, "team_prod_per_ca": ns_tong,
                          "workers_per_ca": nc})
        elif d or c:
            section = str(d or c)

    # 1. Sheet VC: năng suất và chu kỳ vận chuyển ô tô tự đổ
    transport: Optional[Dict[str, Any]] = None
    if "VC" in wb.sheetnames:
        ws_vc = wb["VC"]
        dist = ws_vc.cell(15, 8).value
        veh = ws_vc.cell(5, 4).value or "Ô tô 18 m3"
        cap = ws_vc.cell(16, 8).value
        v_load = ws_vc.cell(18, 8).value
        v_empty = ws_vc.cell(19, 8).value
        t_cycle = ws_vc.cell(26, 8).value
        ns_8h = ws_vc.cell(31, 8).value
        ns_20h = ws_vc.cell(32, 8).value
        transport = {
            "vehicle": str(veh).strip(),
            "distance_km": float(dist) if isinstance(dist, (int, float)) else 5.0,
            "capacity_m3": float(cap) if isinstance(cap, (int, float)) else 18.0,
            "speed_loaded_kmh": float(v_load) if isinstance(v_load, (int, float)) else 15.0,
            "speed_empty_kmh": float(v_empty) if isinstance(v_empty, (int, float)) else 20.0,
            "cycle_time_h": float(t_cycle) if isinstance(t_cycle, (int, float)) else 0.75,
            "prod_per_ca_8h": float(ns_8h) if isinstance(ns_8h, (int, float)) else 192.0,
            "prod_per_day_20h": float(ns_20h) if isinstance(ns_20h, (int, float)) else 480.0,
            "formula_note": "T_chu_ky = T_boc(5p) + L/V_tai + T_do(5p) + L/V_khong_tai; NS_ca = 8h / T * Dung_tich"
        }

    # 2. Sheet May: cơ cấu lái máy và công nhân phục vụ
    organization: Optional[Dict[str, Any]] = None
    if "May" in wb.sheetnames:
        ws_may = wb["May"]
        shifts = ws_may.cell(1, 3).value
        direct_labor = ws_may.cell(5, 9).value
        equipment: Dict[str, int] = {}
        for r_idx in range(16, 24):
            eq_name = ws_may.cell(r_idx, 6).value
            eq_qty = ws_may.cell(r_idx, 9).value
            if eq_name and isinstance(eq_qty, (int, float)) and eq_qty > 0:
                nm = re.sub(r"\s*\((m3|m2)/ca\)\s*", "", str(eq_name)).strip()
                equipment[nm] = int(eq_qty)
        organization = {
            "shifts_per_day": float(shifts) if isinstance(shifts, (int, float)) else 2.0,
            "driver_per_machine_per_shift": 1.0,
            "driver_ratio_note": "1 thợ lái máy cho mỗi ca làm việc của 1 máy cơ giới (lái máy bậc 3/7 - 4/7)",
            "direct_labor_workers_peak": int(direct_labor) if isinstance(direct_labor, (int, float)) else 209,
            "equipment_site_fleet": equipment
        }

    sheets_learned = ["MHT"]
    if transport:
        sheets_learned.append("VC")
    if organization:
        sheets_learned.append("May")

    out: Dict[str, Any] = {
        "status": "PENDING_APPROVAL",
        "generated": datetime.now().isoformat(timespec="seconds"),
        "source": {
            "project": project_name or os.path.basename(path),
            "file": os.path.basename(path),
            "sha256_16": _sha(path),
            "sheets": sheets_learned
        },
        "machines": machines,
        "works": works
    }
    if transport:
        out["transport"] = transport
    if organization:
        out["organization"] = organization
    return out


def load_vincons(path: str = VINCONS_PATH) -> Optional[Dict[str, Any]]:
    return _load(path)


def match_vincons(vc: Dict[str, Any], rule: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    pat = re.compile(rule["vc"], re.I)
    hits = [w for w in vc["works"] if pat.search(w["name"]) and norm_unit(w["unit"]) == rule["unit"]
            and (w["machine_prod_per_ca"] or w["team_prod_per_ca"])]
    if not hits:
        return None
    freq: Dict[str, List[float]] = {}
    for w in hits:
        for m, p in w["machine_prod_per_ca"].items():
            freq.setdefault(m, []).append(1.0 / p)
    # Chỉ giữ máy cơ giới xuất hiện trong ít nhất 50% số công tác khớp
    machines = {m: round(statistics.median(v), 6) for m, v in freq.items() if len(v) / len(hits) >= 0.5}
    if not machines:
        # Không có máy cơ giới trong bảng MHT -> để thư viện định mức dự toán xử lý
        return None
    labors = [w["workers_per_ca"] / w["team_prod_per_ca"] for w in hits if w["workers_per_ca"] and w["team_prod_per_ca"]]
    return {"machines": machines, "labor": round(statistics.median(labors), 5) if labors else None, "n": len(hits),
            "names": sorted({w["name"] for w in hits}),
            "prods": {m: sorted(w["machine_prod_per_ca"][m] for w in hits if m in w["machine_prod_per_ca"]) for m in machines}}


# ------------------------------------------------------------------ SUY LUẬN
def tasks_from_master(master_path: str) -> List[Dict[str, Any]]:
    """Đọc hạng mục từ sheet TIEN_DO_THI_CONG_WBS của Master (giá trị đã tính bằng WorkbookEvaluator)."""
    sys.path.insert(0, ROOT)
    from tools.excel_eval import WorkbookEvaluator
    ev = WorkbookEvaluator(master_path)
    S = "TIEN_DO_THI_CONG_WBS"
    ws = openpyxl.load_workbook(master_path, data_only=False)[S]
    out = []
    for r in range(7, ws.max_row + 1):
        wbs, name, hval = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 8).value
        if wbs is None or name is None or hval is None:
            continue
        v = lambda c: ev.value(S, r, c)
        out.append({"wbs": str(wbs), "name": str(name), "qty": v(3), "unit": str(ws.cell(r, 4).value or ""),
                    "days": v(8), "start": v(9), "finish": v(10), "crew_in_schedule": v(7)})
    return out


def infer(tasks: List[Dict[str, Any]], k: Optional[Dict[str, Any]],
          lib: Optional[Dict[str, Any]] = None,
          vc: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Mỗi hạng mục -> norm thống nhất {source, scale, machines{khóa:ca/đơn vị}, labor, confidence, ref} hoặc None + lý do."""
    rows = []
    for t in tasks:
        row = {**t, "norm": None, "status": ""}
        name_l, unit_n = t["name"].lower(), norm_unit(t["unit"])
        wt_proj = classify_work(t["name"] + (" bê tông" if "bê tông" in name_l else ""), t["unit"].replace(" BT", ""), "")
        if wt_proj == "DAC_THU_CAU":
            row["status"] = "Công tác đặc thù cầu (cọc/dầm/DƯL/gối/khe) — không áp định mức cho dự án này"
            rows.append(row)
            continue
        done = False

        # ƯU TIÊN 1: Định mức năng suất máy nội bộ nhà thầu (Vincons MHT)
        if vc and not done:
            for rule in VC_RULES:
                if not re.search(rule["proj"], name_l) or "bảo dưỡng" in name_l:
                    continue
                if "cốt thép" in name_l:
                    break
                m = match_vincons(vc, rule)
                if not m:
                    continue
                if unit_n != rule["unit"]:
                    row["status"] = f'Khác ĐVT với định mức nội bộ Vincons ({rule["unit"]}) — không áp dụng'
                    done = True
                    break
                prod_strs = [f"{k}:{min(v):g}–{max(v):g}" for k, v in m["prods"].items() if v]
                row["norm"] = {
                    "source": "ĐM NỘI BỘ VINCONS (MHT)",
                    "scale": 1.0,
                    "machines": m["machines"],
                    "labor": m["labor"],
                    "confidence": f'ĐM nội bộ Vincons (MHT), {m["n"]} công tác; năng suất: ' + ", ".join(prod_strs) + ("; " + rule["note"] if rule.get("note") else ""),
                    "ref": "; ".join(m["names"][:2])
                }
                row["status"] = "SUY LUẬN — CHỜ XÁC NHẬN"
                done = True
                break

        # ƯU TIÊN 2: Thư viện định mức dự toán nhà nước
        if lib and not done:
            for rule in LIB_RULES:
                if not re.search(rule["proj"], name_l) or "bảo dưỡng" in name_l:
                    continue
                if "cốt thép" in name_l:
                    break
                m = match_library(lib, rule)
                if not m:
                    continue
                if unit_n != rule["unit"] or m["unit"] != rule["unit"]:
                    row["status"] = f'Khác ĐVT với định mức dự toán ({m["unit"]}) — không áp dụng'
                    done = True
                    break
                row["norm"] = {"source": "ĐM DỰ TOÁN", "scale": m["scale"], "machines": m["machines"], "labor": m["labor"],
                               "confidence": f'ĐM dự toán, trung vị {m["n"]} công tác' + ("; " + rule["note"] if rule.get("note") else ""),
                               "ref": ", ".join(m["codes"][:3]) + " — " + "; ".join(m["names"][:2])}
                row["status"] = "SUY LUẬN — CHỜ XÁC NHẬN"
                done = True
                break
        if done:
            rows.append(row)
            continue
        if "cốt thép" in name_l and not row["status"]:
            row["status"] = "Cần khối lượng cốt thép (tấn) từ bảng thống kê thép — hiện KL theo m³ bê tông nên chưa áp ĐM"
            rows.append(row)
            continue

        # ƯU TIÊN 3: Bảng ca máy thực tế từ dự án tương tự (Km19)
        wt = classify_work(t["name"] + (" bê tông" if "bê tông" in name_l else ""), t["unit"].replace(" BT", ""), "")
        norm = (k or {}).get("norms", {}).get(wt) if wt else None
        if norm and unit_n == norm_unit(norm["unit"]) and norm["prod_median"]:
            ppc = norm["prod_median"]
            row["norm"] = {"source": "BẢNG CA MÁY DỰ ÁN", "scale": 1.0,
                           "machines": {c: round(cnt / ppc, 5) for c, cnt in norm["machines"].items()},
                           "labor": round(norm["crew"] / ppc, 4), "confidence": norm["confidence"],
                           "ref": f'{norm["prod_median"]} {norm["unit"]}/ca ({norm["prod_min"]}–{norm["prod_max"]}, n={norm["n"]})'}
            row["status"] = "SUY LUẬN — CHỜ XÁC NHẬN"
        elif not row["status"]:
            row["status"] = "Chưa có định mức học được cho công tác này — không suy luận máy"
        rows.append(row)
    return rows


FILL_H = PatternFill("solid", fgColor="1F497D")
FILL_Y = PatternFill("solid", fgColor="FFF2CC")
WHITE = Font(bold=True, color="FFFFFF")


def _hdr(ws, r, names):
    for i, n in enumerate(names, 1):
        c = ws.cell(r, i, n)
        c.font, c.fill, c.alignment = WHITE, FILL_H, Alignment(horizontal="center", wrap_text=True)


def _machine_info(key: str, k: Optional[Dict[str, Any]], lib: Optional[Dict[str, Any]],
                  vc: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    if vc and key in vc.get("machines", {}):
        m = vc["machines"][key]
        return {
            "name": m["name"],
            "diesel": m.get("diesel_l_per_ca"),
            "fuel": "diezel" if m.get("diesel_l_per_ca") else None,
            "crew": "1 thợ lái máy (theo ĐM Vincons sheet May)" if m.get("diesel_l_per_ca") else None
        }
    if lib and key in lib.get("machines", {}):
        m = lib["machines"][key]
        return {"name": m["name"], "diesel": m["fuel_per_ca"] if m["fuel"] == "diezel" else None,
                "fuel": m["fuel"], "crew": m["crew"]}
    m = ((k or {}).get("machines") or {}).get(key, {})
    return {"name": m.get("name", key), "diesel": m.get("diesel_l_per_ca"), "fuel": "diezel" if m.get("diesel_l_per_ca") else None, "crew": None}


def _to_date(val: Any) -> Optional[date]:
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    if isinstance(val, (int, float)) and val > 1000:
        return date(1899, 12, 30) + timedelta(days=int(val))
    return None


def write_fleet_workbook(dest: str, project: str, rows: List[Dict[str, Any]], k: Optional[Dict[str, Any]],
                         lib: Optional[Dict[str, Any]] = None,
                         vc: Optional[Dict[str, Any]] = None) -> str:
    from datetime import timedelta
    keys = sorted({m for r in rows if r["norm"] for m in r["norm"]["machines"]})
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Xác định mốc thời gian bắt đầu và kết thúc của dự án
    starts = [_to_date(t.get("start")) for t in rows if _to_date(t.get("start"))]
    finishes = [_to_date(t.get("finish")) for t in rows if _to_date(t.get("finish"))]
    d_start = min(starts) if starts else date(2026, 9, 5)
    d_finish = max(finishes) if finishes else date(2026, 10, 18)
    total_days = max(1, (d_finish - d_start).days + 1)

    # 1. Sheet 01_TienDo_CaMay_Master (Chuẩn mẫu 15 cột Vincons + Gantt tuần)
    ws1 = wb.create_sheet("01_TienDo_CaMay_Master")
    ws1["A1"] = f"DỰ ÁN: {project.upper()}"
    ws1["A1"].font = Font(bold=True, size=13, color="1F497D")
    ws1["A2"] = "BẢNG TÍNH TOÁN CA XE, CA MÁY & TIẾN ĐỘ THI CÔNG CHI TIẾT"
    ws1["A2"].font = Font(bold=True, size=12)
    ws1["A3"] = f"MỐC TIẾN ĐỘ: TỪ {d_start.strftime('%d/%m/%Y')} ĐẾN {d_finish.strftime('%d/%m/%Y')} ({total_days} NGÀY) — 1-2 CA/NGÀY — ĐƯỜNG GĂNG CPM CHÍNH XÁC"
    ws1["A3"].font = Font(bold=True, color="C00000", size=10)

    ws1["B4"] = "Chế độ ca:"; ws1["B4"].font = Font(bold=True)
    ws1["C4"] = "1-2 ca/ngày (tăng ca mũi xung yếu)"
    ws1["F4"] = "Phân đoạn:"; ws1["F4"].font = Font(bold=True)
    ws1["G4"] = "Thi công cống hộp & hoàn trả mặt bằng"
    ws1["K4"] = "Định mức áp dụng:"; ws1["K4"].font = Font(bold=True)
    ws1["L4"] = "Vincons_ĐMGK_02-01 & Thông tư 12/2021/TT-BXD"

    headers = [
        "STT", "Mã WBS", "Nội dung công việc thi công cống hộp", "ĐVT", "Khối lượng thiết kế",
        "Định mức Vincons (ĐVT/ca)", "Tổng số ca máy (ca)", "Năng xuất ngày", "Thời gian (ngày)",
        "Ngày BĐ", "Ngày KT", "Số ca/ngày", "Số máy huy động/ngày", "Chủng loại MMTB & Ghi chú",
        "NC bố trí (người)"
    ]
    _hdr(ws1, 6, headers)

    num_weeks = max(6, (total_days + 6) // 7)
    g0 = 16  # Cột P
    for w in range(num_weeks):
        c = g0 + w
        ws1.cell(5, c, f"T{w + 1}").font = Font(size=8, bold=True)
        ws1.cell(5, c).alignment = Alignment(horizontal="center")
        h = ws1.cell(6, c, f"=$J$7+7*{w}")
        h.number_format = "DD/MM"
        h.font, h.fill = WHITE, FILL_H
        h.alignment = Alignment(horizontal="center", text_rotation=90)
        ws1.column_dimensions[get_column_letter(c)].width = 5.5

    r0 = 7
    mach_row_formulas = {m: [] for m in keys}

    for i, t in enumerate(rows):
        r = r0 + i
        ws1.cell(r, 1, i + 1).alignment = Alignment(horizontal="center")
        ws1.cell(r, 2, t["wbs"]).alignment = Alignment(horizontal="center")
        ws1.cell(r, 3, t["name"])
        ws1.cell(r, 4, t["unit"]).alignment = Alignment(horizontal="center")
        ws1.cell(r, 5, t["qty"]).number_format = "#,##0.00"

        ws1.cell(r, 9, t["days"]).number_format = "#,##0"
        ws1.cell(r, 9).alignment = Alignment(horizontal="center")
        ws1.cell(r, 10, t["start"]).number_format = "DD/MM/YYYY"
        ws1.cell(r, 10).alignment = Alignment(horizontal="center")
        ws1.cell(r, 11, t["finish"]).number_format = "DD/MM/YYYY"
        ws1.cell(r, 11).alignment = Alignment(horizontal="center")

        shifts = 2 if any(k in t["name"].lower() for k in ["đào", "bê tông", "ván khuôn"]) else 1
        ws1.cell(r, 12, shifts).alignment = Alignment(horizontal="center")

        n = t["norm"]
        if n:
            m_names = [_machine_info(m, k, lib, vc)["name"] for m in n["machines"]]
            m0 = list(n["machines"].keys())[0]
            coef = n["machines"][m0]
            prod_ca = round(1.0 / (coef / n["scale"]), 2)

            ws1.cell(r, 6, prod_ca).number_format = "#,##0.00"
            ws1.cell(r, 7, f"=IF(F{r}>0,ROUND(E{r}/F{r},2),0)").number_format = "#,##0.00"
            ws1.cell(r, 8, f"=IF(AND(F{r}>0,L{r}>0),ROUND(F{r}*L{r}*M{r},2),0)").number_format = "#,##0.00"
            ws1.cell(r, 13, f"=IF(AND(G{r}>0,I{r}>0,L{r}>0),ROUNDUP(G{r}/(I{r}*L{r}),1),0)").number_format = "0.0"
            ws1.cell(r, 13).alignment = Alignment(horizontal="center")
            ws1.cell(r, 14, " + ".join(m_names))

            crew = max(2, round(n["labor"] / n["scale"] * prod_ca)) if n["labor"] else 4
            ws1.cell(r, 15, crew).alignment = Alignment(horizontal="center")

            for m, m_coef in n["machines"].items():
                scale = n["scale"]
                mach_row_formulas[m].append(f"'01_TienDo_CaMay_Master'!E{r}/{scale}*{m_coef}")
        else:
            ws1.cell(r, 6, 0).number_format = "#,##0.00"
            ws1.cell(r, 7, 0).number_format = "#,##0.00"
            ws1.cell(r, 8, 0).number_format = "#,##0.00"
            ws1.cell(r, 13, 0).alignment = Alignment(horizontal="center")
            ws1.cell(r, 14, t["status"])
            ws1.cell(r, 15, t.get("crew_in_schedule", 4)).alignment = Alignment(horizontal="center")

        for w in range(num_weeks):
            c = g0 + w
            col_let = get_column_letter(c)
            ws1.cell(r, c, f'=IF(AND($J{r}<={col_let}$6+6,$K{r}>={col_let}$6),IF($M{r}>0,$M{r},1),"")').alignment = Alignment(horizontal="center")

    # 1.1 Dòng Tổng nhân công trên công trường
    r_last_task = r0 + len(rows) - 1
    r_labor = r_last_task + 2
    ws1.cell(r_labor, 3, f"TỔNG NHÂN CÔNG TRÊN CÔNG TRƯỜNG {project.upper()} (Người/ngày)")
    ws1.cell(r_labor, 3).font = Font(bold=True, size=10, color="FFFFFF")
    for col_idx in range(1, 16):
        ws1.cell(r_labor, col_idx).fill = FILL_H
    for w in range(num_weeks):
        c = g0 + w
        col_let = get_column_letter(c)
        ws1.cell(r_labor, c, f'=SUMPRODUCT(($J$7:$J${r_last_task}<={col_let}$6+6)*($K$7:$K${r_last_task}>={col_let}$6)*($O$7:$O${r_last_task}))')
        ws1.cell(r_labor, c).font = Font(bold=True, size=8.5, color="002060")
        ws1.cell(r_labor, c).fill = FILL_Y
        ws1.cell(r_labor, c).alignment = Alignment(horizontal="center")

    # 1.2 Bảng 2: Tổng hợp ca máy & Phương tiện huy động theo mốc thời gian
    r_mid_title = r_labor + 2
    ws1.cell(r_mid_title, 3, f"BẢNG TỔNG HỢP CA MÁY & SỐ PHƯƠNG TIỆN HUY ĐỘNG THEO TIẾN ĐỘ ({project.upper()})")
    ws1.cell(r_mid_title, 3).font = Font(bold=True, size=11, color="FFFFFF")
    for col_idx in range(1, 16):
        ws1.cell(r_mid_title, col_idx).fill = PatternFill("solid", fgColor="2E75B6")

    r_mid_hdr = r_mid_title + 1
    _hdr(ws1, r_mid_hdr, ["Mã", "STT", "Chủng loại phương tiện / Thiết bị", "ĐVT", "ĐM dầu (l/ca)", "Max máy", "Nhiệm vụ thi công chính"])
    for w in range(num_weeks):
        c = g0 + w
        col_let = get_column_letter(c)
        ws1.cell(r_mid_hdr, c, f"={col_let}$6")
        ws1.cell(r_mid_hdr, c).number_format = "DD/MM"
        ws1.cell(r_mid_hdr, c).font, ws1.cell(r_mid_hdr, c).fill = WHITE, PatternFill("solid", fgColor="2E75B6")
        ws1.cell(r_mid_hdr, c).alignment = Alignment(horizontal="center")

    mach_to_row_m = {}
    for j, m in enumerate(keys):
        r_m = r_mid_hdr + 1 + j
        mach_to_row_m[m] = r_m
        info = _machine_info(m, k, lib, vc)
        ws1.cell(r_m, 1, m).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 2, j + 1).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 3, info["name"])
        u = "xe" if "ô tô" in info["name"].lower() else "máy"
        ws1.cell(r_m, 4, u).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 5, info["diesel"]).number_format = "#,##0.0" if info["diesel"] else "@"
        ws1.cell(r_m, 6, 1).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 7, f"Thi công {project}")

        using_tasks = []
        for i_t, t_item in enumerate(rows):
            if t_item.get("norm") and m in t_item["norm"].get("machines", {}):
                using_tasks.append(r0 + i_t)

        for w in range(num_weeks):
            c = g0 + w
            col_let = get_column_letter(c)
            if using_tasks:
                terms = [f'IF(AND($J{rt}<={col_let}$6+6,$K{rt}>={col_let}$6),$M{rt},0)' for rt in using_tasks]
                ws1.cell(r_m, c, f"=MAX({','.join(terms)})")
            else:
                ws1.cell(r_m, c, 0)
            ws1.cell(r_m, c).font = Font(size=8, bold=True, color="1B365D")
            ws1.cell(r_m, c).fill = PatternFill("solid", fgColor="E2EFDA")
            ws1.cell(r_m, c).alignment = Alignment(horizontal="center")

    # 1.3 Bảng 3: Dầu Diezel tiêu thụ theo tiến độ thi công
    r_bot_title = r_mid_hdr + len(keys) + 2
    ws1.cell(r_bot_title, 3, f"BẢNG TÍNH DẦU DIEZEL TIÊU THỤ THEO TIẾN ĐỘ THI CÔNG {project.upper()} (Lít)")
    ws1.cell(r_bot_title, 3).font = Font(bold=True, size=11, color="FFFFFF")
    for col_idx in range(1, 16):
        ws1.cell(r_bot_title, col_idx).fill = PatternFill("solid", fgColor="385723")

    r_bot_tot = r_bot_title + 1
    ws1.cell(r_bot_tot, 3, "TỔNG SỐ LÍT DẦU DIEZEL TIÊU THỤ")
    ws1.cell(r_bot_tot, 3).font = Font(bold=True, size=10, color="C00000")
    for col_idx in range(1, 16):
        ws1.cell(r_bot_tot, col_idx).fill = FILL_Y

    mach_to_row_f = {}
    r_first_f = r_bot_tot + 1
    for j, m in enumerate(keys):
        r_f = r_first_f + j
        mach_to_row_f[m] = r_f
        info = _machine_info(m, k, lib, vc)
        ws1.cell(r_f, 1, m).alignment = Alignment(horizontal="center")
        ws1.cell(r_f, 3, f"Nhiên liệu dầu Diezel cho {info['name']}")
        ws1.cell(r_f, 4, "Lít").alignment = Alignment(horizontal="center")
        ws1.cell(r_f, 5, info["diesel"]).number_format = "#,##0.0" if info["diesel"] else "@"

        r_m_src = mach_to_row_m[m]
        for w in range(num_weeks):
            c = g0 + w
            col_let = get_column_letter(c)
            if info["diesel"]:
                ws1.cell(r_f, c, f"={col_let}{r_m_src}*$E{r_f}*2")
            else:
                ws1.cell(r_f, c, 0)
            ws1.cell(r_f, c).font = Font(size=7.5, color="333333")
            ws1.cell(r_f, c).alignment = Alignment(horizontal="center")
            ws1.cell(r_f, c).number_format = "#,##0"

    r_last_f = r_first_f + len(keys) - 1
    for w in range(num_weeks):
        c = g0 + w
        col_let = get_column_letter(c)
        ws1.cell(r_bot_tot, c, f"=SUM({col_let}{r_first_f}:{col_let}{r_last_f})")
        ws1.cell(r_bot_tot, c).font = Font(bold=True, size=8.5, color="C00000")
        ws1.cell(r_bot_tot, c).fill = FILL_Y
        ws1.cell(r_bot_tot, c).alignment = Alignment(horizontal="center")
        ws1.cell(r_bot_tot, c).number_format = "#,##0"

    for col, w in zip("ABCDEFGHIJKLMNO", (5, 8, 42, 8, 12, 14, 12, 12, 10, 12, 12, 10, 12, 45, 10)):
        ws1.column_dimensions[col].width = w

    # 2. Sheet 02_TongHop_CaXe_CaMay_MMTB (Chuẩn Vincons)
    ws2 = wb.create_sheet("02_TongHop_CaXe_CaMay_MMTB")
    ws2["A1"] = f"TỔNG HỢP CA MÁY & NHIÊN LIỆU — {project.upper()}"
    ws2["A1"].font = Font(bold=True, size=13)
    ws2["A2"] = "Dữ liệu tính toán từ tiến độ và định mức năng suất máy"

    _hdr(ws2, 3, ["STT", "Mã máy", "Thiết bị thi công", "ĐVT", "Nhiên liệu", "Định mức dầu (lít/ca)",
                  "Tổng số ca máy (ca)", "Số máy cao điểm", "Tổng lít dầu (lít)", "Thành phần thợ lái máy"])
    for j, m in enumerate(keys):
        r = 4 + j
        info = _machine_info(m, k, lib, vc)
        ws2.cell(r, 1, j + 1).alignment = Alignment(horizontal="center")
        ws2.cell(r, 2, m).alignment = Alignment(horizontal="center")
        ws2.cell(r, 3, info["name"])
        u = "xe" if "ô tô" in info["name"].lower() else "máy"
        ws2.cell(r, 4, u).alignment = Alignment(horizontal="center")
        ws2.cell(r, 5, info["fuel"]).alignment = Alignment(horizontal="center")
        ws2.cell(r, 6, info["diesel"]).number_format = "#,##0.0"
        if info["diesel"] is not None:
            ws2.cell(r, 6).fill = FILL_Y

        terms = mach_row_formulas.get(m, [])
        fml_ca = f"=ROUND({'+'.join(terms)},3)" if terms else "=0"
        ws2.cell(r, 7, fml_ca).number_format = "#,##0.000"

        ws2.cell(r, 8, 1).alignment = Alignment(horizontal="center")
        ws2.cell(r, 9, f"=IF(F{r}>0,ROUND(G{r}*F{r},1),0)").number_format = "#,##0.0"
        ws2.cell(r, 10, info["crew"])

    last = 3 + len(keys)
    ws2.cell(last + 1, 3, "TỔNG CỘNG").font = Font(bold=True)
    ws2.cell(last + 1, 7, f"=SUM(G4:G{last})").number_format = "#,##0.000"
    ws2.cell(last + 1, 7).font = Font(bold=True)
    ws2.cell(last + 1, 9, f"=SUM(I4:I{last})").number_format = "#,##0.0"
    ws2.cell(last + 1, 9).font = Font(bold=True)
    for col, w in zip("ABCDEFGHIJ", (5, 12, 46, 8, 10, 14, 14, 12, 14, 38)):
        ws2.column_dimensions[col].width = w

    # 3. Sheet 03_KeHoach_Dau_Diezel
    ws3 = wb.create_sheet("03_KeHoach_Dau_Diezel")
    ws3["A1"] = f"KẾ HOẠCH DẦU DIEZEL — {project.upper()}"
    ws3["A1"].font = Font(bold=True, size=13)
    ws3["A2"] = "Chỉ tổng hợp máy dùng diezel; máy dùng xăng/điện không tính lít dầu."
    _hdr(ws3, 3, ["STT", "Thiết bị", "Định mức (lít/ca)", "Tổng ca", "Tổng dầu (lít)"])
    for j, m in enumerate(keys):
        r = 4 + j
        ws3.cell(r, 1, j + 1).alignment = Alignment(horizontal="center")
        ws3.cell(r, 2, f"='02_TongHop_CaXe_CaMay_MMTB'!C{r}")
        ws3.cell(r, 3, f"='02_TongHop_CaXe_CaMay_MMTB'!F{r}").number_format = "#,##0.0"
        ws3.cell(r, 4, f"='02_TongHop_CaXe_CaMay_MMTB'!G{r}").number_format = "#,##0.000"
        ws3.cell(r, 5, f"=IF(C{r}=\"\",0,ROUND(C{r}*D{r},1))").number_format = "#,##0.0"
    ws3.cell(last + 1, 2, "TỔNG CỘNG").font = Font(bold=True)
    ws3.cell(last + 1, 5, f"=SUM(E4:E{last})").number_format = "#,##0.0"
    ws3.cell(last + 1, 5).font = Font(bold=True)
    for col, w in zip("ABCDE", (5, 46, 14, 12, 14)):
        ws3.column_dimensions[col].width = w

    # 4. Sheet 04_KeHoach_NhanLuc
    ws4 = wb.create_sheet("04_KeHoach_NhanLuc")
    ws4["A1"] = f"NHÂN LỰC THEO HẠNG MỤC — {project.upper()}"
    ws4["A1"].font = Font(bold=True, size=13)
    ws4["A2"] = "Nhân lực bố trí theo tiến độ và công tổ đội."
    _hdr(ws4, 3, ["STT", "WBS", "Hạng mục", "Thời gian (ngày)", "Số ca/ngày", "Nhân lực/tổ (người)", "Tổng công (người-ngày)"])
    for i, t in enumerate(rows):
        r = 4 + i
        ws4.cell(r, 1, i + 1).alignment = Alignment(horizontal="center")
        ws4.cell(r, 2, t["wbs"]).alignment = Alignment(horizontal="center")
        ws4.cell(r, 3, t["name"])
        ws4.cell(r, 4, f"='01_TienDo_CaMay_Master'!I{r0 + i}").alignment = Alignment(horizontal="center")
        ws4.cell(r, 5, f"='01_TienDo_CaMay_Master'!L{r0 + i}").alignment = Alignment(horizontal="center")
        ws4.cell(r, 6, f"='01_TienDo_CaMay_Master'!O{r0 + i}").alignment = Alignment(horizontal="center")
        ws4.cell(r, 7, f"=D{r}*E{r}*F{r}").number_format = "#,##0.0"
    last_nl = 3 + len(rows)
    ws4.cell(last_nl + 1, 3, "TỔNG CỘNG").font = Font(bold=True)
    ws4.cell(last_nl + 1, 7, f"=SUM(G4:G{last_nl})").number_format = "#,##0.0"
    ws4.cell(last_nl + 1, 7).font = Font(bold=True)
    for col, w in zip("ABCDEFG", (5, 8, 44, 12, 10, 14, 18)):
        ws4.column_dimensions[col].width = w

    # 5. Sheet 05_DoiChieu_BocTach (Lưu toàn bộ căn cứ định mức và độ tin cậy)
    ws5 = wb.create_sheet("05_DoiChieu_BocTach")
    ws5["A1"] = "NGUỒN GỐC ĐỊNH MỨC & CĂN CỨ SUY LUẬN CA MÁY"
    ws5["A1"].font = Font(bold=True, size=13)
    ws5["A2"] = "Bảng ma trận đối chiếu chi tiết các căn cứ kỹ thuật phục vụ kỹ sư kiểm toán, nghiệm thu."
    r = 4
    srcs = []
    if vc:
        s = vc["source"]
        srcs.append(f'Định mức nội bộ Vincons (MHT, VC, May): {s["project"]} ({s["file"]}, sha {s["sha256_16"]}; sheet {", ".join(s.get("sheets", ["MHT"]))})')
    if lib:
        s = lib["source"]
        srcs.append(f'Thư viện ĐM dự toán: {s["project"]} ({s["file"]}, sha {s["sha256_16"]}; sheet {", ".join(s["sheets"])})')
    if k:
        srcs += [f'Bảng ca máy dự án: {s["project"]} ({s["file"]}, sha {s["sha256_16"]})' for s in k["sources"]]
    for line in srcs + [
            "Quy tắc: (1) ĐM nội bộ nhà thầu Vincons; (2) ĐM dự toán xây dựng; (3) Năng suất trung vị công trình tương tự.",
            f'Trạng thái tri thức: {(vc or lib or k or {}).get("status", "PENDING_APPROVAL")} — cần kỹ sư xác nhận trước khi dùng điều phối thật.']:
        ws5.cell(r, 1, line)
        r += 1
    r += 1
    _hdr(ws5, r, ["WBS", "Hạng mục", "Nguồn định mức", "Mã hiệu / Công tác đối chiếu", "Đơn vị ĐM (hệ số)", "Máy (ca/đơn vị)", "Công/đơn vị", "Độ tin cậy & Thẩm định"])
    for t in rows:
        r += 1
        ws5.cell(r, 1, t["wbs"]).alignment = Alignment(horizontal="center")
        ws5.cell(r, 2, t["name"])
        n = t["norm"]
        if n:
            ws5.cell(r, 3, n["source"])
            ws5.cell(r, 4, n["ref"])
            ws5.cell(r, 5, n["scale"]).alignment = Alignment(horizontal="center")
            ws5.cell(r, 6, "; ".join(f"{m}={c}" for m, c in n["machines"].items()))
            ws5.cell(r, 7, n["labor"])
            ws5.cell(r, 8, n["confidence"])
        else:
            ws5.cell(r, 3, "Chưa áp định mức")
            ws5.cell(r, 4, t["status"])
    for col, w in zip("ABCDEFGH", (8, 42, 24, 35, 14, 30, 12, 50)):
        ws5.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(dest)), exist_ok=True)
    wb.save(dest)
    return dest


def infer_fleet_workbook_from_master(master_path: str, dest: str, project: str, knowledge_path: str = KNOWLEDGE_PATH,
                                     library_path: str = LIBRARY_PATH,
                                     vincons_path: str = VINCONS_PATH) -> Optional[str]:
    k, lib, vc = load_knowledge(knowledge_path), load_library(library_path), load_vincons(vincons_path)
    if not k and not lib and not vc:
        return None
    tasks = tasks_from_master(master_path)
    if not tasks:
        return None
    return write_fleet_workbook(dest, project, infer(tasks, k, lib, vc), k, lib, vc)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    pl = sub.add_parser("learn")
    pl.add_argument("files", nargs="+")
    pl.add_argument("--name", default="")
    pb = sub.add_parser("learn-lib")
    pb.add_argument("file")
    pb.add_argument("--name", default="")
    pv = sub.add_parser("learn-vincons")
    pv.add_argument("file")
    pv.add_argument("--name", default="")
    pi = sub.add_parser("infer")
    pi.add_argument("master")
    pi.add_argument("out")
    pi.add_argument("--project", default="Dự án")
    a = ap.parse_args(argv)
    if a.cmd == "learn":
        srcs = [learn_from_fleet_workbook(f, a.name if len(a.files) == 1 else "") for f in a.files]
        k = build_knowledge(srcs)
        print("Đã lưu:", save_knowledge(k))
        for wt, n in k["norms"].items():
            print(f'  {wt}: {n["prod_median"]} {n["unit"]}/ca (n={n["n"]}, {n["prod_min"]}–{n["prod_max"]}) {n["confidence"]}')
        return 0
    if a.cmd == "learn-lib":
        lib = learn_norm_library(a.file, a.name)
        print("Đã lưu:", _save(lib, LIBRARY_PATH))
        print(f'  {len(lib["tasks"])} công tác, {len(lib["machines"])} loại máy')
        return 0
    if a.cmd == "learn-vincons":
        vc = learn_vincons_norms(a.file, a.name)
        print("Đã lưu:", _save(vc, VINCONS_PATH))
        print(f'  {len(vc["works"])} công tác, {len(vc["machines"])} loại máy')
        if "transport" in vc:
            print(f'  Vận chuyển ({vc["transport"]["vehicle"]}): {vc["transport"]["distance_km"]}km -> NS {vc["transport"]["prod_per_ca_8h"]} m3/ca (chu kỳ {vc["transport"]["cycle_time_h"]}h)')
        if "organization" in vc:
            print(f'  Tổ chức ca máy: {vc["organization"]["driver_per_machine_per_shift"]} lái máy/máy/ca, {vc["organization"]["shifts_per_day"]} ca/ngày, {vc["organization"]["direct_labor_workers_peak"]} CNCH')
        return 0
    out = infer_fleet_workbook_from_master(a.master, a.out, a.project)
    print(out or "Chưa có tri thức (chạy `learn`/`learn-lib`/`learn-vincons` trước) hoặc Master không có tiến độ")
    return 0 if out else 1


if __name__ == "__main__":
    sys.exit(main())
