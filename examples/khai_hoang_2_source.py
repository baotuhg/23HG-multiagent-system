# -*- coding: utf-8 -*-
"""
NGUỒN DỮ LIỆU THẬT — CẦU THÔN KHAI HOANG 2, KM14+363.65 (bản vẽ ghi Km14+363,35)

Đọc 2 nguồn gốc, không tự bịa số liệu:
  1. bang_so_lieu.json  — bảng số liệu trích từ PDF hồ sơ BVTC (đã thẩm định SGTVT Hà Giang)
  2. THKL G9 - CHUẨN.xlsx, sheet "Cau 15m" — bảng tổng hợp khối lượng cầu

Mọi chỗ sửa số liệu OCR đều nằm trong OCR_FIXES kèm lý do, để kỹ sư đối chiếu bản gốc.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import openpyxl

SOURCE_DIR = (r"C:\Users\baotu\Downloads\Documents\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker"
              r"\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker")
TABLES_JSON = os.path.join(SOURCE_DIR, "bang_so_lieu.json")
THKL_XLSX = os.path.join(SOURCE_DIR, "THKL G9 - CHUẨN.xlsx")
THKL_SHEET = "Cau 15m"

# Trọng lượng đơn vị thép tròn (kg/m) = 0.00617 × D², TCVN 1651
UNIT_WEIGHT = {6: 0.222, 8: 0.395, 10: 0.617, 12: 0.888, 14: 1.208, 16: 1.578, 18: 1.998,
               20: 2.466, 22: 2.984, 25: 3.853, 28: 4.834, 32: 6.313}
GRADE = {d: ("CB240-T" if d <= 8 else "CB400-V") for d in UNIT_WEIGHT}
# Bản vẽ ghi toàn bộ là CB400-V (kể cả D6, D10) — giữ theo bản vẽ.
GRADE = {d: "CB400-V" for d in UNIT_WEIGHT}


@dataclass
class BarRow:
    component: str        # Hạng mục kết cấu
    mark: str             # Ký hiệu thanh
    dia: int              # mm
    length_mm: float
    qty_per_unit: int     # số thanh / 1 cấu kiện (theo bảng bản vẽ)
    units: int            # số cấu kiện toàn cầu
    drawing: str          # số hiệu bản vẽ
    page: int             # trang PDF
    drawing_kg: Optional[float] = None   # khối lượng ghi trên bản vẽ (1 cấu kiện)
    note: str = ""

    @property
    def total_qty(self) -> int:
        return self.qty_per_unit * self.units

    @property
    def kg_per_unit(self) -> float:
        return self.qty_per_unit * self.length_mm / 1000 * UNIT_WEIGHT[self.dia]

    @property
    def kg_total(self) -> float:
        return self.kg_per_unit * self.units


# Các bảng thép thường trên bản vẽ: (chỉ số bảng trong bang_so_lieu.json, hạng mục, số cấu kiện,
# cột ký hiệu, cột Ø, cột chiều dài, cột số lượng, cột KL bản vẽ, ghi chú số cấu kiện)
BBS_TABLES = [
    (5,  "Thép neo mố vào đá (toàn cầu)",          1, 1, 2, 4, 3, 6, "Bảng PD-01 tính cho toàn cầu: 98 lỗ khoan"),
    (7,  "Bệ mố - bảng 1 (tr.14)",                 1, 0, 1, 2, 4, 6, "Bảng ghi '01 bệ'; 2 bảng bệ khác nhau cho 2 mố"),
    (10, "Bệ mố - bảng 2 (tr.15)",                 1, 0, 1, 2, 4, 6, "Bảng ghi '01 bệ'; 2 bảng bệ khác nhau cho 2 mố"),
    (11, "Thân mố + tường đỉnh (1 mố)",            2, 0, 1, 2, 4, 6, "Giả định bảng PD-04 tính cho 1 mố, toàn cầu 2 mố"),
    (12, "Bệ kê gối + u neo dầm (1 mố)",           2, 0, 1, 2, 4, 6, "Giả định bảng PD-04 tính cho 1 mố, toàn cầu 2 mố"),
    (13, "Tường cánh (1 tường)",                   4, 0, 1, 2, 4, 6, "Bảng ghi 'tính cho 1 tường cánh'; mố chữ U: 4 tường cánh"),
    (14, "Bản quá độ (1 bản)",                     2, 0, 1, 2, 4, 6, "Bảng ghi 'tính cho 01 bản, toàn cầu có 02 bản'"),
    (15, "Khe co giãn (1 khe)",                    2, 0, 2, 3, 1, 6, "Bảng ghi toàn cầu 02 khe"),
    (18, "Dầm T L=15m - dầm biên (1 dầm)",         2, 1, 2, 4, 3, 6, "Mặt cắt ngang 4 dầm: 2 dầm biên"),
    (19, "Dầm T L=15m - dầm giữa (1 dầm)",         2, 1, 2, 4, 3, 6, "Mặt cắt ngang 4 dầm: 2 dầm giữa"),
    (20, "Lớp phủ mặt cầu + gờ lan can (1 nhịp)",  1, 1, 2, 4, 5, 6, "Cầu 1 nhịp"),
    (25, "Tường chắn đốt L=5,20m",                 1, 1, 2, 3, 4, 7, "1 đốt"),
    (26, "Tường chắn đốt L=5,89m",                 1, 1, 2, 3, 4, 7, "1 đốt"),
    (27, "Tường chắn đốt L=10,60m",                1, 1, 2, 3, 4, 7, "1 đốt"),
    (28, "Tường chắn đốt L=13,50m",                1, 1, 2, 3, 4, 7, "1 đốt"),
]

# Sửa lỗi OCR có căn cứ: (chỉ số bảng, ký hiệu, trường, giá trị đúng, lý do)
OCR_FIXES = [
    (11, "A2", "qty", 53, "OCR đọc '5.3'; tổng chiều dài 28,62m / 0,54m = 53 thanh"),
    (13, "FIB", "mark", "F1B", "OCR đọc 'FIB'"),
    (10, "FIB", "mark", "F1B", "OCR đọc 'FIB'"),
    (7, "FIA", "mark", "F1A", "OCR đọc 'FIA'"),
]

# Ký hiệu bị trùng trên bản vẽ (dầm T có 2 dòng 'D2B'): dòng thứ 2 đổi thành D2D
DUPLICATE_MARK_RENAMES = {"D2B": "D2D"}


def _num(text) -> Optional[float]:
    """Số VN/US trên bản vẽ → float. '8.467,00'→8467, '3,850'→3.85, '1,000.00'→1000, '5.3'→5.3."""
    if text is None:
        return None
    s = str(text).strip().replace(" ", "")
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s or not re.search(r"\d", s):
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")      # 8.467,00
        else:
            s = s.replace(",", "")                        # 1,000.00
    elif "," in s:
        s = s.replace(",", ".")                           # 3,850 → 3.850 ; 7262 giữ nguyên
    try:
        return float(s)
    except ValueError:
        return None


def _length_mm(text) -> Optional[float]:
    """Chiều dài (mm): '8.496' trên bản vẽ là 8496 (dấu chấm phân cách nghìn), không phải 8,496mm."""
    s = str(text or "").strip()
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        return float(s.replace(".", ""))
    return _num(s)


def _dia(text) -> Optional[int]:
    m = re.search(r"(\d{1,2})", str(text or "").replace("Φ", "").replace("Ø", ""))
    return int(m.group(1)) if m and int(m.group(1)) in UNIT_WEIGHT else None


def load_tables() -> List[dict]:
    with open(TABLES_JSON, encoding="utf-8") as f:
        return json.load(f)


def load_bbs() -> Tuple[List[BarRow], List[str]]:
    """Trả về (các dòng BBS, nhật ký sửa/cảnh báo)."""
    tables = load_tables()
    log: List[str] = []
    rows: List[BarRow] = []
    for (ti, comp, units, c_mark, c_dia, c_len, c_qty, c_kg, unit_note) in BBS_TABLES:
        t = tables[ti]
        drawing, page = t.get("sheet") or "", t["page"]
        seen = set()
        for r in t["rows"]:
            if len(r) <= max(c_mark, c_dia, c_len, c_qty):
                continue
            mark = str(r[c_mark]).strip()
            dia = _dia(r[c_dia])
            length = _length_mm(r[c_len])
            qty = _num(r[c_qty])
            if not mark or dia is None or not length or qty is None:
                continue
            if ti in (25, 26, 27, 28):   # tường chắn: cột ký hiệu là số → ghép bộ phận
                mark = f"TC{mark}"
            for fti, fmark, field_, val, why in OCR_FIXES:
                if fti == ti and fmark == mark:
                    if field_ == "qty":
                        log.append(f"[SỬA OCR] {drawing} tr.{page} {mark}: số lượng {qty} → {val} ({why})")
                        qty = val
                    elif field_ == "mark":
                        log.append(f"[SỬA OCR] {drawing} tr.{page} {mark} → {val} ({why})")
                        mark = val
            if mark in seen and mark in DUPLICATE_MARK_RENAMES:
                log.append(f"[TRÙNG KÝ HIỆU] {drawing} tr.{page}: '{mark}' xuất hiện 2 lần, dòng sau đổi thành "
                           f"'{DUPLICATE_MARK_RENAMES[mark]}' (L={length:.0f}mm) — cần TVTK xác nhận")
                mark = DUPLICATE_MARK_RENAMES[mark]
            seen.add(mark)
            kg = _num(r[c_kg]) if len(r) > c_kg else None
            rows.append(BarRow(comp, mark, dia, length, int(round(qty)), units, drawing, page, kg, unit_note))
    return rows, log


def check_bbs_weights(rows: List[BarRow], tol: float = 0.02) -> List[str]:
    """So khối lượng tính lại (L × n × kg/m chuẩn) với khối lượng ghi trên bản vẽ."""
    out = []
    for b in rows:
        if b.drawing_kg is None or b.drawing_kg <= 0:
            continue
        diff = (b.kg_per_unit - b.drawing_kg) / b.drawing_kg
        if abs(diff) > tol:
            out.append(f"[LỆCH KL] {b.drawing} tr.{b.page} {b.component} / {b.mark} Ø{b.dia} "
                       f"L={b.length_mm:.0f} n={b.qty_per_unit}: tính lại {b.kg_per_unit:,.2f} kg, "
                       f"bản vẽ ghi {b.drawing_kg:,.2f} kg ({diff:+.1%})")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# BẢNG THKL "Cau 15m"
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class BOQItem:
    code: str            # TT gốc trong THKL (vd "2.1-3")
    section: str         # phần / nhóm
    name: str
    unit: str
    qty: Optional[float]
    source_row: int


def load_thkl() -> List[BOQItem]:
    wb = openpyxl.load_workbook(THKL_XLSX, read_only=True, data_only=True)
    ws = wb[THKL_SHEET]
    items: List[BOQItem] = []
    part = group = ""
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=130, max_col=5, values_only=True), start=1):
        tt, name, unit, qty = r[0], r[1], r[2], r[3]
        if name is None or str(name).strip() in ("", "Hạng mục công việc"):
            continue
        name = str(name).strip()
        tt_s = "" if tt is None else str(tt).strip()
        if re.fullmatch(r"[IVX]+", tt_s):
            part, group = f"{tt_s}. {name}", ""
            continue
        if re.fullmatch(r"\d+\.\d+", tt_s):
            group = f"{tt_s} {name}"
            if unit is None:          # tiêu đề nhóm, không có KL
                continue
        grp_code = group.split(" ")[0] if group else part.split(".")[0]
        items.append(BOQItem(f"{grp_code}-{tt_s}", group or part, name,
                             "" if unit is None else str(unit).strip(),
                             float(qty) if isinstance(qty, (int, float)) else None, i))
    return items


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    rows, log = load_bbs()
    print("\n".join(log))
    print("\n".join(check_bbs_weights(rows)))
    from collections import defaultdict
    agg = defaultdict(lambda: [0.0, 0.0, 0.0])
    for b in rows:
        k = 0 if b.dia <= 10 else (1 if b.dia <= 18 else 2)
        agg[b.component][k] += b.kg_total
    for c, v in agg.items():
        print(f"{c:45s} D<=10 {v[0]:10.2f}  10<D<=18 {v[1]:10.2f}  D>18 {v[2]:10.2f}")
    print(len(rows), "dòng BBS")
    for it in load_thkl():
        print(it.code, "|", it.name, "|", it.unit, "|", it.qty)
