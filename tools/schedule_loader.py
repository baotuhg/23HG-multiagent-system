# -*- coding: utf-8 -*-
"""
SCHEDULE LOADER — Đọc danh mục công việc tiến độ THẬT từ file
Hỗ trợ: MS Project XML (.xml), Excel (.xlsx/.xlsm), CSV (.csv), JSON (.json)

Pure Python, Zero LLM. Không tự bịa dữ liệu: dòng nào sai được liệt kê trong
`errors` để người dùng sửa file, không âm thầm bỏ qua.

MS Project XML: bỏ qua công việc tổng hợp (Summary), đọc Duration (giờ → ngày theo
MinutesPerDay), PredecessorLink (FS/SS/FF/SF + LinkLag), StartDate của dự án.

Excel/CSV — nhận diện cột theo tiêu đề (không phân biệt hoa/thường, có dấu hay không):
  Mã WBS / Mã công việc / id                 → mã công việc (bắt buộc)
  Danh mục công tác / Tên công việc / name   → tên
  Thời gian (ngày) / Duration / duration     → thời lượng (ngày làm việc, bắt buộc)
  Quan hệ logic / Predecessors               → vd "1.2FS; 1.3SS+3d" ("-" = không có)
  Ngày bắt đầu / Start, Ngày hoàn thành / Finish (tùy chọn, dùng để đối chiếu)

Ví dụ CSV tối thiểu:
  id,name,duration,predecessors
  1.1,Tim mốc,2,
  1.2,Dọn mặt bằng,6,1.1FS
"""

from __future__ import annotations
import csv
import json
import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Sequence

from tools.bbs_loader import _FORMULA, _cell, _norm, _to_number
from tools.cpm_calculator import parse_predecessor

HEADER_SCAN_ROWS = 40
MSP_LINK_TYPES = {"0": "FF", "1": "FS", "2": "SF", "3": "SS"}


class ScheduleLoadError(ValueError):
    """Không đọc được file tiến độ (sai định dạng, thiếu cột bắt buộc...)."""


@dataclass
class ScheduleLoadResult:
    source: str
    tasks: List[Dict[str, Any]] = field(default_factory=list)  # định dạng input của CPMCalculator
    errors: List[str] = field(default_factory=list)
    project_start: str = ""   # ngày khởi công ghi trong file (nếu có)


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────────────────────────────────────

def load_schedule(path: str, sheet: Optional[str] = None) -> ScheduleLoadResult:
    if not os.path.exists(path):
        raise ScheduleLoadError(f"Không tìm thấy file tiến độ: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xml":
        return _load_msproject_xml(path)
    if ext in (".xlsx", ".xlsm"):
        return _load_excel(path, sheet)
    if ext in (".csv", ".txt"):
        return _parse_rows(_read_csv(path), source=path)
    if ext == ".json":
        return _parse_rows(_read_json(path), source=path)
    raise ScheduleLoadError(f"Định dạng tiến độ chưa hỗ trợ: {ext} (dùng .xml, .xlsx, .csv hoặc .json)")


# ─────────────────────────────────────────────────────────────────────────────
# MS PROJECT XML
# ─────────────────────────────────────────────────────────────────────────────

def _load_msproject_xml(path: str) -> ScheduleLoadResult:
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError as e:
        raise ScheduleLoadError(f"{path}: XML không hợp lệ — {e}") from e
    ns = {"p": root.tag[1:root.tag.index("}")]} if root.tag.startswith("{") else {"p": ""}

    def q(tag: str) -> str:
        return f"p:{tag}" if ns["p"] else tag

    def text(el, tag: str, default: str = "") -> str:
        found = el.find(q(tag), ns)
        return found.text.strip() if found is not None and found.text else default

    if root.find(q("Tasks"), ns) is None:
        raise ScheduleLoadError(f"{path}: không phải file MS Project XML (thiếu <Tasks>)")

    minutes_per_day = float(text(root, "MinutesPerDay", "480") or 480)
    result = ScheduleLoadResult(source=path, project_start=text(root, "StartDate")[:10])

    raw = []
    summary_uids = set()
    for t in root.find(q("Tasks"), ns).findall(q("Task"), ns):
        uid = text(t, "UID")
        name = text(t, "Name")
        if uid == "0" or (not name and not text(t, "Duration")):
            continue  # task 0 = dòng tổng dự án; dòng trống
        if text(t, "Summary") == "1":
            summary_uids.add(uid)
            continue
        raw.append((uid, name, t))

    for uid, name, t in raw:
        label = text(t, "WBS") or text(t, "OutlineNumber") or uid
        where = f"Công việc UID {uid} ({label} {name})"
        hours = _iso_duration_hours(text(t, "Duration"))
        if hours is None:
            result.errors.append(f"{where}: không đọc được Duration '{text(t, 'Duration')}'")
            continue
        preds = []
        for link in t.findall(q("PredecessorLink"), ns):
            pred_uid = text(link, "PredecessorUID")
            if pred_uid in summary_uids:
                result.errors.append(
                    f"{where}: liên kết tới công việc tổng hợp UID {pred_uid} — chưa hỗ trợ, "
                    f"hãy liên kết tới công việc chi tiết"
                )
                continue
            link_type = MSP_LINK_TYPES.get(text(link, "Type", "1"))
            if link_type is None:
                result.errors.append(f"{where}: loại liên kết không hợp lệ '{text(link, 'Type')}'")
                continue
            lag_tenth_minutes = float(text(link, "LinkLag", "0") or 0)
            preds.append({
                "id": pred_uid,
                "type": link_type,
                "lag": round(lag_tenth_minutes / 10 / minutes_per_day, 4),
            })
        result.tasks.append({
            "id": uid,
            "code": label,
            "name": f"{label} {name}".strip(),
            "duration": round(hours * 60 / minutes_per_day, 4),
            "predecessors": preds,
            "file_start": text(t, "Start")[:10],
            "file_finish": text(t, "Finish")[:10],
        })

    if not result.tasks and not result.errors:
        raise ScheduleLoadError(f"{path}: không có công việc chi tiết nào")
    return result


def _iso_duration_hours(value: str) -> Optional[float]:
    """PT16H0M0S / P2DT4H → số giờ (ngày trong ISO tính 24h theo chuẩn XML)."""
    m = re.fullmatch(
        r"P(?:(?P<d>\d+(?:\.\d+)?)D)?(?:T(?:(?P<h>\d+(?:\.\d+)?)H)?(?:(?P<m>\d+(?:\.\d+)?)M)?"
        r"(?:(?P<s>\d+(?:\.\d+)?)S)?)?",
        value or "",
    )
    if not m or not value or value == "P":
        return None
    parts = {k: float(v) if v else 0.0 for k, v in m.groupdict().items()}
    return parts["d"] * 24 + parts["h"] + parts["m"] / 60 + parts["s"] / 3600


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL / CSV / JSON
# ─────────────────────────────────────────────────────────────────────────────

def _load_excel(path: str, sheet: Optional[str]) -> ScheduleLoadResult:
    from tools.excel_eval import WorkbookEvaluator

    # Giá trị Excel lưu sẵn; ô công thức chưa có kết quả được tự tính, không tính được → _FORMULA
    ev = WorkbookEvaluator(path)
    if sheet and sheet not in ev.sheetnames:
        raise ScheduleLoadError(f"Không có sheet '{sheet}' trong {path}. Các sheet: {ev.sheetnames}")
    for name in ([sheet] if sheet else ev.sheetnames):
        rows = ev.rows(name, missing=_FORMULA)
        if _find_header(rows) is not None:
            return _parse_rows(rows, source=f"{path}#{name}")
    raise ScheduleLoadError(
        f"Không tìm thấy bảng tiến độ trong {path} — cần các cột mã công việc và thời gian (ngày)"
    )


def _read_csv(path: str) -> List[List[Any]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        text = f.read()
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [row for row in csv.reader(text.splitlines(), dialect)]


def _read_json(path: str) -> List[List[Any]]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    items = data.get("tasks", data.get("items", [])) if isinstance(data, dict) else data
    if not isinstance(items, list) or not all(isinstance(i, dict) for i in items):
        raise ScheduleLoadError("File JSON tiến độ phải là danh sách đối tượng hoặc {\"tasks\": [...]}")
    keys: List[str] = []
    for item in items:
        for k in item:
            if k not in keys:
                keys.append(k)
    rows = [keys]
    for item in items:
        row = []
        for k in keys:
            v = item.get(k)
            row.append("; ".join(map(str, v)) if isinstance(v, list) else v)
        rows.append(row)
    return rows


def _detect_columns(headers: Sequence[Any]) -> Dict[str, int]:
    cols: Dict[str, int] = {}
    for idx, raw in enumerate(headers):
        if raw is None or raw is _FORMULA:
            continue
        h = _norm(raw)
        if not h:
            continue
        if h.startswith("mawbs") or h in ("wbs", "id", "taskid", "macongviec", "ma", "macv"):
            cols.setdefault("id", idx)
        elif "danhmuc" in h or "tencongviec" in h or "congtac" in h or h in ("name", "taskname", "task"):
            cols.setdefault("name", idx)
        elif h.startswith("thoigian") or h.startswith("duration") or h in ("songay", "days"):
            cols.setdefault("duration", idx)
        elif "quanhe" in h or "predecessor" in h or "congviectruoc" in h:
            cols.setdefault("predecessors", idx)
        elif "ngaybatdau" in h or h.startswith("start"):
            cols.setdefault("start", idx)
        elif "ngayhoanthanh" in h or "ngayketthuc" in h or h.startswith("finish"):
            cols.setdefault("finish", idx)
    return cols


def _find_header(rows: Sequence[Sequence[Any]]):
    for r, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        cols = _detect_columns(row)
        if "id" in cols and "duration" in cols:
            return r, cols
    return None


def _as_date_str(v: Any) -> str:
    if v is None or v is _FORMULA:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def _parse_rows(rows: List[List[Any]], source: str) -> ScheduleLoadResult:
    found = _find_header(rows)
    if found is None:
        raise ScheduleLoadError(
            f"{source}: không tìm thấy dòng tiêu đề tiến độ — cần các cột mã công việc (Mã WBS / id) "
            f"và thời gian (Thời gian (ngày) / duration)"
        )
    header_idx, cols = found
    result = ScheduleLoadResult(source=source)

    for offset, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        tid = _cell(row, cols["id"])
        name = _cell(row, cols.get("name"))
        raw_dur = _cell(row, cols["duration"])
        raw_pred = _cell(row, cols.get("predecessors"))
        if tid is None and name is None:
            continue
        if raw_dur is None and raw_pred is None and name is None:
            continue  # dòng nhóm/tổng hợp (vd "1.0")
        tid = str(tid).strip() if tid is not None else ""
        where = f"Dòng {offset} ({tid} {name or ''})".rstrip()
        if not tid:
            result.errors.append(f"{where}: thiếu mã công việc")
            continue

        if raw_dur is _FORMULA:
            result.errors.append(f"{where}: thời gian là công thức chưa được tính — mở và lưu lại file bằng Excel")
            continue
        duration = _to_number(raw_dur)
        if duration is None or duration < 0:
            result.errors.append(f"{where}: thời gian '{raw_dur}' không hợp lệ")
            continue

        preds = []
        text = "" if raw_pred is None else str(raw_pred).strip()
        if text and text not in ("-", "—"):
            for part in re.split(r"[;,]", text):
                if not part.strip():
                    continue
                try:
                    link = parse_predecessor(part)
                    preds.append({"id": link.pred_id, "type": link.type, "lag": link.lag_days})
                except ValueError as e:
                    result.errors.append(f"{where}: {e}")

        start = _as_date_str(_cell(row, cols.get("start")))
        result.tasks.append({
            "id": tid,
            "code": tid,
            "name": f"{tid} {name}".strip() if name else tid,
            "duration": duration,
            "predecessors": preds,
            "file_start": start,
            "file_finish": _as_date_str(_cell(row, cols.get("finish"))),
        })

    starts = [t["file_start"] for t in result.tasks if t["file_start"]]
    result.project_start = min(starts) if starts else ""
    if not result.tasks and not result.errors:
        raise ScheduleLoadError(f"{source}: bảng tiến độ không có công việc nào")
    return result


# ─────────────────────────────────────────────────────────────────────────────
# KIỂM TRA NGÀY GHI TRONG FILE
# ─────────────────────────────────────────────────────────────────────────────

def find_date_violations(tasks: List[Dict[str, Any]]) -> List[str]:
    """
    Tìm các quan hệ logic mà chính ngày bắt đầu/kết thúc ghi trong file vi phạm
    (vd FS nhưng việc sau bắt đầu trước khi việc trước kết thúc). Chỉ xét phần
    chắc chắn sai (bỏ qua lag dương, lịch nghỉ) nên không báo nhầm.
    """
    by_id = {t["id"]: t for t in tasks}
    violations = []
    for t in tasks:
        for p in t.get("predecessors", []):
            pred = by_id.get(p["id"])
            if pred is None or p.get("lag", 0) < 0:
                continue
            s_start, s_finish = t.get("file_start"), t.get("file_finish")
            p_start, p_finish = pred.get("file_start"), pred.get("file_finish")
            kind = p["type"]
            checks = {
                "FS": (s_start, p_finish, lambda a, b: a > b, "bắt đầu", "kết thúc"),
                "SS": (s_start, p_start, lambda a, b: a >= b, "bắt đầu", "bắt đầu"),
                "FF": (s_finish, p_finish, lambda a, b: a >= b, "kết thúc", "kết thúc"),
                "SF": (s_finish, p_start, lambda a, b: a >= b, "kết thúc", "bắt đầu"),
            }
            mine, theirs, ok, my_word, their_word = checks[kind]
            if mine and theirs and not ok(mine, theirs):
                violations.append(
                    f"{t['name']} {my_word} {mine} nhưng quan hệ {kind} với "
                    f"{pred['name']} ({their_word} {theirs})"
                )
    return violations
