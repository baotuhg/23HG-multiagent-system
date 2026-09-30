# -*- coding: utf-8 -*-
"""
QS LOADER — Đọc bảng dự toán / tiên lượng (QS, BOQ) THẬT và tính G_XD
Hỗ trợ: Excel (.xlsx/.xlsm; công thức chưa có kết quả được tự tính bằng tools/excel_eval),
        CSV (.csv), JSON (.json)

Pure Python, Zero LLM. Không dùng tỷ lệ mặc định: tỷ lệ chi phí chung, nhà tạm, công việc
không xác định KL, thu nhập chịu thuế tính trước, VAT lấy từ sheet tổng hợp trong file
hoặc tham số dòng lệnh — thiếu tỷ lệ nào sẽ báo, không tự điền.

Nhận diện cột bảng QS theo tiêu đề (không phân biệt hoa/thường, có dấu hay không):
  TT / STT, Mã hiệu (mã định mức), Nội dung công tác, ĐVT, Khối lượng, Đơn giá, Thành tiền
  (tùy chọn: Đơn giá vật liệu / nhân công / máy — khi không có cột Đơn giá tổng)

Dòng công tác: có Mã hiệu hoặc STT, và có Đơn giá. Dòng bắt đầu bằng "-", "+", "•" là dòng
diễn giải của công tác phía trên (chỉ dùng để đối chiếu khối lượng). Dòng chỉ có chữ là
tiêu đề phần.

Cách tính (TT 36/2026/TT-BXD):
  T  = Σ khối lượng × đơn giá
  GT = C (chi phí chung) + LT (nhà tạm) + TT (công việc không xác định KL) = T × tỷ lệ
  TL = (T + GT) × tỷ lệ thu nhập chịu thuế tính trước
  G  = T + GT + TL;  VAT = G × thuế suất;  G_XD = G + VAT
"""

from __future__ import annotations
import csv
import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

from tools.bbs_loader import _norm, _to_number

HEADER_SCAN_ROWS = 40
RATE_KEYS = ("chung", "nha_tam", "kxd", "tl", "vat")
RATE_LABELS = {
    "chung": "Chi phí chung (% T)",
    "nha_tam": "Chi phí nhà tạm để ở và điều hành thi công (% T)",
    "kxd": "Chi phí một số công việc không xác định được KL từ thiết kế (% T)",
    "tl": "Thu nhập chịu thuế tính trước (% (T+GT))",
    "vat": "Thuế giá trị gia tăng (% G)",
}
DETAIL_PREFIXES = ("-", "+", "•", "*", "–")


class QSLoadError(ValueError):
    """Không đọc được bảng QS (sai định dạng, thiếu cột bắt buộc...)."""


@dataclass
class QSItem:
    row: int
    stt: str
    code: str
    description: str
    unit: str
    quantity: float
    unit_price: float
    amount: float                      # = quantity × unit_price (tính lại, không lấy từ file)
    section: str = ""
    file_amount: Optional[float] = None
    detail_quantity: Optional[float] = None   # tổng khối lượng các dòng diễn giải


@dataclass
class QSEstimate:
    source: str
    items: List[QSItem] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    rates: Dict[str, float] = field(default_factory=dict)          # tỷ lệ (0.051 = 5.1%)
    rate_sources: Dict[str, str] = field(default_factory=dict)
    file_total_T: Optional[float] = None
    file_G_XD: Optional[float] = None
    evaluated_cells: int = 0
    # Kết quả (VNĐ, làm tròn đồng)
    T: int = 0
    GT_components: Dict[str, int] = field(default_factory=dict)
    GT: int = 0
    TL: int = 0
    G: int = 0
    VAT: int = 0
    G_XD: int = 0

    @property
    def missing_rates(self) -> List[str]:
        return [k for k in RATE_KEYS if k not in self.rates]

    def compute(self) -> None:
        """Tính T, GT, TL, G, VAT, G_XD. Gọi sau khi đủ tỷ lệ."""
        if self.missing_rates:
            raise QSLoadError("Thiếu tỷ lệ: " + ", ".join(RATE_LABELS[k] for k in self.missing_rates))
        self.T = round(sum(i.amount for i in self.items))
        self.GT_components = {k: round(self.T * self.rates[k]) for k in ("chung", "nha_tam", "kxd")}
        self.GT = sum(self.GT_components.values())
        self.TL = round((self.T + self.GT) * self.rates["tl"])
        self.G = self.T + self.GT + self.TL
        self.VAT = round(self.G * self.rates["vat"])
        self.G_XD = self.G + self.VAT
        if self.file_total_T is not None and abs(self.file_total_T - self.T) > max(1000, 1e-6 * self.T):
            self.warnings.append(
                f"Tổng chi phí trực tiếp ghi trong file {self.file_total_T:,.0f} ≠ tính lại Σ KL×ĐG {self.T:,.0f}"
            )
        if self.file_G_XD is not None and abs(self.file_G_XD - self.G_XD) > max(1000, 1e-6 * self.G_XD):
            self.warnings.append(f"G_XD ghi trong file {self.file_G_XD:,.0f} ≠ tính lại {self.G_XD:,.0f}")


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────────────────────────────────────

def load_qs(path: str, sheet: Optional[str] = None) -> QSEstimate:
    """Đọc bảng QS + (nếu có) sheet tổng hợp G_XD để lấy tỷ lệ. Chưa tính G_XD (gọi .compute())."""
    if not os.path.exists(path):
        raise QSLoadError(f"Không tìm thấy file QS: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return _load_excel(path, sheet)
    if ext in (".csv", ".txt"):
        return _parse_items(_read_csv(path), source=path)
    if ext == ".json":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        items = data.get("items", []) if isinstance(data, dict) else data
        if not isinstance(items, list) or not all(isinstance(i, dict) for i in items):
            raise QSLoadError("File JSON QS phải là danh sách công tác hoặc {\"items\": [...], \"rates\": {...}}")
        keys: List[str] = []
        for i in items:
            keys.extend(k for k in i if k not in keys)
        est = _parse_items([keys] + [[i.get(k) for k in keys] for i in items], source=path)
        for k, v in (data.get("rates", {}) if isinstance(data, dict) else {}).items():
            if k in RATE_KEYS and _to_number(v) is not None:
                est.rates[k] = _as_rate(_to_number(v))
                est.rate_sources[k] = f"{path} (rates.{k})"
        return est
    raise QSLoadError(f"Định dạng QS chưa hỗ trợ: {ext} (dùng .xlsx, .csv hoặc .json)")


def apply_rate_overrides(est: QSEstimate, overrides: Dict[str, Optional[float]]) -> None:
    """Tỷ lệ từ dòng lệnh (phần trăm, vd 5.1) ghi đè tỷ lệ đọc từ file."""
    for key, value in overrides.items():
        if value is not None:
            est.rates[key] = value / 100
            est.rate_sources[key] = "tham số dòng lệnh"


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def _load_excel(path: str, sheet: Optional[str]) -> QSEstimate:
    from tools.excel_eval import WorkbookEvaluator

    ev = WorkbookEvaluator(path)
    if sheet and sheet not in ev.sheetnames:
        raise QSLoadError(f"Không có sheet '{sheet}' trong {path}. Các sheet: {ev.sheetnames}")

    est = None
    for name in ([sheet] if sheet else ev.sheetnames):
        rows = ev.rows(name, missing=_UNCALC)
        if _find_header(rows) is not None:
            est = _parse_items(rows, source=f"{path}#{name}")
            break
    if est is None:
        raise QSLoadError(
            f"Không tìm thấy bảng QS trong {path} — cần các cột Nội dung công tác, ĐVT, Khối lượng, Đơn giá"
        )

    # Tỷ lệ: lấy từ MỘT sheet tổng hợp G_XD — sheet có nhiều khoản mục chi phí nhận diện được nhất
    qs_sheet = est.source.split("#", 1)[1]
    best = None
    for name in ev.sheetnames:
        if name == qs_sheet:
            continue
        found = _read_rates(ev.rows(name, missing=_UNCALC), [
            [ev.formula(name, r, c) for c in range(1, ev.max_col(name) + 1)]
            for r in range(1, ev.max_row(name) + 1)
        ])
        if len(found[0]) >= 2 and (best is None or len(found[0]) > len(best[1][0])):
            best = (name, found)
    if best is not None:
        name, (rates, g_xd) = best
        for key, rate in rates.items():
            est.rates[key] = rate
            est.rate_sources[key] = f"{path}#{name}"
        est.file_G_XD = g_xd
    est.evaluated_cells = ev.evaluated_cells
    return est


class _Uncalc:
    def __repr__(self):
        return "<công thức không tính được>"


_UNCALC = _Uncalc()


def _read_csv(path: str) -> List[List[Any]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        text = f.read()
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [row for row in csv.reader(text.splitlines(), dialect)]


# ─────────────────────────────────────────────────────────────────────────────
# BẢNG QS
# ─────────────────────────────────────────────────────────────────────────────

def _detect_columns(headers: Sequence[Any]) -> Dict[str, int]:
    cols: Dict[str, int] = {}
    for idx, raw in enumerate(headers):
        if raw is None or raw is _UNCALC:
            continue
        h = _norm(raw)
        if not h:
            continue
        if "thanhtien" in h or h in ("amount", "total"):
            if not any(k in h for k in ("vatlieu", "nhancong", "may")):
                cols.setdefault("amount", idx)
        elif "dongia" in h and "vatlieu" in h or h in ("vl", "dongiavl"):
            cols.setdefault("price_vl", idx)
        elif "dongia" in h and "nhancong" in h or h in ("nc", "dongianc"):
            cols.setdefault("price_nc", idx)
        elif "dongia" in h and "may" in h or h in ("m", "mtc", "dongiamtc"):
            cols.setdefault("price_m", idx)
        elif "dongia" in h or h in ("unitprice", "price"):
            cols.setdefault("price", idx)
        elif h in ("tt", "stt", "no"):
            cols.setdefault("stt", idx)
        elif "mahieu" in h or "madinhmuc" in h or "madongia" in h or h in ("ma", "code"):
            cols.setdefault("code", idx)
        elif "noidung" in h or "tencongviec" in h or "tencongtac" in h or "hangmuc" in h \
                or h in ("description", "congtac", "congviec"):
            cols.setdefault("description", idx)
        elif h.startswith("dvt") or h.startswith("donvi") or h in ("unit",):
            cols.setdefault("unit", idx)
        elif "khoiluong" in h or h in ("quantity", "qty"):
            cols.setdefault("quantity", idx)
    return cols


def _find_header(rows: Sequence[Sequence[Any]]):
    for r, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        cols = _detect_columns(row)
        has_price = "price" in cols or any(k in cols for k in ("price_vl", "price_nc", "price_m"))
        if "description" in cols and "quantity" in cols and has_price:
            return r, cols
    return None


def _get(row: Sequence[Any], idx: Optional[int]) -> Any:
    if idx is None or idx >= len(row):
        return None
    v = row[idx]
    if isinstance(v, str) and not v.strip():
        return None
    return v


def _parse_items(rows: List[List[Any]], source: str) -> QSEstimate:
    found = _find_header(rows)
    if found is None:
        raise QSLoadError(
            f"{source}: không tìm thấy dòng tiêu đề QS — cần các cột Nội dung công tác, Khối lượng, Đơn giá"
        )
    header_idx, cols = found
    est = QSEstimate(source=source)
    section = ""
    current: Optional[QSItem] = None
    detail_sum: Optional[float] = None

    def close_item():
        if current is not None and detail_sum is not None:
            current.detail_quantity = detail_sum

    for offset, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        desc = _get(row, cols.get("description"))
        desc_text = "" if desc in (None, _UNCALC) else str(desc).strip()
        stt = _get(row, cols.get("stt"))
        code = _get(row, cols.get("code"))
        raw_qty = _get(row, cols.get("quantity"))
        raw_amount = _get(row, cols.get("amount"))
        price_cells = [_get(row, cols.get(k)) for k in ("price", "price_vl", "price_nc", "price_m")]
        has_price = any(p not in (None, _UNCALC) for p in price_cells)
        if all(v in (None, _UNCALC) for v in row):
            continue

        # Dòng diễn giải của công tác phía trên
        if desc_text.startswith(DETAIL_PREFIXES) and not has_price and code is None:
            q = _to_number(raw_qty) if raw_qty is not _UNCALC else None
            if q is not None:
                detail_sum = (detail_sum or 0.0) + q
            continue

        is_item = has_price or (code is not None and raw_qty is not None)
        if not is_item:
            if raw_amount not in (None, _UNCALC) and raw_qty in (None, _UNCALC):
                amount = _to_number(raw_amount)
                if amount is not None:
                    est.file_total_T = amount   # dòng tổng cộng (lấy dòng cuối cùng)
                continue
            if desc_text or stt is not None:
                close_item()
                current, detail_sum = None, None
                section = desc_text or str(stt).strip()
            continue

        close_item()
        current, detail_sum = None, None
        where = f"Dòng {offset} ({code or stt or ''} {desc_text[:50]})".strip()
        if raw_qty is _UNCALC or _UNCALC in price_cells:
            est.errors.append(f"{where}: có công thức không tính được — mở và lưu lại file bằng Excel")
            continue
        qty = _to_number(raw_qty)
        if qty is None or qty < 0:
            est.errors.append(f"{where}: khối lượng '{raw_qty}' không hợp lệ")
            continue
        if _get(row, cols.get("price")) is not None:
            price = _to_number(_get(row, cols["price"]))
        else:
            parts = [_to_number(p) for p in price_cells[1:] if p is not None]
            price = sum(parts) if parts and all(p is not None for p in parts) else None
        if price is None or price < 0:
            est.errors.append(f"{where}: chưa có hoặc sai đơn giá")
            continue
        file_amount = _to_number(raw_amount) if raw_amount not in (None, _UNCALC) else None
        current = QSItem(
            row=offset, stt="" if stt is None else str(stt).strip(), code="" if code is None else str(code).strip(),
            description=desc_text, unit=str(_get(row, cols.get("unit")) or "").strip(),
            quantity=qty, unit_price=price, amount=qty * price, section=section, file_amount=file_amount,
        )
        est.items.append(current)
    close_item()

    for item in est.items:
        label = f"{item.code or item.stt} {item.description[:50]}".strip()
        if item.file_amount is not None and abs(item.file_amount - item.amount) > max(1.0, 1e-6 * item.amount):
            est.warnings.append(
                f"{label}: thành tiền trong file {item.file_amount:,.0f} ≠ KL×ĐG {item.amount:,.0f}"
            )
        if item.detail_quantity is not None and abs(item.detail_quantity - item.quantity) > \
                max(1e-6, 0.005 * max(abs(item.quantity), abs(item.detail_quantity))):
            est.warnings.append(
                f"{label}: khối lượng {item.quantity:,.3f} {item.unit} khác tổng diễn giải "
                f"{item.detail_quantity:,.3f} (lệch {item.quantity - item.detail_quantity:+,.3f})"
            )
        if item.quantity == 0:
            est.warnings.append(f"{label}: khối lượng = 0")
    if not est.items and not est.errors:
        raise QSLoadError(f"{source}: bảng QS không có công tác nào có đơn giá")
    return est


# ─────────────────────────────────────────────────────────────────────────────
# TỶ LỆ TỪ SHEET TỔNG HỢP G_XD
# ─────────────────────────────────────────────────────────────────────────────

_RATE_PATTERNS = [
    ("chung", ("chiphichung",)),
    ("nha_tam", ("nhatam", "lantrai")),
    ("kxd", ("khongxacdinh",)),
    ("tl", ("thunhapchiuthue", "loinhuan")),
    ("vat", ("giatrigiatang",)),
]


def _as_rate(x: float) -> float:
    return x / 100 if x > 1 else x


def _rate_from_row(values: Sequence[Any], formulas: Sequence[Any]) -> Optional[float]:
    """Tỷ lệ: hằng số nhân trong công thức (=E6*0.051) hoặc 'x,xx%' trong chữ."""
    for f in formulas:
        if isinstance(f, str) and f.startswith("="):
            consts = [float(c) for c in re.findall(r"\*\s*\(?\s*(\d*\.\d+|\d+)\s*\)?(?![\d.:A-Za-z])", f)]
            consts = [c for c in consts if 0 < c < 1]
            if len(consts) == 1:
                return consts[0]
    for v in values:
        if isinstance(v, str):
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", v)
            if m:
                return float(m.group(1).replace(",", ".")) / 100
    return None


def _read_rates(rows, formula_rows) -> Tuple[Dict[str, float], Optional[float]]:
    """Tìm tỷ lệ các khoản mục và G_XD ghi trong một sheet tổng hợp."""
    rates: Dict[str, float] = {}
    g_xd: Optional[float] = None
    for values, formulas in zip(rows, formula_rows):
        raw = " ".join(str(v) for v in values if isinstance(v, str))
        text = _norm(raw)
        if not text:
            continue
        if "tongcong" in text and ("gxd" in text or "dutoan" in text):
            nums = [v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if nums:
                g_xd = float(nums[-1])
            continue
        for key, needles in _RATE_PATTERNS:
            hit = any(n in text for n in needles) or (key == "vat" and re.search(r"\bVAT\b", raw, re.I))
            if key in rates or not hit:
                continue
            if key == "vat" and "truocthue" in text:
                continue
            rate = _rate_from_row(values, formulas)
            if rate is not None:
                rates[key] = rate
            break
    return rates, g_xd
