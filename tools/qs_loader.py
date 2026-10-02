# -*- coding: utf-8 -*-
"""
QS LOADER — Đọc bảng dự toán / tiên lượng (QS, BOQ) THẬT và tính G_XD
Hỗ trợ: Excel (.xlsx/.xlsm; công thức chưa có kết quả được tự tính bằng tools/excel_eval),
        Excel 97-2003 (.xls; đọc giá trị đã lưu trong file, cần thư viện xlrd),
        CSV (.csv), JSON (.json)

Pure Python, Zero LLM. Không dùng tỷ lệ mặc định: tỷ lệ chi phí chung, nhà tạm, công việc
không xác định KL, thu nhập chịu thuế tính trước, VAT lấy từ sheet tổng hợp trong file
hoặc tham số dòng lệnh — thiếu tỷ lệ nào sẽ báo, không tự điền.

Nhận diện cột bảng QS theo tiêu đề (không phân biệt hoa/thường, có dấu hay không):
  TT / STT, Mã hiệu (mã định mức), Nội dung công tác, ĐVT, Khối lượng, Đơn giá, Thành tiền
  (tùy chọn: Đơn giá vật liệu / vật liệu phụ / nhân công / máy — khi không có cột Đơn giá tổng)
Tiêu đề 2 dòng (dòng trên "Đơn giá", "Thành tiền"; dòng dưới "Vật liệu", "Nhân công", "Máy thi
công") được ghép lại thành "Đơn giá Vật liệu"... như cách phần mềm dự toán thường xuất.

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
from tools.money import D, mul, round_vnd, total

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
    # Đơn giá tách theo thành phần (khi bảng có cột Đơn giá VL / NC / M); VL gồm cả vật liệu phụ
    price_vl: Optional[float] = None
    price_nc: Optional[float] = None
    price_m: Optional[float] = None


@dataclass
class QSEstimate:
    source: str
    hang_muc: str = ""
    items: List[QSItem] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    rates: Dict[str, float] = field(default_factory=dict)          # tỷ lệ (0.051 = 5.1%)
    rate_sources: Dict[str, str] = field(default_factory=dict)
    file_total_T: Optional[float] = None
    file_G_XD: Optional[float] = None
    evaluated_cells: int = 0
    # Kết quả (VNĐ, làm tròn đồng)
    VL: int = 0
    NC: int = 0
    M: int = 0
    T: int = 0
    GT_components: Dict[str, int] = field(default_factory=dict)
    GT: int = 0
    TL: int = 0
    G: int = 0
    VAT: int = 0
    G_XD: int = 0
    line_rounding: bool = False

    @property
    def missing_rates(self) -> List[str]:
        return [k for k in RATE_KEYS if k not in self.rates]

    def line_amount(self, item: "QSItem"):
        """Thành tiền chính xác (Decimal) của một dòng: KL × ĐG; nếu dòng chỉ có thành tiền trong file thì dùng số đó."""
        exact = mul(item.quantity, item.unit_price)
        if abs(float(exact) - item.amount) <= 1e-6 * max(1.0, abs(item.amount)):
            return exact
        return D(item.amount)

    def compute(self, line_rounding: bool = False) -> None:
        """Tính T, GT, TL, G, VAT, G_XD (Decimal, làm tròn half up như ROUND của Excel).

        line_rounding=False (mặc định): T = ROUND(Σ KL×ĐG) — cộng chính xác rồi làm tròn một lần.
        line_rounding=True : làm tròn từng dòng thành tiền đến đồng rồi cộng (cách nhiều bảng dự toán dùng).
        Hai cách có thể lệch vài đồng; chọn theo quy ước của bảng dự toán / hợp đồng cần đối chiếu.
        """
        if self.missing_rates:
            raise QSLoadError("Thiếu tỷ lệ: " + ", ".join(RATE_LABELS[k] for k in self.missing_rates))
        self.line_rounding = line_rounding
        if self.items and all(i.price_nc is not None for i in self.items):
            # Bảng tách đơn giá VL / NC / M: làm tròn từng thành phần từng dòng (như phần mềm dự toán)
            self.VL = sum(round_vnd(mul(i.quantity, i.price_vl)) for i in self.items)
            self.NC = sum(round_vnd(mul(i.quantity, i.price_nc)) for i in self.items)
            self.M = sum(round_vnd(mul(i.quantity, i.price_m)) for i in self.items)
            self.T = self.VL + self.NC + self.M
        else:
            amounts = [self.line_amount(i) for i in self.items]
            self.T = sum(round_vnd(a) for a in amounts) if line_rounding else round_vnd(total(amounts))
        self.GT_components = {k: round_vnd(mul(self.T, self.rates[k])) for k in ("chung", "nha_tam", "kxd")}
        self.GT = sum(self.GT_components.values())
        self.TL = round_vnd(mul(self.T + self.GT, self.rates["tl"]))
        self.G = self.T + self.GT + self.TL
        self.VAT = round_vnd(mul(self.G, self.rates["vat"]))
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

def load_qs(path: str, sheet: Optional[str] = None, hang_muc: Optional[str] = None) -> QSEstimate:
    """Đọc bảng QS + (nếu có) sheet tổng hợp G_XD để lấy tỷ lệ. Chưa tính G_XD (gọi .compute()).

    hang_muc: khi sheet có nhiều hạng mục, chọn hạng mục theo tên (khớp một phần, bỏ dấu).
    """
    if not os.path.exists(path):
        raise QSLoadError(f"Không tìm thấy file QS: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return _load_excel(path, sheet, hang_muc=hang_muc)
    if ext == ".xls":
        return _load_excel(path, sheet, book=_XlsBook(path), hang_muc=hang_muc)
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
    raise QSLoadError(f"Định dạng QS chưa hỗ trợ: {ext} (dùng .xlsx, .xls, .csv hoặc .json)")


def apply_rate_overrides(est: QSEstimate, overrides: Dict[str, Optional[float]]) -> None:
    """Tỷ lệ từ dòng lệnh (phần trăm, vd 5.1) ghi đè tỷ lệ đọc từ file."""
    for key, value in overrides.items():
        if value is not None:
            est.rates[key] = value / 100
            est.rate_sources[key] = "tham số dòng lệnh"


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────────────────────────────────────

_HANG_MUC = re.compile(r"^\s*HẠNG\s*MỤC\s*[:：]", re.I)


def _split_hang_muc(rows: Sequence[Sequence[Any]]) -> List[Tuple[str, List[List[Any]]]]:
    """Tách các khối hạng mục trong một sheet (dòng 'HẠNG MỤC: ...'). [] nếu < 2 hạng mục."""
    starts = []
    for r, row in enumerate(rows):
        for v in row:
            if isinstance(v, str) and _HANG_MUC.match(v):
                name = re.split(r"[:：]", v, maxsplit=1)[1].strip()
                starts.append((r, name))
                break
    if len(starts) < 2:
        return []
    out = []
    for i, (r, name) in enumerate(starts):
        end = starts[i + 1][0] if i + 1 < len(starts) else len(rows)
        block = [list(row) for row in rows[r:end]]
        # Cắt bỏ phần chân (TỔNG HẠNG MỤC + chữ ký) để không lẫn vào danh sách công tác;
        # giữ lại dòng TỔNG để lấy chi phí trực tiếp ghi trong file.
        for k, brow in enumerate(block):
            if any(isinstance(v, str) and re.search(r"(TỔNG|CỘNG)\s+HẠNG\s*MỤC", v, re.I) for v in brow):
                block = block[:k + 1]
                break
        out.append((name, block))
    return out


def _qs_sheet_rows(ev, sheet: Optional[str], path: str):
    """Tìm sheet chứa bảng QS và trả (tên sheet, rows)."""
    if sheet and sheet not in ev.sheetnames:
        raise QSLoadError(f"Không có sheet '{sheet}' trong {path}. Các sheet: {ev.sheetnames}")
    for name in ([sheet] if sheet else ev.sheetnames):
        rows = ev.rows(name, missing=_UNCALC)
        if _find_header(rows) is not None:
            return name, rows
    raise QSLoadError(
        f"Không tìm thấy bảng QS trong {path} — cần các cột Nội dung công tác, ĐVT, Khối lượng, Đơn giá"
    )


def load_all_hang_muc(path: str, sheet: Optional[str] = None) -> List[QSEstimate]:
    """Đọc mọi hạng mục trong file (mỗi hạng mục một dự toán). File một hạng mục → danh sách 1 phần tử."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        book = _XlsBook(path)
    elif ext in (".xlsx", ".xlsm"):
        from tools.excel_eval import WorkbookEvaluator
        book = _XlsxBook(WorkbookEvaluator(path))
    else:
        return [load_qs(path, sheet)]
    name, rows = _qs_sheet_rows(book, sheet, path)
    blocks = _split_hang_muc(rows)
    names = [b[0] for b in blocks] if blocks else [None]
    return [_load_excel(path, sheet, book=book, hang_muc=n) for n in names]


def _load_excel(path: str, sheet: Optional[str], book=None, hang_muc: Optional[str] = None) -> QSEstimate:
    if book is None:
        from tools.excel_eval import WorkbookEvaluator
        book = _XlsxBook(WorkbookEvaluator(path))
    ev = book

    name, rows = _qs_sheet_rows(ev, sheet, path)
    blocks = _split_hang_muc(rows)
    if blocks:
        chosen = next((b for b in blocks if hang_muc and _norm(hang_muc) in _norm(b[0])), blocks[0])
        est = _parse_items(chosen[1], source=f"{path}#{name}")
        est.hang_muc = chosen[0]
        if hang_muc is None and len(blocks) > 1:
            est.warnings.insert(0, f"Sheet '{name}' có {len(blocks)} hạng mục; đang đọc '{chosen[0]}'. "
                                   f"Chọn hạng mục khác bằng tham số hang_muc. Các hạng mục: "
                                   + "; ".join(b[0] for b in blocks))
    else:
        est = _parse_items(rows, source=f"{path}#{name}")

    # Tỷ lệ: lấy từ MỘT sheet tổng hợp G_XD — sheet có nhiều khoản mục chi phí nhận diện được nhất
    qs_sheet = name
    best = None
    for name in ev.sheetnames:
        if name == qs_sheet:
            continue
        found = _read_rates(ev.rows(name, missing=_UNCALC), ev.formula_rows(name))
        if len(found[0]) >= 2 and (best is None or len(found[0]) > len(best[1][0])):
            best = (name, found)
    if any(_looks_like_survey(ev.rows(n, missing=_UNCALC)) for n in ev.sheetnames):
        est.warnings.append(
            "File có bảng tổng hợp dự toán khảo sát (ký hiệu Gks): chi phí chung tính theo nhân công — "
            "dùng tools/survey_estimate (--survey) thay cho cách tính G_XD xây lắp"
        )
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


class _XlsxBook:
    """.xlsx/.xlsm qua WorkbookEvaluator (tự tính công thức chưa có kết quả)."""

    def __init__(self, ev):
        self._ev = ev
        self.sheetnames = ev.sheetnames

    def rows(self, name, missing=None):
        return self._ev.rows(name, missing=missing)

    def formula_rows(self, name):
        ev = self._ev
        return [[ev.formula(name, r, c) for c in range(1, ev.max_col(name) + 1)]
                for r in range(1, ev.max_row(name) + 1)]

    @property
    def evaluated_cells(self):
        return self._ev.evaluated_cells


class _XlsBook:
    """.xls (Excel 97-2003) qua xlrd: chỉ có giá trị đã lưu trong file, không có công thức."""

    def __init__(self, path: str):
        try:
            import xlrd
        except ImportError:
            raise QSLoadError("Đọc file .xls cần thư viện xlrd: pip install xlrd (hoặc lưu file sang .xlsx)")
        try:
            self._book = xlrd.open_workbook(path)
        except xlrd.XLRDError as e:
            raise QSLoadError(f"Không đọc được file .xls {path}: {e}")
        self._xlrd = xlrd
        self.sheetnames = self._book.sheet_names()
        self.evaluated_cells = 0

    def rows(self, name, missing=None):
        xlrd = self._xlrd
        sh = self._book.sheet_by_name(name)
        out = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    row.append(None)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    row.append(xlrd.error_text_from_code.get(cell.value, "#ERR"))
                elif cell.ctype == xlrd.XL_CELL_NUMBER and float(cell.value).is_integer():
                    row.append(int(cell.value))
                else:
                    row.append(cell.value)
            out.append(row)
        return out

    def formula_rows(self, name):
        return [[] for _ in range(self._book.sheet_by_name(name).nrows)]


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
            if "vatlieuphu" in h:
                cols.setdefault("amount_vlp", idx)
            elif "vatlieu" in h:
                cols.setdefault("amount_vl", idx)
            elif "nhancong" in h:
                cols.setdefault("amount_nc", idx)
            elif "may" in h:
                cols.setdefault("amount_m", idx)
            else:
                cols.setdefault("amount", idx)
        elif _is_price(h) and "vatlieuphu" in h:
            cols.setdefault("price_vlp", idx)
        elif _is_price(h) and "vatlieu" in h or h in ("vl", "dongiavl"):
            cols.setdefault("price_vl", idx)
        elif _is_price(h) and "nhancong" in h or h in ("nc", "dongianc"):
            cols.setdefault("price_nc", idx)
        elif _is_price(h) and "may" in h or h in ("m", "mtc", "dongiamtc"):
            cols.setdefault("price_m", idx)
        elif "dongia" in h or h in ("unitprice", "price"):
            cols.setdefault("price", idx)
        elif h in ("tt", "stt", "no"):
            cols.setdefault("stt", idx)
        elif "mahieu" in h or "madinhmuc" in h or "madongia" in h or h in ("ma", "code"):
            cols.setdefault("code", idx)
        elif "noidung" in h or "tencongviec" in h or "tencongtac" in h or "hangmuc" in h \
                or "danhmuccongtac" in h or h in ("description", "congtac", "congviec"):
            cols.setdefault("description", idx)
        elif h.startswith("dvt") or h.startswith("donvi") or h in ("unit",):
            cols.setdefault("unit", idx)
        elif "khoiluong" in h or h in ("quantity", "qty"):
            # Nhiều cột "Khối lượng": ưu tiên "toàn bộ / toàn công trình" (KL để nhân đơn giá)
            # hơn "một bộ phận / cấu kiện" (KL đơn vị). Cột đúng tên "Khối lượng" thắng tất cả.
            if h in ("khoiluong", "quantity", "qty"):
                score = 4
            elif any(w in h for w in ("toanbo", "toancongtrinh", "tongcong", "toanphan")):
                score = 3
            elif any(w in h for w in ("motbophan", "caukien", "mot", "donvi")):
                score = 1
            else:
                score = 2
            if score > cols.get("_qty_score", 0):
                cols["quantity"] = idx
                cols["_qty_score"] = score
    cols.pop("_qty_score", None)
    return cols


def _is_price(h: str) -> bool:
    """Tiêu đề thuộc nhóm đơn giá: 'Đơn giá' hoặc 'Tính trực tiếp' (đơn giá trực tiếp VL/NC/M)."""
    return "dongia" in h or "tinhtructiep" in h


_PRICE_PARTS = ("price_vl", "price_vlp", "price_nc", "price_m")
_AMOUNT_PARTS = ("amount_vl", "amount_vlp", "amount_nc", "amount_m")


def _blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _subheader_row(top: Sequence[Any], sub: Sequence[Any]) -> bool:
    """Dòng dưới là tiêu đề phụ: dòng trên có ≥ 3 ô tiêu đề (không phải dòng tên bảng), dòng dưới
    chỉ có chữ, và ≥ 2 ô nằm dưới ô trống của dòng trên (ô gộp)."""
    if sum(1 for v in top if not _blank(v)) < 3:
        return False
    cells = [v for v in sub if not _blank(v)]
    if len(cells) < 2 or any(not isinstance(v, str) for v in cells):
        return False
    under_blank = sum(1 for c, v in enumerate(sub)
                      if not _blank(v) and (c >= len(top) or _blank(top[c])))
    return under_blank >= 2


def _merge_header(top: Sequence[Any], sub: Sequence[Any]) -> List[Any]:
    """Ghép tiêu đề 2 dòng: ô trống của dòng trên (ô gộp) lấy chữ của ô bên trái gần nhất."""
    merged: List[Any] = []
    last_top = ""
    for c in range(max(len(top), len(sub))):
        t = top[c] if c < len(top) else None
        b = sub[c] if c < len(sub) else None
        if not _blank(t):
            last_top = str(t).strip()
        if _blank(b):
            merged.append(t)
        else:
            merged.append(f"{last_top} {str(b).strip()}".strip())
    return merged


def _find_header(rows: Sequence[Sequence[Any]]):
    """Trả (chỉ số dòng tiêu đề cuối cùng, cột). Tiêu đề 2 dòng được ghép trước khi nhận diện."""
    for r, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        candidates = []
        if r + 1 < len(rows) and _subheader_row(row, rows[r + 1]):
            candidates.append((r + 1, _merge_header(row, rows[r + 1])))
        candidates.append((r, row))
        for last, header in candidates:
            cols = _detect_columns(header)
            has_price = "price" in cols or any(k in cols for k in _PRICE_PARTS)
            if "description" in cols and "quantity" in cols and has_price:
                return last, cols
    return None


def _row_amount(row: Sequence[Any], cols: Dict[str, int]) -> Any:
    """Thành tiền của dòng: cột Thành tiền, hoặc tổng các cột Thành tiền VL / NC / M."""
    if "amount" in cols:
        return _get(row, cols["amount"])
    cells = [_get(row, cols.get(k)) for k in _AMOUNT_PARTS if k in cols]
    if _UNCALC in cells:
        return _UNCALC
    nums = [_to_number(v) for v in cells if v is not None]
    if not nums or any(n is None for n in nums):
        return None
    return sum(nums)


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
        raw_amount = _row_amount(row, cols)
        price_cells = [_get(row, cols.get(k)) for k in ("price",) + _PRICE_PARTS]
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
        # Đơn giá tổng nếu là số; cột "Đơn giá" chứa chữ (vd mã bộ đơn giá) thì dùng đơn giá thành phần
        part_cells = dict(zip(_PRICE_PARTS, price_cells[1:]))
        parts = {k: _to_number(v) for k, v in part_cells.items() if v is not None}
        price = _to_number(price_cells[0]) if price_cells[0] is not None else None
        if price is None and (price_cells[0] is None or parts):
            price = sum(parts.values()) if parts and all(v is not None for v in parts.values()) else None
        file_amount = _to_number(raw_amount) if raw_amount not in (None, _UNCALC) else None
        if price is None or price < 0:
            if file_amount == 0:            # dòng ghi rõ không tính tiền (đã trừ / không tính tiền)
                price = 0.0
                parts = {"price_vl": 0.0, "price_nc": 0.0, "price_m": 0.0}
            else:
                est.errors.append(f"{where}: chưa có hoặc sai đơn giá")
                continue
        current = QSItem(
            row=offset, stt="" if stt is None else str(stt).strip(), code="" if code is None else str(code).strip(),
            description=desc_text, unit=str(_get(row, cols.get("unit")) or "").strip(),
            quantity=qty, unit_price=price, amount=qty * price, section=section, file_amount=file_amount,
        )
        if parts and all(v is not None for v in parts.values()):
            current.price_vl = parts.get("price_vl", 0.0) + parts.get("price_vlp", 0.0)
            current.price_nc = parts.get("price_nc", 0.0)
            current.price_m = parts.get("price_m", 0.0)
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


_NC_BASE = re.compile(r"\bNC\s*[x×*]", re.I)
_SURVEY_MARK = re.compile(r"^\s*Gks\s*$")


def _looks_like_survey(rows) -> bool:
    """Bảng tổng hợp có ký hiệu Gks (giá thành khảo sát) → dự toán khảo sát xây dựng."""
    return any(isinstance(v, str) and _SURVEY_MARK.match(v) for row in rows for v in row)


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
            if key == "chung" and _NC_BASE.search(raw):
                break    # chi phí chung tính theo nhân công (dự toán khảo sát), không phải % của T
            rate = _rate_from_row(values, formulas)
            if rate is not None:
                rates[key] = rate
            break
    return rates, g_xd
