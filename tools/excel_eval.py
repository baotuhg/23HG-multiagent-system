# -*- coding: utf-8 -*-
"""
EXCEL EVAL — Tính giá trị công thức Excel chưa có kết quả lưu sẵn (Pure Python)

Khi nào cần: file .xlsx do phần mềm khác tạo (openpyxl, xuất từ web...) chỉ có công thức,
chưa từng được Excel tính → đọc bằng openpyxl(data_only=True) ra None. File đã mở và lưu
bằng Excel thì dùng luôn giá trị Excel đã tính (không tính lại).

Hỗ trợ (đủ cho bảng QS / BBS / tiến độ thông dụng):
  số, chuỗi, TRUE/FALSE, + - * / ^ & %, so sánh = <> < > <= >=, ngoặc, dấu âm, tham chiếu ô / vùng
  (có $, sang sheet khác); phép toán theo TỪNG PHẦN TỬ khi gặp vùng ô (vd SUMPRODUCT((A1:A9>0)*B1:B9));
  ngày tháng là số seri như Excel (1900, gốc 30/12/1899).
  SUM, SUMIF(S), COUNTIF(S), SUMPRODUCT, ROUND/ROUNDUP/ROUNDDOWN, MIN, MAX, AVERAGE, ABS, SQRT, PI, PRODUCT,
  LET (biến cục bộ, kể cả namespace Office 365), IF, AND, OR, NOT, IFERROR, ISERROR,
  VLOOKUP, MONTH, YEAR, DAY, DATEVALUE (dd/mm/yyyy hoặc yyyy-mm-dd),
  TEXT (định dạng "@", "0", "0.0", "#,##0", "#,##0.00", "0%", "dd/mm/yyyy", "mm/yyyy", "yyyy").
Lỗi Excel (#DIV/0!, #N/A, #VALUE!, #REF!, #NUM!) là GIÁ TRỊ lan truyền như trong Excel (IF/IFERROR xử lý được);
ô có kết quả lỗi → ExcelErrorResult (lớp con của FormulaError, thuộc tính .code).
Hàm khác → FormulaError (báo rõ, không đoán).
"""

from __future__ import annotations
import datetime as _dt
import fnmatch
import math
import re
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.utils import column_index_from_string


class FormulaError(ValueError):
    """Công thức không tính được (hàm chưa hỗ trợ, tham chiếu vòng, lỗi cú pháp...)."""


class ExcelErrorResult(FormulaError):
    """Ô tính ra một lỗi Excel thật (#DIV/0!, #N/A, ...) — khác với 'chưa hỗ trợ'."""

    def __init__(self, message: str, code: str):
        super().__init__(message)
        self.code = code


ERROR_CODES = ("#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NUM!", "#NAME?", "#NULL!")


class ExcelError:
    """Giá trị lỗi Excel lan truyền trong phép tính."""
    __slots__ = ("code",)

    def __init__(self, code: str):
        self.code = code

    def __eq__(self, other):
        return isinstance(other, ExcelError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)

    def __repr__(self):
        return self.code


_TOKEN = re.compile(r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w\.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
  | (?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+)
  | (?P<func>[A-Za-z_][A-Za-z0-9_\.]*)\s*\(
  | (?P<bool>(?i:TRUE|FALSE))\b
  | (?P<name>[A-Za-z_][A-Za-z0-9_\.]*)
  | (?P<op><>|<=|>=|[-+*/^&(),:<>=%])
""", re.VERBOSE)


class WorkbookEvaluator:
    """Đọc workbook 2 lần (công thức + giá trị lưu sẵn) và tính ô thiếu giá trị theo yêu cầu."""

    def __init__(self, path: str):
        import openpyxl
        self._formulas: Dict[str, Dict[Tuple[int, int], Any]] = {}
        self._cached: Dict[str, Dict[Tuple[int, int], Any]] = {}
        fwb = openpyxl.load_workbook(path, data_only=False, read_only=True)
        vwb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            for ws in fwb.worksheets:
                self._formulas[ws.title] = _cells(ws)
                self._cached[ws.title] = _cells(vwb[ws.title])
        finally:
            fwb.close()
            vwb.close()
        self._computed: Dict[Tuple[str, int, int], Any] = {}
        self._stack: set = set()
        self.evaluated_cells = 0   # số ô đã phải tự tính (không có giá trị Excel lưu sẵn)

    @property
    def sheetnames(self) -> List[str]:
        return list(self._formulas)

    def max_row(self, sheet: str) -> int:
        return max((r for r, _ in self._formulas[sheet]), default=0)

    def max_col(self, sheet: str) -> int:
        return max((c for _, c in self._formulas[sheet]), default=0)

    def formula(self, sheet: str, row: int, col: int) -> Any:
        return self._formulas[sheet].get((row, col))

    def value(self, sheet: str, row: int, col: int) -> Any:
        """Giá trị ô: giá trị Excel lưu sẵn, hoặc tự tính công thức.
        Không tính được → FormulaError; tính ra lỗi Excel → ExcelErrorResult (.code)."""
        v = self._value(sheet, row, col)
        if isinstance(v, ExcelError):
            raise ExcelErrorResult(f"{sheet}!{_addr(row, col)} = {v.code}", v.code)
        return v

    def _value(self, sheet: str, row: int, col: int) -> Any:
        """Như value() nhưng trả lỗi Excel dưới dạng ExcelError (để lan truyền trong công thức)."""
        if sheet not in self._formulas:
            raise FormulaError(f"Không có sheet '{sheet}'")
        cached = self._cached[sheet].get((row, col))
        raw = self._formulas[sheet].get((row, col))
        if cached is not None or not (isinstance(raw, str) and raw.startswith("=")):
            v = cached if cached is not None else raw
            return ExcelError(v) if isinstance(v, str) and v in ERROR_CODES else v
        key = (sheet, row, col)
        if key in self._computed:
            return self._computed[key]
        if key in self._stack:
            raise FormulaError(f"Tham chiếu vòng tại {sheet}!{_addr(row, col)}")
        self._stack.add(key)
        try:
            result = _Parser(self, sheet, raw[1:]).parse()
        except ExcelErrorResult:
            raise
        except FormulaError as e:
            raise FormulaError(f"{sheet}!{_addr(row, col)} '{raw}': {e}") from None
        finally:
            self._stack.discard(key)
        if isinstance(result, list):
            result = result[0][0] if result and result[0] else None
        self._computed[key] = result
        self.evaluated_cells += 1
        return result

    def rows(self, sheet: str, missing=None) -> List[List[Any]]:
        """Toàn bộ giá trị sheet theo dòng; ô công thức không tính được → `missing`."""
        out = []
        ncol = self.max_col(sheet)
        for r in range(1, self.max_row(sheet) + 1):
            row = []
            for c in range(1, ncol + 1):
                try:
                    row.append(self.value(sheet, r, c))
                except FormulaError:
                    row.append(missing)
            out.append(row)
        return out

    def range_values(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> List[List[Any]]:
        return [[self._value(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]


def _cells(ws) -> Dict[Tuple[int, int], Any]:
    out = {}
    for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for c, v in enumerate(row, start=1):
            if v is not None:
                out[(r, c)] = v
    return out


def _addr(row: int, col: int) -> str:
    from openpyxl.utils import get_column_letter
    return f"{get_column_letter(col)}{row}"


_EPOCH = _dt.datetime(1899, 12, 30)


def _serial(v: Any) -> float:
    """datetime/date/time → số seri Excel (hệ 1900)."""
    if isinstance(v, _dt.datetime):
        d = v - _EPOCH
        return d.days + d.seconds / 86400 + d.microseconds / 86400e6
    if isinstance(v, _dt.date):
        return float((_dt.datetime(v.year, v.month, v.day) - _EPOCH).days)
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    raise TypeError(v)


def _from_serial(x: float) -> _dt.datetime:
    return _EPOCH + _dt.timedelta(days=float(x))


def serial_to_date(v: Any) -> Optional[_dt.date]:
    """Số seri Excel (vd 46245) → date; trả None nếu không phải số seri hợp lý (năm 1955–2119).
    Dùng cho cột ngày: công thức ngày tháng tự tính ra số seri như Excel."""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    if not 20000 <= v <= 80000:
        return None
    return _from_serial(v).date()


def _num(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return _serial(v)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        raise FormulaError(f"giá trị '{v}' không phải số")


def _is_numlike(v: Any) -> bool:
    return (isinstance(v, (int, float)) and not isinstance(v, bool)) or isinstance(v, (_dt.datetime, _dt.date, _dt.time))


def _arith(v: Any):
    """Toán hạng cho phép toán số học: số, hoặc ExcelError (#VALUE! nếu là chữ không đổi được sang số)."""
    if isinstance(v, ExcelError):
        return v
    try:
        return _num(v)
    except FormulaError:
        return ExcelError("#VALUE!")


def _first_error(*values: Any):
    for v in values:
        for x in _flatten(v):
            if isinstance(x, ExcelError):
                return x
    return None


def _shape(v):
    return (len(v), len(v[0]) if v else 0) if isinstance(v, list) else None


def _elementwise(a: Any, b: Any, fn):
    """Áp fn từng phần tử khi có vùng ô (broadcast vô hướng); hai vùng phải cùng kích thước."""
    sa, sb = _shape(a), _shape(b)
    if sa is None and sb is None:
        return fn(a, b)
    if sa and sb and sa != sb:
        if sa == (1, 1):
            a, sa = a[0][0], None
        elif sb == (1, 1):
            b, sb = b[0][0], None
        else:
            return ExcelError("#VALUE!")
    shape = sa or sb
    get = lambda v, r, c: v[r][c] if isinstance(v, list) else v
    return [[fn(get(a, r, c), get(b, r, c)) for c in range(shape[1])] for r in range(shape[0])]


def _map(v: Any, fn):
    return [[fn(x) for x in row] for row in v] if isinstance(v, list) else fn(v)


def _rank(v: Any) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def _compare(op: str, a: Any, b: Any):
    if isinstance(a, ExcelError):
        return a
    if isinstance(b, ExcelError):
        return b
    a = "" if a is None and isinstance(b, str) else a
    b = "" if b is None and isinstance(a, str) else b
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        x, y = ra, rb
    elif ra == 1:
        x, y = a.casefold(), b.casefold()
    else:
        x, y = _num(a), _num(b)
    return {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]


def _binary(op: str, a: Any, b: Any):
    x, y = _arith(a), _arith(b)
    if isinstance(x, ExcelError):
        return x
    if isinstance(y, ExcelError):
        return y
    if op == "+":
        return x + y
    if op == "-":
        return x - y
    if op == "*":
        return x * y
    if op == "/":
        return ExcelError("#DIV/0!") if y == 0 else x / y
    try:
        return x ** y
    except (OverflowError, ZeroDivisionError):
        return ExcelError("#NUM!")


def _truthy(v: Any):
    v = _scalar(v)
    if isinstance(v, ExcelError):
        return v
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        return ExcelError("#VALUE!")
    return _num(v) != 0


_DATE_FORMATS = {"dd/mm/yyyy": "%d/%m/%Y", "dd/mm/yy": "%d/%m/%y", "mm/yyyy": "%m/%Y", "yyyy": "%Y",
                 "dd-mm-yyyy": "%d-%m-%Y", "yyyy-mm-dd": "%Y-%m-%d"}


def _excel_text(v: Any, fmt: str):
    """TEXT(giá trị, định dạng) cho các định dạng thông dụng; định dạng khác → FormulaError."""
    if isinstance(v, ExcelError):
        return v
    f = fmt.strip()
    if f == "@":
        return _text(v)
    if isinstance(v, str):                      # Excel thử đổi chữ sang ngày / số; không được thì trả nguyên chữ
        d = _datevalue(v)
        if isinstance(d, ExcelError):
            try:
                v = float(v.replace(",", "."))
            except ValueError:
                return v
        else:
            v = d
    key = f.lower()
    if key in _DATE_FORMATS:
        d = v if isinstance(v, _dt.datetime) else (_from_serial(_num(v)) if not isinstance(v, str) else None)
        if d is None:
            return ExcelError("#VALUE!")
        return d.strftime(_DATE_FORMATS[key])
    m = re.fullmatch(r"(#,##)?(0+)(\.(0+))?(%)?", f)
    if not m:
        raise FormulaError(f"TEXT: định dạng '{fmt}' chưa hỗ trợ")
    x = _arith(v)
    if isinstance(x, ExcelError):
        return x
    if m.group(5):
        x *= 100
    decimals = len(m.group(4) or "")
    from tools.money import round_half_up
    x = float(round_half_up(x, decimals))
    if m.group(1):
        out = f"{x:,.{decimals}f}"
    else:                                        # "00" → đệm số 0 tới đủ số chữ số phần nguyên
        width = len(m.group(2)) + (decimals + 1 if decimals else 0) + (1 if x < 0 else 0)
        out = f"{x:0{width}.{decimals}f}"
    return out + ("%" if m.group(5) else "")


def _datevalue(v: Any):
    if isinstance(v, ExcelError):
        return v
    if not isinstance(v, str):
        return ExcelError("#VALUE!")
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return float(int(_serial(_dt.datetime.strptime(v.strip(), fmt))))
        except ValueError:
            continue
    return ExcelError("#VALUE!")


def _lookup_equal(a: Any, b: Any) -> bool:
    if _rank(a) != _rank(b) or a is None or b is None:
        return False
    if isinstance(a, str):
        return a.casefold() == b.casefold()
    return _num(a) == _num(b)


def _vlookup(x: Any, table: Any, col: Any, approx: bool = True):
    err = _first_error(x, col)
    if err:
        return err
    if not isinstance(table, list):
        return ExcelError("#N/A")
    k = int(_num(col))
    if k < 1:
        return ExcelError("#VALUE!")
    if table and k > len(table[0]):
        return ExcelError("#REF!")
    if not approx:
        for row in table:
            if _lookup_equal(row[0], x):
                return row[k - 1]
        return ExcelError("#N/A")
    hit = None
    for row in table:
        key = row[0]
        if key is None or _rank(key) != _rank(x):
            continue
        if _compare("<=", key, x) is True:
            hit = row
        else:
            break
    return hit[k - 1] if hit else ExcelError("#N/A")


def _flatten(v: Any) -> List[Any]:
    if isinstance(v, list):
        return [x for row in v for x in row]
    return [v]


def _match(value: Any, criterion: Any) -> bool:
    """Điều kiện kiểu Excel: số, chuỗi có * ?, hoặc '>5', '<>0', '=abc'."""
    if isinstance(criterion, (int, float)) and not isinstance(criterion, bool):
        try:
            return value is not None and value != "" and _num(value) == float(criterion)
        except FormulaError:
            return False
    text = "" if criterion is None else str(criterion)
    m = re.match(r"^(<>|<=|>=|<|>|=)?(.*)$", text, re.S)
    op, operand = m.group(1) or "=", m.group(2)
    try:
        number = float(operand.replace(",", "."))
        is_number = operand.strip() != ""
    except ValueError:
        number, is_number = None, False
    if is_number and op in ("<", ">", "<=", ">=", "=", "<>"):
        try:
            v = _num(value) if value not in (None, "") else None
        except FormulaError:
            v = None
        if v is None:
            return op == "<>"
        return {"<": v < number, ">": v > number, "<=": v <= number, ">=": v >= number,
                "=": v == number, "<>": v != number}[op]
    s = "" if value is None else str(value)
    hit = fnmatch.fnmatchcase(s.lower(), operand.lower()) if any(ch in operand for ch in "*?") \
        else s.lower() == operand.lower()
    return hit if op == "=" else (not hit if op == "<>" else False)


class _Parser:
    def __init__(self, ev: WorkbookEvaluator, sheet: str, text: str):
        self.ev, self.sheet = ev, sheet
        self.bindings = {}
        self.tokens = []
        pos = 0
        while pos < len(text):
            m = _TOKEN.match(text, pos)
            if not m:
                raise FormulaError(f"không đọc được '{text[pos:pos + 15]}'")
            pos = m.end()
            kind = m.lastgroup
            if kind != "ws":
                self.tokens.append((kind, m.group(kind)))
        self.i = 0

    def peek(self):
        return self.tokens[self.i] if self.i < len(self.tokens) else (None, None)

    def take(self, value=None):
        tok = self.peek()
        if value is not None and tok[1] != value:
            raise FormulaError(f"cần '{value}', gặp '{tok[1]}'")
        self.i += 1
        return tok

    def parse(self):
        v = self.compare()
        if self.i != len(self.tokens):
            raise FormulaError(f"thừa '{self.peek()[1]}'")
        return v

    def compare(self):
        left = self.concat()
        while self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            right = self.concat()
            left = _elementwise(left, right, lambda x, y, op=op: _compare(op, x, y))
        return left

    def concat(self):
        left = self.additive()
        while self.peek()[1] == "&":
            self.take()
            right = self.additive()
            left = _elementwise(left, right, lambda x, y: _first_error(x, y) or f"{_text(x)}{_text(y)}")
        return left

    def additive(self):
        left = self.term()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            right = self.term()
            left = _elementwise(left, right, lambda x, y, op=op: _binary(op, x, y))
        return left

    def term(self):
        left = self.power()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            right = self.power()
            left = _elementwise(left, right, lambda x, y, op=op: _binary(op, x, y))
        return left

    def power(self):
        left = self.unary()
        while self.peek()[1] == "^":
            self.take()
            right = self.unary()
            left = _elementwise(left, right, lambda x, y: _binary("^", x, y))
        return left

    def unary(self):
        if self.peek()[1] == "-":
            self.take()
            return _map(self.unary(), lambda x: _binary("-", 0.0, x))
        if self.peek()[1] == "+":
            self.take()
            return self.unary()
        v = self.primary()
        while self.peek()[1] == "%":
            self.take()
            v = _map(v, lambda x: _binary("/", x, 100.0))
        return v

    def primary(self):
        kind, text = self.take()
        if kind == "num":
            return float(text)
        if kind == "str":
            return text[1:-1].replace('""', '"')
        if kind == "bool":
            return text.upper() == "TRUE"
        if kind == "name":
            key = text.upper().removeprefix("_XLPM.")
            if key not in self.bindings:
                raise FormulaError(f"Tên chưa được khai báo: {text}")
            return self.bindings[key]
        if kind == "ref":
            return self.reference(text)
        if kind == "func":
            return self.function(text.upper())
        if text == "(":
            v = self.compare()
            self.take(")")
            return v
        raise FormulaError(f"không hiểu '{text}'")

    def reference(self, text: str):
        sheet = self.sheet
        if "!" in text:
            sheet, text = text.rsplit("!", 1)
            if sheet.startswith("'"):
                sheet = sheet[1:-1].replace("''", "'")
        parts = text.replace("$", "").split(":")
        coords = []
        for p in parts:
            m = re.match(r"([A-Za-z]+)(\d+)", p)
            coords.append((int(m.group(2)), column_index_from_string(m.group(1).upper())))
        if len(coords) == 1:
            return self.ev._value(sheet, *coords[0])
        (r1, c1), (r2, c2) = coords
        return self.ev.range_values(sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))

    def args(self) -> List[Any]:
        out = []
        if self.peek()[1] == ")":
            self.take()
            return out
        while True:
            out.append(self.compare())
            if self.peek()[1] == ",":
                self.take()
                continue
            self.take(")")
            return out

    def function(self, name: str):
        name = name.removeprefix("_XLFN.")
        if name == "LET":
            outer = self.bindings
            self.bindings = dict(outer)
            count = 0
            try:
                while self.peek()[0] == "name" and self.i + 1 < len(self.tokens) and self.tokens[self.i + 1][1] == ",":
                    key = self.take()[1].upper().removeprefix("_XLPM.")
                    self.take(",")
                    value = self.compare()
                    self.take(",")
                    self.bindings[key] = value
                    count += 1
                if not count:
                    raise FormulaError("LET cần ít nhất một cặp tên/giá trị")
                value = self.compare()
                self.take(")")
                return value
            finally:
                self.bindings = outer
        a = self.args()
        nums = lambda values: [_num(x) for v in values for x in _flatten(v) if _is_numlike(x)]
        # Các hàm xử lý lỗi / điều kiện: KHÔNG lan truyền lỗi của nhánh không dùng
        if name == "IFERROR":
            v = _scalar(a[0])
            return a[1] if isinstance(v, ExcelError) else a[0]
        if name == "ISERROR":
            return isinstance(_scalar(a[0]), ExcelError)
        if name == "IF":
            cond = _truthy(a[0])
            if isinstance(cond, ExcelError):
                return cond
            if cond:
                return a[1] if len(a) > 1 else True
            return a[2] if len(a) > 2 else False
        if name in ("AND", "OR"):
            vals = []
            for v in a:
                for x in _flatten(v):
                    if isinstance(x, ExcelError):
                        return x
                    if x is None or (isinstance(x, str) and isinstance(v, list)):
                        continue
                    t = _truthy(x)
                    if isinstance(t, ExcelError):
                        return t
                    vals.append(t)
            if not vals:
                return ExcelError("#VALUE!")
            return all(vals) if name == "AND" else any(vals)
        if name == "NOT":
            t = _truthy(a[0])
            return t if isinstance(t, ExcelError) else (not t)
        if name == "VLOOKUP":
            approx = True if len(a) < 4 else _truthy(a[3])
            if isinstance(approx, ExcelError):
                return approx
            return _vlookup(_scalar(a[0]), a[1], _scalar(a[2]), bool(approx))
        if name == "TEXT":
            return _excel_text(_scalar(a[0]), _text(_scalar(a[1])))
        if name == "DATEVALUE":
            return _datevalue(_scalar(a[0]))
        # Các hàm còn lại: lỗi trong đối số → lỗi (như Excel)
        err = _first_error(*a)
        if err:
            return err
        if name == "SUM":
            return sum(nums(a))
        if name == "SUMPRODUCT":
            arrays = [v if isinstance(v, list) else [[v]] for v in a]
            shape = _shape(arrays[0])
            if any(_shape(x) != shape for x in arrays):
                return ExcelError("#VALUE!")
            total = 0.0
            for r in range(shape[0]):
                for c in range(shape[1]):
                    prod = 1.0
                    for arr in arrays:
                        x = arr[r][c]
                        prod *= _num(x) if _is_numlike(x) else 0.0
                    total += prod
            return total
        if name == "PRODUCT":
            return math.prod(nums(a))
        if name in ("MIN", "MAX"):
            values = nums(a)
            return (min if name == "MIN" else max)(values) if values else 0.0
        if name == "AVERAGE":
            values = nums(a)
            if not values:
                return ExcelError("#DIV/0!")
            return sum(values) / len(values)
        if name == "ABS":
            return abs(_num(_scalar(a[0])))
        if name == "PI":
            return math.pi
        if name == "SQRT":
            x = _num(_scalar(a[0]))
            return ExcelError("#NUM!") if x < 0 else math.sqrt(x)
        if name in ("MONTH", "YEAR", "DAY"):
            v = _scalar(a[0])
            if isinstance(v, str):
                return ExcelError("#VALUE!")
            d = v if isinstance(v, _dt.datetime) else _from_serial(_num(v))
            return float({"MONTH": d.month, "YEAR": d.year, "DAY": d.day}[name])
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            x, d = _num(_scalar(a[0])), int(_num(_scalar(a[1])) if len(a) > 1 else 0)
            f = 10 ** d
            if name == "ROUND":
                return math.floor(abs(x) * f + 0.5) / f * (1 if x >= 0 else -1)
            op = math.ceil if name == "ROUNDUP" else math.floor
            return op(abs(x) * f - 1e-9 if name == "ROUNDUP" else abs(x) * f + 1e-9) / f * (1 if x >= 0 else -1)
        if name in ("SUMIF", "COUNTIF"):
            rng, crit = _flatten(a[0]), _scalar(a[1])
            if name == "COUNTIF":
                return float(sum(1 for v in rng if _match(v, crit)))
            sums = _flatten(a[2]) if len(a) > 2 else rng
            return sum(_num(s) for v, s in zip(rng, sums) if _match(v, crit) and _is_numlike(s))
        if name in ("SUMIFS", "COUNTIFS"):
            if name == "SUMIFS":
                target, pairs = _flatten(a[0]), a[1:]
            else:
                target, pairs = None, a
            if len(pairs) % 2:
                raise FormulaError(f"{name} thiếu điều kiện")
            ranges = [(_flatten(pairs[k]), _scalar(pairs[k + 1])) for k in range(0, len(pairs), 2)]
            n = len(ranges[0][0])
            hits = [all(_match(rng[i], crit) for rng, crit in ranges) for i in range(n)]
            if target is None:
                return float(sum(hits))
            return sum(_num(target[i]) for i in range(n) if hits[i] and _is_numlike(target[i]))
        raise FormulaError(f"hàm {name} chưa hỗ trợ")


def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _scalar(v: Any) -> Any:
    return _flatten(v)[0] if isinstance(v, list) else v


def _text(v: Any) -> str:
    v = _scalar(v)
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)
