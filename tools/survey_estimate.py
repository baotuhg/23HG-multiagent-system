# -*- coding: utf-8 -*-
"""
SURVEY ESTIMATE — Dự toán chi phí khảo sát xây dựng (KSXD)

Pure Python, Zero LLM. Đọc bảng khối lượng × đơn giá (tách VL / NC / M) bằng tools/qs_loader,
tỷ lệ các khoản mục đọc từ bảng tổng hợp dự toán khảo sát trong file (cột KÝ HIỆU + CÁCH TÍNH)
và sheet "Hệ số" nếu có. Không có tỷ lệ mặc định: thiếu tỷ lệ nào sẽ báo, không tự điền.

Cấu trúc tính (theo bảng tổng hợp dự toán khảo sát; ký hiệu như trong file):
  VL, NC, M = Σ ROUND(KL × đơn giá thành phần)           (làm tròn từng dòng như bảng đo bóc)
  T   = VL + NC + M                                      chi phí trực tiếp
  C   = NC × tỷ lệ                                       chi phí chung (theo chi phí nhân công)
  TL  = (T + C) × tỷ lệ                                  thu nhập chịu thuế tính trước
  Gks = T + C + TL                                       giá thành khảo sát
  Glpa = Gks × tỷ lệ;  Glbc = Gks × tỷ lệ                lập phương án, lập báo cáo kết quả khảo sát
  Ghmc = Gco + Gdc + Ggt + Gbh  (mỗi khoản = Gks × tỷ lệ)  chi phí hạng mục chung
  G   = Gks + Glpa + Glbc + Ghmc                         dự toán khảo sát trước thuế
  GTGT = G × thuế suất;  Gxd = G + GTGT;  Gdp = Gxd × tỷ lệ;  Tổng = Gxd + Gdp

Các giá trị trung gian giữ chính xác (Decimal), chỉ làm tròn tổng cuối khi xuất — như bảng tính.
Nếu CÁCH TÍNH của một khoản trong file dùng cơ sở khác (vd C = T × %) thì báo lỗi thay vì tính
theo cấu trúc trên.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Sequence, Tuple

from tools.bbs_loader import _to_number
from tools.money import D, round_vnd
from tools.qs_loader import QSEstimate, QSLoadError, load_qs

# khóa → (ký hiệu trong bảng tổng hợp, mã trong sheet Hệ số, nhãn, cơ sở bắt buộc trong CÁCH TÍNH)
SURVEY_RATES: Dict[str, Tuple[str, str, str, str]] = {
    "chung_nc": ("C", "hsCNC", "Chi phí chung (% chi phí nhân công)", "NC"),
    "tl": ("TL", "hsTL", "Thu nhập chịu thuế tính trước (% (T + C))", "T"),
    "lpa": ("Glpa", "hsGlpa", "Lập phương án kỹ thuật khảo sát (% Gks)", "Gks"),
    "lbc": ("Glbc", "hsGlbc", "Lập báo cáo kết quả khảo sát (% Gks)", "Gks"),
    "co": ("Gco", "hsGxdnt", "Chỗ ở tạm thời tại hiện trường (% Gks)", "Gks"),
    "dc": ("Gdc", "hsGdc", "Di chuyển máy và thiết bị khảo sát (% Gks)", "Gks"),
    "atgt": ("Ggt", "hsDBATGT", "Đảm bảo an toàn giao thông (% Gks)", "Gks"),
    "bh": ("Gbh", "hsGbhtn", "Bảo hiểm trách nhiệm nghề nghiệp (% Gks)", "Gks"),
    "vat": ("GTGT", "hsGTGT", "Thuế giá trị gia tăng (% G)", "G"),
    "dp": ("Gdp", "hsGdp", "Chi phí dự phòng (% Gxd)", "Gxd"),
}
_SYMBOL_TO_KEY = {v[0]: k for k, v in SURVEY_RATES.items()}
_HS_TO_KEY = {v[1]: k for k, v in SURVEY_RATES.items()}
# Giá trị ghi trong file để đối chiếu (ký hiệu → tên kết quả)
_FILE_VALUES = {"T": "T", "C": "C", "TL": "TL", "Gks": "Gks", "Glpa": "Glpa", "Glbc": "Glbc",
                "Ghmc": "Ghmc", "G": "G", "GTGT": "GTGT", "Gxd": "Gxd", "Gdp": "Gdp"}
_PERCENT = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")


@dataclass
class SurveyEstimate:
    qs: QSEstimate
    rates: Dict[str, float] = field(default_factory=dict)
    rate_sources: Dict[str, str] = field(default_factory=dict)
    file_values: Dict[str, float] = field(default_factory=dict)   # giá trị ghi trong bảng tổng hợp
    file_total: Optional[float] = None                            # dòng "Tổng cộng"
    file_total_rounded: Optional[float] = None                    # dòng "Làm tròn"
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    results: Dict[str, Any] = field(default_factory=dict)         # Decimal (VL/NC/M/T: số nguyên)

    @property
    def missing_rates(self) -> List[str]:
        return [k for k in SURVEY_RATES if k not in self.rates]

    def compute(self) -> Dict[str, Any]:
        if self.missing_rates:
            raise QSLoadError("Thiếu tỷ lệ dự toán khảo sát: " +
                              ", ".join(SURVEY_RATES[k][2] for k in self.missing_rates))
        no_split = [i for i in self.qs.items if i.price_nc is None]
        if no_split:
            raise QSLoadError(
                "Dự toán khảo sát cần đơn giá tách Vật liệu / Nhân công / Máy; thiếu ở: " +
                ", ".join(f"dòng {i.row} {i.code}" for i in no_split[:5])
            )
        r = {k: D(v) for k, v in self.rates.items()}
        vl = sum(round_vnd(D(i.quantity) * D(i.price_vl)) for i in self.qs.items)
        nc = sum(round_vnd(D(i.quantity) * D(i.price_nc)) for i in self.qs.items)
        m = sum(round_vnd(D(i.quantity) * D(i.price_m)) for i in self.qs.items)
        T = D(vl + nc + m)
        C = D(nc) * r["chung_nc"]
        TL = (T + C) * r["tl"]
        Gks = T + C + TL
        Glpa, Glbc = Gks * r["lpa"], Gks * r["lbc"]
        hmc = {k: Gks * r[k] for k in ("co", "dc", "atgt", "bh")}
        Ghmc = sum(hmc.values())
        G = Gks + Glpa + Glbc + Ghmc
        GTGT = G * r["vat"]
        Gxd = G + GTGT
        Gdp = Gxd * r["dp"]
        self.results = {
            "VL": vl, "NC": nc, "M": m, "T": T, "C": C, "TL": TL, "Gks": Gks, "Glpa": Glpa,
            "Glbc": Glbc, "Gco": hmc["co"], "Gdc": hmc["dc"], "Ggt": hmc["atgt"], "Gbh": hmc["bh"],
            "Ghmc": Ghmc, "G": G, "GTGT": GTGT, "Gxd": Gxd, "Gdp": Gdp, "Tong": Gxd + Gdp,
        }
        self._check_against_file()
        return self.results

    def total_rounded(self, unit: int = 1) -> int:
        """Tổng cộng làm tròn đến `unit` đồng (1, 1000...), half up như ROUND của Excel."""
        q = D(self.results["Tong"]) / D(unit)
        return int(q.quantize(D(1), rounding=ROUND_HALF_UP)) * unit

    def _check_against_file(self) -> None:
        for sym, key in _FILE_VALUES.items():
            if sym in self.file_values and abs(D(self.file_values[sym]) - D(self.results[key])) > 1:
                self.warnings.append(
                    f"{sym} ghi trong file {self.file_values[sym]:,.0f} ≠ tính lại {round_vnd(self.results[key]):,}")
        if self.file_total is not None and abs(D(self.file_total) - self.results["Tong"]) > 1:
            self.warnings.append(
                f"Tổng cộng ghi trong file {self.file_total:,.0f} ≠ tính lại {round_vnd(self.results['Tong']):,}")
        if self.file_total_rounded is not None:
            unit = _rounding_unit(self.file_total_rounded)
            if self.total_rounded(unit) != round(self.file_total_rounded):
                self.warnings.append(
                    f"Làm tròn ghi trong file {self.file_total_rounded:,.0f} ≠ tính lại "
                    f"{self.total_rounded(unit):,} (làm tròn đến {unit:,} đồng)")

    def summary_rows(self) -> List[Tuple[str, str, Any]]:
        """(ký hiệu, nội dung, giá trị đã làm tròn đồng) để in / xuất."""
        labels = [
            ("VL", "Chi phí vật liệu"), ("NC", "Chi phí nhân công"), ("M", "Chi phí máy thi công"),
            ("T", "Chi phí trực tiếp"), ("C", "Chi phí chung"), ("TL", "Thu nhập chịu thuế tính trước"),
            ("Gks", "Giá thành khảo sát xây dựng"), ("Glpa", "Lập phương án kỹ thuật khảo sát"),
            ("Glbc", "Lập báo cáo kết quả khảo sát"), ("Ghmc", "Chi phí hạng mục chung"),
            ("G", "Dự toán khảo sát trước thuế"), ("GTGT", "Thuế giá trị gia tăng"),
            ("Gxd", "Chi phí khảo sát sau thuế"), ("Gdp", "Chi phí dự phòng"), ("Tong", "Tổng cộng"),
        ]
        return [(k, name, round_vnd(self.results[k])) for k, name in labels]


def _rounding_unit(value: float) -> int:
    """Đơn vị làm tròn suy ra từ số đã làm tròn trong file (2 479 179 000 → 1 000)."""
    v, unit = int(round(value)), 1
    while unit < 1_000_000 and v % (unit * 10) == 0 and v != 0:
        unit *= 10
    return unit


# ─────────────────────────────────────────────────────────────────────────────
# ĐỌC FILE
# ─────────────────────────────────────────────────────────────────────────────

def load_survey_estimate(path: str, sheet: Optional[str] = None,
                         rate_overrides: Optional[Dict[str, float]] = None) -> SurveyEstimate:
    """Đọc dự toán khảo sát: bảng KL × đơn giá + tỷ lệ từ bảng tổng hợp / sheet Hệ số. Chưa tính."""
    qs = load_qs(path, sheet=sheet)
    est = SurveyEstimate(qs=qs, errors=list(qs.errors),
                         warnings=[w for w in qs.warnings if "Gks" not in w])
    sheets = list(_sheet_rows(path))
    # Nguồn chính: bảng tổng hợp có nhiều ký hiệu khoản mục nhất (≥ 3)
    summary = max((s for s in sheets if not _hs_sheet(s[1])),
                  key=lambda s: _count_symbols(s[1]), default=None)
    if summary is not None and _count_symbols(summary[1]) >= 3:
        _read_summary(est, f"{path}#{summary[0]}", summary[1])
    # Sheet Hệ số (nếu có): bổ sung tỷ lệ còn thiếu và đối chiếu với bảng tổng hợp
    for name, rows in sheets:
        if not _hs_sheet(rows):
            continue
        for key, rate in _read_hs_sheet(rows).items():
            if key not in est.rates:
                est.rates[key], est.rate_sources[key] = rate, f"{path}#{name}"
            elif abs(est.rates[key] - rate) > 1e-9:
                est.warnings.append(
                    f"{SURVEY_RATES[key][2]}: bảng tổng hợp {est.rates[key]:.2%} ≠ sheet {name} {rate:.2%}")
    for key, value in (rate_overrides or {}).items():
        if value is not None:
            est.rates[key] = value / 100
            est.rate_sources[key] = "tham số dòng lệnh"
    return est


def _sheet_rows(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        from tools.qs_loader import _XlsBook
        book = _XlsBook(path)
    else:
        from tools.excel_eval import WorkbookEvaluator
        from tools.qs_loader import _XlsxBook
        book = _XlsxBook(WorkbookEvaluator(path))
    for name in book.sheetnames:
        yield name, book.rows(name)


def _cells(row: Sequence[Any]) -> List[Any]:
    return [v for v in row if v is not None and not (isinstance(v, str) and not v.strip())]


def _hs_sheet(rows) -> bool:
    return sum(1 for row in rows for v in row if isinstance(v, str) and v.strip() in _HS_TO_KEY) >= 3


def _count_symbols(rows) -> int:
    return sum(1 for row in rows for v in row
               if isinstance(v, str) and (v.strip() in _SYMBOL_TO_KEY or v.strip() == "Gks"))


def _read_hs_sheet(rows) -> Dict[str, float]:
    """Sheet Hệ số: dòng 'tên | mã hsXXX | giá trị'."""
    rates: Dict[str, float] = {}
    for row in rows:
        for c, v in enumerate(row):
            if isinstance(v, str) and v.strip() in _HS_TO_KEY:
                nums = [_to_number(x) for x in row[c + 1:] if _to_number(x) is not None]
                if nums:
                    rates.setdefault(_HS_TO_KEY[v.strip()], nums[0])
                break
    return rates


def _read_summary(est: SurveyEstimate, source: str, rows) -> None:
    """Bảng tổng hợp: dòng có KÝ HIỆU (C, TL, Glpa...) và CÁCH TÍNH ghi 'NC x 65%', 'Gks x 1,5%'..."""
    for row in rows:
        cells = _cells(row)
        texts = [v.strip() for v in cells if isinstance(v, str)]
        nums = [float(v) for v in cells if isinstance(v, (int, float)) and not isinstance(v, bool)]
        norm = " ".join(texts).lower()
        symbol = next((t for t in reversed(texts) if t in _SYMBOL_TO_KEY or t in _FILE_VALUES), None)
        if symbol is None:
            if "làm tròn" in norm and nums:
                est.file_total_rounded = nums[-1]
            elif "tổng cộng" in norm and nums:
                est.file_total = nums[-1]
            continue
        if nums and symbol in _FILE_VALUES:
            est.file_values[symbol] = nums[-1]
        key = _SYMBOL_TO_KEY.get(symbol)
        if key is None:
            continue
        formula = next((t for t in texts if _PERCENT.search(t)), None)
        if formula is None:
            continue
        base = SURVEY_RATES[key][3]
        if not re.search(rf"(?<![A-Za-z]){base}(?![a-z])", formula):
            est.errors.append(
                f"{source}: {symbol} tính theo '{formula}' — cấu trúc dự toán khảo sát cần cơ sở {base}; "
                f"chưa hỗ trợ cách tính này")
            continue
        est.rates[key] = float(_PERCENT.search(formula).group(1).replace(",", ".")) / 100
        est.rate_sources[key] = source


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def print_report(est: SurveyEstimate) -> None:
    print(f"\n  DỰ TOÁN KHẢO SÁT XÂY DỰNG — {est.qs.source}")
    print(f"  {len(est.qs.items)} công tác\n")
    for sym, name, value in est.summary_rows():
        rate = ""
        key = _SYMBOL_TO_KEY.get(sym)
        if key:
            rate = f"  ({est.rates[key]:.2%} {SURVEY_RATES[key][3]})"
        print(f"   {sym:<5} {name:<36} {value:>18,}{rate}")
    if est.file_total_rounded is not None:
        unit = _rounding_unit(est.file_total_rounded)
        print(f"   {'':<5} {'Làm tròn (' + format(unit, ',') + ' đ)':<36} {est.total_rounded(unit):>18,}")
    for e in est.errors:
        print(f"  ❌ {e}")
    for w in est.warnings:
        print(f"  ⚠ {w}")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Tính lại dự toán khảo sát xây dựng từ file Excel")
    ap.add_argument("path")
    ap.add_argument("--sheet", default=None)
    a = ap.parse_args()
    e = load_survey_estimate(a.path, a.sheet)
    e.compute()
    print_report(e)
