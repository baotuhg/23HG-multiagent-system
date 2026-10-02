# -*- coding: utf-8 -*-
"""Bộ tính công thức: ngày tháng, lỗi Excel lan truyền, mảng, IF/AND/VLOOKUP/SUMPRODUCT/TEXT.
Kết quả kỳ vọng là kết quả của Excel (tính tay), mốc: 01/01/2024 = số seri 45292."""

import datetime
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import openpyxl

from tools.audit_excels_static import audit_file
from tools.excel_eval import ExcelErrorResult, FormulaError, WorkbookEvaluator, serial_to_date
from tools.lab_qaqc import _as_date
from tools.schedule_loader import _as_date_str

CELLS = {
    # ngày tháng
    "A1": datetime.datetime(2024, 1, 1), "A2": "=A1+1", "A3": "=MONTH(A1+31)", "A4": '=DATEVALUE("15/03/2027")',
    "A5": "=YEAR(A1)", "A6": '=A1>=DATEVALUE("01/01/2024")',
    # lỗi lan truyền & IF chỉ dùng nhánh được chọn
    "B1": 0, "C1": 7, "B2": "=IF(B1=0, 0, (C1/B1)*2)", "B3": '=IFERROR(1/B1,"x")', "B4": "=1/B1", "B5": "=ISERROR(B4)",
    "B6": "=B4+1",
    # mảng: SUMPRODUCT kiểu biểu đồ Gantt / phụ tải
    "D1": 1, "D2": 5, "D3": 10, "E1": 4, "E2": 8, "E3": 12, "F1": 10, "F2": 20, "F3": 30, "G1": 5,
    "G2": "=SUMPRODUCT((G1>=D1:D3)*(G1<=E1:E3)*F1:F3)", "G3": "=SUM(F1:F3*2)",
    # VLOOKUP
    "H1": "C30", "I1": 1, "H2": "C35", "I2": 2, "H3": 45, "I3": 3,
    "J1": '=VLOOKUP("c35",H1:I3,2,FALSE)', "J2": '=VLOOKUP("C40",H1:I3,2,FALSE)',
    "J3": "=IFERROR(VLOOKUP(45,H1:I3,2,FALSE),0)", "J4": '=VLOOKUP("45",H1:I3,2,FALSE)',
    # logic
    "K1": "=AND(TRUE, 1>0, NOT(FALSE))", "K2": '=OR(FALSE, "a">5)', "K3": '="x"&TRUE',
    "K4": '=IF(AND(G1>=D2, G1<=E2), IF(B1=0, "█", F2), "")',
    # TEXT
    "L1": '=TEXT(1234.5,"#,##0")', "L2": '=TEXT(5,"00")', "L3": '=TEXT(0.256,"0.0%")', "L4": '=TEXT(A1,"dd/mm/yyyy")',
    "L5": '=TEXT("15/03/2027","dd/mm/yyyy")', "L6": '=TEXT("abc","0")', "L7": '=TEXT(1234.567,"#,##0.00")',
}


class ExcelEvalExtendedTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.path = os.path.join(cls.tmp.name, "f.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "S"
        for k, v in CELLS.items():
            ws[k] = v
        wb.save(cls.path)
        cls.ev = WorkbookEvaluator(cls.path)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def v(self, coord):
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
        col, row = coordinate_from_string(coord)
        return self.ev.value("S", row, column_index_from_string(col))

    def test_dates_are_excel_serials(self):
        self.assertEqual(self.v("A2"), 45293)                 # 01/01/2024 = 45292
        self.assertEqual(self.v("A3"), 2)                     # 01/02/2024
        self.assertEqual(self.v("A4"), 46461)                 # 15/03/2027 = 45292 + 366 + 365 + 365 + 73
        self.assertEqual(self.v("A5"), 2024)
        self.assertIs(self.v("A6"), True)

    def test_errors_propagate_and_if_uses_only_chosen_branch(self):
        self.assertEqual(self.v("B2"), 0)                     # nhánh (C1/B1) có #DIV/0! nhưng không được chọn
        self.assertEqual(self.v("B3"), "x")
        with self.assertRaises(ExcelErrorResult) as cm:
            self.v("B4")
        self.assertEqual(cm.exception.code, "#DIV/0!")
        self.assertIsInstance(cm.exception, FormulaError)    # tương thích với code cũ bắt FormulaError
        self.assertIs(self.v("B5"), True)
        with self.assertRaises(ExcelErrorResult):
            self.v("B6")                                      # lỗi lan sang ô tham chiếu

    def test_array_arithmetic_is_elementwise(self):
        # ngày 5 nằm trong [5,8] của dòng 2 → 20 ; không chỉ lấy phần tử đầu
        self.assertEqual(self.v("G2"), 20)
        self.assertEqual(self.v("G3"), 120)                   # (10 + 20 + 30) × 2

    def test_vlookup_exact(self):
        self.assertEqual(self.v("J1"), 2)                     # không phân biệt hoa thường
        with self.assertRaises(ExcelErrorResult) as cm:
            self.v("J2")
        self.assertEqual(cm.exception.code, "#N/A")
        self.assertEqual(self.v("J3"), 3)
        with self.assertRaises(ExcelErrorResult):
            self.v("J4")                                      # chữ "45" khác số 45 như Excel

    def test_logic(self):
        self.assertIs(self.v("K1"), True)
        self.assertIs(self.v("K2"), True)                     # Excel: chữ > số
        self.assertEqual(self.v("K3"), "xTRUE")
        self.assertEqual(self.v("K4"), "█")

    def test_text_formats(self):
        self.assertEqual(self.v("L1"), "1,235")               # làm tròn half-up như Excel
        self.assertEqual(self.v("L2"), "05")
        self.assertEqual(self.v("L3"), "25.6%")
        self.assertEqual(self.v("L4"), "01/01/2024")
        self.assertEqual(self.v("L5"), "15/03/2027")          # chữ dạng ngày được đổi như Excel
        self.assertEqual(self.v("L6"), "abc")                 # không đổi được → trả nguyên chữ
        self.assertEqual(self.v("L7"), "1,234.57")

    def test_audit_counts_real_errors_separately_from_unsupported(self):
        r = audit_file(self.path)
        self.assertEqual(r["eval_unsupported"], 0)
        self.assertEqual(r["eval_error_results"], 4)          # B4, B6, J2, J4 tính ra lỗi Excel thật


class SerialDatesInLoadersTest(unittest.TestCase):

    def test_serial_to_date(self):
        self.assertEqual(serial_to_date(45292), datetime.date(2024, 1, 1))
        self.assertEqual(serial_to_date(45292.75), datetime.date(2024, 1, 1))
        self.assertIsNone(serial_to_date(5.0))                # số nhỏ không phải ngày
        self.assertIsNone(serial_to_date(True))

    def test_loaders_read_formula_dates(self):
        self.assertEqual(_as_date_str(45292.0), "2024-01-01")
        self.assertEqual(_as_date(45292), datetime.date(2024, 1, 1))
        self.assertEqual(_as_date_str(12.0), "")


if __name__ == "__main__":
    unittest.main()
