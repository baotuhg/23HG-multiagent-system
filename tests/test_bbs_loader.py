# -*- coding: utf-8 -*-
"""Kiểm thử bộ đọc BBS thật — chạy: python -m unittest discover tests"""

import json
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tools.bbs_loader import BBSLoadError, load_bbs

TEMPLATE = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")


class BBSLoaderTest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, name, text):
        path = os.path.join(self.tmp.name, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path

    def test_csv_vietnamese_headers_in_metres(self):
        path = self.write("bbs.csv", (
            "Hạng mục kết cấu;Ký hiệu thanh;Đường kính Ø (mm);Mác thép;Chiều dài 1 thanh (m);"
            "Số thanh / cấu kiện;Số cấu kiện\n"
            "Mố M1;F1;28;CB400-V;6,578;83;2\n"
            "02. PHẦN TRÊN;;;;;;\n"
            "Mố M1;F5;16;CB400-V;5,85;16;1\n"
        ))
        result = load_bbs(path)
        self.assertEqual(result.errors, [])
        self.assertEqual([(d.mark, d.diameter_mm, d.length_mm, d.quantity) for d in result.demands],
                         [("Mố M1:F1", 28, 6578, 166), ("Mố M1:F5", 16, 5850, 16)])

    def test_csv_english_headers_in_mm(self):
        path = self.write("bbs.csv", "mark,diameter_mm,grade,length_mm,quantity\nT1,20,CB400-V,4500,30\n")
        result = load_bbs(path)
        self.assertEqual(result.demands[0].length_mm, 4500)
        self.assertEqual(result.demands[0].quantity, 30)

    def test_number_formats(self):
        from tools.bbs_loader import _to_number
        self.assertEqual(_to_number("6,578"), 6.578)
        self.assertEqual(_to_number("1.234,5"), 1234.5)
        self.assertEqual(_to_number("1,234.5"), 1234.5)
        self.assertEqual(_to_number("abc"), None)

    def test_length_without_unit_is_rejected(self):
        path = self.write("bbs.csv", "mark,diameter,length,quantity\nT1,20,4500,30\n")
        with self.assertRaises(BBSLoadError):
            load_bbs(path)

    def test_bad_rows_are_reported_not_hidden(self):
        path = self.write("bbs.csv", (
            "mark,diameter_mm,grade,length_mm,quantity\n"
            "OK,16,CB400-V,3000,10\n"
            "SHIFT,128,CB400-V,3000,10\n"
            "TINY,16,CB400-V,2,10\n"
            "HALF,16,CB400-V,3000,2.5\n"
            "CABLE-15.2,15.2,ASTM A416 Gr270,38200,660\n"
        ))
        result = load_bbs(path)
        self.assertEqual([d.mark for d in result.demands], ["OK"])
        self.assertEqual(len(result.errors), 3)
        self.assertEqual(len(result.skipped), 1)

    def test_json(self):
        path = self.write("bbs.json", json.dumps({"items": [
            {"mark": "A", "diameter_mm": 12, "length_mm": 1200, "quantity": 5},
        ]}))
        self.assertEqual(load_bbs(path).demands[0].quantity, 5)

    def test_excel_uncalculated_formula_uses_per_member_count(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ký hiệu thanh", "Đường kính Ø (mm)", "Chiều dài 1 thanh (m)",
                   "Số thanh / cấu kiện", "Số cấu kiện", "Tổng số thanh"])
        ws.append(["F1", 16, 2.5, 10, 3, "=D2*E2"])
        path = os.path.join(self.tmp.name, "bbs.xlsx")
        wb.save(path)
        result = load_bbs(path)
        self.assertEqual(result.errors, [])
        self.assertEqual(result.demands[0].quantity, 30)

    def test_repository_template_bbs(self):
        result = load_bbs(TEMPLATE)
        self.assertTrue(result.source.endswith("#THONG_KE_THEP_CHI_TIET"))
        self.assertEqual(result.rows_read, 390)
        self.assertEqual(len(result.demands), 379)
        self.assertEqual(len(result.errors), 10)   # dòng lệch cột: Ø128, Ø260, dài 2mm...
        self.assertEqual(len(result.skipped), 1)   # cáp DƯL 15.2


if __name__ == "__main__":
    unittest.main()
