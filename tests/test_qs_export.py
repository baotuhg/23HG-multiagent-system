# -*- coding: utf-8 -*-
"""Test xuất bảng G_XD (tools/qs_export.py) từ bảng QS mẫu trong templates/.

Trước đây module này có lỗi cú pháp (import đặt lệch vào giữa hàm) mà không test nào phát hiện,
nên `--qs-out` luôn lỗi. Test này đảm bảo module import được và xuất ra file Excel đọc lại được.
"""

import os
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QS_TEMPLATE = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")


class TestQSExport(unittest.TestCase):

    def test_module_imports(self):
        from tools import qs_export  # noqa: F401  (lỗi cú pháp sẽ làm test này đỏ)

    @unittest.skipUnless(os.path.exists(QS_TEMPLATE), "thiếu bảng QS mẫu trong templates/")
    def test_write_gxd_workbook_from_template(self):
        import openpyxl
        from tools.qs_export import write_gxd_workbook
        from tools.qs_loader import apply_rate_overrides, load_qs

        est = load_qs(QS_TEMPLATE)
        self.assertGreater(len(est.items), 0)
        apply_rate_overrides(est, {"chung": 5.1, "nha_tam": 1.2, "kxd": 1.0, "tl": 5.5, "vat": 8.0})
        est.compute()
        self.assertGreater(est.G_XD, 0)

        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "gxd.xlsx")
            write_gxd_workbook(out, est)
            self.assertTrue(os.path.exists(out))
            wb = openpyxl.load_workbook(out)
            self.assertIn("TONG_HOP_GXD", wb.sheetnames)


if __name__ == "__main__":
    unittest.main()
