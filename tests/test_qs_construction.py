# -*- coding: utf-8 -*-
"""
Golden test dự toán xây lắp — đối chiếu với một dự toán THẬT đã lập (TT 13/2021/TT-BXD).

Fixture tests/fixtures/du_toan_xay_lap_trich.xlsx: trích 2 sheet (bảng chi tiết khối lượng công
tác xây dựng và bảng tổng hợp dự toán hạng mục) của một file dự toán công trình xây lắp, giữ
nguyên toàn bộ số liệu, bỏ tên dự án / địa danh (hạng mục đổi thành "PHẦN A" / "PHẦN B"). Sheet
"Công trình" chứa 2 hạng mục trong cùng một bảng — kiểm tra việc tách hạng mục.

Các con số kỳ vọng lấy từ file gốc (bảng "TH dự toán hạng mục"), không tính từ chính chương trình.

Chạy thêm với file gốc đầy đủ:  AEC_TEST_XAYLAP_XLSX=/đường/dẫn/du_toan.xlsx
"""

import os
import unittest

from tools.qs_loader import load_all_hang_muc, load_qs

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURE = os.path.join(ROOT, "tests", "fixtures", "du_toan_xay_lap_trich.xlsx")

RATES = {"chung": 0.048, "nha_tam": 0.018, "kxd": 0.02, "tl": 0.055, "vat": 0.1}

# Số ghi trong file gốc (bảng TH dự toán hạng mục), theo thứ tự 2 hạng mục trong sheet.
APPROVED = [
    {"items": 217, "VL": 34_261_038_929, "NC": 27_865_442_796, "M": 10_087_867_807,
     "T": 72_214_349_532, "C": 3_466_288_778, "LT": 1_299_858_292, "TT": 1_444_286_991,
     "TL": 4_313_363_098, "G": 82_738_146_691, "G_XD": 91_011_961_360, "round": 91_011_961_000},
    {"items": 208, "VL": 18_735_702_283, "NC": 13_366_694_001, "M": 3_792_808_473,
     "T": 35_895_204_757, "C": 1_722_969_828, "LT": 646_113_686, "TT": 717_904_095,
     "TL": 2_144_020_580, "G": 41_126_212_946, "G_XD": 45_238_834_241, "round": 45_238_834_000},
]


def _check(test, est, exp):
    est.compute()
    test.assertEqual(len(est.items), exp["items"])
    test.assertEqual(est.errors, [])
    test.assertEqual(set(est.rates), set(RATES))
    for k, v in RATES.items():
        test.assertAlmostEqual(est.rates[k], v, places=9, msg=k)
    test.assertEqual((est.VL, est.NC, est.M, est.T), (exp["VL"], exp["NC"], exp["M"], exp["T"]))
    test.assertEqual(est.GT_components["chung"], exp["C"])
    test.assertEqual(est.GT_components["nha_tam"], exp["LT"])
    test.assertEqual(est.GT_components["kxd"], exp["TT"])
    test.assertEqual(est.TL, exp["TL"])
    test.assertEqual(est.G, exp["G"])
    test.assertEqual(est.G_XD, exp["G_XD"])
    test.assertEqual(round(est.G_XD / 1000) * 1000, exp["round"])


class ConstructionGoldenTest(unittest.TestCase):
    def test_both_hang_muc_match_approved(self):
        ests = load_all_hang_muc(FIXTURE)
        self.assertEqual([e.hang_muc for e in ests], ["PHẦN A", "PHẦN B"])
        for est, exp in zip(ests, APPROVED):
            _check(self, est, exp)

    def test_select_one_hang_muc_by_name(self):
        est = load_qs(FIXTURE, hang_muc="PHẦN B")
        self.assertEqual(est.hang_muc, "PHẦN B")
        _check(self, est, APPROVED[1])

    def test_default_reads_first_block_with_warning(self):
        est = load_qs(FIXTURE)
        self.assertEqual(est.hang_muc, "PHẦN A")
        self.assertTrue(any("2 hạng mục" in w for w in est.warnings))

    def test_component_prices_captured(self):
        est = load_qs(FIXTURE, hang_muc="PHẦN A")
        first = est.items[0]           # AB.65120 Đắp đất, 55,781 (100m3) — chỉ có NC và Máy
        self.assertEqual(first.code, "AB.65120")
        self.assertAlmostEqual(first.quantity, 55.781, places=3)
        # Đơn giá tách VL(+VLP) / NC / M lấy từ nhóm cột "Tính trực tiếp"
        self.assertEqual((first.price_vl, first.price_nc, first.price_m), (0, 1376656, 1462546))

    def test_no_charge_line_is_zero_not_error(self):
        est = load_qs(FIXTURE, hang_muc="PHẦN A")
        zero = [i for i in est.items if i.unit_price == 0 and "đá hộc dư thừa" in i.description]
        self.assertTrue(zero)
        self.assertEqual(zero[0].amount, 0)


@unittest.skipUnless(os.environ.get("AEC_TEST_XAYLAP_XLSX"), "đặt AEC_TEST_XAYLAP_XLSX để chạy với file gốc")
class FullConstructionFileTest(unittest.TestCase):
    def test_full_file_matches(self):
        ests = load_all_hang_muc(os.environ["AEC_TEST_XAYLAP_XLSX"])
        self.assertEqual(len(ests), 2)
        for est, exp in zip(ests, APPROVED):
            _check(self, est, exp)


if __name__ == "__main__":
    unittest.main()
