# -*- coding: utf-8 -*-
"""
Golden test quy ước tính thép — đối chiếu với một bảng tính thép QS chuyên nghiệp thực chiến.

Các con số kỳ vọng dưới đây là GIÁ TRỊ ĐÃ TÍNH SẴN trong file gốc (sheet "Chi tiết" cho cốt thép,
sheet "ket cau thep" cho thép tấm), trích ra kèm đường kính / kích thước — không có thông tin dự án.
Chúng kiểm chứng rằng tools/steel_qs tái tạo đúng quy ước của hồ sơ thật, không phải tự nhất quán.
"""

import unittest

from tools import steel_qs


# (D mm, KL thiết kế kg, KL 1 cây 11,7m kg, số cây, dây buộc kg) — từ sheet "Chi tiết"
REBAR_ROWS = [
    (6, 427.6523890821286, 0.0, 0, 6.415),
    (8, 3399.3236669930257, 0.0, 0, 50.990),
    (10, 7902.553867392121, 7.222222222222221, 1094, 118.538),
    (12, 13983.305965884154, 10.4, 1345, 209.750),
    (16, 3543.6550082752924, 18.488888888888887, 192, 53.155),
    (18, 3099.2998744055876, 23.4, 132, 46.489),
    (20, 2036.1769051991507, 28.888888888888886, 70, 30.543),
    (22, 5675.438489650041, 34.955555555555556, 162, 85.132),
    (25, 4537.0998138922505, 45.138888888888886, 101, 68.056),
    (28, 2188.4853545352335, 56.62222222222222, 39, 32.827),
    (30, 0.0, 65.0, 0, 0.0),
]

# (ký hiệu, dày mm, rộng mm, dài mm, số lượng, KL kg) — từ sheet "ket cau thep"
PLATE_ROWS = [
    ("PL12", 12, 200, 10238, 1, 192.88392),
    ("PL6", 6, 276, 10238, 1, 133.0899048),
    ("PL6", 6, 97, 276, 22, 27.7411464),
    ("PL8", 8, 97, 276, 2, 3.3625632),
    ("PL20", 20, 250, 350, 1, 13.7375),
    ("PL10", 10, 450, 550, 1, 19.42875),
]


class RebarConventionGoldenTest(unittest.TestCase):
    def test_unit_mass_d10_is_d2_over_162(self):
        self.assertAlmostEqual(steel_qs.rebar_unit_mass(10), 100 / 162.0, places=9)

    def test_whole_bar_weight_matches_file(self):
        for d, _kg, g, _h, _i in REBAR_ROWS:
            self.assertAlmostEqual(steel_qs.rebar_whole_bar_weight(d), g, places=6, msg=f"D{d}")

    def test_coil_diameters_not_counted_as_bars(self):
        # D ≤ 8 cấp dạng cuộn → 0 cây nguyên dù khối lượng lớn
        self.assertEqual(steel_qs.rebar_whole_bar_weight(8), 0.0)
        self.assertEqual(steel_qs.rebar_bar_count(3399.32, 8), 0)

    def test_bar_count_matches_file(self):
        for d, kg, _g, h, _i in REBAR_ROWS:
            self.assertEqual(steel_qs.rebar_bar_count(kg, d), h, msg=f"D{d}")

    def test_bar_count_uses_round_half_up(self):
        # D12: 13983.306 / 10.4 = 1344.549 → 1345 (nửa lên, không phải cắt xuống)
        self.assertEqual(steel_qs.rebar_bar_count(13983.305965884154, 12), 1345)

    def test_tie_wire_is_1_5_percent(self):
        for d, kg, _g, _h, i in REBAR_ROWS:
            self.assertAlmostEqual(steel_qs.tie_wire_kg(kg), i, places=3, msg=f"D{d}")

    def test_procurement_bundle_with_waste(self):
        r = steel_qs.rebar_procurement(10000.0, 20, waste_ratio=0.05)
        self.assertAlmostEqual(r["purchase_kg"], 10500.0, places=4)
        self.assertEqual(r["bar_count"], steel_qs.rebar_bar_count(10500.0, 20))
        self.assertAlmostEqual(r["tie_wire_kg"], round(10500.0 * 0.015, 3), places=3)


class SteelPlateGoldenTest(unittest.TestCase):
    def test_plate_weight_matches_file(self):
        for mark, t, w, L, n, kg in PLATE_ROWS:
            got = steel_qs.steel_member_weight(mark, t, w, L, n)
            self.assertAlmostEqual(got, round(kg, 4), places=3, msg=mark)

    def test_plate_vs_profile_dispatch(self):
        # "PL.." → công thức tấm; ký hiệu khác → thép hình theo kg/m
        plate = steel_qs.steel_member_weight("PL10", 10, 200, 1000, 1)
        self.assertAlmostEqual(plate, 10 * 200 * 1000 * 7.85 / 1_000_000, places=4)
        profile = steel_qs.steel_member_weight("I300", 42.2, 0, 6000, 2)  # 42,2 kg/m, dài 6m, 2 cây
        self.assertAlmostEqual(profile, 42.2 * 6.0 * 2, places=4)

    def test_parse_plate_thickness(self):
        self.assertEqual(steel_qs.parse_plate_thickness("PL12"), 12.0)
        self.assertIsNone(steel_qs.parse_plate_thickness("H350"))


if __name__ == "__main__":
    unittest.main()
