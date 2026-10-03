# -*- coding: utf-8 -*-
"""
Golden test đo bóc đài móng — đối chiếu với bảng tính khối lượng QS chuyên nghiệp thực chiến.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN trong sheet 'Móng' của hồ sơ thật (đài F1 vuông/chữ nhật và
đài SF2 quả trám kiểu 1), chỉ lấy kích thước hình học và kết quả — không có thông tin dự án.
"""

import unittest

from tools.civil_foundation_qs import (WALL_FORMWORK_FACES, footing_box,
                                       footing_diamond_type1, footing_frustum)


class FoundationGoldenTest(unittest.TestCase):
    def test_box_footing_F1_matches_file(self):
        # F1: L=3.6, W=3.332, H=0.9, lót dày 0.05, mở rộng 0.1
        r = footing_box(3.6, 3.332, 0.9, lean_t=0.05, lean_extend=0.1).as_dict()
        self.assertAlmostEqual(r["concrete_m3"], 10.79568, places=5)
        self.assertAlmostEqual(r["formwork_m2"], 12.4776, places=4)
        self.assertAlmostEqual(r["lean_concrete_m3"], 0.63492, places=5)
        self.assertAlmostEqual(r["lean_formwork_m2"], 0.7132, places=4)

    def test_diamond_type1_SF2_matches_file(self):
        # SF2: L=W=1.9, H=0.9, a=0.85, b=0.95, c=1.15, d=0.9, lót 0.05, mở rộng 0.1
        r = footing_diamond_type1(1.9, 1.9, 0.9, a=0.85, b=0.95, c=1.15, d=0.9,
                                  lean_t=0.05, lean_extend=0.1).as_dict()
        self.assertAlmostEqual(r["concrete_m3"], 2.714625, places=6)
        self.assertAlmostEqual(r["formwork_m2"], 6.165, places=3)
        self.assertAlmostEqual(r["lean_concrete_m3"], 0.16875, places=5)
        self.assertAlmostEqual(r["lean_formwork_m2"], 0.3625, places=4)

    def test_count_and_waste_scale_concrete(self):
        one = footing_box(2.0, 2.0, 1.0).concrete_m3
        self.assertAlmostEqual(footing_box(2.0, 2.0, 1.0, count=3).concrete_m3, one * 3, places=6)
        self.assertAlmostEqual(footing_box(2.0, 2.0, 1.0, waste=1.05).concrete_m3, one * 1.05, places=6)


class FrustumConventionTest(unittest.TestCase):
    """Chóp cụt theo quy ước QS (trung bình diện tích) — hồ sơ không có dòng mẫu nên kiểm công thức."""

    def test_frustum_uses_average_area_approximation(self):
        # Đáy 4x4, cao hộp 1; đỉnh 2x2, cao vát 1
        # V = 4*4*1 + ((16)+(4))/2 * 1 = 16 + 10 = 26 m3  (xấp xỉ QS)
        r = footing_frustum(4.0, 4.0, 1.0, L_top=2.0, W_top=2.0, h_slope=1.0)
        self.assertAlmostEqual(r.concrete_m3, 26.0, places=6)

    def test_frustum_differs_from_exact_cone_engine(self):
        from tools.civil_and_bridge_takeoff_engine import calc_frustum_pyramid
        exact = calc_frustum_pyramid(4.0, 4.0, 1.0, 2.0, 2.0, 1.0)["total_concrete_m3"]
        qs = footing_frustum(4.0, 4.0, 1.0, L_top=2.0, W_top=2.0, h_slope=1.0).concrete_m3
        self.assertAlmostEqual(exact, 25.3333, places=3)   # nón cụt chính xác
        self.assertAlmostEqual(qs, 26.0, places=3)         # xấp xỉ QS
        self.assertGreater(qs, exact)


class FormworkFaceConventionTest(unittest.TestCase):
    def test_wall_face_convention_recorded(self):
        self.assertEqual(WALL_FORMWORK_FACES["chu_L"], 2)
        self.assertEqual(WALL_FORMWORK_FACES["chu_I"], 2)
        self.assertEqual(WALL_FORMWORK_FACES["loi_thang_may"], "bù trừ")


if __name__ == "__main__":
    unittest.main()
