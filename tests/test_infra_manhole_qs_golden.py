# -*- coding: utf-8 -*-
"""
Golden test đo bóc hố ga hạ tầng (kiểu 1) — đối chiếu với bảng tính khối lượng QS chuyên nghiệp.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN trong sheet 'Hố ga (Kiểu 1)' của hồ sơ thật: một hố ga bê tông
và ba hố ga xây gạch (nắp song chắn / nắp bê tông), chỉ lấy kích thước/cốt và kết quả.
"""

import unittest

from tools.infra_manhole_qs import (ManholeType1Params, manhole_takeoff,
                                    slope_factor_mm)


class ManholeGoldenTest(unittest.TestCase):
    def test_concrete_manhole_row8(self):
        r = manhole_takeoff(ManholeType1Params(
            culvert_in_mm=300, bottom_level_mm=-900, top_level_mm=-180, manhole_type="Bê tông",
            ground_level_mm=-500, cross_D_mm=300, cross_count=2, cover_type="Grating"))
        self.assertAlmostEqual(r["concrete_m3"], 0.40780088815692245, places=6)
        self.assertAlmostEqual(r["formwork_m2"], 4.734336293856408, places=5)
        self.assertAlmostEqual(r["lean_concrete_m3"], 0.05, places=6)
        self.assertAlmostEqual(r["lean_formwork_m2"], 0.2, places=6)
        self.assertEqual(r["pile_count"], 30)
        self.assertAlmostEqual(r["pile_md"], 60, places=4)
        self.assertAlmostEqual(r["grating_neck_md"], 2.8, places=4)
        self.assertAlmostEqual(r["excavation_m3"], 1.45561, places=5)
        self.assertAlmostEqual(r["backfill_m3"], 0.96011, places=5)
        self.assertAlmostEqual(r["haul_m3"], 0.4955, places=5)

    def test_brick_manhole_row9(self):
        r = manhole_takeoff(ManholeType1Params(
            culvert_in_mm=300, bottom_level_mm=-960, top_level_mm=-180, manhole_type="Xây gạch",
            ground_level_mm=-500, cross_D_mm=300, cross_count=2, cover_type="Grating"))
        self.assertAlmostEqual(r["masonry_m3"], 0.3632359692968196, places=6)
        self.assertAlmostEqual(r["plaster_m2"], 4.614145175425633, places=5)
        self.assertAlmostEqual(r["bottom_neck_concrete_m3"], 0.22608, places=5)
        self.assertAlmostEqual(r["bottom_neck_formwork_m2"], 1.728, places=4)
        self.assertEqual(r["pile_count"], 38)
        self.assertAlmostEqual(r["excavation_m3"], 1.681364225, places=5)
        self.assertAlmostEqual(r["backfill_m3"], 1.010688225, places=5)
        self.assertAlmostEqual(r["haul_m3"], 0.670676, places=5)

    def test_brick_manhole_concrete_cover_row10(self):
        r = manhole_takeoff(ManholeType1Params(
            culvert_in_mm=300, bottom_level_mm=-980, top_level_mm=-180, manhole_type="Xây gạch",
            ground_level_mm=-500, cross_D_mm=300, cross_count=2, cover_type="Bê tông"))
        self.assertAlmostEqual(r["cover_concrete_m3"], 0.0196, places=5)
        self.assertAlmostEqual(r["cover_neck_md"], 2.8, places=4)
        self.assertAlmostEqual(r["excavation_m3"], 1.751159475, places=5)
        self.assertAlmostEqual(r["backfill_m3"], 1.058851475, places=5)

    def test_brick_manhole_row11(self):
        r = manhole_takeoff(ManholeType1Params(
            culvert_in_mm=300, bottom_level_mm=-1030, top_level_mm=-180, manhole_type="Xây gạch",
            ground_level_mm=-500, cross_D_mm=300, cross_count=2, cover_type="Bê tông"))
        self.assertAlmostEqual(r["masonry_m3"], 0.4137479692968196, places=6)
        self.assertAlmostEqual(r["plaster_m2"], 5.073345175425633, places=5)
        self.assertAlmostEqual(r["excavation_m3"], 1.9300576, places=5)
        self.assertAlmostEqual(r["backfill_m3"], 1.1836696, places=5)
        self.assertAlmostEqual(r["haul_m3"], 0.746388, places=5)

    def test_haul_equals_excavation_minus_backfill(self):
        r = manhole_takeoff(ManholeType1Params(
            culvert_in_mm=300, bottom_level_mm=-900, top_level_mm=-180, manhole_type="Bê tông",
            ground_level_mm=-500, cross_D_mm=300, cross_count=2))
        self.assertAlmostEqual(r["haul_m3"], r["excavation_m3"] - r["backfill_m3"], places=6)

    def test_slope_factor_bands_mm(self):
        self.assertEqual(slope_factor_mm(600), 0.25)
        self.assertEqual(slope_factor_mm(3000), 0.67)
        self.assertEqual(slope_factor_mm(6000), 0.85)


if __name__ == "__main__":
    unittest.main()
