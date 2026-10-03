# -*- coding: utf-8 -*-
"""
Golden test đo bóc hạ tầng đường — đối chiếu bảng tính khối lượng QS chuyên nghiệp.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN ở hai đoạn đường asphalt tải nhẹ (R1) trong sheet
'Hạ tầng- Đường' của hồ sơ thật, chỉ lấy kích thước và kết quả.
"""

import unittest

from tools.infra_road_qs import RoadSegmentParams, road_segment_takeoff


class RoadGoldenTest(unittest.TestCase):
    def test_r1_segment_1(self):
        r = road_segment_takeoff(RoadSegmentParams(
            road_type="R1", area_m2=1714.50875, base_depth_m=0.4,
            ground_level_m=-0.7, road_avg_level_m=-0.36, strip_extra_m=0.3,
            sand_fill_depth_m=0.3, sand_compaction=1.35))
        self.assertAlmostEqual(r["base_m3"], 685.8035, places=4)
        self.assertAlmostEqual(r["asphalt_light_m2"], 1714.50875, places=4)
        self.assertAlmostEqual(r["emulsion_light_m2"], 1714.50875, places=4)
        self.assertAlmostEqual(r["strip_height_m"], 0.36, places=6)
        self.assertAlmostEqual(r["strip_subgrade_m3"], 617.22315, places=4)
        self.assertAlmostEqual(r["sand_fill_m3"], 694.37604375, places=4)
        self.assertAlmostEqual(r["compaction_m2"], 1714.50875, places=4)

    def test_r1_segment_2(self):
        r = road_segment_takeoff(RoadSegmentParams(
            road_type="R1", area_m2=2047.57708, base_depth_m=0.4,
            ground_level_m=-0.7, road_avg_level_m=-0.36))
        self.assertAlmostEqual(r["base_m3"], 819.030832, places=4)
        self.assertAlmostEqual(r["strip_subgrade_m3"], 737.1277488, places=4)
        self.assertAlmostEqual(r["sand_fill_m3"], 829.2687174, places=4)

    def test_concrete_road_r3_branch(self):
        # R3: bê tông mặt + nilon + cốt thép lưới (nhánh công thức, kiểm nội bộ)
        r = road_segment_takeoff(RoadSegmentParams(
            road_type="R3", area_m2=1000.0, base_depth_m=0.3, concrete_depth_m=0.25,
            waste_concrete=1.0, rebar_d_mm=12, rebar_spacing_mm=200, rebar_layers=1,
            rebar_factor=1.15))
        self.assertAlmostEqual(r["concrete_m3"], 1000.0 * 0.25, places=4)
        self.assertEqual(r["asphalt_light_m2"], 0.0)
        self.assertAlmostEqual(r["nilon_m2"], 1000.0, places=4)
        self.assertAlmostEqual(r["rebar_unit_kg_m2"], 12 ** 2 / 162 * (1000 / 200) * 2, places=6)
        self.assertGreater(r["rebar_kg"], 0.0)


if __name__ == "__main__":
    unittest.main()
