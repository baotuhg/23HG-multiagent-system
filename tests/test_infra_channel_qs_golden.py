# -*- coding: utf-8 -*-
"""
Golden test đo bóc mương hộp (dòng đáy) — đối chiếu bảng tính khối lượng QS chuyên nghiệp.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN ở dòng đáy mương B600 trong sheet 'Mương' của hồ sơ thật,
chỉ lấy kích thước và kết quả.
"""

import unittest

from tools.infra_channel_qs import (ChannelBaseParams, channel_base_takeoff,
                                    slope_factor)


class ChannelBaseGoldenTest(unittest.TestCase):
    def setUp(self):
        self.r = channel_base_takeoff(ChannelBaseParams(
            length_m=48.5, width_m=0.6, height_m=1.3, lean_t_m=0.05,
            base_slab_t_m=0.15, wall_t_m=0.145, cover_t_m=0.15, aggregate_base_t_m=0.0,
            count=1, pile_length_m=2.5, ground_level_m=-0.7, water_avg_level_m=-1.35,
            base_widen_m=0.5, rebar_d_mm=10, rebar_spacing_mm=200, rebar_layers=2, rebar_factor=1.12))

    def test_concrete_and_lean(self):
        self.assertAlmostEqual(self.r["lean_concrete_m3"], 2.40075, places=5)
        self.assertAlmostEqual(self.r["base_concrete_m3"], 6.47475, places=5)
        self.assertAlmostEqual(self.r["lean_formwork_m2"], 4.85, places=4)
        self.assertAlmostEqual(self.r["formwork_m2"], 14.55, places=4)
        self.assertAlmostEqual(self.r["nilon_m2"], 48.015, places=4)
        self.assertAlmostEqual(self.r["waterproof_in_m2"], 29.1, places=4)

    def test_piles(self):
        self.assertEqual(self.r["pile_count"], 960)
        self.assertAlmostEqual(self.r["pile_md"], 2400, places=3)

    def test_earthwork(self):
        self.assertAlmostEqual(self.r["dig_height_m"], 0.85, places=6)
        self.assertAlmostEqual(self.r["top_widen_m"], 0.7125, places=6)
        self.assertAlmostEqual(self.r["excavation_m3"], 86.67556250, places=5)
        self.assertAlmostEqual(self.r["backfill_m3"], 49.74281250, places=5)

    def test_rebar(self):
        self.assertAlmostEqual(self.r["rebar_unit_kg_m2"], 12.345679012345679, places=6)
        self.assertAlmostEqual(self.r["rebar_kg"], 596.8493827160494, places=3)

    def test_slope_factor_bands(self):
        self.assertEqual(slope_factor(0.85), 0.25)
        self.assertEqual(slope_factor(3.0), 0.67)
        self.assertEqual(slope_factor(5.5), 0.85)


if __name__ == "__main__":
    unittest.main()
