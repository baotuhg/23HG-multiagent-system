# -*- coding: utf-8 -*-
"""
Golden test đo bóc hàng rào — đối chiếu bảng tính khối lượng QS chuyên nghiệp.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN trong sheet 'Hàng rào' của hồ sơ thật: móng trụ M1/M2 và
tường xây 220, chỉ lấy kích thước và kết quả.
"""

import unittest

from tools.infra_fence_qs import (FencePostFootingParams, FenceWallParams,
                                  fence_post_footing_takeoff, fence_wall_takeoff)


class FencePostFootingGoldenTest(unittest.TestCase):
    def _params(self, count):
        return FencePostFootingParams(
            L_m=0.6, W_m=0.6, footing_box_h_m=0.15, footing_slope_h_m=0.1,
            top_L_m=0.32, top_W_m=0.27, post_w_m=0.22, post_l_m=0.22,
            foundation_to_top_m=1.5, count=count, ground_level_m=-0.2,
            base_widen_m=0.5, top_widen_m=1.0, footing_bottom_level_m=-1.0)

    def test_M1_37_posts(self):
        r = fence_post_footing_takeoff(self._params(37))
        self.assertAlmostEqual(r["lean_concrete_m3"], 0.9065, places=4)
        self.assertAlmostEqual(r["footing_concrete_m3"], 2.82384, places=5)
        self.assertAlmostEqual(r["footing_formwork_m2"], 13.32, places=4)
        self.assertAlmostEqual(r["post_ge200_md"], 46.25, places=4)
        self.assertAlmostEqual(r["post_concrete_m3"], 2.2385, places=4)
        self.assertAlmostEqual(r["post_formwork_m2"], 40.7, places=4)
        self.assertAlmostEqual(r["dig_height_m"], 0.85, places=6)
        self.assertAlmostEqual(r["excavation_m3"], 146.557, places=3)
        self.assertAlmostEqual(r["backfill_m3"], 141.84172, places=4)

    def test_M2_4_posts(self):
        r = fence_post_footing_takeoff(self._params(4))
        self.assertAlmostEqual(r["lean_concrete_m3"], 0.098, places=4)
        self.assertAlmostEqual(r["footing_concrete_m3"], 0.30528, places=5)
        self.assertAlmostEqual(r["post_ge200_md"], 5.0, places=4)
        self.assertAlmostEqual(r["post_concrete_m3"], 0.242, places=4)
        self.assertAlmostEqual(r["excavation_m3"], 15.844, places=3)
        self.assertAlmostEqual(r["backfill_m3"], 15.33424, places=4)

    def test_post_thickness_buckets(self):
        thin = fence_post_footing_takeoff(self._params(1).__class__(
            L_m=0.6, W_m=0.6, footing_box_h_m=0.15, footing_slope_h_m=0.1, top_L_m=0.32, top_W_m=0.27,
            post_w_m=0.15, post_l_m=0.15, foundation_to_top_m=1.5, count=1, footing_bottom_level_m=-1.0))
        self.assertGreater(thin["post_lt200_md"], 0)
        self.assertEqual(thin["post_ge200_md"], 0)
        big = fence_post_footing_takeoff(self._params(1).__class__(
            L_m=0.6, W_m=0.6, footing_box_h_m=0.15, footing_slope_h_m=0.1, top_L_m=0.32, top_W_m=0.27,
            post_w_m=0.3, post_l_m=0.3, foundation_to_top_m=1.5, count=1, footing_bottom_level_m=-1.0))
        self.assertGreater(big["post_ge300_md"], 0)
        self.assertEqual(big["post_ge200_md"], 0)


class FenceWallGoldenTest(unittest.TestCase):
    def test_wall_220(self):
        w = fence_wall_takeoff(FenceWallParams(length_m=117.893, thickness_m=0.22,
                                               height_m=0.55, post_deduction_m=9.02, count=1))
        self.assertAlmostEqual(w["masonry_m3"], 13.173633, places=5)
        self.assertAlmostEqual(w["plaster_m2"], 155.86076, places=4)
        self.assertAlmostEqual(w["paint_m2"], 155.86076, places=4)
        self.assertTrue(w["is_220"])


if __name__ == "__main__":
    unittest.main()
