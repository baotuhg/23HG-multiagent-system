# -*- coding: utf-8 -*-
"""
Golden test đo bóc cống tròn hạ tầng — đối chiếu với bảng tính khối lượng QS chuyên nghiệp.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN trong sheet 'Hạ tầng- Cống tròn' của hồ sơ thật (4 tuyến
cống liên tiếp), chỉ lấy kích thước/loại cống và kết quả — không có thông tin dự án.
"""

import unittest

from tools.infra_culvert_qs import (CulvertParams, culvert_takeoff,
                                    slope_factor)

# (D cống m, Dài m, loại, trừ giao m, cốt hiện trạng, cốt TB đào) + kết quả hồ sơ
CASES = [
    (dict(diameter_m=0.3, length_m=17.30, culvert_type="CL", junction_trim_m=0.40,
          ground_level_m=-0.5, avg_dig_level_m=-2.275),
     dict(bucket="CL-300", eff=16.7, pile=123, md=246, dig=1.775,
          exc=68.04364875, back=66.10940875, haul=1.93424)),
    (dict(diameter_m=0.3, length_m=17.30, culvert_type="TH", junction_trim_m=0.40,
          ground_level_m=-0.5, avg_dig_level_m=-2.325),
     dict(bucket="TH-300", eff=16.7, pile=123, md=246, dig=1.825,
          exc=70.90188875, back=68.96764875, haul=1.93424)),
    (dict(diameter_m=0.4, length_m=14.85, culvert_type="CL", junction_trim_m=0.40,
          ground_level_m=-0.5, avg_dig_level_m=-2.375),
     dict(bucket="CL-400", eff=14.25, pile=120, md=240, dig=1.875,
          exc=64.4970703125, back=61.9556328125, haul=2.5414375)),
    (dict(diameter_m=0.3, length_m=11.017, culvert_type="CL", junction_trim_m=0.35,
          ground_level_m=-0.5, avg_dig_level_m=-2.05),
     dict(bucket="CL-300", eff=10.467, pile=77, md=154, dig=1.55,
          exc=33.227395725, back=32.076020525, haul=1.1513752)),
]


class CulvertGoldenTest(unittest.TestCase):
    def test_all_rows_match_file(self):
        for params, exp in CASES:
            r = culvert_takeoff(CulvertParams(**params))
            tag = f"{params['culvert_type']} D{params['diameter_m']} L{params['length_m']}"
            self.assertEqual(r["bucket"], exp["bucket"], tag)
            self.assertAlmostEqual(r["effective_length_m"], exp["eff"], places=5, msg=tag)
            self.assertEqual(r["pile_count"], exp["pile"], tag)
            self.assertAlmostEqual(r["pile_md"], exp["md"], places=4, msg=tag)
            self.assertAlmostEqual(r["dig_height_m"], exp["dig"], places=5, msg=tag)
            self.assertAlmostEqual(r["excavation_m3"], exp["exc"], places=4, msg=tag)
            self.assertAlmostEqual(r["backfill_m3"], exp["back"], places=4, msg=tag)
            self.assertAlmostEqual(r["haul_m3"], exp["haul"], places=4, msg=tag)

    def test_haul_equals_pipe_displacement(self):
        # Vận chuyển = Đào − Đắp = thể tích thân cống chiếm chỗ
        r = culvert_takeoff(CulvertParams(**CASES[0][0]))
        self.assertAlmostEqual(r["haul_m3"], r["excavation_m3"] - r["backfill_m3"], places=6)

    def test_pvc_has_no_bamboo_pile(self):
        r = culvert_takeoff(CulvertParams(diameter_m=0.2, length_m=10.0, culvert_type="PVC",
                                          junction_trim_m=0.3, ground_level_m=-0.5, avg_dig_level_m=-2.0))
        self.assertEqual(r["pile_count"], 0)
        self.assertEqual(r["bucket"], "PVC-200")

    def test_slope_factor_bands(self):
        self.assertEqual(slope_factor(1.0), 0.25)    # ≤ 1,5 m
        self.assertEqual(slope_factor(1.775), 0.67)  # (1,5; 5) m
        self.assertEqual(slope_factor(6.0), 0.85)    # ≥ 5 m


if __name__ == "__main__":
    unittest.main()
