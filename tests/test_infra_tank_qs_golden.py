# -*- coding: utf-8 -*-
"""
Golden test đo bóc đáy bể nước ngầm — đối chiếu bảng tính khối lượng QS chuyên nghiệp.

Giá trị kỳ vọng là các ô ĐÃ TÍNH SẴN ở dòng đáy bể trong hai sheet 'Bể nước PCCC' và 'Bể XLNT'
của hồ sơ thật, chỉ lấy kích thước và kết quả.
"""

import unittest

from tools.infra_tank_qs import TankBaseParams, tank_base_takeoff


class TankBaseGoldenTest(unittest.TestCase):
    def test_pccc_tank_base(self):
        r = tank_base_takeoff(TankBaseParams(
            length_m=20.85, width_m=12, tank_height_m=3.05, slab_t_m=0.35, lean_t_m=0.05,
            lean_widen_m=0.1, wall_type="IW1", bottom_widen_m=1.2, top_widen_m=2.0,
            level_offset_m=-0.3, base_offset_m=-0.1, rebar_mesh=[(14, 150)] * 4, rebar_factor=1.12))
        self.assertAlmostEqual(r["lean_concrete_m3"], 12.8405, places=4)
        self.assertAlmostEqual(r["concrete_m3"], 87.57, places=4)
        self.assertAlmostEqual(r["formwork_m2"], 22.995, places=4)
        self.assertAlmostEqual(r["waterproof_inner_m2"], 250.2, places=4)
        self.assertAlmostEqual(r["dig_height_m"], 3.25, places=6)
        self.assertAlmostEqual(r["excavation_m3"], 1190.15, places=4)
        self.assertAlmostEqual(r["backfill_m3"], 376.6695, places=4)
        self.assertAlmostEqual(r["rebar_unit_kg_m2"], 32.263374485596714, places=6)
        self.assertAlmostEqual(r["rebar_kg"], 9040.971851851855, places=2)

    def test_xlnt_tank_base(self):
        r = tank_base_takeoff(TankBaseParams(
            length_m=30, width_m=25, tank_height_m=4.15, slab_t_m=0.5, lean_t_m=0.05,
            lean_widen_m=0.1, wall_type="IW1", bottom_widen_m=1.2, top_widen_m=2.0,
            level_offset_m=-0.3, base_offset_m=-0.1, rebar_mesh=[(18, 150)] * 4, rebar_factor=1.12))
        self.assertAlmostEqual(r["lean_concrete_m3"], 38.052, places=3)
        self.assertAlmostEqual(r["concrete_m3"], 375, places=3)
        self.assertAlmostEqual(r["formwork_m2"], 55, places=3)
        self.assertAlmostEqual(r["dig_height_m"], 4.5, places=6)
        self.assertAlmostEqual(r["excavation_m3"], 4215.96, places=2)
        self.assertAlmostEqual(r["backfill_m3"], 840.408, places=2)
        self.assertAlmostEqual(r["rebar_unit_kg_m2"], 53.333333333333336, places=6)
        self.assertAlmostEqual(r["rebar_kg"], 44800.0, places=1)

    def test_waterproof_switches_by_wall_type(self):
        inner = tank_base_takeoff(TankBaseParams(10, 10, 3, 0.3, wall_type="IW1"))
        outer = tank_base_takeoff(TankBaseParams(10, 10, 3, 0.3, wall_type="EW1"))
        self.assertEqual(inner["waterproof_inner_m2"], 100.0)
        self.assertEqual(inner["waterproof_outer_m2"], 0.0)
        self.assertEqual(outer["waterproof_outer_m2"], 100.0)
        self.assertEqual(outer["waterproof_inner_m2"], 0.0)


if __name__ == "__main__":
    unittest.main()
