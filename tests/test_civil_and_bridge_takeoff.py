# -*- coding: utf-8 -*-
"""
UNIT TESTS — CIVIL & BRIDGE TAKEOFF ENGINE
Kiểm thử toàn diện động cơ bóc tách hình học dân dụng & suy luận diễn biến sang cầu đường:
  1. Hình học móng chóp cụt & bệ mố vát
  2. Bệ trụ xẻ nước mũi thuyền (Cutwater)
  3. Thân trụ tròn đôi & Xà mũ cánh hẫng (Pier Cap)
  4. Trừ giao dầm, cột và bản sàn
  5. Thép tấm PL & diện tích sơn chống gỉ
  6. Bóc tách Mố cầu toàn diện (Abutment Engine)
  7. Bóc tách Trụ cầu toàn diện (Pier Engine)
  8. Bóc tách Kết cấu nhịp Dầm Super-T & Cáp DƯL (Superstructure Engine)
  9. Bóc tách Dầm cầu thép chữ I liên hợp (Steel Girder Engine)
"""

import math
import unittest

from tools.civil_and_bridge_takeoff_engine import (
    calc_frustum_pyramid,
    calc_cutwater_pier_footing,
    calc_column_and_corbel,
    calc_beam_with_slab_deductions,
    calc_structural_steel_plate,
    BridgeAbutmentParams,
    BridgeAbutmentEngine,
    BridgePierParams,
    BridgePierEngine,
    BridgeSuperstructureParams,
    BridgeSuperstructureEngine,
    SteelBridgeGirderSegment,
    SteelBridgeGirderEngine,
)


class TestCivilAndBridgeTakeoffEngine(unittest.TestCase):

    def test_calc_frustum_pyramid(self):
        """Kiểm tra tính thể tích móng chóp cụt và bệ mố vát dốc."""
        # Đáy: 4m x 4m, cao thẳng: 1m
        # Đỉnh vát: 2m x 2m, chiều cao vát dốc: 1m
        res = calc_frustum_pyramid(L_base=4.0, W_base=4.0, H_base=1.0, L_top=2.0, W_top=2.0, h_slope=1.0)
        # V_base = 4 * 4 * 1 = 16 m3
        self.assertEqual(res["concrete_base_m3"], 16.0)
        # V_frustum = (16 + 4 + sqrt(64))/3 * 1 = 28/3 ≈ 9.3333 m3
        self.assertAlmostEqual(res["concrete_frustum_m3"], 9.3333, places=3)
        self.assertAlmostEqual(res["total_concrete_m3"], 25.3333, places=3)
        self.assertGreater(res["total_formwork_m2"], 16.0)

    def test_calc_cutwater_pier_footing(self):
        """Kiểm tra tính bệ trụ xẻ nước mũi thuyền bo tròn."""
        # Hình chữ nhật giữa: 6m x 3m x 1.5m
        # Bán kính mũi tròn: R = 1.5m
        res = calc_cutwater_pier_footing(L_rect=6.0, W_rect=3.0, H=1.5)
        # V_rect = 6 * 3 * 1.5 = 27 m3
        # V_cylinder = pi * 1.5^2 * 1.5 ≈ 10.6029 m3
        expected_conc = round(27.0 + math.pi * 2.25 * 1.5, 4)
        self.assertEqual(res["concrete_m3"], expected_conc)
        self.assertGreater(res["formwork_m2"], 0)

    def test_calc_column_and_corbel(self):
        """Kiểm tra cột tròn có trừ giao dầm/sàn và có vai đỡ."""
        res = calc_column_and_corbel(
            L=0, W=0, H=6.0, dia=1.2, h_beam_deduct=0.8,
            corbel_w=0.4, corbel_l=0.5, corbel_h=0.4, corbel_count=2
        )
        # effective_h = 5.2m
        expected_v = (math.pi * (0.6**2)) * 5.2
        self.assertAlmostEqual(res["concrete_col_m3"], round(expected_v, 4), places=3)
        self.assertGreater(res["concrete_corbel_m3"], 0)

    def test_calc_beam_with_slab_deductions(self):
        """Kiểm tra dầm tầng / dầm ngang trừ giao cột và sàn."""
        res = calc_beam_with_slab_deductions(
            length_m=10.0, width_m=0.3, height_m=0.7,
            col_width_deduct_m=0.8, slab_thick_deduct_m=0.15
        )
        # net_l = 9.2m, net_h = 0.55m
        self.assertEqual(res["net_length_m"], 9.2)
        self.assertEqual(res["net_height_m"], 0.55)
        self.assertAlmostEqual(res["concrete_m3"], 9.2 * 0.3 * 0.55, places=3)

    def test_calc_structural_steel_plate(self):
        """Kiểm tra tính khối lượng thép tấm PL và diện tích sơn từ File 2."""
        # Tấm PL20: dày 20mm, rộng 300mm, dài 2000mm, số lượng 2
        res = calc_structural_steel_plate(thickness_mm=20.0, width_mm=300.0, length_mm=2000.0, count=2)
        # Vol = 20 * 300 * 2000 * 2 = 24,000,000 mm3 = 24 dm3
        # Weight = 24 * 7.85 = 188.4 kg
        self.assertEqual(res["weight_kg"], 188.4)
        # Sơn 2 mặt: 2 * 300 * 2000 * 2 / 10^6 = 2.4 m2
        self.assertEqual(res["paint_area_m2"], 2.4)

    def test_bridge_abutment_engine(self):
        """Kiểm thử toàn diện bóc tách Mố cầu M1."""
        params = BridgeAbutmentParams(name="Mố Cầu M1")
        takeoff = BridgeAbutmentEngine.compute_takeoff(params)

        self.assertEqual(takeoff["component"], "Mố Cầu M1")
        self.assertGreater(takeoff["total_structural_concrete_m3"], 50.0)
        self.assertGreater(takeoff["total_formwork_m2"], 100.0)
        self.assertGreater(takeoff["rebar_kg"], 5000.0)
        # Kiểm tra dây thép buộc 1.5%
        self.assertAlmostEqual(takeoff["rebar_tie_wire_kg"], round(takeoff["rebar_kg"] * 0.015, 1), places=1)

    def test_bridge_pier_engine(self):
        """Kiểm thử toàn diện bóc tách Trụ cầu T1 xẻ nước & xà mũ hẫng."""
        params = BridgePierParams(name="Trụ T1 Thân Cột Đôi")
        takeoff = BridgePierEngine.compute_takeoff(params)

        self.assertEqual(takeoff["component"], "Trụ T1 Thân Cột Đôi")
        self.assertGreater(takeoff["footing_concrete_m3"], 30.0)
        self.assertGreater(takeoff["columns_concrete_m3"], 25.0)
        self.assertGreater(takeoff["cap_concrete_m3"], 20.0)
        self.assertGreater(takeoff["total_rebar_kg"], 8000.0)

    def test_bridge_superstructure_engine(self):
        """Kiểm thử bóc tách Kết cấu nhịp 4 dầm Super-T 33m & Cáp DƯL."""
        params = BridgeSuperstructureParams(
            bridge_length_m=33.0, deck_width_m=9.0, girder_count=4
        )
        takeoff = BridgeSuperstructureEngine.compute_takeoff(params)

        # 4 dầm x 23.5m3 = 94 m3
        self.assertEqual(takeoff["girder_concrete_m3"], 94.0)
        # Cáp DƯL: 4 dầm x 1480kg = 5920 kg
        self.assertEqual(takeoff["total_pt_strands_kg"], 5920.0)
        self.assertEqual(takeoff["anchorage_sets"], 40)
        self.assertGreater(takeoff["asphalt_concrete_ton"], 30.0)

    def test_steel_bridge_girder_engine(self):
        """Kiểm thử bóc tách dầm thép chữ I liên hợp từ Sheet ket cau thep."""
        seg = SteelBridgeGirderSegment(name="Dầm thép liên hợp L=30m")
        takeoff = SteelBridgeGirderEngine.compute_takeoff(seg)

        self.assertGreater(takeoff["top_flange_kg"], 2000.0)
        self.assertGreater(takeoff["bottom_flange_kg"], 3500.0)
        self.assertGreater(takeoff["web_plate_kg"], 4500.0)
        self.assertGreater(takeoff["shear_studs_kg"], 80.0)
        self.assertGreater(takeoff["total_structural_steel_ton"], 11.0)
    def test_bridge_rebar_bbs_generation(self):
        """Kiểm thử sinh CutDemands cốt thép cầu nối sang CuttingStockSolver."""
        from tools.civil_and_bridge_takeoff_engine import (
            generate_bridge_pier_rebar_demands,
            generate_super_t_rebar_demands
        )
        from tools.cutting_stock_solver import CuttingStockSolver, CutDemand

        pier_demands = generate_bridge_pier_rebar_demands(column_height_m=8.5, column_dia_m=1.5, column_count=2)
        self.assertEqual(len(pier_demands), 2)
        self.assertEqual(pier_demands[0]["mark"], "TC-TRU-D28")
        self.assertEqual(pier_demands[0]["quantity"], 64) # 32 x 2

        super_t_demands = generate_super_t_rebar_demands(span_length_m=33.0, girder_count=4)
        self.assertGreaterEqual(len(super_t_demands), 3)

        # Chuyển đổi thành CutDemand và đưa vào CuttingStockSolver
        cut_demands = [CutDemand(**d) for d in super_t_demands if d["diameter_mm"] == 20]
        solver = CuttingStockSolver()
        sol = solver.solve(cut_demands)
        self.assertIn(sol.status, ("OPTIMAL", "FEASIBLE"))
        self.assertLessEqual(sol.waste_ratio_pct, 1.5)


if __name__ == "__main__":
    unittest.main()
