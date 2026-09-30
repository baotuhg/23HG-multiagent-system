# -*- coding: utf-8 -*-
"""
UNIT TESTS — OPENBIM IFC LOADER & TAKEOFF ENGINE
Kiểm thử nạp mô hình OpenBIM IFC, bóc tách cấu kiện bê tông và chuyển đổi cốt thép sang Solver OR-Tools.
"""

import os
import unittest
import tempfile

try:
    import ifcopenshell
    HAS_IFCOPENSHELL = True
except ImportError:
    HAS_IFCOPENSHELL = False

from tools.ifc_loader import IFCLoader, IFCTakeoffResult, nominal_weight_kg_per_m
from tools.cutting_stock_solver import CuttingStockSolver


class TestIFCLoader(unittest.TestCase):

    def setUp(self):
        self.loader = IFCLoader()
        self.temp_dir = tempfile.mkdtemp()
        self.test_ifc_path = os.path.join(self.temp_dir, "test_bridge_pier.ifc")
        self._generate_test_ifc_file(self.test_ifc_path)

    def tearDown(self):
        if os.path.exists(self.test_ifc_path):
            os.remove(self.test_ifc_path)
        if os.path.exists(self.temp_dir):
            os.rmdir(self.temp_dir)

    def _generate_test_ifc_file(self, path: str):
        """Sinh tệp IFC4 mẫu đại diện mố trụ cầu đường cho kiểm thử."""
        if not HAS_IFCOPENSHELL:
            # Tạo file STEP cơ bản nếu thiếu ifcopenshell
            with open(path, "w", encoding="utf-8") as f:
                f.write(
                    "ISO-10303-21;\nHEADER;\nFILE_SCHEMA(('IFC4'));\nENDSEC;\nDATA;\n"
                    "#1=IFCPROJECT('111',$,'DuAnCauKhaiHoang2',$,$,$,$,$,$);\n"
                    "#2=IFCBEAM('222',$,'Dam_SuperT_Nhip1',$,$,$,$,$);\n"
                    "#3=IFCCOLUMN('333',$,'Tru_T1',$,$,$,$,$);\n"
                    "#4=IFCREINFORCINGBAR('444',$,'ThepChu_D25',$,$,$,$,$,'RB-01','CB400-V',25.0,490.87,11700.0);\n"
                    "ENDSEC;\nEND-ISO-10303-21;\n"
                )
            return

        f = ifcopenshell.file(schema="IFC4")
        f.create_entity("IfcProject", ifcopenshell.guid.new(), None, "DuAnCauKhaiHoang2")

        # Cấu kiện Bê tông: 1 Dầm, 1 Cột, 1 Móng, 1 Bản mặt cầu
        f.create_entity("IfcBeam", ifcopenshell.guid.new(), None, "Dam_SuperT_Nhip1")
        f.create_entity("IfcColumn", ifcopenshell.guid.new(), None, "Tru_T1")
        f.create_entity("IfcFooting", ifcopenshell.guid.new(), None, "Mong_Tru_T1")
        f.create_entity("IfcSlab", ifcopenshell.guid.new(), None, "Ban_Mat_Cau")

        # Cốt thép 3D: D25 và D12
        for i in range(5):
            f.create_entity(
                "IfcReinforcingBar",
                ifcopenshell.guid.new(),
                None,
                f"ThepChu_D25_{i+1}",
                "Thanh thép chủ D25",
                None, None, None,
                f"RB-25-{i+1}",
                "CB400-V",
                25.0,
                490.87,
                6500.0
            )
        for i in range(5):
            f.create_entity(
                "IfcReinforcingBar",
                ifcopenshell.guid.new(),
                None,
                f"ThepChu_Ngan_D25_{i+1}",
                "Thanh thép chủ ngắn D25",
                None, None, None,
                f"RB-25N-{i+1}",
                "CB400-V",
                25.0,
                490.87,
                5100.0
            )
        for i in range(10):
            f.create_entity(
                "IfcReinforcingBar",
                ifcopenshell.guid.new(),
                None,
                f"ThepDai_D12_{i+1}",
                "Thanh thép đai D12",
                None, None, None,
                f"RB-12-{i+1}",
                "CB300-V",
                12.0,
                113.1,
                2450.0
            )

        f.write(path)

    def test_nominal_weight_kg_per_m(self):
        """Kiểm tra khối lượng riêng danh định theo TCVN 1651:2018."""
        self.assertAlmostEqual(nominal_weight_kg_per_m(10.0), 0.617, places=2)
        self.assertAlmostEqual(nominal_weight_kg_per_m(12.0), 0.888, places=2)
        self.assertAlmostEqual(nominal_weight_kg_per_m(20.0), 2.466, places=2)
        self.assertAlmostEqual(nominal_weight_kg_per_m(25.0), 3.853, places=2)

    def test_load_ifc_concrete_elements(self):
        """Kiểm tra trích xuất cấu kiện bê tông từ mô hình IFC."""
        res = self.loader.load_ifc(self.test_ifc_path)
        self.assertIsInstance(res, IFCTakeoffResult)
        self.assertGreaterEqual(len(res.concrete_elements), 2)

        types = [e.ifc_type for e in res.concrete_elements]
        self.assertIn("IfcBeam", types)
        self.assertIn("IfcColumn", types)

    def test_load_ifc_rebar_elements(self):
        """Kiểm tra trích xuất cốt thép 3D từ mô hình IFC."""
        res = self.loader.load_ifc(self.test_ifc_path)
        self.assertGreater(len(res.rebar_elements), 0)

        # Tổng hợp theo đường kính
        rebar_by_dia = res.summary_rebar_by_diameter()
        self.assertIn(25.0, rebar_by_dia)

    def test_ifc_to_cutting_stock_solver(self):
        """Kiểm tra đường truyền khép kín: Trích xuất thép IFC → Nạp vào Solver Cắt thép 1D."""
        res = self.loader.load_ifc(self.test_ifc_path)
        demands = res.to_cutting_stock_demands()
        self.assertGreater(len(demands), 0)

        # Chỉ lọc nhóm D25 để giải bài toán tối ưu
        demands_d25 = [d for d in demands if d.diameter_mm == 25.0]
        if demands_d25:
            solver = CuttingStockSolver()
            solution = solver.solve(demands_d25)
            self.assertIn(solution.status, ["OPTIMAL", "FEASIBLE"])
            self.assertGreater(solution.total_bars_needed, 0)
            # Với 5 thanh 6.5m và 5 thanh 5.1m ghép cặp vào cây 11.7m (11.6m), tỷ lệ hao hụt cực thấp < 2%
            self.assertLess(solution.waste_ratio_pct, 2.0)


if __name__ == "__main__":
    unittest.main()
