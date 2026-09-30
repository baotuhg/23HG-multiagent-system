# -*- coding: utf-8 -*-
"""
UNIT TESTS — PROJECT EXPERIENCE & SYSTEM EVOLUTION ENGINE
Kiểm thử toàn diện 4 cấp độ tự nâng cấp (Level-Up / Continuous Learning) của hệ thống AEC:
  1. Field Productivity Calibration
  2. Golden Rebar Cutting Pattern Library
  3. Reflexion & Error Immunity Engine
  4. Autonomous Skill Packaging with Human-in-the-Loop Gate
  5. System XP, Rank & Level-up Progression
  6. Integration with EquipmentFleetScheduler & CuttingStockSolver
"""

import os
import tempfile
import unittest
import datetime

from aec_core.experience_store import (
    ProjectExperienceStore,
    ProductivityObservation,
    GoldenRebarPattern,
    ImmunityRule,
    CandidateSkill,
)
from agents.aec_experience_agent import AECExperienceAgent
from core.gates.human_gate import HumanGate, ApprovalRequest
from tools.equipment_fleet_scheduler import EquipmentFleetScheduler, FleetTask
from tools.cutting_stock_solver import CuttingStockSolver, CutDemand


class TestProjectExperienceStore(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.store_file = os.path.join(self.temp_dir.name, "test_experience.json")
        self.store = ProjectExperienceStore(store_path=self.store_file)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_default_rules_initialization(self):
        """Kiểm tra khởi tạo các quy tắc miễn dịch mặc định."""
        active_rules = self.store.get_active_immunity_rules()
        self.assertGreaterEqual(len(active_rules), 5)
        rule_ids = [r.rule_id for r in active_rules]
        self.assertIn("RULE-EXCEL-001", rule_ids)
        self.assertIn("RULE-MATH-002", rule_ids)
        self.assertIn("RULE-REBAR-003", rule_ids)

    def test_productivity_calibration(self):
        """Kiểm tra hiệu chuẩn năng suất thi công thực tế (Level 1)."""
        # Ghi nhận 3 quan trắc thực tế cho tác vụ đào đất
        # Định mức thiết kế: 400 m3/ca.
        # Thực tế 1: 360 m3/ca (ratio 0.90)
        # Thực tế 2: 340 m3/ca (ratio 0.85)
        # Thực tế 3: 380 m3/ca (ratio 0.95)
        self.store.record_productivity("dao_dat", "Đào đất móng", 400.0, 360.0, unit="m3")
        self.store.record_productivity("dao_dat", "Đào đất móng", 400.0, 340.0, unit="m3")
        self.store.record_productivity("dao_dat", "Đào đất móng", 400.0, 380.0, unit="m3")

        # Hệ số alpha có trọng số tuyến tính:
        # weights = [1, 2, 3] -> (0.90*1 + 0.85*2 + 0.95*3) / 6 = (0.90 + 1.70 + 2.85) / 6 = 5.45 / 6 ≈ 0.9083
        factor = self.store.get_calibration_factor("dao_dat")
        self.assertAlmostEqual(factor, 0.9083, places=3)

        # Năng suất hiệu chuẩn từ default 400
        calibrated_rate = self.store.get_calibrated_productivity("dao_dat", 400.0)
        self.assertAlmostEqual(calibrated_rate, round(400.0 * 0.9083, 2), places=1)

        # Tác vụ chưa có mẫu quan trắc -> alpha = 1.0
        self.assertEqual(self.store.get_calibration_factor("chua_co_mau"), 1.0)

    def test_golden_rebar_pattern_library(self):
        """Kiểm tra lưu trữ và tra cứu mẫu cắt thép vàng (Level 2)."""
        demands = {5850: 10, 2925: 20}  # Cắt vừa khít cây 11.7m (5850 + 2925*2 = 11700mm)
        patterns = [{"bars_used": 10, "cuts": [5850, 2925, 2925], "waste_mm": 0, "offcut_class": "PHE"}]

        # Lưu mẫu với hao hụt 0.1% (< 1.5% -> Thành công)
        saved = self.store.save_golden_pattern(
            element_type="coc_khoan_nhoi_d1000",
            diameter_mm=25,
            demands_dict=demands,
            stock_bar_count=10,
            waste_pct=0.1,
            cutting_patterns=patterns,
            project_origin="Cầu Thôn Khai Hoang 2"
        )
        self.assertIsNotNone(saved)
        self.assertEqual(saved.element_type, "coc_khoan_nhoi_d1000")

        # Thử lưu mẫu với hao hụt 3.5% (> 1.5% -> Bị từ chối làm mẫu vàng)
        rejected = self.store.save_golden_pattern(
            element_type="coc_khoan_nhoi_d1000",
            diameter_mm=25,
            demands_dict=demands,
            stock_bar_count=10,
            waste_pct=3.5,
            cutting_patterns=patterns,
        )
        self.assertIsNone(rejected)

        # Tra cứu mẫu vàng đã lưu
        looked_up = self.store.lookup_golden_pattern(
            element_type="coc_khoan_nhoi_d1000",
            diameter_mm=25,
            demands_dict=demands
        )
        self.assertIsNotNone(looked_up)
        self.assertEqual(looked_up.times_reused, 1)

    def test_reflexion_and_immunity_rules(self):
        """Kiểm tra cơ chế miễn dịch lỗi & kiểm toán chuỗi (Level 3)."""
        rule = self.store.register_immunity_rule(
            rule_id="RULE-CUSTOM-99",
            name="Kiểm tra mác vữa xây trát",
            category="LEGAL_NORMS",
            trigger_description="Vữa mác thấp hơn M75 trong môi trường xâm thực mặn",
            severity="CRITICAL",
            fix_recommendation="Nâng lên tối thiểu vữa xi măng M100"
        )
        self.assertEqual(rule.rule_id, "RULE-CUSTOM-99")

        # Kiểm tra ghi nhận vi phạm
        self.assertTrue(self.store.record_immunity_violation("RULE-CUSTOM-99"))
        self.assertEqual(self.store.immunity_rules["RULE-CUSTOM-99"].times_triggered, 1)

        # Kiểm tra kiểm toán text ô tính (chuỗi bắt đầu bằng '=' gây crash XML)
        cell_bad = "= Diễn giải móng M1 khối lượng đào"
        violations = self.store.audit_text_cell(cell_bad)
        self.assertTrue(len(violations) > 0)
        self.assertIn("RULE-EXCEL-001", violations[0])

        cell_good_formula = "=SUM(E6:E10)"
        self.assertEqual(len(self.store.audit_text_cell(cell_good_formula)), 0)

    def test_candidate_skill_and_human_gate(self):
        """Kiểm tra đề xuất và phê duyệt kỹ năng mới (Level 4)."""
        skill = self.store.propose_candidate_skill(
            skill_id="SKILL-BRIDGE-BEARING-01",
            name="Tự động kiểm tra độ lún gối cầu",
            category="QUALITY_CONTROL",
            description="So sánh cao độ gối cầu thực tế với bản vẽ thiết kế",
            workflow_trigger="Sau khi đổ bê tông xà mũ mố trụ",
            prompt_template="Đọc cao độ gối cầu và kiểm tra delta <= 2mm",
        )
        self.assertEqual(skill.status, "PENDING_APPROVAL")

        # Phê duyệt bằng HumanGate (auto mode)
        human_gate = HumanGate(mode="auto")
        agent = AECExperienceAgent(store_path=self.store_file)
        decision = agent.submit_candidate_skill_for_approval(
            skill_id="SKILL-BRIDGE-BEARING-01",
            name="Tự động kiểm tra độ lún gối cầu",
            category="QUALITY_CONTROL",
            description="So sánh cao độ gối cầu thực tế với bản vẽ thiết kế",
            workflow_trigger="Sau khi đổ bê tông xà mũ mố trụ",
            human_gate=human_gate
        )
        self.assertTrue(decision.approved)

        # Kiểm tra trạng thái đã chuyển thành APPROVED
        approved_skills = agent.store.get_approved_skills()
        self.assertEqual(len(approved_skills), 1)
        self.assertEqual(approved_skills[0].skill_id, "SKILL-BRIDGE-BEARING-01")

        # Xuất SKILL.md
        md = agent.store.export_skill_markdown("SKILL-BRIDGE-BEARING-01")
        self.assertIn("# SKILL: Tự động kiểm tra độ lún gối cầu", md)

    def test_system_level_progression(self):
        """Kiểm tra tính điểm XP và cấp bậc Level Up của AI."""
        # Baseline XP với 5 default immunity rules (5 * 10 = 50 XP)
        lvl_init = self.store.get_system_level()
        self.assertEqual(lvl_init["level"], 2)  # 1 + int(sqrt(50/40)) = 1 + 1 = 2
        self.assertIn("Kỹ sư Tập sự", lvl_init["rank"])

        # Hoàn thành 2 dự án (+200 XP), thêm 10 quan trắc (+100 XP), thêm 2 mẫu vàng (+50 XP)
        self.store.record_project_completion("Dự án Cầu Khai Hoang 2")
        self.store.record_project_completion("Dự án Cụm B9 Vina Alpha")
        for i in range(10):
            self.store.record_productivity(f"task_{i}", f"Tác vụ {i}", 100.0, 100.0)
        self.store.save_golden_pattern(
            "dam_i", 20, {4000: 2}, 1, 0.5, []
        )
        self.store.save_golden_pattern(
            "dam_t", 25, {5000: 2}, 1, 0.8, []
        )

        lvl_new = self.store.get_system_level()
        self.assertGreater(lvl_new["total_xp"], lvl_init["total_xp"])
        self.assertGreaterEqual(lvl_new["level"], 3)

    def test_persistence_atomic_reload(self):
        """Kiểm tra tính toàn vẹn dữ liệu khi nạp lại từ đĩa."""
        self.store.record_project_completion("Dự án Tuyến Tránh QL1A")
        self.store.save()

        # Tạo một store instance mới trỏ vào cùng file
        store2 = ProjectExperienceStore(store_path=self.store_file)
        self.assertIn("Dự án Tuyến Tránh QL1A", store2.projects_history)

    def test_equipment_fleet_integration(self):
        """Kiểm tra tích hợp tự động hiệu chuẩn vào EquipmentFleetScheduler."""
        # Ghi nhận thực tế máy đào M1 chỉ đạt 320 m3/ca thay vì 401.07 m3/ca theo định mức (alpha ≈ 0.80)
        self.store.record_productivity(
            task_code="dao_mong_cong",
            task_name="Đào móng cống",
            planned_productivity=401.07,
            actual_productivity=320.85,
        )

        scheduler = EquipmentFleetScheduler(experience_store=self.store)
        task = FleetTask(
            task_id="1",
            code="ĐM-16",
            name="Đào đất hố móng cống hộp",
            unit="m3",
            quantity=73800.0,
            start_date=datetime.date(2026, 9, 20),
            end_date=datetime.date(2026, 10, 5),
            shifts_per_day=2,
            task_norm_key="dao_mong_cong"
        )
        task = scheduler.compute_task_allocations(task)

        # Vì thực tế chậm hơn định mức, số ca yêu cầu phải tăng lên so với gốc (184.0 ca -> ~230 ca)
        m1 = task.machine_allocations[0]
        self.assertGreater(m1.total_shifts_required, 185.0)

    def test_cutting_stock_integration(self):
        """Kiểm tra tích hợp tự động lưu mẫu vàng vào CuttingStockSolver."""
        demands = [
            CutDemand(length_mm=5800, quantity=20, diameter_mm=25, mark="C1"),
        ]
        solver = CuttingStockSolver()
        sol = solver.solve(
            demands,
            experience_store=self.store,
            element_type="coc_d1000_khoan_nhoi"
        )
        self.assertIn(sol.status, ("OPTIMAL", "FEASIBLE"))
        self.assertLessEqual(sol.waste_ratio_pct, 1.5)

        # Xác nhận mẫu vàng đã được tự động lưu vào store!
        patterns = self.store.list_golden_patterns()
        self.assertEqual(len(patterns), 1)
        self.assertEqual(patterns[0].element_type, "coc_d1000_khoan_nhoi")
        self.assertEqual(patterns[0].diameter_mm, 25)

    def test_experience_agent_report_generation(self):
        """Kiểm tra tác tử sinh báo cáo tiến hóa (Evolution Report)."""
        agent = AECExperienceAgent(store_path=self.store_file)
        report = agent.generate_evolution_report()
        self.assertIn("BÁO CÁO TIẾN HÓA & CẤP ĐỘ HỆ THỐNG AEC MULTI-AGENT", report)
        self.assertIn("CẤP ĐỘ HIỆN TẠI (LEVEL)", report)


if __name__ == "__main__":
    unittest.main()
