# -*- coding: utf-8 -*-
"""
TÁC TỬ: AEC EXPERIENCE & SYSTEM EVOLUTION AGENT
Hệ thống Multi-Agent AEC (23HG-multiagent-system).

Nhiệm vụ:
1. Quản trị vòng đời tích lũy kinh nghiệm và tự tiến hóa (Continuous Learning Engine).
2. Tự động thu nạp số liệu hoàn công (As-Built) để tính hệ số hiệu chuẩn năng suất thi công thực tế.
3. Tự động lưu trữ các nghiệm thức cắt thép tối ưu (< 1.5% đề-xê) thành Mẫu cắt thép vàng (Golden Rebar Patterns).
4. Khắc ghi các lỗi kiểm toán & sự cố định dạng vào Bộ quy tắc Miễn dịch (Error Immunity Rules).
5. Đề xuất và đóng gói kỹ năng mới (Candidate Skills), trình duyệt qua Human-in-the-Loop Gate.
6. Tính toán điểm kinh nghiệm (XP), Cấp độ (Level Up) và Danh hiệu nghiệp vụ của hệ thống AI.
"""

from __future__ import annotations

import os
import sys
from typing import Any, Dict, List, Optional

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from aec_core.experience_store import (
    ProjectExperienceStore,
    ProductivityObservation,
    GoldenRebarPattern,
    ImmunityRule,
    CandidateSkill,
)
from core.gates.human_gate import HumanGate, ApprovalRequest, ApprovalDecision


class AECExperienceAgent:
    """
    Tác tử Quản trị Tiến hóa & Tích lũy Kinh nghiệm Thực chiến AEC.
    """

    def __init__(self, store_path: Optional[str] = None):
        self.store = ProjectExperienceStore(store_path=store_path)

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 1: HIỆU CHUẨN NĂNG SUẤT HIỆN TRƯỜNG
    # ─────────────────────────────────────────────────────────────────────────

    def ingest_asbuilt_logs(
        self,
        project_id: str,
        site_logs: List[Dict[str, Any]],
        planned_productivity_map: Optional[Dict[str, float]] = None
    ) -> List[ProductivityObservation]:
        """
        Nạp dữ liệu nhật ký hiện trường từ AsBuiltAgent để tự động cập nhật hệ số hiệu chuẩn.
        """
        planned_map = planned_productivity_map or {
            "dao_mong_cong": 401.07,
            "ep_coc_khoan_nhoi": 12.0,
            "lap_dung_cot_thep": 1.5,
            "do_be_tong_dam_san": 85.0,
            "lap_dat_ong_cong": 25.0,
        }

        observations = []
        for log in site_logs:
            weather = log.get("weather", "nang_tot")
            for task in log.get("tasks_executed", []):
                t_id = task.get("task_id", "")
                t_vol = float(task.get("actual_volume", 0.0))
                if t_vol <= 0:
                    continue

                planned_rate = planned_map.get(t_id, 100.0)
                obs = self.store.record_productivity(
                    task_code=t_id,
                    task_name=f"Hạng mục {t_id}",
                    planned_productivity=planned_rate,
                    actual_productivity=t_vol,
                    unit=task.get("unit", "m3"),
                    conditions={"weather": weather},
                    project_id=project_id,
                )
                observations.append(obs)

        return observations

    def get_calibrated_rate(
        self,
        task_code: str,
        default_rate: float,
        conditions: Optional[Dict[str, Any]] = None
    ) -> float:
        """Lấy định mức năng suất đã được hiệu chuẩn bằng dữ liệu thực chiến."""
        return self.store.get_calibrated_productivity(task_code, default_rate, conditions)

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 2: THƯ VIỆN MẪU CẮT THÉP VÀNG
    # ─────────────────────────────────────────────────────────────────────────

    def cache_optimal_rebar_solution(
        self,
        element_type: str,
        diameter_mm: int,
        demands: Dict[int, int],
        stock_bar_count: int,
        waste_pct: float,
        cutting_patterns: List[Dict[str, Any]],
        steel_grade: str = "CB400-V",
        project_name: str = ""
    ) -> Optional[GoldenRebarPattern]:
        """
        Lưu kết quả cắt thép vào thư viện mẫu vàng nếu đề-xê < 1.5%.
        """
        return self.store.save_golden_pattern(
            element_type=element_type,
            diameter_mm=diameter_mm,
            demands_dict=demands,
            stock_bar_count=stock_bar_count,
            waste_pct=waste_pct,
            cutting_patterns=cutting_patterns,
            steel_grade=steel_grade,
            project_origin=project_name,
            max_waste_threshold=1.5
        )

    def find_cached_rebar_pattern(
        self,
        element_type: str,
        diameter_mm: int,
        demands: Dict[int, int],
        steel_grade: str = "CB400-V"
    ) -> Optional[GoldenRebarPattern]:
        """Tra cứu nhanh mẫu cắt thép vàng (0 ms runtime)."""
        return self.store.lookup_golden_pattern(
            element_type=element_type,
            diameter_mm=diameter_mm,
            demands_dict=demands,
            steel_grade=steel_grade
        )

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 3: MIỄN DỊCH LỖI & PHẢN BIỆN TỰ ĐỘNG
    # ─────────────────────────────────────────────────────────────────────────

    def ingest_audit_findings(self, findings: List[str]) -> int:
        """
        Ghi nhận các phát hiện kiểm toán để tự động kích hoạt bộ đếm vi phạm quy tắc.
        """
        recorded = 0
        for f in findings:
            if "CẢNH BÁO SỐ CHẾT" in f or "không dùng công thức sống" in f:
                self.store.record_immunity_violation("RULE-MATH-002")
                recorded += 1
            elif "HÌNH HỌC" in f or "E*F*G*H*I" in f:
                self.store.record_immunity_violation("RULE-EXCEL-001")
                recorded += 1
        return recorded

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 4: TỰ ĐỘNG ĐÓNG GÓI KỸ NĂNG VỚI CỔNG PHÊ DUYỆT
    # ─────────────────────────────────────────────────────────────────────────

    def submit_candidate_skill_for_approval(
        self,
        skill_id: str,
        name: str,
        category: str,
        description: str,
        workflow_trigger: str,
        human_gate: HumanGate,
        prompt_template: str = "",
        code_snippet: str = ""
    ) -> ApprovalDecision:
        """
        Đề xuất kỹ năng mới và chuyển qua HumanGate để Kỹ sư trưởng phê duyệt.
        """
        skill = self.store.propose_candidate_skill(
            skill_id=skill_id,
            name=name,
            category=category,
            description=description,
            workflow_trigger=workflow_trigger,
            prompt_template=prompt_template,
            code_snippet=code_snippet
        )

        req = ApprovalRequest(
            gate_id=f"SKILL-APPROVAL-{skill_id}",
            gate_name=f"Phê duyệt Kỹ năng Mới: {name}",
            phase="EVOLUTION_SKILL_PACKAGING",
            document_ref=f"skill://{skill_id}",
            summary_data={
                "Skill Name": name,
                "Category": category,
                "Description": description,
                "Workflow Trigger": workflow_trigger,
            }
        )

        decision = human_gate.request_approval(req)
        if decision.approved:
            self.store.approve_skill(
                skill_id=skill_id,
                approver=decision.approver,
                approval_notes=decision.comments or "Phê chuẩn tự động bởi Kỹ sư trưởng"
            )
        else:
            self.store.reject_skill(
                skill_id=skill_id,
                approver=decision.approver,
                reason=decision.comments or "Từ chối phê chuẩn"
            )

        return decision

    # ─────────────────────────────────────────────────────────────────────────
    # BÁO CÁO CẤP ĐỘ HỆ THỐNG (LEVEL-UP REPORT)
    # ─────────────────────────────────────────────────────────────────────────

    def generate_evolution_report(self) -> str:
        """Tạo báo cáo chi tiết cấp độ tiến hóa và thành tựu của Hệ thống."""
        lvl_info = self.store.get_system_level()
        stats = lvl_info["stats"]
        breakdown = lvl_info["xp_breakdown"]

        lines = [
            "# =====================================================================",
            "# 🏆 BÁO CÁO TIẾN HÓA & CẤP ĐỘ HỆ THỐNG AEC MULTI-AGENT (LEVEL-UP)",
            "# =====================================================================",
            f"  ⭐ CẤP ĐỘ HIỆN TẠI (LEVEL)      : LEVEL {lvl_info['level']}",
            f"  🎖️ DANH HIỆU NGHỆP VỤ          : {lvl_info['rank']}",
            f"  ⚡ TỔNG ĐIỂM KINH NGHIỆM (XP)   : {lvl_info['total_xp']:,} XP",
            f"  📈 TIẾN ĐỘ LÊN LEVEL {lvl_info['next_level']}       : {lvl_info['progress_to_next_pct']:.1f}% "
            f"({lvl_info['total_xp']} / {lvl_info['xp_needed_for_next']} XP)",
            "---------------------------------------------------------------------",
            "  📊 THỐNG KÊ TÍCH LŨY KINH NGHIỆM THỰC CHIẾN:",
            f"     - Số dự án đã hoàn thành          : {stats['projects_completed']} dự án (+{breakdown['projects']} XP)",
            f"     - Quan trắc năng suất hiện trường : {stats['productivity_observations']} mẫu (+{breakdown['productivity']} XP)",
            f"     - Mẫu cắt thép vàng tối ưu        : {stats['golden_patterns_stored']} mẫu (+{breakdown['rebar_patterns']} XP)",
            f"     - Số lần tái sử dụng mẫu vàng     : {stats['golden_pattern_reuses']} lần",
            f"     - Bộ quy tắc miễn dịch lỗi active : {stats['active_immunity_rules']} quy tắc (+{breakdown['immunity_rules']} XP)",
            f"     - Kỹ năng mới đã phê duyệt (Skills): {stats['approved_skills']} kỹ năng (+{breakdown['skills']} XP)",
            "=====================================================================",
        ]

        # In tóm tắt hiệu chuẩn năng suất
        prod_summary = self.store.get_productivity_summary()
        if prod_summary:
            lines.append("  ⚙️ KẾT QUẢ HIỆU CHUẨN ĐỊNH MỨC NĂNG SUẤT (FIELD CALIBRATION):")
            for code, data in prod_summary.items():
                lines.append(
                    f"     • {code} ({data['task_name']}): Hệ số alpha = {data['calibration_factor']:.3f} "
                    f"(từ {data['total_observations']} lần quan trắc thực tế)"
                )
            lines.append("---------------------------------------------------------------------")

        # In danh sách kỹ năng đã tốt nghiệp
        approved_skills = self.store.get_approved_skills()
        if approved_skills:
            lines.append("  🎓 DANH MỤC KỸ NĂNG ĐÃ TỐT NGHIỆP (GRADUATED SKILLS):")
            for sk in approved_skills:
                lines.append(f"     ✓ [{sk.skill_id}] {sk.name} — Phê duyệt bởi {sk.approved_by}")
            lines.append("=====================================================================")

        return "\n".join(lines)
