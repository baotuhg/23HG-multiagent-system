# -*- coding: utf-8 -*-
"""
AEC PROJECT EXPERIENCE & CONTINUOUS EVOLUTION ENGINE
Bộ lưu trữ kinh nghiệm thực chiến & Cơ chế tự nâng cấp (Level-Up / Self-Evolving)
Dành cho Hệ thống Multi-Agent AEC (23HG-multiagent-system).

Tuân thủ nguyên tắc:
  - ZERO LLM MATH: Tính toán hiệu chuẩn năng suất, thống kê sai số là công thức toán học xác định 100%.
  - HUMAN-IN-THE-LOOP: Mọi kỹ năng mới (Candidate Skills) bắt buộc phải qua cổng Human Gate phê duyệt.
  - DETERMINISTIC PERSISTENCE: Dữ liệu được lưu trữ nguyên tử (Atomic Write) dưới dạng JSON/SQLite cục bộ,
    không phụ thuộc cloud hay dịch vụ ngoài.

4 Cấp độ Tự Tiến Hóa (4 Evolution Levels):
  Level 1: Field Productivity Calibration (Hiệu chuẩn năng suất thi công thực tế từ As-Built).
  Level 2: Golden Rebar Cutting Patterns (Thư viện mẫu cắt thép vàng < 1.5% đề-xê, tra cứu O(1)).
  Level 3: Reflexion & Error Immunity (Miễn dịch lỗi tự động, mở rộng bộ quy tắc kiểm toán AEC).
  Level 4: Autonomous Skill Packaging (Đóng gói Skill mới với Human Gate kiểm soát).
"""

from __future__ import annotations

import os
import json
import math
import hashlib
import datetime
import tempfile
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union


# ─────────────────────────────────────────────────────────────────────────────
# DATA MODELS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ProductivityObservation:
    """Bản ghi quan trắc năng suất thực tế công trường."""
    observation_id: str
    project_id: str
    task_code: str
    task_name: str
    unit: str
    planned_productivity: float      # Năng suất thiết kế / định mức (đơn vị/ca)
    actual_productivity: float       # Năng suất thực tế đo đạc (đơn vị/ca)
    ratio: float                     # actual / planned
    conditions: Dict[str, Any] = field(default_factory=dict)
    recorded_at: str = ""

    def __post_init__(self):
        if not self.recorded_at:
            self.recorded_at = datetime.datetime.now().isoformat(timespec="seconds")
        if self.planned_productivity > 0 and self.ratio == 0.0:
            self.ratio = round(self.actual_productivity / self.planned_productivity, 4)


@dataclass
class GoldenRebarPattern:
    """Mẫu cắt thép vàng tối ưu (< 1.5% đề-xê) cho cấu kiện chuẩn hóa."""
    pattern_id: str
    element_type: str                # e.g., 'coc_khoan_nhoi_d1000', 'dam_super_t_33m', 'mo_cau'
    diameter_mm: int
    steel_grade: str
    demand_signature: str            # Hash signature của danh sách đoạn cắt {length_mm: count}
    demands_summary: Dict[str, int]  # Tóm tắt số lượng đoạn cắt
    stock_bar_count: int             # Số cây thép nguyên 11.7m
    waste_pct: float                 # Tỷ lệ hao hụt đề-xê (%)
    cutting_patterns: List[Dict[str, Any]] = field(default_factory=list)
    project_origin: str = ""
    times_reused: int = 0
    created_at: str = ""

    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.datetime.now().isoformat(timespec="seconds")


@dataclass
class ImmunityRule:
    """Quy tắc miễn dịch lỗi học được từ các sự cố/lỗi kiểm toán thực tế."""
    rule_id: str
    name: str
    category: str                    # 'EXCEL_INTEGRITY', 'ENGINEERING_SAFETY', 'LEGAL_NORMS', 'CPM_SCHEDULE'
    trigger_description: str         # Điều kiện kích hoạt lỗi
    severity: str                    # 'CRITICAL', 'WARNING', 'INFO'
    fix_recommendation: str          # Hướng dẫn khắc phục triệt để
    times_triggered: int = 0
    status: str = "ACTIVE"           # 'ACTIVE', 'MUTED', 'GRADUATED'
    created_at: str = ""

    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.datetime.now().isoformat(timespec="seconds")


@dataclass
class CandidateSkill:
    """Kỹ năng mới được hệ thống đề xuất từ kinh nghiệm dự án."""
    skill_id: str
    name: str
    category: str                    # 'TAKEOFF', 'OPTIMIZATION', 'QUALITY_CONTROL', 'AUTOMATION'
    description: str
    workflow_trigger: str            # Thời điểm hoặc điều kiện kích hoạt
    prompt_template: str = ""
    code_snippet: str = ""
    status: str = "PENDING_APPROVAL" # 'PENDING_APPROVAL', 'APPROVED', 'REJECTED'
    approved_by: str = ""
    approval_notes: str = ""
    decided_at: str = ""
    created_at: str = ""
    times_executed: int = 0

    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.datetime.now().isoformat(timespec="seconds")


# ─────────────────────────────────────────────────────────────────────────────
# CORE EXPERIENCE STORE
# ─────────────────────────────────────────────────────────────────────────────

class ProjectExperienceStore:
    """
    Kho lưu trữ kinh nghiệm và Quản trị tiến hóa (Self-Evolving / Level-Up Engine).
    
    Tự động ghi nhận bài học qua từng dự án, hiệu chuẩn hệ số định mức,
    lưu trữ nghiệm thức cắt thép vàng và phát hiện quy tắc miễn dịch lỗi.
    """

    DEFAULT_IMMUNITY_RULES = [
        ImmunityRule(
            rule_id="RULE-EXCEL-001",
            name="Ngăn ngừa lỗi XML Formula do text bắt đầu bằng '='",
            category="EXCEL_INTEGRITY",
            trigger_description="Ô text, ghi chú hoặc diễn giải bắt đầu bằng ký tự '=' khiến Excel hiểu nhầm là công thức hỏng.",
            severity="CRITICAL",
            fix_recommendation="Thêm dấu nháy đơn (') phía trước hoặc loại bỏ tiền tố '=' đối với text diễn giải.",
            status="ACTIVE"
        ),
        ImmunityRule(
            rule_id="RULE-MATH-002",
            name="Quy tắc Số Chết (Zero LLM Math Guard)",
            category="LEGAL_NORMS",
            trigger_description="Giá trị thành tiền, khối lượng tổng hợp trong bảng tính không dùng công thức sống mà điền số tĩnh.",
            severity="CRITICAL",
            fix_recommendation="Bắt buộc sử dụng công thức sống =C*D hoặc liên kết động =Sheet!Cell.",
            status="ACTIVE"
        ),
        ImmunityRule(
            rule_id="RULE-REBAR-003",
            name="Cấm nối thép tại vùng mômen uốn cực đại",
            category="ENGINEERING_SAFETY",
            trigger_description="Mối nối chồng cốt thép rơi vào vùng kéo căng lớn nhất (giữa nhịp dầm/đỉnh mômen) không theo BPTC.",
            severity="CRITICAL",
            fix_recommendation="Chuyển mối nối sang vùng nén hoặc bố trí so le theo TCVN 5574:2018.",
            status="ACTIVE"
        ),
        ImmunityRule(
            rule_id="RULE-CPM-004",
            name="Kiểm tra đứt gãy quan hệ tiền định CPM",
            category="CPM_SCHEDULE",
            trigger_description="Tác vụ trung gian không có successor nối tiếp dẫn tới tính sai đường găng.",
            severity="WARNING",
            fix_recommendation="Rà soát toàn bộ tác vụ không phải End Event để đảm bảo có ít nhất 1 successor FS/SS.",
            status="ACTIVE"
        ),
        ImmunityRule(
            rule_id="RULE-QAQC-005",
            name="Thời gian bảo dưỡng bê tông trước khi dỡ cốp pha",
            category="ENGINEERING_SAFETY",
            trigger_description="Lịch tháo cốp pha đáy dầm sàn sớm hơn kết quả nén mẫu R7/R28 theo TCVN 4453:1995.",
            severity="CRITICAL",
            fix_recommendation="Bắt buộc có phiếu thí nghiệm nén mẫu R >= 70-80% R28 trước khi ra lệnh dỡ cốp pha.",
            status="ACTIVE"
        ),
    ]

    def __init__(self, store_path: Optional[str] = None):
        """
        Khởi tạo ProjectExperienceStore.
        Nếu store_path là None, sử dụng đường dẫn mặc định trong .aec_state/experience_store.json.
        """
        self._custom_path = store_path is not None
        if store_path is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            state_dir = os.environ.get("AEC_STATE_DIR") or os.path.join(base_dir, ".aec_state")
            self.store_path = os.path.join(state_dir, "experience_store.json")
        else:
            self.store_path = store_path

        self.projects_history: List[str] = []
        self.productivity_observations: Dict[str, List[ProductivityObservation]] = {}
        self.golden_patterns: Dict[str, GoldenRebarPattern] = {}
        self.immunity_rules: Dict[str, ImmunityRule] = {}
        self.candidate_skills: Dict[str, CandidateSkill] = {}

        self._load()

    # ─────────────────────────────────────────────────────────────────────────
    # PERSISTENCE (ATOMIC READ/WRITE)
    # ─────────────────────────────────────────────────────────────────────────

    def _load(self) -> None:
        """Đọc toàn bộ dữ liệu kinh nghiệm từ file JSON."""
        # Khởi tạo các immunity rule mặc định
        for r in self.DEFAULT_IMMUNITY_RULES:
            self.immunity_rules[r.rule_id] = r

        load_path = self.store_path
        if not os.path.exists(self.store_path):
            if not getattr(self, "_custom_path", False):
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                seed_path = os.path.join(base_dir, "data", "experience_store_seed.json")
                if os.path.exists(seed_path):
                    load_path = seed_path
                else:
                    return
            else:
                return

        try:
            with open(load_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.projects_history = data.get("projects_history", [])

            # Load productivity observations
            prod_data = data.get("productivity_observations", {})
            for task_code, obs_list in prod_data.items():
                self.productivity_observations[task_code] = [
                    ProductivityObservation(**item) for item in obs_list
                ]

            # Load golden patterns
            patterns_data = data.get("golden_patterns", {})
            for pid, pat in patterns_data.items():
                self.golden_patterns[pid] = GoldenRebarPattern(**pat)

            # Load immunity rules
            rules_data = data.get("immunity_rules", {})
            for rid, rule in rules_data.items():
                self.immunity_rules[rid] = ImmunityRule(**rule)

            # Load candidate skills
            skills_data = data.get("candidate_skills", {})
            for sid, sk in skills_data.items():
                self.candidate_skills[sid] = CandidateSkill(**sk)

        except Exception as e:
            print(f"[ProjectExperienceStore] ⚠ Cảnh báo: Lỗi nạp store {self.store_path} ({e}). Dùng bộ nhớ đệm.")

    def save(self) -> None:
        """Lưu toàn bộ kinh nghiệm tích lũy xuống đĩa bằng atomic write."""
        os.makedirs(os.path.dirname(os.path.abspath(self.store_path)), exist_ok=True)

        data = {
            "version": "1.0.0",
            "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "projects_history": self.projects_history,
            "productivity_observations": {
                tc: [asdict(obs) for obs in obs_list]
                for tc, obs_list in self.productivity_observations.items()
            },
            "golden_patterns": {
                pid: asdict(pat) for pid, pat in self.golden_patterns.items()
            },
            "immunity_rules": {
                rid: asdict(rule) for rid, rule in self.immunity_rules.items()
            },
            "candidate_skills": {
                sid: asdict(sk) for sid, sk in self.candidate_skills.items()
            }
        }

        # Ghi file tạm rồi rename để đảm bảo Atomic Write, an toàn trước crash
        dir_name = os.path.dirname(os.path.abspath(self.store_path))
        with tempfile.NamedTemporaryFile("w", dir=dir_name, delete=False, encoding="utf-8") as tf:
            json.dump(data, tf, ensure_ascii=False, indent=2)
            temp_name = tf.name

        os.replace(temp_name, self.store_path)

    def record_project_completion(self, project_name: str) -> None:
        """Ghi nhận dự án đã hoàn thành vào lịch sử."""
        if project_name and project_name not in self.projects_history:
            self.projects_history.append(project_name)
            self.save()

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 1: FIELD PRODUCTIVITY CALIBRATION
    # ─────────────────────────────────────────────────────────────────────────

    def record_productivity(
        self,
        task_code: str,
        task_name: str,
        planned_productivity: float,
        actual_productivity: float,
        unit: str = "ca",
        conditions: Optional[Dict[str, Any]] = None,
        project_id: str = ""
    ) -> ProductivityObservation:
        """
        Ghi nhận quan trắc năng suất hiện trường thực tế.
        """
        obs_id = f"OBS-{task_code}-{len(self.productivity_observations.get(task_code, [])) + 1:04d}"
        obs = ProductivityObservation(
            observation_id=obs_id,
            project_id=project_id,
            task_code=task_code,
            task_name=task_name,
            unit=unit,
            planned_productivity=float(planned_productivity),
            actual_productivity=float(actual_productivity),
            ratio=0.0,
            conditions=conditions or {}
        )
        if task_code not in self.productivity_observations:
            self.productivity_observations[task_code] = []
        self.productivity_observations[task_code].append(obs)
        self.save()
        return obs

    def get_calibration_factor(
        self,
        task_code: str,
        conditions: Optional[Dict[str, Any]] = None,
        min_samples: int = 1
    ) -> float:
        """
        Tính toán hệ số hiệu chuẩn năng suất thực tế (Calibration Factor alpha).
        alpha = actual / planned.
        
        Quy tắc tính toán:
          - Lọc bỏ ngoại lai (outlier filtering): chỉ chấp nhận tỷ lệ trong đoạn [0.35, 2.8].
          - Nếu chưa đủ số mẫu quan trắc (samples < min_samples): trả về 1.0 (chuẩn định mức).
          - Tính trung bình có trọng số (các mẫu gần đây nhất có trọng số cao hơn).
        """
        history = self.productivity_observations.get(task_code, [])
        if not history or len(history) < min_samples:
            return 1.0

        # Lọc mẫu hợp lệ
        valid_ratios: List[float] = []
        for obs in history:
            # Nếu có điều kiện lọc (ví dụ cùng thời tiết mưa hoặc cùng địa hình)
            if conditions:
                match = True
                for k, v in conditions.items():
                    if obs.conditions.get(k) != v:
                        match = False
                        break
                if not match:
                    continue
            if 0.35 <= obs.ratio <= 2.8:
                valid_ratios.append(obs.ratio)

        # Nếu lọc theo điều kiện cụ thể không còn đủ mẫu, fallback về toàn bộ lịch sử của task_code
        if not valid_ratios and conditions:
            for obs in history:
                if 0.35 <= obs.ratio <= 2.8:
                    valid_ratios.append(obs.ratio)

        if not valid_ratios:
            return 1.0

        # Tính trung bình có trọng số tuyến tính (mẫu mới nhất trọng số lớn nhất)
        n = len(valid_ratios)
        weights = [i + 1 for i in range(n)]
        weighted_sum = sum(r * w for r, w in zip(valid_ratios, weights))
        factor = weighted_sum / sum(weights)
        return round(factor, 4)

    def get_calibrated_productivity(
        self,
        task_code: str,
        default_productivity: float,
        conditions: Optional[Dict[str, Any]] = None
    ) -> float:
        """
        Trả về năng suất đã được hiệu chuẩn theo kinh nghiệm thực tế.
        """
        alpha = self.get_calibration_factor(task_code, conditions=conditions)
        return round(default_productivity * alpha, 2)

    def get_productivity_summary(self, task_code: Optional[str] = None) -> Dict[str, Any]:
        """Tổng hợp thống kê năng suất đã quan trắc."""
        codes = [task_code] if task_code else list(self.productivity_observations.keys())
        summary = {}
        for code in codes:
            obs_list = self.productivity_observations.get(code, [])
            if not obs_list:
                continue
            ratios = [o.ratio for o in obs_list]
            summary[code] = {
                "task_name": obs_list[-1].task_name,
                "unit": obs_list[-1].unit,
                "total_observations": len(obs_list),
                "calibration_factor": self.get_calibration_factor(code),
                "min_ratio": min(ratios),
                "max_ratio": max(ratios),
                "last_observed_at": obs_list[-1].recorded_at,
            }
        return summary

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 2: GOLDEN REBAR CUTTING PATTERN LIBRARY
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def compute_demand_signature(demands_dict: Dict[int, int]) -> str:
        """
        Tạo hash signature xác định cho tập hợp nhu cầu cắt thép.
        demands_dict: {length_mm: count}
        """
        sorted_pairs = sorted(demands_dict.items(), key=lambda x: x[0])
        sig_str = "|".join(f"{length}:{qty}" for length, qty in sorted_pairs)
        return hashlib.sha256(sig_str.encode("utf-8")).hexdigest()[:16]

    def save_golden_pattern(
        self,
        element_type: str,
        diameter_mm: int,
        demands_dict: Dict[int, int],
        stock_bar_count: int,
        waste_pct: float,
        cutting_patterns: List[Dict[str, Any]],
        steel_grade: str = "CB400-V",
        project_origin: str = "",
        max_waste_threshold: float = 1.5
    ) -> Optional[GoldenRebarPattern]:
        """
        Lưu mẫu cắt thép vàng nếu tỷ lệ hao hụt đạt chuẩn (< 1.5%).
        """
        if waste_pct > max_waste_threshold:
            return None  # Không đạt chuẩn mẫu vàng

        sig = self.compute_demand_signature(demands_dict)
        pid = f"GOLDEN-{element_type}-D{diameter_mm}-{sig[:8]}"

        summary_dict = {str(k): v for k, v in demands_dict.items()}

        pattern = GoldenRebarPattern(
            pattern_id=pid,
            element_type=element_type,
            diameter_mm=diameter_mm,
            steel_grade=steel_grade,
            demand_signature=sig,
            demands_summary=summary_dict,
            stock_bar_count=stock_bar_count,
            waste_pct=round(waste_pct, 3),
            cutting_patterns=cutting_patterns,
            project_origin=project_origin,
            times_reused=0
        )
        self.golden_patterns[pid] = pattern
        self.save()
        return pattern

    def lookup_golden_pattern(
        self,
        element_type: str,
        diameter_mm: int,
        demands_dict: Dict[int, int],
        steel_grade: str = "CB400-V"
    ) -> Optional[GoldenRebarPattern]:
        """
        Tra cứu mẫu cắt vàng có sẵn cho cùng cấu kiện và tập hợp đoạn cắt.
        Nếu tìm thấy, tăng biến đếm số lần tái sử dụng (reuse counter).
        """
        sig = self.compute_demand_signature(demands_dict)
        for pattern in self.golden_patterns.values():
            if (pattern.element_type == element_type and
                pattern.diameter_mm == diameter_mm and
                pattern.steel_grade == steel_grade and
                pattern.demand_signature == sig):
                pattern.times_reused += 1
                self.save()
                return pattern
        return None

    def list_golden_patterns(self) -> List[GoldenRebarPattern]:
        """Trả về toàn bộ mẫu cắt vàng đã lưu."""
        return list(self.golden_patterns.values())

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 3: REFLEXION & ERROR IMMUNITY ENGINE
    # ─────────────────────────────────────────────────────────────────────────

    def register_immunity_rule(
        self,
        rule_id: str,
        name: str,
        category: str,
        trigger_description: str,
        severity: str = "CRITICAL",
        fix_recommendation: str = ""
    ) -> ImmunityRule:
        """
        Đăng ký một quy tắc kiểm toán / phòng ngừa lỗi mới vào hệ thống.
        """
        rule = ImmunityRule(
            rule_id=rule_id,
            name=name,
            category=category,
            trigger_description=trigger_description,
            severity=severity,
            fix_recommendation=fix_recommendation,
            times_triggered=0,
            status="ACTIVE"
        )
        self.immunity_rules[rule_id] = rule
        self.save()
        return rule

    def record_immunity_violation(self, rule_id: str) -> bool:
        """Ghi nhận một lần quy tắc kiểm toán bị vi phạm (trigger)."""
        if rule_id in self.immunity_rules:
            self.immunity_rules[rule_id].times_triggered += 1
            self.save()
            return True
        return False

    def get_active_immunity_rules(self, category: Optional[str] = None) -> List[ImmunityRule]:
        """Lấy danh sách các quy tắc phòng ngừa đang kích hoạt."""
        rules = [r for r in self.immunity_rules.values() if r.status == "ACTIVE"]
        if category:
            rules = [r for r in rules if r.category == category]
        return rules

    def audit_text_cell(self, cell_value: Any) -> List[str]:
        """
        Kiểm tra nhanh giá trị chuỗi văn bản xem có vi phạm quy tắc phòng vệ không.
        Ví dụ: Text bắt đầu bằng '=' gây lỗi openpyxl XML.
        """
        violations = []
        if isinstance(cell_value, str):
            s = cell_value.strip()
            # Kiểm tra RULE-EXCEL-001
            if s.startswith("=") and not any(op in s for op in ["SUM", "+", "-", "*", "/", "(", ")", "IF", "VLOOKUP", "ROUND"]):
                # Text mô tả tự nhiên nhưng lỡ tay bắt đầu bằng '='
                self.record_immunity_violation("RULE-EXCEL-001")
                violations.append("RULE-EXCEL-001: Chuỗi text bắt đầu bằng '=' nhưng không phải công thức hợp lệ!")
        return violations

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL 4: AUTONOMOUS SKILL PACKAGING & HUMAN GATE APPROVAL
    # ─────────────────────────────────────────────────────────────────────────

    def propose_candidate_skill(
        self,
        skill_id: str,
        name: str,
        category: str,
        description: str,
        workflow_trigger: str,
        prompt_template: str = "",
        code_snippet: str = ""
    ) -> CandidateSkill:
        """
        Đề xuất một kỹ năng mới từ thực tế thi công.
        Mặc định trạng thái là PENDING_APPROVAL chờ Kỹ sư trưởng phê duyệt.
        """
        skill = CandidateSkill(
            skill_id=skill_id,
            name=name,
            category=category,
            description=description,
            workflow_trigger=workflow_trigger,
            prompt_template=prompt_template,
            code_snippet=code_snippet,
            status="PENDING_APPROVAL"
        )
        self.candidate_skills[skill_id] = skill
        self.save()
        return skill

    def approve_skill(
        self,
        skill_id: str,
        approver: str = "KS_TRUONG",
        approval_notes: str = "Đã kiểm tra và phê chuẩn"
    ) -> bool:
        """Phê duyệt kỹ năng mới đưa vào danh mục chính thức."""
        if skill_id not in self.candidate_skills:
            return False
        skill = self.candidate_skills[skill_id]
        skill.status = "APPROVED"
        skill.approved_by = approver
        skill.approval_notes = approval_notes
        skill.decided_at = datetime.datetime.now().isoformat(timespec="seconds")
        self.save()
        return True

    def reject_skill(
        self,
        skill_id: str,
        approver: str = "KS_TRUONG",
        reason: str = "Không phù hợp tiêu chuẩn dự án"
    ) -> bool:
        """Từ chối kỹ năng đề xuất."""
        if skill_id not in self.candidate_skills:
            return False
        skill = self.candidate_skills[skill_id]
        skill.status = "REJECTED"
        skill.approved_by = approver
        skill.approval_notes = reason
        skill.decided_at = datetime.datetime.now().isoformat(timespec="seconds")
        self.save()
        return True

    def get_approved_skills(self) -> List[CandidateSkill]:
        """Lấy danh sách các kỹ năng đã được phê chuẩn."""
        return [s for s in self.candidate_skills.values() if s.status == "APPROVED"]

    def export_skill_markdown(self, skill_id: str, output_path: Optional[str] = None) -> str:
        """Xuất kỹ năng sang tài liệu chuẩn SKILL.md."""
        if skill_id not in self.candidate_skills:
            raise KeyError(f"Không tìm thấy skill {skill_id}")
        skill = self.candidate_skills[skill_id]

        md_content = f"""# SKILL: {skill.name}
**Skill ID:** `{skill.skill_id}`  
**Category:** `{skill.category}`  
**Status:** `{skill.status}` (Approved by: `{skill.approved_by}`)  
**Decided at:** `{skill.decided_at}`  

## 1. MÔ TẢ KỸ NĂNG
{skill.description}

## 2. ĐIỀU KIỆN KÍCH HOẠT (TRIGGER)
`{skill.workflow_trigger}`

## 3. PROMPT TEMPLATE / CHỈ DẪN THỰC THI
```markdown
{skill.prompt_template}
```

## 4. CODE SNIPPET (NẾU CÓ)
```python
{skill.code_snippet}
```

---
*Tài liệu được đóng gói tự động bởi ProjectExperienceStore — 23HG MultiAgent AEC System.*
"""
        if output_path:
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_content)
        return md_content

    # ─────────────────────────────────────────────────────────────────────────
    # LEVEL-UP PROGRESSION (EXP & RANK CALCULATION)
    # ─────────────────────────────────────────────────────────────────────────

    def get_system_level(self) -> Dict[str, Any]:
        """
        Tính toán Điểm kinh nghiệm (XP), Cấp độ (Level) và Danh hiệu của Hệ thống AI.
        
        Công thức điểm kinh nghiệm (XP Formula):
          - Mỗi dự án hoàn thành: 100 XP
          - Mỗi bản ghi hiệu chuẩn năng suất: 10 XP
          - Mỗi mẫu cắt thép vàng được lưu trữ: 25 XP
          - Mỗi lần tái sử dụng mẫu cắt vàng: 15 XP
          - Mỗi quy tắc phòng ngừa lỗi active: 10 XP
          - Mỗi kỹ năng mới được phê duyệt (Approved Skill): 50 XP
        
        Cấp độ (Level):
          Level = 1 + int(sqrt(Total_XP / 40))
        """
        proj_count = len(self.projects_history)
        obs_count = sum(len(obs) for obs in self.productivity_observations.values())
        pat_count = len(self.golden_patterns)
        pat_reuses = sum(p.times_reused for p in self.golden_patterns.values())
        rules_count = len(self.get_active_immunity_rules())
        skills_count = len(self.get_approved_skills())

        xp_projects = proj_count * 100
        xp_obs = obs_count * 10
        xp_patterns = pat_count * 25 + pat_reuses * 15
        xp_rules = rules_count * 10
        xp_skills = skills_count * 50

        total_xp = xp_projects + xp_obs + xp_patterns + xp_rules + xp_skills
        level = 1 + int(math.sqrt(total_xp / 40.0))

        # Danh hiệu cấp bậc theo phong cách Kỹ sư Xây dựng
        if level < 5:
            rank = "Kỹ sư Tập sự (Novice Assistant)"
        elif level < 10:
            rank = "Kỹ sư Giám sát Hiện trường (Field Engineer)"
        elif level < 18:
            rank = "Kỹ sư QS & BPTC Chuyên nghiệp (Senior QS/Site Engineer)"
        elif level < 28:
            rank = "Chỉ huy phó Công trình (Deputy Site Manager)"
        elif level < 40:
            rank = "Chỉ huy trưởng Công trình (Chief Site Manager)"
        else:
            rank = "Chuyên gia AEC Trưởng (Master Chief AEC Architect)"

        # XP cho level tiếp theo
        next_level = level + 1
        xp_needed_next = int(((next_level - 1) ** 2) * 40.0)
        xp_current_level_base = int(((level - 1) ** 2) * 40.0)
        progress_pct = 0.0
        if xp_needed_next > xp_current_level_base:
            progress_pct = round(
                ((total_xp - xp_current_level_base) / (xp_needed_next - xp_current_level_base)) * 100.0, 1
            )

        return {
            "level": level,
            "rank": rank,
            "total_xp": total_xp,
            "next_level": next_level,
            "xp_needed_for_next": xp_needed_next,
            "progress_to_next_pct": min(100.0, max(0.0, progress_pct)),
            "stats": {
                "projects_completed": proj_count,
                "productivity_observations": obs_count,
                "golden_patterns_stored": pat_count,
                "golden_pattern_reuses": pat_reuses,
                "active_immunity_rules": rules_count,
                "approved_skills": skills_count,
            },
            "xp_breakdown": {
                "projects": xp_projects,
                "productivity": xp_obs,
                "rebar_patterns": xp_patterns,
                "immunity_rules": xp_rules,
                "skills": xp_skills,
            }
        }
