# -*- coding: utf-8 -*-
"""
SHARED STATE SCHEMA — AEC MultiAgent System v3.0
Sử dụng TypedDict thuần (không cần Pydantic) để tương thích Python 3.8+
Đây là Single Source of Truth (SSOT) cho toàn bộ State Graph.

Cấu trúc: ProjectSharedState
├── meta            : Thông tin phiên làm việc, trạng thái vòng lặp
├── project_info    : Thông tin dự án cầu cố định
├── execution       : Trạng thái từng Agent (NodeStatus)
├── cad_data        : Kết quả bóc tách CAD (Shoelace/Average-End-Area)
├── rebar_data      : BBS + kết quả cắt thép OR-Tools
├── qs_data         : Dự toán G_XD (T, GT, TL, VAT)
├── schedule_data   : CPM Tiến độ (planned vs actual)
├── qaqc_data       : QA/QC, kết quả thí nghiệm R7/R28, chứng chỉ
├── site_log        : Nhật ký hiện trường hàng ngày (As-Built loop)
└── approvals       : Human-in-the-loop gates (AWAITING / APPROVED / REJECTED)
"""

from __future__ import annotations
from enum import Enum
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Any, Dict, List, Optional
import json
import os
import threading


# ─────────────────────────────────────────────────────────────────────────────
# ENUMS — Trạng thái Node trong State Graph
# ─────────────────────────────────────────────────────────────────────────────

class NodeStatus(str, Enum):
    """Trạng thái thực thi của mỗi Sub-Agent node trong State Graph."""
    IDLE       = "IDLE"              # Chưa chạy
    RUNNING    = "RUNNING"           # Đang thực thi
    SUCCESS    = "SUCCESS"           # Hoàn thành thành công
    FAILED     = "FAILED"            # Lỗi — cần retry hoặc escalate
    REJECTED   = "REJECTED"          # Bị Quality Gate hoặc Inter-Agent phản biện từ chối
    AWAITING   = "AWAITING_APPROVAL" # Chờ Human-in-the-loop Approve
    APPROVED   = "APPROVED"          # Đã được Kỹ sư trưởng duyệt
    SKIPPED    = "SKIPPED"           # Bỏ qua (không cần thiết cho version này)


class ProjectPhase(str, Enum):
    """Giai đoạn dự án — xác định luồng State Graph."""
    INIT           = "INIT"
    CAD_TAKEOFF    = "CAD_TAKEOFF"
    REBAR_CUT      = "REBAR_CUT"
    QS_ESTIMATE    = "QS_ESTIMATE"
    QAQC_REVIEW    = "QAQC_REVIEW"
    HUMAN_GATE     = "HUMAN_GATE"
    SCHEDULE_CPM   = "SCHEDULE_CPM"
    ASBUILT_LOOP   = "ASBUILT_LOOP"
    PAYMENT_03A    = "PAYMENT_03A"
    COMPLETED      = "COMPLETED"
    ERROR          = "ERROR"


class ApprovalStatus(str, Enum):
    """Trạng thái cổng phê duyệt Human-in-the-loop."""
    PENDING  = "PENDING"
    APPROVED = "APPROVED"
    REJECTED = "REJECTED"


# ─────────────────────────────────────────────────────────────────────────────
# SUB-SCHEMAS — Dữ liệu từng miền chuyên môn
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class AgentRunRecord:
    """Ghi nhận kết quả mỗi lần Agent chạy (cho log và retry logic)."""
    agent_id: str = ""
    status: str = NodeStatus.IDLE
    started_at: str = ""
    finished_at: str = ""
    attempt: int = 0
    error_message: str = ""
    output_summary: str = ""


@dataclass
class CADTakeoffData:
    """Kết quả bóc tách khối lượng từ bản vẽ CAD."""
    # Bê tông
    concrete_components: List[Dict[str, Any]] = field(default_factory=list)
    # Format: [{"id": "COC-T1-01", "wbs": "KẾT CẤU MÓNG CỌC", "volume_m3": 45.24, ...}]
    total_concrete_m3: float = 0.0

    # Ván khuôn
    formwork_components: List[Dict[str, Any]] = field(default_factory=list)
    total_formwork_m2: float = 0.0

    # Đào đắp
    excavation_m3: float = 0.0
    backfill_m3: float = 0.0

    # Metadata
    drawings_scanned: int = 0
    drawings_processed: int = 0
    source_dwg_files: List[str] = field(default_factory=list)
    wbs_mapping: Dict[str, Any] = field(default_factory=dict)

    # CAD Diff / Versioning
    revision_tag: str = "Rev00"
    diff_from_previous: Dict[str, Any] = field(default_factory=dict)


@dataclass
class RebarScheduleItem:
    """Một dòng trong bảng thống kê cốt thép (BBS)."""
    mark: str = ""           # Ký hiệu (e.g. "T1", "D2")
    component: str = ""      # Cấu kiện (e.g. "Cọc Ø1200 T1")
    diameter_mm: int = 0     # Đường kính
    grade: str = "CB300-V"   # Cấp thép
    count: int = 0           # Số lượng thanh
    length_mm: int = 0       # Chiều dài đã uốn (mm)
    unit_weight_kg_m: float = 0.0
    total_weight_kg: float = 0.0
    zone: str = ""           # Vùng chịu lực (TENSION / COMPRESSION / SHEAR)


@dataclass
class CuttingStockResult:
    """Kết quả bài toán cắt thép 1D từ OR-Tools Solver."""
    solver_status: str = ""          # OPTIMAL / FEASIBLE / INFEASIBLE
    bar_length_mm: int = 11700       # Cây thép nguyên 11.7m
    total_bars_needed: int = 0
    waste_mm_per_bar: List[int] = field(default_factory=list)
    total_waste_kg: float = 0.0
    waste_ratio_pct: float = 0.0     # Phải < 1.5%
    cutting_patterns: List[Dict[str, Any]] = field(default_factory=list)
    # Format: [{"bar_id": 1, "cuts": [3000, 3000, 5700], "waste_mm": 0}, ...]
    inter_agent_validation: str = ""  # PASS / REJECT (từ BPTC Agent kiểm tra vùng nối)
    rejection_reason: str = ""


@dataclass
class RebarData:
    """Toàn bộ dữ liệu cốt thép."""
    bbs_items: List[RebarScheduleItem] = field(default_factory=list)
    total_rebar_kg: float = 0.0
    cutting_result: CuttingStockResult = field(default_factory=CuttingStockResult)

    # Inter-Agent validation result (từ BPTC_KCS Agent)
    splice_zone_check: str = "NOT_RUN"  # PASS / REJECT
    splice_violations: List[str] = field(default_factory=list)


@dataclass
class QSData:
    """Dự toán tổng hợp G_XD theo TT 11/2021/TT-BXD."""
    direct_cost_T_vnd: float = 0.0       # T — Chi phí trực tiếp
    indirect_cost_GT_vnd: float = 0.0    # GT = T × (tỷ lệ chi phí chung + nhà tạm + KXĐ)
    tax_TL_vnd: float = 0.0              # TL = (T + GT) × tỷ lệ thu nhập chịu thuế tính trước
    subtotal_vnd: float = 0.0            # T + GT + TL
    vat_vnd: float = 0.0                 # VAT = subtotal * 10%
    total_G_XD_vnd: float = 0.0          # G_XD = subtotal + VAT
    payment_period_03a_vnd: float = 0.0  # Phụ lục 03a kỳ thanh toán
    unit_price_breakdown: Dict[str, float] = field(default_factory=dict)
    supplement_items: List[Dict[str, Any]] = field(default_factory=list)  # Phát sinh
    gt_breakdown_vnd: Dict[str, float] = field(default_factory=dict)      # chung / nha_tam / kxd
    rates: Dict[str, float] = field(default_factory=dict)                 # tỷ lệ đã dùng (0.051 = 5.1%)
    rate_sources: Dict[str, str] = field(default_factory=dict)
    items_count: int = 0
    data_source: str = ""
    warnings: List[str] = field(default_factory=list)
    payment_summary: Dict[str, Any] = field(default_factory=dict)  # Mẫu 03a kỳ hiện tại


@dataclass
class LabTestResult:
    """Kết quả thí nghiệm vật liệu (QA/QC Lab Link)."""
    test_id: str = ""
    material_type: str = ""       # CONCRETE / REBAR / BITUMEN
    sample_code: str = ""
    test_date: str = ""
    r7_mpa: float = 0.0
    r28_mpa: float = 0.0
    required_mpa: float = 0.0
    status: str = ""              # PASS / FAIL
    certificate_ref: str = ""     # Số chứng chỉ / phiếu thí nghiệm
    kcs_record_linked: str = ""   # BBNT-XX liên kết


@dataclass
class QAQCData:
    """Quản lý chất lượng và kiểm soát chất lượng."""
    total_inspection_records: int = 0
    records: List[Dict[str, Any]] = field(default_factory=list)
    lab_results: List[LabTestResult] = field(default_factory=list)
    hold_points: List[str] = field(default_factory=list)
    date_cross_check_status: str = "NOT_RUN"  # PASSED / FAILED / NOT_RUN
    audit_score: int = 0
    clashes_detected: List[str] = field(default_factory=list)
    hold_point_status: List[Dict[str, Any]] = field(default_factory=list)  # BBNT → GIẢI TỎA / CHỜ / CHẶN
    lab_summary: Dict[str, Any] = field(default_factory=dict)


@dataclass
class CPMTask:
    """Một công việc trong biểu đồ CPM."""
    task_id: str = ""
    name: str = ""
    duration_days: int = 0
    predecessors: List[str] = field(default_factory=list)
    early_start: str = ""
    early_finish: str = ""
    late_start: str = ""
    late_finish: str = ""
    float_days: int = 0
    is_critical: bool = False
    planned_volume: float = 0.0
    actual_volume: float = 0.0   # Từ vòng lặp As-Built


@dataclass
class ScheduleData:
    """Dữ liệu tiến độ CPM."""
    start_date: str = ""
    finish_date: str = ""
    total_duration_days: int = 0
    tasks: List[CPMTask] = field(default_factory=list)
    critical_path: List[str] = field(default_factory=list)
    overall_progress_pct: float = 0.0  # % hoàn thành thực tế
    delay_days: int = 0


@dataclass
class DailySiteLog:
    """Nhật ký hiện trường ngày (As-Built Loop)."""
    log_date: str = ""
    weather: str = ""
    tasks_executed: List[Dict[str, Any]] = field(default_factory=list)
    # Format: [{"task_id": "CPM-01", "volume_done": 12.5, "unit": "m3", "worker_count": 8}]
    issues: List[str] = field(default_factory=list)
    logged_by: str = ""


@dataclass
class ApprovalGate:
    """Human-in-the-loop Gate — trước khi xuất tài liệu pháp lý."""
    gate_id: str = ""
    gate_name: str = ""           # e.g. "Phê duyệt Dự toán G_XD"
    status: str = ApprovalStatus.PENDING
    requested_at: str = ""
    resolved_at: str = ""
    approver: str = ""
    comments: str = ""
    document_ref: str = ""        # File cần phê duyệt
    clashes_to_review: List[str] = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────────────────
# MASTER SHARED STATE
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ProjectSharedState:
    """
    SHARED STATE BUS — Trung tâm điều phối của AI Supervisor.
    Mọi Sub-Agent chỉ đọc/ghi qua đây. KHÔNG agent nào gọi thẳng agent khác.
    """

    # ── Meta ─────────────────────────────────────────────────────────────────
    session_id: str = ""
    schema_version: str = "3.0.0"
    current_phase: str = ProjectPhase.INIT
    created_at: str = ""
    updated_at: str = ""
    max_retries: int = 3

    # ── Node Status Map ───────────────────────────────────────────────────────
    # Mỗi agent có entry: agent_id -> NodeStatus
    node_status: Dict[str, str] = field(default_factory=lambda: {
        "cad_agent":       NodeStatus.IDLE,
        "rebar_agent":     NodeStatus.IDLE,
        "bptc_kcs_agent":  NodeStatus.IDLE,
        "qs_agent":        NodeStatus.IDLE,
        "scheduler_agent": NodeStatus.IDLE,
        "asbuilt_agent":   NodeStatus.IDLE,
        "payment_agent":   NodeStatus.IDLE,
        "supervisor":      NodeStatus.RUNNING,
    })

    # ── Run History ───────────────────────────────────────────────────────────
    run_history: List[AgentRunRecord] = field(default_factory=list)

    # ── Project Fixed Info ────────────────────────────────────────────────────
    project_name: str = "Cầu Km19+529.080 — Cao tốc Tuyên Quang - Hà Giang"
    structure_id: str = "CAU-KM19-529"
    drawings_folder: str = ""
    excel_master_path: str = ""

    # ── Nguồn dữ liệu ─────────────────────────────────────────────────────────
    # demo_mode=False: agent thiếu dữ liệu thật phải báo lỗi, KHÔNG dùng dữ liệu mẫu.
    demo_mode: bool = False
    # Ghi nhận mọi chỗ đã dùng dữ liệu mẫu: [{"agent_id": ..., "note": ...}]
    sample_data_sources: List[Dict[str, str]] = field(default_factory=list)

    # ── Domain Data ───────────────────────────────────────────────────────────
    cad_data: CADTakeoffData = field(default_factory=CADTakeoffData)
    rebar_data: RebarData = field(default_factory=RebarData)
    qs_data: QSData = field(default_factory=QSData)
    qaqc_data: QAQCData = field(default_factory=QAQCData)
    schedule_data: ScheduleData = field(default_factory=ScheduleData)

    # ── As-Built Loop ─────────────────────────────────────────────────────────
    site_logs: List[DailySiteLog] = field(default_factory=list)

    # ── Human Gates ───────────────────────────────────────────────────────────
    approval_gates: List[ApprovalGate] = field(default_factory=list)

    # ── Error Bus ─────────────────────────────────────────────────────────────
    global_errors: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """Serialize toàn bộ state sang dict (cho JSON persistence)."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ProjectSharedState":
        """Deserialize từ dict (khi load từ file JSON)."""
        state = cls()
        # Simple field mapping — nested dataclasses cần reconstruct
        for k, v in data.items():
            if hasattr(state, k):
                setattr(state, k, v)
        return state
