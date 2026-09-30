# -*- coding: utf-8 -*-
"""
AEC Master Core Engine
Hệ thống lõi tự động hóa tính toán Kỹ thuật, Dự toán, Tiến độ và Pháp lý Xây dựng.
Tuân thủ Luật Xây dựng 135/2025/QH15, NĐ 207/2026/NĐ-CP, NĐ 254/2025/NĐ-CP & TT 36/2026/TT-BXD.
"""

from .audit_verifier import AECAuditVerifier
from .project_state import ProjectStateManager
from .experience_store import (
    ProjectExperienceStore,
    ProductivityObservation,
    GoldenRebarPattern,
    ImmunityRule,
    CandidateSkill,
)

__version__ = "2.1.0"
__all__ = [
    "AECAuditVerifier",
    "ProjectStateManager",
    "ProjectExperienceStore",
    "ProductivityObservation",
    "GoldenRebarPattern",
    "ImmunityRule",
    "CandidateSkill",
]
