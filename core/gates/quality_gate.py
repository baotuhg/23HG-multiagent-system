# -*- coding: utf-8 -*-
"""
QUALITY GATE — Kiểm soát chất lượng kỹ thuật trước khi chuyển phase
Tích hợp với AECAuditVerifier (100/100 điểm) và các kiểm tra chuyên ngành.

Khi nào Quality Gate được kích hoạt:
  - Trước khi chuyển phase CAD_TAKEOFF → REBAR_CUT
  - Trước khi chuyển phase REBAR_CUT → QS_ESTIMATE
  - Trước khi chuyển phase QS_ESTIMATE → HUMAN_GATE
  - Trước khi xuất tài liệu pháp lý (BBNT, Dự toán, Phụ lục 03a)

Cơ chế REJECT:
  - Nếu Gate FAIL → Supervisor gửi REJECTED về Agent tương ứng
  - Agent nhận REJECTED → chạy lại logic (retry ≤ max_retries)
  - Sau max_retries → escalate lên Human Gate
"""

from __future__ import annotations
import os
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class GateCheckResult:
    gate_name: str
    passed: bool
    score: int = 0
    max_score: int = 100
    issues: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    details: Dict[str, Any] = field(default_factory=dict)


class QualityGate:
    """
    Bộ cổng kiểm soát chất lượng — chạy các bộ kiểm tra xác định (deterministic).
    KHÔNG dùng LLM cho bất kỳ phán quyết nào.
    """

    def __init__(self, project_root: str):
        self.project_root = project_root
        # Thêm project root vào sys.path để import aec_core
        if project_root not in sys.path:
            sys.path.insert(0, project_root)

    # ── GATE 1: CAD Takeoff Sanity Check ─────────────────────────────────────

    def check_cad_takeoff(self, cad_data: Dict[str, Any]) -> GateCheckResult:
        """
        Kiểm tra kết quả bóc tách CAD trước khi chuyển sang Rebar Agent.
        Tiêu chí:
          1. Có ít nhất 1 cấu kiện được bóc tách
          2. Tổng bê tông > 0 m³
          3. Tất cả cấu kiện phải có WBS hợp lệ
          4. Không có volume âm
        """
        result = GateCheckResult(gate_name="CAD_TAKEOFF_GATE", passed=False)
        score = 0

        components = cad_data.get("concrete_components", [])
        total_m3 = cad_data.get("total_concrete_m3", 0)
        drawings_processed = cad_data.get("drawings_processed", 0)

        VALID_WBS = {
            "KẾT CẤU CHUNG", "KẾT CẤU NHỊP", "KẾT CẤU MẶT CẦU",
            "KẾT CẤU MỐ CẦU", "KẾT CẤU TRỤ CẦU", "KẾT CẤU MÓNG CỌC",
            "PHỤ TRỢ MẶT CẦU", "ĐẦU CẦU",
        }

        # Check 1: Số cấu kiện
        if len(components) >= 1:
            score += 25
        else:
            result.issues.append("Không có cấu kiện nào được bóc tách")

        # Check 2: Tổng bê tông > 0
        if total_m3 > 0:
            score += 25
        else:
            result.issues.append(f"Tổng bê tông = {total_m3} m³ — không hợp lệ")

        # Check 3: WBS hợp lệ
        invalid_wbs = [c.get("id", "?") for c in components
                       if c.get("wbs", "") not in VALID_WBS]
        if not invalid_wbs:
            score += 25
        else:
            result.warnings.append(f"{len(invalid_wbs)} cấu kiện có WBS không chuẩn: {invalid_wbs[:5]}")
            score += 10  # Partial credit

        # Check 4: Không volume âm
        neg_volumes = [c.get("id", "?") for c in components
                       if c.get("volume_m3", 0) < 0]
        if not neg_volumes:
            score += 25
        else:
            result.issues.append(f"Phát hiện {len(neg_volumes)} cấu kiện có thể tích âm!")

        result.score = score
        result.passed = (score >= 75 and not result.issues)
        result.details = {
            "components_count": len(components),
            "total_concrete_m3": total_m3,
            "drawings_processed": drawings_processed,
        }
        return result

    # ── GATE 2: Rebar / Cutting Stock Check ──────────────────────────────────

    def check_rebar_cutting(
        self,
        cutting_result: Dict[str, Any],
        splice_status: str,
        splice_violations: List[str]
    ) -> GateCheckResult:
        """
        Kiểm tra kết quả cắt thép:
          1. Solver trả về OPTIMAL hoặc FEASIBLE
          2. Số cây thép so với cận dưới (OPTIMAL = đã chứng minh ít nhất có thể)
          3. Inter-agent splice check: không vi phạm TCVN 5574 (NOT_RUN → cảnh báo)
          4. Số đoạn đã cắt khớp đúng số đoạn yêu cầu trong BBS
        Đề-xê > 1.5% chỉ là cảnh báo: với chiều dài thanh cho trước, có những BBS
        không thể đạt 1.5% dù phương án cắt đã tối ưu.
        """
        result = GateCheckResult(gate_name="REBAR_CUTTING_GATE", passed=False)
        score = 0

        status = cutting_result.get("status", "NOT_RUN")
        waste_pct = cutting_result.get("waste_ratio_pct", 999)
        weight_kg = cutting_result.get("total_weight_kg", 0)
        total_bars = cutting_result.get("total_bars_needed", 0)
        lower_bound = cutting_result.get("lower_bound_bars", 0)
        pieces_demanded = cutting_result.get("pieces_demanded")
        pieces_cut = cutting_result.get("pieces_cut")

        # Check 1: Solver status
        if status == "OPTIMAL":
            score += 35
        elif status == "FEASIBLE":
            score += 25
        else:
            result.issues.append(f"Solver status không hợp lệ: {status} — chưa chạy hoặc INFEASIBLE")

        # Check 2: Khoảng cách tới cận dưới số cây thép
        gap_bars = total_bars - lower_bound if lower_bound else None
        if status == "OPTIMAL":
            score += 30
        elif status == "FEASIBLE" and gap_bars is not None:
            gap_pct = gap_bars / lower_bound * 100
            if gap_pct <= 5.0:
                score += 25
                result.warnings.append(
                    f"Chưa chứng minh tối ưu: dùng {total_bars} cây, cận dưới {lower_bound} cây "
                    f"(dư tối đa {gap_bars} cây, {gap_pct:.1f}%)"
                )
            else:
                result.issues.append(
                    f"Phương án cắt cách cận dưới {gap_bars} cây ({gap_pct:.1f}%) — cần tối ưu lại"
                )
        elif status == "FEASIBLE":
            result.warnings.append("Không có cận dưới số cây — không đánh giá được mức tối ưu")

        if status in ("OPTIMAL", "FEASIBLE") and waste_pct > 1.5:
            result.warnings.append(
                f"Đề-xê {waste_pct:.2f}% > 1.5%"
                + (" — đã là mức thấp nhất với chiều dài thanh trong BBS" if status == "OPTIMAL" else "")
            )

        # Check 3: Splice zone validation (INTER-AGENT NEGOTIATION)
        if splice_status == "PASS":
            score += 25
        elif splice_status == "NOT_RUN":
            score += 10
            result.warnings.append(
                "Chưa kiểm tra vị trí nối (chưa có dữ liệu vị trí nối từ bản vẽ) — "
                "KS phải tự kiểm tra theo TCVN 5574:2018"
            )
        else:  # REJECT
            result.issues.append(
                f"REJECT: Nối thép vi phạm TCVN 5574:2018 — {len(splice_violations)} điểm vi phạm"
            )
            for v in splice_violations[:3]:
                result.issues.append(f"  → {v}")

        # Check 4: Đủ số đoạn cắt theo BBS
        if pieces_demanded is not None and pieces_cut is not None:
            if pieces_cut == pieces_demanded and total_bars > 0:
                score += 10
            else:
                result.issues.append(
                    f"Số đoạn đã cắt ({pieces_cut}) khác số đoạn yêu cầu trong BBS ({pieces_demanded})"
                )
        elif total_bars > 0:
            score += 10

        for w in cutting_result.get("input_summary", []):
            result.warnings.append(f"BBS: {w}")

        plan = cutting_result.get("splice_plan") or {}
        if plan.get("splices"):
            result.warnings.append(
                f"PA nối thép (đề xuất) giảm {plan['bars_saved']} cây với {plan['splices']} mối nối — "
                f"phiếu cắt chính thức vẫn là PA không nối cho tới khi kỹ thuật duyệt PA nối"
            )

        result.score = score
        result.passed = (score >= 70 and not result.issues)
        result.details = {
            "solver_status": status,
            "waste_ratio_pct": waste_pct,
            "total_weight_kg": weight_kg,
            "total_bars_needed": total_bars,
            "lower_bound_bars": lower_bound,
            "splice_status": splice_status,
        }
        return result

    # ── GATE 3: QS / Cost Estimate Check ─────────────────────────────────────

    def check_qs_estimate(self, qs_data: Dict[str, Any]) -> GateCheckResult:
        """
        Kiểm tra dự toán G_XD (TT 11/2021/TT-BXD) theo đúng tỷ lệ đã dùng (qs_data.rates):
          1. Có công tác và G_XD > 0
          2. GT = T × (chi phí chung + nhà tạm + KXĐ)
          3. TL = (T + GT) × tỷ lệ thu nhập chịu thuế tính trước
          4. VAT = G × thuế suất và G_XD = G + VAT
        """
        result = GateCheckResult(gate_name="QS_ESTIMATE_GATE", passed=False)
        score = 0

        T = qs_data.get("direct_cost_T_vnd", 0)
        GT = qs_data.get("indirect_cost_GT_vnd", 0)
        TL = qs_data.get("tax_TL_vnd", 0)
        VAT = qs_data.get("vat_vnd", 0)
        G_XD = qs_data.get("total_G_XD_vnd", 0)
        rates = qs_data.get("rates") or {}
        missing = [k for k in ("chung", "nha_tam", "kxd", "tl", "vat") if k not in rates]
        if missing:
            result.issues.append(f"Thiếu tỷ lệ đã dùng để tính G_XD: {missing} — không kiểm tra được")
            result.score = 0
            return result

        def close(actual, expected):
            return abs(actual - expected) <= max(2.0, 1e-6 * abs(expected))

        # Check 1: có dữ liệu
        if G_XD > 0 and T > 0 and qs_data.get("items_count", 1) > 0:
            score += 25
        else:
            result.issues.append(f"T = {T:,.0f}, G_XD = {G_XD:,.0f} — không hợp lệ")

        # Check 2: GT
        gt_rate = rates["chung"] + rates["nha_tam"] + rates["kxd"]
        if close(GT, T * gt_rate):
            score += 25
        else:
            result.issues.append(f"GT = {GT:,.0f} ≠ T × {gt_rate * 100:g}% = {T * gt_rate:,.0f}")

        # Check 3: TL
        if close(TL, (T + GT) * rates["tl"]):
            score += 25
        else:
            result.issues.append(f"TL = {TL:,.0f} ≠ (T+GT) × {rates['tl'] * 100:g}% = {(T + GT) * rates['tl']:,.0f}")

        # Check 4: VAT và G_XD
        G = T + GT + TL
        if close(VAT, G * rates["vat"]) and close(G_XD, G + VAT):
            score += 25
        else:
            result.issues.append(f"VAT/G_XD không khớp: VAT = {VAT:,.0f}, G × {rates['vat'] * 100:g}% = "
                                 f"{G * rates['vat']:,.0f}; G_XD = {G_XD:,.0f}, G + VAT = {G + VAT:,.0f}")

        for w in (qs_data.get("warnings") or [])[:10]:
            result.warnings.append(f"QS: {w}")

        result.score = score
        result.passed = (score >= 75 and not result.issues)
        result.details = {"T": T, "GT": GT, "TL": TL, "VAT": VAT, "G_XD": G_XD, "rates": rates}
        return result

    # ── GATE 4: Excel Audit (100/100) ────────────────────────────────────────

    def check_excel_audit(self, excel_path: str) -> GateCheckResult:
        """
        Chạy AECAuditVerifier — yêu cầu đạt 100/100 mới PASS.
        Đây là gate cứng nhất: 0 số chết, 0 link gãy.
        """
        result = GateCheckResult(gate_name="EXCEL_AUDIT_GATE", passed=False)

        if not os.path.exists(excel_path):
            result.issues.append(f"File Excel không tồn tại: {excel_path}")
            return result

        try:
            from aec_core.audit_verifier import AECAuditVerifier
            auditor = AECAuditVerifier(excel_path)
            auditor.audit_excel_workbook()

            result.score = auditor.score
            result.passed = (auditor.score >= 100)
            result.details = auditor.stats

            if auditor.score < 100:
                result.issues.append(
                    f"Audit score = {auditor.score}/100 — chưa đạt 100/100. "
                    f"Kiểm tra: {auditor.stats.get('issues', [])}"
                )
        except Exception as e:
            result.issues.append(f"Lỗi chạy AECAuditVerifier: {e}")

        return result

    # ── AGGREGATE: Run all gates for a phase ─────────────────────────────────

    def run_phase_gate(
        self,
        phase: str,
        context: Dict[str, Any]
    ) -> Tuple[bool, List[GateCheckResult]]:
        """
        Chạy tất cả gate tương ứng với phase.
        Returns: (all_passed, list_of_results)
        """
        results = []

        if phase == "CAD_TAKEOFF":
            results.append(self.check_cad_takeoff(context.get("cad_data", {})))

        elif phase == "REBAR_CUT":
            results.append(self.check_rebar_cutting(
                context.get("cutting_result", {}),
                context.get("splice_status", "NOT_RUN"),
                context.get("splice_violations", []),
            ))

        elif phase == "QS_ESTIMATE":
            results.append(self.check_qs_estimate(context.get("qs_data", {})))
            if context.get("excel_path"):
                results.append(self.check_excel_audit(context["excel_path"]))

        all_passed = all(r.passed for r in results)
        return all_passed, results
