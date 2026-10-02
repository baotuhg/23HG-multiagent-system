# -*- coding: utf-8 -*-
"""
AI SUPERVISOR — Chỉ huy trưởng ảo của AEC MultiAgent System
Điều phối State Graph: giao việc cho Sub-Agent, kiểm tra Quality Gate,
xử lý REJECT/RETRY, và kích hoạt Human Gate trước khi xuất tài liệu pháp lý.

Kiến trúc State Graph:
  INIT
    ↓
  CAD_TAKEOFF   ← cad_agent.execute()
    ↓ Quality Gate 1
  REBAR_CUT     ← rebar_agent.execute() → BPTC splice check → REJECT/RETRY
    ↓ Quality Gate 2
  QS_ESTIMATE   ← qs_agent.execute()
    ↓ Quality Gate 3
  QAQC_REVIEW   ← bptc_kcs_agent.execute() (phiếu thí nghiệm) + Excel Audit 100/100 (khi có --excel)
    ↓ Human Gate (AWAITING_APPROVAL)
  SCHEDULE_CPM  ← scheduler_agent.execute()
    ↓
  ASBUILT_LOOP  ← asbuilt_agent.execute() [daily loop]
    ↓
  COMPLETED

Mọi nhánh lỗi → ERROR state → Supervisor quyết định retry hoặc escalate.
"""

from __future__ import annotations
import os
import sys
import time
from datetime import datetime
from typing import Any, Dict, List, Optional, Type

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.state.shared_state import ProjectSharedState, NodeStatus, ProjectPhase
from core.state.state_bus import StateBus
from core.gates.quality_gate import QualityGate
from core.gates.human_gate import HumanGate, ApprovalRequest
from core.supervisor.base_agent import BaseAgent
from aec_core.experience_store import ProjectExperienceStore


class AECSupervisor:
    """
    AI Supervisor (Chỉ huy trưởng ảo) — quản lý toàn bộ State Graph.

    Supervisor KHÔNG tự thực hiện nghiệp vụ kỹ thuật.
    Supervisor CHỈ:
      1. Gọi Sub-Agent đúng thứ tự
      2. Kiểm tra Quality Gate sau mỗi phase
      3. Xử lý REJECT (retry hoặc escalate)
      4. Kích hoạt Human Gate trước tài liệu pháp lý
      5. Cập nhật State Bus
    """

    def __init__(
        self,
        project_root: str = _ROOT,
        excel_master_path: str = "",
        drawings_folder: str = "",
        human_gate_mode: str = "cli",    # "auto" chỉ dùng cho demo/test
        max_retries: int = 3,
        persist_path: Optional[str] = None,
        demo_mode: bool = False,
        project_name: Optional[str] = None,
    ):
        self.project_root = project_root
        self.max_retries = max_retries

        # Khởi tạo State
        state = ProjectSharedState(
            max_retries=max_retries,
            excel_master_path=excel_master_path,
            drawings_folder=drawings_folder,
            demo_mode=demo_mode,
        )
        if project_name:
            state.project_name = project_name
        elif not demo_mode:
            state.project_name = "(chưa đặt tên dự án — dùng --project-name)"
        persist = persist_path or os.path.join(
            os.environ.get("AEC_STATE_DIR") or os.path.join(project_root, ".aec_state"), "RUNTIME_STATE.json")
        self.bus = StateBus(state, persist_path=persist)

        # Gates
        self.quality_gate = QualityGate(project_root)
        self.human_gate = HumanGate(mode=human_gate_mode)

        # Experience Store & Evolution Engine
        self.experience_store = ProjectExperienceStore()

        # Agent registry — Supervisor điền vào trước khi chạy
        self._agents: Dict[str, BaseAgent] = {}

    # ─────────────────────────────────────────────────────────────────────────
    # AGENT REGISTRY
    # ─────────────────────────────────────────────────────────────────────────

    def register_agent(self, agent: BaseAgent) -> None:
        """Đăng ký Sub-Agent vào Supervisor. Gọi trước run()."""
        self._agents[agent.agent_id] = agent
        print(f"  [Supervisor] Đăng ký agent: {agent.agent_id} ({agent.description})")

    # ─────────────────────────────────────────────────────────────────────────
    # MAIN ORCHESTRATION LOOP
    # ─────────────────────────────────────────────────────────────────────────

    def run(self, phases: Optional[List[str]] = None) -> bool:
        """
        Chạy toàn bộ State Graph hoặc chỉ các phase được chỉ định.
        Returns True nếu hoàn thành thành công.
        """
        print("\n" + "═" * 65)
        print("  🏗  AEC SUPERVISOR — STATE GRAPH ORCHESTRATOR")
        print(f"  Dự án: {self.bus._state.project_name}")
        print(f"  Session: {self.bus._state.session_id}")
        if self.bus.is_demo_mode():
            print("  Chế độ : ⚠ DEMO — được phép dùng dữ liệu mẫu, KHÔNG dùng cho hồ sơ thật")
        else:
            print("  Chế độ : DỮ LIỆU THẬT — thiếu dữ liệu sẽ dừng, không thay bằng dữ liệu mẫu")
        print("═" * 65 + "\n")

        # Mặc định chạy tất cả phase
        all_phases = [
            ProjectPhase.CAD_TAKEOFF,
            ProjectPhase.REBAR_CUT,
            ProjectPhase.QS_ESTIMATE,
            ProjectPhase.QAQC_REVIEW,
            ProjectPhase.HUMAN_GATE,
            ProjectPhase.SCHEDULE_CPM,
            ProjectPhase.ASBUILT_LOOP,
            ProjectPhase.PAYMENT_03A,
        ]
        run_phases = phases if phases else all_phases

        for phase in run_phases:
            success = self._run_phase(phase)
            if not success:
                self.bus.set_phase(ProjectPhase.ERROR)
                print(f"\n  ✗ SUPERVISOR: Dừng tại phase {phase} — xem log để biết chi tiết")
                self.bus.print_status_board()
                return False

        self.bus.set_phase(ProjectPhase.COMPLETED)
        self.bus.print_status_board()
        if self.experience_store:
            self.experience_store.record_project_completion(self.bus._state.project_name)
            lvl_info = self.experience_store.get_system_level()
            print(f"\n  ⭐ AEC SYSTEM EVOLUTION: Level {lvl_info['level']} ({lvl_info['rank']}) — Total XP: {lvl_info['total_xp']} XP")
        print("\n  ✅ SUPERVISOR: Hoàn tất toàn bộ State Graph!")
        return True

    # ─────────────────────────────────────────────────────────────────────────
    # PHASE HANDLERS
    # ─────────────────────────────────────────────────────────────────────────

    # Dynamic phase handler registry (decorator-based)
    _PHASE_HANDLERS: Dict[str, str] = {
        ProjectPhase.CAD_TAKEOFF: "_phase_cad_takeoff",
        ProjectPhase.REBAR_CUT: "_phase_rebar_cut",
        ProjectPhase.QS_ESTIMATE: "_phase_qs_estimate",
        ProjectPhase.QAQC_REVIEW: "_phase_qaqc_review",
        ProjectPhase.HUMAN_GATE: "_phase_human_gate",
        ProjectPhase.SCHEDULE_CPM: "_phase_schedule_cpm",
        ProjectPhase.ASBUILT_LOOP: "_phase_asbuilt_loop",
        ProjectPhase.PAYMENT_03A: "_phase_payment",
    }

    @classmethod
    def register_phase_handler(cls, phase: str, handler_name_or_func: Any) -> None:
        """Đăng ký phase handler gọn gàng mà không dùng dynamic import magic."""
        cls._PHASE_HANDLERS[phase] = handler_name_or_func

    def _run_phase(self, phase: str) -> bool:
        """Dispatch sang handler đã đăng ký cho từng phase."""
        print(f"\n{'─' * 60}")
        print(f"  ▶ Phase: {phase}")
        print(f"{'─' * 60}")
        self.bus.set_phase(phase)

        handler = self._PHASE_HANDLERS.get(phase)
        if not handler:
            print(f"  [Supervisor] Phase không xác định: {phase}")
            return False

        if callable(handler):
            return handler(self)
        elif isinstance(handler, str) and hasattr(self, handler):
            return getattr(self, handler)()
        else:
            print(f"  [Supervisor] Handler không hợp lệ cho phase {phase}: {handler}")
            return False

    def _phase_cad_takeoff(self) -> bool:
        success = self._run_agent_with_retry("cad_agent", max_retries=self.max_retries)
        if not success:
            return False

        # Quality Gate 1
        cad_dict = self.bus.get_cad_data().__dict__ if hasattr(self.bus.get_cad_data(), "__dict__") else {}
        passed, results = self.quality_gate.run_phase_gate(
            "CAD_TAKEOFF", {"cad_data": cad_dict}
        )
        self._print_gate_results("GATE-1: CAD Takeoff", results, agent_id="cad_agent")

        if not passed:
            # Một lần retry toàn bộ CAD — kết quả mới phải qua lại Gate-1
            print("  [Supervisor] Gate-1 FAIL → Retry cad_agent...")
            success = self._run_agent_with_retry("cad_agent", max_retries=1)
            if not success:
                return False
            cad_dict = self.bus.get_cad_data().__dict__
            passed, results = self.quality_gate.run_phase_gate(
                "CAD_TAKEOFF", {"cad_data": cad_dict}
            )
            self._print_gate_results("GATE-1: CAD Takeoff (retry)", results, agent_id="cad_agent")
        return passed

    def _phase_rebar_cut(self) -> bool:
        splice_violations: list = []  # Khởi tạo trước vòng lặp để tránh UnboundLocalError
        for attempt in range(1, self.max_retries + 1):
            print(f"  [Supervisor] Rebar attempt {attempt}/{self.max_retries}")
            success = self._run_agent_with_retry("rebar_agent", max_retries=1)
            if not success:
                if self._agent_fatal("rebar_agent"):
                    return False  # Lỗi dữ liệu đầu vào: chạy lại cũng vậy
                continue

            # Inter-agent negotiation: đọc kết quả từ StateBus
            rebar = self.bus.get_rebar_data()
            splice_status = getattr(rebar, "splice_zone_check", "NOT_RUN")
            splice_violations = getattr(rebar, "splice_violations", [])

            # Đọc cutting_dict: dùng get_cutting_dict() lấy dict thuần đã lưu
            cutting_dict = self.bus.get_cutting_dict()
            if not cutting_dict:
                # fallback: thử đọc từ object
                cr = getattr(rebar, "cutting_result", None)
                if cr is not None:
                    cutting_dict = cr.__dict__ if hasattr(cr, "__dict__") else {}

            passed, results = self.quality_gate.run_phase_gate(
                "REBAR_CUT",
                {
                    "cutting_result": cutting_dict,
                    "splice_status": splice_status,
                    "splice_violations": splice_violations,
                }
            )
            self._print_gate_results(f"GATE-2: Rebar Cut (attempt {attempt})", results,
                                     agent_id="rebar_agent")

            if passed:
                return True

            # REJECT → Supervisor gửi thông báo REJECTED về rebar_agent
            self.bus.update_node_status(
                "rebar_agent", NodeStatus.REJECTED,
                error_message=f"Gate-2 REJECT attempt {attempt}: splice violations={len(splice_violations)}"
            )
            print(f"  [Supervisor] REJECTED — rebar_agent sẽ chạy lại (attempt {attempt + 1})")

        # Hết retry → escalate lên Human Gate tự động
        print("  [Supervisor] Rebar FAILED sau tất cả retry — escalate lên Human Gate")
        self.bus.create_approval_gate(
            gate_id="REBAR-ESCALATE",
            gate_name="Xác nhận thủ công kết quả cắt thép (hết retry)",
            clashes=[f"Splice violation {i+1}" for i in range(len(splice_violations))]
        )
        return False

    def _phase_qs_estimate(self) -> bool:
        success = self._run_agent_with_retry("qs_agent", max_retries=self.max_retries)
        if not success:
            return False

        qs = self.bus.get_qs_data()
        qs_dict = qs.__dict__ if hasattr(qs, "__dict__") else {}

        passed, results = self.quality_gate.run_phase_gate(
            "QS_ESTIMATE", {"qs_data": qs_dict}
        )
        self._print_gate_results("GATE-3: QS Estimate", results, agent_id="qs_agent")
        return passed

    def _phase_qaqc_review(self) -> bool:
        success = self._run_agent_with_retry("bptc_kcs_agent", max_retries=self.max_retries)
        if not success:
            return False

        # Excel Audit (GATE-4) — chỉ khi người dùng chỉ định Excel master (--excel)
        excel_path = self.bus._state.excel_master_path
        if not excel_path:
            print("  [Supervisor] ℹ Chưa chỉ định Excel master (--excel) — bỏ qua GATE-4 kiểm toán Excel; "
                  "Human Gate sẽ ghi 'chưa kiểm toán'")
            return True
        if not os.path.exists(excel_path):
            print(f"  [Supervisor] ✗ Không tìm thấy Excel master {excel_path} — không thể kiểm toán, dừng phase")
            return False
        audit_result = self.quality_gate.check_excel_audit(excel_path)
        self.bus.set_qaqc_data({"audit_score": audit_result.score, "audit_run": True})
        self._print_gate_results("GATE-4: Excel Audit 100/100", [audit_result],
                                 agent_id="bptc_kcs_agent")
        return audit_result.passed

    def _phase_human_gate(self) -> bool:
        """
        Cổng phê duyệt Human-in-the-loop trước khi xuất tài liệu pháp lý.
        Hiển thị tất cả clash và anomaly để KS trưởng xem xét.
        """
        qaqc = self.bus.get_qaqc_data()
        clashes = getattr(qaqc, "clashes_detected", [])
        errors = self.bus.get_errors()

        qs = self.bus.get_qs_data()
        cad = self.bus.get_cad_data()
        sample_sources = self.bus.get_sample_data_sources()
        summary = {
            "Nguồn dữ liệu": (
                "⚠ CÓ DỮ LIỆU MẪU (demo) — không phải số liệu dự án: "
                + ", ".join(dict.fromkeys(s["agent_id"] for s in sample_sources))
                if sample_sources else "Dữ liệu thật"
            ),
            "Đo bóc khối lượng": (
                "{n} khối lượng từ {f} — hồ sơ quy tắc: {p} ({v})".format(
                    n=len(cad.takeoff_quantities), f=os.path.basename(cad.source_takeoff_file),
                    p=cad.takeoff_profile.get("name", "?"),
                    v="đã đối chiếu bản gốc" if cad.takeoff_profile.get("verified") else "CHƯA đối chiếu bản gốc")
                if getattr(cad, "takeoff_quantities", None) else "Từ bản vẽ CAD (không dùng bảng cấu kiện)"),
            "G_XD (VNĐ)": f"{getattr(qs, 'total_G_XD_vnd', 0):,.0f}",
            "VAT 10%": f"{getattr(qs, 'vat_vnd', 0):,.0f}",
            "Audit score": (f"{getattr(qaqc, 'audit_score', 0)}/100" if getattr(qaqc, "audit_run", False)
                            else "Chưa kiểm toán Excel master (không có --excel)"),
            "Phiếu thí nghiệm": (
                "{total} phiếu — {pass} đạt / {fail} không đạt / {pending} chờ".format(**qaqc.lab_summary)
                if getattr(qaqc, "lab_summary", None) else "Chưa đánh giá"),
            "Số lỗi hệ thống": len(errors),
        }

        req = ApprovalRequest(
            gate_id="LEGAL-DOCS-APPROVAL",
            gate_name="Phê duyệt xuất tài liệu pháp lý",
            phase=ProjectPhase.HUMAN_GATE,
            document_ref=self.bus._state.excel_master_path,
            clashes=clashes,
            anomalies=errors[:5],
            summary_data=summary,
        )

        self.bus.create_approval_gate(
            gate_id=req.gate_id,
            gate_name=req.gate_name,
            document_ref=req.document_ref,
            clashes=clashes,
        )
        self.bus.update_node_status("supervisor", NodeStatus.AWAITING,
                                    output_summary="Chờ phê duyệt KS trưởng")

        decision = self.human_gate.request_approval(req)

        self.bus.resolve_approval_gate(
            gate_id=req.gate_id,
            approved=decision.approved,
            approver=decision.approver,
            comments=decision.comments,
        )

        if decision.approved:
            print(f"\n  ✅ Human Gate APPROVED bởi {decision.approver}: {decision.comments}")
            self.bus.update_node_status("supervisor", NodeStatus.RUNNING,
                                        output_summary="Tiếp tục sau approval")
            return True
        else:
            print(f"\n  ✗ Human Gate REJECTED bởi {decision.approver}: {decision.comments}")
            return False

    def _phase_schedule_cpm(self) -> bool:
        return self._run_agent_with_retry("scheduler_agent", max_retries=self.max_retries)

    def _phase_asbuilt_loop(self) -> bool:
        if "asbuilt_agent" not in self._agents:
            print("  [Supervisor] asbuilt_agent chưa được đăng ký — bỏ qua As-Built loop")
            return True
        return self._run_agent_with_retry("asbuilt_agent", max_retries=1)

    def _phase_payment(self) -> bool:
        if "payment_agent" not in self._agents:
            print("  [Supervisor] payment_agent chưa được đăng ký — bỏ qua Mẫu 03a")
            return True
        if not self._run_agent_with_retry("payment_agent", max_retries=1):
            return False
        summary = self.bus.get_qs_data().payment_summary
        passed, results = self.quality_gate.run_phase_gate("PAYMENT_03A", {"payment": summary})
        self._print_gate_results("GATE-5: Mẫu 03a", results, agent_id="payment_agent")
        return passed

    # ─────────────────────────────────────────────────────────────────────────
    # HELPERS
    # ─────────────────────────────────────────────────────────────────────────

    def _run_agent_with_retry(self, agent_id: str, max_retries: int = 3) -> bool:
        """Chạy agent với retry logic. Trả về True nếu thành công."""
        agent = self._agents.get(agent_id)
        if agent is None:
            print(f"  [Supervisor] WARN: Agent '{agent_id}' chưa được đăng ký — bỏ qua")
            self.bus.update_node_status(agent_id, NodeStatus.SKIPPED,
                                        output_summary="Agent chưa đăng ký")
            return True  # Non-fatal: cho phép pipeline tiếp tục

        for attempt in range(1, max_retries + 1):
            print(f"\n  → Gọi {agent_id} (lần {attempt}/{max_retries})")
            success = agent.execute(self.bus)
            if success:
                return True
            if agent.last_error_fatal:
                print(f"  [Supervisor] {agent_id}: lỗi dữ liệu đầu vào — không retry")
                return False

            status = self.bus.get_node_status(agent_id)
            if status == NodeStatus.FAILED and attempt < max_retries:
                wait_s = 2 ** attempt  # Exponential backoff
                print(f"  [Supervisor] {agent_id} FAILED — retry sau {wait_s}s...")
                time.sleep(wait_s)

        print(f"  [Supervisor] {agent_id} FAILED sau {max_retries} lần thử")
        return False

    def _agent_fatal(self, agent_id: str) -> bool:
        agent = self._agents.get(agent_id)
        return bool(agent is not None and agent.last_error_fatal)

    def _print_gate_results(self, gate_label: str, results: list, agent_id: Optional[str] = None) -> None:
        print(f"\n  {'─'*50}")
        if agent_id and self.bus.uses_sample_data(agent_id):
            gate_label += "  [⚠ DỮ LIỆU MẪU — điểm số không có giá trị cho hồ sơ thật]"
        print(f"  🔍 {gate_label}")
        for r in results:
            icon = "✓" if r.passed else "✗"
            print(f"  {icon} {r.gate_name}: {r.score}/{r.max_score}")
            for issue in r.issues:
                print(f"      ✗ {issue}")
            for warn in r.warnings:
                print(f"      ⚠ {warn}")
        print(f"  {'─'*50}")

    def get_status_report(self) -> Dict[str, Any]:
        """Trả về báo cáo trạng thái toàn hệ thống."""
        snap = self.bus.get_state_snapshot()
        return {
            "session_id": snap.get("session_id"),
            "phase": snap.get("current_phase"),
            "updated_at": snap.get("updated_at"),
            "node_status": snap.get("node_status", {}),
            "errors_count": len(snap.get("global_errors", [])),
            "pending_approvals": len(self.bus.get_pending_gates()),
            "demo_mode": snap.get("demo_mode", False),
            "sample_data_sources": snap.get("sample_data_sources", []),
            "system_level": self.experience_store.get_system_level() if self.experience_store else None,
        }
