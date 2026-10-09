# -*- coding: utf-8 -*-
"""
STATE BUS — Thread-safe Read/Write operations trên ProjectSharedState
Đây là lớp trung gian giữa Supervisor và các Sub-Agent.
Không Agent nào được ghi thẳng vào SharedState — phải qua StateBus.

Pattern: StateBus.update_agent(agent_id, status, payload=...)
"""

from __future__ import annotations
import json
import os
import threading
from datetime import datetime, timezone
from typing import Any, Dict, Optional

from core.state.shared_state import (
    ProjectSharedState, NodeStatus, ProjectPhase,
    AgentRunRecord, ApprovalGate, ApprovalStatus
)


class StateBus:
    """
    Shared State Bus — Quản lý đọc/ghi thread-safe lên ProjectSharedState.
    Supervisor là đối tượng DUY NHẤT khởi tạo StateBus.
    Sub-Agent nhận tham chiếu `bus` từ Supervisor khi được gọi.
    """

    def __init__(self, state: ProjectSharedState, persist_path: Optional[str] = None):
        self._legal_exports = []
        self._state = state
        self._lock = threading.RLock()
        self._persist_path = persist_path  # JSON file để lưu trạng thái liên phiên

        # Khởi tạo session_id và timestamp nếu chưa có
        if not self._state.session_id:
            import uuid
            self._state.session_id = str(uuid.uuid4())[:8]
        if not self._state.created_at:
            self._state.created_at = _now()
        self._state.updated_at = _now()

    # ─────────────────────────────────────────────────────────────────────────
    # READ operations (thread-safe)
    # ─────────────────────────────────────────────────────────────────────────

    def defer_legal_export(self, path, writer) -> None:
        """Queue an in-memory result; no destination is written before approval."""
        with self._lock:
            self._legal_exports.append((os.path.abspath(path), writer))

    def pending_legal_exports(self):
        with self._lock:
            return [path for path, _ in self._legal_exports]

    def discard_legal_exports(self):
        with self._lock:
            self._legal_exports.clear()

    def publish_legal_exports(self):
        """Only the supervisor calls this after approval and successful phases."""
        import tempfile
        staged = []
        try:
            for path, writer in self._legal_exports:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(path))
                os.close(fd)
                staged.append((tmp, path))
                writer(tmp)
            for tmp, path in staged:
                os.replace(tmp, path)
        finally:
            for tmp, _ in staged:
                if os.path.exists(tmp):
                    os.unlink(tmp)
            self.discard_legal_exports()

    def get_phase(self) -> str:
        with self._lock:
            return self._state.current_phase

    def get_node_status(self, agent_id: str) -> str:
        with self._lock:
            return self._state.node_status.get(agent_id, NodeStatus.IDLE)

    def get_state_snapshot(self) -> Dict[str, Any]:
        """Trả về snapshot toàn bộ state (copy) để Agent đọc context."""
        with self._lock:
            return self._state.to_dict()

    def get_cad_data(self):
        with self._lock:
            return self._state.cad_data

    def get_rebar_data(self):
        with self._lock:
            return self._state.rebar_data

    def get_qs_data(self):
        with self._lock:
            return self._state.qs_data

    def get_qaqc_data(self):
        with self._lock:
            return self._state.qaqc_data

    def get_schedule_data(self):
        with self._lock:
            return self._state.schedule_data

    def get_errors(self):
        with self._lock:
            return list(self._state.global_errors)

    def is_demo_mode(self) -> bool:
        with self._lock:
            return bool(self._state.demo_mode)

    def get_sample_data_sources(self) -> list:
        with self._lock:
            return list(self._state.sample_data_sources)

    def uses_sample_data(self, agent_id: str) -> bool:
        with self._lock:
            return any(s.get("agent_id") == agent_id for s in self._state.sample_data_sources)

    # ─────────────────────────────────────────────────────────────────────────
    # WRITE operations (thread-safe, chỉ Supervisor/Agent được phép gọi)
    # ─────────────────────────────────────────────────────────────────────────

    def set_phase(self, phase: ProjectPhase) -> None:
        with self._lock:
            old_phase = self._state.current_phase
            self._state.current_phase = phase
            self._state.updated_at = _now()
            print(f"  [StateBus] Phase: {old_phase} → {phase}")

    def update_node_status(
        self,
        agent_id: str,
        status: NodeStatus,
        output_summary: str = "",
        error_message: str = ""
    ) -> None:
        """Cập nhật trạng thái node và ghi vào run_history."""
        with self._lock:
            self._state.node_status[agent_id] = status
            self._state.updated_at = _now()

            # Tìm record gần nhất của agent này trong run_history
            existing = [r for r in self._state.run_history if r.get("agent_id") == agent_id]
            if existing and existing[-1].get("status") == NodeStatus.RUNNING:
                existing[-1]["status"] = status
                existing[-1]["finished_at"] = _now()
                existing[-1]["output_summary"] = output_summary
                existing[-1]["error_message"] = error_message
            else:
                record = {
                    "agent_id": agent_id,
                    "status": status,
                    "started_at": _now(),
                    "finished_at": _now() if status != NodeStatus.RUNNING else "",
                    "attempt": len([r for r in self._state.run_history if r.get("agent_id") == agent_id]) + 1,
                    "error_message": error_message,
                    "output_summary": output_summary,
                }
                self._state.run_history.append(record)

            print(f"  [StateBus] {agent_id}: {status} — {output_summary or error_message or '...'}")
        self._persist()

    def set_cad_data(self, data_dict: Dict[str, Any]) -> None:
        with self._lock:
            for k, v in data_dict.items():
                if hasattr(self._state.cad_data, k):
                    setattr(self._state.cad_data, k, v)
            self._state.updated_at = _now()
        self._persist()

    def set_rebar_data(self, data_dict: Dict[str, Any]) -> None:
        with self._lock:
            for k, v in data_dict.items():
                if hasattr(self._state.rebar_data, k):
                    setattr(self._state.rebar_data, k, v)
                # Lưu extra keys (không có trong dataclass) vào dict phụ trên state
            # Lưu cutting dict (dict thuần) vào _state để Supervisor đọc được
            if "_cutting_dict_json" in data_dict:
                self._state.__dict__.setdefault("_extra", {})
                self._state.__dict__["_extra"]["cutting_dict"] = data_dict["_cutting_dict_json"]
            self._state.updated_at = _now()
        self._persist()

    def get_cutting_dict(self) -> Dict[str, Any]:
        """Trả về cutting dict thuần (JSON-serializable) từ lần chạy Rebar Agent gần nhất."""
        with self._lock:
            return self._state.__dict__.get("_extra", {}).get("cutting_dict", {})

    def set_qs_data(self, data_dict: Dict[str, Any]) -> None:
        with self._lock:
            for k, v in data_dict.items():
                if hasattr(self._state.qs_data, k):
                    setattr(self._state.qs_data, k, v)
            self._state.updated_at = _now()
        self._persist()

    def set_qaqc_data(self, data_dict: Dict[str, Any]) -> None:
        with self._lock:
            for k, v in data_dict.items():
                if hasattr(self._state.qaqc_data, k):
                    setattr(self._state.qaqc_data, k, v)
            self._state.updated_at = _now()
        self._persist()

    def set_schedule_data(self, data_dict: Dict[str, Any]) -> None:
        with self._lock:
            for k, v in data_dict.items():
                if hasattr(self._state.schedule_data, k):
                    setattr(self._state.schedule_data, k, v)
            self._state.updated_at = _now()
        self._persist()

    def mark_sample_data(self, agent_id: str, note: str) -> None:
        """Ghi nhận agent đã dùng dữ liệu mẫu (chỉ hợp lệ ở chế độ demo)."""
        with self._lock:
            if not self._state.demo_mode:
                raise RuntimeError(f"{agent_id} dùng dữ liệu mẫu khi không bật chế độ demo")
            entry = {"agent_id": agent_id, "note": note}
            if entry not in self._state.sample_data_sources:
                self._state.sample_data_sources.append(entry)
            self._state.updated_at = _now()
        print(f"  [StateBus] ⚠ DỮ LIỆU MẪU — {agent_id}: {note}")
        self._persist()

    def push_error(self, error_msg: str) -> None:
        with self._lock:
            self._state.global_errors.append(f"[{_now()}] {error_msg}")
            self._state.updated_at = _now()

    def push_site_log(self, log_dict: Dict[str, Any]) -> None:
        with self._lock:
            self._state.site_logs.append(log_dict)
            self._state.updated_at = _now()
        self._persist()

    def create_approval_gate(
        self, gate_id: str, gate_name: str,
        document_ref: str = "", clashes: list = None
    ) -> None:
        """Tạo cổng phê duyệt — đưa node vào trạng thái AWAITING_APPROVAL."""
        with self._lock:
            gate = {
                "gate_id": gate_id,
                "gate_name": gate_name,
                "status": ApprovalStatus.PENDING,
                "requested_at": _now(),
                "resolved_at": "",
                "approver": "",
                "comments": "",
                "document_ref": document_ref,
                "clashes_to_review": clashes or [],
            }
            self._state.approval_gates.append(gate)
            self._state.updated_at = _now()
        self._persist()
        print(f"\n  ╔══════════════════════════════════════╗")
        print(f"  ║  ⏳ HUMAN GATE: {gate_name[:38]}")
        print(f"  ║  Tài liệu: {document_ref}")
        print(f"  ║  Trạng thái: AWAITING_APPROVAL")
        print(f"  ╚══════════════════════════════════════╝\n")

    def resolve_approval_gate(
        self, gate_id: str, approved: bool,
        approver: str = "KS_TRUONG", comments: str = ""
    ) -> None:
        with self._lock:
            for gate in self._state.approval_gates:
                if gate.get("gate_id") == gate_id:
                    gate["status"] = ApprovalStatus.APPROVED if approved else ApprovalStatus.REJECTED
                    gate["resolved_at"] = _now()
                    gate["approver"] = approver
                    gate["comments"] = comments
                    break
            self._state.updated_at = _now()
        self._persist()

    def get_pending_gates(self) -> list:
        with self._lock:
            return [g for g in self._state.approval_gates if g.get("status") == ApprovalStatus.PENDING]

    # ─────────────────────────────────────────────────────────────────────────
    # PERSIST / RESTORE
    # ─────────────────────────────────────────────────────────────────────────

    def _persist(self) -> None:
        """Ghi state ra JSON file để khôi phục sau khi crash."""
        if not self._persist_path:
            return
        with self._lock:
            try:
                target = os.path.abspath(self._persist_path)
                os.makedirs(os.path.dirname(target), exist_ok=True)
                # Ghi ra file tạm rồi đổi tên: nếu tiến trình chết giữa chừng, file cũ vẫn nguyên vẹn
                tmp = f"{target}.{os.getpid()}.tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(self._state.to_dict(), f, ensure_ascii=False, indent=2,
                              default=lambda o: str(o))  # Fallback cho non-serializable objects
                os.replace(tmp, target)
            except Exception as e:
                print(f"  [StateBus] WARN: Không thể persist state: {e}")

    @classmethod
    def load_from_file(cls, persist_path: str) -> "StateBus":
        """Khôi phục StateBus từ file JSON (resume sau crash)."""
        if os.path.exists(persist_path):
            with open(persist_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            state = ProjectSharedState.from_dict(data)
            print(f"  [StateBus] Đã khôi phục state từ {persist_path} (session: {state.session_id})")
        else:
            state = ProjectSharedState()
        return cls(state, persist_path=persist_path)

    def print_status_board(self) -> None:
        """In bảng trạng thái tất cả các node — dùng cho debugging."""
        with self._lock:
            print("\n" + "─" * 58)
            print(f"  📊 AEC STATE GRAPH — Phase: {self._state.current_phase}")
            print("─" * 58)
            icons = {
                NodeStatus.IDLE: "○", NodeStatus.RUNNING: "▶",
                NodeStatus.SUCCESS: "✓", NodeStatus.FAILED: "✗",
                NodeStatus.REJECTED: "⊘", NodeStatus.AWAITING: "⏳",
                NodeStatus.APPROVED: "✅", NodeStatus.SKIPPED: "—",
            }
            for agent_id, status in self._state.node_status.items():
                icon = icons.get(status, "?")
                print(f"  {icon}  {agent_id:<22} {status}")
            print("─" * 58)
            if self._state.global_errors:
                print(f"  ⚠ Errors: {len(self._state.global_errors)}")
            pending = self.get_pending_gates()
            if pending:
                print(f"  ⏳ Pending approvals: {len(pending)}")
            print()


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────────────────────────────────────

def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")
