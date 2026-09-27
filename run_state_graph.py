# -*- coding: utf-8 -*-
"""
RUN STATE GRAPH — Entry Point cho 23HG MultiAgent System v3.0
Kiến trúc: State Graph + Supervisor Pattern (thay thế Linear Pipeline cũ)

Dùng lệnh:
  python run_state_graph.py --demo                        # Chạy thử toàn bộ bằng dữ liệu mẫu
  python run_state_graph.py --phase rebar --bbs BBS.xlsx  # Tối ưu cắt thép từ BBS thật
  python run_state_graph.py --solver-test                 # Chỉ test OR-Tools solver

Dữ liệu thật vs dữ liệu mẫu:
  Mặc định hệ thống CHỈ dùng dữ liệu thật: phase nào thiếu dữ liệu sẽ dừng và báo rõ
  cần cung cấp gì. Dữ liệu mẫu (Cầu Km19+529.080) chỉ được dùng khi có cờ --demo,
  và mọi chỗ dùng dữ liệu mẫu đều được đánh dấu trong log, Quality Gate và Human Gate.

Backward compatibility:
  Pipeline cũ (run_pipeline.py) vẫn hoạt động bình thường.
  Script này chạy SONG SONG — không thay thế file cũ.
"""

import argparse
import os
import sys

# Thêm project root vào sys.path
ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.supervisor.supervisor_agent import AECSupervisor
from core.state.shared_state import ProjectPhase


# ── CONFIG ───────────────────────────────────────────────────────────────────

SAMPLE_EXCEL_MASTER = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")
RUNTIME_STATE = os.path.join(ROOT, "agents", "RUNTIME_STATE.json")


PHASE_MAP = {
    "cad":      ProjectPhase.CAD_TAKEOFF,
    "rebar":    ProjectPhase.REBAR_CUT,
    "qs":       ProjectPhase.QS_ESTIMATE,
    "qaqc":     ProjectPhase.QAQC_REVIEW,
    "gate":     ProjectPhase.HUMAN_GATE,
    "schedule": ProjectPhase.SCHEDULE_CPM,
    "asbuilt":  ProjectPhase.ASBUILT_LOOP,
}


# ── MAIN ─────────────────────────────────────────────────────────────────────

def run_solver_test():
    """Quick test OR-Tools + CPM Calculator (dữ liệu thử nghiệm cố định)."""
    print("\n" + "═" * 55)
    print("  TEST: OR-Tools Cutting Stock Solver (dữ liệu thử nghiệm)")
    print("═" * 55)

    from tools.cutting_stock_solver import CuttingStockSolver, CutDemand
    demands = [
        CutDemand(length_mm=4500, quantity=30, diameter_mm=20, mark="T1"),
        CutDemand(length_mm=3200, quantity=50, diameter_mm=20, mark="T2"),
        CutDemand(length_mm=2800, quantity=20, diameter_mm=16, mark="D1"),
    ]
    sol = CuttingStockSolver().solve(demands)
    print(f"  Status: {sol.status} ({sol.solver_name})")
    print(f"  Số cây: {sol.total_bars_needed} (cận dưới {sol.lower_bound_bars})")
    print(f"  Đề-xê: {sol.waste_ratio_pct:.2f}%")
    for g in sol.groups:
        print(f"    Ø{g.diameter_mm}: {g.total_bars_needed} cây, đề-xê {g.waste_ratio_pct:.2f}% [{g.status}]")
    for w in sol.warnings:
        print(f"  ⚠ {w}")

    print("\n" + "═" * 55)
    print("  TEST: CPM Calculator")
    print("═" * 55)

    from tools.cpm_calculator import CPMCalculator
    tasks = [
        {"id": "T01", "name": "Tim mốc", "duration": 3, "predecessors": []},
        {"id": "T02", "name": "Cọc nhồi", "duration": 30, "predecessors": ["T01"]},
        {"id": "T03", "name": "Bệ trụ", "duration": 20, "predecessors": ["T02"]},
        {"id": "T04", "name": "Dầm Super-T", "duration": 10, "predecessors": ["T03"]},
        {"id": "T05", "name": "Thử tải", "duration": 5, "predecessors": ["T04"]},
    ]
    result = CPMCalculator().calculate(tasks, start_date_str="2026-10-01")
    print(f"  Tổng thời gian: {result.total_duration_days} ngày")
    print(f"  Hoàn thành: {result.project_finish}")
    print(f"  Đường găng: {' → '.join(result.critical_path)}")
    print("\n  ✅ Tất cả tools hoạt động bình thường!\n")


def main():
    parser = argparse.ArgumentParser(
        description="23HG MultiAgent System v3.0 — State Graph + Supervisor"
    )
    parser.add_argument(
        "--phase", nargs="*",
        choices=list(PHASE_MAP.keys()),
        help="Chỉ chạy các phase được chỉ định (mặc định: tất cả)"
    )
    parser.add_argument(
        "--demo", action="store_true",
        help="Cho phép dùng dữ liệu mẫu Cầu Km19+529.080 khi thiếu dữ liệu thật "
             "(chỉ để chạy thử — KHÔNG dùng cho hồ sơ thật)"
    )
    parser.add_argument(
        "--human-gate", choices=["cli", "auto", "file"],
        default=None,
        help="Chế độ Human Gate (mặc định: cli; auto chỉ là mặc định khi --demo)"
    )
    parser.add_argument(
        "--solver-test", action="store_true",
        help="Chỉ test OR-Tools và CPM Calculator"
    )
    parser.add_argument(
        "--excel", default=None,
        help="Đường dẫn file Excel master (mặc định khi --demo: workbook mẫu trong templates/)"
    )
    parser.add_argument(
        "--drawings", default="",
        help="Thư mục chứa bản vẽ DWG/DXF"
    )
    parser.add_argument(
        "--bbs", default=None,
        help="File BBS thật (.xlsx/.csv/.json) cho phase cắt thép"
    )
    parser.add_argument(
        "--bbs-sheet", default=None,
        help="Tên sheet BBS trong file Excel (mặc định: tự tìm)"
    )
    parser.add_argument(
        "--bbs-skip-invalid", action="store_true",
        help="Loại các dòng BBS sai dữ liệu (được liệt kê trong cảnh báo) thay vì dừng"
    )
    parser.add_argument(
        "--cut-plan-out", default=None,
        help="Xuất phiếu cắt thép cho xưởng ra file CSV"
    )
    parser.add_argument(
        "--project-name", default=None,
        help="Tên dự án hiển thị trong báo cáo"
    )

    args = parser.parse_args()

    if args.solver_test:
        run_solver_test()
        return

    human_gate_mode = args.human_gate or ("auto" if args.demo else "cli")
    if human_gate_mode == "auto" and not args.demo:
        print("  ⚠ Human Gate 'auto' tự phê duyệt hồ sơ — không dùng cho hồ sơ thật.")
    excel_path = args.excel if args.excel is not None else (SAMPLE_EXCEL_MASTER if args.demo else "")

    # Khởi tạo Supervisor
    supervisor = AECSupervisor(
        project_root=ROOT,
        excel_master_path=excel_path,
        drawings_folder=args.drawings,
        human_gate_mode=human_gate_mode,
        max_retries=3,
        persist_path=RUNTIME_STATE,
        demo_mode=args.demo,
        project_name=args.project_name,
    )

    # Đăng ký tất cả Sub-Agent
    from core.agents.sub_agents import (
        CADAgent, QSAgent, BPTCKCSAgent, SchedulerAgent
    )
    from core.agents.rebar_agent import RebarAgent
    from core.agents.asbuilt_agent import AsBuiltAgent

    supervisor.register_agent(CADAgent(drawings_folder=args.drawings))
    supervisor.register_agent(RebarAgent(
        bbs_path=args.bbs,
        bbs_sheet=args.bbs_sheet,
        skip_invalid_rows=args.bbs_skip_invalid,
        cut_plan_out=args.cut_plan_out,
    ))
    supervisor.register_agent(QSAgent())
    supervisor.register_agent(BPTCKCSAgent())
    supervisor.register_agent(SchedulerAgent())
    supervisor.register_agent(AsBuiltAgent())

    # Chọn phases
    selected_phases = None
    if args.phase:
        selected_phases = [PHASE_MAP[p] for p in args.phase]

    # Chạy State Graph
    success = supervisor.run(phases=selected_phases)

    # In báo cáo cuối
    report = supervisor.get_status_report()
    print(f"\n  📋 Báo cáo cuối:")
    print(f"     Session   : {report['session_id']}")
    print(f"     Phase     : {report['phase']}")
    print(f"     Cập nhật  : {report['updated_at']}")
    print(f"     Lỗi       : {report['errors_count']}")
    print(f"     Chờ duyệt : {report['pending_approvals']}")
    if report["sample_data_sources"]:
        print("\n  ⚠ KẾT QUẢ CÓ DÙNG DỮ LIỆU MẪU — KHÔNG DÙNG CHO HỒ SƠ THẬT:")
        for src in report["sample_data_sources"]:
            print(f"     • {src['agent_id']}: {src['note']}")
    if not success:
        for err in supervisor.bus.get_errors()[-3:]:
            print(f"\n  ✗ {err}")
        if not args.demo:
            print("\n  Gợi ý: chạy thử toàn bộ bằng dữ liệu mẫu:  python run_state_graph.py --demo")
            print("         tối ưu cắt thép từ BBS thật:         python run_state_graph.py --phase rebar --bbs <file>")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
