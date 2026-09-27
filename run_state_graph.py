# -*- coding: utf-8 -*-
"""
RUN STATE GRAPH — Entry Point cho 23HG MultiAgent System v3.0
Kiến trúc: State Graph + Supervisor Pattern (thay thế Linear Pipeline cũ)

Dùng lệnh:
  python run_state_graph.py --demo                        # Chạy thử toàn bộ bằng dữ liệu mẫu
  python run_state_graph.py --phase rebar --bbs BBS.xlsx  # Tối ưu cắt thép từ BBS thật
  python run_state_graph.py --phase schedule --schedule TienDo.xml --non-working-days cn
                                                          # Tính CPM từ MS Project XML / Excel thật
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
from datetime import date, timedelta

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


WEEKDAYS = {
    "t2": 0, "mon": 0, "t3": 1, "tue": 1, "t4": 2, "wed": 2, "t5": 3, "thu": 3,
    "t6": 4, "fri": 4, "t7": 5, "sat": 5, "cn": 6, "sun": 6,
}


def parse_weekdays(text: str) -> set:
    """'cn' / 't7,cn' / 'sat,sun' → {5, 6}."""
    days = set()
    for part in filter(None, (p.strip().lower() for p in (text or "").split(","))):
        if part not in WEEKDAYS:
            raise argparse.ArgumentTypeError(f"Ngày nghỉ không hợp lệ '{part}' (dùng t2..t7, cn hoặc mon..sun)")
        days.add(WEEKDAYS[part])
    return days


def parse_holidays(text: str) -> set:
    """'2027-02-05:2027-02-12,2027-04-30' → tập ngày lễ."""
    days = set()
    for part in filter(None, (p.strip() for p in (text or "").split(","))):
        try:
            first, _, last = part.partition(":")
            d, end = date.fromisoformat(first), date.fromisoformat(last or first)
        except ValueError:
            raise argparse.ArgumentTypeError(f"Ngày lễ không hợp lệ '{part}' (dùng YYYY-MM-DD hoặc YYYY-MM-DD:YYYY-MM-DD)")
        while d <= end:
            days.add(d)
            d += timedelta(days=1)
    return days


def parse_splice_zone(text: str):
    """'0-0.25; 0.75-1' hoặc '0-25%; 75-100%' → [(0, 0.25), (0.75, 1)]."""
    from tools.bbs_loader import _parse_zones
    if "m" in text.lower():
        raise argparse.ArgumentTypeError("--splice-zone dùng tỷ lệ (0-0.25) hoặc % (0-25%), không dùng mét")
    try:
        zones = _parse_zones(text, 1000)
    except ValueError as e:
        raise argparse.ArgumentTypeError(str(e))
    if not zones:
        raise argparse.ArgumentTypeError(f"Không đọc được vùng nối '{text}'")
    return zones


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
    rebar = parser.add_argument_group("Cắt thép (phase rebar)")
    rebar.add_argument("--kerf-mm", type=int, default=3, help="Hao hụt mỗi nhát cắt (mm, mặc định 3)")
    rebar.add_argument("--end-trim-mm", type=int, default=0, help="Cắt bỏ mỗi đầu cây (mm, mặc định 0)")
    rebar.add_argument("--max-pieces-per-bar", type=int, default=None,
                       help="Giới hạn cho tổ cắt: tối đa số đoạn / cây")
    rebar.add_argument("--max-marks-per-bar", type=int, default=None,
                       help="Giới hạn cho tổ cắt: tối đa số Bar Mark / cây")
    rebar.add_argument("--reuse-xd", type=float, default=100, help="Đầu thừa ≥ xD được tái sử dụng (mặc định 100)")
    rebar.add_argument("--short-offcut-xd", type=float, default=20,
                       help="Đầu thừa ≥ xD là đầu thừa ngắn, ngắn hơn là phế (mặc định 20)")
    rebar.add_argument("--splice", action="store_true",
                       help="Tính thêm phương án nối thép tận dụng đầu thừa (đề xuất, cần kỹ thuật duyệt)")
    rebar.add_argument("--lap-xd", type=float, default=40, help="Chiều dài nối chồng mặc định xD (mặc định 40)")
    rebar.add_argument("--max-splice-ratio", type=float, default=0.5,
                       help="Tỷ lệ thanh được nối tối đa mỗi Bar Mark (mặc định 0.5)")
    rebar.add_argument("--min-splice-segment-xd", type=float, default=20,
                       help="Đoạn nối tối thiểu mỗi phía xD (mặc định 20)")
    rebar.add_argument("--splice-zone", type=parse_splice_zone, default=None,
                       help="Vùng cho phép nối mặc định cho mọi Bar Mark 'Cho nối = Có' chưa ghi vùng, "
                            "theo tỷ lệ chiều dài thanh, vd '0-0.25; 0.75-1'")
    rebar.add_argument("--rebarcut-out", default=None,
                       help="Xuất kết quả theo bố cục RebarCut Pro Excel (.xlsx)")
    qs = parser.add_argument_group("Dự toán G_XD (phase qs)")
    qs.add_argument("--qs", default=None, help="Bảng QS / BOQ thật (.xlsx/.csv/.json): khối lượng × đơn giá")
    qs.add_argument("--qs-sheet", default=None, help="Tên sheet QS trong file Excel (mặc định: tự tìm)")
    qs.add_argument("--rate-chung", type=float, default=None, help="Chi phí chung, %% của T (vd 5.1)")
    qs.add_argument("--rate-nha-tam", type=float, default=None, help="Chi phí nhà tạm, %% của T (vd 1.2)")
    qs.add_argument("--rate-kxd", type=float, default=None,
                    help="Chi phí công việc không xác định được KL, %% của T (vd 1.0)")
    qs.add_argument("--rate-tl", type=float, default=None,
                    help="Thu nhập chịu thuế tính trước, %% của (T+GT) (vd 5.5)")
    qs.add_argument("--vat", type=float, default=None, help="Thuế suất VAT, %% của G (vd 10 hoặc 8)")
    qs.add_argument("--qs-out", default=None, help="Xuất bảng tổng hợp G_XD + chi tiết công tác (.xlsx)")
    parser.add_argument(
        "--project-name", default=None,
        help="Tên dự án hiển thị trong báo cáo"
    )
    parser.add_argument(
        "--schedule", default=None,
        help="File tiến độ thật: MS Project XML (.xml), Excel (.xlsx), CSV hoặc JSON"
    )
    parser.add_argument(
        "--schedule-sheet", default=None,
        help="Tên sheet tiến độ trong file Excel (mặc định: tự tìm)"
    )
    parser.add_argument(
        "--start-date", default=None,
        help="Ngày khởi công YYYY-MM-DD (mặc định: lấy từ file tiến độ)"
    )
    parser.add_argument(
        "--non-working-days", type=parse_weekdays, default=set(),
        help="Thứ nghỉ trong tuần, vd 'cn' hoặc 't7,cn' (mặc định: làm cả tuần)"
    )
    parser.add_argument(
        "--holidays", type=parse_holidays, default=set(),
        help="Ngày nghỉ lễ, vd '2027-02-05:2027-02-12,2027-04-30'"
    )
    parser.add_argument(
        "--schedule-out", default=None,
        help="Xuất bảng tiến độ CPM (ES/EF/LS/LF/dự trữ/ngày) ra file CSV"
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
        kerf_mm=args.kerf_mm,
        end_trim_mm=args.end_trim_mm,
        max_pieces_per_bar=args.max_pieces_per_bar,
        max_marks_per_bar=args.max_marks_per_bar,
        reuse_xd=args.reuse_xd,
        short_xd=args.short_offcut_xd,
        splice=args.splice,
        lap_xd=args.lap_xd,
        max_splice_ratio=args.max_splice_ratio,
        min_splice_segment_xd=args.min_splice_segment_xd,
        splice_zones=args.splice_zone,
        rebarcut_out=args.rebarcut_out,
    ))
    supervisor.register_agent(QSAgent(
        qs_path=args.qs,
        qs_sheet=args.qs_sheet,
        rate_overrides={"chung": args.rate_chung, "nha_tam": args.rate_nha_tam, "kxd": args.rate_kxd,
                        "tl": args.rate_tl, "vat": args.vat},
        qs_out=args.qs_out,
    ))
    supervisor.register_agent(BPTCKCSAgent())
    supervisor.register_agent(SchedulerAgent(
        schedule_path=args.schedule,
        schedule_sheet=args.schedule_sheet,
        start_date=args.start_date,
        non_working_weekdays=args.non_working_days,
        holidays=args.holidays,
        schedule_out=args.schedule_out,
    ))
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
            print("         tính tiến độ CPM từ file thật:       python run_state_graph.py --phase schedule --schedule <file>")
            print("         tính dự toán G_XD từ bảng QS thật:   python run_state_graph.py --phase qs --qs <file>")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
