# -*- coding: utf-8 -*-
"""
REBAR AGENT — Sub-Agent Gia công & Cắt thép
Minh họa hoàn chỉnh pattern: Tool Calling + Inter-Agent Validation Gate

Vai trò:
  1. Đọc BBS (Bảng Thống kê Cốt thép) thật từ file (--bbs) hoặc từ SharedState
  2. Gọi CuttingStockSolver (OR-Tools) — PURE PYTHON, ZERO LLM
  3. Kiểm tra vùng nối (TCVN 5574:2018) khi có dữ liệu vị trí nối
  4. Ghi kết quả vào StateBus → Supervisor đọc để chạy Quality Gate

Không có BBS thật → báo lỗi và dừng. BBS mẫu chỉ được dùng ở chế độ --demo.

LLM Role (nếu tích hợp): CHỈ được dùng để:
  - Intent recognition (phân tích yêu cầu từ người dùng)
  - Trình bày kết quả bằng ngôn ngữ tự nhiên
  - KHÔNG được làm toán học (cộng, trừ, nhân số, diện tích, khối lượng)
"""

from __future__ import annotations
import os
import sys
from typing import List, Optional, Tuple

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.supervisor.base_agent import BaseAgent, DataInputError, MissingDataError
from core.state.state_bus import StateBus
from tools.bbs_loader import BBSLoadError, load_bbs
from tools.cutting_stock_solver import (
    DEFAULT_LAP_XD, DEFAULT_MAX_SPLICE_RATIO, MIN_SPLICE_SEGMENT_XD, REUSE_OFFCUT_XD, SAW_KERF_MM,
    SHORT_OFFCUT_XD, CuttingStockSolver, CutDemand, SpliceZoneValidator, write_cut_plan_csv,
)

MAX_LISTED = 10  # Số dòng lỗi/cảnh báo tối đa in ra


class RebarAgent(BaseAgent):
    """
    Sub-Agent Gia công & Cắt thép — dùng OR-Tools Solver.

    Input  : file BBS thật (bbs_path) hoặc rebar_data.bbs_items trong StateBus
    Output (vào StateBus): rebar_data (cutting_result, splice_zone_check, splice_violations)
    """

    def __init__(
        self,
        bar_length_mm: int = 11_700,
        span_length_mm: int = 38_200,  # Chiều dài nhịp dầm Super-T để kiểm tra splice (demo)
        bbs_path: Optional[str] = None,
        bbs_sheet: Optional[str] = None,
        skip_invalid_rows: bool = False,
        cut_plan_out: Optional[str] = None,
        kerf_mm: int = SAW_KERF_MM,
        end_trim_mm: int = 0,
        max_pieces_per_bar: Optional[int] = None,
        max_marks_per_bar: Optional[int] = None,
        reuse_xd: float = REUSE_OFFCUT_XD,
        short_xd: float = SHORT_OFFCUT_XD,
        splice: bool = False,
        lap_xd: float = DEFAULT_LAP_XD,
        max_splice_ratio: float = DEFAULT_MAX_SPLICE_RATIO,
        min_splice_segment_xd: float = MIN_SPLICE_SEGMENT_XD,
        splice_zones: Optional[List[Tuple[float, float]]] = None,
        rebarcut_out: Optional[str] = None,
    ):
        super().__init__(
            agent_id="rebar_agent",
            description="Gia công & Cắt thép OR-Tools (TCVN 1651:2018 + TCVN 5574:2018)"
        )
        self.bar_length_mm = bar_length_mm
        self.span_length_mm = span_length_mm
        self.bbs_path = bbs_path
        self.bbs_sheet = bbs_sheet
        self.skip_invalid_rows = skip_invalid_rows
        self.cut_plan_out = cut_plan_out
        self.splice = splice
        self.rebarcut_out = rebarcut_out
        self.params = {
            "reuse_xd": reuse_xd, "short_xd": short_xd, "lap_xd": lap_xd,
            "max_splice_ratio": max_splice_ratio, "min_splice_segment_xd": min_splice_segment_xd,
            "max_pieces_per_bar": max_pieces_per_bar, "max_marks_per_bar": max_marks_per_bar,
        }
        self.solver = CuttingStockSolver(
            bar_length_mm=bar_length_mm, kerf_mm=kerf_mm, end_trim_mm=end_trim_mm,
            max_pieces_per_bar=max_pieces_per_bar, max_marks_per_bar=max_marks_per_bar,
            reuse_offcut_xd=reuse_xd, short_offcut_xd=short_xd, lap_xd=lap_xd,
            max_splice_ratio=max_splice_ratio, min_splice_segment_xd=min_splice_segment_xd,
            default_splice_zones=splice_zones,
        )
        self.splice_validator = SpliceZoneValidator()

    def run(self, bus: StateBus) -> bool:
        """
        Luồng thực thi:
        1. Lấy danh sách cắt từ BBS thật (hoặc BBS mẫu nếu --demo)
        2. Chạy OR-Tools Solver theo từng nhóm đường kính + mác thép
        3. Kiểm tra vùng nối (chỉ khi có dữ liệu vị trí nối)
        4. Ghi kết quả vào StateBus
        """
        print("  [RebarAgent] Bắt đầu tính toán cắt thép...")

        # ── BƯỚC 1: Lấy demands ───────────────────────────────────────────────
        demands, source, input_warnings, input_summary, is_sample = self._load_demands(bus)
        print(f"  [RebarAgent] Nguồn BBS: {source}")
        print(f"  [RebarAgent] Số dòng BBS: {len(demands)} — tổng số thanh: {sum(d.quantity for d in demands):,}")
        for w in input_warnings[:MAX_LISTED]:
            print(f"    ⚠ {w}")
        if len(input_warnings) > MAX_LISTED:
            print(f"    ... và {len(input_warnings) - MAX_LISTED} cảnh báo khác")

        # ── BƯỚC 2: Tool Calling — CuttingStockSolver ────────────────────────
        # Thanh dài hơn cây thép được tách thành các đoạn nối trong vùng cho phép nối;
        # thanh không có vùng nối / không bố trí được → 'unplanned' (cảnh báo, không bỏ qua lặng lẽ)
        print("  [RebarAgent] Gọi CuttingStockSolver (OR-Tools, tách nhóm đường kính + mác thép)...")
        solution = self.solver.solve(demands, split_long_bars=True)
        if solution.status not in ("OPTIMAL", "FEASIBLE"):
            raise DataInputError(f"Solver trả về {solution.status}: {solution.warning}")

        oversize_items = [f"{u['mark']} Ø{u['diameter_mm']} L={u['length_mm']}mm × {u['quantity']}: {u['reason']}"
                          for u in solution.unplanned]
        if solution.long_bars_split:
            input_summary.append(
                f"Đã tách {solution.long_bars_split:,} thanh dài hơn cây {self.bar_length_mm}mm thành đoạn nối "
                f"({sum(1 for s in solution.splices if s['mandatory']):,} mối nối so le trong vùng cho phép) "
                f"— đối chiếu bản vẽ trước khi gia công"
            )
        if solution.unplanned:
            input_summary.append(
                f"{len(solution.unplanned)} dòng BBS ({sum(u['quantity'] for u in solution.unplanned):,} thanh) "
                f"CHƯA có trong kế hoạch cắt (thanh dài hơn cây nhưng thiếu vùng cho phép nối hoặc "
                f"không bố trí được mối nối)"
            )
            for item in oversize_items[:MAX_LISTED]:
                print(f"    ⚠ CHƯA LẬP KẾ HOẠCH: {item}")

        print(f"  [RebarAgent] Solver: {solution.status} ({solution.solver_name})")
        print(f"  [RebarAgent] Số cây thép: {solution.total_bars_needed:,} (cận dưới {solution.lower_bound_bars:,})")
        print(f"  [RebarAgent] Đề-xê: {solution.waste_ratio_pct:.2f}% — "
              f"KL thép mua {solution.total_weight_kg:,.1f} kg / KL thành phẩm {solution.net_weight_kg:,.1f} kg")
        for g in solution.groups:
            print(f"    Ø{g.diameter_mm:<3} {g.grade:<9} {g.total_pieces:>7,} đoạn → {g.total_bars_needed:>6,} cây "
                  f"(cận dưới {g.lower_bound_bars:,}) đề-xê {g.waste_ratio_pct:5.2f}% [{g.status}]")
        for w in solution.warnings:
            print(f"  [RebarAgent] ⚠ {w}")
        offcuts = ", ".join(f"{cls} {n:,} ({mm / 1000:,.1f} m)" for cls, (n, mm) in solution.offcut_summary.items())
        print(f"  [RebarAgent] Đầu thừa: {offcuts}")

        # ── BƯỚC 2c: Phương án nối thép (đề xuất, cần duyệt) ─────────────────
        spliced = None
        splice_summary = {}
        if self.splice:
            print("  [RebarAgent] Tính phương án nối thép (chỉ Bar Mark có vùng cho phép nối)...")
            spliced = self.solver.solve(demands, allow_splicing=True, split_long_bars=True)
            saved = solution.total_bars_needed - spliced.total_bars_needed
            splice_summary = {
                "status": spliced.status,
                "total_bars_needed": spliced.total_bars_needed,
                "lower_bound_bars": spliced.lower_bound_bars,
                "bars_saved": saved,
                "weight_saved_kg": round(solution.total_weight_kg - spliced.total_weight_kg, 2),
                "splices": spliced.total_splices,
                "splice_extra_m": spliced.splice_extra_mm / 1000,
                "warnings": [w for w in spliced.warnings if w not in solution.warnings],
            }
            print(f"  [RebarAgent] PA nối: {spliced.total_bars_needed:,} cây (giảm {saved:,} cây, "
                  f"{splice_summary['weight_saved_kg']:,.1f} kg), {spliced.total_splices:,} mối nối, "
                  f"thép bù nối {splice_summary['splice_extra_m']:,.1f} m [{spliced.status}]")
            for w in splice_summary["warnings"][:MAX_LISTED]:
                print(f"    ⚠ {w}")

        if self.cut_plan_out:
            write_cut_plan_csv(solution, self.cut_plan_out)
            print(f"  [RebarAgent] Đã xuất phiếu cắt thép: {self.cut_plan_out}")
            if spliced is not None:
                stem, ext = os.path.splitext(self.cut_plan_out)
                write_cut_plan_csv(spliced, f"{stem}_PA_noi{ext or '.csv'}")
                print(f"  [RebarAgent] Đã xuất phiếu cắt PA nối (đề xuất): {stem}_PA_noi{ext or '.csv'}")
        if self.rebarcut_out:
            from tools.rebarcut_export import write_rebarcut_workbook
            write_rebarcut_workbook(self.rebarcut_out, demands, solution, spliced, self.params)
            print(f"  [RebarAgent] Đã xuất file theo bố cục RebarCut: {self.rebarcut_out}")

        # ── BƯỚC 3: Kiểm tra vùng nối ─────────────────────────────────────────
        if is_sample:
            # Chỉ ở demo: vị trí nối giả định 1/3 nhịp — không phải dữ liệu bản vẽ
            splice_positions = self._estimate_splice_positions(demands)
            splice_status, splice_violations = self.splice_validator.validate(
                splice_positions, span_length_mm=self.span_length_mm
            )
            print(f"  [RebarAgent] Splice check (vị trí nối GIẢ ĐỊNH — demo): {splice_status}")
        else:
            splice_status, splice_violations = "NOT_RUN", []
            print("  [RebarAgent] Splice check: NOT_RUN — chưa có dữ liệu vị trí nối từ bản vẽ")
        for v in splice_violations:
            print(f"    {v}")

        # ── BƯỚC 4: Ghi kết quả vào StateBus ─────────────────────────────────
        pieces_demanded = sum(d.quantity for d in demands) - sum(u["quantity"] for u in solution.unplanned)
        pieces_cut = (sum(1 for a in solution.assignment for m in a["marks"] if "(nối-" not in m)
                      + len({s["assembly_id"] for s in solution.splices}))
        cutting_dict = {
            "status": solution.status,
            "solver_name": solution.solver_name,
            "data_source": source,
            "bar_length_mm": solution.bar_length_mm,
            "kerf_mm": solution.kerf_mm,
            "total_bars_needed": solution.total_bars_needed,
            "lower_bound_bars": solution.lower_bound_bars,
            "total_waste_mm": solution.total_waste_mm,
            "waste_ratio_pct": solution.waste_ratio_pct,
            "total_weight_kg": solution.total_weight_kg,
            "net_weight_kg": solution.net_weight_kg,
            "reusable_offcuts": solution.reusable_offcuts,
            "pieces_demanded": pieces_demanded,
            "pieces_cut": pieces_cut,
            "input_warnings": input_warnings,
            "input_summary": input_summary,
            "oversize_items": oversize_items,
            "end_trim_mm": solution.end_trim_mm,
            "offcut_summary": solution.offcut_summary,
            "splice_plan": splice_summary,
            "groups": [g.__dict__ for g in solution.groups],
            "patterns": [
                {"diameter_mm": p.diameter_mm, "grade": p.grade, "cuts": p.cuts,
                 "marks": p.marks, "bars": p.bars_used, "waste_mm": p.waste_mm}
                for p in solution.patterns[:50]  # Giới hạn 50 phương án
            ],
        }

        bus.set_rebar_data({
            "cutting_result": type("CuttingStockResult", (), cutting_dict)(),  # noqa: preserved for StateBus
            "splice_zone_check": splice_status,
            "splice_violations": splice_violations,
            "total_rebar_kg": solution.total_weight_kg,
            # Lưu cutting_dict riêng để persist JSON
            "_cutting_dict_json": cutting_dict,
        })

        # Nếu có violation → trả về False để Supervisor biết cần Inter-Agent REJECT
        if splice_status == "REJECT":
            print("  [RebarAgent] REJECT: Có vi phạm vùng nối — báo cáo về Supervisor")
            # Không raise exception — trả False để Supervisor quyết định retry
            # (Supervisor sẽ escalate lên human gate nếu hết retry)
            return False

        return True

    # ─────────────────────────────────────────────────────────────────────────
    # HELPERS
    # ─────────────────────────────────────────────────────────────────────────

    def _load_demands(self, bus: StateBus) -> Tuple[List[CutDemand], str, List[str], List[str], bool]:
        """Trả về (demands, nguồn, cảnh báo chi tiết, tóm tắt cảnh báo, có phải dữ liệu mẫu)."""
        if self.bbs_path:
            try:
                result = load_bbs(self.bbs_path, sheet=self.bbs_sheet)
            except BBSLoadError as e:
                raise DataInputError(str(e)) from e

            warnings = list(result.skipped)
            summary = [f"Bỏ qua {len(result.skipped)} dòng không phải thép thanh (cáp DƯL...)"] \
                if result.skipped else []
            if result.errors:
                if not self.skip_invalid_rows:
                    listed = "\n    ".join(result.errors[:MAX_LISTED])
                    more = f"\n    ... và {len(result.errors) - MAX_LISTED} dòng khác" \
                        if len(result.errors) > MAX_LISTED else ""
                    raise DataInputError(
                        f"BBS {result.source} có {len(result.errors)} dòng sai dữ liệu — sửa BBS "
                        f"hoặc chạy với --bbs-skip-invalid để loại các dòng này:\n    {listed}{more}"
                    )
                warnings.extend(f"ĐÃ LOẠI (--bbs-skip-invalid): {e}" for e in result.errors)
                summary.append(
                    f"ĐÃ LOẠI {len(result.errors)} dòng BBS sai dữ liệu (--bbs-skip-invalid) — "
                    f"khối lượng các dòng này CHƯA có trong kế hoạch cắt"
                )
            if not result.demands:
                raise DataInputError(f"BBS {result.source} không còn dòng hợp lệ nào")
            return result.demands, result.source, warnings, summary, False

        demands = self._extract_demands_from_state(bus)
        if demands:
            return demands, "rebar_data.bbs_items (StateBus)", [], [], False

        if not bus.is_demo_mode():
            raise MissingDataError(
                "Chưa có BBS thật — chỉ định file bằng --bbs <file.xlsx|.csv|.json>. "
                "Muốn chạy thử với BBS mẫu Cầu Km19+529.080 thì thêm cờ --demo."
            )
        bus.mark_sample_data(self.agent_id, "BBS mẫu 10 loại thanh viết sẵn trong code")
        return self._get_sample_demands(), "BBS mẫu Cầu Km19 (demo)", [], [], True

    def _extract_demands_from_state(self, bus: StateBus) -> list:
        """Trích xuất danh sách cắt từ BBS trong SharedState (object hoặc dict)."""
        rebar_data = bus.get_rebar_data()
        demands = []
        for item in getattr(rebar_data, "bbs_items", []) or []:
            get = item.get if isinstance(item, dict) else (lambda k, d=None, _i=item: getattr(_i, k, d))
            length, count = get("length_mm"), get("count")
            if length and count:
                demands.append(CutDemand(
                    length_mm=int(length),
                    quantity=int(count),
                    diameter_mm=int(get("diameter_mm", 0) or 0),
                    mark=str(get("mark", "") or ""),
                    grade=str(get("grade", "") or ""),
                ))
        return demands

    def _get_sample_demands(self) -> list:
        """
        Dữ liệu BBS mẫu cho Cầu Km19+529.080 — Dầm Super-T + Cọc nhồi.
        CHỈ dùng ở chế độ --demo.
        """
        return [
            # Cọc nhồi Ø1200 — thép dọc Ø25
            CutDemand(length_mm=9000, quantity=104, diameter_mm=25, mark="CL1"),
            # Cọc nhồi — thép đai Ø16
            CutDemand(length_mm=5027, quantity=208, diameter_mm=16, mark="CD1"),
            # Dầm Super-T — thép Ø20
            CutDemand(length_mm=37800, quantity=15, diameter_mm=20, mark="DT1"),
            CutDemand(length_mm=4500, quantity=60, diameter_mm=20, mark="DT2"),
            # Bệ trụ — thép Ø22
            CutDemand(length_mm=6500, quantity=32, diameter_mm=22, mark="BT1"),
            CutDemand(length_mm=3200, quantity=48, diameter_mm=22, mark="BT2"),
            # Mặt cầu — thép Ø12
            CutDemand(length_mm=5800, quantity=120, diameter_mm=12, mark="MC1"),
            CutDemand(length_mm=3000, quantity=200, diameter_mm=12, mark="MC2"),
            # Mố cầu M1/M2 — thép Ø18
            CutDemand(length_mm=7200, quantity=28, diameter_mm=18, mark="MO1"),
            CutDemand(length_mm=4800, quantity=42, diameter_mm=18, mark="MO2"),
        ]

    def _estimate_splice_positions(self, demands: list) -> list:
        """
        Vị trí nối GIẢ ĐỊNH cho demo: thanh dài hơn cây thép được coi là nối tại 1/3 nhịp,
        vùng nén. Không dùng cho dữ liệu thật — thực tế phải đọc từ bản vẽ chi tiết.
        """
        splice_positions = []
        long_pieces = [d for d in demands if d.length_mm > self.bar_length_mm * 0.95]

        for d in long_pieces:
            splice_positions.append({
                "mark": d.mark,
                "splice_at_mm": int(self.span_length_mm * 0.33),
                "zone": "COMPRESSION",  # Dầm Super-T: vùng 1/3 đầu là vùng nén
            })

        return splice_positions


from core.agents.registry import register_agent  # noqa: E402


@register_agent("rebar_agent", order=20)
def _make_rebar_agent(args) -> "RebarAgent":
    return RebarAgent(
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
    )
