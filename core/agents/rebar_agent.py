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
    SAW_KERF_MM, CuttingStockSolver, CutDemand, SpliceZoneValidator, write_cut_plan_csv,
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
        self.solver = CuttingStockSolver(bar_length_mm=bar_length_mm, kerf_mm=kerf_mm)
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

        # ── BƯỚC 2: Tách thanh dài hơn cây thép — phải nối, không cắt trọn từ 1 cây
        valid_demands = [d for d in demands if d.length_mm <= self.bar_length_mm]
        oversize = [d for d in demands if d.length_mm > self.bar_length_mm]
        oversize_items = [f"{d.mark} Ø{d.diameter_mm} L={d.length_mm}mm × {d.quantity}" for d in oversize]
        if oversize:
            input_summary.append(
                f"{len(oversize)} dòng BBS ({sum(d.quantity for d in oversize):,} thanh) dài hơn cây "
                f"{self.bar_length_mm}mm — phải tách đoạn nối theo bản vẽ, CHƯA có trong kế hoạch cắt"
            )
            print(f"  [RebarAgent] ⚠ {input_summary[-1]}: {oversize_items[:3]}")

        if not valid_demands:
            raise DataInputError(
                f"Không có thanh nào ≤ {self.bar_length_mm}mm để lập kế hoạch cắt (nguồn: {source})"
            )

        # ── BƯỚC 2b: Tool Calling — CuttingStockSolver ───────────────────────
        print("  [RebarAgent] Gọi CuttingStockSolver (OR-Tools, tách nhóm đường kính + mác thép)...")
        solution = self.solver.solve(valid_demands)

        print(f"  [RebarAgent] Solver: {solution.status} ({solution.solver_name})")
        print(f"  [RebarAgent] Số cây thép: {solution.total_bars_needed:,} (cận dưới {solution.lower_bound_bars:,})")
        print(f"  [RebarAgent] Đề-xê: {solution.waste_ratio_pct:.2f}% — "
              f"KL thép mua {solution.total_weight_kg:,.1f} kg / KL thành phẩm {solution.net_weight_kg:,.1f} kg")
        for g in solution.groups:
            print(f"    Ø{g.diameter_mm:<3} {g.grade:<9} {g.total_pieces:>7,} đoạn → {g.total_bars_needed:>6,} cây "
                  f"(cận dưới {g.lower_bound_bars:,}) đề-xê {g.waste_ratio_pct:5.2f}% [{g.status}]")
        for w in solution.warnings:
            print(f"  [RebarAgent] ⚠ {w}")

        if self.cut_plan_out:
            write_cut_plan_csv(solution, self.cut_plan_out)
            print(f"  [RebarAgent] Đã xuất phiếu cắt thép: {self.cut_plan_out}")

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
        pieces_demanded = sum(d.quantity for d in valid_demands)
        pieces_cut = sum(len(a["cuts_mm"]) for a in solution.assignment)
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

        if solution.status not in ("OPTIMAL", "FEASIBLE"):
            raise DataInputError(f"Solver trả về {solution.status}: {solution.warning}")

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
