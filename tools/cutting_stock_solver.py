# -*- coding: utf-8 -*-
"""
OR-TOOLS 1D CUTTING STOCK SOLVER
Bài toán cắt thép một chiều (1D Cutting Stock Problem)
Tối ưu ghép các đoạn cắt vào cây thép nguyên 11.7m — mục tiêu: ít cây thép nhất.

Đây là Pure Python Tool — KHÔNG có bất kỳ LLM nào tham gia tính toán.
LLM chỉ được phép: gọi tool này, đọc kết quả, trình bày ngôn ngữ tự nhiên.

Nguyên tắc thực tế xưởng gia công:
  - Chỉ ghép các đoạn CÙNG đường kính và CÙNG mác thép vào một cây thép.
  - Mỗi nhát cắt mất một lượng bằng bề rộng lưỡi cắt (kerf); có thể cắt bỏ đầu cây.
  - Tùy chọn giới hạn cho tổ cắt: tối đa số đoạn / cây, tối đa số Bar Mark / cây.
  - Đầu thừa phân loại theo đường kính: ≥ 100D tái sử dụng, ≥ 20D đầu thừa ngắn, còn lại phế.

Phương án nối thép (allow_splicing=True) — CHỈ là đề xuất, phải được kỹ thuật duyệt:
  - Chỉ nối thanh có splice_allowed=True VÀ có vùng cho phép nối (splice_zones).
  - Một thanh được ghép tối đa từ 2 đoạn A + B = L + L_nối; vùng chồng nối phải
    nằm trọn trong vùng cho phép; mỗi đoạn ≥ max(L_nối, 20D).
  - Số thanh được nối của mỗi Bar Mark ≤ tỷ lệ cho phép × số lượng.
  - Phép nối được đưa vào chính mô hình tối ưu (không phải vá sau), số mối nối
    được giảm tối đa sau khi đã tối thiểu số cây.

Thuật toán (chạy riêng cho từng nhóm đường kính + mác thép):
  1. Heuristic tham lam theo phương án cắt → lời giải khởi đầu (và dự phòng khi
     không có OR-Tools).
  2. Column Generation Gilmore–Gomory: LP (OR-Tools GLOP) + bài toán cái túi
     (quy hoạch động) sinh phương án cắt → cận dưới số cây thép.
  3. OR-Tools CP-SAT chọn số lần lặp nguyên cho từng phương án cắt.

Trạng thái lời giải:
  OPTIMAL  — số cây thép bằng cận dưới → đã chứng minh là ít nhất có thể
             (với phương án nối: trong tập vị trí nối đã sinh).
  FEASIBLE — phương án hợp lệ nhưng chưa chứng minh tối ưu (xem lower_bound_bars).

Input:  demands = [CutDemand(length_mm, quantity, diameter_mm, mark, grade, ...), ...]
Output: CuttingStockSolution
"""

from __future__ import annotations
import csv
import math
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS — TCVN 1651:2018 & thực tế công trường
# ─────────────────────────────────────────────────────────────────────────────

STANDARD_BAR_LENGTH_MM = 11_700    # Cây thép nguyên 11.7m
SAW_KERF_MM = 3                    # Lượng mất mát mỗi nhát cưa (mm)
MAX_WASTE_RATIO_PCT = 1.5          # Mục tiêu đề-xê (%) — chỉ để cảnh báo
REUSE_OFFCUT_XD = 100              # Đầu thừa ≥ 100D: nhập kho tái sử dụng
SHORT_OFFCUT_XD = 20               # Đầu thừa ≥ 20D: đầu thừa ngắn; ngắn hơn: phế
DEFAULT_LAP_XD = 40                # Chiều dài nối chồng mặc định (tham khảo, theo hồ sơ dự án)
DEFAULT_MAX_SPLICE_RATIO = 0.5     # Tỷ lệ thanh được nối tối đa mỗi Bar Mark
MIN_SPLICE_SEGMENT_XD = 20         # Đoạn nối tối thiểu mỗi phía
MAX_SPLIT_OPTIONS = 24             # Số vị trí nối thử cho mỗi Bar Mark

OFFCUT_CLASSES = ("Hết cây", "Tái sử dụng", "Đầu thừa ngắn", "Phế")

REBAR_UNIT_WEIGHT_KG_M: Dict[int, float] = {
    6:   0.222,   8:   0.395,  10:   0.617,
    12:  0.888,  14:   1.208,  16:   1.578,
    18:  1.998,  20:   2.466,  22:   2.984,
    25:  3.853,  28:   4.834,  32:   6.313,
}

STEEL_DENSITY_KG_M3 = 7_850


def rebar_unit_weight_kg_m(diameter_mm: int) -> float:
    """Trọng lượng 1m thép thanh (kg/m) — tra bảng, ngoài bảng tính theo π·d²/4·ρ."""
    if diameter_mm in REBAR_UNIT_WEIGHT_KG_M:
        return REBAR_UNIT_WEIGHT_KG_M[diameter_mm]
    if diameter_mm and diameter_mm > 0:
        return round(math.pi / 4 * (diameter_mm / 1000) ** 2 * STEEL_DENSITY_KG_M3, 3)
    return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# DATA CLASSES
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class CutDemand:
    length_mm: int
    quantity: int
    diameter_mm: int = 0
    mark: str = ""
    grade: str = ""
    # Nối thép (chỉ dùng khi solve(..., allow_splicing=True))
    splice_allowed: bool = False
    splice_kind: str = ""                                   # "Nối chồng" / "Coupler"
    lap_xd: Optional[float] = None                          # None → mặc định của solver
    max_splice_ratio: Optional[float] = None                # None → mặc định của solver
    splice_zones: Optional[List[Tuple[float, float]]] = None  # tỷ lệ chiều dài thanh (0..1)


@dataclass
class CuttingPattern:
    """Một phương án ghép các đoạn cắt vào 1 cây thép, lặp lại `bars_used` cây."""
    cuts: List[int] = field(default_factory=list)  # list chiều dài (mm)
    waste_mm: int = 0                               # phần còn lại của cây (sau cắt đầu)
    bars_used: int = 1
    diameter_mm: int = 0
    grade: str = ""
    marks: List[str] = field(default_factory=list)  # ký hiệu tương ứng từng đoạn cắt
    offcut_class: str = ""


@dataclass
class DiameterGroupResult:
    """Kết quả của một nhóm cùng đường kính + mác thép."""
    diameter_mm: int
    grade: str
    status: str
    solver_name: str
    total_pieces: int
    total_bars_needed: int
    lower_bound_bars: int
    total_waste_mm: int
    waste_ratio_pct: float
    purchased_weight_kg: float
    net_weight_kg: float
    reusable_offcuts: int = 0            # số đầu thừa ≥ 100D
    reusable_offcut_length_mm: int = 0
    offcut_summary: Dict[str, List[int]] = field(default_factory=dict)  # loại → [số, tổng mm]
    splices: int = 0
    splice_extra_mm: int = 0             # thép tăng thêm do chồng nối


@dataclass
class CuttingStockSolution:
    status: str = "NOT_RUN"          # OPTIMAL / FEASIBLE / INFEASIBLE / ERROR
    solver_name: str = "OR-Tools GLOP + CP-SAT"
    bar_length_mm: int = STANDARD_BAR_LENGTH_MM
    kerf_mm: int = SAW_KERF_MM
    end_trim_mm: int = 0
    splicing: bool = False
    total_bars_needed: int = 0
    lower_bound_bars: int = 0        # Cận dưới số cây (tổng các nhóm)
    total_pieces: int = 0
    total_length_used_mm: int = 0
    total_waste_mm: int = 0
    waste_ratio_pct: float = 0.0
    total_weight_kg: float = 0.0     # Khối lượng thép cây phải mua (số cây × 11.7m × kg/m)
    net_weight_kg: float = 0.0       # Khối lượng các đoạn thành phẩm theo BBS
    reusable_offcuts: int = 0        # Số đầu thừa ≥ 100D có thể nhập kho tận dụng
    reusable_offcut_length_mm: int = 0
    offcut_summary: Dict[str, List[int]] = field(default_factory=dict)
    total_splices: int = 0           # số mối nối (thanh k đoạn có k−1 mối nối)
    splice_extra_mm: int = 0
    long_bars_split: int = 0         # số thanh dài hơn cây thép đã được tách đoạn
    patterns: List[CuttingPattern] = field(default_factory=list)  # phương án khác nhau
    assignment: List[Dict] = field(default_factory=list)          # chi tiết từng cây
    splices: List[Dict] = field(default_factory=list)             # chi tiết từng mối nối
    unplanned: List[Dict] = field(default_factory=list)           # thanh CHƯA có trong kế hoạch cắt
    groups: List[DiameterGroupResult] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    warning: str = ""                # Gộp warnings (giữ tương thích)


@dataclass
class _SplitOption:
    """Một cách ghép thanh `owner` từ các đoạn nối (theo thứ tự dọc thanh)."""
    owner: int
    segments: List[int]              # chiều dài từng đoạn, từ đầu thanh
    seg_items: List[int]             # chỉ số item của từng đoạn
    lap: int
    laps: List[Tuple[int, int]]      # vùng chồng nối (từ, đến) đo từ đầu thanh
    mandatory: bool                  # thanh dài hơn cây thép: bắt buộc nối

    @property
    def centers(self) -> List[float]:
        return [(a + b) / 2 for a, b in self.laps]


# ─────────────────────────────────────────────────────────────────────────────
# SOLVER
# ─────────────────────────────────────────────────────────────────────────────

class CuttingStockSolver:
    """
    1D Cutting Stock Solver — tách nhóm theo (đường kính, mác thép), mỗi nhóm:
      Heuristic tham lam → Column Generation (GLOP) → CP-SAT nghiệm nguyên.
    """

    # Hai mối nối chồng coi là cùng một mặt cắt khi tâm cách nhau < 1.3 × L_nối
    SAME_SECTION_FACTOR = 1.3

    def __init__(
        self,
        bar_length_mm: int = STANDARD_BAR_LENGTH_MM,
        kerf_mm: int = SAW_KERF_MM,
        time_limit_s: float = 20.0,
        use_ortools: bool = True,
        end_trim_mm: int = 0,
        max_pieces_per_bar: Optional[int] = None,
        max_marks_per_bar: Optional[int] = None,
        reuse_offcut_xd: float = REUSE_OFFCUT_XD,
        short_offcut_xd: float = SHORT_OFFCUT_XD,
        lap_xd: float = DEFAULT_LAP_XD,
        max_splice_ratio: float = DEFAULT_MAX_SPLICE_RATIO,
        min_splice_segment_xd: float = MIN_SPLICE_SEGMENT_XD,
        default_splice_zones: Optional[List[Tuple[float, float]]] = None,
    ):
        self.bar_length_mm = bar_length_mm
        self.kerf_mm = kerf_mm
        self.time_limit_s = time_limit_s  # giới hạn thời gian cho mỗi nhóm đường kính
        self.use_ortools = use_ortools
        self.end_trim_mm = end_trim_mm
        self.max_pieces_per_bar = max_pieces_per_bar
        self.max_marks_per_bar = max_marks_per_bar
        self.reuse_offcut_xd = reuse_offcut_xd
        self.short_offcut_xd = short_offcut_xd
        self.lap_xd = lap_xd
        self.max_splice_ratio = max_splice_ratio
        self.min_splice_segment_xd = min_splice_segment_xd
        self.default_splice_zones = default_splice_zones

    @property
    def usable_length_mm(self) -> int:
        return self.bar_length_mm - 2 * self.end_trim_mm

    # ── PUBLIC API ─────────────────────────────────────────────────────────

    def solve(
        self,
        demands: List[CutDemand],
        allow_splicing: bool = False,
        split_long_bars: bool = False,
    ) -> CuttingStockSolution:
        """
        allow_splicing : đề xuất nối thanh ngắn hơn cây để tận dụng đầu thừa (cần duyệt).
        split_long_bars: tách thanh dài hơn cây thép thành các đoạn nối trong vùng cho phép nối;
                         thanh không tách được đưa vào `unplanned` (có cảnh báo).
                         False: có thanh dài hơn cây → INFEASIBLE.
        """
        sol = CuttingStockSolution(
            bar_length_mm=self.bar_length_mm, kerf_mm=self.kerf_mm,
            end_trim_mm=self.end_trim_mm, splicing=allow_splicing,
        )

        def fail(status: str, message: str) -> CuttingStockSolution:
            sol.status = status
            sol.warning = message
            sol.warnings.append(message)
            return sol

        if self.usable_length_mm <= 0:
            return fail("ERROR", f"Cắt đầu {self.end_trim_mm}mm × 2 không còn chiều dài cây để cắt")
        for limit in (self.max_pieces_per_bar, self.max_marks_per_bar):
            if limit is not None and limit < 1:
                return fail("ERROR", "Giới hạn số đoạn / số Bar Mark trên cây phải ≥ 1")

        invalid = [d for d in demands if d.length_mm <= 0 or d.quantity < 0]
        if invalid:
            return fail("ERROR", f"Có {len(invalid)} dòng chiều dài ≤ 0 hoặc số lượng âm: "
                                 f"{[(d.mark, d.length_mm, d.quantity) for d in invalid[:5]]}")
        bad_zones = [d.mark for d in demands if d.splice_zones and any(
            not (0 <= a < b <= 1) for a, b in d.splice_zones)]
        if bad_zones:
            return fail("ERROR", f"Vùng cho phép nối phải nằm trong 0..1 và bắt đầu < kết thúc: {bad_zones[:5]}")

        active = [d for d in demands if d.quantity > 0]
        if not active:
            return fail("INFEASIBLE", "Danh sách cần cắt rỗng")

        too_long = [d.length_mm for d in active if d.length_mm > self.usable_length_mm]
        if too_long and not split_long_bars:
            return fail("INFEASIBLE", f"Có {len(too_long)} đoạn dài hơn cây {self.bar_length_mm}mm"
                                      f"{' (sau cắt đầu)' if self.end_trim_mm else ''}: {too_long[:5]}")

        groups: Dict[Tuple[int, str], List[CutDemand]] = defaultdict(list)
        for d in active:
            groups[(int(d.diameter_mm or 0), (d.grade or "").strip())].append(d)

        if any(diameter == 0 for diameter, _ in groups):
            sol.warnings.append(
                "Có đoạn thép không ghi đường kính — được ghép chung một nhóm riêng "
                "và không tính được khối lượng"
            )

        solver_names = set()
        for key in sorted(groups):
            name = self._solve_diameter_group(key, groups[key], allow_splicing, sol)
            if name:
                solver_names.add(name)

        if not sol.groups:
            return fail("INFEASIBLE", "Không có thanh nào lập được kế hoạch cắt: "
                                      + "; ".join(u["reason"] for u in sol.unplanned[:3]))

        sol.total_bars_needed = sum(g.total_bars_needed for g in sol.groups)
        sol.lower_bound_bars = sum(g.lower_bound_bars for g in sol.groups)
        sol.total_pieces = sum(g.total_pieces for g in sol.groups)
        sol.total_waste_mm = sum(g.total_waste_mm for g in sol.groups)
        total_capacity = sol.total_bars_needed * self.bar_length_mm
        sol.total_length_used_mm = total_capacity - sol.total_waste_mm
        sol.waste_ratio_pct = (sol.total_waste_mm / total_capacity * 100) if total_capacity else 0.0
        sol.total_weight_kg = round(sum(g.purchased_weight_kg for g in sol.groups), 2)
        sol.net_weight_kg = round(sum(g.net_weight_kg for g in sol.groups), 2)
        sol.reusable_offcuts = sum(g.reusable_offcuts for g in sol.groups)
        sol.reusable_offcut_length_mm = sum(g.reusable_offcut_length_mm for g in sol.groups)
        for g in sol.groups:
            for cls, (n, mm) in g.offcut_summary.items():
                acc = sol.offcut_summary.setdefault(cls, [0, 0])
                acc[0] += n
                acc[1] += mm
        sol.total_splices = len(sol.splices)
        sol.splice_extra_mm = sum(s["lap_mm"] for s in sol.splices)
        sol.long_bars_split = len({s["assembly_id"] for s in sol.splices if s["mandatory"]})
        sol.status = "OPTIMAL" if all(g.status == "OPTIMAL" for g in sol.groups) else "FEASIBLE"
        sol.solver_name = ", ".join(sorted(solver_names))

        if sol.waste_ratio_pct > MAX_WASTE_RATIO_PCT:
            if sol.status == "OPTIMAL":
                sol.warnings.append(
                    f"Đề-xê {sol.waste_ratio_pct:.2f}% > {MAX_WASTE_RATIO_PCT}% nhưng số cây đã bằng "
                    f"cận dưới (tối ưu đã chứng minh) — hao hụt do chiều dài thanh trong BBS. "
                    f"Muốn giảm phải điều chỉnh chiều dài/vị trí nối hoặc tận dụng đầu thừa "
                    f"({sol.reusable_offcuts} đầu thừa ≥ {self.reuse_offcut_xd:g}D, "
                    f"tổng {sol.reusable_offcut_length_mm / 1000:,.1f} m)."
                )
            else:
                sol.warnings.append(
                    f"Đề-xê {sol.waste_ratio_pct:.2f}% > {MAX_WASTE_RATIO_PCT}%; phương án chưa chứng minh "
                    f"tối ưu (dùng {sol.total_bars_needed} cây, cận dưới {sol.lower_bound_bars} cây)."
                )
        optional = sum(1 for s in sol.splices if not s["mandatory"])
        if optional:
            sol.warnings.append(
                f"Phương án có {optional} mối nối tận dụng đầu thừa — CHỈ là đề xuất vật tư, phải được kỹ thuật "
                f"duyệt vị trí nối, tỷ lệ nối trên mặt cắt và cấu tạo theo hồ sơ dự án."
            )
        if sol.long_bars_split:
            sol.warnings.append(
                f"Đã tách {sol.long_bars_split} thanh dài hơn cây thép "
                f"({sum(1 for s in sol.splices if s['mandatory'])} mối nối, so le tâm ≥ "
                f"{self.SAME_SECTION_FACTOR:g}×L_nối) — đối chiếu vị trí nối với bản vẽ trước khi gia công."
            )
        if sol.unplanned:
            sol.warnings.append(
                f"{len(sol.unplanned)} dòng BBS ({sum(u['quantity'] for u in sol.unplanned):,} thanh) "
                f"CHƯA có trong kế hoạch cắt — xem danh sách 'unplanned'"
            )
        sol.warning = " | ".join(sol.warnings)
        return sol

    # ── ONE DIAMETER + GRADE GROUP ─────────────────────────────────────────

    def _solve_diameter_group(
        self, key: Tuple[int, str], demands: List[CutDemand], allow_splicing: bool,
        sol: CuttingStockSolution,
    ) -> Optional[str]:
        diameter, grade = key
        usable = self.usable_length_mm

        def unplanned(d: CutDemand, reason: str) -> None:
            sol.unplanned.append({"mark": d.mark, "diameter_mm": diameter, "grade": grade,
                                  "length_mm": d.length_mm, "quantity": d.quantity, "reason": reason})

        # Item nguyên: gộp theo chiều dài, trừ khi cần theo dõi từng Bar Mark
        per_mark = self.max_marks_per_bar is not None
        item_keys: Dict[Tuple, int] = {}
        lengths: List[int] = []
        qty: List[int] = []
        marks: List[List[List]] = []
        owner_demand: Dict[int, CutDemand] = {}
        for d in sorted(demands, key=lambda x: -x.length_mm):
            zones = d.splice_zones or self.default_splice_zones
            if d.length_mm > usable:
                if not zones:
                    unplanned(d, f"dài {d.length_mm}mm > cây {usable}mm nhưng chưa có vùng cho phép nối "
                                 f"dạng số để tách đoạn")
                    continue
                spliceable = True
            else:
                spliceable = allow_splicing and d.splice_allowed
                if spliceable and not zones:
                    sol.warnings.append(
                        f"{d.mark} Ø{diameter}: cho nối nhưng chưa có vùng cho phép nối dạng số "
                        f"(vd 0-0.25; 0.75-1) — KHÔNG đề xuất nối"
                    )
                    spliceable = False
            # Thanh được nối: mỗi dòng BBS là một item riêng để tỷ lệ nối / so le tính theo từng dòng
            # (hai dòng trùng ký hiệu có thể là hai cấu kiện khác nhau)
            k = (int(d.length_mm), d.mark if per_mark else None, id(d) if spliceable else None)
            if k not in item_keys:
                item_keys[k] = len(lengths)
                lengths.append(int(d.length_mm))
                qty.append(0)
                marks.append([])
            i = item_keys[k]
            qty[i] += int(d.quantity)
            marks[i].append([d.mark or "?", int(d.quantity)])
            if spliceable:
                owner_demand[i] = d

        n_whole = len(lengths)
        seg_item_of: Dict[int, int] = {}      # chiều dài đoạn nối → item (dùng chung trong nhóm)
        options: List[_SplitOption] = []
        ratio_cap: Dict[int, int] = {}
        stagger_rows: List[Tuple[List[int], int]] = []
        initial_splits: List[int] = []
        for i, d in owner_demand.items():
            ratio = self.max_splice_ratio if d.max_splice_ratio is None else d.max_splice_ratio
            ratio = min(ratio, self.max_splice_ratio)
            lap = int(round((d.lap_xd if d.lap_xd is not None else self.lap_xd) * diameter))
            zones = d.splice_zones or self.default_splice_zones
            mandatory = lengths[i] > usable
            if ratio > 0.5 and not mandatory:
                sol.warnings.append(f"{d.mark}: tỷ lệ nối {ratio:.0%} > 50% — kiểm tra quy định nối trên cùng mặt cắt")

            if not mandatory:
                cap = int(qty[i] * ratio + 1e-9)
                if cap <= 0:
                    continue
                layouts = [(segs, laps) for segs, laps in self._split_candidates(
                    lengths[i], lap, diameter, zones, lengths[:n_whole])]
                if layouts:
                    ratio_cap[i] = cap
            else:
                layouts = self._long_bar_layouts(lengths[i], lap, diameter, zones)
                if not layouts:
                    unplanned(d, f"không bố trí được các đoạn ≤ {usable}mm với mối nối nằm trong vùng "
                                 f"cho phép {zones}")
                    qty[i] = 0
                    continue

            first = len(options)
            for segs, laps in layouts:
                seg_items = []
                for seg in segs:
                    if seg not in seg_item_of:
                        seg_item_of[seg] = len(lengths)
                        lengths.append(seg)
                        qty.append(0)
                        marks.append([])
                    seg_items.append(seg_item_of[seg])
                options.append(_SplitOption(i, list(segs), seg_items, lap, list(laps), mandatory))
                initial_splits.append(0)

            if mandatory:
                own = list(range(first, len(options)))
                sec_cap = max(1, int(qty[i] * ratio + 1e-9))
                rows = self._stagger_rows(options, own, lap, sec_cap)
                start = self._initial_stagger(options, own, lap, sec_cap, qty[i])
                if start is None:
                    unplanned(d, f"không đủ vị trí nối so le (tâm cách ≥ {self.SAME_SECTION_FACTOR:g}×L_nối) "
                                 f"để mỗi mặt cắt ≤ {ratio:.0%} số thanh — mở rộng vùng cho phép nối")
                    qty[i] = 0
                    del options[first:]
                    del initial_splits[first:]
                    continue
                stagger_rows.extend(rows)
                for j, count in start.items():
                    initial_splits[j] = count

        if sum(qty) == 0 and not any(initial_splits):
            return None

        item_bars, split_counts, status, lower_bound, solver_name = self._solve_group(
            lengths, qty, n_whole, options, ratio_cap, stagger_rows, initial_splits
        )

        # ── Gán ký hiệu và ghi kết quả ────────────────────────────────────────
        unit_w = rebar_unit_weight_kg_m(diameter)
        mark_queue = {i: [list(m) for m in marks[i]] for i in range(n_whole)}
        unique_marks = {i: "|".join(dict.fromkeys(m[0] for m in marks[i])) for i in range(n_whole)}
        seg_queue: Dict[int, List[Tuple[int, int, int]]] = defaultdict(list)  # item → [(ghép, lần, vị trí)]
        for j, (opt, count) in enumerate(zip(options, split_counts)):
            for inst in range(count):
                for pos, item in enumerate(opt.seg_items):
                    seg_queue[item].append((j, inst, pos))
        seg_label = {item: "|".join(dict.fromkeys(unique_marks[options[j].owner] for j, _, _ in queue))
                     for item, queue in seg_queue.items()}
        seg_bar: Dict[Tuple[int, int, int], int] = {}   # (ghép, lần, vị trí đoạn) → mã cây

        bar_offset = len(sol.assignment)
        offcuts: List[int] = []
        summary: Dict[str, List[int]] = {c: [0, 0] for c in OFFCUT_CLASSES}
        pattern_counter: Counter = Counter()
        pattern_info: Dict[Tuple, Tuple] = {}
        for n, bar in enumerate(item_bars, start=1):
            bar_id = bar_offset + n
            cuts = [lengths[i] for i in bar]
            used = sum(cuts) + self.kerf_mm * max(len(cuts) - 1, 0)
            leftover = usable - used
            cls = self.classify_offcut(leftover, diameter)
            summary[cls][0] += 1
            summary[cls][1] += leftover
            offcuts.append(leftover)
            labels = []
            for i in bar:
                if i < n_whole:
                    labels.append(_pop_mark(mark_queue[i]))
                else:
                    j, inst, pos = seg_queue[i].pop(0)
                    seg_bar[(j, inst, pos)] = bar_id
                    labels.append(f"{unique_marks[options[j].owner]}(nối-{chr(65 + pos)})")
            sol.assignment.append({
                "bar_id": bar_id,
                "diameter_mm": diameter,
                "grade": grade,
                "cuts_mm": cuts,
                "marks": labels,
                "used_mm": used,
                "waste_mm": leftover,
                "offcut_class": cls,
            })
            pkey = tuple(bar)
            pattern_counter[pkey] += 1
            pattern_info[pkey] = (cuts, leftover, cls)

        for pkey, count in pattern_counter.most_common():
            cuts, leftover, cls = pattern_info[pkey]
            sol.patterns.append(CuttingPattern(
                cuts=cuts, waste_mm=leftover, bars_used=count, diameter_mm=diameter, grade=grade,
                marks=[unique_marks[i] if i < n_whole else f"{seg_label.get(i, '?')}(nối)" for i in pkey],
                offcut_class=cls,
            ))

        group_splices = 0
        for j, (opt, count) in enumerate(zip(options, split_counts)):
            for inst in range(count):
                assembly = f"G{len({s['assembly_id'] for s in sol.splices}) + 1}"
                bars = [seg_bar[(j, inst, pos)] for pos in range(len(opt.seg_items))]
                for pos, (lap_from, lap_to) in enumerate(opt.laps):
                    group_splices += 1
                    sol.splices.append({
                        "splice_id": f"S{len(sol.splices) + 1}",
                        "assembly_id": assembly,
                        "mandatory": opt.mandatory,
                        "diameter_mm": diameter,
                        "grade": grade,
                        "mark": unique_marks[opt.owner],
                        "bar_length_mm": lengths[opt.owner],
                        "segments_mm": list(opt.segments),
                        "segment_bars": bars,
                        "a_mm": opt.segments[pos],
                        "b_mm": opt.segments[pos + 1],
                        "a_label": chr(65 + pos),
                        "b_label": chr(66 + pos),
                        "lap_mm": opt.lap,
                        "lap_from_mm": lap_from,
                        "lap_to_mm": lap_to,
                        "first_segment": "A",
                        "bar_a": bars[pos],
                        "bar_b": bars[pos + 1],
                        "kind": owner_demand[opt.owner].splice_kind or "Nối chồng",
                    })

        pieces = sum(qty[:n_whole])
        net_len = sum(lengths[i] * qty[i] for i in range(n_whole))
        capacity = len(item_bars) * self.bar_length_mm
        group_waste = sum(offcuts) + 2 * self.end_trim_mm * len(item_bars)
        reusable = summary["Tái sử dụng"]
        sol.groups.append(DiameterGroupResult(
            diameter_mm=diameter,
            grade=grade,
            status=status,
            solver_name=solver_name,
            total_pieces=pieces,
            total_bars_needed=len(item_bars),
            lower_bound_bars=lower_bound,
            total_waste_mm=group_waste,
            waste_ratio_pct=(group_waste / capacity * 100) if capacity else 0.0,
            purchased_weight_kg=round(capacity / 1000 * unit_w, 2),
            net_weight_kg=round(net_len / 1000 * unit_w, 2),
            reusable_offcuts=reusable[0],
            reusable_offcut_length_mm=reusable[1],
            offcut_summary={c: v for c, v in summary.items() if v[0]},
            splices=group_splices,
            splice_extra_mm=sum(o.lap * len(o.laps) * c for o, c in zip(options, split_counts)),
        ))
        return solver_name

    def classify_offcut(self, leftover_mm: int, diameter_mm: int) -> str:
        if leftover_mm <= 0:
            return "Hết cây"
        d = diameter_mm or 10  # không có đường kính: ngưỡng theo Ø10
        if leftover_mm >= self.reuse_offcut_xd * d:
            return "Tái sử dụng"
        if leftover_mm >= self.short_offcut_xd * d:
            return "Đầu thừa ngắn"
        return "Phế"

    # ── CÁCH GHÉP THANH TỪ CÁC ĐOẠN NỐI ────────────────────────────────────

    def _min_segment(self, lap: int, diameter: int) -> int:
        return max(lap, int(math.ceil(self.min_splice_segment_xd * diameter)), 1)

    @staticmethod
    def _in_zones(lap_from: float, lap_to: float, length: int, zones) -> bool:
        return any(z0 * length - 1e-6 <= lap_from and lap_to <= z1 * length + 1e-6 for z0, z1 in zones)

    def _split_candidates(
        self, length: int, lap: int, diameter: int, zones: List[Tuple[float, float]],
        other_lengths: Sequence[int],
    ) -> List[Tuple[List[int], List[Tuple[int, int]]]]:
        """
        Nối tùy chọn (tận dụng đầu thừa): tách thanh `length` thành 2 đoạn A + B = length + lap.
        Trả về [([đoạn đầu, đoạn cuối], [vùng chồng nối])] — ưu tiên đoạn lấp vừa phần dư
        của các phương án cắt thông thường.
        """
        usable, kerf = self.usable_length_mm, self.kerf_mm
        min_seg = self._min_segment(lap, diameter)
        total = length + lap

        preferred: List[int] = []
        for other in sorted(set(other_lengths), reverse=True):
            for n in (1, 2, 3):
                remnant = usable - n * other - n * kerf   # phần dư sau n đoạn (trừ nhát cắt)
                if remnant > 0:
                    preferred.extend([remnant, total - remnant])
        preferred.append(usable)
        preferred.append(total - usable)
        grid = list(range(min_seg, total - min_seg + 1, 250))

        result = []
        seen = set()
        for a in preferred + grid:
            b = total - a
            if a < min_seg or b < min_seg or a > usable or b > usable:
                continue
            key = (min(a, b), max(a, b))
            if key in seen:
                continue
            for first, second in ((a, b), (b, a)):   # đoạn nào ở đầu thanh
                lap_from = first - lap
                if self._in_zones(lap_from, first, length, zones):
                    seen.add(key)
                    result.append(([first, second], [(lap_from, first)]))
                    break
            if len(result) >= MAX_SPLIT_OPTIONS:
                break
        return result

    def _long_bar_layouts(
        self, length: int, lap: int, diameter: int, zones: List[Tuple[float, float]],
        max_layouts: int = 40,
    ) -> List[Tuple[List[int], List[Tuple[int, int]]]]:
        """
        Thanh dài hơn cây thép: tách thành k đoạn (k nhỏ nhất có thể, thử thêm k+1),
        mỗi đoạn ≤ cây thép và ≥ max(L_nối, 20D), mọi vùng chồng nối nằm trong vùng cho phép.
        Vùng chồng nối thứ j là [x_j, x_j + lap]; đoạn j dài x_j + lap − x_{j-1}.
        Sinh nhiều vị trí nối đầu tiên khác nhau để có thể bố trí so le giữa các thanh.
        """
        usable = self.usable_length_mm
        min_seg = self._min_segment(lap, diameter)
        if usable <= lap:
            return []
        intervals = sorted((z0 * length, z1 * length - lap) for z0, z1 in zones if z1 * length - lap >= z0 * length)
        if not intervals:
            return []
        step = max(50, int(self.SAME_SECTION_FACTOR * lap / 2))

        def allowed(lo: float, hi: float, spread: bool) -> List[int]:
            """Vị trí bắt đầu vùng nối cho phép trong [lo, hi]: điểm xa nhất trước, rồi rải đều."""
            picks: List[int] = []
            for a, b in intervals:
                a, b = max(a, lo), min(b, hi)
                if a > b + 1e-9:
                    continue
                a, b = int(math.ceil(a)), int(math.floor(b))
                if a > b:
                    continue
                picks.append(b)
                if spread:
                    picks.extend(range(b - step, a - 1, -step))
                    picks.append(a)
                else:
                    picks.append(a)
            return sorted(set(picks), reverse=True)

        layouts: List[Tuple[List[int], List[Tuple[int, int]]]] = []
        seen = set()

        def extend(start: int, xs: List[int], remaining: int) -> None:
            if len(layouts) >= max_layouts:
                return
            if remaining == 1:
                last = length - start
                if min_seg <= last <= usable:
                    segs, prev = [], 0
                    for x in xs:
                        segs.append(x + lap - prev)
                        prev = x
                    segs.append(last)
                    key = tuple(segs)
                    if key not in seen:
                        seen.add(key)
                        layouts.append((segs, [(x, x + lap) for x in xs]))
                return
            # đoạn hiện tại: x + lap − start ∈ [min_seg, usable]; phần còn lại phải phủ được
            max_rest = (remaining - 1) * usable - (remaining - 2) * lap
            lo = max(start + min_seg - lap, length - max_rest, start + 1)
            hi = min(start + usable - lap, length - min_seg)
            for x in allowed(lo, hi, spread=not xs)[: (12 if not xs else 2)]:
                extend(x, xs + [x], remaining - 1)

        k_min = max(2, math.ceil((length - lap) / (usable - lap)))
        for k in (k_min, k_min + 1):
            extend(0, [], k)
            if len(layouts) >= max_layouts // 2:
                break
        return layouts

    def _same_section(self, a: float, b: float, lap: int) -> bool:
        return abs(a - b) < self.SAME_SECTION_FACTOR * lap

    def _stagger_rows(
        self, options: List[_SplitOption], own: List[int], lap: int, cap: int,
    ) -> List[Tuple[List[int], int]]:
        """Mỗi mặt cắt (tâm mối nối): tổng số thanh có mối nối tại đó ≤ cap."""
        rows, seen = [], set()
        for j in own:
            for c in options[j].centers:
                members = tuple(sorted(
                    k for k in own if any(self._same_section(c, c2, lap) for c2 in options[k].centers)
                ))
                if members not in seen:
                    seen.add(members)
                    rows.append((list(members), cap))
        return rows

    def _initial_stagger(
        self, options: List[_SplitOption], own: List[int], lap: int, cap: int, quantity: int,
    ) -> Optional[Dict[int, int]]:
        """Lời giải khởi đầu: chọn các cách ghép có mối nối không cùng mặt cắt, chia đều số thanh."""
        chosen: List[int] = []
        for j in own:
            if all(not self._same_section(c, c2, lap)
                   for k in chosen for c in options[j].centers for c2 in options[k].centers):
                chosen.append(j)
            if len(chosen) * cap >= quantity:
                break
        if len(chosen) * cap < quantity:
            return None
        counts, left = {}, quantity
        for j in chosen:
            counts[j] = min(cap, left)
            left -= counts[j]
        return counts

    # ── GROUP SOLVER ───────────────────────────────────────────────────────

    def _solve_group(
        self,
        lengths: List[int],
        qty: List[int],
        n_whole: int,
        options: List[_SplitOption],
        ratio_cap: Dict[int, int],
        stagger_rows: List[Tuple[List[int], int]],
        initial_splits: List[int],
    ) -> Tuple[List[List[int]], List[int], str, int, str]:
        """
        Giải một nhóm. Item 0..n_whole-1 là thanh nguyên (qty), item còn lại là đoạn nối.
        Trả về (danh sách cây [chỉ số item], số lần dùng mỗi cách ghép, status, cận dưới, tên solver).
        """
        n = len(lengths)
        cap = self.usable_length_mm + self.kerf_mm        # mẹo kerf: mỗi đoạn chiếm l + kerf
        weights = [l + self.kerf_mm for l in lengths]
        split_ub = self._split_upper_bounds(options, ratio_cap, qty)

        material_lb = -(-sum(qty[i] * weights[i] for i in range(n_whole)) // cap)
        greedy_need = self._need(qty, options, initial_splits)
        greedy = self._greedy_patterns(weights, greedy_need, cap)
        greedy_bars = sum(m for _, m in greedy)

        def expand(chosen, splits):
            return self._expand(chosen, self._need(qty, options, splits), lengths)

        if greedy_bars <= material_lb:
            return expand(greedy, initial_splits), list(initial_splits), "OPTIMAL", material_lb, "Greedy"

        if not self.use_ortools:
            return (expand(greedy, initial_splits), list(initial_splits), "FEASIBLE", material_lb,
                    "Greedy (không OR-Tools)")

        try:
            deadline = time.monotonic() + self.time_limit_s
            patterns, lp_lb = self._column_generation(
                weights, qty, n_whole, cap, options, split_ub, ratio_cap, stagger_rows,
                [p for p, _ in greedy], deadline,
            )
            lower_bound = max(material_lb, lp_lb)
            if greedy_bars <= lower_bound:
                return (expand(greedy, initial_splits), list(initial_splits), "OPTIMAL", lower_bound,
                        "Greedy + LP bound")

            solved = self._solve_integer_master(
                patterns, qty, n_whole, options, split_ub, ratio_cap, stagger_rows, lower_bound,
                dict(greedy), initial_splits, max(1.0, deadline - time.monotonic()),
            )
        except (ImportError, RuntimeError):
            # OR-Tools không có hoặc LP lỗi: dùng heuristic, trạng thái vẫn so với cận dưới vật liệu
            return (expand(greedy, initial_splits), list(initial_splits), "FEASIBLE", material_lb,
                    "Greedy (OR-Tools lỗi)")

        if solved is None or sum(solved[0]) > greedy_bars or (
                sum(solved[0]) == greedy_bars and sum(solved[1]) >= sum(initial_splits)):
            chosen, splits = greedy, list(initial_splits)
        else:
            counts, splits = solved
            chosen = [(p, c) for p, c in zip(patterns, counts) if c > 0]

        bars = expand(chosen, splits)
        status = "OPTIMAL" if len(bars) <= lower_bound else "FEASIBLE"
        return bars, list(splits), status, lower_bound, "OR-Tools GLOP + CP-SAT"

    @staticmethod
    def _split_upper_bounds(options, ratio_cap, qty) -> List[int]:
        return [ratio_cap.get(o.owner, 0) if not o.mandatory else qty[o.owner] for o in options]

    @staticmethod
    def _need(qty: Sequence[int], options: List[_SplitOption], splits: Sequence[int]) -> List[int]:
        need = list(qty)
        for opt, s in zip(options, splits):
            if s:
                need[opt.owner] -= s
                for item in opt.seg_items:
                    need[item] += s
        return need

    def _item_bounds(self, weights, qty, n_whole, options, split_ub, cap) -> List[int]:
        bounds = [min(qty[i], cap // weights[i]) if i < n_whole else 0 for i in range(len(weights))]
        seg_need: Counter = Counter()
        for opt, ub in zip(options, split_ub):
            for item in opt.seg_items:
                seg_need[item] += ub
        for item, total in seg_need.items():
            bounds[item] = min(total, cap // weights[item])
        if self.max_pieces_per_bar:
            bounds = [min(b, self.max_pieces_per_bar) for b in bounds]
        return bounds

    def _greedy_patterns(
        self, weights: Sequence[int], qty: Sequence[int], cap: int
    ) -> List[Tuple[Tuple[int, ...], int]]:
        """
        Heuristic tuần tự: dựng một phương án kiểu First-Fit Decreasing từ số lượng còn lại
        (tôn trọng giới hạn số đoạn / số loại mỗi cây), lặp phương án đó nhiều nhất có thể,
        rồi dựng phương án tiếp theo. Trả về [(số đoạn mỗi loại, số cây lặp lại)].
        """
        n = len(weights)
        order = sorted(range(n), key=lambda i: -weights[i])
        max_p = self.max_pieces_per_bar or 10 ** 9
        max_t = self.max_marks_per_bar or 10 ** 9
        remaining = [q if weights[i] <= cap else 0 for i, q in enumerate(qty)]
        result: List[Tuple[Tuple[int, ...], int]] = []
        while any(remaining):
            counts = [0] * n
            free, pieces, types = cap, 0, 0
            for i in order:
                if remaining[i] and weights[i] <= free and pieces < max_p and types < max_t:
                    fit = min(remaining[i], free // weights[i], max_p - pieces)
                    counts[i] = fit
                    free -= fit * weights[i]
                    pieces += fit
                    types += 1
            repeat = min(remaining[i] // counts[i] for i in range(n) if counts[i])
            result.append((tuple(counts), repeat))
            for i in range(n):
                remaining[i] -= repeat * counts[i]
        return result

    def _column_generation(
        self,
        weights: Sequence[int],
        qty: Sequence[int],
        n_whole: int,
        cap: int,
        options: List[_SplitOption],
        split_ub: List[int],
        ratio_cap: Dict[int, int],
        stagger_rows: List[Tuple[List[int], int]],
        initial_patterns: List[Tuple[int, ...]],
        deadline: float,
        max_iterations: int = 2000,
    ) -> Tuple[List[Tuple[int, ...]], int]:
        """Gilmore–Gomory Column Generation. Trả về (tập phương án, cận dưới số cây)."""
        from ortools.linear_solver import pywraplp

        n = len(weights)
        bounds = self._item_bounds(weights, qty, n_whole, options, split_ub, cap)
        patterns: List[Tuple[int, ...]] = []
        seen = set()
        singles = [tuple(bounds[i] if k == i else 0 for k in range(n)) for i in range(n) if bounds[i] > 0]
        for p in list(initial_patterns) + singles:
            if p not in seen and any(p):
                seen.add(p)
                patterns.append(p)

        lp = pywraplp.Solver.CreateSolver("GLOP")
        if lp is None:
            raise ImportError("OR-Tools GLOP không khả dụng")
        infinity = lp.infinity()
        rows = [lp.Constraint(float(qty[i]) if i < n_whole else 0.0, infinity) for i in range(n)]
        objective = lp.Objective()
        objective.SetMinimization()

        # Biến số lần dùng mỗi cách ghép: thay 1 thanh nguyên bằng các đoạn nối
        ratio_rows = {owner: lp.Constraint(-infinity, float(limit)) for owner, limit in ratio_cap.items()}
        split_vars = []
        for opt, ub in zip(options, split_ub):
            s = lp.NumVar(0.0, float(ub), "")
            split_vars.append(s)
            rows[opt.owner].SetCoefficient(s, 1.0)
            for item, mult in Counter(opt.seg_items).items():
                rows[item].SetCoefficient(s, -float(mult))
            if opt.owner in ratio_rows:
                ratio_rows[opt.owner].SetCoefficient(s, 1.0)
        for members, limit in stagger_rows:
            row = lp.Constraint(-infinity, float(limit))
            for j in members:
                row.SetCoefficient(split_vars[j], 1.0)

        def add_column(p: Tuple[int, ...]) -> None:
            x = lp.NumVar(0.0, infinity, "")
            objective.SetCoefficient(x, 1.0)
            for i, a in enumerate(p):
                if a:
                    rows[i].SetCoefficient(x, float(a))

        for p in patterns:
            add_column(p)

        lp_value = float("inf")
        best_value = float("inf")
        for _ in range(max_iterations):
            if lp.Solve() != pywraplp.Solver.OPTIMAL:
                raise RuntimeError("LP master của Column Generation không giải được")
            lp_value = objective.Value()
            duals = [row.dual_value() for row in rows]
            best_value, new_pattern = _knapsack(
                weights, duals, bounds, cap, self.max_pieces_per_bar, self.max_marks_per_bar
            )
            if best_value <= 1.0 + 1e-9 or new_pattern in seen or time.monotonic() > deadline:
                break
            seen.add(new_pattern)
            patterns.append(new_pattern)
            add_column(new_pattern)

        # Hội tụ (best_value ≤ 1): lp_value là cận dưới LP. Chưa hội tụ: dùng cận Farley.
        lower_bound = math.ceil(lp_value / max(1.0, best_value) - 1e-6)
        return patterns, lower_bound

    @staticmethod
    def _solve_integer_master(
        patterns: List[Tuple[int, ...]],
        qty: Sequence[int],
        n_whole: int,
        options: List[_SplitOption],
        split_ub: List[int],
        ratio_cap: Dict[int, int],
        stagger_rows: List[Tuple[List[int], int]],
        lower_bound: int,
        hint: Dict[Tuple[int, ...], int],
        initial_splits: Sequence[int],
        time_limit_s: float,
    ) -> Optional[Tuple[List[int], List[int]]]:
        """CP-SAT: chọn số cây cho từng phương án (và số lần ghép nối), tối thiểu số cây rồi số mối nối."""
        from ortools.sat.python import cp_model

        n = len(qty)
        need = list(qty)
        for opt, ub in zip(options, split_ub):
            for item in opt.seg_items:
                need[item] += ub

        model = cp_model.CpModel()
        xs = []
        for j, p in enumerate(patterns):
            ub = max(-(-need[i] // p[i]) for i in range(n) if p[i])
            x = model.NewIntVar(0, max(ub, 0), f"x{j}")
            model.AddHint(x, hint.get(p, 0))
            xs.append(x)
        ss = []
        for k, (opt, ub) in enumerate(zip(options, split_ub)):
            s = model.NewIntVar(0, ub, f"s{k}")
            model.AddHint(s, initial_splits[k])
            ss.append(s)

        split_of: Dict[int, List] = defaultdict(list)
        seg_of: Dict[int, List] = defaultdict(list)
        for opt, s in zip(options, ss):
            split_of[opt.owner].append(s)
            for item in opt.seg_items:
                seg_of[item].append(s)
        for i in range(n):
            produced = sum(p[i] * xs[j] for j, p in enumerate(patterns) if p[i])
            if i < n_whole:
                if qty[i]:
                    model.Add(produced + sum(split_of[i]) >= qty[i])
            elif i in seg_of:
                model.Add(produced >= sum(seg_of[i]))
        for owner, limit in ratio_cap.items():
            model.Add(sum(split_of[owner]) <= limit)
        for members, limit in stagger_rows:
            model.Add(sum(ss[j] for j in members) <= limit)

        total = sum(xs)
        model.Add(total >= lower_bound)
        joints = sum(len(opt.laps) * s for opt, s in zip(options, ss))
        big = sum(len(opt.laps) * ub for opt, ub in zip(options, split_ub)) + 1
        model.Minimize(total * big + joints)   # ít cây nhất, rồi ít mối nối nhất

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit_s
        solver.parameters.num_search_workers = 1  # 1 luồng: cùng BBS → cùng phiếu cắt
        solver.parameters.random_seed = 0
        status = solver.Solve(model)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return None
        return [int(solver.Value(x)) for x in xs], [int(solver.Value(s)) for s in ss]

    @staticmethod
    def _expand(
        chosen: List[Tuple[Tuple[int, ...], int]],
        need: Sequence[int],
        lengths: Sequence[int],
    ) -> List[List[int]]:
        """Bung phương án thành từng cây (chỉ số item), bỏ bớt các đoạn cắt dư so với nhu cầu."""
        bars: List[List[int]] = []
        for p, count in chosen:
            items = [i for i, a in enumerate(p) for _ in range(a)]
            bars.extend(list(items) for _ in range(count))

        produced = Counter(i for b in bars for i in b)
        surplus = {i: produced[i] - need[i] for i in produced if produced[i] > need[i]}
        for bar in reversed(bars):
            if not surplus:
                break
            for i in list(surplus):
                while surplus.get(i) and i in bar:
                    bar.remove(i)
                    surplus[i] -= 1
                    if surplus[i] == 0:
                        del surplus[i]
        bars = [sorted(b, key=lambda i: (-lengths[i], i)) for b in bars if b]
        bars.sort(key=lambda b: (-sum(lengths[i] for i in b), [lengths[i] for i in b], b))
        return bars


def _knapsack(
    weights: Sequence[int],
    values: Sequence[float],
    bounds: Sequence[int],
    cap: int,
    max_pieces: Optional[int] = None,
    max_types: Optional[int] = None,
) -> Tuple[float, Tuple[int, ...]]:
    """Pricing của Column Generation: cái túi có giới hạn số lượng (và số đoạn / số loại mỗi cây)."""
    if max_pieces is None and max_types is None:
        return _bounded_knapsack(weights, values, bounds, cap)
    return _constrained_knapsack(weights, values, bounds, cap, max_pieces, max_types)


def _bounded_knapsack(
    weights: Sequence[int],
    values: Sequence[float],
    bounds: Sequence[int],
    cap: int,
) -> Tuple[float, Tuple[int, ...]]:
    """
    Bài toán cái túi có giới hạn số lượng (pricing của Column Generation).
    Quy hoạch động 0-1 trên các gói tách nhị phân, vector hóa bằng numpy.
    """
    import numpy as np

    n = len(weights)
    chunks: List[Tuple[int, int]] = []  # (loại đoạn, số đoạn trong gói)
    for i in range(n):
        if values[i] <= 1e-12 or bounds[i] <= 0 or weights[i] > cap:
            continue
        rest, size = bounds[i], 1
        while rest > 0:
            take = min(size, rest)
            chunks.append((i, take))
            rest -= take
            size *= 2

    best = np.zeros(cap + 1)
    taken = np.zeros((len(chunks), cap + 1), dtype=bool)
    for j, (i, count) in enumerate(chunks):
        w = count * weights[i]
        if w > cap:
            continue
        candidate = best[: cap + 1 - w] + count * values[i]
        better = candidate > best[w:] + 1e-12
        taken[j, w:] = better
        best[w:] = np.where(better, candidate, best[w:])

    pattern = [0] * n
    c = cap
    for j in range(len(chunks) - 1, -1, -1):
        if taken[j, c]:
            i, count = chunks[j]
            pattern[i] += count
            c -= count * weights[i]
    return float(best[cap]), tuple(pattern)


def _constrained_knapsack(
    weights: Sequence[int],
    values: Sequence[float],
    bounds: Sequence[int],
    cap: int,
    max_pieces: Optional[int],
    max_types: Optional[int],
) -> Tuple[float, Tuple[int, ...]]:
    """
    Cái túi có thêm giới hạn tổng số đoạn (max_pieces) và số loại đoạn (max_types) mỗi cây.
    Quy hoạch động theo trạng thái (số loại, số đoạn, chiều dài).
    """
    import numpy as np

    n = len(weights)
    T = (max_types + 1) if max_types else 1
    P = (max_pieces + 1) if max_pieces else 1
    neg = -1e18
    best = np.full((T, P, cap + 1), neg)
    best[0, 0, :] = 0.0
    items = [i for i in range(n) if values[i] > 1e-12 and bounds[i] > 0 and weights[i] <= cap]
    choice = np.zeros((len(items), T, P, cap + 1), dtype=np.int16)

    for idx, i in enumerate(items):
        old = best.copy()
        kmax = min(bounds[i], cap // weights[i], (max_pieces or 10 ** 9))
        for k in range(1, kmax + 1):
            w = k * weights[i]
            dt = 1 if max_types else 0
            dp = k if max_pieces else 0
            if dt >= T or dp >= P:
                break
            src = old[: T - dt, : P - dp, : cap + 1 - w] + k * values[i]
            dst = best[dt:, dp:, w:]
            better = src > dst + 1e-12
            best[dt:, dp:, w:] = np.where(better, src, dst)
            choice[idx, dt:, dp:, w:][better] = k

    t, p, c = np.unravel_index(int(np.argmax(best)), best.shape)
    value = float(best[t, p, c])
    pattern = [0] * n
    for idx in range(len(items) - 1, -1, -1):
        k = int(choice[idx, t, p, c])
        if k:
            i = items[idx]
            pattern[i] = k
            t -= 1 if max_types else 0
            p -= k if max_pieces else 0
            c -= k * weights[i]
    return max(value, 0.0), tuple(pattern)


def _pop_mark(items: List[List]) -> str:
    """Lấy ký hiệu thanh cho một đoạn cắt theo thứ tự trong BBS."""
    while items and items[0][1] <= 0:
        items.pop(0)
    if not items:
        return "?"
    items[0][1] -= 1
    return items[0][0]


def _grouped_cuts(p: CuttingPattern) -> List[Tuple[int, int, str]]:
    """Gộp các đoạn cùng chiều dài và ký hiệu trong một phương án: [(chiều dài, số đoạn, ký hiệu)]."""
    grouped: Dict[Tuple[int, str], int] = {}
    for length, mark in zip(p.cuts, p.marks or [""] * len(p.cuts)):
        grouped[(length, mark)] = grouped.get((length, mark), 0) + 1
    return [(length, n, mark) for (length, mark), n in grouped.items()]


def write_cut_plan_csv(solution: CuttingStockSolution, path: str) -> None:
    """Xuất phiếu cắt thép cho xưởng gia công (CSV UTF-8, mở được bằng Excel)."""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "STT", "Đường kính (mm)", "Mác thép", f"Số cây {solution.bar_length_mm / 1000:g}m",
            "Phương án cắt 1 cây (mm × số đoạn)", "Ký hiệu thanh theo chiều dài",
            "Số đoạn/cây", "Chiều dài dùng (mm)", "Dư cuối cây (mm)", "Phân loại đầu thừa",
        ])
        for idx, p in enumerate(solution.patterns, start=1):
            grouped = _grouped_cuts(p)
            w.writerow([
                idx, p.diameter_mm, p.grade, p.bars_used,
                " + ".join(f"{length}×{n}" for length, n, _ in grouped),
                "; ".join(f"{length}: {mark}" for length, _, mark in grouped),
                len(p.cuts), solution.bar_length_mm - 2 * solution.end_trim_mm - p.waste_mm,
                p.waste_mm, p.offcut_class,
            ])
        if solution.splices:
            w.writerow([])
            w.writerow(["MỐI NỐI — CẦN KỸ THUẬT DUYỆT", "Mã thanh ghép", "Loại", "Đường kính (mm)", "Bar Mark",
                        "Các đoạn (mm)", "Đoạn trước (mm)", "Đoạn sau (mm)", "Chồng nối (mm)",
                        "Vùng nối từ (mm)", "Vùng nối đến (mm)", "Cây đoạn trước / sau", "Kiểu nối"])
            for s in solution.splices:
                w.writerow([s["splice_id"], s["assembly_id"],
                            "Bắt buộc (thanh dài)" if s["mandatory"] else "Tận dụng đầu thừa",
                            s["diameter_mm"], s["mark"], " + ".join(map(str, s["segments_mm"])),
                            s["a_mm"], s["b_mm"], s["lap_mm"], s["lap_from_mm"], s["lap_to_mm"],
                            f"C{s['bar_a']} / C{s['bar_b']}", s["kind"]])
        w.writerow([])
        w.writerow([
            "TỔNG HỢP", "Đường kính (mm)", "Mác thép", "Số cây", "Cận dưới", "Trạng thái",
            "Số đoạn", "Đề-xê (%)", "KL mua (kg)", "Đầu thừa tái sử dụng", "Số mối nối",
        ])
        for g in solution.groups:
            w.writerow([
                "", g.diameter_mm, g.grade, g.total_bars_needed, g.lower_bound_bars, g.status,
                g.total_pieces, f"{g.waste_ratio_pct:.2f}", f"{g.purchased_weight_kg:.2f}",
                g.reusable_offcuts, g.splices,
            ])


# ─────────────────────────────────────────────────────────────────────────────
# INTER-AGENT VALIDATION — Kiểm tra vùng nối thép theo TCVN
# ─────────────────────────────────────────────────────────────────────────────

class SpliceZoneValidator:
    """
    Kiểm tra chéo (Inter-Agent Negotiation):
    BPTC_KCS Agent xác nhận vị trí nối thép của Rebar Agent.
    TCVN 5574:2018 §8.6: Không được nối thép tại vùng chịu kéo/cắt nguy hiểm.
    """

    # Vùng cấm nối (tính từ gối tính tỷ lệ chiều dài nhịp)
    FORBIDDEN_ZONES = {
        "TENSION":     "25-75% chiều dài nhịp (vùng chịu kéo dương của dầm đơn giản)",
        "SHEAR":       "0-15% và 85-100% chiều dài nhịp (vùng chịu cắt nguy hiểm gần gối)",
        "COMPRESSION": "Được phép nối nhưng phải so le ≥ 1.3 × L_nối",
    }

    def validate(
        self,
        splice_positions: List[Dict],
        span_length_mm: int = 38_200
    ) -> Tuple[str, List[str]]:
        """
        splice_positions: [{"mark": "T1", "splice_at_mm": 12000, "zone": "TENSION"}, ...]
        Trả về: (status, violations)
        status = "PASS" hoặc "REJECT"
        """
        violations = []
        for sp in splice_positions:
            mark = sp.get("mark", "?")
            pos = sp.get("splice_at_mm", 0)
            zone = sp.get("zone", "").upper()
            pos_ratio = pos / span_length_mm if span_length_mm else 0

            # Kiểm tra vùng chịu kéo: 25-75% nhịp
            if zone == "TENSION" and 0.25 <= pos_ratio <= 0.75:
                violations.append(
                    f"[REJECT] Thanh {mark}: Nối tại {pos}mm ({pos_ratio:.0%} nhịp) "
                    f"— vi phạm TCVN 5574:2018 §8.6 (vùng chịu kéo nguy hiểm)"
                )

            # Kiểm tra vùng chịu cắt: 0-15% và 85-100% nhịp
            if zone == "SHEAR" and (pos_ratio <= 0.15 or pos_ratio >= 0.85):
                violations.append(
                    f"[REJECT] Thanh {mark}: Nối tại {pos}mm ({pos_ratio:.0%} nhịp) "
                    f"— vi phạm TCVN 5574:2018 §8.6 (vùng chịu cắt gần gối)"
                )

        status = "REJECT" if violations else "PASS"
        return status, violations


# ─────────────────────────────────────────────────────────────────────────────
# QUICK TEST
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    demands = [
        CutDemand(length_mm=4500, quantity=20, diameter_mm=20, mark="T1"),
        CutDemand(length_mm=3200, quantity=35, diameter_mm=20, mark="T2"),
        CutDemand(length_mm=2800, quantity=15, diameter_mm=16, mark="D1"),
        CutDemand(length_mm=6000, quantity=10, diameter_mm=25, mark="CU1"),
    ]

    solver = CuttingStockSolver()
    sol = solver.solve(demands)

    print(f"\n{'='*60}")
    print(f"  CUTTING STOCK RESULT — {sol.solver_name}")
    print(f"{'='*60}")
    print(f"  Status        : {sol.status}")
    print(f"  Số cây thép   : {sol.total_bars_needed} cây (cận dưới {sol.lower_bound_bars})")
    print(f"  Tổng đề-xê    : {sol.total_waste_mm:,} mm")
    print(f"  Tỷ lệ đề-xê   : {sol.waste_ratio_pct:.2f}%")
    print(f"  KL thép mua   : {sol.total_weight_kg:,.1f} kg")
    for g in sol.groups:
        print(f"    Ø{g.diameter_mm}: {g.total_bars_needed} cây, đề-xê {g.waste_ratio_pct:.2f}% [{g.status}]")
    if sol.warning:
        print(f"  ⚠ {sol.warning}")

    print(f"\n  Phương án cắt:")
    for p in sol.patterns:
        print(f"    Ø{p.diameter_mm} × {p.bars_used} cây: {p.cuts} — dư {p.waste_mm}mm ({p.offcut_class})")
