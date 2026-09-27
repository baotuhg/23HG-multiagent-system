# -*- coding: utf-8 -*-
"""
OR-TOOLS 1D CUTTING STOCK SOLVER
Bài toán cắt thép một chiều (1D Cutting Stock Problem)
Tối ưu ghép các đoạn cắt vào cây thép nguyên 11.7m — mục tiêu: ít cây thép nhất.

Đây là Pure Python Tool — KHÔNG có bất kỳ LLM nào tham gia tính toán.
LLM chỉ được phép: gọi tool này, đọc kết quả, trình bày ngôn ngữ tự nhiên.

Nguyên tắc thực tế xưởng gia công:
  - Chỉ ghép các đoạn CÙNG đường kính và CÙNG mác thép vào một cây thép.
  - Mỗi nhát cắt mất một lượng bằng bề rộng lưỡi cắt (kerf).

Thuật toán (chạy riêng cho từng nhóm đường kính + mác thép):
  1. Heuristic tham lam theo phương án cắt → lời giải khởi đầu (và dự phòng khi
     không có OR-Tools).
  2. Column Generation Gilmore–Gomory: LP (OR-Tools GLOP) + bài toán cái túi
     (quy hoạch động) sinh phương án cắt → cận dưới số cây thép.
  3. OR-Tools CP-SAT chọn số lần lặp nguyên cho từng phương án cắt.

Trạng thái lời giải:
  OPTIMAL  — số cây thép bằng cận dưới → đã chứng minh là ít nhất có thể.
  FEASIBLE — phương án hợp lệ nhưng chưa chứng minh tối ưu (xem lower_bound_bars).

Input:  demands = [CutDemand(length_mm, quantity, diameter_mm, mark, grade), ...]
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
REUSABLE_OFFCUT_MIN_MM = 1_000     # Mẩu thừa ≥ 1m: nhập kho để tận dụng (thép đai, thép cấu tạo)

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


@dataclass
class CuttingPattern:
    """Một phương án ghép các đoạn cắt vào 1 cây thép, lặp lại `bars_used` cây."""
    cuts: List[int] = field(default_factory=list)  # list chiều dài (mm)
    waste_mm: int = 0
    bars_used: int = 1
    diameter_mm: int = 0
    grade: str = ""
    marks: List[str] = field(default_factory=list)  # ký hiệu tương ứng từng đoạn cắt


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
    reusable_offcuts: int = 0            # số mẩu thừa ≥ REUSABLE_OFFCUT_MIN_MM
    reusable_offcut_length_mm: int = 0


@dataclass
class CuttingStockSolution:
    status: str = "NOT_RUN"          # OPTIMAL / FEASIBLE / INFEASIBLE / ERROR
    solver_name: str = "OR-Tools GLOP + CP-SAT"
    bar_length_mm: int = STANDARD_BAR_LENGTH_MM
    kerf_mm: int = SAW_KERF_MM
    total_bars_needed: int = 0
    lower_bound_bars: int = 0        # Cận dưới số cây (tổng các nhóm)
    total_pieces: int = 0
    total_length_used_mm: int = 0
    total_waste_mm: int = 0
    waste_ratio_pct: float = 0.0
    total_weight_kg: float = 0.0     # Khối lượng thép cây phải mua (số cây × 11.7m × kg/m)
    net_weight_kg: float = 0.0       # Khối lượng các đoạn thành phẩm theo BBS
    reusable_offcuts: int = 0        # Số mẩu thừa ≥ 1m có thể nhập kho tận dụng
    reusable_offcut_length_mm: int = 0
    patterns: List[CuttingPattern] = field(default_factory=list)  # phương án khác nhau
    assignment: List[Dict] = field(default_factory=list)          # chi tiết từng cây
    groups: List[DiameterGroupResult] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    warning: str = ""                # Gộp warnings (giữ tương thích)


# ─────────────────────────────────────────────────────────────────────────────
# SOLVER
# ─────────────────────────────────────────────────────────────────────────────

class CuttingStockSolver:
    """
    1D Cutting Stock Solver — tách nhóm theo (đường kính, mác thép), mỗi nhóm:
      Heuristic tham lam → Column Generation (GLOP) → CP-SAT nghiệm nguyên.
    """

    def __init__(
        self,
        bar_length_mm: int = STANDARD_BAR_LENGTH_MM,
        kerf_mm: int = SAW_KERF_MM,
        time_limit_s: float = 20.0,
        use_ortools: bool = True,
    ):
        self.bar_length_mm = bar_length_mm
        self.kerf_mm = kerf_mm
        self.time_limit_s = time_limit_s  # giới hạn thời gian cho mỗi nhóm đường kính
        self.use_ortools = use_ortools

    # ── PUBLIC API ─────────────────────────────────────────────────────────

    def solve(self, demands: List[CutDemand]) -> CuttingStockSolution:
        sol = CuttingStockSolution(bar_length_mm=self.bar_length_mm, kerf_mm=self.kerf_mm)

        invalid = [d for d in demands if d.length_mm <= 0 or d.quantity < 0]
        if invalid:
            sol.status = "ERROR"
            sol.warning = (
                f"Có {len(invalid)} dòng chiều dài ≤ 0 hoặc số lượng âm: "
                f"{[(d.mark, d.length_mm, d.quantity) for d in invalid[:5]]}"
            )
            sol.warnings.append(sol.warning)
            return sol

        active = [d for d in demands if d.quantity > 0]
        if not active:
            sol.status = "INFEASIBLE"
            sol.warning = "Danh sách cần cắt rỗng"
            sol.warnings.append(sol.warning)
            return sol

        too_long = [d.length_mm for d in active if d.length_mm > self.bar_length_mm]
        if too_long:
            sol.status = "INFEASIBLE"
            sol.warning = f"Có {len(too_long)} đoạn dài hơn cây {self.bar_length_mm}mm: {too_long[:5]}"
            sol.warnings.append(sol.warning)
            return sol

        # Gom nhóm: không bao giờ ghép khác đường kính / khác mác thép vào một cây
        group_demand: Dict[Tuple[int, str], Dict[int, int]] = defaultdict(lambda: defaultdict(int))
        group_marks: Dict[Tuple[int, str], Dict[int, List[List]]] = defaultdict(lambda: defaultdict(list))
        for d in active:
            key = (int(d.diameter_mm or 0), (d.grade or "").strip())
            length = int(d.length_mm)
            group_demand[key][length] += int(d.quantity)
            group_marks[key][length].append([d.mark or "?", int(d.quantity)])

        if any(diameter == 0 for diameter, _ in group_demand):
            sol.warnings.append(
                "Có đoạn thép không ghi đường kính — được ghép chung một nhóm riêng "
                "và không tính được khối lượng"
            )

        solver_names = set()
        bar_counter = 0
        for key in sorted(group_demand):
            diameter, grade = key
            demand = dict(group_demand[key])
            bars, status, lower_bound, solver_name = self._solve_group(demand)
            solver_names.add(solver_name)

            unit_w = rebar_unit_weight_kg_m(diameter)
            pieces = sum(demand.values())
            net_len = sum(l * q for l, q in demand.items())
            offcuts: List[int] = []
            mark_queue = {l: [list(m) for m in ms] for l, ms in group_marks[key].items()}
            unique_marks = {
                l: "|".join(dict.fromkeys(m[0] for m in ms))
                for l, ms in group_marks[key].items()
            }

            for cuts in bars:
                used = sum(cuts) + self.kerf_mm * max(len(cuts) - 1, 0)
                waste = self.bar_length_mm - used
                offcuts.append(waste)
                bar_counter += 1
                sol.assignment.append({
                    "bar_id": bar_counter,
                    "diameter_mm": diameter,
                    "grade": grade,
                    "cuts_mm": cuts,
                    "marks": [_pop_mark(mark_queue, c) for c in cuts],
                    "used_mm": used,
                    "waste_mm": waste,
                })

            for cuts, count in Counter(tuple(b) for b in bars).most_common():
                used = sum(cuts) + self.kerf_mm * max(len(cuts) - 1, 0)
                sol.patterns.append(CuttingPattern(
                    cuts=list(cuts),
                    waste_mm=self.bar_length_mm - used,
                    bars_used=count,
                    diameter_mm=diameter,
                    grade=grade,
                    marks=[unique_marks[c] for c in cuts],
                ))

            capacity = len(bars) * self.bar_length_mm
            group_waste = sum(offcuts)
            reusable = [w for w in offcuts if w >= REUSABLE_OFFCUT_MIN_MM]
            sol.groups.append(DiameterGroupResult(
                diameter_mm=diameter,
                grade=grade,
                status=status,
                solver_name=solver_name,
                total_pieces=pieces,
                total_bars_needed=len(bars),
                lower_bound_bars=lower_bound,
                total_waste_mm=group_waste,
                waste_ratio_pct=(group_waste / capacity * 100) if capacity else 0.0,
                purchased_weight_kg=round(capacity / 1000 * unit_w, 2),
                net_weight_kg=round(net_len / 1000 * unit_w, 2),
                reusable_offcuts=len(reusable),
                reusable_offcut_length_mm=sum(reusable),
            ))

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
        sol.status = "OPTIMAL" if all(g.status == "OPTIMAL" for g in sol.groups) else "FEASIBLE"
        sol.solver_name = ", ".join(sorted(solver_names))

        if sol.waste_ratio_pct > MAX_WASTE_RATIO_PCT:
            if sol.status == "OPTIMAL":
                sol.warnings.append(
                    f"Đề-xê {sol.waste_ratio_pct:.2f}% > {MAX_WASTE_RATIO_PCT}% nhưng số cây đã bằng "
                    f"cận dưới (tối ưu đã chứng minh) — hao hụt do chiều dài thanh trong BBS. "
                    f"Muốn giảm phải điều chỉnh chiều dài/vị trí nối hoặc tận dụng mẩu thừa "
                    f"({sol.reusable_offcuts} mẩu ≥ {REUSABLE_OFFCUT_MIN_MM}mm, "
                    f"tổng {sol.reusable_offcut_length_mm / 1000:,.1f} m)."
                )
            else:
                sol.warnings.append(
                    f"Đề-xê {sol.waste_ratio_pct:.2f}% > {MAX_WASTE_RATIO_PCT}%; phương án chưa chứng minh "
                    f"tối ưu (dùng {sol.total_bars_needed} cây, cận dưới {sol.lower_bound_bars} cây)."
                )
        sol.warning = " | ".join(sol.warnings)
        return sol

    # ── GROUP SOLVER ───────────────────────────────────────────────────────

    def _solve_group(self, demand: Dict[int, int]) -> Tuple[List[List[int]], str, int, str]:
        """
        Giải một nhóm cùng đường kính + mác thép.
        Trả về (danh sách cây [các đoạn cắt], status, cận dưới số cây, tên solver).
        """
        lengths = sorted(demand, reverse=True)
        qty = [demand[l] for l in lengths]
        cap = self.bar_length_mm + self.kerf_mm          # mẹo kerf: mỗi đoạn chiếm l + kerf
        weights = [l + self.kerf_mm for l in lengths]

        material_lb = -(-sum(q * w for q, w in zip(qty, weights)) // cap)
        greedy = self._greedy_patterns(weights, qty, cap)
        greedy_bars = sum(m for _, m in greedy)

        if greedy_bars <= material_lb:
            return self._expand(greedy, lengths, qty), "OPTIMAL", material_lb, "Greedy"

        if not self.use_ortools:
            return self._expand(greedy, lengths, qty), "FEASIBLE", material_lb, "Greedy (không OR-Tools)"

        try:
            deadline = time.monotonic() + self.time_limit_s
            patterns, lp_lb = self._column_generation(
                weights, qty, cap, [p for p, _ in greedy], deadline
            )
            lower_bound = max(material_lb, lp_lb)
            if greedy_bars <= lower_bound:
                return self._expand(greedy, lengths, qty), "OPTIMAL", lower_bound, "Greedy + LP bound"

            counts = self._solve_integer_master(
                patterns, qty, lower_bound, dict(greedy), max(1.0, deadline - time.monotonic())
            )
        except (ImportError, RuntimeError):
            # OR-Tools không có hoặc LP lỗi: dùng heuristic, trạng thái vẫn so với cận dưới vật liệu
            return self._expand(greedy, lengths, qty), "FEASIBLE", material_lb, "Greedy (OR-Tools lỗi)"

        if counts is None or sum(counts) >= greedy_bars:
            chosen, bars_used = greedy, greedy_bars
        else:
            chosen = [(p, c) for p, c in zip(patterns, counts) if c > 0]
            bars_used = sum(counts)

        status = "OPTIMAL" if bars_used <= lower_bound else "FEASIBLE"
        return self._expand(chosen, lengths, qty), status, lower_bound, "OR-Tools GLOP + CP-SAT"

    @staticmethod
    def _greedy_patterns(
        weights: Sequence[int], qty: Sequence[int], cap: int
    ) -> List[Tuple[Tuple[int, ...], int]]:
        """
        Heuristic tuần tự: dựng một phương án kiểu First-Fit Decreasing từ số lượng còn lại,
        lặp phương án đó nhiều nhất có thể, rồi dựng phương án tiếp theo.
        `weights` đã sắp giảm dần. Trả về [(số đoạn mỗi loại, số cây lặp lại)].
        """
        n = len(weights)
        remaining = list(qty)
        result: List[Tuple[Tuple[int, ...], int]] = []
        while any(remaining):
            counts = [0] * n
            free = cap
            for i in range(n):
                if remaining[i] and weights[i] <= free:
                    fit = min(remaining[i], free // weights[i])
                    counts[i] = fit
                    free -= fit * weights[i]
            repeat = min(remaining[i] // counts[i] for i in range(n) if counts[i])
            result.append((tuple(counts), repeat))
            for i in range(n):
                remaining[i] -= repeat * counts[i]
        return result

    def _column_generation(
        self,
        weights: Sequence[int],
        qty: Sequence[int],
        cap: int,
        initial_patterns: List[Tuple[int, ...]],
        deadline: float,
        max_iterations: int = 2000,
    ) -> Tuple[List[Tuple[int, ...]], int]:
        """Gilmore–Gomory Column Generation. Trả về (tập phương án, cận dưới số cây)."""
        from ortools.linear_solver import pywraplp

        n = len(weights)
        bounds = [min(qty[i], cap // weights[i]) for i in range(n)]
        patterns: List[Tuple[int, ...]] = []
        seen = set()
        singles = [tuple(bounds[i] if k == i else 0 for k in range(n)) for i in range(n)]
        for p in list(initial_patterns) + singles:
            if p not in seen:
                seen.add(p)
                patterns.append(p)

        lp = pywraplp.Solver.CreateSolver("GLOP")
        if lp is None:
            raise ImportError("OR-Tools GLOP không khả dụng")
        infinity = lp.infinity()
        rows = [lp.Constraint(float(qty[i]), infinity) for i in range(n)]
        objective = lp.Objective()
        objective.SetMinimization()

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
            best_value, new_pattern = _bounded_knapsack(weights, duals, bounds, cap)
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
        lower_bound: int,
        hint: Dict[Tuple[int, ...], int],
        time_limit_s: float,
    ) -> Optional[List[int]]:
        """CP-SAT: chọn số cây cho từng phương án, tối thiểu tổng số cây."""
        from ortools.sat.python import cp_model

        n = len(qty)
        model = cp_model.CpModel()
        xs = []
        for j, p in enumerate(patterns):
            ub = max(-(-qty[i] // p[i]) for i in range(n) if p[i])
            x = model.NewIntVar(0, ub, f"x{j}")
            model.AddHint(x, hint.get(p, 0))
            xs.append(x)
        for i in range(n):
            model.Add(sum(p[i] * xs[j] for j, p in enumerate(patterns) if p[i]) >= qty[i])
        total = sum(xs)
        model.Add(total >= lower_bound)
        model.Minimize(total)

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit_s
        solver.parameters.num_search_workers = 1  # 1 luồng: cùng BBS → cùng phiếu cắt
        solver.parameters.random_seed = 0
        status = solver.Solve(model)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return None
        return [int(solver.Value(x)) for x in xs]

    @staticmethod
    def _expand(
        chosen: List[Tuple[Tuple[int, ...], int]],
        lengths: Sequence[int],
        qty: Sequence[int],
    ) -> List[List[int]]:
        """Bung phương án thành từng cây, bỏ bớt các đoạn cắt dư so với nhu cầu."""
        bars: List[List[int]] = []
        for p, count in chosen:
            cuts = [lengths[i] for i, a in enumerate(p) for _ in range(a)]
            bars.extend(list(cuts) for _ in range(count))

        produced = Counter(c for b in bars for c in b)
        surplus = {l: produced[l] - q for l, q in zip(lengths, qty) if produced[l] > q}
        for bar in reversed(bars):
            if not surplus:
                break
            for l in list(surplus):
                while surplus.get(l) and l in bar:
                    bar.remove(l)
                    surplus[l] -= 1
                    if surplus[l] == 0:
                        del surplus[l]
        bars = [b for b in bars if b]
        bars.sort(key=lambda b: (-sum(b), b))
        return bars


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


def _pop_mark(queue: Dict[int, List[List]], length: int) -> str:
    """Lấy ký hiệu thanh cho một đoạn cắt theo thứ tự trong BBS."""
    items = queue.get(length, [])
    while items and items[0][1] <= 0:
        items.pop(0)
    if not items:
        return "?"
    items[0][1] -= 1
    return items[0][0]


def _grouped_cuts(p: CuttingPattern) -> List[Tuple[int, int, str]]:
    """Gộp các đoạn cùng chiều dài trong một phương án: [(chiều dài, số đoạn, ký hiệu)]."""
    grouped: Dict[int, List] = {}
    for length, mark in zip(p.cuts, p.marks or [""] * len(p.cuts)):
        grouped.setdefault(length, [0, mark])[0] += 1
    return [(length, n, mark) for length, (n, mark) in grouped.items()]


def write_cut_plan_csv(solution: CuttingStockSolution, path: str) -> None:
    """Xuất phiếu cắt thép cho xưởng gia công (CSV UTF-8, mở được bằng Excel)."""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "STT", "Đường kính (mm)", "Mác thép", f"Số cây {solution.bar_length_mm / 1000:g}m",
            "Phương án cắt 1 cây (mm × số đoạn)", "Ký hiệu thanh theo chiều dài",
            "Số đoạn/cây", "Chiều dài dùng (mm)", "Đề-xê/cây (mm)",
        ])
        for idx, p in enumerate(solution.patterns, start=1):
            grouped = _grouped_cuts(p)
            w.writerow([
                idx, p.diameter_mm, p.grade, p.bars_used,
                " + ".join(f"{length}×{n}" for length, n, _ in grouped),
                "; ".join(f"{length}: {mark}" for length, _, mark in grouped),
                len(p.cuts), solution.bar_length_mm - p.waste_mm, p.waste_mm,
            ])
        w.writerow([])
        w.writerow([
            "TỔNG HỢP", "Đường kính (mm)", "Mác thép", "Số cây", "Cận dưới", "Trạng thái",
            "Số đoạn", "Đề-xê (%)", "KL mua (kg)", f"Mẩu thừa ≥{REUSABLE_OFFCUT_MIN_MM}mm",
        ])
        for g in solution.groups:
            w.writerow([
                "", g.diameter_mm, g.grade, g.total_bars_needed, g.lower_bound_bars, g.status,
                g.total_pieces, f"{g.waste_ratio_pct:.2f}", f"{g.purchased_weight_kg:.2f}",
                g.reusable_offcuts,
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
        print(f"    Ø{p.diameter_mm} × {p.bars_used} cây: {p.cuts} — đề-xê {p.waste_mm}mm")
