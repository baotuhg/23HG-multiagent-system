# -*- coding: utf-8 -*-
"""Kiểm thử CuttingStockSolver — chạy: python -m unittest discover tests"""

import os
import random
import sys
import tempfile
import unittest
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tools.cutting_stock_solver import (
    CuttingStockSolver, CutDemand, rebar_unit_weight_kg_m, write_cut_plan_csv,
)


def random_demands(seed, diameters=(12, 16, 20), types=(5, 25)):
    rnd = random.Random(seed)
    demands = []
    for d in diameters:
        for _ in range(rnd.randint(*types)):
            demands.append(CutDemand(rnd.randint(300, 8000), rnd.randint(1, 60), d,
                                     f"M{len(demands)}", "CB400-V"))
    return demands


class CuttingStockSolverTest(unittest.TestCase):

    def assert_valid_plan(self, sol, demands, bar_length=11_700, kerf=3):
        for bar in sol.assignment:
            used = sum(bar["cuts_mm"]) + kerf * (len(bar["cuts_mm"]) - 1)
            self.assertEqual(bar["used_mm"], used)
            self.assertLessEqual(used, bar_length)
        cut = Counter((b["diameter_mm"], b["grade"], c) for b in sol.assignment for c in b["cuts_mm"])
        wanted = Counter()
        for d in demands:
            wanted[(d.diameter_mm, d.grade, d.length_mm)] += d.quantity
        self.assertEqual(cut, wanted, "phải cắt đúng và đủ số đoạn theo BBS")
        self.assertGreaterEqual(sol.total_bars_needed, sol.lower_bound_bars)

    def test_never_mixes_diameters_or_grades(self):
        demands = [
            CutDemand(2800, 15, 16, "D1", "CB400-V"),
            CutDemand(2800, 15, 16, "D1b", "CB300-V"),
            CutDemand(4500, 20, 20, "T1", "CB400-V"),
            CutDemand(3200, 35, 12, "T2", "CB400-V"),
        ]
        sol = CuttingStockSolver().solve(demands)
        self.assert_valid_plan(sol, demands)
        self.assertEqual(len(sol.groups), 4)
        for bar in sol.assignment:
            lengths = {d.length_mm for d in demands
                       if d.diameter_mm == bar["diameter_mm"] and d.grade == bar["grade"]}
            self.assertTrue(set(bar["cuts_mm"]) <= lengths)

    def test_finds_optimum_where_ffd_does_not(self):
        # FFD dùng 3 cây; tối ưu là 2 cây: 50+25+25 và 34+33+33
        demands = [CutDemand(50, 1, 10), CutDemand(34, 1, 10), CutDemand(33, 2, 10), CutDemand(25, 2, 10)]
        sol = CuttingStockSolver(bar_length_mm=100, kerf_mm=0).solve(demands)
        self.assertEqual(sol.status, "OPTIMAL")
        self.assertEqual(sol.total_bars_needed, 2)
        self.assert_valid_plan(sol, demands, bar_length=100, kerf=0)

    def test_kerf_is_accounted(self):
        # 2 × 5849 + 1 nhát cắt 3mm = 11701 > 11700 → không ghép được vào 1 cây
        self.assertEqual(CuttingStockSolver().solve([CutDemand(5849, 2, 16)]).total_bars_needed, 2)
        self.assertEqual(CuttingStockSolver().solve([CutDemand(5848, 2, 16)]).total_bars_needed, 1)

    def test_weight_is_computed_per_diameter(self):
        sol = CuttingStockSolver().solve([CutDemand(11_700, 2, 20), CutDemand(11_700, 1, 12)])
        expected = 2 * 11.7 * 2.466 + 1 * 11.7 * 0.888
        self.assertAlmostEqual(sol.total_weight_kg, expected, places=1)
        self.assertAlmostEqual(rebar_unit_weight_kg_m(40), 9.865, places=2)

    def test_random_instances_are_valid(self):
        for seed in range(5):
            demands = random_demands(seed)
            sol = CuttingStockSolver(time_limit_s=5).solve(demands)
            self.assertIn(sol.status, ("OPTIMAL", "FEASIBLE"))
            self.assert_valid_plan(sol, demands)

    def test_greedy_fallback_is_valid(self):
        demands = random_demands(42)
        sol = CuttingStockSolver(use_ortools=False).solve(demands)
        self.assert_valid_plan(sol, demands)

    def test_deterministic(self):
        demands = random_demands(7)
        plans = [
            [(p.diameter_mm, p.cuts, p.bars_used) for p in CuttingStockSolver(time_limit_s=5).solve(demands).patterns]
            for _ in range(2)
        ]
        self.assertEqual(plans[0], plans[1])

    def test_rejects_bad_input(self):
        self.assertEqual(CuttingStockSolver().solve([CutDemand(12_000, 1, 16)]).status, "INFEASIBLE")
        self.assertEqual(CuttingStockSolver().solve([CutDemand(0, 1, 16)]).status, "ERROR")
        self.assertEqual(CuttingStockSolver().solve([CutDemand(1000, -1, 16)]).status, "ERROR")
        self.assertEqual(CuttingStockSolver().solve([]).status, "INFEASIBLE")

    def test_cut_plan_csv(self):
        sol = CuttingStockSolver().solve([CutDemand(3000, 7, 16, "A1"), CutDemand(2000, 3, 16, "A2")])
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "plan.csv")
            write_cut_plan_csv(sol, path)
            with open(path, encoding="utf-8-sig") as f:
                text = f.read()
        self.assertIn("3000×", text)
        self.assertIn("3000: A1", text)
        self.assertIn("OPTIMAL", text)


if __name__ == "__main__":
    unittest.main()
