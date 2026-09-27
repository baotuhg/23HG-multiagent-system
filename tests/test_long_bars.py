# -*- coding: utf-8 -*-
"""
Kiểm thử tách thanh dài hơn cây thép thành các đoạn nối (bắt buộc nối):
hình học, vùng cho phép nối, so le mối nối theo từng dòng BBS, thanh không tách được.
Chạy: python -m unittest discover tests
"""

import contextlib
import io
import os
import sys
import tempfile
import unittest
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.rebar_agent import RebarAgent
from core.state.shared_state import ProjectPhase
from core.supervisor.supervisor_agent import AECSupervisor
from tools.cutting_stock_solver import CuttingStockSolver, CutDemand


def assemblies(sol):
    out = defaultdict(list)
    for s in sol.splices:
        out[s["assembly_id"]].append(s)
    return out


class LongBarTest(unittest.TestCase):

    def check(self, sol, demands, zones, ratio=0.5, usable=11_700):
        asm = assemblies(sol)
        by_row = defaultdict(list)
        for joints in asm.values():
            j0 = joints[0]
            segs, lap, length = j0["segments_mm"], j0["lap_mm"], j0["bar_length_mm"]
            self.assertEqual(len(joints), len(segs) - 1)
            self.assertEqual(sum(segs) - lap * (len(segs) - 1), length)
            for seg, bar in zip(segs, j0["segment_bars"]):
                self.assertLessEqual(seg, usable)
                self.assertGreaterEqual(seg, lap)
                self.assertIn(seg, sol.assignment[bar - 1]["cuts_mm"])
            for s in joints:
                self.assertTrue(any(z0 * length - 1e-6 <= s["lap_from_mm"] and s["lap_to_mm"] <= z1 * length + 1e-6
                                    for z0, z1 in zones))
            by_row[(j0["mark"], length)].append([(s["lap_from_mm"] + s["lap_to_mm"]) / 2 for s in joints])
        for d in demands:
            if d.length_mm <= usable or any(u["mark"] == d.mark for u in sol.unplanned):
                continue
            rows = by_row[(d.mark, d.length_mm)]
            self.assertEqual(len(rows), d.quantity, d.mark)
            lap = 40 * d.diameter_mm
            cap = max(1, int(d.quantity * ratio))
            for centers in rows:
                for c in centers:
                    same = sum(1 for other in rows if any(abs(c - c2) < 1.3 * lap for c2 in other))
                    self.assertLessEqual(same, cap, f"{d.mark}: {same} thanh nối cùng mặt cắt")
        for bar in sol.assignment:
            self.assertLessEqual(sum(bar["cuts_mm"]) + 3 * (len(bar["cuts_mm"]) - 1), usable)

    def test_two_segment_bars_are_staggered(self):
        zones = [(0.0, 1.0)]
        demands = [CutDemand(13_571, 40, 18, "F2"), CutDemand(6_336, 83, 18, "F3")]
        sol = CuttingStockSolver(default_splice_zones=zones).solve(demands, split_long_bars=True)
        self.assertEqual(sol.long_bars_split, 40)
        self.assertEqual(sol.total_splices, 40)
        self.assertEqual(sol.unplanned, [])
        self.check(sol, demands, zones)

    def test_very_long_bar_needs_several_splices(self):
        zones = [(0.0, 1.0)]
        demands = [CutDemand(39_850, 10, 14, "S1")]
        sol = CuttingStockSolver(default_splice_zones=zones).solve(demands, split_long_bars=True)
        self.assertEqual(sol.long_bars_split, 10)
        self.assertTrue(all(len(s["segments_mm"]) == 4 for s in sol.splices))
        self.check(sol, demands, zones)

    def test_restricted_zones(self):
        zones = [(0.0, 0.3), (0.7, 1.0)]
        demands = [CutDemand(13_571, 40, 18, "F2"), CutDemand(16_916, 12, 16, "G1"), CutDemand(4_000, 30, 16, "K")]
        sol = CuttingStockSolver(default_splice_zones=zones).solve(demands, split_long_bars=True)
        self.check(sol, demands, zones)
        self.assertEqual(sol.long_bars_split, 52)

    def test_bar_without_zone_is_reported_not_dropped_silently(self):
        demands = [CutDemand(13_571, 40, 18, "F2"), CutDemand(4_000, 10, 18, "K")]
        sol = CuttingStockSolver().solve(demands, split_long_bars=True)
        self.assertEqual([u["mark"] for u in sol.unplanned], ["F2"])
        self.assertTrue(any("CHƯA có trong kế hoạch cắt" in w for w in sol.warnings))
        self.assertEqual(sum(len(a["cuts_mm"]) for a in sol.assignment), 10)

    def test_impossible_layout_is_reported(self):
        # 39.85m cần 3 mối nối nhưng chỉ cho nối ở 2 đầu thanh
        sol = CuttingStockSolver(default_splice_zones=[(0.0, 0.3), (0.7, 1.0)]).solve(
            [CutDemand(39_850, 4, 14, "S1"), CutDemand(4_000, 10, 14, "K")], split_long_bars=True)
        self.assertEqual([u["mark"] for u in sol.unplanned], ["S1"])

    def test_same_mark_in_two_rows_staggered_separately(self):
        zones = [(0.0, 1.0)]
        demands = [CutDemand(13_571, 20, 18, "F2"), CutDemand(13_571, 20, 18, "F2")]
        sol = CuttingStockSolver(default_splice_zones=zones).solve(demands, split_long_bars=True)
        self.assertEqual(sol.long_bars_split, 40)
        # mỗi dòng 20 thanh → mỗi mặt cắt tối đa 10 thanh của dòng đó
        asm = assemblies(sol)
        self.assertEqual(len(asm), 40)

    def test_without_flag_long_bars_are_infeasible(self):
        self.assertEqual(CuttingStockSolver().solve([CutDemand(13_571, 1, 18)]).status, "INFEASIBLE")

    def test_agent_counts_split_bars_as_done(self):
        with tempfile.TemporaryDirectory() as tmp:
            bbs = os.path.join(tmp, "bbs.csv")
            with open(bbs, "w", encoding="utf-8") as f:
                f.write("mark,diameter_mm,length_mm,quantity\nF2,18,13571,40\nF3,18,6336,83\nX,18,25000,2\n")
            sup = AECSupervisor(project_root=ROOT, human_gate_mode="auto", persist_path=os.path.join(tmp, "s.json"))
            with contextlib.redirect_stdout(io.StringIO()):
                sup.register_agent(RebarAgent(bbs_path=bbs, splice_zones=[(0.0, 0.35), (0.65, 1.0)]))
                ok = sup.run(phases=[ProjectPhase.REBAR_CUT])
            cutting = sup.bus.get_cutting_dict()
        self.assertTrue(ok)
        self.assertEqual(cutting["pieces_cut"], cutting["pieces_demanded"])
        self.assertEqual(cutting["pieces_demanded"], 40 + 83 + 2)


if __name__ == "__main__":
    unittest.main()
