# -*- coding: utf-8 -*-
"""Test CAD Diff Engine — so sánh hai phiên bản bản vẽ (cấu kiện thêm / bỏ / sửa / không đổi)."""

import unittest

from tools.cad_diff_engine import CADDiffEngine


def comp(cid, concrete=0.0, formwork=0.0, rebar=0.0, name=None, wbs="MÓNG"):
    return {"id": cid, "name": name or cid, "wbs": wbs,
            "concrete_m3": concrete, "formwork_m2": formwork, "rebar_kg": rebar}


class CADDiffEngineTest(unittest.TestCase):
    def setUp(self):
        self.engine = CADDiffEngine()

    def diff(self, v1, v2, **kw):
        return self.engine.diff({"components": v1}, {"components": v2}, **kw)

    def test_added_removed_modified_unchanged(self):
        v1 = [comp("A", 10, 20, 100), comp("B", 5, 8, 50), comp("C", 3, 4, 30)]
        v2 = [comp("A", 12, 20, 100),            # sửa: bê tông +2
              comp("C", 3, 4, 30),               # không đổi
              comp("D", 7, 9, 70)]               # thêm; B bị bỏ
        r = self.diff(v1, v2, rev_from="Rev01", rev_to="Rev02")
        self.assertEqual((r.rev_from, r.rev_to), ("Rev01", "Rev02"))
        self.assertEqual(r.total_components_compared, 4)
        self.assertEqual((r.added_count, r.removed_count, r.modified_count, r.unchanged_count), (1, 1, 1, 1))
        kinds = {d.component_id: d.change_type for d in r.deltas}
        self.assertEqual(kinds, {"A": "MODIFIED", "B": "REMOVED", "D": "ADDED"})

    def test_deltas_signs_and_net_totals(self):
        v1 = [comp("A", 10, 20, 100), comp("B", 5, 8, 50)]
        v2 = [comp("A", 12, 18, 130), comp("D", 7, 9, 70)]
        r = self.diff(v1, v2)
        by = {d.component_id: d for d in r.deltas}
        self.assertAlmostEqual(by["A"].concrete_delta_m3, 2.0)
        self.assertAlmostEqual(by["A"].formwork_delta_m2, -2.0)
        self.assertAlmostEqual(by["A"].rebar_delta_kg, 30.0)
        self.assertAlmostEqual(by["B"].concrete_delta_m3, -5.0)        # bị bỏ → âm
        self.assertAlmostEqual(by["D"].rebar_delta_kg, 70.0)           # thêm → dương
        # net = tổng các delta: bê tông 2 − 5 + 7, ván khuôn −2 − 8 + 9, thép 30 − 50 + 70
        self.assertAlmostEqual(r.net_concrete_delta_m3, 4.0)
        self.assertAlmostEqual(r.net_formwork_delta_m2, -1.0)
        self.assertAlmostEqual(r.net_rebar_delta_kg, 50.0)

    def test_only_changed_components_are_recalculated(self):
        v1 = [comp("A", 10), comp("B", 5), comp("C", 3)]
        v2 = [comp("A", 10), comp("B", 6), comp("C", 3), comp("D", 1)]
        r = self.diff(v1, v2)
        self.assertEqual(r.components_to_recalculate, ["B", "D"])      # theo thứ tự id; A, C không tính lại

    def test_change_within_tolerance_is_unchanged(self):
        v1 = [comp("A", 10.0, 20.0, 100.0)]
        v2 = [comp("A", 10.0 + 0.0005, 20.0 + 0.005, 100.0 + 0.05)]    # dưới cả 3 ngưỡng
        r = self.diff(v1, v2)
        self.assertEqual((r.unchanged_count, r.modified_count), (1, 0))
        self.assertEqual(r.deltas, [])
        self.assertEqual(r.net_concrete_delta_m3, 0.0)

    def test_change_just_over_any_single_tolerance_is_modified(self):
        base = [comp("A", 10.0, 20.0, 100.0)]
        for over in (comp("A", 10.0 + 0.002, 20.0, 100.0),             # chỉ bê tông vượt 0,001 m³
                     comp("A", 10.0, 20.0 + 0.02, 100.0),              # chỉ ván khuôn vượt 0,01 m²
                     comp("A", 10.0, 20.0, 100.0 + 0.2)):              # chỉ thép vượt 0,1 kg
            r = self.diff(base, [over])
            self.assertEqual(r.modified_count, 1, over)

    def test_missing_quantity_keys_default_to_zero(self):
        r = self.diff([{"id": "A"}], [{"id": "A", "concrete_m3": 3.0}])
        self.assertEqual(r.modified_count, 1)
        self.assertAlmostEqual(r.deltas[0].concrete_delta_m3, 3.0)
        added = self.diff([], [{"id": "N"}])                            # thiếu name/wbs → lấy id / rỗng
        self.assertEqual(added.deltas[0].component_name, "N")
        self.assertEqual(added.deltas[0].wbs, "")

    def test_old_and_new_values_are_kept_for_modified(self):
        r = self.diff([comp("A", 10)], [comp("A", 11)])
        d = r.deltas[0]
        self.assertEqual(d.old_values["concrete_m3"], 10)
        self.assertEqual(d.new_values["concrete_m3"], 11)

    def test_empty_and_identical_snapshots(self):
        self.assertEqual(self.diff([], []).total_components_compared, 0)
        same = [comp("A", 1, 2, 3), comp("B", 4, 5, 6)]
        r = self.diff(same, [dict(c) for c in same])
        self.assertEqual((r.unchanged_count, r.modified_count, r.added_count, r.removed_count), (2, 0, 0, 0))
        self.assertEqual(r.components_to_recalculate, [])
        self.assertEqual(self.engine.diff({}, {}).total_components_compared, 0)   # thiếu khóa 'components'


if __name__ == "__main__":
    unittest.main()
