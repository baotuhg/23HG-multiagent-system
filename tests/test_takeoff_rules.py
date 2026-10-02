# -*- coding: utf-8 -*-
"""Kiểm thử diễn giải khối lượng cống hộp: hình học tính tay, và đối chiếu công thức trong sheet với module."""

import math
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import openpyxl

from tools.excel_eval import WorkbookEvaluator
import json
import tempfile

from tools.takeoff_rules import (
    BANG_6_2_HEADER, BoxCulvert, CULVERT_ROWS, K_BETONG_M3, MeasurementProfile, PROFILE_TT13_2021_PL_VI,
    average_end_volume, bang_6_1, bang_6_2, bored_length, circular_column, load_profile, net_of_buried_works,
    pile, pipe_length, pit_excavation, rect_concrete, rect_formwork, scaffold_column, scaffold_extra_layers,
    scaffold_inner, trench_excavation,
)

A5 = os.path.join(ROOT, "examples", "HO_SO_CONG_HOP_TUYEN_A5", "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO",
                  "03_QS_Dien_Giai_Chi_Tiet_Takeoff_Cong_A5.xlsx")
CASES = {  # (cống, thể tích BT trên 1 m tính tay)
    8: (BoxCulvert(2.5, 2.5, 2.0, 2.0, 1, 0.20), 2.33),        # 6.25 − 4 + 4·½·0.04          = 2.330
    9: (BoxCulvert(3.6, 3.6, 3.0, 3.0, 1, 0.25), 4.085),       # 12.96 − 9 + 4·½·0.0625       = 4.085
    10: (BoxCulvert(6.9, 3.6, 3.0, 3.0, 2, 0.25), 7.09),       # 24.84 − 18 + 8·½·0.0625      = 7.090
}


class BoxCulvertGeometryTest(unittest.TestCase):

    def test_concrete_area_by_hand(self):
        for row, (bc, expected) in CASES.items():
            self.assertAlmostEqual(bc.concrete_area(), expected, places=9, msg=f"dòng {row}")

    def test_inner_perimeter_haunch_replaces_two_legs_by_hypotenuse(self):
        # lòng 2×2, vút 0.2: 8 − 4·0.2·(2 − √2) = 8 − 0.8·0.585786... = 7.531371
        self.assertAlmostEqual(BoxCulvert(2.5, 2.5, 2, 2, 1, 0.2).inner_perimeter(), 7.5313708499, places=9)
        self.assertEqual(BoxCulvert(2.5, 2.5, 2, 2, 1, 0.0).inner_perimeter(), 8.0)   # không vút = chu vi chữ nhật

    def test_formwork_faces_by_hand(self):
        bc = BoxCulvert(2.5, 2.5, 2, 2, 1, 0.2)
        fw = bc.formwork_per_m(11.3)                     # mặc định: trong + hông ngoài + đầu đốt
        self.assertAlmostEqual(fw["trong"], 7.5313708499, places=9)
        self.assertEqual(fw["hong_ngoai"], 5.0)           # 2 × 2.5
        self.assertAlmostEqual(fw["dau_dot"], 2 * 2.33 / 11.3, places=12)
        self.assertEqual(fw["mat_tren"] + fw["mat_day"], 0.0)
        full = bc.formwork_per_m(11.3, top=True, bottom=True)
        self.assertEqual(full["mat_tren"], 2.5)
        self.assertAlmostEqual(full["tong"] - fw["tong"], 5.0, places=12)   # thêm nắp + đáy = 2 × 2.5

    def test_double_cell_has_two_inner_perimeters(self):
        bc = BoxCulvert(6.9, 3.6, 3, 3, 2, 0.25)
        self.assertAlmostEqual(bc.formwork_per_m(11.3)["trong"], 2 * (12 - 4 * 0.25 * (2 - math.sqrt(2))), places=9)


@unittest.skipUnless(os.path.exists(A5), "thiếu workbook ví dụ A5")
class A5SheetDerivationTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.ev = WorkbookEvaluator(A5)
        cls.ws = openpyxl.load_workbook(A5)["QS_TAKEOFF"]

    def test_concrete_quantities_unchanged_and_match_hand_values(self):
        # số đốt × dài × diện tích mặt cắt: 34·11.3·2.33, 50·11.3·4.085, 108·11.3·7.09
        for row, expected in ((8, 895.186), (9, 2308.025), (10, 8652.636)):
            self.assertAlmostEqual(self.ev.value("QS_TAKEOFF", row, 9), expected, places=3)

    def test_no_hardcoded_dimensions_in_concrete_formulas(self):
        for row in CULVERT_ROWS:
            f = self.ws[f"I{row}"].value
            self.assertNotIn("2.0*2.0", f)
            self.assertNotIn("3.0*3.0", f)
            self.assertNotIn("0.25", f)
            self.assertNotIn("0.2*", f)
            for col in "KLMN":
                self.assertIsNotNone(self.ws[f"{col}{row}"].value, f"thiếu đầu vào {col}{row}")

    def test_sheet_formwork_formula_matches_python_module(self):
        for crow, spec in CULVERT_ROWS.items():
            bc, _ = CASES[crow]
            expected = bc.formwork_per_m(11.3)["tong"]
            self.assertAlmostEqual(self.ev.value("QS_TAKEOFF", spec["formwork_row"], 19), expected, places=9)

    def test_legacy_formwork_numbers_are_kept_not_silently_overwritten(self):
        # Số ván khuôn đang dùng cho hợp đồng được giữ nguyên; chênh lệch hiển thị ở cột T để QS quyết định.
        self.assertEqual([self.ws[f"F{r}"].value for r in (11, 12, 13)], [11.14, 16.48, 25.66])


class ElementLibraryTest(unittest.TestCase):
    """Mọi kỳ vọng tính tay (số thật, không lấy từ chương trình)."""

    def test_footing(self):
        # 2 móng 1.2 × 1.0 × 0.4: BT = 2 × 0.48 = 0.960 ; ván khuôn hông = 2 × (2 × 2.2 × 0.4) = 3.520
        self.assertEqual(rect_concrete("M1", 2, 1.2, 1.0, 0.4).value, 0.96)
        self.assertEqual(rect_formwork("M1", "mong", 2, 1.2, 1.0, 0.4).value, 3.52)

    def test_column(self):
        # 10 cột 0.3 × 0.4, cao 3.6: BT = 10 × 0.432 = 4.320 ; VK = 10 × 2 × 0.7 × 3.6 = 50.400
        self.assertEqual(rect_concrete("C1", 10, 0.3, 0.4, 3.6).value, 4.32)
        self.assertEqual(rect_formwork("C1", "cot", 10, 0.3, 0.4, 3.6).value, 50.4)

    def test_beam_formwork_subtracts_slab_embedded_part(self):
        # dầm dài 5, rộng 0.22, cao 0.5, sàn dày 0.1: BT = 0.550 ; VK = 5 × (2 × 0.4 + 0.22) = 5.100
        self.assertEqual(rect_concrete("D1", 1, 5, 0.22, 0.5).value, 0.55)
        self.assertEqual(rect_formwork("D1", "dam", 1, 5, 0.22, 0.5, slab_thickness=0.1).value, 5.1)
        # không trừ phần chìm khi hồ sơ tắt trừ giao nhau: 5 × (2 × 0.5 + 0.22) = 6.100
        no_overlap = MeasurementProfile(deduct_formwork_overlap=False)
        self.assertEqual(rect_formwork("D1", "dam", 1, 5, 0.22, 0.5, slab_thickness=0.1, profile=no_overlap).value, 6.1)

    def test_slab_and_wall(self):
        # sàn 4 × 5 dày 0.1: BT = 2.000 ; VK đáy = 20.000 ; thêm cạnh = 2 × 9 × 0.1 = 1.800 → 21.800
        self.assertEqual(rect_concrete("S1", 1, 4, 5, 0.1).value, 2.0)
        self.assertEqual(rect_formwork("S1", "san", 1, 4, 5, 0.1).value, 20.0)
        self.assertEqual(rect_formwork("S1", "san", 1, 4, 5, 0.1, edges=True).value, 21.8)
        # tường dài 10, cao 3, dày 0.2: BT = 6.000 ; VK 2 mặt = 60.000 ; thêm 2 đầu = 2 × 0.2 × 3 = 1.200
        self.assertEqual(rect_concrete("T1", 1, 10, 0.2, 3).value, 6.0)
        self.assertEqual(rect_formwork("T1", "tuong", 1, 10, 0.2, 3).value, 60.0)
        self.assertEqual(rect_formwork("T1", "tuong", 1, 10, 0.2, 3, ends=True).value, 61.2)

    def test_openings_follow_profile_threshold(self):
        wall = (1, 10, 0.2, 3)   # BT thô 6.000 ; lỗ 0.5 m3 và 0.05 m3
        deduct_all = rect_concrete("T1", *wall, openings_m3=[0.5, 0.05])
        self.assertEqual(deduct_all.value, 5.45)                                  # 6 − 0.5 − 0.05
        over_threshold = MeasurementProfile(no_deduct_below={K_BETONG_M3: 0.1})
        q = rect_concrete("T1", *wall, openings_m3=[0.5, 0.05], profile=over_threshold)
        self.assertEqual(q.value, 5.5)                                            # chỉ trừ lỗ 0.5 > 0.1
        self.assertTrue(any("không trừ 1 lỗ" in n for n in q.notes))

    def test_bored_pile_and_cutoff(self):
        # 8 cọc D1.2 dài 40: π × 0.36 × 320 = 361.911 ; đập đầu 0.8 m: π × 0.36 × 0.8 × 8 = 7.238
        conc, demo = pile("CKN", 8, 1.2, 40, cutoff=0.8)
        self.assertAlmostEqual(conc.value, 361.911, places=3)
        self.assertAlmostEqual(demo.value, 7.238, places=3)
        self.assertIsNone(pile("CKN", 8, 1.2, 40)[1])

    def test_circular_column(self):
        # D0.5, cao 3: BT = π × 0.0625 × 3 = 0.589 ; VK = π × 0.5 × 3 = 4.712
        conc, form = circular_column("CT", 1, 0.5, 3)
        self.assertAlmostEqual(conc.value, 0.589, places=3)
        self.assertAlmostEqual(form.value, 4.712, places=3)

    def test_excavation(self):
        # hào đáy 1.5, sâu 2, mái 1:0.5, dài 100: (1.5 + 1.0) × 2 × 100 = 500.000
        self.assertEqual(trench_excavation("H1", 100, 1.5, 2.0, 0.5).value, 500.0)
        # hố 2 × 3, sâu 1.5, không mái: lăng trụ = 6 × 1.5 = 9.000
        self.assertAlmostEqual(pit_excavation("H2", 2, 3, 1.5).value, 9.0, places=9)
        # hố 2 × 2, sâu 3, mái 1:0.5: đáy trên 5 × 5 = 25 ; V = 3/3 × (4 + 25 + 10) = 39.000
        self.assertAlmostEqual(pit_excavation("H3", 2, 2, 3, 0.5).value, 39.0, places=9)

    def test_average_end_area(self):
        # lý trình 0, 20, 50; diện tích 10, 16, 12: (10+16)/2 × 20 + (16+12)/2 × 30 = 260 + 420 = 680
        q = average_end_volume("DAO", [0, 20, 50], [10, 16, 12])
        self.assertEqual(q.value, 680.0)
        with self.assertRaises(ValueError):
            average_end_volume("X", [0, 0], [1, 1])

    def test_explanation_is_readable_and_unverified_profile_warns(self):
        q = rect_concrete("M1", 2, 1.2, 1.0, 0.4)
        text = q.explain()
        self.assertIn("2 × (1.2 × 1 × 0.4)", text)
        self.assertIn("= 0.960 m3", text)
        self.assertTrue(any("CHƯA đối chiếu" in n for n in q.notes))
        self.assertEqual(rect_concrete("M1", 2, 1.2, 1.0, 0.4, profile=MeasurementProfile(verified=True)).notes, ())

    def test_profile_from_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "p.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"name": "Dự án X", "source": "Điều ... đã đối chiếu", "verified": True,
                           "no_deduct_below": {"be_tong_m3": 0.1}, "ocr_pending": ["Trang 5"]}, f)
            pr = load_profile(path)
            self.assertTrue(pr.verified)
            self.assertEqual(pr.threshold("be_tong_m3"), 0.1)
            self.assertEqual(pr.ocr_pending, ("Trang 5",))
            bad = os.path.join(tmp, "bad.json")
            with open(bad, "w", encoding="utf-8") as f:
                json.dump({"khong_co": 1}, f)
            with self.assertRaises(ValueError):
                load_profile(bad)


class PhuLucVIRulesTest(unittest.TestCase):
    """Quy tắc rút ra từ Phụ lục VI (bản OCR người dùng cung cấp). Số tính tay."""
    P = PROFILE_TT13_2021_PL_VI

    def test_preset_is_unverified_and_lists_ocr_items(self):
        self.assertFalse(self.P.verified)
        text = " | ".join(self.P.warnings())
        self.assertIn("CHƯA đối chiếu", text)
        self.assertIn("Cần đối chiếu", text)
        self.assertEqual(self.P.decimals, 3)
        self.assertEqual(self.P.rebar_no_deduct_below_ratio, 0.02)

    def test_concrete_void_threshold_is_strictly_below(self):
        # tường BT thô 6.000 ; lỗ 0.5, 0.1, 0.05: lỗ < 0.1 không trừ, lỗ ≥ 0.1 thì trừ → 6 − 0.5 − 0.1 = 5.400
        q = rect_concrete("T1", 1, 10, 0.2, 3, openings_m3=[0.5, 0.1, 0.05], profile=self.P)
        self.assertEqual(q.value, 5.4)

    def test_rebar_deducted_only_when_content_reaches_2_percent(self):
        # cột 0.3×0.3×3 = 0.270 m3: cốt thép 0.004 (1.48%) không trừ ; 0.007 (2.59%) thì trừ → 0.263
        self.assertEqual(rect_concrete("C", 1, 0.3, 0.3, 3, rebar_volume_m3=0.004, profile=self.P).value, 0.27)
        self.assertEqual(rect_concrete("C", 1, 0.3, 0.3, 3, rebar_volume_m3=0.007, profile=self.P).value, 0.263)
        # hồ sơ mặc định (không đặt ngưỡng) không trừ cốt thép
        self.assertEqual(rect_concrete("C", 1, 0.3, 0.3, 3, rebar_volume_m3=0.007).value, 0.27)

    def test_formwork_void_threshold(self):
        # sàn 4×5: ván khuôn đáy 20.000 ; lỗ 1.5 (≥ 1 m2) trừ, lỗ 0.5 (< 1 m2) không → 18.500
        q = rect_formwork("S", "san", 1, 4, 5, 0.1, openings_m2=[1.5, 0.5], profile=self.P)
        self.assertEqual(q.value, 18.5)

    def test_bored_pile_drilling_and_buried_works(self):
        self.assertEqual(bored_length("Khoan", 8, 42).value, 336.0)
        # đắp: hào 2170 × 4 × 2.5 = 21,700 trừ phần cống chiếm chỗ 6.25 m2 × 11.3 × 34 = 2,401.25 → 19,298.75
        q = net_of_buried_works("Đắp", 21700, [("cống 2x2 bao ngoài", 2401.25)])
        self.assertEqual(q.value, 19298.75)

    def test_pipe_length_excludes_chambers_only_for_drainage(self):
        self.assertEqual(pipe_length("Ống", 120, 3.0, drainage=True).value, 117.0)
        self.assertEqual(pipe_length("Ống", 120, 3.0, drainage=False).value, 120.0)

    def test_scaffold_layers_rule(self):
        # cao ≤ 3.6 không tính; mỗi 1.2 m tăng thêm = 1 lớp; phần dư < 0.6 không tính, ≥ 0.6 tính 1 lớp
        cases = {3.6: 0, 4.0: 0, 4.2: 1, 4.8: 1, 5.4: 2, 6.0: 2, 6.6: 3}
        for h, layers in cases.items():
            self.assertEqual(scaffold_extra_layers(h), layers, f"cao {h}")
        self.assertEqual(scaffold_inner("DG", 50, 3.0).value, 0.0)
        self.assertEqual(scaffold_inner("DG", 50, 5.4).value, 150.0)          # 50 × (1 + 2)
        self.assertEqual(scaffold_column("DG cột", 1.6, 5).value, 26.0)       # (1.6 + 3.6) × 5

    def test_detail_table_6_2_columns(self):
        q = rect_concrete("Bê tông móng M1", 2, 1.2, 1.0, 0.4).with_ref(drawing="KC-01", code="AF.11110")
        rows = bang_6_2([q])
        self.assertEqual(rows[0], BANG_6_2_HEADER)
        self.assertEqual(len(rows[1]), 10)
        stt, drawing, code, name, unit, count, formula, per_unit, total, _ = rows[1]
        self.assertEqual((stt, drawing, code, name, unit, count), (1, "KC-01", "AF.11110", "Bê tông móng M1", "m3", 2))
        self.assertEqual((per_unit, total), (0.48, 0.96))                       # cột (9) = (6) × (8)
        self.assertIn("2 × (1.2 × 1 × 0.4)", formula)

    def test_summary_table_6_1_merges_same_work(self):
        a = rect_concrete("Bê tông móng", 2, 1.2, 1.0, 0.4).with_ref(code="AF.11110")
        b = rect_concrete("Bê tông móng", 1, 1.0, 1.0, 0.4).with_ref(code="AF.11110")
        rows = bang_6_1([a, b])
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][5], 1.36)                                      # 0.960 + 0.400


if __name__ == "__main__":
    unittest.main()
