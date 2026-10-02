# -*- coding: utf-8 -*-
"""Đo bóc từ bảng cấu kiện (--takeoff): đọc file mẫu, số tính tay, dòng sai dữ liệu, CLI xuất Bảng 6.2 / 6.1."""

import os
import subprocess
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import openpyxl

from tools.takeoff_loader import TakeoffLoadError, load_takeoff, resolve_profile
from tools.takeoff_rules import DEFAULT_PROFILE, PROFILE_TT13_2021_PL_VI

SAMPLE = os.path.join(ROOT, "templates", "Mau_dau_vao_do_boc.csv")


class TakeoffLoaderTest(unittest.TestCase):

    def values(self, profile=DEFAULT_PROFILE):
        return {q.name: q.value for q in load_takeoff(SAMPLE, profile=profile)}

    def test_sample_file_hand_values(self):
        v = self.values(PROFILE_TT13_2021_PL_VI)
        self.assertEqual(v["Bê tông móng M1"], 0.96)                    # 2 × 1.2 × 1.0 × 0.4 (dấu phẩy thập phân VN)
        self.assertEqual(v["Ván khuôn móng M1"], 3.52)
        self.assertEqual(v["Bê tông cột C1"], 4.32)
        self.assertEqual(v["Ván khuôn dầm D1"], 5.1)
        self.assertEqual(v["Ván khuôn sàn S1"], 20.3)                    # 20 + 1.8 − lỗ 1.5 m2 (≥ 1 m2 nên trừ)
        self.assertAlmostEqual(v["Bê tông cọc khoan nhồi D1200"], 361.911, places=3)
        self.assertAlmostEqual(v["Bê tông cọc khoan nhồi D1200 — đập đầu cọc"], 7.238, places=3)
        self.assertEqual(v["Khoan tạo lỗ cọc D1200 trên cạn"], 336.0)
        self.assertEqual(v["Đào hào đặt cống"], 500.0)
        self.assertEqual(v["Đào nền đường"], 680.0)                       # 3 mặt cắt gộp theo tên
        self.assertEqual(v["Ống thoát nước D600"], 117.0)
        self.assertEqual(v["Dàn giáo trong sảnh"], 150.0)

    def test_profile_changes_result(self):
        # hồ sơ mặc định (không ngưỡng) vẫn trừ lỗ 1.5 m2; muốn thấy khác: lỗ nhỏ hơn ngưỡng 1 m2
        self.assertEqual(self.values()["Ván khuôn sàn S1"], 20.3)

    def test_references_are_carried(self):
        q = next(q for q in load_takeoff(SAMPLE) if q.name == "Bê tông móng M1")
        self.assertEqual((q.drawing, q.code, q.count, q.per_unit), ("KC-01", "AF.11110", 2.0, 0.48))

    def test_bad_rows_are_listed_not_guessed(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = os.path.join(tmp, "x.csv")
            with open(p, "w", encoding="utf-8") as f:
                f.write("Loại;Tên công tác;Số lượng;Dài (m);Rộng (m);Cao (m);Kiểu\n"
                        "be_tong;Móng;1;1;1;\n"                 # thiếu cao
                        "van_khuon;VK;1;1;1;1;khung\n"           # kiểu sai
                        "xay;Tường;1;1;1;1\n"                     # loại chưa hỗ trợ
                        "be_tong;Cột;-2;1;1;1\n")                 # số âm
            with self.assertRaises(TakeoffLoadError) as cm:
                load_takeoff(p)
            msg = str(cm.exception)
            self.assertIn("4 dòng sai dữ liệu", msg)
            for frag in ("Dòng 2", "thiếu cột cao", "Dòng 3", "Dòng 4", "loại 'xay'", "Dòng 5"):
                self.assertIn(frag, msg)

    def test_resolve_profile(self):
        self.assertIs(resolve_profile("tt13-2021"), PROFILE_TT13_2021_PL_VI)
        self.assertIs(resolve_profile(None), DEFAULT_PROFILE)
        with self.assertRaises(TakeoffLoadError):
            resolve_profile("khong-co")


class TakeoffCliTest(unittest.TestCase):

    def run_cli(self, *args):
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
        return subprocess.run([sys.executable, os.path.join(ROOT, "run_state_graph.py"), *args],
                              capture_output=True, text=True, encoding="utf-8", env=env, cwd=ROOT)

    def test_cli_writes_forms_6_2_and_6_1(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "kq.xlsx")
            r = self.run_cli("--takeoff", SAMPLE, "--takeoff-out", out, "--takeoff-profile", "tt13-2021")
            self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
            wb = openpyxl.load_workbook(out)
            self.assertEqual(wb.sheetnames, ["BANG_6_2_CHI_TIET", "BANG_6_1_TONG_HOP", "QUY_TAC_DO_BOC"])
            ws = wb["BANG_6_2_CHI_TIET"]
            self.assertEqual(ws["I4"].value, "KHỐI LƯỢNG TOÀN BỘ")
            first = [c.value for c in ws[5]]
            self.assertEqual(first[:6], [1, "KC-01", "AF.11110", "Bê tông móng M1", "m3", 2])
            self.assertEqual((first[7], first[8]), (0.48, 0.96))
            rules = {r[0].value: r[1].value for r in wb["QUY_TAC_DO_BOC"].iter_rows()}
            self.assertEqual(rules["Đã đối chiếu bản gốc"], "CHƯA")

    def test_cli_fails_on_bad_input(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = os.path.join(tmp, "x.csv")
            with open(p, "w", encoding="utf-8") as f:
                f.write("Loại;Tên công tác;Dài (m)\nbe_tong;Móng;1\n")
            r = self.run_cli("--takeoff", p)
            self.assertEqual(r.returncode, 1)
            self.assertIn("thiếu cột", r.stdout)


if __name__ == "__main__":
    unittest.main()
