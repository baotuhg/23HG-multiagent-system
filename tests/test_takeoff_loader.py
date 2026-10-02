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


class TakeoffInSupervisorTest(unittest.TestCase):
    """Bảng cấu kiện đi vào Supervisor ở pha CAD_TAKEOFF: State Bus, Gate-1, dữ liệu sai thì dừng."""

    def run_phase(self, path, profile="tt13-2021", out=None):
        import contextlib
        import io
        from core.agents.sub_agents import CADAgent
        from core.state.shared_state import ProjectPhase
        from core.supervisor.supervisor_agent import AECSupervisor
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        sup = AECSupervisor(project_root=ROOT, human_gate_mode="auto", persist_path=os.path.join(tmp.name, "s.json"))
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            sup.register_agent(CADAgent(takeoff_path=path, takeoff_profile=profile, takeoff_out=out))
            ok = sup.run(phases=[ProjectPhase.CAD_TAKEOFF])
        return ok, sup, buf.getvalue()

    def test_quantities_reach_state_bus_with_derivations(self):
        ok, sup, log = self.run_phase(SAMPLE)
        self.assertTrue(ok, log)
        cad = sup.bus.get_cad_data()
        self.assertAlmostEqual(cad.total_concrete_m3, 367.191, places=3)    # 0.96 + 4.32 + 361.911
        self.assertAlmostEqual(cad.total_formwork_m2, 28.92, places=3)      # 3.52 + 5.1 + 20.3
        self.assertEqual(cad.excavation_m3, 500.0)                           # chỉ đào hào; mặt cắt không gộp vào
        self.assertEqual(len(cad.takeoff_quantities), 12)
        self.assertFalse(cad.takeoff_profile["verified"])
        first = cad.takeoff_quantities[0]
        self.assertEqual((first["kind"], first["formula"]), ("be_tong", "2 × (1.2 × 1 × 0.4)"))
        self.assertIn("CHƯA đối chiếu bản gốc", log)                         # Gate-1 cảnh báo, không chặn
        self.assertFalse(sup.bus.get_sample_data_sources())                  # dữ liệu thật, không phải mẫu

    def test_bad_table_stops_the_phase(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = os.path.join(tmp, "x.csv")
            with open(p, "w", encoding="utf-8") as f:
                f.write("Loại;Tên công tác;Dài (m)\nbe_tong;Móng;1\n")
            ok, _, log = self.run_phase(p)
        self.assertFalse(ok)
        self.assertIn("thiếu cột", log)
        self.assertIn("không retry", log)

    def test_cli_with_phase_runs_through_supervisor(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "kq.xlsx")
            env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", AEC_STATE_DIR=tmp)
            r = subprocess.run([sys.executable, os.path.join(ROOT, "run_state_graph.py"), "--takeoff", SAMPLE,
                                "--phase", "takeoff", "--takeoff-out", out], capture_output=True, text=True,
                               encoding="utf-8", env=env, cwd=ROOT)
            self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
            self.assertIn("GATE-1", r.stdout)
            self.assertTrue(os.path.exists(out))
