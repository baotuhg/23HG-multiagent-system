# -*- coding: utf-8 -*-
"""Kiểm thử đánh giá phiếu thí nghiệm & điểm dừng kỹ thuật — python -m unittest discover tests"""

import contextlib
import io
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.sub_agents import BPTCKCSAgent
from core.state.shared_state import ProjectPhase
from core.supervisor.supervisor_agent import AECSupervisor
from tools.lab_qaqc import (LabCriteria, concrete_requirement, evaluate, group_strength, load_lab_results,
                            write_lab_report)

SAMPLE = os.path.join(ROOT, "templates", "Phieu_thi_nghiem_mau.csv")
HEADER = ("Mã phiếu;Loại thí nghiệm;Cấu kiện;Mã BBNT;Loại mẫu;Ngày đúc;Ngày thí nghiệm;Tuổi (ngày);Yêu cầu;"
          "Mẫu 1;Mẫu 2;Mẫu 3;Mác thép;Giới hạn chảy (MPa);Giới hạn bền (MPa);Độ giãn dài (%);Chỉ tiêu;Kết quả\n")


class LabTest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def run_rows(self, body, criteria=None):
        path = os.path.join(self.tmp.name, "lab.csv")
        with open(path, "w", encoding="utf-8") as f:
            f.write(HEADER + body)
        return evaluate(load_lab_results(path), path, criteria)

    def test_sample_file(self):
        ev = evaluate(load_lab_results(SAMPLE), SAMPLE)
        self.assertEqual((ev.count("PASS"), ev.count("FAIL"), ev.count("PENDING")), (7, 0, 1))
        self.assertEqual({h.bbnt: h.status for h in ev.hold_points}["BBNT-18"], "CHỜ")
        self.assertEqual(ev.errors, [])

    def test_requirements(self):
        self.assertEqual(concrete_requirement("C30/37", cylinder=False), 37)
        self.assertEqual(concrete_requirement("C30/37", cylinder=True), 30)
        self.assertEqual(concrete_requirement("C30", cylinder=False), 37)
        self.assertEqual(concrete_requirement("M300", cylinder=False), 29.42)
        self.assertEqual(concrete_requirement("32.5", cylinder=False), 32.5)
        with self.assertRaises(ValueError):
            concrete_requirement("B25", cylinder=False)

    def test_group_strength_outlier(self):
        self.assertAlmostEqual(group_strength([30, 31, 32], 15)[0], 31)
        value, note = group_strength([30, 31, 40], 15)
        self.assertEqual(value, 31)
        self.assertTrue(note)

    def test_concrete_acceptance(self):
        ev = self.run_rows(
            "A;Nén bê tông;Bệ;BB-1;;2026-01-01;2026-01-29;;C30/37;36.1;35.2;30.9;;;;;;\n"   # viên < 85%
            "B;Nén bê tông;Bệ;BB-2;;2026-01-01;2026-01-29;;C30/37;38;37;37.5;;;;;;\n"      # đạt
            "C;Nén bê tông;Bệ;BB-3;;2026-01-01;2026-01-29;;C30/37;36.5;36.8;36.9;;;;;;\n"  # tổ < 100%
            "D;Nén bê tông;Bệ;BB-4;Lập phương 100;2026-01-01;2026-01-29;;37;40;40;40;;;;;;\n"  # 40×0.91 < 37
            "E;Nén bê tông;Bệ;BB-5;;2026-01-01;2026-01-08;;C30/37;20;21;22;;;;;;\n"      # R7 thấp → chờ + cảnh báo
        )
        status = {r.test_id: r.status for r in ev.records}
        self.assertEqual(status, {"A": "FAIL", "B": "PASS", "C": "FAIL", "D": "FAIL", "E": "PENDING"})
        self.assertTrue(any("cảnh báo sớm" in w for w in ev.warnings))
        holds = {h.bbnt: h.status for h in ev.hold_points}
        self.assertEqual(holds, {"BB-1": "CHẶN", "BB-2": "GIẢI TỎA", "BB-3": "CHẶN", "BB-4": "CHẶN", "BB-5": "CHỜ"})

    def test_criteria_are_configurable(self):
        ev = self.run_rows("C;Nén bê tông;Bệ;BB-3;;2026-01-01;2026-01-29;;C30/37;36.5;36.8;36.9;;;;;;\n",
                           LabCriteria(group_min_ratio=0.95))
        self.assertEqual(ev.records[0].status, "PASS")

    def test_date_errors(self):
        ev = self.run_rows(
            "A;Nén bê tông;Bệ;BB;;2026-01-29;2026-01-01;;C30/37;38;38;38;;;;;;\n"
            "B;Nén bê tông;Bệ;BB;;2026-01-01;2026-01-29;7;C30/37;38;38;38;;;;;;\n"
            "C;Nén bê tông;Bệ;BB;;;;;C30/37;38;38;38;;;;;;\n"
            "A;Kéo thép;Thép;BB;;;;;;;;;CB400-V;450;600;16;;\n")
        self.assertEqual(len(ev.errors), 4)   # nén trước đúc; tuổi ≠ ngày; thiếu tuổi; trùng mã phiếu

    def test_steel_and_other(self):
        ev = self.run_rows(
            "S1;Kéo thép;Thép Ø25;BB;;;;;;;;;CB400-V;452;618;17.5;;\n"
            "S2;Kéo thép;Thép Ø16;BB;;;;;;;;;CB400-V;390;600;16;;\n"
            "P1;PDA;Cọc;BB;;;;;≥ 7800;;;;;;;;Sức chịu tải;7500\n"
            "L1;Độ sụt;Cọc;BB;;;;;18±2;;;;;;;;Độ sụt;19\n"
            "U1;Siêu âm;Cọc;BB;;;;;Loại 1;;;;;;;;Phân loại;loại 1\n")
        status = {r.test_id: r.status for r in ev.records}
        self.assertEqual(status, {"S1": "PASS", "S2": "FAIL", "P1": "FAIL", "L1": "PASS", "U1": "PASS"})
        bad = self.run_rows("S9;Kéo thép;Thép;BB;;;;;;;;;SD390;450;600;16;;\n")
        self.assertEqual(len(bad.errors), 1)

    def test_report(self):
        ev = evaluate(load_lab_results(SAMPLE), SAMPLE)
        path = os.path.join(self.tmp.name, "qaqc.xlsx")
        write_lab_report(path, ev)
        import openpyxl
        self.assertEqual(openpyxl.load_workbook(path).sheetnames[:3], ["HOLD_POINTS", "KET_QUA", "TIEU_CHI"])


class LabAgentTest(unittest.TestCase):

    def run_phase(self, demo, **kwargs):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        sup = AECSupervisor(project_root=ROOT, human_gate_mode="auto",
                            persist_path=os.path.join(tmp.name, "s.json"), demo_mode=demo)
        with contextlib.redirect_stdout(io.StringIO()):
            sup.register_agent(BPTCKCSAgent(**kwargs))
            ok = sup.run(phases=[ProjectPhase.QAQC_REVIEW])
        return ok, sup

    def test_real_file_passes(self):
        ok, sup = self.run_phase(False, lab_path=SAMPLE)
        self.assertTrue(ok)
        self.assertEqual(sup.bus.get_sample_data_sources(), [])
        self.assertEqual(sup.bus.get_qaqc_data().lab_summary["pending"], 1)

    def test_failed_test_stops_pipeline(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "lab.csv")
            with open(path, "w", encoding="utf-8") as f:
                f.write(HEADER + "A;Nén bê tông;Bệ;BB-1;;2026-01-01;2026-01-29;;C30/37;30;30;30;;;;;;\n")
            ok, sup = self.run_phase(False, lab_path=path)
        self.assertFalse(ok)
        self.assertEqual(sup.bus.get_qaqc_data().hold_point_status[0]["status"], "CHẶN")

    def test_missing_file_and_demo(self):
        self.assertFalse(self.run_phase(False)[0])
        ok, sup = self.run_phase(True)
        self.assertTrue(ok)
        self.assertEqual([s["agent_id"] for s in sup.bus.get_sample_data_sources()], ["bptc_kcs_agent"])


if __name__ == "__main__":
    unittest.main()
