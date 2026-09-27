# -*- coding: utf-8 -*-
"""Kiểm thử bộ tính công thức Excel, bộ đọc bảng QS và dự toán G_XD — python -m unittest discover tests"""

import contextlib
import io
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.sub_agents import QSAgent
from core.gates.quality_gate import QualityGate
from core.state.shared_state import NodeStatus, ProjectPhase
from core.supervisor.supervisor_agent import AECSupervisor
from tools.excel_eval import FormulaError, WorkbookEvaluator
from tools.qs_loader import QSLoadError, apply_rate_overrides, load_qs

TEMPLATE = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")


def make_workbook(path, sheets):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


class ExcelEvalTest(unittest.TestCase):

    def test_functions(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "f.xlsx")
            make_workbook(path, {
                "A": [["x", 2, 3], ["Cọc khoan nhồi T1", 5, "=B1*C1+B2"], ["Dầm", 7, "=SUM(B1:B3)"],
                      ["", "=ROUNDUP(10/3,0)", "=ROUND(2.345,2)"],
                      ["", "=SUMIFS(B1:B3,A1:A3,\"*khoan*\")", "=COUNTIFS(B1:B3,\">4\")"],
                      ["", "='Tên sheet'!A1*2", "=FOO(1)"]],
                "Tên sheet": [[21]],
            })
            ev = WorkbookEvaluator(path)
            self.assertEqual(ev.value("A", 2, 3), 11)
            self.assertEqual(ev.value("A", 3, 3), 14)
            self.assertEqual(ev.value("A", 4, 2), 4)
            self.assertEqual(ev.value("A", 4, 3), 2.35)
            self.assertEqual(ev.value("A", 5, 2), 5)
            self.assertEqual(ev.value("A", 5, 3), 2)
            self.assertEqual(ev.value("A", 6, 2), 42)
            with self.assertRaises(FormulaError):
                ev.value("A", 6, 3)

    def test_circular_reference(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "c.xlsx")
            make_workbook(path, {"A": [["=B1", "=A1"]]})
            with self.assertRaises(FormulaError):
                WorkbookEvaluator(path).value("A", 1, 1)


class QSLoaderTest(unittest.TestCase):

    def test_repository_template(self):
        est = load_qs(TEMPLATE)
        self.assertEqual(len(est.items), 32)
        self.assertEqual(est.errors, [])
        self.assertEqual(est.rates, {"chung": 0.051, "nha_tam": 0.012, "kxd": 0.01, "tl": 0.055, "vat": 0.1})
        est.compute()
        self.assertEqual(est.T, 27_284_103_489)
        self.assertLessEqual(abs(est.G_XD - est.file_G_XD), 1)       # khớp sheet tổng hợp (làm tròn đồng)
        self.assertTrue(any("khác tổng diễn giải" in w for w in est.warnings))

    def test_csv_with_breakdown_and_checks(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "qs.csv")
            with open(path, "w", encoding="utf-8") as f:
                f.write("STT;Mã hiệu;Nội dung công tác;ĐVT;Khối lượng;Đơn giá;Thành tiền\n"
                        "PHẦN I;;Móng;;;;\n"
                        "1;AF.11111;Bê tông lót;m3;10;1000000;10000000\n"
                        ";;- Móng M1;m3;6;;\n"
                        ";;- Móng M2;m3;3;;\n"
                        "2;AF.61111;Cốt thép;Tấn;2;20000000;41000000\n"
                        "3;AB.1;Đào đất;m3;5;;\n")
            est = load_qs(path)
        self.assertEqual([i.code for i in est.items], ["AF.11111", "AF.61111"])
        self.assertEqual(len(est.errors), 1)                           # đào đất thiếu đơn giá
        self.assertTrue(any("khác tổng diễn giải" in w for w in est.warnings))   # 10 ≠ 6 + 3
        self.assertTrue(any("thành tiền trong file" in w for w in est.warnings)) # 41tr ≠ 40tr
        self.assertEqual(est.missing_rates, ["chung", "nha_tam", "kxd", "tl", "vat"])
        with self.assertRaises(QSLoadError):
            est.compute()
        apply_rate_overrides(est, {"chung": 5, "nha_tam": 1, "kxd": 1, "tl": 5, "vat": 8})
        est.compute()
        self.assertEqual(est.T, 50_000_000)
        self.assertEqual(est.GT, 3_500_000)
        self.assertEqual(est.TL, 2_675_000)
        self.assertEqual(est.VAT, round(56_175_000 * 0.08))
        self.assertEqual(est.G_XD, 56_175_000 + est.VAT)


class QSAgentTest(unittest.TestCase):

    def run_phase(self, demo, **kwargs):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        sup = AECSupervisor(project_root=ROOT, human_gate_mode="auto",
                            persist_path=os.path.join(tmp.name, "s.json"), demo_mode=demo)
        with contextlib.redirect_stdout(io.StringIO()):
            sup.register_agent(QSAgent(**kwargs))
            ok = sup.run(phases=[ProjectPhase.QS_ESTIMATE])
        return ok, sup

    def test_real_file(self):
        ok, sup = self.run_phase(False, qs_path=TEMPLATE)
        self.assertTrue(ok)
        qs = sup.bus.get_qs_data()
        self.assertEqual(qs.direct_cost_T_vnd, 27_284_103_489)
        self.assertEqual(qs.items_count, 32)
        self.assertEqual(sup.bus.get_sample_data_sources(), [])

    def test_rate_override(self):
        ok, sup = self.run_phase(False, qs_path=TEMPLATE, rate_overrides={"vat": 8})
        self.assertTrue(ok)
        self.assertEqual(sup.bus.get_qs_data().rates["vat"], 0.08)

    def test_missing_rates_stop(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "qs.csv")
            with open(path, "w", encoding="utf-8") as f:
                f.write("Mã hiệu,Nội dung,ĐVT,Khối lượng,Đơn giá\nA1,Việc A,m3,1,100\n")
            ok, sup = self.run_phase(False, qs_path=path)
        self.assertFalse(ok)
        self.assertIn("Thiếu tỷ lệ", sup.bus.get_errors()[0])

    def test_no_file_real_mode_and_demo(self):
        ok, sup = self.run_phase(False)
        self.assertFalse(ok)
        self.assertEqual(sup.bus.get_node_status("qs_agent"), NodeStatus.FAILED)
        ok, sup = self.run_phase(True)
        self.assertTrue(ok)
        self.assertEqual([s["agent_id"] for s in sup.bus.get_sample_data_sources()], ["qs_agent"])

    def test_gate_uses_actual_rates(self):
        gate = QualityGate(ROOT)
        rates = {"chung": 0.05, "nha_tam": 0.01, "kxd": 0.01, "tl": 0.05, "vat": 0.08}
        good = {"direct_cost_T_vnd": 1000, "indirect_cost_GT_vnd": 70, "tax_TL_vnd": 53.5,
                "vat_vnd": 89.88, "total_G_XD_vnd": 1213.38, "rates": rates, "items_count": 1}
        self.assertTrue(gate.check_qs_estimate(good).passed)
        bad = dict(good, indirect_cost_GT_vnd=90)
        self.assertFalse(gate.check_qs_estimate(bad).passed)
        self.assertFalse(gate.check_qs_estimate(dict(good, rates={})).passed)


if __name__ == "__main__":
    unittest.main()
