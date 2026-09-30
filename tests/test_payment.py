# -*- coding: utf-8 -*-
"""Kiểm thử Mẫu 03.a/TT (NĐ 254/2025) — python -m unittest discover tests"""

import contextlib
import io
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.payment_agent import PaymentAgent
from core.gates.quality_gate import QualityGate
from core.state.shared_state import NodeStatus, ProjectPhase
from core.supervisor.supervisor_agent import AECSupervisor
from tools.payment import PaymentError, compute_payment, load_progress, write_payment_workbook
from tools.qs_loader import apply_rate_overrides, load_qs

QS_CSV = ("STT;Mã hiệu;Nội dung công tác;ĐVT;Khối lượng;Đơn giá\n"
          "1;AF.11111;Bê tông lót;m3;10;1000000\n"
          "2;AF.61111;Cốt thép;Tấn;2;20000000\n"
          "3;AB.11111;Đào đất;m3;100;100000\n")
RATES = {"chung": 5, "nha_tam": 1, "kxd": 1, "tl": 5, "vat": 10}


class PaymentTest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.qs = self.write("qs.csv", QS_CSV)
        self.est = load_qs(self.qs)
        apply_rate_overrides(self.est, RATES)
        self.est.compute()

    def write(self, name, text):
        path = os.path.join(self.tmp.name, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path

    def progress(self, text):
        return load_progress(self.write("kl.csv", "Mã hiệu;Nội dung;KL lũy kế kỳ trước;KL thực hiện kỳ này\n" + text))

    def test_contract_basis(self):
        res = compute_payment(self.est, self.progress("AF.11111;;0;4\nAF.61111;;1;1\n"), "contract", 20, 5)
        self.assertEqual(res.this_value, 4 * 1_000_000 + 1 * 20_000_000)
        self.assertEqual(res.this_vat, 2_400_000)
        self.assertEqual(res.this_total, 26_400_000)
        self.assertEqual(res.advance_recovery, 5_280_000)
        self.assertEqual(res.retention, 1_320_000)
        self.assertEqual(res.payable, 26_400_000 - 5_280_000 - 1_320_000)
        self.assertEqual(res.cumulative_value, 4_000_000 + 2 * 20_000_000)
        self.assertEqual(res.contract_value, 60_000_000)

    def test_direct_basis_uses_g_over_t(self):
        res = compute_payment(self.est, self.progress("AF.11111;;0;10\n"), "direct", 0, 0)
        factor = self.est.G / self.est.T
        self.assertAlmostEqual(res.price_factor, factor)
        self.assertEqual(res.this_value, 10 * round(1_000_000 * factor))

    def test_overrun_and_unmatched_are_not_paid(self):
        res = compute_payment(self.est, self.progress("AF.11111;;8;5\nXX.1;Đường tạm;0;7\n"), "contract", 0, 0)
        line = next(l for l in res.lines if l.item.code == "AF.11111")
        self.assertEqual(line.this_qty, 2)          # chỉ còn 2 m3 trong hợp đồng
        self.assertEqual(line.overrun_qty, 3)
        self.assertEqual(res.overrun_value, 3_000_000)
        self.assertEqual([p.code for p in res.unmatched], ["XX.1"])
        self.assertEqual(res.this_value, 2_000_000)
        self.assertEqual(len(res.warnings), 2)

    def test_advance_outstanding_caps_recovery(self):
        res = compute_payment(self.est, self.progress("AF.61111;;0;2\n"), "contract", 20, 0,
                              advance_outstanding=1_000_000)
        self.assertEqual(res.advance_recovery, 1_000_000)

    def test_negative_payable_is_error(self):
        res = compute_payment(self.est, self.progress("AF.61111;;0;2\n"), "contract", 90, 20)
        self.assertTrue(res.errors)

    def test_bad_progress_and_ambiguous_match(self):
        with self.assertRaises(PaymentError):
            self.progress("AF.11111;;0;-1\n")
        est = load_qs(self.write("qs2.csv", "STT;Nội dung công tác;ĐVT;Khối lượng;Đơn giá\n"
                                            "1;Bê tông;m3;1;10\n2;Bê tông;m3;2;10\n"))
        est.rates["vat"] = 0.1
        prog = load_progress(self.write("kl2.csv", "Nội dung;KL thực hiện kỳ này\nBê tông;1\n"))
        self.assertTrue(compute_payment(est, prog, "contract", 0, 0).errors)

    def test_two_periods(self):
        p1 = compute_payment(self.est, self.progress("AF.11111;;0;6\n"), "contract", 0, 0, period="01")
        prev = next(l.cumulative_qty for l in p1.lines if l.item.code == "AF.11111")
        p2 = compute_payment(self.est, self.progress(f"AF.11111;;{prev};4\n"), "contract", 0, 0, period="02")
        self.assertEqual(p1.cumulative_value + p2.this_value, p2.cumulative_value)
        self.assertEqual(p2.cumulative_value, 10_000_000)

    def test_workbook_and_gate(self):
        res = compute_payment(self.est, self.progress("AF.11111;;0;4\nXX.1;;0;1\n"), "contract", 20, 5, period="01")
        path = os.path.join(self.tmp.name, "03a.xlsx")
        write_payment_workbook(path, res)
        import openpyxl
        wb = openpyxl.load_workbook(path)
        self.assertEqual(wb.sheetnames, ["PHU_LUC_03A", "PHAT_SINH_CANH_BAO"])
        summary = {"this_value": res.this_value, "line_values": [l.this_value for l in res.lines],
                   "vat_rate": res.vat_rate, "this_vat": res.this_vat, "this_total": res.this_total,
                   "advance_recovery": res.advance_recovery, "retention": res.retention, "payable": res.payable,
                   "cumulative_value": res.cumulative_value, "contract_value": res.contract_value,
                   "unmatched": 1, "overrun_value": 0}
        gate = QualityGate(ROOT)
        self.assertTrue(gate.check_payment(summary).passed)
        self.assertFalse(gate.check_payment(dict(summary, payable=summary["payable"] + 1)).passed)


class PaymentAgentTest(unittest.TestCase):

    def run_phase(self, demo, **kwargs):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        sup = AECSupervisor(project_root=ROOT, human_gate_mode="auto",
                            persist_path=os.path.join(tmp.name, "s.json"), demo_mode=demo)
        with contextlib.redirect_stdout(io.StringIO()):
            sup.register_agent(PaymentAgent(**kwargs))
            ok = sup.run(phases=[ProjectPhase.PAYMENT_03A])
        return ok, sup

    def test_real_mode_requires_explicit_terms(self):
        ok, sup = self.run_phase(False)
        self.assertFalse(ok)
        self.assertEqual(sup.bus.get_node_status("payment_agent"), NodeStatus.FAILED)
        self.assertIn("--price-basis", sup.bus.get_errors()[0])

    def test_real_mode_full(self):
        with tempfile.TemporaryDirectory() as tmp:
            qs = os.path.join(tmp, "qs.csv")
            kl = os.path.join(tmp, "kl.csv")
            out = os.path.join(tmp, "03a.xlsx")
            with open(qs, "w", encoding="utf-8") as f:
                f.write(QS_CSV)
            with open(kl, "w", encoding="utf-8") as f:
                f.write("Mã hiệu,KL thực hiện kỳ này\nAF.11111,5\n")
            ok, sup = self.run_phase(False, qs_path=qs, progress_path=kl, price_basis="contract",
                                     advance_recovery_pct=0, retention_pct=5,
                                     rate_overrides={"vat": 8}, payment_out=out)
            self.assertTrue(ok)
            self.assertTrue(os.path.exists(out))
        summary = sup.bus.get_qs_data().payment_summary
        self.assertEqual(summary["this_value"], 5_000_000)
        self.assertEqual(summary["this_vat"], 400_000)
        self.assertEqual(summary["payable"], 5_400_000 - 270_000)
        self.assertEqual(sup.bus.get_sample_data_sources(), [])

    def test_demo_marks_sample(self):
        ok, sup = self.run_phase(True)
        self.assertTrue(ok)
        self.assertEqual([s["agent_id"] for s in sup.bus.get_sample_data_sources()], ["payment_agent"])


if __name__ == "__main__":
    unittest.main()
