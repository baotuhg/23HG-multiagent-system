# -*- coding: utf-8 -*-
"""
Kiểm thử chế độ dữ liệu thật / demo của State Graph:
  - Không --demo: thiếu dữ liệu thật → dừng ngay, không retry, không dùng dữ liệu mẫu.
  - --demo: được dùng dữ liệu mẫu nhưng mọi chỗ dùng đều được ghi nhận.
"""

import contextlib
import io
import os
import sys
import tempfile
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.asbuilt_agent import AsBuiltAgent
from core.agents.rebar_agent import RebarAgent
from core.agents.sub_agents import BPTCKCSAgent, CADAgent, QSAgent, SchedulerAgent
from core.state.shared_state import NodeStatus, ProjectPhase, ProjectSharedState
from core.state.state_bus import StateBus
from core.supervisor.supervisor_agent import AECSupervisor

SAMPLE_EXCEL = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")


class DataModeTest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.persist = os.path.join(self.tmp.name, "state.json")

    def tearDown(self):
        self.tmp.cleanup()

    def make_supervisor(self, demo, bbs_path=None, excel=""):
        sup = AECSupervisor(
            project_root=ROOT, excel_master_path=excel, human_gate_mode="auto",
            persist_path=self.persist, demo_mode=demo,
        )
        with contextlib.redirect_stdout(io.StringIO()):
            for agent in (CADAgent(), RebarAgent(bbs_path=bbs_path), QSAgent(), BPTCKCSAgent(),
                          SchedulerAgent(), AsBuiltAgent()):
                sup.register_agent(agent)
        return sup

    def run_quiet(self, sup, phases=None):
        with contextlib.redirect_stdout(io.StringIO()):
            return sup.run(phases=phases)

    def test_real_mode_stops_without_sample_data_or_retries(self):
        sup = self.make_supervisor(demo=False)
        started = time.monotonic()
        ok = self.run_quiet(sup)
        self.assertFalse(ok)
        self.assertLess(time.monotonic() - started, 1.5, "lỗi thiếu dữ liệu không được retry/backoff")
        self.assertEqual(sup.bus.get_node_status("cad_agent"), NodeStatus.FAILED)
        self.assertEqual(sup.bus.get_sample_data_sources(), [])
        self.assertEqual(len([r for r in sup.bus._state.run_history if r["agent_id"] == "cad_agent"]), 1)
        self.assertIn("THIẾU DỮ LIỆU THẬT", sup.bus.get_errors()[0])

    def test_real_mode_every_sample_backed_agent_refuses(self):
        for phase, agent_id in [
            (ProjectPhase.REBAR_CUT, "rebar_agent"),
            (ProjectPhase.QS_ESTIMATE, "qs_agent"),
            (ProjectPhase.QAQC_REVIEW, "bptc_kcs_agent"),
            (ProjectPhase.SCHEDULE_CPM, "scheduler_agent"),
            (ProjectPhase.ASBUILT_LOOP, "asbuilt_agent"),
        ]:
            with self.subTest(agent=agent_id):
                sup = self.make_supervisor(demo=False)
                self.assertFalse(self.run_quiet(sup, phases=[phase]))
                self.assertEqual(sup.bus.get_node_status(agent_id), NodeStatus.FAILED)
                self.assertEqual(sup.bus.get_sample_data_sources(), [])

    def test_real_mode_rebar_with_real_bbs(self):
        bbs = os.path.join(self.tmp.name, "bbs.csv")
        with open(bbs, "w", encoding="utf-8") as f:
            f.write("mark,diameter_mm,grade,length_mm,quantity\nT1,20,CB400-V,4500,30\nT2,20,CB400-V,3200,50\n")
        sup = self.make_supervisor(demo=False, bbs_path=bbs)
        self.assertTrue(self.run_quiet(sup, phases=[ProjectPhase.REBAR_CUT]))
        cutting = sup.bus.get_cutting_dict()
        self.assertEqual(cutting["pieces_cut"], 80)
        self.assertEqual(cutting["status"], "OPTIMAL")
        self.assertEqual(sup.bus.get_sample_data_sources(), [])

    def test_demo_mode_records_all_sample_sources(self):
        sup = self.make_supervisor(demo=True, excel=SAMPLE_EXCEL)
        self.assertTrue(self.run_quiet(sup))
        used = {s["agent_id"] for s in sup.bus.get_sample_data_sources()}
        self.assertEqual(used, {"cad_agent", "rebar_agent", "qs_agent", "bptc_kcs_agent",
                                "scheduler_agent", "asbuilt_agent"})

    def test_sample_data_cannot_be_marked_outside_demo(self):
        bus = StateBus(ProjectSharedState(demo_mode=False))
        with self.assertRaises(RuntimeError):
            bus.mark_sample_data("x", "y")


if __name__ == "__main__":
    unittest.main()
