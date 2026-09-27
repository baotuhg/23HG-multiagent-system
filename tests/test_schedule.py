# -*- coding: utf-8 -*-
"""Kiểm thử CPMCalculator và bộ đọc tiến độ thật — chạy: python -m unittest discover tests"""

import contextlib
import io
import os
import sys
import tempfile
import unittest
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.sub_agents import SchedulerAgent
from core.state.shared_state import ProjectPhase
from core.supervisor.supervisor_agent import AECSupervisor
from tools.cpm_calculator import CPMCalculator, parse_predecessor
from tools.schedule_loader import ScheduleLoadError, find_date_violations, load_schedule

TEMPLATE_XML = os.path.join(ROOT, "templates", "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml")
TEMPLATE_XLSX = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")


def by_id(result):
    return {t.task_id: t for t in result.tasks}


class CPMCalculatorTest(unittest.TestCase):

    def test_parse_predecessor(self):
        cases = {
            "T01": ("T01", "FS", 0), "1.3SS+3d": ("1.3", "SS", 3), "2.1 FF -2 ngày": ("2.1", "FF", -2),
            "A SF+1,5d": ("A", "SF", 1.5), "1.2fs": ("1.2", "FS", 0),
        }
        for text, (pid, kind, lag) in cases.items():
            link = parse_predecessor(text)
            self.assertEqual((link.pred_id, link.type, link.lag_days), (pid, kind, lag), text)

    def test_link_types_and_lags(self):
        res = CPMCalculator().calculate([
            {"id": "A", "duration": 10},
            {"id": "B", "duration": 4, "predecessors": ["ASS+3d"]},     # ES = 3
            {"id": "C", "duration": 5, "predecessors": ["AFF+2d"]},     # EF = 12 → ES = 7
            {"id": "D", "duration": 6, "predecessors": ["ASF+8d"]},     # EF = 8 → ES = 2
            {"id": "E", "duration": 2, "predecessors": ["AFS-2d"]},     # ES = 8
            {"id": "F", "duration": 1, "predecessors": ["B", "C", "D", "E"]},
        ])
        t = by_id(res)
        self.assertEqual(res.status, "OK")
        self.assertEqual([(t[k].es, t[k].ef) for k in "ABCDE"], [(0, 10), (3, 7), (7, 12), (2, 8), (8, 10)])
        self.assertEqual(t["F"].es, 12)
        self.assertEqual(res.total_duration_days, 13)
        self.assertEqual(res.critical_path, ["A", "C", "F"])
        self.assertEqual(t["B"].tf, 5)

    def test_errors_are_not_silently_ignored(self):
        calc = CPMCalculator()
        self.assertEqual(calc.calculate([{"id": "A", "duration": 1, "predecessors": ["X"]}]).status, "ERROR")
        self.assertEqual(calc.calculate([
            {"id": "A", "duration": 1, "predecessors": ["B"]},
            {"id": "B", "duration": 1, "predecessors": ["A"]},
        ]).status, "ERROR")
        self.assertEqual(calc.calculate([{"id": "A", "duration": 1}, {"id": "A", "duration": 2}]).status, "ERROR")
        self.assertEqual(calc.calculate([]).status, "ERROR")

    def test_calendar_skips_sundays_and_holidays(self):
        res = CPMCalculator().calculate(
            [{"id": "A", "duration": 3}, {"id": "B", "duration": 2, "predecessors": ["A"]}],
            start_date_str="2026-10-01",                       # thứ Năm
            non_working_weekdays={6}, holidays={date(2026, 10, 5)},
        )
        t = by_id(res)
        self.assertEqual((t["A"].start_date, t["A"].finish_date), ("2026-10-01", "2026-10-03"))
        self.assertEqual((t["B"].start_date, t["B"].finish_date), ("2026-10-06", "2026-10-07"))
        self.assertEqual(res.project_finish, "2026-10-07")

    def test_old_call_style_still_works(self):
        res = CPMCalculator().calculate(
            [{"id": "T01", "duration": 3, "predecessors": []},
             {"id": "T02", "duration": 7, "predecessors": ["T01"]}],
            start_date_str="2026-10-01",
        )
        self.assertEqual(res.critical_path, ["T01", "T02"])
        self.assertEqual(res.project_finish, "2026-10-10")


class ScheduleLoaderTest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, name, text):
        path = os.path.join(self.tmp.name, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path

    def test_csv_vietnamese_headers(self):
        path = self.write("td.csv", (
            "Mã WBS;Danh mục công tác thi công;Thời gian Duration (ngày);Ngày bắt đầu;Quan hệ logic Predecessors\n"
            "1.0;;;;\n"
            "1.1;Tim mốc;2;2026-10-01;-\n"
            "1.2;Dọn mặt bằng;6;2026-10-03;1.1FS\n"
            "1.4;Bãi đúc;13;2026-10-12;\"1.2SS+3d; 1.1FS\"\n"
        ))
        r = load_schedule(path)
        self.assertEqual(r.errors, [])
        self.assertEqual([t["id"] for t in r.tasks], ["1.1", "1.2", "1.4"])
        self.assertEqual(r.tasks[2]["predecessors"],
                         [{"id": "1.2", "type": "SS", "lag": 3.0}, {"id": "1.1", "type": "FS", "lag": 0.0}])
        self.assertEqual(r.project_start, "2026-10-01")

    def test_bad_rows_are_reported(self):
        path = self.write("td.csv", "id,name,duration,predecessors\nA,Việc A,abc,\nB,Việc B,2,A??\n")
        r = load_schedule(path)
        self.assertEqual(len(r.errors), 1)                # thời gian 'abc'
        res = CPMCalculator().calculate(r.tasks)          # 'A??' không tồn tại → CPM báo lỗi
        self.assertEqual(res.status, "ERROR")
        self.assertIn("A??", res.warnings[0])

    def test_msproject_xml_link_types_and_lag(self):
        path = self.write("td.xml", """<?xml version="1.0"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <StartDate>2026-10-01T08:00:00</StartDate><MinutesPerDay>480</MinutesPerDay>
  <Tasks>
    <Task><UID>0</UID><Name>Dự án</Name><Summary>1</Summary></Task>
    <Task><UID>1</UID><Name>Nhóm</Name><Summary>1</Summary><Duration>PT80H0M0S</Duration></Task>
    <Task><UID>2</UID><Name>A</Name><WBS>1.1</WBS><Duration>PT40H0M0S</Duration></Task>
    <Task><UID>3</UID><Name>B</Name><WBS>1.2</WBS><Duration>PT16H0M0S</Duration>
      <PredecessorLink><PredecessorUID>2</PredecessorUID><Type>3</Type><LinkLag>9600</LinkLag></PredecessorLink>
    </Task>
  </Tasks>
</Project>""")
        r = load_schedule(path)
        self.assertEqual(r.errors, [])
        self.assertEqual(r.project_start, "2026-10-01")
        self.assertEqual([(t["code"], t["duration"]) for t in r.tasks], [("1.1", 5.0), ("1.2", 2.0)])
        self.assertEqual(r.tasks[1]["predecessors"], [{"id": "2", "type": "SS", "lag": 2.0}])

    def test_repository_templates(self):
        xml = load_schedule(TEMPLATE_XML)
        self.assertEqual((len(xml.tasks), xml.errors), (30, []))
        self.assertEqual(len(find_date_violations(xml.tasks)), 4)   # FS nhưng ngày chồng lấn
        res = CPMCalculator().calculate(xml.tasks, start_date_str=xml.project_start)
        self.assertEqual(res.status, "OK")
        self.assertEqual(res.total_duration_days, 187)

        # Sheet Excel: thời gian là công thức chưa có kết quả → được tự tính (tools/excel_eval)
        xlsx = load_schedule(TEMPLATE_XLSX)
        self.assertEqual((len(xlsx.tasks), xlsx.errors), (30, []))
        self.assertEqual(find_date_violations(xlsx.tasks), [])   # Excel ghi 1.3SS+3d đúng ý đồ
        res = CPMCalculator().calculate(xlsx.tasks, start_date_str=xlsx.project_start)
        self.assertEqual((res.status, res.total_duration_days), ("OK", 186))

    def test_unknown_format(self):
        with self.assertRaises(ScheduleLoadError):
            load_schedule(self.write("td.txt", "không có tiêu đề\n1,2,3\n"))


class SchedulerAgentRealDataTest(unittest.TestCase):

    def test_real_schedule_file_in_real_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "cpm.csv")
            sup = AECSupervisor(project_root=ROOT, human_gate_mode="auto",
                                persist_path=os.path.join(tmp, "state.json"))
            with contextlib.redirect_stdout(io.StringIO()):
                sup.register_agent(SchedulerAgent(schedule_path=TEMPLATE_XML, non_working_weekdays={6},
                                                  schedule_out=out))
                ok = sup.run(phases=[ProjectPhase.SCHEDULE_CPM])
            self.assertTrue(ok)
            self.assertEqual(sup.bus.get_sample_data_sources(), [])
            sched = sup.bus.get_schedule_data()
            self.assertEqual(sched.total_duration_days, 187)
            self.assertEqual(sched.start_date, "2026-10-01")
            self.assertEqual(len(sched.tasks), 30)
            self.assertTrue(os.path.exists(out))


if __name__ == "__main__":
    unittest.main()
