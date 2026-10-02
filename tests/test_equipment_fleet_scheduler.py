# -*- coding: utf-8 -*-
"""
UNIT TESTS — EQUIPMENT FLEET & FUEL SCHEDULER
Kiểm thử tính toán ca xe, ca máy, định mức nhiên liệu và tác tử AEC Equipment Fleet Agent.
"""

import os
import unittest
import datetime
import tempfile

from tools.equipment_fleet_scheduler import (
    EquipmentFleetScheduler,
    FleetTask,
    MachineAllocation,
    DailyFleetMatrix
)
from core.agents.aec_equipment_fleet_agent import AECEquipmentFleetAgent


class TestEquipmentFleetScheduler(unittest.TestCase):

    def setUp(self):
        self.scheduler = EquipmentFleetScheduler()
        self.agent = AECEquipmentFleetAgent()

    def test_load_norms(self):
        """Kiểm tra nạp cơ sở dữ liệu định mức Vincons & dầu máy."""
        self.assertIn("tasks", self.scheduler.norms_data)
        self.assertIn("dao_mong_cong", self.scheduler.norms_data["tasks"])
        self.assertIn("equipment_rates", self.scheduler.fuel_data)
        self.assertEqual(self.scheduler.get_fuel_rate("M1"), 89.6)
        self.assertEqual(self.scheduler.get_fuel_rate("M3"), 124.0)

    def test_compute_task_allocations(self):
        """Kiểm tra tính toán số ca máy và máy huy động hàng ngày."""
        task = FleetTask(
            task_id="1",
            code="ĐM-16",
            name="Đào đất hố móng cống hộp",
            unit="m3",
            quantity=73800.0,
            start_date=datetime.date(2026, 9, 20),
            end_date=datetime.date(2026, 10, 5),
            shifts_per_day=2,
            task_norm_key="dao_mong_cong"
        )
        task = self.scheduler.compute_task_allocations(task)
        self.assertEqual(task.duration_days, 16)
        self.assertTrue(len(task.machine_allocations) > 0)

        m1 = task.machine_allocations[0]
        self.assertEqual(m1.machine_code, "M1")
        # 73800 / 401.07 ≈ 184.0 ca
        self.assertAlmostEqual(m1.total_shifts_required, 184.01, places=1)
        # daily machines = 184.01 / (16 * 2) ≈ 5.75 máy
        self.assertAlmostEqual(m1.daily_machines_needed, 5.75, places=2)
        # fuel rate = 89.6 l/shift
        self.assertEqual(m1.fuel_rate_l_per_shift, 89.6)
        # total fuel = 184.01 * 89.6 ≈ 16487.3 lít
        self.assertAlmostEqual(m1.total_fuel_liters, 16487.3, places=0)

    def test_build_daily_matrix(self):
        """Kiểm tra lập ma trận phân bổ ngày và phụ tải thiết bị."""
        tasks = [
            FleetTask(
                task_id="1",
                code="ĐM-16",
                name="Đào cống",
                unit="m3",
                quantity=73800.0,
                start_date=datetime.date(2026, 9, 20),
                end_date=datetime.date(2026, 10, 5),
                shifts_per_day=2,
                task_norm_key="dao_mong_cong"
            ),
            FleetTask(
                task_id="2",
                code="ĐM-29",
                name="Vận chuyển đất <5km",
                unit="m3",
                quantity=86450.0,
                start_date=datetime.date(2026, 9, 20),
                end_date=datetime.date(2026, 10, 5),
                shifts_per_day=2,
                task_norm_key="van_chuyen_dat_5km"
            )
        ]
        start_d = datetime.date(2026, 9, 20)
        end_d = datetime.date(2026, 10, 25)
        matrix = self.scheduler.build_daily_matrix(tasks, start_d, end_d)

        self.assertEqual(len(matrix.dates), 36)
        self.assertIn("M1", matrix.daily_machines)
        self.assertIn("M7", matrix.daily_machines)
        # On 2026-09-20, both tasks active
        self.assertTrue(matrix.daily_machines["M1"][start_d] > 5.0)
        self.assertTrue(matrix.daily_machines["M7"][start_d] > 4.0)
        # On 2026-10-20, both tasks finished
        d_later = datetime.date(2026, 10, 20)
        self.assertEqual(matrix.daily_machines["M1"][d_later], 0.0)
        self.assertTrue(matrix.total_fuel_liters > 0)

    def test_agent_end_to_end_planning(self):
        """Kiểm tra quy trình làm việc trọn gói của AECEquipmentFleetAgent."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            excel_out = os.path.join(tmp_dir, "test_fleet.xlsx")
            xml_out = os.path.join(tmp_dir, "test_fleet.xml")

            tasks_data = [
                {
                    "task_id": "1",
                    "code": "ĐM-16",
                    "name": "Đào móng cống hộp",
                    "unit": "m3",
                    "quantity": 73800.0,
                    "start_date": "2026-09-20",
                    "end_date": "2026-10-05",
                    "task_norm_key": "dao_mong_cong",
                    "shifts_per_day": 2,
                    "labor_per_shift": 4.0
                },
                {
                    "task_id": "2",
                    "code": "ĐM-BT",
                    "name": "Đổ bê tông thân cống B20",
                    "unit": "m3",
                    "quantity": 11870.0,
                    "start_date": "2026-09-26",
                    "end_date": "2026-10-17",
                    "task_norm_key": "do_be_tong_than_cong",
                    "shifts_per_day": 2,
                    "labor_per_shift": 12.0
                }
            ]

            res = self.agent.plan_project_fleet(
                project_name="CỐNG HỘP TUYẾN A5",
                section_name="CỐNG ĐƠN & ĐÔI",
                tasks_data=tasks_data,
                start_date=datetime.date(2026, 9, 20),
                end_date=datetime.date(2026, 10, 25),
                shifts_per_day=2,
                output_excel_path=excel_out,
                output_xml_path=xml_out
            )

            self.assertEqual(res["status"], "SUCCESS")
            self.assertEqual(res["total_days"], 36)
            self.assertTrue(os.path.exists(excel_out))
            self.assertTrue(os.path.exists(xml_out))
            self.assertTrue(os.path.getsize(excel_out) > 5000)
            self.assertTrue(os.path.getsize(xml_out) > 500)

            # Test report generation
            rpt = self.agent.generate_executive_report_markdown(res)
            self.assertIn("BÁO CÁO ĐIỀU PHỐI CA XE", rpt)
            self.assertIn("CỐNG HỘP TUYẾN A5", rpt)


if __name__ == "__main__":
    unittest.main()
