# -*- coding: utf-8 -*-
"""
TÁC TỬ: AEC EQUIPMENT & FLEET AGENT (ĐIỀU PHỐI XE MÁY & NHIÊN LIỆU CÔNG TRƯỜNG)
Hệ thống Multi-Agent AEC (23HG-multiagent-system).

Nhiệm vụ:
1. Tiếp nhận khối lượng bóc tách hình học (BOQ / Takeoff) từ Agent QS.
2. Tra cứu cơ sở dữ liệu định mức ca máy (Vincons_ĐMGK_02-01) và định mức tiêu hao dầu máy.
3. Tính toán tổng ca máy, số máy huy động hàng ngày theo khung thời gian yêu cầu và số ca/ngày.
4. Cân đối biểu đồ phụ tải thiết bị (Equipment Leveling), phát hiện nút thắt cổ chai (bottlenecks).
5. Dự trù và phân kỳ tiêu thụ nhiên liệu Dầu Diezel (lít/ngày, lít/kỳ).
6. Tự động xuất file Excel Master 5 Sheet liên kết động 100% và file MS Project XML.
"""

from __future__ import annotations

import os
import json
import datetime
from typing import Any, Dict, List, Optional

from tools.equipment_fleet_scheduler import (
    EquipmentFleetScheduler,
    FleetTask,
    DailyFleetMatrix,
)


class AECEquipmentFleetAgent:
    """
    Tác tử chuyên trách Điều phối Xe máy, Nhiên liệu & Kế hoạch phụ tải công trường.
    """

    def __init__(
        self,
        norms_path: Optional[str] = None,
        fuel_norms_path: Optional[str] = None,
        material_factors_path: Optional[str] = None,
    ):
        self.scheduler = EquipmentFleetScheduler(
            norms_path=norms_path,
            fuel_norms_path=fuel_norms_path,
            material_factors_path=material_factors_path,
        )
        self.last_matrix: Optional[DailyFleetMatrix] = None
        self.last_tasks: List[FleetTask] = []

    def plan_project_fleet(
        self,
        project_name: str,
        section_name: str,
        tasks_data: List[Dict[str, Any]],
        start_date: datetime.date,
        end_date: datetime.date,
        shifts_per_day: int = 2,
        output_excel_path: Optional[str] = None,
        output_xml_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Lập kế hoạch ca xe, ca máy & nhiên liệu cho dự án.

        tasks_data format:
        [
            {
                "task_id": "1",
                "code": "ĐM-16",
                "name": "Đào đất móng cống",
                "unit": "m3",
                "quantity": 73800.0,
                "start_date": "2026-09-20",
                "end_date": "2026-10-05",
                "task_norm_key": "dao_mong_cong",
                "labor_per_shift": 4.0,
                "shifts_per_day": 2
            }, ...
        ]
        """
        fleet_tasks: List[FleetTask] = []
        for t_dict in tasks_data:
            s_d = datetime.date.fromisoformat(str(t_dict["start_date"]))
            e_d = datetime.date.fromisoformat(str(t_dict["end_date"]))
            ft = FleetTask(
                task_id=str(t_dict.get("task_id", len(fleet_tasks) + 1)),
                code=str(t_dict.get("code", "")),
                name=str(t_dict.get("name", "")),
                unit=str(t_dict.get("unit", "")),
                quantity=float(t_dict.get("quantity", 0.0)),
                start_date=s_d,
                end_date=e_d,
                shifts_per_day=int(t_dict.get("shifts_per_day", shifts_per_day)),
                task_norm_key=t_dict.get("task_norm_key"),
                custom_productivity=t_dict.get("custom_productivity"),
                labor_per_shift=float(t_dict.get("labor_per_shift", 4.0)),
                notes=str(t_dict.get("notes", "")),
            )
            fleet_tasks.append(ft)

        self.last_tasks = fleet_tasks
        self.last_matrix = self.scheduler.build_daily_matrix(fleet_tasks, start_date, end_date)

        # Export Excel if path provided
        excel_exported = None
        if output_excel_path:
            excel_exported = self.scheduler.export_excel(
                project_name=project_name,
                section_name=section_name,
                tasks=fleet_tasks,
                start_date=start_date,
                end_date=end_date,
                output_path=output_excel_path,
                shifts_per_day=shifts_per_day,
            )

        # Export XML if path provided
        xml_exported = None
        if output_xml_path:
            xml_exported = self.scheduler.export_ms_project_xml(
                project_name=project_name,
                tasks=fleet_tasks,
                start_date=start_date,
                end_date=end_date,
                output_path=output_xml_path,
            )

        # Generate summary metrics
        max_daily_labor = max(self.last_matrix.daily_labor.values()) if self.last_matrix.daily_labor else 0
        avg_daily_fuel = (
            sum(self.last_matrix.daily_fuel.values()) / len(self.last_matrix.dates)
            if self.last_matrix.dates
            else 0
        )

        return {
            "project_name": project_name,
            "section_name": section_name,
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "total_days": len(self.last_matrix.dates),
            "shifts_per_day": shifts_per_day,
            "total_machine_shifts": self.last_matrix.total_shifts,
            "total_fuel_liters_net": self.last_matrix.total_fuel_liters,
            "total_fuel_liters_with_safety": round(self.last_matrix.total_fuel_liters * 1.05, 1),
            "peak_fleet": self.last_matrix.peak_machines,
            "peak_daily_labor_persons": max_daily_labor,
            "avg_daily_fuel_liters": round(avg_daily_fuel, 1),
            "excel_path": excel_exported,
            "xml_path": xml_exported,
            "status": "SUCCESS",
        }

    def generate_executive_report_markdown(self, plan_result: Dict[str, Any]) -> str:
        """
        Sinh báo cáo tóm tắt điều hành bằng định dạng Markdown.
        """
        p_name = plan_result.get("project_name", "")
        s_name = plan_result.get("section_name", "")
        tot_days = plan_result.get("total_days", 0)
        tot_shifts = plan_result.get("total_machine_shifts", 0)
        fuel_net = plan_result.get("total_fuel_liters_net", 0)
        fuel_buf = plan_result.get("total_fuel_liters_with_safety", 0)
        peak_labor = plan_result.get("peak_daily_labor_persons", 0)
        peak_fleet = plan_result.get("peak_fleet", {})

        fleet_lines = []
        for code, count in peak_fleet.items():
            fleet_lines.append(f"- **{code}:** Huy động đỉnh điểm **{count:.1f} máy**")
        fleet_txt = "\n".join(fleet_lines)

        md = f"""# BÁO CÁO ĐIỀU PHỐI CA XE, CA MÁY & NHIÊN LIỆU DẦU DIEZEL
## DỰ ÁN: {p_name} | HẠNG MỤC: {s_name}
### TÁC TỬ ĐIỀU PHỐI: AEC EQUIPMENT & FLEET AGENT v1.0

---

### 1. CHỈ SỐ HOẠT ĐỘNG CHÍNH
- **Thời gian thi công:** Từ {plan_result.get('start_date')} đến {plan_result.get('end_date')} (**{tot_days} ngày**).
- **Chế độ làm việc:** {plan_result.get('shifts_per_day')} ca/ngày (20h/ngày xoay ca liên tục).
- **Tổng số ca máy yêu cầu:** **{tot_shifts:,.1f} ca máy**.
- **Tổng lượng Dầu Diezel định mức:** **{fuel_net:,.1f} lít**.
- **Tổng lượng Dầu Diezel dự trù cấp (kèm hao hụt an toàn 5%):** **{fuel_buf:,.1f} lít**.
- **Quân số nhân lực huy động đỉnh điểm:** **{peak_labor:,.0f} người**.

---

### 2. PHỤ TẢI MÁY MÓC THIẾT BỊ ĐỈNH ĐIỂM (PEAK FLEET)
{fleet_txt}

---

### 3. TỆP DỮ LIỆU ĐÃ XUẤT RA HỆ THỐNG
- File Bảng tính Master Excel: `{plan_result.get('excel_path')}`
- File Tiến độ MS Project XML: `{plan_result.get('xml_path')}`
"""
        return md
