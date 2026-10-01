# -*- coding: utf-8 -*-
"""
EQUIPMENT FLEET & FUEL SCHEDULER — Bộ công cụ Quản trị Ca xe, Ca máy & Kế hoạch Nhiên liệu
Dành cho Hệ thống Multi-Agent AEC (23HG-multiagent-system).
Pure Python, Zero LLM — Tính toán kỹ thuật xác định 100%.

Tính toán:
  1. Số ca máy yêu cầu từ Khối lượng BOQ và Định mức năng suất (Vincons / BXD).
  2. Số lượng máy móc thiết bị (MMTB) cần huy động hàng ngày theo khung thời gian & số ca/ngày.
  3. Ma trận phân bổ phụ tải MMTB và Nhân công theo ngày trên toàn bộ timeline (Gantt).
  4. Dự trù và phân kỳ tiêu thụ nhiên liệu Dầu Diezel (lít/ca, lít/ngày, lít/kỳ).
  5. Xuất báo cáo kỹ thuật, file Master Excel 5 Sheet và MS Project XML.
"""

from __future__ import annotations

import os
import json
import datetime
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


@dataclass
class MachineAllocation:
    machine_code: str
    machine_name: str
    productivity_per_shift: float
    total_shifts_required: float
    daily_machines_needed: float
    fuel_rate_l_per_shift: float
    total_fuel_liters: float


@dataclass
class FleetTask:
    task_id: str
    code: str
    name: str
    unit: str
    quantity: float
    start_date: datetime.date
    end_date: datetime.date
    shifts_per_day: int = 2
    task_norm_key: Optional[str] = None
    custom_productivity: Optional[float] = None
    labor_per_shift: float = 4.0
    notes: str = ""
    # Calculated fields
    duration_days: int = 0
    daily_production: float = 0.0
    machine_allocations: List[MachineAllocation] = field(default_factory=list)

    def __post_init__(self):
        if self.duration_days == 0:
            self.duration_days = (self.end_date - self.start_date).days + 1
        if self.duration_days > 0 and self.quantity > 0:
            self.daily_production = round(self.quantity / self.duration_days, 2)


@dataclass
class DailyFleetMatrix:
    dates: List[datetime.date]
    daily_labor: Dict[datetime.date, float]
    daily_machines: Dict[str, Dict[datetime.date, float]]
    daily_fuel: Dict[datetime.date, float]
    peak_machines: Dict[str, float]
    total_fuel_liters: float
    total_shifts: float


class EquipmentFleetScheduler:
    """
    Công cụ tính toán ca xe, ca máy, cân đối thiết bị & nhiên liệu cho công trình xây dựng.
    """

    def __init__(
        self,
        norms_path: Optional[str] = None,
        fuel_norms_path: Optional[str] = None,
        material_factors_path: Optional[str] = None,
        experience_store: Optional[Any] = None,
    ):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        data_dir = os.path.join(base_dir, "data")

        self.norms_path = norms_path or os.path.join(data_dir, "equipment_norms_vincons.json")
        self.fuel_norms_path = fuel_norms_path or os.path.join(data_dir, "fuel_consumption_norms.json")
        self.material_factors_path = material_factors_path or os.path.join(data_dir, "material_conversion_factors.json")
        self.experience_store = experience_store

        self.norms_data = self._load_json(self.norms_path)
        self.fuel_data = self._load_json(self.fuel_norms_path)
        self.material_factors = self._load_json(self.material_factors_path)

    def set_experience_store(self, store: Any) -> None:
        """Thiết lập ProjectExperienceStore để tự động hiệu chuẩn định mức."""
        self.experience_store = store

    @staticmethod
    def _load_json(path: str) -> Dict[str, Any]:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def get_fuel_rate(self, machine_code: str) -> float:
        """Lấy định mức dầu lít/ca cho mã máy."""
        rates = self.fuel_data.get("equipment_rates", {})
        if machine_code in rates:
            return float(rates[machine_code].get("fuel_rate_l_per_shift", 40.0))
        return 40.0

    def compute_task_allocations(self, task: FleetTask) -> FleetTask:
        """
        Tính toán ca máy, số máy và nhiên liệu cho một tác vụ.
        """
        task_key = task.task_norm_key
        task_norms = self.norms_data.get("tasks", {}).get(task_key, {}) if task_key else {}

        machines_spec = task_norms.get("machines", {})
        task.machine_allocations = []

        if machines_spec:
            for m_k, m_info in machines_spec.items():
                m_code = m_info.get("code", "M1")
                m_name = m_info.get("name", m_k)
                prod_shift = float(m_info.get("productivity_per_shift", 100.0))
                if self.experience_store and hasattr(self.experience_store, "get_calibration_factor"):
                    alpha = self.experience_store.get_calibration_factor(task_key)
                    if alpha > 0 and alpha != 1.0:
                        prod_shift = round(prod_shift * alpha, 2)

                shifts_req = round(task.quantity / max(0.01, prod_shift), 2)
                daily_m = round(shifts_req / (task.duration_days * task.shifts_per_day), 2)
                fuel_rate = self.get_fuel_rate(m_code)
                tot_fuel = round(shifts_req * fuel_rate, 1)

                task.machine_allocations.append(
                    MachineAllocation(
                        machine_code=m_code,
                        machine_name=m_name,
                        productivity_per_shift=prod_shift,
                        total_shifts_required=shifts_req,
                        daily_machines_needed=daily_m,
                        fuel_rate_l_per_shift=fuel_rate,
                        total_fuel_liters=tot_fuel
                    )
                )
        elif task.custom_productivity:
            # Custom productivity
            shifts_req = round(task.quantity / task.custom_productivity, 2)
            daily_m = round(shifts_req / (task.duration_days * task.shifts_per_day), 2)
            fuel_rate = 50.0
            tot_fuel = round(shifts_req * fuel_rate, 1)
            task.machine_allocations.append(
                MachineAllocation(
                    machine_code="M_CUSTOM",
                    machine_name="Thiết bị chuyên dụng",
                    productivity_per_shift=task.custom_productivity,
                    total_shifts_required=shifts_req,
                    daily_machines_needed=daily_m,
                    fuel_rate_l_per_shift=fuel_rate,
                    total_fuel_liters=tot_fuel
                )
            )

        return task

    def build_daily_matrix(
        self,
        tasks: List[FleetTask],
        start_date: datetime.date,
        end_date: datetime.date
    ) -> DailyFleetMatrix:
        """
        Lập ma trận phân bổ nhân công, ca máy và nhiên liệu hàng ngày cho toàn dự án.
        """
        total_days = (end_date - start_date).days + 1
        dates = [start_date + datetime.timedelta(days=i) for i in range(total_days)]

        daily_labor: Dict[datetime.date, float] = {d: 0.0 for d in dates}
        daily_machines: Dict[str, Dict[datetime.date, float]] = {}
        daily_fuel: Dict[datetime.date, float] = {d: 0.0 for d in dates}

        # Initialize machine codes
        all_codes: set = set()
        for t in tasks:
            self.compute_task_allocations(t)
            for m in t.machine_allocations:
                all_codes.add(m.machine_code)

        for code in all_codes:
            daily_machines[code] = {d: 0.0 for d in dates}

        total_shifts = 0.0
        total_fuel = 0.0

        for t in tasks:
            for d in dates:
                if t.start_date <= d <= t.end_date:
                    daily_labor[d] += t.labor_per_shift * t.shifts_per_day

            for m in t.machine_allocations:
                total_shifts += m.total_shifts_required
                total_fuel += m.total_fuel_liters
                for d in dates:
                    if t.start_date <= d <= t.end_date:
                        daily_machines[m.machine_code][d] += m.daily_machines_needed
                        daily_fuel[d] += (
                            m.daily_machines_needed * m.fuel_rate_l_per_shift * t.shifts_per_day
                        )

        peak_machines = {
            code: max(daily_machines[code].values()) if daily_machines[code] else 0.0
            for code in all_codes
        }

        return DailyFleetMatrix(
            dates=dates,
            daily_labor=daily_labor,
            daily_machines=daily_machines,
            daily_fuel=daily_fuel,
            peak_machines=peak_machines,
            total_fuel_liters=round(total_fuel, 1),
            total_shifts=round(total_shifts, 1)
        )

    def export_excel(
        self,
        project_name: str,
        section_name: str,
        tasks: List[FleetTask],
        start_date: datetime.date,
        end_date: datetime.date,
        output_path: str,
        shifts_per_day: int = 2
    ) -> str:
        """
        Xuất file Excel Master 5 Sheet liên kết động chuẩn mẫu Vincons / B9.
        """
        if not HAS_OPENPYXL:
            raise RuntimeError("Cần cài đặt openpyxl để xuất file Excel: pip install openpyxl")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        font_family = "Arial"
        NAVY_HEADER = "1B365D"
        BLUE_SUB = "2E75B6"
        LIGHT_BLUE = "D9E1F2"
        LIGHT_GREEN = "E2EFDA"
        ACCENT_GREEN = "385723"
        AMBER_SUM = "FFF2CC"
        PEACH_FILL = "FCE4D6"
        GRAY_BORDER = "D9D9D9"

        thin_border = Border(
            left=Side(style='thin', color=GRAY_BORDER),
            right=Side(style='thin', color=GRAY_BORDER),
            top=Side(style='thin', color=GRAY_BORDER),
            bottom=Side(style='thin', color=GRAY_BORDER)
        )
        thick_bottom = Border(
            left=Side(style='thin', color=GRAY_BORDER),
            right=Side(style='thin', color=GRAY_BORDER),
            top=Side(style='thin', color=GRAY_BORDER),
            bottom=Side(style='medium', color=NAVY_HEADER)
        )

        matrix = self.build_daily_matrix(tasks, start_date, end_date)
        total_days = len(matrix.dates)
        weekday_vn = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

        # =====================================================================
        # SHEET 1: 01_TienDo_CaMay_Master
        # =====================================================================
        ws1 = wb.create_sheet(title="01_TienDo_CaMay_Master")
        ws1.views.sheetView[0].showGridLines = True

        ws1.merge_cells("A1:N1")
        ws1["A1"] = f"DỰ ÁN: {project_name.upper()}"
        ws1["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        ws1["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws1.merge_cells("A2:N2")
        ws1["A2"] = f"HẠNG MỤC: {section_name.upper()} — TIẾN ĐỘ THI CÔNG & PHÂN BỔ CA MÁY"
        ws1["A2"].font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
        ws1["A2"].fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
        ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws1.merge_cells("A3:N3")
        ws1["A3"] = f"TIẾN ĐỘ TỪ {start_date.strftime('%d/%m/%Y')} ĐẾN {end_date.strftime('%d/%m/%Y')} ({total_days} NGÀY) — {shifts_per_day} CA/NGÀY (20H/NGÀY)"
        ws1["A3"].font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws1["A3"].fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws1["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        headers = [
            ("STT", "A5:A6"), ("Mã WBS", "B5:B6"), ("Nội dung công việc", "C5:C6"),
            ("ĐVT", "D5:D6"), ("Khối lượng TK", "E5:E6"), ("Định mức (ĐVT/ca)", "F5:F6"),
            ("Tổng số ca", "G5:G6"), ("Năng xuất ngày", "H5:H6"), ("Số ngày", "I5:I6"),
            ("Ngày BĐ", "J5:J6"), ("Ngày KT", "K5:K6"), ("Số ca/ngày", "L5:L6"),
            ("Máy huy động", "M5:M6"), ("Thiết bị chính", "N5:N6"), ("NC/ngày", "O5:O6")
        ]
        for t_title, rng in headers:
            ws1.merge_cells(rng)
            top = rng.split(":")[0]
            ws1[top] = t_title
            ws1[top].font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
            ws1[top].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            ws1[top].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        date_col_start = 16
        for idx, d in enumerate(matrix.dates):
            c_idx = date_col_start + idx
            ws1.cell(row=5, column=c_idx, value=d.strftime("%d/%m"))
            ws1.cell(row=5, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
            ws1.cell(row=5, column=c_idx).fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
            ws1.cell(row=5, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")

            ws1.cell(row=6, column=c_idx, value=weekday_vn[d.weekday()])
            ws1.cell(row=6, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")
            if d.weekday() == 6:
                ws1.cell(row=6, column=c_idx).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                ws1.cell(row=6, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="C00000")
            else:
                ws1.cell(row=6, column=c_idx).fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
                ws1.cell(row=6, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="1B365D")

        # Task Rows
        row_curr = 7
        task_start_row = row_curr
        for idx, t in enumerate(tasks, 1):
            m_main = t.machine_allocations[0] if t.machine_allocations else None
            norm_val = m_main.productivity_per_shift if m_main else 100.0
            tot_shifts = m_main.total_shifts_required if m_main else (t.quantity / norm_val)
            m_daily = m_main.daily_machines_needed if m_main else (tot_shifts / (t.duration_days * shifts_per_day))
            m_name = m_main.machine_name if m_main else "Thiết bị phụ trợ"

            ws1.cell(row=row_curr, column=1, value=idx)
            ws1.cell(row=row_curr, column=2, value=t.code)
            ws1.cell(row=row_curr, column=3, value=t.name)
            ws1.cell(row=row_curr, column=4, value=t.unit)
            ws1.cell(row=row_curr, column=5, value=t.quantity)
            ws1.cell(row=row_curr, column=6, value=norm_val)
            ws1.cell(row=row_curr, column=7, value=f"=E{row_curr}/F{row_curr}")
            ws1.cell(row=row_curr, column=8, value=f"=E{row_curr}/I{row_curr}")
            ws1.cell(row=row_curr, column=9, value=t.duration_days)
            ws1.cell(row=row_curr, column=10, value=t.start_date.strftime("%Y-%m-%d"))
            ws1.cell(row=row_curr, column=11, value=t.end_date.strftime("%Y-%m-%d"))
            ws1.cell(row=row_curr, column=12, value=shifts_per_day)
            ws1.cell(row=row_curr, column=13, value=f"=G{row_curr}/(I{row_curr}*L{row_curr})")
            ws1.cell(row=row_curr, column=14, value=m_name)
            ws1.cell(row=row_curr, column=15, value=t.labor_per_shift * shifts_per_day)

            for c in range(1, date_col_start):
                cell = ws1.cell(row=row_curr, column=c)
                cell.font = Font(name=font_family, size=8.5)
                cell.border = thin_border

            # Timeline Gantt
            for d_i, d in enumerate(matrix.dates):
                c_idx = date_col_start + d_i
                cell = ws1.cell(row=row_curr, column=c_idx)
                cell.border = thin_border
                if t.start_date <= d <= t.end_date:
                    cell.value = round(m_daily, 2)
                    cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                    cell.font = Font(name=font_family, size=8, bold=True, color="1B365D")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.value = 0
                    cell.font = Font(name=font_family, size=7, color="BFBFBF")
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            row_curr += 1

        task_end_row = row_curr - 1
        row_curr += 1

        # Summary Row: Daily Labor
        row_labor = row_curr
        ws1.cell(row=row_labor, column=3, value="TỔNG NHÂN CÔNG HUY ĐỘNG (Người/ngày)")
        ws1.cell(row=row_labor, column=3).font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        for c in range(1, date_col_start):
            ws1.cell(row=row_labor, column=c).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")

        for d_i, d in enumerate(matrix.dates):
            c_idx = date_col_start + d_i
            c_let = get_column_letter(c_idx)
            cell = ws1.cell(row=row_labor, column=c_idx)
            cell.value = f"=SUM({c_let}{task_start_row}:{c_let}{task_end_row})"
            cell.font = Font(name=font_family, size=8.5, bold=True, color="002060")
            cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
            cell.border = thick_bottom
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row_curr += 2

        # Summary Row: Daily Fuel
        row_fuel = row_curr
        ws1.cell(row=row_fuel, column=3, value="TỔNG LƯỢNG DẦU DIEZEL TIÊU THỤ (Lít/ngày)")
        ws1.cell(row=row_fuel, column=3).font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        for c in range(1, date_col_start):
            ws1.cell(row=row_fuel, column=c).fill = PatternFill(start_color=ACCENT_GREEN, end_color=ACCENT_GREEN, fill_type="solid")

        for d_i, d in enumerate(matrix.dates):
            c_idx = date_col_start + d_i
            cell = ws1.cell(row=row_fuel, column=c_idx)
            cell.value = round(matrix.daily_fuel[d], 0)
            cell.font = Font(name=font_family, size=8, bold=True, color="C00000")
            cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
            cell.border = thick_bottom
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")

        # Column widths
        for c_let, w in {
            "A": 6, "B": 10, "C": 42, "D": 8, "E": 14, "F": 14, "G": 12, "H": 14,
            "I": 9, "J": 12, "K": 12, "L": 9, "M": 14, "N": 32, "O": 10
        }.items():
            ws1.column_dimensions[c_let].width = w

        for idx in range(total_days):
            c_let = get_column_letter(date_col_start + idx)
            ws1.column_dimensions[c_let].width = 6.8

        # Number formats
        for r in range(task_start_row, task_end_row + 1):
            ws1[f"E{r}"].number_format = "#,##0.0"
            ws1[f"F{r}"].number_format = "#,##0.0"
            ws1[f"G{r}"].number_format = "#,##0.0"
            ws1[f"H{r}"].number_format = "#,##0.0"
            ws1[f"M{r}"].number_format = "#,##0.00"
            ws1[f"O{r}"].number_format = "#,##0"

        # =====================================================================
        # SHEET 2: 02_TongHop_CaXe_CaMay_MMTB
        # =====================================================================
        ws2 = wb.create_sheet(title="02_TongHop_CaXe_CaMay_MMTB")
        ws2.views.sheetView[0].showGridLines = True

        ws2.merge_cells("A1:H1")
        ws2["A1"] = f"BẢNG TỔNG HỢP CA XE, CA MÁY & PHƯƠNG TIỆN HUY ĐỘNG — {project_name.upper()}"
        ws2["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        ws2["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws2_headers = ["STT", "Mã máy", "Tên phương tiện / Thiết bị", "ĐVT", "Định mức dầu (l/ca)", "Tổng ca máy (ca)", "Số máy Max", "Tổng dầu (Lít)"]
        for c_i, h in enumerate(ws2_headers, 1):
            cell = ws2.cell(row=3, column=c_i, value=h)
            cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Group machine stats
        mach_stats: Dict[str, Dict[str, Any]] = {}
        for t in tasks:
            for m in t.machine_allocations:
                if m.machine_code not in mach_stats:
                    mach_stats[m.machine_code] = {
                        "name": m.machine_name,
                        "oil_rate": m.fuel_rate_l_per_shift,
                        "shifts": 0.0,
                        "fuel": 0.0,
                        "max_m": matrix.peak_machines.get(m.machine_code, 0.0)
                    }
                mach_stats[m.machine_code]["shifts"] += m.total_shifts_required
                mach_stats[m.machine_code]["fuel"] += m.total_fuel_liters

        r2 = 4
        for idx, (code, data) in enumerate(mach_stats.items(), 1):
            ws2.cell(row=r2, column=1, value=idx)
            ws2.cell(row=r2, column=2, value=code)
            ws2.cell(row=r2, column=3, value=data["name"])
            ws2.cell(row=r2, column=4, value="ca")
            ws2.cell(row=r2, column=5, value=data["oil_rate"])
            ws2.cell(row=r2, column=6, value=round(data["shifts"], 1))
            ws2.cell(row=r2, column=7, value=round(data["max_m"], 1))
            ws2.cell(row=r2, column=8, value=f"=E{r2}*F{r2}")

            for c in range(1, 9):
                cell = ws2.cell(row=r2, column=c)
                cell.font = Font(name=font_family, size=9)
                cell.border = thin_border
                if c in [1, 2, 4]: cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c in [5, 6, 7, 8]:
                    cell.number_format = "#,##0.0" if c in [5, 6, 7] else "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else: cell.alignment = Alignment(horizontal="left", vertical="center")
            r2 += 1

        # Total
        ws2.cell(row=r2, column=3, value="TỔNG CỘNG TOÀN DỰ ÁN")
        ws2.cell(row=r2, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws2.cell(row=r2, column=6, value=f"=SUM(F4:F{r2-1})")
        ws2.cell(row=r2, column=7, value=f"=SUM(G4:G{r2-1})")
        ws2.cell(row=r2, column=8, value=f"=SUM(H4:H{r2-1})")
        for c in [6, 7, 8]:
            cell = ws2.cell(row=r2, column=c)
            cell.font = Font(name=font_family, size=10, bold=True, color="C00000")
            cell.number_format = "#,##0.0" if c in [6, 7] else "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")
        for c in range(1, 9):
            ws2.cell(row=r2, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
            ws2.cell(row=r2, column=c).border = thick_bottom

        for c_let, w in {"A": 6, "B": 10, "C": 44, "D": 8, "E": 18, "F": 18, "G": 16, "H": 22}.items():
            ws2.column_dimensions[c_let].width = w

        # =====================================================================
        # SHEET 3: 03_KeHoach_Dau_Diezel (THEO 4 KỲ CHIẾN LƯỢC)
        # =====================================================================
        ws3 = wb.create_sheet(title="03_KeHoach_Dau_Diezel")
        ws3.views.sheetView[0].showGridLines = True

        ws3.merge_cells("A1:I1")
        ws3["A1"] = f"KẾ HOẠCH CẤP DẦU DIEZEL CHO MÁY MÓC THI CÔNG — {project_name.upper()} (4 KỲ CHIẾN LƯỢC)"
        ws3["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        ws3["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        headers_ws3 = [
            "STT", "Chủng loại phương tiện / Thiết bị", "Định mức (lít/ca)", "Tổng số ca máy",
            "Tổng nhu cầu dầu (Lít)", "Kỳ 1: Khởi công & Chuẩn bị",
            "Kỳ 2: Kết cấu ngầm & Móng", "Kỳ 3: Kết cấu thân & Thượng tầng", "Kỳ 4: Hoàn thiện & Bàn giao"
        ]
        for c_i, h in enumerate(headers_ws3, 1):
            cell = ws3.cell(row=3, column=c_i, value=h)
            cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        r_op = 4
        for idx, (code, data) in enumerate(mach_stats.items(), 1):
            ws3.cell(row=r_op, column=1, value=idx)
            ws3.cell(row=r_op, column=2, value=data["name"])
            ws3.cell(row=r_op, column=3, value=data["oil_rate"])
            ws3.cell(row=r_op, column=4, value=round(data["shifts"], 1))
            ws3.cell(row=r_op, column=5, value=f"=C{r_op}*D{r_op}")
            ws3.cell(row=r_op, column=6, value=f"=E{r_op}*0.30")
            ws3.cell(row=r_op, column=7, value=f"=E{r_op}*0.35")
            ws3.cell(row=r_op, column=8, value=f"=E{r_op}*0.25")
            ws3.cell(row=r_op, column=9, value=f"=E{r_op}*0.10")

            for c in range(1, 10):
                cell = ws3.cell(row=r_op, column=c)
                cell.font = Font(name=font_family, size=9)
                cell.border = thin_border
                if c == 1:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c in [3, 4, 5, 6, 7, 8, 9]:
                    cell.number_format = "#,##0.0" if c in [3, 4] else "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r_op += 1

        # Total oil row
        ws3.cell(row=r_op, column=2, value="TỔNG SỐ LÍT DẦU DIEZEL CẦN CUNG CẤP (LÍT)")
        ws3.cell(row=r_op, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
        for c in [4, 5, 6, 7, 8, 9]:
            c_let = get_column_letter(c)
            ws3.cell(row=r_op, column=c, value=f"=SUM({c_let}4:{c_let}{r_op-1})")
            ws3.cell(row=r_op, column=c).font = Font(name=font_family, size=10, bold=True, color="C00000")
            ws3.cell(row=r_op, column=c).number_format = "#,##0.0" if c == 4 else "#,##0"
            ws3.cell(row=r_op, column=c).alignment = Alignment(horizontal="right", vertical="center")
        for c in range(1, 10):
            ws3.cell(row=r_op, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
            ws3.cell(row=r_op, column=c).border = thick_bottom

        for c_let, w in {"A": 6, "B": 44, "C": 16, "D": 16, "E": 20, "F": 22, "G": 22, "H": 22, "I": 24}.items():
            ws3.column_dimensions[c_let].width = w

        # =====================================================================
        # SHEET 4: 04_KeHoach_NhanLuc (PHÂN BỔ TỔ ĐỘI CHUYÊN NGHIỆP)
        # =====================================================================
        ws4 = wb.create_sheet(title="04_KeHoach_NhanLuc")
        ws4.views.sheetView[0].showGridLines = True

        ws4.merge_cells("A1:H1")
        ws4["A1"] = f"BẢNG PHÂN BỔ NHÂN LỰC THI CÔNG — {project_name.upper()} (CÁC TỔ ĐỘI CHUYÊN NGHIỆP)"
        ws4["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        ws4["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        headers_ws4 = [
            "STT", "Tổ đội / Bộ phận chức năng", "Nhân lực bình quân (người)",
            "Huy động cao điểm (người)", "Chế độ ca kíp", "Nhiệm vụ chính trên công trường",
            "Đội trưởng phụ trách", "Ghi chú an toàn & kỹ thuật"
        ]
        for c_i, h in enumerate(headers_ws4, 1):
            cell = ws4.cell(row=3, column=c_i, value=h)
            cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        labor_rows = [
            (1, "Tổ Cơ giới & Vận hành máy thi công", 12, 18, "2 ca/ngày", "Vận hành máy đào, cần cẩu, máy khoan cọc", "Nguyễn Văn Hùng", "Tuân thủ ATLĐ ca đêm"),
            (2, "Tổ Cốt thép & Gia công bãi tiền chế", 14, 20, "1-2 ca/ngày", "Cắt uốn thép 11.7m theo sơ đồ RebarCut, hàn lồng", "Trần Bá Thắng", "Kiểm tra mối hàn TCVN 5574"),
            (3, "Tổ Ván khuôn & Đà giáo định hình", 12, 16, "1 ca/ngày", "Lắp dựng & tháo dỡ ván khuôn tấm lớn", "Lê Đình Long", "Kiểm tra độ võng & chuyển vị"),
            (4, "Tổ Bê tông & Bảo dưỡng ẩm", 10, 15, "Theo đợt đổ", "Đổ, san gạt, đầm dùi & phủ bạt bảo dưỡng", "Phạm Quốc Tuấn", "Đo độ sụt & đúc mẫu thí nghiệm"),
            (5, "Tổ Kỹ thuật, QA/QC, Trắc đạc & KCS", 6, 8, "Thường trực", "Định vị tim mốc, nghiệm thu Hold Points, lập BBNT", "Kỹ sư Trưởng Hiện trường", "Nghiệm thu chuyển bước theo NĐ 207"),
        ]
        r_l = 4
        for row_data in labor_rows:
            ws4.cell(row=r_l, column=1, value=row_data[0])
            ws4.cell(row=r_l, column=2, value=row_data[1])
            ws4.cell(row=r_l, column=3, value=row_data[2])
            ws4.cell(row=r_l, column=4, value=row_data[3])
            ws4.cell(row=r_l, column=5, value=row_data[4])
            ws4.cell(row=r_l, column=6, value=row_data[5])
            ws4.cell(row=r_l, column=7, value=row_data[6])
            ws4.cell(row=r_l, column=8, value=row_data[7])

            for c in range(1, 9):
                cell = ws4.cell(row=r_l, column=c)
                cell.font = Font(name=font_family, size=9)
                cell.border = thin_border
                if c == 1:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c in [3, 4]:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c == 5:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r_l += 1

        # Total labor row
        ws4.cell(row=r_l, column=2, value="TỔNG CỘNG NHÂN LỰC THI CÔNG TOÀN DỰ ÁN")
        ws4.cell(row=r_l, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws4.cell(row=r_l, column=3, value=f"=SUM(C4:C{r_l-1})")
        ws4.cell(row=r_l, column=4, value=f"=SUM(D4:D{r_l-1})")
        for c in [3, 4]:
            ws4.cell(row=r_l, column=c).font = Font(name=font_family, size=10, bold=True, color="C00000")
            ws4.cell(row=r_l, column=c).number_format = "#,##0"
            ws4.cell(row=r_l, column=c).alignment = Alignment(horizontal="right", vertical="center")
        for c in range(1, 9):
            ws4.cell(row=r_l, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
            ws4.cell(row=r_l, column=c).border = thick_bottom

        for c_let, w in {"A": 6, "B": 38, "C": 24, "D": 24, "E": 18, "F": 42, "G": 24, "H": 30}.items():
            ws4.column_dimensions[c_let].width = w

        # =====================================================================
        # SHEET 5: 05_DoiChieu_BocTach (ĐỐI CHIẾU KHỐI LƯỢNG THIẾT KẾ)
        # =====================================================================
        ws5 = wb.create_sheet(title="05_DoiChieu_BocTach")
        ws5.views.sheetView[0].showGridLines = True

        ws5.merge_cells("A1:H1")
        ws5["A1"] = f"BẢNG ĐỐI CHIẾU KHỐI LƯỢNG THỰC TẾ HỒ SƠ BÓC TÁCH — {project_name.upper()}"
        ws5["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
        ws5["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        headers_ws5 = [
            "STT", "Mã WBS", "Danh mục công tác thi công", "ĐVT",
            "Khối lượng thiết kế", "Tổng số ca máy yêu cầu", "Số ca/ngày", "MMTB chính áp dụng"
        ]
        for c_i, h in enumerate(headers_ws5, 1):
            cell = ws5.cell(row=3, column=c_i, value=h)
            cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        r_boq = 4
        for idx, t in enumerate(tasks, 1):
            m_main = t.machine_allocations[0].machine_name if t.machine_allocations else "Thủ công / Cơ giới nhỏ"
            shifts_tot = sum(m.total_shifts_required for m in t.machine_allocations) if t.machine_allocations else 0.0

            ws5.cell(row=r_boq, column=1, value=idx)
            ws5.cell(row=r_boq, column=2, value=t.code)
            ws5.cell(row=r_boq, column=3, value=t.name)
            ws5.cell(row=r_boq, column=4, value=t.unit)
            ws5.cell(row=r_boq, column=5, value=t.quantity)
            ws5.cell(row=r_boq, column=6, value=round(shifts_tot, 1))
            ws5.cell(row=r_boq, column=7, value=t.shifts_per_day)
            ws5.cell(row=r_boq, column=8, value=m_main)

            for c in range(1, 9):
                cell = ws5.cell(row=r_boq, column=c)
                cell.font = Font(name=font_family, size=9)
                cell.border = thin_border
                if c in [1, 2, 4, 7]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c in [5, 6]:
                    cell.number_format = "#,##0.0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r_boq += 1

        # Total BOQ row
        ws5.cell(row=r_boq, column=3, value="TỔNG CỘNG SỐ CA MÁY YÊU CẦU TOÀN CÔNG TRÌNH")
        ws5.cell(row=r_boq, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws5.cell(row=r_boq, column=6, value=f"=SUM(F4:F{r_boq-1})")
        ws5.cell(row=r_boq, column=6).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws5.cell(row=r_boq, column=6).number_format = "#,##0.0"
        ws5.cell(row=r_boq, column=6).alignment = Alignment(horizontal="right", vertical="center")
        for c in range(1, 9):
            ws5.cell(row=r_boq, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
            ws5.cell(row=r_boq, column=c).border = thick_bottom

        for c_let, w in {"A": 6, "B": 12, "C": 44, "D": 8, "E": 18, "F": 22, "G": 12, "H": 36}.items():
            ws5.column_dimensions[c_let].width = w

        # Save workbook
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        wb.save(output_path)
        return output_path

    def export_ms_project_xml(
        self,
        project_name: str,
        tasks: List[FleetTask],
        start_date: datetime.date,
        end_date: datetime.date,
        output_path: str
    ) -> str:
        """
        Xuất file XML chuẩn Microsoft Project.
        """
        project = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
        ET.SubElement(project, "Name").text = project_name
        ET.SubElement(project, "Title").text = f"Tiến độ thi công & Ca máy — {project_name}"
        ET.SubElement(project, "StartDate").text = f"{start_date.isoformat()}T07:00:00"
        ET.SubElement(project, "FinishDate").text = f"{end_date.isoformat()}T17:00:00"

        tasks_elem = ET.SubElement(project, "Tasks")

        # Root task
        root_t = ET.SubElement(tasks_elem, "Task")
        ET.SubElement(root_t, "UID").text = "1"
        ET.SubElement(root_t, "ID").text = "1"
        ET.SubElement(root_t, "Name").text = project_name
        ET.SubElement(root_t, "OutlineLevel").text = "0"
        ET.SubElement(root_t, "Start").text = f"{start_date.isoformat()}T07:00:00"
        ET.SubElement(root_t, "Finish").text = f"{end_date.isoformat()}T17:00:00"
        dur_h = len(tasks) * 8
        ET.SubElement(root_t, "Duration").text = f"PT{dur_h}H0M0S"

        for idx, t in enumerate(tasks, 2):
            task_node = ET.SubElement(tasks_elem, "Task")
            ET.SubElement(task_node, "UID").text = str(idx)
            ET.SubElement(task_node, "ID").text = str(idx)
            ET.SubElement(task_node, "Name").text = f"{t.code} — {t.name}"
            ET.SubElement(task_node, "OutlineLevel").text = "1"
            ET.SubElement(task_node, "Start").text = f"{t.start_date.isoformat()}T07:00:00"
            ET.SubElement(task_node, "Finish").text = f"{t.end_date.isoformat()}T17:00:00"
            ET.SubElement(task_node, "Duration").text = f"PT{t.duration_days * 8}H0M0S"

        tree = ET.ElementTree(project)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        tree.write(output_path, encoding="utf-8", xml_declaration=True)
        return output_path
