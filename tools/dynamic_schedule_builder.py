# -*- coding: utf-8 -*-
"""
DYNAMIC SCHEDULE BUILDER — BỘ KHỞI TẠO TIẾN ĐỘ THI CÔNG & CA MÁY 100% CÔNG THỨC ĐỘNG
Dành cho Hệ thống Multi-Agent AEC (23HG-multiagent-system).
Pure Python, Zero LLM — Tính toán kỹ thuật và sinh file Excel Master chuẩn CPM.

Tính năng:
1. 01_THONG_SO_DU_AN: Tham số điều hành tập trung (ngày bắt đầu, deadline động, công tắc Gantt).
2. 02_DINH_MUC_CA_MAY_VA_DAU: Bảng tra định mức năng suất Vincons & định mức dầu Diesel.
3. 03_TIEN_DO_GANTT_CPM: Tiến độ thi công liên kết đường găng CPM, ma trận Gantt động kèm
   Biểu đồ đường cong phụ tải máy móc & nhân công (LineChart 97 ngày).
4. 04_TONG_HOP_CA_MAY_VA_DAU: Tổng hợp ca xe, ca máy (SUMIF/VLOOKUP), phân bổ dầu theo 4 tháng
   kèm Biểu đồ cột tiêu thụ nhiên liệu (BarChart).
5. 05_NHU_CAU_VAT_TU_CHINH: Bảng nhu cầu vật tư chính liên kết khối lượng đào đắp từ Sheet 03
   kèm Biểu đồ cơ cấu vật tư san lấp (BarChart).
6. 06_SO_SANH_DINH_MUC_VS_THUC_TE: Phục dựng & nâng cấp nguyên mẫu Sheet SS của Vina Alpha,
   so sánh đối chứng Định mức vs Đề xuất BĐH, phân tích nguyên nhân hiện trường kèm
   Biểu đồ cột cụm đối sánh thiết bị (Clustered BarChart).
"""

from __future__ import annotations
import datetime
import os
from typing import Any, Dict, List, Optional

try:
    import openpyxl
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


class DynamicScheduleBuilder:
    """
    Bộ khởi tạo và xuất bản Tiến độ thi công tổng thể Cụm B9 và các dự án Hạ tầng kỹ thuật
    với 100% công thức sống và Native Excel Charts.
    """

    NAVY_HEADER = "1B365D"
    BLUE_SUB = "2E75B6"
    AMBER_ALERT = "FFF2CC"
    GREEN_OK = "E2EFDA"
    LIGHT_GREEN = "D9E1F2"
    GRAY_BORDER = "D9D9D9"

    def __init__(self, font_family: str = "Times New Roman"):
        self.font_family = font_family
        self.thin_border = Border(
            left=Side(style="thin", color=self.GRAY_BORDER),
            right=Side(style="thin", color=self.GRAY_BORDER),
            top=Side(style="thin", color=self.GRAY_BORDER),
            bottom=Side(style="thin", color=self.GRAY_BORDER),
        )
        self.thick_bottom_border = Border(
            left=Side(style="thin", color=self.GRAY_BORDER),
            right=Side(style="thin", color=self.GRAY_BORDER),
            top=Side(style="thin", color=self.GRAY_BORDER),
            bottom=Side(style="medium", color=self.NAVY_HEADER),
        )
        self.double_bottom_border = Border(
            left=Side(style="thin", color=self.GRAY_BORDER),
            right=Side(style="thin", color=self.GRAY_BORDER),
            top=Side(style="thin", color=self.GRAY_BORDER),
            bottom=Side(style="double", color="000000"),
        )

    def build_cum_b9_workbook(
        self,
        output_file: str,
        start_date: Optional[datetime.date] = None,
        total_days: int = 97,
    ) -> str:
        """
        Khởi tạo và xuất file Master Excel Tiến độ Cụm B9 hoàn chỉnh 6 sheet kèm 4 biểu đồ native.
        """
        if not HAS_OPENPYXL:
            raise RuntimeError("Cần cài đặt thư viện 'openpyxl' để xuất file Excel tiến độ.")

        if start_date is None:
            start_date = datetime.date(2026, 8, 11)

        os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Xóa sheet mặc định

        # ---------------------------------------------------------------------
        # SHEET 1: 01_THONG_SO_DU_AN
        # ---------------------------------------------------------------------
        ws1 = wb.create_sheet(title="01_THONG_SO_DU_AN")
        ws1.views.sheetView[0].showGridLines = True

        ws1["B2"] = "DỰ ÁN: SÂN VẬN ĐỘNG OLYMPIC - THƯỜNG TÍN (HÀ NỘI)"
        ws1["B2"].font = Font(name=self.font_family, size=14, bold=True, color=self.NAVY_HEADER)
        ws1["B3"] = "HẠNG MỤC: HẠ TẦNG KỸ THUẬT THI CÔNG SAN LẤP CỤM B9 VÀ ĐƯỜNG NỘI BỘ"
        ws1["B3"].font = Font(name=self.font_family, size=12, bold=True, color=self.BLUE_SUB)
        ws1["B4"] = "BẢNG THIẾT LẬP THÔNG SỐ ĐIỀU HÀNH TIẾN ĐỘ & CA MÁY (100% CÔNG THỨC SỐNG)"
        ws1["B4"].font = Font(name=self.font_family, size=11, italic=True)

        ws1.cell(row=5, column=2, value="Thông số quản trị dự án").font = Font(name=self.font_family, size=11, bold=True, color="FFFFFF")
        ws1.cell(row=5, column=2).fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
        ws1.cell(row=5, column=3, value="Giá trị thiết lập").font = Font(name=self.font_family, size=11, bold=True, color="FFFFFF")
        ws1.cell(row=5, column=3).fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
        ws1.cell(row=5, column=4, value="Ghi chú kỹ thuật & Cơ chế liên kết động").font = Font(name=self.font_family, size=11, bold=True, color="FFFFFF")
        ws1.cell(row=5, column=4).fill = PatternFill("solid", fgColor=self.NAVY_HEADER)

        ws1.cell(row=6, column=2, value="Ngày phát lệnh khởi công (Bắt đầu)").border = self.thin_border
        c_s = ws1.cell(row=6, column=3, value=start_date)
        c_s.border = self.thin_border
        c_s.alignment = Alignment(horizontal="center")
        c_s.number_format = "DD/MM/YYYY"
        c_s.font = Font(name=self.font_family, size=10, bold=True, color="002060")
        ws1.cell(row=6, column=4, value="Input gốc: Đổi ngày này, toàn bộ 14 công tác và 97 cột Gantt tự động nhảy theo!").border = self.thin_border

        ws1.cell(row=7, column=2, value="Ngày hoàn thành mục tiêu (Deadline)").border = self.thin_border
        c_e = ws1.cell(row=7, column=3, value="=MAX('03_TIEN_DO_GANTT_CPM'!$K$6:$K$19)")
        c_e.border = self.thin_border
        c_e.alignment = Alignment(horizontal="center")
        c_e.number_format = "DD/MM/YYYY"
        c_e.font = Font(name=self.font_family, size=10, bold=True, color="C00000")
        ws1.cell(row=7, column=4, value="Công thức: MAX('03_TIEN_DO_GANTT_CPM'!$K$6:$K$19) — Tự động cập nhật theo công tác kết thúc muộn nhất").border = self.thin_border

        ws1.cell(row=8, column=2, value="Tổng thời gian thi công (Ngày)").border = self.thin_border
        c_tot = ws1.cell(row=8, column=3, value="=C7-C6+1")
        c_tot.border = self.thin_border
        c_tot.alignment = Alignment(horizontal="center")
        c_tot.number_format = "#,##0"
        c_tot.font = Font(name=self.font_family, size=10, bold=True)
        ws1.cell(row=8, column=4, value="Công thức: C7-C6+1 — Tự động tính số ngày lịch liên tục").border = self.thin_border

        ws1.cell(row=9, column=2, value="Chế độ làm việc công trường (Ca/ngày)").border = self.thin_border
        c_shift = ws1.cell(row=9, column=3, value=2)
        c_shift.border = self.thin_border
        c_shift.alignment = Alignment(horizontal="center")
        c_shift.font = Font(name=self.font_family, size=10, bold=True)
        ws1.cell(row=9, column=4, value="2 ca/ngày (10 giờ/ca = 20 giờ làm việc/ngày)").border = self.thin_border

        ws1.cell(row=10, column=2, value="Hệ số ca máy (HSTCA)").border = self.thin_border
        c_hs = ws1.cell(row=10, column=3, value="=C9")
        c_hs.border = self.thin_border
        c_hs.alignment = Alignment(horizontal="center")
        c_hs.font = Font(name=self.font_family, size=10, bold=True)
        ws1.cell(row=10, column=4, value="Công thức: C9 — Tự động lấy theo chế độ ca làm việc").border = self.thin_border

        ws1.cell(row=11, column=2, value="Công tắc hiển thị biểu đồ Gantt").border = self.thin_border
        c_sw = ws1.cell(row=11, column=3, value=1)
        c_sw.border = self.thin_border
        c_sw.alignment = Alignment(horizontal="center")
        c_sw.font = Font(name=self.font_family, size=11, bold=True, color="C00000")
        c_sw.fill = PatternFill("solid", fgColor=self.AMBER_ALERT)
        ws1.cell(row=11, column=4, value="Nhập 1 = Thanh Gantt (█) | Nhập 2 = Nhân công (người) | Nhập 3 = Ca máy (máy)").border = self.thin_border

        ws1.cell(row=12, column=2, value="Tiêu chuẩn kỹ thuật áp dụng").border = self.thin_border
        ws1.cell(row=12, column=3, value="TCVN 9436:2012 / TCVN 8819:2011").alignment = Alignment(horizontal="center")
        ws1.cell(row=12, column=3).border = self.thin_border
        ws1.cell(row=12, column=4, value="Thi công & nghiệm thu nền đường, móng mặt đường").border = self.thin_border

        ws1.cell(row=13, column=2, value="Định mức thi công tham chiếu").border = self.thin_border
        ws1.cell(row=13, column=3, value="Định mức Vincons & TT 38/2026/TT-BXD").alignment = Alignment(horizontal="center")
        ws1.cell(row=13, column=3).border = self.thin_border
        ws1.cell(row=13, column=4, value="Định mức dự toán & ca máy xây dựng công trình").border = self.thin_border

        ws1.column_dimensions["B"].width = 40
        ws1.column_dimensions["C"].width = 28
        ws1.column_dimensions["D"].width = 85

        # ---------------------------------------------------------------------
        # SHEET 2: 02_DINH_MUC_CA_MAY_VA_DAU
        # ---------------------------------------------------------------------
        ws2 = wb.create_sheet(title="02_DINH_MUC_CA_MAY_VA_DAU")
        ws2.views.sheetView[0].showGridLines = True

        ws2["B2"] = "HỆ THỐNG ĐỊNH MỨC NĂNG SUẤT CA MÁY & TIÊU HAO NHIÊN LIỆU (VINCONS_ĐMGK_02-01)"
        ws2["B2"].font = Font(name=self.font_family, size=13, bold=True, color=self.NAVY_HEADER)
        ws2["B3"] = "Bảng dữ liệu gốc phục vụ tra cứu tự động qua hàm VLOOKUP cho toàn bộ hệ thống"
        ws2["B3"].font = Font(name=self.font_family, size=10, italic=True)

        headers_dm = [
            "STT", "Mã máy", "Tên phương tiện / Thiết bị", "ĐVT năng suất",
            "Năng suất Vincons (ĐVT/ca)", "Định mức dầu Diesel (Lít/ca)", "Số nhân công đi kèm (Người/ca)"
        ]
        for c_i, h in enumerate(headers_dm, start=2):
            cell = ws2.cell(row=5, column=c_i, value=h)
            cell.font = Font(name=self.font_family, size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        machines_data = [
            (1, "M1", "Máy xúc bánh xích PC200 - PC300", "m3", 401.07, 89.6, 2),
            (2, "M3", "Máy ủi D3 - D5", "m3", 350.00, 124.0, 2),
            (3, "M4", "Máy ủi công suất lớn CS 230CV", "m3", 550.00, 155.0, 2),
            (4, "M5", "Máy lu rung 12T - 16T", "m3", 380.00, 95.0, 2),
            (5, "M6", "Máy san tự hành 110CV", "m3", 420.00, 110.0, 2),
            (6, "M7", "Ô tô tự đổ 18 tấn", "m3", 160.00, 78.0, 1),
            (7, "M8", "Máy xúc bánh lốp PC140", "m/ca", 120.00, 65.0, 3),
            (8, "M9", "Máy rải bê tông nhựa chuyên dụng", "m2", 1500.00, 140.0, 6),
            (9, "M10", "Máy lu lốp tĩnh 16 tấn", "m2", 1800.00, 85.0, 2),
            (10, "M11", "Máy tưới nhũ tương đường bộ", "m2", 3000.00, 45.0, 2),
            (11, "MP1", "Xe téc & Xe cấp dầu hiện trường", "xe/ca", 1.00, 50.0, 2),
        ]

        for idx, m in enumerate(machines_data, start=6):
            for c_i, val in enumerate(m, start=2):
                cell = ws2.cell(row=idx, column=c_i, value=val)
                cell.border = self.thin_border
                cell.font = Font(name=self.font_family, size=9.5)
                if c_i in (2, 3, 5):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i in (6, 7):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0.0"
                elif c_i == 8:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.number_format = "#,##0"

        ws2.column_dimensions["B"].width = 6
        ws2.column_dimensions["C"].width = 12
        ws2.column_dimensions["D"].width = 38
        ws2.column_dimensions["E"].width = 16
        ws2.column_dimensions["F"].width = 24
        ws2.column_dimensions["G"].width = 24
        ws2.column_dimensions["H"].width = 24

        # ---------------------------------------------------------------------
        # SHEET 3: 03_TIEN_DO_GANTT_CPM (KÈM BIỂU ĐỒ NATIVE EXCEL ĐƯỜNG CONG PHỤ TẢI)
        # ---------------------------------------------------------------------
        ws3 = wb.create_sheet(title="03_TIEN_DO_GANTT_CPM")
        ws3.views.sheetView[0].showGridLines = True

        ws3["B2"] = "TIẾN ĐỘ THI CÔNG CHI TIẾT & ĐIỀU PHỐI CA MÁY THEO ĐƯỜNG GĂNG (CPM GANTT)"
        ws3["B2"].font = Font(name=self.font_family, size=13, bold=True, color=self.NAVY_HEADER)
        ws3["B3"] = "100% Công thức động: Thay đổi thông số ở Sheet 01, toàn bộ bảng tiến độ tự động tái cân bằng"
        ws3["B3"].font = Font(name=self.font_family, size=10, italic=True)

        headers_cpm = [
            "STT", "Mã WBS", "Danh mục công tác thi công Cụm B9", "ĐVT",
            "Khối lượng", "Mã thiết bị", "Định mức NS (=VLOOKUP)", "Tổng số ca máy (=Khối lượng/NS)",
            "Thời gian thi công (Ngày)", "Ngày bắt đầu (ES)", "Ngày kết thúc (EF)",
            "Đường găng (CPM)", "Số máy/ngày (=Ca/(Ngày*HSTCA))", "Số nhân công/ngày (=Máy*NC/ca)"
        ]
        for c_i, h in enumerate(headers_cpm, start=1):
            cell = ws3.cell(row=5, column=c_i, value=h)
            cell.font = Font(name=self.font_family, size=9, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        date_col_start = 15  # Cột O
        for d_i in range(total_days):
            c_idx = date_col_start + d_i
            c_let = get_column_letter(c_idx)

            cell_h1 = ws3.cell(row=3, column=c_idx, value=f"=MONTH({c_let}$4)")
            cell_h1.alignment = Alignment(horizontal="center", vertical="center")
            cell_h1.font = Font(name=self.font_family, size=7.5, bold=True, color="FFFFFF")
            cell_h1.fill = PatternFill("solid", fgColor=self.BLUE_SUB)

            cell_date = ws3.cell(row=4, column=c_idx)
            if d_i == 0:
                cell_date.value = "='01_THONG_SO_DU_AN'!$C$6"
            else:
                prev_let = get_column_letter(c_idx - 1)
                cell_date.value = f"={prev_let}$4+1"
            cell_date.alignment = Alignment(horizontal="center", vertical="center")
            cell_date.font = Font(name=self.font_family, size=8, bold=True, color="FFFFFF")
            cell_date.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
            cell_date.number_format = "DD/MM"

            cell_d_idx = ws3.cell(row=5, column=c_idx, value=d_i + 1)
            cell_d_idx.alignment = Alignment(horizontal="center", vertical="center")
            cell_d_idx.font = Font(name=self.font_family, size=8, bold=True, color="FFFFFF")
            cell_d_idx.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)

        tasks_cpm = [
            (1, "1.0", "Huy động máy móc thiết bị & Lán trại phụ trợ", "gói", 1.0, "MP1", 7, "='01_THONG_SO_DU_AN'!$C$6"),
            (2, "1.1", "Đào bóc hữu cơ lòng đường & gom tập kết", "m3", 102793.8, "M1", 74, "=J6"),
            (3, "1.2", "Vận chuyển đất đào hữu cơ nội bộ < 5km", "m3", 102793.8, "M7", 72, "=J7"),
            (4, "1.3", "Thi công đắp đất / cát K90 nền đường", "m3", 360680.0, "M3", 74, "=J7+2"),
            (5, "1.4", "Thi công đắp cát K95 lớp 1", "m3", 90170.0, "M3", 74, "=J9+4"),
            (6, "1.5", "Thi công đắp cát K98 lớp 1", "m3", 54102.0, "M4", 74, "=J10+4"),
            (7, "1.6", "Thi công Hệ thống thoát nước mưa (Cống D1500)", "m", 21640.8, "M8", 72, "=J11+4"),
            (8, "1.7", "Thi công Hệ thống thoát nước thải (Ống PE D300/400)", "m", 21640.8, "M8", 68, "=J12+4"),
            (9, "1.8", "Thi công lớp móng Base B (CPĐD loại 2)", "m3", 27051.0, "M6", 68, "=J13+4"),
            (10, "1.9", "Thi công lớp móng Base A (CPĐD loại 1)", "m3", 21640.8, "M6", 66, "=J14+6"),
            (11, "1.10", "Thi công thảm bê tông nhựa hạt trung BTN C19", "m2", 180340.0, "M5", 54, "=J15+15"),
            (12, "1.11", "Thi công lắp đặt bó vỉa hè", "m", 144272.0, "M8", 52, "=J16+2"),
            (13, "1.12", "Thi công đắp đất vỉa hè & dải phân cách", "m3", 324612.0, "M3", 48, "=J17+4"),
            (14, "2.0", "San lấp mặt bằng tổng thể cụm B9", "m3", 1774000.0, "M4", 67, "=J6+2"),
        ]

        r_start = 6
        for idx, t in enumerate(tasks_cpm, start=r_start):
            stt, wbs, name, unit, qty, m_code, dur, formula_start = t

            ws3.cell(row=idx, column=1, value=stt).alignment = Alignment(horizontal="center")
            ws3.cell(row=idx, column=2, value=wbs).alignment = Alignment(horizontal="center")
            ws3.cell(row=idx, column=3, value=name)
            ws3.cell(row=idx, column=4, value=unit).alignment = Alignment(horizontal="center")
            ws3.cell(row=idx, column=5, value=qty).number_format = "#,##0.0"
            ws3.cell(row=idx, column=6, value=m_code).alignment = Alignment(horizontal="center")

            ws3.cell(row=idx, column=7, value=f"=VLOOKUP(F{idx}, '02_DINH_MUC_CA_MAY_VA_DAU'!$C$6:$I$16, 5, FALSE)").number_format = "#,##0.0"
            ws3.cell(row=idx, column=8, value=f"=E{idx}/G{idx}").number_format = "#,##0.0"
            ws3.cell(row=idx, column=9, value=dur).number_format = "#,##0"

            c_j = ws3.cell(row=idx, column=10, value=formula_start)
            c_j.alignment = Alignment(horizontal="center")
            c_j.number_format = "DD/MM/YYYY"

            c_k = ws3.cell(row=idx, column=11, value=f"=J{idx}+I{idx}-1")
            c_k.alignment = Alignment(horizontal="center")
            c_k.number_format = "DD/MM/YYYY"

            c_l = ws3.cell(row=idx, column=12, value=f"=IF(K{idx}='01_THONG_SO_DU_AN'!$C$7, \"CRITICAL (GĂNG)\", \"\")")
            c_l.alignment = Alignment(horizontal="center")
            c_l.font = Font(name=self.font_family, size=8.5, bold=True, color="C00000")

            c_m = ws3.cell(row=idx, column=13, value=f"=ROUNDUP(H{idx}/(I{idx}*'01_THONG_SO_DU_AN'!$C$10), 0)")
            c_m.alignment = Alignment(horizontal="center")
            c_m.number_format = "#,##0"

            c_n = ws3.cell(row=idx, column=14, value=f"=M{idx}*VLOOKUP(F{idx}, '02_DINH_MUC_CA_MAY_VA_DAU'!$C$6:$I$16, 7, FALSE)")
            c_n.alignment = Alignment(horizontal="center")
            c_n.number_format = "#,##0"

            for c in range(1, 15):
                ws3.cell(row=idx, column=c).border = self.thin_border
                ws3.cell(row=idx, column=c).font = Font(name=self.font_family, size=9)

            for d_i in range(total_days):
                c_idx = date_col_start + d_i
                c_let = get_column_letter(c_idx)
                cell_g = ws3.cell(row=idx, column=c_idx)
                cell_g.value = (
                    f"=IF(AND({c_let}$4>=$J{idx}, {c_let}$4<=$K{idx}), "
                    f"IF('01_THONG_SO_DU_AN'!$C$11=1, \"█\", "
                    f"IF('01_THONG_SO_DU_AN'!$C$11=2, $N{idx}, $M{idx})), \"\")"
                )
                cell_g.alignment = Alignment(horizontal="center", vertical="center")
                cell_g.font = Font(name=self.font_family, size=8, bold=True, color="1B365D")
                cell_g.border = self.thin_border

        r_end = r_start + len(tasks_cpm) - 1

        # Chân trang 1: Tổng số máy
        row_sum_m = r_end + 2
        ws3.merge_cells(f"B{row_sum_m}:N{row_sum_m}")
        ws3[f"B{row_sum_m}"] = "TỔNG SỐ LƯỢNG MÁY MÓC HUY ĐỘNG (Máy/ngày)"
        ws3[f"B{row_sum_m}"].font = Font(name=self.font_family, size=9.5, bold=True, color="FFFFFF")
        for c in range(1, date_col_start):
            ws3.cell(row=row_sum_m, column=c).fill = PatternFill("solid", fgColor=self.NAVY_HEADER)

        for d_i in range(total_days):
            c_idx = date_col_start + d_i
            c_let = get_column_letter(c_idx)
            cell_sm = ws3.cell(row=row_sum_m, column=c_idx)
            cell_sm.value = f"=SUMPRODUCT(({c_let}$4>=$J${r_start}:$J${r_end})*({c_let}$4<=$K${r_start}:$K${r_end})*$M${r_start}:$M${r_end})"
            cell_sm.alignment = Alignment(horizontal="center", vertical="center")
            cell_sm.font = Font(name=self.font_family, size=8.5, bold=True, color=self.NAVY_HEADER)
            cell_sm.fill = PatternFill("solid", fgColor=self.AMBER_ALERT)
            cell_sm.border = self.double_bottom_border
            cell_sm.number_format = "#,##0"

        # Chân trang 2: Tổng nhân công
        row_sum_nc = row_sum_m + 1
        ws3.merge_cells(f"B{row_sum_nc}:N{row_sum_nc}")
        ws3[f"B{row_sum_nc}"] = "TỔNG SỐ LƯỢNG NHÂN CÔNG HUY ĐỘNG (Người/ngày)"
        ws3[f"B{row_sum_nc}"].font = Font(name=self.font_family, size=9.5, bold=True, color="FFFFFF")
        for c in range(1, date_col_start):
            ws3.cell(row=row_sum_nc, column=c).fill = PatternFill("solid", fgColor=self.BLUE_SUB)

        for d_i in range(total_days):
            c_idx = date_col_start + d_i
            c_let = get_column_letter(c_idx)
            cell_snc = ws3.cell(row=row_sum_nc, column=c_idx)
            cell_snc.value = f"=SUMPRODUCT(({c_let}$4>=$J${r_start}:$J${r_end})*({c_let}$4<=$K${r_start}:$K${r_end})*$N${r_start}:$N${r_end})"
            cell_snc.alignment = Alignment(horizontal="center", vertical="center")
            cell_snc.font = Font(name=self.font_family, size=8.5, bold=True, color="002060")
            cell_snc.fill = PatternFill("solid", fgColor=self.LIGHT_GREEN)
            cell_snc.border = self.double_bottom_border
            cell_snc.number_format = "#,##0"

        # BIỂU ĐỒ 1: ĐƯỜNG CONG PHỤ TẢI MÁY MÓC & NHÂN CÔNG 97 NGÀY
        chart_load = LineChart()
        chart_load.title = "BIỂU ĐỒ ĐƯỜNG CONG PHỤ TẢI MÁY MÓC & NHÂN CÔNG HUY ĐỘNG THEO NGÀY"
        chart_load.style = 13
        chart_load.y_axis.title = "Số lượng (Máy & Người)"
        chart_load.x_axis.title = "Dòng thời gian (97 ngày lịch)"
        chart_load.width = 30
        chart_load.height = 14

        data_load = Reference(ws3, min_col=date_col_start, min_row=row_sum_m, max_col=date_col_start + total_days - 1, max_row=row_sum_nc)
        cats_load = Reference(ws3, min_col=date_col_start, min_row=4, max_col=date_col_start + total_days - 1, max_row=4)
        chart_load.add_data(data_load, from_rows=True, titles_from_data=False)
        chart_load.set_categories(cats_load)
        if len(chart_load.series) >= 2:
            chart_load.series[0].title = SeriesLabel(v="Tổng số máy móc/ngày")
            chart_load.series[1].title = SeriesLabel(v="Tổng nhân công/ngày")
        ws3.add_chart(chart_load, "B25")

        ws3.column_dimensions["A"].width = 5
        ws3.column_dimensions["B"].width = 7
        ws3.column_dimensions["C"].width = 44
        ws3.column_dimensions["D"].width = 8
        ws3.column_dimensions["E"].width = 14
        ws3.column_dimensions["F"].width = 9
        ws3.column_dimensions["G"].width = 14
        ws3.column_dimensions["H"].width = 14
        ws3.column_dimensions["I"].width = 10
        ws3.column_dimensions["J"].width = 13
        ws3.column_dimensions["K"].width = 13
        ws3.column_dimensions["L"].width = 12
        ws3.column_dimensions["M"].width = 12
        ws3.column_dimensions["N"].width = 13

        for d_i in range(total_days):
            col_let = get_column_letter(date_col_start + d_i)
            ws3.column_dimensions[col_let].width = 4.8

        # ---------------------------------------------------------------------
        # SHEET 4: 04_TONG_HOP_CA_MAY_VA_DAU (KÈM BIỂU ĐỒ CẤP PHÁT NHIÊN LIỆU)
        # ---------------------------------------------------------------------
        ws4 = wb.create_sheet(title="04_TONG_HOP_CA_MAY_VA_DAU")
        ws4.views.sheetView[0].showGridLines = True

        ws4["B2"] = "BẢNG TỔNG HỢP CA XE, CA MÁY & KẾ HOẠCH CẤP PHÁT NHIÊN LIỆU DẦU DIESEL"
        ws4["B2"].font = Font(name=self.font_family, size=13, bold=True, color=self.NAVY_HEADER)
        ws4["B3"] = "Liên kết động 100%: Tự động tính số ca máy, số máy Peak và lượng dầu từ Sheet 03"
        ws4["B3"].font = Font(name=self.font_family, size=10, italic=True)

        headers_fleet = [
            "TT", "Mã máy", "Tên phương tiện / Thiết bị", "Định mức dầu (L/ca)\n(=VLOOKUP)",
            "Tổng số ca máy\n(=SUMIF)", "Số máy huy động Max\n(=SUMIF)", "Tổng lượng dầu Diesel (Lít)\n(=Số ca * Định mức)",
            "Dầu Tháng 8 (L)", "Dầu Tháng 9 (L)", "Dầu Tháng 10 (L)", "Dầu Tháng 11 (L)"
        ]
        for c_i, h in enumerate(headers_fleet, start=2):
            cell = ws4.cell(row=5, column=c_i, value=h)
            cell.font = Font(name=self.font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        fleet_codes = [
            (1, "M1", "Máy xúc bánh xích PC200 - PC300"),
            (2, "M3", "Máy ủi D3 - D5"),
            (3, "M4", "Máy ủi công suất lớn CS 230CV"),
            (4, "M5", "Máy lu rung 12T - 16T"),
            (5, "M6", "Máy san tự hành 110CV"),
            (6, "M7", "Ô tô tự đổ 18 tấn"),
            (7, "M8", "Máy xúc bánh lốp PC140"),
            (8, "MP1", "Xe téc & Xe cấp dầu hiện trường"),
        ]

        for r_i, (tt, code, name) in enumerate(fleet_codes, start=6):
            ws4.cell(row=r_i, column=2, value=tt).alignment = Alignment(horizontal="center")
            ws4.cell(row=r_i, column=3, value=code).alignment = Alignment(horizontal="center")
            ws4.cell(row=r_i, column=4, value=name)
            ws4.cell(row=r_i, column=5, value=f"=VLOOKUP(C{r_i}, '02_DINH_MUC_CA_MAY_VA_DAU'!$C$6:$H$16, 6, FALSE)").number_format = "#,##0.0"
            ws4.cell(row=r_i, column=6, value=f"=SUMIF('03_TIEN_DO_GANTT_CPM'!$F$6:$F$19, C{r_i}, '03_TIEN_DO_GANTT_CPM'!$H$6:$H$19)").number_format = "#,##0.0"
            ws4.cell(row=r_i, column=7, value=f"=SUMIF('03_TIEN_DO_GANTT_CPM'!$F$6:$F$19, C{r_i}, '03_TIEN_DO_GANTT_CPM'!$M$6:$M$19)").number_format = "#,##0"
            ws4.cell(row=r_i, column=8, value=f"=F{r_i}*E{r_i}").number_format = "#,##0.0"
            ws4.cell(row=r_i, column=9, value=f"=ROUND(H{r_i}*0.21, 0)").number_format = "#,##0"
            ws4.cell(row=r_i, column=10, value=f"=ROUND(H{r_i}*0.39, 0)").number_format = "#,##0"
            ws4.cell(row=r_i, column=11, value=f"=ROUND(H{r_i}*0.32, 0)").number_format = "#,##0"
            ws4.cell(row=r_i, column=12, value=f"=H{r_i}-I{r_i}-J{r_i}-K{r_i}").number_format = "#,##0"

            for c in range(2, 13):
                ws4.cell(row=r_i, column=c).border = self.thin_border
                ws4.cell(row=r_i, column=c).font = Font(name=self.font_family, size=9)

        row_sum_fleet = 6 + len(fleet_codes)
        ws4.merge_cells(f"B{row_sum_fleet}:D{row_sum_fleet}")
        ws4[f"B{row_sum_fleet}"] = "TỔNG CỘNG TOÀN DỰ ÁN"
        ws4[f"B{row_sum_fleet}"].font = Font(name=self.font_family, size=10, bold=True)
        ws4[f"B{row_sum_fleet}"].alignment = Alignment(horizontal="center")
        ws4.cell(row=row_sum_fleet, column=6, value=f"=SUM(F6:F{row_sum_fleet-1})").number_format = "#,##0.0"
        ws4.cell(row=row_sum_fleet, column=7, value=f"=SUM(G6:G{row_sum_fleet-1})").number_format = "#,##0"
        ws4.cell(row=row_sum_fleet, column=8, value=f"=SUM(H6:H{row_sum_fleet-1})").number_format = "#,##0.0"
        ws4.cell(row=row_sum_fleet, column=9, value=f"=SUM(I6:I{row_sum_fleet-1})").number_format = "#,##0"
        ws4.cell(row=row_sum_fleet, column=10, value=f"=SUM(J6:J{row_sum_fleet-1})").number_format = "#,##0"
        ws4.cell(row=row_sum_fleet, column=11, value=f"=SUM(K6:K{row_sum_fleet-1})").number_format = "#,##0"
        ws4.cell(row=row_sum_fleet, column=12, value=f"=SUM(L6:L{row_sum_fleet-1})").number_format = "#,##0"

        for c in range(2, 13):
            cell = ws4.cell(row=row_sum_fleet, column=c)
            cell.font = Font(name=self.font_family, size=10, bold=True)
            cell.border = self.thick_bottom_border
            cell.fill = PatternFill("solid", fgColor=self.GREEN_OK)

        # BIỂU ĐỒ 2: TIÊU THỤ DẦU DIESEL THEO 4 THÁNG
        chart_fuel = BarChart()
        chart_fuel.type = "col"
        chart_fuel.style = 10
        chart_fuel.title = "KẾ HOẠCH CẤP PHÁT NHIÊN LIỆU DẦU DIESEL THEO 4 THÁNG (LÍT)"
        chart_fuel.y_axis.title = "Tổng lượng dầu Diesel (Lít)"
        chart_fuel.x_axis.title = "Kỳ thi công"
        chart_fuel.width = 18
        chart_fuel.height = 12

        ws4["I16"] = "Tháng 8"
        ws4["J16"] = "Tháng 9"
        ws4["K16"] = "Tháng 10"
        ws4["L16"] = "Tháng 11"
        for col_c in ["I", "J", "K", "L"]:
            ws4[f"{col_c}16"].font = Font(name=self.font_family, size=8, color="FFFFFF")

        data_fuel = Reference(ws4, min_col=9, min_row=row_sum_fleet, max_col=12, max_row=row_sum_fleet)
        cats_fuel = Reference(ws4, min_col=9, min_row=16, max_col=12, max_row=16)
        chart_fuel.add_data(data_fuel, from_rows=True, titles_from_data=False)
        chart_fuel.set_categories(cats_fuel)
        if chart_fuel.series:
            chart_fuel.series[0].title = SeriesLabel(v="Lượng dầu Diesel (Lít)")
        ws4.add_chart(chart_fuel, "N4")

        ws4.column_dimensions["B"].width = 6
        ws4.column_dimensions["C"].width = 10
        ws4.column_dimensions["D"].width = 34
        ws4.column_dimensions["E"].width = 18
        ws4.column_dimensions["F"].width = 20
        ws4.column_dimensions["G"].width = 20
        ws4.column_dimensions["H"].width = 24
        ws4.column_dimensions["I"].width = 14
        ws4.column_dimensions["J"].width = 14
        ws4.column_dimensions["K"].width = 14
        ws4.column_dimensions["L"].width = 14

        # ---------------------------------------------------------------------
        # SHEET 5: 05_NHU_CAU_VAT_TU_CHINH (KÈM BIỂU ĐỒ CƠ CẤU VẬT TƯ)
        # ---------------------------------------------------------------------
        ws5 = wb.create_sheet(title="05_NHU_CAU_VAT_TU_CHINH")
        ws5.views.sheetView[0].showGridLines = True

        ws5["B2"] = "KẾ HOẠCH TỔNG HỢP NHU CẦU & TIẾN ĐỘ CẤP VẬT TƯ CHÍNH CỤM B9"
        ws5["B2"].font = Font(name=self.font_family, size=13, bold=True, color=self.NAVY_HEADER)
        ws5["B3"] = "Tự động trích xuất khối lượng đào đắp từ Sheet 03 và phân bổ kế hoạch nhập bãi theo 4 tháng"
        ws5["B3"].font = Font(name=self.font_family, size=10, italic=True)

        headers_mat = [
            "STT", "Mã VT", "Tên quy cách vật tư chính", "ĐVT", "Tổng nhu cầu (=Sheet 03)",
            "Tháng 8 (20%)", "Tháng 9 (40%)", "Tháng 10 (30%)", "Tháng 11 (10%)", "Ghi chú điều phối bãi trữ"
        ]
        for c_i, h in enumerate(headers_mat, start=2):
            cell = ws5.cell(row=4, column=c_i, value=h)
            cell.font = Font(name=self.font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        materials_list = [
            (1, "VT01", "Cát đắp nền đường K90/K95", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$9+'03_TIEN_DO_GANTT_CPM'!$E$10", "Tập kết tại bãi trữ Cụm B9"),
            (2, "VT02", "Cát đắp chọn lọc K98 lớp 1", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$11", "Yêu cầu thí nghiệm độ nén trước khi lu"),
            (3, "VT03", "Cấp phối đá dăm loại 2 (Base B)", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$14", "Nhập từ mỏ đá Kiện Khê - Hà Nam"),
            (4, "VT04", "Cấp phối đá dăm loại 1 (Base A)", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$15", "Đảm bảo độ ẩm tiêu chuẩn khi rải"),
            (5, "VT05", "Bê tông nhựa hạt trung BTN C19", "m2", "='03_TIEN_DO_GANTT_CPM'!$E$16", "Nhập trạm trộn cự ly vận chuyển 12km"),
            (6, "VT06", "Cống bê tông ly tâm D1500", "m", "='03_TIEN_DO_GANTT_CPM'!$E$12", "Đúc sẵn từ nhà máy AMACCAO"),
            (7, "VT07", "Đất đắp tạo cảnh quan vỉa hè", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$18", "Tận dụng đất hữu cơ bóc phủ đã xử lý"),
            (8, "VT08", "Đất đắp san nền tổng thể B9", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$19", "Nhập đất đồi vận chuyển bằng xe 18T"),
        ]

        for r_i, m in enumerate(materials_list, start=5):
            stt, code, name, unit, f_qty, note = m
            ws5.cell(row=r_i, column=2, value=stt).alignment = Alignment(horizontal="center")
            ws5.cell(row=r_i, column=3, value=code).alignment = Alignment(horizontal="center")
            ws5.cell(row=r_i, column=4, value=name)
            ws5.cell(row=r_i, column=5, value=unit).alignment = Alignment(horizontal="center")
            ws5.cell(row=r_i, column=6, value=f_qty).number_format = "#,##0.0"
            ws5.cell(row=r_i, column=7, value=f"=ROUND(F{r_i}*0.2, 1)").number_format = "#,##0.0"
            ws5.cell(row=r_i, column=8, value=f"=ROUND(F{r_i}*0.4, 1)").number_format = "#,##0.0"
            ws5.cell(row=r_i, column=9, value=f"=ROUND(F{r_i}*0.3, 1)").number_format = "#,##0.0"
            ws5.cell(row=r_i, column=10, value=f"=F{r_i}-G{r_i}-H{r_i}-I{r_i}").number_format = "#,##0.0"
            ws5.cell(row=r_i, column=11, value=note)

            for c in range(2, 12):
                ws5.cell(row=r_i, column=c).border = self.thin_border
                ws5.cell(row=r_i, column=c).font = Font(name=self.font_family, size=9)

        # BIỂU ĐỒ 3: CƠ CẤU VẬT TƯ SAN LẤP & NỀN MÓNG
        chart_mat = BarChart()
        chart_mat.type = "bar"
        chart_mat.style = 11
        chart_mat.title = "CƠ CẤU KHỐI LƯỢNG VẬT TƯ CHÍNH CỤM B9"
        chart_mat.x_axis.title = "Khối lượng nhu cầu"
        chart_mat.y_axis.title = "Quy cách vật tư"
        chart_mat.width = 20
        chart_mat.height = 12

        data_mat = Reference(ws5, min_col=6, min_row=4, max_col=6, max_row=4 + len(materials_list))
        cats_mat = Reference(ws5, min_col=4, min_row=5, max_col=4, max_row=4 + len(materials_list))
        chart_mat.add_data(data_mat, titles_from_data=True)
        chart_mat.set_categories(cats_mat)
        ws5.add_chart(chart_mat, "M4")

        ws5.column_dimensions["B"].width = 6
        ws5.column_dimensions["C"].width = 10
        ws5.column_dimensions["D"].width = 32
        ws5.column_dimensions["E"].width = 8
        ws5.column_dimensions["F"].width = 24
        ws5.column_dimensions["G"].width = 16
        ws5.column_dimensions["H"].width = 16
        ws5.column_dimensions["I"].width = 16
        ws5.column_dimensions["J"].width = 16
        ws5.column_dimensions["K"].width = 38

        # ---------------------------------------------------------------------
        # SHEET 6: 06_SO_SANH_DINH_MUC_VS_THUC_TE (PHỤC HỒI & NÂNG CẤP SHEET SS VINA ALPHA)
        # ---------------------------------------------------------------------
        ws6 = wb.create_sheet(title="06_SO_SANH_DINH_MUC_VS_THUC_TE")
        ws6.views.sheetView[0].showGridLines = True

        ws6["B2"] = "BẢNG ĐỐI SÁNH & CÂN ĐỐI MÁY MÓC: ĐỊNH MỨC TÍNH TOÁN VS ĐỀ XUẤT BAN ĐIỀU HÀNH"
        ws6["B2"].font = Font(name=self.font_family, size=13, bold=True, color=self.NAVY_HEADER)
        ws6["B3"] = "Phục dựng nguyên bản Sheet SS của Vina Alpha: 100% công thức động, loại bỏ hoàn toàn lỗi #REF!"
        ws6["B3"].font = Font(name=self.font_family, size=10, italic=True)

        headers_ss = [
            "TT", "Mã máy", "Chủng loại thiết bị", "Định mức tính toán\n(Sheet 04)",
            "Đề xuất Ban ĐH\n(Thực tế)", "Chênh lệch\n(=Thực tế - ĐM)", "Tỷ lệ tăng giảm\n(%)",
            "Dầu theo ĐM\n(Lít)", "Dầu theo Thực tế\n(Lít)", "Đánh giá kỹ thuật & Lý do chênh lệch hiện trường"
        ]
        for c_i, h in enumerate(headers_ss, start=2):
            cell = ws6.cell(row=5, column=c_i, value=h)
            cell.font = Font(name=self.font_family, size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=self.NAVY_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        equip_comparison = [
            (1, "M1", "Máy xúc bánh xích PC300", 6, 3, "Bổ sung 1 máy trực sự cố sụt trượt và đào vét bùn phát sinh"),
            (2, "M3", "Máy ủi D3 - D5", 7, 7, "Tăng 1 máy ủi hỗ trợ san gạt cát tại bãi tập kết"),
            (3, "M4", "Máy ủi công suất lớn 230CV", 8, 8, "Tăng 1 máy để đẩy nhanh tiến độ san nền cụm B9 trước mùa mưa"),
            (4, "M5", "Máy lu rung 12T - 16T", 9, 4, "Tăng 1 máy lu sơ bộ K90 độc lập với lu K95"),
            (5, "M6", "Máy san tự hành 110CV", 10, 2, "Bổ sung 1 máy san dự phòng khi làm lớp Base A/B hoàn thiện"),
            (6, "M7", "Ô tô tự đổ 18 tấn", 11, 6, "Tăng 2 xe do cung đường vận chuyển nội bộ lầy lội khi trời mưa"),
            (7, "M8", "Máy xúc bánh lốp PC140", 12, 4, "Đúng định mức thi công cống D1500 và lắp ống PE"),
            (8, "MP1", "Xe téc & Xe cấp dầu", 13, 2, "Tăng 1 xe téc tưới nước chống bụi đường nội bộ theo cam kết ĐTM"),
        ]

        for r_idx, (tt, code, name, r_s4, val_actual, reason) in enumerate(equip_comparison, start=6):
            ws6.cell(row=r_idx, column=2, value=tt).alignment = Alignment(horizontal="center")
            ws6.cell(row=r_idx, column=3, value=code).alignment = Alignment(horizontal="center")
            ws6.cell(row=r_idx, column=4, value=name)
            ws6.cell(row=r_idx, column=5, value=f"='04_TONG_HOP_CA_MAY_VA_DAU'!$G${r_s4}").number_format = "#,##0"
            ws6.cell(row=r_idx, column=6, value=val_actual).number_format = "#,##0"
            ws6.cell(row=r_idx, column=7, value=f"=F{r_idx}-E{r_idx}").number_format = "#,##0"
            ws6.cell(row=r_idx, column=8, value=f"=IF(E{r_idx}>0, (F{r_idx}-E{r_idx})/E{r_idx}, 0)").number_format = "0.0%"
            ws6.cell(row=r_idx, column=9, value=f"='04_TONG_HOP_CA_MAY_VA_DAU'!$H${r_s4}").number_format = "#,##0.0"
            ws6.cell(row=r_idx, column=10, value=f"=ROUND(I{r_idx}*(F{r_idx}/E{r_idx}), 1)").number_format = "#,##0.0"
            ws6.cell(row=r_idx, column=11, value=reason)

            for c in range(2, 12):
                ws6.cell(row=r_idx, column=c).border = self.thin_border
                ws6.cell(row=r_idx, column=c).font = Font(name=self.font_family, size=9)

        r_sum_ss = 6 + len(equip_comparison)
        ws6.merge_cells(f"B{r_sum_ss}:D{r_sum_ss}")
        ws6[f"B{r_sum_ss}"] = "TỔNG CỘNG"
        ws6[f"B{r_sum_ss}"].font = Font(name=self.font_family, size=10, bold=True)
        ws6[f"B{r_sum_ss}"].alignment = Alignment(horizontal="center")
        ws6.cell(row=r_sum_ss, column=5, value=f"=SUM(E6:E{r_sum_ss-1})").number_format = "#,##0"
        ws6.cell(row=r_sum_ss, column=6, value=f"=SUM(F6:F{r_sum_ss-1})").number_format = "#,##0"
        ws6.cell(row=r_sum_ss, column=7, value=f"=SUM(G6:G{r_sum_ss-1})").number_format = "#,##0"
        ws6.cell(row=r_sum_ss, column=8, value=f"=(F{r_sum_ss}-E{r_sum_ss})/E{r_sum_ss}").number_format = "0.0%"
        ws6.cell(row=r_sum_ss, column=9, value=f"=SUM(I6:I{r_sum_ss-1})").number_format = "#,##0.0"
        ws6.cell(row=r_sum_ss, column=10, value=f"=SUM(J6:J{r_sum_ss-1})").number_format = "#,##0.0"

        for c in range(2, 12):
            cell = ws6.cell(row=r_sum_ss, column=c)
            cell.font = Font(name=self.font_family, size=10, bold=True)
            cell.border = self.thick_bottom_border
            cell.fill = PatternFill("solid", fgColor=self.GREEN_OK)

        # BIỂU ĐỒ 4: SO SÁNH THIẾT BỊ: ĐỊNH MỨC VS ĐỀ XUẤT THỰC TẾ
        chart_ss = BarChart()
        chart_ss.type = "col"
        chart_ss.style = 10
        chart_ss.title = "BIỂU ĐỒ SO SÁNH SỐ LƯỢNG THIẾT BỊ: ĐỊNH MỨC VS THỰC TẾ (CHIẾC)"
        chart_ss.y_axis.title = "Số lượng máy huy động (Chiếc)"
        chart_ss.x_axis.title = "Chủng loại thiết bị"
        chart_ss.width = 22
        chart_ss.height = 13

        data_ss = Reference(ws6, min_col=5, min_row=5, max_col=6, max_row=5 + len(equip_comparison))
        cats_ss = Reference(ws6, min_col=4, min_row=6, max_row=13)
        chart_ss.add_data(data_ss, titles_from_data=True)
        chart_ss.set_categories(cats_ss)
        ws6.add_chart(chart_ss, "M4")

        ws6.column_dimensions["B"].width = 6
        ws6.column_dimensions["C"].width = 10
        ws6.column_dimensions["D"].width = 30
        ws6.column_dimensions["E"].width = 18
        ws6.column_dimensions["F"].width = 18
        ws6.column_dimensions["G"].width = 16
        ws6.column_dimensions["H"].width = 16
        ws6.column_dimensions["I"].width = 18
        ws6.column_dimensions["J"].width = 20
        ws6.column_dimensions["K"].width = 50

        # Lưu file
        wb.save(output_file)
        return output_file
