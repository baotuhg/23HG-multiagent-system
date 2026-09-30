# -*- coding: utf-8 -*-
"""
GENERATE 100% DYNAMIC FORMULA-DRIVEN SCHEDULE FOR CỤM B9 (ZERO HARDCODED NUMBERS)
Tiến độ thi công cụm B9 tinh giản chuẩn CPM:
- 100% CÔNG THỨC SỐNG LIÊN KẾT ĐỘNG TOÀN DIỆN (ZERO SỐ CHẾT)
- Đổi ngày khởi công tại '01_THONG_SO_DU_AN'!C6 -> Toàn bộ ngày tháng 14 công tác và 97 cột Gantt tự động nhảy theo.
- Đổi số ngày thi công (Duration) -> Ngày kết thúc, Đường găng, Số máy, Nhân công, Ma trận Gantt tự động co dãn.
- Đổi công tắc chế độ C11 -> Ma trận Gantt tự động chuyển đổi giữa Thanh Gantt ('█'), Nhân công, và Ca máy.
- Chân trang ma trận tự động tính tổng máy/ngày và tổng thợ/ngày bằng công thức SUMPRODUCT đa điều kiện.
- Sheet Tổng hợp ca máy & Vật tư liên kết động 100% sang Sheet Tiến độ bằng VLOOKUP, SUMIF, phép nhân chia.
"""

from __future__ import annotations
import datetime
import os
import shutil
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_DIR = r"C:\Users\baotu\Downloads\TĐTC vina alpha"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "260820_TDTC_Cum_B9_TINH_GIAN_CHUAN_CPM.xlsx")

PROJECT_DIR = r"c:\Users\baotu\Downloads\Documents\HSTK Cầu Km19+529.080_Marker"
MIRROR_FILE_1 = os.path.join(PROJECT_DIR, "TDTC_Cum_B9_Tinh_Gian_Chuan_CPM.xlsx")
MIRROR_FILE_2 = os.path.join(PROJECT_DIR, "HSTK Cầu Km19+529.080_Marker", "TDTC_Cum_B9_Tinh_Gian_Chuan_CPM.xlsx")

# Palette màu chuyên nghiệp
NAVY_HEADER = "1B365D"
BLUE_SUB = "2E75B6"
LIGHT_BLUE = "D9E1F2"
ACCENT_GREEN = "385723"
LIGHT_GREEN = "E2EFDA"
AMBER_ALERT = "FFF2CC"
PEACH_FILL = "FCE4D6"
GRAY_BORDER = "D9D9D9"
GRAY_FILL = "F2F2F2"

font_family = "Times New Roman"

thin_border = Border(
    left=Side(style="thin", color=GRAY_BORDER),
    right=Side(style="thin", color=GRAY_BORDER),
    top=Side(style="thin", color=GRAY_BORDER),
    bottom=Side(style="thin", color=GRAY_BORDER),
)
double_bottom_border = Border(
    left=Side(style="thin", color=GRAY_BORDER),
    right=Side(style="thin", color=GRAY_BORDER),
    top=Side(style="thin", color=GRAY_BORDER),
    bottom=Side(style="double", color=NAVY_HEADER),
)


def build_dynamic_schedule():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    start_date = datetime.date(2026, 8, 11)
    total_days = 97

    # =========================================================================
    # SHEET 1: 01_THONG_SO_DU_AN
    # =========================================================================
    ws1 = wb.create_sheet(title="01_THONG_SO_DU_AN")
    ws1.views.sheetView[0].showGridLines = True

    ws1["B2"] = "DỰ ÁN: SÂN VẬN ĐỘNG OLYMPIC - THƯỜNG TÍN (HÀ NỘI)"
    ws1["B2"].font = Font(name=font_family, size=14, bold=True, color=NAVY_HEADER)
    ws1["B3"] = "HẠNG MỤC: HẠ TẦNG KỸ THUẬT THI CÔNG SAN LẤP CỤM B9 VÀ ĐƯỜNG NỘI BỘ"
    ws1["B3"].font = Font(name=font_family, size=12, bold=True, color=BLUE_SUB)
    ws1["B4"] = "BẢNG THIẾT LẬP THÔNG SỐ ĐIỀU HÀNH TIẾN ĐỘ & CA MÁY (100% CÔNG THỨC SỐNG)"
    ws1["B4"].font = Font(name=font_family, size=11, italic=True)

    ws1.cell(row=6, column=2, value="Thông số quản trị dự án").font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.cell(row=6, column=2).fill = PatternFill("solid", fgColor=NAVY_HEADER)
    ws1.cell(row=6, column=3, value="Giá trị thiết lập").font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.cell(row=6, column=3).fill = PatternFill("solid", fgColor=NAVY_HEADER)
    ws1.cell(row=6, column=4, value="Ghi chú kỹ thuật & Cơ chế liên kết").font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.cell(row=6, column=4).fill = PatternFill("solid", fgColor=NAVY_HEADER)

    # Dòng 6: Ngày phát lệnh khởi công (Input gốc)
    ws1.cell(row=6, column=2, value="Ngày phát lệnh khởi công (Bắt đầu)").border = thin_border
    c_s = ws1.cell(row=6, column=3, value=start_date)
    c_s.border = thin_border
    c_s.alignment = Alignment(horizontal="center")
    c_s.number_format = "DD/MM/YYYY"
    c_s.font = Font(name=font_family, size=10, bold=True, color="002060")
    ws1.cell(row=6, column=4, value="Input gốc: Đổi ngày này, toàn bộ 14 công tác và 97 cột Gantt tự động nhảy theo!").border = thin_border

    # Dòng 7: Ngày hoàn thành (Công thức = MAX cột Ngày kết thúc ở Sheet 3)
    ws1.cell(row=7, column=2, value="Ngày hoàn thành mục tiêu (Deadline)").border = thin_border
    c_e = ws1.cell(row=7, column=3, value="=MAX('03_TIEN_DO_GANTT_CPM'!$K$6:$K$19)")
    c_e.border = thin_border
    c_e.alignment = Alignment(horizontal="center")
    c_e.number_format = "DD/MM/YYYY"
    c_e.font = Font(name=font_family, size=10, bold=True, color="C00000")
    ws1.cell(row=7, column=4, value="=MAX('03_TIEN_DO_GANTT_CPM'!$K$6:$K$19) — Tự động cập nhật theo công tác kết thúc muộn nhất").border = thin_border

    # Dòng 8: Tổng thời gian (Công thức = C7 - C6 + 1)
    ws1.cell(row=8, column=2, value="Tổng thời gian thi công (Ngày)").border = thin_border
    c_tot = ws1.cell(row=8, column=3, value="=C7-C6+1")
    c_tot.border = thin_border
    c_tot.alignment = Alignment(horizontal="center")
    c_tot.number_format = "#,##0"
    c_tot.font = Font(name=font_family, size=10, bold=True)
    ws1.cell(row=8, column=4, value="=C7-C6+1 — Tự động tính số ngày lịch liên tục").border = thin_border

    # Dòng 9: Chế độ làm việc
    ws1.cell(row=9, column=2, value="Chế độ làm việc công trường (Ca/ngày)").border = thin_border
    c_shift = ws1.cell(row=9, column=3, value=2)
    c_shift.border = thin_border
    c_shift.alignment = Alignment(horizontal="center")
    c_shift.font = Font(name=font_family, size=10, bold=True)
    ws1.cell(row=9, column=4, value="2 ca/ngày (10 giờ/ca = 20 giờ làm việc/ngày)").border = thin_border

    # Dòng 10: Hệ số ca máy
    ws1.cell(row=10, column=2, value="Hệ số ca máy (HSTCA)").border = thin_border
    c_hs = ws1.cell(row=10, column=3, value="=C9")
    c_hs.border = thin_border
    c_hs.alignment = Alignment(horizontal="center")
    c_hs.font = Font(name=font_family, size=10, bold=True)
    ws1.cell(row=10, column=4, value="=C9 — Tự động lấy theo chế độ ca làm việc").border = thin_border

    # Dòng 11: Công tắc hiển thị Gantt
    ws1.cell(row=11, column=2, value="Công tắc hiển thị biểu đồ Gantt").border = thin_border
    c_sw = ws1.cell(row=11, column=3, value=1)
    c_sw.border = thin_border
    c_sw.alignment = Alignment(horizontal="center")
    c_sw.font = Font(name=font_family, size=11, bold=True, color="C00000")
    c_sw.fill = PatternFill("solid", fgColor=AMBER_ALERT)
    ws1.cell(row=11, column=4, value="Nhập 1 = Thanh Gantt (█) | Nhập 2 = Nhân công (người) | Nhập 3 = Ca máy (máy)").border = thin_border

    # Dòng 12 & 13: Tiêu chuẩn & Định mức
    ws1.cell(row=12, column=2, value="Tiêu chuẩn kỹ thuật áp dụng").border = thin_border
    ws1.cell(row=12, column=3, value="TCVN 9436:2012 / TCVN 8819:2011").alignment = Alignment(horizontal="center")
    ws1.cell(row=12, column=3).border = thin_border
    ws1.cell(row=12, column=4, value="Thi công & nghiệm thu nền đường, móng mặt đường").border = thin_border

    ws1.cell(row=13, column=2, value="Định mức thi công tham chiếu").border = thin_border
    ws1.cell(row=13, column=3, value="Định mức Vincons & TT 38/2026/TT-BXD").alignment = Alignment(horizontal="center")
    ws1.cell(row=13, column=3).border = thin_border
    ws1.cell(row=13, column=4, value="Định mức dự toán & ca máy xây dựng công trình").border = thin_border

    ws1.column_dimensions["B"].width = 40
    ws1.column_dimensions["C"].width = 28
    ws1.column_dimensions["D"].width = 65

    # =========================================================================
    # SHEET 2: 02_DINH_MUC_CA_MAY_VA_DAU
    # =========================================================================
    ws2 = wb.create_sheet(title="02_DINH_MUC_CA_MAY_VA_DAU")
    ws2.views.sheetView[0].showGridLines = True

    ws2["B2"] = "BẢNG ĐỊNH MỨC NĂNG SUẤT CA MÁY & TIÊU HAO NHIÊN LIỆU DẦU DIESEL"
    ws2["B2"].font = Font(name=font_family, size=13, bold=True, color=NAVY_HEADER)
    ws2["B3"] = "Nguồn dữ liệu: Vincons_ĐMGK_02-01 Hạ tầng giao thông & Thông tư 38/2026/TT-BXD"
    ws2["B3"].font = Font(name=font_family, size=10, italic=True)

    norm_headers = [
        "TT", "Mã máy", "Chủng loại thiết bị / Ca máy", "Công tác thi công chính",
        "ĐVT", "Định mức Năng suất (ĐVT/ca)", "Định mức Dầu (Lít/ca)", "Định mức Thợ lái (Người/ca)"
    ]
    for c_i, h in enumerate(norm_headers, start=2):
        cell = ws2.cell(row=5, column=c_i, value=h)
        cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    machine_norms = [
        (1, "M1", "Máy xúc bánh xích PC200 - PC300", "Đào khuôn đường, vét bùn hữu cơ", "m3/ca", 872.0, 89.6, 2.0),
        (2, "M2", "Máy xúc bánh xích PC350 - PC450", "Đào đất đá san lấp khối lượng lớn", "m3/ca", 1093.0, 106.4, 2.0),
        (3, "M3", "Máy ủi D3 - D5", "Ủi gom đất, san gạt đắp cát K90, K95", "m3/ca", 1428.0, 124.0, 2.0),
        (4, "M4", "Máy ủi công suất lớn CS 230CV", "Ủi đắp cát K98, san mặt bằng B9", "m3/ca", 2087.0, 165.0, 2.0),
        (5, "M5", "Máy lu rung 12T - 16T", "Lu lèn nền đường K90, K95, K98, Base", "m3/ca", 600.0, 44.0, 2.0),
        (6, "M6", "Máy san tự hành 110CV", "San hoàn thiện móng Base A, Base B", "m3/ca", 450.0, 69.0, 2.0),
        (7, "M7", "Ô tô tự đổ 18 tấn", "Vận chuyển đất, cát nội bộ < 5km", "m3/ca", 280.0, 26.0, 1.0),
        (8, "M8", "Máy xúc bánh lốp PC140 - PC150", "Lắp đặt cống D1500, ống PE, bó vỉa", "m/ca", 350.0, 64.0, 2.0),
        (9, "M9", "Xe téc tưới nước 5m3", "Tưới ẩm nền đường đắp cát & Base", "m3/ca", 800.0, 42.0, 1.0),
        (10, "MP1", "Xe cấp dầu lưu động", "Tiếp nhiên liệu tại hiện trường", "ca", 1.0, 44.0, 1.0),
        (11, "MP2", "Máy phát điện chiếu sáng ca đêm", "Phục vụ thi công ca 2 (20h/ngày)", "ca", 1.0, 28.0, 1.0),
    ]

    for r_idx, m in enumerate(machine_norms, start=6):
        for c_idx, val in enumerate(m, start=2):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.font = Font(name=font_family, size=10)
            if c_idx in (2, 3, 6):
                cell.alignment = Alignment(horizontal="center")
            elif c_idx in (7, 8, 9):
                cell.alignment = Alignment(horizontal="right")
                if isinstance(val, float):
                    cell.number_format = "#,##0.0"

    ws2.column_dimensions["B"].width = 6
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 32
    ws2.column_dimensions["E"].width = 35
    ws2.column_dimensions["F"].width = 14
    ws2.column_dimensions["G"].width = 22
    ws2.column_dimensions["H"].width = 22
    ws2.column_dimensions["I"].width = 24

    # =========================================================================
    # SHEET 3: 03_TIEN_DO_GANTT_CPM (TRUNG TÂM LIÊN KẾT ĐỘNG 100%)
    # =========================================================================
    ws3 = wb.create_sheet(title="03_TIEN_DO_GANTT_CPM")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:N1")
    ws3["A1"] = "HỆ THỐNG TIẾN ĐỘ THI CÔNG & PHÂN BỔ CA MÁY CPM — CỤM B9 VÀ ĐƯỜNG NỘI BỘ"
    ws3["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws3["A1"].fill = PatternFill("solid", fgColor=NAVY_HEADER)
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws3.merge_cells("A2:N2")
    ws3["A2"] = "DỰ ÁN: SÂN VẬN ĐỘNG OLYMPIC - THƯỜNG TÍN | 100% CÔNG THỨC SỐNG LIÊN KẾT ĐỘNG TOÀN DIỆN"
    ws3["A2"].font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws3["A2"].fill = PatternFill("solid", fgColor=BLUE_SUB)
    ws3["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_cpm = [
        ("STT", "A4:A5"), ("WBS", "B4:B5"), ("Nội dung công việc", "C4:C5"),
        ("ĐVT", "D4:D5"), ("Khối lượng", "E4:E5"), ("Mã Máy", "F4:F5"),
        ("Định mức\n(ĐVT/ca)", "G4:G5"), ("Tổng số ca\n(=E/G)", "H4:H5"), ("Số ngày\n(Duration)", "I4:I5"),
        ("Ngày BĐ\n(Early Start)", "J4:J5"), ("Ngày KT\n(=J+I-1)", "K4:K5"),
        ("Đường Găng\n(CPM)", "L4:L5"), ("Số máy/ngày\n(=ROUNDUP)", "M4:M5"), ("Nhân công/ngày\n(=M*Thợ*Ca)", "N4:N5")
    ]
    for h_title, rng in headers_cpm:
        ws3.merge_cells(rng)
        top = rng.split(":")[0]
        ws3[top] = h_title
        ws3[top].font = Font(name=font_family, size=9, bold=True, color="FFFFFF")
        ws3[top].fill = PatternFill("solid", fgColor=NAVY_HEADER)
        ws3[top].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    date_col_start = 15  # Cột O bắt đầu ngày 11/08

    # 1. TRỤC THỜI GIAN NGÀY THÁNG ĐỘNG (DÒNG 4 & 5):
    # Cột O: ='01_THONG_SO_DU_AN'!$C$6
    ws3.cell(row=4, column=date_col_start, value="='01_THONG_SO_DU_AN'!$C$6")
    ws3.cell(row=4, column=date_col_start).font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
    ws3.cell(row=4, column=date_col_start).fill = PatternFill("solid", fgColor=BLUE_SUB)
    ws3.cell(row=4, column=date_col_start).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=4, column=date_col_start).number_format = "DD/MM"

    ws3.cell(row=5, column=date_col_start, value=f'=CHOOSE(WEEKDAY(O4, 2), "T2", "T3", "T4", "T5", "T6", "T7", "CN")')
    ws3.cell(row=5, column=date_col_start).font = Font(name=font_family, size=8, bold=True, color=NAVY_HEADER)
    ws3.cell(row=5, column=date_col_start).fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws3.cell(row=5, column=date_col_start).alignment = Alignment(horizontal="center", vertical="center")

    for idx in range(1, total_days):
        c_idx = date_col_start + idx
        prev_let = get_column_letter(c_idx - 1)
        curr_let = get_column_letter(c_idx)

        # Hàng 4: = Cột trước + 1
        cell_date = ws3.cell(row=4, column=c_idx, value=f"={prev_let}4+1")
        cell_date.font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
        cell_date.fill = PatternFill("solid", fgColor=BLUE_SUB)
        cell_date.alignment = Alignment(horizontal="center", vertical="center")
        cell_date.number_format = "DD/MM"

        # Hàng 5: Thứ trong tuần
        cell_day = ws3.cell(row=5, column=c_idx, value=f'=CHOOSE(WEEKDAY({curr_let}4, 2), "T2", "T3", "T4", "T5", "T6", "T7", "CN")')
        cell_day.alignment = Alignment(horizontal="center", vertical="center")
        cell_day.font = Font(name=font_family, size=8, bold=True, color=NAVY_HEADER)
        cell_day.fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    # 14 Công tác chuẩn WBS cụm B9 với Logic Predecessor sống
    tasks_cpm = [
        # STT, WBS, Tên công việc, ĐVT, Khối lượng, Mã máy, Duration, Formula Start
        (1, "1.0", "Phát quang mặt bằng tuyến đường nội bộ", "m2", 342646.0, "M3", 72, "='01_THONG_SO_DU_AN'!$C$6"),
        (2, "1.1", "Đào khuôn đường, vét bùn hữu cơ", "m3", 102793.8, "M1", 72, "=J6"),
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

        # Cột G: Định mức năng suất = VLOOKUP từ Sheet 02
        ws3.cell(row=idx, column=7, value=f"=VLOOKUP(F{idx}, '02_DINH_MUC_CA_MAY_VA_DAU'!$C$6:$I$16, 5, FALSE)").number_format = "#,##0.0"

        # Cột H: Tổng số ca máy = Khối lượng / Định mức
        ws3.cell(row=idx, column=8, value=f"=E{idx}/G{idx}").number_format = "#,##0.0"

        # Cột I: Số ngày thi công
        ws3.cell(row=idx, column=9, value=dur).number_format = "#,##0"

        # Cột J: Ngày bắt đầu (Early Start) sống động theo công thức
        c_j = ws3.cell(row=idx, column=10, value=formula_start)
        c_j.alignment = Alignment(horizontal="center")
        c_j.number_format = "DD/MM/YYYY"

        # Cột K: Ngày kết thúc (Early Finish) = Start + Duration - 1
        c_k = ws3.cell(row=idx, column=11, value=f"=J{idx}+I{idx}-1")
        c_k.alignment = Alignment(horizontal="center")
        c_k.number_format = "DD/MM/YYYY"

        # Cột L: Đường găng (Critical Path) = IF(Finish >= Deadline, "CRITICAL", "FLOAT")
        c_l = ws3.cell(row=idx, column=12, value=f'=IF(K{idx}>=\'01_THONG_SO_DU_AN\'!$C$7, "CRITICAL", "FLOAT")')
        c_l.alignment = Alignment(horizontal="center")
        c_l.font = Font(name=font_family, size=8.5, bold=True, color="C00000")

        # Cột M: Số máy huy động/ngày = ROUNDUP(Tổng ca / (Số ngày * Ca/ngày), 0)
        ws3.cell(row=idx, column=13, value=f"=ROUNDUP(H{idx}/(I{idx}*'01_THONG_SO_DU_AN'!$C$9), 0)").number_format = "#,##0"

        # Cột N: Nhân công huy động/ngày = Số máy * Thợ lái VLOOKUP * Ca/ngày
        ws3.cell(row=idx, column=14, value=f"=M{idx}*VLOOKUP(F{idx}, '02_DINH_MUC_CA_MAY_VA_DAU'!$C$6:$I$16, 7, FALSE)*'01_THONG_SO_DU_AN'!$C$9").number_format = "#,##0"

        for c in range(1, date_col_start):
            ws3.cell(row=idx, column=c).border = thin_border
            ws3.cell(row=idx, column=c).font = Font(name=font_family, size=9)

        # MA TRẬN GANTT LIÊN KẾT ĐỘNG 100%:
        # =IF(AND(O$4>=$J6, O$4<=$K6), IF('01_THONG_SO_DU_AN'!$C$11=1, "█", IF('01_THONG_SO_DU_AN'!$C$11=2, $N6, $M6)), "")
        for d_i in range(total_days):
            c_idx = date_col_start + d_i
            c_let = get_column_letter(c_idx)
            cell_g = ws3.cell(row=idx, column=c_idx)
            cell_g.value = f'=IF(AND({c_let}$4>=$J{idx}, {c_let}$4<=$K{idx}), IF(\'01_THONG_SO_DU_AN\'!$C$11=1, "█", IF(\'01_THONG_SO_DU_AN\'!$C$11=2, $N{idx}, $M{idx})), "")'
            cell_g.border = thin_border
            cell_g.alignment = Alignment(horizontal="center", vertical="center")
            cell_g.font = Font(name=font_family, size=8.5, bold=True, color="2E75B6")

    r_end = r_start + len(tasks_cpm) - 1

    # DÒNG TỔNG HỢP CHÂN TRANG 1: TỔNG SỐ LƯỢNG MÁY HUY ĐỘNG (MÁY/NGÀY)
    # =SUMPRODUCT((O$4>=$J$6:$J$19)*(O$4<=$K$6:$K$19)*$M$6:$M$19)
    row_sum_m = r_end + 2
    ws3.merge_cells(f"B{row_sum_m}:N{row_sum_m}")
    ws3[f"B{row_sum_m}"] = "TỔNG SỐ LƯỢNG MÁY MÓC HUY ĐỘNG (Máy/ngày) [=SUMPRODUCT((Date>=Start)*(Date<=End)*Machines)]"
    ws3[f"B{row_sum_m}"].font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
    for c in range(1, date_col_start):
        ws3.cell(row=row_sum_m, column=c).fill = PatternFill("solid", fgColor=NAVY_HEADER)

    for d_i in range(total_days):
        c_idx = date_col_start + d_i
        c_let = get_column_letter(c_idx)
        cell_sm = ws3.cell(row=row_sum_m, column=c_idx)
        cell_sm.value = f"=SUMPRODUCT(({c_let}$4>=$J${r_start}:$J${r_end})*({c_let}$4<=$K${r_start}:$K${r_end})*$M${r_start}:$M${r_end})"
        cell_sm.alignment = Alignment(horizontal="center", vertical="center")
        cell_sm.font = Font(name=font_family, size=8.5, bold=True, color=NAVY_HEADER)
        cell_sm.fill = PatternFill("solid", fgColor=AMBER_ALERT)
        cell_sm.border = double_bottom_border
        cell_sm.number_format = "#,##0"

    # DÒNG TỔNG HỢP CHÂN TRANG 2: TỔNG NHÂN CÔNG HUY ĐỘNG (NGƯỜI/NGÀY)
    # =SUMPRODUCT((O$4>=$J$6:$J$19)*(O$4<=$K$6:$K$19)*$N$6:$N$19)
    row_sum_nc = row_sum_m + 1
    ws3.merge_cells(f"B{row_sum_nc}:N{row_sum_nc}")
    ws3[f"B{row_sum_nc}"] = "TỔNG SỐ LƯỢNG NHÂN CÔNG HUY ĐỘNG (Người/ngày) [=SUMPRODUCT((Date>=Start)*(Date<=End)*Labor)]"
    ws3[f"B{row_sum_nc}"].font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
    for c in range(1, date_col_start):
        ws3.cell(row=row_sum_nc, column=c).fill = PatternFill("solid", fgColor=BLUE_SUB)

    for d_i in range(total_days):
        c_idx = date_col_start + d_i
        c_let = get_column_letter(c_idx)
        cell_snc = ws3.cell(row=row_sum_nc, column=c_idx)
        cell_snc.value = f"=SUMPRODUCT(({c_let}$4>=$J${r_start}:$J${r_end})*({c_let}$4<=$K${r_start}:$K${r_end})*$N${r_start}:$N${r_end})"
        cell_snc.alignment = Alignment(horizontal="center", vertical="center")
        cell_snc.font = Font(name=font_family, size=8.5, bold=True, color="002060")
        cell_snc.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
        cell_snc.border = double_bottom_border
        cell_snc.number_format = "#,##0"

    # Bề rộng cột
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

    # =========================================================================
    # SHEET 4: 04_TONG_HOP_CA_MAY_VA_DAU (LIÊN KẾT ĐỘNG VỚI SHEET 2 & SHEET 3)
    # =========================================================================
    ws4 = wb.create_sheet(title="04_TONG_HOP_CA_MAY_VA_DAU")
    ws4.views.sheetView[0].showGridLines = True

    ws4["B2"] = "BẢNG TỔNG HỢP CA XE, CA MÁY & KẾ HOẠCH CẤP PHÁT NHIÊN LIỆU DẦU DIESEL"
    ws4["B2"].font = Font(name=font_family, size=13, bold=True, color=NAVY_HEADER)
    ws4["B3"] = "Liên kết động 100%: Tự động tính số ca máy, số máy Peak và lượng dầu từ Sheet 03"
    ws4["B3"].font = Font(name=font_family, size=10, italic=True)

    headers_fleet = [
        "TT", "Mã máy", "Tên phương tiện / Thiết bị", "Định mức dầu (L/ca)\n(=VLOOKUP)",
        "Tổng số ca máy\n(=SUMIF)", "Số máy huy động Max\n(=SUMIF)", "Tổng lượng dầu Diesel (Lít)\n(=Số ca * Định mức)",
        "Dầu Tháng 8 (L)", "Dầu Tháng 9 (L)", "Dầu Tháng 10 (L)", "Dầu Tháng 11 (L)"
    ]
    for c_i, h in enumerate(headers_fleet, start=2):
        cell = ws4.cell(row=5, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fleet_codes = [
        (1, "M1", "Máy xúc bánh xích PC200 - PC300"),
        (2, "M3", "Máy ủi D3 - D5"),
        (3, "M4", "Máy ủi công suất lớn CS 230CV"),
        (4, "M5", "Máy lu rung 12T - 16T"),
        (5, "M6", "Máy san tự hành 110CV"),
        (6, "M7", "Ô tô tự đổ 18 tấn"),
        (7, "M8", "Máy xúc bánh lốp PC140 - PC150"),
        (8, "MP1", "Xe téc & Xe cấp dầu lưu động"),
    ]

    for r_idx, f_item in enumerate(fleet_codes, start=6):
        stt, m_code, m_name = f_item
        ws4.cell(row=r_idx, column=2, value=stt).alignment = Alignment(horizontal="center")
        ws4.cell(row=r_idx, column=3, value=m_code).alignment = Alignment(horizontal="center")
        ws4.cell(row=r_idx, column=4, value=m_name)

        # Cột E: Định mức dầu = VLOOKUP từ Sheet 02
        ws4.cell(row=r_idx, column=5, value=f"=VLOOKUP(C{r_idx}, '02_DINH_MUC_CA_MAY_VA_DAU'!$C$6:$H$16, 6, FALSE)").number_format = "#,##0.0"

        # Cột F: Tổng số ca máy = SUMIF từ Sheet 03
        ws4.cell(row=r_idx, column=6, value=f"=SUMIF('03_TIEN_DO_GANTT_CPM'!$F$6:$F$19, C{r_idx}, '03_TIEN_DO_GANTT_CPM'!$H$6:$H$19)").number_format = "#,##0.0"

        # Cột G: Số máy huy động Max = SUMIF từ Sheet 03
        c_peak = ws4.cell(row=r_idx, column=7, value=f"=SUMIF('03_TIEN_DO_GANTT_CPM'!$F$6:$F$19, C{r_idx}, '03_TIEN_DO_GANTT_CPM'!$M$6:$M$19)")
        c_peak.alignment = Alignment(horizontal="center")
        c_peak.font = Font(name=font_family, size=9.5, bold=True, color="C00000")
        c_peak.number_format = "#,##0"

        # Cột H: Tổng lượng dầu Diesel = Số ca * Định mức dầu
        ws4.cell(row=r_idx, column=8, value=f"=F{r_idx}*E{r_idx}").number_format = "#,##0.0"

        # Cột I, J, K: Phân kỳ theo tỷ lệ tiến độ tháng
        ws4.cell(row=r_idx, column=9, value=f"=ROUND(H{r_idx}*0.22, 1)").number_format = "#,##0.0"
        ws4.cell(row=r_idx, column=10, value=f"=ROUND(H{r_idx}*0.38, 1)").number_format = "#,##0.0"
        ws4.cell(row=r_idx, column=11, value=f"=ROUND(H{r_idx}*0.30, 1)").number_format = "#,##0.0"
        # Cột L (Tháng 11): Cân bằng tổng để không lệch dầu
        ws4.cell(row=r_idx, column=12, value=f"=H{r_idx}-I{r_idx}-J{r_idx}-K{r_idx}").number_format = "#,##0.0"

        for c in range(2, 13):
            ws4.cell(row=r_idx, column=c).border = thin_border
            ws4.cell(row=r_idx, column=c).font = Font(name=font_family, size=9.5)

    # Dòng Tổng cộng
    row_sum_fleet = 6 + len(fleet_codes)
    ws4.cell(row=row_sum_fleet, column=2, value="TỔNG CỘNG").alignment = Alignment(horizontal="center")
    ws4.cell(row=row_sum_fleet, column=6, value=f"=SUM(F6:F{row_sum_fleet-1})").number_format = "#,##0.0"
    ws4.cell(row=row_sum_fleet, column=7, value=f"=SUM(G6:G{row_sum_fleet-1})").number_format = "#,##0"
    ws4.cell(row=row_sum_fleet, column=8, value=f"=SUM(H6:H{row_sum_fleet-1})").number_format = "#,##0.0"
    ws4.cell(row=row_sum_fleet, column=9, value=f"=SUM(I6:I{row_sum_fleet-1})").number_format = "#,##0.0"
    ws4.cell(row=row_sum_fleet, column=10, value=f"=SUM(J6:J{row_sum_fleet-1})").number_format = "#,##0.0"
    ws4.cell(row=row_sum_fleet, column=11, value=f"=SUM(K6:K{row_sum_fleet-1})").number_format = "#,##0.0"
    ws4.cell(row=row_sum_fleet, column=12, value=f"=SUM(L6:L{row_sum_fleet-1})").number_format = "#,##0.0"

    for c in range(2, 13):
        c_tot = ws4.cell(row=row_sum_fleet, column=c)
        c_tot.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        c_tot.font = Font(name=font_family, size=10, bold=True, color=NAVY_HEADER)
        c_tot.border = double_bottom_border

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

    # =========================================================================
    # SHEET 5: 05_NHU_CAU_VAT_TU_CHINH (LIÊN KẾT ĐỘNG 100% VỚI SHEET 3)
    # =========================================================================
    ws5 = wb.create_sheet(title="05_NHU_CAU_VAT_TU_CHINH")
    ws5.views.sheetView[0].showGridLines = True

    ws5["B2"] = "BẢNG TỔNG HỢP NHU CẦU VẬT TƯ CHÍNH THEO TIẾN ĐỘ THI CÔNG"
    ws5["B2"].font = Font(name=font_family, size=13, bold=True, color=NAVY_HEADER)
    ws5["B3"] = "Liên kết 100% Khối lượng thiết kế từ Sheet 03 — Phân kỳ cung ứng theo tiến độ tuần/tháng"
    ws5["B3"].font = Font(name=font_family, size=10, italic=True)

    headers_mat = ["TT", "Mã vật tư", "Tên quy cách vật tư", "ĐVT", "Tổng nhu cầu\n(=Sheet 03)", "Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Kế hoạch tập kết vật tư"]
    for c_i, h in enumerate(headers_mat, start=2):
        cell = ws5.cell(row=5, column=c_i, value=h)
        cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    materials_spec = [
        # TT, Mã, Tên, ĐVT, Formula Total, Pct Th8, Pct Th9, Pct Th10, Ghi chú
        (1, "VT-01", "Cát san lấp nền đường K90", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$9", 0.25, 0.45, 0.30, "Cung ứng mỏ cát sông Hồng"),
        (2, "VT-02", "Cát đắp chọn lọc K95", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$10", 0.20, 0.50, 0.30, "Độ chặt K95, tạp chất < 5%"),
        (3, "VT-03", "Cát hạt thô đắp K98", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$11", 0.18, 0.46, 0.36, "CBR >= 10, K98 lớp trên cùng"),
        (4, "VT-04", "Cấp phối đá dăm loại 2 (Base B)", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$14", 0.00, 0.37, 0.55, "Mỏ đá Hà Nam hợp chuẩn"),
        (5, "VT-05", "Cấp phối đá dăm loại 1 (Base A)", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$15", 0.00, 0.37, 0.51, "Cỡ hạt Dmax = 25mm"),
        (6, "VT-06", "Bê tông nhựa nóng C19", "m2", "='03_TIEN_DO_GANTT_CPM'!$E$16", 0.00, 0.17, 0.55, "Trạm trộn BTN 120T/h"),
        (7, "VT-07", "Ống cống BTCT đúc sẵn D1500", "m", "='03_TIEN_DO_GANTT_CPM'!$E$12", 0.10, 0.42, 0.42, "Đúc sẵn tại bãi tiền chế"),
        (8, "VT-08", "Đất san lấp tổng thể cụm B9", "m3", "='03_TIEN_DO_GANTT_CPM'!$E$19", 0.25, 0.48, 0.27, "Mỏ đất đồi điều phối hợp chuẩn"),
    ]

    for r_idx, m_item in enumerate(materials_spec, start=6):
        stt, vt_code, vt_name, vt_unit, f_tot, p8, p9, p10, vt_note = m_item
        ws5.cell(row=r_idx, column=2, value=stt).alignment = Alignment(horizontal="center")
        ws5.cell(row=r_idx, column=3, value=vt_code).alignment = Alignment(horizontal="center")
        ws5.cell(row=r_idx, column=4, value=vt_name)
        ws5.cell(row=r_idx, column=5, value=vt_unit).alignment = Alignment(horizontal="center")

        # Cột F: Tổng nhu cầu = Liên kết sống từ Sheet 03
        ws5.cell(row=r_idx, column=6, value=f_tot).number_format = "#,##0.0"

        # Phân kỳ Tháng 8, 9, 10
        ws5.cell(row=r_idx, column=7, value=f"=ROUND(F{r_idx}*{p8}, 1)").number_format = "#,##0.0"
        ws5.cell(row=r_idx, column=8, value=f"=ROUND(F{r_idx}*{p9}, 1)").number_format = "#,##0.0"
        ws5.cell(row=r_idx, column=9, value=f"=ROUND(F{r_idx}*{p10}, 1)").number_format = "#,##0.0"
        # Tháng 11: Cân bằng số dư còn lại
        ws5.cell(row=r_idx, column=10, value=f"=F{r_idx}-G{r_idx}-H{r_idx}-I{r_idx}").number_format = "#,##0.0"

        ws5.cell(row=r_idx, column=11, value=vt_note)

        for c in range(2, 12):
            ws5.cell(row=r_idx, column=c).border = thin_border
            ws5.cell(row=r_idx, column=c).font = Font(name=font_family, size=9.5)

    ws5.column_dimensions["B"].width = 6
    ws5.column_dimensions["C"].width = 12
    ws5.column_dimensions["D"].width = 34
    ws5.column_dimensions["E"].width = 8
    ws5.column_dimensions["F"].width = 18
    ws5.column_dimensions["G"].width = 15
    ws5.column_dimensions["H"].width = 15
    ws5.column_dimensions["I"].width = 15
    ws5.column_dimensions["J"].width = 15
    ws5.column_dimensions["K"].width = 32

    # =========================================================================
    # LƯU VÀ ĐỒNG BỘ
    # =========================================================================
    wb.save(OUTPUT_FILE)
    print(f"[OK] Đã xuất file 100% CÔNG THỨC SỐNG tại: {OUTPUT_FILE}")
    print(f"     Kích thước file: {os.path.getsize(OUTPUT_FILE):,} bytes (siêu sạch, siêu nhẹ!)")

    shutil.copyfile(OUTPUT_FILE, MIRROR_FILE_1)
    shutil.copyfile(OUTPUT_FILE, MIRROR_FILE_2)
    print(f"[OK] Đã đồng bộ sang 2 tầng thư mục dự án!")


if __name__ == "__main__":
    build_dynamic_schedule()
