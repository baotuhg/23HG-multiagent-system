# -*- coding: utf-8 -*-
"""
TỰ ĐỘNG HÓA TÍNH TOÁN CA XE, CA MÁY & TIẾN ĐỘ THI CÔNG CỐNG HỘP TUYẾN A5
Áp dụng định mức Vincons_ĐMGK_02-01 theo khung tiến độ 20/09/2026 - 25/10/2026 (36 ngày)
Dự án: Cống hộp Tuyến A5 - Khu đô thị Thể thao Quốc tế Hà Nội (Vincons)
Chiều dài: L = 2,170 m (192 đốt 11.3m, 63 hố ga)
"""

import os
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_a5_machine_schedule():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    font_family = "Arial"
    NAVY_HEADER = "1B365D"
    BLUE_SUB = "2E75B6"
    LIGHT_BLUE = "D9E1F2"
    ICE_BLUE = "EDF2F8"
    LIGHT_GREEN = "E2EFDA"
    ACCENT_GREEN = "385723"
    AMBER_SUM = "FFF2CC"
    ORANGE_CRIT = "C65911"
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

    # 36 dates from 2026-09-20 to 2026-10-25
    start_date = datetime.date(2026, 9, 20)
    end_date = datetime.date(2026, 10, 25)
    total_days = (end_date - start_date).days + 1  # 36
    dates = [start_date + datetime.timedelta(days=i) for i in range(total_days)]
    weekday_vn = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

    # =========================================================================
    # SHEET 1: 01_TienDo_CaMay_Master
    # =========================================================================
    ws1 = wb.create_sheet(title="01_TienDo_CaMay_Master")
    ws1.views.sheetView[0].showGridLines = True

    # Title
    ws1.merge_cells("A1:N1")
    ws1["A1"] = "DỰ ÁN: CỐNG HỘP TUYẾN A5 - KHU ĐÔ THỊ THỂ THAO QUỐC TẾ HÀ NỘI"
    ws1["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws1["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A2:N2")
    ws1["A2"] = "BẢNG TÍNH TOÁN CA XE, CA MÁY & TIẾN ĐỘ THI CÔNG TOÀN TUYẾN A5 (L = 2,170 M - 63 HỐ GA)"
    ws1["A2"].font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A3:N3")
    ws1["A3"] = f"MỐC TIẾN ĐỘ THẦN TỐC: TỪ {start_date.strftime('%d/%m/%Y')} ĐẾN {end_date.strftime('%d/%m/%Y')} (36 NGÀY) - 2 CA/NGÀY (20H/NGÀY) - 3 MŨI THI CÔNG ĐỒNG THỜI"
    ws1["A3"].font = Font(name=font_family, size=10, bold=True, color="C00000")
    ws1["A3"].fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
    ws1["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Key parameters row
    ws1["B4"] = "Chế độ ca:"
    ws1["C4"] = "2 ca/ngày (20 giờ/ngày)"
    ws1["B4"].font = Font(name=font_family, size=9.5, bold=True)
    ws1["C4"].font = Font(name=font_family, size=9.5, italic=True)

    ws1["F4"] = "Phân đoạn thi công:"
    ws1["G4"] = "3 mũi song song (Cống 2x2: 385m, Cống 3x3: 560m, Cống đôi 2x(3x3): 1,225m)"
    ws1["F4"].font = Font(name=font_family, size=9.5, bold=True)
    ws1["G4"].font = Font(name=font_family, size=9.5, color="1B365D")

    ws1["K4"] = "Định mức áp dụng:"
    ws1["L4"] = "Vincons_ĐMGK_02-01 Hạ tầng giao thông"
    ws1["K4"].font = Font(name=font_family, size=9.5, bold=True)
    ws1["L4"].font = Font(name=font_family, size=9.5, color="008000")

    # Table Headers
    base_headers = [
        ("STT", "A6:A7"),
        ("Mã ĐM", "B6:B7"),
        ("Nội dung công việc thi công cống hộp", "C6:C7"),
        ("ĐVT", "D6:D7"),
        ("Khối lượng thiết kế", "E6:E7"),
        ("Định mức Vincons (ĐVT/ca)", "F6:F7"),
        ("Tổng số ca máy (ca)", "G6:G7"),
        ("Năng xuất ngày", "H6:H7"),
        ("Thời gian (ngày)", "I6:I7"),
        ("Ngày BĐ", "J6:J7"),
        ("Ngày KT", "K6:K7"),
        ("Số ca/ngày", "L6:L7"),
        ("Số máy huy động/ngày", "M6:M7"),
        ("Chủng loại MMTB & Ghi chú", "N6:N7"),
        ("NC bố trí (người)", "O6:O7")
    ]

    for title, rng in base_headers:
        ws1.merge_cells(rng)
        top_cell = rng.split(":")[0]
        ws1[top_cell] = title
        ws1[top_cell].font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
        ws1[top_cell].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws1[top_cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Date headers starting at column P (column 16)
    date_start_col = 16
    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        ws1.cell(row=6, column=c_idx, value=d.strftime("%d/%m"))
        ws1.cell(row=6, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
        ws1.cell(row=6, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")

        weekday_str = weekday_vn[d.weekday()]
        ws1.cell(row=7, column=c_idx, value=weekday_str)
        ws1.cell(row=7, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")

        if d.weekday() == 6:  # Sunday
            ws1.cell(row=6, column=c_idx).fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            ws1.cell(row=7, column=c_idx).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            ws1.cell(row=7, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="C00000")
        else:
            ws1.cell(row=6, column=c_idx).fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
            ws1.cell(row=7, column=c_idx).fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
            ws1.cell(row=7, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="1B365D")

    # Tasks definition:
    # (stt, code, name, unit, qty, norm_shift, s_str, e_str, shifts_day, m_equip, nc_day, oil_norm, mach_type)
    tasks = [
        (1, "ĐM-16", "Đào đất hố móng cống hộp 3 mũi (H_tb = 3.5m)", "m3", 73800.0, 401.07, "2026-09-20", "2026-10-05", 2, "Máy xúc PC200 - PC300", 24, 89.6, "Xúc đào PC200"),
        (2, "ĐM-15", "Đào hố móng 63 hố ga thu nước & đấu nối", "m3", 12650.0, 313.55, "2026-09-22", "2026-10-03", 2, "Máy xúc PC200", 8, 89.6, "Xúc đào PC200"),
        (3, "ĐM-29", "Vận chuyển đất đào cự ly <5km (bãi thải/đắp)", "m3", 86450.0, 571.43, "2026-09-20", "2026-10-05", 2, "Ô tô tự đổ 18 m3", 10, 26.0, "Ô tô 18m3"),
        (4, "ĐM-BOM", "Bơm hạ mực nước ngầm & chống ngập móng", "ca", 72.0, 1.0, "2026-09-20", "2026-10-25", 2, "Máy bơm nước d80-d100", 4, 14.0, "Máy bơm nước"),
        (5, "ĐM-53", "Đệm cát / đá dăm 4x6 lót móng cống", "m3", 1250.0, 150.0, "2026-09-22", "2026-10-05", 2, "Máy xúc hỗ trợ + Lu cóc", 12, 29.0, "Xúc lốp PC55"),
        (6, "ĐM-37", "Bê tông lót móng M100# dày 50mm", "m3", 592.4, 45.0, "2026-09-23", "2026-10-08", 2, "Xe bồn xả máng + Đầm bàn", 16, 35.0, "Xe bồn BT"),
        (7, "ĐM-CT", "Gia công & lắp dựng cốt thép cống hộp B20", "tấn", 753.36, 12.0, "2026-09-24", "2026-10-15", 2, "Dàn cắt uốn CNC + Cẩu 25T", 45, 45.0, "Cần cẩu 25T"),
        (8, "ĐM-VK", "Lắp dựng ván khuôn thép/phủ phim thân cống", "m2", 44940.7, 450.0, "2026-09-25", "2026-10-16", 2, "Cẩu thùng 15-25T cẩu lắp", 60, 28.0, "Cần cẩu 25T"),
        (9, "ĐM-BT", "Đổ bê tông thân cống B20 (đáy, vách, nắp)", "m3", 11870.0, 180.0, "2026-09-26", "2026-10-17", 2, "Máy bơm bê tông cần 37-43m", 36, 65.0, "Bơm bê tông"),
        (10, "ĐM-VCBT", "Vận chuyển bê tông thương phẩm B20 trạm trộn", "m3", 11870.0, 45.0, "2026-09-26", "2026-10-17", 2, "Xe bồn vận chuyển 8-10m3", 12, 42.0, "Xe bồn BT"),
        (11, "ĐM-46", "Bê tông & cốt thép 63 hố ga BTCT", "m3", 409.5, 30.0, "2026-09-28", "2026-10-12", 2, "Máy bơm bê tông tĩnh", 18, 45.0, "Bơm bê tông"),
        (12, "ĐM-25", "Máy xúc lốp PC140 cẩu lắp & đầm cóc mang cống", "md", 2170.0, 100.0, "2026-09-26", "2026-10-17", 2, "Máy xúc lốp PC140", 8, 64.0, "Xúc lốp PC140"),
        (13, "ĐM-17", "Đắp cát/đất K95 hoàn trả mang cống", "m3", 54200.0, 372.0, "2026-10-05", "2026-10-22", 2, "Máy xúc PC200 đắp mang cống", 24, 89.6, "Xúc đào PC200"),
        (14, "ĐM-ỦI", "Máy ủi D3-D5 san gạt hoàn trả đỉnh móng", "m3", 54200.0, 1428.0, "2026-10-05", "2026-10-22", 2, "Máy ủi bánh xích D3-D5", 6, 124.0, "Máy ủi D3-D5"),
        (15, "ĐM-LU", "Máy lu rung 12-16T đầm nén K95 hoàn trả", "m3", 54200.0, 255.0, "2026-10-05", "2026-10-22", 2, "Máy lu rung bánh thép 14T", 12, 44.0, "Máy lu 14T"),
        (16, "ĐM-DAU", "Xe téc cấp dầu lưu động 9m3 phục vụ máy móc", "ca", 72.0, 1.0, "2026-09-20", "2026-10-25", 2, "Xe téc cấp dầu 9m3", 2, 44.0, "Xe cấp dầu"),
        (17, "ĐM-ĐIỆN", "Máy phát điện 3 pha công nghiệp 25-45kVA", "ca", 72.0, 1.0, "2026-09-20", "2026-10-25", 2, "Máy phát điện 3 pha", 4, 28.0, "Máy phát điện")
    ]

    curr_row = 8
    task_rows = []

    for t in tasks:
        stt, code, name, unit, qty, norm_s, s_str, e_str, shifts, m_note, nc_d, oil_n, m_type = t
        s_d = datetime.date.fromisoformat(s_str)
        e_d = datetime.date.fromisoformat(e_str)
        dur = (e_d - s_d).days + 1
        total_shifts = round(qty / norm_s, 1)
        daily_prod = round(qty / dur, 1)
        m_req = round(total_shifts / (dur * shifts), 2)

        ws1.cell(row=curr_row, column=1, value=stt)
        ws1.cell(row=curr_row, column=2, value=code)
        ws1.cell(row=curr_row, column=3, value=name)
        ws1.cell(row=curr_row, column=4, value=unit)
        ws1.cell(row=curr_row, column=5, value=qty)
        ws1.cell(row=curr_row, column=6, value=norm_s)
        ws1.cell(row=curr_row, column=7, value=f"=E{curr_row}/F{curr_row}")
        ws1.cell(row=curr_row, column=8, value=f"=E{curr_row}/I{curr_row}")
        ws1.cell(row=curr_row, column=9, value=dur)
        ws1.cell(row=curr_row, column=10, value=s_str)
        ws1.cell(row=curr_row, column=11, value=e_str)
        ws1.cell(row=curr_row, column=12, value=shifts)
        ws1.cell(row=curr_row, column=13, value=f"=G{curr_row}/(I{curr_row}*L{curr_row})")
        ws1.cell(row=curr_row, column=14, value=m_note)
        ws1.cell(row=curr_row, column=15, value=nc_d)

        # Style left columns
        is_concrete = "thân cống B20" in name or "cốt thép" in name
        row_fill = PatternFill(start_color=PEACH_FILL, end_color=PEACH_FILL, fill_type="solid") if is_concrete else None

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5, bold=is_concrete)
            cell.border = thin_border
            if row_fill: cell.fill = row_fill

        # Date columns: display machine count on active days
        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if s_d <= d <= e_d:
                cell.value = m_req  # Show daily machines required
                if is_concrete:
                    cell.fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
                    cell.font = Font(name=font_family, size=8, bold=True, color="C00000")
                else:
                    cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                    cell.font = Font(name=font_family, size=8, color="1B365D")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = 0
                cell.font = Font(name=font_family, size=7, color="BFBFBF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

        task_rows.append((curr_row, m_type, oil_n, nc_d, s_d, e_d, m_req))
        curr_row += 1

    curr_row += 1

    # =========================================================================
    # SUMMARY 1: CỘNG NHÂN CÔNG THEO NGÀY
    # =========================================================================
    row_nc_sum = curr_row
    ws1.cell(row=row_nc_sum, column=3, value="TỔNG NHÂN CÔNG TRÊN CÔNG TRƯỜNG TUYẾN A5 (Người/ngày)")
    ws1.cell(row=row_nc_sum, column=3).font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    for c in range(1, date_start_col):
        ws1.cell(row=row_nc_sum, column=c).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")

    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        daily_nc = sum(nc for r_idx, m_t, oil_n, nc, s_d, e_d, m_req in task_rows if s_d <= d <= e_d)
        cell = ws1.cell(row=row_nc_sum, column=c_idx, value=daily_nc)
        cell.font = Font(name=font_family, size=8.5, bold=True, color="002060")
        cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        cell.border = thick_bottom
        cell.alignment = Alignment(horizontal="center", vertical="center")
    curr_row += 2

    # =========================================================================
    # SUMMARY 2: BẢNG TỔNG HỢP CA MÁY & SỐ PHƯƠNG TIỆN HUY ĐỘNG
    # =========================================================================
    ws1.cell(row=curr_row, column=3, value="BẢNG TỔNG HỢP CA MÁY & SỐ PHƯƠNG TIỆN HUY ĐỘNG THEO NGÀY (CÔNG TRƯỜNG A5)")
    ws1.cell(row=curr_row, column=3).font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.merge_cells(f"C{curr_row}:M{curr_row}")
    for c in range(3, date_start_col):
        ws1.cell(row=curr_row, column=c).fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    curr_row += 1

    # Machinery categories
    mach_categories = [
        ("M1", "Máy xúc bánh xích PC200 - PC300", "cái", 89.6, 8, "Đào móng, đào ga, xúc mang cống"),
        ("M7", "Ô tô tự đổ 18 m3", "xe", 26.0, 5, "Vận chuyển đất đào cự ly <5km"),
        ("M5", "Máy lu rung 12T - 16T", "cái", 44.0, 6, "Đầm nén K95 mang cống và đỉnh cống"),
        ("M3", "Máy ủi bánh xích D3 - D5", "cái", 124.0, 2, "San gạt hoàn trả hố móng cống"),
        ("MC", "Cần cẩu 25T & Cẩu tự hành 15T", "cái", 28.0, 3, "Hạ ván khuôn, cẩu lắp cốt thép"),
        ("MB", "Máy bơm bê tông cần 37 - 43m", "cái", 65.0, 2, "Đổ bê tông thân cống B20 & hố ga"),
        ("XB", "Xe bồn vận chuyển bê tông 8-10m3", "xe", 42.0, 6, "Chở bê tông từ trạm trộn về công trường"),
        ("M8", "Máy xúc bánh lốp PC140", "cái", 64.0, 1, "Cẩu lắp, đầm cóc mang cống"),
        ("BP", "Máy bơm nước hố móng d80-d100", "cái", 14.0, 2, "Bơm hạ mực nước ngầm 24/7"),
        ("MP1", "Xe téc cấp dầu lưu động 9m3", "xe", 44.0, 1, "Cấp dầu lưu động 2 ca/ngày"),
        ("MP2", "Máy phát điện 3 pha 25-45kVA", "cái", 28.0, 2, "Chiếu sáng ban đêm 2 ca")
    ]

    ws1.cell(row=curr_row, column=1, value="Mã")
    ws1.cell(row=curr_row, column=3, value="Chủng loại phương tiện / Thiết bị")
    ws1.cell(row=curr_row, column=4, value="ĐVT")
    ws1.cell(row=curr_row, column=5, value="ĐM dầu (l/ca)")
    ws1.cell(row=curr_row, column=6, value="Max máy")
    ws1.cell(row=curr_row, column=7, value="Nhiệm vụ thi công trên tuyến A5")
    for c in range(1, date_start_col):
        cell = ws1.cell(row=curr_row, column=c)
        cell.font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        cell = ws1.cell(row=curr_row, column=c_idx, value=d.strftime("%d/%m"))
        cell.font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    curr_row += 1

    mach_summary_rows = []
    for code, m_name, u, oil_rate, max_m, task_desc in mach_categories:
        ws1.cell(row=curr_row, column=1, value=code)
        ws1.cell(row=curr_row, column=3, value=m_name)
        ws1.cell(row=curr_row, column=4, value=u)
        ws1.cell(row=curr_row, column=5, value=oil_rate)
        ws1.cell(row=curr_row, column=6, value=max_m)
        ws1.cell(row=curr_row, column=7, value=task_desc)

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5)
            cell.border = thin_border

        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            
            # Realistic leveling
            val = 0
            if code == "M1":
                if d <= datetime.date(2026, 10, 5): val = 8
                elif d <= datetime.date(2026, 10, 22): val = 4
            elif code == "M7":
                if d <= datetime.date(2026, 10, 5): val = 5
            elif code == "M5":
                if datetime.date(2026, 10, 5) <= d <= datetime.date(2026, 10, 22): val = 6
            elif code == "M3":
                if datetime.date(2026, 10, 5) <= d <= datetime.date(2026, 10, 22): val = 2
            elif code == "MC":
                if datetime.date(2026, 9, 24) <= d <= datetime.date(2026, 10, 16): val = 3
            elif code == "MB":
                if datetime.date(2026, 9, 26) <= d <= datetime.date(2026, 10, 17): val = 2
            elif code == "XB":
                if datetime.date(2026, 9, 26) <= d <= datetime.date(2026, 10, 17): val = 6
            elif code == "M8":
                if datetime.date(2026, 9, 26) <= d <= datetime.date(2026, 10, 17): val = 1
            elif code == "BP":
                val = 2
            elif code == "MP1":
                val = 1
            elif code == "MP2":
                val = 2

            cell.value = val
            cell.font = Font(name=font_family, size=8, bold=(val > 0), color=("1B365D" if val > 0 else "BFBFBF"))
            if val > 0:
                cell.fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        mach_summary_rows.append((curr_row, code, oil_rate))
        curr_row += 1

    curr_row += 1

    # =========================================================================
    # SUMMARY 3: BẢNG TÍNH LƯỢNG DẦU DIEZEL TIÊU THỤ HÀNG NGÀY
    # =========================================================================
    ws1.cell(row=curr_row, column=3, value="BẢNG TÍNH DẦU DIEZEL TIÊU THỤ THEO TIẾN ĐỘ THI CÔNG TUYẾN A5 (Lít/ngày)")
    ws1.cell(row=curr_row, column=3).font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.merge_cells(f"C{curr_row}:M{curr_row}")
    for c in range(3, date_start_col):
        ws1.cell(row=curr_row, column=c).fill = PatternFill(start_color=ACCENT_GREEN, end_color=ACCENT_GREEN, fill_type="solid")
    curr_row += 1

    row_total_oil = curr_row
    ws1.cell(row=row_total_oil, column=3, value="TỔNG SỐ LÍT DẦU DIEZEL TIÊU THỤ / NGÀY")
    ws1.cell(row=row_total_oil, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in range(1, date_start_col):
        ws1.cell(row=row_total_oil, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")

    oil_calc_rows = []
    curr_row += 1
    for m_row, code, oil_rate in mach_summary_rows:
        ws1.cell(row=curr_row, column=1, value=code)
        ws1.cell(row=curr_row, column=3, value=f"Nhiên liệu dầu Diezel cho máy {code}")
        ws1.cell(row=curr_row, column=4, value="Lít")
        ws1.cell(row=curr_row, column=5, value=oil_rate)

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8)
            cell.border = thin_border

        for idx in range(total_days):
            c_idx = date_start_col + idx
            c_let = get_column_letter(c_idx)
            cell = ws1.cell(row=curr_row, column=c_idx)
            # Fuel = Machine count * oil_rate * 2 shifts
            cell.value = f"={c_let}{m_row}*$E{curr_row}*2"
            cell.font = Font(name=font_family, size=7.5, color="333333")
            cell.border = thin_border
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")

        oil_calc_rows.append(curr_row)
        curr_row += 1

    # Formula in row_total_oil
    for idx in range(total_days):
        c_idx = date_start_col + idx
        c_let = get_column_letter(c_idx)
        cell = ws1.cell(row=row_total_oil, column=c_idx)
        cell.value = f"=SUM({c_let}{oil_calc_rows[0]}:{c_let}{oil_calc_rows[-1]})"
        cell.font = Font(name=font_family, size=8.5, bold=True, color="C00000")
        cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        cell.border = thick_bottom
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # Column widths for Sheet 1
    col_widths = {
        "A": 6, "B": 10, "C": 42, "D": 8, "E": 14, "F": 14, "G": 14,
        "H": 14, "I": 10, "J": 12, "K": 12, "L": 10, "M": 14, "N": 30, "O": 12
    }
    for col_let, w in col_widths.items():
        ws1.column_dimensions[col_let].width = w

    for idx in range(total_days):
        c_let = get_column_letter(date_start_col + idx)
        ws1.column_dimensions[c_let].width = 6.8

    # Formats
    for r in range(8, curr_row):
        ws1[f"E{r}"].number_format = "#,##0.0"
        ws1[f"F{r}"].number_format = "#,##0.0"
        ws1[f"G{r}"].number_format = "#,##0.0"
        ws1[f"H{r}"].number_format = "#,##0.0"
        ws1[f"M{r}"].number_format = "#,##0.00"
        ws1[f"O{r}"].number_format = "#,##0"


    # =========================================================================
    # SHEET 2: 02_TongHop_CaXe_CaMay_MMTB
    # =========================================================================
    ws2 = wb.create_sheet(title="02_TongHop_CaXe_CaMay_MMTB")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:J1")
    ws2["A1"] = "BẢNG TỔNG HỢP CA XE, CA MÁY THI CÔNG TOÀN TUYẾN CỐNG HỘP A5"
    ws2["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws2 = [
        "STT", "Mã máy", "Tên phương tiện / Thiết bị thi công", "ĐVT",
        "Định mức dầu (lít/ca)", "Tổng số ca máy (ca)", "Số máy huy động Max",
        "Số ngày làm việc", "Tổng lít dầu tiêu thụ (lít)", "Nhiệm vụ thi công trên công trường A5"
    ]
    for c_i, h in enumerate(headers_ws2, 1):
        cell = ws2.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    summary_mach_data = [
        (1, "M1", "Máy xúc bánh xích PC200 - PC300", "cái", 89.6, 224.3, 8, 16, 20100, "Đào hố móng cống, đào hố ga, xúc đắp hoàn trả K95"),
        (2, "M7", "Ô tô tự đổ 16 - 18 tấn", "xe", 26.0, 151.3, 5, 16, 3934, "Vận chuyển đất đào cự ly <5km ra bãi thải/bãi đắp"),
        (3, "M5", "Máy lu rung bánh thép & lốp 12T - 16T", "cái", 44.0, 212.5, 6, 18, 9352, "Đầm nén cát/đất K95 hoàn trả mang cống và đỉnh cống"),
        (4, "M3", "Máy ủi bánh xích D3 - D5", "cái", 124.0, 38.0, 2, 18, 4706, "San gạt đất hoàn trả hố móng cống, tạo mặt bằng"),
        (5, "MC", "Cần cẩu 25T & Cẩu tự hành 15T", "cái", 28.0, 162.7, 3, 22, 5621, "Cẩu lắp ván khuôn cống, hạ cốt thép, hỗ trợ đúc cống"),
        (6, "MB", "Máy bơm bê tông cần 37 - 43m", "cái", 65.0, 79.6, 2, 22, 4900, "Bơm bê tông thân cống B20 (đáy, vách, nắp) & hố ga"),
        (7, "XB", "Xe bồn vận chuyển bê tông thương phẩm 8-10m3", "xe", 42.0, 263.8, 6, 22, 11079, "Chở bê tông từ trạm trộn thương phẩm đến công trường"),
        (8, "M8", "Máy xúc bánh lốp PC140", "cái", 64.0, 21.7, 1, 22, 1389, "Trung chuyển vật tư, cẩu lắp và đầm cóc mang cống"),
        (9, "BP", "Máy bơm nước hố móng d80-d100", "cái", 14.0, 72.0, 2, 36, 1008, "Bơm hút hạ mực nước ngầm, chống ngập hố móng 24/7"),
        (10, "MP1", "Xe téc cấp dầu lưu động 9m3", "xe", 44.0, 72.0, 1, 36, 3168, "Cấp phát dầu Diezel trực tiếp cho máy móc 2 ca/ngày"),
        (11, "MP2", "Máy phát điện 3 pha công nghiệp 25-45kVA", "cái", 28.0, 72.0, 2, 36, 2016, "Cấp điện thi công ca đêm & vận hành máy đầm, máy bơm")
    ]

    r_s = 4
    for stt, code, name, u, oil_rate, shifts, max_m, days, total_oil, task_desc in summary_mach_data:
        ws2.cell(row=r_s, column=1, value=stt)
        ws2.cell(row=r_s, column=2, value=code)
        ws2.cell(row=r_s, column=3, value=name)
        ws2.cell(row=r_s, column=4, value=u)
        ws2.cell(row=r_s, column=5, value=oil_rate)
        ws2.cell(row=r_s, column=6, value=shifts)
        ws2.cell(row=r_s, column=7, value=max_m)
        ws2.cell(row=r_s, column=8, value=days)
        ws2.cell(row=r_s, column=9, value=f"=F{r_s}*E{r_s}")
        ws2.cell(row=r_s, column=10, value=task_desc)

        for c in range(1, 11):
            cell = ws2.cell(row=r_s, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 2, 4]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [5, 6, 7, 8, 9]:
                cell.number_format = "#,##0.0" if c in [5, 6] else "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_s += 1

    # Total row
    ws2.cell(row=r_s, column=3, value="TỔNG CỘNG TOÀN CÔNG TRÌNH CỐNG HỘP A5")
    ws2.cell(row=r_s, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
    ws2.cell(row=r_s, column=6, value=f"=SUM(F4:F{r_s-1})")
    ws2.cell(row=r_s, column=7, value=f"=SUM(G4:G{r_s-1})")
    ws2.cell(row=r_s, column=9, value=f"=SUM(I4:I{r_s-1})")
    for c in [6, 7, 9]:
        cell = ws2.cell(row=r_s, column=c)
        cell.font = Font(name=font_family, size=10, bold=True, color="C00000")
        cell.number_format = "#,##0.0" if c == 6 else "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 11):
        ws2.cell(row=r_s, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws2.cell(row=r_s, column=c).border = thick_bottom

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 44
    ws2.column_dimensions["D"].width = 8
    ws2.column_dimensions["E"].width = 16
    ws2.column_dimensions["F"].width = 18
    ws2.column_dimensions["G"].width = 18
    ws2.column_dimensions["H"].width = 16
    ws2.column_dimensions["I"].width = 22
    ws2.column_dimensions["J"].width = 46


    # =========================================================================
    # SHEET 3: 03_KeHoach_Dau_Diezel
    # =========================================================================
    ws3 = wb.create_sheet(title="03_KeHoach_Dau_Diezel")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:H1")
    ws3["A1"] = "KẾ HOẠCH CẤP DẦU DIEZEL CHO MÁY MÓC THI CÔNG TUYẾN A5 (THEO 4 KỲ)"
    ws3["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws3["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws3 = [
        "STT", "Chủng loại thiết bị", "Định mức (lít/ca)", "Tổng số ca máy",
        "Tổng nhu cầu dầu (Lít)", "Kỳ 1 (20/9 - 25/9)", "Kỳ 2 (26/9 - 05/10)", "Kỳ 3 (06/10 - 15/10)", "Kỳ 4 (16/10 - 25/10)"
    ]
    for c_i, h in enumerate(headers_ws3, 1):
        cell = ws3.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    oil_plan = [
        (1, "Máy xúc bánh xích PC200 - PC300", 89.6, 224.3, 20100, 0.25, 0.40, 0.25, 0.10),
        (2, "Ô tô tự đổ 16 - 18 tấn", 26.0, 151.3, 3934, 0.35, 0.50, 0.15, 0.00),
        (3, "Máy lu rung 12T - 16T", 44.0, 212.5, 9352, 0.00, 0.10, 0.50, 0.40),
        (4, "Máy ủi bánh xích D3 - D5", 124.0, 38.0, 4706, 0.00, 0.10, 0.50, 0.40),
        (5, "Cần cẩu 25T & Cẩu tự hành 15T", 28.0, 162.7, 5621, 0.10, 0.45, 0.35, 0.10),
        (6, "Máy bơm bê tông cần 37 - 43m", 65.0, 79.6, 4900, 0.05, 0.45, 0.40, 0.10),
        (7, "Xe bồn vận chuyển bê tông 8-10m3", 42.0, 263.8, 11079, 0.05, 0.45, 0.40, 0.10),
        (8, "Máy xúc bánh lốp PC140", 64.0, 21.7, 1389, 0.00, 0.45, 0.45, 0.10),
        (9, "Máy bơm nước hố móng", 14.0, 72.0, 1008, 0.20, 0.30, 0.30, 0.20),
        (10, "Xe téc cấp dầu lưu động 9m3", 44.0, 72.0, 3168, 0.20, 0.30, 0.30, 0.20),
        (11, "Máy phát điện 3 pha 25-45kVA", 28.0, 72.0, 2016, 0.20, 0.30, 0.30, 0.20)
    ]

    r_op = 4
    for stt, name, dm_oil, shifts, total_oil, k1, k2, k3, k4 in oil_plan:
        ws3.cell(row=r_op, column=1, value=stt)
        ws3.cell(row=r_op, column=2, value=name)
        ws3.cell(row=r_op, column=3, value=dm_oil)
        ws3.cell(row=r_op, column=4, value=shifts)
        ws3.cell(row=r_op, column=5, value=f"=C{r_op}*D{r_op}")
        ws3.cell(row=r_op, column=6, value=f"=E{r_op}*{k1}")
        ws3.cell(row=r_op, column=7, value=f"=E{r_op}*{k2}")
        ws3.cell(row=r_op, column=8, value=f"=E{r_op}*{k3}")
        ws3.cell(row=r_op, column=9, value=f"=E{r_op}*{k4}")

        for c in range(1, 10):
            cell = ws3.cell(row=r_op, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c == 1: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [3, 4, 5, 6, 7, 8, 9]:
                cell.number_format = "#,##0.0" if c in [3, 4] else "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
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

    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 44
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 16
    ws3.column_dimensions["E"].width = 20
    ws3.column_dimensions["F"].width = 18
    ws3.column_dimensions["G"].width = 18
    ws3.column_dimensions["H"].width = 18
    ws3.column_dimensions["I"].width = 18


    # =========================================================================
    # SHEET 4: 04_KeHoach_NhanLuc_3_Mui
    # =========================================================================
    ws4 = wb.create_sheet(title="04_KeHoach_NhanLuc_3_Mui")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:H1")
    ws4["A1"] = "BẢNG PHÂN BỔ NHÂN LỰC THI CÔNG 3 MŨI CỐNG HỘP TUYẾN A5"
    ws4["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws4["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws4 = [
        "STT", "Tổ đội / Bộ phận thi công", "Mũi 1 (Cống 2x2, L=385m)",
        "Mũi 2 (Cống 3x3, L=560m)", "Mũi 3 (Cống đôi 2x(3x3), L=1225m)", "Tổ cơ giới & Phục vụ", "Tổng nhân lực (người)", "Ghi chú bố trí ca kíp"
    ]
    for c_i, h in enumerate(headers_ws4, 1):
        cell = ws4.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    labor_allocation = [
        (1, "Thợ gia công & lắp dựng cốt thép CB300/CB400", 12, 18, 30, 0, "Gia công tại bãi + buộc lồng thép hố móng"),
        (2, "Thợ lắp dựng & tháo dỡ ván khuôn thép/phim", 14, 20, 36, 0, "Ghép ván khuôn đáy, vách, bản nắp cống"),
        (3, "Công nhân đổ bê tông & đầm dùi/đầm bàn B20", 8, 12, 20, 0, "Đổ bê tông theo từng đốt 11.3m"),
        (4, "Thợ thi công 63 hố ga thu nước & đấu nối", 4, 6, 12, 0, "Đúc ga, lắp nắp ga thu nước"),
        (5, "Công nhân đắp đất/cát hoàn trả mang cống K95", 6, 8, 14, 0, "Đầm cóc mang cống & lu lèn"),
        (6, "Lái máy xúc bánh xích PC200 - PC300", 0, 0, 0, 16, "8 máy xúc x 2 tài/máy (2 ca)"),
        (7, "Lái ô tô tự đổ 18 m3", 0, 0, 0, 10, "5 xe x 2 tài/xe"),
        (8, "Lái máy lu rung 12-16T", 0, 0, 0, 12, "6 máy lu x 2 tài/máy"),
        (9, "Lái cần cẩu 25T & Cẩu tự hành 15T", 0, 0, 0, 6, "3 cẩu x 2 tài/cẩu"),
        (10, "Lái máy ủi, xúc lốp, xe cấp dầu, máy phát điện", 0, 0, 0, 10, "Vận hành thiết bị phụ trợ 2 ca"),
        (11, "Chỉ huy trưởng, kỹ sư hiện trường, KCS/QA/QC", 3, 3, 5, 5, "Giám sát kỹ thuật 24/7")
    ]

    r_la = 4
    for stt, dept, m1, m2, m3, m4, gc in labor_allocation:
        ws4.cell(row=r_la, column=1, value=stt)
        ws4.cell(row=r_la, column=2, value=dept)
        ws4.cell(row=r_la, column=3, value=m1)
        ws4.cell(row=r_la, column=4, value=m2)
        ws4.cell(row=r_la, column=5, value=m3)
        ws4.cell(row=r_la, column=6, value=m4)
        ws4.cell(row=r_la, column=7, value=f"=SUM(C{r_la}:F{r_la})")
        ws4.cell(row=r_la, column=8, value=gc)

        for c in range(1, 9):
            cell = ws4.cell(row=r_la, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c == 1: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [3, 4, 5, 6, 7]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_la += 1

    # Total labor
    ws4.cell(row=r_la, column=2, value="TỔNG NHÂN LỰC TOÀN CÔNG TRƯỜNG A5 (NGƯỜI)")
    ws4.cell(row=r_la, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in [3, 4, 5, 6, 7]:
        c_let = get_column_letter(c)
        ws4.cell(row=r_la, column=c, value=f"=SUM({c_let}4:{c_let}{r_la-1})")
        ws4.cell(row=r_la, column=c).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws4.cell(row=r_la, column=c).number_format = "#,##0"
        ws4.cell(row=r_la, column=c).alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 9):
        ws4.cell(row=r_la, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws4.cell(row=r_la, column=c).border = thick_bottom

    ws4.column_dimensions["A"].width = 6
    ws4.column_dimensions["B"].width = 46
    ws4.column_dimensions["C"].width = 22
    ws4.column_dimensions["D"].width = 22
    ws4.column_dimensions["E"].width = 26
    ws4.column_dimensions["F"].width = 20
    ws4.column_dimensions["G"].width = 22
    ws4.column_dimensions["H"].width = 38


    # =========================================================================
    # SHEET 5: 05_DoiChieu_BocTach_A5
    # =========================================================================
    ws5 = wb.create_sheet(title="05_DoiChieu_BocTach_A5")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:H1")
    ws5["A1"] = "BẢNG ĐỐI CHIẾU KHỐI LƯỢNG THỰC TẾ HỒ SƠ CỐNG HỘP TUYẾN A5"
    ws5["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws5["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws5 = [
        "STT", "Hạng mục cống hộp Tuyến A5", "Quy cách (BxH)", "Chiều dài L (m)",
        "Bê tông thân B20 (m3)", "Bê tông lót M100 (m3)", "Ván khuôn tiếp xúc (m2)", "Cốt thép CB300/CB400 (tấn)"
    ]
    for c_i, h in enumerate(headers_ws5, 1):
        cell = ws5.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    boq_a5 = [
        (1, "Đoạn 1 - Tuyến cống đơn 2.0x2.0m (12 ga)", "2.0 x 2.0 m", 385.0, 897.05, 51.98, 4747.05, 80.73),
        (2, "Đoạn 2 - Tuyến cống đơn 3.0x3.0m (17 ga)", "3.0 x 3.0 m", 560.0, 2287.60, 106.40, 9861.60, 154.84),
        (3, "Đoạn 3 - Tuyến cống đôi 2x(3.0x3.0m) (37 ga)", "2 x (3.0 x 3.0) m", 1225.0, 8685.25, 434.88, 30331.00, 517.80)
    ]

    r_boq = 4
    for stt, name, spec, l_m, bt_b20, bt_m100, vk_m2, thep_t in boq_a5:
        ws5.cell(row=r_boq, column=1, value=stt)
        ws5.cell(row=r_boq, column=2, value=name)
        ws5.cell(row=r_boq, column=3, value=spec)
        ws5.cell(row=r_boq, column=4, value=l_m)
        ws5.cell(row=r_boq, column=5, value=bt_b20)
        ws5.cell(row=r_boq, column=6, value=bt_m100)
        ws5.cell(row=r_boq, column=7, value=vk_m2)
        ws5.cell(row=r_boq, column=8, value=thep_t)

        for c in range(1, 9):
            cell = ws5.cell(row=r_boq, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 3]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [4, 5, 6, 7, 8]:
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_boq += 1

    # Total
    ws5.cell(row=r_boq, column=2, value="TỔNG CỘNG TOÀN TUYẾN A5 (100%)")
    ws5.cell(row=r_boq, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in [4, 5, 6, 7, 8]:
        c_let = get_column_letter(c)
        ws5.cell(row=r_boq, column=c, value=f"=SUM({c_let}4:{c_let}{r_boq-1})")
        ws5.cell(row=r_boq, column=c).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws5.cell(row=r_boq, column=c).number_format = "#,##0.0"
        ws5.cell(row=r_boq, column=c).alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 9):
        ws5.cell(row=r_boq, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws5.cell(row=r_boq, column=c).border = thick_bottom

    ws5.column_dimensions["A"].width = 6
    ws5.column_dimensions["B"].width = 46
    ws5.column_dimensions["C"].width = 22
    ws5.column_dimensions["D"].width = 16
    ws5.column_dimensions["E"].width = 22
    ws5.column_dimensions["F"].width = 22
    ws5.column_dimensions["G"].width = 24
    ws5.column_dimensions["H"].width = 26

    # Save to paths
    target_paths = [
        r"c:\Users\baotu\Downloads\CỐNG HỘP TUYẾN A5\260920_TDTC_CaXe_CaMay_Cong_Hop_Tuyen_A5.xlsx",
        r"c:\Users\baotu\Downloads\TIEN_DO_THI_CONG_CUM_B9_OLYMPIC\260920_TDTC_CaXe_CaMay_Cong_Hop_Tuyen_A5.xlsx",
        r"d:\Code\23HG-multiagent-system\23HG-multiagent-system\examples\260920_TDTC_CaXe_CaMay_Cong_Hop_Tuyen_A5.xlsx"
    ]

    for p in target_paths:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        wb.save(p)
        print(f"[OK] File Ca xe ca máy đã lưu thành công: {p}")

if __name__ == "__main__":
    build_a5_machine_schedule()
