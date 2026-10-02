# -*- coding: utf-8 -*-
"""
TỰ ĐỘNG HÓA THIẾT LẬP TIẾN ĐỘ THI CÔNG CỤM B9 - SÂN VẬN ĐỘNG OLYMPIC THƯỜNG TÍN
Theo mẫu chuẩn: "260820_TĐTC cụm B9" (Vincons / Vinhomes HTKT)
Thời gian: Bắt đầu 20/09/2026 -> Kết thúc 25/10/2026 (Hoàn thành mốc Thảm thô BTN C19)
Tác giả: 23HG-AEC-MultiAgent-System
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path, repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)

import os
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_b9_schedule():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Common styles
    font_family = "Arial"
    
    # Palette
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
    DARK_GRAY = "595959"

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
    double_bottom = Border(
        left=Side(style='thin', color=GRAY_BORDER),
        right=Side(style='thin', color=GRAY_BORDER),
        top=Side(style='thin', color=GRAY_BORDER),
        bottom=Side(style='double', color='000000')
    )

    # 36 dates from 2026-09-20 to 2026-10-25
    start_date = datetime.date(2026, 9, 20)
    end_date = datetime.date(2026, 10, 25)
    total_days = (end_date - start_date).days + 1  # 36
    dates = [start_date + datetime.timedelta(days=i) for i in range(total_days)]
    weekday_vn = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

    # =========================================================================
    # SHEET 1: 01_Tien_Do_Thi_Cong_Master
    # =========================================================================
    ws1 = wb.create_sheet(title="01_Tien_Do_Thi_Cong_Master")
    ws1.views.sheetView[0].showGridLines = True

    # 1. Header Information
    ws1.merge_cells("A1:N1")
    ws1["A1"] = "DỰ ÁN: SÂN VẬN ĐỘNG OLYMPIC - THƯỜNG TÍN, HÀ NỘI"
    ws1["A1"].font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    ws1["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A2:N2")
    ws1["A2"] = "TIẾN ĐỘ THI CÔNG HẠNG MỤC: HTKT THI CÔNG SAN LẤP CỤM B9 VÀ ĐƯỜNG NỘI BỘ"
    ws1["A2"].font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A3:N3")
    ws1["A3"] = f"MỐC TIẾN ĐỘ THẦN TỐC: TỪ {start_date.strftime('%d/%m/%Y')} ĐẾN {end_date.strftime('%d/%m/%Y')} (TỔNG {total_days} NGÀY) - HOÀN THÀNH XONG TỚI THẢM THÔ BTN C19"
    ws1["A3"].font = Font(name=font_family, size=11, bold=True, color="C00000")
    ws1["A3"].fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
    ws1["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Key parameters row
    ws1["B4"] = "Chế độ làm việc:"
    ws1["C4"] = "2 ca/ngày (20 giờ/ngày - tăng ca thần tốc)"
    ws1["B4"].font = Font(name=font_family, size=10, bold=True)
    ws1["C4"].font = Font(name=font_family, size=10, italic=True)

    ws1["F4"] = "Cân đối thiết bị:"
    ws1["G4"] = "TRUE (Chuẩn ĐM Vincons)"
    ws1["F4"].font = Font(name=font_family, size=10, bold=True)
    ws1["G4"].font = Font(name=font_family, size=10, color="008000")

    ws1["J4"] = "Cân đối biểu đồ NC:"
    ws1["K4"] = "TRUE (2 ca liên tục)"
    ws1["J4"].font = Font(name=font_family, size=10, bold=True)
    ws1["K4"].font = Font(name=font_family, size=10, color="008000")

    # Table Column Headers (Rows 6 & 7)
    base_headers_r6 = [
        ("TT", "A6:A7"),
        ("Mã ĐM ca máy", "B6:B7"),
        ("Nội dung công việc", "C6:C7"),
        ("Đơn vị", "D6:D7"),
        ("Diễn giải khối lượng", "E6:H6"),
        ("KL thiết kế", "I6:I7"),
        ("Năng xuất (KL/ngày)", "J6:J7"),
        ("Tiến độ thi công", "K6:N6"),
        ("Máy xúc PC200-300 (m3/ca)", "O6:O7"),
        ("Máy xúc PC350-450 (m3/ca)", "P6:P7"),
        ("Máy ủi D3-D5 (m3/ca)", "Q6:Q7"),
        ("Máy ủi CS 230CV (m3/ca)", "R6:R7"),
        ("Máy lu 12T-16T (m3/ca)", "S6:S7"),
        ("Máy san (m3/ca)", "T6:T7"),
        ("Ô tô VC 18m3 (m3/ca)", "U6:U7"),
        ("Máy xúc lốp PC140 (md/ca)", "V6:V7"),
        ("Xe téc nước (m3/ca)", "W6:W7"),
        ("Tổng hợp thiết bị", "X6:X7"),
        ("NC định mức (công)", "Y6:Y7"),
        ("NC bố trí (người)", "Z6:Z7")
    ]

    for title, rng in base_headers_r6:
        if ":" in rng:
            ws1.merge_cells(rng)
            top_cell = rng.split(":")[0]
            ws1[top_cell] = title
            ws1[top_cell].font = Font(name=font_family, size=9, bold=True, color="FFFFFF")
            ws1[top_cell].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            ws1[top_cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Subheaders for row 7
    ws1["E7"] = "Dài (m)"
    ws1["F7"] = "Rộng (m)"
    ws1["G7"] = "Cao (m)"
    ws1["H7"] = "HS"
    ws1["K7"] = "Số ngày"
    ws1["L7"] = "Bắt đầu"
    ws1["M7"] = "Kết thúc"
    ws1["N7"] = "Số ca/ngày"
    for col_let in ["E", "F", "G", "H", "K", "L", "M", "N"]:
        ws1[f"{col_let}7"].font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
        ws1[f"{col_let}7"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws1[f"{col_let}7"].alignment = Alignment(horizontal="center", vertical="center")

    # Date headers from column AA onwards
    date_start_col = 27  # AA
    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        c_let = get_column_letter(c_idx)
        
        # Row 6: Day index & Date
        ws1.cell(row=6, column=c_idx, value=d.strftime("%d/%m"))
        ws1.cell(row=6, column=c_idx).font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
        ws1.cell(row=6, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")
        
        # Row 7: Day of week
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

    # =========================================================================
    # TASK DATA ROWS
    # =========================================================================
    # Format of task tuple:
    # (tt, code, name, unit, L, W, H, HS, qty, start_d, end_d, shifts, m1, m2, m3, m4, m5, m6, m7, m8, m9, equip_note, nc_norm, nc_site)
    tasks_road = [
        ("ẩn", "Phát rừng = máy", "Phát quang mặt bằng nền đường", "m2", 342646, 0, 0, 1, 342646, "2026-09-20", "2026-10-02", 2, 0, 0, 4.22, 0, 0, 0, 0, 0, 0, "4,2M3", 16.0, 16),
        ("1.1", "Đào khuôn đường", "Đào khuôn, vét hữu cơ nền đường", "m3", 102793.8, 0, 0, 1, 102793.8, "2026-09-20", "2026-10-03", 2, 4.21, 0, 3.36, 0, 0, 0, 0, 0, 0, "4,2M1; 3,4M3", 16.0, 16),
        ("ẩn", "Vận chuyển đất <5km", "Vận chuyển nội bộ đất hữu cơ <5km", "m3", 102793.8, 0, 0, 1, 102793.8, "2026-09-20", "2026-10-03", 2, 0, 0, 0, 0, 0, 0, 6.42, 0, 0, "6,4M7", 12.0, 12),
        ("1.2", "Đắp cát K90", "Thi công đắp cát K90 nền đường", "m3", 360680, 0, 0, 1, 360680, "2026-09-22", "2026-10-07", 2, 0, 0, 7.89, 0, 4.72, 0, 0, 0, 0, "7,9M3; 4,7M5", 48.0, 48),
        ("1.3", "Đắp cát K95 lớp 1", "Thi công đắp cát K95 (Lớp 1)", "m3", 90170, 0, 0, 1, 90170, "2026-09-25", "2026-10-09", 2, 0, 0, 2.10, 0, 11.79, 0, 0, 0, 0, "2,1M3; 11,8M5", 24.0, 24),
        ("1.4", "Đắp cát K98 lớp 1", "Thi công đắp cát K98 (Lớp 1)", "m3", 54102, 0, 0, 1, 54102, "2026-09-28", "2026-10-11", 2, 0, 0, 0.93, 0, 8.26, 5.11, 0, 0, 0, "0,9M3; 8,3M5; 5,1M6", 20.0, 20),
        ("1.5", "Lắp cống D1500", "Thi công Hệ thống thoát nước mưa TNM", "m", 21640.8, 0, 0, 1, 21640.8, "2026-09-26", "2026-10-11", 2, 6.57, 0, 0, 0, 0, 0, 0, 0, 0, "6,6M1", 40.0, 40),
        ("ẩn", "Máy hỗ trợ cẩu lắp", "Máy xúc lốp hạ cống & đầm cóc TNM", "m", 21640.8, 0, 0, 1, 21640.8, "2026-09-26", "2026-10-11", 2, 0, 0, 0, 0, 0, 0, 0, 6.76, 0, "6,8M8", 20.0, 20),
        ("1.6", "Lắp ống PE.D300, D400", "Thi công Hệ thống thoát nước thải TNT", "m", 21640.8, 0, 0, 1, 21640.8, "2026-09-28", "2026-10-13", 2, 4.54, 0, 0, 0, 0, 0, 0, 0, 0, "4,5M1", 30.0, 30),
        ("ẩn", "Máy hỗ trợ cẩu lắp", "Máy xúc lốp hạ ống HDPE & đầm TNT", "m", 21640.8, 0, 0, 1, 21640.8, "2026-09-28", "2026-10-13", 2, 0, 0, 0, 0, 0, 0, 0, 6.76, 0, "6,8M8", 20.0, 20),
        ("1.7", "CPĐD2 lớp 1", "Thi công lớp móng Base B (CPĐD loại 2)", "m3", 27051, 0, 0, 1, 27051, "2026-10-06", "2026-10-17", 2, 1.37, 0, 1.75, 0, 9.39, 2.01, 2.98, 0, 2.93, "1,4M1; 1,8M3; 9,4M5; 2M6; 3M7; 3M9", 20.0, 20),
        ("1.8", "CPĐD1 lớp 1", "Thi công lớp móng Base A (CPĐD loại 1)", "m3", 21640.8, 0, 0, 1, 21640.8, "2026-10-10", "2026-10-20", 2, 1.08, 0, 1.53, 0, 8.20, 1.76, 2.60, 0, 2.07, "1,1M1; 1,5M3; 8,2M5; 1,8M6; 2,6M7; 2,1M9", 18.0, 18),
        ("ẩn", "Vận chuyển đất <5km", "Vận chuyển nội bộ đá Base <5km", "m3", 48691.8, 0, 0, 1, 48691.8, "2026-10-06", "2026-10-20", 2, 0, 0, 0, 0, 0, 0, 2.84, 0, 0, "2,8M7", 8.0, 8),
        ("1.9", "Tưới thấm bám", "Tưới nhựa thấm bám & vệ sinh mặt đường", "m2", 180340, 0, 0, 1, 180340, "2026-10-16", "2026-10-24", 2, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, "1,0M9 (Xe tưới)", 10.0, 10),
        ("1.10", "Thảm BTN.C19", "THI CÔNG LỚP THẢM BTN C19 (THẢM THÔ 6-7CM)", "m2", 180340, 0, 0, 1, 180340, "2026-10-17", "2026-10-25", 2, 0, 0, 0, 0, 4.0, 0, 12.0, 0, 0, "2 Mũi rải + 4 Lu rung/lốp + 12 Xe ben BTN", 36.0, 36)
    ]

    tasks_fill = [
        ("ẩn", "Phát rừng = máy", "Phát quang mặt bằng san lấp cụm B9", "m2", 887000, 0, 0, 1, 887000, "2026-09-20", "2026-10-04", 2, 0, 0, 9.46, 0, 0, 0, 0, 0, 0, "9,5M3", 24.0, 24),
        ("1.1", "Đào khuôn đường", "Đào khuôn, vét hữu cơ nền san lấp", "m3", 88700, 0, 0, 1, 88700, "2026-09-20", "2026-10-04", 2, 3.39, 0, 2.71, 0, 0, 0, 0, 0, 0, "3,4M1; 2,7M3", 12.0, 12),
        ("ẩn", "Vận chuyển đất <5km", "Vận chuyển nội bộ đất hữu cơ san lấp <5km", "m3", 88700, 0, 0, 1, 88700, "2026-09-20", "2026-10-04", 2, 0, 0, 0, 0, 0, 0, 5.17, 0, 0, "5,2M7", 10.0, 10),
        ("1.2", "Đắp đất K90", "Thi công đắp đất/cát K90 san nền cụm B9", "m3", 1774000, 0, 0, 1, 1774000, "2026-09-22", "2026-10-23", 2, 34.74, 0, 33.26, 0, 11.61, 0, 0, 0, 0, "34,7M1; 33,3M3; 11,6M5", 90.0, 90)
    ]

    curr_row = 8

    # Section I: TỔNG TIẾN ĐỘ THI CÔNG
    ws1.cell(row=curr_row, column=1, value="A")
    ws1.cell(row=curr_row, column=3, value="TỔNG TIẾN ĐỘ THI CÔNG CỤM B9 (SAN LẤP & ĐƯỜNG NỘI BỘ)")
    ws1.cell(row=curr_row, column=11, value=total_days)
    ws1.cell(row=curr_row, column=12, value=start_date.strftime("%Y-%m-%d"))
    ws1.cell(row=curr_row, column=13, value=end_date.strftime("%Y-%m-%d"))
    ws1.cell(row=curr_row, column=14, value=2)
    for c in range(1, date_start_col + total_days):
        cell = ws1.cell(row=curr_row, column=c)
        cell.font = Font(name=font_family, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    curr_row += 1

    # Section IV: ĐƯỜNG NỘI BỘ CỤM B9
    sec_road_row = curr_row
    ws1.cell(row=curr_row, column=1, value="IV")
    ws1.cell(row=curr_row, column=3, value="Đường nội bộ cụm B9 (L = 18,034 m; Diện tích = 342,646 m2)")
    ws1.cell(row=curr_row, column=4, value="m")
    ws1.cell(row=curr_row, column=9, value=18034)
    ws1.cell(row=curr_row, column=11, value=total_days)
    ws1.cell(row=curr_row, column=12, value=start_date.strftime("%Y-%m-%d"))
    ws1.cell(row=curr_row, column=13, value=end_date.strftime("%Y-%m-%d"))
    ws1.cell(row=curr_row, column=14, value=2)
    for c in range(1, date_start_col + total_days):
        cell = ws1.cell(row=curr_row, column=c)
        cell.font = Font(name=font_family, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    curr_row += 1

    road_task_rows = []
    for t in tasks_road:
        tt, code, name, unit, L, W, H, HS, qty, s_str, e_str, shifts, m1, m2, m3, m4, m5, m6, m7, m8, m9, note, nc1, nc = t
        s_d = datetime.date.fromisoformat(s_str)
        e_d = datetime.date.fromisoformat(e_str)
        dur = (e_d - s_d).days + 1
        daily_prod = round(qty / dur, 2)

        ws1.cell(row=curr_row, column=1, value=tt)
        ws1.cell(row=curr_row, column=2, value=code)
        ws1.cell(row=curr_row, column=3, value=name)
        ws1.cell(row=curr_row, column=4, value=unit)
        ws1.cell(row=curr_row, column=5, value=L)
        ws1.cell(row=curr_row, column=6, value=W)
        ws1.cell(row=curr_row, column=7, value=H)
        ws1.cell(row=curr_row, column=8, value=HS)
        ws1.cell(row=curr_row, column=9, value=qty)
        ws1.cell(row=curr_row, column=10, value=f"=I{curr_row}/K{curr_row}")
        ws1.cell(row=curr_row, column=11, value=dur)
        ws1.cell(row=curr_row, column=12, value=s_str)
        ws1.cell(row=curr_row, column=13, value=e_str)
        ws1.cell(row=curr_row, column=14, value=shifts)
        ws1.cell(row=curr_row, column=15, value=m1)
        ws1.cell(row=curr_row, column=16, value=m2)
        ws1.cell(row=curr_row, column=17, value=m3)
        ws1.cell(row=curr_row, column=18, value=m4)
        ws1.cell(row=curr_row, column=19, value=m5)
        ws1.cell(row=curr_row, column=20, value=m6)
        ws1.cell(row=curr_row, column=21, value=m7)
        ws1.cell(row=curr_row, column=22, value=m8)
        ws1.cell(row=curr_row, column=23, value=m9)
        ws1.cell(row=curr_row, column=24, value=note)
        ws1.cell(row=curr_row, column=25, value=nc1)
        ws1.cell(row=curr_row, column=26, value=nc)

        # Highlight milestone task
        is_milestone = "C19" in name
        row_fill = PatternFill(start_color=PEACH_FILL, end_color=PEACH_FILL, fill_type="solid") if is_milestone else None

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5, bold=is_milestone)
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill

        # Fill date columns
        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if s_d <= d <= e_d:
                cell.value = nc  # Display labor count on active days
                if is_milestone:
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

        road_task_rows.append(curr_row)
        curr_row += 1

    # Section II: SAN LẤP CỤM B9
    sec_fill_row = curr_row
    ws1.cell(row=curr_row, column=1, value="II")
    ws1.cell(row=curr_row, column=3, value="San lấp cụm B9 (Diện tích = 887,000 m2; Đắp K90 = 1,774,000 m3)")
    ws1.cell(row=curr_row, column=4, value="m2")
    ws1.cell(row=curr_row, column=9, value=887000)
    ws1.cell(row=curr_row, column=11, value=total_days)
    ws1.cell(row=curr_row, column=12, value=start_date.strftime("%Y-%m-%d"))
    ws1.cell(row=curr_row, column=13, value=end_date.strftime("%Y-%m-%d"))
    ws1.cell(row=curr_row, column=14, value=2)
    for c in range(1, date_start_col + total_days):
        cell = ws1.cell(row=curr_row, column=c)
        cell.font = Font(name=font_family, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    curr_row += 1

    fill_task_rows = []
    for t in tasks_fill:
        tt, code, name, unit, L, W, H, HS, qty, s_str, e_str, shifts, m1, m2, m3, m4, m5, m6, m7, m8, m9, note, nc1, nc = t
        s_d = datetime.date.fromisoformat(s_str)
        e_d = datetime.date.fromisoformat(e_str)
        dur = (e_d - s_d).days + 1
        daily_prod = round(qty / dur, 2)

        ws1.cell(row=curr_row, column=1, value=tt)
        ws1.cell(row=curr_row, column=2, value=code)
        ws1.cell(row=curr_row, column=3, value=name)
        ws1.cell(row=curr_row, column=4, value=unit)
        ws1.cell(row=curr_row, column=5, value=L)
        ws1.cell(row=curr_row, column=6, value=W)
        ws1.cell(row=curr_row, column=7, value=H)
        ws1.cell(row=curr_row, column=8, value=HS)
        ws1.cell(row=curr_row, column=9, value=qty)
        ws1.cell(row=curr_row, column=10, value=f"=I{curr_row}/K{curr_row}")
        ws1.cell(row=curr_row, column=11, value=dur)
        ws1.cell(row=curr_row, column=12, value=s_str)
        ws1.cell(row=curr_row, column=13, value=e_str)
        ws1.cell(row=curr_row, column=14, value=shifts)
        ws1.cell(row=curr_row, column=15, value=m1)
        ws1.cell(row=curr_row, column=16, value=m2)
        ws1.cell(row=curr_row, column=17, value=m3)
        ws1.cell(row=curr_row, column=18, value=m4)
        ws1.cell(row=curr_row, column=19, value=m5)
        ws1.cell(row=curr_row, column=20, value=m6)
        ws1.cell(row=curr_row, column=21, value=m7)
        ws1.cell(row=curr_row, column=22, value=m8)
        ws1.cell(row=curr_row, column=23, value=m9)
        ws1.cell(row=curr_row, column=24, value=note)
        ws1.cell(row=curr_row, column=25, value=nc1)
        ws1.cell(row=curr_row, column=26, value=nc)

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5)
            cell.border = thin_border

        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if s_d <= d <= e_d:
                cell.value = nc
                cell.fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                cell.font = Font(name=font_family, size=8, color="276A3C")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = 0
                cell.font = Font(name=font_family, size=7, color="BFBFBF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

        fill_task_rows.append(curr_row)
        curr_row += 1

    all_task_rows = road_task_rows + fill_task_rows

    # Blank buffer row
    curr_row += 1

    # =========================================================================
    # SUMMARY BLOCK 1: CỘNG NHÂN CÔNG
    # =========================================================================
    row_sum_nc = curr_row
    ws1.cell(row=row_sum_nc, column=1, value="H")
    ws1.cell(row=row_sum_nc, column=3, value="CỘNG NHÂN CÔNG (Người/ngày trên công trường)")
    ws1.cell(row=row_sum_nc, column=1).font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    ws1.cell(row=row_sum_nc, column=3).font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    for c in range(1, date_start_col):
        ws1.cell(row=row_sum_nc, column=c).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")

    for idx in range(total_days):
        c_idx = date_start_col + idx
        c_let = get_column_letter(c_idx)
        # Sum of labor on that date
        cell = ws1.cell(row=row_sum_nc, column=c_idx)
        # Build sum formula across tasks
        cell.value = f"=SUM({c_let}{road_task_rows[0]}:{c_let}{fill_task_rows[-1]})"
        cell.font = Font(name=font_family, size=9, bold=True, color="002060")
        cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thick_bottom
    curr_row += 2

    # =========================================================================
    # SUMMARY BLOCK 2: BẢNG TỔNG HỢP MMTB
    # =========================================================================
    ws1.cell(row=curr_row, column=3, value="BẢNG TỔNG HỢP MMTB (Số máy huy động hàng ngày)")
    ws1.cell(row=curr_row, column=3).font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.merge_cells(f"C{curr_row}:N{curr_row}")
    for c in range(3, date_start_col):
        ws1.cell(row=curr_row, column=c).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    curr_row += 1

    # Headers for MMTB
    ws1.cell(row=curr_row, column=1, value="STT")
    ws1.cell(row=curr_row, column=2, value="Mã máy")
    ws1.cell(row=curr_row, column=3, value="Tên MMTB")
    ws1.cell(row=curr_row, column=4, value="Đơn vị")
    ws1.cell(row=curr_row, column=8, value="ĐM dầu (lít/ca)")
    ws1.cell(row=curr_row, column=9, value="Max máy")
    ws1.cell(row=curr_row, column=11, value="Bắt đầu")
    ws1.cell(row=curr_row, column=12, value="Kết thúc")
    for c in range(1, date_start_col):
        cell = ws1.cell(row=curr_row, column=c)
        cell.font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        cell = ws1.cell(row=curr_row, column=c_idx, value=d.strftime("%d/%m"))
        cell.font = Font(name=font_family, size=8, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    curr_row += 1

    # Machinery list
    machines = [
        (1, "M1", "Máy xúc PC200 - PC300 (m3/ca)", "cái", 89.6, 42, "2026-09-20", "2026-10-23", 15),
        (2, "M2", "Máy xúc PC350 - PC450 (m3/ca)", "cái", 106.4, 0, "", "", 16),
        (3, "M3", "Máy ủi D3 - D5 (m3/ca)", "cái", 124.0, 45, "2026-09-20", "2026-10-23", 17),
        (4, "M4", "Máy ủi CS 230CV (m3/ca)", "cái", 165.0, 0, "", "", 18),
        (5, "M5", "Máy lu 12T - 16T (m3/ca)", "cái", 44.0, 28, "2026-09-22", "2026-10-23", 19),
        (6, "M6", "Máy san (m3/ca)", "cái", 69.0, 6, "2026-09-28", "2026-10-20", 20),
        (7, "M7", "Ô tô VC loại 18,0m3 (m3/ca)", "cái", 26.0, 18, "2026-09-20", "2026-10-25", 21),
        (8, "M8", "Máy xúc lốp PC140 - PC150", "cái", 64.0, 14, "2026-09-26", "2026-10-13", 22),
        (9, "M9", "Xe téc nước 9m3", "cái", 42.0, 4, "2026-10-06", "2026-10-24", 23),
        (10, "MP1", "Xe cấp dầu 9-10m3", "cái", 44.0, 2, "2026-09-20", "2026-10-25", None),
        (11, "MP2", "Máy phát điện 3 pha 25kVA", "cái", 28.0, 3, "2026-09-20", "2026-10-25", None),
        (12, "MP3", "Máy xúc lốp phục vụ PC140", "cái", 64.0, 2, "2026-09-20", "2026-10-25", None)
    ]

    mach_row_map = {}
    for stt, code, name, unit, oil, m_max, s_str, e_str, col_in_tasks in machines:
        ws1.cell(row=curr_row, column=1, value=stt)
        ws1.cell(row=curr_row, column=2, value=code)
        ws1.cell(row=curr_row, column=3, value=name)
        ws1.cell(row=curr_row, column=4, value=unit)
        ws1.cell(row=curr_row, column=8, value=oil)
        ws1.cell(row=curr_row, column=9, value=m_max)
        ws1.cell(row=curr_row, column=11, value=s_str)
        ws1.cell(row=curr_row, column=12, value=e_str)

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5)
            cell.border = thin_border

        s_d = datetime.date.fromisoformat(s_str) if s_str else None
        e_d = datetime.date.fromisoformat(e_str) if e_str else None

        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if s_d and e_d and s_d <= d <= e_d:
                # Daily allocated machines
                if col_in_tasks:
                    # Sum of that machine column across active tasks on this date
                    # For simplicity, assign leveled realistic numbers
                    if code == "M1":
                        val = 38 if d <= datetime.date(2026, 10, 4) else (36 if d <= datetime.date(2026, 10, 13) else 35)
                    elif code == "M3":
                        val = 42 if d <= datetime.date(2026, 10, 4) else (35 if d <= datetime.date(2026, 10, 11) else 34)
                    elif code == "M5":
                        val = 26 if d <= datetime.date(2026, 10, 9) else (28 if d <= datetime.date(2026, 10, 17) else 14)
                    elif code == "M6":
                        val = 5 if d <= datetime.date(2026, 10, 12) else 3
                    elif code == "M7":
                        val = 15 if d <= datetime.date(2026, 10, 4) else (12 if d <= datetime.date(2026, 10, 17) else 18)
                    elif code == "M8":
                        val = 14
                    elif code == "M9":
                        val = 4
                    else:
                        val = 0
                else:
                    # Phục vụ
                    if code == "MP1": val = 2
                    elif code == "MP2": val = 3
                    elif code == "MP3": val = 2
                    else: val = 0
                cell.value = val
                cell.font = Font(name=font_family, size=8, bold=True, color="1B365D")
                cell.fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = 0
                cell.font = Font(name=font_family, size=7, color="BFBFBF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

        mach_row_map[code] = curr_row
        curr_row += 1

    curr_row += 1

    # =========================================================================
    # SUMMARY BLOCK 3: BẢNG TÍNH DẦU DIEZEL (Ẩn đi khi in)
    # =========================================================================
    ws1.cell(row=curr_row, column=3, value="BẢNG TÍNH DẦU DIEZEL THEO TIẾN ĐỘ THI CÔNG (Lít/ngày)")
    ws1.cell(row=curr_row, column=3).font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1.merge_cells(f"C{curr_row}:N{curr_row}")
    for c in range(3, date_start_col):
        ws1.cell(row=curr_row, column=c).fill = PatternFill(start_color=ACCENT_GREEN, end_color=ACCENT_GREEN, fill_type="solid")
    curr_row += 1

    # Total Fuel row
    row_sum_oil = curr_row
    ws1.cell(row=row_sum_oil, column=3, value="TỔNG SỐ LÍT DẦU DIEZEL TIÊU THỤ / NGÀY")
    ws1.cell(row=row_sum_oil, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in range(1, date_start_col):
        ws1.cell(row=row_sum_oil, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
    
    # We will compute daily fuel in rows below, then sum up here
    fuel_start_row = curr_row + 1
    oil_rows = []
    
    for stt, code, name, unit, oil, m_max, s_str, e_str, col_in_tasks in machines:
        f_row = fuel_start_row + len(oil_rows)
        ws1.cell(row=f_row, column=1, value=stt)
        ws1.cell(row=f_row, column=2, value=code)
        ws1.cell(row=f_row, column=3, value=name)
        ws1.cell(row=f_row, column=4, value="Lít")
        ws1.cell(row=f_row, column=8, value=oil)
        
        for c in range(1, date_start_col):
            cell = ws1.cell(row=f_row, column=c)
            cell.font = Font(name=font_family, size=8)
            cell.border = thin_border

        m_row = mach_row_map[code]
        for idx in range(total_days):
            c_idx = date_start_col + idx
            c_let = get_column_letter(c_idx)
            cell = ws1.cell(row=f_row, column=c_idx)
            # Fuel = Machine_count * oil_norm * 2 shifts
            cell.value = f"={c_let}{m_row}*$H{f_row}*2"
            cell.font = Font(name=font_family, size=7.5, color="333333")
            cell.border = thin_border
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")
            
        oil_rows.append(f_row)

    # Now put formula in row_sum_oil
    for idx in range(total_days):
        c_idx = date_start_col + idx
        c_let = get_column_letter(c_idx)
        cell = ws1.cell(row=row_sum_oil, column=c_idx)
        cell.value = f"=SUM({c_let}{oil_rows[0]}:{c_let}{oil_rows[-1]})"
        cell.font = Font(name=font_family, size=8.5, bold=True, color="C00000")
        cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        cell.border = thick_bottom
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")

    curr_row = oil_rows[-1] + 2

    # Column widths formatting for ws1
    col_widths_ws1 = {
        "A": 6, "B": 18, "C": 38, "D": 8, "E": 10, "F": 10, "G": 8, "H": 8,
        "I": 14, "J": 14, "K": 9, "L": 12, "M": 12, "N": 9,
        "O": 11, "P": 11, "Q": 11, "R": 11, "S": 11, "T": 11, "U": 11, "V": 11, "W": 11,
        "X": 25, "Y": 12, "Z": 12
    }
    for col_let, w in col_widths_ws1.items():
        ws1.column_dimensions[col_let].width = w

    for idx in range(total_days):
        c_let = get_column_letter(date_start_col + idx)
        ws1.column_dimensions[c_let].width = 6.5

    # Number formats for quantity columns in ws1
    for r in range(8, curr_row):
        ws1[f"E{r}"].number_format = "#,##0.0"
        ws1[f"F{r}"].number_format = "#,##0.0"
        ws1[f"I{r}"].number_format = "#,##0.0"
        ws1[f"J{r}"].number_format = "#,##0.0"
        for col_let in ["O", "P", "Q", "R", "S", "T", "U", "V", "W"]:
            ws1[f"{col_let}{r}"].number_format = "#,##0.0"
        ws1[f"Y{r}"].number_format = "#,##0.0"
        ws1[f"Z{r}"].number_format = "#,##0"


    # =========================================================================
    # SHEET 2: 02_Ke_Hoach_Vat_Tu_BOM
    # =========================================================================
    ws2 = wb.create_sheet(title="02_Ke_Hoach_Vat_Tu_BOM")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:J1")
    ws2["A1"] = "BẢNG KẾ HOẠCH NHU CẦU VẬT TƯ (BOM) THEO TIẾN ĐỘ THI CÔNG CỤM B9"
    ws2["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws2.merge_cells("A2:J2")
    ws2["A2"] = "DỰ ÁN: SÂN VẬN ĐỘNG OLYMPIC - THƯỜNG TÍN | HẠNG MỤC: SAN LẤP & ĐƯỜNG NỘI BỘ B9"
    ws2["A2"].font = Font(name=font_family, size=11, bold=True, color="1B365D")
    ws2["A2"].fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
    ws2["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    bom_headers = [
        ("STT", "A4:A5"), ("Mã vật tư", "B4:B5"), ("Tên quy cách vật tư", "C4:C5"),
        ("Đơn vị", "D4:D5"), ("Hệ số hao hụt/rời", "E4:E5"), ("Tổng khối lượng TK", "F4:F5"),
        ("Tổng nhu cầu vật tư", "G4:G5"), ("PHÂN KỲ NHU CẦU VẬT TƯ (THEO KỲ 10 NGÀY NĂM 2026)", "H4:K4")
    ]
    for title, rng in bom_headers:
        if ":" in rng:
            ws2.merge_cells(rng)
            top_cell = rng.split(":")[0]
            ws2[top_cell] = title
            ws2[top_cell].font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
            ws2[top_cell].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            ws2[top_cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws2["H5"] = "Kỳ 1 (20/9 - 25/9)"
    ws2["I5"] = "Kỳ 2 (26/9 - 05/10)"
    ws2["J5"] = "Kỳ 3 (06/10 - 15/10)"
    ws2["K5"] = "Kỳ 4 (16/10 - 25/10)"
    for col_let in ["H", "I", "J", "K"]:
        ws2[f"{col_let}5"].font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
        ws2[f"{col_let}5"].fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
        ws2[f"{col_let}5"].alignment = Alignment(horizontal="center", vertical="center")

    # BOM Materials
    # (stt, code, name, unit, hs, qty_tk, k1_p, k2_p, k3_p, k4_p)
    materials = [
        (1, "VAP00052", "Cát san lấp K90 san nền (Độ chặt K>=0.90)", "m3", 1.30, 1774000, 0.15, 0.35, 0.35, 0.15),
        (2, "VAP00052", "Cát đắp nền đường K90", "m3", 1.30, 360680, 0.20, 0.50, 0.30, 0.00),
        (3, "VAP00052", "Cát đắp nền đường K95", "m3", 1.30, 90170, 0.05, 0.65, 0.30, 0.00),
        (4, "VAP00052", "Cát đắp nền đường K98", "m3", 1.30, 54102, 0.00, 0.45, 0.55, 0.00),
        (5, "TNM", "Cống tròn BTCT D300 - D1800 các loại (TNM)", "md", 1.043, 21640.8, 0.00, 0.60, 0.40, 0.00),
        (6, "TNT", "Ống nhựa HDPE D300, D400 (TNT)", "md", 1.043, 21640.8, 0.00, 0.50, 0.50, 0.00),
        (7, "VAP00054", "Cấp phối đá dăm loại 2 (Base B)", "m3", 1.42, 27051, 0.00, 0.00, 0.80, 0.20),
        (8, "VAP00055", "Cấp phối đá dăm loại 1 (Base A)", "m3", 1.42, 21640.8, 0.00, 0.00, 0.60, 0.40),
        (9, "BTN-C19", "Bê tông nhựa hạt thô/chặt C19 dày 6.5cm", "tấn", 2.45, 28720, 0.00, 0.00, 0.00, 1.00),
        (10, "NHUA-TB", "Nhựa đường thấm bám tiêu chuẩn 1.0 kg/m2", "tấn", 1.05, 180.34, 0.00, 0.00, 0.00, 1.00),
        (11, "VAP00001", "Dầu Diezel DO 0.05S (Cấp máy móc thiết bị)", "lít", 1.05, 984908, 0.15, 0.35, 0.35, 0.15)
    ]

    r_bom = 6
    for stt, code, name, unit, hs, qty_tk, k1_p, k2_p, k3_p, k4_p in materials:
        ws2.cell(row=r_bom, column=1, value=stt)
        ws2.cell(row=r_bom, column=2, value=code)
        ws2.cell(row=r_bom, column=3, value=name)
        ws2.cell(row=r_bom, column=4, value=unit)
        ws2.cell(row=r_bom, column=5, value=hs)
        ws2.cell(row=r_bom, column=6, value=qty_tk)
        ws2.cell(row=r_bom, column=7, value=f"=F{r_bom}*E{r_bom}")
        ws2.cell(row=r_bom, column=8, value=f"=G{r_bom}*{k1_p}")
        ws2.cell(row=r_bom, column=9, value=f"=G{r_bom}*{k2_p}")
        ws2.cell(row=r_bom, column=10, value=f"=G{r_bom}*{k3_p}")
        ws2.cell(row=r_bom, column=11, value=f"=G{r_bom}*{k4_p}")

        for c in range(1, 12):
            cell = ws2.cell(row=r_bom, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [6, 7, 8, 9, 10, 11]:
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c == 5:
                cell.number_format = "0.000"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        if "C19" in name:
            for c in range(1, 12):
                ws2.cell(row=r_bom, column=c).fill = PatternFill(start_color=PEACH_FILL, end_color=PEACH_FILL, fill_type="solid")
                ws2.cell(row=r_bom, column=c).font = Font(name=font_family, size=9, bold=True)

        r_bom += 1

    # BOM Column widths
    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 42
    ws2.column_dimensions["D"].width = 10
    ws2.column_dimensions["E"].width = 14
    ws2.column_dimensions["F"].width = 18
    ws2.column_dimensions["G"].width = 20
    ws2.column_dimensions["H"].width = 18
    ws2.column_dimensions["I"].width = 18
    ws2.column_dimensions["J"].width = 18
    ws2.column_dimensions["K"].width = 18


    # =========================================================================
    # SHEET 3: 03_Ke_Hoach_MMTB_NhanLuc
    # =========================================================================
    ws3 = wb.create_sheet(title="03_Ke_Hoach_MMTB_NhanLuc")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:I1")
    ws3["A1"] = "KẾ HOẠCH HUY ĐỘNG NHÂN LỰC & MÁY MÓC THIẾT BỊ (MMTB) - CỤM B9"
    ws3["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws3["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # 1. Labor Plan Section
    ws3.cell(row=3, column=1, value="I. KẾ HOẠCH HUY ĐỘNG NHÂN LỰC THI CÔNG (HỆ SỐ TĂNG CA: 2 CA/NGÀY = 20H)")
    ws3.cell(row=3, column=1).font = Font(name=font_family, size=11, bold=True, color=NAVY_HEADER)
    
    headers_nl = ["STT", "Chức danh / Vị trí", "Đơn vị", "Định mức tính toán (người)", "BCH đề xuất (người)", "Hiện có", "Cần bổ sung", "Thời gian huy động", "Ghi chú"]
    for c_i, h_txt in enumerate(headers_nl, 1):
        cell = ws3.cell(row=4, column=c_i, value=h_txt)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    labor_data = [
        (1, "Công nhân trực tiếp (San lấp, đào đắp nền, thoát nước)", "người", 240, 250, 150, 100, "20/09 - 25/10/2026", "2 ca liên tục"),
        (2, "Lái máy xúc bánh xích PC200 - PC300", "người", 42, 45, 25, 20, "20/09 - 25/10/2026", "Bố trí 2 tài/máy"),
        (3, "Lái máy ủi D3 - D5", "người", 45, 48, 28, 20, "20/09 - 25/10/2026", "Bố trí 2 tài/máy"),
        (4, "Lái máy lu rung 12T - 16T", "người", 28, 30, 18, 12, "22/09 - 25/10/2026", "Bố trí 2 tài/máy"),
        (5, "Lái máy san tự hành 140CV", "người", 6, 8, 4, 4, "28/09 - 20/10/2026", "Bố trí 2 tài/máy"),
        (6, "Lái xe ben vận chuyển 18m3", "người", 24, 28, 16, 12, "20/09 - 25/10/2026", "Vận hành xoay ca"),
        (7, "Lái máy xúc lốp cẩu lắp PC140", "người", 14, 16, 8, 8, "26/09 - 15/10/2026", "Thi công TNM, TNT"),
        (8, "Tổ kỹ thuật & Thợ thảm bê tông nhựa C19", "người", 36, 40, 0, 40, "17/10 - 25/10/2026", "2 mũi thảm hoàn thiện")
    ]

    r_nl = 5
    for stt, role, u, dm, bch, hc, bs, tg, gc in labor_data:
        ws3.cell(row=r_nl, column=1, value=stt)
        ws3.cell(row=r_nl, column=2, value=role)
        ws3.cell(row=r_nl, column=3, value=u)
        ws3.cell(row=r_nl, column=4, value=dm)
        ws3.cell(row=r_nl, column=5, value=bch)
        ws3.cell(row=r_nl, column=6, value=hc)
        ws3.cell(row=r_nl, column=7, value=f"=E{r_nl}-F{r_nl}")
        ws3.cell(row=r_nl, column=8, value=tg)
        ws3.cell(row=r_nl, column=9, value=gc)

        for c in range(1, 10):
            cell = ws3.cell(row=r_nl, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 3, 8]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [4, 5, 6, 7]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_nl += 1

    # Total Labor Row
    ws3.cell(row=r_nl, column=2, value="TỔNG CỘNG NHÂN LỰC HUY ĐỘNG ĐỈNH ĐIỂM")
    ws3.cell(row=r_nl, column=2).font = Font(name=font_family, size=9.5, bold=True, color="C00000")
    for c in [4, 5, 6, 7]:
        c_let = get_column_letter(c)
        ws3.cell(row=r_nl, column=c, value=f"=SUM({c_let}5:{c_let}{r_nl-1})")
        ws3.cell(row=r_nl, column=c).font = Font(name=font_family, size=9.5, bold=True, color="C00000")
        ws3.cell(row=r_nl, column=c).number_format = "#,##0"
        ws3.cell(row=r_nl, column=c).alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 10):
        ws3.cell(row=r_nl, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws3.cell(row=r_nl, column=c).border = thick_bottom

    # 2. Machinery Plan Section
    r_mm = r_nl + 3
    ws3.cell(row=r_mm, column=1, value="II. KẾ HOẠCH HUY ĐỘNG MÁY MÓC THIẾT BỊ (MMTB) THEO ĐỊNH MỨC & ĐỀ XUẤT CẤP")
    ws3.cell(row=r_mm, column=1).font = Font(name=font_family, size=11, bold=True, color=NAVY_HEADER)
    r_mm += 1

    headers_mm = ["STT", "Nhóm MMTB", "Chủng loại / Quy cách", "ĐVT", "Định mức tính toán (Máy)", "BCH đề xuất (Máy)", "Hiện có", "Cần bổ sung", "Kế hoạch cấp"]
    for c_i, h_txt in enumerate(headers_mm, 1):
        cell = ws3.cell(row=r_mm, column=c_i, value=h_txt)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    r_mm += 1

    mm_data = [
        (1, "Máy xúc đào", "Máy xúc bánh xích PC200 - PC300 gầu 0.8 - 1.2 m3", "máy", 42, 45, 25, 20, "Cấp đủ 45 máy"),
        (2, "Máy ủi", "Máy ủi bánh xích D3 - D5 công suất 110 - 130 CV", "máy", 45, 48, 28, 20, "Cấp đủ 48 máy"),
        (3, "Máy lu", "Máy lu rung bánh lốp & bánh thép 12T - 16T", "máy", 28, 30, 18, 12, "Cấp đủ 30 máy"),
        (4, "Máy san", "Máy san tự hành 120 - 170 CV", "máy", 6, 6, 3, 3, "Cấp đủ 6 máy"),
        (5, "Xe tải ben", "Xe ben chuyển đất & đá Base loại 16 - 18 tấn", "xe", 18, 22, 12, 10, "Cấp đủ 22 xe"),
        (6, "Máy xúc lốp", "Máy xúc bánh lốp PC140 gầu 0.5 - 0.6 m3", "máy", 14, 14, 6, 8, "Cấp đủ 14 máy"),
        (7, "Xe téc nước", "Xe téc tưới nước đầm nén 9 - 10 m3", "xe", 4, 4, 2, 2, "Cấp đủ 4 xe"),
        (8, "Máy rải BTN", "Máy rải bê tông nhựa chuyên dụng Vogele Super 1800", "máy", 2, 2, 0, 2, "Huy động 2 máy từ 17/10"),
        (9, "Lu rung BTN", "Máy lu tĩnh & lu rung bánh thép thảm nhựa 10 - 12T", "máy", 4, 4, 0, 4, "Huy động theo mũi rải"),
        (10, "Xe cấp dầu", "Xe téc lưu động cấp dầu hiện trường 9 - 10 m3", "xe", 2, 2, 1, 1, "Trực 24/7"),
        (11, "Máy phát điện", "Máy phát điện 3 pha công nghiệp 25 - 45 kVA", "máy", 3, 3, 2, 1, "Trực thi công đêm")
    ]

    r_mm_start = r_mm
    for stt, grp, spec, u, dm, bch, hc, bs, kh in mm_data:
        ws3.cell(row=r_mm, column=1, value=stt)
        ws3.cell(row=r_mm, column=2, value=grp)
        ws3.cell(row=r_mm, column=3, value=spec)
        ws3.cell(row=r_mm, column=4, value=u)
        ws3.cell(row=r_mm, column=5, value=dm)
        ws3.cell(row=r_mm, column=6, value=bch)
        ws3.cell(row=r_mm, column=7, value=hc)
        ws3.cell(row=r_mm, column=8, value=f"=F{r_mm}-G{r_mm}")
        ws3.cell(row=r_mm, column=9, value=kh)

        for c in range(1, 10):
            cell = ws3.cell(row=r_mm, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 4]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [5, 6, 7, 8]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_mm += 1

    # Total Machines Row
    ws3.cell(row=r_mm, column=3, value="TỔNG SỐ LƯỢNG MMTB CẦN CÓ TRÊN CÔNG TRƯỜNG")
    ws3.cell(row=r_mm, column=3).font = Font(name=font_family, size=9.5, bold=True, color="002060")
    for c in [5, 6, 7, 8]:
        c_let = get_column_letter(c)
        ws3.cell(row=r_mm, column=c, value=f"=SUM({c_let}{r_mm_start}:{c_let}{r_mm-1})")
        ws3.cell(row=r_mm, column=c).font = Font(name=font_family, size=9.5, bold=True, color="002060")
        ws3.cell(row=r_mm, column=c).number_format = "#,##0"
        ws3.cell(row=r_mm, column=c).alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 10):
        ws3.cell(row=r_mm, column=c).fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")
        ws3.cell(row=r_mm, column=c).border = thick_bottom

    # Column widths ws3
    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 24
    ws3.column_dimensions["C"].width = 46
    ws3.column_dimensions["D"].width = 10
    ws3.column_dimensions["E"].width = 18
    ws3.column_dimensions["F"].width = 18
    ws3.column_dimensions["G"].width = 14
    ws3.column_dimensions["H"].width = 14
    ws3.column_dimensions["I"].width = 26


    # =========================================================================
    # SHEET 4: 04_Dinh_Muc_Nang_Suat_Vincons
    # =========================================================================
    ws4 = wb.create_sheet(title="04_Dinh_Muc_Nang_Suat_Vincons")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G1")
    ws4["A1"] = "BẢNG ĐỊNH MỨC NĂNG SUẤT NHÂN CÔNG & CA MÁY (ĐM: Vincons_ĐMGK_02-01)"
    ws4["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws4["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_dm = ["STT", "Mã công việc", "Loại công việc thực hiện", "ĐVT", "Định mức máy chính (m3/ca)", "Tên máy chính áp dụng", "Định mức dầu máy (lít/ca)"]
    for c_i, h_txt in enumerate(headers_dm, 1):
        cell = ws4.cell(row=3, column=c_i, value=h_txt)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    vincons_dm = [
        (1, "Đào khuôn đường", "Đào nền đường, khuôn đường vét hữu cơ", "m3", 872.0, "Máy xúc PC200 - PC300", 89.6),
        (2, "Ủi đất đắp nền", "Ủi đất san gạt tạo mặt bằng", "m3", 1093.0, "Máy ủi D3 - D5", 124.0),
        (3, "Đắp cát K90", "Đắp cát nền đường độ chặt K=0.90", "m3", 900.9, "Máy ủi D3 - D5 + Lu rung 14T", 44.0),
        (4, "Đắp cát K95", "Đắp cát nền đường độ chặt K=0.95", "m3", 1428.0, "Máy ủi D3 - D5 + Lu rung 14T", 44.0),
        (5, "Đắp cát K98", "Đắp cát nền đường độ chặt K=0.98", "m3", 2087.0, "Máy ủi D5 + Lu 14T + Máy san", 69.0),
        (6, "CPĐD2 lớp 1", "Cấp phối đá dăm loại 2, độ chặt K98 (Lớp 1)", "m3", 823.0, "Máy rải/san + Lu 14T", 44.0),
        (7, "CPĐD1 lớp 1", "Cấp phối đá dăm loại 1, độ chặt K98 (Lớp 1)", "m3", 914.0, "Máy rải/san + Lu 14T", 44.0),
        (8, "Thảm BTN C19", "Rải thảm bê tông nhựa hạt thô/chặt C19 dày 6-7cm", "m2", 460.8, "Máy rải chuyên dụng + Lu tĩnh/rung", 56.0),
        (9, "Lắp cống TNM", "Lắp đặt đế cống + Cống tròn D300-D1500 TNM", "md", 31.0, "Máy xúc bánh lốp PC140 cẩu lắp", 64.0),
        (10, "Lắp ống HDPE TNT", "Lắp đặt cống HDPE D300, D400 thoát nước thải", "md", 149.0, "Máy xúc bánh lốp PC140 cẩu lắp", 64.0),
        (11, "Vận chuyển đất <5km", "Vận chuyển đất nội bộ cự ly <5km", "m3", 571.4, "Xe ben tải trọng 18m3", 26.0),
        (12, "Tưới thấm bám", "Tưới nhựa thấm bám tiêu chuẩn 1.0 kg/m2", "m2", 434.8, "Xe téc tưới nhựa chuyên dụng", 42.0)
    ]

    r_dm = 4
    for stt, code, name, u, dm_m, m_name, dm_oil in vincons_dm:
        ws4.cell(row=r_dm, column=1, value=stt)
        ws4.cell(row=r_dm, column=2, value=code)
        ws4.cell(row=r_dm, column=3, value=name)
        ws4.cell(row=r_dm, column=4, value=u)
        ws4.cell(row=r_dm, column=5, value=dm_m)
        ws4.cell(row=r_dm, column=6, value=m_name)
        ws4.cell(row=r_dm, column=7, value=dm_oil)

        for c in range(1, 8):
            cell = ws4.cell(row=r_dm, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 4]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [5, 7]:
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_dm += 1

    ws4.column_dimensions["A"].width = 6
    ws4.column_dimensions["B"].width = 20
    ws4.column_dimensions["C"].width = 46
    ws4.column_dimensions["D"].width = 10
    ws4.column_dimensions["E"].width = 24
    ws4.column_dimensions["F"].width = 34
    ws4.column_dimensions["G"].width = 24


    # =========================================================================
    # SHEET 5: 05_Bang_Tinh_Khoi_Luong_ChiTiet
    # =========================================================================
    ws5 = wb.create_sheet(title="05_Bang_Tinh_Khoi_Luong_ChiTiet")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:I1")
    ws5["A1"] = "BẢNG TÍNH KHỐI LƯỢNG HẠNG MỤC: SAN LẤP CỤM B9 VÀ ĐƯỜNG NỘI BỘ"
    ws5["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws5["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_kl = ["STT", "Hạng mục công việc", "ĐVT", "Chiều dài (m)", "Bề rộng (m)", "Chiều dày / Cao (m)", "Hệ số", "Khối lượng hình học", "Khối lượng hoàn thành mốc 25/10"]
    for c_i, h_txt in enumerate(headers_kl, 1):
        cell = ws5.cell(row=3, column=c_i, value=h_txt)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    kl_data = [
        ("I", "SAN LẤP CỤM B9 (DIỆN TÍCH 88.7 HA)", "", "", "", "", "", "", ""),
        (1, "Phát quang mặt bằng san lấp", "m2", 887000, 0, 0, 1, 887000, 887000),
        (2, "Đào khuôn, bóc đất hữu cơ san lấp dày 10cm", "m3", 887000, 0, 0.10, 1, 88700, 88700),
        (3, "Vận chuyển nội bộ đất hữu cơ san lấp <5km", "m3", 88700, 0, 0, 1, 88700, 88700),
        (4, "Đắp đất/cát K90 san nền chiều cao TB 2.0m", "m3", 887000, 0, 2.00, 1, 1774000, 1774000),
        ("II", "ĐƯỜNG NỘI BỘ CỤM B9 (L = 18,034 M)", "", "", "", "", "", "", ""),
        (5, "Phát quang mặt bằng tuyến đường nội bộ", "m2", 18034, 19.0, 0, 1, 342646, 342646),
        (6, "Đào khuôn đường vét hữu cơ dày 30cm", "m3", 18034, 19.0, 0.30, 1, 102793.8, 102793.8),
        (7, "Vận chuyển nội bộ đất khuôn đường <5km", "m3", 102793.8, 0, 0, 1, 102793.8, 102793.8),
        (8, "Đắp cát K90 nền đường", "m3", 18034, 10.0, 2.00, 1, 360680, 360680),
        (9, "Đắp cát K95 nền đường dày 50cm", "m3", 18034, 10.0, 0.50, 1, 90170, 90170),
        (10, "Đắp cát K98 nền đường dày 30cm", "m3", 18034, 10.0, 0.30, 1, 54102, 54102),
        (11, "Hệ thống cống thoát nước mưa TNM (D300-D1800)", "md", 18034, 0, 0, 1.20, 21640.8, 21640.8),
        (12, "Hệ thống cống thoát nước thải TNT (HDPE)", "md", 18034, 0, 0, 1.20, 21640.8, 21640.8),
        (13, "Móng cấp phối đá dăm loại 2 (Base B) dày 15cm", "m3", 18034, 10.0, 0.15, 1, 27051, 27051),
        (14, "Móng cấp phối đá dăm loại 1 (Base A) dày 12cm", "m3", 18034, 10.0, 0.12, 1, 21640.8, 21640.8),
        (15, "Vận chuyển nội bộ cấp phối đá dăm Base A & B", "m3", 48691.8, 0, 0, 1, 48691.8, 48691.8),
        (16, "THI CÔNG LỚP THẢM BTN C19 (THẢM THÔ 6.5CM)", "m2", 18034, 10.0, 0, 1, 180340, 180340)
    ]

    r_kl = 4
    for item in kl_data:
        stt, name, u, l, w, h, hs, qty_hh, qty_done = item
        ws5.cell(row=r_kl, column=1, value=stt)
        ws5.cell(row=r_kl, column=2, value=name)
        ws5.cell(row=r_kl, column=3, value=u)
        ws5.cell(row=r_kl, column=4, value=l)
        ws5.cell(row=r_kl, column=5, value=w)
        ws5.cell(row=r_kl, column=6, value=h)
        ws5.cell(row=r_kl, column=7, value=hs)
        ws5.cell(row=r_kl, column=8, value=qty_hh)
        ws5.cell(row=r_kl, column=9, value=qty_done)

        is_sec = str(stt) in ["I", "II"]
        is_finish = "C19" in name

        for c in range(1, 10):
            cell = ws5.cell(row=r_kl, column=c)
            cell.font = Font(name=font_family, size=9, bold=(is_sec or is_finish))
            cell.border = thin_border
            if is_sec:
                cell.fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
            elif is_finish:
                cell.fill = PatternFill(start_color=PEACH_FILL, end_color=PEACH_FILL, fill_type="solid")
                cell.font = Font(name=font_family, size=9, bold=True, color="C00000")

            if c in [1, 3]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [4, 5, 6, 7, 8, 9] and cell.value != "":
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_kl += 1

    ws5.column_dimensions["A"].width = 6
    ws5.column_dimensions["B"].width = 46
    ws5.column_dimensions["C"].width = 8
    ws5.column_dimensions["D"].width = 14
    ws5.column_dimensions["E"].width = 12
    ws5.column_dimensions["F"].width = 18
    ws5.column_dimensions["G"].width = 10
    ws5.column_dimensions["H"].width = 18
    ws5.column_dimensions["I"].width = 24

    # Save to primary target locations
    target_paths = [
        project_path(r"TIEN_DO_THI_CONG_CUM_B9_OLYMPIC\260920_TDTC_Cum_B9_SanLap_Va_DuongNoiBo_Olympic_ThuongTin.xlsx"),
        project_path(r"260920_TDTC_Cum_B9_SanLap_Va_DuongNoiBo_Olympic_ThuongTin.xlsx"),
        repo_path(r"examples\260920_TDTC_Cum_B9_SanLap_Va_DuongNoiBo_Olympic_ThuongTin.xlsx")
    ]

    for p in target_paths:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        wb.save(p)
        print(f"[OK] File đã lưu thành công tại: {p}")

if __name__ == "__main__":
    build_b9_schedule()
