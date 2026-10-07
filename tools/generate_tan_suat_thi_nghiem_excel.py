# -*- coding: utf-8 -*-
"""
TỰ ĐỘNG HÓA LẬP HỒ SƠ TẦN SUẤT THÍ NGHIỆM BÊ TÔNG & VẬT LIỆU ĐẦU VÀO
DỰ ÁN: CẦU KM19+529.080 (3 NHỊP DẦM SUPER-T L=38.2M)
Chuẩn theo mẫu hình ảnh người dùng cung cấp (media_1791357332627.png).

Bao gồm chi tiết 100% cấu kiện:
- Toàn bộ 30 cọc khoan nhồi D1.2m (C1..C7 Mố M1, C1..C8 Trụ T1, C1..C8 Trụ T2, C1..C7 Mố M2)
- Toàn bộ Bê tông lót & Bệ mố M1, M2, Bệ trụ T1, T2
- Toàn bộ các đốt thân trụ T1, T2 (Đốt 1, 2, 3) & Xà mũ trụ T1, T2
- Toàn bộ thân mố M1, M2, tường đỉnh, tường cánh
- Toàn bộ 15 phiến dầm chủ Super-T L=38.2m C45
- Dầm ngang, mối nối liên tục nhiệt, bản mặt cầu 3 nhịp, bản quá độ, gờ lan can, khe co giãn.

100% CÔNG THỨC SỐNG:
- Tuổi R7 (=C+7), R28 (=C+28)
- Lũy kế Xi măng, Cát, Đá theo định mức cấp phối
- Tự động nhảy "Lần 1", "Lần 2", "Lần 3"... khi vượt ngưỡng định mức (XM 50T, Cát 200m3, Đá 350m3)
- Cột ngày N/T/N để mở cho người dùng tự điền.
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def build_tan_suat_workbook(output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Theo_Doi_Tan_Suat_Thi_Nghiem"
    ws.views.sheetView[0].showGridLines = True

    # Font definitions
    FONT_NAME = "Arial"
    FONT_HDR = Font(name=FONT_NAME, size=9, bold=True)
    FONT_HDR_RED = Font(name=FONT_NAME, size=9, bold=True, color="C00000")
    FONT_HDR_GREEN = Font(name=FONT_NAME, size=9, bold=True, color="0070C0")
    FONT_REG = Font(name=FONT_NAME, size=8.5)
    FONT_BOLD = Font(name=FONT_NAME, size=8.5, bold=True)
    FONT_RED = Font(name=FONT_NAME, size=8.5, bold=True, color="C00000")
    FONT_GREEN = Font(name=FONT_NAME, size=8.5, bold=True, color="008000")
    FONT_TITLE = Font(name=FONT_NAME, size=12, bold=True, color="1B365D")

    # Border definitions
    THIN_GRAY = Side(style='thin', color='BFBFBF')
    THIN_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
    MEDIUM_BORDER = Border(left=Side(style='medium', color='1B365D'),
                           right=Side(style='medium', color='1B365D'),
                           top=Side(style='medium', color='1B365D'),
                           bottom=Side(style='medium', color='1B365D'))

    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

    FILL_HDR = PatternFill("solid", fgColor="F2F2F2")
    FILL_SUB = PatternFill("solid", fgColor="EAEAEA")
    FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")

    # 1. HÀNG TIÊU ĐỀ BẢNG (Hàng 1 & 2)
    # Merges
    ws.merge_cells("A1:A2") # STT
    ws.merge_cells("B1:B2") # HẠNG MỤC CÔNG VIỆC
    ws.merge_cells("C1:C2") # N/T/N
    ws.merge_cells("D1:D2") # KL/T.tế
    ws.merge_cells("E1:F1") # Tuổi Bt
    ws.merge_cells("G1:I1") # 30Mpa-18±2 (Cọc nhồi)
    ws.merge_cells("J1:L1") # 30Mpa-14±2 (Bệ, thân, mố, trụ)
    ws.merge_cells("M1:O1") # Khối lượng (Lũy kế)
    ws.merge_cells("P1:R1") # Tần suất
    ws.merge_cells("S1:S2") # Số lô
    ws.merge_cells("T1:T2") # KL

    # Values for Row 1
    ws["A1"] = "STT"
    ws["B1"] = "HẠNG MỤC CÔNG VIỆC"
    ws["C1"] = "N/T/N"
    ws["D1"] = "KL/T.tế\n(m3)"
    ws["E1"] = "Tuổi Bt"
    ws["G1"] = "30Mpa-18±2"
    ws["J1"] = "30Mpa-14±2"
    ws["M1"] = "Khối lượng vật tư lũy kế"
    ws["P1"] = "Tần suất lấy mẫu thí nghiệm"
    ws["S1"] = "Số lô"
    ws["T1"] = "KL\n(Tấn)"

    # Values for Row 2
    ws["E2"] = "R7"
    ws["F2"] = "R28"
    ws["G2"] = "Xi măng\n0,445"
    ws["H2"] = "Cát\n0,525"
    ws["I2"] = "Đá\n0,696"
    ws["J2"] = "Xi măng\n0,425"
    ws["K2"] = "Cát\n0,555"
    ws["L2"] = "Đá\n0,700"
    ws["M2"] = "Xi măng\n(Tấn)"
    ws["N2"] = "Cát\n(m3)"
    ws["O2"] = "Đá\n(m3)"
    ws["P2"] = "Xi măng\n50 tấn\nLần"
    ws["Q2"] = "Cát\n200m3\nLần"
    ws["R2"] = "Đá\n350m3\nLần"

    # Style Row 1 & 2
    for r in [1, 2]:
        ws.row_dimensions[r].height = 28 if r == 2 else 24
        for c in range(1, 21):
            cell = ws.cell(r, c)
            cell.alignment = ALIGN_CENTER
            cell.border = THIN_BORDER
            cell.fill = FILL_HDR
            if c in [16]: # P: Xi măng
                cell.font = FONT_HDR_RED
            elif c in [17]: # Q: Cát
                cell.font = FONT_HDR_GREEN
            elif c in [18]: # R: Đá
                cell.font = FONT_HDR_RED
            else:
                cell.font = FONT_HDR

    # DANH SÁCH 100% CẤU KIỆN CHI TIẾT CẦU KM19+529.080
    # Định dạng mỗi item: (Tên cấu kiện, Khối lượng thiết kế m3, Loại cấp phối: "18" hoặc "14" hoặc None cho section)
    items = [
        # --- PHẦN I: CỌC KHOAN NHỒI D1.2M (30 CỌC TOÀN CẦU) ---
        ("SECTION", "I. HẠNG MỤC CỌC KHOAN NHỒI D1.2M (C30 ĐỘ SỤT 18±2CM)"),
        ("Bê tông cọc khoan nhồi - C1 mố M1 (L=20m)", 24.71, "18"),
        ("Bê tông cọc khoan nhồi - C2 mố M1 (L=20m)", 24.71, "18"),
        ("Bê tông cọc khoan nhồi - C3 mố M1 (L=20m)", 24.71, "18"),
        ("Bê tông cọc khoan nhồi - C4 mố M1 (L=20m)", 24.71, "18"),
        ("Bê tông cọc khoan nhồi - C5 mố M1 (L=20m)", 24.71, "18"),
        ("Bê tông cọc khoan nhồi - C6 mố M1 (L=20m)", 24.71, "18"),
        ("Bê tông cọc khoan nhồi - C7 mố M1 (L=20m)", 24.71, "18"),

        ("Bê tông cọc khoan nhồi - C1 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C2 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C3 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C4 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C5 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C6 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C7 trụ T1 (L=40m)", 47.01, "18"),
        ("Bê tông cọc khoan nhồi - C8 trụ T1 (L=40m)", 47.01, "18"),

        ("Bê tông cọc khoan nhồi - C1 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C2 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C3 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C4 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C5 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C6 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C7 trụ T2 (L=30m)", 35.86, "18"),
        ("Bê tông cọc khoan nhồi - C8 trụ T2 (L=30m)", 35.86, "18"),

        ("Bê tông cọc khoan nhồi - C1 mố M2 (L=36m)", 42.83, "18"),
        ("Bê tông cọc khoan nhồi - C2 mố M2 (L=36m)", 42.83, "18"),
        ("Bê tông cọc khoan nhồi - C3 mố M2 (L=36m)", 42.83, "18"),
        ("Bê tông cọc khoan nhồi - C4 mố M2 (L=36m)", 42.83, "18"),
        ("Bê tông cọc khoan nhồi - C5 mố M2 (L=36m)", 42.83, "18"),
        ("Bê tông cọc khoan nhồi - C6 mố M2 (L=36m)", 42.83, "18"),
        ("Bê tông cọc khoan nhồi - C7 mố M2 (L=36m)", 42.83, "18"),

        # --- PHẦN II: BÊ TÔNG BỆ MỐ, BỆ TRỤ ---
        ("SECTION", "II. HẠNG MỤC BÊ TÔNG BỆ MỐ & BỆ TRỤ (C30 ĐỘ SỤT 14±2CM)"),
        ("Bê tông lót móng bệ mố M1 (M100 đá 4x6)", 9.55, "14"),
        ("Bê tông bệ mố M1 (C30 đá 1x2)", 74.50, "14"),
        ("Bê tông lót móng bệ trụ T1 (M100 đá 4x6)", 12.50, "14"),
        ("Bê tông bệ trụ T1 - Đợt 1 (C30 đá 1x2)", 82.50, "14"),
        ("Bê tông bệ trụ T1 - Đợt 2 (C30 đá 1x2)", 82.50, "14"),
        ("Bê tông lót móng bệ trụ T2 (M100 đá 4x6)", 12.50, "14"),
        ("Bê tông bệ trụ T2 - Đợt 1 (C30 đá 1x2)", 82.50, "14"),
        ("Bê tông bệ trụ T2 - Đợt 2 (C30 đá 1x2)", 82.50, "14"),
        ("Bê tông lót móng bệ mố M2 (M100 đá 4x6)", 9.55, "14"),
        ("Bê tông bệ mố M2 (C30 đá 1x2)", 74.50, "14"),

        # --- PHẦN III: THÂN TRỤ, THÂN MỐ, XÀ MŨ ---
        ("SECTION", "III. HẠNG MỤC THÂN TRỤ, THÂN MỐ & XÀ MŨ (CÁC ĐỐT LÊN HOÀN THIỆN)"),
        ("Bê tông thân trụ T1 - Đốt 1 (C30 đá 1x2)", 42.00, "14"),
        ("Bê tông thân trụ T1 - Đốt 2 (C30 đá 1x2)", 42.00, "14"),
        ("Bê tông thân trụ T1 - Đốt 3 (C30 đá 1x2)", 38.00, "14"),
        ("Bê tông xà mũ trụ T1 & đá kê gối (C35/C30 đá 1x2)", 45.00, "14"),

        ("Bê tông thân trụ T2 - Đốt 1 (C30 đá 1x2)", 45.00, "14"),
        ("Bê tông thân trụ T2 - Đốt 2 (C30 đá 1x2)", 45.00, "14"),
        ("Bê tông thân trụ T2 - Đốt 3 (C30 đá 1x2)", 42.00, "14"),
        ("Bê tông xà mũ trụ T2 & đá kê gối (C35/C30 đá 1x2)", 45.00, "14"),

        ("Bê tông tường thân mố M1 - Đợt 1 (C30 đá 1x2)", 45.00, "14"),
        ("Bê tông tường thân mố M1 - Đợt 2 (C30 đá 1x2)", 39.44, "14"),
        ("Bê tông tường đỉnh mố M1 & đá kê gối (C30)", 14.33, "14"),
        ("Bê tông tường cánh mố M1 (C30 đá 1x2)", 10.13, "14"),

        ("Bê tông tường thân mố M2 - Đợt 1 (C30 đá 1x2)", 45.00, "14"),
        ("Bê tông tường thân mố M2 - Đợt 2 (C30 đá 1x2)", 39.44, "14"),
        ("Bê tông tường đỉnh mố M2 & đá kê gối (C30)", 14.33, "14"),
        ("Bê tông tường cánh mố M2 (C30 đá 1x2)", 10.13, "14"),

        # --- PHẦN IV: KẾT CẤU NHỊP DẦM SUPER-T L=38.2M ---
        ("SECTION", "IV. HẠNG MỤC DẦM CHỦ SUPER-T L=38.2M (15 PHIẾN DẦM C45/C35)"),
        ("Bê tông dầm Super-T Nhịp 1 - Phiến D1 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 1 - Phiến D2 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 1 - Phiến D3 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 1 - Phiến D4 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 1 - Phiến D5 (C45 đá 1x2)", 28.99, "14"),

        ("Bê tông dầm Super-T Nhịp 2 - Phiến D1 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 2 - Phiến D2 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 2 - Phiến D3 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 2 - Phiến D4 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 2 - Phiến D5 (C45 đá 1x2)", 28.99, "14"),

        ("Bê tông dầm Super-T Nhịp 3 - Phiến D1 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 3 - Phiến D2 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 3 - Phiến D3 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 3 - Phiến D4 (C45 đá 1x2)", 28.99, "14"),
        ("Bê tông dầm Super-T Nhịp 3 - Phiến D5 (C45 đá 1x2)", 28.99, "14"),

        # --- PHẦN V: DẦM NGANG, BẢN MẶT CẦU & HOÀN THIỆN ---
        ("SECTION", "V. HẠNG MỤC DẦM NGANG, BẢN MẶT CẦU & KẾT CẤU HOÀN THIỆN"),
        ("Bê tông dầm ngang mố M1 & Trụ T1 (C35 đá 1x2)", 6.46, "14"),
        ("Bê tông dầm ngang Trụ T1 & Trụ T2 (C35 đá 1x2)", 6.46, "14"),
        ("Bê tông dầm ngang Trụ T2 & Mố M2 (C35 đá 1x2)", 6.46, "14"),
        ("Bê tông mối nối ướt liên tục nhiệt dầm Super-T (C45)", 8.50, "14"),
        ("Bê tông bản mặt cầu Nhịp 1 (C35 đá 1x2)", 85.00, "14"),
        ("Bê tông bản mặt cầu Nhịp 2 (C35 đá 1x2)", 85.00, "14"),
        ("Bê tông bản mặt cầu Nhịp 3 (C35 đá 1x2)", 85.00, "14"),
        ("Bê tông bản quá độ mố M1 (C30 đá 1x2)", 16.00, "14"),
        ("Bê tông bản quá độ mố M2 (C30 đá 1x2)", 16.00, "14"),
        ("Bê tông gờ lan can mố M1 & Nhịp 1 (C30 đá 1x2)", 9.50, "14"),
        ("Bê tông gờ lan can Nhịp 2, Nhịp 3 & mố M2 (C30)", 9.50, "14"),
        ("Bê tông khe co giãn răng lược D=100mm mố M1 & M2 (C40)", 4.67, "14"),
    ]

    current_row = 3
    stt_counter = 1
    prev_data_row = None

    for item in items:
        if item[0] == "SECTION":
            # Dòng phân cách Section
            ws.row_dimensions[current_row].height = 22
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=20)
            sec_cell = ws.cell(current_row, 1, item[1])
            sec_cell.font = Font(name=FONT_NAME, size=9.5, bold=True, color="1B365D")
            sec_cell.alignment = Alignment(horizontal="left", vertical="center")
            sec_cell.fill = FILL_SECTION
            for c in range(1, 21):
                ws.cell(current_row, c).border = THIN_BORDER
            current_row += 1
            continue

        name, vol, mix_type = item
        r = current_row
        ws.row_dimensions[r].height = 20

        # Col A: STT
        ws.cell(r, 1, stt_counter).alignment = ALIGN_CENTER
        ws.cell(r, 1).font = FONT_REG
        ws.cell(r, 1).border = THIN_BORDER

        # Col B: Hạng mục
        ws.cell(r, 2, name).alignment = ALIGN_LEFT
        ws.cell(r, 2).font = FONT_REG
        ws.cell(r, 2).border = THIN_BORDER

        # Col C: N/T/N (Để trống cho người dùng tự điền ngày)
        cell_c = ws.cell(r, 3)
        cell_c.alignment = ALIGN_CENTER
        cell_c.font = FONT_REG
        cell_c.border = THIN_BORDER
        cell_c.number_format = "DD/MM/YYYY"

        # Col D: Khối lượng thực tế
        cell_d = ws.cell(r, 4, vol)
        cell_d.alignment = ALIGN_RIGHT
        cell_d.font = FONT_REG
        cell_d.border = THIN_BORDER
        cell_d.number_format = "#,##0.00"

        # Col E: R7 = IF(C="","", C + 7)
        cell_e = ws.cell(r, 5, f'=IF(C{r}="","",C{r}+7)')
        cell_e.alignment = ALIGN_CENTER
        cell_e.font = FONT_REG
        cell_e.border = THIN_BORDER
        cell_e.number_format = "DD/MM/YYYY"

        # Col F: R28 = IF(C="","", C + 28)
        cell_f = ws.cell(r, 6, f'=IF(C{r}="","",C{r}+28)')
        cell_f.alignment = ALIGN_CENTER
        cell_f.font = FONT_REG
        cell_f.border = THIN_BORDER
        cell_f.number_format = "DD/MM/YYYY"

        # Định mức cấp phối
        if mix_type == "18":
            ws.cell(r, 7, 0.445).number_format = "0.000" # G: XM
            ws.cell(r, 8, 0.525).number_format = "0.000" # H: Cát
            ws.cell(r, 9, 0.696).number_format = "0.000" # I: Đá
            ws.cell(r, 10, "")
            ws.cell(r, 11, "")
            ws.cell(r, 12, "")
        else: # "14"
            ws.cell(r, 7, "")
            ws.cell(r, 8, "")
            ws.cell(r, 9, "")
            ws.cell(r, 10, 0.425).number_format = "0.000" # J: XM
            ws.cell(r, 11, 0.555).number_format = "0.000" # K: Cát
            ws.cell(r, 12, 0.700).number_format = "0.000" # L: Đá

        for c_mix in range(7, 13):
            c_cell = ws.cell(r, c_mix)
            c_cell.alignment = ALIGN_CENTER
            c_cell.font = FONT_REG
            c_cell.border = THIN_BORDER

        # Cột M, N, O: Khối lượng vật tư tích lũy sống động
        if prev_data_row is None:
            # Dòng dữ liệu đầu tiên
            ws.cell(r, 13, f'=IF(D{r}="","",ROUND(IF(G{r}>0,D{r}*G{r},D{r}*J{r}),2))')
            ws.cell(r, 14, f'=IF(D{r}="","",ROUND(IF(H{r}>0,D{r}*H{r},D{r}*K{r}),2))')
            ws.cell(r, 15, f'=IF(D{r}="","",ROUND(IF(I{r}>0,D{r}*I{r},D{r}*L{r}),2))')
        else:
            p_r = prev_data_row
            ws.cell(r, 13, f'=IF(D{r}="","",ROUND(M{p_r}+IF(G{r}>0,D{r}*G{r},D{r}*J{r}),2))')
            ws.cell(r, 14, f'=IF(D{r}="","",ROUND(N{p_r}+IF(H{r}>0,D{r}*H{r},D{r}*K{r}),2))')
            ws.cell(r, 15, f'=IF(D{r}="","",ROUND(O{p_r}+IF(I{r}>0,D{r}*I{r},D{r}*L{r}),2))')

        for c_kl in range(13, 16):
            cell_kl = ws.cell(r, c_kl)
            cell_kl.alignment = ALIGN_RIGHT
            cell_kl.font = FONT_REG
            cell_kl.border = THIN_BORDER
            cell_kl.number_format = "#,##0.00"

        # Cột P, Q, R: Tần suất thí nghiệm tự động nhảy khi vượt ngưỡng
        # Xi măng 50T / lần
        if prev_data_row is None:
            ws.cell(r, 16, f'=IF(M{r}="","",IF(INT(M{r}/50)>0,"Lần " & INT(M{r}/50),""))')
            ws.cell(r, 17, f'=IF(N{r}="","",IF(INT(N{r}/200)>0,"Lần " & INT(N{r}/200),""))')
            ws.cell(r, 18, f'=IF(O{r}="","",IF(INT(O{r}/350)>0,"Lần " & INT(O{r}/350),""))')
        else:
            p_r = prev_data_row
            ws.cell(r, 16, f'=IF(M{r}="","",IF(INT(M{r}/50)>INT(M{p_r}/50),"Lần " & INT(M{r}/50),""))')
            ws.cell(r, 17, f'=IF(N{r}="","",IF(INT(N{r}/200)>INT(N{p_r}/200),"Lần " & INT(N{r}/200),""))')
            ws.cell(r, 18, f'=IF(O{r}="","",IF(INT(O{r}/350)>INT(O{p_r}/350),"Lần " & INT(O{r}/350),""))')

        ws.cell(r, 16).font = FONT_RED
        ws.cell(r, 17).font = FONT_GREEN
        ws.cell(r, 18).font = FONT_RED
        for c_ts in range(16, 19):
            ws.cell(r, c_ts).alignment = ALIGN_CENTER
            ws.cell(r, c_ts).border = THIN_BORDER

        # Cột S: Số lô (để trống)
        ws.cell(r, 19, "").alignment = ALIGN_CENTER
        ws.cell(r, 19).font = FONT_REG
        ws.cell(r, 19).border = THIN_BORDER

        # Cột T: Khối lượng lô Tấn (để trống)
        ws.cell(r, 20, "").alignment = ALIGN_RIGHT
        ws.cell(r, 20).font = FONT_REG
        ws.cell(r, 20).border = THIN_BORDER
        ws.cell(r, 20).number_format = "#,##0.00"

        prev_data_row = r
        stt_counter += 1
        current_row += 1

    # Dòng Tổng cộng cuối bảng
    r_tot = current_row
    ws.row_dimensions[r_tot].height = 24
    ws.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=3)
    ws.cell(r_tot, 1, "TỔNG CỘNG TOÀN CẦU").alignment = ALIGN_CENTER
    ws.cell(r_tot, 1).font = FONT_HDR
    ws.cell(r_tot, 4, f"=SUM(D4:D{r_tot-1})").alignment = ALIGN_RIGHT
    ws.cell(r_tot, 4).font = FONT_HDR
    ws.cell(r_tot, 4).number_format = "#,##0.00"

    # Lũy kế cuối
    if prev_data_row:
        ws.cell(r_tot, 13, f"=M{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 14, f"=N{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 15, f"=O{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 13).font = FONT_HDR
        ws.cell(r_tot, 14).font = FONT_HDR
        ws.cell(r_tot, 15).font = FONT_HDR
        ws.cell(r_tot, 13).alignment = ALIGN_RIGHT
        ws.cell(r_tot, 14).alignment = ALIGN_RIGHT
        ws.cell(r_tot, 15).alignment = ALIGN_RIGHT

    for c in range(1, 21):
        cell = ws.cell(r_tot, c)
        cell.border = THIN_BORDER
        cell.fill = PatternFill("solid", fgColor="FFF2CC")

    # Set Widths
    col_widths = {
        1: 6,   # STT
        2: 46,  # HẠNG MỤC CÔNG VIỆC
        3: 13,  # N/T/N
        4: 12,  # KL/T.tế
        5: 13,  # R7
        6: 13,  # R28
        7: 10,  # XM 18
        8: 10,  # Cát 18
        9: 10,  # Đá 18
        10: 10, # XM 14
        11: 10, # Cát 14
        12: 10, # Đá 14
        13: 13, # XM lũy kế
        14: 13, # Cát lũy kế
        15: 13, # Đá lũy kế
        16: 12, # Tần suất XM
        17: 12, # Tần suất Cát
        18: 12, # Tần suất Đá
        19: 18, # Số lô
        20: 10  # KL lô
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze Panes tại hàng 3
    ws.freeze_panes = "C3"

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    wb.close()
    print(f"[OK] Đã tạo thành công Bảng theo dõi Tần suất thí nghiệm: {output_path}")

if __name__ == "__main__":
    out_file = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Cau_Km19.xlsx"
    build_tan_suat_workbook(out_file)
