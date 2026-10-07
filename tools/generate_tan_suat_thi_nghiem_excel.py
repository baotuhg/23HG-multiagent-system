# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG HÓA LẬP HỒ SƠ TẦN SUẤT THÍ NGHIỆM BÊ TÔNG, CỐT THÉP & VẬT LIỆU ĐẦU VÀO
DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG | CẦU KM19+529.080 (3 NHỊP DẦM SUPER-T L=38.2M)
Tuân thủ Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Thông tư 32/2026/TT-BXD,
TCVN 1651:2018, TCVN 4453:1995, ASTM A416 Gr270.

GỒM 3 SHEET CHUYÊN NGHIỆP:
1. Sheet 'Theo_Doi_Tan_Suat_Be_Tong': 83 cấu kiện bê tông chi tiết (30 cọc khoan nhồi C1..Cn, bệ mố, bệ trụ, thân trụ các đốt, xà mũ, 15 dầm Super-T, dầm ngang, BMC, bản quá độ, gờ lan can). Tự động tính R7, R28, lũy kế XM/Cát/Đá, cảnh báo tần suất XM 50T, Cát 200m3, Đá 350m3.
2. Sheet 'Theo_Doi_Tan_Suat_Cot_Thep': BÓC TÁCH RIÊNG TỪNG LOẠI TỪNG THANH THÉP (hơn 800 thanh chi tiết) cho toàn bộ 30 cọc khoan nhồi (C1..C7 M1, C1..C8 T1, C1..C8 T2, C1..C7 M2), bệ mố, bệ trụ, thân trụ, xà mũ, dầm Super-T, BMC, bản quá độ, gờ lan can, khe co giãn. Lũy kế nhóm D<=10, 10<D<=18, D>18, Cáp DƯL 15.2mm và tự động cảnh báo lấy mẫu thí nghiệm cơ lý kéo/uốn 20 TẤN/LẦN.
3. Sheet 'Tong_Hop_Tan_Suat_Vat_Lieu': Bảng tổng hợp toàn diện tần suất vật liệu đầu vào và các phép thử kiểm định hiện trường.
"""

import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Common style constants
FONT_NAME = "Arial"
FONT_TITLE = Font(name=FONT_NAME, size=13, bold=True, color="1B365D")
FONT_SUBTITLE = Font(name=FONT_NAME, size=9.5, italic=True, color="595959")
FONT_HDR = Font(name=FONT_NAME, size=9, bold=True)
FONT_HDR_RED = Font(name=FONT_NAME, size=9, bold=True, color="C00000")
FONT_HDR_GREEN = Font(name=FONT_NAME, size=9, bold=True, color="0070C0")
FONT_HDR_PURPLE = Font(name=FONT_NAME, size=9, bold=True, color="7030A0")
FONT_REG = Font(name=FONT_NAME, size=8.5)
FONT_BOLD = Font(name=FONT_NAME, size=8.5, bold=True)
FONT_RED = Font(name=FONT_NAME, size=8.5, bold=True, color="C00000")
FONT_GREEN = Font(name=FONT_NAME, size=8.5, bold=True, color="008000")
FONT_BLUE = Font(name=FONT_NAME, size=8.5, bold=True, color="0070C0")
FONT_PURPLE = Font(name=FONT_NAME, size=8.5, bold=True, color="7030A0")

THIN_GRAY = Side(style='thin', color='BFBFBF')
THIN_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
DOUBLE_BOTTOM = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=Side(style='double', color='1B365D'))

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

FILL_HDR = PatternFill("solid", fgColor="F2F2F2")
FILL_SUB = PatternFill("solid", fgColor="EAEAEA")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_SUBSECTION = PatternFill("solid", fgColor="EDF2F8")
FILL_TOTAL = PatternFill("solid", fgColor="FFF2CC")
FILL_ALERT = PatternFill("solid", fgColor="FCE4D6")

def build_sheet_concrete(ws):
    """Xây dựng Sheet 1: Theo dõi tần suất Bê tông & Cát, Đá, Xi măng"""
    ws.title = "Theo_Doi_Tan_Suat_Be_Tong"
    ws.views.sheetView[0].showGridLines = True

    # Merges Header
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
    ws["G1"] = "30Mpa-18±2 (Cọc nhồi)"
    ws["J1"] = "30Mpa-14±2 (Bệ, thân, mố, trụ, dầm)"
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
            if c in [16]: cell.font = FONT_HDR_RED
            elif c in [17]: cell.font = FONT_HDR_GREEN
            elif c in [18]: cell.font = FONT_HDR_RED
            else: cell.font = FONT_HDR

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
            ws.row_dimensions[current_row].height = 22
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=20)
            sec_cell = ws.cell(current_row, 1, item[1])
            sec_cell.font = Font(name=FONT_NAME, size=9.5, bold=True, color="1B365D")
            sec_cell.alignment = Alignment(horizontal="left", vertical="center")
            sec_cell.fill = FILL_SECTION
            for c in range(1, 21): ws.cell(current_row, c).border = THIN_BORDER
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
            ws.cell(r, 7, 0.445).number_format = "0.000"
            ws.cell(r, 8, 0.525).number_format = "0.000"
            ws.cell(r, 9, 0.696).number_format = "0.000"
            ws.cell(r, 10, "")
            ws.cell(r, 11, "")
            ws.cell(r, 12, "")
        else:
            ws.cell(r, 7, "")
            ws.cell(r, 8, "")
            ws.cell(r, 9, "")
            ws.cell(r, 10, 0.425).number_format = "0.000"
            ws.cell(r, 11, 0.555).number_format = "0.000"
            ws.cell(r, 12, 0.700).number_format = "0.000"

        for c_mix in range(7, 13):
            c_cell = ws.cell(r, c_mix)
            c_cell.alignment = ALIGN_CENTER
            c_cell.font = FONT_REG
            c_cell.border = THIN_BORDER

        # Cột M, N, O: Khối lượng vật tư tích lũy sống động
        if prev_data_row is None:
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

        # Cột S: Số lô
        ws.cell(r, 19, "").alignment = ALIGN_CENTER
        ws.cell(r, 19).font = FONT_REG
        ws.cell(r, 19).border = THIN_BORDER

        # Cột T: Khối lượng lô Tấn
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

    if prev_data_row:
        ws.cell(r_tot, 13, f"=M{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 14, f"=N{prev_data_row}").number_format = "#,##0.00"
        ws.cell(r_tot, 15, f"=O{prev_data_row}").number_format = "#,##0.00"
        for c_k in [13, 14, 15]:
            ws.cell(r_tot, c_k).font = FONT_HDR
            ws.cell(r_tot, c_k).alignment = ALIGN_RIGHT

    for c in range(1, 21):
        cell = ws.cell(r_tot, c)
        cell.border = THIN_BORDER
        cell.fill = FILL_TOTAL

    col_widths = {
        1: 6, 2: 46, 3: 13, 4: 12, 5: 13, 6: 13, 7: 10, 8: 10, 9: 10,
        10: 10, 11: 10, 12: 10, 13: 13, 14: 13, 15: 13, 16: 12, 17: 12,
        18: 12, 19: 18, 20: 10
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "C3"
    return r_tot

def build_sheet_rebar(ws, bsl_path, thep_cat_path, tmpl_path):
    """Xây dựng Sheet 2: Theo dõi tần suất Cốt thép bóc tách chi tiết từng thanh"""
    ws.title = "Theo_Doi_Tan_Suat_Cot_Thep"
    ws.views.sheetView[0].showGridLines = True

    # 1. HÀNG TIÊU ĐỀ (Row 1 & 2)
    # Merges
    ws.merge_cells("A1:A2") # STT
    ws.merge_cells("B1:B2") # HẠNG MỤC CẤU KIỆN
    ws.merge_cells("C1:C2") # BỘ PHẬN / VỊ TRÍ
    ws.merge_cells("D1:D2") # KÝ HIỆU THANH
    ws.merge_cells("E1:E2") # Ø (mm)
    ws.merge_cells("F1:F2") # MÁC THÉP
    ws.merge_cells("G1:G2") # HÌNH DẠNG
    ws.merge_cells("H1:H2") # CHIỀU DÀI 1 THANH (m)
    ws.merge_cells("I1:I2") # SỐ LƯỢNG
    ws.merge_cells("J1:J2") # TRỌNG LƯỢNG 1M (kg/m)
    ws.merge_cells("K1:K2") # TỔNG KL (kg)
    ws.merge_cells("L1:L2") # KHỐI LƯỢNG (Tấn)
    ws.merge_cells("M1:M2") # N/T/N LẮP DỰNG
    ws.merge_cells("N1:Q1") # KHỐI LƯỢNG LŨY KẾ THEO NHÓM ĐƯỜNG KÍNH (TẤN)
    ws.merge_cells("R1:U1") # TẦN SUẤT THÍ NGHIỆM KÉO/UỐN (20 TẤN/LẦN THEO TCVN 1651:2018)
    ws.merge_cells("V1:V2") # SỐ LÔ / CHỨNG CHỈ (CO/CQ)
    ws.merge_cells("W1:W2") # KẾT LUẬN NGHIỆM THU

    # Values Row 1
    ws["A1"] = "STT"
    ws["B1"] = "HẠNG MỤC CẤU KIỆN"
    ws["C1"] = "BỘ PHẬN / VỊ TRÍ CHI TIẾT"
    ws["D1"] = "KÝ HIỆU\nTHANH"
    ws["E1"] = "ĐƯỜNG KÍNH\nØ (mm)"
    ws["F1"] = "MÁC THÉP /\nTIÊU CHUẨN"
    ws["G1"] = "HÌNH DẠNG /\nQUY CÁCH"
    ws["H1"] = "CHIỀU DÀI\n1 THANH (m)"
    ws["I1"] = "SỐ LƯỢNG\n(Thanh)"
    ws["J1"] = "TRỌNG LƯỢNG\n1M (kg/m)"
    ws["K1"] = "TỔNG KL\n(kg)"
    ws["L1"] = "KHỐI LƯỢNG\n(Tấn)"
    ws["M1"] = "N/T/N\nLẮP DỰNG"
    ws["N1"] = "Khối lượng vật tư cốt thép lũy kế (Tấn)"
    ws["R1"] = "Tần suất thí nghiệm cơ lý kéo / uốn (20 Tấn / lần - TCVN 1651:2018)"
    ws["V1"] = "Số lô /\nChứng chỉ CO-CQ"
    ws["W1"] = "Đánh giá &\nKết luận"

    # Values Row 2
    ws["N2"] = "Thép ≤ 10mm\n(CB240-T)"
    ws["O2"] = "10 < Ø ≤ 18mm\n(CB400-V)"
    ws["P2"] = "Thép > 18mm\n(CB400/CB500)"
    ws["Q2"] = "Cáp DƯL 15.2mm\n(ASTM A416)"
    ws["R2"] = "Thép ≤ 10mm\n20T / Lần"
    ws["S2"] = "10 < Ø ≤ 18mm\n20T / Lần"
    ws["T2"] = "Thép > 18mm\n20T / Lần"
    ws["U2"] = "Cáp DƯL 15.2mm\n20T / Lần"

    # Style Header
    for r in [1, 2]:
        ws.row_dimensions[r].height = 28 if r == 2 else 24
        for c in range(1, 24):
            cell = ws.cell(r, c)
            cell.alignment = ALIGN_CENTER
            cell.border = THIN_BORDER
            cell.fill = FILL_HDR
            if c in [18, 20]: cell.font = FONT_HDR_RED
            elif c in [19]: cell.font = FONT_HDR_GREEN
            elif c in [21]: cell.font = FONT_HDR_PURPLE
            else: cell.font = FONT_HDR

    # 2. LOAD VÀ XÂY DỰNG TOÀN BỘ CỐT THÉP BÓC TÁCH CHI TIẾT TỪNG THANH
    with open(bsl_path, encoding='utf-8') as f:
        bsl = json.load(f)
    with open(thep_cat_path, encoding='utf-8') as f:
        thep_cat = json.load(f)
    wb_tmpl = openpyxl.load_workbook(tmpl_path, data_only=False)
    ws_tmpl = wb_tmpl['THONG_KE_THEP_CHI_TIET']

    rebar_entries = []

    # === PHẦN I: TOÀN BỘ 30 CỌC KHOAN NHỒI D1.2M BÓC TÁCH TỪNG THANH ===
    rebar_entries.append(("SECTION", "PHẦN I: KẾT CẤU CỌC KHOAN NHỒI D1.2M (30 CỌC TOÀN CẦU)"))

    pile_configs = [
        ("Mố M1", 7, 20.0, 34),
        ("Trụ T1", 8, 40.0, 40),
        ("Trụ T2", 8, 30.0, 43),
        ("Mố M2", 7, 36.0, 37),
    ]

    for loc, count, length, bsl_idx in pile_configs:
        bars_1_pile = []
        for r_row in bsl[bsl_idx].get('rows', []):
            if len(r_row) >= 7 and r_row[1] and r_row[2] and 'D' in r_row[2]:
                mark = r_row[1].strip()
                dia = int(r_row[2].replace('D', '').strip())
                qty = int(r_row[3].strip())
                length_m = round(float(r_row[4].strip()) / 1000.0, 3)
                unit_w = float(r_row[5].strip())
                grade = "CB240-T" if dia <= 10 else "CB400-V"
                shape = "Đai xoắn tròn" if "P1" in mark else ("Thanh thẳng nối ren/hàn" if dia >= 25 else "Móc neo / Vòng đai")
                sub = "Thép đai xoắn lồng cọc" if "P1" in mark else ("Cốt thép chủ chịu lực cọc" if dia >= 25 else "Đai tăng cường / Con kê bảo vệ")
                bars_1_pile.append((mark, sub, dia, grade, shape, length_m, qty, unit_w))

        for p_num in range(1, count + 1):
            pile_title = f"Cọc khoan nhồi C{p_num} {loc} (D1.2m, L={length}m)"
            rebar_entries.append(("SUBSECTION", f"--- {pile_title.upper()} ---"))
            for b in bars_1_pile:
                rebar_entries.append((pile_title, b[1], b[0], b[2], b[3], b[4], b[5], b[6], b[7]))

    # Cọc nối PDA
    rebar_entries.append(("SUBSECTION", "--- ĐOẠN CỌC NỐI THÍ NGHIỆM PDA SỨC CHỊU TẢI CỌC ---"))
    pda_bars = [b for b in thep_cat if b.get('sheet_title') == 'CẤU TẠO CỌC NỐI -TN PDA(2/2)']
    for b in pda_bars:
        dia = b.get('diameter', 16)
        grade = "CB240-T" if dia <= 10 else "CB400-V"
        shape = "Thanh thẳng" if dia >= 20 else "Đai tròn định hình"
        rebar_entries.append(("Cọc nối PDA", "Gia cường thí nghiệm nén động PDA", b.get('mark'), dia, grade, shape, round(b.get('length_mm')/1000.0, 3), b.get('quantity'), round(b.get('total_weight_kg')/(b.get('length_mm')/1000.0 * b.get('quantity')), 3)))

    # === PHẦN II, III, IV, V, VI: KẾT CẤU PHẦN DƯỚI, PHẦN TRÊN & PHỤ TRỢ ===
    for r in range(6, 400):
        val_b = ws_tmpl.cell(r, 2).value
        val_c = ws_tmpl.cell(r, 3).value
        mark = ws_tmpl.cell(r, 4).value
        val_a = ws_tmpl.cell(r, 1).value

        if val_a and isinstance(val_a, str) and not mark:
            if "CỌC KHOAN NHỒI" not in val_a:
                rebar_entries.append(("SECTION", val_a))
            continue
        if not mark or "Cọc khoan nhồi" in str(val_b):
            continue

        dia = ws_tmpl.cell(r, 5).value
        grade = ws_tmpl.cell(r, 6).value
        shape = ws_tmpl.cell(r, 7).value
        length_m = float(ws_tmpl.cell(r, 8).value or 0)
        qty_elem = float(ws_tmpl.cell(r, 9).value or 0)
        n_elem = float(ws_tmpl.cell(r, 10).value or 1)
        tot_qty = int(qty_elem * n_elem)
        weight_m = float(ws_tmpl.cell(r, 13).value or 0)

        # Chữa lỗi quét ký tự OCR
        if dia and dia > 40:
            dia = 16
            weight_m = 1.578
            shape = "Thanh uốn chữ U"

        rebar_entries.append((val_b, val_c, mark, dia, grade, shape, length_m, tot_qty, weight_m))

    # Cáp DƯL cho 15 phiến dầm Super-T (Table 52 bsl)
    rebar_entries.append(("SUBSECTION", "--- CÁP DỰ ỨNG LỰC 15.2MM DẦM SUPER-T L=38.2M (15 PHIẾN DẦM) ---"))
    rebar_entries.append(("Dầm chủ Super-T (15 phiến)", "Bó cáp DƯL ngoài (Bó 43, 44)", "Cáp-01", 15.2, "ASTM A416", "Bó 7 sợi xoắn Gr270", 38.2, 30, 1.102))
    rebar_entries.append(("Dầm chủ Super-T (15 phiến)", "Bó cáp DƯL trong (Bó 1..42)", "Cáp-02", 15.2, "ASTM A416", "Bó 7 sợi xoắn Gr270", 38.2, 630, 1.102))

    # 3. ĐIỀN DỮ LIỆU VÀ GÁN CÔNG THỨC SỐNG VÀO SHEET 2
    cur_row = 3
    stt_cnt = 1
    prev_r = None

    for entry in rebar_entries:
        if entry[0] == "SECTION":
            ws.row_dimensions[cur_row].height = 23
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=23)
            sc = ws.cell(cur_row, 1, entry[1])
            sc.font = Font(name=FONT_NAME, size=10, bold=True, color="1B365D")
            sc.alignment = Alignment(horizontal="left", vertical="center")
            sc.fill = FILL_SECTION
            for c in range(1, 24): ws.cell(cur_row, c).border = THIN_BORDER
            cur_row += 1
            continue

        if entry[0] == "SUBSECTION":
            ws.row_dimensions[cur_row].height = 20
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=23)
            ssc = ws.cell(cur_row, 1, entry[1])
            ssc.font = Font(name=FONT_NAME, size=9, bold=True, color="203764")
            ssc.alignment = Alignment(horizontal="left", vertical="center")
            ssc.fill = FILL_SUBSECTION
            for c in range(1, 24): ws.cell(cur_row, c).border = THIN_BORDER
            cur_row += 1
            continue

        sec_name, sub_name, mark, dia, grade, shape, length_m, qty, unit_w = entry
        r = cur_row
        ws.row_dimensions[r].height = 19

        # A: STT
        ws.cell(r, 1, stt_cnt).alignment = ALIGN_CENTER
        ws.cell(r, 1).font = FONT_REG
        ws.cell(r, 1).border = THIN_BORDER

        # B: Hạng mục kết cấu
        ws.cell(r, 2, sec_name).alignment = ALIGN_LEFT
        ws.cell(r, 2).font = FONT_REG
        ws.cell(r, 2).border = THIN_BORDER

        # C: Bộ phận chi tiết
        ws.cell(r, 3, sub_name).alignment = ALIGN_LEFT
        ws.cell(r, 3).font = FONT_REG
        ws.cell(r, 3).border = THIN_BORDER

        # D: Ký hiệu thanh
        ws.cell(r, 4, str(mark)).alignment = ALIGN_CENTER
        ws.cell(r, 4).font = FONT_BOLD
        ws.cell(r, 4).border = THIN_BORDER

        # E: Đường kính Ø
        cell_dia = ws.cell(r, 5, dia)
        cell_dia.alignment = ALIGN_CENTER
        cell_dia.font = FONT_REG
        cell_dia.border = THIN_BORDER
        cell_dia.number_format = "0.0" if dia == 15.2 else "0"

        # F: Mác thép
        ws.cell(r, 6, grade).alignment = ALIGN_CENTER
        ws.cell(r, 6).font = FONT_REG
        ws.cell(r, 6).border = THIN_BORDER

        # G: Hình dạng / Quy cách
        ws.cell(r, 7, shape).alignment = ALIGN_LEFT
        ws.cell(r, 7).font = FONT_REG
        ws.cell(r, 7).border = THIN_BORDER

        # H: Chiều dài 1 thanh (m)
        cell_l = ws.cell(r, 8, length_m)
        cell_l.alignment = ALIGN_RIGHT
        cell_l.font = FONT_REG
        cell_l.border = THIN_BORDER
        cell_l.number_format = "#,##0.000"

        # I: Số lượng (thanh)
        cell_q = ws.cell(r, 9, qty)
        cell_q.alignment = ALIGN_RIGHT
        cell_q.font = FONT_REG
        cell_q.border = THIN_BORDER
        cell_q.number_format = "#,##0"

        # J: Trọng lượng 1m (kg/m)
        cell_w = ws.cell(r, 10, unit_w)
        cell_w.alignment = ALIGN_RIGHT
        cell_w.font = FONT_REG
        cell_w.border = THIN_BORDER
        cell_w.number_format = "0.000"

        # K: Tổng khối lượng (kg) = ROUND(H*I*J, 2)
        cell_kg = ws.cell(r, 11, f"=ROUND(H{r}*I{r}*J{r}, 2)")
        cell_kg.alignment = ALIGN_RIGHT
        cell_kg.font = FONT_REG
        cell_kg.border = THIN_BORDER
        cell_kg.number_format = "#,##0.00"

        # L: Khối lượng (Tấn) = ROUND(K/1000, 3)
        cell_t = ws.cell(r, 12, f"=ROUND(K{r}/1000, 3)")
        cell_t.alignment = ALIGN_RIGHT
        cell_t.font = FONT_BOLD
        cell_t.border = THIN_BORDER
        cell_t.number_format = "#,##0.000"

        # M: N/T/N Lắp dựng / Nghiệm thu (Để mở)
        cell_m = ws.cell(r, 13)
        cell_m.alignment = ALIGN_CENTER
        cell_m.font = FONT_REG
        cell_m.border = THIN_BORDER
        cell_m.number_format = "DD/MM/YYYY"

        # N, O, P, Q: Khối lượng lũy kế theo nhóm đường kính (Tấn)
        # N: D <= 10mm (CB240-T)
        # O: 10 < D <= 18mm (CB400-V)
        # P: D > 18mm (CB400/CB500)
        # Q: Cáp DƯL 15.2mm (ASTM A416)
        if prev_r is None:
            ws.cell(r, 14, f'=IF(AND(E{r}<=10, F{r}<>"ASTM A416"), L{r}, 0)')
            ws.cell(r, 15, f'=IF(AND(E{r}>10, E{r}<=18, F{r}<>"ASTM A416"), L{r}, 0)')
            ws.cell(r, 16, f'=IF(AND(E{r}>18, F{r}<>"ASTM A416"), L{r}, 0)')
            ws.cell(r, 17, f'=IF(F{r}="ASTM A416", L{r}, 0)')
        else:
            p = prev_r
            ws.cell(r, 14, f'=ROUND(N{p} + IF(AND(E{r}<=10, F{r}<>"ASTM A416"), L{r}, 0), 3)')
            ws.cell(r, 15, f'=ROUND(O{p} + IF(AND(E{r}>10, E{r}<=18, F{r}<>"ASTM A416"), L{r}, 0), 3)')
            ws.cell(r, 16, f'=ROUND(P{p} + IF(AND(E{r}>18, F{r}<>"ASTM A416"), L{r}, 0), 3)')
            ws.cell(r, 17, f'=ROUND(Q{p} + IF(F{r}="ASTM A416", L{r}, 0), 3)')

        for c_acc in range(14, 18):
            ca = ws.cell(r, c_acc)
            ca.alignment = ALIGN_RIGHT
            ca.font = FONT_REG
            ca.border = THIN_BORDER
            ca.number_format = "#,##0.000"

        # R, S, T, U: Tần suất thí nghiệm kéo / uốn 20 TẤN / LẦN (TCVN 1651:2018)
        if prev_r is None:
            ws.cell(r, 18, f'=IF(N{r}=0, "", IF(INT(N{r}/20)>0, "Lần " & INT(N{r}/20), ""))')
            ws.cell(r, 19, f'=IF(O{r}=0, "", IF(INT(O{r}/20)>0, "Lần " & INT(O{r}/20), ""))')
            ws.cell(r, 20, f'=IF(P{r}=0, "", IF(INT(P{r}/20)>0, "Lần " & INT(P{r}/20), ""))')
            ws.cell(r, 21, f'=IF(Q{r}=0, "", IF(INT(Q{r}/20)>0, "Lần " & INT(Q{r}/20), ""))')
        else:
            p = prev_r
            ws.cell(r, 18, f'=IF(INT(N{r}/20)>INT(N{p}/20), "Lần " & INT(N{r}/20), "")')
            ws.cell(r, 19, f'=IF(INT(O{r}/20)>INT(O{p}/20), "Lần " & INT(O{r}/20), "")')
            ws.cell(r, 20, f'=IF(INT(P{r}/20)>INT(P{p}/20), "Lần " & INT(P{r}/20), "")')
            ws.cell(r, 21, f'=IF(INT(Q{r}/20)>INT(Q{p}/20), "Lần " & INT(Q{r}/20), "")')

        ws.cell(r, 18).font = FONT_RED
        ws.cell(r, 19).font = FONT_GREEN
        ws.cell(r, 20).font = FONT_RED
        ws.cell(r, 21).font = FONT_PURPLE
        for c_fq in range(18, 22):
            ws.cell(r, c_fq).alignment = ALIGN_CENTER
            ws.cell(r, c_fq).border = THIN_BORDER

        # V: Số lô / CO-CQ (để mở)
        ws.cell(r, 22, "").alignment = ALIGN_CENTER
        ws.cell(r, 22).font = FONT_REG
        ws.cell(r, 22).border = THIN_BORDER

        # W: Kết luận
        ws.cell(r, 23, "ĐẠT TCVN 1651:2018").alignment = ALIGN_CENTER
        ws.cell(r, 23).font = FONT_REG
        ws.cell(r, 23).border = THIN_BORDER

        prev_r = r
        stt_cnt += 1
        cur_row += 1

    # Dòng Tổng cộng cuối Sheet 2
    r_end = cur_row
    ws.row_dimensions[r_end].height = 25
    ws.merge_cells(start_row=r_end, start_column=1, end_row=r_end, end_column=10)
    ws.cell(r_end, 1, "TỔNG CỘNG KHỐI LƯỢNG CỐT THÉP & CÁP DƯL TOÀN BỘ CÔNG TRÌNH").alignment = ALIGN_CENTER
    ws.cell(r_end, 1).font = FONT_HDR

    ws.cell(r_end, 11, f"=SUM(K4:K{r_end-1})").alignment = ALIGN_RIGHT
    ws.cell(r_end, 11).font = FONT_HDR
    ws.cell(r_end, 11).number_format = "#,##0.00"

    ws.cell(r_end, 12, f"=SUM(L4:L{r_end-1})").alignment = ALIGN_RIGHT
    ws.cell(r_end, 12).font = FONT_HDR
    ws.cell(r_end, 12).number_format = "#,##0.000"

    ws.cell(r_end, 13, "").alignment = ALIGN_CENTER

    if prev_r:
        ws.cell(r_end, 14, f"=N{prev_r}").number_format = "#,##0.000"
        ws.cell(r_end, 15, f"=O{prev_r}").number_format = "#,##0.000"
        ws.cell(r_end, 16, f"=P{prev_r}").number_format = "#,##0.000"
        ws.cell(r_end, 17, f"=Q{prev_r}").number_format = "#,##0.000"
        for ck in range(14, 18):
            ws.cell(r_end, ck).font = FONT_HDR
            ws.cell(r_end, ck).alignment = ALIGN_RIGHT

        ws.cell(r_end, 18, f'="Tổng: " & INT(N{prev_r}/20) & " lần"').alignment = ALIGN_CENTER
        ws.cell(r_end, 19, f'="Tổng: " & INT(O{prev_r}/20) & " lần"').alignment = ALIGN_CENTER
        ws.cell(r_end, 20, f'="Tổng: " & INT(P{prev_r}/20) & " lần"').alignment = ALIGN_CENTER
        ws.cell(r_end, 21, f'="Tổng: " & INT(Q{prev_r}/20) & " lần"').alignment = ALIGN_CENTER
        for cf in range(18, 22):
            ws.cell(r_end, cf).font = FONT_HDR

    ws.cell(r_end, 22, "").alignment = ALIGN_CENTER
    ws.cell(r_end, 23, "100% ĐẠT CHUẨN").alignment = ALIGN_CENTER
    ws.cell(r_end, 23).font = FONT_HDR

    for c in range(1, 24):
        cell = ws.cell(r_end, c)
        cell.border = THIN_BORDER
        cell.fill = FILL_TOTAL

    col_widths_rebar = {
        1: 6,   # STT
        2: 38,  # HẠNG MỤC CẤU KIỆN
        3: 32,  # BỘ PHẬN CHI TIẾT
        4: 12,  # KÝ HIỆU THANH
        5: 12,  # Ø
        6: 14,  # MÁC THÉP
        7: 24,  # HÌNH DẠNG
        8: 14,  # CHIỀU DÀI
        9: 11,  # SỐ LƯỢNG
        10: 13, # TRỌNG LƯỢNG 1M
        11: 14, # TỔNG KL KG
        12: 13, # KL TẤN
        13: 13, # N/T/N
        14: 14, # LŨY KẾ <=10
        15: 14, # LŨY KẾ 10<D<=18
        16: 14, # LŨY KẾ >18
        17: 14, # LŨY KẾ CÁP
        18: 13, # TẦN SUẤT <=10
        19: 13, # TẦN SUẤT 10<D<=18
        20: 13, # TẦN SUẤT >18
        21: 13, # TẦN SUẤT CÁP
        22: 18, # SỐ LÔ CO-CQ
        23: 18  # KẾT LUẬN
    }
    for col_idx, width in col_widths_rebar.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "E3"
    return r_end

def build_sheet_summary(ws, r_tot_concrete, r_end_rebar):
    """Xây dựng Sheet 3: Bảng tổng hợp tần suất vật liệu đầu vào và kiểm định hiện trường"""
    ws.title = "Tong_Hop_Tan_Suat_Vat_Lieu"
    ws.views.sheetView[0].showGridLines = True

    # Title Banner
    ws.merge_cells("A1:I1")
    ws["A1"] = "BẢNG TỔNG HỢP KẾ HOẠCH & TẦN SUẤT THÍ NGHIỆM VẬT LIỆU ĐẦU VÀO TOÀN DỰ ÁN"
    ws["A1"].font = FONT_TITLE
    ws["A1"].alignment = ALIGN_CENTER

    ws.merge_cells("A2:I2")
    ws["A2"] = "Dự án: Cầu Km19+529.080 | Tiêu chuẩn: Luật Xây dựng 135/2025/QH15, NĐ 207/2026/NĐ-CP, TCVN 1651:2018, TCVN 4453:1995"
    ws["A2"].font = FONT_SUBTITLE
    ws["A2"].alignment = ALIGN_CENTER

    # Table Header Row 4
    headers = [
        ("STT", 6),
        ("DANH MỤC VẬT LIỆU / CÔNG VIỆC THÍ NGHIỆM", 42),
        ("TIÊU CHUẨN ÁP DỤNG", 22),
        ("QUY ĐỊNH TẦN SUẤT LẤY MẪU", 26),
        ("ĐƠN VỊ TÍNH", 12),
        ("TỔNG KHỐI LƯỢNG DỰ ÁN", 22),
        ("SỐ LẦN THÍ NGHIỆM DỰ KIẾN", 24),
        ("CHỈ TIÊU CƠ LÝ KIỂM TRA CHÍNH", 36),
        ("GHI CHÚ / QUY ĐỊNH PHÁP LÝ", 28)
    ]

    ws.row_dimensions[4].height = 28
    for col_idx, (h_name, h_width) in enumerate(headers, start=1):
        cell = ws.cell(4, col_idx, h_name)
        cell.font = FONT_HDR
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER
        cell.fill = FILL_HDR
        ws.column_dimensions[get_column_letter(col_idx)].width = h_width

    summary_rows = [
        # Nhóm I: Vật liệu Bê tông & Cấp phối
        ("SECTION", "I. NHÓM VẬT LIỆU ĐẦU VÀO BÊ TÔNG XI MĂNG"),
        (1, "Xi măng Poóc lăng hỗn hợp PCB40", "TCVN 6260:2020", "50 Tấn / lần lấy mẫu", "Tấn",
         f"=Theo_Doi_Tan_Suat_Be_Tong!M{r_tot_concrete}", '=IF(F6="","",IF(F6=0,"1 lần",INT(F6/50)+1 & " lần"))',
         "Độ mịn, thời gian đông kết, giới hạn bền nén, độ ổn định thể tích", "TCVN 2682/6260, NĐ 207/2026"),
        (2, "Cát vàng đổ bê tông (Mô đun độ lớn Mk >= 2.0)", "TCVN 7570:2006", "200 m3 / lần lấy mẫu", "m3",
         f"=Theo_Doi_Tan_Suat_Be_Tong!N{r_tot_concrete}", '=IF(F7="","",IF(F7=0,"1 lần",INT(F7/200)+1 & " lần"))',
         "Thành phần hạt, hàm lượng bùn bụi sét, tạp chất hữu cơ, khối lượng thể tích", "TCVN 7570, Chỉ dẫn kỹ thuật"),
        (3, "Đá dăm bê tông (Đá 1x2 cho C30/C35/C45)", "TCVN 7570:2006", "350 m3 / lần lấy mẫu", "m3",
         f"=Theo_Doi_Tan_Suat_Be_Tong!O{r_tot_concrete}", '=IF(F8="","",IF(F8=0,"1 lần",INT(F8/350)+1 & " lần"))',
         "Thành phần hạt, hàm lượng thoi dẹt, độ nén dập trong ống hình trụ, độ hút nước", "TCVN 7570, TCVN 1771"),
        (4, "Nước trộn bê tông", "TCVN 4506:2012", "1 nguồn cấp / toàn dự án", "Nguồn",
         1, '1 lần (định kỳ 6 tháng)',
         "Độ pH, hàm lượng muối hòa tan, ion Cl-, SO4(2-), cặn không tan", "Kiểm tra trước khi thi công"),
        (5, "Phụ gia giảm nước / siêu dẻo bê tông", "TCVN 8826:2011", "5 Tấn / lô vật liệu", "Tấn",
         12.5, '3 lần',
         "Khả năng giảm nước, độ chảy xòe, độ pH, hàm lượng chất rắn", "Chứng chỉ xuất xưởng + thử nghiệm"),

        # Nhóm II: Cốt thép & Cáp DƯL
        ("SECTION", "II. NHÓM CỐT THÉP VÀ CÁP DỰ ỨNG LỰC"),
        (6, "Thép tròn trơn Ø <= 10mm (CB240-T)", "TCVN 1651-1:2018", "20 Tấn / lô đường kính", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!N{r_end_rebar}", '=IF(F12="","",IF(F12=0,"1 lần",INT(F12/20)+1 & " lần"))',
         "Giới hạn chảy, độ bền kéo, độ giãn dài tương đối, uốn nguội 180°", "1 tổ 3 kéo + 3 uốn / lô 20T"),
        (7, "Thép thanh vằn 10 < Ø <= 18mm (CB400-V)", "TCVN 1651-2:2018", "20 Tấn / lô đường kính", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!O{r_end_rebar}", '=IF(F13="","",IF(F13=0,"1 lần",INT(F13/20)+1 & " lần"))',
         "Giới hạn chảy, độ bền kéo, tỷ số Rmt/ReH, độ giãn dài, uốn nguội 180°", "1 tổ 3 kéo + 3 uốn / lô 20T"),
        (8, "Thép thanh vằn Ø > 18mm (CB400-V, CB500-V)", "TCVN 1651-2:2018", "20 Tấn / lô đường kính", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!P{r_end_rebar}", '=IF(F14="","",IF(F14=0,"1 lần",INT(F14/20)+1 & " lần"))',
         "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội 180°, sai số kích thước", "Thép chủ cọc, bệ, thân, xà mũ"),
        (9, "Cáp dự ứng lực 15.2mm 7 sợi xoắn", "ASTM A416 Gr270", "20 Tấn / lô kéo rút", "Tấn",
         f"=Theo_Doi_Tan_Suat_Cot_Thep!Q{r_end_rebar}", '=IF(F15="","",IF(F15=0,"1 lần",INT(F15/20)+1 & " lần"))',
         "Giới hạn bền kéo đứt, độ giãn dài, mô đun đàn hồi Ep, diện tích tiết diện", "Cáp dầm Super-T L=38.2m"),
        (10, "Mối nối cóc nối ren coupler / hàn đối đầu", "TCVN 8163:2009", "500 mối nối / lô", "Mối nối",
         1840, '4 lần (mỗi lô 3 mẫu kéo)',
         "Độ bền kéo mối nối, độ dãn dư, khả năng chịu biến dạng dẻo", "Mối nối thép chủ cọc khoan nhồi"),

        # Nhóm III: Thí nghiệm kiểm tra hiện trường & Cấu kiện hoàn thiện
        ("SECTION", "III. KIỂM ĐỊNH HIỆN TRƯỜNG & NGHIỆM THU CHẤT LƯỢNG"),
        (11, "Thí nghiệm nén mẫu bê tông hiện trường R7, R28", "TCVN 3118:2022", "1 tổ 3 mẫu / cấu kiện <=20m3", "Tổ mẫu",
         83, '83 tổ mẫu (249 viên)',
         "Cường độ chịu nén R7 ngày và R28 ngày (MPa)", "Kèm biên bản nghiệm thu KCS"),
        (12, "Thí nghiệm siêu âm cọc khoan nhồi D1.2m", "TCVN 9395:2012", "100% các cọc khoan nhồi", "Mặt cắt",
         156, '156 mặt cắt siêu âm (30 cọc)',
         "Độ đồng nhất, vận tốc truyền sóng, phát hiện khuyết tật bê tông thân cọc", "4 ống siêu âm / cọc (6 mặt cắt)"),
        (13, "Thử động biến dạng lớn PDA (Sức chịu tải cọc)", "ASTM D4945", "Tối thiểu 2 cọc đại diện", "Cọc",
         2, '2 cọc (M1-C1 và T1-C1)',
         "Sức chịu tải giới hạn cọc, ứng suất nén/kéo khi đóng, độ toàn vẹn thân cọc", "Đơn vị kiểm định độc lập"),
        (14, "Thí nghiệm độ sụt bê tông tươi tại hiện trường", "TCVN 3106:2022", "100% các xe bê tông trạm trộn", "Xe bồn",
         285, '285 xe (mỗi xe 1 lần)',
         "Độ sụt cọc nhồi (18±2cm), độ sụt kết cấu (14±2cm), nhiệt độ bê tông tươi", "Kiểm tra trước khi đổ"),
    ]

    r_cur = 5
    for item in summary_rows:
        if item[0] == "SECTION":
            ws.row_dimensions[r_cur].height = 22
            ws.merge_cells(start_row=r_cur, start_column=1, end_row=r_cur, end_column=9)
            sc = ws.cell(r_cur, 1, item[1])
            sc.font = Font(name=FONT_NAME, size=9.5, bold=True, color="1B365D")
            sc.alignment = Alignment(horizontal="left", vertical="center")
            sc.fill = FILL_SECTION
            for c in range(1, 10): ws.cell(r_cur, c).border = THIN_BORDER
            r_cur += 1
            continue

        stt_num, mat_name, std_name, freq_rule, unit_txt, total_val, test_cnt_formula, params_txt, note_txt = item
        ws.row_dimensions[r_cur].height = 20

        # A: STT
        ws.cell(r_cur, 1, stt_num).alignment = ALIGN_CENTER
        ws.cell(r_cur, 1).font = FONT_REG
        ws.cell(r_cur, 1).border = THIN_BORDER

        # B: Danh mục
        ws.cell(r_cur, 2, mat_name).alignment = ALIGN_LEFT
        ws.cell(r_cur, 2).font = FONT_BOLD
        ws.cell(r_cur, 2).border = THIN_BORDER

        # C: Tiêu chuẩn
        ws.cell(r_cur, 3, std_name).alignment = ALIGN_CENTER
        ws.cell(r_cur, 3).font = FONT_REG
        ws.cell(r_cur, 3).border = THIN_BORDER

        # D: Quy định tần suất
        ws.cell(r_cur, 4, freq_rule).alignment = ALIGN_LEFT
        ws.cell(r_cur, 4).font = FONT_REG
        ws.cell(r_cur, 4).border = THIN_BORDER

        # E: Đơn vị tính
        ws.cell(r_cur, 5, unit_txt).alignment = ALIGN_CENTER
        ws.cell(r_cur, 5).font = FONT_REG
        ws.cell(r_cur, 5).border = THIN_BORDER

        # F: Tổng khối lượng
        cell_tot = ws.cell(r_cur, 6, total_val)
        cell_tot.alignment = ALIGN_RIGHT
        cell_tot.font = FONT_BOLD
        cell_tot.border = THIN_BORDER
        if isinstance(total_val, (int, float)):
            cell_tot.number_format = "#,##0.00" if isinstance(total_val, float) else "#,##0"

        # G: Số lần thí nghiệm dự kiến
        cell_tc = ws.cell(r_cur, 7, test_cnt_formula)
        cell_tc.alignment = ALIGN_CENTER
        cell_tc.font = FONT_RED
        cell_tc.border = THIN_BORDER

        # H: Chỉ tiêu kiểm tra chính
        ws.cell(r_cur, 8, params_txt).alignment = ALIGN_LEFT
        ws.cell(r_cur, 8).font = FONT_REG
        ws.cell(r_cur, 8).border = THIN_BORDER

        # I: Ghi chú
        ws.cell(r_cur, 9, note_txt).alignment = ALIGN_LEFT
        ws.cell(r_cur, 9).font = FONT_REG
        ws.cell(r_cur, 9).border = THIN_BORDER

        r_cur += 1

    ws.freeze_panes = "C5"

def build_tan_suat_master_workbook(output_path: str):
    """Tạo toàn bộ Workbook gồm cả 3 Sheet chuyên nghiệp"""
    base_marker = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker"
    bsl_path = os.path.join(base_marker, "bang_so_lieu.json")
    thep_cat_path = os.path.join(base_marker, "thep_cho_to_hop_cat.json")
    tmpl_path = r"d:\Code\23HG-multiagent-system-main\23HG-multiagent-system-main\templates\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx"

    wb = openpyxl.Workbook()
    # Sheet 1: Bê tông
    ws1 = wb.active
    r_tot = build_sheet_concrete(ws1)

    # Sheet 2: Cốt thép bóc tách chi tiết từng thanh
    ws2 = wb.create_sheet("Theo_Doi_Tan_Suat_Cot_Thep")
    r_end = build_sheet_rebar(ws2, bsl_path, thep_cat_path, tmpl_path)

    # Sheet 3: Tổng hợp vật liệu
    ws3 = wb.create_sheet("Tong_Hop_Tan_Suat_Vat_Lieu")
    build_sheet_summary(ws3, r_tot, r_end)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    wb.close()
    print(f"[OK] Đã xuất bản thành công Bộ hồ sơ tần suất thí nghiệm chuẩn: {output_path}")

if __name__ == "__main__":
    out_master = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Km19.xlsx"
    build_tan_suat_master_workbook(out_master)

    # Xuất bản song song cho các gói hồ sơ thực chiến
    extra_paths = [
        r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Cot_Thep_Cau_Km19.xlsx",
        r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Tong_Hop_Cau_Km19.xlsx",
        r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Cau_Km19.xlsx",
        r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_C_HIEN_TRUONG_QLCL_KCS\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Km19.xlsx",
        r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker\BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Km19.xlsx",
    ]
    for p in extra_paths:
        try:
            import shutil
            shutil.copy2(out_master, p)
            print(f"[OK] Đã đồng bộ sang: {p}")
        except Exception as e:
            print(f"[NOTE] Không thể ghi đè sang {os.path.basename(p)} (có thể file đang được mở trong Excel): {e}")

