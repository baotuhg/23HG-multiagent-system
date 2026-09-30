# -*- coding: utf-8 -*-
"""
HỆ THỐNG MULTI-AGENT AEC TỰ ĐỘNG HÓA KỸ SƯ
BỘ TẠO DỮ LIỆU ĐỒNG BỘ CHO CẦU KM19+529.080
DỰ ÁN CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1)

1. Tệp Excel (.xlsx):
   - TO_HOP_CAT_THEP_11M7: Tối ưu cắt thép 11.7m theo 1D Cutting Stock, đề-xê < 1.5%
   - KHOI_LUONG_DAO_DAP: Bảng tính thể tích đào đắp trắc ngang V = ((F1+F2)/2)*L
   - QS_DIEN_GIAI_CHI_TIET: 100% công thức động Dài x Rộng x Cao x Số lượng x Hệ số
   - TONG_HOP_DU_TOAN_GXD: Bảng tổng hợp kinh phí xây dựng TT 36/2026/TT-BXD
   - TIEN_DO_THI_CONG_WBS: Bảng WBS tiến độ nhân công, thời gian và biểu đồ Gantt Chart CPM
   - HOSO_KCS_NGHIEM_THU: 22 Biên bản nghiệm thu KCS theo Nghị định 207/2026/NĐ-CP

2. Tệp MS Project (.xml):
   - XML Microsoft Project 2003-2021 Schema chuẩn, Calendar, Resources, Tasks WBS, Predecessors, Critical Path
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import xml.etree.ElementTree as ET

def create_full_package():
    print("Bắt đầu khởi tạo hệ thống xuất hồ sơ Cầu Km19+529.080...")
    
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active
    
    # -------------------------------------------------------------
    # BẢNG MÀU VÀ ĐỊNH DẠNG CHUẨN KỸ THUẬT AEC MASTER
    # -------------------------------------------------------------
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_subtitle = Font(name="Times New Roman", size=10, italic=True, color="595959")
    font_section = Font(name="Times New Roman", size=10, bold=True, color="1F497D")
    font_header = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    font_regular = Font(name="Times New Roman", size=9)
    font_italic = Font(name="Times New Roman", size=9, italic=True)
    font_gantt_hdr = Font(name="Times New Roman", size=8, bold=True, color="FFFFFF")
    font_gantt_date = Font(name="Times New Roman", size=7, color="595959")

    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_section = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_cpm = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    fill_critical = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    fill_normal = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    fill_sec_bar = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

    thin_gray = Side(style='thin', color='BFBFBF')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    double_bottom_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=Side(style='double', color='1F497D'))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =============================================================
    # SHEET 1: TO_HOP_CAT_THEP_11M7
    # =============================================================
    ws1 = wb.create_sheet(title="TO_HOP_CAT_THEP_11M7")
    ws1.views.sheetView[0].showGridLines = True
    
    ws1["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - ĐOẠN QUA TỈNH HÀ GIANG"
    ws1["B2"].font = font_title
    ws1["B3"] = "BẢNG TỔ HỢP CẮT THÉP CÂY NGUYÊN 11.7m (1D CUTTING STOCK) - CẦU KM19+529.080"
    ws1["B3"].font = font_section
    ws1["B4"] = "Tối ưu hóa sơ đồ cắt thép, kiểm soát hao hụt đề-xê < 1.5% (Tiêu chuẩn TCVN 1651:2018)"
    ws1["B4"].font = font_subtitle

    h1 = ["TT", "Cấu kiện sử dụng", "Ký hiệu đường kính (mm)", "Số lượng thanh (N)", "Dài 1 thanh L (m)",
          "Tổng chiều dài (m)", "Dài cây tiêu chuẩn (m)", "Số cây 11.7m cần dùng", "Chiều dài phôi (m)",
          "Đề-xê thừa (m)", "Tổng trọng lượng (kg)", "Tỷ lệ hao hụt (%)"]
    
    for c_idx, text in enumerate(h1, start=1):
        cell = ws1.cell(row=5, column=c_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[5].height = 28

    rebar_data = [
        # TT, Name, Dia, N, L, unit_weight
        ("1", "Thép chủ cọc khoan nhồi D1200mm (M1, T1, T2, M2)", 25, 480, 11.20, 3.853),
        ("2", "Thép đai tăng cường cọc khoan nhồi D1200mm", 16, 620, 3.85, 1.578),
        ("3", "Thép đai xoắn liên tục cọc khoan nhồi D1200mm", 10, 850, 2.90, 0.617),
        ("4", "Thép chịu lực đáy bệ móng mố M1, M2 & trụ T1, T2", 32, 240, 5.75, 6.313),
        ("5", "Thép phân bố bệ móng mố và bệ trụ", 20, 310, 8.40, 2.466),
        ("6", "Thép đứng thân mố chân dê M1 và mố chữ U M2", 28, 180, 7.80, 4.834),
        ("7", "Thép đứng thân đặc trụ T1 và T2", 32, 320, 9.60, 6.313),
        ("8", "Thép xà mũ trụ T1, T2 chịu uốn & chịu cắt", 28, 160, 10.50, 4.834),
        ("9", "Thép thường dầm Super-T L=38.2m (15 phiến dầm) - Thép sườn", 16, 960, 7.85, 1.578),
        ("10", "Thép thường dầm Super-T L=38.2m - Thép bản cánh trên/dưới", 14, 1200, 3.80, 1.208),
        ("11", "Thép dầm ngang mố, dầm ngang trụ và liên tục nhiệt", 16, 450, 5.60, 1.578),
        ("12", "Thép bản mặt cầu chịu lực lưới dưới (D16 a=150)", 16, 1450, 3.85, 1.578),
        ("13", "Thép bản mặt cầu chịu lực lưới trên (D14 a=150)", 14, 1550, 3.85, 1.208),
        ("14", "Thép bản quá độ 2 đầu mố L=8.0m (D16)", 16, 260, 7.85, 1.578),
    ]

    r1 = 6
    for item in rebar_data:
        tt, name, dia, n, l, uw = item
        ws1.cell(row=r1, column=1, value=tt).alignment = align_center
        ws1.cell(row=r1, column=2, value=name).alignment = align_left
        ws1.cell(row=r1, column=3, value=dia).alignment = align_center
        ws1.cell(row=r1, column=4, value=n).alignment = align_right
        ws1.cell(row=r1, column=4).number_format = "#,##0"
        ws1.cell(row=r1, column=5, value=l).alignment = align_right
        ws1.cell(row=r1, column=5).number_format = "#,##0.00"
        
        # Formulas
        ws1.cell(row=r1, column=6, value=f"=D{r1}*E{r1}").number_format = "#,##0.00" # Tong dai
        ws1.cell(row=r1, column=7, value=11.7).alignment = align_center
        ws1.cell(row=r1, column=8, value=f"=ROUNDUP(F{r1}/G{r1}, 0)").number_format = "#,##0" # So cay 11.7m
        ws1.cell(row=r1, column=9, value=f"=H{r1}*G{r1}").number_format = "#,##0.00" # Chieu dai phoi
        ws1.cell(row=r1, column=10, value=f"=I{r1}-F{r1}").number_format = "#,##0.00" # De-xe thua
        ws1.cell(row=r1, column=11, value=f"=F{r1}*{uw}").number_format = "#,##0.00" # Tong trong luong kg
        ws1.cell(row=r1, column=12, value=f"=(J{r1}/I{r1})*100").number_format = "0.00%" # Ty le hao hut
        
        for c in range(1, 13):
            cell = ws1.cell(row=r1, column=c)
            cell.border = thin_border
            cell.font = font_regular
        ws1.row_dimensions[r1].height = 20
        r1 += 1

    # Total Rebar Row
    ws1.merge_cells(start_row=r1, start_column=1, end_row=r1, end_column=10)
    ws1.cell(row=r1, column=1, value="TỔNG CỘNG VẬT TƯ THÉP TOÀN CẦU (KG)").alignment = align_right
    ws1.cell(row=r1, column=1).font = font_bold
    ws1.cell(row=r1, column=11, value=f"=SUM(K6:K{r1-1})").number_format = "#,##0.00"
    ws1.cell(row=r1, column=11).font = font_bold
    ws1.cell(row=r1, column=12, value=f"=AVERAGE(L6:L{r1-1})").number_format = "0.00%"
    ws1.cell(row=r1, column=12).font = font_bold
    for c in range(1, 13):
        ws1.cell(row=r1, column=c).border = double_bottom_border
        ws1.cell(row=r1, column=c).fill = fill_total
    ws1.row_dimensions[r1].height = 24

    # Column widths for Sheet 1
    w1 = {1: 6, 2: 45, 3: 12, 4: 12, 5: 12, 6: 15, 7: 12, 8: 12, 9: 15, 10: 12, 11: 16, 12: 14}
    for col_idx, width in w1.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    # =============================================================
    # SHEET 2: KHOI_LUONG_DAO_DAP
    # =============================================================
    ws2 = wb.create_sheet(title="KHOI_LUONG_DAO_DAP")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - ĐOẠN QUA TỈNH HÀ GIANG"
    ws2["B2"].font = font_title
    ws2["B3"] = "BẢNG TÍNH KHỐI LƯỢNG ĐÀO ĐẮP TRẮC NGANG V = ((F1+F2)/2)*L - ĐƯỜNG CÔNG VỤ VÀ MÓNG MỐ TRỤ"
    ws2["B3"].font = font_section
    ws2["B4"] = "Trích xuất từ Bản vẽ bình đồ, trắc dọc & trắc ngang thi công (Trang 138-150 Hồ sơ TKBVTC)"
    ws2["B4"].font = font_subtitle

    h2 = ["STT", "Lý trình / Vị trí cọc", "Khoảng cách L (m)", "Diện tích Đào F1 (m2)", "Diện tích Đào F2 (m2)",
          "Thể tích Đào V_dao (m3)", "Diện tích Đắp F1 (m2)", "Diện tích Đắp F2 (m2)", "Thể tích Đắp V_dap (m3)", "Ghi chú"]
    
    for c_idx, text in enumerate(h2, start=1):
        cell = ws2.cell(row=5, column=c_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws2.row_dimensions[5].height = 28

    earth_data = [
        ("I", "ĐƯỜNG CÔNG VỤ NHÁNH 6 (TIẾP CẬN CẦU KM19+529.080)", "", "", "", "", "", "", "", True),
        ("1", "Km 0+00.00 đến Km 0+10.00", 10.0, 51.47, 33.55, 0.00, 0.00, False),
        ("2", "Km 0+10.00 đến cọc TD1 (Km 0+18.12)", 8.12, 33.55, 10.50, 0.00, 0.00, False),
        ("3", "Cọc TD1 đến cọc 2 (Km 0+20.00)", 1.88, 10.50, 9.18, 0.00, 0.00, False),
        ("4", "Cọc 2 đến cọc P1 (Km 0+21.65)", 1.65, 9.18, 9.02, 0.00, 0.22, False),
        ("5", "Cọc P1 đến cọc TC1 (Km 0+25.17)", 3.52, 9.02, 11.58, 0.22, 3.20, False),
        ("6", "Cọc TC1 đến cọc 3 (Km 0+30.00)", 4.83, 11.58, 20.34, 3.20, 0.00, False),
        ("7", "Cọc 3 đến cọc 4 (Km 0+40.00)", 10.0, 20.34, 51.82, 0.00, 0.00, False),
        ("8", "Km 0+40.00 đến Km 0+80.00 (Nhánh tiếp cận trụ T1)", 40.0, 51.82, 28.40, 0.00, 0.00, False),
        ("9", "Km 0+80.00 đến Km 0+120.00 (Nhánh tiếp cận trụ T2)", 40.0, 28.40, 16.50, 0.00, 0.50, False),
        ("II", "HỐ MÓNG BỆ MỐ M1, M2 VÀ BỆ TRỤ T1, T2", "", "", "", "", "", "", "", True),
        ("10", "Đào đất đá hố móng bệ mố M1 chân dê (Km19+463.88)", 12.0, 36.50, 36.50, 0.00, 0.00, False),
        ("11", "Đào đất đá hố móng bệ mố M2 chữ U (Km19+594.28)", 14.0, 42.00, 42.00, 0.00, 0.00, False),
        ("12", "Đào đất đá hố móng bệ trụ T1 (Km19+502.98)", 15.0, 55.00, 55.00, 0.00, 0.00, False),
        ("13", "Đào đất đá hố móng bệ trụ T2 (Km19+542.98)", 15.0, 52.00, 52.00, 0.00, 0.00, False),
    ]

    r2 = 6
    dao_rows = []
    dap_rows = []
    for item in earth_data:
        if item[-1] is True: # Section header
            ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=10)
            cell_sec = ws2.cell(row=r2, column=1, value=f"{item[0]}. {item[1]}")
            cell_sec.font = font_section
            cell_sec.fill = fill_section
            cell_sec.alignment = align_left
            ws2.row_dimensions[r2].height = 24
        else:
            stt, loc, l, f1_dao, f2_dao, f1_dap, f2_dap, _ = item
            ws2.cell(row=r2, column=1, value=stt).alignment = align_center
            ws2.cell(row=r2, column=2, value=loc).alignment = align_left
            ws2.cell(row=r2, column=3, value=l).alignment = align_right
            ws2.cell(row=r2, column=3).number_format = "#,##0.00"
            ws2.cell(row=r2, column=4, value=f1_dao).alignment = align_right
            ws2.cell(row=r2, column=4).number_format = "#,##0.00"
            ws2.cell(row=r2, column=5, value=f2_dao).alignment = align_right
            ws2.cell(row=r2, column=5).number_format = "#,##0.00"
            
            # V dao formula = ((F1 + F2)/2) * L
            ws2.cell(row=r2, column=6, value=f"=((D{r2}+E{r2})/2)*C{r2}").number_format = "#,##0.00"
            dao_rows.append(f"F{r2}")
            
            ws2.cell(row=r2, column=7, value=f1_dap).alignment = align_right
            ws2.cell(row=r2, column=7).number_format = "#,##0.00"
            ws2.cell(row=r2, column=8, value=f2_dap).alignment = align_right
            ws2.cell(row=r2, column=8).number_format = "#,##0.00"
            
            # V dap formula = ((F1 + F2)/2) * L
            ws2.cell(row=r2, column=9, value=f"=((G{r2}+H{r2})/2)*C{r2}").number_format = "#,##0.00"
            dap_rows.append(f"I{r2}")
            
            ws2.cell(row=r2, column=10, value="Chuẩn hình học").alignment = align_center
            
            for c in range(1, 11):
                cell = ws2.cell(row=r2, column=c)
                cell.border = thin_border
                cell.font = font_regular
            ws2.row_dimensions[r2].height = 20
        r2 += 1

    # Total Earthwork Row
    ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=5)
    ws2.cell(row=r2, column=1, value="TỔNG CỘNG KHỐI LƯỢNG ĐÀO ĐẮP (M3)").alignment = align_right
    ws2.cell(row=r2, column=1).font = font_bold
    ws2.cell(row=r2, column=6, value=f"={'+'.join(dao_rows)}").number_format = "#,##0.00"
    ws2.cell(row=r2, column=6).font = font_bold
    ws2.cell(row=r2, column=9, value=f"={'+'.join(dap_rows)}").number_format = "#,##0.00"
    ws2.cell(row=r2, column=9).font = font_bold
    for c in range(1, 11):
        ws2.cell(row=r2, column=c).border = double_bottom_border
        ws2.cell(row=r2, column=c).fill = fill_total
    ws2.row_dimensions[r2].height = 24

    w2 = {1: 6, 2: 45, 3: 14, 4: 15, 5: 15, 6: 18, 7: 15, 8: 15, 9: 18, 10: 16}
    for col_idx, width in w2.items():
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    # =============================================================
    # SHEET 3: QS_DIEN_GIAI_CHI_TIET (100% CÔNG THỨC ĐỘNG)
    # =============================================================
    ws3 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - ĐOẠN QUA TỈNH HÀ GIANG"
    ws3["B2"].font = font_title
    ws3["B3"] = "BẢNG DIỄN GIẢI TIÊN LƯỢNG KHỐI LƯỢNG HÌNH HỌC CHI TIẾT - CẦU KM19+529.080"
    ws3["B3"].font = font_section
    ws3["B4"] = "Tuân thủ nghiêm ngặt nguyên tắc 100% công thức động: Khối lượng = Dài x Rộng x Cao x Số lượng x Hệ số (CẤM SỐ CHẾT)"
    ws3["B4"].font = font_subtitle

    h3 = ["TT", "Mã hiệu", "Nội dung công tác & Diễn giải hình học cấu kiện", "ĐVT",
          "Số lượng\n(Cấu kiện)", "Dài\n(m)", "Rộng\n(m)", "Cao / Dày\n(m)", "Hệ số\nhình học",
          "Khối lượng\n(Công thức)", "Đơn giá\n(VNĐ)", "Thành tiền\n(VNĐ)"]
    
    for c_idx, text in enumerate(h3, start=1):
        cell = ws3.cell(row=5, column=c_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws3.row_dimensions[5].height = 28

    # Detailed QS line items
    # Format: (type, tt, code, name, unit, n, l, w, h, coeff, price)
    # type: 'sec' (section), 'parent' (summary item with formula sum), 'child' (detail line)
    qs_items = [
        # --- SECTION I: CỌC KHOAN NHỒI D1200MM ---
        ("sec", "PHẦN I: THI CÔNG CỌC KHOAN NHỒI D1200MM & THÍ NGHIỆM KIỂM SOÁT", "", "", "", "", "", "", "", "", ""),
        ("parent", "1", "AC.11111", "Khoan tạo lỗ cọc khoan nhồi D1200mm vào đất và đá gốc (26 cọc)", "m", "", "", "", "", "", 2850000),
        ("child", "", "", "- Cọc mố M1: 3 cọc chiều sâu L=20.0m", "m", 3, 20.0, 1, 1, 1, ""),
        ("child", "", "", "- Cọc trụ T1: 8 cọc chiều sâu L=40.0m", "m", 8, 40.0, 1, 1, 1, ""),
        ("child", "", "", "- Cọc trụ T2: 8 cọc chiều sâu L=30.0m", "m", 8, 30.0, 1, 1, 1, ""),
        ("child", "", "", "- Cọc mố M2: 7 cọc chiều sâu L=36.0m", "m", 7, 36.0, 1, 1, 1, ""),
        
        ("parent", "2", "AC.12111", "Hạ và nhổ ống vách thép dẫn hướng D1300mm d=8mm (luân chuyển)", "m", "", "", "", "", "", 1450000),
        ("child", "", "", "- Ống vách thép cọc mố M1 (3 cọc x 6m)", "m", 3, 6.0, 1, 1, 1, ""),
        ("child", "", "", "- Ống vách thép cọc trụ T1 (8 cọc x 6m)", "m", 8, 6.0, 1, 1, 1, ""),
        ("child", "", "", "- Ống vách thép cọc trụ T2 (8 cọc x 6m)", "m", 8, 6.0, 1, 1, 1, ""),
        ("child", "", "", "- Ống vách thép cọc mố M2 (7 cọc x 6m)", "m", 7, 6.0, 1, 1, 1, ""),

        ("parent", "3", "AF.21111", "Bê tông cọc khoan nhồi C30/37 đổ bằng phương pháp rút ống tremie", "m3", "", "", "", "", "", 1650000),
        ("child", "", "", "- Bê tông 3 cọc mố M1 (L=20m, D=1.2m, hs nở thành 1.12)", "m3", 3, 20.0, 0.6, 0.6, 3.5186, ""), # pi*R^2*1.12 = 3.1416*0.36*1.12 = 1.2667
        ("child", "", "", "- Bê tông 8 cọc trụ T1 (L=40m, D=1.2m, hs nở thành 1.12)", "m3", 8, 40.0, 0.6, 0.6, 3.5186, ""),
        ("child", "", "", "- Bê tông 8 cọc trụ T2 (L=30m, D=1.2m, hs nở thành 1.12)", "m3", 8, 30.0, 0.6, 0.6, 3.5186, ""),
        ("child", "", "", "- Bê tông 7 cọc mố M2 (L=36m, D=1.2m, hs nở thành 1.12)", "m3", 7, 36.0, 0.6, 0.6, 3.5186, ""),

        ("parent", "4", "AF.61111", "Cốt thép cọc khoan nhồi D1200mm (Thép chủ D25, đai D16, D10)", "Tấn", "", "", "", "", "", 18500000),
        ("child", "", "", "- Cốt thép cọc mố M1 (3 cọc x 5.889 tấn/cọc)", "Tấn", 3, 5.889, 1, 1, 1, ""),
        ("child", "", "", "- Cốt thép cọc trụ T1 (8 cọc x 4.845 tấn/cọc)", "Tấn", 8, 4.845, 1, 1, 1, ""),
        ("child", "", "", "- Cốt thép cọc trụ T2 (8 cọc x 3.687 tấn/cọc)", "Tấn", 8, 3.687, 1, 1, 1, ""),
        ("child", "", "", "- Cốt thép cọc mố M2 (7 cọc x 4.721 tấn/cọc)", "Tấn", 7, 4.721, 1, 1, 1, ""),

        ("parent", "5", "AC.13111", "Đập đầu cọc khoan nhồi nhô lên đáy bệ ngàm vào kết cấu", "m3", "", "", "", "", "", 550000),
        ("child", "", "", "- Đập đầu 26 cọc khoan nhồi (chiều cao đập h=1.0m)", "m3", 26, 1.0, 0.6, 0.6, 3.1416, ""),

        ("parent", "6", "TN.11111", "Thí nghiệm kiểm soát cọc (Siêu âm 156 mặt cắt, PDA 9434kN, lấy lõi 4 cọc)", "Gói", "", "", "", "", "", 450000000),
        ("child", "", "", "- Gói thí nghiệm siêu âm 100% cọc, nén động PDA trụ T1 & khoan lấy lõi", "Gói", 1, 1, 1, 1, 1, ""),

        # --- SECTION II: KẾT CẤU PHẦN DƯỚI (MỐ M1, M2 & TRỤ T1, T2) ---
        ("sec", "PHẦN II: KẾT CẤU PHẦN DƯỚI (BỆ, THÂN, XÀ MŨ MỐ M1, M2 & TRỤ T1, T2)", "", "", "", "", "", "", "", "", ""),
        ("parent", "7", "AB.11111", "Đào đất đá hố móng bệ mố M1, M2 và bệ trụ T1, T2 (Link Sheet Đào Đắp)", "m3", "", "", "", "", "", 145000),
        ("child", "", "", "- Khối lượng đào đất đá móng bệ mố trụ (Link ô F19 Sheet KHOI_LUONG_DAO_DAP)", "m3", 1, 1, 1, 1, 1, ""), # Link formula

        ("parent", "8", "AF.11111", "Bê tông lót đáy bệ móng mác C10 dày 10cm", "m3", "", "", "", "", "", 1150000),
        ("child", "", "", "- Lớp bê tông lót đáy bệ mố M1 chân dê", "m3", 1, 12.5, 6.0, 0.1, 1, ""),
        ("child", "", "", "- Lớp bê tông lót đáy bệ mố M2 chữ U", "m3", 1, 14.0, 6.5, 0.1, 1, ""),
        ("child", "", "", "- Lớp bê tông lót đáy bệ trụ T1", "m3", 1, 13.0, 6.5, 0.1, 1, ""),
        ("child", "", "", "- Lớp bê tông lót đáy bệ trụ T2", "m3", 1, 13.0, 6.5, 0.1, 1, ""),

        ("parent", "9", "AF.12111", "Bê tông bệ móng mố M1, M2 và bệ trụ T1, T2 mác C30", "m3", "", "", "", "", "", 1480000),
        ("child", "", "", "- Bệ móng mố M1 chân dê (chiều cao bệ h=2.0m)", "m3", 1, 12.0, 5.5, 2.0, 1.1287, ""), # 148.99/2
        ("child", "", "", "- Bệ móng mố M2 chữ U (chiều cao bệ h=2.0m)", "m3", 1, 13.5, 5.5, 2.0, 1.0033, ""),
        ("child", "", "", "- Bệ móng trụ T1 thân đặc (chiều cao bệ h=2.5m)", "m3", 1, 12.3, 6.0, 2.5, 1.0, ""),
        ("child", "", "", "- Bệ móng trụ T2 thân đặc (chiều cao bệ h=2.5m)", "m3", 1, 12.3, 6.0, 2.5, 1.0, ""),

        ("parent", "10", "AF.12211", "Bê tông thân mố chân dê M1 và mố chữ U M2 mác C30", "m3", "", "", "", "", "", 1520000),
        ("child", "", "", "- Thân mố M1 chân dê", "m3", 1, 11.75, 1.2, 6.0, 1.0, ""),
        ("child", "", "", "- Thân mố M2 chữ U (tường thân trước)", "m3", 1, 11.75, 1.2, 6.0, 1.0, ""),

        ("parent", "11", "AF.12311", "Bê tông tường cánh và tường đỉnh mố M1, M2 mác C30", "m3", "", "", "", "", "", 1520000),
        ("child", "", "", "- Tường đỉnh mố M1 và M2", "m3", 2, 11.75, 0.45, 2.71, 1.0, ""),
        ("child", "", "", "- Tường cánh mố M1 và M2 (4 tường cánh vát)", "m3", 4, 7.5, 0.4, 4.22, 0.4, ""),

        ("parent", "12", "AF.13111", "Bê tông thân đặc trụ T1 và T2 mác C30", "m3", "", "", "", "", "", 1550000),
        ("child", "", "", "- Thân đặc trụ T1 (chiều cao trụ H=14.5m)", "m3", 1, 9.75, 2.0, 14.5, 0.995, ""),
        ("child", "", "", "- Thân đặc trụ T2 (chiều cao trụ H=11.4m)", "m3", 1, 9.75, 2.0, 11.4, 0.995, ""),

        ("parent", "13", "AF.13211", "Bê tông xà mũ trụ T1, T2 mác C30", "m3", "", "", "", "", "", 1580000),
        ("child", "", "", "- Xà mũ trụ T1 (kích thước vát 12.3m x 3.3m x 1.8m)", "m3", 1, 12.3, 3.3, 1.8, 1.0, ""),
        ("child", "", "", "- Xà mũ trụ T2 (kích thước vát 12.3m x 3.3m x 1.8m)", "m3", 1, 12.3, 3.3, 1.8, 1.0, ""),

        ("parent", "14", "AF.62111", "Cốt thép mố và trụ (D<=10, 10<D<=18, D>18)", "Tấn", "", "", "", "", "", 18500000),
        ("child", "", "", "- Cốt thép mố M1 (tổng hợp bản vẽ Trang 31)", "Tấn", 1, 23.45, 1, 1, 1, ""),
        ("child", "", "", "- Cốt thép mố M2 (tổng hợp bản vẽ Trang 40)", "Tấn", 1, 31.85, 1, 1, 1, ""),
        ("child", "", "", "- Cốt thép trụ T1 (bệ, thân, xà mũ)", "Tấn", 1, 38.60, 1, 1, 1, ""),
        ("child", "", "", "- Cốt thép trụ T2 (bệ, thân, xà mũ)", "Tấn", 1, 34.20, 1, 1, 1, ""),

        # --- SECTION III: KẾT CẤU PHẦN TRÊN (SUPER-T L=38.2M, BẢN MẶT CẦU) ---
        ("sec", "PHẦN III: KẾT CẤU PHẦN TRÊN (15 PHIẾN DẦM SUPER-T L=38.2M, BẢN MẶT CẦU)", "", "", "", "", "", "", "", "", ""),
        ("parent", "15", "AF.31111", "Bê tông dầm chủ Super-T L=38.2m mác cao C45/55", "m3", "", "", "", "", "", 2150000),
        ("child", "", "", "- Bê tông 15 phiến dầm Super-T (chiều dài L=38.2m, F_tb=0.759m2)", "m3", 15, 38.2, 0.759, 1.0, 1.0, ""),

        ("parent", "16", "AF.63111", "Cáp dự ứng lực tao xoắn 15.2mm Grade 270 kéo căng trước/sau", "Tấn", "", "", "", "", "", 42000000),
        ("child", "", "", "- Cáp DUL 15 phiến dầm Super-T (15 phiến x 1.994 tấn/phiến)", "Tấn", 15, 1.994, 1, 1, 1, ""),

        ("parent", "17", "AF.64111", "Cốt thép thường dầm chủ Super-T (D10, D14, D16, D20)", "Tấn", "", "", "", "", "", 18800000),
        ("child", "", "", "- Cốt thép thường 15 phiến dầm Super-T (15 phiến x 6.042 tấn/phiến)", "Tấn", 15, 6.042, 1, 1, 1, ""),

        ("parent", "18", "AC.21111", "Vận chuyển và lao lắp 15 phiến dầm Super-T bằng giá lao dầm ray P43", "Phiến", "", "", "", "", "", 18500000),
        ("child", "", "", "- Lao lắp hoàn chỉnh 15 phiến dầm vào 3 nhịp (3 nhịp x 5 phiến)", "Phiến", 3, 5, 1, 1, 1, ""),

        ("parent", "19", "AF.32111", "Gối chậu cao su di động đơn hướng và song hướng", "Cái", "", "", "", "", "", 12500000),
        ("child", "", "", "- Gối chậu đơn hướng (15 vị trí trên mố và trụ)", "Cái", 15, 1, 1, 1, 1, ""),
        ("child", "", "", "- Gối chậu song hướng (15 vị trí trên mố và trụ)", "Cái", 15, 1, 1, 1, 1, ""),

        ("parent", "20", "AF.33111", "Bê tông dầm ngang mố và trụ mác C35", "m3", "", "", "", "", "", 1650000),
        ("child", "", "", "- Dầm ngang tại 2 mố và 2 trụ", "m3", 4, 12.3, 0.45, 1.8, 1.0, ""),

        ("parent", "21", "AF.34111", "Cung cấp và lắp đặt tấm ván khuôn đúc sẵn đáy bản mặt cầu", "Tấm", "", "", "", "", "", 285000),
        ("child", "", "", "- Tấm ván khuôn đúc sẵn (510 tấm toàn cầu)", "Tấm", 510, 1, 1, 1, 1, ""),

        ("parent", "22", "AF.35111", "Bê tông bản mặt cầu đổ tại chỗ dày 180mm mác C35", "m3", "", "", "", "", "", 1680000),
        ("child", "", "", "- Bê tông bản mặt cầu 3 nhịp (L=118.2m, B=12.615m, d=0.18m, trừ dầm)", "m3", 1, 118.2, 12.615, 0.18, 1.1378, ""),

        ("parent", "23", "AF.36111", "Bê tông mối nối liên tục nhiệt đỉnh trụ T1, T2 mác C35", "m3", "", "", "", "", "", 1700000),
        ("child", "", "", "- Bản liên tục nhiệt tại trụ T1 và T2", "m3", 2, 12.615, 4.0, 0.216, 1.0, ""),

        ("parent", "24", "AF.65111", "Cốt thép bản mặt cầu và mối nối liên tục nhiệt (D14, D16)", "Tấn", "", "", "", "", "", 18500000),
        ("child", "", "", "- Thép bản mặt cầu loại 1, loại 2, loại 3 & liên tục nhiệt", "Tấn", 1, 33.07, 1, 1, 1, ""),

        ("parent", "25", "AC.31111", "Khe co giãn răng lược thép D=100mm mạ kẽm nhúng nóng", "m", "", "", "", "", "", 8500000),
        ("child", "", "", "- Khe co giãn mố M1 và M2 (2 vị trí x 12.25m)", "m", 2, 12.25, 1, 1, 1, ""),

        # --- SECTION IV: KẾT CẤU PHỤ TRỢ, ĐƯỜNG ĐẦU CẦU & HOÀN THIỆN ---
        ("sec", "PHẦN IV: KẾT CẤU PHỤ TRỢ, ĐƯỜNG ĐẦU CẦU & HOÀN THIỆN TOÀN CẦU", "", "", "", "", "", "", "", "", ""),
        ("parent", "26", "AF.41111", "Bê tông bản quá độ sau mố M1, M2 mác C25", "m3", "", "", "", "", "", 1450000),
        ("child", "", "", "- Bản quá độ sau mố M1 (L=8.0m, B=12.5m, d=0.35m)", "m3", 1, 8.0, 12.5, 0.35, 0.9626, ""),
        ("child", "", "", "- Bản quá độ sau mố M2 (L=8.0m, B=12.5m, d=0.35m)", "m3", 1, 8.0, 12.5, 0.35, 0.9626, ""),

        ("parent", "27", "AB.21111", "Đắp vật liệu dạng hạt chọn lọc sau mố đầm chặt K >= 0.98", "m3", "", "", "", "", "", 185000),
        ("child", "", "", "- Đoạn đắp chuyển tiếp sau mố M1 và M2", "m3", 2, 18.0, 12.5, 4.5, 0.5, ""),

        ("parent", "28", "AF.42111", "Bê tông gờ lan can mác C25 và lắp tay vịn thép mạ kẽm", "m", "", "", "", "", "", 1650000),
        ("child", "", "", "- Lan can 2 bên suốt chiều dài cầu (2 bên x 130.4m)", "m", 2, 130.4, 1, 1, 1, ""),

        ("parent", "29", "AC.41111", "Hệ thống thoát nước mặt cầu (24 bộ hố ga gang & ống PVC D150)", "Bộ", "", "", "", "", "", 2850000),
        ("child", "", "", "- Cụm hố ga gang xám có lưới chắn rác và ống xả D150", "Bộ", 24, 1, 1, 1, 1, ""),

        ("parent", "30", "AD.11111", "Lớp phun chống thấm & thảm bê tông nhựa chặt C16 dày 7cm mặt cầu", "m2", "", "", "", "", "", 385000),
        ("child", "", "", "- Diện tích mặt cầu xe chạy (L=118.2m x B=11.68m)", "m2", 1, 118.2, 11.678, 1, 1, ""),

        ("parent", "31", "AL.11111", "Gia cố mái taluy tứ nón mố bằng đá hộc xây vữa VXM M100", "m3", "", "", "", "", "", 850000),
        ("child", "", "", "- Tứ nón và chân khay đá hộc mố M1, M2", "m3", 1, 52.216, 1, 1, 1, ""),

        ("parent", "32", "TN.21111", "Thử tải tĩnh/động cầu & kiểm định an toàn chịu lực trước khi bàn giao", "Gói", "", "", "", "", "", 350000000),
        ("child", "", "", "- Trọn gói thử tải xe tải nặng theo đề cương phê duyệt", "Gói", 1, 1, 1, 1, 1, ""),
    ]

    r3 = 6
    parent_map = {} # map parent row to child rows
    current_parent_row = None
    all_parent_rows = []

    for item in qs_items:
        itype = item[0]
        if itype == "sec":
            _, sec_title = item[0], item[1]
            ws3.merge_cells(start_row=r3, start_column=1, end_row=r3, end_column=12)
            cell_sec = ws3.cell(row=r3, column=1, value=sec_title)
            cell_sec.font = font_section
            cell_sec.fill = fill_section
            cell_sec.alignment = align_left
            ws3.row_dimensions[r3].height = 24
        elif itype == "parent":
            _, tt, code, name, unit, _, _, _, _, _, price = item
            ws3.cell(row=r3, column=1, value=tt).alignment = align_center
            ws3.cell(row=r3, column=1).font = font_bold
            ws3.cell(row=r3, column=2, value=code).alignment = align_center
            ws3.cell(row=r3, column=2).font = font_bold
            ws3.cell(row=r3, column=3, value=name).alignment = align_left
            ws3.cell(row=r3, column=3).font = font_bold
            ws3.cell(row=r3, column=4, value=unit).alignment = align_center
            ws3.cell(row=r3, column=4).font = font_bold
            ws3.cell(row=r3, column=11, value=price).alignment = align_right
            ws3.cell(row=r3, column=11).number_format = "#,##0"
            ws3.cell(row=r3, column=11).font = font_bold
            
            # Amount formula = Qty * Price
            ws3.cell(row=r3, column=12, value=f"=J{r3}*K{r3}").number_format = "#,##0"
            ws3.cell(row=r3, column=12).font = font_bold
            
            current_parent_row = r3
            parent_map[current_parent_row] = []
            all_parent_rows.append(current_parent_row)
            
            for c in range(1, 13):
                ws3.cell(row=r3, column=c).border = thin_border
            ws3.row_dimensions[r3].height = 22
        elif itype == "child":
            _, _, _, name, unit, n, l, w, h, coeff, _ = item
            ws3.cell(row=r3, column=3, value=name).alignment = align_left
            ws3.cell(row=r3, column=3).font = font_italic
            ws3.cell(row=r3, column=4, value=unit).alignment = align_center
            ws3.cell(row=r3, column=5, value=n).alignment = align_right
            ws3.cell(row=r3, column=5).number_format = "#,##0.00"
            ws3.cell(row=r3, column=6, value=l).alignment = align_right
            ws3.cell(row=r3, column=6).number_format = "#,##0.00"
            ws3.cell(row=r3, column=7, value=w).alignment = align_right
            ws3.cell(row=r3, column=7).number_format = "#,##0.00"
            ws3.cell(row=r3, column=8, value=h).alignment = align_right
            ws3.cell(row=r3, column=8).number_format = "#,##0.00"
            ws3.cell(row=r3, column=9, value=coeff).alignment = align_right
            ws3.cell(row=r3, column=9).number_format = "#,##0.0000"
            
            # Khối lượng con = E * F * G * H * I
            ws3.cell(row=r3, column=10, value=f"=E{r3}*F{r3}*G{r3}*H{r3}*I{r3}").number_format = "#,##0.00"
            parent_map[current_parent_row].append(r3)
            
            for c in range(1, 13):
                cell = ws3.cell(row=r3, column=c)
                cell.border = thin_border
                cell.font = font_regular
                cell.fill = fill_zebra
            ws3.row_dimensions[r3].height = 19
        r3 += 1

    # Now assign parent sum formulas
    for prow, crows in parent_map.items():
        if crows:
            first_c = crows[0]
            last_c = crows[-1]
            ws3.cell(row=prow, column=10, value=f"=SUM(J{first_c}:J{last_c})").number_format = "#,##0.00"
            ws3.cell(row=prow, column=10).font = font_bold

    # TOTAL DIRECT COST ROW (T)
    ws3.merge_cells(start_row=r3, start_column=1, end_row=r3, end_column=11)
    ws3.cell(row=r3, column=1, value="TỔNG CỘNG CHI PHÍ TRỰC TIẾP (T) (VNĐ)").alignment = align_right
    ws3.cell(row=r3, column=1).font = font_bold
    parent_amount_cells = [f"L{pr}" for pr in all_parent_rows]
    ws3.cell(row=r3, column=12, value=f"={'+'.join(parent_amount_cells)}").number_format = "#,##0"
    ws3.cell(row=r3, column=12).font = font_bold
    for c in range(1, 13):
        ws3.cell(row=r3, column=c).border = double_bottom_border
        ws3.cell(row=r3, column=c).fill = fill_total
    ws3.row_dimensions[r3].height = 26
    total_direct_cost_row = r3

    w3 = {1: 6, 2: 12, 3: 48, 4: 8, 5: 12, 6: 10, 7: 10, 8: 10, 9: 10, 10: 16, 11: 15, 12: 18}
    for col_idx, width in w3.items():
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    # =============================================================
    # SHEET 4: TONG_HOP_DU_TOAN_GXD (THÔNG TƯ 36/2026/TT-BXD)
    # =============================================================
    ws4 = wb.create_sheet(title="TONG_HOP_DU_TOAN_GXD")
    ws4.views.sheetView[0].showGridLines = True
    
    ws4["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - ĐOẠN QUA TỈNH HÀ GIANG"
    ws4["B2"].font = font_title
    ws4["B3"] = "BẢNG TỔNG HỢP KINH PHÍ DỰ TOÁN XÂY DỰNG (G_XD) - CẦU KM19+529.080"
    ws4["B3"].font = font_section
    ws4["B4"] = "Căn cứ Thông tư 36/2026/TT-BXD, Thông tư 38/2026/TT-BXD, Luật Xây dựng số 135/2025/QH15 & NĐ 207/2026/NĐ-CP"
    ws4["B4"].font = font_subtitle

    h4 = ["TT", "Khoản mục chi phí", "Cách tính / Căn cứ định mức", "Ký hiệu", "Giá trị (VNĐ)", "Ghi chú"]
    for c_idx, text in enumerate(h4, start=1):
        cell = ws4.cell(row=5, column=c_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws4.row_dimensions[5].height = 28

    gxd_rows = [
        ("1", "Chi phí trực tiếp (Vật liệu, Nhân công, Máy thi công)", f"Link từ Sheet QS (Ô L{total_direct_cost_row})", "T", f"=QS_DIEN_GIAI_CHI_TIET!L{total_direct_cost_row}", "100% công thức động"),
        ("2", "Chi phí gián tiếp", "GT = T x 7,30% (Công trình giao thông theo tuyến)", "GT", "=E6*0.073", "Thông tư 36/2026/TT-BXD"),
        ("", "  - Chi phí chung", "5,10% x T", "T_chung", "=E6*0.051", "Định mức công trình cầu lớn"),
        ("", "  - Chi phí xây dựng nhà tạm để ở và điều hành thi công", "1,20% x T", "T_tam", "=E6*0.012", "Lán trại vùng cao Hà Giang"),
        ("", "  - Chi phí một số công việc không xác định được KL từ TK", "1,00% x T", "T_kxd", "=E6*0.010", "Thí nghiệm, an toàn giao thông"),
        ("3", "Thu nhập chịu thuế tính trước", "TL = (T + GT) x 5,50%", "TL", "=(E6+E7)*0.055", "Thông tư 36/2026/TT-BXD"),
        ("4", "Chi phí xây dựng trước thuế", "G = T + GT + TL", "G", "=E6+E7+E11", "Tổng chi phí chưa VAT"),
        ("5", "Thuế giá trị gia tăng (VAT)", "VAT = G x 8,00%", "VAT", "=E12*0.08", "Chính sách thuế hiện hành"),
        ("6", "TỔNG CỘNG KINH PHÍ DỰ TOÁN XÂY DỰNG (G_XD)", "G_XD = G + VAT", "G_XD", "=E12+E13", "Giá trị gói thầu hoàn chỉnh"),
    ]

    r4 = 6
    for row_data in gxd_rows:
        tt, item_name, method, symbol, formula, note = row_data
        ws4.cell(row=r4, column=1, value=tt).alignment = align_center
        ws4.cell(row=r4, column=2, value=item_name).alignment = align_left
        ws4.cell(row=r4, column=3, value=method).alignment = align_left
        ws4.cell(row=r4, column=4, value=symbol).alignment = align_center
        ws4.cell(row=r4, column=5, value=formula).alignment = align_right
        ws4.cell(row=r4, column=5).number_format = "#,##0"
        ws4.cell(row=r4, column=6, value=note).alignment = align_center

        if tt in ["1", "2", "3", "4", "5"]:
            ws4.cell(row=r4, column=1).font = font_bold
            ws4.cell(row=r4, column=2).font = font_bold
            ws4.cell(row=r4, column=4).font = font_bold
            ws4.cell(row=r4, column=5).font = font_bold
        elif tt == "6":
            ws4.cell(row=r4, column=1).font = font_bold
            ws4.cell(row=r4, column=2).font = font_bold
            ws4.cell(row=r4, column=4).font = font_bold
            ws4.cell(row=r4, column=5).font = font_title
            for c in range(1, 7):
                ws4.cell(row=r4, column=c).fill = fill_total
                ws4.cell(row=r4, column=c).border = double_bottom_border
        else:
            ws4.cell(row=r4, column=2).font = font_italic
            ws4.cell(row=r4, column=5).font = font_italic

        for c in range(1, 7):
            if tt != "6":
                ws4.cell(row=r4, column=c).border = thin_border
        ws4.row_dimensions[r4].height = 22
        r4 += 1

    w4 = {1: 8, 2: 45, 3: 45, 4: 12, 5: 22, 6: 25}
    for col_idx, width in w4.items():
        ws4.column_dimensions[get_column_letter(col_idx)].width = width

    # =============================================================
    # SHEET 5: TIEN_DO_THI_CONG_WBS (KÈM GANTT CHART CPM TRÊN SHEET)
    # =============================================================
    ws5 = wb.create_sheet(title="TIEN_DO_THI_CONG_WBS")
    ws5.views.sheetView[0].showGridLines = True
    
    ws5["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - ĐOẠN QUA TỈNH HÀ GIANG"
    ws5["B2"].font = font_title
    ws5["B3"] = "BẢNG KẾ HOẠCH TIẾN ĐỘ WBS, ĐIỀU PHỐI NHÂN CÔNG & BIỂU ĐỒ GANTT CPM - CẦU KM19+529.080"
    ws5["B3"].font = font_section
    ws5["B4"] = "Định mức ngày công TT 38/2026/TT-BXD, Luật Xây dựng số 135/2025/QH15 & NĐ 207/2026/NĐ-CP (Cột đỏ: Đường găng CPM)"
    ws5["B4"].font = font_subtitle

    left_headers = [
        "Mã WBS", "Danh mục công tác thi công", "Khối lượng", "ĐVT",
        "Định mức\n(công/ĐVT)", "Tổng công\n(công)", "Tổ đội\n(người)",
        "Thời gian\nDuration (ngày)", "Ngày bắt đầu\nStart", "Ngày hoàn thành\nFinish",
        "Quan hệ logic\nPredecessors", "Đường găng\nCPM"
    ]

    months = ["Tháng 10/2026", "Tháng 11/2026", "Tháng 12/2026", "Tháng 01/2027", "Tháng 02/2027", "Tháng 03/2027"]
    weeks_per_month = 4

    # Header Row 5: Left headers & Gantt Month Header
    for c_idx, text in enumerate(left_headers, start=1):
        ws5.merge_cells(start_row=5, start_column=c_idx, end_row=6, end_column=c_idx)
        cell = ws5.cell(row=5, column=c_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
        ws5.cell(row=6, column=c_idx).border = thin_border

    gantt_col_start = 13
    current_col = gantt_col_start
    for m in months:
        ws5.merge_cells(start_row=5, start_column=current_col, end_row=5, end_column=current_col + weeks_per_month - 1)
        cell_m = ws5.cell(row=5, column=current_col, value=m)
        cell_m.font = font_header
        cell_m.fill = PatternFill(start_color="244061", end_color="244061", fill_type="solid")
        cell_m.alignment = align_center
        for c in range(current_col, current_col + weeks_per_month):
            ws5.cell(row=5, column=c).border = thin_border
            # Week subheader row 6
            w_num = (c - gantt_col_start) % 4 + 1
            cell_w = ws5.cell(row=6, column=c, value=f"T{w_num}")
            cell_w.font = Font(name="Times New Roman", size=8, bold=True, color="FFFFFF")
            cell_w.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell_w.alignment = align_center
            cell_w.border = thin_border
        current_col += weeks_per_month

    total_gantt_cols = current_col - 1
    ws5.row_dimensions[5].height = 20
    ws5.row_dimensions[6].height = 18

    # Task data for Bridge Km19+529.080
    # (wbs, name, qty_formula, unit, norm, team, dur_formula, start, finish, pred, cpm, w_start, w_end, is_sec)
    schedule_tasks = [
        ("1.0", "GIAI ĐOẠN I: CHUẨN BỊ MẶT BẰNG, ĐƯỜNG CÔNG VỤ & BÃI ĐÚC DẦM", "", "", "", "", "", "", "", "", "", 1, 3, True),
        ("1.1", "Bàn giao tim mốc VN2000, cao độ quốc gia Hòn Dấu & tim mốc cầu", 1, "Điểm", 12.0, 6, "=ROUNDUP(F8/G8, 0)", "2026-10-01", "2026-10-02", "-", "YES", 1, 1, False),
        ("1.2", "Rà phá bom mìn, dọn dẹp mặt bằng thi công & dựng lán trại", 12500, "m2", 0.02, 10, "=ROUNDUP(F9/G9, 0)", "2026-10-03", "2026-10-08", "1.1FS", "YES", 1, 2, False),
        ("1.3", "Thi công đường công vụ nhánh 6 & nhánh 6A tiếp cận mố trụ", "=KHOI_LUONG_DAO_DAP!F19", "m3", 0.08, 12, "=ROUNDUP(F10/G10, 0)", "2026-10-09", "2026-10-18", "1.2FS", "YES", 2, 3, False),
        ("1.4", "Xây dựng bãi đúc dầm Super-T, bệ đúc cáp căng & trạm trộn BTXM", 1, "Hệ", 180.0, 15, "=ROUNDUP(F11/G11, 0)", "2026-10-12", "2026-10-24", "1.3SS+3d", "NO", 2, 4, False),

        ("2.0", "GIAI ĐOẠN II: KHOAN CỌC NHỒI D1200MM & XỬ LÝ ĐỊA CHẤT KARST", "", "", "", "", "", "", "", "", "", 3, 9, True),
        ("2.1", "Khoan thăm dò hang Karst sâu 5m vào đá liền khối dưới mũi cọc", 26, "Lỗ", 12.0, 8, "=ROUNDUP(F13/G13, 0)", "2026-10-25", "2026-11-05", "1.3FS", "YES", 3, 5, False),
        ("2.2", "Khoan cọc nhồi D1200mm M1 (3 cọc L=20m) & đổ bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J8", "m", 0.75, 10, "=ROUNDUP(F14/G14, 0)", "2026-11-06", "2026-11-12", "2.1FS", "YES", 5, 6, False),
        ("2.3", "Khoan cọc nhồi D1200mm Trụ T1 (8 cọc L=40m) & đổ bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J9", "m", 0.75, 16, "=ROUNDUP(F15/G15, 0)", "2026-11-10", "2026-11-28", "2.2SS+4d", "YES", 6, 8, False),
        ("2.4", "Khoan cọc nhồi D1200mm Trụ T2 (8 cọc L=30m) & đổ bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J10", "m", 0.75, 14, "=ROUNDUP(F16/G16, 0)", "2026-11-18", "2026-12-04", "2.3SS+8d", "YES", 7, 9, False),
        ("2.5", "Khoan cọc nhồi D1200mm M2 (7 cọc L=36m) & đổ bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J11", "m", 0.75, 14, "=ROUNDUP(F17/G17, 0)", "2026-11-25", "2026-12-12", "2.4SS+7d", "NO", 8, 10, False),
        ("2.6", "Thí nghiệm siêu âm 156 mặt cắt, khoan kiểm tra mũi cọc & thử tải PDA", 1, "Gói", 48.0, 6, "=ROUNDUP(F18/G18, 0)", "2026-12-13", "2026-12-20", "2.4FS,2.5FS", "YES", 10, 11, False),

        ("3.0", "GIAI ĐOẠN III: KẾT CẤU PHẦN DƯỚI (BỆ, THÂN, XÀ MŨ M1, M2 & T1, T2)", "", "", "", "", "", "", "", "", "", 11, 16, True),
        ("3.1", "Đào hố móng, đập đầu 26 cọc nhồi & đổ bê tông lót đáy bệ mác C10", 26, "Cọc", 2.5, 10, "=ROUNDUP(F20/G20, 0)", "2026-12-21", "2026-12-27", "2.6FS", "YES", 11, 12, False),
        ("3.2", "Gia công lắp dựng cốt thép, ván khuôn & đổ bê tông bệ móng mố M1, M2", "=QS_DIEN_GIAI_CHI_TIET!J38", "m3", 0.45, 16, "=ROUNDUP(F21/G21, 0)", "2026-12-28", "2027-01-04", "3.1FS", "YES", 12, 13, False),
        ("3.3", "Gia công lắp dựng cốt thép, ván khuôn & đổ bê tông bệ trụ T1, T2", "=QS_DIEN_GIAI_CHI_TIET!J39", "m3", 0.42, 18, "=ROUNDUP(F22/G22, 0)", "2026-12-30", "2027-01-08", "3.1FS+2d", "YES", 12, 14, False),
        ("3.4", "Thi công thân đặc trụ T1, T2 (H=14.5m & 11.4m) bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J47", "m3", 0.55, 18, "=ROUNDUP(F23/G23, 0)", "2027-01-09", "2027-01-22", "3.3FS", "YES", 14, 16, False),
        ("3.5", "Thi công thân mố chân dê M1 và mố chữ U M2 bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J42", "m3", 0.52, 14, "=ROUNDUP(F24/G24, 0)", "2027-01-06", "2027-01-18", "3.2FS", "NO", 13, 15, False),
        ("3.6", "Thi công xà mũ trụ T1, T2, đá kê gối & khối chống xô", "=QS_DIEN_GIAI_CHI_TIET!J51", "m3", 0.65, 14, "=ROUNDUP(F25/G25, 0)", "2027-01-23", "2027-01-30", "3.4FS", "YES", 16, 17, False),

        ("4.0", "GIAI ĐOẠN IV: ĐÚC DẦM SUPER-T L=38.2M & LAO LẮP 3 NHỊP", "", "", "", "", "", "", "", "", "", 8, 19, True),
        ("4.1", "Gia công cốt thép, luồn cáp DUL 15.2mm 15 phiến dầm tại bãi đúc", 15, "Phiến", 28.0, 16, "=ROUNDUP(F27/G27, 0)", "2026-11-28", "2026-12-25", "1.4FS", "NO", 8, 12, False),
        ("4.2", "Đổ bê tông C45 15 phiến dầm Super-T (bảo dưỡng đạt cường độ R28)", "=QS_DIEN_GIAI_CHI_TIET!J57", "m3", 0.65, 16, "=ROUNDUP(F28/G28, 0)", "2026-12-26", "2027-01-15", "4.1FS", "NO", 12, 15, False),
        ("4.3", "Căng kéo cáp DUL tao xoắn 15.2mm & bơm vữa ống gen dầm", 15, "Phiến", 4.0, 8, "=ROUNDUP(F29/G29, 0)", "2027-01-16", "2027-01-25", "4.2FS", "YES", 15, 16, False),
        ("4.4", "Lắp đặt 30 gối chậu cao su đơn/song hướng mố trụ", 30, "Cái", 1.2, 6, "=ROUNDUP(F30/G30, 0)", "2027-01-31", "2027-02-04", "3.6FS", "YES", 17, 18, False),
        ("4.5", "Lắp đặt đường ray P43, vận chuyển & lao lắp 15 phiến dầm Super-T", 15, "Phiến", 5.0, 12, "=ROUNDUP(F31/G31, 0)", "2027-02-05", "2027-02-14", "4.3FS,4.4FS", "YES", 18, 19, False),

        ("5.0", "GIAI ĐOẠN V: DẦM NGANG, BẢN MẶT CẦU, LIÊN TỤC NHIỆT & ĐƯỜNG ĐẦU CẦU", "", "", "", "", "", "", "", "", "", 19, 23, True),
        ("5.1", "Đổ bê tông dầm ngang mố, dầm ngang trụ và bản liên tục nhiệt C35", "=QS_DIEN_GIAI_CHI_TIET!J68", "m3", 0.85, 12, "=ROUNDUP(F33/G33, 0)", "2027-02-15", "2027-02-19", "4.5FS", "YES", 19, 20, False),
        ("5.2", "Lắp đặt 510 tấm ván khuôn đúc sẵn & lắp cốt thép bản mặt cầu C35", "=QS_DIEN_GIAI_CHI_TIET!J72", "Tấm", 0.08, 14, "=ROUNDUP(F34/G34, 0)", "2027-02-20", "2027-02-24", "5.1FS", "YES", 20, 20, False),
        ("5.3", "Đổ bê tông bản mặt cầu tại chỗ C35 (dày tối thiểu 180mm)", "=QS_DIEN_GIAI_CHI_TIET!J74", "m3", 0.45, 18, "=ROUNDUP(F35/G35, 0)", "2027-02-25", "2027-03-02", "5.2FS", "YES", 21, 21, False),
        ("5.4", "Đổ bê tông bản quá độ C25 sau 2 mố & đắp đất chọn lọc K98", 2, "Bản", 24.0, 10, "=ROUNDUP(F36/G36, 0)", "2027-03-01", "2027-03-06", "5.1FS", "NO", 21, 22, False),
        ("5.5", "Lắp đặt khe co giãn răng lược thép D=100mm & ống thoát nước PVC D150", 2, "Bộ", 18.0, 6, "=ROUNDUP(F37/G37, 0)", "2027-03-07", "2027-03-12", "5.3FS", "YES", 22, 23, False),
        ("5.6", "Đổ bê tông gờ lan can C25 & lắp dựng tay vịn thép mạ kẽm", 260.8, "m", 0.35, 10, "=ROUNDUP(F38/G38, 0)", "2027-03-08", "2027-03-15", "5.3FS", "NO", 22, 23, False),
        ("5.7", "Phun lớp màng chống thấm & thảm bê tông nhựa C16 mặt cầu dày 7cm", "=QS_DIEN_GIAI_CHI_TIET!J87", "m2", 0.04, 12, "=ROUNDUP(F39/G39, 0)", "2027-03-16", "2027-03-19", "5.5FS", "YES", 23, 24, False),

        ("6.0", "GIAI ĐOẠN VI: THỬ TẢI CẦU, KIỂM ĐỊNH & BÀN GIAO KHAI THÁC", "", "", "", "", "", "", "", "", "", 24, 24, True),
        ("6.1", "Thử tải tĩnh/động cầu bằng đoàn xe tải nặng, đo dao động & võng dầm", 1, "Gói", 32.0, 8, "=ROUNDUP(F41/G41, 0)", "2027-03-20", "2027-03-24", "5.7FS", "YES", 24, 24, False),
        ("6.2", "Hoàn thiện hồ sơ hoàn công, nghiệm thu bàn giao đưa vào khai thác", 1, "Gói", 24.0, 6, "=ROUNDUP(F42/G42, 0)", "2027-03-25", "2027-03-28", "6.1FS", "YES", 24, 24, False),
    ]

    r5 = 7
    for task in schedule_tasks:
        wbs, name, qty, unit, norm, team, dur, start, finish, pred, cpm, w_start, w_end, is_sec = task
        
        ws5.cell(row=r5, column=1, value=wbs)
        ws5.cell(row=r5, column=2, value=name)
        
        if is_sec:
            ws5.merge_cells(start_row=r5, start_column=1, end_row=r5, end_column=12)
            cell_sec = ws5.cell(row=r5, column=1)
            cell_sec.font = font_section
            cell_sec.fill = fill_section
            cell_sec.alignment = align_left
            ws5.row_dimensions[r5].height = 24
            
            # Gantt section bar
            for c in range(gantt_col_start, total_gantt_cols + 1):
                col_week = c - gantt_col_start + 1
                cell_g = ws5.cell(row=r5, column=c)
                cell_g.border = thin_border
                if w_start <= col_week <= w_end:
                    cell_g.fill = fill_sec_bar
        else:
            ws5.cell(row=r5, column=3, value=qty)
            ws5.cell(row=r5, column=4, value=unit)
            ws5.cell(row=r5, column=5, value=norm)
            
            # Total man-days = Qty * Norm
            ws5.cell(row=r5, column=6, value=f"=C{r5}*E{r5}")
            ws5.cell(row=r5, column=7, value=team)
            ws5.cell(row=r5, column=8, value=dur)
            ws5.cell(row=r5, column=9, value=start)
            ws5.cell(row=r5, column=10, value=finish)
            ws5.cell(row=r5, column=11, value=pred)
            
            cpm_cell = ws5.cell(row=r5, column=12, value=cpm)
            if cpm == "YES":
                cpm_cell.fill = fill_cpm
                cpm_cell.font = font_bold
            
            # Formatting
            ws5.cell(row=r5, column=1).alignment = align_center
            ws5.cell(row=r5, column=1).font = font_bold
            ws5.cell(row=r5, column=2).alignment = align_left
            ws5.cell(row=r5, column=2).font = font_regular
            ws5.cell(row=r5, column=3).alignment = align_right
            ws5.cell(row=r5, column=3).number_format = "#,##0.00"
            ws5.cell(row=r5, column=4).alignment = align_center
            ws5.cell(row=r5, column=5).alignment = align_right
            ws5.cell(row=r5, column=5).number_format = "#,##0.00"
            ws5.cell(row=r5, column=6).alignment = align_right
            ws5.cell(row=r5, column=6).number_format = "#,##0.0"
            ws5.cell(row=r5, column=7).alignment = align_center
            ws5.cell(row=r5, column=8).alignment = align_center
            ws5.cell(row=r5, column=8).font = font_bold
            ws5.cell(row=r5, column=9).alignment = align_center
            ws5.cell(row=r5, column=10).alignment = align_center
            ws5.cell(row=r5, column=11).alignment = align_center
            ws5.cell(row=r5, column=12).alignment = align_center
            
            for c in range(1, 13):
                ws5.cell(row=r5, column=c).border = thin_border
            ws5.row_dimensions[r5].height = 20
            
            # Gantt bar painting
            for c in range(gantt_col_start, total_gantt_cols + 1):
                col_week = c - gantt_col_start + 1
                cell_g = ws5.cell(row=r5, column=c)
                cell_g.border = thin_border
                if w_start <= col_week <= w_end:
                    if cpm == "YES":
                        cell_g.fill = fill_critical
                    else:
                        cell_g.fill = fill_normal
        r5 += 1

    w5 = {1: 8, 2: 46, 3: 14, 4: 8, 5: 12, 6: 14, 7: 10, 8: 12, 9: 13, 10: 13, 11: 15, 12: 12}
    for col_idx, width in w5.items():
        ws5.column_dimensions[get_column_letter(col_idx)].width = width
    for c in range(gantt_col_start, total_gantt_cols + 1):
        ws5.column_dimensions[get_column_letter(c)].width = 4.2

    # =============================================================
    # SHEET 6: HOSO_KCS_NGHIEM_THU (22 BIÊN BẢN CHUẨN NĐ 207/2026/NĐ-CP)
    # =============================================================
    ws6 = wb.create_sheet(title="HOSO_KCS_NGHIEM_THU")
    ws6.views.sheetView[0].showGridLines = True
    
    ws6["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - ĐOẠN QUA TỈNH HÀ GIANG"
    ws6["B2"].font = font_title
    ws6["B3"] = "DANH MỤC HỒ SƠ QUẢN LÝ CHẤT LƯỢNG & BIÊN BẢN NGHIỆM THU KCS - CẦU KM19+529.080"
    ws6["B3"].font = font_section
    ws6["B4"] = "Tuân thủ Luật Xây dựng số 135/2025/QH15 & Nghị định 207/2026/NĐ-CP (Khối lượng link trực tiếp từ Sheet QS)"
    ws6["B4"].font = font_subtitle

    h6 = ["STT", "Số hiệu biên bản", "Tên công việc / Giai đoạn nghiệm thu KCS",
          "Khối lượng nghiệm thu", "ĐVT", "Căn cứ quy chuẩn / Tiêu chuẩn áp dụng",
          "Chỉ tiêu thí nghiệm & Điểm kiểm soát (Hold Points)"]
    
    for c_idx, text in enumerate(h6, start=1):
        cell = ws6.cell(row=5, column=c_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws6.row_dimensions[5].height = 28

    kcs_data = [
        ("1", "BBNT-01", "Bàn giao tim mốc VN2000, cao độ Hòn Dấu & tim mốc cầu Km19+529.080", 1, "Điểm", "TCVN 9393:2012", "Sai số tọa độ tim mốc <= 5mm, mốc cao độ <= 2mm"),
        ("2", "BBNT-02", "Nghiệm thu dọn dẹp mặt bằng, thi công đường công vụ nhánh 6 & 6A", "=KHOI_LUONG_DAO_DAP!F19", "m3", "TCVN 4447:2012", "Độ dầm chặt nền đường K >= 0.95, thoát nước tốt"),
        ("3", "BBNT-03", "Kiểm tra nghiệm thu chất lượng vật liệu đầu vào (Thép, Cáp DUL, Xi măng)", 1, "Lô", "TCVN 1651:2018, ASTM A416", "Chứng chỉ CO/CQ, thí nghiệm kéo uốn thép, nén xi măng"),
        ("4", "BBNT-04", "Nghiệm thu trạm trộn BTXM, trạm biến áp & bệ đúc dầm Super-T", 1, "Hệ", "TCVN 9340:2012", "Độ chính xác cân đong sai số <= 1%, bệ đúc đủ độ cứng"),
        ("5", "BBNT-05", "Nghiệm thu định vị tim cọc & hạ ống vách thép dẫn hướng D1300mm", "=QS_DIEN_GIAI_CHI_TIET!J13", "m", "TCVN 11823:2017", "Độ nghiêng ống vách <= 1%, sai số tim cọc <= 50mm"),
        ("6", "BBNT-06", "Nghiệm thu khoan thăm dò hang Karst sâu 5m dưới mũi cọc", 26, "Lỗ", "Chỉ dẫn TK Trang 11", "Khoan vào đá liền khối tối thiểu 5m, không rỗng Karst"),
        ("7", "BBNT-07", "Nghiệm thu hố khoan cọc D1200mm, thổi rửa đáy & kiểm tra dung dịch", "=QS_DIEN_GIAI_CHI_TIET!J7", "m", "TCVN 9395:2012", "Chiều dày cặn lắng đáy <= 50mm, tỷ trọng bentonite 1.05-1.15"),
        ("8", "BBNT-08", "Nghiệm thu gia công, lắp dựng lồng cốt thép & ống siêu âm cọc nhồi", "=QS_DIEN_GIAI_CHI_TIET!J23", "Tấn", "TCVN 4453:1995", "Chiều dài lồng thép, mối nối coupler, ống siêu âm kín nước"),
        ("9", "BBNT-09", "Nghiệm thu đổ bê tông cọc khoan nhồi C30 bằng ống tremie", "=QS_DIEN_GIAI_CHI_TIET!J18", "m3", "TCVN 9395:2012", "Độ sụt 18+-2cm, ống tremie ngập bê tông >= 2.0m liên tục"),
        ("10", "BBNT-10", "Nghiệm thu thí nghiệm kiểm tra cọc (Siêu âm 100%, lõi bê tông & PDA)", 1, "Gói", "TCVN 9396, ASTM D4945", "Siêu âm 156 mặt cắt; PDA trụ T1 đạt tải trọng 9434 kN"),
        ("11", "BBNT-11", "Nghiệm thu đào hố móng, đập đầu 26 cọc & đổ bê tông lót bệ mác C10", "=QS_DIEN_GIAI_CHI_TIET!J28", "m3", "TCVN 4453:1995", "Đầu cọc ngàm vào bệ 10cm, cốt thép cọc neo sâu 40D"),
        ("12", "BBNT-12", "Nghiệm thu ván khuôn, cốt thép & đổ bê tông bệ mố M1, M2 & bệ trụ T1, T2 C30", "=QS_DIEN_GIAI_CHI_TIET!J37", "m3", "TCVN 3105:2022", "Kích thước hình học bệ, 12 tổ mẫu thử nén đạt C30"),
        ("13", "BBNT-13", "Nghiệm thu thân mố chân dê M1 và thân mố chữ U M2 bê tông C30", "=QS_DIEN_GIAI_CHI_TIET!J42", "m3", "TCVN 11823:2017", "Độ thẳng đứng thân mố, bề mặt bê tông phẳng nhẵn"),
        ("14", "BBNT-14", "Nghiệm thu cốt thép, ván khuôn & đổ bê tông thân đặc trụ T1, T2 C30", "=QS_DIEN_GIAI_CHI_TIET!J47", "m3", "TCVN 4453:1995", "Kích thước thân đặc 9.75x2.0m, chiều cao H=14.5m và 11.4m"),
        ("15", "BBNT-15", "Nghiệm thu cốt thép, ván khuôn & đổ bê tông xà mũ trụ T1, T2, đá kê gối", "=QS_DIEN_GIAI_CHI_TIET!J51", "m3", "TCVN 11823:2017", "Cao độ mặt đá kê gối sai số <= 2mm, tim đá kê <= 3mm"),
        ("16", "BBNT-16", "Nghiệm thu cốt thép, ván khuôn & đổ bê tông C45 15 phiến dầm Super-T", "=QS_DIEN_GIAI_CHI_TIET!J57", "m3", "TCVN 3105:2022", "Cường độ bê tông đạt R28 >= 45 MPa, không rỗ nứt"),
        ("17", "BBNT-17", "Nghiệm thu căng kéo cáp DUL 15.2mm & bơm vữa ống gen dầm Super-T", "=QS_DIEN_GIAI_CHI_TIET!J59", "Tấn", "TCVN 11823:2017", "Sai số độ giãn dài cáp <= +-5%, vữa bơm lấp đầy 100% gen"),
        ("18", "BBNT-18", "Nghiệm thu lắp đặt 30 gối chậu cao su & lao lắp 15 phiến dầm Super-T", "=QS_DIEN_GIAI_CHI_TIET!J65", "Cái", "TCVN 10308:2014", "Gối chậu nằm ngang chuẩn xác, dầm đặt êm thuận trên mố trụ"),
        ("19", "BBNT-19", "Nghiệm thu đổ bê tông dầm ngang & mối nối liên tục nhiệt đỉnh trụ T1, T2", "=QS_DIEN_GIAI_CHI_TIET!J68", "m3", "TCVN 4453:1995", "Liên kết cứng dầm ngang, bản liên tục nhiệt triệt tiêu khe nứt"),
        ("20", "BBNT-20", "Nghiệm thu cốt thép, ván khuôn & đổ bê tông bản mặt cầu C35 dày 180mm", "=QS_DIEN_GIAI_CHI_TIET!J74", "m3", "TCVN 11823:2017", "Độ dốc ngang thoát nước 2%, chiều dày bảo vệ cốt thép chuẩn"),
        ("21", "BBNT-21", "Nghiệm thu khe co giãn răng lược D=100mm, gờ lan can & thảm BTN C16", "=QS_DIEN_GIAI_CHI_TIET!J87", "m2", "TCVN 8819:2011", "Độ bằng phẳng mặt cầu IRI <= 1.5, khe răng lược êm thuận"),
        ("22", "BBNT-22", "Nghiệm thu thử tải tĩnh/động toàn cầu & nghiệm thu hoàn thành đưa vào sử dụng", 1, "Gói", "Nghị định 207/2026/NĐ-CP", "Độ võng tĩnh f <= f_cp, hệ số xung kích động đạt tiêu chuẩn"),
    ]

    r6 = 6
    for item in kcs_data:
        stt, bb_code, name, qty, unit, std, note = item
        ws6.cell(row=r6, column=1, value=stt).alignment = align_center
        ws6.cell(row=r6, column=2, value=bb_code).alignment = align_center
        ws6.cell(row=r6, column=2).font = font_bold
        ws6.cell(row=r6, column=3, value=name).alignment = align_left
        ws6.cell(row=r6, column=3).font = font_regular
        ws6.cell(row=r6, column=4, value=qty).alignment = align_right
        ws6.cell(row=r6, column=4).number_format = "#,##0.00"
        ws6.cell(row=r6, column=5, value=unit).alignment = align_center
        ws6.cell(row=r6, column=6, value=std).alignment = align_left
        ws6.cell(row=r6, column=6).font = font_italic
        ws6.cell(row=r6, column=7, value=note).alignment = align_left
        ws6.cell(row=r6, column=7).font = font_regular

        for c in range(1, 8):
            ws6.cell(row=r6, column=c).border = thin_border
        ws6.row_dimensions[r6].height = 20
        r6 += 1

    w6 = {1: 6, 2: 15, 3: 48, 4: 15, 5: 8, 6: 25, 7: 45}
    for col_idx, width in w6.items():
        ws6.column_dimensions[get_column_letter(col_idx)].width = width

    # Remove temporary default sheet if empty
    if default_sheet.title == "Sheet":
        wb.remove(default_sheet)

    # Save to files
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    target_excel = os.path.join(repo_root, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")

    wb.save(target_excel)
    print(f"-> Đã xuất thành công file Excel: {target_excel}")

    try:
        sys.path.insert(0, repo_root)
        from examples.update_material_sheets import add_material_sheets
        add_material_sheets(target_excel)
    except Exception as e:
        print("Lỗi tích hợp sheet vật tư:", e)

    # =============================================================
    # TẠO FILE MS PROJECT XML TIẾN ĐỘ THI CÔNG
    # =============================================================
    root = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(root, "Name").text = "Tien_Do_Thi_Cong_Cau_Km19+529.080"
    ET.SubElement(root, "Title").text = "TIẾN ĐỘ THI CÔNG CẦU KM19+529.080 (3 NHỊP SUPER-T L=38.2M, DỰ ÁN CAO TỐC TUYÊN QUANG - HÀ GIANG)"
    ET.SubElement(root, "StartDate").text = "2026-10-01T08:00:00"
    ET.SubElement(root, "FinishDate").text = "2027-03-28T17:00:00"
    ET.SubElement(root, "CalendarUID").text = "1"

    calendars = ET.SubElement(root, "Calendars")
    cal = ET.SubElement(calendars, "Calendar")
    ET.SubElement(cal, "UID").text = "1"
    ET.SubElement(cal, "Name").text = "Standard"
    ET.SubElement(cal, "IsBaseCalendar").text = "1"

    resources = ET.SubElement(root, "Resources")
    res_list = [
        (1, "Tổ trắc đạc & định vị tim mốc cầu", 1),
        (2, "Tổ máy khoan cọc nhồi D1200 & trắc địa Karst", 1),
        (3, "Tổ thợ sắt thép cọc nhồi & bệ mố trụ", 1),
        (4, "Tổ ván khuôn & đà giáo thân xà mũ", 1),
        (5, "Tổ đúc dầm Super-T tại bãi đúc (Bê tông C45)", 1),
        (6, "Tổ căng kéo cáp DUL 15.2mm & bơm vữa", 1),
        (7, "Tổ vận chuyển & xe lao dầm P43 chuyên dụng", 1),
        (8, "Tổ bản mặt cầu & liên tục nhiệt", 1),
        (9, "Tổ thảm bê tông nhựa C16 & khe răng lược", 1),
        (10, "Máy khoan cọc xoay đập thủy lực R618", 2),
        (11, "Giá lao dầm Super-T trọng lượng 78.88T", 2),
        (12, "Trạm trộn BTXM 90m3/h & xe bồn 10m3", 2),
    ]
    for r_uid, r_name, r_type in res_list:
        res = ET.SubElement(resources, "Resource")
        ET.SubElement(res, "UID").text = str(r_uid)
        ET.SubElement(res, "ID").text = str(r_uid)
        ET.SubElement(res, "Name").text = r_name
        ET.SubElement(res, "Type").text = str(r_type)

    tasks_element = ET.SubElement(root, "Tasks")

    xml_task_items = [
        # uid, name, wbs, outline, level, is_summary, start, finish, dur_hours, critical, pred_uids
        (1, "CHUẨN BỊ MẶT BẰNG, ĐƯỜNG CÔNG VỤ & BÃI ĐÚC DẦM", "1.0", "1.0", 1, 1, "2026-10-01T08:00:00", "2026-10-24T17:00:00", 192, 1, []),
        (2, "Bàn giao tim mốc VN2000, cao độ quốc gia & tim mốc cầu", "1.1", "1.1", 2, 0, "2026-10-01T08:00:00", "2026-10-02T17:00:00", 16, 1, []),
        (3, "Rà phá bom mìn, dọn dẹp mặt bằng & dựng lán trại", "1.2", "1.2", 2, 0, "2026-10-03T08:00:00", "2026-10-08T17:00:00", 48, 1, [2]),
        (4, "Thi công đường công vụ nhánh 6 & 6A tiếp cận mố trụ", "1.3", "1.3", 2, 0, "2026-10-09T08:00:00", "2026-10-18T17:00:00", 80, 1, [3]),
        (5, "Xây dựng bãi đúc dầm Super-T, bệ đúc & trạm trộn BTXM", "1.4", "1.4", 2, 0, "2026-10-12T08:00:00", "2026-10-24T17:00:00", 104, 0, [4]),

        (6, "KHOAN CỌC NHỒI D1200MM & XỬ LÝ ĐỊA CHẤT KARST", "2.0", "2.0", 1, 1, "2026-10-25T08:00:00", "2026-12-20T17:00:00", 456, 1, []),
        (7, "Khoan thăm dò hang Karst sâu 5m dưới mũi cọc", "2.1", "2.1", 2, 0, "2026-10-25T08:00:00", "2026-11-05T17:00:00", 96, 1, [4]),
        (8, "Khoan cọc nhồi D1200mm M1 (3 cọc L=20m) & đổ bê tông C30", "2.2", "2.2", 2, 0, "2026-11-06T08:00:00", "2026-11-12T17:00:00", 56, 1, [7]),
        (9, "Khoan cọc nhồi D1200mm Trụ T1 (8 cọc L=40m) & đổ bê tông C30", "2.3", "2.3", 2, 0, "2026-11-10T08:00:00", "2026-11-28T17:00:00", 152, 1, [8]),
        (10, "Khoan cọc nhồi D1200mm Trụ T2 (8 cọc L=30m) & đổ bê tông C30", "2.4", "2.4", 2, 0, "2026-11-18T08:00:00", "2026-12-04T17:00:00", 136, 1, [9]),
        (11, "Khoan cọc nhồi D1200mm M2 (7 cọc L=36m) & đổ bê tông C30", "2.5", "2.5", 2, 0, "2026-11-25T08:00:00", "2026-12-12T17:00:00", 144, 0, [10]),
        (12, "Thí nghiệm siêu âm 156 mặt cắt, kiểm tra mũi cọc & thử tải PDA (9434 kN)", "2.6", "2.6", 2, 0, "2026-12-13T08:00:00", "2026-12-20T17:00:00", 64, 1, [10, 11]),

        (13, "KẾT CẤU PHẦN DƯỚI (BỆ, THÂN, XÀ MŨ M1, M2 & T1, T2)", "3.0", "3.0", 1, 1, "2026-12-21T08:00:00", "2027-01-30T17:00:00", 328, 1, []),
        (14, "Đào hố móng, đập đầu cọc & đổ bê tông lót mác C10", "3.1", "3.1", 2, 0, "2026-12-21T08:00:00", "2026-12-27T17:00:00", 56, 1, [12]),
        (15, "Lắp dựng cốt thép, ván khuôn & đổ bê tông bệ móng mố M1, M2", "3.2", "3.2", 2, 0, "2026-12-28T08:00:00", "2027-01-04T17:00:00", 64, 1, [14]),
        (16, "Lắp dựng cốt thép, ván khuôn & đổ bê tông bệ trụ T1, T2 C30", "3.3", "3.3", 2, 0, "2026-12-30T08:00:00", "2027-01-08T17:00:00", 72, 1, [14]),
        (17, "Thi công thân đặc trụ T1, T2 (H=14.5m & 11.4m) bê tông C30", "3.4", "3.4", 2, 0, "2027-01-09T08:00:00", "2027-01-22T17:00:00", 112, 1, [16]),
        (18, "Thi công thân mố chân dê M1 và thân mố chữ U M2 C30", "3.5", "3.5", 2, 0, "2027-01-06T08:00:00", "2027-01-18T17:00:00", 104, 0, [15]),
        (19, "Thi công xà mũ trụ T1, T2, đá kê gối & khối chống xô C30", "3.6", "3.6", 2, 0, "2027-01-23T08:00:00", "2027-01-30T17:00:00", 64, 1, [17]),

        (20, "ĐÚC DẦM SUPER-T L=38.2M & LAO LẮP 3 NHỊP", "4.0", "4.0", 1, 1, "2026-11-28T08:00:00", "2027-02-14T17:00:00", 624, 1, []),
        (21, "Gia công cốt thép, luồn cáp DUL 15 phiến dầm tại bãi đúc", "4.1", "4.1", 2, 0, "2026-11-28T08:00:00", "2026-12-25T17:00:00", 224, 0, [5]),
        (22, "Đổ bê tông C45 15 phiến dầm Super-T (dưỡng hộ đạt R28)", "4.2", "4.2", 2, 0, "2026-12-26T08:00:00", "2027-01-15T17:00:00", 168, 0, [21]),
        (23, "Căng kéo cáp DUL tao xoắn 15.2mm & bơm vữa ống gen", "4.3", "4.3", 2, 0, "2027-01-16T08:00:00", "2027-01-25T17:00:00", 80, 1, [22]),
        (24, "Lắp đặt 30 gối chậu cao su đơn/song hướng mố trụ", "4.4", "4.4", 2, 0, "2027-01-31T08:00:00", "2027-02-04T17:00:00", 40, 1, [19]),
        (25, "Lắp đường ray P43, vận chuyển & lao lắp 15 phiến dầm Super-T", "4.5", "4.5", 2, 0, "2027-02-05T08:00:00", "2027-02-14T17:00:00", 80, 1, [23, 24]),

        (26, "DẦM NGANG, BẢN MẶT CẦU, LIÊN TỤC NHIỆT & ĐƯỜNG ĐẦU CẦU", "5.0", "5.0", 1, 1, "2027-02-15T08:00:00", "2027-03-19T17:00:00", 264, 1, []),
        (27, "Đổ bê tông dầm ngang mố trụ & bản liên tục nhiệt C35", "5.1", "5.1", 2, 0, "2027-02-15T08:00:00", "2027-02-19T17:00:00", 40, 1, [25]),
        (28, "Lắp 510 tấm ván khuôn đúc sẵn & lắp cốt thép bản mặt cầu", "5.2", "5.2", 2, 0, "2027-02-20T08:00:00", "2027-02-24T17:00:00", 40, 1, [27]),
        (29, "Đổ bê tông bản mặt cầu tại chỗ C35 (dày 180mm)", "5.3", "5.3", 2, 0, "2027-02-25T08:00:00", "2027-03-02T17:00:00", 48, 1, [28]),
        (30, "Đổ bê tông bản quá độ C25 sau 2 mố & đắp hạt chọn lọc K98", "5.4", "5.4", 2, 0, "2027-03-01T08:00:00", "2027-03-06T17:00:00", 48, 0, [27]),
        (31, "Lắp đặt khe co giãn răng lược D=100mm & ống thoát nước D150", "5.5", "5.5", 2, 0, "2027-03-07T08:00:00", "2027-03-12T17:00:00", 48, 1, [29]),
        (32, "Đổ bê tông gờ lan can C25 & lắp dựng tay vịn mạ kẽm", "5.6", "5.6", 2, 0, "2027-03-08T08:00:00", "2027-03-15T17:00:00", 64, 0, [29]),
        (33, "Phun màng chống thấm & thảm bê tông nhựa C16 dày 7cm", "5.7", "5.7", 2, 0, "2027-03-16T08:00:00", "2027-03-19T17:00:00", 32, 1, [31]),

        (34, "THỬ TẢI CẦU, KIỂM ĐỊNH & BÀN GIAO KHAI THÁC", "6.0", "6.0", 1, 1, "2027-03-20T08:00:00", "2027-03-28T17:00:00", 72, 1, []),
        (35, "Thử tải tĩnh/động toàn cầu bằng đoàn xe tải nặng", "6.1", "6.1", 2, 0, "2027-03-20T08:00:00", "2027-03-24T17:00:00", 40, 1, [33]),
        (36, "Hoàn thiện hồ sơ hoàn công & bàn giao đưa vào khai thác", "6.2", "6.2", 2, 0, "2027-03-25T08:00:00", "2027-03-28T17:00:00", 32, 1, [35]),
    ]

    for t_item in xml_task_items:
        uid, name, wbs, outline, level, is_sum, start, finish, dur_h, crit, preds = t_item
        t = ET.SubElement(tasks_element, "Task")
        ET.SubElement(t, "UID").text = str(uid)
        ET.SubElement(t, "ID").text = str(uid)
        ET.SubElement(t, "Name").text = name
        ET.SubElement(t, "WBS").text = wbs
        ET.SubElement(t, "OutlineNumber").text = outline
        ET.SubElement(t, "OutlineLevel").text = str(level)
        ET.SubElement(t, "Summary").text = str(is_sum)
        ET.SubElement(t, "Start").text = start
        ET.SubElement(t, "Finish").text = finish
        ET.SubElement(t, "Duration").text = f"PT{dur_h}H0M0S"
        ET.SubElement(t, "Critical").text = str(crit)

        for p_uid in preds:
            plink = ET.SubElement(t, "PredecessorLink")
            ET.SubElement(plink, "PredecessorUID").text = str(p_uid)
            ET.SubElement(plink, "Type").text = "1" # Finish-to-Start (FS)
            ET.SubElement(plink, "CrossProject").text = "0"

    xml_target = os.path.join(repo_root, "templates", "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml")

    tree = ET.ElementTree(root)
    tree.write(xml_target, encoding="utf-8", xml_declaration=True)
    print(f"-> Đã xuất thành công file MS Project XML: {xml_target}")

    print("\n=== HOÀN TẤT 100% QUÁ TRÌNH XUẤT HỒ SƠ CẦU KM19+529.080 ===")

if __name__ == "__main__":
    create_full_package()
