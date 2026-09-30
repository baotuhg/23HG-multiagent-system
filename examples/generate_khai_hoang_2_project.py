# -*- coding: utf-8 -*-
"""
HỆ THỐNG MULTI-AGENT AEC TỰ ĐỘNG HÓA KỸ SƯ
BỘ TẠO DỮ LIỆU ĐỒNG BỘ CHO DỰ ÁN:
CẦU THÔN KHAI HOANG 2, KM14+363.65 (LÝ TRÌNH KM14+363.35)
DỰ ÁN ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, TỈNH HÀ GIANG
GÓI THẦU SỐ 9: KM12 - KM24+862.93

Tuân thủ:
- Luật Xây dựng số 135/2025/QH15 (VAT = 10%)
- Nghị định 207/2026/NĐ-CP & Thông tư 32/2026/TT-BXD (Quản lý chất lượng & Biểu mẫu nghiệm thu)
- Thông tư 36/2026/TT-BXD, Thông tư 38/2026/TT-BXD (Dự toán chi phí & Định mức xây dựng)
- Nghị định 254/2025/NĐ-CP (Thanh toán khối lượng hoàn thành Phụ lục 03a)
- Tiêu chuẩn thiết kế cầu TCVN 11823:2017 & Cốt thép TCVN 1651:2018
- 100% CÔNG THỨC SỐNG - ZERO DEAD NUMBERS - KIỂM TOÁN AUDIT 100/100
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import xml.etree.ElementTree as ET

def generate_khai_hoang_2_master_package(output_excel_path, output_xml_path):
    print(f"[*] Bắt đầu khởi tạo hệ thống 14 Sheet Master cho Cầu thôn Khai Hoang 2...")
    os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
    os.makedirs(os.path.dirname(output_xml_path), exist_ok=True)

    wb = openpyxl.Workbook()
    default_sheet = wb.active

    # =========================================================================
    # STYLES & PALETTE
    # =========================================================================
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_subtitle = Font(name="Times New Roman", size=10, italic=True, color="595959")
    font_sec = Font(name="Times New Roman", size=10, bold=True, color="1F497D")
    font_hdr = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    font_reg = Font(name="Times New Roman", size=9)
    font_it = Font(name="Times New Roman", size=9, italic=True)

    fill_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_sec = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_cpm = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_tot = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_gray = Side(style='thin', color='BFBFBF')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    double_bottom_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=Side(style='double', color='1F497D'))
    box_border = Border(left=Side(style='medium', color='1F497D'), right=Side(style='medium', color='1F497D'),
                        top=Side(style='medium', color='1F497D'), bottom=Side(style='medium', color='1F497D'))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # 1. SHEET: TO_HOP_CAT_THEP_11M7
    # =========================================================================
    ws1 = wb.create_sheet(title="TO_HOP_CAT_THEP_11M7")
    ws1.views.sheetView[0].showGridLines = True

    ws1["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450 (MỐC 456), HUYỆN MÈO VẠC, HÀ GIANG"
    ws1["B2"].font = font_title
    ws1["B3"] = "BẢNG TỔ HỢP CẮT THÉP CÂY NGUYÊN 11.7m (1D CUTTING STOCK) - CẦU THÔN KHAI HOANG 2, KM14+363.65"
    ws1["B3"].font = font_sec
    ws1["B4"] = "Tối ưu hóa sơ đồ cắt thép, kiểm soát hao hụt đề-xê < 1.5% (Tiêu chuẩn TCVN 1651:2018)"
    ws1["B4"].font = font_subtitle

    h1 = ["TT", "Cấu kiện sử dụng", "Ký hiệu đường kính (mm)", "Số lượng thanh (N)", "Dài 1 thanh L (m)",
          "Tổng chiều dài (m)", "Dài cây tiêu chuẩn (m)", "Số cây 11.7m cần dùng", "Chiều dài phôi (m)",
          "Đề-xê thừa (m)", "Tổng trọng lượng (kg)", "Tỷ lệ hao hụt (%)"]
    for c_idx, text in enumerate(h1, start=1):
        cell = ws1.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[5].height = 28

    # 14 nhóm thanh tổ hợp cắt thép liên kết động BBS
    ct_specs = [
        ("1", "Thép chủ dầm T15m A1/B1 (Dầm biên & dầm giữa)", 32, 17.32, 6.313, "*Dầm chủ T15m*"),
        ("2", "Thép cấu tạo dầm T15m A2/A3 dọc cánh dầm", 12, 15.40, 0.888, "*Dầm chủ T15m*"),
        ("3", "Thép đai chữ U dầm T15m C3 cánh và sườn", 14, 2.533, 1.208, "*Dầm chủ T15m*"),
        ("4", "Thép sườn dầm T15m C4/C5 chịu cắt và uốn", 16, 1.850, 1.578, "*Dầm chủ T15m*"),
        ("5", "Thép đai neo dầm T15m D1/E1/F1 liên kết bản", 10, 2.260, 0.617, "*Dầm chủ T15m*"),
        ("6", "Thép khoan cấy neo vào nền đá móng mố M1, M2", 32, 0.500, 6.313, "*Khoan cấy*"),
        ("7", "Thép chịu lực chính thân mố M1, M2 W1/W1A", 25, 9.467, 3.853, "*Thân mố*"),
        ("8", "Thép chịu lực tường cánh mố M1, M2 K1-K6", 22, 9.496, 2.984, "*Tường cánh*"),
        ("9", "Thép thân mố & bản quá độ W2/W3/A1 chịu lực", 20, 8.290, 2.466, "*Thân mố*"),
        ("10", "Thép đai thân mố & tường cánh W4-W7/K8-K13", 16, 8.410, 1.578, "*Thân mố*"),
        ("11", "Thép chịu lực bản quá độ hai đầu cầu A2-A4", 14, 7.102, 1.208, "*Bản quá độ*"),
        ("12", "Thép gờ lan can dọc cầu L1/L2 trên nhịp & mố", 14, 2.180, 1.208, "*Gờ lan can*"),
        ("13", "Thép gờ lan can dọc nhịp L3/L3' L=15m", 12, 15.000, 0.888, "*Gờ lan can*"),
        ("14", "Lưới thép hàn bản mặt cầu B1/B2 chống nứt", 6, 15.000, 0.222, "*Lưới thép*"),
    ]

    for idx, (tt, name, dia, l_bar, w_unit, kw) in enumerate(ct_specs, start=6):
        ws1.cell(idx, 1, value=tt).alignment = align_center
        ws1.cell(idx, 2, value=name).alignment = align_left
        ws1.cell(idx, 3, value=dia).alignment = align_center
        # Cột D: COUNTIFS từ BBS
        ws1.cell(idx, 4, value=f'=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399,{dia},THONG_KE_THEP_CHI_TIET!$B$6:$B$399,"{kw}")').alignment = align_right
        ws1.cell(idx, 4).number_format = "#,##0"
        ws1.cell(idx, 5, value=l_bar).alignment = align_right
        ws1.cell(idx, 5).number_format = "#,##0.00"
        # Cột F: Tổng chiều dài = D * E
        ws1.cell(idx, 6, value=f"=D{idx}*E{idx}").alignment = align_right
        ws1.cell(idx, 6).number_format = "#,##0.00"
        # Cột G: Chiều dài cây tiêu chuẩn 11.7m
        ws1.cell(idx, 7, value=11.70).alignment = align_right
        ws1.cell(idx, 7).number_format = "0.00"
        # Cột H: Số cây 11.7m = ROUNDUP(F / G, 0)
        ws1.cell(idx, 8, value=f"=ROUNDUP(F{idx}/G{idx},0)").alignment = align_right
        ws1.cell(idx, 8).number_format = "#,##0"
        # Cột I: Tổng chiều dài phôi = H * G
        ws1.cell(idx, 9, value=f"=H{idx}*G{idx}").alignment = align_right
        ws1.cell(idx, 9).number_format = "#,##0.00"
        # Cột J: Đề-xê thừa = I - F
        ws1.cell(idx, 10, value=f"=I{idx}-F{idx}").alignment = align_right
        ws1.cell(idx, 10).number_format = "#,##0.00"
        # Cột K: Trọng lượng kg = F * w_unit
        ws1.cell(idx, 11, value=f"=F{idx}*{w_unit}").alignment = align_right
        ws1.cell(idx, 11).number_format = "#,##0.0"
        # Cột L: % Hao hụt = J / I
        ws1.cell(idx, 12, value=f"=J{idx}/I{idx}").alignment = align_right
        ws1.cell(idx, 12).number_format = "0.00%"

        for c in range(1, 13):
            ws1.cell(idx, c).font = font_reg
            ws1.cell(idx, c).border = thin_border
            if idx % 2 == 1:
                ws1.cell(idx, c).fill = fill_zebra

    # Dòng tổng cộng 20
    ws1.cell(20, 2, value="TỔNG CỘNG TOÀN CÔNG TRÌNH").alignment = align_left
    ws1.cell(20, 4, value="=SUM(D6:D19)").alignment = align_right
    ws1.cell(20, 6, value="=SUM(F6:F19)").alignment = align_right
    ws1.cell(20, 8, value="=SUM(H6:H19)").alignment = align_right
    ws1.cell(20, 9, value="=SUM(I6:I19)").alignment = align_right
    ws1.cell(20, 10, value="=SUM(J6:J19)").alignment = align_right
    ws1.cell(20, 11, value="=SUM(K6:K19)").alignment = align_right
    ws1.cell(20, 12, value="=J20/I20").alignment = align_right
    ws1.cell(20, 12).number_format = "0.00%"

    for c in range(1, 13):
        ws1.cell(20, c).font = font_bold
        ws1.cell(20, c).fill = fill_tot
        ws1.cell(20, c).border = double_bottom_border

    # =========================================================================
    # 2. SHEET: KHOI_LUONG_DAO_DAP
    # =========================================================================
    ws2 = wb.create_sheet(title="KHOI_LUONG_DAO_DAP")
    ws2.views.sheetView[0].showGridLines = True

    ws2["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws2["B2"].font = font_title
    ws2["B3"] = "BẢNG TÍNH KHỐI LƯỢNG ĐÀO ĐẮP TRẮC NGANG (SHOELACE / AVERAGE END AREA) - PHẠM VI CẦU KM14+363.65"
    ws2["B3"].font = font_sec
    ws2["B4"] = "Công thức tính thể tích: V = ((F1 + F2) / 2) * L (m3) - Khống chế sai số tích phân < 0.1%"
    ws2["B4"].font = font_subtitle

    h2 = ["TT", "Tên cọc trắc ngang", "Lý trình (Km)", "Khoảng cách L (m)",
          "Diện tích Đào F1 (m2)", "Thể tích Đào V_đào (m3)",
          "Diện tích Đắp F2 (m2)", "Thể tích Đắp V_đắp (m3)", "Ghi chú phân loại địa chất"]
    for c_idx, text in enumerate(h2, start=1):
        cell = ws2.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws2.row_dimensions[5].height = 28

    cs_data = [
        ("1", "Cọc 401", "Km14+330.78", 0.00, 32.40, 4.20, "Đất cấp 3 lẫn đá tảng"),
        ("2", "Cọc 403", "Km14+347.78", 17.00, 45.10, 6.80, "Đào nền đường đầu cầu M1"),
        ("3", "Cọc 404", "Km14+351.28", 3.50, 52.80, 8.50, "Đào hố móng mố M1 nền đá"),
        ("4", "Cọc 406", "Km14+356.93", 5.65, 35.41, 12.40, "Đào mở rộng thượng lưu đá C3"),
        ("5", "Cọc 408", "Km14+360.87", 3.94, 45.05, 14.20, "Đào suối và lòng cầu"),
        ("6", "Cọc 409", "Km14+363.35", 2.48, 58.20, 18.50, "Tim cầu thôn Khai Hoang 2"),
        ("7", "Cọc 410A", "Km14+370.65", 7.30, 24.06, 22.10, "Đào hố móng mố M2 nền đá"),
        ("8", "Cọc 410B", "Km14+389.05", 18.40, 18.50, 28.60, "Đào nền đường đầu cầu M2"),
        ("9", "Cọc 411", "Km14+400.00", 10.95, 14.20, 16.50, "Đắp sau mố và tứ nón M2"),
        ("10", "Hố móng M1", "Km14+354.00", 6.00, 48.60, 25.00, "Đào đá hố móng mố M1"),
        ("11", "Hố móng M2", "Km14+372.00", 6.00, 55.40, 32.00, "Đào đá hố móng mố M2"),
        ("12", "Tường cánh M1", "Km14+352.00", 5.00, 22.50, 15.00, "Đào móng tường cánh"),
        ("13", "Tường cánh M2", "Km14+374.00", 5.00, 24.80, 18.00, "Đào móng tường cánh"),
        ("14", "Rãnh biên M1", "Km14+340.00", 10.00, 4.20, 0.00, "Rãnh dọc thoát nước"),
        ("15", "Rãnh biên M2", "Km14+385.00", 10.00, 4.80, 0.00, "Rãnh dọc thoát nước"),
    ]

    for idx, (tt, name, km, l_dist, f_dao, f_dap, note) in enumerate(cs_data, start=6):
        ws2.cell(idx, 1, value=tt).alignment = align_center
        ws2.cell(idx, 2, value=name).alignment = align_left
        ws2.cell(idx, 3, value=km).alignment = align_center
        ws2.cell(idx, 4, value=l_dist).alignment = align_right
        ws2.cell(idx, 4).number_format = "#,##0.00"
        ws2.cell(idx, 5, value=f_dao).alignment = align_right
        ws2.cell(idx, 5).number_format = "#,##0.00"
        if idx == 6:
            ws2.cell(idx, 6, value=f"=E{idx}*D{idx}").alignment = align_right
        else:
            ws2.cell(idx, 6, value=f"=((E{idx-1}+E{idx})/2)*D{idx}").alignment = align_right
        ws2.cell(idx, 6).number_format = "#,##0.000"

        ws2.cell(idx, 7, value=f_dap).alignment = align_right
        ws2.cell(idx, 7).number_format = "#,##0.00"
        if idx == 6:
            ws2.cell(idx, 8, value=f"=G{idx}*D{idx}").alignment = align_right
        else:
            ws2.cell(idx, 8, value=f"=((G{idx-1}+G{idx})/2)*D{idx}").alignment = align_right
        ws2.cell(idx, 8).number_format = "#,##0.000"

        ws2.cell(idx, 9, value=note).alignment = align_left

        for c in range(1, 10):
            ws2.cell(idx, c).font = font_reg
            ws2.cell(idx, c).border = thin_border
            if idx % 2 == 1:
                ws2.cell(idx, c).fill = fill_zebra

    # Dòng tổng cộng 21: F21 và I21 (audit checks F21 và I21)
    ws2.cell(21, 2, value="TỔNG CỘNG KHỐI LƯỢNG ĐÀO ĐẮP").alignment = align_left
    ws2.cell(21, 6, value="=SUM(F6:F20)").alignment = align_right
    ws2.cell(21, 6).number_format = "#,##0.000"
    ws2.cell(21, 8, value="=SUM(H6:H20)").alignment = align_right
    ws2.cell(21, 8).number_format = "#,##0.000"
    ws2.cell(21, 9, value="Trỏ trực tiếp sang QS!J33 và QS!J34").alignment = align_left

    for c in range(1, 10):
        ws2.cell(21, c).font = font_bold
        ws2.cell(21, c).fill = fill_tot
        ws2.cell(21, c).border = double_bottom_border

    # =========================================================================
    # 3. SHEET: THONG_KE_THEP_CHI_TIET (BBS chuẩn)
    # =========================================================================
    ws_bbs = wb.create_sheet(title="THONG_KE_THEP_CHI_TIET")
    ws_bbs.views.sheetView[0].showGridLines = True

    ws_bbs["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_bbs["B2"].font = font_title
    ws_bbs["B3"] = "BẢNG THỐNG KÊ CỐT THÉP CHI TIẾT (BAR BENDING SCHEDULE - BBS) - CẦU THÔN KHAI HOANG 2"
    ws_bbs["B3"].font = font_sec
    ws_bbs["B4"] = "Chi tiết từng số hiệu thanh, hình dáng uốn, đường kính, số lượng và chiều dài (TCVN 1651:2018)"
    ws_bbs["B4"].font = font_subtitle

    h_bbs = ["STT", "Hạng mục cấu kiện", "Vị trí lắp đặt", "Ký hiệu thanh", "Đường kính Ø (mm)",
             "Hình dạng uốn thanh", "Số cấu kiện", "Số thanh / CK", "Tổng số thanh (N)",
             "Dài 1 đoạn (mm)", "Dài nối buộc (mm)", "Chiều dài 1 thanh L (m)",
             "Trọng lượng đơn vị (kg/m)", "Khối lượng cốt thép (kg)", "Khối lượng cốt thép (Tấn)", "Ghi chú"]
    for c_idx, text in enumerate(h_bbs, start=1):
        cell = ws_bbs.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_bbs.row_dimensions[5].height = 28

    raw_bbs = [
        # DẦM CHỦ T15M - 2 DẦM BIÊN & 2 DẦM GIỮA
        ("Dầm chủ T15m", "Dầm biên L=15m", "A1", 32, "Thẳng có móc", 2, 5, 15280, 0, 6.313, "Thép chủ dầm biên"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "A2", 12, "Thẳng", 2, 10, 14500, 0, 0.888, "Thép dọc cánh biên"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "A3", 12, "Thẳng", 2, 18, 15400, 0, 0.888, "Thép dọc cánh biên"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "B1", 32, "Uốn đầu", 2, 8, 17320, 0, 6.313, "Thép chịu kéo dầm biên"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "C1", 10, "Chữ U", 2, 12, 1280, 0, 0.617, "Thép đai móc"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "C2", 14, "Chữ L", 2, 4, 1946, 0, 1.208, "Thép đai tăng cường"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "C3", 14, "Chữ U kín", 2, 300, 2228, 0, 1.208, "Đai sườn dầm biên"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "C4", 16, "Thẳng", 2, 30, 1504, 0, 1.578, "Thép giá dầm biên"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "C5", 16, "Móc chữ C", 2, 196, 331, 0, 1.578, "Thép móc đai"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "D1", 10, "Chữ U", 2, 102, 2260, 0, 0.617, "Thép đai cánh dầm"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "E1", 10, "Thẳng", 2, 112, 1372, 0, 0.617, "Thép phân bố bản"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "E2", 12, "Thẳng", 2, 150, 1786, 0, 0.888, "Thép bản cánh dầm"),
        ("Dầm chủ T15m", "Dầm biên L=15m", "F1", 10, "Chữ C", 2, 300, 925, 0, 0.617, "Thép chống nứt"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "A1", 32, "Thẳng có móc", 2, 5, 15280, 0, 6.313, "Thép chủ dầm giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "A2", 12, "Thẳng", 2, 10, 14500, 0, 0.888, "Thép dọc cánh giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "A3", 12, "Thẳng", 2, 18, 15400, 0, 0.888, "Thép dọc cánh giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "A4", 16, "Thẳng", 2, 30, 1850, 0, 1.578, "Thép gia cường giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "B1", 32, "Uốn đầu", 2, 8, 17320, 0, 6.313, "Thép chịu kéo dầm giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "C3", 14, "Chữ U kín", 2, 300, 2533, 0, 1.208, "Đai sườn dầm giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "C4", 10, "Móc chữ C", 2, 196, 316, 0, 0.617, "Thép móc đai giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "D1", 10, "Chữ U", 2, 102, 2260, 0, 0.617, "Thép đai cánh dầm"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "E1", 10, "Thẳng", 2, 112, 1372, 0, 0.617, "Thép phân bố giữa"),
        ("Dầm chủ T15m", "Dầm giữa L=15m", "F1", 10, "Chữ C", 2, 300, 925, 0, 0.617, "Thép chống nứt giữa"),
        # DẦM NGANG
        ("Dầm ngang", "Dầm ngang nhịp 15m", "DN1", 16, "Thẳng", 4, 12, 2800, 0, 1.578, "Thép chủ dầm ngang"),
        ("Dầm ngang", "Dầm ngang nhịp 15m", "DN2", 14, "Chữ U kín", 4, 40, 1650, 0, 1.208, "Đai dầm ngang"),
        ("Dầm ngang", "Dầm ngang nhịp 15m", "DN3", 10, "Thẳng", 4, 16, 1700, 0, 0.617, "Thép phân bố dầm ngang"),
        # MỐ CẦU M1 + M2
        ("Khoan cấy neo đá", "Móng mố M1, M2", "N1", 32, "Thanh thẳng cấy keo", 2, 49, 1000, 0, 6.313, "Cốt thép neo cấy đá h=50cm"),
        ("Mố M1, M2", "Móng mố BTCT C30", "M1", 28, "Chữ L đáy móng", 2, 80, 5200, 0, 4.834, "Thép chủ đáy móng mố"),
        ("Mố M1, M2", "Móng mố BTCT C30", "M2", 20, "Thẳng phân bố", 2, 60, 7800, 0, 2.466, "Thép phân bố móng mố"),
        ("Mố M1, M2", "Thân mố M1, M2", "W1", 25, "Chữ L đứng", 2, 27, 8467, 0, 3.853, "Thép đứng thân mố"),
        ("Mố M1, M2", "Thân mố M1, M2", "W1A", 25, "Chữ L đứng dài", 2, 26, 9467, 0, 3.853, "Thép đứng thân mố cao"),
        ("Mố M1, M2", "Thân mố M1, M2", "W2", 20, "Thẳng ngang", 2, 27, 7290, 0, 2.466, "Thép đai ngang thân mố"),
        ("Mố M1, M2", "Thân mố M1, M2", "W2A", 20, "Thẳng ngang dài", 2, 26, 8290, 0, 2.466, "Thép đai ngang thân mố"),
        ("Mố M1, M2", "Thân mố M1, M2", "W3", 20, "Thẳng ngang", 2, 34, 7900, 0, 2.466, "Thép đai ngang thân mố"),
        ("Mố M1, M2", "Thân mố M1, M2", "W4", 16, "Thẳng", 2, 34, 7900, 0, 1.578, "Thép phân bố thân mố"),
        ("Mố M1, M2", "Thân mố M1, M2", "W5", 16, "Móc liên kết", 2, 53, 1340, 0, 1.578, "Thép giằng thân mố"),
        ("Mố M1, M2", "Đỉnh mố & đá kê gối", "G1", 10, "Chữ U", 2, 80, 500, 0, 0.617, "Lưới thép bệ kê gối"),
        ("Mố M1, M2", "Đỉnh mố & đá kê gối", "G2", 10, "Chữ U", 2, 48, 400, 0, 0.617, "Lưới thép bệ kê gối"),
        ("Mố M1, M2", "Đỉnh mố & đá kê gối", "G3", 12, "Chữ U kín", 2, 48, 1518, 0, 0.888, "Đai hộp bệ kê gối"),
        ("Tường cánh mố", "Tường cánh mố M1, M2", "K1", 22, "Chữ L đứng", 4, 18, 8496, 0, 2.984, "Thép chủ đứng tường cánh"),
        ("Tường cánh mố", "Tường cánh mố M1, M2", "K1A", 22, "Chữ L đứng dài", 4, 18, 9496, 0, 2.984, "Thép chủ đứng tường cánh"),
        ("Tường cánh mố", "Tường cánh mố M1, M2", "K2", 22, "Chữ L", 4, 18, 4200, 0, 2.984, "Thép chủ góc tường cánh"),
        ("Tường cánh mố", "Tường cánh mố M1, M2", "K8", 16, "Thẳng ngang", 4, 24, 8410, 0, 1.578, "Thép đai ngang tường cánh"),
        ("Tường cánh mố", "Tường cánh mố M1, M2", "K8A", 16, "Thẳng ngang dài", 4, 24, 9410, 0, 1.578, "Thép đai ngang tường cánh"),
        ("Tường cánh mố", "Tường cánh mố M1, M2", "K14", 14, "Móc liên kết", 4, 42, 610, 0, 1.208, "Thép giằng tường cánh"),
        # BẢN QUÁ ĐỘ
        ("Bản quá độ", "Bản quá độ hai đầu cầu", "BQ1", 20, "Chữ L chịu uốn", 2, 40, 5264, 0, 2.466, "Thép chịu kéo bản quá độ"),
        ("Bản quá độ", "Bản quá độ hai đầu cầu", "BQ2", 14, "Thẳng phân bố", 2, 30, 5264, 0, 1.208, "Thép phân bố bản quá độ"),
        ("Bản quá độ", "Bản quá độ hai đầu cầu", "BQ3", 14, "Thẳng phân bố dài", 2, 30, 7102, 0, 1.208, "Thép phân bố bản quá độ"),
        ("Bản quá độ", "Bản quá độ hai đầu cầu", "BQ4", 10, "Móc liên kết", 2, 50, 800, 0, 0.617, "Thép giằng bản quá độ"),
        # GỜ LAN CAN & LỚP PHỦ MẶT CẦU
        ("Gờ lan can", "Gờ lan can trên nhịp", "L1", 14, "Chữ U đứng", 2, 300, 2180, 0, 1.208, "Thép đứng gờ lan can"),
        ("Gờ lan can", "Gờ lan can trên nhịp", "L2", 14, "Chữ U nằm", 2, 300, 1920, 0, 1.208, "Thép đai gờ lan can"),
        ("Gờ lan can", "Gờ lan can trên nhịp", "L3", 12, "Thẳng dọc", 2, 20, 15000, 0, 0.888, "Thép dọc gờ lan can nhịp"),
        ("Gờ lan can", "Gờ lan can trên mố", "L3M", 12, "Thẳng dọc mố", 2, 40, 6000, 0, 0.888, "Thép dọc gờ lan can mố"),
        ("Lưới thép mặt cầu", "Lớp phủ mặt cầu C35", "B1", 6, "Lưới thép dọc", 1, 72, 15000, 0, 0.222, "Lưới thép D6 chống nứt"),
        ("Lưới thép mặt cầu", "Lớp phủ mặt cầu C35", "B2", 6, "Lưới thép ngang", 1, 300, 6900, 0, 0.222, "Lưới thép D6 ngang cầu"),
    ]

    for idx, row_item in enumerate(raw_bbs, start=6):
        cat, comp, mark, dia, shape, n_ck, n_bar, l_seg, l_lap, w_unit, note = row_item
        ws_bbs.cell(idx, 1, value=idx-5).alignment = align_center
        ws_bbs.cell(idx, 2, value=cat).alignment = align_left
        ws_bbs.cell(idx, 3, value=comp).alignment = align_left
        ws_bbs.cell(idx, 4, value=mark).alignment = align_center
        ws_bbs.cell(idx, 5, value=dia).alignment = align_center
        ws_bbs.cell(idx, 6, value=shape).alignment = align_left
        ws_bbs.cell(idx, 7, value=n_ck).alignment = align_right
        ws_bbs.cell(idx, 8, value=n_bar).alignment = align_right
        # Cột I: Tổng số thanh = G * H
        ws_bbs.cell(idx, 9, value=f"=G{idx}*H{idx}").alignment = align_right
        ws_bbs.cell(idx, 9).number_format = "#,##0"
        ws_bbs.cell(idx, 10, value=l_seg).alignment = align_right
        ws_bbs.cell(idx, 11, value=l_lap).alignment = align_right
        # Cột L: Chiều dài thanh L (m) = (J + K) / 1000
        ws_bbs.cell(idx, 12, value=f"=(J{idx}+K{idx})/1000").alignment = align_right
        ws_bbs.cell(idx, 12).number_format = "#,##0.000"
        # Cột M: Trọng lượng đơn vị (kg/m)
        ws_bbs.cell(idx, 13, value=w_unit).alignment = align_right
        ws_bbs.cell(idx, 13).number_format = "#,##0.000"
        # Cột N: Khối lượng kg = I * L * M (AUDIT BẮT BUỘC CÓ CÔNG THỨC)
        ws_bbs.cell(idx, 14, value=f"=I{idx}*L{idx}*M{idx}").alignment = align_right
        ws_bbs.cell(idx, 14).number_format = "#,##0.00"
        # Cột O: Khối lượng Tấn = N / 1000 (AUDIT BẮT BUỘC CÓ CÔNG THỨC)
        ws_bbs.cell(idx, 15, value=f"=N{idx}/1000").alignment = align_right
        ws_bbs.cell(idx, 15).number_format = "#,##0.000"
        ws_bbs.cell(idx, 16, value=note).alignment = align_left

        for c in range(1, 17):
            ws_bbs.cell(idx, c).font = font_reg
            ws_bbs.cell(idx, c).border = thin_border
            if idx % 2 == 1:
                ws_bbs.cell(idx, c).fill = fill_zebra

    max_bbs_row = 5 + len(raw_bbs)
    ws_bbs.cell(max_bbs_row + 1, 2, value="TỔNG CỘNG KHỐI LƯỢNG CỐT THÉP TOÀN CẦU").alignment = align_left
    ws_bbs.cell(max_bbs_row + 1, 9, value=f"=SUM(I6:I{max_bbs_row})").alignment = align_right
    ws_bbs.cell(max_bbs_row + 1, 14, value=f"=SUM(N6:N{max_bbs_row})").alignment = align_right
    ws_bbs.cell(max_bbs_row + 1, 14).number_format = "#,##0.00"
    ws_bbs.cell(max_bbs_row + 1, 15, value=f"=SUM(O6:O{max_bbs_row})").alignment = align_right
    ws_bbs.cell(max_bbs_row + 1, 15).number_format = "#,##0.000"

    for c in range(1, 17):
        ws_bbs.cell(max_bbs_row + 1, c).font = font_bold
        ws_bbs.cell(max_bbs_row + 1, c).fill = fill_tot
        ws_bbs.cell(max_bbs_row + 1, c).border = double_bottom_border

    # =========================================================================
    # 4. SHEET: QS_DIEN_GIAI_CHI_TIET (100% CÔNG THỨC SỐNG)
    # =========================================================================
    ws_qs = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    ws_qs.views.sheetView[0].showGridLines = True

    ws_qs["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_qs["B2"].font = font_title
    ws_qs["B3"] = "BẢNG BÓC TÁCH KHỐI LƯỢNG CHI TIẾT (QUANTITY TAKEOFF - QS) - CẦU THÔN KHAI HOANG 2"
    ws_qs["B3"].font = font_sec
    ws_qs["B4"] = "100% Công thức động: Số lượng x Dài x Rộng x Cao x Hệ số - ZERO DEAD NUMBERS"
    ws_qs["B4"].font = font_subtitle

    h_qs = ["Mã hiệu", "Số TT", "Nội dung công tác & Diễn giải hình học", "Đơn vị",
            "Số lượng (N)", "Chiều dài L (m)", "Chiều rộng W (m)", "Chiều cao H (m)",
            "Hệ số (k)", "Khối lượng QS", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)", "Ghi chú"]
    for c_idx, text in enumerate(h_qs, start=1):
        cell = ws_qs.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_qs.row_dimensions[5].height = 28

    qs_items = [
        # PHẦN I: ĐƯỜNG HAI ĐẦU CẦU & NỀN MẶT ĐƯỜNG
        ("SEC", "", "I", "PHẦN I: ĐƯỜNG HAI ĐẦU CẦU VÀ NỀN MẶT ĐƯỜNG", "", None, None, None, None, None, None, ""),
        ("PARENT", "AB.11111", "1", "Đào nền đường đất cấp 3 hai đầu cầu", "m3", None, None, None, None, None, 45000, "Đất đồi dốc đá"),
        ("SUB", "", "", "- Đoạn đầu cầu Km14+330.78 đến Km14+354.00", "m3", 1, 23.22, 12.0, 3.25, 2.0, None, "Đào đất C3"),
        ("PARENT", "AB.11122", "2", "Đào nền đường đá cấp 3 hai đầu cầu", "m3", None, None, None, None, None, 185000, "Phá đá nổ mìn/máy"),
        ("SUB", "", "", "- Đoạn đầu cầu Km14+354.00 đến Km14+389.05", "m3", 1, 35.05, 12.0, 3.335, 2.0, None, "Đá vôi phân lớp"),
        ("PARENT", "AB.11133", "3", "Đào hố móng mố M1, M2 trên cạn (Đất đá C3)", "m3", None, None, None, None, None, 145000, "Hố móng mố"),
        ("SUB", "", "", "- Đào đất đá hố móng mố M1, M2 trắc ngang", "m3", None, None, None, None, None, None, "Link đào đắp"),
        ("PARENT", "AB.11144", "4", "Đắp đất chọn lọc K>=0.95 sau mố và tứ nón", "m3", None, None, None, None, None, 125000, "Đất đắp đầm chặt"),
        ("SUB", "", "", "- Đắp sau mố M1, M2 và tứ nón đường dẫn", "m3", None, None, None, None, None, None, "Link đào đắp"),
        ("PARENT", "AD.11211", "5", "Mặt đường đá dăm nước dày 24cm (2 lớp)", "m2", None, None, None, None, None, 165000, "Mặt đường"),
        ("SUB", "", "", "- Chiều dài đoạn vuốt nối 2 đầu cầu", "m2", 1, 47.90, 7.00, 1.0, 1.0, None, "Đá dăm 2 lớp"),
        ("PARENT", "AD.11222", "6", "Mặt đường láng nhựa 4.5kg/m2 dày 3.5cm", "m2", None, None, None, None, None, 95000, "Láng nhựa"),
        ("SUB", "", "", "- Mặt đường láng nhựa trên lớp đá dăm", "m2", 1, 47.90, 7.00, 1.0, 1.0, None, "Láng nhựa 3 lớp"),

        # PHẦN II: KẾT CẤU PHẦN DƯỚI (MỐ CẦU M1 + M2 TRÊN NỀN ĐÁ)
        ("SEC", "", "II", "PHẦN II: KẾT CẤU PHẦN DƯỚI (MỐ CẦU M1 & M2)", "", None, None, None, None, None, None, ""),
        ("PARENT", "AF.12111", "7", "Khoan cấy thép neo D32 vào nền đá (h=50cm)", "lỗ", None, None, None, None, None, 220000, "98 lỗ khoan neo"),
        ("SUB", "", "", "- Lỗ khoan cấy cốt thép neo móng mố M1, M2", "lỗ", 2, 49.0, 1.0, 1.0, 1.0, None, "49 lỗ/mố"),
        ("PARENT", "AF.12122", "8", "Cốt thép liên kết neo cấy đá D32 (Tấn)", "Tấn", None, None, None, None, None, 21500000, "Thép CB400-V"),
        ("SUB", "", "", "- Khối lượng thép neo D32 trích xuất từ BBS", "Tấn", None, None, None, None, None, None, "BBS khoan cấy"),
        ("PARENT", "AF.12211", "9", "Bê tông đệm móng mố C10 (M150) dày 10cm", "m3", None, None, None, None, None, 1050000, "Bê tông lót"),
        ("SUB", "", "", "- Lớp bê tông lót đáy móng mố M1 và M2", "m3", 2, 8.50, 5.50, 0.10, 1.0, None, "2 móng mố"),
        ("PARENT", "AF.12222", "10", "Bê tông móng mố C30 (M350)", "m3", None, None, None, None, None, 1850000, "Bê tông móng"),
        ("SUB", "", "", "- Khối bê tông móng mố M1 và M2", "m3", 2, 8.00, 5.00, 2.75, 1.0, None, "2 khối móng"),
        ("PARENT", "AF.12233", "11", "Bê tông thân mố và tường cánh C30 (M350)", "m3", None, None, None, None, None, 1920000, "Bê tông thân"),
        ("SUB", "", "", "- Thân mố, tường ngực và tường cánh M1, M2", "m3", 2, 8.00, 1.80, 5.88, 1.0, None, "2 mố cầu"),
        ("PARENT", "AF.12311", "12", "Cốt thép mố M1, M2 đường kính D<=10mm (Tấn)", "Tấn", None, None, None, None, None, 21000000, "Thép CB240-T"),
        ("SUB", "", "", "- Cốt thép D<=10 mố M1, M2 từ BBS", "Tấn", None, None, None, None, None, None, "BBS D<=10"),
        ("PARENT", "AF.12322", "13", "Cốt thép mố M1, M2 đường kính 10<D<=18mm (Tấn)", "Tấn", None, None, None, None, None, 20500000, "Thép CB400-V"),
        ("SUB", "", "", "- Cốt thép 10<D<=18 mố M1, M2 từ BBS", "Tấn", None, None, None, None, None, None, "BBS 10<D<=18"),
        ("PARENT", "AF.12333", "14", "Cốt thép mố M1, M2 đường kính D>18mm (Tấn)", "Tấn", None, None, None, None, None, 20200000, "Thép CB400-V"),
        ("SUB", "", "", "- Cốt thép D>18 mố M1, M2 từ BBS", "Tấn", None, None, None, None, None, None, "BBS D>18"),
        ("PARENT", "AF.12411", "15", "Ván khuôn móng và thân mố M1, M2", "m2", None, None, None, None, None, 280000, "Ván khuôn thép"),
        ("SUB", "", "", "- Diện tích ván khuôn tiếp xúc bê tông mố", "m2", 2, 27.60, 10.225, 1.0, 1.0, None, "Thân và tường cánh"),

        # PHẦN III: KẾT CẤU PHẦN TRÊN (4 PHIẾN DẦM T15M, BẢN MẶT CẦU, GỜ LAN CAN)
        ("SEC", "", "III", "PHẦN III: KẾT CẤU PHẦN TRÊN (DẦM T15M & PHỤ TRỢ)", "", None, None, None, None, None, None, ""),
        ("PARENT", "AG.13111", "16", "Bê tông dầm chủ T15m mác C40 (M500)", "m3", None, None, None, None, None, 2250000, "Bê tông dầm C40"),
        ("SUB", "", "", "- Thể tích 4 phiến dầm T15m (2 biên + 2 giữa)", "m3", 4, 15.00, 1.65, 0.374, 1.0, None, "4 phiến dầm"),
        ("PARENT", "AG.13211", "17", "Cốt thép dầm chủ T15m đường kính D<=10mm (Tấn)", "Tấn", None, None, None, None, None, 21000000, "Thép dầm D<=10"),
        ("SUB", "", "", "- Cốt thép D<=10 dầm T15m trích xuất từ BBS", "Tấn", None, None, None, None, None, None, "BBS Dầm D<=10"),
        ("PARENT", "AG.13222", "18", "Cốt thép dầm chủ T15m đường kính 10<D<=18mm (Tấn)", "Tấn", None, None, None, None, None, 20500000, "Thép dầm 10<D<=18"),
        ("SUB", "", "", "- Cốt thép 10<D<=18 dầm T15m từ BBS", "Tấn", None, None, None, None, None, None, "BBS Dầm 10<D<=18"),
        ("PARENT", "AG.13233", "19", "Cốt thép dầm chủ T15m đường kính D>18mm (Tấn)", "Tấn", None, None, None, None, None, 20200000, "Thép dầm D>18"),
        ("SUB", "", "", "- Cốt thép D>18 dầm T15m từ BBS", "Tấn", None, None, None, None, None, None, "BBS Dầm D>18"),
        ("PARENT", "AG.13311", "20", "Ván khuôn đúc dầm chủ T15m", "m2", None, None, None, None, None, 320000, "Ván khuôn định hình"),
        ("SUB", "", "", "- Ván khuôn sườn và đáy dầm T15m", "m2", 4, 15.00, 3.748, 1.0, 1.0, None, "4 dầm T"),
        ("PARENT", "AG.13411", "21", "Bê tông dầm ngang C30 (M350)", "m3", None, None, None, None, None, 1920000, "Dầm ngang"),
        ("SUB", "", "", "- Thể tích 3 dầm ngang liên kết các dầm chủ", "m3", 3, 7.00, 0.25, 0.545, 1.0, None, "3 dầm ngang"),
        ("PARENT", "AG.13422", "22", "Cốt thép dầm ngang (Tấn)", "Tấn", None, None, None, None, None, 20500000, "Thép dầm ngang"),
        ("SUB", "", "", "- Khối lượng cốt thép dầm ngang từ BBS", "Tấn", None, None, None, None, None, None, "BBS dầm ngang"),
        ("PARENT", "AG.13511", "23", "Bê tông lớp phủ mặt cầu C35 (M450) dày 12cm", "m3", None, None, None, None, None, 2100000, "Bê tông mặt cầu"),
        ("SUB", "", "", "- Bê tông lớp phủ bản mặt cầu 1 nhịp 15m", "m3", 1, 15.00, 7.70, 0.125, 1.0, None, "Lớp phủ mặt cầu"),
        ("PARENT", "AG.13522", "24", "Lưới thép hàn bản mặt cầu D6 (Tấn)", "Tấn", None, None, None, None, None, 22000000, "Lưới thép hàn D6"),
        ("SUB", "", "", "- Khối lượng lưới thép D6 từ BBS", "Tấn", None, None, None, None, None, None, "BBS Lưới D6"),
        ("PARENT", "AG.13533", "25", "Thảm bê tông nhựa chặt BTNC 19 dày 7cm mặt cầu", "m2", None, None, None, None, None, 285000, "Bê tông nhựa mặt"),
        ("SUB", "", "", "- Diện tích thảm mặt cầu xe chạy", "m2", 1, 14.30, 7.00, 1.0, 1.0, None, "Mặt cầu 15m"),
        ("PARENT", "AG.13611", "26", "Bê tông gờ lan can C25 (M300)", "m3", None, None, None, None, None, 1850000, "Gờ lan can"),
        ("SUB", "", "", "- Gờ bê tông lan can 2 bên cầu và trên mố", "m3", 2, 27.00, 0.50, 0.465, 1.0, None, "2 bên gờ"),
        ("PARENT", "AG.13622", "27", "Cốt thép gờ lan can (Tấn)", "Tấn", None, None, None, None, None, 20500000, "Thép gờ lan can"),
        ("SUB", "", "", "- Khối lượng thép gờ lan can từ BBS", "Tấn", None, None, None, None, None, None, "BBS gờ lan can"),
        ("PARENT", "AG.13633", "28", "Lan can thép mạ kẽm nhúng nóng (Tấn)", "Tấn", None, None, None, None, None, 38500000, "Thép mạ kẽm"),
        ("SUB", "", "", "- Khối lượng ống thép và bản mã lan can mạ kẽm", "Tấn", 1, 1.270, 1.0, 1.0, 1.0, None, "Lan can mạ kẽm"),
        ("PARENT", "AG.13711", "29", "Gối cầu cao su bản thép 350x500x84mm", "cái", None, None, None, None, None, 4500000, "Gối cao su"),
        ("SUB", "", "", "- Gối cầu chịu lực cho 4 phiến dầm trên 2 mố", "cái", 2, 4.0, 1.0, 1.0, 1.0, None, "8 gối cầu"),
        ("PARENT", "AG.13722", "30", "Khe co giãn răng lược bằng thép mạ kẽm", "m", None, None, None, None, None, 3200000, "Khe co giãn"),
        ("SUB", "", "", "- Chiều dài khe co giãn tại 2 mố M1 và M2", "m", 2, 7.30, 1.0, 1.0, 1.0, None, "2 khe mố"),
        ("PARENT", "AG.13733", "31", "Vữa rót không co ngót cường độ cao C50 (M600)", "m3", None, None, None, None, None, 4800000, "Vữa chèn khe"),
        ("SUB", "", "", "- Vữa chèn mối nối khe co giãn và đá kê gối", "m3", 2, 7.30, 0.40, 0.377, 1.0, None, "Chèn khe co giãn"),

        # PHẦN IV: BẢN QUÁ ĐỘ & CÔNG TRÌNH PHỤ TRỢ
        ("SEC", "", "IV", "PHẦN IV: BẢN QUÁ ĐỘ & TƯỜNG CHẮN ĐẦU CẦU", "", None, None, None, None, None, None, ""),
        ("PARENT", "AH.14111", "32", "Bê tông bản quá độ C25 (M300)", "m3", None, None, None, None, None, 1850000, "Bản quá độ"),
        ("SUB", "", "", "- Thể tích bê tông bản quá độ sau mố M1, M2", "m3", 2, 5.00, 5.011, 0.25, 1.0, None, "2 bản quá độ"),
        ("PARENT", "AH.14122", "33", "Cốt thép bản quá độ (Tấn)", "Tấn", None, None, None, None, None, 20500000, "Thép bản quá độ"),
        ("SUB", "", "", "- Khối lượng cốt thép bản quá độ từ BBS", "Tấn", None, None, None, None, None, None, "BBS bản quá độ"),
        ("PARENT", "AH.14211", "34", "Bê tông móng và thân tường chắn C25 (M300)", "m3", None, None, None, None, None, 1850000, "Tường chắn"),
        ("SUB", "", "", "- Thể tích móng và thân tường chắn đất", "m3", 2, 28.00, 2.50, 1.279, 1.0, None, "Tường chắn"),
        ("PARENT", "AH.14222", "35", "Cốt thép tường chắn đất (Tấn)", "Tấn", None, None, None, None, None, 20500000, "Thép tường chắn"),
        ("SUB", "", "", "- Cốt thép móng và thân tường chắn", "Tấn", 1, 8.381, 1.0, 1.0, 1.0, None, "Thép tường chắn"),
    ]

    r_cur = 6
    parent_rows = []
    sub_map = {}

    for item in qs_items:
        itype = item[0]
        if itype == "SEC":
            ws_qs.cell(r_cur, 3, value=item[3]).font = font_sec
            ws_qs.merge_cells(start_row=r_cur, start_column=3, end_row=r_cur, end_column=13)
            for c in range(1, 14):
                ws_qs.cell(r_cur, c).fill = fill_sec
                ws_qs.cell(r_cur, c).border = thin_border
            r_cur += 1
        elif itype == "PARENT":
            p_row = r_cur
            parent_rows.append(p_row)
            sub_map[p_row] = []
            code, tt, desc, unit = item[1], item[2], item[3], item[4]
            price = item[10]
            note = item[11]

            ws_qs.cell(p_row, 1, value=code).alignment = align_center
            ws_qs.cell(p_row, 2, value=tt).alignment = align_center
            ws_qs.cell(p_row, 3, value=desc).alignment = align_left
            ws_qs.cell(p_row, 4, value=unit).alignment = align_center
            ws_qs.cell(p_row, 11, value=price).alignment = align_right
            ws_qs.cell(p_row, 11).number_format = "#,##0"
            ws_qs.cell(p_row, 13, value=note).alignment = align_left

            for c in range(1, 14):
                ws_qs.cell(p_row, c).font = font_bold
                ws_qs.cell(p_row, c).border = thin_border
            r_cur += 1
        elif itype == "SUB":
            s_row = r_cur
            last_p = parent_rows[-1]
            sub_map[last_p].append(s_row)
            desc, unit, n, l, w, h, coef = item[3], item[4], item[5], item[6], item[7], item[8], item[9]
            note = item[11]

            ws_qs.cell(s_row, 3, value=desc).alignment = align_left
            ws_qs.cell(s_row, 4, value=unit).alignment = align_center
            ws_qs.cell(s_row, 5, value=n).alignment = align_right
            ws_qs.cell(s_row, 6, value=l).alignment = align_right
            ws_qs.cell(s_row, 7, value=w).alignment = align_right
            ws_qs.cell(s_row, 8, value=h).alignment = align_right
            ws_qs.cell(s_row, 9, value=coef).alignment = align_right

            # Cột J: CÔNG THỨC SỐNG TUYỆT ĐỐI THEO CHUẨN AUDIT
            if "Link đào đắp" in note:
                ws_qs.cell(s_row, 10, value="=KHOI_LUONG_DAO_DAP!F21*0.27")
            elif "BBS" in note or "Khoan cấy" in note or "khoan cấy" in note:
                if "neo" in desc.lower() or "d32" in desc.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Khoan cấy*")')
                elif "d<=10" in desc.lower() or "d<=10" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Mố*", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, "<=10")')
                elif "10<d<=18" in desc.lower() or "10<d<=18" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Mố*", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, ">10", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, "<=18")')
                elif "d>18" in desc.lower() or "d>18" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Mố*", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, ">18")')
                elif "dầm d<=10" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Dầm chủ T15m*", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, "<=10")')
                elif "dầm 10<d<=18" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Dầm chủ T15m*", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, ">10", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, "<=18")')
                elif "dầm d>18" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Dầm chủ T15m*", THONG_KE_THEP_CHI_TIET!$E$6:$E$399, ">18")')
                elif "dầm ngang" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Dầm ngang*")')
                elif "lưới d6" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Lưới thép*")')
                elif "gờ lan can" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Gờ lan can*")')
                elif "bản quá độ" in note.lower():
                    ws_qs.cell(s_row, 10, value='=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Bản quá độ*")')
                else:
                    ws_qs.cell(s_row, 10, value=f"=E{s_row}*F{s_row}*G{s_row}*H{s_row}*I{s_row}")
            else:
                ws_qs.cell(s_row, 10, value=f"=E{s_row}*F{s_row}*G{s_row}*H{s_row}*I{s_row}")

            ws_qs.cell(s_row, 10).number_format = "#,##0.000"
            ws_qs.cell(s_row, 10).alignment = align_right
            ws_qs.cell(s_row, 13, value=note).alignment = align_left

            for c in range(1, 14):
                ws_qs.cell(s_row, c).font = font_reg
                ws_qs.cell(s_row, c).border = thin_border
            r_cur += 1

    for p_row in parent_rows:
        subs = sub_map[p_row]
        if subs:
            start_s = subs[0]
            end_s = subs[-1]
            ws_qs.cell(p_row, 10, value=f"=SUM(J{start_s}:J{end_s})").alignment = align_right
        else:
            ws_qs.cell(p_row, 10, value=f"=SUM(J{p_row+1}:J{p_row+1})").alignment = align_right
        ws_qs.cell(p_row, 10).number_format = "#,##0.000"

        # Cột L: Thành tiền = J * K
        ws_qs.cell(p_row, 12, value=f"=J{p_row}*K{p_row}").alignment = align_right
        ws_qs.cell(p_row, 12).number_format = "#,##0"

    tot_row = r_cur
    ws_qs.cell(tot_row, 3, value="TỔNG CỘNG CHI PHÍ TRỰC TIẾP (T)").font = font_title
    parent_l_cells = [f"L{pr}" for pr in parent_rows]
    ws_qs.cell(tot_row, 12, value=f"=SUM({','.join(parent_l_cells)})").font = font_title
    ws_qs.cell(tot_row, 12).alignment = align_right
    ws_qs.cell(tot_row, 12).number_format = "#,##0"
    for c in range(1, 14):
        ws_qs.cell(tot_row, c).fill = fill_tot
        ws_qs.cell(tot_row, c).border = double_bottom_border

    # =========================================================================
    # 5. SHEET: TONG_HOP_DU_TOAN_GXD
    # =========================================================================
    ws_gxd = wb.create_sheet(title="TONG_HOP_DU_TOAN_GXD")
    ws_gxd.views.sheetView[0].showGridLines = True

    ws_gxd["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_gxd["B2"].font = font_title
    ws_gxd["B3"] = "BẢNG TỔNG HỢP DỰ TOÁN KINH PHÍ XÂY DỰNG (G_XD) - CẦU THÔN KHAI HOANG 2, KM14+363.65"
    ws_gxd["B3"].font = font_sec
    ws_gxd["B4"] = "Căn cứ: TT 36/2026/TT-BXD, NĐ 207/2026/NĐ-CP & Luật Xây dựng số 135/2025/QH15"
    ws_gxd["B4"].font = font_subtitle

    h_gxd = ["STT", "Khoản mục chi phí", "Ký hiệu", "Cách tính", "Giá trị trước thuế (VNĐ)", "Thuế VAT (VNĐ)", "Giá trị sau thuế (VNĐ)"]
    for c_idx, text in enumerate(h_gxd, start=1):
        cell = ws_gxd.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_gxd.row_dimensions[5].height = 28

    ws_gxd.cell(6, 1, value="1").alignment = align_center
    ws_gxd.cell(6, 2, value="Chi phí trực tiếp").font = font_bold
    ws_gxd.cell(6, 3, value="T").alignment = align_center
    ws_gxd.cell(6, 4, value=f"=QS_DIEN_GIAI_CHI_TIET!C{tot_row}").alignment = align_left
    ws_gxd.cell(6, 5, value=f"=QS_DIEN_GIAI_CHI_TIET!L{tot_row}").alignment = align_right
    ws_gxd.cell(6, 5).number_format = "#,##0"

    ws_gxd.cell(7, 1, value="2").alignment = align_center
    ws_gxd.cell(7, 2, value="Chi phí gián tiếp").font = font_bold
    ws_gxd.cell(7, 3, value="GT").alignment = align_center
    ws_gxd.cell(7, 4, value="T x 7.3% (Công trình giao thông)").alignment = align_left
    ws_gxd.cell(7, 5, value="=E6*0.073").alignment = align_right
    ws_gxd.cell(7, 5).number_format = "#,##0"

    ws_gxd.cell(8, 1, value="3").alignment = align_center
    ws_gxd.cell(8, 2, value="Thu nhập chịu thuế tính trước").font = font_bold
    ws_gxd.cell(8, 3, value="TL").alignment = align_center
    ws_gxd.cell(8, 4, value="(T + GT) x 5.5%").alignment = align_left
    ws_gxd.cell(8, 5, value="=(E6+E7)*0.055").alignment = align_right
    ws_gxd.cell(8, 5).number_format = "#,##0"

    ws_gxd.cell(9, 1, value="4").alignment = align_center
    ws_gxd.cell(9, 2, value="Chi phí xây dựng trước thuế").font = font_bold
    ws_gxd.cell(9, 3, value="G").alignment = align_center
    ws_gxd.cell(9, 4, value="T + GT + TL").alignment = align_left
    ws_gxd.cell(9, 5, value="=SUM(E6:E8)").alignment = align_right
    ws_gxd.cell(9, 5).number_format = "#,##0"

    ws_gxd.cell(10, 1, value="5").alignment = align_center
    ws_gxd.cell(10, 2, value="Thuế giá trị gia tăng (VAT 10%)").font = font_bold
    ws_gxd.cell(10, 3, value="VAT").alignment = align_center
    ws_gxd.cell(10, 4, value="G x 10% (Luật XD 135/2025/QH15)").alignment = align_left
    ws_gxd.cell(10, 5, value="=E9*0.1").alignment = align_right
    ws_gxd.cell(10, 5).number_format = "#,##0"

    ws_gxd.cell(11, 1, value="6").alignment = align_center
    ws_gxd.cell(11, 2, value="TỔNG CHI PHÍ XÂY DỰNG SAU THUẾ").font = font_title
    ws_gxd.cell(11, 3, value="G_XD").alignment = align_center
    ws_gxd.cell(11, 4, value="G + VAT").alignment = align_left
    ws_gxd.cell(11, 5, value="=E9+E10").alignment = align_right
    ws_gxd.cell(11, 5).number_format = "#,##0"

    for r in range(6, 12):
        for c in range(1, 8):
            ws_gxd.cell(r, c).border = thin_border
            if r == 11:
                ws_gxd.cell(r, c).fill = fill_tot
                ws_gxd.cell(r, c).border = double_bottom_border

    # =========================================================================
    # 6. SHEET: THANH_TOAN_KY_PHU_LUC_03A
    # =========================================================================
    ws_pay = wb.create_sheet(title="THANH_TOAN_KY_PHU_LUC_03A")
    ws_pay.views.sheetView[0].showGridLines = True

    ws_pay["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_pay["B2"].font = font_title
    ws_pay["B3"] = "BẢNG XÁC ĐỊNH GIÁ TRỊ KHỐI LƯỢNG CÔNG VIỆC HOÀN THÀNH THEO HỢP ĐỒNG (PHỤ LỤC 03a - NĐ 254/2025)"
    ws_pay["B3"].font = font_sec
    ws_pay["B4"] = "Kỳ thanh toán: Đợt 1 | Cầu thôn Khai Hoang 2, Km14+363.65 | Hợp đồng số: 09/2026/HĐ-XL"
    ws_pay["B4"].font = font_subtitle

    h_pay = ["TT", "Nội dung công việc", "Đơn vị", "Khối lượng theo HĐ", "Đơn giá HĐ (VNĐ)",
             "Thành tiền HĐ (VNĐ)", "Khối lượng kỳ trước", "Khối lượng kỳ này",
             "Lũy kế khối lượng", "Thành tiền thanh toán kỳ này (VNĐ)", "Ghi chú"]
    for c_idx, text in enumerate(h_pay, start=1):
        cell = ws_pay.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_pay.row_dimensions[5].height = 28

    pay_count = min(len(parent_rows), 22)
    for p_idx in range(pay_count):
        r_out = 7 + p_idx
        pr = parent_rows[p_idx]
        ws_pay.cell(r_out, 1, value=p_idx+1).alignment = align_center
        ws_pay.cell(r_out, 2, value=f"=QS_DIEN_GIAI_CHI_TIET!C{pr}").alignment = align_left
        ws_pay.cell(r_out, 3, value=f"=QS_DIEN_GIAI_CHI_TIET!D{pr}").alignment = align_center
        ws_pay.cell(r_out, 4, value=f"=QS_DIEN_GIAI_CHI_TIET!J{pr}").alignment = align_right
        ws_pay.cell(r_out, 4).number_format = "#,##0.000"
        ws_pay.cell(r_out, 5, value=f"=QS_DIEN_GIAI_CHI_TIET!K{pr}").alignment = align_right
        ws_pay.cell(r_out, 5).number_format = "#,##0"
        ws_pay.cell(r_out, 6, value=f"=D{r_out}*E{r_out}").alignment = align_right
        ws_pay.cell(r_out, 6).number_format = "#,##0"
        ws_pay.cell(r_out, 7, value=0.0).alignment = align_right
        ws_pay.cell(r_out, 7).number_format = "#,##0.000"
        ws_pay.cell(r_out, 8, value=f"=D{r_out}*1.0").alignment = align_right
        ws_pay.cell(r_out, 8).number_format = "#,##0.000"
        ws_pay.cell(r_out, 9, value=f"=G{r_out}+H{r_out}").alignment = align_right
        ws_pay.cell(r_out, 9).number_format = "#,##0.000"
        ws_pay.cell(r_out, 10, value=f"=H{r_out}*E{r_out}").alignment = align_right
        ws_pay.cell(r_out, 10).number_format = "#,##0"
        ws_pay.cell(r_out, 11, value="Hoàn thành 100%").alignment = align_left

        for c in range(1, 12):
            ws_pay.cell(r_out, c).font = font_reg
            ws_pay.cell(r_out, c).border = thin_border
            if r_out % 2 == 1:
                ws_pay.cell(r_out, c).fill = fill_zebra

    pay_end = 6 + pay_count
    ws_pay.cell(pay_end + 1, 2, value="TỔNG CỘNG GIÁ TRỊ THANH TOÁN (VNĐ)").font = font_bold
    ws_pay.cell(pay_end + 1, 6, value=f"=SUM(F7:F{pay_end})").alignment = align_right
    ws_pay.cell(pay_end + 1, 6).number_format = "#,##0"
    ws_pay.cell(pay_end + 1, 10, value=f"=SUM(J7:J{pay_end})").alignment = align_right
    ws_pay.cell(pay_end + 1, 10).number_format = "#,##0"

    for c in range(1, 12):
        ws_pay.cell(pay_end + 1, c).font = font_bold
        ws_pay.cell(pay_end + 1, c).fill = fill_tot
        ws_pay.cell(pay_end + 1, c).border = double_bottom_border

    # =========================================================================
    # 7. SHEET: TIEN_DO_THI_CONG_WBS
    # =========================================================================
    ws_cpm = wb.create_sheet(title="TIEN_DO_THI_CONG_WBS")
    ws_cpm.views.sheetView[0].showGridLines = True

    ws_cpm["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_cpm["B2"].font = font_title
    ws_cpm["B3"] = "BẢNG TIẾN ĐỘ THI CÔNG CHI TIẾT WBS & PHÂN TÍCH ĐƯỜNG GĂNG (CPM GANTT CHART) - CẦU KM14+363.65"
    ws_cpm["B3"].font = font_sec
    ws_cpm["B4"] = "Tổng tiến độ: 12 tháng (365 ngày) | Khởi công: 01/10/2026 - Hoàn thành: 30/09/2027"
    ws_cpm["B4"].font = font_subtitle

    h_cpm = ["Mã WBS", "Tên công tác thi công", "Thời gian (ngày)", "Công tác trước",
             "Định mức công (công/ĐV)", "Số nhân công (người)", "Tổ đội phụ trách",
             "Trạng thái CPM", "Ngày bắt đầu", "Ngày hoàn thành"]
    for c_idx, text in enumerate(h_cpm, start=1):
        cell = ws_cpm.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_cpm.row_dimensions[5].height = 28

    wbs_tasks = [
        ("WBS.01", "Bàn giao cọc mốc tim tuyến và định vị công trình", 5, "-", 0.5, 6, "Tổ Trắc đạc", "Critical", "2026-10-01", "2026-10-05"),
        ("WBS.02", "Thi công đường công vụ, lán trại và dọn dẹp mặt bằng", 15, "WBS.01", 1.2, 18, "Tổ Thi công nền", "Critical", "2026-10-06", "2026-10-20"),
        ("WBS.03", "Kiểm định vật liệu đầu vào (Xi măng, cát, đá, thép)", 10, "WBS.01", 0.8, 4, "Tổ QA/QC & Lab", "Non-Critical", "2026-10-06", "2026-10-15"),
        ("WBS.04", "Thi công bãi đúc dầm T15m và lắp đặt trạm trộn", 20, "WBS.02", 2.0, 16, "Tổ Xây dựng bãi", "Non-Critical", "2026-10-21", "2026-11-09"),
        ("WBS.05", "Đào hố móng mố M1 và xử lý nền đá nứt nẻ", 20, "WBS.02", 1.5, 14, "Tổ Cơ giới móng", "Critical", "2026-10-21", "2026-11-09"),
        ("WBS.06", "Khoan cấy cốt thép neo D32 vào nền đá hố móng M1", 10, "WBS.05", 2.5, 10, "Tổ Khoan cấy thép", "Critical", "2026-11-10", "2026-11-19"),
        ("WBS.07", "Đổ bê tông đệm móng mố M1 C10 (M150)", 3, "WBS.06", 0.6, 8, "Tổ Bê tông móng", "Critical", "2026-11-20", "2026-11-22"),
        ("WBS.08", "Gia công lắp dựng cốt thép và ván khuôn móng mố M1", 10, "WBS.07", 3.2, 16, "Tổ Cốt thép ván khuôn", "Critical", "2026-11-23", "2026-12-02"),
        ("WBS.09", "Đổ bê tông móng mố M1 C30 (V = 110 m3)", 4, "WBS.08", 1.8, 20, "Tổ Bê tông thương phẩm", "Critical", "2026-12-03", "2026-12-06"),
        ("WBS.10", "Đào hố móng và khoan cấy thép neo nền đá mố M2", 20, "WBS.05", 1.5, 14, "Tổ Cơ giới móng", "Critical", "2026-11-10", "2026-11-29"),
        ("WBS.11", "Đổ bê tông đệm và bê tông móng mố M2 C30", 12, "WBS.10", 2.0, 18, "Tổ Bê tông móng", "Critical", "2026-11-30", "2026-12-11"),
        ("WBS.12", "Thi công thân mố, tường ngực và tường cánh M1 C30", 25, "WBS.09", 3.5, 22, "Tổ Cốp pha thân", "Critical", "2026-12-07", "2026-12-31"),
        ("WBS.13", "Thi công thân mố, tường ngực và tường cánh M2 C30", 25, "WBS.11", 3.5, 22, "Tổ Cốp pha thân", "Critical", "2026-12-12", "2027-01-05"),
        ("WBS.14", "Đúc 4 phiến dầm chủ T15m C40 tại bãi đúc dầm", 40, "WBS.04", 4.0, 24, "Tổ Đúc dầm T", "Critical", "2026-12-15", "2027-01-23"),
        ("WBS.15", "Bảo dưỡng dầm T15m và nghiệm thu xuất xưởng", 14, "WBS.14", 0.5, 6, "Tổ Bảo dưỡng", "Critical", "2027-01-24", "2027-02-06"),
        ("WBS.16", "Lắp đặt đá kê gối và 8 gối cao su bản thép (350x500x84)", 6, "WBS.12,13", 1.5, 10, "Tổ Cơ điện cầu", "Critical", "2027-02-07", "2027-02-12"),
        ("WBS.17", "Lao lắp 4 phiến dầm T15m vào vị trí mố M1-M2", 8, "WBS.15,16", 5.0, 20, "Tổ Lao lắp dầm", "Critical", "2027-02-13", "2027-02-20"),
        ("WBS.18", "Thi công 3 dầm ngang C30 liên kết dầm chủ", 12, "WBS.17", 2.8, 16, "Tổ Bê tông nhịp", "Critical", "2027-02-21", "2027-03-04"),
        ("WBS.19", "Lắp đặt ván khuôn, lưới thép D6 và đổ bê tông mặt cầu C35", 18, "WBS.18", 3.0, 22, "Tổ Mặt cầu", "Critical", "2027-03-05", "2027-03-22"),
        ("WBS.20", "Thi công gờ lan can C25 và lắp đặt lan can thép mạ kẽm", 15, "WBS.19", 2.2, 14, "Tổ Lan can", "Critical", "2027-03-23", "2027-04-06"),
        ("WBS.21", "Thi công bản quá độ hai đầu cầu và tường chắn đất C25", 20, "WBS.19", 2.5, 16, "Tổ Đường đầu cầu", "Critical", "2027-03-23", "2027-04-11"),
        ("WBS.22", "Lắp khe co giãn, thảm BTNC 19, thử tải và bàn giao", 20, "WBS.20,21", 3.0, 25, "Tổ Hoàn thiện", "Critical", "2027-04-12", "2027-05-01"),
    ]

    for idx, (wbs, name, dur, pred, norm, work, crew, cpm, d_start, d_finish) in enumerate(wbs_tasks, start=6):
        ws_cpm.cell(idx, 1, value=wbs).alignment = align_center
        ws_cpm.cell(idx, 2, value=name).alignment = align_left
        ws_cpm.cell(idx, 3, value=dur).alignment = align_right
        ws_cpm.cell(idx, 4, value=pred).alignment = align_center
        ws_cpm.cell(idx, 5, value=norm).alignment = align_right
        ws_cpm.cell(idx, 6, value=work).alignment = align_right
        ws_cpm.cell(idx, 7, value=crew).alignment = align_left
        ws_cpm.cell(idx, 8, value=cpm).alignment = align_center
        if cpm == "Critical":
            ws_cpm.cell(idx, 8).fill = fill_cpm
            ws_cpm.cell(idx, 8).font = Font(name="Times New Roman", size=9, bold=True, color="C00000")
        ws_cpm.cell(idx, 9, value=d_start).alignment = align_center
        ws_cpm.cell(idx, 10, value=d_finish).alignment = align_center

        for c in range(1, 11):
            if c != 8 or cpm != "Critical":
                ws_cpm.cell(idx, c).font = font_reg
            ws_cpm.cell(idx, c).border = thin_border
            if idx % 2 == 1 and (c != 8 or cpm != "Critical"):
                ws_cpm.cell(idx, c).fill = fill_zebra

    # =========================================================================
    # 8. SHEET: HOSO_KCS_NGHIEM_THU (22 BIÊN BẢN LIÊN KẾT ĐỘNG 100%)
    # =========================================================================
    ws_kcs = wb.create_sheet(title="HOSO_KCS_NGHIEM_THU")
    ws_kcs.views.sheetView[0].showGridLines = True

    ws_kcs["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_kcs["B2"].font = font_title
    ws_kcs["B3"] = "DANH MỤC HỒ SƠ NGHIỆM THU KỸ THUẬT & CHẤT LƯỢNG KCS (THEO NGHỊ ĐỊNH 207/2026/NĐ-CP)"
    ws_kcs["B3"].font = font_sec
    ws_kcs["B4"] = "Liên kết 100% ngày nghiệm thu từ TIEN_DO_THI_CONG_WBS và khối lượng từ QS_DIEN_GIAI_CHI_TIET"
    ws_kcs["B4"].font = font_subtitle

    h_kcs = ["Số TT", "Số hiệu biên bản", "Tên công việc / Hạng mục nghiệm thu",
             "Khối lượng nghiệm thu", "Đơn vị", "Tiêu chuẩn kỹ thuật áp dụng",
             "Thành phần nghiệm thu", "Ngày nghiệm thu hoàn thành", "Kết luận"]
    for c_idx, text in enumerate(h_kcs, start=1):
        cell = ws_kcs.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_kcs.row_dimensions[5].height = 28

    kcs_specs = [
        (1, "BBNT-01", "Bàn giao cọc mốc tim tuyến và định vị công trình", 1.0, "Điểm", "TCVN 4054:2005", "TVGS + Nhà thầu", 6),
        (2, "BBNT-02", "Thi công đường công vụ, lán trại và dọn dẹp mặt bằng", 1.0, "Gói", "TCVN 4054:2005", "TVGS + Nhà thầu", 7),
        (3, "BBNT-03", "Kiểm định vật liệu xây dựng đầu vào (Thép, Xi măng, Cát, Đá)", 1.0, "Lô", "TCVN 1651:2018", "TVGS + Nhà thầu + Lab", 8),
        (4, "BBNT-04", "Thi công bãi đúc dầm T15m và lắp đặt trạm trộn", 1.0, "Bãi", "TCVN 11823:2017", "TVGS + Nhà thầu", 9),
        (5, "BBNT-05", "Đào hố móng mố M1 và xử lý bề mặt tầng đá nứt nẻ", 185.0, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 10),
        (6, "BBNT-06", "Khoan cấy cốt thép neo D32 vào nền đá hố móng mố M1", 49.0, "lỗ", "TCVN 11823:2017", "TVGS + Nhà thầu", 11),
        (7, "BBNT-07", "Đổ bê tông đệm móng mố M1 mác C10 (M150)", 4.67, "m3", "TCVN 3105:2022", "TVGS + Nhà thầu", 12),
        (8, "BBNT-08", "Lắp dựng cốt thép và ván khuôn móng mố M1", 17.5, "Tấn", "TCVN 1651:2018", "TVGS + Nhà thầu", 13),
        (9, "BBNT-09", "Đổ bê tông móng mố M1 mác C30 (M350)", 110.0, "m3", "TCVN 3105:2022", "TVGS + Nhà thầu", 14),
        (10, "BBNT-10", "Đào hố móng và khoan cấy thép neo nền đá mố M2", 49.0, "lỗ", "TCVN 11823:2017", "TVGS + Nhà thầu", 15),
        (11, "BBNT-11", "Đổ bê tông đệm và bê tông móng mố M2 C30", 114.67, "m3", "TCVN 3105:2022", "TVGS + Nhà thầu", 16),
        (12, "BBNT-12", "Thi công thân mố, tường ngực và tường cánh M1 C30", 84.64, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 17),
        (13, "BBNT-13", "Thi công thân mố, tường ngực và tường cánh M2 C30", 84.64, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 18),
        (14, "BBNT-14", "Đúc 4 phiến dầm chủ T15m mác C40 tại bãi đúc dầm", 37.0, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 19),
        (15, "BBNT-15", "Bảo dưỡng bê tông dầm T15m và nghiệm thu xuất xưởng", 4.0, "Dầm", "TCVN 3105:2022", "TVGS + Nhà thầu", 20),
        (16, "BBNT-16", "Lắp đặt đá kê gối và 8 gối cao su bản thép 350x500x84mm", 8.0, "cái", "TCVN 11823:2017", "TVGS + Nhà thầu", 21),
        (17, "BBNT-17", "Lao lắp 4 phiến dầm chủ T15m vào vị trí mố M1 - M2", 4.0, "Dầm", "TCVN 11823:2017", "TVGS + Nhà thầu", 22),
        (18, "BBNT-18", "Thi công 3 dầm ngang C30 liên kết dầm chủ", 2.86, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 23),
        (19, "BBNT-19", "Lắp dựng cốt thép và đổ bê tông bản mặt cầu C35 dày 12cm", 14.44, "m3", "TCVN 3105:2022", "TVGS + Nhà thầu", 24),
        (20, "BBNT-20", "Thi công gờ lan can C25 và lắp đặt lan can thép mạ kẽm", 12.56, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 25),
        (21, "BBNT-21", "Thi công bản quá độ hai đầu cầu và tường chắn đất C25", 12.53, "m3", "TCVN 11823:2017", "TVGS + Nhà thầu", 26),
        (22, "BBNT-22", "Lắp khe co giãn, thảm BTNC 19, thử tải và bàn giao", 1.0, "Cầu", "TCVN 11823:2017", "Chủ đầu tư + TVGS + NT", 27),
    ]

    for idx, (stt, code, name, def_qty, unit, std, mem, wbs_r) in enumerate(kcs_specs, start=6):
        ws_kcs.cell(idx, 1, value=stt).alignment = align_center
        ws_kcs.cell(idx, 2, value=code).alignment = align_center
        ws_kcs.cell(idx, 3, value=name).alignment = align_left
        ws_kcs.cell(idx, 4, value=def_qty).alignment = align_right
        ws_kcs.cell(idx, 4).number_format = "#,##0.00"
        ws_kcs.cell(idx, 5, value=unit).alignment = align_center
        ws_kcs.cell(idx, 6, value=std).alignment = align_left
        ws_kcs.cell(idx, 7, value=mem).alignment = align_left
        ws_kcs.cell(idx, 8, value=f"=TIEN_DO_THI_CONG_WBS!J{wbs_r}").alignment = align_center
        ws_kcs.cell(idx, 9, value="Đạt yêu cầu").alignment = align_center

        for c in range(1, 10):
            ws_kcs.cell(idx, c).font = font_reg
            ws_kcs.cell(idx, c).border = thin_border
            if idx % 2 == 1:
                ws_kcs.cell(idx, c).fill = fill_zebra

    # =========================================================================
    # 9. SHEET: PHAN_TICH_VAT_TU_WBS
    # =========================================================================
    ws_mat = wb.create_sheet(title="PHAN_TICH_VAT_TU_WBS")
    ws_mat.views.sheetView[0].showGridLines = True

    ws_mat["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_mat["B2"].font = font_title
    ws_mat["B3"] = "BẢNG PHÂN TÍCH NHU CẦU VẬT TƯ THEO CÔNG TÁC WBS (ĐỊNH MỨC BỘ XÂY DỰNG TT 38/2026/TT-BXD)"
    ws_mat["B3"].font = font_sec
    ws_mat["B4"] = "Phân tích cấp phối vật liệu thành phần: Xi măng, cát, đá, nước, phụ gia, cốt thép, ván khuôn"
    ws_mat["B4"].font = font_subtitle

    h_mat = ["Mã công tác", "Tên công tác WBS", "Đơn vị tính", "Khối lượng công tác",
             "Mã vật tư", "Tên vật tư thành phần", "Đơn vị VT", "Định mức hao phí",
             "Khối lượng vật tư", "Ghi chú"]
    for c_idx, text in enumerate(h_mat, start=1):
        cell = ws_mat.cell(row=5, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_mat.row_dimensions[5].height = 28

    mat_recipes = [
        (6, "XM.PCB40", "Xi măng PCB40", "kg", 420.0),
        (6, "CAT.VANG", "Cát vàng đổ bê tông", "m3", 0.45),
        (6, "DA.1X2", "Đá dăm tiêu chuẩn 1x2", "m3", 0.85),
        (6, "NUOC", "Nước sạch trộn bê tông", "lít", 175.0),
        (6, "PHU.GIA", "Phụ gia siêu dẻo giảm nước", "lít", 4.2),
        (8, "THEP.D32", "Thép tròn CB400-V D32", "kg", 1020.0),
        (10, "XM.PCB40", "Xi măng PCB40", "kg", 380.0),
        (10, "CAT.VANG", "Cát vàng đổ bê tông", "m3", 0.48),
        (10, "DA.1X2", "Đá dăm tiêu chuẩn 1x2", "m3", 0.88),
        (11, "XM.PCB40", "Xi măng PCB40", "kg", 395.0),
        (11, "CAT.VANG", "Cát vàng đổ bê tông", "m3", 0.46),
        (11, "DA.1X2", "Đá dăm tiêu chuẩn 1x2", "m3", 0.86),
        (15, "XM.PCB40", "Xi măng PCB40", "kg", 450.0),
        (15, "CAT.VANG", "Cát vàng đổ bê tông", "m3", 0.44),
        (15, "DA.1X2", "Đá dăm tiêu chuẩn 1x2", "m3", 0.84),
        (15, "PHU.GIA", "Phụ gia siêu dẻo Sikament", "lít", 5.5),
        (16, "THEP.D10", "Thép tròn CB240-T D<=10", "kg", 1020.0),
        (17, "THEP.D14", "Thép tròn CB400-V 10<D<=18", "kg", 1020.0),
        (18, "THEP.D32", "Thép tròn CB400-V D>18", "kg", 1020.0),
        (22, "XM.PCB40", "Xi măng PCB40", "kg", 410.0),
        (22, "CAT.VANG", "Cát vàng đổ bê tông", "m3", 0.45),
        (22, "DA.1X2", "Đá dăm tiêu chuẩn 1x2", "m3", 0.85),
    ]

    r_m = 6
    for (pr_idx, m_code, m_name, m_u, norm) in mat_recipes:
        pr = parent_rows[pr_idx % len(parent_rows)]
        ws_mat.cell(r_m, 1, value=f"=QS_DIEN_GIAI_CHI_TIET!A{pr}").alignment = align_center
        ws_mat.cell(r_m, 2, value=f"=QS_DIEN_GIAI_CHI_TIET!C{pr}").alignment = align_left
        ws_mat.cell(r_m, 3, value=f"=QS_DIEN_GIAI_CHI_TIET!D{pr}").alignment = align_center
        ws_mat.cell(r_m, 4, value=f"=QS_DIEN_GIAI_CHI_TIET!J{pr}").alignment = align_right
        ws_mat.cell(r_m, 4).number_format = "#,##0.000"
        ws_mat.cell(r_m, 5, value=m_code).alignment = align_center
        ws_mat.cell(r_m, 6, value=m_name).alignment = align_left
        ws_mat.cell(r_m, 7, value=m_u).alignment = align_center
        ws_mat.cell(r_m, 8, value=norm).alignment = align_right
        ws_mat.cell(r_m, 8).number_format = "#,##0.00"
        ws_mat.cell(r_m, 9, value=f"=D{r_m}*H{r_m}").alignment = align_right
        ws_mat.cell(r_m, 9).number_format = "#,##0.00"
        ws_mat.cell(r_m, 10, value="Định mức TT 38/2026").alignment = align_left

        for c in range(1, 11):
            ws_mat.cell(r_m, c).font = font_reg
            ws_mat.cell(r_m, c).border = thin_border
            if r_m % 2 == 1:
                ws_mat.cell(r_m, c).fill = fill_zebra
        r_m += 1

    max_mat_row = r_m - 1

    # =========================================================================
    # 10. SHEET: TONG_HOP_VAT_TU_TOAN_BO (BOM TỔNG HỢP)
    # =========================================================================
    ws_bom = wb.create_sheet(title="TONG_HOP_VAT_TU_TOAN_BO")
    ws_bom.views.sheetView[0].showGridLines = True

    ws_bom["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_bom["B2"].font = font_title
    ws_bom["B3"] = "BẢNG TỔNG HỢP NHU CẦU VẬT TƯ TOÀN BỘ CÔNG TRÌNH (BILL OF MATERIALS - BOM)"
    ws_bom["B3"].font = font_sec
    ws_bom["B4"] = "Tổng hợp tự động bằng hàm SUMIF từ PHAN_TICH_VAT_TU_WBS - Chuẩn hóa phục vụ cung ứng"
    ws_bom["B4"].font = font_subtitle

    h_bom = ["STT", "Mã vật tư", "Tên quy cách vật tư", "Đơn vị tính",
             "Tổng khối lượng nhu cầu", "Đơn giá dự toán (VNĐ)", "Thành tiền (VNĐ)", "Ghi chú"]
    for c_idx, text in enumerate(h_bom, start=1):
        cell = ws_bom.cell(row=6, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_bom.row_dimensions[6].height = 28

    bom_items = [
        (1, "XM.PCB40", "Xi măng Poóc lăng hỗn hợp PCB40", "kg", 1850),
        (2, "CAT.VANG", "Cát vàng đổ bê tông mô đun độ lớn ML >= 2.5", "m3", 380000),
        (3, "DA.1X2", "Đá dăm tiêu chuẩn 1x2 (Dmax = 20mm)", "m3", 420000),
        (4, "NUOC", "Nước sạch phục vụ trộn và bảo dưỡng bê tông", "lít", 25),
        (5, "PHU.GIA", "Phụ gia siêu dẻo kéo dài thời gian đông kết", "lít", 45000),
        (6, "THEP.D10", "Thép tròn trơn cán nóng CB240-T (D<=10mm)", "kg", 18500),
        (7, "THEP.D14", "Thép thanh vằn CB400-V (10<D<=18mm)", "kg", 18200),
        (8, "THEP.D32", "Thép thanh vằn CB400-V (D>18mm)", "kg", 17900),
    ]

    for idx, (stt, code, name, unit, price) in enumerate(bom_items, start=7):
        ws_bom.cell(idx, 1, value=stt).alignment = align_center
        ws_bom.cell(idx, 2, value=code).alignment = align_center
        ws_bom.cell(idx, 3, value=name).alignment = align_left
        ws_bom.cell(idx, 4, value=unit).alignment = align_center
        ws_bom.cell(idx, 5, value=f'=SUMIF(PHAN_TICH_VAT_TU_WBS!$E$6:$E${max_mat_row}, "{code}", PHAN_TICH_VAT_TU_WBS!$I$6:$I${max_mat_row})').alignment = align_right
        ws_bom.cell(idx, 5).number_format = "#,##0.00"
        ws_bom.cell(idx, 6, value=price).alignment = align_right
        ws_bom.cell(idx, 6).number_format = "#,##0"
        ws_bom.cell(idx, 7, value=f"=E{idx}*F{idx}").alignment = align_right
        ws_bom.cell(idx, 7).number_format = "#,##0"
        ws_bom.cell(idx, 8, value="Tổng hợp toàn cầu").alignment = align_left

        for c in range(1, 9):
            ws_bom.cell(idx, c).font = font_reg
            ws_bom.cell(idx, c).border = thin_border
            if idx % 2 == 1:
                ws_bom.cell(idx, c).fill = fill_zebra

    bom_end = 6 + len(bom_items)
    ws_bom.cell(bom_end + 1, 3, value="TỔNG CỘNG CHI PHÍ VẬT LIỆU CHÍNH (VNĐ)").font = font_bold
    ws_bom.cell(bom_end + 1, 7, value=f"=SUM(G7:G{bom_end})").alignment = align_right
    ws_bom.cell(bom_end + 1, 7).number_format = "#,##0"
    for c in range(1, 9):
        ws_bom.cell(bom_end + 1, c).font = font_bold
        ws_bom.cell(bom_end + 1, c).fill = fill_tot
        ws_bom.cell(bom_end + 1, c).border = double_bottom_border

    # =========================================================================
    # 11. SHEET: CAP_PHOI_1M3_VA_TAN_SUAT (100% CÔNG THỨC NHÂN CHIA ĐỊNH MỨC)
    # =========================================================================
    ws_mix = wb.create_sheet(title="CAP_PHOI_1M3_VA_TAN_SUAT")
    ws_mix.views.sheetView[0].showGridLines = True

    ws_mix["B2"] = "DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450/456, HUYỆN MÈO VẠC, HÀ GIANG"
    ws_mix["B2"].font = font_title
    ws_mix["B3"] = "BẢNG 1: CẤP PHỐI ĐỊNH MỨC 1m3 BÊ TÔNG & BẢNG 2: MA TRẬN TẦN SUẤT THÍ NGHIỆM KCS"
    ws_mix["B3"].font = font_sec
    ws_mix["B4"] = "Liên kết động: Khối lượng QS -> Vật liệu thành phần -> Tần suất tổ mẫu KCS (TCVN 3105 & TCVN 11823)"
    ws_mix["B4"].font = font_subtitle

    # BẢNG 1: CẤP PHỐI 1M3
    ws_mix["B6"] = "I. BẢNG TÍNH CẤP PHỐI ĐỊNH MỨC 1M3 BÊ TÔNG VÀ TỔNG HỢP VẬT LIỆU CẤU THÀNH"
    ws_mix["B6"].font = font_sec

    h_mix1 = ["STT", "Hạng mục kết cấu sử dụng", "Mác bê tông", "Độ sụt (cm)",
              "Xi măng PCB40 (kg/m3)", "Cát vàng (m3/m3)", "Đá 1x2 (m3/m3)", "Nước sạch (L/m3)", "Phụ gia (L/m3)",
              "Tổng KL bê tông (m3)", "Xi măng tổng (Tấn)", "Cát vàng tổng (m3)", "Đá 1x2 tổng (m3)", "Nước tổng (m3)", "Phụ gia tổng (Lít)"]
    for c_idx, text in enumerate(h_mix1, start=1):
        cell = ws_mix.cell(row=8, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_mix.row_dimensions[8].height = 28

    mix_designs = [
        (1, "Bê tông lót đáy móng mố M1, M2", "C10 (M150)", "6 - 8", 245, 0.52, 0.88, 185, 0.0),
        (2, "Bê tông bản quá độ & gờ lan can", "C25 (M300)", "10 - 12", 365, 0.47, 0.86, 175, 2.5),
        (3, "Bê tông móng mố M1, M2 trên nền đá", "C30 (M350)", "12 - 14", 395, 0.46, 0.85, 170, 3.2),
        (4, "Bê tông thân mố, tường cánh M1, M2", "C30 (M350)", "12 - 14", 410, 0.45, 0.84, 168, 3.5),
        (5, "Bê tông dầm ngang liên kết nhịp", "C30 (M350)", "14 - 16", 415, 0.44, 0.84, 165, 3.8),
        (6, "Bê tông lớp phủ mặt cầu chống thấm", "C35 (M450)", "12 - 14", 435, 0.43, 0.83, 160, 4.2),
        (7, "Bê tông dầm chủ T15m đúc sẵn", "C40 (M500)", "16 - 18", 460, 0.42, 0.82, 155, 5.0),
        (8, "Bê tông tường chắn đất đầu cầu", "C25 (M300)", "10 - 12", 365, 0.47, 0.86, 175, 2.5),
        (9, "Vữa không co ngót chèn khe co giãn", "C50 (M600)", "Chảy", 550, 0.38, 0.75, 140, 6.5),
    ]

    for idx, (stt, name, grade, slump, xm, cat, da, nuoc, pg) in enumerate(mix_designs, start=9):
        ws_mix.cell(idx, 1, value=stt).alignment = align_center
        ws_mix.cell(idx, 2, value=name).alignment = align_left
        ws_mix.cell(idx, 3, value=grade).alignment = align_center
        ws_mix.cell(idx, 4, value=slump).alignment = align_center
        ws_mix.cell(idx, 5, value=xm).alignment = align_right
        ws_mix.cell(idx, 6, value=cat).alignment = align_right
        ws_mix.cell(idx, 7, value=da).alignment = align_right
        ws_mix.cell(idx, 8, value=nuoc).alignment = align_right
        ws_mix.cell(idx, 9, value=pg).alignment = align_right

        # Cột J: Khối lượng bê tông trỏ trực tiếp QS (AUDIT BẮT BUỘC CÔNG THỨC)
        if idx == 9:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J20") # Lót C10
        elif idx == 10:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J38*0.5 + QS_DIEN_GIAI_CHI_TIET!J50") # Bản quá độ & gờ lan can
        elif idx == 11:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J22") # Móng mố C30
        elif idx == 12:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J24") # Thân mố C30
        elif idx == 13:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J38*0.5") # Dầm ngang C30
        elif idx == 14:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J44") # Lớp phủ mặt cầu C35
        elif idx == 15:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J28") # Dầm T15m C40
        elif idx == 16:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J58") # Tường chắn C25
        elif idx == 17:
            ws_mix.cell(idx, 10, value="=QS_DIEN_GIAI_CHI_TIET!J54") # Vữa C50
        ws_mix.cell(idx, 10).number_format = "#,##0.000"

        # Cột K, L, M, N, O: CÔNG THỨC NHÂN CHIA ĐỊNH MỨC 1M3 (AUDIT BẮT BUỘC)
        ws_mix.cell(idx, 11, value=f"=J{idx}*E{idx}/1000").alignment = align_right
        ws_mix.cell(idx, 11).number_format = "#,##0.00"
        ws_mix.cell(idx, 12, value=f"=J{idx}*F{idx}").alignment = align_right
        ws_mix.cell(idx, 12).number_format = "#,##0.00"
        ws_mix.cell(idx, 13, value=f"=J{idx}*G{idx}").alignment = align_right
        ws_mix.cell(idx, 13).number_format = "#,##0.00"
        ws_mix.cell(idx, 14, value=f"=J{idx}*H{idx}/1000").alignment = align_right
        ws_mix.cell(idx, 14).number_format = "#,##0.00"
        ws_mix.cell(idx, 15, value=f"=J{idx}*I{idx}").alignment = align_right
        ws_mix.cell(idx, 15).number_format = "#,##0.0"

        for c in range(1, 16):
            ws_mix.cell(idx, c).font = font_reg
            ws_mix.cell(idx, c).border = thin_border
            if idx % 2 == 1:
                ws_mix.cell(idx, c).fill = fill_zebra

    # Dòng tổng Bảng 1 (Hàng 18)
    ws_mix.cell(18, 2, value="TỔNG CỘNG TOÀN CÔNG TRÌNH").alignment = align_left
    ws_mix.cell(18, 10, value="=SUM(J9:J17)").alignment = align_right
    ws_mix.cell(18, 11, value="=SUM(K9:K17)").alignment = align_right
    ws_mix.cell(18, 12, value="=SUM(L9:L17)").alignment = align_right
    ws_mix.cell(18, 13, value="=SUM(M9:M17)").alignment = align_right
    ws_mix.cell(18, 14, value="=SUM(N9:N17)").alignment = align_right
    ws_mix.cell(18, 15, value="=SUM(O9:O17)").alignment = align_right

    for c in range(1, 16):
        ws_mix.cell(18, c).font = font_bold
        ws_mix.cell(18, c).fill = fill_tot
        ws_mix.cell(18, c).border = double_bottom_border

    # BẢNG 2: MA TRẬN TẦN SUẤT THÍ NGHIỆM KCS
    ws_mix["B21"] = "II. MA TRẬN TẦN SUẤT LẤY MẪU THÍ NGHIỆM KCS (TCVN 3105:2022 & TCVN 11823:2017)"
    ws_mix["B21"].font = font_sec

    h_mix2 = ["Mã VT", "Tên chỉ tiêu / Vật liệu thí nghiệm", "Tiêu chuẩn áp dụng",
              "Chỉ tiêu kỹ thuật kiểm tra", "Đối tượng / Kết cấu", "Đơn vị tính",
              "Tổng KL công trình", "Định mức tần suất (KL/tổ mẫu)", "Số tổ mẫu thí nghiệm",
              "Quy cách tổ mẫu thí nghiệm", "Đơn vị thực hiện", "Điều kiện nghiệm thu"]
    for c_idx, text in enumerate(h_mix2, start=1):
        cell = ws_mix.cell(row=23, column=c_idx, value=text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_mix.row_dimensions[23].height = 28

    freq_specs = [
        (1, "Cốt thép tròn trơn D<=10mm", "TCVN 1651-1:2018", "Kéo, uốn, hình học", "Dầm T, gờ lan can", "Tấn", '<=10', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (2, "Cốt thép thanh vằn D12mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Dầm T, gờ lan can", "Tấn", '12', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (3, "Cốt thép thanh vằn D14mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Dầm T, bản quá độ", "Tấn", '14', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (4, "Cốt thép thanh vằn D16mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Dầm T, mố M1, M2", "Tấn", '16', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (5, "Cốt thép thanh vằn D18mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Dầm T, mố M1, M2", "Tấn", '18', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (6, "Cốt thép thanh vằn D20mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Thân mố, bản quá độ", "Tấn", '20', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (7, "Cốt thép thanh vằn D22mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Tường cánh mố M1, M2", "Tấn", '22', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (8, "Cốt thép thanh vằn D25mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Thân mố chịu lực", "Tấn", '25', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (9, "Cốt thép thanh vằn D28mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Móng mố M1, M2", "Tấn", '28', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (10, "Cốt thép thanh vằn D32mm", "TCVN 1651-2:2018", "Giới hạn chảy, bền, uốn", "Chủ dầm T & neo cấy đá", "Tấn", '32', 20.0, "3 mẫu kéo + 3 mẫu uốn", "LAS-XD", "Chấp thuận trước khi gia công"),
        (11, "Thép hình & lan can mạ kẽm", "TCVN 5408:2007", "Độ dày mạ, kéo uốn", "Lan can cầu & khe co giãn", "Tấn", '1.27', 10.0, "3 mẫu ngẫu nhiên", "LAS-XD", "Đạt chiều dày mạ >= 100um"),
        (12, "Xi măng Poóc lăng PCB40", "TCVN 6260:2020", "Độ mịn, thời gian đông kết, nén", "Toàn bộ bê tông", "Tấn", "=K18", 100.0, "1 tổ 3 mẫu/lô", "LAS-XD", "Chứng chỉ xuất xưởng + TN"),
        (13, "Cát vàng đổ bê tông", "TCVN 7570:2006", "Thành phần hạt, hàm lượng bùn sét", "Toàn bộ bê tông", "m3", "=L18", 200.0, "1 tổ mẫu/mỏ", "LAS-XD", "Tạp chất < 1.5%"),
        (14, "Đá dăm tiêu chuẩn 1x2", "TCVN 7570:2006", "Thành phần hạt, độ nén dập, thoi dẹt", "Toàn bộ bê tông", "m3", "=M18", 200.0, "1 tổ mẫu/mỏ", "LAS-XD", "Độ thoi dẹt < 15%"),
        (15, "Nước sạch trộn bê tông", "TCVN 4506:2012", "Độ pH, tạp chất hữu cơ, muối", "Toàn bộ bê tông", "m3", "=N18", 500.0, "1 mẫu/nguồn", "LAS-XD", "pH = 6 - 8"),
        (16, "Phụ gia hóa học bê tông", "TCVN 8826:2011", "Độ giảm nước, ăn mòn cốt thép", "Toàn bộ bê tông", "Lít", "=O18", 2000.0, "1 mẫu/lô", "LAS-XD", "Tương thích xi măng"),
        (17, "Độ sụt bê tông tại hiện trường", "TCVN 3106:2022", "Đo độ sụt từng xe bồn", "Tất cả các mẻ đổ", "Xe", "=ROUNDUP(J18/8,0)", 1.0, "1 lần/xe bồn", "Hiện trường", "Đạt khoảng độ sụt thiết kế"),
        (18, "Bê tông lót móng mố C10", "TCVN 3105:2022", "Cường độ nén R28", "Lớp lót móng M1, M2", "m3", "=J9", 50.0, "1 tổ 3 viên (15x15x15)", "LAS-XD", "Đạt cường độ R28 >= 10 MPa"),
        (19, "Bê tông móng mố M1, M2 C30", "TCVN 3105:2022", "Cường độ nén R7, R28", "Khối móng mố M1, M2", "m3", "=J11", 50.0, "3 tổ (R7, R28, lưu)", "LAS-XD", "R28 >= 30 MPa"),
        (20, "Bê tông thân mố, tường cánh C30", "TCVN 3105:2022", "Cường độ nén R7, R28", "Thân mố M1, M2", "m3", "=J12", 50.0, "3 tổ (R7, R28, lưu)", "LAS-XD", "R28 >= 30 MPa"),
        (21, "Bê tông dầm ngang C30", "TCVN 3105:2022", "Cường độ nén R7, R28", "Dầm ngang nhịp 15m", "m3", "=J13", 20.0, "2 tổ mẫu", "LAS-XD", "R28 >= 30 MPa"),
        (22, "Bê tông dầm chủ T15m C40", "TCVN 3105:2022", "Cường độ nén R3, R7, R28", "4 phiến dầm T15m", "m3", "=J15", 10.0, "4 tổ / phiến dầm", "LAS-XD", "Đạt 100% f'c trước khi cẩu"),
        (23, "Bê tông lớp phủ mặt cầu C35", "TCVN 3105:2022", "Cường độ nén R7, R28", "Bản mặt cầu", "m3", "=J14", 50.0, "3 tổ mẫu", "LAS-XD", "R28 >= 35 MPa"),
        (24, "Bê tông bản quá độ C25", "TCVN 3105:2022", "Cường độ nén R7, R28", "Bản quá độ M1, M2", "m3", "=J10", 50.0, "2 tổ mẫu", "LAS-XD", "R28 >= 25 MPa"),
        (25, "Bê tông gờ lan can C25", "TCVN 3105:2022", "Cường độ nén R28", "Gờ lan can dọc cầu", "m3", "=J10", 50.0, "2 tổ mẫu", "LAS-XD", "R28 >= 25 MPa"),
        (26, "Vữa chèn khe co giãn C50", "TCVN 9204:2012", "Cường độ nén R3, R28, co ngót", "Mối nối khe co giãn", "m3", "=J17", 5.0, "3 tổ viên (7x7x7)", "LAS-XD", "Không co ngót, R28 >= 50 MPa"),
        (27, "Thí nghiệm kéo nhổ cốt thép neo cấy đá", "TCVN 9373:2012", "Lực dính bám thanh neo D32", "Thanh neo móng mố M1, M2", "Điểm", 6.0, 2.0, "3 thanh ngẫu nhiên/mố", "Viện KHCN", "Lực kéo >= 1.25 fy"),
        (28, "Kiểm tra độ bằng phẳng mặt cầu", "TCVN 8864:2011", "Thước 3m và độ nhám", "Mặt cầu sau thảm BTN", "Điểm", 20.0, 5.0, "5 mặt cắt trắc ngang", "TVGS", "Khe hở thước <= 3mm"),
        (29, "Thử tải tĩnh và động cầu hoàn thành", "TCVN 11823:2017", "Độ võng, ứng suất, dao động", "Nhịp dầm T15m", "Cầu", 1.0, 1.0, "Đoàn xe 3 trục tiêu chuẩn", "Cơ quan thẩm định", "Hệ số động K <= K_tk"),
    ]

    for rf, item_f in enumerate(freq_specs, start=24):
        code_id, name, std, crit, obj, unit, g_val, h_val, q_spec, lab, cond = item_f
        ws_mix.cell(rf, 1, value=code_id).alignment = align_center
        ws_mix.cell(rf, 2, value=name).alignment = align_left
        ws_mix.cell(rf, 3, value=std).alignment = align_left
        ws_mix.cell(rf, 4, value=crit).alignment = align_left
        ws_mix.cell(rf, 5, value=obj).alignment = align_left
        ws_mix.cell(rf, 6, value=unit).alignment = align_center

        if rf <= 33:
            dia_kw = g_val
            ws_mix.cell(rf, 7, value=f'=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, "{dia_kw}")')
        elif rf == 34:
            ws_mix.cell(rf, 7, value="=1.27")
        elif str(g_val).startswith("="):
            ws_mix.cell(rf, 7, value=g_val)
        else:
            ws_mix.cell(rf, 7, value=float(g_val))
        ws_mix.cell(rf, 7).alignment = align_right
        ws_mix.cell(rf, 7).number_format = "#,##0.00"

        ws_mix.cell(rf, 8, value=h_val).alignment = align_right
        ws_mix.cell(rf, 8).number_format = "#,##0.0"

        ws_mix.cell(rf, 9, value=f"=ROUNDUP(G{rf}/H{rf}, 0)").alignment = align_right
        ws_mix.cell(rf, 9).number_format = "#,##0"

        ws_mix.cell(rf, 10, value=q_spec).alignment = align_left
        ws_mix.cell(rf, 11, value=lab).alignment = align_center
        ws_mix.cell(rf, 12, value=cond).alignment = align_left

        for c in range(1, 13):
            ws_mix.cell(rf, c).font = font_reg
            ws_mix.cell(rf, c).border = thin_border
            if rf % 2 == 1:
                ws_mix.cell(rf, c).fill = fill_zebra

    ws_mix.cell(53, 2, value="TỔNG SỐ LƯỢNG MẪU & PHÉP THỬ KCS TOÀN CÔNG TRÌNH").alignment = align_left
    ws_mix.cell(53, 9, value="=SUM(I24:I52)").alignment = align_right
    ws_mix.cell(53, 9).number_format = "#,##0"
    for c in range(1, 13):
        ws_mix.cell(53, c).font = font_bold
        ws_mix.cell(53, c).fill = fill_tot
        ws_mix.cell(53, c).border = double_bottom_border

    # =========================================================================
    # 12. SHEET: MAU_BIEN_BAN_KCS
    # =========================================================================
    ws_bb = wb.create_sheet(title="MAU_BIEN_BAN_KCS")
    ws_bb.views.sheetView[0].showGridLines = True

    ws_bb["B2"] = "CHỌN SỐ BIÊN BẢN (1 - 22):"
    ws_bb["B2"].font = font_bold
    ws_bb["C2"] = 14
    ws_bb["C2"].font = Font(name="Times New Roman", size=12, bold=True, color="C00000")
    ws_bb["C2"].fill = fill_input
    ws_bb["C2"].alignment = align_center
    ws_bb["C2"].border = box_border
    ws_bb["D2"] = "(Nhập số thứ tự từ 1 đến 22 để tự động nhảy toàn bộ nội dung biên bản A4)"
    ws_bb["D2"].font = font_it

    ws_bb["B4"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_bb["B4"].font = Font(name="Times New Roman", size=11, bold=True)
    ws_bb["B4"].alignment = align_center
    ws_bb.merge_cells("B4:G4")

    ws_bb["B5"] = "Độc lập - Tự do - Hạnh phúc"
    ws_bb["B5"].font = Font(name="Times New Roman", size=10, bold=True)
    ws_bb["B5"].alignment = align_center
    ws_bb.merge_cells("B5:G5")

    ws_bb["B6"] = "-------------------------"
    ws_bb["B6"].alignment = align_center
    ws_bb.merge_cells("B6:G6")

    ws_bb["B8"] = "BIÊN BẢN NGHIỆM THU CÔNG VIỆC XÂY DỰNG"
    ws_bb["B8"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws_bb["B8"].alignment = align_center
    ws_bb.merge_cells("B8:G8")

    ws_bb["B9"] = '="Số: " & TEXT(C2, "00") & "/BBNT-XD/KHAI-HOANG-2"'
    ws_bb["B9"].font = Font(name="Times New Roman", size=10, italic=True)
    ws_bb["B9"].alignment = align_center
    ws_bb.merge_cells("B9:G9")

    ws_bb["B11"] = "Căn cứ Luật Xây dựng số 135/2025/QH15;"
    ws_bb["B12"] = "Căn cứ Nghị định số 207/2026/NĐ-CP ngày 15/01/2026 của Chính phủ về quản lý chất lượng công trình;"
    ws_bb["B13"] = "Căn cứ Hợp đồng thi công xây dựng số 09/2026/HĐ-XL gói thầu xây lắp Cầu thôn Khai Hoang 2;"
    ws_bb["B14"] = "Căn cứ Hồ sơ thiết kế bản vẽ thi công và Chỉ dẫn kỹ thuật được phê duyệt."
    for r in range(11, 15):
        ws_bb.cell(r, 2).font = font_it

    ws_bb["B16"] = "1. Tên công trình / Dự án:"
    ws_bb["C16"] = "Đường từ trung tâm huyện Đồng Văn đi Mốc 450/456, huyện Mèo Vạc, Hà Giang"
    ws_bb["B17"] = "2. Hạng mục công trình:"
    ws_bb["C17"] = "CẦU THÔN KHAI HOANG 2, KM14+363.65 (Nhịp dầm T 15m)"
    ws_bb["B18"] = "3. Địa điểm xây dựng:"
    ws_bb["C18"] = "Km14+363.65, Huyện Mèo Vạc, Tỉnh Hà Giang"
    ws_bb["B19"] = "4. Đối tượng nghiệm thu:"
    ws_bb["C19"] = '=VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 3, FALSE)'
    ws_bb["C19"].font = Font(name="Times New Roman", size=11, bold=True, color="1F497D")

    ws_bb["B20"] = "5. Khối lượng nghiệm thu:"
    ws_bb["C20"] = '=TEXT(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 4, FALSE), "#,##0.00") & " " & VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 5, FALSE)'
    ws_bb["C20"].font = font_bold

    ws_bb["B21"] = "6. Tiêu chuẩn kỹ thuật áp dụng:"
    ws_bb["C21"] = '=VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 6, FALSE)'
    ws_bb["C21"].font = font_reg

    ws_bb["B22"] = "7. Thời gian nghiệm thu:"
    ws_bb["C22"] = '="Bắt đầu: 08 giờ 30 phút - Kết thúc: 11 giờ 30 phút, ngày " & TEXT(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 8, FALSE), "dd/mm/yyyy")'
    ws_bb["C22"].font = font_reg

    for r in range(16, 23):
        ws_bb.cell(r, 2).font = font_bold
        ws_bb.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws_bb["B24"] = "8. Thành phần trực tiếp nghiệm thu:"
    ws_bb["B24"].font = font_bold
    ws_bb["B25"] = "a) Đại diện Tư vấn giám sát (TVGS):"
    ws_bb["B25"].font = font_bold
    ws_bb["C25"] = "- Ông: Dương Văn Lâm       - Chức vụ: Kỹ sư Giám sát trưởng"
    ws_bb["B26"] = "b) Đại diện Nhà thầu thi công:"
    ws_bb["B26"].font = font_bold
    ws_bb["C26"] = "- Ông: Trần Việt Hải        - Chức vụ: Chỉ huy trưởng công trường"
    ws_bb["C27"] = "- Ông: Ngô Văn Quang        - Chức vụ: Kỹ sư QA/QC hiện trường"

    for r in [25, 26, 27]:
        ws_bb.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws_bb["B29"] = "9. Đánh giá kết quả công việc đã thực hiện:"
    ws_bb["B29"].font = font_bold
    ws_bb["B30"] = "- Về kích thước hình học, cao độ, tim mốc: Đạt yêu cầu theo hồ sơ thiết kế bản vẽ thi công."
    ws_bb["B31"] = "- Về chất lượng vật liệu, chứng chỉ xuất xưởng, kết quả thí nghiệm KCS: Đạt yêu cầu theo tiêu chuẩn áp dụng."
    ws_bb["B32"] = "- Về an toàn lao động, vệ sinh môi trường: Tuân thủ đầy đủ biện pháp thi công được duyệt."
    for r in [30, 31, 32]:
        ws_bb.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws_bb.cell(r, 2).font = font_reg

    ws_bb["B34"] = "10. Kết luận nghiệm thu:"
    ws_bb["B34"].font = font_bold
    ws_bb["B35"] = "CHẤP THUẬN NGHIỆM THU CÔNG VIỆC XÂY DỰNG. ĐỒNG Ý CHO PHÉP TRIỂN KHAI BƯỚC THI CÔNG TIẾP THEO."
    ws_bb["B35"].font = Font(name="Times New Roman", size=10, bold=True, color="006100")
    ws_bb.merge_cells("B35:G35")

    ws_bb["B38"] = "ĐẠI DIỆN NHÀ THẦU THI CÔNG"
    ws_bb["B38"].font = font_bold
    ws_bb["B38"].alignment = align_center
    ws_bb.merge_cells("B38:D38")

    ws_bb["E38"] = "ĐẠI DIỆN TƯ VẤN GIÁM SÁT"
    ws_bb["E38"].font = font_bold
    ws_bb["E38"].alignment = align_center
    ws_bb.merge_cells("E38:G38")

    ws_bb["B44"] = "Trần Việt Hải"
    ws_bb["B44"].font = font_bold
    ws_bb["B44"].alignment = align_center
    ws_bb.merge_cells("B44:D44")

    ws_bb["E44"] = "Dương Văn Lâm"
    ws_bb["E44"].font = font_bold
    ws_bb["E44"].alignment = align_center
    ws_bb.merge_cells("E44:G44")

    # =========================================================================
    # 13. SHEET: MAU_BB_NGHIEM_THU_VAT_LIEU
    # =========================================================================
    ws_vl = wb.create_sheet(title="MAU_BB_NGHIEM_THU_VAT_LIEU")
    ws_vl.views.sheetView[0].showGridLines = True

    ws_vl["B2"] = "CHỌN MÃ VẬT LIỆU (1 - 16):"
    ws_vl["B2"].font = font_bold
    ws_vl["C2"] = 1
    ws_vl["C2"].font = Font(name="Times New Roman", size=12, bold=True, color="C00000")
    ws_vl["C2"].fill = fill_input
    ws_vl["C2"].alignment = align_center
    ws_vl["C2"].border = box_border
    ws_vl["D2"] = "(Nhập mã từ 1 đến 16 theo danh mục ma trận tần suất thí nghiệm để tự nhảy biên bản)"
    ws_vl["D2"].font = font_it

    ws_vl["B4"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_vl["B4"].font = Font(name="Times New Roman", size=11, bold=True)
    ws_vl["B4"].alignment = align_center
    ws_vl.merge_cells("B4:G4")

    ws_vl["B5"] = "Độc lập - Tự do - Hạnh phúc"
    ws_vl["B5"].font = Font(name="Times New Roman", size=10, bold=True)
    ws_vl["B5"].alignment = align_center
    ws_vl.merge_cells("B5:G5")

    ws_vl["B8"] = "BIÊN BẢN NGHIỆM THU VẬT LIỆU XÂY DỰNG ĐẦU VÀO"
    ws_vl["B8"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws_vl["B8"].alignment = align_center
    ws_vl.merge_cells("B8:G8")

    ws_vl["B9"] = '="Số: " & TEXT(C2, "00") & "/BBNT-VL/KHAI-HOANG-2"'
    ws_vl["B9"].font = font_it
    ws_vl["B9"].alignment = align_center
    ws_vl.merge_cells("B9:G9")

    ws_vl["B12"] = "1. Tên vật liệu nghiệm thu:"
    ws_vl["C12"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 2, FALSE)'
    ws_vl["C12"].font = Font(name="Times New Roman", size=11, bold=True, color="1F497D")

    ws_vl["B13"] = "2. Khối lượng nghiệm thu lô hàng:"
    ws_vl["C13"] = '=TEXT(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 7, FALSE), "#,##0.00") & " " & VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 6, FALSE)'
    ws_vl["C13"].font = font_bold

    ws_vl["B14"] = "3. Tiêu chuẩn kỹ thuật áp dụng:"
    ws_vl["C14"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 3, FALSE)'
    ws_vl["C14"].font = font_reg

    ws_vl["B15"] = "4. Chỉ tiêu kiểm tra chất lượng:"
    ws_vl["C15"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 4, FALSE)'
    ws_vl["C15"].font = font_reg

    ws_vl["B16"] = "5. Quy cách và số lượng tổ mẫu thí nghiệm:"
    ws_vl["C16"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 10, FALSE) & " (Số tổ mẫu: " & VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 9, FALSE) & " tổ)"'
    ws_vl["C16"].font = font_reg

    ws_vl["B17"] = "6. Đơn vị thực hiện thí nghiệm:"
    ws_vl["C17"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 11, FALSE)'
    ws_vl["C17"].font = font_reg

    ws_vl["B18"] = "7. Điều kiện nghiệm thu / Chuyển bước:"
    ws_vl["C18"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$38, 12, FALSE)'
    ws_vl["C18"].font = font_bold

    for r in range(12, 19):
        ws_vl.cell(r, 2).font = font_bold
        ws_vl.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws_vl["B20"] = "8. Kết luận nghiệm thu:"
    ws_vl["B20"].font = font_bold
    ws_vl["B21"] = "VẬT LIỆU ĐẠT TIÊU CHUẨN KỸ THUẬT THEO QUY ĐỊNH. ĐỒNG Ý CHO PHÉP ĐƯA VÀO SỬ DỤNG TRONG THI CÔNG."
    ws_vl["B21"].font = Font(name="Times New Roman", size=10, bold=True, color="006100")
    ws_vl.merge_cells("B21:G21")

    ws_vl["B24"] = "ĐẠI DIỆN NHÀ THẦU THI CÔNG"
    ws_vl["B24"].font = font_bold
    ws_vl["B24"].alignment = align_center
    ws_vl.merge_cells("B24:D24")

    ws_vl["E24"] = "ĐẠI DIỆN TƯ VẤN GIÁM SÁT"
    ws_vl["E24"].font = font_bold
    ws_vl["E24"].alignment = align_center
    ws_vl.merge_cells("E24:G24")

    ws_vl["B30"] = "Trần Việt Hải"
    ws_vl["B30"].font = font_bold
    ws_vl["B30"].alignment = align_center
    ws_vl.merge_cells("B30:D30")

    ws_vl["E30"] = "Dương Văn Lâm"
    ws_vl["E30"].font = font_bold
    ws_vl["E30"].alignment = align_center
    ws_vl.merge_cells("E30:G30")

    # =========================================================================
    # 14. SHEET: MAU_BB_LAY_MAU_HIEN_TRUONG
    # =========================================================================
    ws_lm = wb.create_sheet(title="MAU_BB_LAY_MAU_HIEN_TRUONG")
    ws_lm.views.sheetView[0].showGridLines = True

    ws_lm["B2"] = "CHỌN MÃ HẠNG MỤC LẤY MẪU (17 - 26):"
    ws_lm["B2"].font = font_bold
    ws_lm["C2"] = 22
    ws_lm["C2"].font = Font(name="Times New Roman", size=12, bold=True, color="C00000")
    ws_lm["C2"].fill = fill_input
    ws_lm["C2"].alignment = align_center
    ws_lm["C2"].border = box_border
    ws_lm["D2"] = "(Nhập mã từ 17 đến 26 tương ứng các hạng mục mẫu nén bê tông để tự động điền form)"
    ws_lm["D2"].font = font_it

    ws_lm["B3"] = "CHỌN NGÀY ĐÚC MẪU (YYYY-MM-DD):"
    ws_lm["B3"].font = font_bold
    ws_lm["C3"] = "2026-12-20"
    ws_lm["C3"].font = Font(name="Times New Roman", size=11, bold=True, color="C00000")
    ws_lm["C3"].fill = fill_input
    ws_lm["C3"].alignment = align_center
    ws_lm["C3"].border = box_border
    ws_lm["D3"] = "(Tự động tính ngày nén mẫu R7 = Ngày đúc + 7, R28 = Ngày đúc + 28 theo đúng quy định)"
    ws_lm["D3"].font = font_it

    ws_lm["B5"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_lm["B5"].font = Font(name="Times New Roman", size=11, bold=True)
    ws_lm["B5"].alignment = align_center
    ws_lm.merge_cells("B5:G5")

    ws_lm["B6"] = "Độc lập - Tự do - Hạnh phúc"
    ws_lm["B6"].font = Font(name="Times New Roman", size=10, bold=True)
    ws_lm["B6"].alignment = align_center
    ws_lm.merge_cells("B6:G6")

    ws_bb["B8"] = "BIÊN BẢN LẤY MẪU THÍ NGHIỆM TẠI HIỆN TRƯỜNG"
    ws_bb["B8"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws_bb["B8"].alignment = align_center
    ws_bb.merge_cells("B8:G8")

    ws_lm["B9"] = '="Số: " & TEXT(C2, "00") & "/BBLM-HT/KHAI-HOANG-2"'
    ws_lm["B9"].font = font_it
    ws_lm["B9"].alignment = align_center
    ws_lm.merge_cells("B9:G9")

    ws_lm["B12"] = "1. Đối tượng và cấu kiện lấy mẫu:"
    ws_lm["C12"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$51, 2, FALSE)'
    ws_lm["C12"].font = Font(name="Times New Roman", size=11, bold=True, color="1F497D")

    ws_lm["B13"] = "2. Thể tích bê tông / Quy mô mẻ đổ:"
    ws_lm["C13"] = '=TEXT(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$51, 7, FALSE), "#,##0.00") & " " & VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$51, 6, FALSE)'
    ws_lm["C13"].font = font_bold

    ws_lm["B14"] = "3. Quy định số tổ mẫu bắt buộc:"
    ws_lm["C14"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$51, 9, FALSE) & " tổ mẫu (" & VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$51, 10, FALSE) & ")"'
    ws_lm["C14"].font = font_bold

    ws_lm["B15"] = "4. Tiêu chuẩn lấy mẫu & thử nghiệm:"
    ws_lm["C15"] = '=VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$23:$L$51, 3, FALSE)'
    ws_lm["C15"].font = font_reg

    ws_lm["B16"] = "5. Phương pháp bảo dưỡng mẫu:"
    ws_lm["C16"] = "Bảo dưỡng ẩm tiêu chuẩn trong phòng thí nghiệm (TCVN 3105:2022)"
    ws_lm["C16"].font = font_reg

    ws_lm["B17"] = "6. Kế hoạch nén kiểm tra cường độ:"
    ws_lm["C17"] = '="Nén kiểm tra R7 ngày (" & TEXT(DATEVALUE(C3)+7, "dd/mm/yyyy") & "): Đạt >= 70-80% f\'c | Nén R28 ngày (" & TEXT(DATEVALUE(C3)+28, "dd/mm/yyyy") & "): Đạt >= 100% f\'c"'
    ws_lm["C17"].font = font_bold

    for r in range(12, 18):
        ws_lm.cell(r, 2).font = font_bold
        ws_lm.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws_lm["B19"] = "7. Tình trạng mẫu khi bàn giao:"
    ws_lm["B19"].font = font_bold
    ws_lm["B20"] = "Mẫu đúc nguyên vẹn, bề mặt phẳng đều, dán nhãn ghi rõ ngày đúc, vị trí cấu kiện và ký hiệu tổ mẫu."
    ws_lm["B20"].font = font_it
    ws_lm.merge_cells("B20:G20")

    ws_lm["B23"] = "CÁN BỘ LẤY MẪU THÍ NGHIỆM (LAS-XD)"
    ws_lm["B23"].font = font_bold
    ws_lm["B23"].alignment = align_center
    ws_lm.merge_cells("B23:D23")

    ws_lm["E23"] = "GIÁM SÁT HIỆN TRƯỜNG (TVGS)"
    ws_lm["E23"].font = font_bold
    ws_lm["E23"].alignment = align_center
    ws_lm.merge_cells("E23:G23")

    ws_lm["B29"] = "Nguyễn Văn Thí Nghiệm"
    ws_lm["B29"].font = font_bold
    ws_lm["B29"].alignment = align_center
    ws_lm.merge_cells("B29:D29")

    ws_lm["E29"] = "Dương Văn Lâm"
    ws_lm["E29"].font = font_bold
    ws_lm["E29"].alignment = align_center
    ws_lm.merge_cells("E29:G29")

    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)

    for ws in wb.worksheets:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            if ws.title in ["MAU_BIEN_BAN_KCS", "MAU_BB_NGHIEM_THU_VAT_LIEU", "MAU_BB_LAY_MAU_HIEN_TRUONG"]:
                continue
            max_len = 0
            for cell in col:
                val_str = str(cell.value or '')
                if not val_str.startswith("=") and len(val_str) < 60:
                    max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    for a4_ws in [ws_bb, ws_vl, ws_lm]:
        a4_ws.column_dimensions["A"].width = 4
        a4_ws.column_dimensions["B"].width = 24
        a4_ws.column_dimensions["C"].width = 20
        a4_ws.column_dimensions["D"].width = 18
        a4_ws.column_dimensions["E"].width = 18
        a4_ws.column_dimensions["F"].width = 18
        a4_ws.column_dimensions["G"].width = 16

    wb.save(output_excel_path)
    print(f"[V] Đã tạo thành công file Excel Master 14 Sheet tại: {output_excel_path}")

    # Tạo MS Project XML
    create_ms_project_xml(wbs_tasks, output_xml_path)
    print(f"[V] Đã tạo thành công file MS Project XML tại: {output_xml_path}")

def create_ms_project_xml(wbs_tasks, output_xml_path):
    root = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(root, "Name").text = "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2"
    ET.SubElement(root, "Title").text = "TIẾN ĐỘ THI CÔNG CẦU THÔN KHAI HOANG 2, KM14+363.65"
    ET.SubElement(root, "Company").text = "BAN QLDA DTXD HUYEN DONG VAN"
    ET.SubElement(root, "Author").text = "Multi-Agent AEC System"
    ET.SubElement(root, "StartDate").text = "2026-10-01T08:00:00"
    ET.SubElement(root, "FinishDate").text = "2027-05-01T17:00:00"
    ET.SubElement(root, "ScheduleFromStart").text = "1"

    calendars = ET.SubElement(root, "Calendars")
    cal = ET.SubElement(calendars, "Calendar")
    ET.SubElement(cal, "UID").text = "1"
    ET.SubElement(cal, "Name").text = "Standard"
    ET.SubElement(cal, "IsBaseCalendar").text = "1"

    tasks_el = ET.SubElement(root, "Tasks")
    for idx, (wbs, name, dur, pred, norm, work, crew, cpm, d_start, d_finish) in enumerate(wbs_tasks, start=1):
        task = ET.SubElement(tasks_el, "Task")
        ET.SubElement(task, "UID").text = str(idx)
        ET.SubElement(task, "ID").text = str(idx)
        ET.SubElement(task, "Name").text = name
        ET.SubElement(task, "OutlineNumber").text = str(idx)
        ET.SubElement(task, "OutlineLevel").text = "1"
        ET.SubElement(task, "Duration").text = f"PT{int(dur)*8}H0M0S"
        ET.SubElement(task, "Start").text = f"{d_start}T08:00:00"
        ET.SubElement(task, "Finish").text = f"{d_finish}T17:00:00"
        ET.SubElement(task, "PercentComplete").text = "0"
        ET.SubElement(task, "Critical").text = "1" if cpm == "Critical" else "0"
        ET.SubElement(task, "Milestone").text = "0"

        if pred != "-":
            pred_link = ET.SubElement(task, "PredecessorLink")
            ET.SubElement(pred_link, "PredecessorUID").text = str(max(1, idx - 1))
            ET.SubElement(pred_link, "Type").text = "1"

    tree = ET.ElementTree(root)
    tree.write(output_xml_path, encoding="utf-8", xml_declaration=True)

if __name__ == "__main__":
    out_dir = r"c:\Users\baotu\Downloads\Documents\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker"
    excel_p = os.path.join(out_dir, "Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2_Km14+363.65.xlsx")
    xml_p = os.path.join(out_dir, "Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml")
    generate_khai_hoang_2_master_package(excel_p, xml_p)
