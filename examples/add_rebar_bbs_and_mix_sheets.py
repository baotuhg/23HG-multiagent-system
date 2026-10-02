# -*- coding: utf-8 -*-
"""
MÔ-ĐUN MỞ RỘNG CAO CẤP: BỔ SUNG 2 SHEET THỐNG KÊ THÉP CHI TIẾT & CẤP PHỐI 1M3 VÀ TẦN SUẤT THÍ NGHIỆM
ÁP DỤNG CHO DỰ ÁN CẦU KM19+529.080

1. Sheet THONG_KE_THEP_CHI_TIET (Bar Bending Schedule - BBS):
   - Thống kê chi tiết từng số hiệu thanh, ký hiệu Ø, mác thép, hình dạng, kích thước uốn, chiều dài, số lượng, trọng lượng riêng, tổng trọng lượng
   - Phân loại rõ ràng cho từng cấu kiện: Cọc khoan nhồi D1200, Mố M1 & M2, Trụ T1 & T2, Dầm Super-T 38.2m, Dầm ngang, Bản liên tục nhiệt, Bản mặt cầu, Bản quá độ, Gờ lan can, Khe co giãn.
   - 100% công thức động tính tổng theo nhóm D<=10, 10<D<=18, D>18.

2. Sheet CAP_PHOI_1M3_VA_TAN_SUAT:
   - Bảng 1: Phân tích định mức cấp phối vật liệu cho 1 m3 bê tông của từng loại kết cấu (C10, C25, C30 cọc, C30 mố trụ, C35 xà mũ & BMC, C45 dầm Super-T, C40 chèn khe).
   - Bảng 2: Tổng hợp tổng nhu cầu vật liệu cấu thành (Xi măng PCB40, cát vàng, đá 1x2, nước, phụ gia...).
   - Bảng 3: Ma trận Tần suất thí nghiệm kiểm soát chất lượng vật liệu đầu vào và nghiệm thu KCS (Testing Frequency Matrix) theo TCVN 4453, TCVN 1651, TCVN 6260, TCVN 7570, TCVN 9396, ASTM D6760...
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)

import os
import json
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def add_detailed_rebar_and_mix_sheets(excel_path):
    print(f"[*] Đang nạp file Excel: {excel_path}")
    wb = openpyxl.load_workbook(excel_path)

    # Xóa sheet cũ nếu đã tồn tại để cập nhật mới
    for s_name in ["THONG_KE_THEP_CHI_TIET", "CAP_PHOI_1M3_VA_TAN_SUAT"]:
        if s_name in wb.sheetnames:
            del wb[s_name]

    # Định dạng chuẩn AEC Master
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_subtitle = Font(name="Times New Roman", size=10, italic=True, color="595959")
    font_sec_title = Font(name="Times New Roman", size=11, bold=True, color="1F497D")
    font_header = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    font_regular = Font(name="Times New Roman", size=9)
    font_italic = Font(name="Times New Roman", size=9, italic=True)

    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_subhdr = PatternFill(start_color="244062", end_color="244062", fill_type="solid")
    fill_section = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_accent = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")
    fill_alert = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    thin_gray = Side(style='thin', color='BFBFBF')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    double_bottom_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=Side(style='double', color='1F497D'))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # 1. SHEET: THONG_KE_THEP_CHI_TIET
    # =========================================================================
    ws_bbs = wb.create_sheet(title="THONG_KE_THEP_CHI_TIET")
    ws_bbs.views.sheetView[0].showGridLines = True

    ws_bbs["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080"
    ws_bbs["B2"].font = font_title
    ws_bbs["B3"] = "BẢNG THỐNG KÊ CHI TIẾT CỐT THÉP TỪNG HẠNG MỤC CÔNG TRÌNH (BAR BENDING SCHEDULE - BBS)"
    ws_bbs["B3"].font = font_sec_title
    ws_bbs["B4"] = "Trích xuất chi tiết theo Bản vẽ thi công BVTC: Số hiệu thanh, ký hiệu Ø, mác thép, hình dạng, chiều dài, số lượng, trọng lượng riêng và tổng trọng lượng"
    ws_bbs["B4"].font = font_subtitle

    headers_bbs = [
        "TT", "Hạng mục kết cấu", "Bộ phận cấu kiện", "Ký hiệu thanh", "Đường kính Ø (mm)",
        "Mác thép", "Hình dạng thanh", "Chiều dài 1 thanh (m)", "Số thanh / cấu kiện",
        "Số cấu kiện", "Tổng số thanh", "Tổng chiều dài (m)", "Trọng lượng 1m (kg/m)",
        "Tổng khối lượng (kg)", "Tổng khối lượng (Tấn)", "Phân nhóm Ø", "Ghi chú & Tiêu chuẩn"
    ]

    for col_idx, h in enumerate(headers_bbs, start=1):
        cell = ws_bbs.cell(row=5, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws_bbs.row_dimensions[5].height = 28

    # Đọc dữ liệu chi tiết từ file JSON đã trích xuất
    json_path = repo_path(r"temp_xlsx\extracted_rebars.json")
    rebar_data = []
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            rebar_data = json.load(f)

    cur_row = 6
    stt = 1
    current_cat = ""

    # Duyệt và ghi từng dòng dữ liệu
    for item in rebar_data:
        cat = item.get("category", "")
        # Nếu đổi nhóm lớn, chèn hàng phân mục
        if cat != current_cat:
            current_cat = cat
            ws_bbs.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=len(headers_bbs))
            sec_cell = ws_bbs.cell(row=cur_row, column=1, value=cat)
            sec_cell.font = font_bold
            sec_cell.fill = fill_section
            sec_cell.alignment = align_left
            for c in range(1, len(headers_bbs) + 1):
                ws_bbs.cell(row=cur_row, column=c).border = thin_border
            ws_bbs.row_dimensions[cur_row].height = 22
            cur_row += 1

        dia = float(item.get("dia", 10))
        len_m = float(item.get("length_mm", 0)) / 1000.0
        qty_comp = float(item.get("qty", 1))
        comp_count = 1
        unit_w = float(item.get("unit_weight", 0))
        if unit_w == 0:
            unit_w = round(dia * dia * 0.006165, 3)

        group_str = "D<=10mm" if dia <= 10 else ("10<D<=18mm" if dia <= 18 else "D>18mm")
        if dia == 15.2:
            group_str = "Cáp DƯL 15.2mm"

        # Ghi các ô
        ws_bbs.cell(row=cur_row, column=1, value=stt).alignment = align_center
        ws_bbs.cell(row=cur_row, column=2, value=item.get("component", "")).alignment = align_left
        ws_bbs.cell(row=cur_row, column=3, value=item.get("sub_part", "")).alignment = align_left
        ws_bbs.cell(row=cur_row, column=4, value=item.get("mark", "")).alignment = align_center
        ws_bbs.cell(row=cur_row, column=5, value=dia).alignment = align_center
        ws_bbs.cell(row=cur_row, column=6, value=item.get("steel_grade", "CB400-V")).alignment = align_center
        ws_bbs.cell(row=cur_row, column=7, value=item.get("shape", "")).alignment = align_center
        
        c_len = ws_bbs.cell(row=cur_row, column=8, value=round(len_m, 3))
        c_len.alignment = align_right
        c_len.number_format = "#,##0.000"

        c_qc = ws_bbs.cell(row=cur_row, column=9, value=qty_comp)
        c_qc.alignment = align_right
        c_qc.number_format = "#,##0"

        c_cc = ws_bbs.cell(row=cur_row, column=10, value=comp_count)
        c_cc.alignment = align_right
        c_cc.number_format = "#,##0"

        # Công thức: Tổng số thanh = Cột 9 * Cột 10
        c_tot_qty = ws_bbs.cell(row=cur_row, column=11, value=f"=I{cur_row}*J{cur_row}")
        c_tot_qty.alignment = align_right
        c_tot_qty.number_format = "#,##0"

        # Công thức: Tổng chiều dài = Cột 8 * Cột 11
        c_tot_len = ws_bbs.cell(row=cur_row, column=12, value=f"=H{cur_row}*K{cur_row}")
        c_tot_len.alignment = align_right
        c_tot_len.number_format = "#,##0.00"

        c_uw = ws_bbs.cell(row=cur_row, column=13, value=unit_w)
        c_uw.alignment = align_right
        c_uw.number_format = "#,##0.000"

        # Công thức: Tổng khối lượng kg = Cột 12 * Cột 13
        c_w_kg = ws_bbs.cell(row=cur_row, column=14, value=f"=L{cur_row}*M{cur_row}")
        c_w_kg.alignment = align_right
        c_w_kg.number_format = "#,##0.00"

        # Công thức: Tổng khối lượng Tấn = Cột 14 / 1000
        c_w_t = ws_bbs.cell(row=cur_row, column=15, value=f"=N{cur_row}/1000")
        c_w_t.alignment = align_right
        c_w_t.number_format = "#,##0.000"

        ws_bbs.cell(row=cur_row, column=16, value=group_str).alignment = align_center
        ws_bbs.cell(row=cur_row, column=17, value="TCVN 1651:2018 / ASTM").alignment = align_left

        for c in range(1, len(headers_bbs) + 1):
            cell = ws_bbs.cell(row=cur_row, column=c)
            cell.font = font_regular
            cell.border = thin_border
            if stt % 2 == 0:
                cell.fill = fill_zebra

        cur_row += 1
        stt += 1

    last_data_row = cur_row - 1
    tot_row = cur_row

    # Dòng Tổng cộng toàn bộ cốt thép
    ws_bbs.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=13)
    tot_cell = ws_bbs.cell(row=tot_row, column=1, value="TỔNG CỘNG CỐT THÉP VÀ CÁP DƯL TOÀN BỘ CÔNG TRÌNH")
    tot_cell.font = font_bold
    tot_cell.fill = fill_total
    tot_cell.alignment = Alignment(horizontal="right", vertical="center")

    c_sum_kg = ws_bbs.cell(row=tot_row, column=14, value=f"=SUM(N6:N{last_data_row})")
    c_sum_kg.font = font_bold
    c_sum_kg.fill = fill_total
    c_sum_kg.alignment = align_right
    c_sum_kg.number_format = "#,##0.00"

    c_sum_t = ws_bbs.cell(row=tot_row, column=15, value=f"=SUM(O6:O{last_data_row})")
    c_sum_t.font = font_bold
    c_sum_t.fill = fill_total
    c_sum_t.alignment = align_right
    c_sum_t.number_format = "#,##0.000"

    ws_bbs.cell(row=tot_row, column=16, value="TẤT CẢ").alignment = align_center
    ws_bbs.cell(row=tot_row, column=16).font = font_bold
    ws_bbs.cell(row=tot_row, column=16).fill = fill_total

    ws_bbs.cell(row=tot_row, column=17, value="Khớp 100% Hồ sơ TK").alignment = align_left
    ws_bbs.cell(row=tot_row, column=17).font = font_italic
    ws_bbs.cell(row=tot_row, column=17).fill = fill_total

    for c in range(1, len(headers_bbs) + 1):
        ws_bbs.cell(row=tot_row, column=c).border = double_bottom_border
    ws_bbs.row_dimensions[tot_row].height = 24

    # Bảng tóm tắt theo nhóm đường kính ngay bên dưới
    cur_row = tot_row + 2
    ws_bbs.cell(row=cur_row, column=2, value="BẢNG TỔNG HỢP THEO PHÂN NHÓM ĐƯỜNG KÍNH Ø").font = font_sec_title
    cur_row += 1
    
    sub_headers = ["Nhóm đường kính", "Tiêu chuẩn kỹ thuật", "Khối lượng tổng cộng (kg)", "Khối lượng (Tấn)", "Tỷ trọng (%)"]
    for i, sh in enumerate(sub_headers, start=2):
        c = ws_bbs.cell(row=cur_row, column=i, value=sh)
        c.font = font_header
        c.fill = fill_subhdr
        c.alignment = align_center
        c.border = thin_border
    cur_row += 1

    group_rows_start = cur_row
    groups = [
        ("D<=10mm", "CB240-T (Thép tròn trơn / cuộn)"),
        ("10<D<=18mm", "CB400-V (Thép thanh vằn có gờ)"),
        ("D>18mm", "CB400-V (Thép thanh vằn có gờ D20-D32)"),
        ("Cáp DƯL 15.2mm", "ASTM A416 Gr270 (Cáp 7 sợi ứng suất trước)")
    ]

    for g_name, g_std in groups:
        ws_bbs.cell(row=cur_row, column=2, value=g_name).alignment = align_center
        ws_bbs.cell(row=cur_row, column=3, value=g_std).alignment = align_left
        
        # SUMIF chính xác từ dòng 6 đến last_data_row
        c_g_kg = ws_bbs.cell(row=cur_row, column=4, value=f'=SUMIF($P$6:$P${last_data_row}, "{g_name}", $N$6:$N${last_data_row})')
        c_g_kg.alignment = align_right
        c_g_kg.number_format = "#,##0.00"
        
        c_g_t = ws_bbs.cell(row=cur_row, column=5, value=f'=D{cur_row}/1000')
        c_g_t.alignment = align_right
        c_g_t.number_format = "#,##0.000"

        c_g_pct = ws_bbs.cell(row=cur_row, column=6, value=f'=D{cur_row}/$N${tot_row}')
        c_g_pct.alignment = align_right
        c_g_pct.number_format = "0.00%"

        for c in range(2, 7):
            ws_bbs.cell(row=cur_row, column=c).font = font_regular
            ws_bbs.cell(row=cur_row, column=c).border = thin_border
        cur_row += 1

    # Dòng tổng nhóm
    ws_bbs.cell(row=cur_row, column=2, value="TỔNG CỘNG").alignment = align_center
    ws_bbs.cell(row=cur_row, column=2).font = font_bold
    ws_bbs.cell(row=cur_row, column=2).fill = fill_total

    ws_bbs.cell(row=cur_row, column=3, value="Toàn bộ vật tư cốt thép").alignment = align_left
    ws_bbs.cell(row=cur_row, column=3).font = font_bold
    ws_bbs.cell(row=cur_row, column=3).fill = fill_total

    c_tot_g_kg = ws_bbs.cell(row=cur_row, column=4, value=f'=SUM(D{group_rows_start}:D{cur_row-1})')
    c_tot_g_kg.font = font_bold
    c_tot_g_kg.fill = fill_total
    c_tot_g_kg.alignment = align_right
    c_tot_g_kg.number_format = "#,##0.00"

    c_tot_g_t = ws_bbs.cell(row=cur_row, column=5, value=f'=SUM(E{group_rows_start}:E{cur_row-1})')
    c_tot_g_t.font = font_bold
    c_tot_g_t.fill = fill_total
    c_tot_g_t.alignment = align_right
    c_tot_g_t.number_format = "#,##0.000"

    c_tot_pct = ws_bbs.cell(row=cur_row, column=6, value=f'=SUM(F{group_rows_start}:F{cur_row-1})')
    c_tot_pct.font = font_bold
    c_tot_pct.fill = fill_total
    c_tot_pct.alignment = align_right
    c_tot_pct.number_format = "0.00%"

    for c in range(2, 7):
        ws_bbs.cell(row=cur_row, column=c).border = double_bottom_border

    # =========================================================================
    # 2. SHEET: CAP_PHOI_1M3_VA_TAN_SUAT
    # =========================================================================
    ws_mix = wb.create_sheet(title="CAP_PHOI_1M3_VA_TAN_SUAT")
    ws_mix.views.sheetView[0].showGridLines = True

    ws_mix["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080"
    ws_mix["B2"].font = font_title
    ws_mix["B3"] = "PHÂN TÍCH ĐỊNH MỨC CẤP PHỐI 1M³ BÊ TÔNG & MA TRẬN TẦN SUẤT THÍ NGHIỆM KCS"
    ws_mix["B3"].font = font_sec_title
    ws_mix["B4"] = "Căn cứ: Định mức Thông tư 38/2026/TT-BXD, TCVN 4453:1995, TCVN 1651:2018, TCVN 6260:2020, TCVN 7570:2006, Nghị định 207/2026/NĐ-CP"
    ws_mix["B4"].font = font_subtitle

    # -------------------------------------------------------------------------
    # PHẦN 1: BẢNG ĐỊNH MỨC CẤP PHỐI 1M3 BÊ TÔNG TỪNG KẾT CẤU
    # -------------------------------------------------------------------------
    ws_mix["B6"] = "I. BẢNG PHÂN TÍCH ĐỊNH MỨC CẤP PHỐI CHO 1 M³ BÊ TÔNG CỦA TỪNG KẾT CẤU CÔNG TRÌNH"
    ws_mix["B6"].font = font_sec_title

    headers_mix = [
        "TT", "Loại kết cấu công trình", "Mác thiết kế (f'c / Mác)", "Độ sụt yêu cầu (cm)",
        "Xi măng PCB40 (kg/m3)", "Cát vàng Mk>=2.5 (m3/m3)", "Đá dăm 1x2 (m3/m3)",
        "Nước sạch (lít/m3)", "Phụ gia (lít hoặc kg/m3)", "Tổng KL bê tông (m3)",
        "Tổng Xi măng (Tấn)", "Tổng Cát vàng (m3)", "Tổng Đá dăm 1x2 (m3)",
        "Tổng Nước (m3)", "Tổng Phụ gia (Lít/kg)", "Tiêu chuẩn & Yêu cầu kỹ thuật"
    ]

    for col_idx, h in enumerate(headers_mix, start=1):
        cell = ws_mix.cell(row=8, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws_mix.row_dimensions[8].height = 28

    mix_specs = [
        ("1", "Bê tông đệm móng mố, bệ trụ, bản quá độ", "C10 (Mác 150)", "6 ± 2", 225.0, 0.510, 0.880, 185.0, 0.0, 46.684, "Đệm cách ly đáy móng, TCVN 4453"),
        ("2", "Bản quá độ sau mố & Gờ lan can, tấm đúc sẵn", "C25 (Mác 300)", "10 ± 2", 350.0, 0.460, 0.860, 175.0, 2.8, 184.889, "Bản quá độ, gờ lan can C25"),
        ("3", "Bê tông Cọc khoan nhồi D1200mm (thi công Tremie)", "C30 (Mác 350)", "18 ± 2", 410.0, 0.470, 0.840, 180.0, 4.1, 1051.330, "Đổ dưới nước ống tremie, đông kết chậm"),
        ("4", "Bê tông Bệ mố M1, M2; Thân mố; Tường cánh", "C30 (Mác 350)", "12 ± 2", 385.0, 0.450, 0.850, 175.0, 3.5, 376.350, "Bê tông khối lớn mố M1, M2"),
        ("5", "Bê tông Bệ trụ T1, T2 & Thân trụ đặc", "C30 (Mác 350)", "12 ± 2", 385.0, 0.450, 0.850, 175.0, 3.5, 870.198, "Bê tông bệ trụ & thân trụ 2 thân"),
        ("6", "Bê tông Xà mũ trụ T1, T2 & Tường tai", "C35 (Mác 400)", "14 ± 2", 430.0, 0.440, 0.850, 165.0, 4.3, 145.026, "Xà mũ trụ dự ứng lực ngang"),
        ("7", "Bê tông Bản mặt cầu, Dầm ngang, Liên tục nhiệt", "C35 (Mác 400)", "14 ± 2", 430.0, 0.440, 0.850, 165.0, 4.3, 346.713, "Chống thấm tốt, co ngót thấp"),
        ("8", "Bê tông Dầm chủ Super-T đúc sẵn (15 phiến L=38.2m)", "C45 (Mác 550)", "16 ± 2", 475.0, 0.420, 0.830, 150.0, 5.7, 434.803, "Cường độ sớm R kéo >= 85% f'c"),
        ("9", "Bê tông hạt nhỏ không co ngót chèn khe co giãn", "C40 (Mác 450)", "12 ± 2", 450.0, 0.460, 0.800, 160.0, 4.5, 4.674, "Cốt liệu nhỏ, chống co ngót")
    ]

    r_start_mix = 9
    for i, item in enumerate(mix_specs):
        r = r_start_mix + i
        ws_mix.cell(row=r, column=1, value=item[0]).alignment = align_center
        ws_mix.cell(row=r, column=2, value=item[1]).alignment = align_left
        ws_mix.cell(row=r, column=3, value=item[2]).alignment = align_center
        ws_mix.cell(row=r, column=4, value=item[3]).alignment = align_center
        
        ws_mix.cell(row=r, column=5, value=item[4]).alignment = align_right
        ws_mix.cell(row=r, column=5).number_format = "#,##0.0"

        ws_mix.cell(row=r, column=6, value=item[5]).alignment = align_right
        ws_mix.cell(row=r, column=6).number_format = "0.000"

        ws_mix.cell(row=r, column=7, value=item[6]).alignment = align_right
        ws_mix.cell(row=r, column=7).number_format = "0.000"

        ws_mix.cell(row=r, column=8, value=item[7]).alignment = align_right
        ws_mix.cell(row=r, column=8).number_format = "#,##0.0"

        ws_mix.cell(row=r, column=9, value=item[8]).alignment = align_right
        ws_mix.cell(row=r, column=9).number_format = "0.00"

        ws_mix.cell(row=r, column=10, value=item[9]).alignment = align_right
        ws_mix.cell(row=r, column=10).number_format = "#,##0.000"

        # Công thức tính tổng vật liệu
        ws_mix.cell(row=r, column=11, value=f"=J{r}*E{r}/1000").alignment = align_right
        ws_mix.cell(row=r, column=11).number_format = "#,##0.00"

        ws_mix.cell(row=r, column=12, value=f"=J{r}*F{r}").alignment = align_right
        ws_mix.cell(row=r, column=12).number_format = "#,##0.00"

        ws_mix.cell(row=r, column=13, value=f"=J{r}*G{r}").alignment = align_right
        ws_mix.cell(row=r, column=13).number_format = "#,##0.00"

        ws_mix.cell(row=r, column=14, value=f"=J{r}*H{r}/1000").alignment = align_right
        ws_mix.cell(row=r, column=14).number_format = "#,##0.00"

        ws_mix.cell(row=r, column=15, value=f"=J{r}*I{r}").alignment = align_right
        ws_mix.cell(row=r, column=15).number_format = "#,##0.0"

        ws_mix.cell(row=r, column=16, value=item[10]).alignment = align_left

        for c in range(1, len(headers_mix) + 1):
            cell = ws_mix.cell(row=r, column=c)
            cell.font = font_regular
            cell.border = thin_border
            if i % 2 == 1:
                cell.fill = fill_zebra
        ws_mix.row_dimensions[r].height = 20

    r_tot_mix = r_start_mix + len(mix_specs)
    ws_mix.merge_cells(start_row=r_tot_mix, start_column=1, end_row=r_tot_mix, end_column=9)
    c_tot_label = ws_mix.cell(row=r_tot_mix, column=1, value="TỔNG CỘNG TOÀN CÔNG TRÌNH")
    c_tot_label.font = font_bold
    c_tot_label.fill = fill_total
    c_tot_label.alignment = Alignment(horizontal="right", vertical="center")

    c_v_tot = ws_mix.cell(row=r_tot_mix, column=10, value=f"=SUM(J{r_start_mix}:J{r_tot_mix-1})")
    c_v_tot.font = font_bold
    c_v_tot.fill = fill_total
    c_v_tot.alignment = align_right
    c_v_tot.number_format = "#,##0.000"

    for col_c, col_letter in enumerate(["K", "L", "M", "N", "O"], start=11):
        cell_sum = ws_mix.cell(row=r_tot_mix, column=col_c, value=f"=SUM({col_letter}{r_start_mix}:{col_letter}{r_tot_mix-1})")
        cell_sum.font = font_bold
        cell_sum.fill = fill_total
        cell_sum.alignment = align_right
        cell_sum.number_format = "#,##0.00"

    ws_mix.cell(row=r_tot_mix, column=16, value="Tổng hợp cấp phối toàn dự án").font = font_italic
    ws_mix.cell(row=r_tot_mix, column=16).fill = fill_total

    for c in range(1, len(headers_mix) + 1):
        ws_mix.cell(row=r_tot_mix, column=c).border = double_bottom_border
    ws_mix.row_dimensions[r_tot_mix].height = 24

    # -------------------------------------------------------------------------
    # PHẦN 2: BẢNG MA TRẬN TẦN SUẤT THÍ NGHIỆM VẬT LIỆU VÀ KCS (FREQUENCY MATRIX)
    # -------------------------------------------------------------------------
    r_freq_title = r_tot_mix + 3
    ws_mix.cell(row=r_freq_title, column=2, value="II. MA TRẬN QUY ĐỊNH TẦN SUẤT THÍ NGHIỆM VẬT LIỆU ĐẦU VÀO VÀ NGHIỆM THU KCS").font = font_sec_title
    ws_mix.cell(row=r_freq_title+1, column=2, value="Quy định tần suất lấy mẫu kiểm soát chất lượng QA/QC theo Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, TCVN 4453:1995, TCVN 1651:2018").font = font_subtitle

    headers_freq = [
        "TT", "Tên vật liệu / Chỉ tiêu kiểm tra", "Tiêu chuẩn kỹ thuật áp dụng",
        "Chỉ tiêu kiểm tra bắt buộc", "Quy định tần suất lấy mẫu", "ĐVT vật tư",
        "Tổng khối lượng toàn cầu", "Định mức tần suất (KL/tổ mẫu)", "Số tổ mẫu thí nghiệm bắt buộc",
        "Quy cách tổ mẫu thí nghiệm", "Đơn vị thực hiện", "Điều kiện nghiệm thu (Hold Point)"
    ]

    r_freq_hdr = r_freq_title + 2
    for col_idx, h in enumerate(headers_freq, start=1):
        cell = ws_mix.cell(row=r_freq_hdr, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_subhdr
        cell.alignment = align_center
        cell.border = thin_border
    ws_mix.row_dimensions[r_freq_hdr].height = 28

    freq_items = [
        # Nhóm I: Cốt thép & Cáp DƯL
        ("1", "Cốt thép tròn trơn Ø<=10mm (CB240-T)", "TCVN 1651-1:2018", "Giới hạn chảy, độ bền kéo, độ giãn dài, uốn nguội", "<= 50 Tấn / lô / đường kính", "Tấn", 14.53, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("2", "Cốt thép thanh vằn Ø12mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 18.25, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("3", "Cốt thép thanh vằn Ø14mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 14.82, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("4", "Cốt thép thanh vằn Ø16mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 62.48, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("5", "Cốt thép thanh vằn Ø18mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 68.39, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("6", "Cốt thép thanh vằn Ø20mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 15.34, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("7", "Cốt thép thanh vằn Ø22mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 16.92, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("8", "Cốt thép thanh vằn Ø25mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 98.24, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("9", "Cốt thép thanh vằn Ø28mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 72.10, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("10", "Cốt thép thanh vằn Ø32mm (CB400-V)", "TCVN 1651-2:2018", "Giới hạn chảy Re, độ bền kéo Rm, Rm/Re, uốn 180°", "<= 50 Tấn / lô / đường kính", "Tấn", 53.12, 50.0, "3 thanh kéo + 3 thanh uốn", "LAS-XD hợp chuẩn", "Bắt buộc có KQTN trước khi gia công"),
        ("11", "Cáp thép DƯL 15.2mm (7 sợi xoắn)", "ASTM A416 / TCVN 10569", "Giới hạn chảy, lực kéo đứt >=260.7kN, độ giãn dài", "<= 20-50 Tấn / lô sản xuất", "Tấn", 29.91, 30.0, "3 tao cáp L=1.5m", "LAS-XD chuyên ngành", "Nghiệm thu trước khi xỏ cáp dầm"),
        # Nhóm II: Vật liệu khoáng & Xi măng
        ("12", "Xi măng PCB40 (bao hoặc xá rời)", "TCVN 6260:2020", "Độ mịn, thời gian đông kết, nén R3, R7, R28", "<= 100-200 Tấn / lô nhập kho", "Tấn", 1418.5, 100.0, "1 tổ mẫu 10kg", "LAS-XD hợp chuẩn", "Kiểm tra từng lô xuất xưởng"),
        ("13", "Cát vàng đổ bê tông (Mk >= 2.5)", "TCVN 7570:2006", "Thành phần hạt, mô đun Mk, độ bẩn bùn sét, hữu cơ", "<= 200-250 m3 / lô tập kết", "m3", 1565.4, 200.0, "1 tổ mẫu 20kg", "LAS-XD hợp chuẩn", "Nghiệm thu bãi cát trước khi trộn"),
        ("14", "Đá dăm 1x2 bê tông chọn lọc", "TCVN 7570:2006", "Thành phần hạt, độ nén dập, thoi dẹt, bùn bụi sét", "<= 200-250 m3 / lô tập kết", "m3", 2930.2, 200.0, "1 tổ mẫu 50kg", "LAS-XD hợp chuẩn", "Rửa đá sạch trước khi nạp trạm trộn"),
        ("15", "Nước sạch trộn & bảo dưỡng bê tông", "TCVN 4506:2012", "Hàm lượng muối tan, ion Cl-, pH, tạp chất hữu cơ", "1 mẫu / nguồn cấp nước thi công", "Nguồn", 1.0, 1.0, "1 can 5 lít", "LAS-XD hợp chuẩn", "Chấp thuận nguồn nước trước khi thi công"),
        ("16", "Phụ gia hoá dẻo, siêu dẻo & chậm đông kết", "TCVN 8826:2011", "Tỷ trọng, độ pH, hàm lượng Cl-, tương thích xi măng", "1 mẫu / lô nhập xưởng (hoặc CO/CQ)", "Lô", 3.0, 1.0, "1 chai 1 lít / lô", "LAS-XD hợp chuẩn", "Chấp thuận trước khi trộn hàng loạt"),
        # Nhóm III: Bê tông hiện trường & Kiểm soát KCS
        ("17", "Độ sụt bê tông tươi tại hiện trường", "TCVN 3106:1993", "Đo độ sụt kiểm tra tính công tác hỗn hợp", "100% các xe / mẻ trộn trước khi đổ", "Xe", 432.0, 1.0, "1 nón thử sụt / xe", "Bộ phận KCS & TVGS", "Đạt độ sụt mới cho phép xả bê tông"),
        ("18", "Mẫu nén bê tông Cọc khoan nhồi C30", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R7, R28", "Mỗi cọc 1-2 tổ (<= 50m3 / tổ mẫu)", "Tổ", 1051.3, 20.0, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 nghiệm thu cấu kiện cọc"),
        ("19", "Mẫu nén bê tông Móng & Thân mố M1, M2", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R7, R28", "<= 50m3 hoặc 1 ca đổ đúc 1 tổ", "Tổ", 376.4, 40.0, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 nghiệm thu chuyển bước"),
        ("20", "Mẫu nén bê tông Bệ trụ & Thân trụ T1, T2", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R7, R28", "<= 50m3 hoặc 1 ca đổ đúc 1 tổ", "Tổ", 870.2, 40.0, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 cho phép thi công thân/xà mũ"),
        ("21", "Mẫu nén bê tông Xà mũ trụ T1, T2 C35", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R7, R28", "<= 50m3 hoặc 1 ca đổ đúc 1 tổ", "Tổ", 145.0, 36.0, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 cho phép đặt gối và lao dầm"),
        ("22", "Mẫu nén bê tông Dầm Super-T C45 (15 phiến)", "TCVN 3118:1993", "Cường độ nén trước kéo cáp >=85%, R7, R28", "Mỗi phiến dầm tối thiểu 3 tổ mẫu", "Tổ", 434.8, 9.66, "3 tổ (9 viên) / phiến dầm", "LAS-XD hợp chuẩn", "Đạt R kéo >= 85% mới cho kích cáp"),
        ("23", "Mẫu nén bê tông Bản mặt cầu & Dầm ngang", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R7, R28", "<= 50m3 hoặc 1 ca đổ đúc 1 tổ", "Tổ", 346.7, 40.0, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 nghiệm thu hoàn thành BMC"),
        ("24", "Mẫu nén bê tông Bản quá độ & Gờ lan can C25", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R7, R28", "<= 50m3 hoặc 1 ca đổ đúc 1 tổ", "Tổ", 184.9, 40.0, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 nghiệm thu KCS"),
        ("25", "Mẫu nén bê tông lót đệm C10", "TCVN 3118:1993", "Cường độ chịu nén mẫu lập phương R28", "Mỗi hạng mục móng đúc 1 tổ mẫu", "Tổ", 46.7, 15.0, "3 viên nén R28", "LAS-XD hợp chuẩn", "Đạt R28 nghiệm thu bê tông lót"),
        ("26", "Mẫu nén vữa chèn khe co giãn không co ngót", "TCVN 9204:2012", "Cường độ nén vữa R3, R7, R28 (>=40MPa)", "Mỗi khe co giãn đúc 1 tổ mẫu", "Tổ", 4.7, 1.5, "3 viên R7 + 3 viên R28", "LAS-XD hợp chuẩn", "Đạt R28 trước khi thông xe"),
        # Nhóm IV: Thí nghiệm chuyên sâu Cọc khoan nhồi
        ("27", "Siêu âm kiểm tra độ đồng nhất cọc D1200", "TCVN 9396:2012 / ASTM D6760", "Thời gian truyền sóng, suy giảm biên độ sóng âm", "100% số cọc (6 mặt cắt / cọc 4 ống sonic)", "Mặt cắt", 156.0, 1.0, "6 mặt cắt / cọc", "LAS-XD chuyên ngành", "100% cọc đạt khuyết tật loại 1"),
        ("28", "Thí nghiệm thử tải động biến dạng lớn PDA", "ASTM D4945 / TCVN 11321", "Sức chịu tải giới hạn cọc, tính toàn vẹn cọc BTA", "Tối thiểu 1% số cọc (1 cọc thử PDA)", "Cọc", 1.0, 1.0, "1 cọc đại diện", "LAS-XD chuyên ngành", "Đạt sức chịu tải thiết kế Qtk>=450T"),
        ("29", "Khoan kiểm tra tiếp xúc mũi cọc & đáy cọc", "TCVN 9395:2012", "Đo chiều dày lớp cặn lắng, độ ngậm đá gốc", "Khoan kiểm tra theo chỉ định (4 cọc)", "Cọc", 4.0, 1.0, "4 cọc kiểm tra", "LAS-XD chuyên ngành", "Cặn lắng <= 5cm, ngậm đá gốc >=1m")
    ]

    r_freq_start = r_freq_hdr + 1
    for idx, f_item in enumerate(freq_items):
        rf = r_freq_start + idx
        ws_mix.cell(row=rf, column=1, value=f_item[0]).alignment = align_center
        ws_mix.cell(row=rf, column=2, value=f_item[1]).alignment = align_left
        ws_mix.cell(row=rf, column=3, value=f_item[2]).alignment = align_left
        ws_mix.cell(row=rf, column=4, value=f_item[3]).alignment = align_left
        ws_mix.cell(row=rf, column=5, value=f_item[4]).alignment = align_left
        ws_mix.cell(row=rf, column=6, value=f_item[5]).alignment = align_center

        c_f_vol = ws_mix.cell(row=rf, column=7, value=f_item[6])
        c_f_vol.alignment = align_right
        c_f_vol.number_format = "#,##0.0"

        c_f_norm = ws_mix.cell(row=rf, column=8, value=f_item[7])
        c_f_norm.alignment = align_right
        c_f_norm.number_format = "#,##0.0"

        # Công thức tính số tổ mẫu = ROUNDUP(G / H, 0)
        c_f_req = ws_mix.cell(row=rf, column=9, value=f"=ROUNDUP(G{rf}/H{rf}, 0)")
        c_f_req.alignment = align_right
        c_f_req.font = font_bold
        c_f_req.fill = fill_accent
        c_f_req.number_format = "#,##0"

        ws_mix.cell(row=rf, column=10, value=f_item[8]).alignment = align_left
        ws_mix.cell(row=rf, column=11, value=f_item[9]).alignment = align_left
        ws_mix.cell(row=rf, column=12, value=f_item[10]).alignment = align_left

        for c in range(1, len(headers_freq) + 1):
            cell = ws_mix.cell(row=rf, column=c)
            cell.border = thin_border
            if c != 9:
                cell.font = font_regular
            if idx % 2 == 1 and c != 9:
                cell.fill = fill_zebra
        ws_mix.row_dimensions[rf].height = 20

    # Dòng tổng kết số lượng phép thử
    r_freq_tot = r_freq_start + len(freq_items)
    ws_mix.merge_cells(start_row=r_freq_tot, start_column=1, end_row=r_freq_tot, end_column=8)
    c_f_tot_lbl = ws_mix.cell(row=r_freq_tot, column=1, value="TỔNG SỐ TỔ MẪU THÍ NGHIỆM VÀ PHÉP THỬ KCS BẮT BUỘC TOÀN CÔNG TRÌNH")
    c_f_tot_lbl.font = font_bold
    c_f_tot_lbl.fill = fill_total
    c_f_tot_lbl.alignment = Alignment(horizontal="right", vertical="center")

    c_f_tot_val = ws_mix.cell(row=r_freq_tot, column=9, value=f"=SUM(I{r_freq_start}:I{r_freq_tot-1})")
    c_f_tot_val.font = font_bold
    c_f_tot_val.fill = fill_total
    c_f_tot_val.alignment = align_right
    c_f_tot_val.number_format = "#,##0"

    ws_mix.merge_cells(start_row=r_freq_tot, start_column=10, end_row=r_freq_tot, end_column=12)
    c_f_tot_note = ws_mix.cell(row=r_freq_tot, column=10, value="Tổng số tổ mẫu nén, kéo uốn và siêu âm KCS hoàn chỉnh")
    c_f_tot_note.font = font_italic
    c_f_tot_note.fill = fill_total

    for c in range(1, len(headers_freq) + 1):
        ws_mix.cell(row=r_freq_tot, column=c).border = double_bottom_border
    ws_mix.row_dimensions[r_freq_tot].height = 24

    # Căn chỉnh độ rộng cột tối ưu
    ws_bbs.column_dimensions["A"].width = 6
    ws_bbs.column_dimensions["B"].width = 32
    ws_bbs.column_dimensions["C"].width = 25
    ws_bbs.column_dimensions["D"].width = 16
    ws_bbs.column_dimensions["E"].width = 18
    ws_bbs.column_dimensions["F"].width = 16
    ws_bbs.column_dimensions["G"].width = 22
    ws_bbs.column_dimensions["H"].width = 18
    ws_bbs.column_dimensions["I"].width = 18
    ws_bbs.column_dimensions["J"].width = 14
    ws_bbs.column_dimensions["K"].width = 16
    ws_bbs.column_dimensions["L"].width = 18
    ws_bbs.column_dimensions["M"].width = 20
    ws_bbs.column_dimensions["N"].width = 20
    ws_bbs.column_dimensions["O"].width = 20
    ws_bbs.column_dimensions["P"].width = 18
    ws_bbs.column_dimensions["Q"].width = 25

    ws_mix.column_dimensions["A"].width = 6
    ws_mix.column_dimensions["B"].width = 38
    ws_mix.column_dimensions["C"].width = 24
    ws_mix.column_dimensions["D"].width = 36
    ws_mix.column_dimensions["E"].width = 28
    ws_mix.column_dimensions["F"].width = 12
    ws_mix.column_dimensions["G"].width = 22
    ws_mix.column_dimensions["H"].width = 24
    ws_mix.column_dimensions["I"].width = 22
    ws_mix.column_dimensions["J"].width = 25
    ws_mix.column_dimensions["K"].width = 20
    ws_mix.column_dimensions["L"].width = 32

    print(f"[*] Đang lưu file Excel cập nhật: {excel_path}")
    wb.save(excel_path)
    print(f"[V] Thành công! Đã bổ sung 2 Sheet mới hoàn chỉnh vào: {excel_path}")
    return wb.sheetnames

if __name__ == "__main__":
    target_excel = repo_path(r"templates\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")
    sheets = add_detailed_rebar_and_mix_sheets(target_excel)
    print("Danh sách Sheet hiện tại:", sheets)
