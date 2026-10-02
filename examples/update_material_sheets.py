# -*- coding: utf-8 -*-
"""
MÔ-ĐUN MỞ RỘNG: BỔ SUNG 2 SHEET PHÂN TÍCH ĐỊNH MỨC & TỔNG HỢP VẬT TƯ (BOM)
VÀO FILE EXCEL MASTER CẦU KM19+529.080

1. Sheet PHAN_TICH_VAT_TU_WBS:
   - Phân tích chi tiết vật liệu cấu thành từng hạng mục WBS theo Thông tư 38/2026/TT-BXD
   - Chi tiết sắt thép từng loại Ø (Ø10, Ø12, Ø14, Ø16, Ø18, Ø20, Ø22, Ø25, Ø28, Ø32), cáp DƯL, xi măng, cát, đá, nước, phụ gia...
   - 100% công thức động trỏ sang Sheet QS_DIEN_GIAI_CHI_TIET

2. Sheet TONG_HOP_VAT_TU_TOAN_BO:
   - Tổng hợp toàn bộ nhu cầu vật liệu theo 4 nhóm lớn (Sắt thép, Bê tông & Khoáng, Phụ gia hóa chất, Phụ kiện hoàn thiện)
   - Sử dụng hàm =SUMIF trỏ sang Sheet Phân tích vật tư
   - Tính khối lượng cung ứng có hao hụt thi công theo TT 38/2026
   - Phân bổ kế hoạch cung ứng theo 4 giai đoạn thi công (Procurement Schedule)
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def add_material_sheets(excel_path):
    print(f"[*] Đang nạp file Excel: {excel_path}")
    wb = openpyxl.load_workbook(excel_path)

    # Xóa sheet cũ nếu đã tồn tại để tạo mới tinh gọn
    for s_name in ["PHAN_TICH_VAT_TU_WBS", "TONG_HOP_VAT_TU_TOAN_BO"]:
        if s_name in wb.sheetnames:
            del wb[s_name]

    # Định dạng chuẩn AEC Master
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_subtitle = Font(name="Times New Roman", size=10, italic=True, color="595959")
    font_section = Font(name="Times New Roman", size=10, bold=True, color="1F497D")
    font_header = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    font_regular = Font(name="Times New Roman", size=9)
    font_italic = Font(name="Times New Roman", size=9, italic=True)

    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_subhdr = PatternFill(start_color="244062", end_color="244062", fill_type="solid")
    fill_section = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_procure = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")

    thin_gray = Side(style='thin', color='BFBFBF')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    double_bottom_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=Side(style='double', color='1F497D'))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # 1. SHEET: PHAN_TICH_VAT_TU_WBS
    # =========================================================================
    ws_wbs = wb.create_sheet(title="PHAN_TICH_VAT_TU_WBS")
    ws_wbs.views.sheetView[0].showGridLines = True

    ws_wbs["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080"
    ws_wbs["B2"].font = font_title
    ws_wbs["B3"] = "BẢNG PHÂN TÍCH ĐỊNH MỨC HAO PHÍ VẬT LIỆU CHO TỪNG HẠNG MỤC CÔNG TÁC WBS"
    ws_wbs["B3"].font = font_section
    ws_wbs["B4"] = "Chi tiết sắt thép từng loại đường kính Ø, xi măng, cát, đá, cáp DƯL (Theo Định mức Thông tư 38/2026/TT-BXD)"
    ws_wbs["B4"].font = font_subtitle

    headers_wbs = [
        "TT", "Mã định mức", "Hạng mục công tác WBS", "Khối lượng công tác", "ĐVT công tác",
        "Tên quy cách vật tư cấu thành", "ĐVT vật tư", "Định mức cho 1 ĐVT", "Hao phí vật tư hạng mục", "Tiêu chuẩn kỹ thuật áp dụng"
    ]
    for col_idx, h in enumerate(headers_wbs, start=1):
        cell = ws_wbs.cell(row=5, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws_wbs.row_dimensions[5].height = 28

    # Danh mục hạng mục công tác và định mức vật liệu cấu thành
    # Mỗi item: (wbs_code, wbs_name, qs_cell, wbs_unit, materials_list)
    # materials_list: [(mat_name, mat_unit, norm_rate, standard)]
    wbs_items = [
        # 1. Cọc khoan nhồi D1200
        ("AC.11111", "Khoan tạo lỗ cọc khoan nhồi D1200mm (26 cọc L=20m-40m)", "QS_DIEN_GIAI_CHI_TIET!J7", "m", [
            ("Bột sét Bentonite tạo vách khoan", "Tấn", 0.045, "TCVN 9395:2012"),
            ("Nước sạch thi công khoan cọc", "m3", 0.250, "TCVN 4506:2012")
        ]),
        ("AC.12111", "Hạ và nhổ ống vách thép dẫn hướng D1300mm d=8mm", "QS_DIEN_GIAI_CHI_TIET!J12", "m", [
            ("Ống vách thép dẫn hướng D1300 d=8mm (khấu hao)", "Tấn", 0.052, "TCVN 9395:2012")
        ]),
        ("AF.21111", "Bê tông cọc khoan nhồi C30/37 đổ bằng ống Tremie", "QS_DIEN_GIAI_CHI_TIET!J17", "m3", [
            ("Xi măng PCB40 cọc khoan nhồi", "kg", 395.0, "TCVN 6260:2020"),
            ("Cát vàng hạt trung Mk>=2.0", "m3", 0.470, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch chọn lọc", "m3", 0.830, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 185.0, "TCVN 4506:2012"),
            ("Phụ gia hoá dẻo chậm đông kết", "kg", 4.20, "ASTM C494 Type G")
        ]),
        ("AF.61111", "Cốt thép cọc khoan nhồi D1200mm (Lồng thép)", "QS_DIEN_GIAI_CHI_TIET!J22", "Tấn", [
            ("Thép chủ cọc nhồi vằn Ø25 (CB500-V)", "Tấn", 0.785, "TCVN 1651:2018"),
            ("Thép đai tăng cường vằn Ø16 (CB400-V)", "Tấn", 0.120, "TCVN 1651:2018"),
            ("Thép đai xoắn liên tục tròn trơn Ø10 (CB240-T)", "Tấn", 0.095, "TCVN 1651:2018"),
            ("Ống siêu âm cọc nhồi thép Ø60x2", "m", 36.50, "ASTM A53 / TCVN 9395"),
            ("Dây thép buộc 1 ly", "kg", 15.0, "TCVN 1651:2018"),
            ("Que hàn điện E42/E50", "kg", 8.0, "TCVN 3223:2000")
        ]),
        # 2. Móng mố trụ
        ("AF.11111", "Bê tông lót đáy bệ móng mác C10 dày 10cm", "QS_DIEN_GIAI_CHI_TIET!J34", "m3", [
            ("Xi măng PCB40 lót móng", "kg", 240.0, "TCVN 6260:2020"),
            ("Cát vàng xây dựng", "m3", 0.510, "TCVN 7570:2006"),
            ("Đá dăm 1x2", "m3", 0.870, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 175.0, "TCVN 4506:2012")
        ]),
        ("AF.12111", "Bê tông bệ móng mố M1, M2 và bệ trụ T1, T2 mác C30", "QS_DIEN_GIAI_CHI_TIET!J39", "m3", [
            ("Xi măng PCB40 bệ móng", "kg", 385.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.460, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.840, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 3.80, "ASTM C494 Type F")
        ]),
        ("AF.12211", "Bê tông thân mố chân dê M1 và mố chữ U M2 mác C30", "QS_DIEN_GIAI_CHI_TIET!J44", "m3", [
            ("Xi măng PCB40 thân mố", "kg", 385.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.460, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.840, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 3.80, "ASTM C494 Type F")
        ]),
        ("AF.12311", "Bê tông tường cánh và tường đỉnh mố M1, M2 mác C30", "QS_DIEN_GIAI_CHI_TIET!J47", "m3", [
            ("Xi măng PCB40 tường cánh", "kg", 385.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.460, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.840, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 3.80, "ASTM C494 Type F")
        ]),
        ("AF.13111", "Bê tông thân đặc trụ T1 và T2 mác C30", "QS_DIEN_GIAI_CHI_TIET!J50", "m3", [
            ("Xi măng PCB40 thân trụ", "kg", 385.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.460, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.840, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 3.80, "ASTM C494 Type F")
        ]),
        ("AF.13211", "Bê tông xà mũ trụ T1, T2 mác C35", "QS_DIEN_GIAI_CHI_TIET!J53", "m3", [
            ("Xi măng PCB40 xà mũ trụ", "kg", 425.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.450, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch chọn lọc", "m3", 0.830, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 4.50, "ASTM C494 Type F")
        ]),
        ("AF.62111", "Cốt thép mố và trụ (Bệ móng, Thân mố, Thân trụ, Xà mũ)", "QS_DIEN_GIAI_CHI_TIET!J56", "Tấn", [
            ("Thép lưới đáy bệ móng vằn Ø32 (CB500-V)", "Tấn", 0.365, "TCVN 1651:2018"),
            ("Thép đứng thân trụ vằn Ø32 (CB500-V)", "Tấn", 0.245, "TCVN 1651:2018"),
            ("Thép đứng thân mố vằn Ø28 (CB500-V)", "Tấn", 0.145, "TCVN 1651:2018"),
            ("Thép xà mũ chịu uốn vằn Ø28 (CB500-V)", "Tấn", 0.085, "TCVN 1651:2018"),
            ("Thép phân bố bệ móng vằn Ø20 (CB500-V)", "Tấn", 0.100, "TCVN 1651:2018"),
            ("Thép đai mố trụ vằn Ø16 (CB400-V)", "Tấn", 0.040, "TCVN 1651:2018"),
            ("Thép đai cấu tạo tròn trơn Ø10 (CB240-T)", "Tấn", 0.020, "TCVN 1651:2018"),
            ("Dây thép buộc 1 ly", "kg", 15.0, "TCVN 1651:2018"),
            ("Que hàn điện E42/E50", "kg", 8.0, "TCVN 3223:2000")
        ]),
        # 3. Kết cấu nhịp dầm Super-T & Bản mặt cầu
        ("AF.31111", "Bê tông dầm chủ Super-T L=38.2m mác C45 (15 phiến)", "QS_DIEN_GIAI_CHI_TIET!J62", "m3", [
            ("Xi măng PCB40 dầm Super-T", "kg", 485.0, "TCVN 6260:2020"),
            ("Cát vàng chọn lọc thô Mk>=2.6", "m3", 0.440, "TCVN 7570:2006"),
            ("Đá dăm 1x2 nghiền rửa chọn lọc", "m3", 0.830, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 165.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo Polycarboxylate thế hệ 3", "kg", 6.80, "ASTM C494 Type F/G")
        ]),
        ("AF.63111", "Cáp dự ứng lực tao xoắn 15.2mm Grade 270 dầm Super-T", "QS_DIEN_GIAI_CHI_TIET!J64", "Tấn", [
            ("Cáp dự ứng lực tao 15.2mm Gr270", "Tấn", 1.000, "ASTM A416 / TCVN 11823"),
            ("Bộ neo chùm dự ứng lực (12 bộ/phiến dầm)", "Bộ", 3.75, "TCVN 11823:2017"),
            ("Ống ghen tôn mạ kẽm luồn cáp DƯL", "m", 36.0, "TCVN 11823:2017"),
            ("Vữa không co ngót bơm lấp ống ghen SikaGrout", "kg", 380.0, "ASTM C1107")
        ]),
        ("AF.64111", "Cốt thép thường dầm chủ Super-T (Thép sườn, cánh)", "QS_DIEN_GIAI_CHI_TIET!J66", "Tấn", [
            ("Thép sườn dầm Super-T vằn Ø16 (CB400-V)", "Tấn", 0.550, "TCVN 1651:2018"),
            ("Thép bản cánh dầm Super-T vằn Ø14 (CB400-V)", "Tấn", 0.350, "TCVN 1651:2018"),
            ("Thép đai và móc cẩu dầm tròn trơn Ø10 (CB240-T)", "Tấn", 0.100, "TCVN 1651:2018"),
            ("Dây thép buộc 1 ly", "kg", 15.0, "TCVN 1651:2018")
        ]),
        ("AF.32111", "Gối chậu cao su di động đơn hướng & cố định", "QS_DIEN_GIAI_CHI_TIET!J70", "Cái", [
            ("Gối chậu cao su cốt bản thép chịu lực", "Cái", 1.00, "TCVN 11823:2017 / AASHTO")
        ]),
        ("AF.33111", "Bê tông dầm ngang mố và trụ mác C35", "QS_DIEN_GIAI_CHI_TIET!J73", "m3", [
            ("Xi măng PCB40 dầm ngang", "kg", 420.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.450, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.830, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 4.50, "ASTM C494 Type F")
        ]),
        ("AF.35111", "Bê tông bản mặt cầu đổ tại chỗ dày 20cm mác C35", "QS_DIEN_GIAI_CHI_TIET!J77", "m3", [
            ("Xi măng PCB40 bản mặt cầu", "kg", 420.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.450, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.830, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 4.50, "ASTM C494 Type F")
        ]),
        ("AF.36111", "Bê tông mối nối liên tục nhiệt đỉnh trụ T1, T2 mác C35", "QS_DIEN_GIAI_CHI_TIET!J79", "m3", [
            ("Xi măng PCB40 liên tục nhiệt", "kg", 420.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.450, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.830, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Phụ gia siêu dẻo giảm nước", "kg", 4.50, "ASTM C494 Type F")
        ]),
        ("AF.65111", "Cốt thép bản mặt cầu và mối nối liên tục nhiệt", "QS_DIEN_GIAI_CHI_TIET!J81", "Tấn", [
            ("Thép bản mặt cầu chịu lực lưới dưới vằn Ø16 (CB400-V)", "Tấn", 0.480, "TCVN 1651:2018"),
            ("Thép bản mặt cầu lưới trên vằn Ø14 (CB400-V)", "Tấn", 0.420, "TCVN 1651:2018"),
            ("Thép đai chống nứt tròn trơn Ø10 (CB240-T)", "Tấn", 0.100, "TCVN 1651:2018"),
            ("Dây thép buộc 1 ly", "kg", 15.0, "TCVN 1651:2018")
        ]),
        ("AC.31111", "Khe co giãn răng lược thép D=100mm", "QS_DIEN_GIAI_CHI_TIET!J83", "m", [
            ("Thép tấm răng lược hợp kim mạ kẽm D=100mm", "m", 1.000, "TCVN 11823:2017"),
            ("Máng thoát nước cao su đàn hồi", "m", 1.050, "TCVN 11823:2017"),
            ("Bu-lông neo M24 cường độ cao", "Cái", 6.0, "ASTM A325")
        ]),
        # 4. Hoàn thiện, Bản quá độ, Lan can & Thoát nước
        ("AF.41111", "Bê tông bản quá độ sau mố M1, M2 mác C30", "QS_DIEN_GIAI_CHI_TIET!J86", "m3", [
            ("Xi măng PCB40 bản quá độ", "kg", 385.0, "TCVN 6260:2020"),
            ("Cát vàng sàng rửa Mk>=2.0", "m3", 0.460, "TCVN 7570:2006"),
            ("Đá dăm 1x2 sạch", "m3", 0.840, "TCVN 7570:2006"),
            ("Nước sạch trộn bê tông", "Lít", 180.0, "TCVN 4506:2012"),
            ("Thép bản quá độ dọc vằn Ø16 (CB400-V)", "kg", 110.0, "TCVN 1651:2018"),
            ("Thép bản quá độ phân bố vằn Ø12 (CB400-V)", "kg", 25.0, "TCVN 1651:2018")
        ]),
        ("AB.21111", "Đắp vật liệu dạng hạt chọn lọc sau mố K98", "QS_DIEN_GIAI_CHI_TIET!J89", "m3", [
            ("Đất đắp nền đường sau mố K98", "m3", 1.150, "TCVN 9436:2012"),
            ("Vải địa kỹ thuật không dệt lọc ngược", "m2", 0.450, "TCVN 9865:2013")
        ]),
        ("AF.42111", "Bê tông gờ lan can mác C25 và lắp tay vịn thép", "QS_DIEN_GIAI_CHI_TIET!J91", "m", [
            ("Xi măng PCB40 gờ lan can", "kg", 33.25, "TCVN 6260:2020"), # 0.095 m3 x 350
            ("Cát vàng", "m3", 0.046, "TCVN 7570:2006"),
            ("Đá dăm 1x2", "m3", 0.081, "TCVN 7570:2006"),
            ("Thép cốt gờ lan can vằn Ø12 (CB400-V)", "kg", 12.50, "TCVN 1651:2018"),
            ("Tay vịn lan can thép mạ kẽm nhúng nóng Ø114", "m", 1.050, "TCVN 11823:2017")
        ]),
        ("AC.41111", "Hệ thống thoát nước mặt cầu (24 cụm hố thu & ống dẫn)", "QS_DIEN_GIAI_CHI_TIET!J93", "Bộ", [
            ("Phễu gang thu nước mặt cầu có lưới chắn", "Bộ", 1.000, "TCVN 11823:2017"),
            ("Ống thoát nước nhựa uPVC Ø110 dày 3.2mm", "m", 3.50, "TCVN 8491:2011")
        ]),
        ("AD.11111", "Lớp phun chống thấm & thảm bê tông nhựa chặt C12.5 dày 7cm", "QS_DIEN_GIAI_CHI_TIET!J95", "m2", [
            ("Màng chống thấm mặt cầu Polymer cải tiến", "m2", 1.150, "TCVN 11823:2017"),
            ("Bê tông nhựa chặt C12.5 rải nóng", "Tấn", 0.175, "TCVN 8819:2011"),
            ("Nhựa dính bám Tack coat", "kg", 0.50, "TCVN 8819:2011")
        ]),
        ("AL.11111", "Gia cố mái taluy tứ nón mố bằng đá hộc xây VXM M100", "QS_DIEN_GIAI_CHI_TIET!J97", "m3", [
            ("Đá hộc 20x30cm kè tứ nón", "m3", 1.050, "TCVN 11823:2017"),
            ("Xi măng PCB40 vữa xây M100", "kg", 119.70, "TCVN 6260:2020"), # 0.38 m3 vữa x 315 kg
            ("Cát vàng vữa xây M100", "m3", 0.400, "TCVN 7570:2006")
        ])
    ]

    curr_row = 6
    stt_cnt = 1
    for item in wbs_items:
        w_code, w_name, qs_link, w_unit, mats = item
        
        # Dòng chính hạng mục WBS
        ws_wbs.cell(row=curr_row, column=1, value=stt_cnt).alignment = align_center
        ws_wbs.cell(row=curr_row, column=2, value=w_code).alignment = align_center
        ws_wbs.cell(row=curr_row, column=3, value=w_name).alignment = align_left
        
        cell_d = ws_wbs.cell(row=curr_row, column=4, value=f"={qs_link}")
        cell_d.alignment = align_right
        cell_d.number_format = "#,##0.00"
        
        ws_wbs.cell(row=curr_row, column=5, value=w_unit).alignment = align_center

        for c in range(1, 11):
            cell = ws_wbs.cell(row=curr_row, column=c)
            cell.font = font_bold
            cell.fill = fill_section
            cell.border = thin_border
        
        parent_row = curr_row
        curr_row += 1

        # Các dòng vật tư cấu thành con
        for m in mats:
            m_name, m_unit, m_norm, m_std = m
            ws_wbs.cell(row=curr_row, column=3, value=f"   • {m_name}").alignment = align_left
            ws_wbs.cell(row=curr_row, column=6, value=m_name).alignment = align_left
            ws_wbs.cell(row=curr_row, column=7, value=m_unit).alignment = align_center
            
            c_norm = ws_wbs.cell(row=curr_row, column=8, value=m_norm)
            c_norm.alignment = align_right
            c_norm.number_format = "#,##0.000" if isinstance(m_norm, float) and m_norm < 1 else "#,##0.00"
            
            # CÔNG THỨC ĐỘNG: Khối lượng công tác (Cột D của parent_row) * Định mức (Cột H)
            c_qty = ws_wbs.cell(row=curr_row, column=9, value=f"=D{parent_row}*H{curr_row}")
            c_qty.alignment = align_right
            c_qty.number_format = "#,##0.00"
            
            ws_wbs.cell(row=curr_row, column=10, value=m_std).alignment = align_left

            for c in range(1, 11):
                cell = ws_wbs.cell(row=curr_row, column=c)
                if not cell.font:
                    cell.font = font_regular
                cell.border = thin_border
            curr_row += 1

        stt_cnt += 1

    max_wbs_row = curr_row - 1

    # Căn chỉnh độ rộng cột Sheet PHAN_TICH_VAT_TU_WBS
    widths_wbs = [6, 12, 45, 18, 12, 45, 12, 16, 20, 24]
    for i, w in enumerate(widths_wbs, start=1):
        ws_wbs.column_dimensions[get_column_letter(i)].width = w
    ws_wbs.freeze_panes = "D6"

    # =========================================================================
    # 2. SHEET: TONG_HOP_VAT_TU_TOAN_BO (CONSOLIDATED BOM & PROCUREMENT)
    # =========================================================================
    ws_bom = wb.create_sheet(title="TONG_HOP_VAT_TU_TOAN_BO")
    ws_bom.views.sheetView[0].showGridLines = True

    ws_bom["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080"
    ws_bom["B2"].font = font_title
    ws_bom["B3"] = "BẢNG TỔNG HỢP NHU CẦU VẬT TƯ TOÀN CÔNG TRÌNH & KẾ HOẠCH CUNG ỨNG (BOM)"
    ws_bom["B3"].font = font_section
    ws_bom["B4"] = "Tổng hợp tự động 100% công thức (=SUMIF) từ bảng phân rã WBS; Kế hoạch phân bổ theo 4 giai đoạn thi công"
    ws_bom["B4"].font = font_subtitle

    headers_bom = [
        "TT", "Tên chủng loại vật tư", "Tiêu chuẩn kỹ thuật", "ĐVT",
        "Hao phí định mức lý thuyết", "Tỷ lệ hao hụt TT 38/2026 (%)", "Tổng nhu cầu cung ứng (BOM)",
        "GĐ 1: Cọc nhồi", "GĐ 2: Mố & Trụ", "GĐ 3: Dầm & Mặt cầu", "GĐ 4: Hoàn thiện"
    ]
    for col_idx, h in enumerate(headers_bom, start=1):
        cell = ws_bom.cell(row=5, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws_bom.row_dimensions[5].height = 28

    # Danh mục tổng hợp gom vật tư:
    # (Nhóm, [(STT, Mat_Name, Mat_Search_Pattern, Unit, Std, Scrap_Rate, p1, p2, p3, p4)])
    # p1..p4 là tỷ trọng phân bổ theo giai đoạn
    categories = [
        ("I. SẮT THÉP CÁC LOẠI (CỐT THÉP, CÁP DỰ ỨNG LỰC & PHỤ KIỆN)", [
            ("1", "Thép đai tròn trơn Ø10 (CB240-T)", "*Ø10*", "Tấn", "TCVN 1651:2018", 0.015, 0.40, 0.20, 0.40, 0.0),
            ("2", "Thép cốt vằn Ø12 (CB400-V)", "*Ø12*", "Tấn", "TCVN 1651:2018", 0.015, 0.0, 0.0, 0.0, 1.0),
            ("3", "Thép bản cánh & mặt cầu vằn Ø14 (CB400-V)", "*Ø14*", "Tấn", "TCVN 1651:2018", 0.015, 0.0, 0.0, 1.0, 0.0),
            ("4", "Thép đai & sườn dầm vằn Ø16 (CB400-V)", "*Ø16*", "Tấn", "TCVN 1651:2018", 0.015, 0.25, 0.15, 0.50, 0.10),
            ("5", "Thép phân bố bệ móng vằn Ø20 (CB500-V)", "*Ø20*", "Tấn", "TCVN 1651:2018", 0.015, 0.0, 1.0, 0.0, 0.0),
            ("6", "Thép chủ cọc khoan nhồi vằn Ø25 (CB500-V)", "*Ø25*", "Tấn", "TCVN 1651:2018", 0.015, 1.0, 0.0, 0.0, 0.0),
            ("7", "Thép thân mố & xà mũ vằn Ø28 (CB500-V)", "*Ø28*", "Tấn", "TCVN 1651:2018", 0.015, 0.0, 1.0, 0.0, 0.0),
            ("8", "Thép bệ móng & thân trụ vằn Ø32 (CB500-V)", "*Ø32*", "Tấn", "TCVN 1651:2018", 0.015, 0.0, 1.0, 0.0, 0.0),
            ("9", "Cáp dự ứng lực tao 15.2mm Grade 270", "*Cáp dự ứng lực*", "Tấn", "ASTM A416 Gr270", 0.015, 0.0, 0.0, 1.0, 0.0),
            ("10", "Bộ neo chùm dự ứng lực công tác", "*Bộ neo chùm*", "Bộ", "TCVN 11823:2017", 0.010, 0.0, 0.0, 1.0, 0.0),
            ("11", "Ống ghen tôn mạ kẽm luồn cáp DƯL", "*Ống ghen tôn*", "m", "TCVN 11823:2017", 0.020, 0.0, 0.0, 1.0, 0.0),
            ("12", "Ống siêu âm cọc nhồi thép Ø60x2", "*Ống siêu âm*", "m", "ASTM A53 / TCVN 9395", 0.010, 1.0, 0.0, 0.0, 0.0),
            ("13", "Ống vách thép dẫn hướng D1300 d=8mm", "*Ống vách thép*", "Tấn", "TCVN 9395:2012", 0.020, 1.0, 0.0, 0.0, 0.0),
            ("14", "Tay vịn lan can thép mạ kẽm nhúng nóng Ø114", "*Tay vịn*", "m", "TCVN 11823:2017", 0.020, 0.0, 0.0, 0.0, 1.0),
            ("15", "Thép tấm khe co giãn răng lược D=100mm", "*Thép tấm răng lược*", "m", "TCVN 11823:2017", 0.010, 0.0, 0.0, 0.0, 1.0),
            ("16", "Dây thép buộc 1 ly", "*Dây thép buộc*", "kg", "TCVN 1651:2018", 0.030, 0.35, 0.25, 0.35, 0.05),
            ("17", "Que hàn điện E42/E50", "*Que hàn điện*", "kg", "TCVN 3223:2000", 0.030, 0.50, 0.30, 0.20, 0.0)
        ]),
        ("II. KHOÁNG XÂY DỰNG & CẤP PHỐI BÊ TÔNG", [
            ("18", "Xi măng PCB40 (quy đổi Tấn)", "*Xi măng PCB40*", "Tấn", "TCVN 6260:2020", 0.010, 0.25, 0.30, 0.40, 0.05),
            ("19", "Cát vàng sàng rửa đúc bê tông (Mk >= 2.0)", "*Cát vàng*", "m3", "TCVN 7570:2006", 0.025, 0.25, 0.30, 0.40, 0.05),
            ("20", "Đá dăm 1x2 sạch chọn lọc", "*Đá dăm 1x2*", "m3", "TCVN 7570:2006", 0.025, 0.25, 0.30, 0.40, 0.05),
            ("21", "Đá hộc 20x30cm kè tứ nón mố", "*Đá hộc*", "m3", "TCVN 11823:2017", 0.020, 0.0, 0.0, 0.0, 1.0),
            ("22", "Đất đắp nền đường sau mố K98", "*Đất đắp*", "m3", "TCVN 9436:2012", 0.030, 0.0, 0.0, 0.0, 1.0),
            ("23", "Nước sạch thi công & dưỡng hộ (quy đổi m3)", "*Nước sạch*", "m3", "TCVN 4506:2012", 0.020, 0.25, 0.30, 0.40, 0.05)
        ]),
        ("III. PHỤ GIA VÀ HÓA CHẤT XÂY DỰNG", [
            ("24", "Phụ gia siêu dẻo Polycarboxylate thế hệ 3", "*Phụ gia siêu dẻo*", "kg", "ASTM C494 Type F/G", 0.010, 0.0, 0.30, 0.70, 0.0),
            ("25", "Phụ gia hoá dẻo chậm đông kết", "*Phụ gia hoá dẻo*", "kg", "ASTM C494 Type G", 0.010, 1.0, 0.0, 0.0, 0.0),
            ("26", "Vữa rót không co ngót bơm ghen SikaGrout 214-11", "*Vữa không co ngót*", "kg", "ASTM C1107", 0.015, 0.0, 0.0, 1.0, 0.0),
            ("27", "Bột sét Bentonite tạo vách cọc khoan nhồi", "*Bentonite*", "Tấn", "TCVN 9395:2012", 0.020, 1.0, 0.0, 0.0, 0.0)
        ]),
        ("IV. PHỤ KIỆN & VẬT LIỆU HOÀN THIỆN CẦU", [
            ("28", "Gối chậu cao su cốt bản thép chịu lực", "*Gối chậu*", "Cái", "TCVN 11823:2017", 0.000, 0.0, 0.0, 1.0, 0.0),
            ("29", "Màng / Sơn chống thấm mặt cầu Polymer cải tiến", "*Màng chống thấm*", "m2", "TCVN 11823:2017", 0.020, 0.0, 0.0, 0.0, 1.0),
            ("30", "Bê tông nhựa chặt C12.5 rải nóng", "*Bê tông nhựa*", "Tấn", "TCVN 8819:2011", 0.025, 0.0, 0.0, 0.0, 1.0),
            ("31", "Vải địa kỹ thuật không dệt lọc ngược", "*Vải địa*", "m2", "TCVN 9865:2013", 0.030, 0.0, 0.0, 0.0, 1.0),
            ("32", "Phễu gang thu nước mặt cầu có lưới chắn", "*Phễu gang*", "Bộ", "TCVN 11823:2017", 0.000, 0.0, 0.0, 0.0, 1.0),
            ("33", "Ống thoát nước nhựa uPVC Ø110 dày 3.2mm", "*uPVC Ø110*", "m", "TCVN 8491:2011", 0.020, 0.0, 0.0, 0.0, 1.0),
            ("34", "Nhựa dính bám Tack coat", "*Nhựa dính bám*", "kg", "TCVN 8819:2011", 0.020, 0.0, 0.0, 0.0, 1.0)
        ])
    ]

    r_bom = 6
    for group_title, items in categories:
        # Tiêu đề nhóm
        ws_bom.cell(row=r_bom, column=2, value=group_title).alignment = align_left
        for c in range(1, 12):
            cell = ws_bom.cell(row=r_bom, column=c)
            cell.font = font_section
            cell.fill = fill_section
            cell.border = thin_border
        ws_bom.row_dimensions[r_bom].height = 22
        r_bom += 1

        # Các dòng vật tư
        for it in items:
            stt_val, m_title, pattern, unit_val, std_val, scrap, p1, p2, p3, p4 = it

            ws_bom.cell(row=r_bom, column=1, value=stt_val).alignment = align_center
            ws_bom.cell(row=r_bom, column=2, value=m_title).alignment = align_left
            ws_bom.cell(row=r_bom, column=3, value=std_val).alignment = align_left
            ws_bom.cell(row=r_bom, column=4, value=unit_val).alignment = align_center

            # CÔNG THỨC ĐỘNG =SUMIF trỏ sang Sheet PHAN_TICH_VAT_TU_WBS
            # Chú ý: Xi măng từ kg đổi ra Tấn (/1000), Nước từ Lít đổi ra m3 (/1000), Thép Ø12 và Ø16 từ kg đổi ra Tấn nếu cần
            if m_title.startswith("Xi măng") or m_title.startswith("Nước sạch") or m_title.startswith("Bột sét Bentonite"):
                formula_sum = f"=SUMIF(PHAN_TICH_VAT_TU_WBS!$F$6:$F${max_wbs_row}, \"{pattern}\", PHAN_TICH_VAT_TU_WBS!$I$6:$I${max_wbs_row})/1000"
            elif "Thép cốt vằn Ø12" in m_title:
                formula_sum = f"=SUMIF(PHAN_TICH_VAT_TU_WBS!$F$6:$F${max_wbs_row}, \"{pattern}\", PHAN_TICH_VAT_TU_WBS!$I$6:$I${max_wbs_row})/1000"
            else:
                formula_sum = f"=SUMIF(PHAN_TICH_VAT_TU_WBS!$F$6:$F${max_wbs_row}, \"{pattern}\", PHAN_TICH_VAT_TU_WBS!$I$6:$I${max_wbs_row})"

            c_theo = ws_bom.cell(row=r_bom, column=5, value=formula_sum)
            c_theo.alignment = align_right
            c_theo.number_format = "#,##0.00"

            c_scrap = ws_bom.cell(row=r_bom, column=6, value=scrap)
            c_scrap.alignment = align_right
            c_scrap.number_format = "0.0%"

            # Cột G: Tổng nhu cầu cung ứng BOM = E * (1 + F)
            c_total = ws_bom.cell(row=r_bom, column=7, value=f"=E{r_bom}*(1+F{r_bom})")
            c_total.alignment = align_right
            c_total.number_format = "#,##0.00"
            c_total.font = font_bold
            c_total.fill = fill_total

            # Cột H, I, J, K: Phân bổ theo 4 giai đoạn thi công = G * tỷ trọng
            c_p1 = ws_bom.cell(row=r_bom, column=8, value=f"=G{r_bom}*{p1}")
            c_p1.alignment = align_right
            c_p1.number_format = "#,##0.00"

            c_p2 = ws_bom.cell(row=r_bom, column=9, value=f"=G{r_bom}*{p2}")
            c_p2.alignment = align_right
            c_p2.number_format = "#,##0.00"

            c_p3 = ws_bom.cell(row=r_bom, column=10, value=f"=G{r_bom}*{p3}")
            c_p3.alignment = align_right
            c_p3.number_format = "#,##0.00"

            c_p4 = ws_bom.cell(row=r_bom, column=11, value=f"=G{r_bom}*{p4}")
            c_p4.alignment = align_right
            c_p4.number_format = "#,##0.00"

            for c in range(1, 12):
                cell = ws_bom.cell(row=r_bom, column=c)
                if not cell.font:
                    cell.font = font_regular
                cell.border = thin_border
            ws_bom.row_dimensions[r_bom].height = 20
            r_bom += 1

    # Căn chỉnh độ rộng cột Sheet TONG_HOP_VAT_TU_TOAN_BO
    widths_bom = [6, 45, 24, 10, 24, 20, 26, 18, 18, 18, 18]
    for i, w in enumerate(widths_bom, start=1):
        ws_bom.column_dimensions[get_column_letter(i)].width = w
    ws_bom.freeze_panes = "E6"

    # Lưu lại workbook
    wb.save(excel_path)
    print(f"[+] Đã cập nhật thành công 2 sheet vật tư mới vào: {excel_path}")
    print(f"    -> Tổng số sheet hiện tại: {len(wb.sheetnames)}: {wb.sheetnames}")

if __name__ == "__main__":
    p1 = repo_path(r"templates\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")
    add_material_sheets(p1)
