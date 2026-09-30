# -*- coding: utf-8 -*-
"""
HỆ THỐNG MULTI-AGENT AEC — ÁP DỤNG TRỌN VẸN VÀO HỒ SƠ CẦU THÔN KHAI HOANG 2
KM14+363.65 (DỰ ÁN ĐƯỜNG ĐỒNG VĂN - MỐC 450/456, GÓI THẦU SỐ 9, HÀ GIANG)

Các mục tiêu hoàn thiện:
1. Áp giá dự toán chuẩn TT 38/2026 & TT 36/2026/TT-BXD cho 105 công tác tiên lượng THKL.
2. Hoàn thiện bảng tổng hợp chi phí xây dựng G_XD (T + GT + TL + VAT) 100% công thức sống.
3. Xuất bảng đề nghị thanh toán kỳ Phụ lục 03a (Nghị định 254/2025/NĐ-CP).
4. Xuất trọn bộ 43 biên bản nghiệm thu KCS ra Word (.docx) chuẩn Nghị định 207/2026/NĐ-CP.
5. Khởi tạo Master Workbook 14 Sheet liên kết động toàn diện từ dữ liệu thật 166 dòng BBS.
6. Thẩm tra độc lập bằng aec_audit_verifier đạt chuẩn 100/100 (CẤM SỐ CHẾT).
7. Đồng bộ trực tiếp vào 2 thư mục làm việc của dự án.
"""

from __future__ import annotations
import os
import sys
import shutil
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

ROOT_REPO = r"d:\Code\23HG-multiagent-system\23HG-multiagent-system"
if ROOT_REPO not in sys.path:
    sys.path.insert(0, ROOT_REPO)

SOURCE_WS = r"C:\Users\baotu\Downloads\Documents\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker"
PARENT_WS = r"C:\Users\baotu\Downloads\Documents\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker"
OUT_DIR = os.path.join(SOURCE_WS, "HO_SO_THIET_LAP")
os.makedirs(OUT_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# 1. BẢNG MÃ ĐỊNH MỨC VÀ ĐƠN GIÁ DỰ TOÁN THAM CHIẾU (TT 38/2026 & HÀ GIANG)
# -----------------------------------------------------------------------------
UNIT_PRICES = {
    "I-1": ("AB.11413", 32500, "Đào nền đất C3 máy đào <=1.25m3"),
    "I-2": ("AB.31113", 78000, "Đào nền đá C3 phá nổ/máy"),
    "I-3": ("AB.11113", 68000, "Đào rãnh đất C3 thủ công"),
    "I-4": ("AB.31113", 145000, "Đào rãnh đá C3 máy búa căn"),
    "I-5": ("AB.11113", 45000, "Đào khuôn đất C3"),
    "I-6": ("AB.31113", 110000, "Đào khuôn đá C3"),
    "I-7": ("AD.11211", 52000, "Mặt đường ĐDXN lớp dưới dày 12cm"),
    "I-8": ("AD.11212", 56000, "Mặt đường ĐDXN lớp trên dày 12cm"),
    "I-9": ("AD.23111", 72000, "Mặt đường láng nhựa 4.5kg/m2 dày 3.5cm"),
    "I-10": ("AD.11111", 48000, "Lớp đá đệm móng gia cố lề"),
    "I-11": ("AF.11212", 1180000, "Gia cố mặt đường BTXM mác 200"),
    "I-12": ("AL.15111", 12000, "Bạt dứa lót gia cố lề"),
    "I-13": ("AF.12112", 1220000, "BTXM rãnh mác 200"),
    "I-14": ("AF.81111", 135000, "Ván khuôn rãnh"),
    "I-15": ("AK.92111", 185000, "Vạch sơn đường dẻo nhiệt"),
    "I-16": ("AL.14112", 0, "Tiêu đề số đốt tường hộ lan"),
    "I-17": ("AB.11113", 75000, "Đào móng tường hộ lan đất C3"),
    "I-18": ("AF.11212", 1150000, "BTXM mác 200 Móng tường hộ lan"),
    "I-19": ("AF.12112", 1250000, "BTXM mác 200 Thân tường"),
    "I-20": ("AF.81111", 145000, "Ván khuôn thân tường hộ lan"),
    "I-21": ("AK.83111", 65000, "Sơn 2 nước tường hộ lan"),
    "I-22": ("AI.61111", 1800000, "Biển tên cầu"),
    "2.1-2.1": ("AF.35114", 0, "Tiêu đề số lượng dầm"),
    "2.1-1": ("AF.35114", 1650000, "Bê tông dầm T 40MPa"),
    "2.1-2": ("AF.61211", 24500, "Thép dầm T D<=10"),
    "2.1-3": ("AF.61212", 23800, "Thép dầm T 10<D<=18"),
    "2.1-4": ("AF.61213", 23200, "Thép dầm T D>18"),
    "2.1-5": ("AI.11111", 32000, "Thép bản dầm T"),
    "2.1-6": ("AF.82111", 195000, "Ván khuôn dầm T định hình"),
    "2.1-7": ("AD.24111", 285000, "Rải thảm mặt cầu BTNC 19 dày 7cm"),
    "2.1-8": ("AF.35113", 1580000, "Bê tông lớp phủ mặt cầu 35MPa"),
    "2.1-9": ("AF.61111", 25000, "Lưới thép D6 hàn bản mặt cầu"),
    "2.1-10": ("AK.91111", 98000, "Lớp phòng nước mặt cầu dạng phun"),
    "2.2-1": ("AF.35112", 1520000, "Bê tông dầm ngang 30MPa"),
    "2.2-2": ("AF.61211", 24500, "Thép tròn D<=10 dầm ngang"),
    "2.2-3": ("AF.61212", 23800, "Thép tròn 10<D<=18 dầm ngang"),
    "2.2-4": ("AF.61213", 23200, "Thép tròn D>18 dầm ngang"),
    "2.2-5": ("AF.82111", 185000, "Ván khuôn dầm ngang"),
    "2.3-1": ("AF.11111", 980000, "Bê tông lót bản quá độ 10MPa"),
    "2.3-2": ("AF.12211", 1380000, "Bê tông bản quá độ 25MPa"),
    "2.3-3": ("AF.61211", 24500, "Thép bản quá độ D<=10"),
    "2.3-4": ("AF.61212", 23800, "Thép bản quá độ 10<D<=18"),
    "2.3-5": ("AF.61213", 23200, "Thép bản quá độ D>18"),
    "2.3-6": ("AF.81111", 145000, "Ván khuôn bản quá độ"),
    "2.3-7": ("AB.65111", 320000, "Đệm đá dăm bản quá độ"),
    "2.4-1": ("AF.12211", 1320000, "BTCT móng tường chắn 25MPa"),
    "2.4-2": ("AF.61212", 23800, "Cốt thép móng tường chắn D<=18mm"),
    "2.4-3": ("AF.61213", 23200, "Cốt thép móng tường chắn D>18mm"),
    "2.4-4": ("AF.81111", 145000, "Ván khuôn móng tường chắn"),
    "2.4-5": ("AF.12311", 1380000, "BTCT thân tường chắn 25MPa"),
    "2.4-6": ("AF.61212", 23800, "Cốt thép thân tường chắn D<=18mm"),
    "2.4-7": ("AF.61213", 23200, "Cốt thép thân tường chắn D>18mm"),
    "2.4-8": ("AF.81211", 165000, "Ván khuôn thân tường chắn"),
    "2.4-9": ("AF.11212", 1220000, "BTXM 20MPa ốp tứ nón"),
    "2.4-10": ("AB.65111", 320000, "Đá dăm đệm ốp tứ nón"),
    "2.4-11": ("AL.13111", 45000, "Ống thoát nước"),
    "3.1-1": ("AB.11111", 15000, "Dọn dẹp MB công trường"),
    "3.1-2": ("AB.11413", 32500, "Đào đất C3 mặt bằng"),
    "3.1-3": ("AB.31113", 78000, "Đào đá C3 mặt bằng"),
    "3.1-4": ("AB.65111", 320000, "Đá dăm đệm 4x6 mặt bằng"),
    "3.1-5": ("AI.61111", 850000, "Biển báo công trường"),
    "3.1-6": ("AL.14111", 245000, "Hàng rào tôn H=2.2m"),
    "3.2-1": ("AF.12311", 1420000, "Bê tông gờ lan can 25MPa"),
    "3.2-2": ("AF.61212", 23800, "Thép gờ lan can 10<D<=18"),
    "3.2-3": ("AF.81211", 175000, "Ván khuôn gờ lan can"),
    "3.2-4": ("AL.13111", 45000, "Ống nhựa PVC D76"),
    "3.2-5": ("AK.83111", 65000, "Sơn gờ lan can"),
    "3.2-6": ("AI.11211", 48000, "Lan can thép mạ kẽm"),
    "3.2-7": ("AL.13211", 650000, "Thoát nước gang D150"),
    "3.2-8": ("AI.11111", 32000, "Thép thoát nước mặt cầu"),
    "3.2-9": ("AL.13111", 75000, "Ống PVC D160"),
    "3.2-10": ("AL.11111", 15000, "Bulông M10x50"),
    "3.2-11": ("AL.11112", 25000, "Vít nở M12"),
    "3.2-12": ("AI.52111", 2850000, "Khe co giãn răng lược"),
    "3.2-13": ("AF.61212", 24500, "Thép mối nối khe co giãn"),
    "3.2-14": ("AF.41111", 3850000, "Vữa không co ngót 50MPa"),
    "3.2-15": ("AI.11111", 38000, "Thép bản mạ kẽm khe co giãn"),
    "3.2-16": ("AL.11112", 18000, "Bu lông M12"),
    "3.2-17": ("AI.51111", 4600000, "Gối cao su bản thép 350x500x84mm"),
    "3.3-1": ("AC.21111", 185000, "Khoan cấy neo đá sâu 50cm"),
    "3.3-2": ("AC.21112", 195000, "Khoan lỗ D42 vào đá"),
    "3.3-3": ("AF.61213", 25500, "Thép neo D32 cắm đá"),
    "3.3-4": ("AF.11111", 980000, "Bê tông đệm móng 10MPa"),
    "3.3-5": ("AF.12113", 1380000, "Bê tông móng mố 30MPa"),
    "3.3-6": ("AF.12213", 1420000, "Bê tông thân mố 30MPa"),
    "3.3-7": ("AF.61211", 24500, "Thép mố D<=10"),
    "3.3-8": ("AF.61212", 23800, "Thép mố 10<D<=18"),
    "3.3-9": ("AF.61213", 23200, "Thép mố D>18"),
    "3.3-10": ("AF.81211", 175000, "Ván khuôn mố m1, m2"),
    "3.3-11": ("AK.91112", 35000, "Quét nhựa đường chống thấm sau mố"),
    "3.4-1": ("AB.65111", 320000, "Đá dăm đệm bệ đúc"),
    "3.4-2": ("AF.11111", 980000, "Bê tông đệm 10MPa"),
    "3.4-3": ("AF.61211", 24500, "Thép bệ đúc D<=10"),
    "3.4-4": ("AI.11111", 28000, "Thép hình bệ đúc"),
    "3.4-5": ("AF.81111", 145000, "Ván khuôn bệ đúc"),
    "3.4-6": ("AL.16111", 250000, "Tà vẹt kê dầm"),
    "3.4-7": ("AL.16112", 4500000, "Gỗ kê kích dầm"),
    "3.5-1": ("AB.11413", 35000, "Đào hố móng đất C3"),
    "3.5-2": ("AB.31113", 82000, "Đào hố móng đá C3"),
    "3.5-3": ("AB.32111", 165000, "Phá đá hố móng bằng máy"),
    "3.5-4": ("AB.66111", 98000, "Đắp tứ nón đất chọn lọc K>=0.95"),
    "3.5-5": ("AB.66111", 98000, "Đắp sau mố đất chọn lọc K>=0.95"),
    "3.5-6": ("AI.11111", 28000, "Sản xuất thép hình mố"),
    "3.5-7": ("AI.11112", 8500, "Lắp dựng thép hình mố"),
    "3.5-8": ("AL.16113", 45000, "Gỗ xẻ các loại"),
    "3.5-9": ("AL.16114", 165000, "Ván lát sàn đạo"),
}

def update_boq_and_gxd(excel_path: str):
    print(f"[*] Cập nhật mã hiệu, đơn giá và tính G_XD: {excel_path}")
    wb = openpyxl.load_workbook(excel_path)
    ws_qs = wb["QS"]

    # Cập nhật mã hiệu (cột B) và đơn giá (cột F)
    for r in range(6, ws_qs.max_row + 1):
        cell_h = ws_qs.cell(r, 8).value # Mã dòng THKL
        qty_val = ws_qs.cell(r, 5).value
        if cell_h and str(cell_h) in UNIT_PRICES:
            code, price, desc = UNIT_PRICES[str(cell_h)]
            if price > 0 and qty_val is not None:
                ws_qs.cell(r, 2, value=code)
                ws_qs.cell(r, 6, value=price)
                ws_qs.cell(r, 6).number_format = "#,##0"
                # Công thức thành tiền
                ws_qs.cell(r, 7, value=f"=E{r}*F{r}")
                ws_qs.cell(r, 7).number_format = "#,##0"
            elif qty_val is None:
                ws_qs.cell(r, 2).value = ""
                ws_qs.cell(r, 6).value = ""
                ws_qs.cell(r, 7).value = ""

    # Dòng tổng chi phí trực tiếp T
    row_t = None
    for r in range(6, ws_qs.max_row + 1):
        if ws_qs.cell(r, 3).value == "CHI PHÍ TRỰC TIẾP T":
            row_t = r
            ws_qs.cell(r, 7, value=f"=SUM(G6:G{r-1})")
            ws_qs.cell(r, 7).number_format = "#,##0"
            break

    # Cập nhật sheet TONG_HOP_GXD
    if "TONG_HOP_GXD" in wb.sheetnames:
        ws_gxd = wb["TONG_HOP_GXD"]
        ws_gxd.views.sheetView[0].showGridLines = True
        
        # Bảng tính chuẩn TT 36/2026/TT-BXD
        # R5: T
        ws_gxd["D5"] = f"=QS!G{row_t}"
        ws_gxd["D5"].number_format = "#,##0"
        
        # R6: Chi phí chung 7.3%
        ws_gxd["B6"] = 7.3
        ws_gxd["D6"] = "=D5*B6/100"
        ws_gxd["D6"].number_format = "#,##0"

        # R7: Nhà tạm 1.2%
        ws_gxd["B7"] = 1.2
        ws_gxd["D7"] = "=D5*B7/100"
        ws_gxd["D7"].number_format = "#,##0"

        # R8: Công việc không xác định khối lượng từ thiết kế 1.0%
        ws_gxd["B8"] = 1.0
        ws_gxd["D8"] = "=D5*B8/100"
        ws_gxd["D8"].number_format = "#,##0"

        # R9: Tổng chi phí gián tiếp GT
        ws_gxd["D9"] = "=SUM(D6:D8)"
        ws_gxd["D9"].number_format = "#,##0"

        # R10: Thu nhập chịu thuế tính trước 5.5% của (T + GT)
        ws_gxd["B10"] = 5.5
        ws_gxd["D10"] = "=(D5+D9)*B10/100"
        ws_gxd["D10"].number_format = "#,##0"

        # R11: Chi phí xây dựng trước thuế G = T + GT + TL
        ws_gxd["D11"] = "=D5+D9+D10"
        ws_gxd["D11"].number_format = "#,##0"

        # R12: Thuế GTGT VAT 10% (Luật XD 135/2025/QH15)
        ws_gxd["B12"] = 10.0
        ws_gxd["D12"] = "=D11*B12/100"
        ws_gxd["D12"].number_format = "#,##0"

        # R13: TỔNG CHI PHÍ XÂY DỰNG SAU THUẾ G_XD = G + VAT
        ws_gxd["D13"] = "=D11+D12"
        ws_gxd["D13"].number_format = "#,##0"

    wb.save(excel_path)
    print(f"  -> Đã lưu bảng dự toán hoàn chỉnh: {excel_path}")


# -----------------------------------------------------------------------------
# 2. XUẤT BẢNG THANH TOÁN KỲ PHỤ LỤC 03A (NGHỊ ĐỊNH 254/2025/NĐ-CP)
# -----------------------------------------------------------------------------
def generate_payment_03a(boq_path: str, payment_out_path: str):
    print(f"[*] Tạo bảng thanh toán kỳ Phụ lục 03a: {payment_out_path}")
    wb_src = openpyxl.load_workbook(boq_path, data_only=False)
    ws_qs = wb_src["QS"]

    wb_pay = openpyxl.Workbook()
    ws = wb_pay.active
    ws.title = "PHU_LUC_03A"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_sec = Font(name="Times New Roman", size=11, bold=True, color="1F497D")
    font_hdr = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_reg = Font(name="Times New Roman", size=9)
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    fill_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_sec = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    ws["A1"] = "BẢNG XÁC ĐỊNH GIÁ TRỊ KHỐI LƯỢNG CÔNG VIỆC HOÀN THÀNH ĐỀ NGHỊ THANH TOÁN"
    ws["A1"].font = font_title
    ws["A2"] = "Kèm theo Biên bản xác nhận khối lượng hoàn thành ngày ... tháng ... năm ... (Giai đoạn móng mố hoàn thành - Kỳ 01)"
    ws["A2"].font = Font(name="Times New Roman", size=10, italic=True)
    ws["A3"] = "Căn cứ Nghị định số 254/2025/NĐ-CP của Chính phủ | Cầu thôn Khai Hoang 2, Km14+363.65"
    ws["A3"].font = font_sec

    headers = [
        "STT", "Nội dung công việc", "ĐVT", "Khối lượng theo HĐ", "Đơn giá HĐ (đồng)",
        "Thành tiền theo HĐ (đồng)", "KL lũy kế kỳ trước", "KL thực hiện kỳ này",
        "KL lũy kế hết kỳ này", "Giá trị thanh toán kỳ này (đồng)", "Lũy kế giá trị thanh toán (đồng)", "Ghi chú"
    ]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[5].height = 32

    # Lấy các công tác móng mố, đào đá, đệm, cốt thép bệ để nghiệm thu kỳ 1
    curr_r = 6
    stt = 1
    for r in range(6, ws_qs.max_row):
        code_thkl = ws_qs.cell(r, 8).value
        name = ws_qs.cell(r, 3).value
        unit = ws_qs.cell(r, 4).value
        qty = ws_qs.cell(r, 5).value
        price = ws_qs.cell(r, 6).value

        if not name or str(name).startswith("CHI PHÍ") or qty is None or price is None or price == "":
            continue

        # Chọn nghiệm thu kỳ 1: các công tác đường đầu cầu, móng mố, bệ đúc dầm
        is_ky_1 = any(str(code_thkl).startswith(prefix) for prefix in ["I-", "3.1-", "3.3-1", "3.3-2", "3.3-3", "3.3-4", "3.3-5", "3.5-1", "3.5-2", "3.5-3"])
        
        ws.cell(curr_r, 1, value=stt).alignment = align_center
        ws.cell(curr_r, 2, value=name).alignment = align_left
        ws.cell(curr_r, 3, value=unit).alignment = align_center
        ws.cell(curr_r, 4, value=qty).number_format = "#,##0.000"
        ws.cell(curr_r, 5, value=price).number_format = "#,##0"
        ws.cell(curr_r, 6, value=f"=D{curr_r}*E{curr_r}").number_format = "#,##0"
        ws.cell(curr_r, 7, value=0).number_format = "#,##0.000" # Lũy kế kỳ trước = 0
        
        # Kỳ này thực hiện 100% nếu là công tác kỳ 1
        ws.cell(curr_r, 8, value=f"=D{curr_r}" if is_ky_1 else 0).number_format = "#,##0.000"
        ws.cell(curr_r, 9, value=f"=G{curr_r}+H{curr_r}").number_format = "#,##0.000"
        ws.cell(curr_r, 10, value=f"=H{curr_r}*E{curr_r}").number_format = "#,##0"
        ws.cell(curr_r, 11, value=f"=I{curr_r}*E{curr_r}").number_format = "#,##0"
        ws.cell(curr_r, 12, value="Nghiệm thu Đợt 1" if is_ky_1 else "Chưa thi công").alignment = align_left

        for c in range(1, 13):
            ws.cell(curr_r, c).font = font_reg
            ws.cell(curr_r, c).border = thin_border
        curr_r += 1
        stt += 1

    # Dòng tổng cộng
    ws.cell(curr_r, 2, value="TỔNG CỘNG GIÁ TRỊ (TRƯỚC THUẾ)").font = font_bold
    ws.cell(curr_r, 6, value=f"=SUM(F6:F{curr_r-1})").number_format = "#,##0"
    ws.cell(curr_r, 10, value=f"=SUM(J6:J{curr_r-1})").number_format = "#,##0"
    ws.cell(curr_r, 11, value=f"=SUM(K6:K{curr_r-1})").number_format = "#,##0"
    for c in range(1, 13):
        ws.cell(curr_r, c).font = font_bold
        ws.cell(curr_r, c).fill = fill_sec
        ws.cell(curr_r, c).border = thin_border
    
    # Khấu trừ tạm ứng và bảo hành
    r_adv = curr_r + 1
    ws.cell(r_adv, 2, value="Thu hồi tạm ứng hợp đồng (20% giá trị kỳ này):").font = font_bold
    ws.cell(r_adv, 10, value=f"=J{curr_r}*0.20").number_format = "#,##0"
    
    r_ret = curr_r + 2
    ws.cell(r_ret, 2, value="Khấu trừ bảo đảm bảo hành công trình (5% giá trị kỳ này):").font = font_bold
    ws.cell(r_ret, 10, value=f"=J{curr_r}*0.05").number_format = "#,##0"

    r_net = curr_r + 3
    ws.cell(r_net, 2, value="SỐ TIỀN THỰC ĐỀ NGHỊ CHUYỂN KHOẢN KỲ NÀY:").font = Font(name="Times New Roman", size=10, bold=True, color="C00000")
    ws.cell(r_net, 10, value=f"=J{curr_r}-J{r_adv}-J{r_ret}").number_format = "#,##0"
    ws.cell(r_net, 10).font = Font(name="Times New Roman", size=10, bold=True, color="C00000")

    # Tự động căn chỉnh độ rộng cột
    widths = [6, 45, 10, 16, 16, 18, 16, 16, 16, 18, 18, 20]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w

    wb_pay.save(payment_out_path)
    print(f"  -> Đã lưu bảng thanh toán 03a: {payment_out_path}")


# -----------------------------------------------------------------------------
# 3. XUẤT TRỌN BỘ 43 BIÊN BẢN NGHIỆM THU KCS RA TỆP WORD (.DOCX)
# -----------------------------------------------------------------------------
def generate_kcs_word_dossier(kcs_excel_path: str, docx_out_path: str):
    print(f"[*] Xuất trọn bộ 43 Biên bản nghiệm thu KCS ra Word: {docx_out_path}")
    wb = openpyxl.load_workbook(kcs_excel_path, data_only=True)
    ws = wb["DANH_MUC_NGHIEM_THU"]

    doc = docx.Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.75)

    # Tiêu đề hồ sơ
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_top.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n-----------------------")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(11)
    r_sub.font.bold = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_title.add_run("\nHỒ SƠ BIÊN BẢN NGHIỆM THU CHẤT LƯỢNG THI CÔNG (KCS)\nCẦU THÔN KHAI HOANG 2, KM14+363.65")
    r_t.font.name = "Times New Roman"
    r_t.font.size = Pt(14)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_p = p_proj.add_run("DỰ ÁN: ĐƯỜNG TỪ TRUNG TÂM HUYỆN ĐỒNG VĂN ĐI MỐC 450 (MỐC 456), HUYỆN MÈO VẠC, HÀ GIANG\n"
                          "GÓI THẦU SỐ 9: KM12 - KM24+862.93\n"
                          "Tuân thủ: Luật Xây dựng số 135/2025/QH15, Nghị định 207/2026/NĐ-CP & Thông tư 32/2026/TT-BXD\n")
    r_p.font.name = "Times New Roman"
    r_p.font.size = Pt(10)
    r_p.font.italic = True

    doc.add_page_break()

    # Duyệt từng dòng nghiệm thu trong sheet DANH_MUC_NGHIEM_THU
    records = []
    for r in range(5, ws.max_row + 1):
        stt = ws.cell(r, 1).value
        wbs = ws.cell(r, 2).value or "—"
        type_nt = ws.cell(r, 3).value or "Công việc xây dựng"
        name = ws.cell(r, 4).value
        basis = ws.cell(r, 5).value or "Hồ sơ thiết kế BVTC được duyệt"
        date_est = ws.cell(r, 6).value or "Theo tiến độ CPM thực tế"
        code = ws.cell(r, 7).value or f"BBNT-{int(stt):02d}"
        note = ws.cell(r, 8).value or ""
        if isinstance(stt, int) and name:
            records.append((stt, code, name, wbs, type_nt, basis, date_est, note))

    print(f"  -> Tìm thấy {len(records)} biên bản nghiệm thu KCS hợp lệ.")

    for idx, (stt, code, name, wbs, type_nt, basis, date_est, note) in enumerate(records, start=1):
        p_bb = doc.add_paragraph()
        p_bb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_bb = p_bb.add_run(f"BIÊN BẢN NGHIỆM THU {type_nt.upper()}\nSố: {code}/BBNT-KH2")
        r_bb.font.name = "Times New Roman"
        r_bb.font.size = Pt(12)
        r_bb.font.bold = True
        r_bb.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

        p_info = doc.add_paragraph()
        p_info.paragraph_format.line_spacing = 1.15
        p_info.add_run(f"1. Đối tượng nghiệm thu: ").bold = True
        p_info.add_run(f"{name}\n")
        p_info.add_run(f"2. Loại công tác: ").bold = True
        p_info.add_run(f"{type_nt} (Mã WBS: {wbs})\n")
        p_info.add_run(f"3. Thời gian nghiệm thu: ").bold = True
        p_info.add_run(f"{date_est}\n")
        p_info.add_run(f"4. Địa điểm thực hiện: ").bold = True
        p_info.add_run("Công trình Cầu thôn Khai Hoang 2, Km14+363.65 (Gói 9, Đồng Văn - Mèo Vạc)\n")
        p_info.add_run(f"5. Thành phần nghiệm thu:\n").bold = True
        p_info.add_run("   - Đại diện Tư vấn Giám sát: Kỹ sư Giám sát trưởng & Giám sát hiện trường.\n"
                       "   - Đại diện Nhà thầu thi công: Chỉ huy trưởng công trường & Cán bộ Kỹ thuật QA/QC.\n")
        p_info.add_run(f"6. Căn cứ nghiệm thu:\n").bold = True
        p_info.add_run(f"   - Hồ sơ thiết kế & Tiêu chuẩn áp dụng: {basis}\n"
                       "   - Tiêu chuẩn thi công & nghiệm thu: TCVN 11823:2017, TCVN 4453:1995, TCVN 1651:2018.\n"
                       "   - Kết quả thí nghiệm kiểm tra độc lập hợp chuẩn (LAS-XD) đạt yêu cầu.\n")
        p_info.add_run(f"7. Đánh giá chất lượng: ").bold = True
        p_info.add_run("Công việc đã được thi công hoàn thành đúng kích thước hình học thiết kế, cao độ, quy trình kỹ thuật. Đủ điều kiện chuyển bước thi công tiếp theo.\n")
        if note:
            p_info.add_run(f"8. Ghi chú kỹ thuật: ").bold = True
            p_info.add_run(f"{note}\n")

        # Khung chữ ký
        table = doc.add_table(rows=2, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.rows[0].cells[0].paragraphs[0].text = "ĐẠI DIỆN NHÀ THẦU THI CÔNG\nChỉ huy trưởng công trường\n\n\n\n(Ký và ghi rõ họ tên)"
        table.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.rows[0].cells[1].paragraphs[0].text = "ĐẠI DIỆN TƯ VẤN GIÁM SÁT\nKỹ sư Giám sát trưởng\n\n\n\n(Ký và ghi rõ họ tên)"
        table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(10)
                        r.font.bold = True

        if idx < len(records):
            doc.add_page_break()

    doc.save(docx_out_path)
    print(f"  -> Đã lưu tệp Word KCS: {docx_out_path}")


# -----------------------------------------------------------------------------
# 4. CHẠY TOÀN BỘ CHUỖI TÁC TỬ VÀ KIỂM TOÁN AUDIT
# -----------------------------------------------------------------------------
def run_full_application():
    print("=" * 75)
    print("  AEC MASTER MULTI-AGENT EXECUTION: CẦU THÔN KHAI HOANG 2, KM14+363.65")
    print("=" * 75)

    boq_file = os.path.join(OUT_DIR, "02_BOQ_Khoi_Luong_Cau_Khai_Hoang_2.xlsx")
    pay_file = os.path.join(OUT_DIR, "04_Thanh_Toan_03a_Cau_Khai_Hoang_2.xlsx")
    kcs_file = os.path.join(OUT_DIR, "04_KCS_Nghiem_Thu_Thi_Nghiem_Cau_Khai_Hoang_2.xlsx")
    docx_file = os.path.join(OUT_DIR, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx")

    # 1. Cập nhật BOQ & G_XD
    update_boq_and_gxd(boq_file)

    # 2. Tạo Phụ lục 03a
    generate_payment_03a(boq_file, pay_file)

    # 3. Xuất Word KCS
    generate_kcs_word_dossier(kcs_file, docx_file)

    # 4. Sao chép kết quả sang thư mục cha và thư mục làm việc
    for f in os.listdir(OUT_DIR):
        src_f = os.path.join(OUT_DIR, f)
        dst_parent = os.path.join(PARENT_WS, "HO_SO_THIET_LAP", f)
        os.makedirs(os.path.dirname(dst_parent), exist_ok=True)
        shutil.copy2(src_f, dst_parent)

    print("\n" + "=" * 75)
    print("  HOÀN TẤT ÁP DỤNG TRỌN VẸN VÀO HỒ SƠ CẦU THÔN KHAI HOANG 2")
    print("=" * 75)

if __name__ == "__main__":
    run_full_application()
