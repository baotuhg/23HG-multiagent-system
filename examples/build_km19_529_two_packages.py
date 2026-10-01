# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT 2 GÓI HỒ SƠ DỰ ÁN CẦU KM19+529.080 (CAO TỐC TUYÊN QUANG - HÀ GIANG):
1. GÓI 01 - VĨ MÔ / MASTER ĐIỀU HÀNH:
   - Thư mục: BO_HO_SO_01_MACRO_MASTER_14_SHEET
   - Chứa Master Excel 14 Sheet liên kết động toàn diện (0 số chết, Audit Score 100/100)
   - Chứa Tiến độ thi công MS Project XML & MPP
   - Chứa Hồ sơ KCS 22 Biên bản nghiệm thu định dạng Word .docx
   - Chứa Thuyết minh biện pháp thi công & Báo cáo Thẩm tra AEC Audit 100/100

2. GÓI 02 - VI MÔ / CHUYÊN SÂU SẢN XUẤT:
   - Thư mục: BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO
   - Đầy đủ 14 bộ hồ sơ riêng biệt khớp 1-1 với 14 Sheet Master, sâu đến từng chi tiết thi công:
     01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx (kèm file CSV lệnh cắt CNC)
     02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx (Mố M1, M2, Trụ T1, T2 & Đắp K95/K98)
     03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx (101 dòng hình học & chi tiết dầm Super-T)
     04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx (396 dòng cốt thép & bảng phân nhóm Ø)
     05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx (Cấp phối C10-C50 & kế hoạch 809 mẫu QA/QC)
     06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx (140 dòng phân tích vật tư định mức TT 38/2026)
     07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx (BOM toàn bộ & kế hoạch cung ứng 4 phase)
     08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx (Dự toán chi phí xây dựng G_XD chuẩn TT11)
     09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx (Thanh toán Kỳ 01 theo NĐ 254/2025)
     10_Tien_Do_Thi_Cong_CPM_Gantt_Chart.xlsx (36 công tác WBS, CPM, Gantt, kèm .xml, .mpp, .csv)
     11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx (22 biên bản nghiệm thu KCS, kèm .docx)
     12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx (Mẫu in A4 tự động tra cứu, nhúng data local)
     13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx (Mẫu in A4 nghiệm thu vật liệu, nhúng data local)
     14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx (Mẫu in A4 lấy mẫu nén bê tông R7/R28)
"""

from __future__ import annotations
import os
import sys
import shutil
import csv
from copy import copy
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Đường dẫn nguồn và đích
REPO_DIR = r"d:\Code\23HG-multiagent-system\23HG-multiagent-system"
if REPO_DIR not in sys.path:
    sys.path.insert(0, REPO_DIR)
TEMPLATES_DIR = os.path.join(REPO_DIR, "templates")

TARGET_PARENT = r"c:\Users\baotu\Downloads\Documents\HSTK Cầu Km19+529.080_Marker"
TARGET_NESTED = os.path.join(TARGET_PARENT, "HSTK Cầu Km19+529.080_Marker")

DIR_MACRO_P = os.path.join(TARGET_PARENT, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
DIR_MICRO_P = os.path.join(TARGET_PARENT, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")

DIR_MACRO_N = os.path.join(TARGET_NESTED, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
DIR_MICRO_N = os.path.join(TARGET_NESTED, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")

for d in [DIR_MACRO_P, DIR_MICRO_P, DIR_MACRO_N, DIR_MICRO_N]:
    os.makedirs(d, exist_ok=True)

MASTER_SOURCE = os.path.join(TEMPLATES_DIR, "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")

# -----------------------------------------------------------------------------
# STYLES CHUẨN
# -----------------------------------------------------------------------------
FONT_TITLE = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
FONT_SUBTITLE = Font(name="Times New Roman", size=10, italic=True, color="595959")
FONT_SEC = Font(name="Times New Roman", size=11, bold=True, color="1F497D")
FONT_HDR = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
FONT_BOLD = Font(name="Times New Roman", size=9, bold=True)
FONT_REG = Font(name="Times New Roman", size=9)
FONT_IT = Font(name="Times New Roman", size=9, italic=True)

FILL_HDR = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
FILL_SUBHDR = PatternFill(start_color="244062", end_color="244062", fill_type="solid")
FILL_SEC = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
FILL_TOT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_INPUT = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_WARN = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
FILL_OK = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

THIN_GRAY = Side(style='thin', color='BFBFBF')
THIN_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
DOUBLE_BOTTOM_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=Side(style='double', color='1F497D'))
BOX_BORDER = Border(left=Side(style='medium', color='1F497D'), right=Side(style='medium', color='1F497D'),
                    top=Side(style='medium', color='1F497D'), bottom=Side(style='medium', color='1F497D'))

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

def title_block(ws, title, subtitle, ncols):
    ws["A1"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) — GÓI THẦU SỐ 09-XL"
    ws["A1"].font = FONT_SUBTITLE
    ws["A2"] = title
    ws["A2"].font = FONT_TITLE
    ws["A3"] = subtitle
    ws["A3"].font = FONT_SUBTITLE
    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)

def header_row(ws, row, headers, widths=None):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row, c, h)
        cell.font, cell.fill, cell.alignment, cell.border = FONT_HDR, FILL_HDR, ALIGN_CENTER, THIN_BORDER
    ws.row_dimensions[row].height = 28
    if widths:
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row + 1, 1)

def put(ws, row, col, value, fmt=None, font=FONT_REG, align=None, fill=None, border=THIN_BORDER):
    cell = ws.cell(row, col, value)
    cell.font = font
    cell.border = border
    cell.alignment = align or (ALIGN_RIGHT if isinstance(value, (int, float)) or str(value).startswith("=") else ALIGN_LEFT)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell

def clone_worksheet(src, dst):
    """Sao chép toàn diện 1 worksheet từ workbook nguồn sang đích giữ nguyên vẹn 100% định dạng."""
    dst.views.sheetView[0].showGridLines = src.views.sheetView[0].showGridLines
    for col_letter, col_dim in src.column_dimensions.items():
        dst.column_dimensions[col_letter].width = col_dim.width
        dst.column_dimensions[col_letter].hidden = col_dim.hidden
    for row_idx, row_dim in src.row_dimensions.items():
        dst.row_dimensions[row_idx].height = row_dim.height
    for row in src.iter_rows():
        for cell in row:
            new_cell = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    for merge_range in src.merged_cells.ranges:
        dst.merge_cells(str(merge_range))

def extract_master_eval_cache(master_path):
    """Trích xuất ma trận giá trị tính toán sạch 100% từ Microsoft Excel COM Engine."""
    import win32com.client
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(master_path), UpdateLinks=0, ReadOnly=True)
        excel.CalculateFullRebuild()
        cache = {}
        for ws in wb.Worksheets:
            rng = ws.UsedRange
            val = rng.Value
            if not isinstance(val, tuple):
                val = ((val,),)
            elif val and not isinstance(val[0], tuple):
                val = (val,)
            cache[ws.Name] = {
                'row_start': rng.Row,
                'col_start': rng.Column,
                'values': val
            }
        wb.Close(False)
        return cache
    finally:
        excel.Quit()

def sanitize_workbook_formulas(wb, eval_cache):
    """
    Tự động rà soát và khử toàn bộ lỗi liên kết chéo (#REF!, #VALUE!) trong các tệp vi mô độc lập:
    - Nếu công thức tham chiếu đến một worksheet KHÔNG TỒN TẠI trong file này: Thay thế bằng giá trị số học đã tính toán chính xác từ Master.
    - Nếu công thức nội bộ (cùng sheet hoặc trỏ sang sheet CÓ MẶT trong file này): Giữ nguyên 100% công thức sống.
    """
    import re
    available_sheets = set(wb.sheetnames)
    for sname in wb.sheetnames:
        ws = wb[sname]
        cache_s = eval_cache.get(sname)
        if not cache_s:
            continue
        values_matrix = cache_s['values']
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    raw_refs = re.findall(r"(?:'([^']+)'|([A-Za-z0-9_]+))!", val)
                    referenced_sheets = set(r[0] if r[0] else r[1] for r in raw_refs)
                    missing = referenced_sheets - available_sheets
                    if missing:
                        r_idx = cell.row - cache_s['row_start']
                        c_idx = cell.column - cache_s['col_start']
                        if 0 <= r_idx < len(values_matrix):
                            row_vals = values_matrix[r_idx]
                            if 0 <= c_idx < len(row_vals):
                                eval_val = row_vals[c_idx]
                                if eval_val is not None:
                                    cell.value = eval_val

# =============================================================================
# XÂY DỰNG GÓI 01: VĨ MÔ / MASTER ĐIỀU HÀNH
# =============================================================================
def build_package_macro(dest_dir):
    print(f"[*] Đang đóng gói GÓI 01 (VĨ MÔ / MASTER) vào: {dest_dir}...")
    
    # 1. File Master Excel 14 Sheet
    dst_excel = os.path.join(dest_dir, "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")
    shutil.copyfile(MASTER_SOURCE, dst_excel)
    
    # 2. File Tiến độ MS Project XML & MPP
    src_xml = os.path.join(TEMPLATES_DIR, "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml")
    if os.path.exists(src_xml):
        shutil.copyfile(src_xml, os.path.join(dest_dir, "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml"))
        
    src_mpp = os.path.join(TEMPLATES_DIR, "Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp")
    if os.path.exists(src_mpp):
        shutil.copyfile(src_mpp, os.path.join(dest_dir, "Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp"))
        
    # 3. File Hồ sơ KCS Word .docx
    src_docx = os.path.join(TEMPLATES_DIR, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx")
    if os.path.exists(src_docx):
        shutil.copyfile(src_docx, os.path.join(dest_dir, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx"))
        
    # 4. Báo cáo Thẩm tra AEC Audit & Thuyết minh biện pháp thi công
    src_audit = os.path.join(TEMPLATES_DIR, "BAO_CAO_THAM_TRA_AEC_AUDIT.md")
    if os.path.exists(src_audit):
        shutil.copyfile(src_audit, os.path.join(dest_dir, "BAO_CAO_THAM_TRA_AEC_AUDIT.md"))
        
    src_bptc = os.path.join(TEMPLATES_DIR, "Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md")
    if os.path.exists(src_bptc):
        shutil.copyfile(src_bptc, os.path.join(dest_dir, "Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md"))
        
    print(f"-> GÓI 01 hoàn tất tại: {dest_dir}")


# =============================================================================
# XÂY DỰNG GÓI 02: VI MÔ / CHUYÊN SÂU SẢN XUẤT (14 BỘ HỒ SƠ)
# =============================================================================

def build_dossier_01(wb_master, dest_dir):
    print("  [01/14] Xây dựng Dossier 01: Tổ hợp cắt thép 11.7m chuẩn RebarCut Pro (OR-Tools CP-SAT)...")
    dst_xlsx = os.path.join(dest_dir, "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx")
    
    from tools.bbs_loader import load_bbs
    from tools.cutting_stock_solver import CuttingStockSolver
    from tools.rebarcut_export import write_rebarcut_workbook
    
    res = load_bbs(MASTER_SOURCE, sheet="THONG_KE_THEP_CHI_TIET")
    solver = CuttingStockSolver(bar_length_mm=11700, kerf_mm=5, end_trim_mm=0, time_limit_s=10.0)
    sol_master = solver.solve(res.demands, split_long_bars=True)
    write_rebarcut_workbook(dst_xlsx, res.demands, sol_master)
    
    # Xuất file CSV lệnh cắt thép CNC kèm theo
    dst_csv = os.path.join(dest_dir, "01_Phieu_Cat_Thep_Cau_Km19+529.080.csv")
    with open(dst_csv, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Ma_Cay", "Duong_Kinh_mm", "Bar_Mark", "Dai_Cat_m", "De_Xe_m", "Cay_Goc_m"])
        for a in sol_master.assignment:
            bar_name = f"C{a['bar_id']}"
            for length, mark in zip(a["cuts_mm"], a["marks"]):
                writer.writerow([bar_name, a["diameter_mm"], mark, length / 1000.0, a["waste_mm"] / 1000.0, 11.7])
    return dst_xlsx

def build_dossier_02(wb_master, eval_cache, dest_dir):
    print("  [02/14] Xây dựng Dossier 02: Bóc tách khối lượng đào đắp trình diễn (M1, M2, T1, T2)...")
    dst_xlsx = os.path.join(dest_dir, "02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: KHOI_LUONG_DAO_DAP
    clone_worksheet(wb_master["KHOI_LUONG_DAO_DAP"], wb.active)
    wb.active.title = "KHOI_LUONG_DAO_DAP"
    
    # Sheet 2: CHI_TIET_DAO_MONG_MO_TRU
    ws2 = wb.create_sheet(title="CHI_TIET_DAO_MONG_MO_TRU")
    ws2.views.sheetView[0].showGridLines = True
    headers2 = ["STT", "Hố móng cấu kiện", "Lý trình", "Kích thước Dài (m)", "Kích thước Rộng (m)", "Chiều sâu TB (m)", "Độ dốc taluy", "Khối lượng Đất C3 (m3)", "Khối lượng Đá C3 (m3)", "Biện pháp thi công"]
    title_block(ws2, "DIỄN GIẢI KÍCH THƯỚC HÌNH HỌC HỐ MÓNG MỐ M1, M2 & BỆ TRỤ T1, T2",
                "Hồ sơ thiết kế BVTC; Bệ móng cọc D1200 ngàm sâu vào tầng cuội sỏi / đá phong hóa", len(headers2))
    header_row(ws2, 5, headers2, [6, 28, 14, 16, 16, 16, 14, 18, 18, 30])
    
    hm_items = [
        (1, "Hố móng bệ mố M1 (chân dê)", "Km19+463.88", 12.00, 7.50, 4.20, 1.0, "=D6*E6*1.8*G6", "=D6*E6*(F6-1.8)*G6", "Máy đào 1.25m3 kết hợp búa đập đá"),
        (2, "Hố móng bệ trụ T1", "Km19+502.98", 15.00, 9.50, 4.80, 1.0, "=D7*E7*2.2*G7", "=D7*E7*(F7-2.2)*G7", "Máy đào gầu sâu hạ tầng phủ đá"),
        (3, "Hố móng bệ trụ T2", "Km19+542.98", 15.00, 9.50, 4.50, 1.0, "=D8*E8*2.0*G8", "=D8*E8*(F8-2.0)*G8", "Máy đào gầu sâu hạ tầng phủ đá"),
        (4, "Hố móng bệ mố M2 (chữ U)", "Km19+594.28", 14.00, 8.50, 4.60, 1.0, "=D9*E9*2.1*G9", "=D9*E9*(F9-2.1)*G9", "Máy đào kết hợp phá đá ngàm móng"),
    ]
    r = 6
    for it in hm_items:
        put(ws2, r, 1, it[0], align=ALIGN_CENTER)
        put(ws2, r, 2, it[1])
        put(ws2, r, 3, it[2], align=ALIGN_CENTER)
        put(ws2, r, 4, it[3], "#,##0.00")
        put(ws2, r, 5, it[4], "#,##0.00")
        put(ws2, r, 6, it[5], "#,##0.00")
        put(ws2, r, 7, it[6], "#,##0.00", align=ALIGN_CENTER)
        put(ws2, r, 8, it[7], "#,##0.000", font=FONT_BOLD)
        put(ws2, r, 9, it[8], "#,##0.000", font=FONT_BOLD)
        put(ws2, r, 10, it[9])
        r += 1
    put(ws2, r, 2, "TỔNG CỘNG ĐÀO HỐ MÓNG", font=FONT_BOLD)
    put(ws2, r, 8, f"=SUM(H6:H{r-1})", "#,##0.000", font=FONT_BOLD, border=DOUBLE_BOTTOM_BORDER)
    put(ws2, r, 9, f"=SUM(I6:I{r-1})", "#,##0.000", font=FONT_BOLD, border=DOUBLE_BOTTOM_BORDER)
    
    # Sheet 3: CHI_TIET_DAP_DAU_CAU_K95_K98
    ws3 = wb.create_sheet(title="CHI_TIET_DAP_DAU_CAU_K95_K98")
    ws3.views.sheetView[0].showGridLines = True
    headers3 = ["STT", "Hạng mục đắp", "Vị trí áp dụng", "Chiều dày lớp đầm (cm)", "Diện tích TB (m2)", "Chiều cao đắp (m)", "Độ chặt K yêu cầu", "Thể tích đắp (m3)", "Thiết bị lu lèn"]
    title_block(ws3, "KẾ HOẠCH ĐẮP ĐẤT CHỌN LỌC ĐẦU CẦU & SAU MỐ — CẦU KM19+529.080",
                "TCVN 9436:2012; Đầm cóc Mikasa 80kg trong phạm vi 2.0m sau mang mố, lu rung 10T ngoài phạm vi", len(headers3))
    header_row(ws3, 5, headers3, [6, 28, 20, 16, 16, 16, 16, 18, 30])
    
    dap_items = [
        (1, "Lòng mố chữ U mố M2 - Đợt 1", "Mố M2 (Km19+594.28)", 15, 68.5, 2.20, "K >= 0.98", "=E6*F6", "Đầm cóc Mikasa 80kg"),
        (2, "Lòng mố chữ U mố M2 - Đợt 2", "Mố M2 (Km19+594.28)", 15, 65.0, 2.40, "K >= 0.98", "=E7*F7", "Đầm cóc Mikasa 80kg"),
        (3, "Đắp sau mố M1 đoạn quá độ", "Sau mố M1 L=20m", 20, 110.0, 3.80, "K >= 0.95", "=E8*F8", "Lu rung 10T kết hợp lu nhỏ"),
        (4, "Đắp sau mố M2 đoạn quá độ", "Sau mố M2 L=20m", 20, 115.0, 4.20, "K >= 0.95", "=E9*F9", "Lu rung 10T kết hợp lu nhỏ"),
        (5, "Đắp đất tạo mái taluy tứ nón M1, M2", "Hai đầu cầu", 20, 85.0, 3.50, "K >= 0.95", "=E10*F10", "Đầm cóc kết hợp vỗ mái"),
    ]
    r = 6
    for d in dap_items:
        put(ws3, r, 1, d[0], align=ALIGN_CENTER)
        put(ws3, r, 2, d[1])
        put(ws3, r, 3, d[2], align=ALIGN_CENTER)
        put(ws3, r, 4, d[3], "0", align=ALIGN_CENTER)
        put(ws3, r, 5, d[4], "#,##0.00")
        put(ws3, r, 6, d[5], "#,##0.00")
        put(ws3, r, 7, d[6], align=ALIGN_CENTER)
        put(ws3, r, 8, d[7], "#,##0.000", font=FONT_BOLD)
        put(ws3, r, 9, d[8])
        r += 1
    put(ws3, r, 2, "TỔNG CỘNG THỂ TÍCH ĐẤT ĐẮP", font=FONT_BOLD)
    put(ws3, r, 8, f"=SUM(H6:H{r-1})", "#,##0.000", font=FONT_BOLD, border=DOUBLE_BOTTOM_BORDER)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_03(wb_master, eval_cache, dest_dir):
    print("  [03/14] Xây dựng Dossier 03: QS Tiên lượng hình học chi tiết Takeoff (101 dòng)...")
    dst_xlsx = os.path.join(dest_dir, "03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: QS_DIEN_GIAI_CHI_TIET
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], wb.active)
    wb.active.title = "QS_DIEN_GIAI_CHI_TIET"
    
    # Sheet 2: KHOI_LUONG_DAO_DAP
    ws2 = wb.create_sheet(title="KHOI_LUONG_DAO_DAP")
    clone_worksheet(wb_master["KHOI_LUONG_DAO_DAP"], ws2)
    
    # Sheet 3: THONG_KE_THEP_CHI_TIET
    ws3 = wb.create_sheet(title="THONG_KE_THEP_CHI_TIET")
    clone_worksheet(wb_master["THONG_KE_THEP_CHI_TIET"], ws3)
    
    # Sheet 4: CHI_TIET_SUPER_T_3X38M2
    ws4 = wb.create_sheet(title="CHI_TIET_SUPER_T_3X38M2")
    ws4.views.sheetView[0].showGridLines = True
    headers4 = ["STT", "Hạng mục kết cấu dầm", "Số lượng", "Chiều dài L (m)", "Diện tích MC (m2)", "Thể tích BT (m3)", "Chu vi VK (m)", "Diện tích VK (m2)", "Khối lượng cáp DUL (kg)", "Ghi chú kỹ thuật"]
    title_block(ws4, "BÓC TÁCH HÌNH HỌC CHI TIẾT 15 PHIẾN DẦM SUPER-T L=38.2M — CẦU KM19+529.080",
                "Sơ đồ 3 nhịp: 3 x 38.2m = 114.6m; Mỗi nhịp 5 phiến dầm Super-T mác C45/50", len(headers4))
    header_row(ws4, 5, headers4, [6, 28, 10, 14, 16, 16, 14, 16, 20, 30])
    
    st_items = [
        (1, "Dầm biên nhịp 1 (Nhịp M1-T1)", 2, 38.20, 0.912, "=C6*D6*E6", 8.45, "=C6*D6*G6", 2450.0, "Dầm Super-T biên có gờ"),
        (2, "Dầm giữa nhịp 1 (Nhịp M1-T1)", 3, 38.20, 0.912, "=C7*D7*E7", 8.45, "=C7*D7*G7", 3675.0, "Dầm Super-T điển hình"),
        (3, "Dầm biên nhịp 2 (Nhịp T1-T2)", 2, 38.20, 0.912, "=C8*D8*E8", 8.45, "=C8*D8*G8", 2450.0, "Dầm Super-T biên có gờ"),
        (4, "Dầm giữa nhịp 2 (Nhịp T1-T2)", 3, 38.20, 0.912, "=C9*D9*E9", 8.45, "=C9*D9*G9", 3675.0, "Dầm Super-T điển hình"),
        (5, "Dầm biên nhịp 3 (Nhịp T2-M2)", 2, 38.20, 0.912, "=C10*D10*E10", 8.45, "=C10*D10*G10", 2450.0, "Dầm Super-T biên có gờ"),
        (6, "Dầm giữa nhịp 3 (Nhịp T2-M2)", 3, 38.20, 0.912, "=C11*D11*E11", 8.45, "=C11*D11*G11", 3675.0, "Dầm Super-T điển hình"),
        (7, "Dầm ngang mố M1, M2", 2, 11.20, 0.650, "=C12*D12*E12", 3.20, "=C12*D12*G12", 0.0, "Bê tông C35 đổ bù tại chỗ"),
        (8, "Dầm ngang trụ T1, T2", 2, 11.20, 0.850, "=C13*D13*E13", 3.60, "=C13*D13*G13", 0.0, "Bê tông C35 liên tục nhiệt"),
    ]
    r = 6
    for st in st_items:
        put(ws4, r, 1, st[0], align=ALIGN_CENTER)
        put(ws4, r, 2, st[1])
        put(ws4, r, 3, st[2], "0", align=ALIGN_CENTER)
        put(ws4, r, 4, st[3], "#,##0.00")
        put(ws4, r, 5, st[4], "#,##0.000")
        put(ws4, r, 6, st[5], "#,##0.000", font=FONT_BOLD)
        put(ws4, r, 7, st[6], "#,##0.00")
        put(ws4, r, 8, st[7], "#,##0.00", font=FONT_BOLD)
        put(ws4, r, 9, st[8], "#,##0.0", font=FONT_BOLD)
        put(ws4, r, 10, st[9])
        r += 1
    put(ws4, r, 2, "TỔNG CỘNG 15 PHIẾN DẦM SUPER-T", font=FONT_BOLD)
    put(ws4, r, 6, f"=SUM(F6:F{r-1})", "#,##0.000", font=FONT_BOLD)
    put(ws4, r, 8, f"=SUM(H6:H{r-1})", "#,##0.00", font=FONT_BOLD)
    put(ws4, r, 9, f"=SUM(I6:I{r-1})", "#,##0.0", font=FONT_BOLD, border=DOUBLE_BOTTOM_BORDER)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_04(wb_master, eval_cache, dest_dir):
    print("  [04/14] Xây dựng Dossier 04: Thống kê cốt thép chi tiết BBS (396 dòng cốt thép)...")
    dst_xlsx = os.path.join(dest_dir, "04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: THONG_KE_THEP_CHI_TIET
    clone_worksheet(wb_master["THONG_KE_THEP_CHI_TIET"], wb.active)
    wb.active.title = "THONG_KE_THEP_CHI_TIET"
    
    # Sheet 2: TONG_HOP_THEP_THEO_DUONG_KINH
    ws2 = wb.create_sheet(title="TONG_HOP_THEP_THEO_DUONG_KINH")
    ws2.views.sheetView[0].showGridLines = True
    headers2 = ["STT", "Đường kính Ø (mm)", "Mác thép tiêu chuẩn", "Trọng lượng đơn vị (kg/m)", "Tổng chiều dài (m)", "Tổng trọng lượng (kg)", "Tổng trọng lượng (Tấn)", "Tỷ trọng (%)", "Quy cách đóng gói"]
    title_block(ws2, "BẢNG TỔNG HỢP VẬT TƯ CỐT THÉP THEO TỪNG ĐƯỜNG KÍNH Ø — CẦU KM19+529.080",
                "Trích xuất từ 396 thanh cốt thép BBS theo TCVN 1651:2018 và ASTM A416", len(headers2))
    header_row(ws2, 5, headers2, [6, 18, 22, 20, 20, 22, 20, 16, 25])
    
    bar_sizes = [
        (1, 10, "CB240-T", 0.617, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 10, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 10, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F6/1000", "=F6/$F$17", "Thép cuộn"),
        (2, 12, "CB400-V", 0.888, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 12, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 12, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F7/1000", "=F7/$F$17", "Cây 11.7m"),
        (3, 14, "CB400-V", 1.208, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 14, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 14, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F8/1000", "=F8/$F$17", "Cây 11.7m"),
        (4, 16, "CB400-V", 1.578, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F9/1000", "=F9/$F$17", "Cây 11.7m"),
        (5, 18, "CB400-V", 2.000, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 18, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 18, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F10/1000", "=F10/$F$17", "Cây 11.7m"),
        (6, 20, "CB400-V", 2.466, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 20, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 20, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F11/1000", "=F11/$F$17", "Cây 11.7m"),
        (7, 22, "CB400-V", 2.984, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 22, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 22, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F12/1000", "=F12/$F$17", "Cây 11.7m"),
        (8, 25, "CB400-V", 3.853, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 25, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 25, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F13/1000", "=F13/$F$17", "Cây 11.7m"),
        (9, 28, "CB400-V", 4.834, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 28, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 28, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F14/1000", "=F14/$F$17", "Cây 11.7m"),
        (10, 32, "CB400-V", 6.313, "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 32, THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 32, THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F15/1000", "=F15/$F$17", "Cây 11.7m"),
        (11, 15.2, "Cáp DƯL", 1.101, "=SUMIF(THONG_KE_THEP_CHI_TIET!$P$6:$P$399, \"Cáp DƯL 15.2mm\", THONG_KE_THEP_CHI_TIET!$L$6:$L$399)", "=SUMIF(THONG_KE_THEP_CHI_TIET!$P$6:$P$399, \"Cáp DƯL 15.2mm\", THONG_KE_THEP_CHI_TIET!$N$6:$N$399)", "=F16/1000", "=F16/$F$17", "Cuộn cáp 7 sợi"),
    ]
    r = 6
    for b in bar_sizes:
        put(ws2, r, 1, b[0], align=ALIGN_CENTER)
        put(ws2, r, 2, f"Ø{b[1]}" if isinstance(b[1], (int, float)) else str(b[1]), align=ALIGN_CENTER)
        put(ws2, r, 3, b[2], align=ALIGN_CENTER)
        put(ws2, r, 4, b[3], "#,##0.000")
        put(ws2, r, 5, b[4], "#,##0.00")
        put(ws2, r, 6, b[5], "#,##0.0", font=FONT_BOLD)
        put(ws2, r, 7, b[6], "#,##0.000", font=FONT_BOLD)
        put(ws2, r, 8, b[7], "0.00%", align=ALIGN_CENTER)
        put(ws2, r, 9, b[8])
        r += 1
    put(ws2, r, 2, "TỔNG CỘNG TOÀN CÔNG TRÌNH", font=FONT_BOLD)
    put(ws2, r, 6, f"=SUM(F6:F{r-1})", "#,##0.0", font=FONT_BOLD)
    put(ws2, r, 7, f"=SUM(G6:G{r-1})", "#,##0.000", font=FONT_BOLD, border=DOUBLE_BOTTOM_BORDER)
    put(ws2, r, 8, f"=SUM(H6:H{r-1})", "0.00%", font=FONT_BOLD, align=ALIGN_CENTER)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_05(wb_master, eval_cache, dest_dir):
    print("  [05/14] Xây dựng Dossier 05: Cấp phối 1m3 & Kế hoạch tần suất thí nghiệm (809 mẫu QA/QC)...")
    dst_xlsx = os.path.join(dest_dir, "05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: CAP_PHOI_1M3_VA_TAN_SUAT
    clone_worksheet(wb_master["CAP_PHOI_1M3_VA_TAN_SUAT"], wb.active)
    wb.active.title = "CAP_PHOI_1M3_VA_TAN_SUAT"
    
    # Sheet 2: QS_DIEN_GIAI_CHI_TIET
    ws2 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], ws2)
    
    # Sheet 3: THONG_KE_THEP_CHI_TIET
    ws3 = wb.create_sheet(title="THONG_KE_THEP_CHI_TIET")
    clone_worksheet(wb_master["THONG_KE_THEP_CHI_TIET"], ws3)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_06(wb_master, eval_cache, dest_dir):
    print("  [06/14] Xây dựng Dossier 06: Phân tích vật tư chi tiết WBS (140 dòng phân tích)...")
    dst_xlsx = os.path.join(dest_dir, "06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: PHAN_TICH_VAT_TU_WBS
    clone_worksheet(wb_master["PHAN_TICH_VAT_TU_WBS"], wb.active)
    wb.active.title = "PHAN_TICH_VAT_TU_WBS"
    
    # Sheet 2: QS_DIEN_GIAI_CHI_TIET
    ws2 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], ws2)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_07(wb_master, eval_cache, dest_dir):
    print("  [07/14] Xây dựng Dossier 07: Tổng hợp nhu cầu vật tư BOM & Kế hoạch cung ứng 4 giai đoạn...")
    dst_xlsx = os.path.join(dest_dir, "07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: TONG_HOP_VAT_TU_TOAN_BO
    clone_worksheet(wb_master["TONG_HOP_VAT_TU_TOAN_BO"], wb.active)
    wb.active.title = "TONG_HOP_VAT_TU_TOAN_BO"
    
    # Sheet 2: PHAN_TICH_VAT_TU_WBS
    ws2 = wb.create_sheet(title="PHAN_TICH_VAT_TU_WBS")
    clone_worksheet(wb_master["PHAN_TICH_VAT_TU_WBS"], ws2)
    
    # Sheet 3: CUNG_UNG_4_GIAI_DOAN
    ws3 = wb.create_sheet(title="CUNG_UNG_4_GIAI_DOAN")
    ws3.views.sheetView[0].showGridLines = True
    headers3 = ["STT", "Chủng loại vật tư chính", "ĐVT", "Tổng nhu cầu công trình", "Giai đoạn 1: Móng cọc & Bệ móng", "Giai đoạn 2: Thân mố, Thân trụ & Đúc dầm Super-T", "Giai đoạn 3: Lao dầm & Bản mặt cầu", "Giai đoạn 4: Khe co giãn & Hoàn thiện", "Quy định dự trữ kho an toàn"]
    title_block(ws3, "KẾ HOẠCH CUNG ỨNG VÀ DỰ TRỮ VẬT TƯ THEO 4 GIAI ĐOÀN THI CÔNG — CẦU KM19+529.080",
                "Phù hợp tiến độ CPM đường găng; Đảm bảo không bị gián đoạn thi công do thời tiết vùng cao", len(headers3))
    header_row(ws3, 5, headers3, [6, 28, 8, 18, 20, 22, 20, 20, 25])
    
    phases = [
        (1, "Xi măng PCB40", "Tấn", "=TONG_HOP_VAT_TU_TOAN_BO!G23", "=D6*0.35", "=D6*0.35", "=D6*0.20", "=D6*0.10", "Tồn kho tối thiểu 50 tấn tại trạm trộn"),
        (2, "Cát vàng bê tông", "m3", "=TONG_HOP_VAT_TU_TOAN_BO!G24", "=D7*0.35", "=D7*0.35", "=D7*0.20", "=D7*0.10", "Dự trữ bãi chứa 100 m3 chống mưa lũ"),
        (3, "Đá dăm 1x2 bê tông", "m3", "=TONG_HOP_VAT_TU_TOAN_BO!G25", "=D8*0.35", "=D8*0.35", "=D8*0.20", "=D8*0.10", "Dự trữ bãi chứa 150 m3 sàng sạch"),
        (4, "Cốt thép CB400-V (các Ø)", "kg", "=TONG_HOP_VAT_TU_TOAN_BO!G10", "=D9*0.40", "=D9*0.35", "=D9*0.20", "=D9*0.05", "Gia công sẵn lồng cọc & dầm tại bãi"),
        (5, "Cáp DƯL tao xoắn 15.2mm", "kg", "=TONG_HOP_VAT_TU_TOAN_BO!G11", 0, "=D10*0.80", "=D10*0.20", 0, "Bảo quản phòng khô, tra mỡ bảo vệ"),
        (6, "Ván khuôn thép định hình", "m2", 1500, 600, 600, 300, 0, "Luân chuyển tuần hoàn thân trụ & dầm"),
        (7, "Gối chậu cao su (30 bộ)", "Bộ", 30, 0, 10, 20, 0, "Nhập đủ lô trước khi đúc dầm nhịp 1"),
        (8, "Khe co giãn răng lược (2 bộ)", "Bộ", 2, 0, 0, 0, 2, "Gia công chuẩn theo bản vẽ cơ khí"),
        (9, "Bê tông nhựa C16 mặt cầu", "Tấn", "=TONG_HOP_VAT_TU_TOAN_BO!G32", 0, 0, 0, "=D14", "Trạm trộn thảm liên tục trong 2 ngày"),
    ]
    r = 6
    for p in phases:
        put(ws3, r, 1, p[0], align=ALIGN_CENTER)
        put(ws3, r, 2, p[1])
        put(ws3, r, 3, p[2], align=ALIGN_CENTER)
        put(ws3, r, 4, p[3], "#,##0.00" if isinstance(p[3], (int, float)) or str(p[3]).startswith("=") else "@", font=FONT_BOLD)
        put(ws3, r, 5, p[4], "#,##0.00" if isinstance(p[4], (int, float)) or str(p[4]).startswith("=") else "@")
        put(ws3, r, 6, p[5], "#,##0.00" if isinstance(p[5], (int, float)) or str(p[5]).startswith("=") else "@")
        put(ws3, r, 7, p[6], "#,##0.00" if isinstance(p[6], (int, float)) or str(p[6]).startswith("=") else "@")
        put(ws3, r, 8, p[7], "#,##0.00" if isinstance(p[7], (int, float)) or str(p[7]).startswith("=") else "@")
        put(ws3, r, 9, p[8])
        r += 1
        
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_08(wb_master, eval_cache, dest_dir):
    print("  [08/14] Xây dựng Dossier 08: Dự toán chi phí xây dựng G_XD (Thông tư 36/2026/TT-BXD)...")
    dst_xlsx = os.path.join(dest_dir, "08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: TONG_HOP_DU_TOAN_GXD
    clone_worksheet(wb_master["TONG_HOP_DU_TOAN_GXD"], wb.active)
    wb.active.title = "TONG_HOP_DU_TOAN_GXD"
    
    # Sheet 2: QS_DIEN_GIAI_CHI_TIET
    ws2 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], ws2)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_09(wb_master, eval_cache, dest_dir):
    print("  [09/14] Xây dựng Dossier 09: Bảng xác định giá trị thanh toán Phụ lục 03.a (NĐ 254/2025)...")
    dst_xlsx = os.path.join(dest_dir, "09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: THANH_TOAN_KY_PHU_LUC_03A
    clone_worksheet(wb_master["THANH_TOAN_KY_PHU_LUC_03A"], wb.active)
    wb.active.title = "THANH_TOAN_KY_PHU_LUC_03A"
    
    # Sheet 2: QS_DIEN_GIAI_CHI_TIET
    ws2 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], ws2)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_10(wb_master, eval_cache, dest_dir):
    print("  [10/14] Xây dựng Dossier 10: Tiến độ thi công CPM & Biểu đồ Gantt (36 WBS, XML, MPP, CSV)...")
    dst_xlsx = os.path.join(dest_dir, "10_Tien_Do_Thi_Cong_CPM_Gantt_Chart.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: TIEN_DO_THI_CONG_WBS
    clone_worksheet(wb_master["TIEN_DO_THI_CONG_WBS"], wb.active)
    wb.active.title = "TIEN_DO_THI_CONG_WBS"
    
    # Sheet 2: QS_DIEN_GIAI_CHI_TIET
    ws2 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], ws2)
    
    # Sheet 3: KHOI_LUONG_DAO_DAP
    ws3 = wb.create_sheet(title="KHOI_LUONG_DAO_DAP")
    clone_worksheet(wb_master["KHOI_LUONG_DAO_DAP"], ws3)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    
    # Sao chép các file vệ tinh XML & MPP
    src_xml = os.path.join(TEMPLATES_DIR, "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml")
    if os.path.exists(src_xml):
        shutil.copyfile(src_xml, os.path.join(dest_dir, "10_Tien_Do_Thi_Cong_Cau_Km19+529.080.xml"))
        
    src_mpp = os.path.join(TEMPLATES_DIR, "Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp")
    if os.path.exists(src_mpp):
        shutil.copyfile(src_mpp, os.path.join(dest_dir, "10_Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp"))
        
    # Xuất file CSV tiến độ CPM
    dst_cpm_csv = os.path.join(dest_dir, "10_Tien_Do_Thi_Cong_Cau_Km19+529.080.csv")
    ws_cpm = wb_master["TIEN_DO_THI_CONG_WBS"]
    with open(dst_cpm_csv, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Ma_WBS", "Ten_Cong_Tac", "Khoi_Luong", "DVT", "Dinh_Muc_Cong", "Tong_Cong", "To_Doi", "Thoi_Gian_Ngay", "Ngay_Bat_Dau", "Ngay_Hoan_Thanh", "Quan_He_Logic", "Duong_Gang_CPM"])
        for r in range(6, ws_cpm.max_row + 1):
            wbs = ws_cpm.cell(r, 1).value
            name = ws_cpm.cell(r, 2).value
            if wbs or name:
                row_data = [ws_cpm.cell(r, c).value for c in range(1, 13)]
                writer.writerow(row_data)
                
    return dst_xlsx

def build_dossier_11(wb_master, eval_cache, dest_dir):
    print("  [11/14] Xây dựng Dossier 11: Danh mục KCS 22 Biên bản nghiệm thu (NĐ 207/2026/NĐ-CP, kèm Word .docx)...")
    dst_xlsx = os.path.join(dest_dir, "11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: HOSO_KCS_NGHIEM_THU
    clone_worksheet(wb_master["HOSO_KCS_NGHIEM_THU"], wb.active)
    wb.active.title = "HOSO_KCS_NGHIEM_THU"
    
    # Sheet 2: QS_DIEN_GIAI_CHI_TIET
    ws2 = wb.create_sheet(title="QS_DIEN_GIAI_CHI_TIET")
    clone_worksheet(wb_master["QS_DIEN_GIAI_CHI_TIET"], ws2)
    
    # Sheet 3: KHOI_LUONG_DAO_DAP
    ws3 = wb.create_sheet(title="KHOI_LUONG_DAO_DAP")
    clone_worksheet(wb_master["KHOI_LUONG_DAO_DAP"], ws3)
    
    # Sheet 4: TIEN_DO_THI_CONG_WBS
    ws4 = wb.create_sheet(title="TIEN_DO_THI_CONG_WBS")
    clone_worksheet(wb_master["TIEN_DO_THI_CONG_WBS"], ws4)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    
    # Sao chép file Word 22 biên bản
    src_docx = os.path.join(TEMPLATES_DIR, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx")
    if os.path.exists(src_docx):
        shutil.copyfile(src_docx, os.path.join(dest_dir, "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx"))
        
    return dst_xlsx

def build_dossier_12(wb_master, eval_cache, dest_dir):
    print("  [12/14] Xây dựng Dossier 12: Mẫu in A4 Biên bản nghiệm thu công việc xây dựng...")
    dst_xlsx = os.path.join(dest_dir, "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: MAU_BIEN_BAN_KCS
    clone_worksheet(wb_master["MAU_BIEN_BAN_KCS"], wb.active)
    wb.active.title = "MAU_BIEN_BAN_KCS"
    
    # Sheet 2: HOSO_KCS_NGHIEM_THU (Nhúng để VLOOKUP nội bộ không bao giờ bị #REF!)
    ws2 = wb.create_sheet(title="HOSO_KCS_NGHIEM_THU")
    clone_worksheet(wb_master["HOSO_KCS_NGHIEM_THU"], ws2)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_13(wb_master, eval_cache, dest_dir):
    print("  [13/14] Xây dựng Dossier 13: Mẫu in A4 Biên bản nghiệm thu vật liệu xây dựng đầu vào...")
    dst_xlsx = os.path.join(dest_dir, "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: MAU_BB_NGHIEM_THU_VAT_LIEU
    clone_worksheet(wb_master["MAU_BB_NGHIEM_THU_VAT_LIEU"], wb.active)
    wb.active.title = "MAU_BB_NGHIEM_THU_VAT_LIEU"
    
    # Sheet 2: CAP_PHOI_1M3_VA_TAN_SUAT (Nhúng để VLOOKUP nội bộ không bao giờ bị #REF!)
    ws2 = wb.create_sheet(title="CAP_PHOI_1M3_VA_TAN_SUAT")
    clone_worksheet(wb_master["CAP_PHOI_1M3_VA_TAN_SUAT"], ws2)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_dossier_14(wb_master, eval_cache, dest_dir):
    print("  [14/14] Xây dựng Dossier 14: Mẫu in A4 Biên bản lấy mẫu & Nén mẫu bê tông R7, R28...")
    dst_xlsx = os.path.join(dest_dir, "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx")
    wb = openpyxl.Workbook()
    
    # Sheet 1: MAU_BB_LAY_MAU_HIEN_TRUONG
    clone_worksheet(wb_master["MAU_BB_LAY_MAU_HIEN_TRUONG"], wb.active)
    wb.active.title = "MAU_BB_LAY_MAU_HIEN_TRUONG"
    
    # Sheet 2: CAP_PHOI_1M3_VA_TAN_SUAT (Nhúng để tính toán tự động ngày nén và quy cách mẫu)
    ws2 = wb.create_sheet(title="CAP_PHOI_1M3_VA_TAN_SUAT")
    clone_worksheet(wb_master["CAP_PHOI_1M3_VA_TAN_SUAT"], ws2)
    
    sanitize_workbook_formulas(wb, eval_cache)
    wb.save(dst_xlsx)
    return dst_xlsx

def build_package_micro(wb_master, eval_cache, dest_dir):
    print(f"[*] Đang xuất toàn diện GÓI 02 (14 BỘ HỒ SƠ VI MÔ CHUYÊN SÂU) vào: {dest_dir}...")
    build_dossier_01(wb_master, dest_dir)
    build_dossier_02(wb_master, eval_cache, dest_dir)
    build_dossier_03(wb_master, eval_cache, dest_dir)
    build_dossier_04(wb_master, eval_cache, dest_dir)
    build_dossier_05(wb_master, eval_cache, dest_dir)
    build_dossier_06(wb_master, eval_cache, dest_dir)
    build_dossier_07(wb_master, eval_cache, dest_dir)
    build_dossier_08(wb_master, eval_cache, dest_dir)
    build_dossier_09(wb_master, eval_cache, dest_dir)
    build_dossier_10(wb_master, eval_cache, dest_dir)
    build_dossier_11(wb_master, eval_cache, dest_dir)
    build_dossier_12(wb_master, eval_cache, dest_dir)
    build_dossier_13(wb_master, eval_cache, dest_dir)
    build_dossier_14(wb_master, eval_cache, dest_dir)
    print(f"-> GÓI 02 hoàn tất với đầy đủ 14 bộ hồ sơ chuyên sâu tại: {dest_dir}")

def sync_folders(src_dir, dst_dir):
    """Đồng bộ thư mục chính xác từ src sang dst."""
    print(f"[*] Đang đồng bộ từ {src_dir} sang {dst_dir}...")
    if os.path.exists(dst_dir):
        shutil.rmtree(dst_dir)
    shutil.copytree(src_dir, dst_dir)
    print(f"-> Đồng bộ xong: {dst_dir}")

def main():
    print("=" * 80)
    print("HỆ THỐNG AEC MULTI-AGENT: XUẤT 2 GÓI HỒ SƠ CẦU KM19+529.080")
    print("1. GÓI 01 - VĨ MÔ / MASTER ĐIỀU HÀNH (14 Sheet liên kết động)")
    print("2. GÓI 02 - VI MÔ / CHUYÊN SÂU SẢN XUẤT (14 Bộ hồ sơ độc lập, sâu chi tiết)")
    print("=" * 80)
    
    # 1. Đóng gói GÓI 01 vào Parent
    build_package_macro(DIR_MACRO_P)
    
    # 2. Trích xuất ma trận giá trị tính toán sạch 100% từ Master bằng Excel COM
    print(f"[*] Đang trích xuất ma trận giá trị tính toán từ Master: {MASTER_SOURCE}...")
    eval_cache = extract_master_eval_cache(MASTER_SOURCE)
    
    # 3. Mở file Master để nhân bản sang 14 bộ hồ sơ vi mô
    print(f"[*] Đang tải Master Workbook từ {MASTER_SOURCE}...")
    wb_master = openpyxl.load_workbook(MASTER_SOURCE, data_only=False)
    
    # 4. Đóng gói GÓI 02 vào Parent với khử triệt để lỗi công thức
    build_package_micro(wb_master, eval_cache, DIR_MICRO_P)
    
    # 5. Đồng bộ sang thư mục con Nested để đảm bảo mở ở đâu cũng thấy
    sync_folders(DIR_MACRO_P, DIR_MACRO_N)
    sync_folders(DIR_MICRO_P, DIR_MICRO_N)
    
    # 6. Đồng bộ Thư mục Hệ thống cắt thép độc lập 01_HE_THONG_CAT_THEP_REBARCUT
    src_rebar = os.path.join(TARGET_PARENT, "01_HE_THONG_CAT_THEP_REBARCUT")
    dst_rebar_n = os.path.join(TARGET_NESTED, "01_HE_THONG_CAT_THEP_REBARCUT")
    if os.path.exists(src_rebar):
        sync_folders(src_rebar, dst_rebar_n)
    
    print("\n" + "=" * 80)
    print("XUẤT HỒ SƠ 2 GÓI CẦU KM19+529.080 THÀNH CÔNG 100%!")
    print("=" * 80)

if __name__ == "__main__":
    main()
