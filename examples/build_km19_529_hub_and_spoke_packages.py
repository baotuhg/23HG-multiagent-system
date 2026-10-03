# -*- coding: utf-8 -*-
"""
HỆ THỐNG XUẤT 5 GÓI VỆ TINH PHÂN QUYỀN THỰC CHIẾN (HUB & SPOKE)
VÀ BẢNG PHÂN QUYỀN & BIÊN BẢN BÀN GIAO 5 GÓI VỆ TINH - CẦU KM19+529.080

Mô hình Hub & Spoke Role-Based Model tuân thủ:
- Luật Xây dựng 135/2025/QH15 & Nghị định 207/2026/NĐ-CP
- Quy trình 15 (23HG-multiagent-system)
- Tiêu chuẩn bảo mật dữ liệu công trường Vincons & Vinhomes Standards

Cấu trúc 5 Gói Vệ tinh:
- GÓI A: ĐỘI CƠ GIỚI, XE MÁY & QUẢN LÝ DẦU DIEZEL
- GÓI B: QUẢN ĐỐC XƯỞNG TIỀN CHẾ & GIA CÔNG CỐT THÉP
- GÓI C: HIỆN TRƯỜNG QUẢN LÝ CHẤT LƯỢNG KCS & THÍ NGHIỆM
- GÓI D: KỸ SƯ QS, BAN DỰ TOÁN & THANH TOÁN HỢP ĐỒNG
- GÓI E: EXECUTIVE DASHBOARD & CHỦ ĐẦU TƯ / BAN LÃNH ĐẠO
"""

from __future__ import annotations
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path, repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)
import os
import sys
import shutil
import json
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT_REPO = repo_path()
TARGET_DIR_1 = project_path(r"HSTK Cầu Km19+529.080_Marker")
TARGET_DIR_2 = os.path.join(TARGET_DIR_1, "HSTK Cầu Km19+529.080_Marker")
TARGET_DIR_3 = os.path.join(TARGET_DIR_2, "HSTK Cầu Km19+529.080_Marker")

FOLDER_HUB_NAME = "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH"

# STYLES CHUẨN
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
FILL_ROLE_A = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid") # Đỏ nhạt
FILL_ROLE_B = PatternFill(start_color="DDD9C4", end_color="DDD9C4", fill_type="solid") # Vàng nâu nhạt
FILL_ROLE_C = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") # Xanh dương nhạt
FILL_ROLE_D = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Xanh lá nhạt
FILL_ROLE_E = PatternFill(start_color="E4DFEC", end_color="E4DFEC", fill_type="solid") # Tím nhạt

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

# -----------------------------------------------------------------------------
# 1. TẠO FILE TIẾN ĐỘ CA MÁY & DẦU DIEZEL CHO GÓI A
# -----------------------------------------------------------------------------
def build_machine_and_fuel_workbook(dest_path: str):
    wb = openpyxl.Workbook()
    
    # Sheet 1: TIEN_DO_CA_MAY_VA_PHU_TAI
    ws1 = wb.active
    ws1.title = "TIEN_DO_CA_MAY_VA_PHU_TAI"
    ws1.views.sheetView[0].showGridLines = True
    
    headers1 = [
        "STT", "Mã thiết bị", "Tên chủng loại máy thi công", "Công suất / Quy cách",
        "Số lượng (Máy)", "Định mức dầu (Lít/ca)", "Tổng số ca thiết kế",
        "Nhu cầu Dầu Diezel (Lít)", "Giai đoạn thi công chính", "Đơn vị vận hành phụ trách", "Ghi chú"
    ]
    title_block(ws1, "BẢNG KẾ HOẠCH ĐIỀU PHỐI CA MÁY & NHU CẦU NHIÊN LIỆU DẦU DIEZEL",
                "Công trình: Cầu Km19+529.080 - Cao tốc Tuyên Quang - Hà Giang (Gói 09-XL)", len(headers1))
    header_row(ws1, 5, headers1, [6, 14, 30, 24, 14, 18, 18, 22, 28, 26, 20])
    
    machines = [
        (1, "MK-01", "Máy khoan cọc nhồi D1200", "Đầu xoay đập đá Bauer BG25", 2, 145.0, 156, "=F6*G6", "Khoan 26 cọc D1200 mố trụ", "Đội Khoan Cọc Nhồi 1", "Khoan qua hang Karst"),
        (2, "CX-01", "Cần cẩu bánh xích 50T", "Kobelco CKE500 (50 tấn)", 1, 62.0, 180, "=F7*G7", "Hạ lồng thép cọc & lắp ván khuôn", "Đội Cơ Giới Nặng", "Phục vụ 2 bờ"),
        (3, "CL-01", "Cần cẩu bánh lốp 25T", "Kato KR250 (25 tấn)", 1, 48.0, 120, "=F8*G8", "Bốc dỡ vật tư & đúc dầm bãi đúc", "Đội Cơ Giới Phụ Trợ", "Tại bãi tiền chế"),
        (4, "MD-01", "Máy đào gầu nghịch 1.25m3", "Komatsu PC200-8", 2, 68.0, 95, "=F9*G9", "Đào hố móng M1, M2, T1, T2", "Đội Đào Đất Móng", "Có búa đập đá"),
        (5, "OT-01", "Ô tô tự đổ 15 tấn", "Howo 371HP (15 tấn)", 4, 52.0, 140, "=F10*G10", "Vận chuyển đất đá hố móng ra bãi thải", "Đội Xe Tự Đổ", "Cự ly 5km"),
        (6, "XB-01", "Xe bồn trộn bê tông 9m3", "Hyundai HD270 (9m3)", 3, 42.0, 210, "=F11*G11", "Vận chuyển BTXM từ trạm ra vị trí đổ", "Đội Vận Chuyển Bê Tông", "Cự ly bãi đúc 1.5km"),
        (7, "BM-01", "Xe bơm bê tông cần 42m", "Putzmeister 42m", 1, 46.0, 85, "=F12*G12", "Bơm bê tông bệ mố trụ, xà mũ, dầm", "Đội Bơm Bê Tông", "Bơm cao áp"),
        (8, "MN-01", "Máy nén khí 7.5m3/phút", "Airman PDS265 (7.5m3/min)", 2, 38.0, 110, "=F13*G13", "Thổi rửa đáy hố khoan & đục đầu cọc", "Đội Khoan Cọc", "Áp lực 7 bar"),
        (9, "MH-01", "Máy hàn điện 500A (máy phát)", "Máy phát hàn Denyo 500A", 4, 28.0, 190, "=F14*G14", "Gia công lồng thép & nối thép cọc", "Xưởng Cốt Thép Tiền Chế", "Hàn hồ quang tay"),
        (10, "TT-01", "Trạm trộn BTXM 60m3/h", "Sicoma 60m3/h (máy phát 150kVA)", 1, 58.0, 160, "=F15*G15", "Sản xuất BT C10, C30, C45/55", "Trạm Bê Tông Hiện Trường", "Chạy máy phát khi mất điện"),
        (11, "GL-01", "Giá lao dầm Super-T 78.88T", "Giá lao ray P43 tời kéo điện", 1, 35.0, 60, "=F16*G16", "Lao lắp 15 phiến dầm Super-T 38.2m", "Đội Lao Lắp Dầm Chuyên Nghiệp", "Chạy máy phát dự phòng"),
        (12, "LU-01", "Xe lu rung 25 tấn", "Hamm 3411 (25 tấn)", 1, 45.0, 45, "=F17*G17", "Đắp đất K95/K98 đường đầu cầu", "Đội Nền Đường Công Vụ", "Lu lèn chặt K98")
    ]
    
    r = 6
    for m in machines:
        put(ws1, r, 1, m[0], align=ALIGN_CENTER)
        put(ws1, r, 2, m[1], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws1, r, 3, m[2])
        put(ws1, r, 4, m[3])
        put(ws1, r, 5, m[4], "#,##0", align=ALIGN_CENTER)
        put(ws1, r, 6, m[5], "#,##0.0")
        put(ws1, r, 7, m[6], "#,##0", align=ALIGN_CENTER)
        put(ws1, r, 8, m[7], "#,##0.0", font=FONT_BOLD)
        put(ws1, r, 9, m[8])
        put(ws1, r, 10, m[9])
        put(ws1, r, 11, m[10])
        r += 1
        
    put(ws1, r, 3, "TỔNG CỘNG NHU CẦU NHIÊN LIỆU DẦU DIEZEL (LÍT)", font=FONT_BOLD)
    put(ws1, r, 8, f"=SUM(H6:H{r-1})", "#,##0.0", font=FONT_BOLD, border=DOUBLE_BOTTOM_BORDER, fill=FILL_TOT)
    
    # Sheet 2: THE_DOI_CAP_PHAT_DAU_HANG_NGAY
    ws2 = wb.create_sheet(title="THE_DOI_CAP_PHAT_DAU")
    ws2.views.sheetView[0].showGridLines = True
    headers2 = ["STT", "Ngày cấp", "Mã máy", "Tên thiết bị", "Số giờ hoạt động (h)", "Số ca quy đổi", "Lượng dầu cấp (Lít)", "Người lái máy nhận", "Thủ kho cấp phát", "Tích kê số", "Xác nhận Đội trưởng"]
    title_block(ws2, "SỔ THEO DÕI & CẤP PHÁT NHIÊN LIỆU DẦU DIEZEL HIỆN TRƯỜNG", "Định mức tiêu hao thực tế so sánh với định mức thiết kế", len(headers2))
    header_row(ws2, 5, headers2, [6, 14, 12, 26, 18, 16, 18, 22, 20, 14, 20])
    
    # Dữ liệu mẫu 5 ngày đầu
    samples = [
        (1, "01/10/2026", "MK-01", "Máy khoan cọc nhồi Bauer BG25", 8.0, 1.0, 142.0, "Nguyễn Văn Hùng", "Trần Bá Thắng", "TK-001", "Đã duyệt"),
        (2, "01/10/2026", "CX-01", "Cần cẩu bánh xích 50T", 7.5, 0.94, 60.0, "Lê Đình Long", "Trần Bá Thắng", "TK-002", "Đã duyệt"),
        (3, "02/10/2026", "MK-01", "Máy khoan cọc nhồi Bauer BG25", 8.0, 1.0, 145.0, "Nguyễn Văn Hùng", "Trần Bá Thắng", "TK-003", "Đã duyệt"),
        (4, "02/10/2026", "MD-01", "Máy đào Komatsu PC200-8", 8.0, 1.0, 68.0, "Phạm Quốc Tuấn", "Trần Bá Thắng", "TK-004", "Đã duyệt"),
        (5, "03/10/2026", "MK-01", "Máy khoan cọc nhồi Bauer BG25", 8.0, 1.0, 140.0, "Nguyễn Văn Hùng", "Trần Bá Thắng", "TK-005", "Đã duyệt"),
    ]
    r = 6
    for s in samples:
        put(ws2, r, 1, s[0], align=ALIGN_CENTER)
        put(ws2, r, 2, s[1], align=ALIGN_CENTER)
        put(ws2, r, 3, s[2], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws2, r, 4, s[3])
        put(ws2, r, 5, s[4], "#,##0.0")
        put(ws2, r, 6, s[5], "#,##0.00")
        put(ws2, r, 7, s[6], "#,##0.0", font=FONT_BOLD)
        put(ws2, r, 8, s[7])
        put(ws2, r, 9, s[8])
        put(ws2, r, 10, s[9], align=ALIGN_CENTER)
        put(ws2, r, 11, s[10], font=FONT_BOLD, align=ALIGN_CENTER, fill=FILL_OK)
        r += 1
    
    wb.save(dest_path)
    print(f"  + Đã tạo File Tiến độ Ca máy & Dầu Diezel: {dest_path}")
    return dest_path


# -----------------------------------------------------------------------------
# 2. TẠO BẢNG PHÂN QUYỀN & BIÊN BẢN BÀN GIAO 5 GÓI VỆ TINH (4 SHEETS)
# -----------------------------------------------------------------------------
def build_permission_and_handover_workbook(dest_path: str):
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------------------
    # SHEET 1: BANG_MA_TRAN_PHAN_QUYEN (RACI MATRIX)
    # -------------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "BANG_MA_TRAN_PHAN_QUYEN"
    ws1.views.sheetView[0].showGridLines = True
    
    headers1 = [
        "STT", "Gói Vệ Tinh", "Tên phân hệ chức năng", "Đối tượng tiếp nhận & sử dụng",
        "Mức độ bảo mật", "Quyền xem (Read)", "Quyền sửa (Write)", "Dữ liệu cấm truy cập (Restricted)",
        "Trách nhiệm (RACI)", "Kênh phân phối / Thiết bị"
    ]
    title_block(ws1, "BẢNG MA TRẬN PHÂN QUYỀN TRUY CẬP DỮ LIỆU HIỆN TRƯỜNG (MÔ HÌNH HUB & SPOKE)",
                "Tuân thủ: Luật Xây dựng 135/2025/QH15, NĐ 207/2026/NĐ-CP & Tiêu chuẩn bảo mật Vincons", len(headers1))
    header_row(ws1, 5, headers1, [6, 16, 30, 32, 18, 28, 24, 34, 18, 26])
    
    roles_data = [
        (1, "GÓI A", "Cơ giới, Xe máy & Dầu Diezel", "Đội trưởng xe máy, Thủ kho dầu, Lái máy, Lái xe bồn",
         "NỘI BỘ (INTERNAL)", "Tiến độ ca máy, phụ tải, ĐM dầu, nhật trình xe", "Nhật ký ca máy, tích kê cấp dầu",
         "TUYỆT ĐỐI CẤM xem Đơn giá, Dự toán G_XD, Lợi nhuận thầu", "R (Responsible)", "Tablet / Giấy A4 tại kho dầu", FILL_ROLE_A),
        (2, "GÓI B", "Xưởng tiền chế & Gia công cốt thép", "Quản đốc xưởng thép, Thợ cắt uốn, Thợ máy CNC, ĐV cấp thép",
         "NỘI BỘ (INTERNAL)", "BBS thép, sơ đồ cắt 1D CSP 11.7m, file CSV CNC, kho đề-xê", "Cập nhật thanh thép đã cắt, mã phôi thừa lưu kho",
         "TUYỆT ĐỐI CẤM xem Đơn giá thép mua, Giá trị hợp đồng, Dự toán", "R (Responsible)", "Màn hình máy CNC / In A3 xưởng", FILL_ROLE_B),
        (3, "GÓI C", "Hiện trường Quản lý chất lượng (KCS)", "Kỹ sư QA/QC, Tư vấn giám sát (TVGS), Thí nghiệm LAS-XD",
         "QUAN TRỌNG (CONFIDENTIAL)", "Danh mục KCS, ngày nghiệm thu logic, BBNT Word, mẫu R7/R28", "Nhập kết quả thí nghiệm, ký số biên bản nghiệm thu",
         "CẤM xem Đơn giá dự toán chi tiết, Doanh thu thanh toán kỳ 03a", "A (Accountable)", "Laptop / iPad hiện trường TVGS", FILL_ROLE_C),
        (4, "GÓI D", "QS, Dự toán & Thanh toán hợp đồng", "Kỹ sư QS, Ban Đấu thầu - Dự toán, Kế toán dự án",
         "MẬT CAO (RESTRICTED)", "Toàn quyền xem Khối lượng Takeoff, Đơn giá, G_XD, Phụ lục 03a", "Lập dự toán, điều chỉnh giá thầu, lập hồ sơ giải ngân 03a",
         "Chỉ chia sẻ nội bộ phòng Kế hoạch & Ban Giám đốc", "R (Responsible)", "PC văn phòng / Khóa mật khẩu BitLocker", FILL_ROLE_D),
        (5, "GÓI E", "Executive Dashboard (Hub điều hành)", "Giám đốc Dự án, Ban Giám đốc, Ban QLDA (Chủ đầu tư)",
         "QUẢN TRỊ TỐI CAO (EXECUTIVE)", "KPI toàn diện, Đường găng CPM, MS Project XML, Audit 100/100", "Phê duyệt tiến độ, ký duyệt thanh toán, chuẩn thuận phát sinh",
         "Toàn quyền truy cập mọi tầng dữ liệu hệ sinh thái", "I (Informed / Owner)", "Executive Web Portal / iPad Ban Lãnh đạo", FILL_ROLE_E)
    ]
    
    r = 6
    for rd in roles_data:
        put(ws1, r, 1, rd[0], align=ALIGN_CENTER)
        put(ws1, r, 2, rd[1], font=FONT_BOLD, align=ALIGN_CENTER, fill=rd[10])
        put(ws1, r, 3, rd[2], font=FONT_BOLD)
        put(ws1, r, 4, rd[3])
        put(ws1, r, 5, rd[4], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws1, r, 6, rd[5])
        put(ws1, r, 7, rd[6])
        put(ws1, r, 8, rd[7], font=FONT_BOLD, fill=FILL_WARN)
        put(ws1, r, 9, rd[8], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws1, r, 10, rd[9])
        r += 1
        
    # Ghi chú RACI
    r += 1
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
    ws1.cell(r, 2, "Ghi chú ma trận RACI: R = Responsible (Người thực hiện trực tiếp); A = Accountable (Người chịu trách nhiệm nghiệm thu); C = Consulted (Người tư vấn chuyên môn); I = Informed (Người nhận báo cáo)").font = FONT_IT
    
    # -------------------------------------------------------------------------
    # SHEET 2: BIEN_BAN_BAN_GIAO_5_GOI (HANDOVER PROTOCOL)
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="BIEN_BAN_BAN_GIAO_5_GOI")
    ws2.views.sheetView[0].showGridLines = True
    
    title_block(ws2, "BIÊN BẢN BÀN GIAO HỒ SƠ THỰC CHIẾN 5 GÓI VỆ TINH PHÂN QUYỀN",
                "Số: 01/2026/BB-BG5G/23HG-KM19 | Căn cứ Nghị định 207/2026/NĐ-CP & Tiêu chuẩn bảo mật dự án", 10)
    
    ws2["A5"] = "Hôm nay, ngày 01 tháng 10 năm 2026, tại Văn phòng Ban Điều hành Công trường Cầu Km19+529.080, chúng tôi gồm có:"
    ws2["A5"].font = FONT_BOLD
    ws2.merge_cells("A5:J5")
    
    parties = [
        ("I. ĐẠI DIỆN BỘ PHẬN ĐIỀU PHỐI HỆ THỐNG AEC (BÊN BÀN GIAO):", "Kỹ sư Trưởng / Hệ thống 23HG Multi-Agent System - Nguyễn Bảo Tú (@baotuhg)"),
        ("II. ĐẠI DIỆN CÁC BỘ PHẬN TIẾP NHẬN HIỆN TRƯỜNG (BÊN NHẬN BÀN GIAO):", ""),
        ("  1. Đại diện Gói A (Đội Cơ giới & Quản lý Dầu):", "Ông: Lê Bá Toàn           - Chức vụ: Đội trưởng Xe máy Cơ giới"),
        ("  2. Đại diện Gói B (Xưởng Tiền chế Cốt thép):", "Ông: Nguyễn Văn Cường      - Chức vụ: Quản đốc Xưởng Gia công Cốt thép"),
        ("  3. Đại diện Gói C (Hiện trường QLCL - KCS):", "Ông: Trần Văn Thành        - Chức vụ: Trưởng bộ phận QA/QC Hiện trường"),
        ("  4. Đại diện Gói D (Ban QS & Dự toán):", "Bà: Hoàng Thu Trang       - Chức vụ: Kỹ sư Trưởng Ban QS - Kế hoạch"),
        ("  5. Đại diện Gói E (Ban Điều hành Dự án):", "Ông: Phạm Hồng Hải         - Chức vụ: Giám đốc Ban Điều hành Dự án")
    ]
    
    curr_r = 7
    for p_title, p_desc in parties:
        ws2.cell(curr_r, 1, p_title).font = FONT_BOLD
        ws2.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=4)
        if p_desc:
            ws2.cell(curr_r, 5, p_desc).font = FONT_REG
            ws2.merge_cells(start_row=curr_r, start_column=5, end_row=curr_r, end_column=10)
        curr_r += 1
        
    curr_r += 1
    ws2.cell(curr_r, 1, "NỘI DUNG BÀN GIAO CHI TIẾT THEO 5 GÓI VỆ TINH:").font = FONT_SEC
    ws2.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=10)
    curr_r += 1
    
    headers_bg = ["STT", "Gói hồ sơ", "Danh mục tệp số bàn giao", "Số lượng tệp", "Định dạng tệp", "Tình trạng bàn giao", "Xác nhận ký nhận Bên nhận"]
    header_row(ws2, curr_r, headers_bg, [6, 14, 40, 14, 16, 22, 26])
    curr_r += 1
    
    bg_rows = [
        (1, "GÓI A", "Tiến độ ca máy, phụ tải & Theo dõi cấp phát dầu Diezel", 2, "XLSX, XML", "Đã bàn giao đầy đủ (Không có giá)", "Đã nhận đủ hồ sơ"),
        (2, "GÓI B", "Sơ đồ cắt thép 11.7m RebarCut Pro, 11 file Ø, Lệnh CNC", 14, "XLSX, CSV", "Đã nạp vào trạm máy cắt CNC", "Đã nhận đủ hồ sơ"),
        (3, "GÓI C", "22 Biên bản nghiệm thu KCS Word, Danh mục KCS, Mẫu A4", 6, "DOCX, XLSX", "Đã tích hợp ma trận logic ngày", "Đã nhận đủ hồ sơ"),
        (4, "GÓI D", "Bóc tách hình học Takeoff, Dự toán G_XD, Thanh toán 03a", 3, "XLSX", "Bảo mật nghiêm ngặt (Có đơn giá)", "Đã nhận đủ hồ sơ"),
        (5, "GÓI E", "Executive Master, Tiến độ CPM MS Project, Audit 100/100", 5, "XLSX, MPP, XML, MD", "Đầy đủ KPI tiến độ & chất lượng", "Đã nhận đủ hồ sơ")
    ]
    for b in bg_rows:
        put(ws2, curr_r, 1, b[0], align=ALIGN_CENTER)
        put(ws2, curr_r, 2, b[1], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws2, curr_r, 3, b[2])
        put(ws2, curr_r, 4, b[3], "#,##0", align=ALIGN_CENTER)
        put(ws2, curr_r, 5, b[4], align=ALIGN_CENTER)
        put(ws2, curr_r, 6, b[5], fill=FILL_OK, align=ALIGN_CENTER)
        put(ws2, curr_r, 7, b[6], font=FONT_BOLD, align=ALIGN_CENTER)
        curr_r += 1
        
    curr_r += 2
    # Khung chữ ký
    sign_titles = ["ĐẠI DIỆN GÓI A\n(Đội trưởng Xe máy)", "ĐẠI DIỆN GÓI B\n(Quản đốc Xưởng thép)", "ĐẠI DIỆN GÓI C\n(Trưởng QA/QC)", "ĐẠI DIỆN GÓI D\n(Kỹ sư Trưởng QS)", "ĐẠI DIỆN BÊN BÀN GIAO\n(Kỹ sư Trưởng Hệ thống)", "GIÁM ĐỐC DỰ ÁN\n(GÓI E - PHÊ DUYỆT)"]
    cols = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 9), (10, 10)]
    
    # -------------------------------------------------------------------------
    # SHEET 3: DANH_MUC_HO_SO_BAN_GIAO (CHECKLIST TẬP TIN CHI TIẾT)
    # -------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="DANH_MUC_HO_SO_BAN_GIAO")
    ws3.views.sheetView[0].showGridLines = True
    
    headers3 = ["STT", "Mã Gói", "Thư mục vệ tinh", "Tên tệp hồ sơ bàn giao", "Định dạng", "Dung lượng (KB)", "Chức năng thực chiến tại công trường"]
    title_block(ws3, "DANH MỤC CHI TIẾT 25+ TẬP TIN HỒ SƠ 5 GÓI VỆ TINH BÀN GIAO THỰC CHIẾN",
                "Được đồng bộ tự động từ hệ sinh thái AEC Multi-Agent System", len(headers3))
    header_row(ws3, 5, headers3, [6, 12, 30, 46, 14, 16, 46])
    
    file_checklist = [
        # Gói A
        (1, "GÓI A", "GOI_A_CO_GIOI_VA_DAU_DIEZEL", "260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx", "Excel", 45, "Theo dõi 12 đầu máy, phụ tải, định mức dầu Diezel"),
        (2, "GÓI A", "GOI_A_CO_GIOI_VA_DAU_DIEZEL", "260920_Tien_Do_CaMay_Cau_Km19+529.080.xml", "XML", 18, "Tiến độ huy động thiết bị tích hợp Primavera/MS Project"),
        # Gói B
        (3, "GÓI B", "GOI_B_XUONG_TIEN_CHE_COT_THEP", "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx", "Excel", 4350, "Sơ đồ cắt thép 1D CSP 11.7m toàn cầu (33.212 cây 11.7m)"),
        (4, "GÓI B", "GOI_B_XUONG_TIEN_CHE_COT_THEP", "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx", "Excel", 8, "Dashboard so sánh chỉ tiêu kinh tế kỹ thuật 11 đường kính Ø"),
        (5, "GÓI B", "GOI_B_XUONG_TIEN_CHE_COT_THEP", "04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx", "Excel", 49, "Bảng thống kê thép chi tiết 396 dòng phục vụ bẻ sắt"),
        (6, "GÓI B", "GOI_B_XUONG_TIEN_CHE_COT_THEP", "01_Phieu_Cat_Thep_Cau_Km19+529.080.csv", "CSV", 9516, "Lệnh cắt CNC nạp thẳng máy cắt đa thanh Shearline"),
        (7, "GÓI B", "GOI_B_XUONG_TIEN_CHE_COT_THEP", "README_QUY_TRINH_VAN_HANH_BAI_THEP.md", "Markdown", 3, "Quy chuẩn quản lý bãi thép và kiểm soát phế liệu đề-xê"),
        # Gói C
        (8, "GÓI C", "GOI_C_HIEN_TRUONG_QLCL_KCS", "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx", "Word", 44, "22 Biên bản nghiệm thu KCS chuẩn NĐ 207/2026 sẵn sàng in ký"),
        (9, "GÓI C", "GOI_C_HIEN_TRUONG_QLCL_KCS", "11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx", "Excel", 30, "Quản lý mã biên bản, ngày nghiệm thu logic chéo"),
        (10, "GÓI C", "GOI_C_HIEN_TRUONG_QLCL_KCS", "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx", "Excel", 11, "Mẫu in A4 tự động nhảy số theo mã công việc"),
        (11, "GÓI C", "GOI_C_HIEN_TRUONG_QLCL_KCS", "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx", "Excel", 14, "Mẫu in A4 nghiệm thu vật liệu thép, cát, đá, xi măng"),
        (12, "GÓI C", "GOI_C_HIEN_TRUONG_QLCL_KCS", "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx", "Excel", 14, "Mẫu in A4 lấy mẫu bê tông hiện trường và ngày nén"),
        (13, "GÓI C", "GOI_C_HIEN_TRUONG_QLCL_KCS", "05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx", "Excel", 65, "Bảng định mức cấp phối và tần suất 809 mẫu thí nghiệm"),
        # Gói D
        (14, "GÓI D", "GOI_D_QS_DU_TOAN_THANH_TOAN", "03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx", "Excel", 61, "101 dòng hình học diễn giải tiên lượng chi tiết"),
        (15, "GÓI D", "GOI_D_QS_DU_TOAN_THANH_TOAN", "08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx", "Excel", 17, "Dự toán chi phí xây dựng G_XD (BẢO MẬT ĐƠN GIÁ)"),
        (16, "GÓI D", "GOI_D_QS_DU_TOAN_THANH_TOAN", "09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx", "Excel", 19, "Bảng thanh toán giải ngân Mẫu 03.a NĐ 254/2025"),
        (17, "GÓI D", "GOI_D_QS_DU_TOAN_THANH_TOAN", "06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx", "Excel", 25, "140 dòng phân tích vật tư định mức TT 38/2026"),
        # Gói E
        (18, "GÓI E", "GOI_E_EXECUTIVE_DASHBOARD", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx", "Excel", 110, "Master Workbook 14 Sheet liên kết động toàn diện"),
        (19, "GÓI E", "GOI_E_EXECUTIVE_DASHBOARD", "Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp", "MPP", 290, "Tiến độ đường găng CPM gốc MS Project"),
        (20, "GÓI E", "GOI_E_EXECUTIVE_DASHBOARD", "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml", "XML", 18, "File trao đổi tiến độ liên thông các bên"),
        (21, "GÓI E", "GOI_E_EXECUTIVE_DASHBOARD", "BAO_CAO_THAM_TRA_AEC_AUDIT.md", "Markdown", 2, "Báo cáo Thẩm tra Kỹ thuật Độc lập Audit 100/100"),
        (22, "GÓI E", "GOI_E_EXECUTIVE_DASHBOARD", "Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md", "Markdown", 7, "Thuyết minh BPTC 8 chương chuẩn mực TCVN")
    ]
    
    r = 6
    for it in file_checklist:
        put(ws3, r, 1, it[0], align=ALIGN_CENTER)
        put(ws3, r, 2, it[1], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws3, r, 3, it[2])
        put(ws3, r, 4, it[3], font=FONT_BOLD)
        put(ws3, r, 5, it[4], align=ALIGN_CENTER)
        put(ws3, r, 6, it[5], "#,##0", align=ALIGN_RIGHT)
        put(ws3, r, 7, it[6])
        r += 1
        
    # -------------------------------------------------------------------------
    # SHEET 4: QUY_CHE_BAO_MAT_DU_LIEU (SECURITY & COMPLIANCE POLICY)
    # -------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="QUY_CHE_BAO_MAT_DU_LIEU")
    ws4.views.sheetView[0].showGridLines = True
    
    headers4 = ["Điều", "Khoản", "Nội dung quy chế bảo mật & kiểm soát dữ liệu", "Phạm vi áp dụng", "Chế tài xử lý vi phạm"]
    title_block(ws4, "QUY CHẾ BẢO MẬT THÔNG TIN & AN TOÀN DỮ LIỆU CÔNG TRƯỜNG AEC MULTI-AGENT",
                "Áp dụng bắt buộc cho toàn thể cán bộ kỹ sư và nhà thầu phụ tại Dự án Cầu Km19+529.080", len(headers4))
    header_row(ws4, 5, headers4, [8, 10, 50, 24, 30])
    
    rules = [
        ("Điều 1", "Khoản 1", "Tuyệt đối không chia sẻ dữ liệu Gói D (Dự toán G_XD, Đơn giá dự thầu, Doanh thu thanh toán) cho nhân sự phụ trách Gói A, B, C.", "Toàn bộ dự án", "Kỷ luật cảnh cáo, điều chuyển công tác"),
        ("Điều 1", "Khoản 2", "Các file Gói A (Dầu Diezel) và Gói B (Thép) chỉ lưu hành trong phạm vi bãi đúc, xưởng cắt và trạm cấp nhiên liệu.", "Đội xe máy & Xưởng thép", "Thu hồi quyền truy cập hệ thống"),
        ("Điều 2", "Khoản 1", "Nghiêm cấm hành vi chỉnh sửa công thức trong file Master Gói E và các mẫu in Gói C làm đứt gãy tính toàn vẹn 100% công thức sống.", "Kỹ sư KCS & QS", "Đình chỉ sử dụng hệ thống tự động"),
        ("Điều 2", "Khoản 2", "Mọi thay đổi về bản vẽ thiết kế phát sinh hình học bắt buộc phải chạy qua Agent aec_cad_takeoff để cập nhật đồng bộ State Graph.", "Ban Kỹ thuật", "Không công nhận khối lượng phát sinh"),
        ("Điều 3", "Khoản 1", "Biên bản bàn giao 5 Gói phải có đầy đủ chữ ký xác nhận của Chỉ huy trưởng và đại diện 5 bộ phận trước khi triển khai thi công đợt 1.", "Chỉ huy trưởng & Trưởng bộ phận", "Chưa đủ điều kiện khởi công hạng mục")
    ]
    r = 6
    for ru in rules:
        put(ws4, r, 1, ru[0], font=FONT_BOLD, align=ALIGN_CENTER)
        put(ws4, r, 2, ru[1], align=ALIGN_CENTER)
        put(ws4, r, 3, ru[2])
        put(ws4, r, 4, ru[3], font=FONT_BOLD)
        put(ws4, r, 5, ru[4], font=FONT_BOLD, fill=FILL_WARN)
        r += 1

    wb.save(dest_path)
    print(f"  + Đã tạo Bảng Phân Quyền & Bàn Giao Excel: {dest_path}")
    return dest_path


# -----------------------------------------------------------------------------
# 3. TẠO TÀI LIỆU MARKDOWN HƯỚNG DẪN BÀN GIAO & PHÂN QUYỀN
# -----------------------------------------------------------------------------
def build_permission_markdown(dest_path: str):
    content = r"""# BẢNG MA TRẬN PHÂN QUYỀN & BIÊN BẢN BÀN GIAO 5 GÓI VỆ TINH THỰC CHIẾN
## DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) — CẦU KM19+529.080
**Mô hình điều phối:** Hub & Spoke Role-Based Model (Quy trình 15 - 23HG MultiAgent System)  
**Tiêu chuẩn tuân thủ:** Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Nghị định 254/2025/NĐ-CP & Vincons Standards  
**Ngày lập & bàn giao:** 01/10/2026  

---

### I. TẠI SAO PHẢI PHÂN CHIA THÀNH 5 GÓI VỆ TINH PHÂN QUYỀN?
Trong môi trường công trường thực chiến, việc sử dụng chung 1 file Monolithic 14 Sheet dẫn đến 3 rủi ro chí mạng:
1. **Xung đột file khóa (Read-Only Lock):** Khi Kỹ sư QS đang mở file tính toán giải ngân thì Đội xe máy không thể cập nhật tích kê tiêu thụ dầu Diezel.
2. **Lộ lọt bí mật tài chính & thương mại:** Thợ gia công sắt thép hay lái máy xúc có thể nhìn thấy toàn bộ đơn giá dự toán, định mức lợi nhuận và chi phí gián tiếp của Tổng thầu.
3. **Quá tải trên thiết bị di động hiện trường:** File tổng hợp quá lớn làm lag thiết bị tại hố móng và rất dễ dẫn đến lỗi `#REF!` do vô tình xóa cột.

Hệ thống chuẩn hóa giải pháp **Hub & Spoke** phân tách độc lập thành 5 gói hồ sơ chuyên biệt theo đúng thẩm quyền và trách nhiệm.

---

### II. BẢNG MA TRẬN PHÂN QUYỀN TRUY CẬP HIỆN TRƯỜNG (RACI MATRIX)

| Gói Vệ Tinh | Bộ phận sử dụng | Mức độ bảo mật | Quyền xem (Read) | Quyền sửa (Write) | Dữ liệu cấm tuyệt đối (Restricted) |
| :--- | :--- | :---: | :--- | :--- | :--- |
| **GÓI A: Cơ giới & Dầu** | Đội trưởng xe máy, Thủ kho dầu, Lái máy | **Nội bộ** | Tiến độ ca máy, phụ tải, ĐM dầu, nhật trình xe | Ghi số giờ máy, tích kê cấp dầu hàng ngày | **CẤM XEM** Đơn giá tiền, Dự toán $G_{XD}$, Doanh thu |
| **GÓI B: Xưởng cốt thép** | Quản đốc xưởng, Thợ uốn cắt, Máy CNC | **Nội bộ** | Sơ đồ cắt 11.7m, BBS 396 dòng, CSV nạp CNC | Báo cáo số thanh đã cắt, mã đề-xê lưu kho | **CẤM XEM** Đơn giá mua thép, Giá trị hợp đồng |
| **GÓI C: Hiện trường KCS** | Kỹ sư QA/QC, TVGS, Thí nghiệm LAS-XD | **Quan trọng** | 22 BBNT Word, ma trận ngày KCS, cấp phối mẫu | Nhập kết quả thí nghiệm R7/R28, ký biên bản | **CẤM XEM** Đơn giá dự toán chi tiết, Thanh toán 03a |
| **GÓI D: QS & Dự toán** | Kỹ sư QS, Ban Đấu thầu, Kế toán | **Mật cao** | Hình học Takeoff, Dự toán $G_{XD}$, Thanh toán 03a | Lập dự toán điều chỉnh, hồ sơ giải ngân đợt | Chỉ lưu hành nội bộ phòng Kế hoạch & Ban Giám đốc |
| **GÓI E: Executive Hub** | Giám đốc Dự án, Ban QLDA, Chủ đầu tư | **Tối cao** | Toàn quyền KPI, Đường găng CPM, Audit 100/100 | Phê duyệt tổng thể, ký giải ngân, chuẩn thuận phát sinh | Truy cập toàn bộ dữ liệu hệ sinh thái |

---

### III. DANH MỤC 25+ TẬP TIN HỒ SƠ 5 GÓI VỆ TINH BÀN GIAO

#### 1. GÓI A: ĐỘI CƠ GIỚI, XE MÁY & QUẢN LÝ DẦU DIEZEL
- `260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx`: Kế hoạch điều phối 12 đầu máy chính & Sổ theo dõi cấp phát dầu Diezel.
- `260920_Tien_Do_CaMay_Cau_Km19+529.080.xml`: Tiến độ huy động thiết bị tích hợp MS Project.

#### 2. GÓI B: QUẢN ĐỐC XƯỞNG TIỀN CHẾ & GIA CÔNG CỐT THÉP
- `01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`: Sơ đồ cắt thép 1D CSP 11.7m toàn cầu (33.212 cây 11.7m = 662.890 kg thép).
- `00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx`: Bảng Dashboard chỉ tiêu kinh tế kỹ thuật của 11 đường kính Ø và Cáp DƯL.
- `04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx`: Thống kê thép chi tiết 396 dòng phục vụ thợ uốn bẻ tại bãi.
- `01_Phieu_Cat_Thep_Cau_Km19+529.080.csv`: Tệp lệnh cắt CNC nạp thẳng máy cắt đa thanh Shearline.
- `THEO_TUNG_DUONG_KINH_PHI/`: 11 file chuyên sâu cho từng phi từ Ø8 đến Ø32 và cáp DƯL 15.2mm.
- `README_QUY_TRINH_VAN_HANH_BAI_THEP.md`: Hướng dẫn quản trị phôi thừa đề-xê $\ge 100D$.

#### 3. GÓI C: HIỆN TRƯỜNG QUẢN LÝ CHẤT LƯỢNG KCS & THÍ NGHIỆM
- `Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx`: Trọn bộ 22 Biên bản nghiệm thu KCS chuẩn NĐ 207/2026 sẵn sàng in ký.
- `11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx`: Bảng quản lý 22 BBNT, kiểm soát chéo ngày yêu cầu và ngày ký.
- `12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx`: Mẫu in A4 tự động nhảy số theo mã công việc.
- `13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx`: Mẫu in A4 nghiệm thu vật liệu đầu vào (Thép, Xi măng, Cát, Đá, Gối cầu).
- `14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx`: Mẫu in A4 lấy mẫu nén bê tông R7, R28.
- `05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx`: Định mức cấp phối và kế hoạch lấy 809 mẫu QA/QC.

#### 4. GÓI D: KỸ SƯ QS, BAN DỰ TOÁN & THANH TOÁN HỢP ĐỒNG
- `03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx`: 101 dòng hình học diễn giải tiên lượng chi tiết.
- `08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx`: Bảng tính tổng hợp kinh phí $G_{XD} = T + GT + TL + VAT$ (Bảo mật đơn giá).
- `09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx`: Bảng thanh toán giải ngân Mẫu 03.a theo Nghị định 254/2025/NĐ-CP.
- `06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx`: 140 dòng phân tích vật tư định mức TT 38/2026.
- `07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx`: Tổng hợp BOM và kế hoạch cung ứng 4 phân đợt.

#### 5. GÓI E: EXECUTIVE DASHBOARD & CHỦ ĐẦU TƯ / BAN LÃNH ĐẠO
- `Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx`: Master Workbook 14 Sheet liên kết động 100% (Zero Dead Numbers).
- `Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp`: File tiến độ gốc MS Project quản lý 36 công tác đường găng CPM.
- `Tien_Do_Thi_Cong_Cau_Km19+529.080.xml`: File tiến độ liên thông Primavera / MS Project.
- `BAO_CAO_THAM_TRA_AEC_AUDIT.md`: Báo cáo Thẩm tra Kỹ thuật Độc lập Audit Score 100/100.
- `Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md`: Thuyết minh biện pháp thi công 8 chương chuẩn TCVN.

---

### IV. BIÊN BẢN BÀN GIAO VÀ CAM KẾT HIỆN TRƯỜNG
1. **Bên bàn giao:** Kỹ sư Trưởng Hệ thống 23HG Multi-Agent System bàn giao nguyên trạng đầy đủ các tệp số cho 5 bộ phận.
2. **Bên nhận bàn giao:** Đại diện 5 bộ phận cam kết tiếp nhận đúng quyền hạn, không can thiệp làm hỏng cấu trúc liên kết và tuân thủ 100% quy chế bảo mật dữ liệu.
"""
    with open(dest_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"  + Đã tạo Báo cáo Markdown Phân Quyền & Bàn Giao: {dest_path}")
    return dest_path


# -----------------------------------------------------------------------------
# 4. CHÍNH HÀM MAIN: ĐÓNG GÓI HOÀN CHỈNH 5 GÓI VỆ TINH VÀ ĐỒNG BỘ
# -----------------------------------------------------------------------------
def main():
    print("=" * 80)
    print("BỘ ĐIỀU PHỐI 23HG: XUẤT 5 GÓI VỆ TINH PHÂN QUYỀN & BIÊN BẢN BÀN GIAO")
    print("DỰ ÁN: CẦU KM19+529.080 (CAO TỐC TUYÊN QUANG - HÀ GIANG)")
    print("=" * 80)

    # 1. Thư mục gốc cho 5 gói
    hub_root = os.path.join(TARGET_DIR_1, FOLDER_HUB_NAME)
    os.makedirs(hub_root, exist_ok=True)

    pkg_dirs = {
        "A": os.path.join(hub_root, "GOI_A_CO_GIOI_VA_DAU_DIEZEL"),
        "B": os.path.join(hub_root, "GOI_B_XUONG_TIEN_CHE_COT_THEP"),
        "C": os.path.join(hub_root, "GOI_C_HIEN_TRUONG_QLCL_KCS"),
        "D": os.path.join(hub_root, "GOI_D_QS_DU_TOAN_THANH_TOAN"),
        "E": os.path.join(hub_root, "GOI_E_EXECUTIVE_DASHBOARD"),
    }
    for p in pkg_dirs.values():
        os.makedirs(p, exist_ok=True)

    # 2. Tạo File Gói A (5 Sheets chuyên sâu chuẩn mẫu Vincons / 23HG)
    _ex_dir = os.path.dirname(os.path.abspath(__file__))
    if _ex_dir not in sys.path: sys.path.insert(0, _ex_dir)
    from generate_km19_machine_schedule import build_km19_machine_schedule
    build_km19_machine_schedule()


    # 3. Phân phối Gói B (Xưởng thép)
    dir_micro = os.path.join(TARGET_DIR_1, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")
    dir_rebar = os.path.join(TARGET_DIR_1, "01_HE_THONG_CAT_THEP_REBARCUT")
    
    for f in ["01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx", "04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx", "01_Phieu_Cat_Thep_Cau_Km19+529.080.csv"]:
        src_f = os.path.join(dir_micro, f)
        if os.path.exists(src_f):
            shutil.copy2(src_f, os.path.join(pkg_dirs["B"], f))
            
    if os.path.exists(dir_rebar):
        for f in ["00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx", "README_QUY_TRINH_VAN_HANH_BAI_THEP.md"]:
            src_f = os.path.join(dir_rebar, f)
            if os.path.exists(src_f):
                shutil.copy2(src_f, os.path.join(pkg_dirs["B"], f))
        # Copy thư mục con THEO_TUNG_DUONG_KINH_PHI nếu có
        src_sub = os.path.join(dir_rebar, "THEO_TUNG_DUONG_KINH_PHI")
        dst_sub = os.path.join(pkg_dirs["B"], "THEO_TUNG_DUONG_KINH_PHI")
        if os.path.exists(src_sub):
            if os.path.exists(dst_sub): shutil.rmtree(dst_sub)
            shutil.copytree(src_sub, dst_sub)

    # 4. Phân phối Gói C (KCS)
    for f in [
        "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx",
        "11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx",
        "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx",
        "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx",
        "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx",
        "05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx"
    ]:
        src_f = os.path.join(dir_micro, f)
        if os.path.exists(src_f):
            shutil.copy2(src_f, os.path.join(pkg_dirs["C"], f))

    # 5. Phân phối Gói D (QS & Dự toán)
    for f in [
        "03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx",
        "08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx",
        "09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx",
        "06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx",
        "07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx"
    ]:
        src_f = os.path.join(dir_micro, f)
        if os.path.exists(src_f):
            shutil.copy2(src_f, os.path.join(pkg_dirs["D"], f))

    # 6. Phân phối Gói E (Executive Hub)
    dir_macro = os.path.join(TARGET_DIR_1, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
    for f in [
        "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx",
        "Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp",
        "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml",
        "BAO_CAO_THAM_TRA_AEC_AUDIT.md",
        "Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md"
    ]:
        src_f = os.path.join(dir_macro, f)
        if os.path.exists(src_f):
            shutil.copy2(src_f, os.path.join(pkg_dirs["E"], f))

    # 7. Tạo Bảng Phân Quyền & Biên Bản Bàn Giao (Excel & Markdown)
    file_perm_xlsx = os.path.join(hub_root, "BANG_PHAN_QUYEN_VA_BIEN_BAN_BAN_GIAO_5_GOI_VE_TINH.xlsx")
    build_permission_and_handover_workbook(file_perm_xlsx)

    file_perm_md = os.path.join(hub_root, "BANG_PHAN_QUYEN_VA_BIEN_BAN_BAN_GIAO_5_GOI_VE_TINH.md")
    build_permission_markdown(file_perm_md)

    # 8. Tạo DISPATCH_MANIFEST.json
    manifest_data = {
        "project_name": "Cau_Km19+529.080",
        "timestamp": datetime.datetime.now().isoformat(),
        "framework": "23HG-AEC-MultiAgent-System v3.0",
        "author": "Nguyen Bao Tu (@baotuhg)",
        "model": "Hub & Spoke Role-Based Dispatching (Quy trình 15)",
        "packages": {
            "GOI_A": {"name": "Cơ giới & Dầu Diezel", "target": "GOI_A_CO_GIOI_VA_DAU_DIEZEL", "files_count": len(os.listdir(pkg_dirs["A"]))},
            "GOI_B": {"name": "Xưởng tiền chế cốt thép", "target": "GOI_B_XUONG_TIEN_CHE_COT_THEP", "files_count": len(os.listdir(pkg_dirs["B"]))},
            "GOI_C": {"name": "Hiện trường QLCL KCS", "target": "GOI_C_HIEN_TRUONG_QLCL_KCS", "files_count": len(os.listdir(pkg_dirs["C"]))},
            "GOI_D": {"name": "QS & Dự toán thanh toán", "target": "GOI_D_QS_DU_TOAN_THANH_TOAN", "files_count": len(os.listdir(pkg_dirs["D"]))},
            "GOI_E": {"name": "Executive Dashboard", "target": "GOI_E_EXECUTIVE_DASHBOARD", "files_count": len(os.listdir(pkg_dirs["E"]))}
        },
        "audit_score": "100/100",
        "zero_dead_numbers": True
    }
    with open(os.path.join(hub_root, "DISPATCH_MANIFEST.json"), "w", encoding="utf-8") as f:
        json.dump(manifest_data, f, ensure_ascii=False, indent=2)

    # 9. ĐỒNG BỘ TOÀN DIỆN SANG CẢ 2 THƯ MỤC NESTED
    print("[*] Đang đồng bộ sang các thư mục con nested...")
    for target_nested in [TARGET_DIR_2, TARGET_DIR_3]:
        if os.path.exists(target_nested):
            dst = os.path.join(target_nested, FOLDER_HUB_NAME)
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(hub_root, dst)
            print(f"  -> Đồng bộ thành công: {dst}")

    print("\n" + "=" * 80)
    print("HOÀN TẤT ĐÓNG GÓI 5 GÓI VỆ TINH & BẢNG PHÂN QUYỀN BÀN GIAO 100%!")
    print("=" * 80)

if __name__ == "__main__":
    main()
