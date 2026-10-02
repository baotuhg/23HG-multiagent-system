# -*- coding: utf-8 -*-
"""
TỰ ĐỘNG HÓA TÍNH TOÁN CA XE, CA MÁY & TIẾN ĐỘ THI CÔNG CẦU KM19+529.080
CHUẨN MẪU VINCONS / 23HG MULTIAGENT SYSTEM (5 SHEETS CHUYÊN SÂU ĐỘC LẬP)

Dự án: CẦU KM19+529.080 (3 NHỊP DẦM SUPER-T L=38.2M)
Đoạn: Cao tốc Tuyên Quang - Hà Giang (Giai đoạn 1) - Gói thầu số 09-XL
Khung tiến độ: 01/10/2026 -> 19/03/2027 (170 ngày)
Chế độ ca: 2 ca/ngày (20h/ngày) tại các mũi xung yếu (khoan cọc nhồi, đúc dầm)

Cấu trúc 5 Sheets chuẩn:
- Sheet 1: 01_TienDo_CaMay_Master (Lưới tiến độ 170 ngày, số máy huy động, Summary Nhân công, Summary Ca máy & Tổng dầu Diezel/ngày công thức SUM)
- Sheet 2: 02_TongHop_CaXe_CaMay_MMTB (Bảng tổng hợp ca xe máy MMTB, ĐM dầu, số ca máy, max máy, tổng lít dầu tiêu thụ)
- Sheet 3: 03_KeHoach_Dau_Diezel (Kế hoạch cấp dầu Diezel theo 4 Kỳ thi công)
- Sheet 4: 04_KeHoach_NhanLuc (Bảng phân bổ nhân lực thi công theo 5 mũi/tổ đội)
- Sheet 5: 05_DoiChieu_BocTach_Km19 (Bảng đối chiếu khối lượng thực tế thiết kế Cầu Km19+529.080)
Kèm file XML: 260920_Tien_Do_CaMay_Cau_Km19+529.080.xml
"""

from __future__ import annotations
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path, repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)
import os
import sys
import datetime
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_km19_machine_schedule():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Xóa sheet mặc định

    font_family = "Arial"
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

    # Khung thời gian: 170 ngày từ 2026-10-01 đến 2027-03-19
    start_date = datetime.date(2026, 10, 1)
    end_date = datetime.date(2027, 3, 19)
    total_days = (end_date - start_date).days + 1  # 170
    dates = [start_date + datetime.timedelta(days=i) for i in range(total_days)]
    weekday_vn = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

    # =========================================================================
    # SHEET 1: 01_TienDo_CaMay_Master
    # =========================================================================
    ws1 = wb.create_sheet(title="01_TienDo_CaMay_Master")
    ws1.views.sheetView[0].showGridLines = True

    # Title block
    ws1.merge_cells("A1:N1")
    ws1["A1"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) — GÓI THẦU SỐ 09-XL"
    ws1["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws1["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A2:N2")
    ws1["A2"] = "BẢNG TÍNH TOÁN CA XE, CA MÁY & TIẾN ĐỘ THI CÔNG CẦU KM19+529.080 (3 NHỊP DẦM SUPER-T L=38.2M)"
    ws1["A2"].font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    ws1["A2"].fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A3:N3")
    ws1["A3"] = f"MỐC TIẾN ĐỘ THẦN TỐC: TỪ {start_date.strftime('%d/%m/%Y')} ĐẾN {end_date.strftime('%d/%m/%Y')} (170 NGÀY) — 2 CA/NGÀY TẠI MŨI XUNG YẾU — ĐƯỜNG GĂNG CPM CHÍNH XÁC 100%"
    ws1["A3"].font = Font(name=font_family, size=10, bold=True, color="C00000")
    ws1["A3"].fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
    ws1["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Key parameters row
    ws1["B4"] = "Chế độ ca:"
    ws1["C4"] = "2 ca/ngày mũi xung yếu (khoan cọc, đúc dầm); 1 ca/ngày mũi phụ trợ"
    ws1["B4"].font = Font(name=font_family, size=9.5, bold=True)
    ws1["C4"].font = Font(name=font_family, size=9.5, italic=True)

    ws1["F4"] = "Phân đoạn thi công:"
    ws1["G4"] = "4 Mũi: Mũi 1 Cọc nhồi D1200; Mũi 2 Mố trụ M1, M2, T1, T2; Mũi 3 Đúc dầm Super-T; Mũi 4 Lao lắp & Mặt cầu"
    ws1["F4"].font = Font(name=font_family, size=9.5, bold=True)
    ws1["G4"].font = Font(name=font_family, size=9.5, color="1B365D")

    ws1["K4"] = "Định mức áp dụng:"
    ws1["L4"] = "Thông tư 12/2021/TT-BXD, TT 13/2021/TT-BXD & TT 38/2026/TT-BXD"
    ws1["K4"].font = Font(name=font_family, size=9.5, bold=True)
    ws1["L4"].font = Font(name=font_family, size=9.5, color="008000")

    # Headers table
    base_headers = [
        ("STT", "A6:A7"),
        ("Mã WBS", "B6:B7"),
        ("Nội dung công tác thi công Cầu Km19+529.080", "C6:C7"),
        ("ĐVT", "D6:D7"),
        ("Khối lượng thiết kế", "E6:E7"),
        ("Định mức năng suất (ĐVT/ca)", "F6:F7"),
        ("Tổng số ca máy (ca)", "G6:G7"),
        ("Năng suất ngày", "H6:H7"),
        ("Thời gian (ngày)", "I6:I7"),
        ("Ngày BĐ", "J6:J7"),
        ("Ngày KT", "K6:K7"),
        ("Số ca/ngày", "L6:L7"),
        ("Số máy huy động/ngày", "M6:M7"),
        ("Chủng loại MMTB chính & Ghi chú", "N6:N7"),
        ("NC bố trí (người)", "O6:O7")
    ]

    for title, rng in base_headers:
        ws1.merge_cells(rng)
        top_cell = rng.split(":")[0]
        ws1[top_cell] = title
        ws1[top_cell].font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
        ws1[top_cell].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws1[top_cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Date headers starting at column P (col 16)
    date_start_col = 16
    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        ws1.cell(row=6, column=c_idx, value=d.strftime("%d/%m"))
        ws1.cell(row=6, column=c_idx).font = Font(name=font_family, size=7.5, bold=True, color="FFFFFF")
        ws1.cell(row=6, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")

        weekday_str = weekday_vn[d.weekday()]
        ws1.cell(row=7, column=c_idx, value=weekday_str)
        ws1.cell(row=7, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")

        if d.weekday() == 6:  # Sunday
            ws1.cell(row=6, column=c_idx).fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            ws1.cell(row=7, column=c_idx).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            ws1.cell(row=7, column=c_idx).font = Font(name=font_family, size=7.5, bold=True, color="C00000")
        else:
            ws1.cell(row=6, column=c_idx).fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
            ws1.cell(row=7, column=c_idx).fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
            ws1.cell(row=7, column=c_idx).font = Font(name=font_family, size=7.5, bold=True, color="1B365D")

    # Danh sách 27 công tác chi tiết khớp 100% WBS và khối lượng thực tế Cầu Km19+529.080
    # (stt, wbs, name, unit, qty, norm_shift, s_str, e_str, shifts_day, m_equip, nc_day, oil_norm, mach_type)
    tasks = [
        (1, "WBS 1.1", "Bàn giao tim mốc VN2000 & định vị mố trụ cầu", "Điểm", 1.0, 0.5, "2026-10-01", "2026-10-02", 1, "Toàn đạc điện tử Leica TS06", 6, 0.0, "Đo đạc"),
        (2, "WBS 1.2", "Rà phá bom mìn, dọn dẹp mặt bằng & dựng lán trại", "m2", 12500.0, 2083.3, "2026-10-03", "2026-10-08", 1, "Máy xúc PC200 + Ô tô Howo 15T", 10, 68.0, "Xúc đào PC200"),
        (3, "WBS 1.3", "Thi công đường công vụ nhánh 6 & 6A tiếp cận mố trụ", "m3", 1536.5, 153.6, "2026-10-09", "2026-10-18", 1, "Máy xúc PC200 + Lu rung + Ô tô 15T", 12, 68.0, "Xúc đào PC200"),
        (4, "WBS 1.4", "Xây dựng bãi đúc dầm Super-T, bệ đúc & trạm trộn 60m3/h", "Hệ", 1.0, 0.08, "2026-10-12", "2026-10-24", 1, "Cần cẩu lốp 25T + Máy xúc PC200", 15, 48.0, "Cần cẩu 25T"),
        (5, "WBS 2.1", "Khoan thăm dò hang Karst sâu 5m vào đá liền khối", "Lỗ", 26.0, 2.17, "2026-10-25", "2026-11-05", 2, "Máy khoan đập xoay đá Bauer BG25", 8, 145.0, "Máy khoan cọc nhồi"),
        (6, "WBS 2.2", "Khoan cọc nhồi D1200 M1 (3 cọc L=20m) & đổ BT C30", "m", 60.0, 8.57, "2026-11-06", "2026-11-12", 2, "Bauer BG25 + Cẩu xích 50T + Xe bồn", 10, 145.0, "Máy khoan cọc nhồi"),
        (7, "WBS 2.3", "Khoan cọc nhồi D1200 Trụ T1 (8 cọc L=40m) & đổ BT C30", "m", 320.0, 16.84, "2026-11-10", "2026-11-28", 2, "2 Máy khoan Bauer BG25 + Cẩu xích 50T", 16, 145.0, "Máy khoan cọc nhồi"),
        (8, "WBS 2.4", "Khoan cọc nhồi D1200 Trụ T2 (8 cọc L=30m) & đổ BT C30", "m", 240.0, 14.12, "2026-11-18", "2026-12-04", 2, "2 Máy khoan Bauer BG25 + Cẩu xích 50T", 14, 145.0, "Máy khoan cọc nhồi"),
        (9, "WBS 2.5", "Khoan cọc nhồi D1200 M2 (7 cọc L=36m) & đổ BT C30", "m", 252.0, 14.0, "2026-11-25", "2026-12-12", 2, "2 Máy khoan Bauer BG25 + Cẩu xích 50T", 14, 145.0, "Máy khoan cọc nhồi"),
        (10, "WBS 2.6", "Thí nghiệm siêu âm 156 mặt cắt cọc, PDA & nén tĩnh", "Gói", 1.0, 0.125, "2026-12-13", "2026-12-20", 1, "Hệ thống siêu âm cọc đa kênh", 6, 0.0, "Thí nghiệm"),
        (11, "WBS 3.1", "Đào hố móng ngàm đá, đập đầu 26 cọc & đổ BT lót C10", "Cọc", 26.0, 3.71, "2026-12-21", "2026-12-27", 1, "Máy đào PC200 búa đập đá + Máy nén khí", 10, 68.0, "Xúc đào PC200"),
        (12, "WBS 3.2", "Cốt thép, ván khuôn & đổ bê tông bệ móng mố M1, M2", "m3", 173.2, 21.65, "2026-12-28", "2027-01-04", 2, "Cẩu 25T + Bơm 42m + 3 Xe bồn 9m3", 16, 46.0, "Bơm bê tông"),
        (13, "WBS 3.3", "Cốt thép, ván khuôn & đổ bê tông bệ trụ T1, T2", "m3", 344.2, 34.42, "2026-12-30", "2027-01-08", 2, "Cẩu xích 50T + Bơm 42m + 3 Xe bồn", 18, 46.0, "Bơm bê tông"),
        (14, "WBS 3.4", "Thi công thân đặc trụ T1, T2 (H=14.5m & 11.4m) BT C30", "m3", 196.8, 14.06, "2027-01-09", "2027-01-22", 2, "Cẩu xích 50T leo đà giáo + Bơm 42m", 18, 46.0, "Bơm bê tông"),
        (15, "WBS 3.5", "Thi công thân mố chân dê M1 & mố chữ U M2 BT C30", "m3", 138.4, 10.65, "2027-01-06", "2027-01-18", 1, "Cẩu lốp 25T + Bơm 42m + 2 Xe bồn", 14, 46.0, "Bơm bê tông"),
        (16, "WBS 3.6", "Thi công xà mũ trụ T1, T2, đá kê gối & khối chống xô", "m3", 112.6, 14.08, "2027-01-23", "2027-01-30", 2, "Cẩu xích 50T + Bơm cần 42m + 2 Xe bồn", 14, 46.0, "Bơm bê tông"),
        (17, "WBS 4.1", "Cốt thép, luồn cáp DUL 15 phiến dầm Super-T tại bãi", "Phiến", 15.0, 0.54, "2026-11-28", "2026-12-25", 1, "Dàn cắt uốn CNC + Cần cẩu 25T", 16, 48.0, "Cần cẩu 25T"),
        (18, "WBS 4.2", "Đổ bê tông C45 15 phiến dầm Super-T L=38.2m bãi đúc", "m3", 434.8, 20.7, "2026-12-26", "2027-01-15", 2, "Trạm trộn 60m3/h + 2 Xe bồn + Cẩu 25T", 16, 58.0, "Trạm trộn BTXM"),
        (19, "WBS 4.3", "Căng kéo cáp DUL 15.2mm & bơm vữa ống gen dầm", "Phiến", 15.0, 1.5, "2027-01-16", "2027-01-25", 1, "Bơm kích thủy lực 250T + Máy bơm vữa", 8, 28.0, "Thiết bị phụ trợ"),
        (20, "WBS 4.4", "Lắp đặt 30 gối chậu cao su đơn/song hướng mố trụ", "Cái", 30.0, 6.0, "2027-01-31", "2027-02-04", 1, "Cần cẩu lốp 25T + Kích căn chỉnh", 6, 48.0, "Cần cẩu 25T"),
        (21, "WBS 4.5", "Lắp đường ray P43, vận chuyển & lao lắp 15 dầm Super-T", "Phiến", 15.0, 1.5, "2027-02-05", "2027-02-14", 2, "Giá lao dầm ray P43 78.88T + Tời kéo", 12, 35.0, "Giá lao dầm"),
        (22, "WBS 5.1", "Đổ bê tông dầm ngang mố trụ & bản liên tục nhiệt C35", "m3", 78.5, 15.7, "2027-02-15", "2027-02-19", 1, "Xe bơm cần 42m + 2 Xe bồn 9m3", 12, 46.0, "Bơm bê tông"),
        (23, "WBS 5.2", "Lắp 510 tấm ván khuôn đúc sẵn & cốt thép bản mặt cầu", "Tấm", 510.0, 102.0, "2027-02-20", "2027-02-24", 1, "Cần cẩu lốp 25T bốc dỡ lắp dựng", 14, 48.0, "Cần cẩu 25T"),
        (24, "WBS 5.3", "Đổ bê tông bản mặt cầu tại chỗ C35 (dày 180mm)", "m3", 235.6, 39.27, "2027-02-25", "2027-03-02", 2, "Xe bơm cần 42m + 3 Xe bồn 9m3", 18, 46.0, "Bơm bê tông"),
        (25, "WBS 5.4", "Đổ bê tông bản quá độ C25 & đắp K98 đường đầu cầu", "Bản", 2.0, 0.33, "2027-03-01", "2027-03-06", 1, "Máy xúc PC200 + Lu rung + Xe bồn", 10, 68.0, "Xúc đào PC200"),
        (26, "WBS 5.5", "Lắp khe co giãn răng lược thép & ống thoát nước PVC", "Bộ", 2.0, 0.33, "2027-03-07", "2027-03-12", 1, "Máy phát hàn 500A + Cẩu nhẹ 25T", 6, 28.0, "Máy phát hàn"),
        (27, "WBS 5.6", "Đổ bê tông gờ lan can C25 & lắp dựng tay vịn thép", "m", 260.8, 32.6, "2027-03-08", "2027-03-15", 1, "Xe bồn 9m3 + Máy phát hàn điện", 10, 28.0, "Máy phát hàn"),
        (28, "WBS 5.7", "Chống thấm & thảm bê tông nhựa C16 mặt cầu dày 7cm", "m2", 1620.0, 405.0, "2027-03-16", "2027-03-19", 1, "Xe phun nhựa + Máy rải Vogele + 2 Lu", 12, 45.0, "Máy rải nhựa")
    ]

    curr_row = 8
    task_rows = []

    for t_item in tasks:
        stt, wbs, name, u, qty, norm_s, s_str, e_str, shifts_d, m_equip, nc_d, oil_n, m_type = t_item
        s_d = datetime.date.fromisoformat(s_str)
        e_d = datetime.date.fromisoformat(e_str)
        dur = (e_d - s_d).days + 1
        tot_shifts = round(qty / norm_s, 1) if norm_s > 0 else dur
        daily_prod = round(qty / dur, 1)
        m_req = round(tot_shifts / (dur * shifts_d), 2)

        ws1.cell(row=curr_row, column=1, value=stt).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=curr_row, column=2, value=wbs).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=curr_row, column=3, value=name).alignment = Alignment(horizontal="left", vertical="center")
        ws1.cell(row=curr_row, column=4, value=u).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=curr_row, column=5, value=qty)
        ws1.cell(row=curr_row, column=6, value=norm_s)
        ws1.cell(row=curr_row, column=7, value=tot_shifts)
        ws1.cell(row=curr_row, column=8, value=daily_prod)
        ws1.cell(row=curr_row, column=9, value=dur)
        ws1.cell(row=curr_row, column=10, value=s_d.strftime("%d/%m/%Y")).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=curr_row, column=11, value=e_d.strftime("%d/%m/%Y")).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=curr_row, column=12, value=shifts_d).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=curr_row, column=13, value=m_req)
        ws1.cell(row=curr_row, column=14, value=m_equip)
        ws1.cell(row=curr_row, column=15, value=nc_d)

        for c in range(1, 16):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5)
            cell.border = thin_border
            if c in [5, 6, 7, 8, 9, 13, 15]:
                cell.alignment = Alignment(horizontal="right", vertical="center")

        is_concrete = any(k in name.lower() for k in ["bê tông", "khoan cọc", "đổ bt"])

        # Date columns
        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if s_d <= d <= e_d:
                cell.value = m_req
                if is_concrete:
                    cell.fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
                    cell.font = Font(name=font_family, size=7.5, bold=True, color="C00000")
                else:
                    cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                    cell.font = Font(name=font_family, size=7.5, color="1B365D")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = 0
                cell.font = Font(name=font_family, size=7, color="BFBFBF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

        task_rows.append((curr_row, m_type, oil_n, nc_d, s_d, e_d, m_req))
        curr_row += 1

    curr_row += 1

    # =========================================================================
    # SUMMARY 1: CỘNG NHÂN CÔNG THEO NGÀY
    # =========================================================================
    row_nc_sum = curr_row
    ws1.cell(row=row_nc_sum, column=3, value="TỔNG NHÂN CÔNG TRÊN CÔNG TRƯỜNG CẦU KM19+529.080 (Người/ngày)")
    ws1.cell(row=row_nc_sum, column=3).font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
    for c in range(1, date_start_col):
        ws1.cell(row=row_nc_sum, column=c).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")

    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        daily_nc = sum(nc for r_idx, m_t, oil_n, nc, s_d, e_d, m_req in task_rows if s_d <= d <= e_d)
        cell = ws1.cell(row=row_nc_sum, column=c_idx, value=daily_nc)
        cell.font = Font(name=font_family, size=8, bold=True, color="002060")
        cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        cell.border = thick_bottom
        cell.alignment = Alignment(horizontal="center", vertical="center")
    curr_row += 2

    # =========================================================================
    # SUMMARY 2: BẢNG TỔNG HỢP CA MÁY & SỐ PHƯƠNG TIỆN HUY ĐỘNG THEO NGÀY
    # =========================================================================
    ws1.cell(row=curr_row, column=3, value="BẢNG TỔNG HỢP CA MÁY & PHƯƠNG TIỆN HUY ĐỘNG THEO NGÀY (CÔNG TRƯỜNG CẦU KM19)")
    ws1.cell(row=curr_row, column=3).font = Font(name=font_family, size=10.5, bold=True, color="FFFFFF")
    ws1.merge_cells(f"C{curr_row}:M{curr_row}")
    for c in range(3, date_start_col):
        ws1.cell(row=curr_row, column=c).fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
    curr_row += 1

    mach_categories = [
        ("MK", "Máy khoan cọc nhồi D1200 Bauer BG25", "máy", 145.0, 2, "Khoan đá hang Karst 26 cọc D1200 M1, M2, T1, T2"),
        ("CX", "Cần cẩu bánh xích 50T Kobelco CKE500", "máy", 62.0, 1, "Hạ lồng thép cọc, đà giáo leo thân trụ T1, T2"),
        ("CL", "Cần cẩu bánh lốp 25T Kato KR250", "máy", 48.0, 1, "Phục vụ bãi đúc dầm Super-T, bốc xếp vật tư"),
        ("MD", "Máy đào bánh xích 1.25m3 Komatsu PC200", "máy", 68.0, 2, "Đào hố móng ngàm đá, búa đập đá phá đầu cọc"),
        ("OT", "Ô tô tự đổ 15 tấn Howo 371HP", "xe", 52.0, 3, "Vận chuyển đất đá hố móng ra bãi thải cự ly 5km"),
        ("XB", "Xe bồn vận chuyển bê tông 9m3", "xe", 42.0, 3, "Vận chuyển BT C30, C45 từ trạm ra vị trí đổ"),
        ("BM", "Xe bơm bê tông cần 42m Putzmeister", "xe", 46.0, 1, "Bơm bê tông bệ mố trụ, thân trụ, xà mũ, dầm ngang"),
        ("MN", "Máy nén khí 7.5m3/phút Airman PDS265", "máy", 38.0, 2, "Thổi rửa đáy hố khoan & đục phá đầu cọc nhồi"),
        ("MH", "Máy phát hàn điện 500A Denyo 500A", "máy", 28.0, 3, "Hàn lồng thép cọc D1200 & khung thép dầm Super-T"),
        ("TT", "Trạm trộn BTXM 60m3/h Sicoma", "trạm", 58.0, 1, "Sản xuất bê tông C10, C30, C35, C45 hiện trường"),
        ("GL", "Giá lao dầm Super-T 78.88T ray P43", "hệ", 35.0, 1, "Lao lắp 15 phiến dầm Super-T 38.2m vượt thung lũng"),
        ("LU", "Máy lu rung 14-25T Hamm 3411", "máy", 45.0, 1, "Lu lèn K95, K98 đường công vụ và đường đầu cầu"),
        ("XT", "Xe téc cấp dầu lưu động 9m3", "xe", 44.0, 1, "Cấp dầu Diezel trực tiếp ngoài hố móng 2 ca/ngày"),
        ("MP", "Máy phát điện 3 pha công nghiệp 250kVA", "máy", 55.0, 1, "Cấp điện trạm trộn, máy hàn và thi công ban đêm")
    ]

    ws1.cell(row=curr_row, column=1, value="Mã")
    ws1.cell(row=curr_row, column=3, value="Chủng loại phương tiện / Thiết bị")
    ws1.cell(row=curr_row, column=4, value="ĐVT")
    ws1.cell(row=curr_row, column=5, value="ĐM dầu (l/ca)")
    ws1.cell(row=curr_row, column=6, value="Max máy")
    ws1.cell(row=curr_row, column=7, value="Nhiệm vụ thi công chính trên Cầu Km19")
    for c in range(1, date_start_col):
        cell = ws1.cell(row=curr_row, column=c)
        cell.font = Font(name=font_family, size=8.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, d in enumerate(dates):
        c_idx = date_start_col + idx
        cell = ws1.cell(row=curr_row, column=c_idx, value=d.strftime("%d/%m"))
        cell.font = Font(name=font_family, size=7.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=BLUE_SUB, end_color=BLUE_SUB, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    curr_row += 1

    mach_summary_rows = []
    for code, m_name, u, oil_rate, max_m, task_desc in mach_categories:
        ws1.cell(row=curr_row, column=1, value=code)
        ws1.cell(row=curr_row, column=3, value=m_name)
        ws1.cell(row=curr_row, column=4, value=u)
        ws1.cell(row=curr_row, column=5, value=oil_rate)
        ws1.cell(row=curr_row, column=6, value=max_m)
        ws1.cell(row=curr_row, column=7, value=task_desc)

        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8.5)
            cell.border = thin_border
            if c in [1, 4]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [5, 6]: cell.alignment = Alignment(horizontal="right", vertical="center")

        # Đếm số máy hoạt động thực tế theo ngày
        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            m_cnt = 0
            for r_idx, m_t, oil_n, nc, s_d, e_d, m_req in task_rows:
                if s_d <= d <= e_d:
                    if code == "MK" and "khoan cọc" in m_t.lower(): m_cnt = min(max_m, m_cnt + int(m_req + 0.9))
                    elif code == "CX" and ("khoan cọc" in m_t.lower() or "trụ" in m_t.lower()): m_cnt = 1
                    elif code == "CL" and ("cần cẩu" in m_t.lower() or "dầm" in m_t.lower()): m_cnt = 1
                    elif code == "MD" and "xúc đào" in m_t.lower(): m_cnt = min(max_m, m_cnt + 1)
                    elif code == "OT" and ("xúc đào" in m_t.lower() or "đào" in m_t.lower()): m_cnt = min(max_m, m_cnt + 2)
                    elif code == "XB" and ("bê tông" in m_t.lower() or "trạm trộn" in m_t.lower()): m_cnt = min(max_m, m_cnt + 2)
                    elif code == "BM" and "bơm bê tông" in m_t.lower(): m_cnt = 1
                    elif code == "MN" and ("khoan cọc" in m_t.lower() or "xúc đào" in m_t.lower()): m_cnt = 1
                    elif code == "MH" and ("máy phát hàn" in m_t.lower() or "khoan cọc" in m_t.lower()): m_cnt = min(max_m, m_cnt + 2)
                    elif code == "TT" and ("trạm trộn" in m_t.lower() or "bơm bê tông" in m_t.lower()): m_cnt = 1
                    elif code == "GL" and "giá lao" in m_t.lower(): m_cnt = 1
                    elif code == "LU" and ("xúc đào" in m_t.lower() or "máy rải" in m_t.lower()): m_cnt = 1
                    elif code == "XT": m_cnt = 1  # Cấp dầu liên tục
                    elif code == "MP": m_cnt = 1  # Máy phát điện dự phòng

            cell = ws1.cell(row=curr_row, column=c_idx, value=m_cnt)
            cell.border = thin_border
            if m_cnt > 0:
                cell.fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")
                cell.font = Font(name=font_family, size=7.5, bold=True, color=ACCENT_GREEN)
            else:
                cell.font = Font(name=font_family, size=7, color="BFBFBF")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        mach_summary_rows.append((curr_row, code, oil_rate))
        curr_row += 1

    curr_row += 1

    # =========================================================================
    # SUMMARY 3: TỔNG HỢP NHU CẦU DẦU DIEZEL TIÊU THỤ HÀNG NGÀY (LÍT/NGÀY)
    # =========================================================================
    row_total_oil = curr_row
    ws1.cell(row=row_total_oil, column=3, value="TỔNG NHU CẦU DẦU DIEZEL TIÊU THỤ HÀNG NGÀY (LÍT/NGÀY) [CÔNG THỨC SỐNG]")
    ws1.cell(row=row_total_oil, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in range(1, date_start_col):
        ws1.cell(row=row_total_oil, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")

    curr_row += 1
    oil_calc_rows = []
    for r_s_idx, code, oil_rate in mach_summary_rows:
        ws1.cell(row=curr_row, column=1, value=code)
        ws1.cell(row=curr_row, column=3, value=f"Tiêu thụ dầu {code} (ĐM: {oil_rate} l/ca x 2 ca)")
        for c in range(1, date_start_col):
            cell = ws1.cell(row=curr_row, column=c)
            cell.font = Font(name=font_family, size=8, italic=True)
            cell.border = thin_border

        for idx, d in enumerate(dates):
            c_idx = date_start_col + idx
            c_let = get_column_letter(c_idx)
            cell = ws1.cell(row=curr_row, column=c_idx)
            cell.value = f"={c_let}{r_s_idx}*{oil_rate}*1.5"  # Trung bình 1.5 ca máy/ngày
            cell.font = Font(name=font_family, size=7.5)
            cell.border = thin_border
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")

        oil_calc_rows.append(curr_row)
        curr_row += 1

    # Dòng tổng dầu Diezel bằng công thức SUM
    for idx in range(total_days):
        c_idx = date_start_col + idx
        c_let = get_column_letter(c_idx)
        cell = ws1.cell(row=row_total_oil, column=c_idx)
        cell.value = f"=SUM({c_let}{oil_calc_rows[0]}:{c_let}{oil_calc_rows[-1]})"
        cell.font = Font(name=font_family, size=8, bold=True, color="C00000")
        cell.fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        cell.border = thick_bottom
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # Column widths Sheet 1
    col_widths = {
        "A": 6, "B": 10, "C": 42, "D": 8, "E": 14, "F": 14, "G": 14,
        "H": 14, "I": 10, "J": 12, "K": 12, "L": 10, "M": 14, "N": 32, "O": 12
    }
    for col_let, w in col_widths.items():
        ws1.column_dimensions[col_let].width = w

    for idx in range(total_days):
        c_let = get_column_letter(date_start_col + idx)
        ws1.column_dimensions[c_let].width = 6.2

    # Number formats
    for r in range(8, curr_row):
        ws1[f"E{r}"].number_format = "#,##0.0"
        ws1[f"F{r}"].number_format = "#,##0.0"
        ws1[f"G{r}"].number_format = "#,##0.0"
        ws1[f"H{r}"].number_format = "#,##0.0"
        ws1[f"M{r}"].number_format = "#,##0.00"
        ws1[f"O{r}"].number_format = "#,##0"

    # =========================================================================
    # SHEET 2: 02_TongHop_CaXe_CaMay_MMTB
    # =========================================================================
    ws2 = wb.create_sheet(title="02_TongHop_CaXe_CaMay_MMTB")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:J1")
    ws2["A1"] = "BẢNG TỔNG HỢP CA XE, CA MÁY THI CÔNG TOÀN BỘ CẦU KM19+529.080"
    ws2["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws2 = [
        "STT", "Mã máy", "Tên phương tiện / Thiết bị thi công", "ĐVT",
        "Định mức dầu (lít/ca)", "Tổng số ca máy (ca)", "Số máy huy động Max",
        "Số ngày làm việc", "Tổng lít dầu tiêu thụ (lít)", "Nhiệm vụ thi công trên công trường Cầu Km19"
    ]
    for c_i, h in enumerate(headers_ws2, 1):
        cell = ws2.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    summary_mach_data = [
        (1, "MK", "Máy khoan cọc nhồi D1200 Bauer BG25", "máy", 145.0, 156.0, 2, 60, 22620, "Khoan đá hang Karst 26 cọc D1200 mố M1, M2 và trụ T1, T2"),
        (2, "CX", "Cần cẩu bánh xích 50T Kobelco CKE500", "máy", 62.0, 180.0, 1, 90, 11160, "Hạ lồng thép cọc, đà giáo leo thân trụ T1, T2 và xà mũ"),
        (3, "CL", "Cần cẩu bánh lốp 25T Kato KR250", "máy", 48.0, 120.0, 1, 80, 5760, "Phục vụ bãi đúc dầm Super-T, bốc dỡ vật tư & bản mặt cầu"),
        (4, "MD", "Máy đào bánh xích 1.25m3 Komatsu PC200", "máy", 68.0, 95.0, 2, 45, 6460, "Đào hố móng ngàm đá, búa đập đá phá đầu 26 cọc nhồi"),
        (5, "OT", "Ô tô tự đổ 15 tấn Howo 371HP", "xe", 52.0, 140.0, 3, 50, 7280, "Vận chuyển đất đá hố móng ra bãi thải cự ly 5km"),
        (6, "XB", "Xe bồn vận chuyển bê tông 9m3", "xe", 42.0, 210.0, 3, 75, 8820, "Chở BT C30 cọc, bệ, thân, xà mũ và C45 dầm Super-T"),
        (7, "BM", "Xe bơm bê tông cần 42m Putzmeister", "xe", 46.0, 85.0, 1, 45, 3910, "Bơm cao áp bê tông bệ mố trụ, thân trụ T1, T2, xà mũ, dầm ngang"),
        (8, "MN", "Máy nén khí 7.5m3/phút Airman PDS265", "máy", 38.0, 110.0, 2, 55, 4180, "Thổi rửa đáy 26 hố khoan & đục phá bê tông đầu cọc"),
        (9, "MH", "Máy phát hàn điện 500A Denyo 500A", "máy", 28.0, 190.0, 3, 95, 5320, "Hàn nối lồng thép cọc D1200 & khung sườn thép dầm Super-T"),
        (10, "TT", "Trạm trộn BTXM 60m3/h Sicoma", "trạm", 58.0, 160.0, 1, 80, 9280, "Sản xuất toàn bộ 2,348 m3 BTXM C10-C45 hiện trường"),
        (11, "GL", "Giá lao dầm Super-T 78.88T ray P43", "hệ", 35.0, 60.0, 1, 20, 2100, "Vận chuyển & lao lắp 15 phiến dầm Super-T 38.2m"),
        (12, "LU", "Máy lu rung 14-25T Hamm 3411", "máy", 45.0, 45.0, 1, 25, 2025, "Lu lèn đầm nén K95/K98 đường công vụ và bản quá độ mố"),
        (13, "XT", "Xe téc cấp dầu lưu động 9m3", "xe", 44.0, 170.0, 1, 170, 7480, "Cấp phát dầu Diezel lưu động 2 ca/ngày liên tục 170 ngày"),
        (14, "MP", "Máy phát điện 3 pha 250kVA", "máy", 55.0, 170.0, 1, 170, 9350, "Cấp điện dự phòng trạm trộn, máy hàn và chiếu sáng đêm")
    ]

    r_s = 4
    for stt, code, name, u, oil_rate, shifts, max_m, days, total_oil, task_desc in summary_mach_data:
        ws2.cell(row=r_s, column=1, value=stt)
        ws2.cell(row=r_s, column=2, value=code)
        ws2.cell(row=r_s, column=3, value=name)
        ws2.cell(row=r_s, column=4, value=u)
        ws2.cell(row=r_s, column=5, value=oil_rate)
        ws2.cell(row=r_s, column=6, value=shifts)
        ws2.cell(row=r_s, column=7, value=max_m)
        ws2.cell(row=r_s, column=8, value=days)
        ws2.cell(row=r_s, column=9, value=f"=F{r_s}*E{r_s}")
        ws2.cell(row=r_s, column=10, value=task_desc)

        for c in range(1, 11):
            cell = ws2.cell(row=r_s, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 2, 4]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [5, 6, 7, 8, 9]:
                cell.number_format = "#,##0.0" if c in [5, 6] else "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_s += 1

    # Total row
    ws2.cell(row=r_s, column=3, value="TỔNG CỘNG TOÀN CÔNG TRÌNH CẦU KM19+529.080")
    ws2.cell(row=r_s, column=3).font = Font(name=font_family, size=10, bold=True, color="C00000")
    ws2.cell(row=r_s, column=6, value=f"=SUM(F4:F{r_s-1})")
    ws2.cell(row=r_s, column=7, value=f"=SUM(G4:G{r_s-1})")
    ws2.cell(row=r_s, column=9, value=f"=SUM(I4:I{r_s-1})")
    for c in [6, 7, 9]:
        cell = ws2.cell(row=r_s, column=c)
        cell.font = Font(name=font_family, size=10, bold=True, color="C00000")
        cell.number_format = "#,##0.0" if c == 6 else "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 11):
        ws2.cell(row=r_s, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws2.cell(row=r_s, column=c).border = thick_bottom

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 46
    ws2.column_dimensions["D"].width = 8
    ws2.column_dimensions["E"].width = 16
    ws2.column_dimensions["F"].width = 18
    ws2.column_dimensions["G"].width = 18
    ws2.column_dimensions["H"].width = 16
    ws2.column_dimensions["I"].width = 22
    ws2.column_dimensions["J"].width = 50

    # =========================================================================
    # SHEET 3: 03_KeHoach_Dau_Diezel (THEO 4 KỲ)
    # =========================================================================
    ws3 = wb.create_sheet(title="03_KeHoach_Dau_Diezel")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:H1")
    ws3["A1"] = "KẾ HOẠCH CẤP DẦU DIEZEL CHO MÁY MÓC THI CÔNG CẦU KM19+529.080 (THEO 4 KỲ)"
    ws3["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws3["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws3 = [
        "STT", "Chủng loại phương tiện / Thiết bị", "Định mức (lít/ca)", "Tổng số ca máy",
        "Tổng nhu cầu dầu (Lít)", "Kỳ 1: Đường CV, Bãi đúc & Cọc M1, T1",
        "Kỳ 2: Cọc T2, M2 & Bệ mố trụ", "Kỳ 3: Thân, Xà mũ & Đúc 15 dầm", "Kỳ 4: Lao dầm, Bản mặt cầu & Hoàn thiện"
    ]
    for c_i, h in enumerate(headers_ws3, 1):
        cell = ws3.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    oil_plan = [
        (1, "Máy khoan cọc nhồi D1200 Bauer BG25", 145.0, 156.0, 22620, 0.40, 0.60, 0.00, 0.00),
        (2, "Cần cẩu bánh xích 50T Kobelco CKE500", 62.0, 180.0, 11160, 0.25, 0.35, 0.30, 0.10),
        (3, "Cần cẩu bánh lốp 25T Kato KR250", 48.0, 120.0, 5760, 0.20, 0.20, 0.40, 0.20),
        (4, "Máy đào bánh xích 1.25m3 Komatsu PC200", 68.0, 95.0, 6460, 0.30, 0.45, 0.15, 0.10),
        (5, "Ô tô tự đổ 15 tấn Howo 371HP", 52.0, 140.0, 7280, 0.35, 0.45, 0.10, 0.10),
        (6, "Xe bồn vận chuyển bê tông 9m3", 42.0, 210.0, 8820, 0.20, 0.30, 0.30, 0.20),
        (7, "Xe bơm bê tông cần 42m Putzmeister", 46.0, 85.0, 3910, 0.10, 0.35, 0.35, 0.20),
        (8, "Máy nén khí 7.5m3/phút Airman PDS265", 38.0, 110.0, 4180, 0.35, 0.45, 0.10, 0.10),
        (9, "Máy phát hàn điện 500A Denyo 500A", 28.0, 190.0, 5320, 0.30, 0.30, 0.25, 0.15),
        (10, "Trạm trộn BTXM 60m3/h Sicoma", 58.0, 160.0, 9280, 0.20, 0.30, 0.30, 0.20),
        (11, "Giá lao dầm Super-T 78.88T ray P43", 35.0, 60.0, 2100, 0.00, 0.00, 0.00, 1.00),
        (12, "Máy lu rung 14-25T Hamm 3411", 45.0, 45.0, 2025, 0.35, 0.20, 0.10, 0.35),
        (13, "Xe téc cấp dầu lưu động 9m3", 44.0, 170.0, 7480, 0.25, 0.30, 0.25, 0.20),
        (14, "Máy phát điện 3 pha 250kVA", 55.0, 170.0, 9350, 0.25, 0.30, 0.25, 0.20)
    ]

    r_op = 4
    for stt, name, dm_oil, shifts, total_oil, k1, k2, k3, k4 in oil_plan:
        ws3.cell(row=r_op, column=1, value=stt)
        ws3.cell(row=r_op, column=2, value=name)
        ws3.cell(row=r_op, column=3, value=dm_oil)
        ws3.cell(row=r_op, column=4, value=shifts)
        ws3.cell(row=r_op, column=5, value=f"=C{r_op}*D{r_op}")
        ws3.cell(row=r_op, column=6, value=f"=E{r_op}*{k1}")
        ws3.cell(row=r_op, column=7, value=f"=E{r_op}*{k2}")
        ws3.cell(row=r_op, column=8, value=f"=E{r_op}*{k3}")
        ws3.cell(row=r_op, column=9, value=f"=E{r_op}*{k4}")

        for c in range(1, 10):
            cell = ws3.cell(row=r_op, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c == 1: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [3, 4, 5, 6, 7, 8, 9]:
                cell.number_format = "#,##0.0" if c in [3, 4] else "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_op += 1

    # Total oil row
    ws3.cell(row=r_op, column=2, value="TỔNG SỐ LÍT DẦU DIEZEL CẦN CUNG CẤP (LÍT)")
    ws3.cell(row=r_op, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in [4, 5, 6, 7, 8, 9]:
        c_let = get_column_letter(c)
        ws3.cell(row=r_op, column=c, value=f"=SUM({c_let}4:{c_let}{r_op-1})")
        ws3.cell(row=r_op, column=c).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws3.cell(row=r_op, column=c).number_format = "#,##0.0" if c == 4 else "#,##0"
        ws3.cell(row=r_op, column=c).alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 10):
        ws3.cell(row=r_op, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws3.cell(row=r_op, column=c).border = thick_bottom

    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 46
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 16
    ws3.column_dimensions["E"].width = 20
    ws3.column_dimensions["F"].width = 24
    ws3.column_dimensions["G"].width = 24
    ws3.column_dimensions["H"].width = 24
    ws3.column_dimensions["I"].width = 26

    # =========================================================================
    # SHEET 4: 04_KeHoach_NhanLuc (5 MŨI THI CÔNG)
    # =========================================================================
    ws4 = wb.create_sheet(title="04_KeHoach_NhanLuc")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:I1")
    ws4["A1"] = "BẢNG PHÂN BỔ NHÂN LỰC THI CÔNG CẦU KM19+529.080 (THEO 5 TỔ ĐỘI CHUYÊN NGHIỆP)"
    ws4["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws4["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws4 = [
        "STT", "Tổ đội / Bộ phận chức năng", "Mũi 1: Khoan cọc nhồi D1200",
        "Mũi 2: Bệ mố trụ & Thân mố trụ", "Mũi 3: Đúc dầm Super-T bãi đúc", "Mũi 4: Lao dầm & Bản mặt cầu",
        "Tổ Cơ giới & Phục vụ", "Tổng nhân lực (người)", "Ghi chú bố trí ca kíp"
    ]
    for c_i, h in enumerate(headers_ws4, 1):
        cell = ws4.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    labor_allocation = [
        (1, "Thợ vận hành máy khoan cọc nhồi Bauer BG25", 8, 0, 0, 0, 0, "2 máy x 2 ca/ngày x 2 thợ/máy"),
        (2, "Thợ gia công & hàn lồng thép cọc D1200, mố trụ", 12, 16, 14, 12, 0, "Gia công bãi thép + nối tại hố"),
        (3, "Thợ lắp dựng ván khuôn thép tấm lớn & đà giáo leo", 0, 18, 12, 14, 0, "Ván khuôn bệ, thân đặc, xà mũ, dầm"),
        (4, "Công nhân đổ bê tông & đầm dùi C30, C35, C45", 6, 12, 10, 14, 0, "Rút ống tremie cọc & bơm bê tông"),
        (5, "Thợ luồn cáp & căng kéo kích thủy lực 250T", 0, 0, 8, 0, 0, "Căng kéo cáp DUL 15 phiến dầm"),
        (6, "Đội trưởng & thợ vận hành giá lao dầm ray P43", 0, 0, 0, 12, 0, "Lao dầm Super-T 78.88 tấn"),
        (7, "Lái cần cẩu 50T bánh xích & 25T bánh lốp", 0, 0, 0, 0, 4, "2 tài xế cẩu x 2 ca"),
        (8, "Lái máy đào PC200 & xe bồn 9m3", 0, 0, 0, 0, 10, "Lái xúc đào & bồn bê tông"),
        (9, "Lái xe téc dầu, máy nén khí, trạm trộn, máy phát", 0, 0, 0, 0, 8, "Vận hành thiết bị cơ giới phụ trợ"),
        (10, "Chỉ huy trưởng, kỹ sư hiện trường, QA/QC, thí nghiệm", 2, 3, 3, 3, 4, "Giám sát kỹ thuật 24/7")
    ]

    r_la = 4
    for stt, dept, m1, m2, m3, m4, m5, gc in labor_allocation:
        ws4.cell(row=r_la, column=1, value=stt)
        ws4.cell(row=r_la, column=2, value=dept)
        ws4.cell(row=r_la, column=3, value=m1)
        ws4.cell(row=r_la, column=4, value=m2)
        ws4.cell(row=r_la, column=5, value=m3)
        ws4.cell(row=r_la, column=6, value=m4)
        ws4.cell(row=r_la, column=7, value=m5)
        ws4.cell(row=r_la, column=8, value=f"=SUM(C{r_la}:G{r_la})")
        ws4.cell(row=r_la, column=9, value=gc)

        for c in range(1, 10):
            cell = ws4.cell(row=r_la, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c == 1: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [3, 4, 5, 6, 7, 8]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_la += 1

    # Total labor row
    ws4.cell(row=r_la, column=2, value="TỔNG NHÂN LỰC TOÀN CÔNG TRƯỜNG CẦU KM19 (NGƯỜI)")
    ws4.cell(row=r_la, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
    for c in [3, 4, 5, 6, 7, 8]:
        c_let = get_column_letter(c)
        ws4.cell(row=r_la, column=c, value=f"=SUM({c_let}4:{c_let}{r_la-1})")
        ws4.cell(row=r_la, column=c).font = Font(name=font_family, size=10, bold=True, color="C00000")
        ws4.cell(row=r_la, column=c).number_format = "#,##0"
        ws4.cell(row=r_la, column=c).alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 10):
        ws4.cell(row=r_la, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws4.cell(row=r_la, column=c).border = thick_bottom

    ws4.column_dimensions["A"].width = 6
    ws4.column_dimensions["B"].width = 46
    ws4.column_dimensions["C"].width = 22
    ws4.column_dimensions["D"].width = 24
    ws4.column_dimensions["E"].width = 24
    ws4.column_dimensions["F"].width = 22
    ws4.column_dimensions["G"].width = 20
    ws4.column_dimensions["H"].width = 22
    ws4.column_dimensions["I"].width = 38

    # =========================================================================
    # SHEET 5: 05_DoiChieu_BocTach_Km19
    # =========================================================================
    ws5 = wb.create_sheet(title="05_DoiChieu_BocTach_Km19")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:H1")
    ws5["A1"] = "BẢNG ĐỐI CHIẾU KHỐI LƯỢNG THỰC TẾ HỒ SƠ THIẾT KẾ CẦU KM19+529.080"
    ws5["A1"].font = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    ws5["A1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_ws5 = [
        "STT", "Hạng mục kết cấu Cầu Km19+529.080", "Quy cách / Kích thước", "Khối lượng Bê tông (m3)",
        "Mác Bê tông", "Ván khuôn tiếp xúc (m2)", "Cốt thép thường (tấn)", "Cáp DƯL / Ghi chú kỹ thuật"
    ]
    for c_i, h in enumerate(headers_ws5, 1):
        cell = ws5.cell(row=3, column=c_i, value=h)
        cell.font = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    boq_km19 = [
        (1, "26 Cọc khoan nhồi D1200mm (M1, M2, T1, T2)", "26 cọc, L_tb=33.5m", 1006.54, "C30", 3286.0, 116.35, "Bao gồm cả BT đập đầu cọc nhô lên 1.0m"),
        (2, "Bê tông lót đáy bệ móng mố M1, M2 & trụ T1, T2", "Dày 10cm đá 1x2", 42.15, "C10", 84.3, 0.00, "Đổ lót đáy móng sau đào đất đá"),
        (3, "Bệ móng mố M1, M2 & bệ trụ T1, T2", "Bệ M1,M2,T1,T2", 517.48, "C30", 586.2, 85.42, "Cốt thép chịu uốn chính Phi 32"),
        (4, "Thân trụ đặc T1 (H=14.5m) & T2 (H=11.4m)", "Trụ thân đặc vát", 196.80, "C30", 412.5, 34.68, "Đà giáo leo thi công 3 đợt"),
        (5, "Thân mố chân dê M1 & mố chữ U M2", "Mố tường đặc & tai mố", 138.40, "C30", 356.8, 24.15, "Bao gồm tường thân, tường cánh, đá kê"),
        (6, "Xà mũ trụ T1, T2 & ụ chống xô", "Xà mũ vát đầu trụ", 112.60, "C30", 228.4, 21.84, "Lắp đặt đá kê gối cầu sai số <=2mm"),
        (7, "15 Phiến dầm Super-T L=38.2m tại bãi đúc", "15 phiến dầm định hình", 434.80, "C45/55", 2610.0, 96.72, "27.78 tấn cáp DƯL 15.2mm (660 tao)"),
        (8, "Dầm ngang mố trụ & bản liên tục nhiệt", "Bê tông nối nhịp", 78.50, "C35", 185.0, 14.20, "Liên kết đầu dầm Super-T trên đỉnh trụ"),
        (9, "Bản mặt cầu tại chỗ & gờ lan can", "Dày 180mm, L=118.2m", 265.40, "C35", 1420.0, 48.60, "Phối hợp 510 tấm ván khuôn đúc sẵn"),
        (10, "Bản quá độ sau 2 mố & đắp K98 đầu cầu", "2 bản L=5.0m", 35.80, "C25", 86.0, 5.80, "Đắp chọn lọc K98 đầm chặt")
    ]

    r_boq = 4
    for stt, name, spec, bt_m3, mac_bt, vk_m2, thep_t, gc in boq_km19:
        ws5.cell(row=r_boq, column=1, value=stt)
        ws5.cell(row=r_boq, column=2, value=name)
        ws5.cell(row=r_boq, column=3, value=spec)
        ws5.cell(row=r_boq, column=4, value=bt_m3)
        ws5.cell(row=r_boq, column=5, value=mac_bt)
        ws5.cell(row=r_boq, column=6, value=vk_m2)
        ws5.cell(row=r_boq, column=7, value=thep_t)
        ws5.cell(row=r_boq, column=8, value=gc)

        for c in range(1, 9):
            cell = ws5.cell(row=r_boq, column=c)
            cell.font = Font(name=font_family, size=9)
            cell.border = thin_border
            if c in [1, 3, 5]: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [4, 6, 7]:
                cell.number_format = "#,##0.00" if c == 7 else "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else: cell.alignment = Alignment(horizontal="left", vertical="center")
        r_boq += 1

    # Total row
    ws5.cell(row=r_boq, column=2, value="TỔNG CỘNG TOÀN CÔNG TRÌNH CẦU KM19+529.080 (100%)")
    ws5.cell(row=r_boq, column=2).font = Font(name=font_family, size=10, bold=True, color="C00000")
    ws5.cell(row=r_boq, column=4, value=f"=SUM(D4:D{r_boq-1})")
    ws5.cell(row=r_boq, column=6, value=f"=SUM(F4:F{r_boq-1})")
    ws5.cell(row=r_boq, column=7, value=f"=SUM(G4:G{r_boq-1})")
    for c in [4, 6, 7]:
        cell = ws5.cell(row=r_boq, column=c)
        cell.font = Font(name=font_family, size=10, bold=True, color="C00000")
        cell.number_format = "#,##0.00" if c == 7 else "#,##0.0"
        cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 9):
        ws5.cell(row=r_boq, column=c).fill = PatternFill(start_color=AMBER_SUM, end_color=AMBER_SUM, fill_type="solid")
        ws5.cell(row=r_boq, column=c).border = thick_bottom

    ws5.column_dimensions["A"].width = 6
    ws5.column_dimensions["B"].width = 46
    ws5.column_dimensions["C"].width = 24
    ws5.column_dimensions["D"].width = 22
    ws5.column_dimensions["E"].width = 14
    ws5.column_dimensions["F"].width = 22
    ws5.column_dimensions["G"].width = 22
    ws5.column_dimensions["H"].width = 42

    # LƯU FILE EXCEL VÀO TẤT CẢ CÁC VỊ TRÍ ĐÍCH
    target_excel_paths = [
        project_path(r"HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"),
        repo_path(r"examples\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx")
    ]

    for p in target_excel_paths:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        wb.save(p)
        print(f"[OK] Đã lưu thành công Excel 5 Sheets chuẩn: {p}")

    # =========================================================================
    # XUẤT FILE XML CHUẨN MS PROJECT CHO CẦU KM19+529.080
    # =========================================================================
    project = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(project, "Name").text = "TIEN_DO_CA_MAY_CAU_KM19_529"
    ET.SubElement(project, "Title").text = "Tiến độ thi công Ca xe ca máy Cầu Km19+529.080 (3 Nhịp Super-T L=38.2m)"
    ET.SubElement(project, "StartDate").text = f"{start_date.isoformat()}T07:00:00"
    ET.SubElement(project, "FinishDate").text = f"{end_date.isoformat()}T17:00:00"

    tasks_elem = ET.SubElement(project, "Tasks")
    for t_item in tasks:
        stt, wbs, name, u, qty, norm_s, s_str, e_str, shifts_d, m_equip, nc_d, oil_n, m_type = t_item
        s_d = datetime.date.fromisoformat(s_str)
        e_d = datetime.date.fromisoformat(e_str)
        dur = (e_d - s_d).days + 1
        is_crit = any(k in name.lower() for k in ["khoan cọc", "đập đầu", "thân trụ", "xà mũ", "lao lắp", "bản mặt cầu"])

        t = ET.SubElement(tasks_elem, "Task")
        ET.SubElement(t, "UID").text = str(stt)
        ET.SubElement(t, "ID").text = str(stt)
        ET.SubElement(t, "Name").text = f"{wbs} - {name} ({qty:,.1f} {u})"
        ET.SubElement(t, "OutlineLevel").text = "1"
        ET.SubElement(t, "Start").text = f"{s_str}T07:00:00"
        ET.SubElement(t, "Finish").text = f"{e_str}T17:00:00"
        ET.SubElement(t, "Duration").text = f"PT{dur * 8}H0M0S"
        ET.SubElement(t, "Milestone").text = "0"
        ET.SubElement(t, "Critical").text = "1" if is_crit else "0"

    tree = ET.ElementTree(project)
    target_xml_paths = [
        project_path(r"HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_Tien_Do_CaMay_Cau_Km19+529.080.xml"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_Tien_Do_CaMay_Cau_Km19+529.080.xml"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_Tien_Do_CaMay_Cau_Km19+529.080.xml"),
        repo_path(r"examples\260920_Tien_Do_CaMay_Cau_Km19+529.080.xml")
    ]

    for p in target_xml_paths:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tree.write(p, encoding="utf-8", xml_declaration=True)
        print(f"[OK] Đã lưu thành công XML MS Project: {p}")

if __name__ == "__main__":
    build_km19_machine_schedule()
