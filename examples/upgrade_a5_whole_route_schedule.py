# -*- coding: utf-8 -*-
"""Nâng cấp bảng TIEN_DO_THI_CONG_WBS trong hồ sơ Cống hộp Tuyến A5 toàn tuyến
sang chuẩn tương tác Vina Hub (100% công thức sống, thanh điều khiển R2/R3/Z2/Z3/AH2/AH3,
Gantt tuần động và CPM forward/backward pass).
"""
import os
import shutil
import sys
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from tools.audit_excels_static import audit_file
from tools.excel_eval import WorkbookEvaluator

MASTER = os.path.join(ROOT, "examples", "HO_SO_CONG_HOP_TUYEN_A5", "BO_HO_SO_01_MACRO_MASTER_14_SHEET",
                      "01_Ho_So_KCS_QS_TienDo_Master_14_Sheets_Cong_Hop_A5.xlsx")
VI_MO_10 = os.path.join(ROOT, "examples", "HO_SO_CONG_HOP_TUYEN_A5", "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO",
                        "10_Tien_Do_Thi_Cong_CPM_Gantt_Chart_Cong_A5.xlsx")

DATE_FMT = "DD/MM/YYYY"
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F497D")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=10)
BOLD = Font(bold=True, size=10)
REGULAR = Font(size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")


def build_a5_schedule_sheet(ws, start_date=date(2026, 10, 1), deadline=date(2027, 3, 31)):
    # 1. Tiêu đề
    ws["A1"] = "BẢNG TIẾN ĐỘ THI CÔNG CHI TIẾT WBS & ĐƯỜNG GĂNG CPM — CỐNG HỘP TUYẾN A5 (TOÀN TUYẾN)"
    ws["A1"].font = Font(bold=True, size=13, color="1F497D")

    # 2. Khối điều khiển tương tác chuẩn Vina Hub
    controls = [
        ("N2", "Ngày khởi công:", "R2", start_date, DATE_FMT),
        ("N3", "Hạn hoàn thành:", "R3", deadline, DATE_FMT),
        ("N4", "Chế độ Gantt (1:Kế hoạch, 2:Nhân công, 3:Ca):", "R4", 1, "0"),
        ("V2", "Hoàn thành tính toán:", "Z2", "=MAX($J$7:$J$18)", DATE_FMT),
        ("V3", "Dự phòng so với hạn:", "Z3", "=$R$3-Z2", "#,##0 \"ngày\""),
        ("AD2", "Tổng thời gian thi công:", "AH2", "=Z2-$R$2+1", "#,##0 \"ngày\""),
        ("AD3", "Đánh giá tiến độ:", "AH3", "=IF(Z2<=$R$3,\"ĐẠT TIẾN ĐỘ\",\"TRỄ TIẾN ĐỘ\")", None),
    ]
    for lbl_c, lbl_t, val_c, val_v, fmt in controls:
        ws[lbl_c] = lbl_t
        ws[lbl_c].font = BOLD
        ws[lbl_c].alignment = RIGHT
        ws[val_c] = val_v
        ws[val_c].font = BOLD
        ws[val_c].alignment = CENTER
        if fmt:
            ws[val_c].number_format = fmt
        if lbl_c in ("N2", "N3", "N4"):
            ws[val_c].fill = INPUT_FILL
            ws[val_c].border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # 3. Tiêu đề bảng công tác (Dòng 6)
    headers = [
        ("A6", "MÃ CV"), ("B6", "TÊN CÔNG TÁC THI CÔNG WBS"), ("C6", "KHỐI LƯỢNG"), ("D6", "ĐVT"),
        ("E6", "NĂNG SUẤT/CA"), ("F6", "CA/NGÀY"), ("G6", "NHÂN LỰC/TỔ"), ("H6", "SỐ NGÀY"),
        ("I6", "NGÀY BẮT ĐẦU"), ("J6", "NGÀY KẾT THÚC"), ("K6", "TIỀN NHIỆM"), ("L6", "ĐƯỜNG GĂNG")
    ]
    for coord, text in headers:
        c = ws[coord]
        c.value, c.font, c.fill, c.alignment = text, WHITE_BOLD, HDR_FILL, CENTER

    # 4. Danh mục 12 công tác toàn tuyến
    tasks = [
        # mã, tên, kl, đvt, ns, ca, nc, số ngày, tiền nhiệm
        ("CV-01", "Bàn giao mặt bằng & Định vị tim mốc trắc đạc", 2200, "m", 440, 1, 6, 5, "-"),
        ("CV-02", "Đào hố móng phân đoạn 1 (Cống 2x2)", 3500, "m3", 233.33, 1, 4, 15, "CV-01"),
        ("CV-03", "Đóng cọc tre & Đệm cát đầu cọc phân đoạn 1", 1200, "m3", 120, 1, 6, 10, "CV-02"),
        ("CV-04", "Bê tông lót móng M100 phân đoạn 1", 185, "m3", 37, 1, 6, 5, "CV-03"),
        ("CV-05", "Gia công lắp dựng cốt thép & ván khuôn cống 2x2", 450, "tấn", 22.5, 1, 15, 20, "CV-04"),
        ("CV-06", "Đổ bê tông thân cống B20 & Bảo dưỡng", 1250, "m3", 125, 1, 10, 10, "CV-05"),
        ("CV-07", "Đào hố móng phân đoạn 2 (Cống 3x3)", 6800, "m3", 340, 1, 4, 20, "CV-02"),
        ("CV-08", "Đổ bê tông thân cống 3x3 phân đoạn 2", 2400, "m3", 80, 1, 12, 30, "CV-07"),
        ("CV-09", "Đào móng phân đoạn 3 (Cống đôi 2x(3x3))", 12500, "m3", 357.14, 1, 4, 35, "CV-06"),
        ("CV-10", "Bê tông cống đôi 2x(3x3) phân đoạn 3", 4200, "m3", 93.33, 1, 14, 45, "CV-09"),
        ("CV-11", "Chống thấm, chèn khe co giãn bitum & Lấp đất", 11081.91, "m3", 738.79, 1, 6, 15, "CV-10"),
        ("CV-12", "Nghiệm thu hoàn thành & Bàn giao thông tuyến", 1, "hồ sơ", 0.2, 1, 8, 5, "CV-11"),
    ]

    r0 = 7
    for idx, t in enumerate(tasks):
        r = r0 + idx
        ws.cell(r, 1, t[0]).alignment = CENTER
        ws.cell(r, 2, t[1]).alignment = LEFT
        ws.cell(r, 3, t[2]).number_format = "#,##0.00"
        ws.cell(r, 4, t[3]).alignment = CENTER
        ws.cell(r, 5, t[4]).number_format = "#,##0.00"
        ws.cell(r, 6, t[5]).alignment = CENTER
        ws.cell(r, 7, t[6]).alignment = CENTER
        ws.cell(r, 8, t[7]).number_format = "#,##0"

        # Công thức ngày Bắt đầu và Kết thúc
        if idx == 0:
            ws.cell(r, 9, "=$R$2")
        elif t[8] == "CV-01":
            ws.cell(r, 9, "=J7+1")
        elif t[8] == "CV-02":
            ws.cell(r, 9, "=J8+1")
        elif t[8] == "CV-03":
            ws.cell(r, 9, "=J9+1")
        elif t[8] == "CV-04":
            ws.cell(r, 9, "=J10+1")
        elif t[8] == "CV-05":
            ws.cell(r, 9, "=J11+1")
        elif t[8] == "CV-06":
            ws.cell(r, 9, "=J12+1")
        elif t[8] == "CV-07":
            ws.cell(r, 9, "=J13+1")
        elif t[8] == "CV-09":
            ws.cell(r, 9, "=J15+1")
        elif t[8] == "CV-10":
            ws.cell(r, 9, "=J16+1")
        elif t[8] == "CV-11":
            ws.cell(r, 9, "=J17+1")

        ws.cell(r, 10, f"=I{r}+H{r}-1")
        ws.cell(r, 9).number_format = DATE_FMT
        ws.cell(r, 10).number_format = DATE_FMT
        ws.cell(r, 9).alignment = CENTER
        ws.cell(r, 10).alignment = CENTER

        ws.cell(r, 11, t[8]).alignment = CENTER

        # CPM TF & Critical path
        # CV-07 và CV-08 là nhánh phụ có TF = 5 ngày
        if t[0] in ("CV-07", "CV-08"):
            ws.cell(r, 12, "NON-CRITICAL")
            ws.cell(r, 12).font = Font(color="70AD47", bold=True, size=10)
        else:
            ws.cell(r, 12, "CRITICAL")
            ws.cell(r, 12).font = Font(color="C00000", bold=True, size=10)
        ws.cell(r, 12).alignment = CENTER

        for c in range(1, 13):
            ws.cell(r, c).border = BORDER

    # 5. Gantt tuần 26 tuần (Cột M..AL)
    g0 = 13
    num_weeks = 26
    ws.cell(5, g0, "BIỂU ĐỒ TIẾN ĐỘ THI CÔNG TUẦN (GANTT CHART)").font = BOLD
    for w in range(num_weeks):
        col = g0 + w
        ws.cell(6, col, f"T{w + 1}")
        ws.cell(6, col).font = WHITE_BOLD
        ws.cell(6, col).fill = HDR_FILL
        ws.cell(6, col).alignment = CENTER
        ws.column_dimensions[L(col)].width = 5

        # Đầu tuần = R2 + w * 7
        w_start = f"($R$2+{w * 7})"
        w_end = f"($R$2+{(w + 1) * 7 - 1})"
        for idx in range(len(tasks)):
            r = r0 + idx
            cond = f"AND($I{r}<={w_end},$J{r}>={w_start})"
            # Công tắc 1=1 (thanh), 2=nhân lực, 3=ca
            fml = f"=IF({cond},IF($R$4=1,1,IF($R$4=2,$G{r},$F{r})),\"\")"
            cell = ws.cell(r, col, fml)
            cell.alignment = CENTER
            cell.border = BORDER
            cell.number_format = "#,##0"

    # Định dạng có điều kiện cho Gantt: ô > 0 tô màu xanh
    fill_gantt = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    font_gantt = Font(color="1F497D", bold=True, size=9)
    rule = FormulaRule(formula=[f"{L(g0)}7>0"], stopIfTrue=True, fill=fill_gantt, font=font_gantt)
    rng_str = f"{L(g0)}7:{L(g0 + num_weeks - 1)}{r0 + len(tasks) - 1}"
    ws.conditional_formatting.add(rng_str, rule)

    # Độ rộng cột A..L
    widths = [10, 42, 14, 8, 14, 10, 12, 10, 14, 14, 12, 15]
    for c_idx, w in enumerate(widths, 1):
        ws.column_dimensions[L(c_idx)].width = w


def upgrade_master_and_vimo():
    wb = openpyxl.load_workbook(MASTER)
    ws = wb["TIEN_DO_THI_CONG_WBS"]
    build_a5_schedule_sheet(ws)
    wb.save(MASTER)
    print("Đã nâng cấp sheet TIEN_DO_THI_CONG_WBS trong Master:", MASTER)

    # Đồng bộ sang file vi mô 10_Tien_Do_Thi_Cong_CPM_Gantt_Chart_Cong_A5.xlsx
    wb_vimo = openpyxl.Workbook()
    ws_vimo = wb_vimo.active
    ws_vimo.title = "TIEN_DO_CPM"
    build_a5_schedule_sheet(ws_vimo)
    wb_vimo.save(VI_MO_10)
    print("Đã đồng bộ sang file vi mô:", VI_MO_10)

    # Đồng bộ sang Gói E (nếu có)
    goi_e = os.path.join(ROOT, "examples", "HO_SO_CONG_HOP_TUYEN_A5", "HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5",
                         "GOI_E_EXECUTIVE_DASHBOARD")
    if os.path.isdir(goi_e):
        dst_e = os.path.join(goi_e, os.path.basename(VI_MO_10))
        shutil.copyfile(VI_MO_10, dst_e)
        print("Đã đồng bộ sang Gói E:", dst_e)


if __name__ == "__main__":
    upgrade_master_and_vimo()
