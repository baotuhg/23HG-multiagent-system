"""
Cập nhật tiến độ toàn tuyến Cống hộp A5 sang mốc: 05/09/2026 đến 12/11/2026 (69 ngày).
100% CÔNG THỨC SỐNG ĐỘNG (ZERO SỐ CHẾT) trên cả 3 tầng của Sheet 01 và liên kết xuyên suốt Sheet 02-05.
"""
import os
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_FILES = [
    os.path.join(ROOT, "examples/HO_SO_CONG_HOP_TUYEN_A5/HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5/GOI_A_CO_GIOI_VA_DAU_DIEZEL/260920_TDTC_CaXe_CaMay_Cong_Hop_Tuyen_A5.xlsx"),
    os.path.join(ROOT, "examples/HO_SO_CONG_HOP_TUYEN_A5/HUB_VINA_ALPHA_05-09_12-11/HUB/03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH/GOI_A_CO_GIOI_VA_DAU_DIEZEL/TDTC_CaXe_CaMay_DauDiezel_Cong_Hop_A5.xlsx")
]

D_START = datetime.date(2026, 9, 5)
D_FINISH = datetime.date(2026, 11, 12)
TOTAL_DAYS = (D_FINISH - D_START).days + 1  # 69 ngày
DATES = [D_START + datetime.timedelta(days=i) for i in range(TOTAL_DAYS)]

# Colors & Fills
FILL_NAVY = PatternFill("solid", fgColor="001B365D")
FILL_BLUE = PatternFill("solid", fgColor="002E75B6")
FILL_GREEN = PatternFill("solid", fgColor="00385723")
FILL_YELLOW = PatternFill("solid", fgColor="00FFF2CC")
FILL_RED_HDR = PatternFill("solid", fgColor="00C00000")
FILL_CORAL_HDR = PatternFill("solid", fgColor="00FCE4D6")
FILL_GANTT_BLUE = PatternFill("solid", fgColor="00BDD7EE")
FILL_GANTT_ORANGE = PatternFill("solid", fgColor="00FCE4D6")
FILL_MACH_GREEN = PatternFill("solid", fgColor="00E2EFDA")
FILL_GREY_LIGHT = PatternFill("solid", fgColor="00F2F2F2")

FONT_WHITE_13 = Font(name="Arial", size=13, bold=True, color="00FFFFFF")
FONT_WHITE_11 = Font(name="Arial", size=11, bold=True, color="00FFFFFF")
FONT_WHITE_85 = Font(name="Arial", size=8.5, bold=True, color="00FFFFFF")
FONT_WHITE_8 = Font(name="Arial", size=8, bold=True, color="00FFFFFF")
FONT_RED_TITLE = Font(name="Arial", size=10, bold=True, color="00C00000")
FONT_RED_85 = Font(name="Arial", size=8.5, bold=True, color="00C00000")
FONT_RED_8 = Font(name="Arial", size=8, bold=True, color="00C00000")
FONT_NAVY_BOLD = Font(name="Arial", size=8.5, bold=True, color="00002060")
FONT_MACH_VAL = Font(name="Arial", size=8, bold=True, color="001B365D")
FONT_DIESEL_VAL = Font(name="Arial", size=7.5, bold=False, color="00333333")
FONT_GANTT_BLUE = Font(name="Arial", size=8, bold=False, color="001B365D")
FONT_REGULAR_85 = Font(name="Arial", size=8.5, bold=False)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

WEEKDAYS_VN = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

# 17 Công tác WBS
TASKS = [
    {
        "row": 8, "wbs": "ĐM-16", "name": "Đào đất hố móng cống hộp 3 mũi (H_tb = 3.5m)",
        "unit": "m3", "qty": 73800.0, "norm": 401.07, "shifts": 2,
        "start": datetime.date(2026, 9, 5), "finish": datetime.date(2026, 10, 4),
        "mach": "Máy xúc PC200 - PC300", "crew": 24, "critical": False
    },
    {
        "row": 9, "wbs": "ĐM-15", "name": "Đào hố móng 63 hố ga thu nước & đấu nối",
        "unit": "m3", "qty": 12650.0, "norm": 313.55, "shifts": 2,
        "start": datetime.date(2026, 9, 8), "finish": datetime.date(2026, 9, 30),
        "mach": "Máy xúc PC200", "crew": 12, "critical": False
    },
    {
        "row": 10, "wbs": "ĐM-29", "name": "Vận chuyển đất đào cự ly <5km (bãi thải/đắp)",
        "unit": "m3", "qty": 86450.0, "norm": 571.43, "shifts": 2,
        "start": datetime.date(2026, 9, 5), "finish": datetime.date(2026, 10, 4),
        "mach": "Ô tô tự đổ Howo 18 m3", "crew": 10, "critical": False
    },
    {
        "row": 11, "wbs": "ĐM-BOM", "name": "Bơm hạ mực nước ngầm & chống ngập móng",
        "unit": "ca", "qty": 138.0, "norm": 1.0, "shifts": 2,
        "start": datetime.date(2026, 9, 5), "finish": datetime.date(2026, 11, 12),
        "mach": "Máy bơm nước d80-d100", "crew": 4, "critical": False
    },
    {
        "row": 12, "wbs": "ĐM-53", "name": "Đệm cát / đá dăm 4x6 lót móng cống",
        "unit": "m3", "qty": 1250.0, "norm": 150.0, "shifts": 2,
        "start": datetime.date(2026, 9, 10), "finish": datetime.date(2026, 10, 8),
        "mach": "Đầm cóc + đầm bàn", "crew": 12, "critical": False
    },
    {
        "row": 13, "wbs": "ĐM-37", "name": "Bê tông lót móng M100# dày 50mm",
        "unit": "m3", "qty": 592.4, "norm": 45.0, "shifts": 2,
        "start": datetime.date(2026, 9, 12), "finish": datetime.date(2026, 10, 12),
        "mach": "Xe bồn + đầm bàn", "crew": 16, "critical": False
    },
    {
        "row": 14, "wbs": "ĐM-CT", "name": "Gia công & lắp dựng cốt thép cống hộp B20",
        "unit": "tấn", "qty": 753.36, "norm": 12.0, "shifts": 2,
        "start": datetime.date(2026, 9, 16), "finish": datetime.date(2026, 10, 26),
        "mach": "Cần cẩu 25T & Máy uốn cắt", "crew": 60, "critical": True
    },
    {
        "row": 15, "wbs": "ĐM-VK", "name": "Lắp dựng ván khuôn thép/phủ phim thân cống",
        "unit": "m2", "qty": 44940.7, "norm": 450.0, "shifts": 2,
        "start": datetime.date(2026, 9, 18), "finish": datetime.date(2026, 10, 28),
        "mach": "Cần cẩu 15T & Dàn giáo", "crew": 70, "critical": True
    },
    {
        "row": 16, "wbs": "ĐM-BT", "name": "Đổ bê tông thân cống B20 (đáy, vách, nắp)",
        "unit": "m3", "qty": 11870.0, "norm": 150.0, "shifts": 2,
        "start": datetime.date(2026, 9, 20), "finish": datetime.date(2026, 10, 31),
        "mach": "Máy bơm cần 37-43m + đầm", "crew": 40, "critical": True
    },
    {
        "row": 17, "wbs": "ĐM-VCBT", "name": "Vận chuyển bê tông thương phẩm B20 trạm trộn",
        "unit": "m3", "qty": 11870.0, "norm": 45.0, "shifts": 2,
        "start": datetime.date(2026, 9, 20), "finish": datetime.date(2026, 10, 31),
        "mach": "Xe bồn 8-10m3 (6 xe)", "crew": 12, "critical": True
    },
    {
        "row": 18, "wbs": "ĐM-46", "name": "Bê tông & cốt thép 63 hố ga BTCT",
        "unit": "m3", "qty": 409.5, "norm": 25.0, "shifts": 2,
        "start": datetime.date(2026, 9, 25), "finish": datetime.date(2026, 10, 25),
        "mach": "Xe bồn + đầm bàn", "crew": 22, "critical": False
    },
    {
        "row": 19, "wbs": "ĐM-25", "name": "Máy xúc lốp PC140 cẩu lắp & đầm cóc mang cống",
        "unit": "md", "qty": 2170.0, "norm": 100.0, "shifts": 2,
        "start": datetime.date(2026, 9, 20), "finish": datetime.date(2026, 10, 31),
        "mach": "Máy xúc lốp PC140", "crew": 8, "critical": False
    },
    {
        "row": 20, "wbs": "ĐM-17", "name": "Đắp cát/đất K95 hoàn trả mang cống",
        "unit": "m3", "qty": 54200.0, "norm": 255.0, "shifts": 2,
        "start": datetime.date(2026, 10, 20), "finish": datetime.date(2026, 11, 10),
        "mach": "Đầm cóc + xúc lật mang cống", "crew": 28, "critical": False
    },
    {
        "row": 21, "wbs": "ĐM-ỦI", "name": "Máy ủi D3-D5 san gạt hoàn trả đỉnh móng",
        "unit": "m3", "qty": 54200.0, "norm": 1425.0, "shifts": 2,
        "start": datetime.date(2026, 10, 20), "finish": datetime.date(2026, 11, 10),
        "mach": "Máy ủi bánh xích D3-D5", "crew": 4, "critical": False
    },
    {
        "row": 22, "wbs": "ĐM-LU", "name": "Máy lu rung 12-16T đầm nén K95 hoàn trả",
        "unit": "m3", "qty": 54200.0, "norm": 255.0, "shifts": 2,
        "start": datetime.date(2026, 10, 21), "finish": datetime.date(2026, 11, 11),
        "mach": "Máy lu rung 12T - 16T", "crew": 12, "critical": False
    },
    {
        "row": 23, "wbs": "ĐM-DAU", "name": "Xe téc cấp dầu lưu động 9m3 phục vụ máy móc",
        "unit": "ca", "qty": 138.0, "norm": 1.0, "shifts": 2,
        "start": datetime.date(2026, 9, 5), "finish": datetime.date(2026, 11, 12),
        "mach": "Xe téc cấp dầu 9 m3", "crew": 2, "critical": False
    },
    {
        "row": 24, "wbs": "ĐM-ĐIỆN", "name": "Máy phát điện 3 pha công nghiệp 25-45kVA",
        "unit": "ca", "qty": 138.0, "norm": 1.0, "shifts": 2,
        "start": datetime.date(2026, 9, 5), "finish": datetime.date(2026, 11, 12),
        "mach": "Máy phát điện 25-45kVA", "crew": 2, "critical": False
    }
]

# 11 Chủng loại MMTB
MACHINES = [
    {"row_m": 30, "row_f": 44, "code": "M1", "name": "Máy xúc bánh xích PC200 - PC300", "unit": "cái", "fuel_l": 89.6, "max_m": 8, "desc": "Đào móng, đào ga, xúc mang cống",
     "fml_active": lambda c: f"=IF(AND($J$8<={c}$6,$K$8>={c}$6),8,IF(AND($J$20<={c}$6,$K$20>={c}$6),4,0))"},
    {"row_m": 31, "row_f": 45, "code": "M7", "name": "Ô tô tự đổ 18 m3", "unit": "xe", "fuel_l": 26.0, "max_m": 5, "desc": "Vận chuyển đất đào cự ly <5km",
     "fml_active": lambda c: f"=IF(AND($J$10<={c}$6,$K$10>={c}$6),5,IF(AND($J$20<={c}$6,$K$20>={c}$6),4,0))"},
    {"row_m": 32, "row_f": 46, "code": "M5", "name": "Máy lu rung 12T - 16T", "unit": "cái", "fuel_l": 44.0, "max_m": 6, "desc": "Đầm nén K95 mang cống và đỉnh cống",
     "fml_active": lambda c: f"=IF(AND($J$22<={c}$6,$K$22>={c}$6),6,0)"},
    {"row_m": 33, "row_f": 47, "code": "M3", "name": "Máy ủi bánh xích D3 - D5", "unit": "cái", "fuel_l": 124.0, "max_m": 2, "desc": "San gạt hoàn trả hố móng cống",
     "fml_active": lambda c: f"=IF(AND($J$21<={c}$6,$K$21>={c}$6),2,0)"},
    {"row_m": 34, "row_f": 48, "code": "MC", "name": "Cần cẩu 25T & Cẩu tự hành 15T", "unit": "cái", "fuel_l": 28.0, "max_m": 3, "desc": "Hạ ván khuôn, cẩu lắp cốt thép",
     "fml_active": lambda c: f"=IF(OR(AND($J$14<={c}$6,$K$14>={c}$6),AND($J$15<={c}$6,$K$15>={c}$6)),3,0)"},
    {"row_m": 35, "row_f": 49, "code": "MB", "name": "Máy bơm bê tông cần 37 - 43m", "unit": "cái", "fuel_l": 65.0, "max_m": 2, "desc": "Đổ bê tông thân cống B20 & hố ga",
     "fml_active": lambda c: f"=IF(AND($J$16<={c}$6,$K$16>={c}$6),2,0)"},
    {"row_m": 36, "row_f": 50, "code": "XB", "name": "Xe bồn vận chuyển bê tông 8-10m3", "unit": "xe", "fuel_l": 42.0, "max_m": 6, "desc": "Chở bê tông từ trạm trộn về công trường",
     "fml_active": lambda c: f"=IF(AND($J$17<={c}$6,$K$17>={c}$6),6,0)"},
    {"row_m": 37, "row_f": 51, "code": "M8", "name": "Máy xúc bánh lốp PC140", "unit": "cái", "fuel_l": 64.0, "max_m": 1, "desc": "Cẩu lắp, đầm cóc mang cống",
     "fml_active": lambda c: f"=IF(AND($J$19<={c}$6,$K$19>={c}$6),1,0)"},
    {"row_m": 38, "row_f": 52, "code": "BP", "name": "Máy bơm nước hố móng d80-d100", "unit": "cái", "fuel_l": 14.0, "max_m": 2, "desc": "Bơm hạ mực nước ngầm 24/7",
     "fml_active": lambda c: f"=IF(AND($J$11<={c}$6,$K$11>={c}$6),2,0)"},
    {"row_m": 39, "row_f": 53, "code": "MP1", "name": "Xe téc cấp dầu lưu động 9m3", "unit": "xe", "fuel_l": 44.0, "max_m": 1, "desc": "Cấp dầu lưu động 2 ca/ngày",
     "fml_active": lambda c: f"=IF(AND($J$23<={c}$6,$K$23>={c}$6),1,0)"},
    {"row_m": 40, "row_f": 54, "code": "MP2", "name": "Máy phát điện 3 pha 25-45kVA", "unit": "cái", "fuel_l": 28.0, "max_m": 2, "desc": "Chiếu sáng ban đêm 2 ca",
     "fml_active": lambda c: f"=IF(AND($J$24<={c}$6,$K$24>={c}$6),2,0)"}
]


def update_workbook_live_formulas(wb_path: str):
    print(f"[*] Đang nạp 100% CÔNG THỨC SỐNG cho: {wb_path}")
    wb = openpyxl.load_workbook(wb_path, data_only=False)
    ws1 = wb["01_TienDo_CaMay_Master"]

    # 1. Tiêu đề Row 3
    ws1["A3"] = f"MỐC TIẾN ĐỘ THI CÔNG TOÀN TUYẾN: TỪ 05/09/2026 ĐẾN 12/11/2026 (69 NGÀY) - 2 CA/NGÀY (20H/NGÀY) - 3 MŨI THI CÔNG ĐỒNG THỜI"
    ws1["A3"].font = FONT_RED_TITLE
    ws1["A3"].fill = FILL_YELLOW

    # 2. Xóa các cột timeline cũ
    max_c = ws1.max_column
    for r in range(6, 56):
        for c in range(16, max_c + 1):
            cell = ws1.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()

    # 3. Tạo Headers Timeline 69 ngày với CÔNG THỨC SỐNG LIÊN TỤC
    # Col 16 (P) là mốc gốc, các cột tiếp theo liên kết công thức sống: =P6+1, =Q6+1...
    for i, dt in enumerate(DATES):
        col = 16 + i
        col_letter = get_column_letter(col)
        ws1.column_dimensions[col_letter].width = 6.8
        is_sun = (dt.weekday() == 6)

        # Row 6: Ngày thực tế dạng Date Serial sống
        if i == 0:
            c6 = ws1.cell(6, col, dt)
        else:
            prev_letter = get_column_letter(col - 1)
            c6 = ws1.cell(6, col, f"={prev_letter}6+1")

        c6.number_format = "DD/MM"
        c6.font = FONT_WHITE_8
        c6.fill = FILL_RED_HDR if is_sun else FILL_NAVY
        c6.alignment = Alignment(horizontal="center", vertical="center")
        c6.border = THIN_BORDER

        # Row 7: Nhãn Thứ
        c7 = ws1.cell(7, col, WEEKDAYS_VN[dt.weekday()])
        c7.font = FONT_RED_8 if is_sun else Font(name="Arial", size=8, bold=True)
        c7.fill = FILL_CORAL_HDR if is_sun else FILL_GREY_LIGHT
        c7.alignment = Alignment(horizontal="center", vertical="center")
        c7.border = THIN_BORDER

        # Row 29: Header cho Bảng 2
        c29 = ws1.cell(29, col, f"={col_letter}6")
        c29.number_format = "DD/MM"
        c29.font = FONT_WHITE_8
        c29.fill = FILL_BLUE
        c29.alignment = Alignment(horizontal="center", vertical="center")
        c29.border = THIN_BORDER

    # 4. Ghi 17 Task với 100% CÔNG THỨC SỐNG TOÀN DIỆN
    for t in TASKS:
        r = t["row"]
        ws1.cell(r, 1, r - 7).alignment = Alignment(horizontal="center")
        ws1.cell(r, 2, t["wbs"]).alignment = Alignment(horizontal="center")
        ws1.cell(r, 3, t["name"])
        ws1.cell(r, 4, t["unit"]).alignment = Alignment(horizontal="center")
        ws1.cell(r, 5, t["qty"]).number_format = "#,##0.00" if isinstance(t["qty"], float) else "#,##0"
        ws1.cell(r, 6, t["norm"]).number_format = "#,##0.00"

        # Cột 7: Tổng số ca máy = FORMULA SỐNG
        if t["unit"] == "ca":
            ws1.cell(r, 7, f"=E{r}").number_format = "#,##0.0"
        else:
            ws1.cell(r, 7, f"=IF(F{r}>0,ROUND(E{r}/F{r},1),0)").number_format = "#,##0.0"

        # Cột 10, 11: Ngày BĐ, Ngày KT
        c10 = ws1.cell(r, 10, t["start"])
        c10.number_format = "DD/MM/YYYY"
        c10.alignment = Alignment(horizontal="center")
        c11 = ws1.cell(r, 11, t["finish"])
        c11.number_format = "DD/MM/YYYY"
        c11.alignment = Alignment(horizontal="center")

        # Cột 9: Thời gian (ngày) = FORMULA SỐNG: =K8-J8+1
        ws1.cell(r, 9, f"=K{r}-J{r}+1").number_format = "#,##0"
        ws1.cell(r, 9).alignment = Alignment(horizontal="center")

        # Cột 8: Năng xuất ngày = FORMULA SỐNG: =IF(I8>0,ROUND(E8/I8,1),0)
        ws1.cell(r, 8, f"=IF(I{r}>0,ROUND(E{r}/I{r},1),0)").number_format = "#,##0.0"

        # Cột 12: Số ca/ngày
        ws1.cell(r, 12, t["shifts"]).alignment = Alignment(horizontal="center")

        # Cột 13: Số máy huy động/ngày = FORMULA SỐNG: =IF(AND(G8>0,I8>0,L8>0),ROUNDUP(G8/(I8*L8),2),0)
        ws1.cell(r, 13, f"=IF(AND(G{r}>0,I{r}>0,L{r}>0),ROUNDUP(G{r}/(I{r}*L{r}),2),0)").number_format = "0.00"
        ws1.cell(r, 13).alignment = Alignment(horizontal="center")

        ws1.cell(r, 14, t["mach"])
        ws1.cell(r, 15, t["crew"]).alignment = Alignment(horizontal="center")

        # Tô Gantt cho task này = FORMULA SỐNG: =IF(AND($J8<=col$6,$K8>=col$6),$M8,"")
        for i, dt in enumerate(DATES):
            col = 16 + i
            col_letter = get_column_letter(col)
            cell = ws1.cell(r, col)
            cell.value = f'=IF(AND($J{r}<={col_letter}$6,$K{r}>={col_letter}$6),$M{r},"")'
            cell.number_format = "0.00"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

            # Giữ màu sắc thẩm mỹ
            if t["start"] <= dt <= t["finish"]:
                cell.font = FONT_GANTT_BLUE
                cell.fill = FILL_GANTT_ORANGE if t["critical"] else FILL_GANTT_BLUE
            else:
                cell.fill = PatternFill(fill_type=None)

    # 5. Row 26: Tổng nhân công = FORMULA SỐNG SUMPRODUCT TOÀN TUYẾN
    ws1.cell(26, 3, "TỔNG NHÂN CÔNG TRÊN CÔNG TRƯỜNG TUYẾN A5 (Người/ngày)")
    ws1.cell(26, 3).font = Font(name="Arial", size=10, bold=True, color="00FFFFFF")
    ws1.cell(26, 3).fill = FILL_NAVY
    for col_idx in range(1, 16):
        ws1.cell(26, col_idx).fill = FILL_NAVY

    for i, dt in enumerate(DATES):
        col = 16 + i
        col_letter = get_column_letter(col)
        # =SUMPRODUCT(($J$8:$J$24<=col$6)*($K$8:$K$24>=col$6)*$O$8:$O$24)
        cell = ws1.cell(26, col, f'=SUMPRODUCT(($J$8:$J$24<={col_letter}$6)*($K$8:$K$24>={col_letter}$6)*$O$8:$O$24)')
        cell.font = FONT_NAVY_BOLD
        cell.fill = FILL_YELLOW
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0"
        cell.border = THIN_BORDER

    # 6. Bảng 2: Tổng hợp ca máy & Phương tiện huy động theo ngày (R28..R40) = FORMULA SỐNG
    for m in MACHINES:
        r_m = m["row_m"]
        ws1.cell(r_m, 1, m["code"]).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 3, m["name"])
        ws1.cell(r_m, 4, m["unit"]).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 5, m["fuel_l"]).number_format = "#,##0.0"
        ws1.cell(r_m, 6, m["max_m"]).alignment = Alignment(horizontal="center")
        ws1.cell(r_m, 7, m["desc"])

        for i, dt in enumerate(DATES):
            col = 16 + i
            col_letter = get_column_letter(col)
            cell = ws1.cell(r_m, col)
            cell.border = THIN_BORDER
            cell.value = m["fml_active"](col_letter)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "0"
            cell.font = FONT_MACH_VAL

            # Giữ highlight xanh lá nhạt cho các ô đang hoạt động
            active_dt = False
            code = m["code"]
            if code == "M1" and (datetime.date(2026, 9, 5) <= dt <= datetime.date(2026, 10, 4) or datetime.date(2026, 10, 20) <= dt <= datetime.date(2026, 11, 10)): active_dt = True
            elif code == "M7" and (datetime.date(2026, 9, 5) <= dt <= datetime.date(2026, 10, 4) or datetime.date(2026, 10, 20) <= dt <= datetime.date(2026, 11, 10)): active_dt = True
            elif code == "M5" and (datetime.date(2026, 10, 21) <= dt <= datetime.date(2026, 11, 11)): active_dt = True
            elif code == "M3" and (datetime.date(2026, 10, 20) <= dt <= datetime.date(2026, 11, 10)): active_dt = True
            elif code == "MC" and (datetime.date(2026, 9, 16) <= dt <= datetime.date(2026, 10, 28)): active_dt = True
            elif code in ("MB", "XB", "M8") and (datetime.date(2026, 9, 20) <= dt <= datetime.date(2026, 10, 31)): active_dt = True
            elif code in ("BP", "MP1", "MP2"): active_dt = True

            if active_dt:
                cell.fill = FILL_MACH_GREEN
            else:
                cell.fill = PatternFill(fill_type=None)

    # 7. Bảng 3: Dầu Diezel tiêu thụ theo ngày (R42..R54) = 100% FORMULA SỐNG
    ws1.cell(42, 3, f"BẢNG TÍNH DẦU DIEZEL TIÊU THỤ THEO TIẾN ĐỘ THI CÔNG TUYẾN A5 (Lít/ngày)")
    ws1.cell(43, 3, "TỔNG SỐ LÍT DẦU DIEZEL TIÊU THỤ / NGÀY")
    ws1.cell(43, 3).font = FONT_RED_TITLE
    ws1.cell(43, 3).fill = FILL_YELLOW
    for col_idx in range(1, 16):
        ws1.cell(43, col_idx).fill = FILL_YELLOW

    for m in MACHINES:
        r_f = m["row_f"]
        ws1.cell(r_f, 1, m["code"]).alignment = Alignment(horizontal="center")
        ws1.cell(r_f, 3, f"Nhiên liệu dầu Diezel cho máy {m['code']}")
        ws1.cell(r_f, 4, "Lít").alignment = Alignment(horizontal="center")
        ws1.cell(r_f, 5, m["fuel_l"]).number_format = "#,##0.0"

    for i, dt in enumerate(DATES):
        col = 16 + i
        col_letter = get_column_letter(col)

        # Tổng lít dầu / ngày = SUM(P44:P54)
        c43 = ws1.cell(43, col, f"=SUM({col_letter}44:{col_letter}54)")
        c43.font = FONT_RED_85
        c43.fill = FILL_YELLOW
        c43.alignment = Alignment(horizontal="center", vertical="center")
        c43.number_format = "#,##0"
        c43.border = THIN_BORDER

        # Từng dòng máy: ={col}row_m*$Erow_f*2
        for m in MACHINES:
            r_m = m["row_m"]
            r_f = m["row_f"]
            cf = ws1.cell(r_f, col, f"={col_letter}{r_m}*$E{r_f}*2")
            cf.font = FONT_DIESEL_VAL
            cf.alignment = Alignment(horizontal="center", vertical="center")
            cf.number_format = "#,##0"
            cf.border = THIN_BORDER

    # 8. Cập nhật Sheet 02 (Tổng hợp Ca máy): LIÊN KẾT CÔNG THỨC SỐNG TỪ SHEET 01
    if "02_TongHop_CaXe_CaMay_MMTB" in wb.sheetnames:
        ws2 = wb["02_TongHop_CaXe_CaMay_MMTB"]
        # Ca máy liên kết trực tiếp các hạng mục trên Sheet 01
        ws2["G4"] = "='01_TienDo_CaMay_Master'!G8+'01_TienDo_CaMay_Master'!G9+'01_TienDo_CaMay_Master'!G20"  # M1
        ws2["G5"] = "='01_TienDo_CaMay_Master'!G10"  # M7
        ws2["G6"] = "='01_TienDo_CaMay_Master'!G22"  # M5
        ws2["G7"] = "='01_TienDo_CaMay_Master'!G21"  # M3
        ws2["G8"] = "='01_TienDo_CaMay_Master'!G14+'01_TienDo_CaMay_Master'!G15"  # MC
        ws2["G9"] = "='01_TienDo_CaMay_Master'!G16"  # MB
        ws2["G10"] = "='01_TienDo_CaMay_Master'!G17"  # XB
        ws2["G11"] = "='01_TienDo_CaMay_Master'!G19"  # M8
        ws2["G12"] = "='01_TienDo_CaMay_Master'!G11"  # BP
        ws2["G13"] = "='01_TienDo_CaMay_Master'!G23"  # MP1
        ws2["G14"] = "='01_TienDo_CaMay_Master'!G24"  # MP2

        for r_row in range(4, 15):
            ws2[f"G{r_row}"].number_format = "#,##0.0"
            ws2[f"I{r_row}"] = f"=ROUND(F{r_row}*G{r_row},1)"
            ws2[f"I{r_row}"].number_format = "#,##0.0"

        ws2["G15"] = "=SUM(G4:G14)"
        ws2["G15"].number_format = "#,##0.0"
        ws2["I15"] = "=SUM(I4:I14)"
        ws2["I15"].number_format = "#,##0.0"

    # 9. Cập nhật Sheet 03 (Kế hoạch dầu diezel): LIÊN KẾT CÔNG THỨC SỐNG TỪ SHEET 02
    if "03_KeHoach_Dau_Diezel" in wb.sheetnames:
        ws3 = wb["03_KeHoach_Dau_Diezel"]
        ws3["A1"] = "KẾ HOẠCH CẤP DẦU DIEZEL CHO MÁY MÓC THI CÔNG TUYẾN A5 (TỪ 05/09/2026 ĐẾN 12/11/2026)"
        ws3["F3"] = "Kỳ 1 (05/9 - 20/9)"
        ws3["G3"] = "Kỳ 2 (21/9 - 10/10)"
        ws3["H3"] = "Kỳ 3 (11/10 - 31/10)"
        ws3["I3"] = "Kỳ 4 (01/11 - 12/11)"
        for r_row in range(4, 15):
            ws3[f"D{r_row}"] = f"='02_TongHop_CaXe_CaMay_MMTB'!G{r_row}"
            ws3[f"D{r_row}"].number_format = "#,##0.0"
            ws3[f"E{r_row}"] = f"=ROUND(C{r_row}*D{r_row},1)"
            ws3[f"E{r_row}"].number_format = "#,##0.0"

        ws3["D15"] = "=SUM(D4:D14)"
        ws3["E15"] = "=SUM(E4:E14)"

    wb.save(wb_path)
    wb.close()
    print(f"[OK] Đã nạp thành công 100% CÔNG THỨC SỐNG cho: {wb_path}")


def main():
    for f in SRC_FILES:
        if os.path.exists(f):
            update_workbook_live_formulas(f)
        else:
            print(f"[!] Không tìm thấy tệp: {f}")


if __name__ == "__main__":
    main()
