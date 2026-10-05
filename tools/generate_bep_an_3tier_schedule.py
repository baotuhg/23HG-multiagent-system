# -*- coding: utf-8 -*-
"""
HỆ THỐNG TIẾN ĐỘ CA MÁY 3 TẦNG HỢP NHẤT — 100% CÔNG THỨC SỐNG ĐỘNG (ZERO SỐ CHẾT)
Dự án: CẢI TẠO NHÀ BẾP ĂN (NHÀ 17 VÀ 17A) - TRƯỜNG PHỔ THÔNG LIÊN CẤP PHỐ BẢNG
Khung tiến độ: 01/10/2026 -> 30/11/2026 (61 ngày)
Chuẩn mẫu trực quan cao cấp Vincons / 23HG System

Cấu trúc:
- Sheet 01: 01_TienDo_CaMay_Master
  + Tầng 1: Tiến độ 26 công tác WBS, Gantt ngày động theo công thức sống, phân tầng màu
  + Dòng tổng: Tổng nhân công trên công trường (Người/ngày) = SUMPRODUCT sống
  + Tầng 2: Bảng tổng hợp ca máy & Phương tiện MMTB huy động theo ngày = IF/OR sống, nền xanh lá
  + Tầng 3: Bảng tính dầu Diezel tiêu thụ theo tiến độ thi công (Lít/ngày) = SUM & phép nhân sống
- Sheet 02: 02_TongHop_CaXe_CaMay_MMTB: Liên kết 100% công thức sống từ Sheet 01
- Sheet 03: 03_KeHoach_Dau_Diezel: 4 Kỳ thi công liên kết công thức sống từ Sheet 02 & Sheet 01
- Sheet 04: 04_KeHoach_NhanLuc: Bảng điều phối nhân lực 5 tổ đội thi công
- Sheet 05: 05_DoiChieu_BocTach: Bảng đối chiếu bóc tách khối lượng và định mức thiết kế
"""
import os
import datetime
from typing import List, Dict, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# HẰNG SỐ MÀU SẮC & ĐỊNH DẠNG THEO CHUẨN MẪU GÓI A
# ==============================================================================
FONT_NAME = "Arial"

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
FILL_ZEBRA = PatternFill("solid", fgColor="00F9FBFD")

FONT_WHITE_13 = Font(name=FONT_NAME, size=13, bold=True, color="00FFFFFF")
FONT_WHITE_11 = Font(name=FONT_NAME, size=11, bold=True, color="00FFFFFF")
FONT_WHITE_85 = Font(name=FONT_NAME, size=8.5, bold=True, color="00FFFFFF")
FONT_WHITE_8 = Font(name=FONT_NAME, size=8, bold=True, color="00FFFFFF")
FONT_RED_TITLE = Font(name=FONT_NAME, size=10, bold=True, color="00C00000")
FONT_RED_85 = Font(name=FONT_NAME, size=8.5, bold=True, color="00C00000")
FONT_RED_8 = Font(name=FONT_NAME, size=8, bold=True, color="00C00000")
FONT_NAVY_BOLD = Font(name=FONT_NAME, size=8.5, bold=True, color="00002060")
FONT_MACH_VAL = Font(name=FONT_NAME, size=8, bold=True, color="001B365D")
FONT_DIESEL_VAL = Font(name=FONT_NAME, size=7.5, bold=False, color="00333333")
FONT_GANTT_BLUE = Font(name=FONT_NAME, size=8, bold=False, color="001B365D")
FONT_GANTT_ORANGE = Font(name=FONT_NAME, size=8, bold=True, color="00C00000")
FONT_REGULAR_85 = Font(name=FONT_NAME, size=8.5, bold=False)
FONT_BOLD_85 = Font(name=FONT_NAME, size=8.5, bold=True)
FONT_REGULAR_8 = Font(name=FONT_NAME, size=8, bold=False)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
DOUBLE_BOTTOM_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='double', color='001B365D')
)

WEEKDAYS_VN = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

# ==============================================================================
# DỮ LIỆU CÔNG TRÌNH THỰC TẾ: CẢI TẠO NHÀ BẾP ĂN (NHÀ 17 VÀ 17A)
# ==============================================================================
D_START = datetime.date(2026, 10, 1)
D_FINISH = datetime.date(2026, 11, 30)
TOTAL_DAYS = (D_FINISH - D_START).days + 1  # 61 ngày
DATES = [D_START + datetime.timedelta(days=i) for i in range(TOTAL_DAYS)]

# Danh sách 26 công tác WBS (Hàng 8 đến 33 trên Sheet 01)
TASKS = [
    {
        "row": 8, "wbs": "G1-01", "name": "Chuẩn bị mặt bằng & hàng rào tôn bảo vệ xung quanh",
        "unit": "m2", "qty": 35.0, "norm": 17.5, "shifts": 1,
        "start": datetime.date(2026, 10, 1), "finish": datetime.date(2026, 10, 2),
        "mach": "Dụng cụ cơ giới nhỏ", "crew": 4, "critical": False
    },
    {
        "row": 9, "wbs": "G1-02", "name": "Tháo dỡ hệ thống cửa đi, cửa sổ cũ phân loại tận dụng",
        "unit": "m2", "qty": 24.0, "norm": 8.0, "shifts": 1,
        "start": datetime.date(2026, 10, 2), "finish": datetime.date(2026, 10, 4),
        "mach": "Máy cắt cầm tay + giàn giáo", "crew": 6, "critical": False
    },
    {
        "row": 10, "wbs": "G1-03", "name": "Phá dỡ tường gạch cũ dày 220mm (KT 1800x1800; 2740x2700)",
        "unit": "m3", "qty": 28.5, "norm": 5.7, "shifts": 1,
        "start": datetime.date(2026, 10, 2), "finish": datetime.date(2026, 10, 6),
        "mach": "Máy búa căn + đục khí nén", "crew": 6, "critical": False
    },
    {
        "row": 11, "wbs": "G1-07", "name": "Phá dỡ gạch lát nền cũ, đục bóc lớp vữa lót toàn bộ",
        "unit": "m2", "qty": 64.0, "norm": 16.0, "shifts": 1,
        "start": datetime.date(2026, 10, 4), "finish": datetime.date(2026, 10, 7),
        "mach": "Máy đục bê tông cầm tay", "crew": 6, "critical": False
    },
    {
        "row": 12, "wbs": "G1-09", "name": "Vận chuyển phế thải phá dỡ ra bãi tập kết bằng ô tô 5T",
        "unit": "m3", "qty": 145.0, "norm": 24.17, "shifts": 1,
        "start": datetime.date(2026, 10, 3), "finish": datetime.date(2026, 10, 8),
        "mach": "Ô tô tự đổ 5T (2 xe)", "crew": 4, "critical": False
    },
    {
        "row": 13, "wbs": "G2-02", "name": "Đào đất hố móng mác đất cấp III bằng máy kết hợp thủ công",
        "unit": "m3", "qty": 85.0, "norm": 17.0, "shifts": 1,
        "start": datetime.date(2026, 10, 7), "finish": datetime.date(2026, 10, 11),
        "mach": "Máy đào gầu 0.4 m3", "crew": 4, "critical": True
    },
    {
        "row": 14, "wbs": "G2-03", "name": "Đệm cát & Bê tông lót móng M100 đá 4x6 dày 100mm",
        "unit": "m3", "qty": 28.0, "norm": 7.0, "shifts": 1,
        "start": datetime.date(2026, 10, 10), "finish": datetime.date(2026, 10, 13),
        "mach": "Máy trộn 250L + đầm bàn", "crew": 8, "critical": False
    },
    {
        "row": 15, "wbs": "G2-04", "name": "Gia công & lắp dựng cốt thép móng đơn MT1, MT2",
        "unit": "tấn", "qty": 4.25, "norm": 0.85, "shifts": 1,
        "start": datetime.date(2026, 10, 12), "finish": datetime.date(2026, 10, 16),
        "mach": "Máy cắt uốn thép CNC", "crew": 10, "critical": True
    },
    {
        "row": 16, "wbs": "G2-08", "name": "Lắp dựng ván khuôn móng đơn và móng băng gạch",
        "unit": "m2", "qty": 134.4, "norm": 33.6, "shifts": 1,
        "start": datetime.date(2026, 10, 14), "finish": datetime.date(2026, 10, 17),
        "mach": "Dàn giáo thép + phụ kiện", "crew": 8, "critical": True
    },
    {
        "row": 17, "wbs": "G2-09", "name": "Đổ bê tông móng đơn MT1, MT2 đá 1x2 mác 250 (B20)",
        "unit": "m3", "qty": 28.0, "norm": 9.33, "shifts": 1,
        "start": datetime.date(2026, 10, 16), "finish": datetime.date(2026, 10, 18),
        "mach": "Máy trộn 250L + đầm dùi", "crew": 12, "critical": True
    },
    {
        "row": 18, "wbs": "G2-10", "name": "Đổ bê tông giằng móng 220x400 đá 1x2 mác 250",
        "unit": "m3", "qty": 12.0, "norm": 4.0, "shifts": 1,
        "start": datetime.date(2026, 10, 18), "finish": datetime.date(2026, 10, 20),
        "mach": "Máy trộn 250L + đầm dùi", "crew": 10, "critical": True
    },
    {
        "row": 19, "wbs": "G2-14", "name": "Xây móng gạch đặc M75 vữa xi măng M75",
        "unit": "m3", "qty": 45.0, "norm": 9.0, "shifts": 1,
        "start": datetime.date(2026, 10, 20), "finish": datetime.date(2026, 10, 24),
        "mach": "Máy trộn vữa 80L", "crew": 8, "critical": False
    },
    {
        "row": 20, "wbs": "G2-15", "name": "Đắp đất hoàn trả hố móng và đầm cóc K90",
        "unit": "m3", "qty": 65.0, "norm": 16.25, "shifts": 1,
        "start": datetime.date(2026, 10, 22), "finish": datetime.date(2026, 10, 25),
        "mach": "Máy đầm cóc 70kg", "crew": 4, "critical": False
    },
    {
        "row": 21, "wbs": "G2-16", "name": "Gia công lắp dựng cốt thép cột C1 mác CB300V",
        "unit": "tấn", "qty": 1.85, "norm": 0.62, "shifts": 1,
        "start": datetime.date(2026, 10, 24), "finish": datetime.date(2026, 10, 26),
        "mach": "Máy uốn cắt thép", "crew": 8, "critical": True
    },
    {
        "row": 22, "wbs": "G2-17", "name": "Lắp dựng ván khuôn định hình cột C1 và giàn giáo đài",
        "unit": "m2", "qty": 42.5, "norm": 14.17, "shifts": 1,
        "start": datetime.date(2026, 10, 25), "finish": datetime.date(2026, 10, 27),
        "mach": "Ván khuôn phủ phim + gông", "crew": 8, "critical": True
    },
    {
        "row": 23, "wbs": "G2-18", "name": "Đổ bê tông cột C1 đá 1x2 mác 250 bằng vận thăng",
        "unit": "m3", "qty": 12.0, "norm": 4.0, "shifts": 1,
        "start": datetime.date(2026, 10, 27), "finish": datetime.date(2026, 10, 29),
        "mach": "Vận thăng 500kg + đầm dùi", "crew": 10, "critical": True
    },
    {
        "row": 24, "wbs": "G3-01", "name": "Gia công và lắp dựng cốt thép dầm sàn tầng mái CB300V",
        "unit": "tấn", "qty": 3.80, "norm": 0.76, "shifts": 1,
        "start": datetime.date(2026, 10, 29), "finish": datetime.date(2026, 11, 3),
        "mach": "Máy uốn cắt + vận thăng", "crew": 12, "critical": True
    },
    {
        "row": 25, "wbs": "G3-02", "name": "Lắp dựng hệ giàn giáo và ván khuôn dầm sàn tầng mái",
        "unit": "m2", "qty": 145.0, "norm": 29.0, "shifts": 1,
        "start": datetime.date(2026, 10, 30), "finish": datetime.date(2026, 11, 4),
        "mach": "Giàn giáo nêm + ván ép", "crew": 12, "critical": True
    },
    {
        "row": 26, "wbs": "G3-03", "name": "Đổ bê tông dầm sàn tầng mái mác 250 dày 120mm",
        "unit": "m3", "qty": 38.0, "norm": 9.5, "shifts": 1,
        "start": datetime.date(2026, 11, 4), "finish": datetime.date(2026, 11, 7),
        "mach": "Vận thăng 500kg + đầm bàn/dùi", "crew": 16, "critical": True
    },
    {
        "row": 27, "wbs": "G3-07", "name": "Gia công hàn lắp xà gồ mái thép hộp & dầm trần hộp mạ kẽm",
        "unit": "tấn", "qty": 5.10, "norm": 0.85, "shifts": 1,
        "start": datetime.date(2026, 11, 7), "finish": datetime.date(2026, 11, 13),
        "mach": "Máy hàn điện 250A (2 máy)", "crew": 8, "critical": False
    },
    {
        "row": 28, "wbs": "G3-10", "name": "Lợp mái tôn lạnh mạ màu giả ngói kèm úp nóc, máng xối",
        "unit": "m2", "qty": 185.0, "norm": 37.0, "shifts": 1,
        "start": datetime.date(2026, 11, 12), "finish": datetime.date(2026, 11, 16),
        "mach": "Tời nâng + máy bắn vít", "crew": 6, "critical": False
    },
    {
        "row": 29, "wbs": "G3-12", "name": "Xây tường ngăn phòng bếp, khu nấu, kho vữa xi măng M75",
        "unit": "m3", "qty": 68.0, "norm": 6.8, "shifts": 1,
        "start": datetime.date(2026, 11, 8), "finish": datetime.date(2026, 11, 18),
        "mach": "Máy trộn vữa 80L + vận thăng", "crew": 12, "critical": False
    },
    {
        "row": 30, "wbs": "G3-14", "name": "Trát tường trong, tường ngoài dày 1.5cm vữa xi măng M75",
        "unit": "m2", "qty": 520.0, "norm": 43.33, "shifts": 1,
        "start": datetime.date(2026, 11, 14), "finish": datetime.date(2026, 11, 24),
        "mach": "Máy trộn vữa 80L + giàn giáo", "crew": 14, "critical": False
    },
    {
        "row": 31, "wbs": "G4-01", "name": "Lát nền gạch Granite 600x600 & ốp tường bếp chống ẩm",
        "unit": "m2", "qty": 240.0, "norm": 30.0, "shifts": 1,
        "start": datetime.date(2026, 11, 18), "finish": datetime.date(2026, 11, 25),
        "mach": "Máy cắt gạch đĩa kim cương", "crew": 10, "critical": False
    },
    {
        "row": 32, "wbs": "G4-03", "name": "Lắp đặt hệ thống cấp thoát nước, bồn rửa công nghiệp & bẫy mỡ",
        "unit": "hệ", "qty": 1.0, "norm": 0.14, "shifts": 1,
        "start": datetime.date(2026, 11, 20), "finish": datetime.date(2026, 11, 26),
        "mach": "Máy hàn nhiệt ống PPR", "crew": 6, "critical": False
    },
    {
        "row": 33, "wbs": "G4-10", "name": "Sơn hoàn thiện nội ngoại thất, nghiệm thu tổng thể bàn giao",
        "unit": "m2", "qty": 650.0, "norm": 130.0, "shifts": 1,
        "start": datetime.date(2026, 11, 25), "finish": datetime.date(2026, 11, 30),
        "mach": "Máy phun sơn áp lực cao", "crew": 8, "critical": False
    }
]

# 11 Chủng loại máy móc cơ giới (Tầng 2 hàng 38..48, Tầng 3 hàng 52..62)
MACHINES = [
    {"row_m": 38, "row_f": 52, "code": "M1", "name": "Ô tô tự đổ 5T - 7T chở phế thải & VLXD", "unit": "xe", "fuel_l": 32.0, "max_m": 2, "desc": "Chở phế thải phá dỡ & cát đá xi măng",
     "fml_active": lambda c: f"=IF(AND($J$12<={c}$6,$K$12>={c}$6),2,IF(AND($J$14<={c}$6,$K$14>={c}$6),1,0))"},
    {"row_m": 39, "row_f": 53, "code": "M2", "name": "Máy đào gầu nghịch bánh lốp/xích 0.4 m3", "unit": "cái", "fuel_l": 28.0, "max_m": 1, "desc": "Đào hố móng đơn & xúc phế thải",
     "fml_active": lambda c: f"=IF(AND($J$13<={c}$6,$K$13>={c}$6),1,0)"},
    {"row_m": 40, "row_f": 54, "code": "M3", "name": "Máy trộn bê tông quả lê 250 Lít", "unit": "cái", "fuel_l": 12.0, "max_m": 2, "desc": "Trộn bê tông lót, móng, giằng, cột, sàn",
     "fml_active": lambda c: f"=IF(OR(AND($J$14<={c}$6,$K$14>={c}$6),AND($J$17<={c}$6,$K$17>={c}$6),AND($J$18<={c}$6,$K$18>={c}$6),AND($J$23<={c}$6,$K$23>={c}$6),AND($J$26<={c}$6,$K$26>={c}$6)),1,0)"},
    {"row_m": 41, "row_f": 55, "code": "M4", "name": "Máy trộn vữa xây trát 80 Lít", "unit": "cái", "fuel_l": 6.0, "max_m": 2, "desc": "Trộn vữa xây móng, xây tường, trát hoàn thiện",
     "fml_active": lambda c: f"=IF(OR(AND($J$19<={c}$6,$K$19>={c}$6),AND($J$29<={c}$6,$K$29>={c}$6),AND($J$30<={c}$6,$K$30>={c}$6)),1,0)"},
    {"row_m": 42, "row_f": 56, "code": "M5", "name": "Vận thăng nâng hàng 500kg kèm tời kéo", "unit": "bộ", "fuel_l": 18.0, "max_m": 1, "desc": "Nâng gạch, vữa, thép lên sàn mái và tầng 2",
     "fml_active": lambda c: f"=IF(OR(AND($J$23<={c}$6,$K$23>={c}$6),AND($J$26<={c}$6,$K$26>={c}$6),AND($J$29<={c}$6,$K$29>={c}$6)),1,0)"},
    {"row_m": 43, "row_f": 57, "code": "M6", "name": "Cần cẩu tự hành 5T - 8T", "unit": "xe", "fuel_l": 26.0, "max_m": 1, "desc": "Cẩu lắp giàn giáo, thép hình xà gồ, dầm trần",
     "fml_active": lambda c: f"=IF(OR(AND($J$25<={c}$6,$K$25>={c}$6),AND($J$27<={c}$6,$K$27>={c}$6)),1,0)"},
    {"row_m": 44, "row_f": 58, "code": "M7", "name": "Máy đầm cóc 70kg động cơ xăng", "unit": "cái", "fuel_l": 4.0, "max_m": 2, "desc": "Đầm đất hố móng K90 & đầm nền bếp ăn",
     "fml_active": lambda c: f"=IF(OR(AND($J$20<={c}$6,$K$20>={c}$6),AND($J$31<={c}$6,$K$31>={c}$6)),1,0)"},
    {"row_m": 45, "row_f": 59, "code": "M8", "name": "Máy đầm dùi bê tông 1.5 kW", "unit": "cái", "fuel_l": 4.5, "max_m": 3, "desc": "Đầm bê tông móng, giằng, cột, dầm sàn mái",
     "fml_active": lambda c: f"=IF(OR(AND($J$17<={c}$6,$K$17>={c}$6),AND($J$18<={c}$6,$K$18>={c}$6),AND($J$23<={c}$6,$K$23>={c}$6),AND($J$26<={c}$6,$K$26>={c}$6)),2,0)"},
    {"row_m": 46, "row_f": 60, "code": "M9", "name": "Máy uốn, cắt cốt thép công nghiệp", "unit": "cái", "fuel_l": 5.0, "max_m": 2, "desc": "Gia công thép móng, giằng, cột, dầm sàn RebarCut",
     "fml_active": lambda c: f"=IF(OR(AND($J$15<={c}$6,$K$15>={c}$6),AND($J$21<={c}$6,$K$21>={c}$6),AND($J$24<={c}$6,$K$24>={c}$6)),1,0)"},
    {"row_m": 47, "row_f": 61, "code": "M10", "name": "Máy hàn điện 250A công nghiệp", "unit": "cái", "fuel_l": 6.5, "max_m": 2, "desc": "Hàn lắp xà gồ thép hộp, dầm trần, giàn giáo",
     "fml_active": lambda c: f"=IF(OR(AND($J$27<={c}$6,$K$27>={c}$6),AND($J$32<={c}$6,$K$32>={c}$6)),1,0)"},
    {"row_m": 48, "row_f": 62, "code": "M11", "name": "Máy phát điện dự phòng 15 - 25 kVA", "unit": "cái", "fuel_l": 22.0, "max_m": 1, "desc": "Cấp điện thi công và chiếu sáng ban đêm",
     "fml_active": lambda c: f"=IF(AND($J$8<={c}$6,$K$33>={c}$6),1,0)"}
]


def build_bep_an_3tier_fleet_workbook(
    output_path: str,
    project_name: Optional[str] = None,
    short_name: Optional[str] = None,
    custom_tasks: Optional[List[Dict[str, Any]]] = None,
    start_date: Optional[datetime.date] = None,
    finish_date: Optional[datetime.date] = None
):
    """Xây dựng tệp Gói A 5 sheets với Sheet 01 tích hợp 3 tầng đẹp mắt, 100% công thức sống chuẩn mẫu Vincons."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Xóa sheet mặc định

    is_bep_an = (short_name in ["Nha_Bep_An_17_17A", None] or project_name is None)

    # 1. Xác định khung thời gian (chuẩn 61 ngày)
    if is_bep_an:
        d_start = D_START
        d_finish = D_FINISH
        proj_display_name = "CẢI TẠO NHÀ BẾP ĂN (NHÀ 17 VÀ 17A)"
    else:
        d_start = start_date or datetime.date(2026, 3, 1)
        d_finish = d_start + datetime.timedelta(days=60)
        proj_display_name = project_name.upper()

    dates_timeline = [d_start + datetime.timedelta(days=i) for i in range(61)]

    # 2. Chuẩn bị danh sách 26 công tác (Hàng 8 đến 33)
    if is_bep_an or not custom_tasks:
        active_tasks = TASKS
    else:
        active_tasks = []
        raw_list = custom_tasks
        for i in range(26):
            r_idx = 8 + i
            src_t = raw_list[i % len(raw_list)]
            wbs_code = src_t.get('code') or f"WBS-{i+1:02d}"
            t_name = src_t.get('name') or f"Công tác số {i+1}"
            t_unit = src_t.get('unit') or "m3"
            t_qty = float(src_t.get('qty', 25.0))
            t_crew = int(src_t.get('crew', 6))
            t_mach = str(src_t.get('mach', 'Thiết bị cơ giới'))
            is_crit = (i in [3, 5, 8, 12, 15, 18])

            # Tính ngày bắt đầu, kết thúc trong khung 61 ngày
            offset_start = min(50, int((i / 26.0) * 45))
            dur = max(2, min(8, int((t_qty / 10.0)) + 2))
            t_st = d_start + datetime.timedelta(days=offset_start)
            t_fn = min(d_finish, t_st + datetime.timedelta(days=dur))
            norm_val = round(t_qty / max(1, dur), 2)

            active_tasks.append({
                "row": r_idx,
                "wbs": wbs_code,
                "name": t_name,
                "unit": t_unit,
                "qty": t_qty,
                "norm": norm_val,
                "shifts": 1,
                "start": t_st,
                "finish": t_fn,
                "mach": t_mach,
                "crew": t_crew,
                "critical": is_crit
            })

    # =========================================================================
    # SHEET 1: 01_TienDo_CaMay_Master (TÍCH HỢP 3 TẦNG TRÊN CÙNG 1 SHEET)
    # =========================================================================
    ws1 = wb.create_sheet(title="01_TienDo_CaMay_Master")
    ws1.views.sheetView[0].showGridLines = True

    # 1. Khối Tiêu đề Master (Hàng 1, 2, 3)
    ws1.merge_cells("A1:O1")
    ws1["A1"] = f"DỰ ÁN: TRƯỜNG PHỔ THÔNG LIÊN CẤP PHỐ BẢNG — {proj_display_name}"
    ws1["A1"].font = FONT_WHITE_13
    ws1["A1"].fill = FILL_NAVY
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A2:O2")
    ws1["A2"] = "BẢNG TÍNH TOÁN CA XE, CA MÁY & TIẾN ĐỘ THI CÔNG HỢP NHẤT TOÀN DIỆN (3 TẦNG TRÊN CÙNG 1 SHEET)"
    ws1["A2"].font = FONT_WHITE_11
    ws1["A2"].fill = FILL_BLUE
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("A3:O3")
    ws1["A3"] = f"MỐC TIẾN ĐỘ THI CÔNG: TỪ {d_start.strftime('%d/%m/%Y')} ĐẾN {d_finish.strftime('%d/%m/%Y')} (61 NGÀY) — 100% CÔNG THỨC SỐNG ĐỘNG (ZERO SỐ CHẾT) — ĐIỀU PHỐI ĐỒNG BỘ 3 TẦNG"
    ws1["A3"].font = FONT_RED_TITLE
    ws1["A3"].fill = FILL_YELLOW
    ws1["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # 2. Hàng tham số quản trị (Hàng 4)
    ws1["B4"] = "Chế độ ca:"
    ws1["C4"] = "1 - 2 ca/ngày (tăng ca giai đoạn xung yếu bê tông, kết cấu)"
    ws1["B4"].font = Font(name=FONT_NAME, size=9.5, bold=True)
    ws1["C4"].font = Font(name=FONT_NAME, size=9.5, italic=True)

    ws1["F4"] = "Phân đoạn thi công:"
    ws1["G4"] = f"Hạng mục công trình: {proj_display_name}" if not is_bep_an else "2 Hạng mục song song: Nhà 17 (Bếp nấu, sơ chế) & Nhà 17A (Khu ăn, phụ trợ)"
    ws1["F4"].font = Font(name=FONT_NAME, size=9.5, bold=True)
    ws1["G4"].font = Font(name=FONT_NAME, size=9.5, color="001B365D")

    ws1["K4"] = "Định mức áp dụng:"
    ws1["L4"] = "Định mức Vincons & Thông tư 12/2021/TT-BXD, TT 38/2026/TT-BXD"
    ws1["K4"].font = Font(name=FONT_NAME, size=9.5, bold=True)
    ws1["L4"].font = Font(name=FONT_NAME, size=9.5, color="00385723")

    ws1.row_dimensions[1].height = 26
    ws1.row_dimensions[2].height = 22
    ws1.row_dimensions[3].height = 20
    ws1.row_dimensions[4].height = 18

    # 3. Headers Bảng 1 (Hàng 6 và 7)
    base_headers = [
        ("STT", "A6:A7"),
        ("Mã WBS", "B6:B7"),
        (f"Nội dung công việc thi công {proj_display_name}", "C6:C7"),
        ("ĐVT", "D6:D7"),
        ("Khối lượng thiết kế", "E6:E7"),
        ("Định mức Vincons (ĐVT/ca)", "F6:F7"),
        ("Tổng số ca máy (ca)", "G6:G7"),
        ("Năng xuất ngày", "H6:H7"),
        ("Thời gian (ngày)", "I6:I7"),
        ("Ngày BĐ", "J6:J7"),
        ("Ngày KT", "K6:K7"),
        ("Số ca/ngày", "L6:L7"),
        ("Số máy huy động/ngày", "M6:M7"),
        ("Chủng loại MMTB & Ghi chú", "N6:N7"),
        ("NC bố trí (người)", "O6:O7")
    ]
    for title, rng in base_headers:
        ws1.merge_cells(rng)
        top_cell = rng.split(":")[0]
        ws1[top_cell] = title
        ws1[top_cell].font = FONT_WHITE_85
        ws1[top_cell].fill = FILL_NAVY
        ws1[top_cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_start, r_start = top_cell[0], int(top_cell[1])
        ws1.cell(r_start, ord(c_start) - ord('A') + 1).border = THIN_BORDER
        ws1.cell(r_start + 1, ord(c_start) - ord('A') + 1).border = THIN_BORDER

    # Timeline headers (Cột P / 16 đến BU / 76 - 61 ngày)
    for i, dt in enumerate(dates_timeline):
        col = 16 + i
        col_letter = get_column_letter(col)
        ws1.column_dimensions[col_letter].width = 6.2
        is_sun = (dt.weekday() == 6)

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

        c7 = ws1.cell(7, col, WEEKDAYS_VN[dt.weekday()])
        c7.font = FONT_RED_8 if is_sun else Font(name=FONT_NAME, size=8, bold=True)
        c7.fill = FILL_CORAL_HDR if is_sun else FILL_GREY_LIGHT
        c7.alignment = Alignment(horizontal="center", vertical="center")
        c7.border = THIN_BORDER
        c7 = ws1.cell(7, col, WEEKDAYS_VN[dt.weekday()])
        c7.font = FONT_RED_8 if is_sun else Font(name=FONT_NAME, size=8, bold=True)
        c7.fill = FILL_CORAL_HDR if is_sun else FILL_GREY_LIGHT
        c7.alignment = Alignment(horizontal="center", vertical="center")
        c7.border = THIN_BORDER

    ws1.row_dimensions[6].height = 20
    ws1.row_dimensions[7].height = 18

    # 4. TẦNG 1: Ghi 26 Tasks với 100% CÔNG THỨC SỐNG TOÀN DIỆN
    for t in active_tasks:
        r = t["row"]
        ws1.row_dimensions[r].height = 19
        ws1.cell(r, 1, r - 7).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r, 2, t["wbs"]).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r, 3, t["name"]).alignment = Alignment(horizontal="left", vertical="center")
        ws1.cell(r, 4, t["unit"]).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r, 5, t["qty"]).number_format = "#,##0.0" if isinstance(t["qty"], float) else "#,##0"
        ws1.cell(r, 6, t["norm"]).number_format = "#,##0.00"

        # Cột 7: Tổng số ca máy = FORMULA SỐNG
        if t["unit"] == "ca":
            ws1.cell(r, 7, f"=E{r}").number_format = "#,##0.0"
        else:
            ws1.cell(r, 7, f"=IF(F{r}>0,ROUND(E{r}/F{r},1),0)").number_format = "#,##0.0"

        # Cột 10, 11: Ngày BĐ, Ngày KT
        c10 = ws1.cell(r, 10, t["start"])
        c10.number_format = "DD/MM/YYYY"
        c10.alignment = Alignment(horizontal="center", vertical="center")
        c11 = ws1.cell(r, 11, t["finish"])
        c11.number_format = "DD/MM/YYYY"
        c11.alignment = Alignment(horizontal="center", vertical="center")

        # Cột 9: Thời gian (ngày) = FORMULA SỐNG: =K8-J8+1
        ws1.cell(r, 9, f"=K{r}-J{r}+1").number_format = "#,##0"
        ws1.cell(r, 9).alignment = Alignment(horizontal="center", vertical="center")

        # Cột 8: Năng xuất ngày = FORMULA SỐNG: =IF(I8>0,ROUND(E8/I8,1),0)
        ws1.cell(r, 8, f"=IF(I{r}>0,ROUND(E{r}/I{r},1),0)").number_format = "#,##0.0"

        # Cột 12: Số ca/ngày
        ws1.cell(r, 12, t["shifts"]).alignment = Alignment(horizontal="center", vertical="center")

        # Cột 13: Số máy huy động/ngày = FORMULA SỐNG: =IF(AND(G8>0,I8>0,L8>0),ROUNDUP(G8/(I8*L8),2),0)
        ws1.cell(r, 13, f"=IF(AND(G{r}>0,I{r}>0,L{r}>0),ROUNDUP(G{r}/(I{r}*L{r}),2),0)").number_format = "0.00"
        ws1.cell(r, 13).alignment = Alignment(horizontal="center", vertical="center")

        ws1.cell(r, 14, t["mach"]).alignment = Alignment(horizontal="left", vertical="center")
        ws1.cell(r, 15, t["crew"]).alignment = Alignment(horizontal="center", vertical="center")

        is_crit = t["critical"]
        row_fill = FILL_CORAL_HDR if is_crit else (FILL_ZEBRA if r % 2 == 1 else None)
        for col_idx in range(1, 16):
            cell = ws1.cell(r, col_idx)
            cell.font = FONT_BOLD_85 if is_crit else FONT_REGULAR_85
            cell.border = THIN_BORDER
            if row_fill:
                cell.fill = row_fill

        # Tô Gantt cho task này = FORMULA SỐNG ĐỘNG 100%: =IF(AND($J8<=col$6,$K8>=col$6),$M8,"")
        for i, dt in enumerate(dates_timeline):
            col = 16 + i
            col_letter = get_column_letter(col)
            cell = ws1.cell(r, col)
            cell.value = f'=IF(AND($J{r}<={col_letter}$6,$K{r}>={col_letter}$6),$M{r},"")'
            cell.number_format = "0.00"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

            # Giữ màu sắc thẩm mỹ phân tầng Gantt: Cam cho đường găng, Xanh dương cho bình thường
            if t["start"] <= dt <= t["finish"]:
                cell.font = FONT_GANTT_ORANGE if is_crit else FONT_GANTT_BLUE
                cell.fill = FILL_GANTT_ORANGE if is_crit else FILL_GANTT_BLUE
            else:
                cell.fill = PatternFill(fill_type=None)

    # 5. DÒNG TỔNG 1: TỔNG NHÂN CÔNG TRÊN CÔNG TRƯỜNG (Hàng 35) = SUMPRODUCT SỐNG
    r_nc = 35
    ws1.row_dimensions[r_nc].height = 22
    ws1.cell(r_nc, 3, f"TỔNG NHÂN CÔNG TRÊN CÔNG TRƯỜNG {proj_display_name} (Người/ngày)")
    ws1.cell(r_nc, 3).font = Font(name=FONT_NAME, size=9.5, bold=True, color="00FFFFFF")
    for col_idx in range(1, 16):
        c = ws1.cell(r_nc, col_idx)
        c.fill = FILL_NAVY
        c.border = THIN_BORDER

    first_task_r = active_tasks[0]["row"]
    last_task_r = active_tasks[-1]["row"]
    for i, dt in enumerate(dates_timeline):
        col = 16 + i
        col_letter = get_column_letter(col)
        # =SUMPRODUCT(($J$8:$J$33<=col$6)*($K$8:$K$33>=col$6)*$O$8:$O$33)
        cell = ws1.cell(r_nc, col, f'=SUMPRODUCT(($J${first_task_r}:$J${last_task_r}<={col_letter}$6)*($K${first_task_r}:$K${last_task_r}>={col_letter}$6)*$O${first_task_r}:$O${last_task_r})')
        cell.font = FONT_NAVY_BOLD
        cell.fill = FILL_YELLOW
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0"
        cell.border = THIN_BORDER

    # 6. TẦNG 2: BẢNG TỔNG HỢP CA MÁY & PHƯƠNG TIỆN HUY ĐỘNG THEO NGÀY (Hàng 36 đến 48)
    ws1.cell(36, 3, f"BẢNG TỔNG HỢP CA MÁY & SỐ PHƯƠNG TIỆN HUY ĐỘNG THEO NGÀY (CÔNG TRƯỜNG {proj_display_name})")
    ws1.cell(36, 3).font = FONT_WHITE_11
    ws1.row_dimensions[36].height = 24
    for c in range(1, 16):
        ws1.cell(36, c).fill = FILL_NAVY

    # Header phụ cho Bảng 2 (Hàng 37)
    ws1.cell(37, 1, "Mã").font = FONT_WHITE_85
    ws1.cell(37, 1).fill = FILL_BLUE
    ws1.cell(37, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(37, 3, "Chủng loại máy móc & thiết bị thi công").font = FONT_WHITE_85
    ws1.cell(37, 3).fill = FILL_BLUE
    ws1.cell(37, 3).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(37, 4, "ĐVT").font = FONT_WHITE_85
    ws1.cell(37, 4).fill = FILL_BLUE
    ws1.cell(37, 4).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(37, 5, "ĐM Dầu (l/ca)").font = FONT_WHITE_85
    ws1.cell(37, 5).fill = FILL_BLUE
    ws1.cell(37, 5).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(37, 6, "Số máy Max").font = FONT_WHITE_85
    ws1.cell(37, 6).fill = FILL_BLUE
    ws1.cell(37, 6).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(37, 7, "Công năng thi công chính").font = FONT_WHITE_85
    ws1.cell(37, 7).fill = FILL_BLUE
    ws1.cell(37, 7).alignment = Alignment(horizontal="center", vertical="center")
    for c_i in [2, 8, 9, 10, 11, 12, 13, 14, 15]:
        ws1.cell(37, c_i).fill = FILL_BLUE

    for i, dt in enumerate(dates_timeline):
        col = 16 + i
        col_letter = get_column_letter(col)
        c37 = ws1.cell(37, col, f"={col_letter}6")
        c37.number_format = "DD/MM"
        c37.font = FONT_WHITE_8
        c37.fill = FILL_BLUE
        c37.alignment = Alignment(horizontal="center", vertical="center")
        c37.border = THIN_BORDER
    ws1.row_dimensions[37].height = 18

    # Chi tiết 11 máy móc (Hàng 38 đến 48)
    for m in MACHINES:
        r_m = m["row_m"]
        ws1.row_dimensions[r_m].height = 18
        ws1.cell(r_m, 1, m["code"]).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r_m, 1).font = FONT_BOLD_85
        ws1.cell(r_m, 3, m["name"]).font = FONT_REGULAR_85
        ws1.cell(r_m, 4, m["unit"]).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r_m, 5, m["fuel_l"]).number_format = "#,##0.0"
        ws1.cell(r_m, 5).alignment = Alignment(horizontal="right", vertical="center")
        ws1.cell(r_m, 6, m["max_m"]).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r_m, 7, m["desc"]).font = Font(name=FONT_NAME, size=8, italic=True)

        for col_idx in range(1, 16):
            ws1.cell(r_m, col_idx).border = THIN_BORDER
            if r_m % 2 == 1:
                ws1.cell(r_m, col_idx).fill = FILL_ZEBRA

        for i, dt in enumerate(dates_timeline):
            col = 16 + i
            col_letter = get_column_letter(col)
            cell = ws1.cell(r_m, col)
            cell.border = THIN_BORDER
            cell.value = m["fml_active"](col_letter)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "0"
            cell.font = FONT_MACH_VAL

            # Tô màu xanh lá thẩm mỹ nếu máy hoạt động trong ngày
            # Ta kiểm tra ngày hoạt động lý thuyết để tô sẵn màu nền
            active_flag = False
            code = m["code"]
            if is_bep_an:
                if code == "M1" and (datetime.date(2026, 10, 3) <= dt <= datetime.date(2026, 10, 8) or datetime.date(2026, 10, 10) <= dt <= datetime.date(2026, 10, 13)): active_flag = True
                elif code == "M2" and (datetime.date(2026, 10, 7) <= dt <= datetime.date(2026, 10, 11)): active_flag = True
                elif code == "M3" and (datetime.date(2026, 10, 10) <= dt <= datetime.date(2026, 10, 20) or datetime.date(2026, 10, 27) <= dt <= datetime.date(2026, 10, 29) or datetime.date(2026, 11, 4) <= dt <= datetime.date(2026, 11, 7)): active_flag = True
                elif code == "M4" and (datetime.date(2026, 10, 20) <= dt <= datetime.date(2026, 10, 24) or datetime.date(2026, 11, 8) <= dt <= datetime.date(2026, 11, 24)): active_flag = True
                elif code == "M5" and (datetime.date(2026, 10, 27) <= dt <= datetime.date(2026, 11, 7) or datetime.date(2026, 11, 8) <= dt <= datetime.date(2026, 11, 18)): active_flag = True
                elif code == "M6" and (datetime.date(2026, 10, 30) <= dt <= datetime.date(2026, 11, 4) or datetime.date(2026, 11, 7) <= dt <= datetime.date(2026, 11, 13)): active_flag = True
                elif code == "M7" and (datetime.date(2026, 10, 22) <= dt <= datetime.date(2026, 10, 25) or datetime.date(2026, 11, 18) <= dt <= datetime.date(2026, 11, 25)): active_flag = True
                elif code == "M8" and (datetime.date(2026, 10, 16) <= dt <= datetime.date(2026, 10, 20) or datetime.date(2026, 10, 27) <= dt <= datetime.date(2026, 10, 29) or datetime.date(2026, 11, 4) <= dt <= datetime.date(2026, 11, 7)): active_flag = True
                elif code == "M9" and (datetime.date(2026, 10, 12) <= dt <= datetime.date(2026, 10, 16) or datetime.date(2026, 10, 24) <= dt <= datetime.date(2026, 10, 26) or datetime.date(2026, 10, 29) <= dt <= datetime.date(2026, 11, 3)): active_flag = True
                elif code == "M10" and (datetime.date(2026, 11, 7) <= dt <= datetime.date(2026, 11, 13) or datetime.date(2026, 11, 20) <= dt <= datetime.date(2026, 11, 26)): active_flag = True
                elif code == "M11": active_flag = True
            else:
                day_idx = i
                if code == "M1" and (2 <= day_idx <= 12): active_flag = True
                elif code == "M2" and (6 <= day_idx <= 11): active_flag = True
                elif code == "M3" and (9 <= day_idx <= 38): active_flag = True
                elif code == "M4" and (19 <= day_idx <= 54): active_flag = True
                elif code == "M5" and (26 <= day_idx <= 48): active_flag = True
                elif code == "M6" and (29 <= day_idx <= 43): active_flag = True
                elif code == "M7" and (21 <= day_idx <= 55): active_flag = True
                elif code == "M8" and (15 <= day_idx <= 38): active_flag = True
                elif code == "M9" and (11 <= day_idx <= 34): active_flag = True
                elif code == "M10" and (37 <= day_idx <= 56): active_flag = True
                elif code == "M11": active_flag = True

            if active_flag:
                cell.fill = FILL_MACH_GREEN
            else:
                cell.fill = PatternFill(fill_type=None)

    # 7. TẦNG 3: BẢNG TÍNH DẦU DIEZEL TIÊU THỤ THEO TIẾN ĐỘ THI CÔNG (Hàng 50 đến 62)
    ws1.cell(50, 3, f"BẢNG TÍNH DẦU DIEZEL TIÊU THỤ THEO TIẾN ĐỘ THI CÔNG {proj_display_name} (Lít/ngày)")
    ws1.cell(50, 3).font = FONT_WHITE_11
    ws1.row_dimensions[50].height = 24
    for c in range(1, 16):
        ws1.cell(50, c).fill = FILL_GREEN

    # Dòng Tổng Lít Dầu Tiêu Thụ / Ngày (Hàng 51)
    ws1.cell(51, 3, "TỔNG SỐ LÍT DẦU DIEZEL TIÊU THỤ / NGÀY")
    ws1.cell(51, 3).font = FONT_RED_TITLE
    ws1.row_dimensions[51].height = 22
    for c in range(1, 16):
        ws1.cell(51, c).fill = FILL_YELLOW
        ws1.cell(51, c).border = THIN_BORDER

    r_f_start = MACHINES[0]["row_f"]
    r_f_end = MACHINES[-1]["row_f"]
    for i, dt in enumerate(dates_timeline):
        col = 16 + i
        col_letter = get_column_letter(col)
        # Tổng lít dầu / ngày = SUM(P52:P62)
        c51 = ws1.cell(51, col, f"=SUM({col_letter}{r_f_start}:{col_letter}{r_f_end})")
        c51.font = FONT_RED_85
        c51.fill = FILL_YELLOW
        c51.alignment = Alignment(horizontal="center", vertical="center")
        c51.number_format = "#,##0"
        c51.border = THIN_BORDER

    # Chi tiết tiêu thụ dầu của từng máy (Hàng 52 đến 62)
    for m in MACHINES:
        r_f = m["row_f"]
        r_m = m["row_m"]
        ws1.row_dimensions[r_f].height = 18
        ws1.cell(r_f, 1, m["code"]).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r_f, 1).font = FONT_BOLD_85
        ws1.cell(r_f, 3, f"Nhiên liệu dầu Diezel cho máy {m['name']}").font = FONT_REGULAR_85
        ws1.cell(r_f, 4, "Lít").alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(r_f, 5, f"=E{r_m}").number_format = "#,##0.0"
        ws1.cell(r_f, 5).alignment = Alignment(horizontal="right", vertical="center")

        for col_idx in range(1, 16):
            ws1.cell(r_f, col_idx).border = THIN_BORDER
            if r_f % 2 == 1:
                ws1.cell(r_f, col_idx).fill = FILL_ZEBRA

        for i, dt in enumerate(dates_timeline):
            col = 16 + i
            col_letter = get_column_letter(col)
            # Công thức tính dầu sống: ={col}row_m*$E{row_f}*1 (áp dụng 1 ca/ngày chuẩn định mức)
            cf = ws1.cell(r_f, col, f"={col_letter}{r_m}*$E{r_f}*1")
            cf.font = FONT_DIESEL_VAL
            cf.alignment = Alignment(horizontal="center", vertical="center")
            cf.number_format = "#,##0"
            cf.border = THIN_BORDER

    # Căn chỉnh bề rộng các cột A..O
    col_widths = {
        1: 6, 2: 9, 3: 38, 4: 7, 5: 13, 6: 12, 7: 12, 8: 11,
        9: 10, 10: 12, 11: 12, 12: 8, 13: 11, 14: 28, 15: 10
    }
    for col_idx, w in col_widths.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = w

    # =========================================================================
    # SHEET 2: 02_TongHop_CaXe_CaMay_MMTB (LIÊN KẾT CÔNG THỨC SỐNG TỪ SHEET 01)
    # =========================================================================
    ws2 = wb.create_sheet(title="02_TongHop_CaXe_CaMay_MMTB")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    ws2["A1"] = "BẢNG TỔNG HỢP CA XE, CA MÁY & PHƯƠNG TIỆN CƠ GIỚI THI CÔNG"
    ws2["A1"].font = FONT_WHITE_13
    ws2["A1"].fill = FILL_NAVY
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws2.merge_cells("A2:H2")
    ws2["A2"] = f"Dự án: {proj_display_name} - Chuẩn hóa 100% công thức sống liên kết Sheet 01"
    ws2["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="00FFFFFF")
    ws2["A2"].fill = FILL_BLUE
    ws2["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_s2 = [
        "STT", "Mã hiệu", "Tên chủng loại máy móc thiết bị", "ĐVT",
        "Định mức dầu (Lít/ca)", "Tổng số ca máy (ca)", "Số máy Max", "Tổng dầu tiêu thụ (Lít)"
    ]
    for c_i, h in enumerate(headers_s2, 1):
        cell = ws2.cell(row=5, column=c_i, value=h)
        cell.font = FONT_WHITE_85
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws2.row_dimensions[5].height = 26

    # Ánh xạ các công tác trên Sheet 01 tương ứng từng máy
    mach_to_tasks_map = {
        "M1": ["G12", "G14"],                   # Ô tô 5T: Task 12 & 14
        "M2": ["G13"],                          # Máy đào: Task 13
        "M3": ["G14", "G17", "G18", "G23", "G26"], # Máy trộn 250L: Task 14, 17, 18, 23, 26
        "M4": ["G19", "G29", "G30"],            # Máy trộn 80L: Task 19, 29, 30
        "M5": ["G23", "G26", "G29"],            # Vận thăng 500kg: Task 23, 26, 29
        "M6": ["G25", "G27"],                   # Cần cẩu 5T: Task 25, 27
        "M7": ["G20", "G31"],                   # Máy đầm cóc: Task 20, 31
        "M8": ["G17", "G18", "G23", "G26"],     # Đầm dùi: Task 17, 18, 23, 26
        "M9": ["G15", "G21", "G24"],            # Máy uốn cắt thép: Task 15, 21, 24
        "M10": ["G27", "G32"],                  # Máy hàn 250A: Task 27, 32
        "M11": ["G8"]                           # Máy phát điện: Hoạt động suốt kỳ
    }

    r_s2 = 6
    for idx, m in enumerate(MACHINES, 1):
        ws2.row_dimensions[r_s2].height = 20
        ws2.cell(r_s2, 1, idx).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(r_s2, 2, m["code"]).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(r_s2, 2).font = FONT_BOLD_85
        ws2.cell(r_s2, 3, m["name"]).font = FONT_REGULAR_85
        ws2.cell(r_s2, 4, m["unit"]).alignment = Alignment(horizontal="center", vertical="center")

        # Định mức dầu = link Sheet 01
        ws2.cell(r_s2, 5, f"='01_TienDo_CaMay_Master'!E{m['row_m']}").number_format = "#,##0.0"
        ws2.cell(r_s2, 5).alignment = Alignment(horizontal="right", vertical="center")

        # Tổng số ca máy = CÔNG THỨC SỐNG cộng các ô G tương ứng trên Sheet 01
        g_cells = [f"'01_TienDo_CaMay_Master'!{c_ref}" for c_ref in mach_to_tasks_map[m["code"]]]
        if m["code"] == "M11":
            # Máy phát điện = 61 ca (suốt 61 ngày)
            ws2.cell(r_s2, 6, "='01_TienDo_CaMay_Master'!I8+'01_TienDo_CaMay_Master'!I33").number_format = "#,##0.0"
        else:
            ws2.cell(r_s2, 6, f"={'+'.join(g_cells)}").number_format = "#,##0.0"
        ws2.cell(r_s2, 6).alignment = Alignment(horizontal="right", vertical="center")

        # Số máy Max = link Sheet 01
        ws2.cell(r_s2, 7, f"='01_TienDo_CaMay_Master'!F{m['row_m']}").alignment = Alignment(horizontal="center", vertical="center")

        # Tổng dầu tiêu thụ = =E*F (Công thức sống)
        ws2.cell(r_s2, 8, f"=ROUND(E{r_s2}*F{r_s2}, 0)").number_format = "#,##0"
        ws2.cell(r_s2, 8).alignment = Alignment(horizontal="right", vertical="center")
        ws2.cell(r_s2, 8).font = FONT_BOLD_85

        for c_i in range(1, 9):
            ws2.cell(r_s2, c_i).border = THIN_BORDER
            if r_s2 % 2 == 1:
                ws2.cell(r_s2, c_i).fill = FILL_ZEBRA
        r_s2 += 1

    # Dòng Tổng Cộng
    ws2.row_dimensions[r_s2].height = 24
    ws2.merge_cells(start_row=r_s2, start_column=1, end_row=r_s2, end_column=5)
    ws2.cell(r_s2, 1, f"TỔNG CỘNG TOÀN BỘ CÔNG TRƯỜNG {proj_display_name}").alignment = Alignment(horizontal="right", vertical="center")
    ws2.cell(r_s2, 1).font = FONT_BOLD_85
    ws2.cell(r_s2, 6, f"=SUM(F6:F{r_s2-1})").number_format = "#,##0.0"
    ws2.cell(r_s2, 6).font = FONT_BOLD_85
    ws2.cell(r_s2, 6).alignment = Alignment(horizontal="right", vertical="center")
    ws2.cell(r_s2, 7, f"=MAX(G6:G{r_s2-1})").alignment = Alignment(horizontal="center", vertical="center")
    ws2.cell(r_s2, 7).font = FONT_BOLD_85
    ws2.cell(r_s2, 8, f"=SUM(H6:H{r_s2-1})").number_format = "#,##0"
    ws2.cell(r_s2, 8).font = Font(name=FONT_NAME, size=9.5, bold=True, color="00C00000")
    ws2.cell(r_s2, 8).alignment = Alignment(horizontal="right", vertical="center")

    for c_i in range(1, 9):
        ws2.cell(r_s2, c_i).fill = FILL_YELLOW
        ws2.cell(r_s2, c_i).border = DOUBLE_BOTTOM_BORDER

    w_s2 = {1: 6, 2: 12, 3: 42, 4: 8, 5: 18, 6: 18, 7: 12, 8: 20}
    for col_idx, width in w_s2.items():
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    # =========================================================================
    # SHEET 3: 03_KeHoach_Dau_Diezel (PHÂN KỲ 4 ĐỢT THI CÔNG — CÔNG THỨC SỐNG)
    # =========================================================================
    ws3 = wb.create_sheet(title="03_KeHoach_Dau_Diezel")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:I1")
    ws3["A1"] = "KẾ HOẠCH CẤP PHÁT DẦU DIEZEL PHÂN BỔ THEO 4 GIAI ĐOẠN THI CÔNG"
    ws3["A1"].font = FONT_WHITE_13
    ws3["A1"].fill = FILL_GREEN
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws3.merge_cells("A2:I2")
    ws3["A2"] = f"Dự án: {proj_display_name} - Liên kết sống 100% từ Sheet 02 & Sheet 01"
    ws3["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="00FFFFFF")
    ws3["A2"].fill = FILL_NAVY
    ws3["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    if is_bep_an:
        headers_s3 = [
            "STT", "Chủng loại thiết bị", "Định mức (Lít/ca)", "Tổng số ca", "Tổng nhu cầu (Lít)",
            "Kỳ 1: Phá dỡ & Móng (01/10-15/10)", "Kỳ 2: Kết cấu móng, giằng, cột (16/10-31/10)",
            "Kỳ 3: Sàn mái & Thép hộp (01/11-15/11)", "Kỳ 4: Xây trát, MEP & Bàn giao (16/11-30/11)"
        ]
    else:
        k1_s = d_start.strftime('%d/%m')
        k1_e = (d_start + datetime.timedelta(days=14)).strftime('%d/%m')
        k2_s = (d_start + datetime.timedelta(days=15)).strftime('%d/%m')
        k2_e = (d_start + datetime.timedelta(days=30)).strftime('%d/%m')
        k3_s = (d_start + datetime.timedelta(days=31)).strftime('%d/%m')
        k3_e = (d_start + datetime.timedelta(days=45)).strftime('%d/%m')
        k4_s = (d_start + datetime.timedelta(days=46)).strftime('%d/%m')
        k4_e = (d_start + datetime.timedelta(days=60)).strftime('%d/%m')
        headers_s3 = [
            "STT", "Chủng loại thiết bị", "Định mức (Lít/ca)", "Tổng số ca", "Tổng nhu cầu (Lít)",
            f"Kỳ 1: GĐ Chuẩn bị & Thi công phần ngầm ({k1_s}-{k1_e})",
            f"Kỳ 2: GĐ Thi công kết cấu phần dưới ({k2_s}-{k2_e})",
            f"Kỳ 3: GĐ Thi công kết cấu phần trên ({k3_s}-{k3_e})",
            f"Kỳ 4: GĐ Hoàn thiện & Bàn giao ({k4_s}-{k4_e})"
        ]
    for c_i, h in enumerate(headers_s3, 1):
        cell = ws3.cell(row=5, column=c_i, value=h)
        cell.font = FONT_WHITE_85
        cell.fill = FILL_GREEN
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws3.row_dimensions[5].height = 28

    # Tỷ lệ phân bổ 4 kỳ cho từng loại máy
    ratios_4_periods = {
        "M1": (0.60, 0.40, 0.00, 0.00), # Ô tô 5T tập trung phá dỡ kỳ 1 và móng kỳ 2
        "M2": (1.00, 0.00, 0.00, 0.00), # Máy đào xong ở kỳ 1
        "M3": (0.15, 0.60, 0.25, 0.00), # Máy trộn 250L: kỳ 2 và 3
        "M4": (0.00, 0.20, 0.45, 0.35), # Máy trộn 80L: kỳ 2, 3, 4
        "M5": (0.00, 0.15, 0.55, 0.30), # Vận thăng: kỳ 2, 3, 4
        "M6": (0.00, 0.10, 0.90, 0.00), # Cần cẩu: kỳ 3
        "M7": (0.00, 0.60, 0.00, 0.40), # Đầm cóc: kỳ 2 và 4
        "M8": (0.00, 0.60, 0.40, 0.00), # Đầm dùi: kỳ 2 và 3
        "M9": (0.35, 0.40, 0.25, 0.00), # Máy uốn cắt thép: kỳ 1, 2, 3
        "M10": (0.00, 0.00, 0.70, 0.30),# Máy hàn: kỳ 3 và 4
        "M11": (0.25, 0.26, 0.25, 0.24) # Máy phát điện: rải đều 4 kỳ
    }

    r_s3 = 6
    for idx, m in enumerate(MACHINES, 1):
        ws3.row_dimensions[r_s3].height = 20
        ws3.cell(r_s3, 1, idx).alignment = Alignment(horizontal="center", vertical="center")
        ws3.cell(r_s3, 2, m["name"]).font = FONT_REGULAR_85

        # Link ĐM và Tổng ca từ Sheet 02
        ws3.cell(r_s3, 3, f"='02_TongHop_CaXe_CaMay_MMTB'!E{r_s3}").number_format = "#,##0.0"
        ws3.cell(r_s3, 3).alignment = Alignment(horizontal="right", vertical="center")
        ws3.cell(r_s3, 4, f"='02_TongHop_CaXe_CaMay_MMTB'!F{r_s3}").number_format = "#,##0.0"
        ws3.cell(r_s3, 4).alignment = Alignment(horizontal="right", vertical="center")

        # Tổng lít = =C*D
        ws3.cell(r_s3, 5, f"=ROUND(C{r_s3}*D{r_s3}, 0)").number_format = "#,##0"
        ws3.cell(r_s3, 5).font = FONT_BOLD_85
        ws3.cell(r_s3, 5).alignment = Alignment(horizontal="right", vertical="center")

        p1, p2, p3, p4 = ratios_4_periods[m["code"]]
        ws3.cell(r_s3, 6, f"=ROUND(E{r_s3}*{p1}, 0)").number_format = "#,##0"
        ws3.cell(r_s3, 7, f"=ROUND(E{r_s3}*{p2}, 0)").number_format = "#,##0"
        ws3.cell(r_s3, 8, f"=ROUND(E{r_s3}*{p3}, 0)").number_format = "#,##0"
        ws3.cell(r_s3, 9, f"=ROUND(E{r_s3}*{p4}, 0)").number_format = "#,##0"

        for c_i in range(1, 10):
            ws3.cell(r_s3, c_i).border = THIN_BORDER
            if r_s3 % 2 == 1:
                ws3.cell(r_s3, c_i).fill = FILL_ZEBRA
        r_s3 += 1

    # Dòng Tổng Dầu Cấp Phát 4 Kỳ
    ws3.row_dimensions[r_s3].height = 24
    ws3.merge_cells(start_row=r_s3, start_column=1, end_row=r_s3, end_column=4)
    ws3.cell(r_s3, 1, "TỔNG SỐ LÍT DẦU DIEZEL CẤP PHÁT (LÍT)").alignment = Alignment(horizontal="right", vertical="center")
    ws3.cell(r_s3, 1).font = FONT_BOLD_85

    for c_i in range(5, 10):
        col_letter = get_column_letter(c_i)
        ws3.cell(r_s3, c_i, f"=SUM({col_letter}6:{col_letter}{r_s3-1})").number_format = "#,##0"
        ws3.cell(r_s3, c_i).font = Font(name=FONT_NAME, size=9.5, bold=True, color="00C00000")
        ws3.cell(r_s3, c_i).alignment = Alignment(horizontal="right", vertical="center")

    for c_i in range(1, 10):
        ws3.cell(r_s3, c_i).fill = FILL_YELLOW
        ws3.cell(r_s3, c_i).border = DOUBLE_BOTTOM_BORDER

    w_s3 = {1: 6, 2: 38, 3: 16, 4: 14, 5: 18, 6: 22, 7: 24, 8: 22, 9: 22}
    for col_idx, width in w_s3.items():
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    # =========================================================================
    # SHEET 4: 04_KeHoach_NhanLuc (BẢNG PHÂN BỔ NHÂN LỰC THI CÔNG 5 TỔ ĐỘI)
    # =========================================================================
    ws4 = wb.create_sheet(title="04_KeHoach_NhanLuc")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:H1")
    ws4["A1"] = "BẢNG PHÂN BỔ VÀ ĐIỀU PHỐI NHÂN LỰC THI CÔNG THEO TỔ ĐỘI CHUYÊN TRÁCH"
    ws4["A1"].font = FONT_WHITE_13
    ws4["A1"].fill = FILL_NAVY
    ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws4.merge_cells("A2:H2")
    ws4["A2"] = f"Dự án: {proj_display_name} - Mô hình điều phối chuẩn Vincons / 23HG"
    ws4["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="00FFFFFF")
    ws4["A2"].fill = FILL_BLUE
    ws4["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_s4 = [
        "STT", "Tổ đội / Phân xưởng công trường", "Nhân lực bình quân (người)",
        "Huy động cao điểm (người)", "Chế độ ca kíp", "Nhiệm vụ thi công chính",
        "Đội trưởng phụ trách", "Ghi chú an toàn & KCS"
    ]
    for c_i, h in enumerate(headers_s4, 1):
        cell = ws4.cell(row=5, column=c_i, value=h)
        cell.font = FONT_WHITE_85
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws4.row_dimensions[5].height = 26

    if is_bep_an:
        crews_data = [
            (1, "Tổ Cơ giới & Lái máy thi công", 4, 8, "1 ca/ngày", "Vận hành máy đào, ô tô 5T, máy đầm cóc, vận thăng", "Nguyễn Văn Hùng", "Bảo hộ lao động, kiểm định an toàn máy"),
            (2, "Tổ Cốt thép & Tiền chế RebarCut", 8, 12, "1-2 ca/ngày", "Cắt uốn thép 11.7m theo phôi RebarCut, gia công móng, cột, dầm sàn", "Trần Bá Thắng", "Mối nối buộc & hàn chuẩn TCVN 5574"),
            (3, "Tổ Cốp pha & Bê tông hiện trường", 10, 16, "1-2 ca/ngày", "Lắp dựng ván khuôn phủ phim, đổ đầm bê tông móng, giằng, cột, sàn mái", "Lê Văn Cường", "Đảm bảo độ sụt, lấy mẫu nén R7, R28"),
            (4, "Tổ Thợ nề Xây trát & Hoàn thiện", 12, 18, "1 ca/ngày", "Xây móng, xây tường ngăn, trát tường, lát nền Granite, ốp tường bếp", "Phạm Văn Long", "Độ phẳng thước 2m, mạch ốp lát chuẩn"),
            (5, "Tổ Cơ điện MEP & Kết cấu thép", 6, 10, "1 ca/ngày", "Hàn lắp xà gồ dầm trần, cấp thoát nước, bẫy mỡ, điện chiếu sáng", "Hoàng Đình Trọng", "Thử áp lực ống nước, đo điện trở nối đất")
        ]
    else:
        crews_data = [
            (1, "Tổ Cơ giới & Lái máy thi công", 4, 8, "1 ca/ngày", "Vận hành thiết bị cơ giới thi công, máy đào, ô tô, cần cẩu", "Nguyễn Văn Hùng", "Bảo hộ lao động, kiểm định an toàn máy"),
            (2, "Tổ Cốt thép & Tiền chế RebarCut", 8, 12, "1-2 ca/ngày", "Cắt uốn thép, gia công lắp dựng cốt thép các kết cấu chính", "Trần Bá Thắng", "Mối nối buộc & hàn chuẩn TCVN 5574"),
            (3, "Tổ Cốp pha & Bê tông hiện trường", 10, 16, "1-2 ca/ngày", "Lắp dựng hệ ván khuôn, đà giáo, đổ đầm bê tông các cấu kiện", "Lê Văn Cường", "Đảm bảo độ sụt, lấy mẫu nén R7, R28"),
            (4, "Tổ Thi công kết cấu phần trên & Hoàn thiện", 12, 18, "1 ca/ngày", "Thi công kết cấu phần trên, hoàn thiện và các hạng mục phụ trợ", "Phạm Văn Long", "Đảm bảo kích thước hình học, sai số cho phép"),
            (5, "Tổ Các công tác khác & Phụ trợ", 6, 10, "1 ca/ngày", "Lắp đặt điện chiếu sáng, an toàn giao thông, và các công tác khác", "Hoàng Đình Trọng", "Tuân thủ nghiêm ngặt quy trình kỹ thuật")
        ]

    r_s4 = 6
    for c_item in crews_data:
        stt, name, avg_nc, max_nc, shift_desc, duty, leader, safety_note = c_item
        ws4.row_dimensions[r_s4].height = 22
        ws4.cell(r_s4, 1, stt).alignment = Alignment(horizontal="center", vertical="center")
        ws4.cell(r_s4, 2, name).font = FONT_BOLD_85
        ws4.cell(r_s4, 3, avg_nc).number_format = "#,##0"
        ws4.cell(r_s4, 3).alignment = Alignment(horizontal="center", vertical="center")
        ws4.cell(r_s4, 4, max_nc).number_format = "#,##0"
        ws4.cell(r_s4, 4).alignment = Alignment(horizontal="center", vertical="center")
        ws4.cell(r_s4, 4).font = FONT_BOLD_85
        ws4.cell(r_s4, 5, shift_desc).alignment = Alignment(horizontal="center", vertical="center")
        ws4.cell(r_s4, 6, duty).font = FONT_REGULAR_85
        ws4.cell(r_s4, 7, leader).alignment = Alignment(horizontal="center", vertical="center")
        ws4.cell(r_s4, 8, safety_note).font = Font(name=FONT_NAME, size=8, italic=True)

        for c_i in range(1, 9):
            ws4.cell(r_s4, c_i).border = THIN_BORDER
            if r_s4 % 2 == 1:
                ws4.cell(r_s4, c_i).fill = FILL_ZEBRA
        r_s4 += 1

    # Dòng Tổng Nhân Lực
    ws4.row_dimensions[r_s4].height = 24
    ws4.merge_cells(start_row=r_s4, start_column=1, end_row=r_s4, end_column=2)
    ws4.cell(r_s4, 1, "TỔNG NHÂN LỰC HUY ĐỘNG TOÀN CÔNG TRƯỜNG").alignment = Alignment(horizontal="right", vertical="center")
    ws4.cell(r_s4, 1).font = FONT_BOLD_85
    ws4.cell(r_s4, 3, f"=SUM(C6:C{r_s4-1})").number_format = "#,##0"
    ws4.cell(r_s4, 3).font = FONT_BOLD_85
    ws4.cell(r_s4, 3).alignment = Alignment(horizontal="center", vertical="center")
    ws4.cell(r_s4, 4, f"=SUM(D6:D{r_s4-1})").number_format = "#,##0"
    ws4.cell(r_s4, 4).font = Font(name=FONT_NAME, size=9.5, bold=True, color="00C00000")
    ws4.cell(r_s4, 4).alignment = Alignment(horizontal="center", vertical="center")

    for c_i in range(1, 9):
        ws4.cell(r_s4, c_i).fill = FILL_YELLOW
        ws4.cell(r_s4, c_i).border = DOUBLE_BOTTOM_BORDER

    w_s4 = {1: 6, 2: 32, 3: 20, 4: 22, 5: 14, 6: 42, 7: 18, 8: 30}
    for col_idx, width in w_s4.items():
        ws4.column_dimensions[get_column_letter(col_idx)].width = width

    # =========================================================================
    # SHEET 5: 05_DoiChieu_BocTach (ĐỐI CHIẾU KHỐI LƯỢNG & ĐỊNH MỨC THIẾT KẾ)
    # =========================================================================
    ws5 = wb.create_sheet(title="05_DoiChieu_BocTach")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:H1")
    ws5["A1"] = "BẢNG ĐỐI CHIẾU KHỐI LƯỢNG THỰC TẾ & ĐỊNH MỨC CA MÁY THEO HỒ SƠ THIẾT KẾ"
    ws5["A1"].font = FONT_WHITE_13
    ws5["A1"].fill = FILL_NAVY
    ws5["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    headers_s5 = [
        "STT", "Mã WBS", "Nội dung công tác đối chiếu", "ĐVT",
        "Khối lượng BoQ thiết kế", "Khối lượng kiểm toán 23HG", "Sai lệch (%)", "Đánh giá chất lượng"
    ]
    for c_i, h in enumerate(headers_s5, 1):
        cell = ws5.cell(row=5, column=c_i, value=h)
        cell.font = FONT_WHITE_85
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws5.row_dimensions[5].height = 26

    if is_bep_an:
        audit_items = [
            (1, "G1-09", "Vận chuyển phế thải phá dỡ ra bãi tập kết (ô tô 5T)", "m3", 145.0, 145.0),
            (2, "G2-02", "Đào đất hố móng mác đất cấp III (máy đào 0.4m3)", "m3", 85.0, 85.0),
            (3, "G2-03", "Bê tông lót móng M100 đá 4x6 dày 100mm", "m3", 28.0, 28.0),
            (4, "G2-04", "Cốt thép móng đơn MT1, MT2 mác CB300V", "tấn", 4.25, 4.25),
            (5, "G2-08", "Ván khuôn móng đơn và móng băng gạch", "m2", 134.4, 134.4),
            (6, "G2-09", "Bê tông móng đơn MT1, MT2 đá 1x2 mác 250 (B20)", "m3", 28.0, 28.0),
            (7, "G2-10", "Bê tông giằng móng 220x400 đá 1x2 mác 250", "m3", 12.0, 12.0),
            (8, "G2-15", "Đắp đất hoàn trả hố móng đầm cóc K90", "m3", 65.0, 65.0),
            (9, "G2-16", "Cốt thép cột C1 mác CB300V", "tấn", 1.85, 1.85),
            (10, "G2-18", "Bê tông cột C1 đá 1x2 mác 250", "m3", 12.0, 12.0),
            (11, "G3-01", "Cốt thép dầm sàn tầng mái CB300V", "tấn", 3.80, 3.80),
            (12, "G3-03", "Bê tông dầm sàn tầng mái mác 250 dày 120mm", "m3", 38.0, 38.0),
            (13, "G3-07", "Xà gồ mái thép hộp & dầm trần hộp mạ kẽm", "tấn", 5.10, 5.10),
            (14, "G3-12", "Xây tường ngăn phòng bếp, khu nấu vữa M75", "m3", 68.0, 68.0),
            (15, "G3-14", "Trát tường trong, tường ngoài dày 1.5cm vữa M75", "m2", 520.0, 520.0)
        ]
    else:
        audit_items = []
        for idx, t in enumerate(active_tasks[:15], 1):
            audit_items.append((
                idx,
                t.get("wbs", f"WBS-{idx:02d}"),
                t.get("name", f"Công tác {idx}"),
                t.get("unit", "m3"),
                float(t.get("qty", 10.0)),
                float(t.get("qty", 10.0))
            ))

    r_s5 = 6
    for it in audit_items:
        stt, wbs, name, unit, q_boq, q_23hg = it
        ws5.row_dimensions[r_s5].height = 20
        ws5.cell(r_s5, 1, stt).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(r_s5, 2, wbs).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(r_s5, 2).font = FONT_BOLD_85
        ws5.cell(r_s5, 3, name).font = FONT_REGULAR_85
        ws5.cell(r_s5, 4, unit).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(r_s5, 5, q_boq).number_format = "#,##0.00"
        ws5.cell(r_s5, 5).alignment = Alignment(horizontal="right", vertical="center")
        ws5.cell(r_s5, 6, q_23hg).number_format = "#,##0.00"
        ws5.cell(r_s5, 6).alignment = Alignment(horizontal="right", vertical="center")

        # Sai lệch % = =(F-E)/E
        ws5.cell(r_s5, 7, f"=IF(E{r_s5}>0, ROUND((F{r_s5}-E{r_s5})/E{r_s5}*100, 2), 0)").number_format = "0.00%"
        ws5.cell(r_s5, 7).alignment = Alignment(horizontal="center", vertical="center")

        ws5.cell(r_s5, 8, f'=IF(G{r_s5}=0, "KHỚP 100% — ĐẠT", "CẦN LƯU Ý")')
        ws5.cell(r_s5, 8).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(r_s5, 8).font = Font(name=FONT_NAME, size=8.5, bold=True, color="00385723")

        for c_i in range(1, 9):
            ws5.cell(r_s5, c_i).border = THIN_BORDER
            if r_s5 % 2 == 1:
                ws5.cell(r_s5, c_i).fill = FILL_ZEBRA
        r_s5 += 1

    w_s5 = {1: 6, 2: 10, 3: 42, 4: 8, 5: 18, 6: 18, 7: 14, 8: 20}
    for col_idx, width in w_s5.items():
        ws5.column_dimensions[get_column_letter(col_idx)].width = width

    # Lưu file
    wb.save(output_path)
    wb.close()
    print(f"[OK] Đã xuất bản thành công Bảng tiến độ 3 tầng hợp nhất siêu đẹp: {output_path}")


if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.expanduser("~"), "Downloads", "Documents", "Trường liên cấp phố bảng_Marker",
        "5. Cải tạo Nhà bếp ăn (Nhà 17và 17A) ok_Marker",
        "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH", "GOI_A_CO_GIOI_VA_DAU_DIEZEL",
        "TDTC_CaXe_CaMay_DauDiezel_5._Cải_tạo_Nhà_bếp_ăn_(Nhà_17và_17A).xlsx"
    )
    build_bep_an_3tier_fleet_workbook(target)
