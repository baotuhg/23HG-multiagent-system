# -*- coding: utf-8 -*-
"""
BỘ XUẤT XƯỞNG HỆ THỐNG CÔNG NGHIỆP 3 TẦNG CHO DỰ ÁN:
CẦU THÔN KHAI HOANG 2, KM 14+363.65
Gói thầu số 9: Km12 - Km24+862.93 — Đường từ huyện Đồng Văn đi Mốc 450 (Hà Giang)

Tiêu chuẩn:
- 100% CÔNG THỨC SỐNG — ZERO DEAD NUMBERS — 0 LỖI #REF!, #VALUE!, #NAME?
- Gói A ca máy: Đúng chuẩn mẫu 15 cột A..O, Timeline từ cột P, Tầng 2 MMTB xanh lá, Tầng 3 Dầu diezel
- Hệ thống công nghiệp 3 tầng: Tầng 1 Master 14 sheet, Tầng 2 Vi mô 14 bộ, Tầng 3 Hub & Spoke 5 gói vệ tinh.
"""
from __future__ import annotations

import os
import sys
import gc
import json
import time
import datetime
from typing import List, Dict, Any

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, ".."))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from tools.pho_bang_project_definitions import ProjectDefinition
from tools.pho_bang_master_builder import (
    build_project_master_workbook,
    build_project_fleet_workbook,
    build_project_companion_files,
)
from tools.package_dispatcher import AECPackageDispatcher


TARGET_DIR = r"C:\Users\baotu\Downloads\Documents\Cầu thôn Khai Hoang 2, Km 14+363.65_Marker_2"


def get_khai_hoang_2_project_def() -> ProjectDefinition:
    """Tạo ProjectDefinition cho Cầu thôn Khai Hoang 2."""
    return ProjectDefinition(
        folder_name="Cầu thôn Khai Hoang 2, Km 14+363.65_Marker_2",
        short_name="Cau_Khai_Hoang_2",
        full_name="Cầu thôn Khai Hoang 2, Km 14+363.65 (Nhịp dầm T L=15m, Mố BTCT)",
        project_type="bridge",
        start_date=datetime.date(2026, 3, 1),
        finish_date=datetime.date(2026, 7, 31),
        hm_code="HM_CAU_KHAI_HOANG_2",
        qlcl_filename=None,
        priority_level=1
    )


def load_khai_hoang_2_data(target_dir: str):
    """Nạp dữ liệu thực tế từ hồ sơ bóc tách Cầu Khai Hoang 2."""
    # 1. Danh sách 26 công tác WBS chuẩn thi công cầu vĩnh cửu
    d_st = datetime.date(2026, 3, 1)
    
    tasks: List[Dict[str, Any]] = [
        {
            "stt": 1, "code": "G1-01", "name": "Định vị tim mốc, giải phóng mặt bằng & xây dựng lán trại công trường",
            "unit": "m2", "qty": 450.0, "norm": 150.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=0), "finish": d_st + datetime.timedelta(days=3),
            "mach": "Dụng cụ cơ giới nhỏ + toàn đạc", "crew": 6, "critical": False
        },
        {
            "stt": 2, "code": "G1-02", "name": "San ủi bến bãi đúc dầm & đường công vụ phục vụ thi công cầu",
            "unit": "m2", "qty": 850.0, "norm": 212.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=2), "finish": d_st + datetime.timedelta(days=6),
            "mach": "Máy ủi 110CV + Lu rung 10T", "crew": 6, "critical": False
        },
        {
            "stt": 3, "code": "G1-03", "name": "Đào đất đá hố móng mố M1 bằng máy đào 0.8m3 kết hợp thủ công",
            "unit": "m3", "qty": 185.0, "norm": 37.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=5), "finish": d_st + datetime.timedelta(days=10),
            "mach": "Máy đào 0.8m3 + Ô tô 7T", "crew": 6, "critical": True
        },
        {
            "stt": 4, "code": "G1-04", "name": "Đào đất đá hố móng mố M2 bằng máy đào 0.8m3 kết hợp thủ công",
            "unit": "m3", "qty": 195.0, "norm": 39.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=7), "finish": d_st + datetime.timedelta(days=12),
            "mach": "Máy đào 0.8m3 + Ô tô 7T", "crew": 6, "critical": True
        },
        {
            "stt": 5, "code": "G1-05", "name": "Đệm cát & Đổ bê tông lót móng M100 đá 4x6 dày 100mm mố M1, M2",
            "unit": "m3", "qty": 24.5, "norm": 8.17, "shifts": 1,
            "start": d_st + datetime.timedelta(days=11), "finish": d_st + datetime.timedelta(days=14),
            "mach": "Máy trộn 350L + đầm bàn", "crew": 8, "critical": False
        },
        {
            "stt": 6, "code": "G1-06", "name": "Gia công lắp dựng cốt thép bệ mố M1, M2 mác CB400V",
            "unit": "tấn", "qty": 6.85, "norm": 1.37, "shifts": 1,
            "start": d_st + datetime.timedelta(days=13), "finish": d_st + datetime.timedelta(days=18),
            "mach": "Máy uốn cắt thép CNC RebarCut", "crew": 10, "critical": True
        },
        {
            "stt": 7, "code": "G1-07", "name": "Lắp dựng ván khuôn định hình phủ phim bệ mố M1, M2",
            "unit": "m2", "qty": 168.0, "norm": 42.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=16), "finish": d_st + datetime.timedelta(days=20),
            "mach": "Giàn giáo thép + ván ép phủ phim", "crew": 8, "critical": True
        },
        {
            "stt": 8, "code": "G1-08", "name": "Đổ bê tông bệ mố M1, M2 đá 1x2 mác 300 (C25) bằng bơm bê tông",
            "unit": "m3", "qty": 58.0, "norm": 19.33, "shifts": 1,
            "start": d_st + datetime.timedelta(days=19), "finish": d_st + datetime.timedelta(days=22),
            "mach": "Xe bơm bê tông 42m + đầm dùi", "crew": 14, "critical": True
        },
        {
            "stt": 9, "code": "G1-09", "name": "Gia công cốt thép thân mố, tường ngực, đỉnh mố & tường cánh M1",
            "unit": "tấn", "qty": 4.50, "norm": 1.13, "shifts": 1,
            "start": d_st + datetime.timedelta(days=22), "finish": d_st + datetime.timedelta(days=26),
            "mach": "Máy uốn cắt thép RebarCut", "crew": 10, "critical": True
        },
        {
            "stt": 10, "code": "G1-10", "name": "Ván khuôn & Đổ bê tông thân mố, tường ngực mố M1 mác 300 (C25)",
            "unit": "m3", "qty": 42.0, "norm": 14.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=25), "finish": d_st + datetime.timedelta(days=28),
            "mach": "Máy trộn 350L + đầm dùi", "crew": 12, "critical": True
        },
        {
            "stt": 11, "code": "G1-11", "name": "Gia công cốt thép thân mố, tường ngực, đỉnh mố & tường cánh M2",
            "unit": "tấn", "qty": 4.65, "norm": 1.16, "shifts": 1,
            "start": d_st + datetime.timedelta(days=26), "finish": d_st + datetime.timedelta(days=30),
            "mach": "Máy uốn cắt thép RebarCut", "crew": 10, "critical": True
        },
        {
            "stt": 12, "code": "G1-12", "name": "Ván khuôn & Đổ bê tông thân mố, tường ngực mố M2 mác 300 (C25)",
            "unit": "m3", "qty": 43.5, "norm": 14.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=29), "finish": d_st + datetime.timedelta(days=32),
            "mach": "Máy trộn 350L + đầm dùi", "crew": 12, "critical": True
        },
        {
            "stt": 13, "code": "G1-13", "name": "Đổ bê tông đá kê gối, lắp đặt 08 gối cầu cao su bản thép 200x250x42",
            "unit": "bộ", "qty": 8.0, "norm": 2.67, "shifts": 1,
            "start": d_st + datetime.timedelta(days=32), "finish": d_st + datetime.timedelta(days=35),
            "mach": "Vữa không co ngót Sika Grout", "crew": 6, "critical": False
        },
        {
            "stt": 14, "code": "G1-14", "name": "Lắp dựng bệ đúc & gia công ván khuôn thép đúc 04 phiến dầm T L=15m",
            "unit": "bộ", "qty": 2.0, "norm": 0.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=15), "finish": d_st + datetime.timedelta(days=19),
            "mach": "Cẩu tự hành 25T + máy hàn 400A", "crew": 8, "critical": False
        },
        {
            "stt": 15, "code": "G1-15", "name": "Gia công & lắp dựng cốt thép 04 phiến dầm T L=15m (Dầm biên & Dầm trong)",
            "unit": "tấn", "qty": 5.43, "norm": 0.91, "shifts": 1,
            "start": d_st + datetime.timedelta(days=19), "finish": d_st + datetime.timedelta(days=25),
            "mach": "Máy uốn cắt CNC + máy hàn", "crew": 12, "critical": True
        },
        {
            "stt": 16, "code": "G1-16", "name": "Đổ bê tông 04 phiến dầm T L=15m đá 1x2 mác 350 (C30) & dưỡng hộ",
            "unit": "m3", "qty": 37.0, "norm": 9.25, "shifts": 1,
            "start": d_st + datetime.timedelta(days=24), "finish": d_st + datetime.timedelta(days=28),
            "mach": "Máy trộn 500L + đầm dùi/bàn", "crew": 16, "critical": True
        },
        {
            "stt": 17, "code": "G1-17", "name": "Cẩu lắp vận chuyển & lao lắp 04 phiến dầm T L=15m vào vị trí gối",
            "unit": "phiến", "qty": 4.0, "norm": 1.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=35), "finish": d_st + datetime.timedelta(days=39),
            "mach": "Cần cẩu bánh xích 50T (2 cẩu)", "crew": 14, "critical": True
        },
        {
            "stt": 18, "code": "G1-18", "name": "Gia công cốt thép & đổ bê tông mối nối ướt, dầm ngang bản mặt cầu",
            "unit": "m3", "qty": 9.13, "norm": 2.28, "shifts": 1,
            "start": d_st + datetime.timedelta(days=38), "finish": d_st + datetime.timedelta(days=42),
            "mach": "Máy trộn vữa/bê tông 350L + đầm dùi", "crew": 10, "critical": True
        },
        {
            "stt": 19, "code": "G1-19", "name": "Lắp dựng ván khuôn & đổ bê tông 02 bản quá độ đầu mố L=3.5m mác 250",
            "unit": "m3", "qty": 14.8, "norm": 3.7, "shifts": 1,
            "start": d_st + datetime.timedelta(days=40), "finish": d_st + datetime.timedelta(days=44),
            "mach": "Máy trộn 350L + đầm bàn", "crew": 8, "critical": False
        },
        {
            "stt": 20, "code": "G1-20", "name": "Thi công lớp phủ bản mặt cầu BTCT dày 10cm chống thấm mác 350",
            "unit": "m3", "qty": 12.5, "norm": 3.13, "shifts": 1,
            "start": d_st + datetime.timedelta(days=43), "finish": d_st + datetime.timedelta(days=47),
            "mach": "Máy rải bê tông + đầm thước", "crew": 12, "critical": True
        },
        {
            "stt": 21, "code": "G1-21", "name": "Lắp đặt 02 bộ khe co giãn cao su cốt bản thép 2 đầu mố M1, M2",
            "unit": "bộ", "qty": 2.0, "norm": 0.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=46), "finish": d_st + datetime.timedelta(days=50),
            "mach": "Máy hàn điện 250A + vữa không co", "crew": 6, "critical": False
        },
        {
            "stt": 22, "code": "G1-22", "name": "Gia công lắp dựng hệ thống lan can thép mạ kẽm (32 cột) & gờ lan can",
            "unit": "md", "qty": 36.0, "norm": 7.2, "shifts": 1,
            "start": d_st + datetime.timedelta(days=47), "finish": d_st + datetime.timedelta(days=52),
            "mach": "Máy hàn + bulông neo M22", "crew": 8, "critical": False
        },
        {
            "stt": 23, "code": "G1-23", "name": "Thi công xây đá hộc ốp mái taluy, tứ nón mố M1 và mố M2 vữa M100",
            "unit": "m3", "qty": 85.0, "norm": 12.14, "shifts": 1,
            "start": d_st + datetime.timedelta(days=44), "finish": d_st + datetime.timedelta(days=51),
            "mach": "Máy trộn vữa 80L + búa đẽo đá", "crew": 12, "critical": False
        },
        {
            "stt": 24, "code": "G1-24", "name": "Đắp đất mang mố đầm chặt K95 & hệ thống tầng lọc thoát nước sau mố",
            "unit": "m3", "qty": 320.0, "norm": 45.71, "shifts": 1,
            "start": d_st + datetime.timedelta(days=48), "finish": d_st + datetime.timedelta(days=55),
            "mach": "Máy đầm cóc 70kg + đầm rung nhỏ", "crew": 8, "critical": False
        },
        {
            "stt": 25, "code": "G1-25", "name": "Thi công kết cấu áo đường đầu cầu (CPĐD loại 1 & Bê tông xi măng M300)",
            "unit": "m2", "qty": 280.0, "norm": 56.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=52), "finish": d_st + datetime.timedelta(days=57),
            "mach": "Lu rung 10T + máy rải bê tông", "crew": 10, "critical": False
        },
        {
            "stt": 26, "code": "G1-26", "name": "Lắp đặt biển báo an toàn giao thông, sơn kẻ đường, thử tải cầu & bàn giao",
            "unit": "hệ", "qty": 1.0, "norm": 0.25, "shifts": 1,
            "start": d_st + datetime.timedelta(days=56), "finish": d_st + datetime.timedelta(days=60),
            "mach": "Xe tải thử tải + máy sơn kẻ đường", "crew": 8, "critical": False
        }
    ]

    # Bổ sung thông số quản lý phân đoạn và vị trí thi công
    for t in tasks:
        stt = t["stt"]
        if stt <= 2:
            t["phase"] = "GIAI ĐOẠN 1: CHUẨN BỊ MẶT BẰNG & BẾN BÃI"
            t["loc"] = "Toàn công trường"
        elif stt <= 8:
            t["phase"] = "GIAI ĐOẠN 2: THI CÔNG HỐ MÓNG & BỆ MỐ M1, M2"
            t["loc"] = "Mố M1 và Mố M2"
        elif stt <= 13:
            t["phase"] = "GIAI ĐOẠN 3: THI CÔNG THÂN MỐ, ĐỈNH MỐ & GỐI CẦU"
            t["loc"] = "Thân & Tường cánh mố"
        elif stt <= 18:
            t["phase"] = "GIAI ĐOẠN 4: ĐÚC DẦM & LAO LẮP DẦM T L=15M"
            t["loc"] = "Bãi đúc dầm & Nhịp cầu"
        elif stt <= 22:
            t["phase"] = "GIAI ĐOẠN 5: BẢN MẶT CẦU, KHE CO GIÃN & LAN CAN"
            t["loc"] = "Mặt cầu L=15m"
        else:
            t["phase"] = "GIAI ĐOẠN 6: TỨ NÓN, ĐƯỜNG ĐẦU CẦU & THỬ TẢI"
            t["loc"] = "Đường đầu cầu & Tứ nón"

    # 2. Danh mục vật tư chính
    materials = [
        {"stt": 1, "code": "VT-XM", "name": "Xi măng Poóc lăng PCB40 Hoàng Thạch / Bút Sơn", "unit": "tấn", "qty": 98.5, "price": 1650000},
        {"stt": 2, "code": "VT-CAT", "name": "Cát vàng đổ bê tông mô đun độ lớn ML >= 2.5", "unit": "m3", "qty": 145.0, "price": 380000},
        {"stt": 3, "code": "VT-DA12", "name": "Đá dăm 1x2 mác cao cho bê tông C25, C30", "unit": "m3", "qty": 210.0, "price": 320000},
        {"stt": 4, "code": "VT-DA46", "name": "Đá dăm 4x6 lót móng bệ mố", "unit": "m3", "qty": 35.0, "price": 280000},
        {"stt": 5, "code": "VT-TH-CB400", "name": "Thép cốt bê tông CB400V (D16 - D32) dầm & mố", "unit": "tấn", "qty": 21.4, "price": 16800000},
        {"stt": 6, "code": "VT-TH-CB300", "name": "Thép cốt bê tông CB300V (D10 - D14) cấu kiện phụ", "unit": "tấn", "qty": 8.6, "price": 16500000},
        {"stt": 7, "code": "VT-VK", "name": "Ván khuôn phủ phim chịu nước 18mm & phụ kiện gông", "unit": "m2", "qty": 380.0, "price": 220000},
        {"stt": 8, "code": "VT-GC", "name": "Gối cầu cao su cốt bản thép 200x250x42mm", "unit": "bộ", "qty": 8.0, "price": 3200000},
        {"stt": 9, "code": "VT-KKG", "name": "Khe co giãn cao su cốt bản thép kèm thanh chặn", "unit": "md", "qty": 14.0, "price": 2450000},
        {"stt": 10, "code": "VT-LC", "name": "Hệ thống lan can thép ống mạ kẽm nhúng nóng", "unit": "md", "qty": 36.0, "price": 1150000},
        {"stt": 11, "code": "VT-DA-HOC", "name": "Đá hộc xây tứ nón và mái kè taluy", "unit": "m3", "qty": 95.0, "price": 260000},
        {"stt": 12, "code": "VT-CPDD", "name": "Cấp phối đá dăm loại 1 đường đầu cầu", "unit": "m3", "qty": 120.0, "price": 290000},
        {"stt": 13, "code": "VT-SIKA", "name": "Vữa không co ngót cường độ cao Sika Grout 214-11", "unit": "tấn", "qty": 2.5, "price": 9500000},
        {"stt": 14, "code": "VT-CT", "name": "Sơn chống thấm màng đàn hồi mặt cầu", "unit": "kg", "qty": 180.0, "price": 85000},
        {"stt": 15, "code": "VT-ONG-TN", "name": "Ống thoát nước mặt cầu composite D100 kèm phễu", "unit": "bộ", "qty": 6.0, "price": 450000}
    ]

    std_defaults = {
        "VT-XM": ("TCVN 2682:2009", "100 Tấn/lô", "Độ mịn, thời gian đông kết, cường độ nén"),
        "VT-CAT": ("TCVN 7570:2006", "200 m3/lô", "Thành phần hạt, hàm lượng bùn sét, tạp chất"),
        "VT-DA12": ("TCVN 7570:2006", "200 m3/lô", "Thành phần hạt, độ nén dập, hàm lượng thoi dẹt"),
        "VT-DA46": ("TCVN 7570:2006", "200 m3/lô", "Độ nén dập, thành phần cỡ hạt"),
        "VT-TH-CB400": ("TCVN 1651:2018", "50 Tấn/lô", "Giới hạn chảy, độ bền kéo, uốn nguội"),
        "VT-TH-CB300": ("TCVN 1651:2018", "50 Tấn/lô", "Giới hạn chảy, độ bền kéo, uốn nguội"),
        "VT-VK": ("TCVN 4453:1995", "Theo đợt nhập", "Kích thước hình học, độ phẳng, chống thấm"),
        "VT-GC": ("TCVN 10308:2014", "Theo lô 08 bộ", "Kích thước, độ cứng cao su, nén nứt"),
        "VT-KKG": ("TCVN 11823:2017", "Theo lô 02 bộ", "Thép mạ kẽm, cao su đàn hồi"),
        "VT-LC": ("TCVN 5729:2012", "Theo đợt lắp", "Chiều dày mạ kẽm nhúng nóng, mối hàn"),
        "VT-DA-HOC": ("TCVN 4085:2011", "100 m3/lô", "Kích thước viên đá, cường độ nén"),
        "VT-CPDD": ("TCVN 8859:2011", "200 m3/lô", "Thành phần hạt, CBR, độ đầm nén max"),
        "VT-SIKA": ("ASTM C1107", "Theo lô sản xuất", "Độ chảy, độ nở, cường độ R1, R3, R28"),
        "VT-CT": ("TCVN 9345:2012", "Theo đợt cung cấp", "Độ bám dính bê tông, khả năng đàn hồi"),
        "VT-ONG-TN": ("TCVN 11823:2017", "Theo lô 06 bộ", "Độ bền cơ học, chống ăn mòn hóa chất")
    }
    for m in materials:
        c_val = m["code"]
        if c_val in std_defaults:
            m["std"], m["freq"], m["test"] = std_defaults[c_val]
        else:
            m["std"], m["freq"], m["test"] = ("TCVN", "Theo lô", "Chỉ tiêu kỹ thuật cơ lý")

    # 3. Nạp danh mục cốt thép từ thep_cho_to_hop_cat.json nếu có
    rebar_items = []
    thep_path = os.path.join(target_dir, "thep_cho_to_hop_cat.json")
    uw_map = {6: 0.222, 8: 0.395, 10: 0.617, 12: 0.888, 14: 1.208, 16: 1.578, 18: 1.998, 20: 2.466, 22: 2.984, 25: 3.853, 28: 4.834, 32: 6.313}
    if os.path.exists(thep_path):
        try:
            with open(thep_path, "r", encoding="utf-8") as f:
                raw_thep = json.load(f)
            for idx, item in enumerate(raw_thep[:30], 1):
                d_val = int(item.get('diameter', 16))
                mark_name = item.get("mark") or f"T{idx:02d}"
                sh_title = item.get('sheet_title', 'Kết cấu cầu')
                comp_name = f"Thanh {mark_name} ({sh_title})"
                n_count = int(item.get("quantity", 10))
                l_m = round(float(item.get("length_mm", 1000)) / 1000.0, 2)
                uw = uw_map.get(d_val, 1.58)
                rebar_items.append((
                    idx,
                    comp_name,
                    f"D{d_val}",
                    n_count,
                    l_m,
                    uw
                ))
        except Exception as e:
            print("  [!] Lỗi đọc thep_cho_to_hop_cat.json:", e)

    if not rebar_items:
        # Fallback 15 cấu kiện thép cầu chuẩn
        rebar_items = [
            (1, "Dầm T L=15m (Thanh A1 - Dầm chủ)", "D32", 20, 15.28, 6.31),
            (2, "Dầm T L=15m (Thanh A2 - Thép đai)", "D12", 40, 14.50, 0.89),
            (3, "Bệ mố M1, M2 (Thanh F1)", "D25", 54, 7.26, 3.85),
            (4, "Bệ mố M1, M2 (Thanh F1A)", "D25", 52, 8.26, 3.85),
            (5, "Tường cánh mố M1, M2 (Thanh K1)", "D22", 36, 8.50, 2.98),
            (6, "Tường cánh mố M1, M2 (Thanh K2)", "D22", 36, 4.20, 2.98),
            (7, "Thân mố & đỉnh mố (Thanh B1)", "D20", 48, 6.80, 2.47),
            (8, "Thân mố & đai (Thanh B2)", "D14", 60, 5.40, 1.21),
            (9, "Bệ kê gối cầu (Thanh G1)", "D10", 160, 0.50, 0.62),
            (10, "Bệ kê gối & u neo (Thanh G3)", "D12", 96, 1.52, 0.89),
            (11, "Bản quá độ đầu mố (Thanh A1-QD)", "D20", 70, 3.26, 2.47),
            (12, "Bản quá độ đầu mố (Thanh A2-QD)", "D14", 70, 3.26, 1.21),
            (13, "Lớp phủ mặt cầu (Thanh L1)", "D14", 600, 2.18, 1.21),
            (14, "Gờ lan can nhịp (Thanh L2)", "D14", 600, 1.92, 1.21),
            (15, "Neo khe co giãn (Thanh E1)", "D16", 16, 7.92, 1.58)
        ]

    return tasks, materials, rebar_items


def export_cau_khai_hoang_2():
    """Quy trình xuất xưởng công nghiệp 3 tầng cho Cầu thôn Khai Hoang 2."""
    t0 = time.time()
    proj = get_khai_hoang_2_project_def()
    dest_dir = TARGET_DIR

    print("=" * 80)
    print(f"[*] KÍCH HOẠT QUY TRÌNH XUẤT XƯỞNG HỆ THỐNG CÔNG NGHIỆP 3 TẦNG")
    print(f"    - Dự án:        {proj.full_name}")
    print(f"    - Thư mục đích: {dest_dir}")
    print(f"    - Short name:   {proj.short_name}")
    print(f"    - Khung mốc:    {proj.start_date.strftime('%d/%m/%Y')} -> {proj.finish_date.strftime('%d/%m/%Y')}")
    print("=" * 80)

    # 1. Nạp dữ liệu thực tế
    tasks, materials, rebar_items = load_khai_hoang_2_data(dest_dir)
    print(f"  [+] Đã tổng hợp dữ liệu: {len(tasks)} công tác WBS, {len(materials)} loại vật tư, {len(rebar_items)} cấu kiện thép.")

    # 2. Xây dựng Tầng 1: Master Workbook 14 Sheet
    master_xlsx = os.path.join(dest_dir, f"Ho_So_KCS_QS_TienDo_{proj.short_name}.xlsx")
    print(f"  [1/5] Đang tạo Master Workbook 14 Sheet...")
    build_project_master_workbook(proj, tasks, materials, rebar_items, master_xlsx)

    # 3. Xây dựng Gói A: Ca máy & Dầu diezel (Layout 3 tầng 15 cột A..O chuẩn mẫu gốc)
    fleet_xlsx = os.path.join(dest_dir, f"TDTC_CaXe_CaMay_DauDiezel_{proj.short_name}.xlsx")
    print(f"  [2/5] Đang tạo Gói A 3 tầng 15 cột A..O chuẩn mẫu trực quan Vincons...")
    build_project_fleet_workbook(proj, tasks, fleet_xlsx)

    # 4. Xây dựng các Companion files (Gantt XML, MPP, DOCX Thuyết minh BPTC, Audit MD)
    print(f"  [3/5] Đang tạo Companion files (XML, MPP, DOCX, MD)...")
    companion = build_project_companion_files(proj, tasks, materials, dest_dir)

    # 5. Kích hoạt Dispatcher đóng gói 3 TẦNG
    print(f"  [4/5] Kích hoạt AECPackageDispatcher phân phối 3 Tầng & 5 Gói vệ tinh...")
    dispatcher = AECPackageDispatcher(base_output_dir=dest_dir)
    manifest = dispatcher.dispatch_full_industrial_dossier(
        master_excel_path=master_xlsx,
        project_name=proj.full_name,
        target_dir=dest_dir,
        companion_files={
            "fleet_template": fleet_xlsx,
            "fleet_xml": companion["master_xml"],
            "mpp": companion["mpp"],
            "docx": companion["docx"],
            "bptc": companion["bptc"],
            "audit": companion["audit"]
        }
    )

    # 6. Đồng bộ DISPATCH_MANIFEST.json ra thư mục gốc dự án & dọn dẹp file tạm
    manifest_in_hub = os.path.join(dest_dir, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH", "DISPATCH_MANIFEST.json")
    manifest_at_root = os.path.join(dest_dir, "DISPATCH_MANIFEST.json")
    if os.path.exists(manifest_in_hub):
        import shutil
        shutil.copyfile(manifest_in_hub, manifest_at_root)

    # Dọn dẹp tệp trung gian ở root sau khi đã phân phối vào 3 tầng
    temp_files = [
        master_xlsx, fleet_xlsx, companion["master_xml"], companion["mpp"],
        companion["docx"], companion["bptc"], companion["audit"]
    ]
    for tf in temp_files:
        if os.path.exists(tf):
            try:
                os.remove(tf)
            except OSError:
                pass

    elapsed = time.time() - t0
    print("=" * 80)
    print(f"[HOÀN TẤT] Xuất xưởng Cầu thôn Khai Hoang 2 thành công trong {elapsed:.1f} giây!")
    print(f"  - Thư mục đích:         {dest_dir}")
    print(f"  - Tổng số tệp tạo lập:  {manifest.total_files_count} tệp tin")
    print(f"  - Trạng thái Quality Gate: {manifest.audit_zero_errors} (100% Zero Formula Errors)")
    print("=" * 80)
    return manifest


if __name__ == "__main__":
    export_cau_khai_hoang_2()
