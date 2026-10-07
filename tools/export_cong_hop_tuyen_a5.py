# -*- coding: utf-8 -*-
"""
Script xuất hồ sơ công nghiệp 3 tầng cho dự án [Cống Hộp Tuyến A5]
26 công tác chi tiết chuyên biệt cho công trình cống hộp thoát nước & tấm giảm tải.
Chuẩn Vincons / 23HG Multi-Agent System - ZERO FORMULA ERRORS.
"""

import os
import sys
import datetime
from typing import List, Dict, Any

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, ".."))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from tools.pho_bang_project_definitions import ProjectDefinition
from tools.pho_bang_master_builder import (
    build_project_master_workbook,
    build_project_fleet_workbook,
    build_project_companion_files,
)
from tools.package_dispatcher import AECPackageDispatcher

def get_project_def() -> ProjectDefinition:
    return ProjectDefinition(
        folder_name="[Cống Hộp Tuyến A5]",
        short_name="Cong_Hop_A5",
        full_name="Cống Hộp Tuyến A5 - Hệ thống thoát nước",
        project_type="infra_water",
        start_date=datetime.date(2026, 10, 1),
        finish_date=datetime.date(2026, 11, 30),
        hm_code="HM_CH_A5",
        qlcl_filename="BBNT_Ho_So_QLCL_[Cong_Hop_Tuyen_A5].docx",
        priority_level=1
    )

def export_cong_hop():
    proj = get_project_def()
    target_dir = os.environ.get("AEC_PROJECTS_DIR", os.path.join(os.path.expanduser("~"), "Downloads", "Documents", "2026.09.12.OLP_SD_HT_ChiTietCongHop+TamGiamTai Model (1)_Marker"))
    
    d_st = proj.start_date
    
    # 26 CÔNG TÁC CHUẨN THI CÔNG CỐNG HỘP TUYẾN A5 (Không lặp, phân đoạn thực tế)
    tasks: List[Dict[str, Any]] = [
        {
            "stt": 1, "code": "CH-01", "name": "Định vị tim mốc, ranh giới tuyến cống hộp A5 & hệ thống cọc mốc dẫn",
            "unit": "m2", "qty": 850.0, "norm": 283.33, "shifts": 1,
            "start": d_st, "finish": d_st + datetime.timedelta(days=2),
            "loc": "Km0+00 - Km0+350", "mach": "Máy toàn đạc điện tử + Máy thủy bình", "crew": 4, "critical": False
        },
        {
            "stt": 2, "code": "CH-02", "name": "Dọn dẹp phát quang mặt bằng, đắp bờ vây chống tràn & rãnh tiêu thoát nước",
            "unit": "m2", "qty": 1200.0, "norm": 300.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=1), "finish": d_st + datetime.timedelta(days=4),
            "loc": "Km0+00 - Km0+350", "mach": "Máy đào 0.8m3", "crew": 6, "critical": False
        },
        {
            "stt": 3, "code": "CH-03", "name": "Đào đất hố móng cống hộp phân đoạn 1 (Đoạn cống đơn 2.0x2.0m)",
            "unit": "m3", "qty": 650.0, "norm": 130.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=3), "finish": d_st + datetime.timedelta(days=7),
            "loc": "Đoạn 2.0x2.0m", "mach": "Máy đào 0.8m3 + Ô tô 7T", "crew": 6, "critical": True
        },
        {
            "stt": 4, "code": "CH-04", "name": "Đào đất hố móng cống hộp phân đoạn 2 (Đoạn cống đơn 3.0x3.0m)",
            "unit": "m3", "qty": 780.0, "norm": 156.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=6), "finish": d_st + datetime.timedelta(days=10),
            "loc": "Đoạn 3.0x3.0m", "mach": "Máy đào 0.8m3 + Ô tô 7T", "crew": 6, "critical": True
        },
        {
            "stt": 5, "code": "CH-05", "name": "Đào đất hố móng cống hộp phân đoạn 3 (Đoạn cống đôi 2x(3.0x3.0m))",
            "unit": "m3", "qty": 920.0, "norm": 153.33, "shifts": 1,
            "start": d_st + datetime.timedelta(days=9), "finish": d_st + datetime.timedelta(days=14),
            "loc": "Đoạn cống đôi", "mach": "Máy đào 0.8m3 + Ô tô 7T", "crew": 6, "critical": True
        },
        {
            "stt": 6, "code": "CH-06", "name": "Vận chuyển đất đào không thích hợp đổ đi bằng ô tô tự đổ 7T (Cự ly 5km)",
            "unit": "m3", "qty": 1850.0, "norm": 154.17, "shifts": 1,
            "start": d_st + datetime.timedelta(days=4), "finish": d_st + datetime.timedelta(days=15),
            "loc": "Bãi đổ quy định", "mach": "Ô tô tự đổ 7T (2 xe)", "crew": 4, "critical": False
        },
        {
            "stt": 7, "code": "CH-07", "name": "Bơm hạ mực nước ngầm & vét dọn bùn đáy hố móng cống hộp",
            "unit": "ca", "qty": 25.0, "norm": 1.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=3), "finish": d_st + datetime.timedelta(days=27),
            "loc": "Dọc tuyến cống", "mach": "Máy bơm chìm D80 - D100", "crew": 4, "critical": False
        },
        {
            "stt": 8, "code": "CH-08", "name": "Trải vải địa kỹ thuật & đệm cát hạt trung đầm chặt đáy móng cống dày 20cm",
            "unit": "m3", "qty": 185.0, "norm": 37.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=8), "finish": d_st + datetime.timedelta(days=12),
            "loc": "Đáy móng cống", "mach": "Máy đầm cóc 70kg + thủ công", "crew": 8, "critical": False
        },
        {
            "stt": 9, "code": "CH-09", "name": "Đổ bê tông lót móng cống M100 đá 4x6 dày 100mm (Đoạn cống 2.0x2.0m)",
            "unit": "m3", "qty": 28.5, "norm": 9.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=11), "finish": d_st + datetime.timedelta(days=13),
            "loc": "Đoạn 2.0x2.0m", "mach": "Máy trộn bê tông 350L", "crew": 8, "critical": False
        },
        {
            "stt": 10, "code": "CH-10", "name": "Đổ bê tông lót móng cống M100 đá 4x6 dày 100mm (Đoạn cống 3x3m & cống đôi)",
            "unit": "m3", "qty": 56.0, "norm": 14.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=13), "finish": d_st + datetime.timedelta(days=16),
            "loc": "Đoạn 3x3m & cống đôi", "mach": "Máy trộn bê tông 350L", "crew": 8, "critical": False
        },
        {
            "stt": 11, "code": "CH-11", "name": "Gia công cốt thép cống hộp các loại mác CB500-V tại xưởng bãi tiền chế",
            "unit": "tấn", "qty": 38.6, "norm": 2.76, "shifts": 1,
            "start": d_st + datetime.timedelta(days=10), "finish": d_st + datetime.timedelta(days=23),
            "loc": "Xưởng thép công trường", "mach": "Máy uốn cắt thép + Máy hàn 250A", "crew": 12, "critical": True
        },
        {
            "stt": 12, "code": "CH-12", "name": "Lắp dựng cốt thép bản đáy & chân thành cống hộp mác CB500-V",
            "unit": "tấn", "qty": 18.2, "norm": 3.03, "shifts": 1,
            "start": d_st + datetime.timedelta(days=15), "finish": d_st + datetime.timedelta(days=20),
            "loc": "Toàn tuyến", "mach": "Máy uốn cắt thép + thủ công", "crew": 10, "critical": True
        },
        {
            "stt": 13, "code": "CH-13", "name": "Lắp dựng ván khuôn định hình bản đáy & chân thành cống hộp",
            "unit": "m2", "qty": 210.0, "norm": 42.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=17), "finish": d_st + datetime.timedelta(days=21),
            "loc": "Toàn tuyến", "mach": "Ván khuôn thép + máy hàn 250A", "crew": 8, "critical": True
        },
        {
            "stt": 14, "code": "CH-14", "name": "Đổ bê tông bản đáy cống hộp M250 đá 1x2 bằng máy bơm bê tông (Đợt 1)",
            "unit": "m3", "qty": 145.0, "norm": 36.25, "shifts": 1,
            "start": d_st + datetime.timedelta(days=21), "finish": d_st + datetime.timedelta(days=24),
            "loc": "Toàn tuyến", "mach": "Máy bơm bê tông + Máy đầm dùi", "crew": 14, "critical": True
        },
        {
            "stt": 15, "code": "CH-15", "name": "Lắp đặt băng cản nước Waterstop PVC V200 tại mạch ngừng chân thành cống",
            "unit": "m", "qty": 360.0, "norm": 120.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=23), "finish": d_st + datetime.timedelta(days=25),
            "loc": "Mạch ngừng thi công", "mach": "Dao hàn nhiệt + thủ công", "crew": 4, "critical": False
        },
        {
            "stt": 16, "code": "CH-16", "name": "Lắp dựng cốt thép thành và bản nắp cống hộp mác CB500-V",
            "unit": "tấn", "qty": 20.4, "norm": 3.4, "shifts": 1,
            "start": d_st + datetime.timedelta(days=25), "finish": d_st + datetime.timedelta(days=30),
            "loc": "Toàn tuyến", "mach": "Máy uốn cắt thép + máy hàn", "crew": 12, "critical": True
        },
        {
            "stt": 17, "code": "CH-17", "name": "Lắp dựng ván khuôn trong ngoài thành cống & đáy bản nắp cống hộp",
            "unit": "m2", "qty": 680.0, "norm": 113.33, "shifts": 1,
            "start": d_st + datetime.timedelta(days=27), "finish": d_st + datetime.timedelta(days=32),
            "loc": "Toàn tuyến", "mach": "Cần cẩu 5T + Giàn giáo định hình", "crew": 14, "critical": True
        },
        {
            "stt": 18, "code": "CH-18", "name": "Đổ bê tông thành và bản nắp cống hộp M250 đá 1x2 bằng máy bơm (Đợt 2)",
            "unit": "m3", "qty": 210.0, "norm": 52.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=32), "finish": d_st + datetime.timedelta(days=35),
            "loc": "Toàn tuyến", "mach": "Máy bơm bê tông + Máy đầm dùi", "crew": 15, "critical": True
        },
        {
            "stt": 19, "code": "CH-19", "name": "Bảo dưỡng bê tông cống bằng bao tải giữ ẩm và tưới nước liên tục 7 ngày",
            "unit": "m2", "qty": 750.0, "norm": 107.14, "shifts": 1,
            "start": d_st + datetime.timedelta(days=34), "finish": d_st + datetime.timedelta(days=40),
            "loc": "Toàn tuyến", "mach": "Máy bơm nước tưới ẩm", "crew": 4, "critical": False
        },
        {
            "stt": 20, "code": "CH-20", "name": "Tháo dỡ hệ thống ván khuôn trong lòng cống và ván khuôn thành ngoài",
            "unit": "m2", "qty": 680.0, "norm": 136.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=39), "finish": d_st + datetime.timedelta(days=43),
            "loc": "Toàn tuyến", "mach": "Thủ công + tời kéo", "crew": 10, "critical": False
        },
        {
            "stt": 21, "code": "CH-21", "name": "Xử lý khe co giãn cống hộp (chèn xốp cao su, matit bitum chèn khe)",
            "unit": "m", "qty": 180.0, "norm": 36.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=42), "finish": d_st + datetime.timedelta(days=46),
            "loc": "Vị trí khe co giãn", "mach": "Nồi nấu nhựa + thủ công", "crew": 6, "critical": False
        },
        {
            "stt": 22, "code": "CH-22", "name": "Thi công quét sơn chống thấm bitum mặt ngoài thành và bản nắp cống hộp",
            "unit": "m2", "qty": 850.0, "norm": 170.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=44), "finish": d_st + datetime.timedelta(days=48),
            "loc": "Thành ngoài cống", "mach": "Thủ công lăn sơn", "crew": 6, "critical": False
        },
        {
            "stt": 23, "code": "CH-23", "name": "Thi công hố ga thu thăm cống hộp (đào móng, BT đáy, xây thành ga, trát ga)",
            "unit": "hố", "qty": 18.0, "norm": 2.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=36), "finish": d_st + datetime.timedelta(days=45),
            "loc": "Các vị trí ga", "mach": "Máy trộn bê tông 350L + Máy trộn 80L", "crew": 12, "critical": False
        },
        {
            "stt": 24, "code": "CH-24", "name": "Cẩu lắp tấm đan bê tông cốt thép giảm tải và nắp ga gang thoát nước D400",
            "unit": "bộ", "qty": 18.0, "norm": 4.5, "shifts": 1,
            "start": d_st + datetime.timedelta(days=45), "finish": d_st + datetime.timedelta(days=48),
            "loc": "Đỉnh các hố ga", "mach": "Cần cẩu tự hành 5T", "crew": 6, "critical": False
        },
        {
            "stt": 25, "code": "CH-25", "name": "Đắp cát hạt trung đầm chặt K95 hai bên hông thành cống hộp dày lớp 20cm",
            "unit": "m3", "qty": 420.0, "norm": 70.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=47), "finish": d_st + datetime.timedelta(days=52),
            "loc": "Hai bên hông cống", "mach": "Máy đầm cóc 70kg + Ô tô 7T", "crew": 8, "critical": True
        },
        {
            "stt": 26, "code": "CH-26", "name": "Đắp đất hoàn trả hố móng đầm chặt K90/K95 bằng máy lu rung và máy đầm cóc",
            "unit": "m3", "qty": 980.0, "norm": 140.0, "shifts": 1,
            "start": d_st + datetime.timedelta(days=50), "finish": d_st + datetime.timedelta(days=56),
            "loc": "Toàn tuyến cống", "mach": "Máy đào + Máy đầm cóc + Máy lu rung", "crew": 8, "critical": True
        }
    ]
    
    # DANH MỤC VẬT TƯ CHUYÊN BIỆT CHO CỐNG HỘP TUYẾN A5
    materials: List[Dict[str, Any]] = [
        {"stt": 1, "code": "VT-01", "test": "Las-XD", "freq": "1 mẫu/200m3", "std": "TCVN 7570:2006", "name": "Cát đệm móng và đắp hông cống", "unit": "m3", "qty": 605.0, "price": 180000},
        {"stt": 2, "code": "VT-02", "test": "Las-XD", "freq": "1 mẫu/100m3", "std": "TCVN 7570:2006", "name": "Đá 4x6 lót móng cống hộp", "unit": "m3", "qty": 85.0, "price": 280000},
        {"stt": 3, "code": "VT-03", "test": "Las-XD", "freq": "1 mẫu/50 tấn", "std": "TCVN 2682:2009", "name": "Xi măng PC40 đổ bê tông cống", "unit": "tấn", "qty": 145.0, "price": 1650000},
        {"stt": 4, "code": "VT-04", "test": "Las-XD", "freq": "1 mẫu/20 tấn", "std": "TCVN 1651:2018", "name": "Thép thanh vằn CB500-V (D14-D22)", "unit": "tấn", "qty": 38.6, "price": 15800000},
        {"stt": 5, "code": "VT-05", "test": "Las-XD", "freq": "Theo lô giao hàng", "std": "TCVN 4453:1995", "name": "Ván khuôn thép định hình phủ phim", "unit": "m2", "qty": 890.0, "price": 130000},
        {"stt": 6, "code": "VT-06", "test": "Las-XD", "freq": "1 mẫu/lô cuộn", "std": "TCVN 9386:2012", "name": "Băng cản nước Waterstop PVC V200", "unit": "m", "qty": 360.0, "price": 125000},
        {"stt": 7, "code": "VT-07", "test": "Las-XD", "freq": "Theo lô sản phẩm", "std": "TCVN 9974:2013", "name": "Xốp chèn khe & Matit bitum chèn khe co giãn", "unit": "m", "qty": 180.0, "price": 85000},
        {"stt": 8, "code": "VT-08", "test": "Las-XD", "freq": "Theo đợt nghiệm thu", "std": "BS EN 124:2015", "name": "Tấm đan BTCT & Nắp ga gang thoát nước D400", "unit": "bộ", "qty": 18.0, "price": 4500000},
    ]
    
    # TỔ HỢP CỐT THÉP BÓC TÁCH CỐNG HỘP TUYẾN A5
    rebar_items = [
        (1, "Thép chịu lực bản đáy cống Ø20 CB500-V", 20, 320, 11.7, 2.47),
        (2, "Thép chịu lực thành và bản nắp Ø18 CB500-V", 18, 450, 9.8, 2.00),
        (3, "Thép đai cống hộp Ø12 CB240-T", 12, 680, 4.2, 0.888),
        (4, "Thép phân bố dọc cống Ø14 CB500-V", 14, 280, 11.7, 1.21),
        (5, "Thép gia cường vát góc cống Ø16 CB500-V", 16, 540, 2.8, 1.58),
        (6, "Thép tấm đan giảm tải đỉnh cống Ø16 CB500-V", 16, 160, 3.5, 1.58),
    ]
    
    t1_dir = os.path.join(target_dir, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
    os.makedirs(t1_dir, exist_ok=True)
    
    master_path = os.path.join(target_dir, f"MACRO_MASTER_14_SHEET_{proj.short_name}.xlsx")
    build_project_master_workbook(proj, tasks, materials, rebar_items, master_path)
    
    fleet_path = os.path.join(target_dir, f"MACRO_FLEET_MASTER_{proj.short_name}.xlsx")
    build_project_fleet_workbook(proj, tasks, fleet_path)
    
    companion = build_project_companion_files(proj, tasks, materials, target_dir)
    companion['fleet_template'] = fleet_path
    
    dispatcher = AECPackageDispatcher(base_output_dir=target_dir)
    manifest = dispatcher.dispatch_full_industrial_dossier(
        master_excel_path=master_path,
        project_name=proj.full_name, 
        target_dir=target_dir,
        companion_files=companion
    )
    
    print(f"Hoàn thành xuất 3 Tầng hồ sơ công nghiệp cho {proj.full_name}.")

if __name__ == "__main__":
    export_cong_hop()
