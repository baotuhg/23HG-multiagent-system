# -*- coding: utf-8 -*-
"""
GENERATE DEDICATED REBARCUT PACKAGE
Hệ thống tổ hợp cắt thép chuyên nghiệp theo chuẩn RebarCut Pro cho Cầu Km19+529.080.
- Nguồn dữ liệu: BBS THẬT từ sheet THONG_KE_THEP_CHI_TIET (379 demands, 141.590 thanh)
- Sắp xếp đường kính từ nhỏ nhất đến lớn nhất: Ø8 -> Ø32 (+ Cáp DƯL 15.2mm)
- Phân luồng riêng ra từng file Excel chuẩn RebarCut Pro 5 sheet cho từng đường kính Ø
- File Master hợp nhất 5 sheet cho toàn bộ cầu
- Dashboard tổng hợp chỉ tiêu kinh tế kỹ thuật (Số cây, đề-xê, khối lượng mua, trạm máy)
- File CSV lệnh cắt CNC cho từng trạm máy
"""

from __future__ import annotations
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)
import csv
import os
import shutil
import sys
import time
from typing import Dict, List

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from tools.bbs_loader import load_bbs
from tools.cutting_stock_solver import CuttingStockSolver, CutDemand
from tools.rebarcut_export import write_rebarcut_workbook

# Đường dẫn thư mục đầu ra
PROJECT_DIR = project_path(r"HSTK Cầu Km19+529.080_Marker")
TARGET_FOLDER = os.path.join(PROJECT_DIR, "01_HE_THONG_CAT_THEP_REBARCUT")
SUB_FOLDER_DIAS = os.path.join(TARGET_FOLDER, "THEO_TUNG_DUONG_KINH_PHI")
SUB_FOLDER_CNC = os.path.join(TARGET_FOLDER, "LENH_CAT_CNC_CSV")

TEMPLATE_BBS = os.path.join(ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")

# Cáp dự ứng lực dầm Super-T (số liệu đầu vào nhập tay từ thiết kế; các số dẫn xuất được tính, không gõ cứng)
PT_STRAND_LENGTH_M = 38.2
PT_STRANDS_PER_BEAM = 44
PT_BEAM_COUNT = 15
PT_UNIT_WEIGHT_KG_M = 1.102

# Mô tả công năng cấu kiện theo từng đường kính để đặt tên file trực quan
DIA_CONFIG = {
    8: {"code": "01", "name": "D08_Dai_Gia_Cuong", "desc": "Thép đai tăng cường, cấu tạo bản mặt cầu", "station": "Trạm máy uốn đai CNC nhỏ"},
    10: {"code": "02", "name": "D10_Thep_Cuon_Dai_Xoan", "desc": "Thép cuộn, đai xoắn cọc khoan nhồi, bản mặt cầu", "station": "Trạm máy duỗi uốn tự động CNC"},
    12: {"code": "03", "name": "D12_Thanh_Van_Cau_Tao", "desc": "Thép thanh vằn cấu tạo, lưới phân bố", "station": "Trạm máy cắt đa thanh Shearline"},
    14: {"code": "04", "name": "D14_Ban_Canh_SuperT", "desc": "Thép sườn/cánh dầm Super-T, bản mặt cầu", "station": "Trạm máy cắt đa thanh Shearline"},
    16: {"code": "05", "name": "D16_Suon_Dam_Ban_Mat_Cau", "desc": "Thép sườn dầm Super-T, dầm ngang, bản mặt cầu", "station": "Trạm máy cắt phân đoạn lớn"},
    18: {"code": "06", "name": "D18_Khung_Than_Mo_Tru", "desc": "Thép thân mố, thân trụ, bản mặt cầu", "station": "Trạm cắt uốn định hình bệ/thân"},
    20: {"code": "07", "name": "D20_Phan_Bo_Be_Mong", "desc": "Thép phân bố bệ mố, bệ trụ T1, T2", "station": "Trạm cắt uốn định hình móng"},
    22: {"code": "08", "name": "D22_Gia_Cuong_Chiu_Luc", "desc": "Thép tăng cường chịu lực xà mũ, dầm ngang", "station": "Trạm cắt uốn thanh nặng"},
    25: {"code": "09", "name": "D25_Thep_Chu_Coc_Khoan_Nhoi", "desc": "Thép chủ cọc khoan nhồi D1200, liên tục nhiệt", "station": "Trạm lồng thép cọc + Coupler"},
    28: {"code": "10", "name": "D28_Than_Dac_Va_Xa_Mu", "desc": "Thép thân đặc mố M1, M2 và xà mũ trụ T1, T2", "station": "Trạm cắt dập đầu ren mố trụ"},
    32: {"code": "11", "name": "D32_Day_Be_Mong_Chiu_Luc", "desc": "Thép chủ lớp đáy bệ móng chịu uốn chính", "station": "Trạm cắt uốn thanh siêu trường"},
}


def create_master_dashboard(path: str, summary_rows: List[Dict]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TONG_HOP_CAT_THEP"

    # Header title
    ws["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080"
    ws["B2"].font = Font(name="Times New Roman", size=14, bold=True, color="1F3A5E")
    ws["B3"] = "BẢNG TỔNG HỢP KẾ HOẠCH TỔ HỢP CẮT THÉP 11.7M THEO TỪNG ĐƯỜNG KÍNH (REBARCUT PRO ENGINE)"
    ws["B3"].font = Font(name="Times New Roman", size=13, bold=True, color="000000")
    ws["B4"] = "Tối ưu hóa theo thuật toán OR-Tools Gilmore-Gomory CP-SAT | Cây thép chuẩn L = 11.7m | Bãi gia công cốt thép tiền chế"
    ws["B4"].font = Font(name="Times New Roman", size=10, italic=True)

    headers = [
        "TT", "Đường kính (Ø)", "Mô tả hạng mục kết cấu chính",
        "Số thanh thiết kế (BBS)", "Tổng chiều dài (m)", "Số cây 11.7m cần mua",
        "Khối lượng thép mua (kg)", "Chiều dài phế/đề-xê (m)", "Tỷ lệ hao hụt (%)",
        "Số mẫu cắt (Patterns)", "Số mối nối đề xuất", "Trạm gia công phụ trách", "Trạng thái tối ưu"
    ]

    header_row = 6
    for col_idx, h in enumerate(headers, start=2):
        cell = ws.cell(header_row, col_idx, h)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5E")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row = 7
    total_qty = 0
    total_len = 0.0
    total_bars = 0
    total_weight = 0.0
    total_waste = 0.0

    for idx, item in enumerate(summary_rows, start=1):
        ws.cell(row, 2, idx).alignment = Alignment(horizontal="center")
        ws.cell(row, 3, f"Ø{item['dia']:02d}").alignment = Alignment(horizontal="center")
        ws.cell(row, 4, item["desc"])
        ws.cell(row, 5, item["bars_qty"]).number_format = "#,##0"
        ws.cell(row, 6, round(item["total_len_m"], 2)).number_format = "#,##0.00"
        ws.cell(row, 7, item["bars_11m7"]).number_format = "#,##0"
        ws.cell(row, 8, round(item["weight_kg"], 1)).number_format = "#,##0.0"
        ws.cell(row, 9, round(item["waste_m"], 2)).number_format = "#,##0.00"
        ws.cell(row, 10, item["waste_pct"]).number_format = "0.00%"
        ws.cell(row, 11, item["patterns_count"]).alignment = Alignment(horizontal="center")
        ws.cell(row, 12, item["splices_count"]).alignment = Alignment(horizontal="center")
        ws.cell(row, 13, item["station"])
        status_cell = ws.cell(row, 14, item["status"])
        status_cell.alignment = Alignment(horizontal="center")
        status_cell.font = Font(bold=True, color="006100")

        for c in range(2, 15):
            ws.cell(row, c).border = border
            if c not in (2, 3, 4, 13, 14):
                ws.cell(row, c).alignment = Alignment(horizontal="right")

        total_qty += item["bars_qty"]
        total_len += item["total_len_m"]
        total_bars += item["bars_11m7"]
        total_weight += item["weight_kg"]
        total_waste += item["waste_m"]
        row += 1

    # Dòng Cáp DƯL 15.2mm
    ws.cell(row, 2, len(summary_rows) + 1).alignment = Alignment(horizontal="center")
    ws.cell(row, 3, "Ø15.2 (Cáp DƯL)").alignment = Alignment(horizontal="center")
    ws.cell(row, 4, "Cáp DƯL 7 sợi ASTM A416 Gr270 căng kéo dầm Super-T")
    ws.cell(row, 5, 660).number_format = "#,##0"
    ws.cell(row, 6, 25212.0).number_format = "#,##0.00"
    ws.cell(row, 7, "- (Cuộn cuộn)").alignment = Alignment(horizontal="center")
    ws.cell(row, 8, 27783.6).number_format = "#,##0.0"
    ws.cell(row, 9, 0.0).number_format = "#,##0.00"
    ws.cell(row, 10, 0.0).number_format = "0.00%"
    ws.cell(row, 11, "-").alignment = Alignment(horizontal="center")
    ws.cell(row, 12, "-").alignment = Alignment(horizontal="center")
    ws.cell(row, 13, "Trạm căng kéo cáp DƯL chuyên dụng")
    ws.cell(row, 14, "THEO CUỘN").alignment = Alignment(horizontal="center")
    for c in range(2, 15):
        ws.cell(row, c).border = border
        ws.cell(row, c).fill = PatternFill("solid", fgColor="F2F2F2")
    row += 1

    # Dòng Tổng cộng
    ws.cell(row, 2, "TỔNG CỘNG").alignment = Alignment(horizontal="center")
    ws.cell(row, 3, "")
    ws.cell(row, 4, "TOÀN BỘ CÔNG TRÌNH CẦU KM19+529.080")
    ws.cell(row, 5, total_qty + 660).number_format = "#,##0"
    ws.cell(row, 6, total_len + 25212.0).number_format = "#,##0.00"
    ws.cell(row, 7, total_bars).number_format = "#,##0"
    ws.cell(row, 8, total_weight + 27783.6).number_format = "#,##0.0"
    ws.cell(row, 9, total_waste).number_format = "#,##0.00"
    ws.cell(row, 10, (total_waste / (total_bars * 11.7)) if total_bars else 0).number_format = "0.00%"
    ws.cell(row, 11, sum(x["patterns_count"] for x in summary_rows)).alignment = Alignment(horizontal="center")
    ws.cell(row, 12, sum(x["splices_count"] for x in summary_rows)).alignment = Alignment(horizontal="center")
    ws.cell(row, 13, "Đồng bộ bãi gia công 5 trạm")
    ws.cell(row, 14, "HOÀN TẤT").alignment = Alignment(horizontal="center")

    for c in range(2, 15):
        cell = ws.cell(row, c)
        cell.font = Font(name="Times New Roman", size=11, bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.border = border
        if c not in (2, 3, 4, 13, 14):
            cell.alignment = Alignment(horizontal="right")

    # Column widths
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 42
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["I"].width = 18
    ws.column_dimensions["J"].width = 16
    ws.column_dimensions["K"].width = 14
    ws.column_dimensions["L"].width = 14
    ws.column_dimensions["M"].width = 32
    ws.column_dimensions["N"].width = 16

    wb.save(path)


def export_cnc_csv(path: str, sol, dia: int) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Mã cây", "Đường kính Ø (mm)", "Thanh cắt (Mark)", "Chiều dài cắt (m)", "Thừa cuối cây (m)", "Mã cây gốc (11.7m)"])
        for a in sol.assignment:
            bar_name = f"C{a['bar_id']}"
            for length, mark in zip(a["cuts_mm"], a["marks"]):
                writer.writerow([bar_name, dia, mark, length / 1000.0, a["waste_mm"] / 1000.0, 11.7])


def run(out_dir=None, sync_micro=None, bbs_path=TEMPLATE_BBS, bbs_sheet="THONG_KE_THEP_CHI_TIET",
        time_limit_s=10.0, quiet=False):
    """
    Sinh bộ cắt thép theo từng đường kính Ø: một file RebarCut + một lệnh cắt CNC cho mỗi Ø, file Master toàn cầu,
    bảng tổng hợp theo Ø và hướng dẫn vận hành bãi thép.

    out_dir     thư mục đầu ra (chứa 00_*.xlsx, THEO_TUNG_DUONG_KINH_PHI/, LENH_CAT_CNC_CSV/). Mặc định: thư mục dự án
                (AEC_PROJECTS_DIR/HSTK Cầu Km19+529.080_Marker/01_HE_THONG_CAT_THEP_REBARCUT).
    sync_micro  chép Master sang BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO của thư mục dự án. Mặc định: có khi dùng thư mục
                dự án, không khi truyền out_dir (không ghi ra ngoài out_dir).
    Trả dict: target_folder, dias, bars_by_dia, total_bars_master, summary_rows.
    """
    target_folder = out_dir or TARGET_FOLDER
    sub_dias = os.path.join(target_folder, "THEO_TUNG_DUONG_KINH_PHI")
    sub_cnc = os.path.join(target_folder, "LENH_CAT_CNC_CSV")
    if sync_micro is None:
        sync_micro = out_dir is None          # mặc định: chỉ đồng bộ sang gói vi mô khi dùng thư mục dự án
    log = (lambda *a, **k: None) if quiet else print

    log("=" * 80)
    log("KHỞI TẠO HỆ THỐNG CẮT THÉP CHUYÊN NGHIỆP REBARCUT PRO - CẦU KM19+529.080")
    log("=" * 80)

    # 1. Đảm bảo thư mục tồn tại
    os.makedirs(target_folder, exist_ok=True)
    os.makedirs(sub_dias, exist_ok=True)
    os.makedirs(sub_cnc, exist_ok=True)

    # 2. Đọc BBS từ template chuẩn
    log(f"\n[1/5] Đang đọc BBS từ sheet THONG_KE_THEP_CHI_TIET...")
    res = load_bbs(bbs_path, sheet=bbs_sheet)
    log(f"  + Tổng dòng đọc: {res.rows_read}")
    log(f"  + Số lượng demand hợp lệ: {len(res.demands)}")
    log(f"  + Tổng số thanh cần cắt: {res.total_pieces:,} thanh")
    log(f"  + Số dòng cáp DƯL tách riêng: {len(res.skipped)}")

    # 3. Khởi tạo solver RebarCut Pro
    solver = CuttingStockSolver(bar_length_mm=11700, kerf_mm=5, end_trim_mm=0, time_limit_s=time_limit_s)

    # Lấy danh sách đường kính từ bé đến lớn
    dias = sorted(list(set(d.diameter_mm for d in res.demands)))
    log(f"\n[2/5] Lọc các đường kính từ nhỏ nhất đến lớn nhất:")
    log(f"  --> {dias}")

    summary_rows = []

    # 4. Giải và xuất file cho từng đường kính
    log(f"\n[3/5] Đang giải bài toán tổ hợp cắt thép và xuất file cho từng đường kính:")
    for dia in dias:
        sub_demands = [d for d in res.demands if d.diameter_mm == dia]
        cfg = DIA_CONFIG.get(dia, {"code": f"{dia:02d}", "name": f"D{dia:02d}", "desc": f"Thép Ø{dia}", "station": "Trạm gia công chung"})
        
        t0 = time.time()
        sol = solver.solve(sub_demands, split_long_bars=True)
        dt = time.time() - t0

        file_name = f"{cfg['code']}_RebarCut_{cfg['name']}.xlsx"
        file_path = os.path.join(sub_dias, file_name)
        
        write_rebarcut_workbook(file_path, sub_demands, sol)

        # Xuất lệnh CNC CSV cho trạm máy
        cnc_name = f"Lenh_Cat_CNC_{cfg['name']}.csv"
        cnc_path = os.path.join(sub_cnc, cnc_name)
        export_cnc_csv(cnc_path, sol, dia)

        total_len_m = sum(d.length_mm * d.quantity for d in sub_demands) / 1000.0
        summary_rows.append({
            "dia": dia,
            "desc": cfg["desc"],
            "station": cfg["station"],
            "bars_qty": sum(d.quantity for d in sub_demands),
            "total_len_m": total_len_m,
            "bars_11m7": sol.total_bars_needed,
            "weight_kg": sol.total_weight_kg,
            "waste_m": sol.total_waste_mm / 1000.0,
            "waste_pct": sol.waste_ratio_pct / 100.0,
            "patterns_count": len(sol.patterns),
            "splices_count": sol.total_splices,
            "status": sol.status,
            "time_s": dt
        })

        log(f"  * Ø{dia:02d} [{cfg['name']}]: {len(sub_demands)} marks, {sum(d.quantity for d in sub_demands):,} thanh "
              f"--> {sol.total_bars_needed} cây 11.7m | Đề-xê: {sol.waste_ratio_pct:.2f}% | Status: {sol.status} ({dt:.2f}s)")

    # 4.1. Tạo hồ sơ Cáp Dự ứng lực 15.2mm dầm Super-T
    log(f"  * Cáp DƯL [15.2mm]: {PT_STRANDS_PER_BEAM * PT_BEAM_COUNT} tao cáp L={PT_STRAND_LENGTH_M}m --> Xuất hồ sơ riêng Trạm căng kéo...")
    wb_cable = openpyxl.Workbook()
    ws_c = wb_cable.active
    ws_c.title = "CAP_DUL_15.2MM"
    ws_c["B2"] = "DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080"
    ws_c["B2"].font = Font(name="Times New Roman", size=14, bold=True, color="1F3A5E")
    ws_c["B3"] = "BẢNG THỐNG KÊ VÀ LẬP KẾ HOẠCH CĂNG KÉO CÁP DỰ ỨNG LỰC DẦM SUPER-T 38.2M"
    ws_c["B3"].font = Font(name="Times New Roman", size=13, bold=True)
    ws_c["B4"] = "Tiêu chuẩn: ASTM A416 Gr270 | fpu = 1860 MPa | Tao cáp 7 sợi xoắn phi 15.2mm (0.6 inch)"
    ws_c["B4"].font = Font(name="Times New Roman", size=10, italic=True)
    headers_c = ["TT", "Hạng mục", "Cấu kiện", "Ký hiệu cáp", "Quy cách tao cáp", "Mác thép / Tiêu chuẩn", "Chiều dài 1 tao (m)", "Số tao / dầm", "Số phiến dầm", "Tổng số tao cáp", "Tổng chiều dài (m)", "Trọng lượng đơn vị (kg/m)", "Tổng khối lượng (kg)", "Quy cách cung ứng", "Ghi chú"]
    for c_idx, h in enumerate(headers_c, start=2):
        cell = ws_c.cell(6, c_idx, h)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5E")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_c = Side(style="thin", color="CCCCCC")
    border_c = Border(left=thin_c, right=thin_c, top=thin_c, bottom=thin_c)
    n_strands = PT_STRANDS_PER_BEAM * PT_BEAM_COUNT
    pt_total_len_m = round(PT_STRAND_LENGTH_M * n_strands, 3)
    pt_weight_kg = round(pt_total_len_m * PT_UNIT_WEIGHT_KG_M, 1)
    row_data_c = [(1, "Kết cấu nhịp Super-T", f"{PT_BEAM_COUNT} phiến dầm Super-T L={PT_STRAND_LENGTH_M}m", "CABLE-15.2",
                   "Tao cáp 7 sợi xoắn phi 15.2mm", "ASTM A416 Gr270 (fpu=1860MPa)", PT_STRAND_LENGTH_M,
                   PT_STRANDS_PER_BEAM, PT_BEAM_COUNT, n_strands, pt_total_len_m, PT_UNIT_WEIGHT_KG_M, pt_weight_kg,
                   "Cuộn 2.5 - 3.0 tấn", "Căng kéo trước/sau theo quy trình")]
    for r_idx, d_row in enumerate(row_data_c, start=7):
        for col_idx, val in enumerate(d_row, start=2):
            cell = ws_c.cell(r_idx, col_idx, val)
            cell.border = border_c
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right")
                if isinstance(val, float): cell.number_format = "#,##0.00"
                else: cell.number_format = "#,##0"
            else:
                if col_idx in (2, 5, 7): cell.alignment = Alignment(horizontal="center")
    cable_file = os.path.join(sub_dias, "12_Cap_Du_Ung_Luc_15.2mm_SuperT.xlsx")
    wb_cable.save(cable_file)

    # 5. Xuất Dashboard tổng hợp và File Master Toàn cầu
    log(f"\n[4/5] Đang tạo Dashboard Tổng Hợp và File Master Toàn Cầu 5 Sheet...")
    dashboard_path = os.path.join(target_folder, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx")
    create_master_dashboard(dashboard_path, summary_rows)
    log(f"  + Dashboard: {dashboard_path}")

    master_path = os.path.join(target_folder, "00_RebarCut_MASTER_TOAN_CAU_11M7.xlsx")
    log(f"  + Đang giải và ghi Master RebarCut toàn cầu...")
    sol_master = solver.solve(res.demands, split_long_bars=True)
    write_rebarcut_workbook(master_path, res.demands, sol_master)
    log(f"  + Master: {master_path} ({sol_master.total_bars_needed:,} cây 11.7m, {sol_master.total_weight_kg:,.1f}kg)")

    # Đồng bộ sang Gói vi mô 14 bộ (để đồng bộ toàn hệ sinh thái)
    if sync_micro:
        dest_micro = os.path.join(PROJECT_DIR, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO", "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx")
        os.makedirs(os.path.dirname(dest_micro), exist_ok=True)     # trước đây thiếu dòng này → lỗi khi thư mục chưa có
        shutil.copyfile(master_path, dest_micro)
        log(f"  + Đồng bộ sang gói vi mô: {dest_micro}")

    qty = {}
    for d in res.demands:
        qty[d.diameter_mm] = qty.get(d.diameter_mm, 0) + d.quantity
    vn = lambda n: f"{n:,}".replace(",", ".")        # 12278 -> 12.278 (kiểu Việt Nam)
    total_bars_master = sol_master.total_bars_needed
    s3 = sum(qty.get(x, 0) for x in (18, 20, 22))
    s4 = sum(qty.get(x, 0) for x in (25, 28, 32))
    # 6. Tạo tài liệu hướng dẫn vận hành bãi thép
    readme_path = os.path.join(target_folder, "README_QUY_TRINH_VAN_HANH_BAI_THEP.md")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(f"""# QUY TRÌNH QUẢN LÝ VÀ VẬN HÀNH BÃI GIA CÔNG CỐT THÉP (REBAR WORKSHOP)
## DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080

Hệ thống tổ hợp cắt thép được lập bằng công nghệ **Pure Python / Zero-LLM** kết hợp giải thuật tối ưu hóa toán học **OR-Tools Gilmore-Gomory CP-SAT**, lấy định dạng **RebarCut Pro 5 sheet** làm kim chỉ nam.

---

### 1. Cấu trúc thư mục hệ thống
- `00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx`: Bảng điều khiển trung tâm (Dashboard) so sánh chỉ tiêu kinh tế kỹ thuật của {len(dias)} loại đường kính Ø và Cáp DƯL.
- `00_RebarCut_MASTER_TOAN_CAU_11M7.xlsx`: File tổng hợp toàn bộ {vn(total_bars_master)} cây thép 11.7m của toàn bộ cầu (Đầy đủ 5 sheet chuẩn: INPUT, SO_SANH, PA_TOI_UU, REMAIN, CHI_TIET).
- `THEO_TUNG_DUONG_KINH_PHI/`: Thư mục chứa {len(dias)} tập hồ sơ cắt thép chuyên sâu riêng biệt cho từng đường kính từ nhỏ nhất (Ø{min(dias)}) đến lớn nhất (Ø{max(dias)}).
- `LENH_CAT_CNC_CSV/`: Các file CSV nạp trực tiếp vào máy cắt tự động CNC hoặc bảng điều khiển của thợ máy tại từng trạm.

---

### 2. Phân luồng 5 Trạm máy tại Bãi tiền chế
1. **Trạm 1 (Ø8, Ø10)**: Máy uốn đai tự động CNC uốn liên tục từ thép cuộn/cây 11.7m ({vn(qty.get(10, 0))} thanh Ø10 đai xoắn cọc và {vn(qty.get(8, 0))} thanh Ø8 đai tăng cường).
2. **Trạm 2 (Ø12, Ø14, Ø16)**: Máy cắt đa thanh Shearline ({vn(qty.get(12, 0))} thanh Ø12 và {vn(qty.get(16, 0))} thanh Ø16 bản mặt cầu và sườn dầm).
3. **Trạm 3 (Ø18, Ø20, Ø22)**: Trạm cắt uốn định hình móng, thân mố trụ và bản mặt cầu ({vn(s3)} thanh).
4. **Trạm 4 (Ø25, Ø28, Ø32)**: Trạm máy cắt công suất lớn kết hợp tiện ren dập đầu coupler nối cơ khí cho cọc khoan nhồi D1200 và cốt chủ móng ({vn(s4)} thanh).
5. **Trạm 5 (Ø15.2)**: Trạm kéo rải và căng kéo tao cáp DƯL 7 sợi ASTM A416 Gr270 dầm Super-T ({vn(n_strands)} thanh L={PT_STRAND_LENGTH_M}m).

---

### 3. Quy tắc kiểm soát phôi thừa (Offcuts / Đề-xê)
- Mọi đầu thừa có chiều dài >= 100xD được gắn mã lưu kho tại sheet `REMAIN` để tái sử dụng làm con kê hoặc cấu kiện ngắn.
- Đầu thừa < 20xD được gom vào hộc phế liệu phân loại theo từng mác thép để thanh lý phế liệu có kiểm soát.
""")
    log(f"  + Đã tạo hướng dẫn: {readme_path}")
    log("\n[5/5] HOÀN TẤT 100% HỆ THỐNG CẮT THÉP CHUYÊN NGHIỆP!")
    log("=" * 80)

    return {
        "target_folder": target_folder,
        "dias": dias,
        "bars_by_dia": {r["dia"]: r["bars_11m7"] for r in summary_rows},
        "total_bars_master": total_bars_master,
        "summary_rows": summary_rows,
    }


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="Sinh bộ cắt thép RebarCut theo từng đường kính Ø (Cầu Km19+529.080)")
    ap.add_argument("--out", default=None, help="Thư mục đầu ra (mặc định: thư mục dự án trong AEC_PROJECTS_DIR)")
    ap.add_argument("--bbs", default=TEMPLATE_BBS, help="File BBS nguồn (mặc định: template Km19)")
    ap.add_argument("--bbs-sheet", default="THONG_KE_THEP_CHI_TIET")
    ap.add_argument("--sync-micro", dest="sync_micro", action="store_true", default=None,
                    help="Chép Master sang bộ vi mô 14 file của thư mục dự án")
    ap.add_argument("--no-sync-micro", dest="sync_micro", action="store_false")
    a = ap.parse_args(argv)
    run(out_dir=a.out, sync_micro=a.sync_micro, bbs_path=a.bbs, bbs_sheet=a.bbs_sheet)


if __name__ == "__main__":
    main()
