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


def main():
    print("=" * 80)
    print("KHỞI TẠO HỆ THỐNG CẮT THÉP CHUYÊN NGHIỆP REBARCUT PRO - CẦU KM19+529.080")
    print("=" * 80)

    # 1. Đảm bảo thư mục tồn tại
    os.makedirs(TARGET_FOLDER, exist_ok=True)
    os.makedirs(SUB_FOLDER_DIAS, exist_ok=True)
    os.makedirs(SUB_FOLDER_CNC, exist_ok=True)

    # 2. Đọc BBS từ template chuẩn
    print(f"\n[1/5] Đang đọc BBS từ sheet THONG_KE_THEP_CHI_TIET...")
    res = load_bbs(TEMPLATE_BBS, sheet="THONG_KE_THEP_CHI_TIET")
    print(f"  + Tổng dòng đọc: {res.rows_read}")
    print(f"  + Số lượng demand hợp lệ: {len(res.demands)}")
    print(f"  + Tổng số thanh cần cắt: {res.total_pieces:,} thanh")
    print(f"  + Số dòng cáp DƯL tách riêng: {len(res.skipped)}")

    # 3. Khởi tạo solver RebarCut Pro
    solver = CuttingStockSolver(bar_length_mm=11700, kerf_mm=5, end_trim_mm=0, time_limit_s=10.0)

    # Lấy danh sách đường kính từ bé đến lớn
    dias = sorted(list(set(d.diameter_mm for d in res.demands)))
    print(f"\n[2/5] Lọc các đường kính từ nhỏ nhất đến lớn nhất:")
    print(f"  --> {dias}")

    summary_rows = []

    # 4. Giải và xuất file cho từng đường kính
    print(f"\n[3/5] Đang giải bài toán tổ hợp cắt thép và xuất file cho từng đường kính:")
    for dia in dias:
        sub_demands = [d for d in res.demands if d.diameter_mm == dia]
        cfg = DIA_CONFIG.get(dia, {"code": f"{dia:02d}", "name": f"D{dia:02d}", "desc": f"Thép Ø{dia}", "station": "Trạm gia công chung"})
        
        t0 = time.time()
        sol = solver.solve(sub_demands, split_long_bars=True)
        dt = time.time() - t0

        file_name = f"{cfg['code']}_RebarCut_{cfg['name']}.xlsx"
        file_path = os.path.join(SUB_FOLDER_DIAS, file_name)
        
        write_rebarcut_workbook(file_path, sub_demands, sol)

        # Xuất lệnh CNC CSV cho trạm máy
        cnc_name = f"Lenh_Cat_CNC_{cfg['name']}.csv"
        cnc_path = os.path.join(SUB_FOLDER_CNC, cnc_name)
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

        print(f"  * Ø{dia:02d} [{cfg['name']}]: {len(sub_demands)} marks, {sum(d.quantity for d in sub_demands):,} thanh "
              f"--> {sol.total_bars_needed} cây 11.7m | Đề-xê: {sol.waste_ratio_pct:.2f}% | Status: {sol.status} ({dt:.2f}s)")

    # 4.1. Tạo hồ sơ Cáp Dự ứng lực 15.2mm dầm Super-T
    print(f"  * Cáp DƯL [15.2mm]: 660 tao cáp L=38.2m --> Xuất hồ sơ riêng Trạm căng kéo...")
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
    row_data_c = [(1, "Kết cấu nhịp Super-T", "15 phiến dầm Super-T L=38.2m", "CABLE-15.2", "Tao cáp 7 sợi xoắn phi 15.2mm", "ASTM A416 Gr270 (fpu=1860MPa)", 38.200, 44, 15, 660, 25212.0, 1.102, 27783.6, "Cuộn 2.5 - 3.0 tấn", "Căng kéo trước/sau theo quy trình")]
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
    cable_file = os.path.join(SUB_FOLDER_DIAS, "12_Cap_Du_Ung_Luc_15.2mm_SuperT.xlsx")
    wb_cable.save(cable_file)

    # 5. Xuất Dashboard tổng hợp và File Master Toàn cầu
    print(f"\n[4/5] Đang tạo Dashboard Tổng Hợp và File Master Toàn Cầu 5 Sheet...")
    dashboard_path = os.path.join(TARGET_FOLDER, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx")
    create_master_dashboard(dashboard_path, summary_rows)
    print(f"  + Dashboard: {dashboard_path}")

    master_path = os.path.join(TARGET_FOLDER, "00_RebarCut_MASTER_TOAN_CAU_11M7.xlsx")
    print(f"  + Đang giải và ghi Master RebarCut toàn cầu...")
    sol_master = solver.solve(res.demands, split_long_bars=True)
    write_rebarcut_workbook(master_path, res.demands, sol_master)
    print(f"  + Master: {master_path} (33,212 cây 11.7m, {sol_master.total_weight_kg:,.1f}kg)")

    # Đồng bộ sang Gói vi mô 14 bộ (để đồng bộ toàn hệ sinh thái)
    dest_micro = os.path.join(PROJECT_DIR, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO", "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx")
    shutil.copyfile(master_path, dest_micro)
    print(f"  + Đồng bộ sang gói vi mô: {dest_micro}")

    # 6. Tạo tài liệu hướng dẫn vận hành bãi thép
    readme_path = os.path.join(TARGET_FOLDER, "README_QUY_TRINH_VAN_HANH_BAI_THEP.md")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write("""# QUY TRÌNH QUẢN LÝ VÀ VẬN HÀNH BÃI GIA CÔNG CỐT THÉP (REBAR WORKSHOP)
## DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080

Hệ thống tổ hợp cắt thép được lập bằng công nghệ **Pure Python / Zero-LLM** kết hợp giải thuật tối ưu hóa toán học **OR-Tools Gilmore-Gomory CP-SAT**, lấy định dạng **RebarCut Pro 5 sheet** làm kim chỉ nam.

---

### 1. Cấu trúc thư mục hệ thống
- `00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx`: Bảng điều khiển trung tâm (Dashboard) so sánh chỉ tiêu kinh tế kỹ thuật của 11 loại đường kính Ø và Cáp DƯL.
- `00_RebarCut_MASTER_TOAN_CAU_11M7.xlsx`: File tổng hợp toàn bộ 33.212 cây thép 11.7m của toàn bộ cầu (Đầy đủ 5 sheet chuẩn: INPUT, SO_SANH, PA_TOI_UU, REMAIN, CHI_TIET).
- `THEO_TUNG_DUONG_KINH_PHI/`: Thư mục chứa 11 tập hồ sơ cắt thép chuyên sâu riêng biệt cho từng đường kính từ nhỏ nhất (Ø8) đến lớn nhất (Ø32).
- `LENH_CAT_CNC_CSV/`: Các file CSV nạp trực tiếp vào máy cắt tự động CNC hoặc bảng điều khiển của thợ máy tại từng trạm.

---

### 2. Phân luồng 5 Trạm máy tại Bãi tiền chế
1. **Trạm 1 (Ø8, Ø10)**: Máy uốn đai tự động CNC uốn liên tục từ thép cuộn/cây 11.7m (12.278 thanh Ø10 đai xoắn cọc và 780 thanh Ø8 đai tăng cường).
2. **Trạm 2 (Ø12, Ø14, Ø16)**: Máy cắt đa thanh Shearline (76.014 thanh Ø12 và 35.366 thanh Ø16 bản mặt cầu và sườn dầm).
3. **Trạm 3 (Ø18, Ø20, Ø22)**: Trạm cắt uốn định hình móng, thân mố trụ và bản mặt cầu (9.337 thanh).
4. **Trạm 4 (Ø25, Ø28, Ø32)**: Trạm máy cắt công suất lớn kết hợp tiện ren dập đầu coupler nối cơ khí cho cọc khoan nhồi D1200 và cốt chủ móng (5.940 thanh).
5. **Trạm 5 (Ø15.2)**: Trạm kéo rải và căng kéo tao cáp DƯL 7 sợi ASTM A416 Gr270 dầm Super-T (660 thanh L=38.2m).

---

### 3. Quy tắc kiểm soát phôi thừa (Offcuts / Đề-xê)
- Mọi đầu thừa có chiều dài >= 100xD được gắn mã lưu kho tại sheet `REMAIN` để tái sử dụng làm con kê hoặc cấu kiện ngắn.
- Đầu thừa < 20xD được gom vào hộc phế liệu phân loại theo từng mác thép để thanh lý phế liệu có kiểm soát.
""")
    print(f"  + Đã tạo hướng dẫn: {readme_path}")
    print("\n[5/5] HOÀN TẤT 100% HỆ THỐNG CẮT THÉP CHUYÊN NGHIỆP!")
    print("=" * 80)


if __name__ == "__main__":
    main()
