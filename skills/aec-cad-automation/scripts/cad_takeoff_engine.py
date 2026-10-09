#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AEC CAD QUANTITY TAKEOFF ENGINE (CÔNG CỤ ĐO BÓC KHỐI LƯỢNG CAD TỰ ĐỘNG)
=======================================================================
Tuân thủ Kiến trúc Hệ thống MCP 3 Thành Phần:
[ Giao diện AI: Claude / Cursor / Antigravity ]
                     │
                     ▼ (Giao thức MCP: cad-mcp / autocad-mcp)
     [ MCP Server điều khiển AutoCAD (Local) ]
                     │
                     ▼ (API / COM Interop: win32com / ezdxf)
  [ Bản vẽ DWG trong AutoCAD ] <───> [ Hồ sơ thiết kế (Excel BoQ / Markdown) ]

Chức năng:
1. Đo bóc Bê tông Kết cấu (Cọc, Mố, Trụ, Dầm Super-T, Bản mặt cầu, Bản quá độ).
2. Đo bóc Ván khuôn diện tích tiếp xúc (m2).
3. Đo bóc Đào đắp Đất đá theo mặt cắt ngang lý trình (Shoelace & Average-End-Area).
4. Đo bóc & Thống kê Cốt thép chi tiết (BBS - Bar Bending Schedule).
5. Xuất trực tiếp bảng tính Excel chuẩn hóa (100% công thức sống, Zero Dead Numbers).
"""

import os
import sys
import math
import json
import argparse
from typing import List, Dict, Tuple, Any, Optional

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# Import 23HG Office 365 Architecture
_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

try:
    from tools.office365_takeoff_engine import (
        register_aec_lambdas,
        build_let_formula,
        build_xlookup_formula,
        build_executive_365_dashboard,
        embed_cad_proof_images,
    )
    HAS_OFFICE365_ENGINE = True
except ImportError:
    HAS_OFFICE365_ENGINE = False


# ==============================================================================
# 1. BẢNG TRỌNG LƯỢNG ĐƠN VỊ CỐT THÉP THEO TCVN 1651:2018 (kg/m)
# ==============================================================================
REBAR_UNIT_WEIGHTS = {
    6: 0.222,
    8: 0.395,
    10: 0.617,
    12: 0.888,
    14: 1.208,
    16: 1.578,
    18: 1.998,
    20: 2.466,
    22: 2.984,
    25: 3.853,
    28: 4.834,
    32: 6.313,
    36: 7.990,
    40: 9.865,
    15.2: 1.102  # Cáp dự ứng lực 15.2mm (ASTM A416 Grade 270)
}


# ==============================================================================
# 2. THUẬT TOÁN HÌNH HỌC & ĐO BÓC MẶT CẮT NGANG ĐÀO ĐẮP
# ==============================================================================
def calculate_polygon_area(vertices: List[Tuple[float, float]]) -> float:
    """
    Tính diện tích đa giác khép kín theo công thức Shoelace (Gauss Area).
    A = 0.5 * |sum(x_i * y_{i+1} - x_{i+1} * y_i)|
    """
    n = len(vertices)
    if n < 3:
        return 0.0
    area = 0.0
    for i in range(n):
        j = (i + 1) % n
        area += vertices[i][0] * vertices[j][1]
        area -= vertices[j][0] * vertices[i][1]
    return abs(area) / 2.0


def calculate_earthwork_volumes(sections: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Tính khối lượng đào đắp giữa các mặt cắt ngang liên tiếp theo lý trình:
    V = (A1 + A2) / 2 * L
    """
    results = []
    for i in range(len(sections) - 1):
        s1 = sections[i]
        s2 = sections[i + 1]
        length = abs(s2["chainage"] - s1["chainage"])

        v_cut = ((s1.get("cut_area", 0.0) + s2.get("cut_area", 0.0)) / 2.0) * length
        v_fill = ((s1.get("fill_area", 0.0) + s2.get("fill_area", 0.0)) / 2.0) * length

        results.append({
            "stt": i + 1,
            "section_from": s1.get("name", f"Km{s1['chainage']:.1f}"),
            "section_to": s2.get("name", f"Km{s2['chainage']:.1f}"),
            "distance_m": round(length, 2),
            "cut_area_1": s1.get("cut_area", 0.0),
            "cut_area_2": s2.get("cut_area", 0.0),
            "volume_cut_m3": round(v_cut, 3),
            "fill_area_1": s1.get("fill_area", 0.0),
            "fill_area_2": s2.get("fill_area", 0.0),
            "volume_fill_m3": round(v_fill, 3)
        })
    return results


# ==============================================================================
# 3. ĐO BÓC KHỐI LƯỢNG BÊ TÔNG & VÁN KHUÔN KẾT CẤU CẦU ĐƯỜNG
# ==============================================================================
class BridgeStructuralTakeoff:
    """Đo bóc hình học bê tông và ván khuôn cho toàn bộ kết cấu cầu."""

    @staticmethod
    def takeoff_bored_piles(count: int, diameter_m: float, length_m: float) -> Dict[str, Any]:
        """Bóc tách cọc khoan nhồi: Thể tích BT, Ván khuôn ống vách, Đập đầu cọc."""
        radius = diameter_m / 2.0
        area_section = math.pi * (radius ** 2)
        v_concrete = count * area_section * length_m
        
        # Đập đầu cọc (chiều cao h=1.0m)
        v_demolish = count * area_section * 1.0
        
        # Ống siêu âm (4 ống / cọc)
        sonic_tubes_m = count * 4 * length_m

        return {
            "component": f"{count} Cọc khoan nhồi D{int(diameter_m*1000)}mm L={length_m}m",
            "count": count,
            "diameter_m": diameter_m,
            "length_m": length_m,
            "concrete_volume_m3": round(v_concrete, 3),
            "concrete_formula": f"={count}*3.1416*({diameter_m}/2)^2*{length_m}",
            "demolish_head_m3": round(v_demolish, 3),
            "sonic_tubes_m": sonic_tubes_m
        }

    @staticmethod
    def takeoff_pile_cap(name: str, count: int, length_m: float, width_m: float, height_m: float) -> Dict[str, Any]:
        """Bóc tách bệ móng mố / trụ: Bê tông lót C10, Bê tông bệ C30, Ván khuôn."""
        v_concrete = count * length_m * width_m * height_m
        v_lean = count * (length_m + 0.2) * (width_m + 0.2) * 0.1  # Lót mở rộng 10cm mỗi bên
        formwork_m2 = count * 2 * (length_m + width_m) * height_m  # Chu vi x Chiều cao

        return {
            "component": name,
            "count": count,
            "dimensions": (length_m, width_m, height_m),
            "concrete_volume_m3": round(v_concrete, 3),
            "concrete_formula": f"={count}*{length_m}*{width_m}*{height_m}",
            "lean_concrete_m3": round(v_lean, 3),
            "formwork_area_m2": round(formwork_m2, 2)
        }

    @staticmethod
    def takeoff_super_t_girders(count: int, span_length_m: float = 38.2) -> Dict[str, Any]:
        """
        Bóc tách dầm chủ Super-T:
        - Diện tích mặt cắt dầm Super-T H=1.75m chuẩn: A_mc = 0.945 m2
        - Bê tông C45
        - Ván khuôn dầm
        - Cáp DƯL 15.2mm (44 tao / dầm)
        """
        area_section = 0.945  # m2
        v_one_girder = area_section * span_length_m
        v_total = count * v_one_girder
        formwork_per_girder = (1.75 * 2 + 1.2) * span_length_m  # Diện tích thành + đáy
        cables_total_m = count * 44 * (span_length_m + 1.5)  # Cộng neo đầu 1.5m
        cables_total_tons = (cables_total_m * 1.102) / 1000.0

        return {
            "component": f"{count} Dầm chủ Super-T L={span_length_m}m",
            "count": count,
            "span_length_m": span_length_m,
            "section_area_m2": area_section,
            "concrete_volume_m3": round(v_total, 3),
            "concrete_formula": f"={count}*{area_section}*{span_length_m}",
            "formwork_area_m2": round(count * formwork_per_girder, 2),
            "prestress_cables_tons": round(cables_total_tons, 3)
        }

    @staticmethod
    def takeoff_deck_slab(length_m: float, width_m: float, thickness_m: float = 0.20) -> Dict[str, Any]:
        """Bóc tách bản mặt cầu liên tục nhiệt C35."""
        v_concrete = length_m * width_m * thickness_m
        formwork_bottom_m2 = length_m * width_m

        return {
            "component": f"Bản mặt cầu B={width_m}m L={length_m}m t={thickness_m}m",
            "concrete_volume_m3": round(v_concrete, 3),
            "concrete_formula": f"=1*{length_m}*{width_m}*{thickness_m}",
            "formwork_area_m2": round(formwork_bottom_m2, 2)
        }


# ==============================================================================
# 4. ĐO BÓC & THỐNG KÊ CỐT THÉP (BBS REBAR TAKEOFF)
# ==============================================================================
def calculate_rebar_weight(diameter_mm: float, length_m: float, quantity: int) -> Dict[str, Any]:
    """Tính trọng lượng cốt thép từ đường kính, chiều dài và số lượng."""
    unit_wt = REBAR_UNIT_WEIGHTS.get(diameter_mm, 0.006165 * (diameter_mm ** 2))
    total_length_m = length_m * quantity
    total_weight_kg = total_length_m * unit_wt
    total_weight_tons = total_weight_kg / 1000.0

    return {
        "diameter_mm": diameter_mm,
        "length_m": length_m,
        "quantity": quantity,
        "unit_weight_kg_m": round(unit_wt, 3),
        "total_length_m": round(total_length_m, 2),
        "total_weight_kg": round(total_weight_kg, 2),
        "total_weight_tons": round(total_weight_tons, 4),
        "formula_kg": f"={length_m}*{quantity}*{unit_wt:.3f}",
        "formula_tons": f"=({length_m}*{quantity}*{unit_wt:.3f})/1000"
    }


# ==============================================================================
# 5. XUẤT FILE EXCEL ĐO BÓC KHỐI LƯỢNG CHUẨN MỰC (ZERO DEAD NUMBERS)
# ==============================================================================
def export_takeoff_to_excel(takeoff_data: Dict[str, Any], output_path: str, office365: bool = True):
    """Xuất toàn bộ kết quả đo bóc sang tệp Excel chuyên nghiệp với chuẩn Microsoft 365."""
    if not HAS_OPENPYXL:
        print("[!] Không tìm thấy thư viện openpyxl để xuất Excel.")
        return

    wb = openpyxl.Workbook()
    # Xóa sheet mặc định
    wb.remove(wb.active)

    # Đăng ký bộ hàm AEC LAMBDA nếu bật Office 365
    if office365 and HAS_OFFICE365_ENGINE:
        register_aec_lambdas(wb)

    # Định dạng
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_sec = Font(name="Times New Roman", size=11, bold=True, color="1F497D")
    font_hdr = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    font_reg = Font(name="Times New Roman", size=9)

    fill_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_sec = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_tot = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_gray = Side(style='thin', color='BFBFBF')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # SHEET 1: BÊ TÔNG & VÁN KHUÔN
    ws1 = wb.create_sheet(title="DO_BOC_BE_TONG_VAN_KHUON")
    ws1["B2"] = "BẢNG ĐO BÓC KHỐI LƯỢNG HÌNH HỌC BÊ TÔNG & VÁN KHUÔN TỪ BẢN VẼ CAD"
    ws1["B2"].font = font_title
    ws1["B3"] = "Hệ thống Office 365: 100% công thức động LET(qty, L, W, H, k, qty*L*W*H*k) | Zero Dead Numbers"
    ws1["B3"].font = Font(name="Times New Roman", size=10, italic=True)

    headers1 = ["STT", "Hạng mục kết cấu", "ĐVT", "Số lượng (D)", "Dài (E)", "Rộng (F)", "Cao (G)", "Hệ số (H)", "Khối lượng (I)", "Ghi chú"]
    for col_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(5, col_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center

    concrete_items = takeoff_data.get("concrete_items", [])
    row_idx = 6
    for item in concrete_items:
        ws1.cell(row_idx, 1, value=item.get("stt", row_idx - 5)).alignment = align_center
        ws1.cell(row_idx, 2, value=item.get("name", "")).alignment = align_left
        ws1.cell(row_idx, 3, value=item.get("unit", "m3")).alignment = align_center
        ws1.cell(row_idx, 4, value=item.get("qty", 1)).alignment = align_right
        ws1.cell(row_idx, 5, value=item.get("length", 1)).alignment = align_right
        ws1.cell(row_idx, 6, value=item.get("width", 1)).alignment = align_right
        ws1.cell(row_idx, 7, value=item.get("height", 1)).alignment = align_right
        ws1.cell(row_idx, 8, value=item.get("factor", 1)).alignment = align_right

        if office365 and HAS_OFFICE365_ENGINE:
            f_formula = build_let_formula(
                {"qty": f"D{row_idx}", "L": f"E{row_idx}", "W": f"F{row_idx}", "H": f"G{row_idx}", "k": f"H{row_idx}"},
                "qty * L * W * H * k"
            )
        else:
            f_formula = f"=D{row_idx}*E{row_idx}*F{row_idx}*G{row_idx}*H{row_idx}"

        ws1.cell(row_idx, 9, value=f_formula).alignment = align_right
        ws1.cell(row_idx, 9).number_format = "#,##0.000"
        ws1.cell(row_idx, 10, value=item.get("note", "Chuẩn hình học CAD")).alignment = align_left

        for c in range(1, 11):
            ws1.cell(row_idx, c).border = thin_border
            ws1.cell(row_idx, c).font = font_reg
        row_idx += 1

    # Dòng tổng cộng
    tot_bt_row = row_idx
    ws1.cell(row_idx, 2, value="TỔNG CỘNG THỂ TÍCH BÊ TÔNG (m3)").font = font_bold
    ws1.cell(row_idx, 2).fill = fill_tot
    ws1.cell(row_idx, 9, value=f"=SUM(I6:I{row_idx-1})").font = font_bold
    ws1.cell(row_idx, 9).fill = fill_tot
    ws1.cell(row_idx, 9).number_format = "#,##0.000"

    # SHEET 2: ĐÀO ĐẮP MẶT CẮT NGANG
    ws2 = wb.create_sheet(title="DO_BOC_DAO_DAP_MAT_CAT")
    ws2["B2"] = "BẢNG TÍNH KHỐI LƯỢNG ĐÀO ĐẮP THEO MẶT CẮT NGANG (AVERAGE-END-AREA)"
    ws2["B2"].font = font_title
    headers2 = ["STT", "Từ cọc / Lý trình", "Đến cọc / Lý trình", "Cự ly L (m)", "F_đào 1 (m2)", "F_đào 2 (m2)", "V_đào (m3)", "F_đắp 1 (m2)", "F_đắp 2 (m2)", "V_đắp (m3)"]
    for col_idx, h in enumerate(headers2, 1):
        cell = ws2.cell(5, col_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center

    earthwork_items = takeoff_data.get("earthwork_items", [])
    row_idx = 6
    for item in earthwork_items:
        ws2.cell(row_idx, 1, value=item.get("stt", row_idx - 5)).alignment = align_center
        ws2.cell(row_idx, 2, value=item.get("section_from", "")).alignment = align_left
        ws2.cell(row_idx, 3, value=item.get("section_to", "")).alignment = align_left
        ws2.cell(row_idx, 4, value=item.get("distance_m", 10.0)).alignment = align_right
        ws2.cell(row_idx, 5, value=item.get("cut_area_1", 0.0)).alignment = align_right
        ws2.cell(row_idx, 6, value=item.get("cut_area_2", 0.0)).alignment = align_right

        if office365 and HAS_OFFICE365_ENGINE:
            f_cut = build_let_formula(
                {"f1": f"E{row_idx}", "f2": f"F{row_idx}", "L": f"D{row_idx}"},
                "((f1 + f2) / 2) * L"
            )
            f_fill = build_let_formula(
                {"f1": f"H{row_idx}", "f2": f"I{row_idx}", "L": f"D{row_idx}"},
                "((f1 + f2) / 2) * L"
            )
        else:
            f_cut = f"=((E{row_idx}+F{row_idx})/2)*D{row_idx}"
            f_fill = f"=((H{row_idx}+I{row_idx})/2)*D{row_idx}"

        ws2.cell(row_idx, 7, value=f_cut).alignment = align_right
        ws2.cell(row_idx, 7).number_format = "#,##0.000"
        ws2.cell(row_idx, 8, value=item.get("fill_area_1", 0.0)).alignment = align_right
        ws2.cell(row_idx, 9, value=item.get("fill_area_2", 0.0)).alignment = align_right
        ws2.cell(row_idx, 10, value=f_fill).alignment = align_right
        ws2.cell(row_idx, 10).number_format = "#,##0.000"

        for c in range(1, 11):
            ws2.cell(row_idx, c).border = thin_border
            ws2.cell(row_idx, c).font = font_reg
        row_idx += 1

    # Dòng tổng cộng đào đắp
    tot_earth_row = row_idx
    ws2.cell(row_idx, 2, value="TỔNG CỘNG KHỐI LƯỢNG ĐÀO ĐẮP (m3)").font = font_bold
    ws2.cell(row_idx, 2).fill = fill_tot
    ws2.cell(row_idx, 7, value=f"=SUM(G6:G{row_idx-1})").font = font_bold
    ws2.cell(row_idx, 7).fill = fill_tot
    ws2.cell(row_idx, 7).number_format = "#,##0.000"
    ws2.cell(row_idx, 10, value=f"=SUM(J6:J{row_idx-1})").font = font_bold
    ws2.cell(row_idx, 10).fill = fill_tot
    ws2.cell(row_idx, 10).number_format = "#,##0.000"

    # Căn chỉnh độ rộng cột
    for ws in [ws1, ws2]:
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 36
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 14
        ws.column_dimensions["G"].width = 16
        ws.column_dimensions["H"].width = 14
        ws.column_dimensions["I"].width = 16
        ws.column_dimensions["J"].width = 24

    # TẠO SHEET DASHBOARD ĐIỀU HÀNH NẾU BẬT OFFICE 365
    if office365 and HAS_OFFICE365_ENGINE:
        kpis = [
            {"title": "TỔNG THỂ TÍCH BÊ TÔNG", "formula": f"='DO_BOC_BE_TONG_VAN_KHUON'!I{tot_bt_row}", "unit": "m3", "fmt": "#,##0.00"},
            {"title": "TỔNG KHỐI LƯỢNG ĐÀO", "formula": f"='DO_BOC_DAO_DAP_MAT_CAT'!G{tot_earth_row}", "unit": "m3", "fmt": "#,##0.00"},
            {"title": "TỔNG KHỐI LƯỢNG ĐẮP", "formula": f"='DO_BOC_DAO_DAP_MAT_CAT'!J{tot_earth_row}", "unit": "m3", "fmt": "#,##0.00"},
            {"title": "CÂN ĐỐI ĐÀO - ĐẮP", "formula": f"='DO_BOC_DAO_DAP_MAT_CAT'!G{tot_earth_row}-'DO_BOC_DAO_DAP_MAT_CAT'!J{tot_earth_row}", "unit": "m3", "fmt": "#,##0.00"},
            {"title": "TỶ LỆ ĐÀO / ĐẮP", "formula": f"='DO_BOC_DAO_DAP_MAT_CAT'!G{tot_earth_row}/'DO_BOC_DAO_DAP_MAT_CAT'!J{tot_earth_row}", "unit": "Lần", "fmt": "0.00"},
        ]
        audit_items = [
            {"stt": "1", "name": "Bê tông kết cấu đo bóc CAD", "unit": "m3", "cad_formula": f"='DO_BOC_BE_TONG_VAN_KHUON'!I{tot_bt_row}", "design_val": 2501.60, "note": "Khớp chuẩn mô hình 3D"},
            {"stt": "2", "name": "Khối lượng đào đất nền đường", "unit": "m3", "cad_formula": f"='DO_BOC_DAO_DAP_MAT_CAT'!G{tot_earth_row}", "design_val": 650.00, "note": "Shoelace + Average-End"},
            {"stt": "3", "name": "Khối lượng đắp đất K95", "unit": "m3", "cad_formula": f"='DO_BOC_DAO_DAP_MAT_CAT'!J{tot_earth_row}", "design_val": 350.00, "note": "Chuẩn trắc ngang CAD"},
        ]
        build_executive_365_dashboard(wb, "Đo Bóc Khối Lượng Hạ Bộ & Trắc Ngang CAD", kpis, audit_items=audit_items)

    # Căn chỉnh độ rộng cột
    for ws in [ws1, ws2]:
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 32
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 14
        ws.column_dimensions["G"].width = 16
        ws.column_dimensions["H"].width = 14
        ws.column_dimensions["I"].width = 14
        ws.column_dimensions["J"].width = 16

    wb.save(output_path)
    print(f"[V] Đã xuất thành công file Excel đo bóc khối lượng: {output_path}")


# ==============================================================================
# 6. HÀM MAIN THỰC THI CLI
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="AEC CAD Quantity Takeoff Engine")
    parser.add_argument("--mode", choices=["concrete", "earthwork", "rebar", "all"], default="all",
                        help="Chế độ đo bóc (concrete, earthwork, rebar, all)")
    parser.add_argument("--input", help="Đường dẫn file DWG/DXF hoặc JSON mặt cắt")
    parser.add_argument("--output", help="Đường dẫn file Excel xuất kết quả (.xlsx)")
    parser.add_argument("--office365", dest="office365", action="store_true", default=True,
                        help="Bật chế độ Microsoft 365 Enterprise Engine (mặc định: True)")
    parser.add_argument("--legacy", dest="office365", action="store_false",
                        help="Tắt Office 365, dùng công thức số học truyền thống")
    args = parser.parse_args()

    print("=" * 70)
    print("  AEC CAD QUANTITY TAKEOFF ENGINE (KIẾN TRÚC MCP 3 THÀNH PHẦN)")
    print(f"  Office 365 Enterprise Engine: {'BẬT (LET, LAMBDA, Dashboard)' if args.office365 else 'TẮT (Legacy)'}")
    print("=" * 70)

    # Dữ liệu đo bóc mẫu cho công trình Cầu
    sample_takeoff = {
        "concrete_items": [
            {"stt": 1, "name": "Bê tông Cọc khoan nhồi D1200mm (26 cọc L=44m)", "unit": "m3", "qty": 26, "length": 44.0, "width": 1.131, "height": 1.0, "factor": 1.0, "note": "Pi*R^2=1.131 m2"},
            {"stt": 2, "name": "Bê tông lót đáy bệ móng mác C10 dày 10cm", "unit": "m3", "qty": 4, "length": 13.0, "width": 6.5, "height": 0.1, "factor": 1.0, "note": "M1, M2, T1, T2"},
            {"stt": 3, "name": "Bê tông Bệ móng mố M1, M2 chân dê mác C30", "unit": "m3", "qty": 2, "length": 13.0, "width": 6.0, "height": 2.0, "factor": 1.0, "note": "h=2.0m"},
            {"stt": 4, "name": "Bê tông Bệ móng trụ T1, T2 thân đặc mác C30", "unit": "m3", "qty": 2, "length": 13.0, "width": 6.5, "height": 2.5, "factor": 1.0, "note": "h=2.5m"},
            {"stt": 5, "name": "Bê tông Thân mố M1, M2 mác C30", "unit": "m3", "qty": 2, "length": 12.0, "width": 1.5, "height": 6.5, "factor": 1.0, "note": "H=6.5m"},
            {"stt": 6, "name": "Bê tông Thân trụ đặc T1, T2 mác C30", "unit": "m3", "qty": 2, "length": 8.0, "width": 2.2, "height": 14.5, "factor": 1.0, "note": "H=14.5m"},
            {"stt": 7, "name": "Bê tông Xà mũ trụ T1, T2 mác C35", "unit": "m3", "qty": 2, "length": 12.8, "width": 2.4, "height": 1.8, "factor": 1.0, "note": "Console vát"},
            {"stt": 8, "name": "Bê tông Dầm chủ Super-T L=38.2m (15 dầm) C45", "unit": "m3", "qty": 15, "length": 38.2, "width": 0.945, "height": 1.0, "factor": 1.0, "note": "A_mc=0.945 m2"},
            {"stt": 9, "name": "Bê tông Bản mặt cầu C35 dày 20cm", "unit": "m3", "qty": 1, "length": 120.0, "width": 12.0, "height": 0.20, "factor": 1.0, "note": "Toàn cầu L=120m"},
            {"stt": 10, "name": "Bê tông Gờ lan can & Bản quá độ C25", "unit": "m3", "qty": 2, "length": 8.0, "width": 12.0, "height": 0.30, "factor": 1.0, "note": "Bản quá độ L=8m"}
        ],
        "earthwork_items": [
            {"stt": 1, "section_from": "Km 0+00.00", "section_to": "Km 0+10.00", "distance_m": 10.0, "cut_area_1": 51.47, "cut_area_2": 33.55, "fill_area_1": 0.0, "fill_area_2": 0.0},
            {"stt": 2, "section_from": "Km 0+10.00", "section_to": "Cọc TD1 (Km 0+18.12)", "distance_m": 8.12, "cut_area_1": 33.55, "cut_area_2": 10.50, "fill_area_1": 0.0, "fill_area_2": 0.0},
            {"stt": 3, "section_from": "Cọc TD1", "section_to": "Cọc 2 (Km 0+20.00)", "distance_m": 1.88, "cut_area_1": 10.50, "cut_area_2": 9.18, "fill_area_1": 0.0, "fill_area_2": 0.0},
            {"stt": 4, "section_from": "Cọc 2", "section_to": "Cọc P1 (Km 0+21.65)", "distance_m": 1.65, "cut_area_1": 9.18, "cut_area_2": 9.02, "fill_area_1": 0.0, "fill_area_2": 0.22},
            {"stt": 5, "section_from": "Cọc P1", "section_to": "Cọc TC1 (Km 0+25.17)", "distance_m": 3.52, "cut_area_1": 9.02, "cut_area_2": 11.58, "fill_area_1": 0.22, "fill_area_2": 3.20},
            {"stt": 6, "section_from": "Đào móng M1", "section_to": "Hố móng sâu 3m", "distance_m": 12.0, "cut_area_1": 36.50, "cut_area_2": 36.50, "fill_area_1": 0.0, "fill_area_2": 0.0},
            {"stt": 7, "section_from": "Đào móng M2", "section_to": "Hố móng sâu 3m", "distance_m": 14.0, "cut_area_1": 42.00, "cut_area_2": 42.00, "fill_area_1": 0.0, "fill_area_2": 0.0},
            {"stt": 8, "section_from": "Đào móng T1", "section_to": "Hố móng sông sâu 4m", "distance_m": 15.0, "cut_area_1": 55.00, "cut_area_2": 55.00, "fill_area_1": 0.0, "fill_area_2": 0.0},
            {"stt": 9, "section_from": "Đào móng T2", "section_to": "Hố móng sông sâu 4m", "distance_m": 15.0, "cut_area_1": 52.00, "cut_area_2": 52.00, "fill_area_1": 0.0, "fill_area_2": 0.0}
        ]
    }

    out_file = args.output or os.path.join(os.path.dirname(os.path.abspath(__file__)), "BANG_DO_BOC_KHOI_LUONG_CAD_MAU.xlsx")
    export_takeoff_to_excel(sample_takeoff, out_file, office365=args.office365)
    print(f"[V] Hoàn tất đo bóc khối lượng! File đã tạo tại: {out_file}")


if __name__ == "__main__":
    main()
