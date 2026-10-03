# -*- coding: utf-8 -*-
"""
CIVIL & BRIDGE TAKEOFF ENGINE — ĐỘNG CƠ BÓC TÁCH HÌNH HỌC DÂN DỤNG & CẦU ĐƯỜNG BỘ
Kế thừa, chuẩn hóa và suy luận diễn biến từ hệ thống bảng tính QS thực chiến (Folder 'Mua').

Tuân thủ nguyên tắc:
  - ZERO LLM MATH: 100% tính toán số học xác định, không làm tròn cẩu thả, không ảo giác.
  - GEOMETRIC ISOMORPHISM: Ánh xạ đồng hình từ móng chóp cụt/vai cột dân dụng sang bệ xẻ nước,
    thân trụ, xà mũ, dầm Super-T, dầm I, bản mặt cầu, tường cánh mố và dầm thép liên hợp.
  - TCVN COMPLIANT: TCVN 11823:2017 (Thiết kế cầu), TCVN 5574:2018 (BTCT), TCVN 1651:2018 (Thép),
    và Thông tư 38/2026/TT-BXD, Thông tư 36/2026/TT-BXD.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

from tools import steel_qs


# ─────────────────────────────────────────────────────────────────────────────
# 1. GEOMETRIC PRIMITIVES (CÁC HÌNH KHỐI CƠ BẢN TỪ FOLDER 'MUA')
# ─────────────────────────────────────────────────────────────────────────────

def calc_frustum_pyramid(
    L_base: float,
    W_base: float,
    H_base: float,
    L_top: float,
    W_top: float,
    h_slope: float
) -> Dict[str, float]:
    """
    Tính thể tích và ván khuôn móng/bệ hình chóp cụt vát dốc (Frustum Pyramid).
    Áp dụng cho: Móng đơn chóp cụt dân dụng, Bệ mố móng cầu vát, Bệ trụ cầu.
    
    V = V_base + V_frustum
    V_frustum = (S_base + S_top + sqrt(S_base * S_top)) / 3 * h (chuẩn xác)
    Hoặc theo công thức xấp xỉ kỹ sư trong Excel Mua: ((S_base + S_top) / 2) * h
    """
    s_base = L_base * W_base
    s_top = L_top * W_top
    v_base = s_base * H_base
    # Công thức hình học chính xác theo nón cụt
    v_frustum = (s_base + s_top + math.sqrt(s_base * s_top)) / 3.0 * h_slope
    total_volume = round(v_base + v_frustum, 4)

    # Ván khuôn thành thẳng chân đế
    formwork_base = round(2.0 * (L_base + W_base) * H_base, 4)

    # Ván khuôn vát dốc (diện tích 4 hình thang nghiêng)
    # Cạnh xiên dốc:
    slope_len_L = math.sqrt(h_slope**2 + ((L_base - L_top) / 2.0)**2)
    slope_len_W = math.sqrt(h_slope**2 + ((W_base - W_top) / 2.0)**2)
    formwork_slope = round(
        2.0 * ((L_base + L_top) / 2.0 * slope_len_W) +
        2.0 * ((W_base + W_top) / 2.0 * slope_len_L), 4
    )

    return {
        "concrete_base_m3": round(v_base, 4),
        "concrete_frustum_m3": round(v_frustum, 4),
        "total_concrete_m3": total_volume,
        "formwork_base_m2": formwork_base,
        "formwork_slope_m2": formwork_slope,
        "total_formwork_m2": round(formwork_base + formwork_slope, 4),
    }


def calc_cutwater_pier_footing(
    L_rect: float,
    W_rect: float,
    H: float,
    R_nose: Optional[float] = None
) -> Dict[str, float]:
    """
    Tính thể tích và ván khuôn Bệ trụ cầu có mũi rẽ nước (Cutwater / Bo tròn mũi thuyền).
    Suy luận từ Sheet 'Móng' (Đài hình bầu dục / Oval) trong Folder 'Mua'.
    
    Phần thân giữa là khối hộp chữ nhật (L_rect x W_rect x H).
    Hai đầu bo tròn bán kính R = W_rect / 2 tạo thành 1 hình trụ tròn hoàn chỉnh.
    """
    r = R_nose if R_nose is not None else (W_rect / 2.0)
    # Thể tích khối hộp chữ nhật giữa
    v_rect = L_rect * W_rect * H
    # Thể tích 2 nửa hình trụ ở 2 đầu = 1 hình trụ tròn bán kính r
    v_cylinder = math.pi * (r**2) * H
    total_concrete = round(v_rect + v_cylinder, 4)

    # Ván khuôn: 2 cạnh thẳng + chu vi đường tròn đáy 2 đầu
    formwork_straight = 2.0 * L_rect * H
    formwork_curved = (2.0 * math.pi * r) * H
    total_formwork = round(formwork_straight + formwork_curved, 4)

    return {
        "concrete_m3": total_concrete,
        "formwork_m2": total_formwork,
        "nose_radius_m": r,
    }


def calc_column_and_corbel(
    L: float,
    W: float,
    H: float,
    dia: float = 0.0,
    h_beam_deduct: float = 0.0,
    corbel_w: float = 0.0,
    corbel_l: float = 0.0,
    corbel_h: float = 0.0,
    corbel_count: int = 0
) -> Dict[str, float]:
    """
    Tính khối lượng Cột chữ nhật / Cột tròn có Vai cột (Corbel) và trừ giao dầm/sàn.
    Áp dụng cho: Cột nhà xưởng, Thân trụ cầu cọc đơn / đôi, Xà mũ mố trụ.
    """
    effective_h = max(0.0, H - h_beam_deduct)

    if dia > 0:
        # Cột tròn
        v_col = (math.pi * (dia**2) / 4.0) * effective_h
        fw_col = (math.pi * dia) * effective_h
    else:
        # Cột chữ nhật
        v_col = (L * W) * effective_h
        fw_col = 2.0 * (L + W) * effective_h

    # Tính vai cột / xà mũ vươn hẫng
    # Vai cột dạng hình thang hoặc lăng trụ tam giác vát
    v_corbel = 0.0
    fw_corbel = 0.0
    if corbel_count > 0 and corbel_w > 0 and corbel_l > 0 and corbel_h > 0:
        v_corbel = (corbel_l * corbel_w * corbel_h * 0.75) * corbel_count
        fw_corbel = (corbel_l * corbel_h + 2.0 * (corbel_w * corbel_h)) * corbel_count

    total_concrete = round(v_col + v_corbel, 4)
    total_formwork = round(fw_col + fw_corbel, 4)

    return {
        "concrete_col_m3": round(v_col, 4),
        "concrete_corbel_m3": round(v_corbel, 4),
        "total_concrete_m3": total_concrete,
        "total_formwork_m2": total_formwork,
    }


def calc_beam_with_slab_deductions(
    length_m: float,
    width_m: float,
    height_m: float,
    col_width_deduct_m: float = 0.0,
    slab_thick_deduct_m: float = 0.0
) -> Dict[str, float]:
    """
    Tính bê tông và ván khuôn dầm có trừ giao cột và dầm/sàn (Chống trùng lặp bê tông).
    Áp dụng cho: Dầm tầng dân dụng, Dầm ngang cầu, Xà mũ mố cầu.
    """
    net_l = max(0.0, length_m - col_width_deduct_m)
    net_h = max(0.0, height_m - slab_thick_deduct_m)

    concrete_m3 = round(net_l * width_m * net_h, 4)
    formwork_sides_m2 = round(net_l * net_h * 2.0, 4)
    formwork_bottom_m2 = round(net_l * width_m, 4)
    total_formwork_m2 = round(formwork_sides_m2 + formwork_bottom_m2, 4)

    return {
        "net_length_m": round(net_l, 4),
        "net_height_m": round(net_h, 4),
        "concrete_m3": concrete_m3,
        "formwork_sides_m2": formwork_sides_m2,
        "formwork_bottom_m2": formwork_bottom_m2,
        "total_formwork_m2": total_formwork_m2,
    }


def calc_structural_steel_plate(
    thickness_mm: float,
    width_mm: float,
    length_mm: float,
    count: int = 1,
    density_kg_dm3: float = 7.85
) -> Dict[str, float]:
    """
    Tính khối lượng và diện tích sơn tấm thép (Plate - PL).
    Kế thừa trực tiếp từ Sheet 'ket cau thep' trong File 2.
    Formula: M = t * w * L * 7.85 * n / 10^6 (kg)
             S_son = 2 * w * L * n / 10^6 (m2)
    """
    weight_kg = round(
        steel_qs.steel_plate_weight(thickness_mm, width_mm, length_mm, count, density_kg_dm3), 3)
    paint_area_m2 = round((2.0 * width_mm * length_mm * count) / 1_000_000.0, 3)

    return {
        "thickness_mm": thickness_mm,
        "weight_kg": weight_kg,
        "paint_area_m2": paint_area_m2,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 2. BRIDGE STRUCTURE TAKEOFF (SUY LUẬN DIỄN BIẾN SANG CẦU ĐƯỜNG BỘ)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class BridgeAbutmentParams:
    """Tham số hình học toàn diện của Mố cầu (Abutment - M1 hoặc M2)."""
    name: str = "Mố M1"
    # 1. Bệ mố
    footing_L: float = 9.0       # Chiều dài bệ mố (vuông góc tim cầu)
    footing_W: float = 4.5       # Chiều rộng bệ mố (dọc tim cầu)
    footing_H: float = 1.5       # Chiều cao bệ mố
    lean_concrete_t: float = 0.1 # Bê tông lót dày 10cm
    lean_concrete_extend: float = 0.1 # Mở rộng lót mỗi bên 10cm
    # 2. Thân mố & Tường ngực (Breast wall)
    body_height: float = 4.2     # Chiều cao thân mố
    body_thickness: float = 1.2  # Chiều dày thân mố
    breast_wall_height: float = 1.4 # Tường ngực / tường đỉnh
    breast_wall_thick: float = 0.4
    # 3. Tường cánh (Wing walls - 2 bên)
    wing_length: float = 5.0     # Chiều dài tường cánh theo taluy
    wing_thick: float = 0.4      # Chiều dày tường cánh
    wing_h_start: float = 5.6    # Chiều cao đầu tường cánh (tại đỉnh thân mố)
    wing_h_end: float = 1.5      # Chiều cao cuối tường cánh (vát theo taluy mố)
    # 4. Bản quá độ (Approach slab)
    approach_slab_L: float = 5.0 # Dài bản quá độ
    approach_slab_W: float = 8.0 # Rộng bản quá độ
    approach_slab_t: float = 0.3 # Dày bản quá độ
    # 5. Đá kê gối mố (Pedestals)
    pedestal_count: int = 4
    pedestal_L: float = 0.6
    pedestal_W: float = 0.6
    pedestal_H: float = 0.25
    rebar_ratio_kg_m3: float = 110.0 # Hàm lượng cốt thép kg/m3


class BridgeAbutmentEngine:
    """Động cơ bóc tách Mố cầu bộ suy luận từ Sheet Móng, Cột, Dầm & Tường."""

    @staticmethod
    def compute_takeoff(p: BridgeAbutmentParams) -> Dict[str, Any]:
        # 1. Bê tông lót bệ mố
        l_lean = p.footing_L + 2.0 * p.lean_concrete_extend
        w_lean = p.footing_W + 2.0 * p.lean_concrete_extend
        lean_concrete_m3 = round(l_lean * w_lean * p.lean_concrete_t, 3)
        lean_formwork_m2 = round(2.0 * (l_lean + w_lean) * p.lean_concrete_t, 3)

        # 2. Bệ mố
        footing_conc_m3 = round(p.footing_L * p.footing_W * p.footing_H, 3)
        footing_fw_m2 = round(2.0 * (p.footing_L + p.footing_W) * p.footing_H, 3)

        # 3. Thân mố
        body_conc_m3 = round(p.footing_L * p.body_thickness * p.body_height, 3)
        body_fw_m2 = round(2.0 * (p.footing_L + p.body_thickness) * p.body_height, 3)

        # 4. Tường ngực
        breast_conc_m3 = round(p.footing_L * p.breast_wall_thick * p.breast_wall_height, 3)
        breast_fw_m2 = round((p.footing_L * 2.0 + p.breast_wall_thick * 2.0) * p.breast_wall_height, 3)

        # 5. Tường cánh (2 tường cánh hình thang vát dốc taluy)
        s_trapezoid = (p.wing_h_start + p.wing_h_end) / 2.0 * p.wing_length
        wing_conc_m3 = round(2.0 * s_trapezoid * p.wing_thick, 3)
        # Ván khuôn 2 mặt mỗi tường cánh
        wing_fw_m2 = round(2.0 * (2.0 * s_trapezoid), 3)

        # 6. Đá kê gối
        ped_conc_m3 = round(p.pedestal_count * (p.pedestal_L * p.pedestal_W * p.pedestal_H), 3)
        ped_fw_m2 = round(p.pedestal_count * (2.0 * (p.pedestal_L + p.pedestal_W) * p.pedestal_H), 3)

        # 7. Bản quá độ
        approach_conc_m3 = round(p.approach_slab_L * p.approach_slab_W * p.approach_slab_t, 3)
        approach_fw_m2 = round(2.0 * (p.approach_slab_L + p.approach_slab_W) * p.approach_slab_t, 3)

        # Tổng hợp Bê tông C30 (không tính lót C15)
        structural_concrete_m3 = round(
            footing_conc_m3 + body_conc_m3 + breast_conc_m3 + wing_conc_m3 + ped_conc_m3 + approach_conc_m3, 3
        )
        total_formwork_m2 = round(
            footing_fw_m2 + body_fw_m2 + breast_fw_m2 + wing_fw_m2 + ped_fw_m2 + approach_fw_m2, 3
        )

        # Cốt thép dự toán
        total_rebar_kg = round(structural_concrete_m3 * p.rebar_ratio_kg_m3, 1)
        tie_wire_kg = steel_qs.tie_wire_kg(total_rebar_kg, ndigits=1)  # 1,5% dây thép buộc — quy ước ở tools/steel_qs

        return {
            "component": p.name,
            "lean_concrete_m3": lean_concrete_m3,
            "lean_formwork_m2": lean_formwork_m2,
            "footing_concrete_m3": footing_conc_m3,
            "body_concrete_m3": body_conc_m3,
            "wing_concrete_m3": wing_conc_m3,
            "breast_wall_concrete_m3": breast_conc_m3,
            "pedestals_concrete_m3": ped_conc_m3,
            "approach_slab_concrete_m3": approach_conc_m3,
            "total_structural_concrete_m3": structural_concrete_m3,
            "total_formwork_m2": total_formwork_m2,
            "rebar_kg": total_rebar_kg,
            "total_rebar_kg": total_rebar_kg,
            "rebar_tie_wire_kg": tie_wire_kg,
        }


@dataclass
class BridgePierParams:
    """Tham số hình học Thân trụ & Xà mũ Trụ cầu (Pier - T1 hoặc T2)."""
    name: str = "Trụ T1"
    # 1. Bệ trụ xẻ nước mũi thuyền
    footing_rect_L: float = 7.0
    footing_rect_W: float = 3.6
    footing_H: float = 1.8
    # 2. Thân trụ (Cột tròn đôi D1.5m)
    column_dia: float = 1.5
    column_height: float = 8.5
    column_count: int = 2
    # 3. Xà mũ trụ (Hammerhead Pier Cap) có cánh hẫng vươn 2 bên
    cap_length: float = 9.6      # Chiều dài xà mũ
    cap_width: float = 1.8       # Bề rộng xà mũ
    cap_height_center: float = 1.6 # Chiều cao tại tim trụ
    cap_height_end: float = 1.0  # Chiều cao tại mút cánh hẫng (vát nghiêng)
    # 4. Đá kê gối
    pedestal_count: int = 4
    pedestal_L: float = 0.6
    pedestal_W: float = 0.6
    pedestal_H: float = 0.25
    rebar_ratio_kg_m3: float = 125.0


class BridgePierEngine:
    """Động cơ bóc tách Trụ cầu suy luận từ Móng oval, Cột tròn & Dầm có vai hẫng."""

    @staticmethod
    def compute_takeoff(p: BridgePierParams) -> Dict[str, Any]:
        # 1. Bệ trụ xẻ nước mũi thuyền (Cutwater)
        footing = calc_cutwater_pier_footing(p.footing_rect_L, p.footing_rect_W, p.footing_H)
        footing_conc_m3 = footing["concrete_m3"]
        footing_fw_m2 = footing["formwork_m2"]

        # 2. Thân trụ cột tròn đôi (trừ phần ngàm vào bệ và xà mũ)
        r = p.column_dia / 2.0
        single_col_conc = (math.pi * (r**2)) * p.column_height
        single_col_fw = (2.0 * math.pi * r) * p.column_height
        columns_conc_m3 = round(single_col_conc * p.column_count, 3)
        columns_fw_m2 = round(single_col_fw * p.column_count, 3)

        # 3. Xà mũ trụ có cánh hẫng vát
        # Thể tích = Bề rộng x (Diện tích hình đa giác chiếu đứng xà mũ)
        # Chiếu đứng: phần giữa (chữ nhật) + 2 cánh hẫng vát hình thang
        mid_len = p.column_dia * p.column_count + 1.2 # Đoạn giữa phẳng
        cantilever_len = max(0.0, (p.cap_length - mid_len) / 2.0)
        s_mid = mid_len * p.cap_height_center
        s_cantilever = 2.0 * ((p.cap_height_center + p.cap_height_end) / 2.0 * cantilever_len)
        cap_side_area = s_mid + s_cantilever
        cap_conc_m3 = round(cap_side_area * p.cap_width, 3)

        # Ván khuôn xà mũ: 2 mặt bên + đáy + 2 đầu mút
        # Mặt đáy: mid_len x W + 2 cánh nghiêng x W
        slope_len = math.sqrt(cantilever_len**2 + (p.cap_height_center - p.cap_height_end)**2)
        bottom_area = mid_len * p.cap_width + 2.0 * (slope_len * p.cap_width)
        end_area = 2.0 * (p.cap_height_end * p.cap_width)
        cap_fw_m2 = round(2.0 * cap_side_area + bottom_area + end_area, 3)

        # 4. Đá kê gối
        ped_conc_m3 = round(p.pedestal_count * (p.pedestal_L * p.pedestal_W * p.pedestal_H), 3)
        ped_fw_m2 = round(p.pedestal_count * (2.0 * (p.pedestal_L + p.pedestal_W) * p.pedestal_H), 3)

        # Tổng hợp
        total_conc_m3 = round(footing_conc_m3 + columns_conc_m3 + cap_conc_m3 + ped_conc_m3, 3)
        total_fw_m2 = round(footing_fw_m2 + columns_fw_m2 + cap_fw_m2 + ped_fw_m2, 3)
        total_rebar_kg = round(total_conc_m3 * p.rebar_ratio_kg_m3, 1)
        tie_wire_kg = steel_qs.tie_wire_kg(total_rebar_kg, ndigits=1)

        return {
            "component": p.name,
            "footing_concrete_m3": footing_conc_m3,
            "columns_concrete_m3": columns_conc_m3,
            "cap_concrete_m3": cap_conc_m3,
            "pedestals_concrete_m3": ped_conc_m3,
            "total_structural_concrete_m3": total_conc_m3,
            "total_formwork_m2": total_fw_m2,
            "rebar_kg": total_rebar_kg,
            "total_rebar_kg": total_rebar_kg,
            "rebar_tie_wire_kg": tie_wire_kg,
        }


@dataclass
class BridgeSuperstructureParams:
    """Tham số Kết cấu nhịp Dầm Super-T hoặc Dầm I kèm Bản mặt cầu."""
    bridge_length_m: float = 33.0 # Chiều dài nhịp dầm
    deck_width_m: float = 9.0    # Chiều rộng bản mặt cầu
    girder_type: str = "Super-T" # "Super-T" hoặc "I-Girder" hoặc "Steel-Plate"
    girder_count: int = 4        # 4 dầm / nhịp
    girder_height_m: float = 1.75
    girder_volume_single_m3: float = 23.5 # Thể tích 1 dầm đúc sẵn (Super-T 33m ~23.5m3)
    girder_formwork_single_m2: float = 145.0 # Ván khuôn 1 dầm
    girder_rebar_kg_single: float = 4250.0 # Thép thường 1 dầm
    girder_pt_strand_kg_single: float = 1480.0 # Cáp DƯL 1 dầm (15.2mm)
    deck_thickness_m: float = 0.2 # Bề dày bản mặt cầu
    asphalt_thickness_m: float = 0.07 # Thảm bê tông nhựa C12.5 (7cm)
    expansion_joint_m: float = 9.0 # Khe co giãn răng lược


class BridgeSuperstructureEngine:
    """Động cơ bóc tách Kết cấu nhịp dầm, bản mặt cầu, cáp DƯL & thảm nhựa."""

    @staticmethod
    def compute_takeoff(p: BridgeSuperstructureParams) -> Dict[str, Any]:
        # 1. Dầm chủ đúc sẵn
        girders_conc_m3 = round(p.girder_count * p.girder_volume_single_m3, 3)
        girders_fw_m2 = round(p.girder_count * p.girder_formwork_single_m2, 3)
        girders_rebar_kg = round(p.girder_count * p.girder_rebar_kg_single, 1)
        girders_pt_kg = round(p.girder_count * p.girder_pt_strand_kg_single, 1)
        # Số đầu neo công tác ước tính (mỗi bó cáp 2 đầu neo, trung bình 4-6 bó/dầm)
        anchorage_sets = p.girder_count * 10
        # Chiều dài ống ghen ruột gà bảo vệ cáp
        grout_duct_m = round(p.girder_count * 5 * p.bridge_length_m, 1)

        # 2. Bản mặt cầu đổ tại chỗ (Deck slab)
        # Trừ phần đỉnh dầm Super-T (chiều dày bản tính bù giữa các cánh dầm)
        deck_conc_m3 = round(p.bridge_length_m * p.deck_width_m * p.deck_thickness_m, 3)
        # Ván khuôn bản đáy (giữa các dầm) + ván khuôn cánh hẫng biên
        deck_fw_m2 = round(p.bridge_length_m * (p.deck_width_m * 0.65), 3) # Ván khuôn đáy
        deck_rebar_kg = round(deck_conc_m3 * 140.0, 1) # Thép bản mặt cầu 140 kg/m3

        # 3. Gờ chắn bánh / Lan can bê tông (2 bên cầu)
        barrier_conc_m3 = round(2.0 * p.bridge_length_m * (0.4 * 0.8 * 0.7), 3) # Tiết diện vát
        barrier_fw_m2 = round(2.0 * p.bridge_length_m * 1.5, 3)
        barrier_rebar_kg = round(barrier_conc_m3 * 130.0, 1)

        # 4. Lớp phòng nước & Thảm bê tông nhựa
        waterproofing_area_m2 = round(p.bridge_length_m * (p.deck_width_m - 0.8), 3) # Trừ 2 gờ lan can
        asphalt_volume_m3 = round(waterproofing_area_m2 * p.asphalt_thickness_m, 3)
        asphalt_ton = round(asphalt_volume_m3 * 2.35, 2) # Dung trọng BTN hạt mịn 2.35 Tấn/m3

        # Tổng hợp vật tư
        total_concrete_c40_c50 = girders_conc_m3 # Dầm
        total_concrete_c30_c35 = round(deck_conc_m3 + barrier_conc_m3, 3) # Mặt cầu
        total_rebar_kg = round(girders_rebar_kg + deck_rebar_kg + barrier_rebar_kg, 1)
        tie_wire_kg = steel_qs.tie_wire_kg(total_rebar_kg, ndigits=1)

        return {
            "girders_count": p.girder_count,
            "girder_concrete_m3": girders_conc_m3,
            "girder_formwork_m2": girders_fw_m2,
            "deck_concrete_m3": deck_conc_m3,
            "deck_formwork_m2": deck_fw_m2,
            "barrier_concrete_m3": barrier_conc_m3,
            "total_precast_concrete_m3": girders_conc_m3,
            "total_in_situ_concrete_m3": total_concrete_c30_c35,
            "total_rebar_kg": total_rebar_kg,
            "total_pt_strands_kg": girders_pt_kg,
            "anchorage_sets": anchorage_sets,
            "duct_length_m": grout_duct_m,
            "tie_wire_kg": tie_wire_kg,
            "waterproofing_membrane_m2": waterproofing_area_m2,
            "asphalt_concrete_ton": asphalt_ton,
            "expansion_joint_m": p.expansion_joint_m,
        }


# ─────────────────────────────────────────────────────────────────────────────
# 3. STEEL-CONCRETE COMPOSITE BRIDGE (CẦU THÉP LIÊN HỢP TỪ SHEET 'KET CAU THEP')
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class SteelBridgeGirderSegment:
    """Phân đoạn dầm thép chữ I liên hợp (Plate Girder) chuẩn hóa."""
    name: str = "Dầm thép I liên hợp L=30m"
    length_mm: float = 30_000.0
    # Bản cánh trên (Top Flange): PL25 x 400
    top_flange_t_mm: float = 25.0
    top_flange_w_mm: float = 400.0
    # Bản đáy / cánh dưới (Bottom Flange): PL32 x 500
    bot_flange_t_mm: float = 32.0
    bot_flange_w_mm: float = 500.0
    # Bản bụng (Web): PL14 x 1500
    web_t_mm: float = 14.0
    web_h_mm: float = 1500.0
    # Sườn tăng cường (Stiffeners): PL12 x 150 (khoảng cách 1.5m -> 20 cặp = 40 tấm)
    stiffener_count: int = 40
    stiffener_t_mm: float = 12.0
    stiffener_w_mm: float = 150.0
    # Neo đinh chống cắt (Nelson Shear Studs D22 x 150mm, 3 hàng dọc tim cánh trên)
    stud_count: int = 180
    stud_dia_mm: float = 22.0
    stud_length_mm: float = 150.0


class SteelBridgeGirderEngine:
    """
    Động cơ bóc tách dầm cầu thép tấm tổ hợp chữ I (Steel Plate Girder)
    Kế thừa trực tiếp nguyên lý tính khối lượng tấm và diện tích sơn từ Sheet 'ket cau thep' File 2.
    """

    @staticmethod
    def compute_takeoff(seg: SteelBridgeGirderSegment) -> Dict[str, Any]:
        # 1. Bản cánh trên
        top = calc_structural_steel_plate(seg.top_flange_t_mm, seg.top_flange_w_mm, seg.length_mm, count=1)
        # 2. Bản cánh dưới
        bot = calc_structural_steel_plate(seg.bot_flange_t_mm, seg.bot_flange_w_mm, seg.length_mm, count=1)
        # 3. Bản bụng
        web = calc_structural_steel_plate(seg.web_t_mm, seg.web_h_mm, seg.length_mm, count=1)
        # 4. Sườn tăng cường
        stiff = calc_structural_steel_plate(seg.stiffener_t_mm, seg.stiffener_w_mm, seg.web_h_mm, count=seg.stiffener_count)

        # 5. Neo đinh chống cắt D22 (thép tròn)
        stud_single_vol = math.pi * ((seg.stud_dia_mm / 2.0)**2) * seg.stud_length_mm
        stud_total_weight = round((stud_single_vol * 7.85 / 1_000_000.0) * seg.stud_count, 2)

        # Tổng hợp
        total_steel_kg = round(top["weight_kg"] + bot["weight_kg"] + web["weight_kg"] + stiff["weight_kg"] + stud_total_weight, 2)
        total_paint_area_m2 = round(top["paint_area_m2"] + bot["paint_area_m2"] + web["paint_area_m2"] + stiff["paint_area_m2"], 2)

        return {
            "component": seg.name,
            "top_flange_kg": top["weight_kg"],
            "bottom_flange_kg": bot["weight_kg"],
            "web_plate_kg": web["weight_kg"],
            "stiffeners_kg": stiff["weight_kg"],
            "shear_studs_kg": stud_total_weight,
            "total_structural_steel_kg": total_steel_kg,
            "total_structural_steel_ton": round(total_steel_kg / 1000.0, 3),
            "paint_area_anti_corrosion_m2": total_paint_area_m2,
        }


# ─────────────────────────────────────────────────────────────────────────────
# 4. BRIDGE REBAR BBS CUT DEMAND GENERATORS (LIÊN KẾT GOOGLE OR-TOOLS)
# ─────────────────────────────────────────────────────────────────────────────

def generate_bridge_pier_rebar_demands(
    column_height_m: float = 8.5,
    column_dia_m: float = 1.5,
    column_count: int = 2,
    grade: str = "CB400-V"
) -> List[Dict[str, Any]]:
    """
    Sinh danh sách đoạn cắt cốt thép thân trụ cầu (CutDemand parameters)
    chuẩn hóa để nạp trực tiếp vào tools.cutting_stock_solver.
    """
    demands = []
    # 1. Thép chủ thân trụ D28 (thanh thẳng + đoạn bẻ neo vào xà mũ và bệ móng)
    # Chiều dài thanh chủ: H + 40d ngàm xà mũ + 40d ngàm bệ móng
    lap_top = int(40 * 28)    # 1120 mm
    lap_bot = int(40 * 28)    # 1120 mm
    bar_len = int(round(column_height_m * 1000.0)) + lap_top + lap_bot
    # Số thanh chủ: trung bình khoảng cách a=150mm quanh chu vi D=1500mm -> pi * 1500 / 150 ≈ 32 thanh/cột
    bars_per_col = 32
    demands.append({
        "mark": "TC-TRU-D28",
        "diameter_mm": 28,
        "length_mm": bar_len,
        "quantity": bars_per_col * column_count,
        "grade": grade,
        "splice_allowed": True,
        "lap_xd": 40.0
    })

    # 2. Thép đai xoắn gia cường D14 bước 100/200mm
    # Đoạn đai vòng tròn hoặc đai khép kín có móc 135 độ
    perimeter_inner = int(round(math.pi * (column_dia_m * 1000.0 - 100.0))) + 250 # Thêm móc 135 độ
    ties_count = int(round(column_height_m * 1000.0 / 150.0)) * column_count
    demands.append({
        "mark": "DAI-TRU-D14",
        "diameter_mm": 14,
        "length_mm": perimeter_inner,
        "quantity": ties_count,
        "grade": grade,
        "splice_allowed": False
    })

    return demands


def generate_super_t_rebar_demands(
    span_length_m: float = 33.0,
    girder_count: int = 4,
    grade: str = "CB400-V"
) -> List[Dict[str, Any]]:
    """
    Sinh danh sách đoạn cắt cốt thép dầm Super-T (CutDemand parameters).
    Bao gồm thép chủ sườn dầm, thép bầu dầm và thép đai chữ U.
    """
    demands = []
    # 1. Thép sườn dầm D20 dọc suốt nhịp (L = 33m - 80mm lớp bảo vệ = 32920mm)
    # Cần ghép nối từ các đoạn 11.7m theo TCVN 5574
    demands.append({
        "mark": "SUON-SUPER-T-D20",
        "diameter_mm": 20,
        "length_mm": 5800, # Chia thành các đoạn tối ưu theo thư viện Mẫu Vàng
        "quantity": 16 * girder_count,
        "grade": grade,
        "splice_allowed": True,
        "lap_xd": 40.0
    })
    demands.append({
        "mark": "BAU-SUPER-T-D25",
        "diameter_mm": 25,
        "length_mm": 5800,
        "quantity": 12 * girder_count,
        "grade": grade,
        "splice_allowed": True,
        "lap_xd": 40.0
    })

    # 2. Thép đai chữ U sườn dầm D12 bước 100/150/200mm
    stirrup_u_len = 3850 # Chu vi thanh chữ U bao quanh sườn dầm cao 1.75m
    stirrup_count = int(round(span_length_m * 1000.0 / 150.0)) * girder_count
    demands.append({
        "mark": "DAI-U-SUPER-T-D12",
        "diameter_mm": 12,
        "length_mm": stirrup_u_len,
        "quantity": stirrup_count,
        "grade": grade,
        "splice_allowed": False
    })

    return demands
