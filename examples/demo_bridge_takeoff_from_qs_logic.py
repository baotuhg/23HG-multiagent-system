# -*- coding: utf-8 -*-
"""
DEMO: BÓC TÁCH HÌNH HỌC CẦU ĐƯỜNG BỘ TỪ ĐỘNG CƠ DIỄN BIẾN SUY LUẬN
Dự án mẫu: Cầu dầm Super-T 33m — 2 Mố chữ U + 1 Trụ xẻ nước mũi thuyền
Tuân thủ nguyên tắc Pure Python, Zero LLM Math, TCVN 11823:2017 & TT 38/2026/TT-BXD.
"""

from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from tools.civil_and_bridge_takeoff_engine import (
    BridgeAbutmentParams,
    BridgeAbutmentEngine,
    BridgePierParams,
    BridgePierEngine,
    BridgeSuperstructureParams,
    BridgeSuperstructureEngine,
    SteelBridgeGirderSegment,
    SteelBridgeGirderEngine,
    generate_bridge_pier_rebar_demands,
    generate_super_t_rebar_demands,
)
from tools.cutting_stock_solver import CuttingStockSolver, CutDemand


def main():
    print("=" * 75)
    print("🌉 HỆ THỐNG BÓC TÁCH HÌNH HỌC CẦU ĐƯỜNG BỘ — SUY LUẬN DIỄN BIẾN TỪ THƯ VIỆN QS")
    print("   (Kế thừa từ nguyên lý hình học & kết cấu thép folder 'Mua')")
    print("=" * 75)

    # 1. Bóc tách Mố Cầu M1 & M2 (2 Mố chữ U)
    p_abut = BridgeAbutmentParams(name="Mố M1 (Mố chữ U)")
    takeoff_abut = BridgeAbutmentEngine.compute_takeoff(p_abut)

    print("\n[1] 🏛️ KẾT QUẢ BÓC TÁCH MỐ CẦU (1 MỐ):")
    print(f"    • Bê tông lót bệ mố (C15)       : {takeoff_abut['lean_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông bệ mố (C30)          : {takeoff_abut['footing_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông thân mố (C30)        : {takeoff_abut['body_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông tường cánh vát taluy : {takeoff_abut['wing_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông tường ngực           : {takeoff_abut['breast_wall_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông bản quá độ           : {takeoff_abut['approach_slab_concrete_m3']:>8.2f} m3")
    print(f"    -> TỔNG BÊ TÔNG KẾT CẤU MỐ      : {takeoff_abut['total_structural_concrete_m3']:>8.2f} m3")
    print(f"    -> TỔNG VÁN KHUÔN MỐ            : {takeoff_abut['total_formwork_m2']:>8.2f} m2")
    print(f"    -> TỔNG CỐT THÉP DỰ TOÁN        : {takeoff_abut['total_rebar_kg']:>8.1f} kg")
    print(f"    -> DÂY THÉP BUỘC CÔNG TRƯỜNG   : {takeoff_abut['rebar_tie_wire_kg']:>8.1f} kg (1.5%)")

    # 2. Bóc tách Trụ Cầu T1 (Thân cột đôi + Bệ xẻ nước mũi thuyền)
    p_pier = BridgePierParams(name="Trụ T1 (Bệ xẻ nước mũi thuyền + 2 Cột D1.5m)")
    takeoff_pier = BridgePierEngine.compute_takeoff(p_pier)

    print("\n[2] ⚓ KẾT QUẢ BÓC TÁCH TRỤ CẦU T1:")
    print(f"    • Bê tông bệ xẻ nước mũi thuyền: {takeoff_pier['footing_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông thân 2 cột tròn D1.5m: {takeoff_pier['columns_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông xà mũ vươn cánh hẫng : {takeoff_pier['cap_concrete_m3']:>8.2f} m3")
    print(f"    -> TỔNG BÊ TÔNG KẾT CẤU TRỤ     : {takeoff_pier['total_structural_concrete_m3']:>8.2f} m3")
    print(f"    -> TỔNG VÁN KHUÔN TRỤ           : {takeoff_pier['total_formwork_m2']:>8.2f} m2")
    print(f"    -> TỔNG CỐT THÉP TRỤ            : {takeoff_pier['total_rebar_kg']:>8.1f} kg")
    print(f"    -> DÂY THÉP BUỘC TRỤ           : {takeoff_pier['rebar_tie_wire_kg']:>8.1f} kg (1.5%)")

    # 3. Bóc tách Kết cấu nhịp Dầm Super-T 33m & Bản mặt cầu
    p_super = BridgeSuperstructureParams(bridge_length_m=33.0, deck_width_m=9.0, girder_count=4)
    takeoff_super = BridgeSuperstructureEngine.compute_takeoff(p_super)

    print("\n[3] 🚀 KẾT QUẢ BÓC TÁCH KẾT CẤU NHỊP (4 DẦM SUPER-T 33M):")
    print(f"    • Bê tông 4 dầm đúc sẵn (C50)   : {takeoff_super['girder_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông bản mặt cầu (C35)     : {takeoff_super['deck_concrete_m3']:>8.2f} m3")
    print(f"    • Bê tông gờ lan can            : {takeoff_super['barrier_concrete_m3']:>8.2f} m3")
    print(f"    • Cáp dự ứng lực DƯL 15.2mm     : {takeoff_super['total_pt_strands_kg']:>8.1f} kg")
    print(f"    • Đầu neo công tác DƯL          : {takeoff_super['anchorage_sets']:>8d} bộ")
    print(f"    • Ống ghen ruột gà bơm vữa      : {takeoff_super['duct_length_m']:>8.1f} md")
    print(f"    • Màng phòng nước mặt cầu       : {takeoff_super['waterproofing_membrane_m2']:>8.2f} m2")
    print(f"    • Thảm bê tông nhựa C12.5 (7cm) : {takeoff_super['asphalt_concrete_ton']:>8.2f} Tấn")
    print(f"    • Khe co giãn răng lược thép    : {takeoff_super['expansion_joint_m']:>8.1f} md")

    # 4. Bóc tách Dầm Cầu Thép Chữ I Liên Hợp (Suy luận từ Sheet 'ket cau thep')
    p_steel = SteelBridgeGirderSegment(name="Dầm Thép Chữ I Liên Hợp Nhịp 30m")
    takeoff_steel = SteelBridgeGirderEngine.compute_takeoff(p_steel)

    print("\n[4] 🏗️ KẾT QUẢ BÓC TÁCH DẦM CẦU THÉP LIÊN HỢP (PLATE GIRDER L=30M):")
    print(f"    • Bản cánh trên (PL25x400)      : {takeoff_steel['top_flange_kg']:>8.1f} kg")
    print(f"    • Bản cánh dưới (PL32x500)      : {takeoff_steel['bottom_flange_kg']:>8.1f} kg")
    print(f"    • Bản bụng (PL14x1500)          : {takeoff_steel['web_plate_kg']:>8.1f} kg")
    print(f"    • Sườn tăng cường (40 tấm PL12) : {takeoff_steel['stiffeners_kg']:>8.1f} kg")
    print(f"    • Đinh neo chống cắt Nelson D22 : {takeoff_steel['shear_studs_kg']:>8.1f} kg (180 chiếc)")
    print(f"    -> TỔNG KHỐI LƯỢNG THÉP DẦM     : {takeoff_steel['total_structural_steel_ton']:>8.3f} Tấn")
    print(f"    -> DIỆN TÍCH SƠN CHỐNG GỈ 3 LỚP : {takeoff_steel['paint_area_anti_corrosion_m2']:>8.2f} m2")

    # 5. Tối ưu Cắt Thép BBS Cầu Đường với Google OR-Tools CP-SAT
    print("\n[5] ⚡ TỔNG HỢP CỐT THÉP VÀ TỐI ƯU CẮT 1D (GOOGLE OR-TOOLS CP-SAT):")
    pier_bbs = generate_bridge_pier_rebar_demands()
    super_t_bbs = generate_super_t_rebar_demands()

    # Tối ưu thép sườn dầm D20
    d20_demands = [CutDemand(**d) for d in super_t_bbs if d["diameter_mm"] == 20]
    solver = CuttingStockSolver()
    sol_d20 = solver.solve(d20_demands)

    print(f"    • Thép sườn Dầm Super-T (Phi 20): Cần {sol_d20.total_bars_needed} cây 11.7m")
    print(f"      - Tỷ lệ hao hụt đề-xê xưởng   : {sol_d20.waste_ratio_pct:.2f}% (Mục tiêu < 1.5%)")
    print(f"      - Trạng thái giải toán        : {sol_d20.status}")

    print("\n" + "=" * 75)
    print("ℹ  Kết quả ước tính sơ bộ với tham số mặc định — chưa đối chiếu bản vẽ cầu thật.")
    print("   Danh mục thép là giả định điển hình; khi làm thật dùng BBS thật (--phase rebar --bbs).")
    print("=" * 75)


if __name__ == "__main__":
    main()
