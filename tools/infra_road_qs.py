# -*- coding: utf-8 -*-
"""
INFRA ROAD QS — đo bóc hạ tầng đường (áo đường & xử lý nền)

Chuẩn hóa từ sheet 'Hạ tầng- Đường' của bảng tính khối lượng QS chuyên nghiệp. Tái tạo đúng một
đoạn đường theo loại kết cấu: R1 (asphalt tải nhẹ), R2 (asphalt tải nặng), R3 (bê tông).

Mỗi đoạn tính: base (cấp phối, có hệ số lu lèn), lớp asphalt / nhũ tương theo loại, bê tông mặt
đường + nilon + cốt thép (chỉ R3), bóc nền, cát san lấp (có hệ số lu lèn), diện tích lu lèn.
Khớp các ô đã tính sẵn trong hồ sơ thật (xem tests/test_infra_road_qs_golden.py).

Quy ước (đúng hồ sơ):
  • Base (m³) = diện_tích × base_dày × hệ_số_lu_lèn.
  • Bê tông R3 (m³) = diện_tích × bê_tông_dày × hệ_số_hao_hụt.
  • H bóc nền = −cốt_đường_TB + base_dày + bê_tông_dày + cốt_hiện_trạng + phần_cộng_thêm.
  • Cát san lấp (m³) = diện_tích × chiều_dày_cát × hệ_số_lu_lèn_cát.
  • Cốt thép R3 (kg/m²) = 1000·D²/162·10⁶·(1000/@)·2·lớp / 10⁹.

Thuần Python, không phụ thuộc ngoài.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict


@dataclass
class RoadSegmentParams:
    road_type: str                  # "R1" | "R2" | "R3"
    area_m2: float                  # J — diện tích mặt đường
    base_depth_m: float             # H — chiều dày lớp base (cấp phối)
    concrete_depth_m: float = 0.0   # I — chiều dày bê tông mặt (R3)
    compaction_base: float = 1.0    # K — hệ số lu lèn chặt base
    waste_concrete: float = 1.0     # L — hệ số hao hụt bê tông
    count: int = 1                  # E
    ground_level_m: float = 0.0     # AA — cốt hiện trạng
    road_avg_level_m: float = 0.0   # AB — cốt đường trung bình
    strip_extra_m: float = 0.3      # AC10 — phần cộng thêm khi bóc nền
    sand_fill_depth_m: float = 0.3  # AE10 — chiều dày cát san lấp (mặc định = strip_extra)
    sand_compaction: float = 1.35   # Z — hệ số lu lèn cát
    rebar_d_mm: float = 0.0         # T (R3)
    rebar_spacing_mm: float = 0.0   # U
    rebar_layers: int = 0           # V
    rebar_factor: float = 1.15      # X


def road_segment_takeoff(p: RoadSegmentParams) -> Dict[str, float]:
    J, H, I, E = p.area_m2, p.base_depth_m, p.concrete_depth_m, p.count
    t = p.road_type
    base = J * H * E * p.compaction_base
    asphalt_light = J * E if t == "R1" else 0.0
    asphalt_heavy = J * E if t == "R2" else 0.0
    concrete = J * I * E * p.waste_concrete if t == "R3" else 0.0
    nilon = J * E if t == "R3" else 0.0

    if t == "R3" and p.rebar_layers > 0 and p.rebar_spacing_mm > 0:
        rebar_unit = (1000 * p.rebar_d_mm ** 2 / 162 * 1e6 * (1000 / p.rebar_spacing_mm)
                      * 2 * p.rebar_layers) / 1e9
    else:
        rebar_unit = 0.0
    rebar_kg = J * rebar_unit * p.rebar_factor * E

    strip_h = -p.road_avg_level_m + H + I + p.ground_level_m + p.strip_extra_m if J else 0.0
    strip = J * strip_h * E
    sand_fill = J * p.sand_fill_depth_m * E * p.sand_compaction
    compaction = J * E

    return {
        "base_m3": round(base, 8),
        "asphalt_light_m2": round(asphalt_light, 8),
        "asphalt_heavy_m2": round(asphalt_heavy, 8),
        "emulsion_light_m2": round(asphalt_light, 8),
        "emulsion_heavy_m2": round(asphalt_heavy, 8),
        "concrete_m3": round(concrete, 8),
        "nilon_m2": round(nilon, 8),
        "rebar_unit_kg_m2": round(rebar_unit, 8),
        "rebar_kg": round(rebar_kg, 8),
        "strip_height_m": round(strip_h, 8),
        "strip_subgrade_m3": round(strip, 8),
        "sand_fill_m3": round(sand_fill, 8),
        "compaction_m2": round(compaction, 8),
    }
