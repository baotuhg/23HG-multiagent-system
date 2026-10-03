# -*- coding: utf-8 -*-
"""
INFRA CHANNEL QS — đo bóc mương hộp hạ tầng (dòng đáy mương)

Chuẩn hóa từ sheet 'Mương' của bảng tính khối lượng QS chuyên nghiệp. Hàm dưới đây tái tạo đúng
DÒNG ĐÁY MƯƠNG (đá dăm nền, bê tông lót, bê tông đáy, ván khuôn, nilon, chống thấm, cọc tre, đào
đắp toàn tuyến mương, và cốt thép đáy) — các ô đã tính sẵn trong hồ sơ thật.

Thành mương và nắp mương trong hồ sơ được lập thành các dòng cấu kiện riêng (khối lăng trụ: dài ×
dày × cao); người dùng cộng thêm như các đầu việc độc lập. Mô-đun này tập trung phần đã kiểm chứng
đến từng m³ (xem tests/test_infra_channel_qs_golden.py).

Quy ước (đúng hồ sơ):
  • Bề rộng phủ bì đáy: width_total = W + thành_dày×2 + mở_rộng_lót.
  • H đào = −(cốt_dòng_nước − lót_dày − đáy_dày − đá_dăm_dày − cốt_hiện_trạng).
  • Mở rộng đỉnh = H_đào × hệ_số_mái(H_đào) + mở_rộng_chân.
  • Đắp = Đào − đá dăm − bê tông lót − bê tông đáy − phần lòng mương dưới cao trình đáy.
  • Cốt thép đáy (kg/m²): 1000·D²/162·10⁶·(1000/@)·2·lớp / 10⁹, với D(mm), @ (khoảng cách mm).
  Hệ số mái taluy theo H đào (m): ≤1,5→0,25; (1,5;5)→0,67; ≥5→0,85.

Thuần Python, không phụ thuộc ngoài, không làm tròn trung gian (trừ ROUND số cọc như hồ sơ).
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import Dict

PILE_DENSITY = 20.0      # mật độ cọc tre mương (Z10)


def _round_half_up(value: float, ndigits: int = 0) -> float:
    q = Decimal(1).scaleb(-ndigits)
    r = Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP)
    return int(r) if ndigits <= 0 else float(r)


def slope_factor(dig_height_m: float) -> float:
    if dig_height_m >= 5:
        return 0.85
    return 0.25 if dig_height_m <= 1.5 else 0.67


@dataclass
class ChannelBaseParams:
    length_m: float                 # E — chiều dài tuyến mương
    width_m: float                  # F — bề rộng lòng mương
    height_m: float                 # G — chiều cao mương (chỉ để tham khảo)
    lean_t_m: float = 0.05          # H — bê tông lót dày
    lean_widen_m: float = None      # I — mở rộng lót (mặc định = lót_dày×2)
    base_slab_t_m: float = 0.15     # J — đáy dày
    wall_t_m: float = 0.145         # K — thành dày
    cover_t_m: float = 0.15         # L — nắp dày (không dùng ở dòng đáy)
    aggregate_base_t_m: float = 0.0  # M — đá dăm nền dày
    waste_base: float = 1.0         # N — hao hụt base
    waste_concrete: float = 1.0     # O — hao hụt bê tông
    count: int = 1                  # D — số lượng
    pile_length_m: float = 0.0      # Y — chiều dài cọc tre
    ground_level_m: float = 0.0     # AB — cốt hiện trạng
    water_avg_level_m: float = 0.0  # AC — cốt trung bình dòng nước (thường âm)
    base_widen_m: float = 0.5       # AD — mở rộng chân
    rebar_d_mm: float = 0.0         # AI
    rebar_spacing_mm: float = 0.0   # AJ
    rebar_layers: int = 0           # AK
    rebar_factor: float = 1.0       # AM

    def _lean_widen(self) -> float:
        return self.lean_t_m * 2 if self.lean_widen_m is None else self.lean_widen_m


def channel_base_takeoff(p: ChannelBaseParams) -> Dict[str, float]:
    E, F, H, J, K = p.length_m, p.width_m, p.lean_t_m, p.base_slab_t_m, p.wall_t_m
    M, N, O, D = p.aggregate_base_t_m, p.waste_base, p.waste_concrete, p.count
    I = p._lean_widen()
    width_total = F + K * 2 + I
    base_out = F + K * 2

    aggregate = E * width_total * M * D * N
    lean_concrete = E * width_total * H * D * O
    base_concrete = E * base_out * J * D * O
    lean_formwork = E * 2 * H * D
    formwork = E * 2 * J * D
    nilon = E * width_total * D
    waterproof_in = E * F * D

    pile_count = int(_round_half_up(width_total * E * PILE_DENSITY, 0)) if p.pile_length_m > 0 else 0
    pile_md = p.pile_length_m * pile_count * D

    dig_h = -(p.water_avg_level_m - H - J - M - p.ground_level_m)
    top_widen = dig_h * slope_factor(dig_h) + p.base_widen_m
    excavation = ((base_out + p.base_widen_m * 2) + (base_out + top_widen * 2)) / 2 * dig_h * E * D
    backfill = excavation - aggregate - lean_concrete - base_concrete \
        - E * base_out * (dig_h - H - J - M) * D

    if p.rebar_layers > 0 and p.rebar_spacing_mm > 0:
        rebar_unit = (1000 * p.rebar_d_mm ** 2 / 162 * 1e6 * (1000 / p.rebar_spacing_mm)
                      * 2 * p.rebar_layers) / 1e9
    else:
        rebar_unit = 0.0
    rebar_kg = E * base_out * D * rebar_unit * p.rebar_factor

    return {
        "aggregate_base_m3": round(aggregate, 8),
        "lean_concrete_m3": round(lean_concrete, 8),
        "base_concrete_m3": round(base_concrete, 8),
        "lean_formwork_m2": round(lean_formwork, 8),
        "formwork_m2": round(formwork, 8),
        "nilon_m2": round(nilon, 8),
        "waterproof_in_m2": round(waterproof_in, 8),
        "pile_count": pile_count,
        "pile_md": round(pile_md, 6),
        "dig_height_m": round(dig_h, 8),
        "top_widen_m": round(top_widen, 8),
        "excavation_m3": round(excavation, 8),
        "backfill_m3": round(backfill, 8),
        "rebar_unit_kg_m2": round(rebar_unit, 8),
        "rebar_kg": round(rebar_kg, 8),
    }
