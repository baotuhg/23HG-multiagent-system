# -*- coding: utf-8 -*-
"""
INFRA FENCE QS — đo bóc hàng rào (móng trụ + trụ/giằng, và tường xây/trát)

Chuẩn hóa từ sheet 'Hàng rào' của bảng tính khối lượng QS chuyên nghiệp. Hai nhóm cấu kiện:
  • Móng trụ: bê tông lót + ván khuôn lót, bê tông móng (hộp + chóp cụt vát), ván khuôn, trụ/giằng
    (phân nhóm theo bề dày), đào/đắp/tân đất.
  • Tường xây + trát + sơn bả (trừ phần trụ chiếm chỗ).

Khớp các ô đã tính sẵn trong hồ sơ thật (xem tests/test_infra_fence_qs_golden.py).

Quy ước (đúng hồ sơ):
  • Bê tông móng = (L·W·H1 + (L·W + L1·W1)/2·h_vát) × SL  (hộp + chóp cụt vát).
  • Chiều cao trụ = H_móng_đỉnh − H1 − h_vát; phân nhóm bề dày: <0,2 / ≥0,2 / ≥0,3 m.
  • H đào = −H_đáy_móng + cốt_hiện_trạng + 0,05.
  • Tường: xây = (dài − trừ_trụ)·dày·cao; trát/sơn = (dài·cao·2 + dài·dày + cao·dày·2).

Thuần Python, không phụ thuộc ngoài.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict


@dataclass
class FencePostFootingParams:
    L_m: float                      # E — dài móng
    W_m: float                      # F — rộng móng
    footing_box_h_m: float          # G — H1 (phần hộp)
    footing_slope_h_m: float        # H — h vát
    top_L_m: float                  # I — W1 (đỉnh vát, dài)
    top_W_m: float                  # J — L1 (đỉnh vát, rộng)
    post_w_m: float                 # K — bề rộng trụ
    post_l_m: float                 # L — bề dày trụ
    foundation_to_top_m: float      # M — H móng → đỉnh trụ
    count: int = 1                  # D
    waste: float = 1.0              # O
    ground_level_m: float = -0.2    # AC
    base_widen_m: float = 0.5       # AD
    top_widen_m: float = 1.0        # AE
    footing_bottom_level_m: float = -1.0   # AF (âm)


def fence_post_footing_takeoff(p: FencePostFootingParams) -> Dict[str, float]:
    E, F, G, H, I, J = p.L_m, p.W_m, p.footing_box_h_m, p.footing_slope_h_m, p.top_L_m, p.top_W_m
    K, Lp, M, D, O = p.post_w_m, p.post_l_m, p.foundation_to_top_m, p.count, p.waste

    lean_concrete = (E + 0.1) * (F + 0.1) * 0.05 * D * O
    lean_formwork = ((E + 0.1) + (F + 0.1)) * 2 * 0.05 * D
    footing_concrete = (E * F * G + (E * F + I * J) / 2 * H) * D * O
    footing_formwork = (E + F) * 2 * G * D

    post_len = M - G - H
    post_lt_200 = post_len * D if K < 0.2 else 0.0
    post_ge_300 = post_len * D if K >= 0.3 else 0.0
    post_ge_200 = post_len * D if (K >= 0.2 and post_ge_300 == 0.0) else 0.0
    post_total = post_lt_200 + post_ge_200 + post_ge_300
    post_concrete = K * Lp * post_total * O
    post_formwork = (K + Lp) * 2 * post_total

    dig_h = -p.footing_bottom_level_m + p.ground_level_m + 0.05
    AD, AE = p.base_widen_m, p.top_widen_m
    if dig_h > 0:
        excavation = (((E + AD * 2) * (F + AD * 2)) + ((E + AE * 2) * (F + AE * 2))) / 2 * dig_h * D
        backfill = (excavation - lean_concrete - footing_concrete
                    - K * Lp * (-p.footing_bottom_level_m + p.ground_level_m - G - H) * D)
        surplus = 0.0
    else:
        excavation = 0.0
        backfill = 0.0
        surplus = (E + 0.1) * (F + 1) * (-dig_h) * D

    return {
        "lean_concrete_m3": round(lean_concrete, 8),
        "lean_formwork_m2": round(lean_formwork, 8),
        "footing_concrete_m3": round(footing_concrete, 8),
        "footing_formwork_m2": round(footing_formwork, 8),
        "post_lt200_md": round(post_lt_200, 8),
        "post_ge200_md": round(post_ge_200, 8),
        "post_ge300_md": round(post_ge_300, 8),
        "post_concrete_m3": round(post_concrete, 8),
        "post_formwork_m2": round(post_formwork, 8),
        "dig_height_m": round(dig_h, 8),
        "excavation_m3": round(excavation, 8),
        "backfill_m3": round(backfill, 8),
        "surplus_m3": round(surplus, 8),
    }


@dataclass
class FenceWallParams:
    length_m: float                 # E — dài tường
    thickness_m: float              # F — dày tường (0,22 hoặc 0,11)
    height_m: float                 # G — cao tường
    post_deduction_m: float = 0.0   # N — trừ trụ (tổng bề rộng trụ cắt vào tường)
    count: int = 1                  # D


def fence_wall_takeoff(p: FenceWallParams) -> Dict[str, float]:
    E, F, G, N, D = p.length_m, p.thickness_m, p.height_m, p.post_deduction_m, p.count
    masonry = (E - N) * F * G * D
    plaster = (E * G * 2 + E * F + G * F * 2) * D
    return {
        "masonry_m3": round(masonry, 8),
        "plaster_m2": round(plaster, 8),
        "paint_m2": round(plaster, 8),
        "is_220": abs(F - 0.22) < 1e-9,
        "is_110": abs(F - 0.11) < 1e-9,
    }
