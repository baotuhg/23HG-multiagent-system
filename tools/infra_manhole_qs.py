# -*- coding: utf-8 -*-
"""
INFRA MANHOLE QS — đo bóc hố ga hạ tầng kiểu 1 (bê tông / xây gạch)

Chuẩn hóa từ sheet 'Hố ga (Kiểu 1)' của bảng tính khối lượng QS chuyên nghiệp. Tái tạo đúng các ô
đã tính trong hồ sơ thật (xem tests/test_infra_manhole_qs_golden.py). Kích thước nhập bằng mm;
khối lượng ra m³ (÷10⁹) và m² (÷10⁶). Dùng math.pi cho π như hàm PI() của hồ sơ.

Hỗ trợ:
  • Hố ga BÊ TÔNG: bê tông thành+đáy (trừ lỗ cống), ván khuôn.
  • Hố ga XÂY GẠCH: khối xây, trát, bê tông đáy+cổ ga, ván khuôn đáy+cổ ga.
  • Bê tông lót + ván khuôn lót; nắp ga (bê tông hoặc song chắn 'Grating'); cọc tre;
    đào/đắp/vận chuyển đất hố (mặt cắt hình thang, hệ số mái taluy theo chiều cao đào).

Quy ước (đúng hồ sơ):
  - Bao ngoài G: bê tông = D + 150×4; xây gạch = D + 150×2 + 220×2;  H = G.
  - Đáy dày J, thành dày K: bê tông (150, 150); xây gạch (100, 220).
  - Chiều cao H hố = cốt đỉnh − cốt đáy.
  - Hệ số mái taluy theo H đào (mm): ≤1500 → 0,25; (1500; 5000) → 0,67; ≥5000 → 0,85.

Thuần Python, không phụ thuộc ngoài, không làm tròn trung gian (trừ ROUND số cọc như hồ sơ).
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import Dict

PI = math.pi
BASE_WIDEN_MM = 500.0        # mở rộng chân hố (AN)
PILE_DENSITY = 25.0          # mật độ cọc tre (AE11 bên cống)


def _round_half_up(value: float, ndigits: int = 0) -> float:
    q = Decimal(1).scaleb(-ndigits)
    r = Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP)
    return int(r) if ndigits <= 0 else float(r)


def slope_factor_mm(dig_height_mm: float) -> float:
    """Hệ số mái taluy theo chiều cao đào (mm): ≤1500→0,25; (1500;5000)→0,67; ≥5000→0,85."""
    if dig_height_mm >= 5000:
        return 0.85
    return 0.25 if dig_height_mm <= 1500 else 0.67


@dataclass
class ManholeType1Params:
    culvert_in_mm: float               # D cống vào ga (F)
    bottom_level_mm: float             # cốt đáy theo thiết kế (M, thường âm)
    top_level_mm: float                # cốt đỉnh ga (O)
    manhole_type: str = "Bê tông"      # "Bê tông" | "Xây gạch"
    lean_t_mm: float = 50.0            # chiều dày BT lót (L)
    ground_level_mm: float = 0.0       # cốt hiện trạng (AM)
    waste: float = 1.0                 # hao hụt vật liệu (P)
    cross_D_mm: float = 0.0            # trừ giao cắt cống — đường kính (Q)
    cross_count: int = 0               # số giao cắt cống (S)
    cover_type: str = "Grating"        # "Grating" | "Bê tông"
    base_widen_mm: float = BASE_WIDEN_MM
    pile_length_m: float = 2.0         # chiều dài 1 cọc (AJ)
    count: int = 1                     # số lượng (E)

    # ----- hình học dẫn xuất -----
    @property
    def G(self) -> float:
        if self.culvert_in_mm <= 0:
            return 0.0
        return (self.culvert_in_mm + 150 * 4 if self.manhole_type == "Bê tông"
                else self.culvert_in_mm + 150 * 2 + 220 * 2)

    @property
    def H_plan(self) -> float:         # bao ngoài theo phương còn lại (= G)
        return self.G

    @property
    def wall_H(self) -> float:         # chiều cao hố (I)
        return self.top_level_mm - self.bottom_level_mm

    @property
    def J(self) -> float:              # đáy dày
        return 0.0 if self.G == 0 else (150 if self.manhole_type == "Bê tông" else 100)

    @property
    def K(self) -> float:              # thành dày
        return 0.0 if self.G == 0 else (150 if self.manhole_type == "Bê tông" else 220)

    @property
    def R(self) -> float:              # đường kính lỗ cống lớn nhất khoét thành
        if self.cross_D_mm == 0:
            return 0.0
        return self.cross_D_mm + 50 * 2 if self.cross_D_mm < 1000 else self.cross_D_mm + 100 * 2

    @property
    def dig_bottom_mm(self) -> float:  # cốt đào đất (N)
        return self.bottom_level_mm - self.J - self.lean_t_mm

    @property
    def dig_height_mm(self) -> float:  # H đào (AP)
        return -self.dig_bottom_mm + self.ground_level_mm


def manhole_takeoff(p: ManholeType1Params) -> Dict[str, float]:
    G, H, I, J, K, L = p.G, p.H_plan, p.wall_H, p.J, p.K, p.lean_t_mm
    R, S, E, P = p.R, p.cross_count, p.count, p.waste
    out: Dict[str, float] = {}
    if G == 0:
        return {k: 0.0 for k in ("lean_concrete_m3", "lean_formwork_m2", "masonry_m3",
                                 "plaster_m2", "bottom_neck_concrete_m3", "bottom_neck_formwork_m2",
                                 "concrete_m3", "formwork_m2", "cover_concrete_m3", "cover_neck_md",
                                 "grating_neck_md", "pile_count", "pile_md",
                                 "excavation_m3", "backfill_m3", "haul_m3")}

    # Bê tông lót + ván khuôn lót
    out["lean_concrete_m3"] = round((G + L * 2) * (H + L * 2) * L / 1e9 * E * P, 8)
    out["lean_formwork_m2"] = round(((G + L * 2) + (H + L * 2)) * 2 * L / 1e6 * E, 8)

    if p.manhole_type == "Xây gạch":
        out["masonry_m3"] = round(((G * K * (I - 200) * 2) + ((H - K * 2) * K * (I - 200) * 2)
                                   - R ** 2 * PI / 4 * K * S) / 1e9 * E, 8)
        out["plaster_m2"] = round((((G + H) * 2 + ((G - K * 2) + (H - K * 2)) * 2) * I
                                   - R ** 2 * PI / 4 * 2 * S) / 1e6 * E, 8)
        out["bottom_neck_concrete_m3"] = round((G * H * J + G * K * 200 * 2 + (H - K * 2) * K * 200)
                                               / 1e9 * P * E, 8)
        out["bottom_neck_formwork_m2"] = round(((G + H) * 2 * J + (G + H) * 2 * 200
                                                + (G - K * 2) * 4 * 200) / 1e6 * E, 8)
        out["concrete_m3"] = 0.0
        out["formwork_m2"] = 0.0
    else:   # Bê tông
        out["masonry_m3"] = 0.0
        out["plaster_m2"] = 0.0
        out["bottom_neck_concrete_m3"] = 0.0
        out["bottom_neck_formwork_m2"] = 0.0
        out["concrete_m3"] = round(((G * H * J) + (G * K * I * 2) + ((H - K * 2) * K * I) * 2
                                    - R ** 2 * PI / 4 * K * S) / 1e9 * P * E, 8)
        wall = (G * (I + J) * 4 + (G - K * 2) * I * 4)
        if p.cross_D_mm > 0:
            wall += -R ** 2 * PI / 4 * 2 * S + R * PI * K * S
        out["formwork_m2"] = round(wall / 1e6 * E, 8)

    # Nắp ga
    AE = G - K * 2 + 100
    AF = AE if AE <= 700 else AE / 2 - 3
    AD = 1 if AE == AF else 2
    out["cover_concrete_m3"] = round((AE * AF * 40) / 1e9 * AD * P * E if p.cover_type == "Bê tông" else 0.0, 8)
    out["cover_neck_md"] = round((AE + AF) * 2 * AD / 1000 * E if p.cover_type == "Bê tông" else 0.0, 8)
    out["grating_neck_md"] = round(AE * 4 / 1000 * E, 8)

    # Cọc tre
    pile = int(_round_half_up((G / 1000 + 0.2) * (H / 1000 + 0.2) * PILE_DENSITY, 0))
    out["pile_count"] = pile
    out["pile_md"] = round(p.pile_length_m * pile * E, 6)

    # Đào đắp
    AP = p.dig_height_mm
    AN = p.base_widen_mm
    AO = AP * slope_factor_mm(AP) + AN
    exc = ((G + AN * 2) * (H + AN * 2) + ((G + AO * 2) * (H + AO * 2)) / 2 * AP) / 1e9 * E
    backfill = exc - out["lean_concrete_m3"] - G * H * (AP - L) / 1e9 * E
    out["excavation_m3"] = round(exc, 8)
    out["backfill_m3"] = round(backfill, 8)
    out["haul_m3"] = round(exc - backfill, 8)
    return out
