# -*- coding: utf-8 -*-
"""
STEEL QS CONVENTIONS — quy ước tính khối lượng thép (cốt thép & kết cấu thép tiền chế)

Chuẩn hóa từ bảng tính thép QS thực chiến, đúng quy ước phổ biến ở Việt Nam:

  • Khối lượng đơn vị cốt thép:   m = D² / 162  (kg/m), D tính bằng mm.
  • Cây thép tiêu chuẩn 11,7 m; chỉ thanh vằn D > 8 mới đếm theo cây nguyên,
    D ≤ 8 cấp ở dạng cuộn (không quy ra số cây 11,7 m).
  • Số cây = ROUND(khối lượng / khối lượng 1 cây, 0)  (làm tròn nửa lên như ROUND của Excel).
  • Dây thép buộc = 1,5% khối lượng cốt thép.
  • Thép tấm (ký hiệu bắt đầu "PL"):  M = t(mm) × rộng(mm) × dài(mm) × 7,85 / 10⁶ (kg).
  • Thép hình (ký hiệu khác):         M = khối lượng riêng(kg/m) × dài(m).

Thuần Python, không phụ thuộc ngoài. Không tự làm tròn trung gian (chỉ làm tròn ở bước xuất).
"""

from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal
from typing import Dict, Optional

REBAR_DENSITY_DIVISOR = 162.0       # D²/162 (kg/m)
STANDARD_BAR_LENGTH_M = 11.7        # cây thép nguyên
COIL_MAX_D_MM = 8                   # D ≤ 8: cấp dạng cuộn, không quy ra cây 11,7 m
STEEL_DENSITY_KG_DM3 = 7.85         # khối lượng riêng thép
TIE_WIRE_RATIO = 0.015             # dây thép buộc 1,5%


def _round_half_up(value: float, ndigits: int = 0) -> float:
    """Làm tròn nửa lên như hàm ROUND của Excel (khác round() của Python)."""
    q = Decimal(1).scaleb(-ndigits)
    r = Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP)
    return int(r) if ndigits <= 0 else float(r)


def rebar_unit_mass(d_mm: float) -> float:
    """Khối lượng đơn vị cốt thép theo đường kính danh nghĩa: D²/162 (kg/m)."""
    if d_mm <= 0:
        raise ValueError(f"Đường kính thép phải > 0, nhận {d_mm}")
    return d_mm * d_mm / REBAR_DENSITY_DIVISOR


def rebar_whole_bar_weight(d_mm: float, bar_length_m: float = STANDARD_BAR_LENGTH_M) -> float:
    """Khối lượng 1 cây nguyên (mặc định 11,7 m); 0 với D ≤ 8 (cấp dạng cuộn)."""
    if d_mm <= COIL_MAX_D_MM:
        return 0.0
    return rebar_unit_mass(d_mm) * bar_length_m


def rebar_bar_count(total_kg: float, d_mm: float, bar_length_m: float = STANDARD_BAR_LENGTH_M) -> int:
    """Số cây nguyên cần mua = ROUND(tổng kg / kg một cây, 0); 0 với D ≤ 8 hoặc tổng kg = 0."""
    one = rebar_whole_bar_weight(d_mm, bar_length_m)
    if one <= 0 or total_kg <= 0:
        return 0
    return int(_round_half_up(total_kg / one, 0))


def tie_wire_kg(total_rebar_kg: float, ratio: float = TIE_WIRE_RATIO, ndigits: int = 3) -> float:
    """Khối lượng dây thép buộc = tỷ lệ (mặc định 1,5%) × khối lượng cốt thép."""
    return round(total_rebar_kg * ratio, ndigits)


def rebar_procurement(total_kg: float, d_mm: float,
                      waste_ratio: float = 0.0,
                      bar_length_m: float = STANDARD_BAR_LENGTH_M) -> Dict[str, float]:
    """Gói kết quả mua thép cho một đường kính: kg thiết kế → kg nhập (hao hụt) → số cây → dây buộc."""
    purchase_kg = total_kg * (1.0 + waste_ratio)
    return {
        "d_mm": d_mm,
        "unit_mass_kg_m": round(rebar_unit_mass(d_mm), 6),
        "design_kg": round(total_kg, 4),
        "purchase_kg": round(purchase_kg, 4),
        "whole_bar_kg": round(rebar_whole_bar_weight(d_mm, bar_length_m), 6),
        "bar_count": rebar_bar_count(purchase_kg, d_mm, bar_length_m),
        "tie_wire_kg": tie_wire_kg(purchase_kg),
    }


# ─────────────────────────────────────────────────────────────────────────────
# KẾT CẤU THÉP TIỀN CHẾ
# ─────────────────────────────────────────────────────────────────────────────

_PLATE_RE = re.compile(r"^\s*PL\s*([0-9]+(?:[.,][0-9]+)?)", re.I)


def parse_plate_thickness(mark: str) -> Optional[float]:
    """Trích chiều dày (mm) từ ký hiệu thép tấm, vd 'PL12' → 12.0; None nếu không phải tấm."""
    if mark is None:
        return None
    m = _PLATE_RE.match(str(mark))
    return float(m.group(1).replace(",", ".")) if m else None


def steel_plate_weight(thickness_mm: float, width_mm: float, length_mm: float, count: int = 1,
                       density_kg_dm3: float = STEEL_DENSITY_KG_DM3) -> float:
    """Khối lượng thép tấm: t × rộng × dài × 7,85 / 10⁶ (kg)."""
    return thickness_mm * width_mm * length_mm * count * density_kg_dm3 / 1_000_000.0


def steel_profile_weight(mass_per_m_kg: float, length_mm: float, count: int = 1) -> float:
    """Khối lượng thép hình: khối lượng riêng (kg/m) × dài (m) × số lượng."""
    return mass_per_m_kg * (length_mm / 1000.0) * count


def steel_member_weight(mark: str, size_mm_or_mass: float, width_mm: float, length_mm: float,
                        count: int = 1, ndigits: int = 4) -> float:
    """
    Khối lượng một chi tiết kết cấu thép, tự phân loại theo ký hiệu (như sheet 'ket cau thep'):
      - "PL.." → thép tấm: size_mm_or_mass là chiều dày (mm).
      - khác   → thép hình: size_mm_or_mass là khối lượng riêng (kg/m); bỏ qua width_mm.
    """
    if parse_plate_thickness(mark) is not None:
        w = steel_plate_weight(size_mm_or_mass, width_mm, length_mm, count)
    else:
        w = steel_profile_weight(size_mm_or_mass, length_mm, count)
    return round(w, ndigits)
