# -*- coding: utf-8 -*-
"""
INFRA CULVERT QS — đo bóc cống tròn hạ tầng (phân loại, cọc tre đế cống, đào/đắp/vận chuyển)

Chuẩn hóa từ sheet 'Hạ tầng- Cống tròn' của bảng tính khối lượng QS chuyên nghiệp. Tái tạo đúng
các ô đã tính trong hồ sơ thật (xem tests/test_infra_culvert_qs_golden.py).

Quy ước (đúng công thức trong hồ sơ):
  • Chiều dài tính toán của cống = Dài − (trừ giao hố ga + 0,2).
  • Phân loại theo loại cống: TH (cống thường), CL (cống chịu lực qua đường), PVC — xếp chiều dài
    vào đúng nhóm đường kính.
  • Cọc tre đế cống (khi không phải PVC):
      SL = ROUND( (Dài − (trừ_giao + 0,3)) / 2,5 × 2 × (0,33 + 0,2) × (D + 0,4) × mật_độ , 0 )
      Khối lượng cọc (md) = SL × chiều_dài_cọc.
  • Đào đắp rãnh đặt cống (mặt cắt hình thang):
      H_đào = −cốt_TB_đào + cốt_hiện_trạng
      mở_rộng_đỉnh = H_đào × hệ_số_mái(H_đào) + mở_rộng_chân
      dài_rãnh = Dài − (trừ + 0,3) − 1,0
      Đào   = ((D + chân×2) + (D + đỉnh×2))/2 × H_đào × dài_rãnh
      Đắp   = Đào − 3,14 × (D + 0,1)² / 4 × dài_rãnh      (trừ thể tích thân cống)
      Vận chuyển = Đào − Đắp
  • Hệ số mái taluy theo chiều cao đào: ≤1,5 m → 0,25; (1,5; 5) m → 0,67; ≥5 m → 0,85.

Hồ sơ dùng hằng số 3,14 cho số pi (không phải math.pi) — giữ nguyên để khớp số.
Thuần Python, không phụ thuộc ngoài, không làm tròn trung gian (trừ ROUND số cọc như hồ sơ).
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import Dict, Optional

PI_QS = 3.14                       # hồ sơ dùng 3,14 cho π
JUNCTION_EXTRA = 0.2               # cộng thêm vào "trừ giao hố ga" để ra đoạn trừ E
PILE_SPACING = 2.5                 # bước cọc tre (m)
PILE_ROWS = 2                      # số hàng cọc
PILE_BED_FACTOR = 0.33 + 0.2       # hệ số bề rộng đế cọc
PILE_WIDTH_EXTRA = 0.4             # (D + 0,4): bề rộng đế cống
TRENCH_END_ALLOWANCE = 0.3         # (E_extra + 0,15×2) ~ phần trừ hai đầu rãnh ngoài E
TRENCH_SHORTEN = 1.0               # trừ 0,5×2 hai đầu rãnh
PIPE_OUTER_EXTRA = 0.1             # (D + 0,1): đường kính ngoài quy ước để trừ thân cống

# Nhóm đường kính (mm) theo loại cống — như hàng tiêu đề trong hồ sơ
BUCKETS = {
    "TH": [200, 300, 400, 500, 600, 800, 1000, 1100, 1200],
    "CL": [200, 300, 400, 500, 600, 800, 1000, 1100, 1200],
    "PVC": [200, 250, 300],
}


def _round_half_up(value: float, ndigits: int = 0) -> float:
    q = Decimal(1).scaleb(-ndigits)
    r = Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP)
    return int(r) if ndigits <= 0 else float(r)


def slope_factor(dig_height_m: float, f_low: float = 0.25, f_mid: float = 0.67,
                 f_high: float = 0.85) -> float:
    """Hệ số mái taluy theo chiều cao đào: ≤1,5→0,25; (1,5;5)→0,67; ≥5→0,85."""
    if dig_height_m >= 5:
        return f_high
    return f_low if dig_height_m <= 1.5 else f_mid


@dataclass
class CulvertParams:
    diameter_m: float                  # D cống (F)
    length_m: float                    # Dài (H)
    culvert_type: str = "TH"           # TH | CL | PVC
    junction_trim_m: float = 0.0       # trừ giao hố ga (D cột)
    ground_level_m: float = 0.0        # cốt hiện trạng (AG)
    avg_dig_level_m: float = 0.0       # cốt trung bình đào cống (AH, thường âm)
    base_widen_m: float = 0.5          # mở rộng chân (AI)
    pile_length_m: float = 2.0         # chiều dài cọc tre (AD)
    pile_density: float = 25.0         # mật độ cọc (cọc/md², AE11)


def culvert_effective_length(p: CulvertParams) -> float:
    """Chiều dài tính toán = Dài − (trừ giao + 0,2)."""
    return p.length_m - (p.junction_trim_m + JUNCTION_EXTRA)


def culvert_bucket(p: CulvertParams) -> Optional[str]:
    """Nhãn nhóm 'LOẠI-Dmm' mà chiều dài được xếp vào, hoặc None nếu D không thuộc nhóm."""
    d_mm = round(p.diameter_m * 1000)
    if d_mm in BUCKETS.get(p.culvert_type, []):
        return f"{p.culvert_type}-{d_mm}"
    return None


def culvert_pile(p: CulvertParams) -> Dict[str, float]:
    """Cọc tre đế cống: số cọc (ROUND) và khối lượng (md). PVC không đóng cọc."""
    if p.culvert_type == "PVC" or p.pile_length_m <= 0:
        return {"pile_count": 0, "pile_md": 0.0}
    raw = ((p.length_m - (p.junction_trim_m + 0.3)) / PILE_SPACING * PILE_ROWS
           * PILE_BED_FACTOR * (p.diameter_m + PILE_WIDTH_EXTRA) * p.pile_density)
    count = int(_round_half_up(raw, 0)) if raw > 0 else 0
    return {"pile_count": count, "pile_md": round(count * p.pile_length_m, 4)}


def culvert_earthwork(p: CulvertParams) -> Dict[str, float]:
    """Đào / đắp / vận chuyển đất rãnh đặt cống (mặt cắt hình thang, trừ thân cống)."""
    if p.length_m == 0:
        return {"dig_height_m": 0.0, "excavation_m3": 0.0, "backfill_m3": 0.0, "haul_m3": 0.0}
    e_trim = p.junction_trim_m + JUNCTION_EXTRA
    dig_h = -p.avg_dig_level_m + p.ground_level_m
    top_widen = dig_h * slope_factor(dig_h) + p.base_widen_m
    trench_len = p.length_m - (e_trim + 0.15 * 2) - TRENCH_SHORTEN
    width_avg = ((p.diameter_m + p.base_widen_m * 2) + (p.diameter_m + top_widen * 2)) / 2.0
    excavation = width_avg * dig_h * trench_len
    pipe_vol = PI_QS * (p.diameter_m + PIPE_OUTER_EXTRA) ** 2 / 4.0 * trench_len
    backfill = excavation - pipe_vol
    return {
        "dig_height_m": round(dig_h, 6),
        "top_widen_m": round(top_widen, 6),
        "excavation_m3": round(excavation, 6),
        "backfill_m3": round(backfill, 6),
        "haul_m3": round(excavation - backfill, 6),
    }


def culvert_takeoff(p: CulvertParams) -> Dict[str, object]:
    """Gộp toàn bộ kết quả đo bóc một tuyến cống."""
    out: Dict[str, object] = {
        "bucket": culvert_bucket(p),
        "effective_length_m": round(culvert_effective_length(p), 6),
    }
    out.update(culvert_pile(p))
    out.update(culvert_earthwork(p))
    return out
