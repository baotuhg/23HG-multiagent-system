# -*- coding: utf-8 -*-
"""
INFRA TANK QS — đo bóc bể nước ngầm (PCCC / xử lý nước thải) — dòng đáy bể

Chuẩn hóa từ các sheet 'Bể nước PCCC' / 'Bể XLNT' của bảng tính khối lượng QS chuyên nghiệp. Hàm
dưới đây tái tạo đúng DÒNG ĐÁY BỂ: bê tông lót + ván khuôn lót, bê tông đáy + ván khuôn, chống thấm
trong/ngoài, đào/đắp đất hố bể (mặt cắt hình thang hai cao trình), và cốt thép đáy (lưới 2 lớp).
Các ô đã tính sẵn trong hồ sơ thật được khớp đến từng m³/kg (xem test).

Thành bể và nắp bể trong hồ sơ là các dòng cấu kiện riêng (khối lăng trụ); mô-đun này tập trung
phần đáy đã kiểm chứng.

Quy ước (đúng hồ sơ):
  • Bê tông lót: (L + lót_mở_rộng×2)(W + lót_mở_rộng×2) × lót_dày.
  • Bê tông đáy: L × W × đáy_dày.
  • H đào = cao_bể + đáy_dày + lót_dày + mực_offset − đáy_offset.
  • Đào = trung bình hai hình chữ nhật mở rộng (chân & đỉnh) × H đào.
  • Đắp = Đào − lót − bê tông đáy − thể tích lòng bể dưới cao trình đáy.
  • Cốt thép lưới (kg/m²) = Σ d²/162 × (1000/khoảng_cách) cho từng phương/lớp.

Thuần Python, không phụ thuộc ngoài.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

REBAR_DIVISOR = 162.0


@dataclass
class TankBaseParams:
    length_m: float                 # F
    width_m: float                  # G
    tank_height_m: float            # H — chiều cao bể
    slab_t_m: float                 # I — đáy dày
    lean_t_m: float = 0.05          # J — bê tông lót dày
    lean_widen_m: float = 0.1       # K — mở rộng lót mỗi phía
    waste: float = 1.0              # N — hao hụt
    count: int = 1                  # E
    wall_type: str = "IW1"          # IW1 (trong) | EW1 (ngoài) — cho chống thấm
    bottom_widen_m: float = 1.2     # AA — mở rộng hố ở cao trình dưới
    top_widen_m: float = 2.0        # AB — mở rộng hố ở cao trình trên
    level_offset_m: float = -0.3    # Z — mực chuẩn
    base_offset_m: float = -0.1     # AC
    # Lưới thép đáy: danh sách (đường kính mm, khoảng cách mm); mặc định 4 phương/lớp như hồ sơ
    rebar_mesh: List[Tuple[float, float]] = field(
        default_factory=lambda: [(14, 150), (14, 150), (14, 150), (14, 150)])
    rebar_factor: float = 1.12      # AP


def tank_base_takeoff(p: TankBaseParams) -> Dict[str, float]:
    F, G, I, J, K = p.length_m, p.width_m, p.slab_t_m, p.lean_t_m, p.lean_widen_m
    N, E = p.waste, p.count
    AA, AB = p.bottom_widen_m, p.top_widen_m

    lean_concrete = (F + K * 2) * (G + K * 2) * J * N * E
    lean_formwork = ((F + K * 2) + (G + K * 2)) * 2 * J * E
    concrete = F * G * I * E * N
    formwork = (F + G) * 2 * I * E
    area = F * G * E
    waterproof_outer = area if p.wall_type == "EW1" else 0.0
    waterproof_inner = area if p.wall_type == "IW1" else 0.0

    dig_h = p.tank_height_m + I + J + p.level_offset_m - p.base_offset_m
    excavation = (((F + AA * 2) * (G + AA * 2) + (F + AB * 2) * (G + AB * 2)) / 2 * dig_h * E)
    backfill = excavation - lean_concrete - concrete - (F * G * (dig_h - I - J) * E)

    rebar_unit = sum(d * d / REBAR_DIVISOR * (1000.0 / s) for d, s in p.rebar_mesh if d > 0 and s > 0)
    rebar_kg = F * G * E * rebar_unit * p.rebar_factor

    return {
        "lean_concrete_m3": round(lean_concrete, 8),
        "lean_formwork_m2": round(lean_formwork, 8),
        "concrete_m3": round(concrete, 8),
        "formwork_m2": round(formwork, 8),
        "area_m2": round(area, 8),
        "waterproof_outer_m2": round(waterproof_outer, 8),
        "waterproof_inner_m2": round(waterproof_inner, 8),
        "dig_height_m": round(dig_h, 8),
        "excavation_m3": round(excavation, 8),
        "backfill_m3": round(backfill, 8),
        "rebar_unit_kg_m2": round(rebar_unit, 8),
        "rebar_kg": round(rebar_kg, 8),
    }
