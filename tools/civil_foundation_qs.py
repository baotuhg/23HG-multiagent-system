# -*- coding: utf-8 -*-
"""
CIVIL FOUNDATION QS — quy ước đo bóc đài móng dân dụng (bê tông, bê tông lót, ván khuôn)

Chuẩn hóa từ bảng tính khối lượng QS chuyên nghiệp (sheet 'Móng'), đúng công thức thực chiến:

  • Đài vuông/chữ nhật:  BT = L·W·H;  VK = (L+W)·2·H
  • Đài móng chóp cụt:   BT = L·W·H + ((L·W)+(L1·W1))/2·h_vát   ← QS dùng CÔNG THỨC XẤP XỈ
       (trung bình diện tích hai đáy). Lưu ý: khác công thức nón cụt chính xác
       (S1+S2+√(S1·S2))/3·h dùng trong tools/civil_and_bridge_takeoff_engine; chênh nhỏ.
  • Đài quả trám (kiểu 1): BT = (L·d + b·(L+a)/2)·H;  VK = (L + d·2 + c·2 + a)·H
  • Bê tông lót mở rộng mỗi phương một lượng `lean_extend` (Z), dày `lean_t` (Y);
    ván khuôn lót tính quanh chu vi khối lót.

Mọi công thức dưới đây tái tạo đúng các ô đã tính trong hồ sơ thật (xem tests/test_civil_foundation_qs_golden.py).
Khối lượng nhân số đài `count` và hệ số hao hụt `waste` (mặc định 1,0 = không hao hụt).

Thuần Python, không phụ thuộc ngoài. Không làm tròn trung gian.

GHI CHÚ QUY ƯỚC VÁN KHUÔN (sheet 'Hình dạng' của hồ sơ):
  - Vách chữ L, chữ I: tính theo quy cách hình vẽ, số cạnh ván khuôn = 2.
  - Vách lõi thang máy: các cạnh ván khuôn bù trừ cho nhau, không cần khai số cạnh.
  Các hằng số này ghi ở WALL_FORMWORK_FACES để dùng khi bóc ván khuôn vách.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict

WALL_FORMWORK_FACES: Dict[str, object] = {
    "chu_L": 2,          # vách chữ L — 2 cạnh ván khuôn
    "chu_I": 2,          # vách chữ I — 2 cạnh ván khuôn
    "loi_thang_may": "bù trừ",   # lõi thang máy — các cạnh bù trừ, không khai số cạnh
}


def _lean(concrete: float, formwork: float) -> Dict[str, float]:
    return {"lean_concrete_m3": round(concrete, 6), "lean_formwork_m2": round(formwork, 6)}


@dataclass
class FootingResult:
    concrete_m3: float
    formwork_m2: float
    lean_concrete_m3: float
    lean_formwork_m2: float

    def as_dict(self) -> Dict[str, float]:
        return {
            "concrete_m3": round(self.concrete_m3, 6),
            "formwork_m2": round(self.formwork_m2, 6),
            "lean_concrete_m3": round(self.lean_concrete_m3, 6),
            "lean_formwork_m2": round(self.lean_formwork_m2, 6),
        }


def footing_box(L: float, W: float, H: float,
                lean_t: float = 0.0, lean_extend: float = 0.0,
                count: int = 1, waste: float = 1.0) -> FootingResult:
    """Đài vuông/chữ nhật."""
    concrete = L * W * H * count * waste
    formwork = (L + W) * 2 * H * count
    lc = (L + lean_extend) * (W + lean_extend) * lean_t * count * waste
    lf = ((L + lean_extend) + (W + lean_extend)) * 2 * lean_t * count
    return FootingResult(concrete, formwork, lc, lf)


def footing_frustum(L: float, W: float, H: float,
                    L_top: float, W_top: float, h_slope: float,
                    lean_t: float = 0.0, lean_extend: float = 0.0,
                    count: int = 1, waste: float = 1.0) -> FootingResult:
    """
    Đài móng chóp cụt — BT = phần hộp đáy + phần vát theo công thức trung bình diện tích
    (quy ước QS trong Excel): V = L·W·H + ((L·W)+(L_top·W_top))/2 · h_slope.
    Ván khuôn tính theo chu vi đáy × chiều cao hộp (quy ước hồ sơ, không cộng mặt vát).
    """
    concrete = (L * W * H + ((L * W) + (L_top * W_top)) / 2.0 * h_slope) * count * waste
    formwork = (L + W) * 2 * H * count
    lc = (L + lean_extend) * (W + lean_extend) * lean_t * count * waste
    lf = ((L + lean_extend) + (W + lean_extend)) * 2 * lean_t * count
    return FootingResult(concrete, formwork, lc, lf)


def footing_diamond_type1(L: float, W: float, H: float,
                          a: float, b: float, c: float, d: float,
                          lean_t: float = 0.0, lean_extend: float = 0.0,
                          count: int = 1, waste: float = 1.0) -> FootingResult:
    """
    Đài quả trám (kiểu 1) — theo sheet 'Móng':
      BT = (L·d + b·(L+a)/2) · H
      VK = (L + d·2 + c·2 + a) · H
    với a,b,c,d là các kích thước vát của quả trám (cột L,M,N,O trong hồ sơ).
    """
    concrete = (L * d + b * (L + a) / 2.0) * H * count * waste
    formwork = (L + d * 2 + c * 2 + a) * H * count
    lc = ((L + lean_extend) * (d + lean_t)
          + ((L + lean_extend) + (a + lean_extend)) / 2.0 * (b + lean_t)) * lean_t * count * waste
    lf = ((L + lean_extend) + (d + lean_t) * 2 + (c + lean_extend) * 2 + a) * lean_t * count
    return FootingResult(concrete, formwork, lc, lf)
