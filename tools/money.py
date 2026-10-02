# -*- coding: utf-8 -*-
"""
MONEY — số học tiền (VNĐ) bằng Decimal, làm tròn "half up" giống hàm ROUND của Excel.

Vì sao không dùng round() của Python: round() làm tròn về số chẵn (round(2.5) = 2, round(0.5) = 0),
còn Excel/kế toán làm tròn ra xa số 0 (ROUND(2.5, 0) = 3, ROUND(-2.5, 0) = -3). Ngoài ra tích
khối lượng × đơn giá bằng float có thể sai số nhị phân (vd 592.4 × 980000), nên mọi phép tính
tiền đi qua Decimal dựng từ biểu diễn thập phân của số gốc (str(x)).

Pure Python, Zero LLM.
"""

from __future__ import annotations
from decimal import Decimal, ROUND_HALF_UP
from typing import Iterable, Union

Number = Union[int, float, Decimal, str]


def D(x: Number) -> Decimal:
    """Số → Decimal theo biểu diễn thập phân (0.1 → Decimal('0.1'), không phải 0.1000000000000000055...)."""
    if isinstance(x, Decimal):
        return x
    if isinstance(x, bool):
        raise TypeError("bool không phải số tiền")
    return Decimal(str(x))


def round_half_up(x: Number, ndigits: int = 0) -> Decimal:
    """ROUND của Excel: ra xa số 0 khi đúng một nửa."""
    q = Decimal(1).scaleb(-ndigits)
    return D(x).quantize(q, rounding=ROUND_HALF_UP)


def round_vnd(x: Number) -> int:
    """Làm tròn đến đồng (số nguyên), half up."""
    return int(round_half_up(x, 0))


def mul(*factors: Number) -> Decimal:
    """Tích chính xác của các thừa số (khối lượng × đơn giá × hệ số...)."""
    out = Decimal(1)
    for f in factors:
        out *= D(f)
    return out


def total(values: Iterable[Number]) -> Decimal:
    """Tổng chính xác."""
    out = Decimal(0)
    for v in values:
        out += D(v)
    return out
