---
name: skill-civil-to-bridge-takeoff
description: Bóc tách sơ bộ khối lượng mố cầu chữ U, trụ cầu (bệ hai đầu bo tròn, 2 cột tròn, xà mũ có cánh hẫng vát) và kết cấu nhịp dầm Super-T bằng các công thức hình học kế thừa từ bảng tính QS dân dụng. Dùng cho ước tính/kiểm tra chéo, chưa thay thế bóc tách từ bản vẽ thi công.
---

# Bóc tách sơ bộ mố, trụ, kết cấu nhịp cầu (kế thừa công thức QS dân dụng)

**Mã:** `SKILL-CIVIL-TO-BRIDGE-TAKEOFF` · **Nhóm:** `QUANTITY_SURVEYING`
**Trạng thái:** `PENDING_APPROVAL` — chờ Kỹ sư trưởng duyệt. Chưa đối chiếu với bản vẽ cầu thật nào.
**Mã nguồn:** [`tools/civil_and_bridge_takeoff_engine.py`](../../tools/civil_and_bridge_takeoff_engine.py) · **Test:** [`tests/test_civil_and_bridge_takeoff.py`](../../tests/test_civil_and_bridge_takeoff.py)

## 1. Code thực sự tính gì

| Cấu kiện | Cách tính trong code | Giới hạn cần biết |
|---|---|---|
| Bệ trụ | Khối hộp `L×W×H` + 2 nửa trụ tròn bán kính `W/2` ở hai đầu (`calc_cutwater_pier_footing`) | Chỉ có mũi bo tròn; chưa có mũi tam giác/vát |
| Thân trụ | `n` cột tròn `π·D²/4·H`; ván khuôn `π·D·H` | Chưa trừ phần ngàm vào bệ/xà mũ |
| Xà mũ | Mặt đứng = đoạn giữa chữ nhật (dài `n·D + 1,2 m`) + 2 cánh hẫng hình thang, nhân bề rộng | Đoạn giữa `+1,2 m` là giả định cố định trong code |
| Mố chữ U | Lót, bệ, thân, tường ngực, 2 tường cánh hình thang, đá kê gối, bản quá độ — đều là khối hộp/lăng trụ | Không trừ giao giữa các khối |
| Dầm Super-T | **Nhập sẵn** thể tích, ván khuôn, thép, cáp DƯL cho **1 dầm** rồi nhân số dầm | Không tính mặt cắt dầm; số liệu 1 dầm phải lấy từ hồ sơ thiết kế |
| Bản mặt cầu | `L × B × t`; ván khuôn đáy lấy `65% × B × L` | Hệ số 65% là giả định |
| Cốt thép | `Bê tông × hàm lượng kg/m³` (mố 110, trụ 125, bản mặt cầu 140) | Chỉ là ước tính theo hàm lượng, không phải BBS |
| Dây thép buộc | `1,5% × khối lượng thép` | Hệ số lấy từ bảng tính folder "Mua"; cần kỹ sư đối chiếu định mức áp dụng |

## 2. Cách dùng (đã chạy thử)

```python
from tools.civil_and_bridge_takeoff_engine import (
    BridgePierParams, BridgePierEngine,
    BridgeAbutmentParams, BridgeAbutmentEngine,
    BridgeSuperstructureParams, BridgeSuperstructureEngine,
)

# Trụ cầu: bệ 7,0 x 3,6 x 1,8 m; 2 cột D1,5 m cao 8,5 m; xà mũ dài 9,6 m
tru = BridgePierEngine.compute_takeoff(BridgePierParams(
    name="Trụ T1",
    footing_rect_L=7.0, footing_rect_W=3.6, footing_H=1.8,
    column_dia=1.5, column_height=8.5, column_count=2,
    cap_length=9.6, cap_width=1.8, cap_height_center=1.6, cap_height_end=1.0,
))
print("Bê tông trụ :", tru["total_structural_concrete_m3"], "m3")   # 118.815
print("Ván khuôn   :", tru["total_formwork_m2"], "m2")             # 176.666

# Mố chữ U (dùng giá trị mặc định của BridgeAbutmentParams, sửa theo bản vẽ)
mo = BridgeAbutmentEngine.compute_takeoff(BridgeAbutmentParams(name="Mố M1"))
print("Bê tông mố  :", mo["total_structural_concrete_m3"], "m3")

# Kết cấu nhịp: 4 dầm Super-T 33 m — số liệu 1 dầm phải lấy từ hồ sơ thiết kế
nhip = BridgeSuperstructureEngine.compute_takeoff(BridgeSuperstructureParams(
    bridge_length_m=33.0, deck_width_m=9.0, girder_count=4,
    girder_volume_single_m3=23.5, girder_formwork_single_m2=145.0,
))
print("BT dầm đúc sẵn:", nhip["girder_concrete_m3"], "m3")  # 94.0
print("BT bản mặt cầu:", nhip["deck_concrete_m3"], "m3")    # 59.4
```

Demo đầy đủ: `python examples/demo_bridge_takeoff_from_qs_logic.py`

## 3. Khi nào dùng / không dùng

- **Dùng:** ước tính nhanh giai đoạn chuẩn bị, kiểm tra chéo bảng khối lượng của tư vấn.
- **Không dùng:** làm khối lượng thanh toán hoặc hồ sơ pháp lý khi chưa được kỹ sư đối chiếu với bản vẽ thi công.
