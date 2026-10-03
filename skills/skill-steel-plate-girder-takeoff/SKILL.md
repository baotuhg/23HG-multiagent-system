---
name: skill-steel-plate-girder-takeoff
description: Tính khối lượng thép tấm (PL t×w×L, 7,85 kg/dm³), diện tích sơn và khối lượng đinh neo chống cắt cho một dầm thép chữ I tổ hợp, kế thừa công thức từ bảng tính kết cấu thép tiền chế.
---

# Bóc tách dầm cầu thép chữ I tổ hợp và đinh neo chống cắt

**Mã:** `SKILL-STEEL-PLATE-GIRDER-TAKEOFF` · **Nhóm:** `STRUCTURAL_STEEL`
**Trạng thái:** `PENDING_APPROVAL` — chờ Kỹ sư trưởng duyệt.
**Mã nguồn:** [`tools/civil_and_bridge_takeoff_engine.py`](../../tools/civil_and_bridge_takeoff_engine.py)

## 1. Code thực sự tính gì

- **Thép tấm** (`calc_structural_steel_plate`): `m = t × w × L × 7,85 × n / 10⁶` (kg; t, w, L tính bằng mm).
- **Diện tích sơn:** `S = 2 × w × L × n / 10⁶` (m²). Công thức chỉ tính **2 mặt chính, chưa tính cạnh mép**. Code không phân biệt số lớp sơn: S là diện tích bề mặt, số lớp do hệ sơn quyết định.
- **Dầm** (`SteelBridgeGirderEngine` + `SteelBridgeGirderSegment`): cánh trên + cánh dưới + bụng + sườn tăng cường (chiều cao sườn lấy bằng chiều cao bụng) + đinh neo.
- **Đinh neo:** `π·d²/4 × L × 7,85 / 10⁶ × số lượng`. Số lượng đinh **phải nhập**, code không tự bố trí theo bước. Với D22×150 thì mỗi đinh khoảng 0,448 kg.
- **Chưa có:** bản nối, bu lông cường độ cao, hệ liên kết ngang, hao hụt cắt tấm.

## 2. Cách dùng (đã chạy thử)

```python
from tools.civil_and_bridge_takeoff_engine import (
    calc_structural_steel_plate, SteelBridgeGirderSegment, SteelBridgeGirderEngine,
)

# Một tấm cánh PL25 x 400, dài 30 m
pl = calc_structural_steel_plate(thickness_mm=25, width_mm=400, length_mm=30000, count=1)
print(pl["weight_kg"], "kg;", pl["paint_area_m2"], "m2")    # 2355.0 kg; 24.0 m2

# Dầm I tổ hợp L = 30 m
dam = SteelBridgeGirderEngine.compute_takeoff(SteelBridgeGirderSegment(
    length_mm=30000,
    top_flange_t_mm=25, top_flange_w_mm=400,
    bot_flange_t_mm=32, bot_flange_w_mm=500,
    web_t_mm=14, web_h_mm=1500,
    stiffener_count=40, stiffener_t_mm=12, stiffener_w_mm=150,
    stud_count=180, stud_dia_mm=22, stud_length_mm=150,
))
print("Tổng thép  :", dam["total_structural_steel_ton"], "tấn")       # 11.997
print("Diện tích sơn:", dam["paint_area_anti_corrosion_m2"], "m2")     # 162.0
print("Đinh neo   :", dam["shear_studs_kg"], "kg")                     # 80.57
```

## 3. Lưu ý

- Số lượng và bước đinh neo phải lấy từ bản vẽ thiết kế (TCVN 11823-6:2017 là tiêu chuẩn thiết kế; công cụ này chỉ đo bóc).
- Diện tích sơn thực tế cần cộng thêm cạnh mép và bản nối nếu hồ sơ yêu cầu.
