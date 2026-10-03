---
name: skill-steel-plate-girder-takeoff
description: Bóc tách chuỗi định danh thép tấm PL{t}x{w}x{L} theo công thức 7.85 t/m3, tự động tính diện tích sơn chống gỉ 3 lớp (S = 2*(w+t)*L) và bóc tách đinh neo chống cắt Nelson D22x150mm theo TCVN 11823:2017.
---

# Kỹ năng Bóc tách Dầm Cầu Thép Tấm Chữ I Liên Hợp & Đinh Neo Nelson

**Mã Kỹ năng:** `SKILL-STEEL-PLATE-GIRDER-TAKEOFF`  
**Nhóm nghiệp vụ:** `STRUCTURAL_STEEL`  
**Trạng thái:** `APPROVED` (Kỹ sư trưởng phê duyệt)  
**Tiêu chuẩn áp dụng:** TCVN 11823-6:2017, TCVN 8789:2011 (Sơn bảo vệ kết cấu thép), Thông tư 38/2026/TT-BXD.

---

## 1. Mục tiêu & Nguyên lý Tính toán

Mở rộng thuật toán bóc tách kết cấu thép tiền chế (Zamil KCT) sang kết cấu Dầm cầu thép chữ I liên hợp (Steel Plate Girder Bridge):

1. **Bóc tách Thép tấm (`PL{t}`):**
   - Quy đổi quy cách thép tấm: $\text{PL}t \times w \times L$ (bề dày $t\text{ mm}$, bề rộng $w\text{ mm}$, chiều dài $L\text{ mm}$).
   - Khối lượng thép tấm với tỷ trọng $7.85\text{ tấn/m}^3$:
     $$m = t \times w \times L \times 7.85 \times 10^{-6}\quad (\text{tấn})$$
2. **Diện tích Sơn chống gỉ 3 lớp (Hệ sơn Polyurethane / Epoxy):**
   - Bắt buộc tính toán sơn bảo vệ bề mặt toàn phần cho thép tấm (gồm 2 mặt chính và các cạnh mép) (`RULE-STRUCT-STEEL-PAINT-009`):
     $$S_{\text{sơn}} = 2 \times (w + t) \times L \times 10^{-6}\quad (\text{m}^2)$$
3. **Đinh neo chống cắt Nelson (Shear Studs):**
   - Khắc phục liên kết chống cắt giữa dầm thép và bản bê tông mặt cầu theo TCVN 11823-6:2017:
     - Đường kính chuẩn: $\Phi22\text{ mm}$, chiều dài $150\text{ mm}$.
     - Khối lượng đinh neo: $m_{\text{stud}} = \frac{\pi \cdot d^2}{4} \cdot L_{\text{stud}} \cdot 7850 \approx 0.448\text{ kg/cái}$.

---

## 2. Hướng dẫn Sử dụng Module Python

```python
from tools.civil_and_bridge_takeoff_engine import (
    calc_structural_steel_plate,
    SteelBridgeGirderEngine,
)

# 1. Bóc tách đơn chiếc một tấm bản cánh hoặc bản bụng dầm thép
flange_plate = calc_structural_steel_plate(
    thickness_mm=25.0, width_mm=500.0, length_mm=30000.0, count=2
)
print(f"Khối lượng 2 bản cánh dầm: {flange_plate['weight_ton']:.3f} tấn")
print(f"Diện tích sơn bảo vệ:      {flange_plate['paint_area_m2']:.2f} m2")

# 2. Bóc tách hoàn chỉnh dầm cầu thép tấm chữ I 30m liên hợp
girder = SteelBridgeGirderEngine(
    span_L=30.0,
    web_H=1600.0, web_t=14.0,
    top_flange_W=500.0, top_flange_t=25.0,
    bot_flange_W=600.0, bot_flange_t=32.0,
    stiffener_spacing=1.5,
    stud_spacing=0.5, studs_per_row=3
)
res = girder.compute_takeoff()
print(f"Tổng khối lượng dầm thép: {res['total_steel_weight_ton']:.3f} tấn")
print(f"Tổng diện tích sơn 3 lớp: {res['total_paint_area_m2']:.2f} m2")
print(f"Số lượng đinh neo Nelson: {res['shear_studs']['count']} cái ({res['shear_studs']['total_weight_kg']:.1f} kg)")
```

---

## 3. Điều kiện Kích hoạt & Phê duyệt

- **Trigger:** Khi lập dự toán hoặc nghiệm thu kết cấu dầm thép, cầu liên hợp thép - bê tông cốt thép, bản liên kết bu lông cường độ cao.
- **Quy tắc bảo vệ:** Không bỏ sót diện tích sơn lót và sơn chống gỉ mặt lưng; đối chiếu số lượng đinh neo Nelson theo đúng bước thiết kế.
