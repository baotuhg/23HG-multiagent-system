---
name: skill-civil-to-bridge-takeoff
description: Chuyển giao và mở rộng nguyên lý hình học móng chóp cụt, vai cột corbel, dầm trừ giao sàn từ QS dân dụng sang mố chữ U, trụ xẻ nước mũi thuyền, xà mũ vươn hẫng hammerhead và dầm Super-T 33m theo TCVN 11823:2017 & TT 38/2026/TT-BXD.
---

# Kỹ năng Bóc tách Hình học Cầu Đường từ Nguyên lý QS Dân dụng (Bridge Takeoff Engine)

**Mã Kỹ năng:** `SKILL-CIVIL-TO-BRIDGE-TAKEOFF`  
**Nhóm nghiệp vụ:** `QUANTITY_SURVEYING`  
**Trạng thái:** `APPROVED` (Kỹ sư trưởng phê duyệt)  
**Tiêu chuẩn áp dụng:** TCVN 11823:2017, TCVN 5574:2018, Thông tư 38/2026/TT-BXD, Thông tư 36/2026/TT-BXD.

---

## 1. Mục tiêu & Nguyên lý Đồng hình Hình học (Isomorphism)

Chuyển đổi các công thức bóc tách hình học truyền thống của công trình dân dụng / công nghiệp sang kết cấu hạ tầng giao thông cầu đường:

1. **Móng chóp cụt & Móng Oval $\longrightarrow$ Bệ mố / Bệ trụ xẻ nước mũi thuyền**:
   - Khối chóp cụt: $V = \frac{h}{3} (S_1 + S_2 + \sqrt{S_1 S_2})$.
   - Mũi xẻ dòng thủy lực: Tam giác / Bán nguyệt giảm lực cản dòng chảy theo TCVN 11823-3:2017.
2. **Vai cột Corbel $\longrightarrow$ Cánh hẫng xà mũ trụ cầu (Hammerhead)**:
   - Thân trụ chịu lực uốn nén kết hợp cánh hẫng dạng nêm lăng trụ tam giác vươn đỡ gối chậu và dầm Super-T.
3. **Dầm tầng trừ giao cột/sàn $\longrightarrow$ Dầm Super-T 33m & Dầm I trừ giao bản mặt cầu**:
   - Đa giác mặt cắt chia mảnh (cánh trên, cánh dưới, sườn nghiêng).
   - Ván khuôn thành ngoài và vòm trong trừ diện tích tiếp giáp bê tông bản mặt cầu (Zero Duplication).
4. **Tường xây 8 công tác $\longrightarrow$ Tường thân, tường ngực & tường cánh mố chữ U**:
   - Tự động phái sinh 6 công tác liên hoàn: Bê tông $\to$ Ván khuôn $\to$ Xử lý mạch ngừng thi công $\to$ Sơn chống thấm Bitum $\to$ Vải địa kỹ thuật lọc $\to$ Tầng phòng nước.

---

## 2. Hướng dẫn Sử dụng Module Python

Sử dụng trực tiếp trong mã nguồn từ module [`tools.civil_and_bridge_takeoff_engine`](tools/civil_and_bridge_takeoff_engine.py):

```python
from tools.civil_and_bridge_takeoff_engine import (
    BridgePierEngine,
    BridgeAbutmentEngine,
    BridgeSupersuperstructureEngine,
)

# 1. Bóc tách Trụ cầu xẻ nước & Xà mũ vươn hẫng Hammerhead
pier = BridgePierEngine(
    footing_L=10.0, footing_W=4.5, footing_H=2.0, cutwater_len=2.25,
    col_diam=1.5, col_H=8.5, num_cols=2,
    cap_L=12.5, cap_W=2.2, cap_H_center=1.8, cap_H_tip=1.0, cant_L=3.5
)
pier_res = pier.compute_takeoff()
print(f"Bê tông bệ trụ: {pier_res['footing']['concrete_m3']:.2f} m3")
print(f"Bê tông xà mũ:  {pier_res['cap']['concrete_m3']:.2f} m3")
print(f"Tổng bê tông:   {pier_res['total_concrete_m3']:.2f} m3")
print(f"Tổng ván khuôn: {pier_res['total_formwork_m2']:.2f} m2")

# 2. Bóc tách Mố chữ U & Tường cánh vát taluy
abutment = BridgeAbutmentEngine(
    footing_L=11.0, footing_W=4.0, footing_H=1.5,
    stem_H=5.5, stem_W=1.2,
    breast_wall_H=1.6, breast_wall_W=0.4,
    wing_L=4.5, wing_H_front=7.1, wing_H_back=2.0, wing_thick=0.4
)
abut_res = abutment.compute_takeoff()
print(f"Bê tông mố: {abut_res['total_concrete_m3']:.2f} m3")

# 3. Bóc tách Kết cấu nhịp Super-T 33m
super_t = BridgeSuperstructureEngine(span_length=33.0, num_girders=4, bridge_width=10.5)
super_res = super_t.compute_takeoff()
print(f"Bê tông 4 dầm Super-T: {super_res['girders_concrete_c50_m3']:.2f} m3 (C50)")
print(f"Bê tông bản mặt cầu:    {super_res['deck_concrete_c35_m3']:.2f} m3 (C35)")
print(f"Thảm BTN C12.5 7cm:     {super_res['asphalt_wearing_m2']:.2f} m2")
```

---

## 3. Điều kiện Kích hoạt & Phê duyệt

- **Điều kiện kích hoạt:** Khi lập dự toán, bóc tách khối lượng hình học từ bản vẽ CAD/IFC hoặc rà soát bảng BoQ cầu đường.
- **Bảo toàn Số học (Zero LLM Math Guard):** 100% phép tính thực thi bằng code Python xác định, không làm tròn trung gian, bảo toàn công thức liên kết.
