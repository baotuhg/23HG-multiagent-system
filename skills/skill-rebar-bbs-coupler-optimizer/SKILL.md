---
name: skill-rebar-bbs-coupler-optimizer
description: Tự động bù trừ mối nối buộc 40d cho cây thép dài > 11.7m, tự động phái sinh định mức thép buộc 1.5% theo TT 38/2026/TT-BXD và giải bài toán Cutting Stock bằng Google OR-Tools CP-SAT đạt đề-xê < 1.5%.
---

# Kỹ năng Tối ưu hóa Cắt Thép 1D với Mối Nối 40d & Thép Buộc 1.5%

**Mã Kỹ năng:** `SKILL-REBAR-BBS-COUPLER-OPTIMIZER`  
**Nhóm nghiệp vụ:** `REBAR_OPTIMIZATION`  
**Trạng thái:** `APPROVED` (Kỹ sư trưởng phê duyệt)  
**Tiêu chuẩn áp dụng:** TCVN 5574:2018, TCVN 1651:2018, Thông tư 38/2026/TT-BXD.

---

## 1. Mục tiêu & Bài toán Kỹ thuật

1. **Chiều dài thương mại cố định:** Cây thép tại Việt Nam dài cố định $11.7\text{m} = 11700\text{mm}$.
2. **Quy tắc nối chồng $40d$:**
   - Khi thanh cốt thép yêu cầu có chiều dài $L > 11.7\text{m}$ (ví dụ cọc khoan nhồi sâu 30-40m, cốt dọc thân trụ cao), bắt buộc chia đoạn và bổ sung chiều dài nối chồng:
     $$L_{\text{thực}} = L + n_{\text{nối}} \times 40d$$
   - Mối nối phải được bố trí so le theo TCVN 5574:2018, tránh vùng mô men uốn cực đại (`RULE-REBAR-003`).
3. **Phái sinh định mức dây thép buộc 1 ly ($1.5\%$):**
   - Theo định mức Thông tư 38/2026/TT-BXD: Khối lượng dây thép buộc $1.5\%$ tổng khối lượng cốt thép dự toán (`RULE-TIE-WIRE-008`).
4. **Tối ưu hóa cắt thép 1D (Cutting Stock Problem):**
   - Nạp toàn bộ danh mục thanh cần cắt (`CutDemand`) vào solver Google OR-Tools (kết hợp Column Generation / GLOP và CP-SAT).
   - Mục tiêu: Tối thiểu hóa số cây thép nguyên 11.7m cần dùng và giảm tỷ lệ mẩu thừa phế thải đề-xê $< 1.5\%$.
   - Khi đề-xê $< 1.5\%$, tự động lưu thành **Mẫu cắt thép vàng (Golden Rebar Pattern)** để tái sử dụng trong $O(1)$.

---

## 2. Hướng dẫn Sử dụng Module Python

```python
from tools.cutting_stock_solver import CuttingStockSolver, CutDemand
from tools.civil_and_bridge_takeoff_engine import (
    generate_bridge_pier_rebar_demands,
    generate_super_t_rebar_demands
)

# 1. Sinh danh mục thanh thép thân trụ D25 và dầm Super-T (đã tính nối 40d)
demands_pier = generate_bridge_pier_rebar_demands(col_diam=1.5, col_H=8.5, num_cols=2)
demands_super_t = generate_super_t_rebar_demands(span_length=33.0, num_girders=4)

# 2. Khởi tạo Solver Google OR-Tools CP-SAT
solver = CuttingStockSolver(stock_length=11700, kerf=5)
for d in demands_pier + demands_super_t:
    solver.add_demand(d)

# 3. Giải bài toán tổ hợp tối ưu
result = solver.solve()
print(f"Trạng thái:            {result.status}")
print(f"Số cây thép nguyên:    {result.stock_bars_used} cây 11.7m")
print(f"Tỷ lệ phế liệu đề-xê:  {result.waste_pct:.2f}% (Đạt chuẩn Golden Pattern nếu < 1.5%)")

# 4. Tính dây thép buộc 1 ly theo TT 38/2026/TT-BXD
total_rebar_weight_kg = result.stock_bars_used * 11.7 * 3.853  # D25: 3.853 kg/m
tie_wire_kg = total_rebar_weight_kg * 0.015
print(f"Khối lượng thép D25:   {total_rebar_weight_kg:.2f} kg")
print(f"Dây thép buộc 1.5%:    {tie_wire_kg:.2f} kg")
```

---

## 3. Điều kiện Kích hoạt & Phê duyệt

- **Trigger:** Khi lập bảng thống kê cốt thép (BBS), bóc tách dự toán hoặc xuất sơ đồ cắt phôi cho xưởng gia công cốt thép.
- **Tiêu chuẩn nghiệm thu:** Tỷ lệ đề-xê phế thải $\le 1.5\%$; không nối thép tại vùng kéo căng lớn nhất.
