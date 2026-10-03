---
name: skill-rebar-bbs-coupler-optimizer
description: Sinh danh mục đoạn cắt cốt thép sơ bộ cho thân trụ và dầm Super-T rồi tối ưu cắt từ cây 11,7 m bằng CuttingStockSolver (OR-Tools), có tùy chọn tách thanh dài và nối chồng 40d.
---

# Tối ưu cắt thép 1D cho cấu kiện cầu (nối chồng 40d, dây buộc 1,5%)

**Mã:** `SKILL-REBAR-BBS-COUPLER-OPTIMIZER` · **Nhóm:** `REBAR_OPTIMIZATION`
**Trạng thái:** `PENDING_APPROVAL` — chờ Kỹ sư trưởng duyệt.
**Mã nguồn:** [`tools/cutting_stock_solver.py`](../../tools/cutting_stock_solver.py), [`tools/civil_and_bridge_takeoff_engine.py`](../../tools/civil_and_bridge_takeoff_engine.py)

## 1. Code thực sự làm gì

- `generate_bridge_pier_rebar_demands(...)` sinh 2 nhóm thanh cho thân trụ: thép chủ D28 dài `H + 2×40d` (giả định **32 thanh/cột**) và đai D14 bước 150 mm.
- `generate_super_t_rebar_demands(...)` sinh thép sườn D20, thép bầu D25 (đoạn **5.800 mm cố định**) và đai chữ U D12 dài 3.850 mm.
- ⚠️ Các số này là **giả định điển hình trong code, không lấy từ bản vẽ**. Muốn dùng thật phải đưa BBS thật vào (`--bbs`).
- Hai hàm trên trả về `list[dict]`, phải chuyển sang `CutDemand(**d)` trước khi đưa vào solver.
- Solver tách theo đường kính và mác thép, tính hao hụt lưỡi cắt và báo cận dưới số cây. `OPTIMAL` nghĩa là đã chứng minh không thể dùng ít cây hơn.
- Thanh dài hơn 11,7 m: dùng `solve(..., split_long_bars=True)`. Phương án nối tận dụng đầu thừa: `allow_splicing=True`. Nối chồng mặc định `lap_xd=40`, kỹ thuật phải duyệt vị trí nối.

> Kết quả demo "đề-xê 0,83%" là trường hợp dễ: 64 đoạn D20 dài 5.800 mm, cắt 2 đoạn mỗi cây. Không nên dùng con số này làm chỉ tiêu chung. Ví dụ thân trụ dưới đây cho kết quả **OPTIMAL nhưng đề-xê khoảng 14%**: D28 dài 10.740 mm chỉ cắt được 1 thanh mỗi cây, đai D14 cũng hao nhiều. Đó là giới hạn của chính danh mục thanh, không phải lỗi solver.
>
> ⚠️ Cùng một trụ, `BridgePierEngine` ước tính thép theo hàm lượng 125 kg/m³ ra khoảng **14,9 tấn**. Trong khi đó danh mục thanh giả định ở đây chỉ khoảng **4,4 tấn**. Hai cách đang không khớp nhau, phải lấy BBS thật làm chuẩn.

## 2. Cách dùng (đã chạy thử)

```python
from tools.cutting_stock_solver import CuttingStockSolver, CutDemand
from tools.civil_and_bridge_takeoff_engine import generate_bridge_pier_rebar_demands

raw = generate_bridge_pier_rebar_demands(column_height_m=8.5, column_dia_m=1.5, column_count=2)
demands = [CutDemand(**d) for d in raw]

solver = CuttingStockSolver(bar_length_mm=11700, kerf_mm=3)
sol = solver.solve(demands)

print("Trạng thái :", sol.status)
print("Số cây 11,7m:", sol.total_bars_needed, "(cận dưới", sol.lower_bound_bars, ")")
print("Đề-xê       :", round(sol.waste_ratio_pct, 2), "%")
print("KL thép mua :", round(sol.total_weight_kg, 1), "kg")
print("Dây buộc 1,5%:", round(sol.total_weight_kg * 0.015, 1), "kg")
for g in sol.groups:
    print(f"  Ø{g.diameter_mm}: {g.total_bars_needed} cây, đề-xê {g.waste_ratio_pct:.2f}% [{g.status}]")
```

Với BBS thật, dùng dòng lệnh:

```bash
python run_state_graph.py --phase rebar --bbs "duong_dan/BBS.xlsx" --kerf-mm 3 --cut-plan-out phieu_cat.csv
# thêm --splice để tính phương án nối tận dụng đầu thừa (cần kỹ thuật duyệt)
```

## 3. Lưu ý kỹ thuật

- Không nối thép tại vùng mô men lớn nhất. Dùng `--splice-zone` để giới hạn vùng được nối.
- Hệ số dây buộc 1,5% lấy từ bảng tính trong folder "Mua"; kỹ sư cần đối chiếu với định mức áp dụng cho dự án.
