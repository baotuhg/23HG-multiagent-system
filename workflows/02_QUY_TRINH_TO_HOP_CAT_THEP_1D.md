# QUY TRÌNH 02: TỔ HỢP CẮT THÉP 1D TRÊN CÂY NGUYÊN 11.7M
## BÀI TOÁN 1D CUTTING STOCK - ÉP TỶ LỆ HAO HỤT ĐỀ-XÊ < 1.5%

---


> **Trạng thái:** ✅ **Đã hiện thực** — `tools/cutting_stock_solver.py` (OR-Tools),
> `core/agents/rebar_agent.py`; có kiểm tra vùng nối chồng và kiểm thử đơn vị.

### I. NGUYÊN TẮC CỐT LÕI
1. **Quy chuẩn cây nguyên:** Chiều dài cây thép tiêu chuẩn của các nhà máy sản xuất (Hòa Phát, Việt Ý, Pomina) là **$11.70\text{ m}$**.
2. **Mục tiêu:** Ghép nối các đoạn cắt theo bảng thống kê thép từ bản vẽ thiết kế vào các cây nguyên sao cho tổng số cây nguyên là ít nhất và lượng đầu thừa phôi cắt (đề-xê) là nhỏ nhất.
3. **Mối nối:**
   - Cọc khoan nhồi sâu $> 11.7\text{m}$: Dùng ống nối ren cơ khí (Coupler) hoặc hàn chập nối so le theo TCVN 4453:1995.
   - Thép chủ dầm và mố trụ: Chiều dài mối nối buộc tối thiểu $30D - 40D$ hoặc hàn nối so le.

### II. CẤU TRÚC BẢNG TÍNH & CÔNG THỨC SỐNG
- Cột A (1): TT
- Cột B (2): Cấu kiện sử dụng
- Cột C (3): Ký hiệu đường kính $\Phi$ (mm)
- Cột D (4): Số lượng thanh cắt $N$
- Cột E (5): Chiều dài 1 đoạn cắt $L$ (m)
- Cột F (6): Tổng chiều dài yêu cầu $= D \times E$ (m)
- Cột G (7): Chiều dài cây tiêu chuẩn $= 11.70\text{ m}$
- Cột H (8): Số cây nguyên 11.7m cần dùng:
  ```excel
  =ROUNDUP(F{row} / G{row}, 0)
  ```
- Cột I (9): Chiều dài phôi thép xuất kho $= H \times G$ (m)
- Cột J (10): Đề-xê thừa $= I - F$ (m)
- Cột K (11): Tổng trọng lượng thép $= F \times \text{Trọng lượng riêng theo TCVN 1651}$ (kg)
- Cột L (12): Tỷ lệ hao hụt đề-xê:
  ```excel
  =(J{row} / I{row}) * 100
  ```

### III. BẢNG TRA TRỌNG LƯỢNG THÉP TIÊU CHUẨN (TCVN 1651:2018)
- $\Phi 10$: $0.617\text{ kg/m}$
- $\Phi 12$: $0.888\text{ kg/m}$
- $\Phi 14$: $1.208\text{ kg/m}$
- $\Phi 16$: $1.578\text{ kg/m}$
- $\Phi 18$: $2.000\text{ kg/m}$
- $\Phi 20$: $2.466\text{ kg/m}$
- $\Phi 25$: $3.853\text{ kg/m}$
- $\Phi 28$: $4.834\text{ kg/m}$
- $\Phi 32$: $6.313\text{ kg/m}$
