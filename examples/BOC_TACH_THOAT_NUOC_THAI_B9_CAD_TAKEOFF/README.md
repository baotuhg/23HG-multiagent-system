# 🌊 Bóc Tách Khối Lượng Toàn Diện Cống & Hố Ga Thoát Nước Thải Cụm B9 (Olympic)

**Tác giả & Chủ quyền hệ thống:** Nguyễn Bảo Tú ([@baotuhg](https://github.com/baotuhg))  
**Hệ thống điều phối:** `23HG-AEC-MultiAgent-System` (Module: `AEC-CAD-Takeoff & QS Engine`)  
**Dự án:** Hạ tầng kỹ thuật Thoát nước thải Cụm B9 (Lô B9.2, B9.3, B9.4) — Khu đô thị Olympic  
**Căn cứ pháp lý & định mức:** Luật Xây dựng 135/2025/QH15, Thông tư 36/2026/TT-BXD, Thông tư 37/2026/TT-BXD, Thông tư 38/2026/TT-BXD.

---

## 🏆 Chiến Tích Thực Chiến (Production Case Study)

Trong quá trình triển khai bóc tách hình học tự động từ bản vẽ AutoCAD `261008.MB TNT LÔ B9.2-3-4.dwg`, hệ thống **23HG Multi-Agent System** đã phát hiện một **sai lệch kỹ thuật nghiêm trọng mang tính hệ thống** của hồ sơ Tư vấn Thiết kế:

### 1. Hiện tượng sai lệch "Kinh điển"
- Trên bản vẽ, nhiều đoạn cống chéo ghi chú text là **`D300-L19M`** (Đường kính 300, Chiều dài 19m).
- Khi kỹ sư dùng lệnh đo kích thước thông thường (`DAL`/`DLI`), kích thước hiển thị cũng ra **`19.00 m`**.
- Tuy nhiên, khi hệ thống 23HG chạy thuật toán trắc đạc hình học CAD (`(vlax-curve-getDistAtParam ...)` và lệnh `tg` / `c:dnbntg`), chiều dài thực tế của nét vẽ Polyline lại là **`23.80 m`** (chênh lệch tới $4.8\text{ m}$ trên một đoạn cống!).

### 2. Vạch trần nguyên nhân gốc rễ (Root Cause Analysis)
Hệ thống tiến hành kiểm tra sâu các biến hệ thống CAD và phát hiện:
- Toàn bộ **1.285 đối tượng Dimension** trên bản vẽ đã bị bên Thiết kế cài đè hệ số tỷ lệ đo tuyến tính:
  $$\text{DIMLFAC} = \mathbf{0.80} \quad \left(\text{tương ứng hệ số thu nhỏ } \frac{1}{1.25}\right)$$
- Đây là lỗi do người vẽ tạo DimStyle để đo trên khung in Viewport Layout (tỷ lệ 1:1250) nhưng lại đem đo trực tiếp ngoài không gian Model:
  $$L_{\text{hiển thị trên Dim}} = L_{\text{thực}} \times 0.8 = 23.80\text{ m} \times 0.8 = 19.04\text{ m} \longrightarrow \text{Làm tròn thành } \mathbf{D300-L19M}$$
- Đoạn cống thực sự nằm xiên góc $37^\circ$ với tọa độ trắc địa VN-2000:
  - Đầu 1 ($P_1$): $X = 585.966,13\text{ m} \,;\, Y = 2.305.036,60\text{ m}$
  - Đầu 2 ($P_2$): $X = 585.952,87\text{ m} \,;\, Y = 2.305.056,37\text{ m}$
  - Chiều dài cạnh huyền Pitago: $\sqrt{13.256^2 + 19.767^2} = \mathbf{23.80\text{ m}}$.

### 3. Hậu quả thực tế & Giá trị bảo vệ dự án
| Tiêu chí | Tổng L theo Text / Dim ghi chú | Tổng L theo Hình học CAD 1:1 | Chênh lệch thiếu hụt |
| :--- | :---: | :---: | :---: |
| **Cống D300 vỉa hè & lòng đường** | $17.447,00\text{ m}$ | $21.898,06\text{ m}$ | **$+4.379,61\text{ m}$** |
| **Cống BTCT D800 qua đường** | $347,34\text{ m}$ | $434,20\text{ m}$ | **$+86,84\text{ m}$** |
| **TỔNG TOÀN KHU B9** | **`17.794,34 m`** | **`22.332,26 m`** | **`+4.466,45 m` (hơn 4,46 km!)** |

> [!CAUTION]
> Nếu bóc tách theo Text Dim của hồ sơ thiết kế ($17.794\text{ m}$), công trường khi triển khai thi công sẽ **bị thiếu hụt tới 4,46 km ống cống**, gây gián đoạn thi công, vỡ phương án mua sắm vật tư và gây tranh chấp hợp đồng nghiêm trọng. Nhờ số liệu bóc tách hình học chính xác của hệ thống 23HG, **bên Tư vấn Thiết kế đã phải thừa nhận sai sót và cập nhật lại toàn bộ hồ sơ bản vẽ**.

---

## 📊 Quy Mô Dữ Liệu Bóc Tách Thực Tế (100% CAD Geometry)

Bảng tính Master Excel `Boc_Tach_Khoi_Luong_Cong_HoGa_B9_CAD_Master.xlsx` được thiết lập theo chuẩn **WBS 12 cột Bộ Xây dựng**, liên kết động **100% công thức sống (Zero Dead Numbers)**:

### 1. Cơ sở dữ liệu 990 Hố ga Thoát nước thải (`📋 CSDL Hố Ga`)
- **990 Block hố ga** chuẩn `TB41` trích xuất trực tiếp qua AutoLISP / MCP.
- Tọa độ trắc địa nhà nước **VN-2000**:
  - Trục X (Bắc): $2.304.635,509\text{ m} \div 2.305.726,026\text{ m}$
  - Trục Y (Đông): $584.907,305\text{ m} \div 585.966,065\text{ m}$
- Cao độ tự nhiên đỉnh nắp ga ($MG$), cốt đáy cống ($DC$) và chiều sâu đào $H = MG - DC$.
- Phân loại chiều sâu đào đất theo định mức BXD:
  - Cấp 1 ($H \le 1.5\text{m}$): 655 ga
  - Cấp 2 ($1.5\text{m} < H \le 2.5\text{m}$): 332 ga
  - Cấp 3 ($H > 2.5\text{m}$): 3 ga

### 2. Cơ sở dữ liệu 967 Tuyến cống Thoát nước thải (`📋 CSDL Tuyến Cống`)
- **967 Đoạn tim cống chính** với tổng chiều dài hình học: **`22.016,90 m`**.
- Đầy đủ 4 tọa độ tim mốc định vị VN-2000:
  - Trục X Điểm Đầu, Trục Y Điểm Đầu
  - Trục X Điểm Cuối, Trục Y Điểm Cuối
- Ghép nối không gian (Spatial Matching): Tự động khớp Ga đầu (Upstream) và Ga cuối (Downstream) trong bán kính $R \le 1.5\text{m}$.
- Tính toán độ dốc thủy lực $i = \frac{DC_1 - DC_2}{L}$.

### 3. Bảng Tổng Hợp Tiên Lượng BoQ (15 Đầu Việc Chuẩn Bộ Xây Dựng)
1. **AB.24111** — Đào đất rãnh đặt cống bằng máy + thủ công: **$38.386,72\text{ m}^3$**
2. **AB.65111** — Đắp cát đệm bảo vệ cống & đắp đất hoàn trả K95: **$37.508,08\text{ m}^3$**
3. **AB.71112** — Vận chuyển đất đào thừa rãnh cống bằng ô tô 10T: **$10.336,23\text{ m}^3$**
4. **BB.11101** — Cung cấp & lắp đặt cống HDPE 2 lớp SN4 D300 vỉa hè: **$20.442,48\text{ m}$**
5. **BB.11102** — Cung cấp & lắp đặt cống HDPE 2 vách SN8 D300 qua đường: **$1.617,13\text{ m}$**
6. **BB.12104** — Cung cấp & lắp đặt cống BTCT đúc sẵn D800 qua đường: **$177,46\text{ m}$**
7. **AB.25111** — Đào đất hố móng ga thoát nước thải: **$5.409,69\text{ m}^3$**
8. **AF.11111** — Bê tông lót móng hố ga M100 đá 4x6: **$222,75\text{ m}^3$**
9. **AF.81111** — Ván khuôn bê tông lót móng hố ga: **$594,00\text{ m}^2$**
10. **AF.21111** — Bê tông thành và đáy hố ga C20/M200: **$1.038,28\text{ m}^3$**
11. **AF.82111** — Ván khuôn thân và thành hố ga 2 mặt: **$10.497,48\text{ m}^2$**
12. **AF.61111** — Cốt thép hố ga CB300-V: **$67.487,92\text{ kg}$** (~ 67.49 tấn)
13. **BB.81111** — Sản xuất & lắp đặt bộ nắp hố ga gang cầu / composite: **$990\text{ bộ}$**
14. **AB.65112** — Đắp đất hố móng ga hoàn trả K95: **$3.007,64\text{ m}^3$**
15. **AB.71113** — Vận chuyển đất thừa hố ga đổ đi bãi tập kết bằng ô tô 10T: **$2.762,36\text{ m}^3$**

---

## 🛠️ Danh Mục File & Mã Nguồn Đi Kèm

- **[`Boc_Tach_Khoi_Luong_Cong_HoGa_B9_CAD_Master.xlsx`](Boc_Tach_Khoi_Luong_Cong_HoGa_B9_CAD_Master.xlsx):** File Excel Master 5 Sheets đầy đủ công thức sống, bảng BoQ, bảng bóc tách chi tiết và 2 bảng CSDL trắc địa.
- **[`generate_cad_takeoff_master.py`](generate_cad_takeoff_master.py):** Script Python tự động sinh file Excel Master, định dạng form mẫu chuẩn Bộ Xây dựng và thiết lập toàn bộ công thức liên kết động.
- **[`cad_blocks.csv`](cad_blocks.csv):** Dữ liệu hình học gốc của 990 hố ga trích xuất từ AutoCAD.
- **[`cad_plines.csv`](cad_plines.csv):** Dữ liệu hình học gốc của 967 đoạn cống (chiều dài, layer, handle).
- **[`cad_pline_coords.csv`](cad_pline_coords.csv):** Tọa độ điểm đầu $P_1(X_1, Y_1)$ và điểm cuối $P_2(X_2, Y_2)$ của từng đoạn cống.
