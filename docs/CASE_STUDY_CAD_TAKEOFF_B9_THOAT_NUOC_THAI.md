# 🌟 BÁO CÁO THỰC CHIẾN: BÓC TÁCH KHỐI LƯỢNG HÌNH HỌC CAD VÀ VẠCH TRẦN SAI LỆCH 4,5 KM CỐNG THOÁT NƯỚC THẢI CỤM B9 (OLYMPIC)

**Chủ quyền tác giả & Hệ thống:** Nguyễn Bảo Tú ([@baotuhg](https://github.com/baotuhg))  
**Hệ thống điều phối:** `23HG-AEC-MultiAgent-System`  
**Dự án áp dụng:** Mạng lưới Thoát nước thải Lô B9.2, B9.3, B9.4 — Khu đô thị Olympic Thường Tín  
**Tập tin bản vẽ gốc:** `261008.MB TNT LÔ B9.2-3-4.dwg` (12.0 MB)  
**Thời gian xử lý & đối soát:** Tháng 10/2026  

---

## Executive Summary (Tóm Tắt Chiến Tích)

Trong công tác quản lý chi phí, đấu thầu và thi công xây dựng (AEC Cost & Tender Engineering), ranh giới giữa **"Số liệu trên giấy/text thuyết minh"** và **"Hình học thực tế dưới lòng đất"** thường tiềm ẩn những sai lệch khổng lồ. 

Tại dự án Thoát nước thải Cụm B9 (Olympic), hệ thống `23HG-AEC-MultiAgent-System` đã:
1. **Quét trực tiếp toàn bộ thực thể CAD 100% qua giao thức MCP:** Đọc chính xác **990 hố ga thu gom chuẩn TB41** và **967 tuyến cống chính** với tổng chiều dài **$22.016,90\text{ m}$**.
2. **Trích xuất tọa độ trắc địa VN-2000 chuẩn xác:** Cung cấp đầy đủ cặp tọa độ Trục X (Bắc $\approx 2.305.000\text{m}$) và Trục Y (Đông $\approx 585.000\text{m}$) cho từng hố ga và 2 đầu từng đoạn cống.
3. **Vạch trần sai lệch kỹ thuật nghiêm trọng của Tư vấn Thiết kế:** Phát hiện toàn bộ hệ thống kích thước (1.285 dimension) bị gán hệ số thu nhỏ $\text{DIMLFAC} = 0.8$, khiến tổng chiều dài cống ghi chú trên bản vẽ bị hụt tới **`4.466,45 m` (hơn 4,46 km ống cống)** so với hình học thực tế.
4. **Bảo vệ an toàn tài chính & tiến độ dự án:** Buộc Tư vấn Thiết kế phải thừa nhận và cập nhật lại toàn bộ hồ sơ bản vẽ; ngăn chặn rủi ro thiếu hụt vật tư trầm trọng ngoài công trường khi đặt mua ống HDPE D300.

```
                    +----------------------------------------------+
                    | BẢN VẼ CAD: MB TNT LÔ B9.2-3-4 (12.0 MB)    |
                    +----------------------+-----------------------+
                                           |
                    +----------------------v-----------------------+
                    |  MCP CAD CONNECTOR & AUTOLISP BATCH ENGINE   |
                    |  - Quét 990 Block hố ga TB41 (MG, DC, VN2000)|
                    |  - Đo 967 Polyline cống (vlax-curve length)  |
                    +----------------------+-----------------------+
                                           |
                        [ PHÁT HIỆN BẤT THƯỜNG ]
                 +-----------------------------------------+
                 | Line chéo vẽ thật : L = 23.80 m         |
                 | Text ghi chú Dim  : L = 19.00 m (L19M)  |
                 | Hệ số cài đặt     : DIMLFAC = 0.80      |
                 | Độ vênh toàn mạng : +4.466,45 m cống    |
                 +--------------------+--------------------+
                                      |
                    +-----------------v----------------------------+
                    | EXCEL MASTER BOQ (100% CÔNG THỨC SỐNG WBS)   |
                    | - 15 Đầu việc Tiên lượng chuẩn TT 36/2026    |
                    | - 990 Ga phân cấp 3 độ sâu theo TT 12/2021   |
                    | - Tích hợp đầy đủ tọa độ trắc địa VN-2000    |
                    +----------------------------------------------+
```

---

## I. Giải Phẫu Sai Lệch Hình Học "Kinh Điển": 23.8m vs 19m

### 1. Hiện trường vụ việc
Kỹ sư kiểm tra trên bản vẽ phát hiện một đoạn cống chéo nối giữa 2 hố ga `AR1.1` và `AR1.2` (Layer `HTKT_TN_D300`, Handle `C8FA52`):
- **Nhãn text ghi chú:** `D300-L19M` (Ống D300, Chiều dài 19m).
- **Đo bằng lệnh `DAL` / `DLI`:** Hiển thị con số `19.00 m`.
- **Đo bằng lệnh `tg` (AutoLISP) hoặc thuộc tính `Length` trong `Ctrl + 1`:** Hiển thị chính xác **`23.80 m`**.

### 2. Chứng minh toán học & Trắc địa VN-2000
Tọa độ trắc địa thực tế của 2 đầu mút đoạn cống `C8FA52`:
- Điểm đầu $P_1$: $X_1 = 585.966,128\text{ m} \,;\, Y_1 = 2.305.036,602\text{ m}$
- Điểm cuối $P_2$: $X_2 = 585.952,872\text{ m} \,;\, Y_2 = 2.305.056,369\text{ m}$

Tính khoảng cách Euclid giữa 2 điểm:
$$\Delta X = |585.966,128 - 585.952,872| = 13,256\text{ m}$$
$$\Delta Y = |2.305.056,369 - 2.305.036,602| = 19,767\text{ m}$$
$$L_{\text{thực}} = \sqrt{\Delta X^2 + \Delta Y^2} = \sqrt{13,256^2 + 19,767^2} = \sqrt{175,72 + 390,73} = \mathbf{23,800\text{ m}}$$

Đoạn cống nằm xiên góc $\alpha \approx 37,1^\circ$ so với trục tọa độ.

### 3. Nguyên nhân: Bí mật của biến `DIMLFAC = 0.8`
Khi truy vấn thuộc tính Dimension Style trong bản vẽ, hệ thống phát hiện toàn bộ kiểu Dim ghi chú cống bị gán biến:
$$\text{DIMLFAC} = \mathbf{0.80} \quad \left(\text{Linear Scale Factor} = \frac{1}{1.25}\right)$$
AutoCAD tự động tính số đo hiển thị:
$$\text{Measurement Text} = L_{\text{thực}} \times \text{DIMLFAC} = 23,80\text{ m} \times 0,8 = 19,04\text{ m}$$
Người thiết kế làm tròn số thành `19m` và gán nhãn `D300-L19M`.

### 4. Đối soát quy mô toàn mạng lưới
Khi tổng hợp toàn bộ 1.283 đoạn ghi chú kích thước trên mặt bằng khu B9:
- **Tổng chiều dài theo Text ghi chú làm tròn:** **`17.794,34 m`**
- **Tổng chiều dài theo Dim (nhân 0.8):** **`17.865,81 m`**
- **Tổng chiều dài theo Hình học thực tế (Line 1:1):** **`22.332,26 m`**
- **Tổng sai lệch thiếu hụt vật tư:** **`+4.466,45 m` (Vượt $20\%$ so với hồ sơ thiết kế)**!

---

## II. Quy Trình Đo Bóc 4 Giai Đoạn Chuẩn (Glass-box Engineering)

Thay vì xử lý dạng "hộp đen" âm thầm, hệ thống 23HG áp dụng quy trình minh bạch 4 giai đoạn:

### Giai đoạn 1: Khám sức khỏe bản vẽ (Pre-flight Inspection)
- Đọc biến đơn vị `$INSUNITS` (mã 6 = mét).
- Quét biến `DIMLFAC`, `DIMSCALE` trong bảng `AcDbDimStyleTable`. Cảnh báo ngay khi `DIMLFAC != 1.0`.
- Kiểm tra cao độ đỉnh $Z$ của các Polyline (`Elevation = 0.00`), loại trừ nguy cơ sai lệch do lỗi bắt điểm 3D.

### Giai đoạn 2: Trích xuất hình học 100% qua AutoLISP / MCP
1. **Module Hố ga (990 Ga):**
   - Lọc đối tượng `AcDbBlockReference` tên `TB41` và `TengaK`.
   - Bóc tách thuộc tính: Tên ga (`T`), Cốt nắp ga (`MG`), Cốt đáy ga (`DC`), Tọa độ thiết kế (`X=...`, `Y=...`).
   - Phân cấp theo định mức Bộ Xây dựng:
     + Cấp 1 ($H \le 1.5\text{m}$): 655 ga (chiều sâu trung bình $1.09\text{m}$)
     + Cấp 2 ($1.5\text{m} < H \le 2.5\text{m}$): 332 ga (chiều sâu trung bình $1.86\text{m}$)
     + Cấp 3 ($H > 2.5\text{m}$): 3 ga (chiều sâu trung bình $2.75\text{m}$)
2. **Module Tuyến cống (967 Tuyến):**
   - Đọc chiều dài thực thể 1:1 qua hàm `(vlax-curve-getDistAtParam ent (vlax-curve-getEndParam ent))`.
   - Lấy tọa độ 2 đầu mút $P_1(X_1, Y_1)$ và $P_2(X_2, Y_2)$.
3. **Module Ghép nối không gian (Spatial Matching):**
   - Thuật toán tìm kiếm bán kính gần nhất $R \le 1.5\text{m}$ giữa mút cống và điểm đặt hố ga.
   - Tự động xác định Ga thượng lưu (Upstream) và Ga hạ lưu (Downstream).
   - Ghép cốt đáy cống $DC_1, DC_2$ để tính toán độ dốc thủy lực thực tế $i = \frac{DC_1 - DC_2}{L}$.

### Giai đoạn 3: Bảng đối soát chéo & Bật cờ cảnh báo (Cross-Check & Red Flag)
- Lập bảng so sánh song song giữa số liệu hình học CAD và số liệu thuyết minh/Text Dim.
- Xuất bảng phân tích chênh lệch vật tư chi tiết để phục vụ kỹ sư lập hồ sơ RFI (Request For Information) gửi Ban Quản lý dự án và Tư vấn Thiết kế.

### Giai đoạn 4: Xuất Master Excel BoQ 100% Công thức sống & Tự động dọn dẹp
- Khởi tạo file Excel Master 5 Sheets:
  + `📑 Tổng Hợp BoQ`: 15 đầu việc định mức dự toán chuẩn Thông tư 36/2026/TT-BXD.
  + `📊 Bóc Tách Cống`: WBS 12 cột bóc tách khối lượng cống HDPE D300 hè, qua đường và cống BTCT D800.
  + `📊 Bóc Tách Hố Ga`: Bóc tách đào móng, bê tông lót M100, ván khuôn, bê tông C20 thành & đáy, cốt thép, nắp ga và đất đắp hoàn trả.
  + `📋 CSDL Hố Ga`: Cơ sở dữ liệu 990 hố ga kèm tọa độ Trục X - Y VN-2000.
  + `📋 CSDL Tuyến Cống`: Cơ sở dữ liệu 967 tuyến cống kèm 4 cột tọa độ điểm đầu và cuối.
- Gọi công cụ **Excel Live MCP** để Microsoft Excel chạy ngầm tính toán và lưu cache toàn bộ 3.048 công thức sống, đảm bảo không có lỗi `#REF!` hay `#VALUE!`.
- **Tự động dọn dẹp:** Xóa sạch 100% file rác trung gian, không để tồn đọng file bừa bộn trong máy tính người dùng.

---

## III. Bảng Tổng Hợp Khối Lượng Tiên Lượng BoQ Khu B9

*(Trích xuất trực tiếp từ file `Boc_Tach_Khoi_Luong_Cong_HoGa_B9_CAD_Master.xlsx`)*

| STT | Mã Hiệu ĐM | Tên Công Tác Thi Công | ĐVT | Khối Lượng | Công Thức Liên Kết Động |
| :---: | :---: | :--- | :---: | :---: | :--- |
| **A** | | **HẠNG MỤC I: HỆ THỐNG CỐNG THOÁT NƯỚC THẢI** | | | **Tổng chiều dài: 22.016,90 m** |
| 1 | `AB.24111` | Đào đất rãnh đặt cống bằng máy + thủ công (đất cấp II) | $\text{m}^3$ | **38.386,72** | `='📊 Bóc Tách Cống'!L13` |
| 2 | `AB.65111` | Đắp cát đệm bảo vệ cống & đắp đất hoàn trả đầm K95 | $\text{m}^3$ | **37.508,08** | `='📊 Bóc Tách Cống'!L18` |
| 3 | `AB.71112` | Vận chuyển đất đào thừa rãnh cống đổ bãi bằng ô tô 10T | $\text{m}^3$ | **10.336,23** | `='📊 Bóc Tách Cống'!L22` |
| 4 | `BB.11101` | Cung cấp & lắp đặt cống HDPE 2 lớp SN4 D300 vỉa hè | $\text{m}$ | **20.442,48** | `='📊 Bóc Tách Cống'!L26` |
| 5 | `BB.11102` | Cung cấp & lắp đặt cống HDPE 2 vách SN8 D300 qua đường | $\text{m}$ | **1.617,13** | `='📊 Bóc Tách Cống'!L30` |
| 6 | `BB.12104` | Cung cấp & lắp đặt cống BTCT đúc sẵn D800 qua đường | $\text{m}$ | **177,46** | `='📊 Bóc Tách Cống'!L34` |
| **B** | | **HẠNG MỤC II: HỆ THỐNG HỐ GA THOÁT NƯỚC THẢI** | | | **Tổng số lượng: 990 Hố ga** |
| 7 | `AB.25111` | Đào đất hố móng ga thoát nước thải (đất cấp II) | $\text{m}^3$ | **5.409,69** | `='📊 Bóc Tách Hố Ga'!L14` |
| 8 | `AF.11111` | Bê tông lót móng hố ga M100 đá 4x6 dày 100mm | $\text{m}^3$ | **222,75** | `='📊 Bóc Tách Hố Ga'!L18` |
| 9 | `AF.81111` | Ván khuôn bê tông lót móng hố ga | $\text{m}^2$ | **594,00** | `='📊 Bóc Tách Hố Ga'!L22` |
| 10 | `AF.21111` | Bê tông thành và đáy hố ga C20/M200 | $\text{m}^3$ | **1.038,28** | `='📊 Bóc Tách Hố Ga'!L27` |
| 11 | `AF.82111` | Ván khuôn thân và thành hố ga (2 mặt trong & ngoài) | $\text{m}^2$ | **10.497,48** | `='📊 Bóc Tách Hố Ga'!L31` |
| 12 | `AF.61111` | Cốt thép hố ga CB300-V (đường kính $\le$ 18mm) | $\text{kg}$ | **67.487,92** | `='📊 Bóc Tách Hố Ga'!L35` |
| 13 | `BB.81111` | Sản xuất & lắp đặt bộ nắp hố ga gang cầu / composite | $\text{bộ}$ | **990,00** | `='📊 Bóc Tách Hố Ga'!L39` |
| 14 | `AB.65112` | Đắp đất hố móng xung quanh mang ga hoàn trả K95 | $\text{m}^3$ | **3.007,64** | `='📊 Bóc Tách Hố Ga'!L43` |
| 15 | `AB.71113` | Vận chuyển đất thừa hố ga đổ đi bãi tập kết bằng ô tô 10T | $\text{m}^3$ | **2.762,36** | `='📊 Bóc Tách Hố Ga'!L47` |

---

## IV. Kết Luận & Bài Học Kinh Nghiệm Quản Trị Kỹ Thuật

1. **Tuyệt đối không tin tưởng mù quáng vào text ghi chú:** Trong các bản vẽ quy hoạch phân lô hoặc hạ tầng kỹ thuật, text ghi chú rất dễ bị sai lệch do thói quen copy paste hàng loạt hoặc áp đặt kích thước phân lô thẳng vào các tuyến cống đi chéo.
2. **Luôn kiểm tra biến `DIMLFAC` trước khi lấy số liệu:** Hệ số tỷ lệ kích thước là nguyên nhân hàng đầu gây sai lệch giữa người thiết kế và đơn vị thi công.
3. **Sức mạnh của tự động hóa đo bóc hình học:** Bằng cách quét trực tiếp hình học CAD và liên kết dữ liệu vào Excel 100% công thức sống, hệ thống `23HG-AEC-MultiAgent-System` đã chứng minh năng lực kiểm toán độc lập mạnh mẽ, bảo vệ quyền lợi hợp pháp và an toàn tài chính cho dự án.
