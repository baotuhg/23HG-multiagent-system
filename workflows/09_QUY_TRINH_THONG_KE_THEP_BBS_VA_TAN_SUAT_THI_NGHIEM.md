# QUY TRÌNH THỐNG KÊ CHI TIẾT CỐT THÉP (BBS), CẤP PHỐI 1M³ VÀ MA TRẬN TẦN SUẤT THÍ NGHIỆM KCS

> **Trục pháp lý và Tiêu chuẩn áp dụng:**
> - **Luật Xây dựng số 135/2025/QH15**
> - **Nghị định số 207/2026/NĐ-CP** về Quản lý chất lượng, thi công xây dựng và bảo trì công trình xây dựng
> - **Thông tư số 38/2026/TT-BXD** về Định mức dự toán xây dựng công trình (Ban hành kèm theo Định mức dự toán xây dựng)
> - **TCVN 4453:1995**: Kết cấu bê tông và bê tông cốt thép toàn khối - Quy phạm thi công và nghiệm thu
> - **TCVN 1651:2018**: Thép cốt bê tông (Thép tròn trơn & thép thanh vằn)
> - **TCVN 6260:2020** & **TCVN 2682:2020**: Xi măng Poóc lăng hỗn hợp và Xi măng Poóc lăng
> - **TCVN 7570:2006**: Cốt liệu cho bê tông và vữa - Yêu cầu kỹ thuật
> - **TCVN 9396:2012** / **ASTM D6760**: Thử nghiệm không phá hủy - Phương pháp xung siêu âm kiểm tra độ đồng nhất cọc khoan nhồi
> - **ASTM D4945 / TCVN 11321:2016**: Phương pháp thử tải trọng động biến dạng lớn đối với cọc (PDA)

---

## 1. Mục tiêu và Nguyên tắc Quản lý Vật tư - KCS

Trong quản lý chất lượng thi công công trình xây dựng (AEC QA/QC), để đảm bảo công trình an toàn chịu lực và nghiệm thu thanh quyết toán chặt chẽ, hệ thống bắt buộc phải giải quyết 3 bài toán cốt lõi:

1. **Thống kê chi tiết cốt thép (Bar Bending Schedule - BBS):**
   - Không chỉ dừng lại ở tổng số tấn theo nhóm đường kính ($D \le 10$, $10 < D \le 18$, $D > 18$), mà phải bóc tách đến **từng số hiệu thanh (Bar Mark)**, đường kính $\varnothing$, cấp độ bền mác thép, hình dạng uốn (Shape Code), chiều dài cắt thanh $L$, số lượng thanh cho từng cấu kiện và toàn bộ công trình.
2. **Phân tích định mức cấp phối vật liệu cho $1\text{ m}^3$ bê tông từng kết cấu:**
   - Mỗi kết cấu công trình chịu các điều kiện làm việc khác nhau: Cọc khoan nhồi (đổ dưới nước bằng ống tremie cần độ sụt lớn và chậm đông kết), Móng mố trụ (bê tông khối lớn), Xà mũ và Bản mặt cầu (bê tông cường độ cao chống thấm), Dầm Super-T DƯL (bê tông mác cao $C45$ cần đạt cường độ sớm $R_{kéo} \ge 85\%$ để kích kéo cáp DƯL).
   - Phân tích rõ định mức $1\text{ m}^3$ gồm: Lượng xi măng PCB40 (kg), Cát vàng ($m^3$), Đá dăm 1x2 ($m^3$), Nước sạch (lít), Phụ gia hóa dẻo/siêu dẻo (lít/kg).
3. **Thiết lập Ma trận Tần suất Thí nghiệm Kiểm soát Chất lượng (Testing Frequency Matrix):**
   - Từ tổng khối lượng vật tư và cấu kiện hoàn thành, tự động tính toán chính xác số lượng tổ mẫu thí nghiệm bắt buộc phải thực hiện theo quy chuẩn quốc gia, phục vụ lập Kế hoạch thí nghiệm KCS (ITP - Inspection and Test Plan) và nghiệm thu chuyển bước thi công.

---

## 2. Bảng Thống kê Chi tiết Cốt thép (BBS) - Sheet `THONG_KE_THEP_CHI_TIET`

Hệ thống trích xuất dữ liệu trực tiếp từ các file bản vẽ CAD và bảng tính chi tiết của Đơn vị Tư vấn Thiết kế (TEDI):
- `03. SUB\03. COC\Coc KN D=1200-19.5.xls` & `04.D1200-19.5.dwg`
- `03. SUB\01. mo\CT-km19.5.xls` & `01.MO-KM19.5.dwg`
- `03. SUB\02.tru\KL-tru-2than-KM19+500.xls` & `Trụ-19+500.dwg`
- `02. SUP\SUPER T 38.2\KLDam.xls` & `04. COT THEP DAM.dwg`, `03. CAP DUL .dwg`, `05 DAM NGANG.dwg`, `07.BAN MAT CAU + LTN.dwg`
- `04. MIS\04.kck\06.BAN QUA DO\KL BAN QUA DO.xls` & `BẢN QUÁ ĐỘ.dwg`
- `04. MIS\04.kck\02.GO LAN CAN\CT GO LAN CAN-19.5.xls` & `Lan Can-19.5.dwg`
- `04. MIS\04.kck\03.KHE CO GIAN\KCGF1.xls` & `Khe Co Gian.dwg`

### Cấu trúc dữ liệu BBS trong Sheet `THONG_KE_THEP_CHI_TIET`:
| Cột | Tên trường | Ý nghĩa kỹ thuật |
| :--- | :--- | :--- |
| **A** | `TT` | Số thứ tự thanh thép |
| **B** | `Hạng mục kết cấu` | Cọc D1200, Mố M1/M2, Trụ T1/T2, Dầm Super-T, Bản mặt cầu, Bản quá độ, Gờ lan can... |
| **C** | `Bộ phận cấu kiện` | Bệ móng, Thân đặc, Xà mũ, Sườn dầm, Cánh dầm, Lưới chịu lực, Đai xoắn... |
| **D** | `Ký hiệu thanh` | Bar Mark (P1, P2a, F1, T1a, G1a, A1, L1...) theo bản vẽ thi công |
| **E** | `Đường kính Ø` | Kích thước danh nghĩa $\varnothing$ (10, 12, 14, 16, 18, 20, 22, 25, 28, 32 mm, Cáp 15.2mm) |
| **F** | `Mác thép` | CB240-T, CB400-V, ASTM A416 Grade 270 |
| **G** | `Hình dạng thanh` | Shape code (S01 thanh thẳng, S02 móc 1 đầu, S06 chữ U, S07 chữ C, S28 bẻ 2 đầu, Đai xoắn...) |
| **H** | `Chiều dài 1 thanh (m)` | Chiều dài thực tế sau khi bẻ uốn theo tiêu chuẩn |
| **I** | `Số thanh / cấu kiện` | Số lượng thanh bố trí trong 1 cấu kiện đơn vị |
| **J** | `Số cấu kiện` | Số lượng cấu kiện tương ứng trên toàn công trình |
| **K** | `Tổng số thanh` | Công thức `=I... * J...` |
| **L** | `Tổng chiều dài (m)` | Công thức `=H... * K...` |
| **M** | `Trọng lượng 1m (kg/m)` | Khối lượng lý thuyết theo TCVN 1651: $q = \frac{\pi \times \varnothing^2}{4} \times 0.00785$ |
| **N** | `Tổng khối lượng (kg)` | Công thức `=L... * M...` |
| **O** | `Tổng khối lượng (Tấn)` | Công thức `=N... / 1000` |
| **P** | `Phân nhóm Ø` | Phân loại `D<=10mm`, `10<D<=18mm`, `D>18mm`, `Cáp DƯL 15.2mm` |
| **Q** | `Tiêu chuẩn & Ghi chú` | TCVN 1651:2018, quy định nghiệm thu vật tư đầu vào |

---

## 3. Phân tích Cấp phối $1\text{ m}^3$ Bê tông - Sheet `CAP_PHOI_1M3_VA_TAN_SUAT`

Bảng định mức cấp phối cho $1\text{ m}^3$ bê tông được thiết kế khoa học cho từng loại kết cấu công trình Cầu Km19+529.080:

```
[Bê tông C10 lót đệm]        --> 225 kg Xi măng | 0.510 m3 Cát | 0.880 m3 Đá 1x2 | 185 L Nước
[Bê tông C25 Bản quá độ/GLC] --> 350 kg Xi măng | 0.460 m3 Cát | 0.860 m3 Đá 1x2 | 175 L Nước | 2.8 L Phụ gia
[Bê tông C30 Cọc khoan nhồi] --> 410 kg Xi măng | 0.470 m3 Cát | 0.840 m3 Đá 1x2 | 180 L Nước | 4.1 L Siêu dẻo chậm đông
[Bê tông C30 Bệ/Thân mố trụ] --> 385 kg Xi măng | 0.450 m3 Cát | 0.850 m3 Đá 1x2 | 175 L Nước | 3.5 L Hóa dẻo
[Bê tông C35 Xà mũ / Mặt cầu] --> 430 kg Xi măng | 0.440 m3 Cát | 0.850 m3 Đá 1x2 | 165 L Nước | 4.3 L Siêu dẻo
[Bê tông C45 Dầm Super-T]    --> 475 kg Xi măng | 0.420 m3 Cát | 0.830 m3 Đá 1x2 | 150 L Nước | 5.7 L Siêu dẻo sớm
[Bê tông C40 Chèn khe co giãn] --> 450 kg Xi măng | 0.460 m3 Cát | 0.800 m3 Đá mi  | 160 L Nước | 4.5 kg Chống co ngót
```

### Tổng hợp nhu cầu vật liệu cấu thành toàn bộ công trình:
$$\text{Khối lượng vật tư} = \sum (\text{Thể tích kết cấu } V_i \times \text{Định mức } 1\text{ m}^3)$$

- **Tổng thể tích bê tông:** $3,460.67\text{ m}^3$
- **Tổng Xi măng PCB40:** $1,406.26$ Tấn ($\approx 28,125$ bao 50kg)
- **Tổng Cát vàng:** $1,565.06\text{ m}^3$
- **Tổng Đá dăm 1x2:** $2,925.37\text{ m}^3$
- **Tổng Nước trộn thi công:** $595.48\text{ m}^3$
- **Tổng Phụ gia hóa dẻo & siêu dẻo:** $13,804.9\text{ lít}$

---

## 4. Ma trận Tần suất Thí nghiệm Vật liệu & KCS Nghiệm thu

Căn cứ các tiêu chuẩn hiện hành, hệ thống thiết lập bảng ma trận tần suất kiểm tra chất lượng (Testing Frequency Matrix):

### 4.1. Cốt thép xây dựng (TCVN 1651:2018)
- **Quy định tần suất:** Cứ $\le 50$ tấn thép cho mỗi lô nhập xưởng, mỗi đường kính $\varnothing$, mỗi mác thép lấy 1 tổ mẫu (gồm 3 thanh thử kéo giới hạn chảy/độ bền kéo, 3 thanh thử uốn nguội $180^\circ$).
- **Số tổ mẫu theo từng cỡ $\varnothing$:**
  * $\varnothing \le 10\text{mm}$ (CB240-T, 14.53T): 1 tổ mẫu
  * $\varnothing 12\text{mm}$ (CB400-V, 18.25T): 1 tổ mẫu
  * $\varnothing 14\text{mm}$ (CB400-V, 14.82T): 1 tổ mẫu
  * $\varnothing 16\text{mm}$ (CB400-V, 62.48T): 2 tổ mẫu
  * $\varnothing 18\text{mm}$ (CB400-V, 68.39T): 2 tổ mẫu
  * $\varnothing 20\text{mm}$ (CB400-V, 15.34T): 1 tổ mẫu
  * $\varnothing 22\text{mm}$ (CB400-V, 16.92T): 1 tổ mẫu
  * $\varnothing 25\text{mm}$ (CB400-V, 98.24T): 2 tổ mẫu
  * $\varnothing 28\text{mm}$ (CB400-V, 72.10T): 2 tổ mẫu
  * $\varnothing 32\text{mm}$ (CB400-V, 53.12T): 2 tổ mẫu
  * Cáp DƯL $15.2\text{mm}$ (ASTM A416, 29.91T): 1 tổ mẫu kéo đứt tao cáp
  * **Tổng số tổ mẫu thí nghiệm thép:** 16 tổ mẫu.

### 4.2. Vật liệu cấu thành bê tông
- **Xi măng PCB40 (TCVN 6260:2020):** Cứ $\le 100 \div 200$ tấn / 1 lô xuất xưởng lấy 1 tổ mẫu thí nghiệm cơ lý (độ mịn, thời gian đông kết, cường độ nén R3, R7, R28) $\rightarrow$ Với $1,406.26$ tấn xi măng cần tối thiểu **15 tổ mẫu**.
- **Cát vàng bê tông (TCVN 7570:2006):** Cứ $\le 200\text{ m}^3$ lấy 1 tổ mẫu thí nghiệm thành phần hạt, mô đun độ lớn, độ bẩn bùn sét $\rightarrow$ Với $1,565.06\text{ m}^3$ cát vàng cần tối thiểu **8 tổ mẫu**.
- **Đá dăm 1x2 (TCVN 7570:2006):** Cứ $\le 200\text{ m}^3$ lấy 1 tổ mẫu thí nghiệm độ nén dập, hạt thoi dẹt, thành phần hạt $\rightarrow$ Với $2,925.37\text{ m}^3$ đá cần tối thiểu **15 tổ mẫu**.
- **Nước thi công (TCVN 4506:2012):** 1 mẫu kiểm tra nguồn nước cấp.
- **Phụ gia bê tông (TCVN 8826:2011):** 3 mẫu kiểm tra tương thích xi măng theo từng lô.

### 4.3. Kiểm tra Bê tông hiện trường & Nghiệm thu KCS (TCVN 3118:1993, TCVN 4453:1995)
- **Đo độ sụt bê tông tươi (TCVN 3106:1993):** Kiểm tra 100% các xe bồn bê tông trước khi cho phép xả ($\approx 432$ lượt kiểm tra).
- **Mẫu nén cọc khoan nhồi C30:** Mỗi cọc lấy 2 tổ mẫu (3 viên R7 + 3 viên R28) $\rightarrow 26 \text{ cọc} \times 2 = \mathbf{52 \text{ tổ mẫu}}$.
- **Mẫu nén mố M1, M2 C30:** $10$ tổ mẫu (chia theo đợt đổ bệ, thân, đỉnh, cánh).
- **Mẫu nén trụ T1, T2 C30 & C35:** $24$ tổ mẫu (8 tổ bệ trụ, 12 tổ thân trụ các đốt đúc, 4 tổ xà mũ trụ).
- **Mẫu nén Dầm Super-T C45:** Mỗi phiến dầm đúc tối thiểu 3 tổ mẫu:
  * 1 tổ nén kiểm tra cường độ đạt $R \ge 85\% f'_c \ge 38.25\text{ MPa}$ trước khi kích kéo cáp DƯL.
  * 1 tổ nén nghiệm thu tuổi 7 ngày (R7).
  * 1 tổ nén nghiệm thu tuổi 28 ngày (R28).
  * $\rightarrow 15 \text{ phiến dầm} \times 3 = \mathbf{45 \text{ tổ mẫu}}$ ($270$ viên mẫu).
- **Mẫu nén Bản mặt cầu, dầm ngang, liên tục nhiệt C35:** $13$ tổ mẫu.
- **Mẫu nén Bản quá độ & Gờ lan can C25:** $10$ tổ mẫu.
- **Mẫu nén bê tông lót đệm C10:** $4$ tổ mẫu.
- **Mẫu nén vữa chèn khe co giãn C40:** $4$ tổ mẫu.

### 4.4. Thí nghiệm Chuyên sâu Cọc khoan nhồi
- **Siêu âm cọc khoan nhồi (ASTM D6760 / TCVN 9396:2012):**
  * Tỷ lệ siêu âm: **100% số cọc** (26/26 cọc D1200).
  * Mỗi cọc đặt 4 ống siêu âm $\rightarrow$ Đo 6 mặt cắt xung siêu âm qua thân cọc $\rightarrow$ **156 mặt cắt siêu âm**.
- **Thử tải động biến dạng lớn PDA (ASTM D4945 / TCVN 11321:2016):**
  * Tỷ lệ thử: Tối thiểu 1% số cọc $\rightarrow$ **1 cọc thử PDA** kiểm tra sức chịu tải thiết kế $Q_{tk} \ge 450$ Tấn và độ nguyên vẹn thân cọc BTA.
- **Khoan kiểm tra đáy cọc và tiếp xúc mũi cọc:**
  * **4 cọc chỉ định** kiểm tra chiều dày lớp cặn lắng ($\le 5\text{ cm}$) và độ ngậm đá gốc ($\ge 1.0\text{ m}$).

---

## 5. Hướng dẫn Tự động hóa qua Python API

Để tạo mới hoặc cập nhật 2 sheet này vào bất kỳ file Excel dự án nào, gọi trực tiếp từ terminal hoặc mã Python:

```bash
python examples/add_rebar_bbs_and_mix_sheets.py
```

Hoặc tích hợp vào code Agent:
```python
from aec_core.material_frequency import ConcreteMixDesigner, QualityTestingFrequencyMatrix
from examples.add_rebar_bbs_and_mix_sheets import add_detailed_rebar_and_mix_sheets

# 1. Bổ sung 2 sheet vào Workbook Master
add_detailed_rebar_and_mix_sheets("templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")

# 2. Tính toán ma trận cấp phối và tần suất
volumes = {"C10": 46.68, "C25": 184.89, "C30_PILE": 1051.33, "C30_SUB": 1246.55, "C35_SUPER": 491.74, "C45_GIRDER": 434.80}
mat_summary = ConcreteMixDesigner.calculate_total_materials(volumes)
testing_plan = QualityTestingFrequencyMatrix.calculate_testing_plan(mat_summary, {"piles": 26, "girders": 15})
```
