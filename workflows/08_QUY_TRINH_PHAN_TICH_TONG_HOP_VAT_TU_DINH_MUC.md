# QUY TRÌNH 08: PHÂN TÍCH ĐỊNH MỨC & TỔNG HỢP VẬT LIỆU TOÀN CÔNG TRÌNH (BOM)
## HỆ THỐNG QUẢN TRỊ KỸ THUẬT & CUNG ỨNG VẬT TƯ: SẮT THÉP TỪNG Ø - XI MĂNG - CÁT - ĐÁ - PHỤ GIA
### Tuân thủ: Thông tư số 38/2026/TT-BXD, Thông tư số 36/2026/TT-BXD & TCVN hiện hành

---

## 1. MỤC TIÊU VÀ NGUYÊN TẮC CỐT LÕI

Trong quản lý dự án xây dựng và công trường, việc chỉ có tổng khối lượng công tác (ví dụ $1.500 m^3$ bê tông hoặc 250 tấn cốt thép) là **chưa đủ để thi công và quản lý tài chính**. Ban Chỉ huy công trường, phòng vật tư và kế toán dự án cần biết chính xác:
1. **Phân tích chi tiết vật liệu theo từng hạng mục (WBS Material Breakdown):** Cọc khoan nhồi cần bao nhiêu tấn xi măng, bao nhiêu khối cát, bao nhiêu khối đá, bao nhiêu tấn thép chủ $\varnothing 25$, bao nhiêu thép đai $\varnothing 10$, bao nhiêu ống siêu âm $\varnothing 60$; Mố M1 cần bao nhiêu thép $\varnothing 32, \varnothing 28, \varnothing 20...$; Dầm Super-T cần bao nhiêu cáp dự ứng lực 15.2mm, bao nhiêu neo chùm, bao nhiêu xi măng cường độ cao...
2. **Tổng hợp nhu cầu vật tư toàn công trình (Consolidated Bill of Materials - BOM):** Tổng hợp lượng xi măng PCB40, cát vàng, đá 1x2, sắt thép từng loại đường kính $\varnothing$ ($\varnothing 10, \varnothing 12, \varnothing 14, \varnothing 16, \varnothing 18, \varnothing 20, \varnothing 22, \varnothing 25, \varnothing 28, \varnothing 32$), cáp DƯL, gối chậu, khe co giãn, vật tư phụ...
3. **Phân kỳ cung ứng vật tư theo giai đoạn thi công (Procurement Schedule):** Chia nhỏ nhu cầu vật tư thành các đợt gọi hàng theo đường găng CPM (Giai đoạn cọc nhồi $\rightarrow$ Giai đoạn mố trụ $\rightarrow$ Giai đoạn dầm nhịp $\rightarrow$ Giai đoạn hoàn thiện mặt cầu).

### NGUYÊN TẮC BẮT BUỘC: 100% CÔNG THỨC ĐỘNG
- **Khối lượng hạng mục:** Lấy trực tiếp bằng công thức liên kết `=QS_DIEN_GIAI_CHI_TIET!J...` (CẤM GÕ SỐ CHẾT).
- **Hao phí vật tư từng cấu kiện:** `=Khối lượng hạng mục * Định mức hao phí TT 38/2026`.
- **Tổng hợp toàn công trình:** Dùng hàm `=SUMIF(Vùng_Tên_Vật_Tư, Tên_Vật_Tư, Vùng_Khối_Lượng)`.
- **Dự trù cung ứng có hao hụt:** `=Tổng_Hao_Phí_Định_Mức * (1 + % Hao_Hụt_Thi_Công_TT12)`.

---

## 2. BẢNG TIÊU CHUẨN ĐỊNH MỨC CẤP PHỐI BÊ TÔNG & HAO PHÍ VẬT LIỆU (TT 38/2026/TT-BXD)

### A. Cấp phối vật liệu cho 1 $m^3$ Bê tông các loại (Xi măng PCB40, Đá 1x2):
| Loại bê tông / Kết cấu | Xi măng PCB40 (kg) | Cát vàng ($m^3$) | Đá dăm 1x2 ($m^3$) | Nước ($lít$) | Phụ gia siêu dẻo / hóa dẻo (kg/lít) |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Bê tông lót C10 (M150)** | 240 | 0.51 | 0.87 | 175 | 0 |
| **Bê tông C30 (M350)** (Cọc nhồi, Bệ móng, Thân mố trụ, Bản quá độ) | 385 - 395 | 0.46 - 0.47 | 0.83 - 0.84 | 180 - 185 | 3.8 - 4.2 |
| **Bê tông C35 (M400)** (Xà mũ trụ, Dầm ngang, Bản mặt cầu, Liên tục nhiệt) | 420 - 425 | 0.45 | 0.83 | 180 | 4.5 |
| **Bê tông C45/C50 (M500/M600)** (Dầm chủ Super-T 38.2m dự ứng lực) | 485 | 0.44 | 0.83 | 165 | 6.8 |

### B. Cơ cấu đường kính cốt thép theo từng kết cấu cầu (TCVN 1651:2018):
| Kết cấu cấu kiện | Đường kính chủ đạo | Loại thép | Mục đích chịu lực |
| :--- | :--- | :--- | :--- |
| **Cọc khoan nhồi D1200** | $\varnothing 25$ (78.5%), $\varnothing 16$ (12.0%), $\varnothing 10$ (9.5%) | CB500-V, CB400-V, CB240-T | Thép chủ chịu uốn dọc, đai tăng cường định hình lồng, đai xoắn liên tục |
| **Bệ móng mố & trụ** | $\varnothing 32$ (68.0%), $\varnothing 20$ (32.0%) | CB500-V | Thép lưới đáy chịu uốn chọc thủng, thép lưới trên & phân bố |
| **Thân mố & Thân trụ** | $\varnothing 32$ (Trụ), $\varnothing 28$ (Mố), $\varnothing 16$ (Đai) | CB500-V, CB400-V | Thép đứng chịu lực lệch tâm lớn, đai giữ ổn định thanh đứng |
| **Xà mũ trụ T1, T2** | $\varnothing 28$ (72.0%), $\varnothing 16$ (28.0%) | CB500-V, CB400-V | Thép chủ chịu mô-men uốn âm gối, thép đai chịu lực cắt dầm console |
| **Dầm chủ Super-T 38.2m** | Cáp $\varnothing 15.2$ (DƯL), $\varnothing 16$ (55%), $\varnothing 14$ (35%), $\varnothing 10$ (10%) | ASTM A416 Gr270, CB400-V, CB240-T | Ứng suất trước căng sau, thép sườn chống cắt, thép cánh dầm, móc cẩu |
| **Bản mặt cầu & Liên tục nhiệt**| $\varnothing 16$ (48.0%), $\varnothing 14$ (42.0%), $\varnothing 10$ (10.0%) | CB400-V, CB240-T | Lưới dưới chịu uốn thuận, lưới trên chịu uốn âm trên sườn dầm, cốt chống nứt |
| **Bản quá độ sau mố** | $\varnothing 16$ (85.0%), $\varnothing 12$ (15.0%) | CB400-V | Thép dọc chịu uốn nhịp bản quá độ khi nền đắp lún, thép phân bố |

---

## 3. CẤU TRÚC 2 SHEET VẬT TƯ TRÊN BẢNG TÍNH EXCEL MASTER

### SHEET 1: `PHAN_TICH_VAT_TU_WBS`
Bảng phân rã chi tiết vật liệu cấu thành cho từng công tác trong WBS:
- Cột A: STT
- Cột B: Mã hiệu định mức TT 38/2026 (AF.21111, AF.12111, AF.61111...)
- Cột C: Tên hạng mục công tác
- Cột D: Khối lượng công tác (`=QS_DIEN_GIAI_CHI_TIET!J...`)
- Cột E: Đơn vị tính công tác ($m^3$, Tấn, m...)
- Cột F: Tên quy cách vật tư thành phần (Xi măng, Cát, Đá, Thép $\varnothing$..., Cáp DƯL, Neo, Phụ gia...)
- Cột G: Đơn vị tính vật tư (kg, Tấn, $m^3$, lít, cái, bộ, m)
- Cột H: Định mức tiêu hao cho 1 đơn vị công tác
- Cột I: Khối lượng tiêu hao cấu kiện `=D... * H...` (Công thức động 100%)
- Cột J: Ghi chú tiêu chuẩn kỹ thuật áp dụng

### SHEET 2: `TONG_HOP_VAT_TU_TOAN_BO`
Bảng gom vật tư toàn bộ công trình và phân kỳ cấp hàng:
1. **Phần A - Sắt thép các loại:**
   - Thép đai tròn trơn $\varnothing \le 10$ (CB240-T)
   - Thép vằn $\varnothing 12$ (CB400-V)
   - Thép vằn $\varnothing 14$ (CB400-V)
   - Thép vằn $\varnothing 16$ (CB400-V)
   - Thép vằn $\varnothing 18$ (CB400-V)
   - Thép vằn $\varnothing 20$ (CB500-V)
   - Thép vằn $\varnothing 22$ (CB500-V)
   - Thép vằn $\varnothing 25$ (CB500-V)
   - Thép vằn $\varnothing 28$ (CB500-V)
   - Thép vằn $\varnothing 32$ (CB500-V)
   - Cáp dự ứng lực tao 15.2mm Gr270 (ASTM A416)
   - Bộ neo chùm DƯL
   - Ống ghen luồn cáp DƯL
   - Ống siêu âm cọc nhồi $\varnothing 60 \times 2$
   - Ống vách thép dẫn hướng D1300
   - Tay vịn lan can thép mạ kẽm nhúng nóng
   - Thép khe co giãn răng lược
   - Dây buộc 1 ly, que hàn
2. **Phần B - Khoáng chất & Cấp phối bê tông:**
   - Xi măng PCB40 (Tấn)
   - Cát vàng sàng rửa $M_k \ge 2.0$ ($m^3$)
   - Đá dăm 1x2 chọn lọc ($m^3$)
   - Nước sạch thi công ($m^3$)
   - Đất đắp nền đường sau mố K98/K95 ($m^3$)
   - Đá hộc $20 \times 30$ cm kè tứ nón ($m^3$)
3. **Phần C - Phụ gia & Hóa chất chuyên dụng:**
   - Phụ gia siêu dẻo giảm nước thế hệ 3 (lít/kg)
   - Phụ gia hoá dẻo chậm đông kết (lít/kg)
   - Vữa rót không co ngót SikaGrout 214-11 (kg)
   - Bột sét Bentonite khoan cọc (Tấn)
4. **Phần D - Phụ kiện & Hoàn thiện cầu:**
   - Gối chậu cao su cốt bản thép (Cái)
   - Màng/Sơn chống thấm mặt cầu đàn hồi ($m^2$)
   - Bê tông nhựa asphalt C12.5 (Tấn)
   - Vải địa kỹ thuật không dệt lọc ngược ($m^2$)
   - Phễu gang + Ống thoát nước $\varnothing 110$ (Bộ / m)
   - Ván khuôn thép / ván phủ phim ($m^2$)

---

## 4. MA TRẬN LIÊN KẾT DỮ LIỆU ĐỘNG VỚI CÁC PHÂN HỆ KHÁC

```
[Bóc tách Takeoff QS_DIEN_GIAI_CHI_TIET]
                   │
                   ▼ (100% Công thức liên kết)
[PHAN_TICH_VAT_TU_WBS] (Khối lượng từng vật tư mỗi hạng mục)
                   │
                   ▼ (=SUMIF)
[TONG_HOP_VAT_TU_TOAN_BO] (Bảng tổng hợp vật tư & phân kỳ cung ứng)
                   │
                   ├──────────────────────────────┬──────────────────────────────┐
                   ▼                              ▼                              ▼
      [Dự toán G_XD TT 36/2026]           [Kế hoạch Cung ứng Vật tư]     [Tiến độ WBS CPM]
      Tính chi phí Vật liệu          Gọi hàng theo mẻ thi công      Kiểm soát chuỗi cung ứng
```

Khi bất kỳ kích thước hình học nào trong hồ sơ thiết kế thay đổi (ví dụ: cọc dài thêm 2m, hoặc mở rộng bề rộng mặt cầu thêm 0.5m), toàn bộ lượng xi măng, cát, đá, sắt thép từng loại đường kính $\varnothing$ và cáp DƯL sẽ **tự động cập nhật 100%** trên toàn bộ hệ thống!
