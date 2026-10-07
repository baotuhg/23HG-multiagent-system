# BẢNG MA TRẬN PHÂN QUYỀN & BIÊN BẢN BÀN GIAO 5 GÓI VỆ TINH THỰC CHIẾN
## DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) — CẦU KM19+529.080
**Mô hình điều phối:** Hub & Spoke Role-Based Model (Quy trình 15 - 23HG MultiAgent System)  
**Tiêu chuẩn tuân thủ:** Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Nghị định 254/2025/NĐ-CP & Vincons Standards  
**Ngày lập & bàn giao:** 01/10/2026  

---

### I. TẠI SAO PHẢI PHÂN CHIA THÀNH 5 GÓI VỆ TINH PHÂN QUYỀN?
Trong môi trường công trường thực chiến, việc sử dụng chung 1 file Monolithic 14 Sheet dẫn đến 3 rủi ro chí mạng:
1. **Xung đột file khóa (Read-Only Lock):** Khi Kỹ sư QS đang mở file tính toán giải ngân thì Đội xe máy không thể cập nhật tích kê tiêu thụ dầu Diezel.
2. **Lộ lọt bí mật tài chính & thương mại:** Thợ gia công sắt thép hay lái máy xúc có thể nhìn thấy toàn bộ đơn giá dự toán, định mức lợi nhuận và chi phí gián tiếp của Tổng thầu.
3. **Quá tải trên thiết bị di động hiện trường:** File tổng hợp quá lớn làm lag thiết bị tại hố móng và rất dễ dẫn đến lỗi `#REF!` do vô tình xóa cột.

Hệ thống chuẩn hóa giải pháp **Hub & Spoke** phân tách độc lập thành 5 gói hồ sơ chuyên biệt theo đúng thẩm quyền và trách nhiệm.

---

### II. BẢNG MA TRẬN PHÂN QUYỀN TRUY CẬP HIỆN TRƯỜNG (RACI MATRIX)

| Gói Vệ Tinh | Bộ phận sử dụng | Mức độ bảo mật | Quyền xem (Read) | Quyền sửa (Write) | Dữ liệu cấm tuyệt đối (Restricted) |
| :--- | :--- | :---: | :--- | :--- | :--- |
| **GÓI A: Cơ giới & Dầu** | Đội trưởng xe máy, Thủ kho dầu, Lái máy | **Nội bộ** | Tiến độ ca máy, phụ tải, ĐM dầu, nhật trình xe | Ghi số giờ máy, tích kê cấp dầu hàng ngày | **CẤM XEM** Đơn giá tiền, Dự toán $G_{XD}$, Doanh thu |
| **GÓI B: Xưởng cốt thép** | Quản đốc xưởng, Thợ uốn cắt, Máy CNC | **Nội bộ** | Sơ đồ cắt 11.7m, BBS 396 dòng, CSV nạp CNC | Báo cáo số thanh đã cắt, mã đề-xê lưu kho | **CẤM XEM** Đơn giá mua thép, Giá trị hợp đồng |
| **GÓI C: Hiện trường KCS** | Kỹ sư QA/QC, TVGS, Thí nghiệm LAS-XD | **Quan trọng** | 22 BBNT Word, ma trận ngày KCS, cấp phối mẫu | Nhập kết quả thí nghiệm R7/R28, ký biên bản | **CẤM XEM** Đơn giá dự toán chi tiết, Thanh toán 03a |
| **GÓI D: QS & Dự toán** | Kỹ sư QS, Ban Đấu thầu, Kế toán | **Mật cao** | Hình học Takeoff, Dự toán $G_{XD}$, Thanh toán 03a | Lập dự toán điều chỉnh, hồ sơ giải ngân đợt | Chỉ lưu hành nội bộ phòng Kế hoạch & Ban Giám đốc |
| **GÓI E: Executive Hub** | Giám đốc Dự án, Ban QLDA, Chủ đầu tư | **Tối cao** | Toàn quyền KPI, Đường găng CPM, Audit 100/100 | Phê duyệt tổng thể, ký giải ngân, chuẩn thuận phát sinh | Truy cập toàn bộ dữ liệu hệ sinh thái |

---

### III. DANH MỤC 25+ TẬP TIN HỒ SƠ 5 GÓI VỆ TINH BÀN GIAO

#### 1. GÓI A: ĐỘI CƠ GIỚI, XE MÁY & QUẢN LÝ DẦU DIEZEL
- `260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx`: Kế hoạch điều phối 12 đầu máy chính & Sổ theo dõi cấp phát dầu Diezel.
- `260920_Tien_Do_CaMay_Cau_Km19+529.080.xml`: Tiến độ huy động thiết bị tích hợp MS Project.

#### 2. GÓI B: QUẢN ĐỐC XƯỞNG TIỀN CHẾ & GIA CÔNG CỐT THÉP
- `01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`: Sơ đồ cắt thép 1D CSP 11.7m toàn cầu (33.212 cây 11.7m = 662.890 kg thép).
- `00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx`: Bảng Dashboard chỉ tiêu kinh tế kỹ thuật của 11 đường kính Ø và Cáp DƯL.
- `04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx`: Thống kê thép chi tiết 396 dòng phục vụ thợ uốn bẻ tại bãi.
- `01_Phieu_Cat_Thep_Cau_Km19+529.080.csv`: Tệp lệnh cắt CNC nạp thẳng máy cắt đa thanh Shearline.
- `THEO_TUNG_DUONG_KINH_PHI/`: 11 file chuyên sâu cho từng phi từ Ø8 đến Ø32 và cáp DƯL 15.2mm.
- `README_QUY_TRINH_VAN_HANH_BAI_THEP.md`: Hướng dẫn quản trị phôi thừa đề-xê $\ge 100D$.

#### 3. GÓI C: HIỆN TRƯỜNG QUẢN LÝ CHẤT LƯỢNG KCS & THÍ NGHIỆM
- `Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx`: Trọn bộ 22 Biên bản nghiệm thu KCS chuẩn NĐ 207/2026 sẵn sàng in ký.
- `11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx`: Bảng quản lý 22 BBNT, kiểm soát chéo ngày yêu cầu và ngày ký.
- `12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx`: Mẫu in A4 tự động nhảy số theo mã công việc.
- `13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx`: Mẫu in A4 nghiệm thu vật liệu đầu vào (Thép, Xi măng, Cát, Đá, Gối cầu).
- `14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx`: Mẫu in A4 lấy mẫu nén bê tông R7, R28.
- `05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx`: Định mức cấp phối và kế hoạch lấy 809 mẫu QA/QC.

#### 4. GÓI D: KỸ SƯ QS, BAN DỰ TOÁN & THANH TOÁN HỢP ĐỒNG
- `03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx`: 101 dòng hình học diễn giải tiên lượng chi tiết.
- `08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx`: Bảng tính tổng hợp kinh phí $G_{XD} = T + GT + TL + VAT$ (Bảo mật đơn giá).
- `09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx`: Bảng thanh toán giải ngân Mẫu 03.a theo Nghị định 254/2025/NĐ-CP.
- `06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx`: 140 dòng phân tích vật tư định mức TT 38/2026.
- `07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx`: Tổng hợp BOM và kế hoạch cung ứng 4 phân đợt.

#### 5. GÓI E: EXECUTIVE DASHBOARD & CHỦ ĐẦU TƯ / BAN LÃNH ĐẠO
- `Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx`: Master Workbook 14 Sheet liên kết động 100% (Zero Dead Numbers).
- `Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp`: File tiến độ gốc MS Project quản lý 36 công tác đường găng CPM.
- `Tien_Do_Thi_Cong_Cau_Km19+529.080.xml`: File tiến độ liên thông Primavera / MS Project.
- `BAO_CAO_THAM_TRA_AEC_AUDIT.md`: Báo cáo Thẩm tra Kỹ thuật Độc lập Audit Score 100/100.
- `Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md`: Thuyết minh biện pháp thi công 8 chương chuẩn TCVN.

---

### IV. BIÊN BẢN BÀN GIAO VÀ CAM KẾT HIỆN TRƯỜNG
1. **Bên bàn giao:** Kỹ sư Trưởng Hệ thống 23HG Multi-Agent System bàn giao nguyên trạng đầy đủ các tệp số cho 5 bộ phận.
2. **Bên nhận bàn giao:** Đại diện 5 bộ phận cam kết tiếp nhận đúng quyền hạn, không can thiệp làm hỏng cấu trúc liên kết và tuân thủ 100% quy chế bảo mật dữ liệu.
