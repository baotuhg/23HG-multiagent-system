# TRẠNG THÁI PHÁP LÝ CỦA CÁC VĂN BẢN HỆ THỐNG ĐANG TRÍCH DẪN

> **Mục đích:** ghi rõ văn bản nào hệ thống đang dùng, đã được kiểm chứng ở mức nào, và cần ai xác nhận.
> **Quy tắc:** một văn bản chỉ được ghi là *ĐÃ ĐỐI CHIẾU BẢN GỐC* khi đã đọc nội dung từ nguồn chính thức
> (Công báo, Cổng pháp luật Chính phủ, hoặc văn bản của Bộ). Nguồn thứ cấp (trang luật, trang tổng hợp) chỉ là gợi ý.
> **Ngày rà soát lần cuối:** 2026-10-08. Môi trường phát triển không truy cập được `datafiles.chinhphu.vn`
> và `qlda.gxd.vn`, nên *không có văn bản nào trong bảng dưới đây được đối chiếu bản gốc*.

## Trạng thái: CHƯA ĐỐI CHIẾU BẢN GỐC cho mọi văn bản

| Văn bản | Dùng ở đâu trong hệ thống | Hiệu lực (theo nguồn thứ cấp) | Mức kiểm chứng hiện tại | Cần ai xác nhận |
|---|---|---|---|---|
| **TT 13/2021/TT-BXD**, Phụ lục VI | Quy tắc đo bóc (`PROFILE_TT13_2021_PL_VI`, `tt13-2021`) | Có hiệu lực từ 15/10/2021; sửa đổi bởi TT 01/2025 và (theo một nguồn) TT 60/2025 | Chỉ có bản OCR Phụ lục VI do người dùng cung cấp; ngưỡng ở trạng thái `verified=False` | Kỹ sư QS, đối chiếu bản gốc |
| **TT 37/2026/TT-BXD** | Ghi trong cảnh báo của `takeoff_rules`: "từ 01/07/2026 cần kiểm tra văn bản thay thế" | Ban hành 26/06/2026, có hiệu lực 01/07/2026 (nhiều nguồn thứ cấp thống nhất) | Nguồn thứ cấp cho thấy văn bản thay thế một số thông tư cũ, nhưng **chưa xác nhận** có bãi bỏ phần đo bóc (Phụ lục VI) hay không | Người có văn bản gốc; cần xác định phạm vi thay thế |
| **TT 38/2026/TT-BXD** | Định mức xây dựng (nhiều nơi trong code) | Có hiệu lực 01/07/2026 (nguồn thứ cấp) | Chưa đọc nội dung định mức | Kỹ sư định mức |
| **TT 36/2026/TT-BXD** | Công thức G_XD xây lắp (`tools/qs_loader.py`) | Chưa xác minh | Chưa đọc nội dung | Kỹ sư dự toán |
| **TT 12/2021/TT-BXD** | Tham chiếu hướng dẫn đo bóc trong `takeoff_rules` | Chưa xác minh | Chưa đọc nội dung | Kỹ sư QS |
| **NĐ 207/2026/NĐ-CP** | Quy trình KCS và nghiệm thu (sổ KCS, biên bản) | Chưa xác minh | Chưa đọc nội dung | Chủ đầu tư / QLCL |
| **NĐ 254/2025/NĐ-CP** | Thanh toán Mẫu 03a | Chưa xác minh | Chưa đọc nội dung | Kế toán / kiểm soát chi |
| **NĐ 99/2021/NĐ-CP** | Khấu trừ thanh toán (đối chiếu với NĐ 254/2025) | Chưa xác minh | Chưa đọc nội dung | Kế toán |
| **TT 01/2025, TT 60/2025, TT 32/2026, TT 17/2019** | Được nhắc như lịch sử sửa đổi hoặc phương pháp cũ | Chưa xác minh | Chưa đọc nội dung | Kỹ sư QS |
| **QĐ 1041/QĐ-BXD** | Nhắc trong cảnh báo đo bóc (nguồn thứ cấp) | Chưa xác minh | Chưa xác nhận tồn tại | Người có văn bản gốc |
| **TCVN 1651:2018** | Thép cốt bê tông, cắt thép | Tiêu chuẩn | Chưa đối chiếu bản gốc | Kỹ sư kết cấu |
| **TCVN 5574:2018** | Kết cấu bê tông cốt thép | Tiêu chuẩn | Chưa đối chiếu bản gốc | Kỹ sư kết cấu |
| **TCVN 11823:2017** | Thiết kế cầu (engine cầu) | Tiêu chuẩn | Chưa đối chiếu bản gốc | Kỹ sư cầu |
| **TCVN 4453:1995** | Thi công và nghiệm thu bê tông toàn khối (biên bản KCS) | Tiêu chuẩn; chưa kiểm tra tình trạng hiện hành | Chưa đối chiếu bản gốc | Kỹ sư kết cấu |

## Hệ quả đối với kết quả hệ thống

- Mọi bảng đo bóc dùng quy tắc `tt13-2021` phải mang cảnh báo "CHƯA đối chiếu bản gốc" (đang làm vậy).
- Sổ KCS chỉ kết luận "đủ điều kiện" với những bản ghi đã đạt kiểm tra logic chéo; còn lại phải ghi rõ chưa đủ điều kiện hoặc chưa kiểm tự động. Không dùng nhãn "Hợp lệ" gõ tay.
- Dự toán G_XD dùng tỷ lệ đọc từ file mẫu; các tỷ lệ này phải được đối chiếu với văn bản đang hiệu lực trước khi dùng cho hồ sơ thật.

## Việc cần làm để nâng trạng thái

1. Đọc nội dung chính thức của TT 37/2026 và TT 38/2026 (đặc biệt điều khoản thay thế và hiệu lực).
2. Xác định Phụ lục VI của TT 13/2021 có còn hiệu lực sau 01/07/2026 không.
3. Đọc NĐ 207/2026 và NĐ 254/2025 cho đúng điều khoản đang dùng.
4. Cập nhật cột "Mức kiểm chứng" thành *ĐÃ ĐỐI CHIẾU BẢN GỐC* kèm số điều/khoản và ngày.
