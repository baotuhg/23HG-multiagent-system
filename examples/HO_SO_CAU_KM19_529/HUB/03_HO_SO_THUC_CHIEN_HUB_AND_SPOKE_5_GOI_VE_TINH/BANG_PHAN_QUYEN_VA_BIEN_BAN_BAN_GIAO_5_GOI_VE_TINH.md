# BẢNG PHÂN QUYỀN & BIÊN BẢN BÀN GIAO 5 GÓI VỆ TINH
### DỰ ÁN: CAU KM19+529.080
**Hệ thống điều phối:** 23HG-AEC-MultiAgent-System (Quy trình 15 - Mô hình Hub & Spoke)  
**Tiêu chuẩn bảo mật:** RBAC Data Partitioning, Luật Xây dựng 135/2025/QH15 & NĐ 207/2026/NĐ-CP  

---

### I. NGUYÊN TẮC PHÂN QUYỀN CÔNG TRƯỜNG
1. **Chống xung đột khóa file (Zero File Locks):** Mỗi bộ phận vận hành độc lập trên gói chuyên trách, không tranh chấp tệp.
2. **Bảo mật giá thầu & chi phí tài chính (Cost Isolation):** Đội xe máy, thợ sắt và TVGS không thể nhìn thấy đơn giá dự toán và lợi nhuận nhà thầu.
3. **Mượt mà trên thiết bị di động:** Dung lượng tệp nhẹ, mở tức thì trên điện thoại và máy tính bảng ngoài hiện trường.

### II. BẢNG MA TRẬN PHÂN QUYỀN TRUY CẬP (RACI MATRIX)
| Gói Vệ Tinh | Phân hệ chức năng | Bộ phận tiếp nhận | Mức bảo mật | Dữ liệu được xem | Dữ liệu cấm tuyệt đối |
| :--- | :--- | :--- | :---: | :--- | :--- |
| **Gói A** | Cơ giới & Dầu | Đội xe máy, Thủ kho dầu | Nội bộ | Tiến độ ca máy, phụ tải, tiêu thụ dầu | CẤM XEM Đơn giá, Dự toán, Lợi nhuận |
| **Gói B** | Xưởng cốt thép | Quản đốc, Thợ cắt, ĐV cấp thép | Nội bộ | BBS, sơ đồ cắt 11.7m, CSV CNC | CẤM XEM Đơn giá mua, Giá trị hợp đồng |
| **Gói C** | Hiện trường KCS | Kỹ sư QA/QC, TVGS, Thí nghiệm | Quan trọng | Danh mục KCS, 22 BBNT Word, mẫu R7/R28 | CẤM XEM Dự toán chi tiết, Doanh thu 03a |
| **Gói D** | QS & Dự toán | Kỹ sư QS, Ban Kế hoạch, Kế toán | Mật cao | Hình học Takeoff, Dự toán G_XD, Phụ lục 03a | Chỉ lưu hành nội bộ phòng Kế hoạch |
| **Gói E** | Executive Hub | Giám đốc DA, Ban QLDA, Chủ đầu tư | Tối cao | KPI, Đường găng CPM, Audit Score 100/100 | Toàn quyền kiểm soát hệ thống |
