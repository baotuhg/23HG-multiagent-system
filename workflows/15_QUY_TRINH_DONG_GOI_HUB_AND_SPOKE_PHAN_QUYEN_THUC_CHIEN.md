# QUY TRÌNH 15: ĐÓNG GÓI HỒ SƠ THỰC CHIẾN MÔ HÌNH "HUB & SPOKE" PHÂN QUYỀN
## GIẢI PHÁP TỐI ƯU HÓA QUẢN TRỊ DỮ LIỆU & ĐIỀU HÀNH HIỆN TRƯỜNG AEC MULTI-AGENT
### Tuân thủ: Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Vincons & Vinhomes Standards

---

## 1. NGUYÊN LÝ THIẾT KẾ: TẠI SAO PHẢI DÙNG MÔ HÌNH "HUB & SPOKE"?

File Monolithic 14 Sheet truyền thống chỉ phù hợp cho **Thẩm tra & Lưu trữ pháp lý** (`mode="audit_archive"`). Khi ra công trường thực tế, nếu tất cả các bộ phận dùng chung 1 file sẽ gây ra:
1. **Xung đột file khóa (Read-Only Lock):** Kỹ sư QS đang làm thanh toán thì Đội trưởng cơ giới không thể cập nhật tích kê dầu máy.
2. **Lộ lọt thông tin tài chính:** Thợ gia công sắt, đội cơ giới thấy toàn bộ đơn giá dự toán, lợi nhuận thầu và chi phí gián tiếp của Tổng thầu.
3. **Nặng và chậm trên thiết bị di động:** Mở file 14 Sheet ngoài hố móng rất dễ lag và làm gãy công thức `#REF!`.

Hệ thống **23HG-multiagent-system** chuẩn hóa quy trình đóng gói thành **5 Gói Vệ tinh Phân quyền (Hub & Spoke Role-Based Model)**:

```
                            +-------------------------------------+
                            |        CENTRAL STATE BUS / DB       |
                            |   (SQLite / JSON / DuckDB Backend)  |
                            |       "Single Source of Truth"      |
                            +------------------+------------------+
                                               |
             +---------------------------------+---------------------------------+
             |                                 |                                 |
             v                                 v                                 v
+------------------------+        +------------------------+        +------------------------+
| GÓI A: CƠ GIỚI & DẦU   |        | GÓI B: XƯỞNG CỐT THÉP  |        | GÓI C: HIỆN TRƯỜNG KCS |
| • Quản lý máy móc      |        | • Quản đốc xưởng       |        | • Cán bộ QA/QC, TVGS   |
| • Cấp phát Dầu Diezel  |        | • Lệnh cắt CNC (.csv)  |        | • BBNT Word A4 (.docx) |
| • Biểu đồ phụ tải máy  |        | • Sơ đồ 1D CSP 11.7m   |        | • Ma trận nén R7/R28   |
+------------------------+        +------------------------+        +------------------------+
             |                                 |                                 |
             +---------------------------------+---------------------------------+
                                               |
                                               v
                            +-------------------------------------+
                            | GÓI D: QS, DỰ TOÁN & THANH TOÁN     |
                            | • Kỹ sư QS, Kế toán dự án           |
                            | • Bảo mật đơn giá dự toán G_XD      |
                            | • Bảng thanh toán Phụ lục 03a       |
                            +------------------+------------------+
                                               |
                                               v
                            +-------------------------------------+
                            | GÓI E: EXECUTIVE DASHBOARD (HUB)    |
                            | • Giám đốc Dự án, Ban QLDA, CĐT     |
                            | • Tiến độ CPM, Báo cáo Audit 100/100|
                            +-------------------------------------+
```

---

## 2. BẢNG PHÂN QUYỀN TRUY CẬP HỒ SƠ CÔNG TRƯỜNG

| Gói hồ sơ | Đối tượng sử dụng | Quyền hạn dữ liệu | Tệp bàn giao chính |
| :--- | :--- | :--- | :--- |
| **Gói A: Cơ giới & Dầu** | Đội trưởng xe máy, Thủ kho dầu, Lái xe máy | Xem tiến độ ca máy, phụ tải, ký nhận dầu. **Không xem đơn giá tiền.** | `260920_TDTC_CaXe_CaMay_...xlsx`<br>`260920_Tien_Do_CaMay_...xml` |
| **Gói B: Xưởng cốt thép** | Quản đốc xưởng, Thợ uốn cắt, Nhà cấp thép | Xem BBS, sơ đồ cắt thép, file CSV nạp máy CNC. **Không xem đơn giá tiền.** | `01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`<br>`01_Phieu_Cat_Thep_Xuong_CNC_...csv`<br>`04_Thong_Ke_Thep_Chi_Tiet_BBS_...xlsx` |
| **Gói C: Hiện trường KCS** | Kỹ sư QA/QC, Tư vấn giám sát, Thí nghiệm | Xem ngày nghiệm thu, nén mẫu, in ấn biên bản. **Không xem đơn giá tiền.** | `11_Danh_Muc_KCS_Bien_Ban_Nghiem_Thu.xlsx`<br>`Ho_So_Bien_Ban_Nghiem_Thu_KCS_...docx`<br>`Mau_A4_Bien_Ban_...xlsx` |
| **Gói D: QS & Dự toán** | Kỹ sư QS, Trưởng phòng Kế hoạch, Kế toán | **Toàn quyền xem đơn giá, doanh thu, thanh toán.** | `03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx`<br>`08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx`<br>`09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx` |
| **Gói E: Executive Hub** | Giám đốc Dự án, Ban Giám đốc, Chủ đầu tư | Xem KPI tổng thể, tiến độ đường găng CPM, báo cáo thẩm tra Audit 100/100. | `02_Tien_Do_Thi_Cong_Master_...xml`<br>`03_BAO_CAO_THAM_TRA_AEC_AUDIT_...md` |

---

## 3. CÁCH THỨC TRIỂN KHAI BẰNG CODE (PYTHON API)

```python
from tools.package_dispatcher import AECPackageDispatcher

dispatcher = AECPackageDispatcher(base_output_dir=r"c:\Users\baotu\Downloads\DU_AN_CONG_TRINH")

manifest = dispatcher.dispatch_site_operation_packages(
    project_name="Cong_Hop_Tuyen_A5",
    artifacts_source_dir=r"c:\Users\baotu\Downloads\CỐNG HỘP TUYẾN A5",
    custom_subfolder="HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5"
)

print(manifest.summary_report)
```
