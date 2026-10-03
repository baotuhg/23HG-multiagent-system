# Đối chiếu gói Hub & Spoke: bản đầy đủ của người dùng vs bản hệ thống xuất (03/10/2026)

Nguồn "đầy đủ": `Downloads/Documents/HSTK Cầu Km19+529.080_Marker/03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH`
Bản hệ thống: `examples/HO_SO_CAU_KM19_529/HUB/03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH`
(Master gốc dùng để xuất = Master của người dùng, chỉ thay sheet `TIEN_DO_THI_CONG_WBS` bằng bản công thức sống.)

| Gói | Tệp | Đầy đủ (KB) | Hệ thống (KB) | Trạng thái |
|---|---|---|---|---|
| A | TDTC ca xe/ca máy/dầu diesel `.xlsx` | 60 | 17 | THIẾU chi tiết — dữ liệu ca máy không nằm trong Master |
| A | Tiến độ ca máy `.xml` | 9 | 17 (xml tiến độ cầu) | Khác nội dung |
| B | Phiếu cắt thép CNC `.csv` | 9.293 | 0 | THIẾU — cần dữ liệu cắt thép từ BBS thật |
| B | Tổ hợp cắt thép 11m7 `.xlsx` | 4.248 | 8 | THIẾU — như trên |
| B | Thống kê thép BBS | 48 | 46 | Tương đương |
| C | 5 tệp KCS + `.docx` | 63/29/11/14/14/43 | 63/18/11/10/8/43 | Đủ, một số tệp mẫu A4 gọn hơn |
| D | 5 tệp QS/dự toán/thanh toán | ~200 | ~136 | Đủ tệp; dự toán GXD gắn TT 36/2026 (bản đầy đủ: TT 11/2021) |
| E | Master, XML, MPP, audit, BPTC | có | có | Đủ tệp |
| — | Bảng phân quyền & bàn giao `.md` | 7 | 2 | Gọn hơn nhiều |

## Đã sửa
- Lỗi lẫn hồ sơ dự án khác (`Khai_Hoang_2`): `run_state_graph.py` nhặt tệp XML/MPP/DOCX đầu tiên trong `templates/`. Nay chỉ nhận tệp khớp tên dự án.
- Lỗi ghi file vi mô với ngày có múi giờ (`tools/package_dispatcher.py`).

## Còn tồn tại (chưa sửa)
1. Gói B (CSV 0 KB, RebarCut 8 KB) và Gói A (ca máy) cần dữ liệu đầu vào thật (bảng cắt thép, danh sách thiết bị). Bộ xuất chỉ cắt các sheet có sẵn trong Master nên không tự sinh ra chúng.
2. `.xml`/`.mpp` trong Gói E là bản cũ (ngày 10/2026–03/2027), KHÔNG khớp tiến độ công thức sống mới (15/03–14/09/2026).
3. Bảng phân quyền/bàn giao chỉ là bản tóm tắt.
