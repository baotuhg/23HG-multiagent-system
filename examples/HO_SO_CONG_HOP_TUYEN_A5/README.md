# Hồ sơ mẫu Cống hộp Tuyến A5 — trạng thái thật

Đây là **hồ sơ mẫu minh họa**, chưa phải hồ sơ đã được kỹ sư QS thẩm định. Số liệu đầu vào (đơn giá, khối lượng hợp đồng,
chi phí trực tiếp T) là số nhập, không liên kết với một dự toán chi tiết.

## Nội dung có dữ liệu

| Tầng | Nội dung |
|---|---|
| 1 — Master | `BO_HO_SO_01_.../01_Ho_So_KCS_QS_TienDo_Master_14_Sheets_Cong_Hop_A5.xlsx`: 8 sheet có dữ liệu, 6 sheet đánh dấu **CHƯA LẬP** (cấp phối & tần suất thí nghiệm, phân tích vật tư, tổng hợp vật tư, 3 mẫu biên bản A4) |
| 2 — Vi mô | 8 file: 01 cắt thép, 02 đào đắp, 03 diễn giải bóc tách, 04 thống kê thép, 08 dự toán G_XD, 09 thanh toán 03a, 10 tiến độ CPM, 11 danh mục KCS |
| 3 — Hub & Spoke | Gói A–E. Gói D chứa 03, 08, 09 (bản sao giống hệt Tầng 2). **Gói A, B, C không chứa sheet đơn giá** |

Sáu hồ sơ vi mô không có dữ liệu (05, 06, 07, 12, 13, 14) **không được xuất**; danh sách nằm trong `DISPATCH_MANIFEST.json`
(`not_generated`). Bản sao giữa Tầng 2 và Gói B/C/D là chủ ý (mỗi gói phải độc lập); `tests/test_examples_consistency.py`
kiểm tra các bản sao không bị lệch nhau.

## Việc đã nối và việc còn mở

- **Đắp lưng cống** (`KHOI_LUONG_DAO_DAP`, dòng 14-20): tính lại theo từng loại cống, trừ phần cống bao ngoài chiếm chỗ.
  Tổng 11.081,9 m³ so với số cũ 20.804,8 m³ (−46,7%). Khoảng thao tác 0,75 m mỗi bên là **giả định suy từ số cũ**, cần QS xác nhận.
- **Mẫu 03a** (cột J-L): khối lượng hợp đồng được đối chiếu với bóc tách. Đào, cọc tre, BT lót, BT thân, ván khuôn khớp
  trong 0,5%. **Cốt thép: hợp đồng 753,4 tấn, thống kê thép 861,7 tấn (+14,4%)** — chưa giải thích được.
- **Ván khuôn** trong bảng diễn giải: số đang dùng thấp hơn 16-22% so với tính lại theo hình học (xem cột P-U của sheet diễn giải).
- **G_XD**: chi phí trực tiếp T = 42,5 tỷ đồng là số nhập; chưa liên kết khối lượng × đơn giá.
- Mẫu 03a ghi căn cứ NĐ 99/2021 và khấu trừ tạm ứng/giữ lại trước VAT; code `tools/payment.py` dùng NĐ 254/2025 và khấu trừ sau VAT.

## Tạo lại

```bash
python examples/clean_a5_dossier.py        # chạy lặp được; áp dụng mọi chỉnh sửa ở trên
python -m tools.audit_excels_static examples
python -m unittest tests.test_examples_consistency
```
