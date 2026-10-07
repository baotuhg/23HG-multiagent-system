# QUY TRÌNH QUẢN LÝ VÀ VẬN HÀNH BÃI GIA CÔNG CỐT THÉP (REBAR WORKSHOP)
## DỰ ÁN: CAO TỐC TUYÊN QUANG - HÀ GIANG (GIAI ĐOẠN 1) - CẦU KM19+529.080

Hệ thống tổ hợp cắt thép được lập bằng công nghệ **Pure Python / Zero-LLM** kết hợp giải thuật tối ưu hóa toán học **OR-Tools Gilmore-Gomory CP-SAT**, lấy định dạng **RebarCut Pro 5 sheet** làm kim chỉ nam.

---

### 1. Cấu trúc thư mục hệ thống
- `00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx`: Bảng điều khiển trung tâm (Dashboard) so sánh chỉ tiêu kinh tế kỹ thuật của 11 loại đường kính Ø và Cáp DƯL.
- `00_RebarCut_MASTER_TOAN_CAU_11M7.xlsx`: File tổng hợp toàn bộ 33.212 cây thép 11.7m của toàn bộ cầu (Đầy đủ 5 sheet chuẩn: INPUT, SO_SANH, PA_TOI_UU, REMAIN, CHI_TIET).
- `THEO_TUNG_DUONG_KINH_PHI/`: Thư mục chứa 11 tập hồ sơ cắt thép chuyên sâu riêng biệt cho từng đường kính từ nhỏ nhất (Ø8) đến lớn nhất (Ø32).
- `LENH_CAT_CNC_CSV/`: Các file CSV nạp trực tiếp vào máy cắt tự động CNC hoặc bảng điều khiển của thợ máy tại từng trạm.

---

### 2. Phân luồng 5 Trạm máy tại Bãi tiền chế
1. **Trạm 1 (Ø8, Ø10)**: Máy uốn đai tự động CNC uốn liên tục từ thép cuộn/cây 11.7m (12.278 thanh Ø10 đai xoắn cọc và 780 thanh Ø8 đai tăng cường).
2. **Trạm 2 (Ø12, Ø14, Ø16)**: Máy cắt đa thanh Shearline (76.014 thanh Ø12 và 35.366 thanh Ø16 bản mặt cầu và sườn dầm).
3. **Trạm 3 (Ø18, Ø20, Ø22)**: Trạm cắt uốn định hình móng, thân mố trụ và bản mặt cầu (9.337 thanh).
4. **Trạm 4 (Ø25, Ø28, Ø32)**: Trạm máy cắt công suất lớn kết hợp tiện ren dập đầu coupler nối cơ khí cho cọc khoan nhồi D1200 và cốt chủ móng (5.940 thanh).
5. **Trạm 5 (Ø15.2)**: Trạm kéo rải và căng kéo tao cáp DƯL 7 sợi ASTM A416 Gr270 dầm Super-T (660 thanh L=38.2m).

---

### 3. Quy tắc kiểm soát phôi thừa (Offcuts / Đề-xê)
- Mọi đầu thừa có chiều dài >= 100xD được gắn mã lưu kho tại sheet `REMAIN` để tái sử dụng làm con kê hoặc cấu kiện ngắn.
- Đầu thừa < 20xD được gom vào hộc phế liệu phân loại theo từng mác thép để thanh lý phế liệu có kiểm soát.
