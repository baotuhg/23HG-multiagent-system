# QUY TRÌNH 04: BẢNG THANH TOÁN KHỐI LƯỢNG HOÀN THÀNH KỲ (PHỤ LỤC 03A)
## CĂN CỨ PHÁP LÝ: NGHỊ ĐỊNH SỐ 254/2025/NĐ-CP CỦA CHÍNH PHỦ VỀ QUẢN LÝ, THANH TOÁN VỐN ĐẦU TƯ CÔNG

---

### I. NGUYÊN TẮC CỐT LÕI
1. **Biểu mẫu bắt buộc:** Phụ lục 03a là biểu mẫu pháp lý cao nhất để Kho bạc Nhà nước, Chủ đầu tư và Ban QLDA giải ngân tiền cho Nhà thầu thi công xây dựng.
2. **Kiểm soát lũy kế:** Khối lượng lũy kế đến hết kỳ này không bao giờ được vượt quá Khối lượng Hợp đồng được duyệt (trừ trường hợp có Phụ lục hợp đồng bổ sung khối lượng phát sinh).

### II. CẤU TRÚC 12 CỘT CHUẨN CỦA PHỤ LỤC 03A
- Cột 1: STT
- Cột 2: Nội dung công việc theo hợp đồng
- Cột 3: Đơn vị tính
- Cột 4: Khối lượng theo Hợp đồng
- Cột 5: Đơn giá theo Hợp đồng
- Cột 6: Thành tiền theo Hợp đồng:
  ```excel
  =D{row} * E{row}
  ```
- Cột 7: Khối lượng thực hiện lũy kế đến hết kỳ trước
- Cột 8: Khối lượng thực hiện nghiệm thu kỳ này
- Cột 9: Khối lượng thực hiện lũy kế đến hết kỳ này:
  ```excel
  =G{row} + H{row}
  ```
- Cột 10: Giá trị thực hiện lũy kế đến hết kỳ này:
  ```excel
  =I{row} * E{row}
  ```
- Cột 11: Giá trị đề nghị thanh toán kỳ này:
  ```excel
  =H{row} * E{row}
  ```
- Cột 12: Ghi chú tình trạng nghiệm thu

### III. BẢNG TỔNG HỢP DÒNG TIỀN THANH TOÁN VỐN ĐẦU TƯ CÔNG
1. **Giá trị khối lượng công việc hoàn thành đề nghị thanh toán kỳ này ($A$):**
   ```excel
   =SUM(Cột_11_từ_đầu_đến_cuối)
   ```
2. **Thu hồi số tiền tạm ứng hợp đồng ($B$):**
   - Theo tỷ lệ thỏa thuận trong hợp đồng (thường khấu trừ $20\% - 25\%$ giá trị thanh toán mỗi kỳ cho đến khi thu hồi hết số tiền tạm ứng ban đầu):
     ```excel
     =A * 20%
     ```
3. **Số tiền giữ lại bảo hành công trình ($C$):**
   - Giữ lại theo quy định pháp luật (thường là $5\%$ giá trị thanh toán):
     ```excel
     =A * 5%
     ```
4. **Số tiền thực đề nghị Kho bạc Nhà nước chuyển kỳ này ($D$):**
   ```excel
   =A - B - C
   ```
