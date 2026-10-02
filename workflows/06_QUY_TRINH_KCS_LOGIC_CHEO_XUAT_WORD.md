# QUY TRÌNH 06: QUẢN LÝ CHẤT LƯỢNG KCS, LOGIC NGÀY CHÉO & BIỂU MẪU EXCEL A4
## CĂN CỨ PHÁP LÝ: NGHỊ ĐỊNH SỐ 207/2026/NĐ-CP & LUẬT XÂY DỰNG SỐ 135/2025/QH15

---


> **Trạng thái:** ✅ **Đã hiện thực** phần kiểm tra logic ngày chéo và phiếu thí nghiệm / điểm
> dừng kỹ thuật: `tools/lab_qaqc.py`, `BPTCKCSAgent`.
> 🟡 Xuất biểu mẫu nghiệm thu ra Word (`.docx`) hiện chỉ có trong các script `examples/`, chưa tích
> hợp vào pipeline chính.

### I. NGUYÊN TẮC CỐT LÕI
1. **Kiểm soát tính hợp pháp của hồ sơ KCS:** Biên bản nghiệm thu là căn cứ pháp lý để thanh quyết toán và bàn giao công trình. Bất kỳ sự mâu thuẫn nào về ngày tháng giữa Biên bản nghiệm thu, Phiếu kết quả thí nghiệm và Nhật ký thi công đều là lỗi nghiêm trọng bị xuất toán khi kiểm toán.
2. **Nguyên tắc "Không được đá ngày":**
   - Nghiệm thu công tác sau bắt buộc phải diễn ra sau khi công tác trước đã được nghiệm thu đạt yêu cầu.
   - Các công tác bê tông bắt buộc phải có kết quả thử độ sụt cùng ngày và kết quả nén mẫu R7 / R28 hợp chuẩn trước khi cho phép chịu lực hoặc chuyển bước thi công kết cấu bên trên.
3. **Nguyên tắc Đồng bộ hóa Biểu mẫu Excel A4 (Không dùng Word):**
   - Theo chuẩn hóa vận hành AEC hiện đại, toàn bộ biểu mẫu KCS được tích hợp trực tiếp dưới dạng các Sheet Excel chuẩn in ấn A4 (Portrait) ngay trong cùng Workbook Master.
   - Cơ chế chọn số biên bản / mã vật tư bằng danh sách thả xuống (`Dropdown/Input Cell`), toàn bộ thông tin đối tượng nghiệm thu, khối lượng, tiêu chuẩn, ngày tháng được điền tự động bằng hàm `=VLOOKUP` không có số chết.

### II. MA TRẬN KIỂM SOÁT LOGIC CHÉO NGÀY THÁNG (DATE CROSS-CHECK MATRIX)
1. **Chuỗi thi công Cọc khoan nhồi:**
   - Ngày hạ lồng thép cọc nhồi $\rightarrow$ Ngày đổ bê tông cọc nhồi (cùng ngày hoặc hôm sau).
   - Ngày đổ bê tông cọc $\rightarrow$ Tối thiểu sau $7 - 14$ ngày mới làm **Siêu âm cọc** và **Thử tải động PDA**.
   - Có kết quả siêu âm và PDA đạt yêu cầu mới được nghiệm thu công tác **Đập đầu cọc nhồi** và **Đổ bê tông lót đáy bệ**.
2. **Chuỗi thi công Bệ và Thân mố trụ:**
   - Ngày đổ bê tông bệ $\rightarrow$ Tối thiểu sau 7 ngày (đạt R7) mới cho phép lắp dựng cốt thép và ván khuôn thân mố/thân trụ.
   - Thân đặc trụ cao ($11.4\text{m} - 14.5\text{m}$) đổ bê tông xong $\rightarrow$ Đạt R14 mới cho phép thi công xà mũ trụ và đá kê gối.
3. **Chuỗi thi công Dầm Super-T:**
   - Đúc dầm tại bãi $\rightarrow$ Đạt $100\%$ cường độ R28 (hoặc tối thiểu $85\%$ nếu bảo dưỡng hơi nước) mới được **Căng kéo cáp DUL $\Phi 15.2\text{mm}$**.
   - Căng cáp xong $\rightarrow$ Bơm vữa ống gen kín khít $\rightarrow$ Sau 24-48h vữa đông cứng mới được vận chuyển và cẩu lao dầm.
   - Lao dầm vào nhịp $\rightarrow$ Yên vị trên gối chậu $\rightarrow$ Mới thi công dầm ngang, bản liên tục nhiệt và bản mặt cầu C35.
4. **Chuỗi hoàn thiện mặt cầu:**
   - Bản mặt cầu đổ bê tông xong $\rightarrow$ Đạt R14 $\rightarrow$ Mới phun màng chống thấm và thảm bê tông nhựa chặt C16.
   - Toàn bộ kết cấu hoàn thành $\rightarrow$ Bê tông đạt R28 $\rightarrow$ Mới thực hiện **Thử tải tĩnh và động toàn cầu**.

### III. HỆ THỐNG 3 BIỂU MẪU EXCEL CHUẨN IN ẤN A4 (THAY THẾ TOÀN BỘ FILE WORD)
1. **Sheet `MAU_BIEN_BAN_KCS` (Biên bản nghiệm thu công việc xây dựng):**
   - Căn cứ pháp lý: Nghị định 207/2026/NĐ-CP, Luật Xây dựng 135/2025/QH15.
   - Ô điều khiển `C2`: Nhập số thứ tự biên bản từ `1` đến `22`.
   - Toàn bộ nội dung: Tên công việc, khối lượng nghiệm thu, đơn vị tính, căn cứ kỹ thuật, ngày giờ nghiệm thu được tra cứu tự động qua `=VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, col, FALSE)`.
   - Tự động nhảy khung chữ ký của Tư vấn Giám sát trưởng và Chỉ huy trưởng công trường.

2. **Sheet `MAU_BB_NGHIEM_THU_VAT_LIEU` (Biên bản nghiệm thu vật liệu xây dựng đầu vào):**
   - Ô điều khiển `C2`: Nhập mã số vật liệu từ `1` đến `16` (thép các loại $\varnothing$, xi măng, cát, đá, nước, phụ gia).
   - Tự động nhảy tên vật tư, khối lượng nghiệm thu của lô, tiêu chuẩn kỹ thuật kiểm tra (TCVN 1651:2018, TCVN 6260:2020,...), quy cách tổ mẫu và điều kiện chuyển bước thi công từ bảng ma trận tần suất KCS.

3. **Sheet `MAU_BB_LAY_MAU_HIEN_TRUONG` (Biên bản lấy mẫu thí nghiệm tại hiện trường):**
   - Ô điều khiển `C2`: Nhập mã hạng mục bê tông từ `17` đến `26` (bê tông cọc, mố, trụ, xà mũ, dầm Super-T, bản mặt cầu,...).
   - Ô điều khiển `C3`: Nhập ngày đúc mẫu / đổ bê tông thực tế (dạng `yyyy-mm-dd`).
   - Tự động tính toán ngày nén mẫu hợp chuẩn:
     $$\text{Ngày nén } R_7 = C_3 + 7$$
     $$\text{Ngày nén } R_{28} = C_3 + 28$$
   - Định dạng chuẩn A4 Portrait, căn lề in ấn tiêu chuẩn, có thể kết xuất trực tiếp ra PDF hoặc in ngay trên công trường.
