---
name: aec-qlcl
description: Hỗ trợ lập, rà soát, bóc tách, chuẩn hóa và tự động hóa hồ sơ quản lý chất lượng (QLCL) công trình xây dựng tại Việt Nam (Giao thông - Cầu đường, San nền, Dân dụng & Hạ tầng kỹ thuật) theo Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP và Thông tư 32/2026/TT-BXD. Dùng khi cần lập danh mục nghiệm thu, kiểm tra logic chéo ngày tháng, đối chiếu kết quả thí nghiệm với biên bản, rà soát nhật ký thi công và bản vẽ hoàn công.
---

# Quản lý Chất lượng & Nghiệm thu Công trình Xây dựng (AEC-QLCL)

## 1. Mục tiêu
Thiết lập quy trình chuẩn để lập, soát xét và tự động hóa hồ sơ QLCL cho các dự án xây dựng tại Việt Nam, bao gồm:
- **Công trình Giao thông & Cầu đường:** Nền đường, áo đường, công trình thoát nước (hệ thống cống bản, cống tròn theo lý trình Km), kết cấu cầu, mố trụ, gia cố ta luy, kè đá.
- **Công trình Dân dụng & Hạ tầng kỹ thuật:** San nền, tường chắn đất, kè điểm trường, cổng hàng rào, cột cờ, nhà lớp học và hạ tầng đồng bộ.

## 2. Trục pháp lý và Tiêu chuẩn cốt lõi

1. **Trục pháp lý hiện hành:**
   - **Luật Xây dựng số 135/2025/QH15**.
   - **Nghị định số 207/2026/NĐ-CP** về quản lý chất lượng, thi công xây dựng và bảo trì công trình xây dựng.
   - **Thông tư số 32/2026/TT-BXD** quy định chi tiết về quản lý chất lượng và phân loại công trình.
2. **Quy chuẩn kỹ thuật quốc gia (Bắt buộc áp dụng):**
   - **QCVN 06:2022/BXD** (và sửa đổi) về an toàn cháy cho nhà và công trình.
   - **QCVN 07:2023/BXD** về các công trình hạ tầng kỹ thuật.
   - **QCVN 03:2022/BXD** về phân cấp công trình xây dựng.
3. **Phân loại 4 nhóm căn cứ khi lập hồ sơ:**
   - `[PHÁP LUẬT/QCVN BẮT BUỘC]`: Luật, Nghị định, Thông tư, Quy chuẩn quốc gia.
   - `[TIÊU CHUẨN ÁP DỤNG CỦA DỰ ÁN]`: TCVN (ví dụ TCVN 4453:1995 kết cấu bê tông, TCVN 9340:2012, TCVN 8859:2011 móng cấp phối đá dăm...) được phê duyệt trong Chỉ dẫn kỹ thuật.
   - `[HỒ SƠ THIẾT KẾ ĐƯỢC DUYỆT]`: Bản vẽ thi công, biện pháp thi công (BPTC), kế hoạch kiểm tra thí nghiệm (ITP).
   - `[THÔNG LỆ QA/QC NỘI BỘ]`: Mẫu biểu nội bộ, quy trình quản lý chất lượng doanh nghiệp.

## 3. Quy trình 7 bước lập và soát xét hồ sơ QLCL

### Bước 1: Xác định phạm vi và Cây phân rã công việc (WBS)
- Thu thập dữ liệu: Gói thầu, lý trình (Km... đến Km...), hạng mục (Đào nền, Đắp K95/K98, Cống tròn D100, Móng mố, Xà mũ, Mặt đường CPĐD, Kè đá hộc, San nền...).
- Xác định cấp công trình và các mốc nghiệm thu giai đoạn/bộ phận.

### Bước 2: Xác định căn cứ kiểm tra và Chỉ dẫn kỹ thuật
- Rà soát hồ sơ thiết kế được duyệt để lấy: Mác bê tông, nhóm thép, độ chặt yêu cầu $K$, cường độ nén $R_{28}$, dung sai kích thước hình học.
- Tuyệt đối không tự suy đoán số liệu kỹ thuật nếu chưa có trong hồ sơ thiết kế.

### Bước 3: Lập danh mục và Mã định danh văn bản (Doc ID Matrix)
- Đặt mã ID duy nhất cho từng tài liệu:
  - `VL-xxx`: Biên bản nghiệm thu vật liệu đầu vào (Cát, đá, xi măng, thép, ống cống, đá hộc...).
  - `TN-xxx`: Phiếu kết quả thí nghiệm (Nén mẫu bê tông R7/R28, độ chặt K, kéo cốt thép...).
  - `NT-xxx`: Biên bản nghiệm thu công việc xây dựng.
  - `GĐ-xxx`: Biên bản nghiệm thu giai đoạn/bộ phận công trình.
  - `NK-xxx`: Nhật ký thi công theo ngày.
  - `HC-xxx`: Bản vẽ hoàn công.

### Bước 4: Kiểm soát Chuỗi vật liệu đầu vào
Mọi công việc sử dụng vật tư phải truy xuất được nguồn gốc:
1. Chứng chỉ xuất xưởng, CO/CQ của nhà sản xuất.
2. Hóa đơn/phiếu xuất kho đến công trường.
3. Biên bản lấy mẫu thí nghiệm tại hiện trường.
4. Phiếu kết quả thí nghiệm kiểm tra độc lập (LAS-XD).
5. Biên bản chấp thuận nghiệm thu vật liệu trước khi đưa vào thi công.

### Bước 5: Kiểm tra Nghiệm thu công việc & Chuyển bước thi công
- Đối với công việc bị che khuất (đào hố móng, đắp đất, cốt thép móng, cống chìm): Phải đo đạc hình học, nghiệm thu hoàn công và lập biên bản trước khi cho phép đổ bê tông hoặc lấp đất.
- Kiểm tra chữ ký: Đầy đủ đại diện TVGS (Kỹ sư giám sát trưởng/giám sát viên) và Chỉ huy trưởng/Cán bộ kỹ thuật nhà thầu.

### Bước 6: Nhật ký thi công và Bản vẽ hoàn công
- **Nhật ký thi công:** Ghi nhận điều kiện thời tiết, nhân lực, thiết bị, vị trí thi công cụ thể và các biên bản nghiệm thu được ký trong ngày theo mẫu Phụ lục Nghị định 207/2026/NĐ-CP.
- **Bản vẽ hoàn công:** Kiểm tra đóng dấu hoàn công, đối chiếu kích thước thực tế so với bản vẽ thiết kế, người lập và người giám sát ký xác nhận.

### Bước 7: Audit logic ngày tháng bằng Tool tự động
Trước khi in ấn hoặc trình nộp hồ sơ, bắt buộc chạy script kiểm tra chéo:

```bash
python .openspace/skills/aec-qlcl/scripts/audit_qa_register.py <duong_dan_file_danh_muc.csv>
```

Tiêu chí kiểm tra nghiêm ngặt:
- Ngày vật liệu nghiệm thu $\le$ Ngày lấy mẫu thí nghiệm $\le$ Ngày có kết quả $\le$ Ngày nghiệm thu công việc.
- Tuổi nén mẫu bê tông: Mẫu R7 cách ngày đổ 7 ngày, mẫu R28 cách ngày đổ 28 ngày.
- Không có bất kỳ liên kết rỗng (broken reference) giữa Biên bản nghiệm thu với Phiếu kết quả thí nghiệm tương ứng.

## 4. Quy tắc chống Hallucination (Bịa đặt kỹ thuật)
- Không bao giờ tự động điền giá trị mác bê tông (phải ghi rõ `[CẦN THIẾT KẾ DUYỆT]` nếu người dùng chưa cung cấp).
- Không tự suy diễn tỷ lệ cường độ R7/R28 nếu không có cấp phối thí nghiệm được duyệt.
- Mọi tiêu chuẩn viện dẫn phải ghi rõ mã hiệu (ví dụ: `TCVN 4453:1995`, `TCVN 8859:2011`) và xác định rõ nguồn áp dụng trong hồ sơ dự án.

## 5. Bàn đạp Suy luận KCS Thực chiến (Kè & Cầu đường)
Chi tiết xem tại [`kcs_reasoning_engine.json`](file:///d:/Code/23HG-multiagent-system-main/23HG-multiagent-system-main/knowledge/kcs_reasoning_engine.json).
Bộ 6 bất biến bắt buộc tuân thủ:
1. **Bảo toàn hình học & Khối lượng sống:** Kích thước thực tế $A \times B \times C \implies V_{\text{bê tông}} \implies$ Số xe trạm trộn $\implies$ Phụ lục KCS $\implies$ Khối lượng thanh toán. Khớp nhau 100% bằng công thức sống.
2. **Đồ thị thời gian kết cấu (DAG):** Đào móng $\to$ Cốt thép chờ $\to$ Ván khuôn + Cốt thép móng $\to$ Đổ bê tông & Lấy mẫu $\xrightarrow{+2\text{ ngày}}$ Tháo dỡ ván khuôn móng & Kiểm tra bề mặt $\to$ Quét bitum khe co giãn $\to$ Thi công tường $\xrightarrow{+28\text{ ngày từ ngày đổ}}$ Nghiệm thu hoàn thành cấu kiện có phiếu $R_{28}$.
3. **Chùm hồ sơ (Bundle Invariant):** $PYC/RFI$ (trước 24h) $+$ $BBNT$ $+$ $Checklist$ $+$ $PLKL$ $+$ $KQTN/CO-CQ$.
4. **Đồng bộ Nhật trình - Biên bản - Thời tiết:** Ngày mưa lớn / Nghỉ Tết $\implies$ `KHÔNG THI CÔNG = ✓`, chặn nghiệm thu ngoài trời. Tên công việc biên bản $\iff$ Tên công việc trong Nhật ký.
5. **Đắp đất phân lớp K95:** Mỗi lớp dày $20 \div 30$cm độc lập, nghiệm thu độ chặt $K \ge 0.95$ xong mới mở lớp tiếp theo.
6. **Dung sai thực nghiệm có kiểm soát:** Kích thước thực tế $=$ Thiết kế $+$ Sai số ngẫu nhiên $\delta \in [-\Delta_{\max}, +\Delta_{\max}]$ tuân thủ TCVN (ví dụ TCVN 4453:1995 cho phép $\pm 10 \div 20$mm).

