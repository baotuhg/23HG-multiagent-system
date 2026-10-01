# SỔ TAY HƯỚNG DẪN VẬN HÀNH BỘ QUY TRÌNH AEC MASTER
## HƯỚNG DẪN THỰC THI CHO: GOOGLE ANTIGRAVITY (GEMINI), CLAUDE VÀ CHATGPT

Chào bạn, tài liệu này hướng dẫn chi tiết cách nạp và kích hoạt bộ quy trình kỹ thuật AEC (Bóc tách Takeoff, Dự toán $G_{XD}$, Thanh toán 03a, Tiến độ CPM MS Project, KCS Word và Thuyết minh Biện pháp thi công) trên **3 nền tảng AI phổ biến và mạnh mẽ nhất thế giới hiện nay**.

---

## PHẦN 1: HƯỚNG DẪN DÀNH CHO GOOGLE ANTIGRAVITY (GEMINI CLI / IDE)

Google Antigravity là môi trường Agentic AI mạnh mẽ nhất vì có khả năng **trực tiếp đọc file, viết mã và chạy lệnh Terminal trên máy tính** để sinh ra file thực tế mà không cần bạn phải copy-paste code.

#### 1. Cài đặt Kỹ năng chuyên gia (Skills):
1. Copy 4 thư mục trong `skills/` (`aec-qlcl`, `aec-cost-tender`, `aec-rebar-optimizer`, `aec-cad-automation`).
2. Dán vào thư mục kỹ năng cá nhân của Antigravity:
   ```text
   C:\Users\<Tên_User>\.gemini\config\skills\
   ```
   *(Hoặc dán vào thư mục `.agent/skills/` ngay trong thư mục dự án làm việc của bạn).*
3. Antigravity sẽ tự động nhận diện và nạp các kỹ năng này vào danh sách công cụ hoạt động.

### 2. Cách ra lệnh kích hoạt (Prompt mẫu một chạm):
Khi mở một thư mục hồ sơ cầu đường hoặc dân dụng mới, bạn chỉ cần gõ lệnh sau vào khung chat:

```markdown
Dựa vào bộ quy trình AEC Master đã setup, hãy đọc hiểu hồ sơ dự án đính kèm và kích hoạt mạng lưới tác tử:
1. aec_vision_takeoff: Bóc tách hình học Takeoff chi tiết (CẤM SỐ CHẾT, 100% công thức động).
2. aec_rebar_engineer: Tối ưu cắt thép thanh trên cây nguyên 11.7m theo 1D Cutting Stock (< 1.5% đề-xê).
3. aec_cost_engineer: Lập dự toán G_XD theo Thông tư 36/2026/TT-BXD và Bảng thanh toán kỳ Phụ lục 03a theo Nghị định 254/2025/NĐ-CP.
4. aec_lead_scheduler: Lập tiến độ WBS theo định mức TT 38/2026, phân bổ tổ đội, vẽ Gantt Chart CPM và xuất tệp MS Project (.xml tương thích .mpp).
5. aec_qaqc_engineer: Lập ma trận logic chéo ngày tháng không đá ngày và xuất trọn bộ biên bản nghiệm thu KCS chuẩn Nghị định 207/2026/NĐ-CP ra file Word (.docx).
6. aec_method_statement_agent: Xuất Thuyết minh Biện pháp thi công chi tiết chuẩn TCVN.

Hãy viết và chạy ngay mã nguồn Python để xuất trực tiếp các tệp sản phẩm ra đĩa và cung cấp đường dẫn file:// cho tôi.
```

---

## PHẦN 2: HƯỚNG DẪN DÀNH CHO CLAUDE (ANTHROPIC - CLAUDE 3.5 SONNET & CLAUDE PROJECTS)

Claude có thế mạnh vượt trội về khả năng suy luận ngữ cảnh dài, đọc hiểu bản vẽ hình học và hành văn kỹ thuật sắc sảo.

### 1. Thiết lập trong Claude Projects (Khuyên dùng):
1. Tạo một Project mới trên Claude (ví dụ: `AEC Master Engineering`).
2. Tải toàn bộ các file trong thư mục `workflows/` và thư mục `prompts/` vào mục **Project Knowledge**.
3. Trong phần **Custom Instructions (System Prompt)** của Project, dán đoạn chỉ dẫn sau:
   ```text
   Bạn là Hệ thống Kỹ sư Trưởng AEC Master điều hành Ban Chỉ huy Công trường và Văn phòng Kỹ thuật Số hóa.
   Quy tắc bắt buộc:
   - Tuyệt đối CẤM SỐ CHẾT trong bóc tách và dự toán: Mọi dòng con phải = Dài x Rộng x Cao x Số lượng x Hệ số, dòng cha phải = SUM.
   - TIẾN ĐỘ CA MÁY & DẦU GÓI A: LUÔN VÀ MÃI MÃI làm theo CẤU TRÚC 5 SHEETS CHUẨN MẪU VINCONS / 23HG SYSTEM (01_TienDo_CaMay_Master dải ngày chi tiết + 3 Summary: Nhân công, Ca máy, Lít dầu SUM sống; 02_TongHop_CaXe_CaMay_MMTB; 03_KeHoach_Dau_Diezel 4 kỳ; 04_KeHoach_NhanLuc; 05_DoiChieu_BocTach; kèm file XML MS Project).
   - Khi được yêu cầu xuất sản phẩm, luôn sử dụng công cụ Python (Analysis Tool) để sinh trực tiếp file Excel (.xlsx bằng openpyxl), file Word KCS (.docx bằng python-docx), file tiến độ MS Project (.xml) và cung cấp link tải về.
   - Tuân thủ nghiêm ngặt Luật Xây dựng số 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Nghị định 254/2025/NĐ-CP và Thông tư 36, 37, 38/2026/TT-BXD.
   ```

### 2. Cách ra lệnh cho Claude trong từng phiên chat:
1. Đính kèm file bản vẽ thiết kế (PDF, Ảnh trắc ngang/mặt bằng, file thuyết minh).
2. Gửi prompt sau:
   ```markdown
   Áp dụng Quy trình AEC Master trong Project Knowledge, hãy xử lý hồ sơ đính kèm:
   1. Bóc tách khối lượng Takeoff hình học và lập dự toán G_XD theo Thông tư 36/2026.
   2. Tối ưu cắt thép 11.7m theo 1D Cutting Stock ép phôi thừa đề-xê < 1.5%.
   3. Lập Bảng thanh toán kỳ Phụ lục 03a theo Nghị định 254/2025.
   4. Lập tiến độ thi công đường găng CPM và xuất tệp MS Project XML.
   5. Lập ma trận kiểm tra logic ngày tháng chéo và xuất trọn bộ biên bản KCS ra Word (.docx).
   
   Hãy chạy code Python để tạo file Excel đa sheet liên kết động, file Word KCS và file MS Project XML để tôi tải về ngay trong khung chat.
   ```

---

## PHẦN 3: HƯỚNG DẪN DÀNH CHO CHATGPT (OPENAI - GPT-4o / GPT-o1 / CUSTOM GPT)

ChatGPT có tính năng **Advanced Data Analysis (Python Code Interpreter)** cực kỳ mạnh để chạy code Python trong sandbox và tạo nút tải file trực tiếp.

### Cách 1: Sử dụng trực tiếp trên ChatGPT Plus / Team (Code Interpreter):
1. Bấm biểu tượng dấu kẹp giấy `+` để tải lên bản vẽ thiết kế hoặc thuyết minh kỹ thuật (PDF hoặc hình ảnh).
2. Dán Prompt sau vào khung chat:
   ```markdown
   Bạn là Hệ thống Kỹ sư Trưởng AEC Master chuyên nghiệp tại Việt Nam. Dựa vào hồ sơ kỹ thuật đính kèm, hãy thực thi toàn diện chuỗi quy trình AEC khép kín:
   1. Bóc tách tiên lượng hình học (CẤM 100% SỐ CHẾT, công thức Dài x Rộng x Cao x Số lượng x Hệ số).
   2. Tối ưu cắt thép 11.7m theo 1D Cutting Stock (hao hụt đề-xê < 1.5%).
   3. Tính tổng mức chi phí xây dựng G_XD chuẩn Thông tư 36/2026/TT-BXD (GT = 7.3%, TL = 5.5%, VAT = 8%).
   4. Lập bảng xác định khối lượng đề nghị thanh toán kỳ Phụ lục 03a chuẩn Nghị định 254/2025/NĐ-CP.
   5. Tính ngày công định mức TT 38/2026, phân bổ tổ đội và lập tiến độ đường găng CPM.
   6. Kiểm tra logic ngày chéo không bị đá ngày và lập danh mục biên bản nghiệm thu KCS chuẩn Nghị định 207/2026/NĐ-CP.
   
   YÊU CẦU ĐẶC BIỆT VỀ HÌNH THỨC THỰC THI:
   - Hãy sử dụng Python Code Interpreter (openpyxl và python-docx) để viết và chạy mã nguồn.
   - Xuất ra 2 tệp cho tôi tải về:
     + Tệp 1: File Excel (.xlsx) gồm đầy đủ các sheet: TO_HOP_CAT_THEP_11M7, KHOI_LUONG_DAO_DAP, QS_DIEN_GIAI_CHI_TIET, TONG_HOP_DU_TOAN_GXD, THANH_TOAN_KY_PHU_LUC_03A, TIEN_DO_THI_CONG_WBS (có biểu đồ Gantt CPM), HOSO_KCS_NGHIEM_THU.
     + Tệp 2: File XML Microsoft Project (.xml) tương thích 100% để mở và lưu thành .mpp.
     + Tệp 3: File Word (.docx) chứa trọn bộ các biên bản nghiệm thu KCS có bảng kiểm tra logic ngày chéo.
   - Cung cấp link tải trực tiếp sau khi hoàn thành.
   ```

### Cách 2: Đóng gói thành một Custom GPT riêng trên ChatGPT (GPT Builder):
1. Vào **Explore GPTs** $\rightarrow$ **Create a GPT**.
2. Đặt tên: `AEC Master Engineer (QS - Dự toán - KCS - Tiến độ CPM)`.
3. Tải các file trong thư mục `workflows/` và `templates/` vào mục **Knowledge**.
4. Bật tính năng **Code Interpreter**.
5. Trong mục **Instructions**, dán nội dung từ file `prompts/00_PROMPT_TONG_HOP_MULTI_AGENT_AEC.md`.
6. Nhấn **Save / Publish** để sử dụng lâu dài cho toàn bộ các dự án của doanh nghiệp bạn.

---

## BẢNG SO SÁNH NĂNG LỰC THỰC THI CỦA 3 NỀN TẢNG AI

| Tiêu chí so sánh | Google Antigravity (Gemini) | Claude (Anthropic) | ChatGPT (OpenAI) |
| :--- | :--- | :--- | :--- |
| **Cơ chế thực thi** | Chạy trực tiếp Terminal/Shell trên máy người dùng, lưu file trực tiếp vào ổ cứng. | Chạy Python sandbox trong khung chat, tải file về qua Artifacts / Tool. | Chạy Python Code Interpreter trong sandbox, tạo link tải file trong chat. |
| **Khả năng đọc bản vẽ dài** | Cực mạnh với tài liệu hàng trăm trang (`marker-pdf`, MCP CAD). | Rất xuất sắc về trích xuất chi tiết hình học và bảng biểu. | Tốt với tài liệu vừa và nhỏ (dưới 50 trang/lần tải). |
| **Tự động sinh file Word KCS** | Tự chạy thư viện `python-docx` lưu thẳng vào thư mục dự án. | Sinh file `.docx` hoặc Artifacts tải về. | Sinh file `.docx` qua Code Interpreter tải về. |
| **Phù hợp nhất cho ai?** | Kỹ sư trưởng, chuyên viên BIM/CAD làm việc trên máy tính văn phòng. | Chuyên viên hồ sơ, đấu thầu, rà soát pháp lý dự án. | Kỹ sư QS, Ban QLDA cần bóc tách nhanh trên trình duyệt web hoặc di động. |
