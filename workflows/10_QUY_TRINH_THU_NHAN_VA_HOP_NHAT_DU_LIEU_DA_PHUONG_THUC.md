# QUY TRÌNH THU NHẬN & HỢP NHẤT DỮ LIỆU ĐA PHƯƠNG THỨC (MULTI-MODAL INGESTION & DATA FUSION)
> **Trạng thái hiện thực:** `[NGUYÊN MẪU / MỘT PHẦN]`
> *Lưu ý:* Hệ thống hiện đọc trực tiếp dữ liệu cấu kiện/bảng khối lượng từ Excel/CSV/JSON/DWG/DXF/IFC. Khâu OCR xử lý hình ảnh chụp/PDF đang ở mức cấu hình định hướng.

> **Căn cứ kiến trúc hệ thống:**
> - **Kiến trúc Bảng tin tập trung (Blackboard Architecture - Single Source of Truth)**
> - **Nguyên lý Đơn nhiệm (Single Responsibility Principle - SRP)**
> - **Cơ chế Đối chiếu chéo Đa phương thức (Cross-Modal Reconciliation)**
> - **Chuẩn mã hóa ký tự Quốc gia (TCVN3 / VNTIME sang Unicode UTF-8)**

---

## 1. Vấn đề cốt lõi trong Xử lý Dữ liệu Xây dựng (AEC Data Dilemma)

Trong các dự án giao thông, cầu đường và hạ tầng kỹ thuật tại Việt Nam, hồ sơ thiết kế và thi công được lưu trữ dưới nhiều định dạng hoàn toàn khác nhau:
1. **Bản vẽ hình học (.dwg, .dxf):** Chứa các nét vẽ mặt bằng, trắc dọc, mặt cắt ngang, khối lượng đào đắp, chi tiết bố trí cốt thép và bảng thống kê thép (BBS) vẽ tay. Thường dùng font chữ cổ TCVN3 (`.VnTime`, `.VnArial`).
2. **Bảng tính toán khối lượng (.xlsx, .xls):** Các file Excel tính toán của Tư vấn Thiết kế (TEDI) chứa hàng chục sheet, công thức liên kết phức tạp giữa các cấu kiện.
3. **Thuyết minh kỹ thuật & Tiêu chuẩn (.md, .docx, .pdf):** Thuyết minh biện pháp thi công, chỉ dẫn kỹ thuật dự án (Technical Specifications) quy định mác bê tông, mác thép, độ sụt, điều kiện căng kéo cáp DƯL và các tiêu chuẩn kiểm soát chất lượng (TCVN, QCVN, ASTM).

Nếu để các tác tử nghiệp vụ (Cắt thép 1D, Dự toán G_XD, Lập tiến độ, Làm biên bản KCS) tự đọc trực tiếp các tệp thô này:
- ❌ **Tràn bộ nhớ ngữ cảnh (Context Bloat):** Token của mô hình AI bị ngập trong hàng ngàn nét vẽ CAD và cấu trúc XML của Office.
- ❌ **Lệch pha số liệu (Design Discrepancy):** Bản vẽ CAD đã sửa sang lần hiệu chỉnh 2 (Rev 02), nhưng bảng tính Excel vẫn giữ số liệu Rev 01, dẫn tới sai sót dây chuyền.
- ❌ **Lỗi font chữ (Encoding Errors):** Không đọc được tiếng Việt do xung đột mã font TCVN3/VNI.

---

---

## 2. ⚙️ Kiến Trúc Hệ Thống MCP 3 Thành Phần (3-Tier MCP Architecture)

Để giải quyết bài toán đọc hiểu bản vẽ khi trong thư mục hồ sơ có thêm hàng chục đến hàng trăm bản vẽ kỹ thuật, hệ thống thiết lập kiến trúc 3 tầng chuẩn hóa:

```text
[ Giao diện AI: Claude Desktop / Windsurf / Cursor / Antigravity ]
                               │
                               ▼ (Giao thức MCP: cad-mcp / autocad-mcp)
               [ MCP Server điều khiển AutoCAD (Local) ]
                               │
                               ▼ (API / COM Interop: win32com / ezdxf)
      [ Bản vẽ DWG trong AutoCAD ] <───> [ Hồ sơ thiết kế (Excel/PDF/MD) ]
```

### Chi tiết 3 tầng hoạt động:
1. **Tầng 1 - Giao diện AI (AI Client Interface):**
   - Các nền tảng trợ lý AI chuyên gia: Google Antigravity, Claude Desktop, Cursor, Windsurf.
   - Tiếp nhận lệnh từ kỹ sư, điều phối tác tử và gửi yêu cầu truy vấn bản vẽ qua giao thức mở MCP.
2. **Tầng 2 - MCP Server Cục bộ (Local MCP Server Bridge):**
   - Đóng vai trò cầu nối chuẩn hóa (chạy các server như `cad-mcp` hoặc `autocad-mcp`).
   - Cung cấp các công cụ: `smart_cad_command`, `batch_execute`, `query_api_commands`, `list_open_drawings`, `list_layers`.
3. **Tầng 3 - Tương tác AutoCAD & Tệp Bản vẽ DWG (API / COM Interop):**
   - **Chế độ Phiên hoạt động (Active Session):** Kết nối trực tiếp vào phần mềm AutoCAD đang mở trên máy tính qua COM Interop (`win32com.client.Dispatch("AutoCAD.Application")`), đọc đối tượng từ `ModelSpace` (Text, MText, Polyline, Block, Dim).
   - **Chế độ Quét đệ quy (Batch Drawing Folder Scanner):** Khi trong thư mục dự án có thêm hàng loạt bản vẽ mới, tác tử tự động duyệt toàn bộ cây thư mục (`os.walk`), phân loại từng bản vẽ theo WBS (Cọc, Mố, Trụ, Dầm Super-T, Mặt cầu, Lan can, Đào đắp...).
   - **Đối soát 2 chiều (`<───>`):** Tự động so khớp dữ liệu hình học và bảng thống kê thép (BBS) trích từ bản vẽ DWG với các bảng tính khối lượng (Excel BoQ) và Thuyết minh kỹ thuật (Markdown/PDF), phát hiện ngay nếu có sai lệch kích thước hay mác vật liệu.

---

## 3. Sơ đồ Luồng Dữ liệu Hợp nhất Đa phương thức

```
[Bản vẽ CAD DWG/DXF]       [Bảng tính Excel/Word]       [Hồ sơ Thuyết minh Markdown]
         │                            │                               │
         ▼                            ▼                               ▼
 ┌───────────────┐            ┌────────────────┐             ┌─────────────────┐
 │aec_cad_       │            │aec_office_     │             │aec_markdown_    │
 │extractor      │            │extractor       │             │ingestor         │
 └───────┬───────┘            └───────┬────────┘             └────────┬────────┘
         │                            │                               │
         └────────────────────────────┼───────────────────────────────┘
                                      │ (Dữ liệu bóc tách thô)
                                      ▼
                      ┌───────────────────────────────┐
                      │     aec_data_aggregator       │
                      │  (Master Data Synthesizer)    │
                      │  • Đối chiếu chéo CAD vs Excel│
                      │  • Giải mã TCVN3 -> Unicode   │
                      │  • Gắn cờ cảnh báo sai khác   │
                      └───────────────┬───────────────┘
                                      │ (Dữ liệu chuẩn hóa - Canonical Data)
                                      ▼
                      ┌───────────────────────────────┐
                      │    BLACKBOARD ARCHITECTURE    │
                      │      (PROJECT_STATE.json)     │
                      └───────────────┬───────────────┘
                                      │
         ┌────────────────────────────┼───────────────────────────────┐
         ▼                            ▼                               ▼
 ┌───────────────┐            ┌────────────────┐             ┌─────────────────┐
 │aec_rebar_     │            │aec_cost_       │             │aec_qaqc_        │
 │engineer       │            │engineer        │             │engineer         │
 │(Cắt thép 1D)  │            │(Dự toán & 03a) │             │(Tần suất & KCS) │
 └───────────────┘            └────────────────┘             └─────────────────┘
```

---

## 4. Đặc tả Nhiệm vụ 4 Tác tử Thu nhận & Hợp nhất

### 3.1. Tác tử `aec_cad_extractor` (Chuyên gia CAD DWG)
- **Vị trí tệp:** `core/agents/aec_cad_extractor.py`
- **Chức năng:**
  - Tích hợp bộ giải mã font `TCVN3Decoder` tự động chuyển các chuỗi ký tự `.VnTime` thành tiếng Việt Unicode chuẩn.
  - Quét cấu trúc bản vẽ, phát hiện các block bảng thống kê thép (BBS) vẽ trong AutoCAD.
  - Bóc tách tọa độ đỉnh của mặt cắt ngang đào đắp nền đường, móng kè, tính diện tích theo thuật toán Shoelace ($A = \frac{1}{2} | \sum (x_i y_{i+1} - x_{i+1} y_i) |$).
  - Nhận diện các cấu kiện chủ đạo: Dầm Super-T 38.2m, Mố M1/M2, Trụ T1/T2, Cọc khoan nhồi D1200.

### 3.2. Tác tử `aec_office_extractor` (Chuyên gia Office Excel & Word)
- **Vị trí tệp:** `core/agents/aec_office_extractor.py`
- **Chức năng:**
  - Bóc tách bảng thống kê thép từ các file Excel thiết kế gốc (`CT-km19.5.xls`, `KLDam.xls`, `KL-tru-2than.xls`, `Coc KN D=1200.xls`).
  - Kiểm tra cây công thức, phân loại rạch ròi giữa ô tính có công thức sống (`=SUM`, `=PRODUCT`) và số chết (hardcoded numbers).
  - Trích xuất bảng tiên lượng mời thầu (BoQ) và mẫu biên bản nghiệm thu Word.

### 3.3. Tác tử `aec_markdown_ingestor` (Chuyên gia Hồ sơ Markdown)
- **Vị trí tệp:** `core/agents/aec_markdown_ingestor.py`
- **Chức năng:**
  - Đọc hồ sơ thiết kế, thuyết minh biện pháp thi công và quy chuẩn kỹ thuật dạng Markdown.
  - Tự động bóc tách các bảng biểu Markdown thành danh sách bản ghi có cấu trúc.
  - Trích xuất tự động các tham số kỹ thuật then chốt: mác bê tông từng cấu kiện (C10, C25, C30, C35, C45), mác cốt thép (CB240-T, CB400-V, ASTM A416), các tiêu chuẩn viện dẫn (TCVN 4453, TCVN 1651, TCVN 9396).

### 3.4. Tác tử `aec_data_aggregator` (Trọng tài Hợp nhất & Trưởng ban Dữ liệu)
- **Vị trí tệp:** `core/agents/aec_data_aggregator.py`
- **Chức năng cốt lõi:**
  - **Cửa khẩu kiểm dịch dữ liệu (Quality Gate):** Tiếp nhận dữ liệu từ 3 tác tử trên và tiến hành đối chiếu chéo (Cross-modal Reconciliation).
  - **Phát hiện lệch pha (Discrepancy Detection):** Nếu CAD ghi cọc $L=44\text{ m}$ nhưng Excel ghi $L=40\text{ m}$, tự động gắn cờ cảnh báo `[WARNING]` và đề xuất phương án xử lý theo hồ sơ được phê duyệt.
  - **Đồng bộ Blackboard:** Xuất dữ liệu sạch đã được kiểm chứng vào `templates/PROJECT_STATE.json` để toàn bộ các tác tử kỹ thuật phía sau sử dụng.

---

## 4. Hướng dẫn Thực thi và Kiểm thử

Chạy trực tiếp chuỗi quy trình thu nhận và đối chiếu chéo từ terminal:

```bash
python examples/run_data_ingestion_pipeline.py
```

Kết quả mẫu:
```text
===========================================================================
  AEC MULTI-MODAL DATA INGESTION & RECONCILIATION PIPELINE
===========================================================================
--- BƯỚC 1: TÁC TỬ AEC_CAD_EXTRACTOR QUÉT BẢN VẼ CAD ---
  -> Nhận diện cấu kiện CAD: DẦM CHỦ SUPER-T 38.2M
--- BƯỚC 2: TÁC TỬ AEC_OFFICE_EXTRACTOR BÓC TÁCH BẢNG TÍNH EXCEL ---
  -> Tổng số sheet trong tệp Excel: 3
--- BƯỚC 3: TÁC TỬ AEC_MARKDOWN_INGESTOR ĐỌC HỒ SƠ THUYẾT MINH ---
  -> Mác bê tông phát hiện: ['C10', 'C25', 'C30', 'C35', 'C40', 'C45', 'C50']
  -> Mác thép phát hiện: ['ASTM A416', 'CB240-T', 'CB400-V']
--- BƯỚC 4: TÁC TỬ AEC_DATA_AGGREGATOR ĐỐI CHIẾU CHÉO & HỢP NHẤT ---
[aec_data_aggregator] [V] Đối chiếu hoàn tất: 100% số liệu KHỚP CHUẨN, KHÔNG CÓ XUNG ĐỘT!
  -> Đã tổng hợp thành công Single Source of Truth:
     * Thống kê thép BBS: 390 số hiệu thanh (883.902 tấn)
     * Bê tông các loại: 3460.667 m3 (1406.26 tấn Xi măng)
     * Ma trận tần suất KCS: 809 tổ mẫu kiểm tra bắt buộc
===========================================================================
```
