---
name: aec-cad-automation
description: Tự động hóa điều khiển AutoCAD/CAD trực tiếp qua giao thức MCP (autocad-mcp và cad-mcp), bóc tách khối lượng hình học bê tông, ván khuôn, cốt thép BBS và đào đắp mặt cắt ngang từ bản vẽ DWG/DXF không cần xuất file trung gian. Hỗ trợ giải mã font TCVN3/VNI sang Unicode và xuất bảng tính Excel 100% công thức sống.
---

# Tự động hóa CAD & Đo bóc Khối lượng Toàn diện (AEC-CAD Takeoff)

## 1. Mục tiêu & Kiến Trúc Hệ Thống MCP 3 Thành Phần

```text
[ Giao diện AI: Claude Desktop / Windsurf / Cursor / Antigravity ]
                               │
                               ▼ (Giao thức MCP: cad-mcp / autocad-mcp)
               [ MCP Server điều khiển AutoCAD (Local) ]
                               │
                               ▼ (API / COM Interop: win32com / ezdxf)
      [ Bản vẽ DWG trong AutoCAD ] <───> [ Hồ sơ thiết kế (Excel/PDF) ]
```

Cung cấp khả năng cho AI Agent tương tác trực tiếp 2 chiều với phần mềm AutoCAD (AutoCAD 2026/GstarCAD) thông qua giao thức MCP:
- **Đọc thông tin bản vẽ đang mở:** Danh sách layer, blocks, text (tự động giải mã font TCVN3 `.VnTime` sang Unicode), dim, polyline.
- **Quét hàng loạt bản vẽ DWG trong thư mục hồ sơ:** Tự động duyệt đệ quy và phân loại cấu kiện theo Modular WBS (Cọc, Mố, Trụ, Dầm Super-T, Bản mặt cầu, Lan can, Đào đắp).
- **Đo bóc Khối lượng Bê tông & Ván khuôn:** Tự động tính toán diện tích mặt cắt dầm, thể tích cọc khoan nhồi, bệ móng, thân mố trụ, xà mũ và diện tích ván khuôn tiếp xúc.
- **Đo bóc Đào đắp Đất đá Mặt cắt ngang:** Tự động tính diện tích khép kín (Shoelace) và khối lượng giữa các lý trình theo phương pháp diện tích trung bình 2 đầu (Average-End-Area).
- **Bóc tách Thống kê Cốt thép (BBS):** Trích xuất số hiệu thanh, đường kính $\Phi$, chiều dài, số lượng và quy đổi ra tổng chiều dài và trọng lượng (kg, tấn) theo TCVN 1651:2018.
- **Đối soát 2 chiều (`<───>`):** So khớp dữ liệu kích thước hình học từ CAD với bảng tiên lượng khối lượng Excel BoQ và thuyết minh Markdown.
- **Xuất bảng tính Excel chuẩn mực:** 100% công thức sống (`=E*F*G*H*I`, `=SUM(...)`), tuyệt đối tuân thủ nguyên tắc Zero Dead Numbers.

---

## 2. Công cụ MCP Khả dụng & Lệnh Thực thi

Agent có thể sử dụng các công cụ từ 2 kênh MCP:

### A. Kênh `cad-mcp` (GsLc API)
- `smart_cad_command`: Lệnh chính để thực thi từng thao tác CAD (vẽ circle, line, polyline, move, query layer, zoom...).
  - Cú pháp tham số: camelCase (ví dụ: `startPoint: [x, y, z]`, `endPoint: [x, y, z]`).
  - Điểm tọa độ: Mảng `[x, y, z]` (mặc định z=0).
- `batch_execute`: Thực thi chuỗi lệnh phụ thuộc liên tiếp để tăng tốc độ.
- `query_api_commands`: Tra cứu cú pháp lệnh API trước khi gọi.

### B. Kênh `autocad-mcp` (Slacker-LLC)
- `autocad_status`: Kiểm tra trạng thái kết nối với phiên bản AutoCAD đang mở.
- `list_open_drawings`: Liệt kê các bản vẽ DWG đang mở.
- `list_layers`: Đọc danh sách tất cả các layer trong bản vẽ.
- `draw_geometry` / `query_geometry`: Vẽ và trích xuất hình học 2D/3D.

---

## 3. Quy trình Đo bóc Khối lượng Chi tiết

### Bước 1: Khảo sát hiện trạng bản vẽ & Thư mục hồ sơ
1. **Kiểm tra kết nối CAD:** Kiểm tra xem AutoCAD có đang mở bản vẽ hợp lệ trên máy tính không (qua COM Interop `win32com.client`).
2. **Quét thư mục hồ sơ (Batch Scan):** Nếu có thư mục chứa nhiều file bản vẽ `.dwg`/`.dxf`, tự động quét đệ quy và phân loại theo hạng mục công trình.
3. **Quét danh sách Layer:** Tìm các layer chứa nét hình học kết cấu (`KET_CAU`, `THIET_KE`, `COT_THEP`, `MAT_CAT`, `DIA_CHAT`).

### Bước 2: Đo bóc Bê tông & Ván khuôn Kết cấu
Áp dụng công thức giải tích hình học chuẩn mực:
- **Cọc khoan nhồi:** $V = N \times \pi \times \left(\frac{D}{2}\right)^2 \times L$
- **Bệ móng mố, trụ:** $V = N \times L \times W \times H$; Diện tích ván khuôn: $S_{vk} = N \times 2 \times (L + W) \times H$.
- **Dầm Super-T (L=38.2m):** $V = N \times A_{mc} \times L$ với diện tích mặt cắt dầm $A_{mc} = 0.945\text{ m}^2$.
- **Bản mặt cầu liên tục nhiệt:** $V = L \times B \times t$ (với $t = 0.20\text{ m}$).

### Bước 3: Đo bóc Khối lượng Đào đắp Mặt cắt ngang
1. **Truy vấn Polyline:** Đọc tọa độ danh sách đỉnh $[(x_1, y_1), (x_2, y_2), ..., (x_n, y_n)]$.
2. **Tính diện tích hình học (Shoelace Formula):**
   $$A = \frac{1}{2} \left| \sum_{i=1}^{n-1} (x_i y_{i+1} - x_{i+1} y_i) + (x_n y_1 - x_1 y_n) \right|$$
3. **Tính khối lượng đào đắp giữa các lý trình:**
   $$V = \frac{A_1 + A_2}{2} \times L$$

### Bước 4: Đo bóc Thống kê Cốt thép (BBS)
1. Giải mã text block thống kê thép TCVN3 sang Unicode.
2. Trích xuất: Số hiệu thanh, Đường kính $d$ (mm), Chiều dài $L$ (m), Số lượng $N$.
3. Tính trọng lượng cốt thép theo TCVN 1651:2018:
   $$W = L \times N \times \left(0.006165 \times d^2\right) \quad (\text{kg})$$

### Bước 5: Xuất Bảng tính Excel Đo bóc Chuẩn hóa (Microsoft 365 Enterprise Engine)
Gọi script động cơ đo bóc CAD chuyên dụng tích hợp chuẩn Office 365:
```bash
# Xuất với động cơ Office 365 hiện đại (LET, LAMBDA, 00_DASHBOARD_365)
python skills/aec-cad-automation/scripts/cad_takeoff_engine.py --mode all --office365 --output templates/BANG_DO_BOC_KHOI_LUONG_CAD.xlsx

# Xuất chế độ công thức truyền thống (Legacy)
python skills/aec-cad-automation/scripts/cad_takeoff_engine.py --mode all --legacy --output templates/BANG_DO_BOC_KHOI_LUONG_CAD_LEGACY.xlsx
```

**Tính năng Office 365 Enterprise tích hợp:**
- **Custom AEC LAMBDAs:** Tự động đăng ký 8 hàm hình học (`V_PRISM`, `V_CYLINDER`, `V_FRUSTUM`, `S_FORMWORK_BOX`, `S_FORMWORK_TRI`, `STEEL_RATIO`, `V_AVERAGE_END`, `REBAR_WEIGHT`) vào Excel Name Manager.
- **Mô hình hóa hàm `LET()`:** Tự diễn giải hình học qua tên biến rõ ràng, tối ưu hóa tốc độ và không lặp lại phép tính.
- **Truy vấn liên sheet `XLOOKUP()`:** Khớp chuỗi ký tự đại diện wildcard (`"*...*"`), chống đứt gãy công thức khi chèn/xóa dòng.
- **Trang bìa `00_DASHBOARD_365`:** Tự động tạo 5 thẻ KPI điều hành thời gian thực và Bảng kiểm toán đối soát chéo (Audit Reconciliation Grid) với hồ sơ duyệt.
- **Visual Proof Anchoring:** Chèn trực tiếp ảnh trích xuất từ bản vẽ CAD vào bảng tính để làm minh chứng đối soát.

---

## 4. Xử lý Lỗi & Cơ chế Tự tiến hóa (Self-Evolution)
- **Lỗi UIPI (User Interface Privilege Isolation):** Đảm bảo tiến trình CAD và MCP Agent cùng cấp quyền (Elevation level).
- **Lỗi Handle không hợp lệ:** Khi một đối tượng bị xóa hoặc Undo, handle cũ sẽ mất hiệu lực; Agent phải truy vấn lại đối tượng mới nhất thay vì dùng handle cũ.
- **Xử lý font TCVN3:** Tự động dùng `TCVN3Decoder` chuyển đổi các ký tự lỗi hiển thị sang tiếng Việt Unicode chuẩn trước khi đưa vào bảng tính.
- **Tự động vá lỗi (FIX):** Khi một lệnh vẽ bị từ chối do trùng tên block hoặc sai kiểu dữ liệu tọa độ, Agent tự động hiệu chỉnh và thử lại.
