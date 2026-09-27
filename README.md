# 🏗️ 23HG-AEC-MultiAgent-System
### Nền tảng Đa tác tử Kỹ thuật số hóa & Quản lý Dự án Xây dựng (Closed-Loop ConTech System)
**Kỹ sư Trưởng Số hóa: Bóc tách Hình học • Cắt thép 1D OR-Tools • Dự toán G_xd • Thanh toán 03a • Tiến độ CPM MS Project • Hồ sơ KCS & BPTC • Vòng lặp Hiện trường As-Built**

---

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Standards: TCVN & BXD](https://img.shields.io/badge/Standards-TCVN%20%7C%20Lu%E1%BA%ADt%20XD%20135%2F2025-brightgreen)](workflows/)
[![Optimization: Google OR-Tools](https://img.shields.io/badge/Optimization-OR--Tools%20Column%20Generation%20%2B%20CP--SAT-blue)](tools/cutting_stock_solver.py)
[![Rebar Cutting: proven lower bound](https://img.shields.io/badge/Rebar%20Cutting-C%E1%BA%ADn%20d%C6%B0%E1%BB%9Bi%20%2B%20t%E1%BB%91i%20%C6%B0u%20ch%E1%BB%A9ng%20minh-success)](tools/cutting_stock_solver.py)
[![Zero Dead Numbers](https://img.shields.io/badge/Math-100%25%20Dynamic%20Formulas-red.svg)](templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx)

---

## 📖 1. Giới thiệu Tổng quan (Overview)

**23HG-AEC-MultiAgent-System** là nền tảng ConTech mã nguồn mở chuyên sâu cho quản lý kỹ thuật, dự toán và điều hành thi công xây dựng tại Việt Nam. Hệ thống chuyển đổi mô hình tự động hóa từ chuỗi tuyến tính (Linear Pipeline) sang **Hệ thống Đa tác tử Khép kín (Closed-Loop Multi-Agent System)** vận hành trên **Đồ thị Trạng thái (State Graph v3.0)** với sự điều phối tập trung của **AI Supervisor (Chỉ huy trưởng ảo)**.

### Cơ sở Pháp lý & Tiêu chuẩn Kỹ thuật:
- **Luật Xây dựng số 135/2025/QH15** & **Nghị định số 207/2026/NĐ-CP**: Quản lý chất lượng thi công, nhật ký thi công, giám sát và nghiệm thu KCS.
- **Nghị định số 99/2021/NĐ-CP**: Quản lý, thanh toán, quyết toán dự án vốn đầu tư công (Bảng thanh toán khối lượng hoàn thành Phụ lục 03a).
- **Thông tư số 11/2021/TT-BXD**: Phương pháp xác định và quản lý chi phí đầu tư xây dựng (Dự toán chi phí xây dựng `G_xd = T + GT + TL + VAT 10%`).
- **Thông tư số 12/2021/TT-BXD**: Định mức dự toán xây dựng công trình, định mức hao phí vật tư (xi măng, cát, đá, cốt thép, cáp DƯL), nhân công và ca máy.
- **Tiêu chuẩn thiết kế & thi công**: **TCVN 11823:2017** (Cầu đường bộ), **TCVN 5574:2018** (Kết cấu BTCT - Quy chuẩn nối cốt thép), **TCVN 1651:2018** (Thép thanh vằn cốt bê tông), **TCVN 9395:2012** (Cọc khoan nhồi), **TCVN 4453:1995** (Toàn khối).

---

## 🏛️ 2. Sơ đồ Kiến trúc Hệ thống (System Architecture)

Hệ thống hoạt động theo mô hình **Supervisor & Shared State Bus**, phân định tuyệt đối giữa năng lực suy luận nhận dạng (LLM Intent Recognition) và năng lực tính toán số học xác định (Deterministic Pure Python Code, Zero LLM Math):

```text
                           +-------------------------------+
                           |      BẢN VẼ / HỒ SƠ DỰ ÁN     |
                           |   (CAD / BIM / Yêu cầu KTXD)  |
                           +---------------+---------------+
                                           |
                                           v
                  +-------------------------------------------------+
                  |       AI SUPERVISOR (CHỈ HUY TRƯỞNG ẢO)         |
                  | - Tiếp nhận mục tiêu, phân bổ đầu việc Modular  |
                  | - Điều phối State dự án qua Shared State Bus    |
                  | - Kiểm soát 4 Cổng Chất lượng (Quality Gates)   |
                  +-----------------------+-------------------------+
                                          |
        +---------------------------------+---------------------------------+
        |                                 |                                 |
        v                                 v                                 v
+------------------+             +------------------+             +------------------+
| CAD/BIM PARSER   |             | KỸ THUẬT BPTC    |             | QS DỰ TOÁN G_XD  |
| (Trắc đạc CAD)   |             | & KCS LAB LINK   |             | & PHỤ LỤC 03A    |
+--------+---------+             +--------+---------+             +--------+---------+
| * Đọc DWG/DXF    |             | * Kiểm soát BPTC |             | * Đơn giá TT 12  |
| * Shoelace diện  |             | * Lab Link R7/R28|             | * Tính G_xd      |
|   tích, thể tích |             | * 22 BBNT chuẩn  |             | * 100% công thức |
| * Average-End    |             | * Hold Points NT |             | * Phụ lục 03a    |
+--------+---------+             +--------+---------+             +--------+---------+
    |    |                                ^                                ^
    |    |                                |                                |
    |    v                                | (Phản biện vị trí nối thép)    | (BOM thép)
    |  +------------------+               | TCVN 5574: CẤM nối vùng kéo    |
    |  | CẮT THÉP 1D      +---------------+--------------------------------+
    |  | (OR-TOOLS SOLVER)|
    |  +------------------+
    |  | * Google OR-Tools| ---> Tối ưu số cây 11.7m theo từng Ø + mác thép
    |  | * CP-SAT / FFD   | ---> Kiểm tra chéo (Inter-Agent): REJECT nếu sai
    |  +--------+---------+
    |           |
    +-----------+
    |
    v
+---------------------------------------------------------------------------+
|                          LẬP TIẾN ĐỘ CPM & GANTT                          |
| * Sắp xếp Topo, Forward Pass (ES/EF) & Backward Pass (LS/LF), Total Float |
| * Nhận diện Đường găng (Critical Path Chain) & Phân bổ nhân lực thiết bị  |
+------------------------------------+--------------------------------------+
                                     |
                                     v
                  +-------------------------------------------------+
                  |       SHARED STATE BUS (SINGLE SOURCE OF TRUTH) |
                  |  8 Miền: Meta • CAD • Rebar • QS • QAQC • CPM   |
                  |          As-Built • Human Approvals (RLock Bus) |
                  +------------------+------------------------------+
                                     |
            [Xung đột / Vi phạm?] ---+---> [Có] ---> REJECT / RETRY LOOP
                                     |               (Solver chạy lại hoặc
                                     |                BPTC điều chỉnh biện pháp)
                                   [Không]
                                     |
                                     v
                  +-------------------------------------------------+
                  |      HUMAN-IN-THE-LOOP QUALITY GATE             |
                  |  Trạng thái AWAITING_APPROVAL: Cảnh báo clash,  |
                  |  Kỹ sư trưởng / Giám đốc ký số phê duyệt        |
                  +------------------+------------------------------+
                                     |
                                     v
                  +-------------------------------------------------+
                  |      VÒNG LẶP ĐỐI SOÁT HIỆN TRƯỜNG (AS-BUILT)   |
                  |  Daily Site Log + Khối lượng thi công thực tế    |
                  |  -> Cập nhật CPM thực tế & Phát sinh Phụ lục 03a|
                  +-------------------------------------------------+
```

---

## 📥 3. Chuẩn I/O (Input/Output Specifications)

| Thành phần | Định dạng Đầu vào (Input) | Định dạng Đầu ra (Output) | Công cụ & Đặc tả Kỹ thuật |
|---|---|---|---|
| **Bóc tách CAD/BIM** | Bản vẽ CAD `.dwg`, `.dxf`, mô hình `.ifc` | Bảng khối lượng hình học Bê tông, Ván khuôn, Đào đắp | `AutoCAD COM Interop`, `ezdxf`, Shoelace & Average-End-Area |
| **Gia công Cốt thép** | File BBS thật `.xlsx` / `.csv` / `.json` (`--bbs`) | Phiếu cắt từng phương án cây 11.7m (CSV), số cây, cận dưới, đề-xê, mẩu thừa tận dụng | OR-Tools Column Generation (GLOP) + CP-SAT, tách nhóm Ø + mác thép, tính lưỡi cắt 3mm |
| **Dự toán Chi phí** | Khối lượng trích xuất, Đơn giá định mức | Bảng dự toán tổng hợp chi phí xây dựng `G_xd` | Excel 100% công thức động (`G_xd = T + GT + TL + VAT 10%`) |
| **Thanh toán Hợp đồng**| Khối lượng thiết kế vs Khối lượng hoàn công | Bảng xác định khối lượng hoàn thành Phụ lục 03a | Nghị định 99/2021/NĐ-CP, tính phát sinh tự động |
| **Quản lý Tiến độ** | File tiến độ thật `MS Project .xml` / `.xlsx` / `.csv` (`--schedule`) | Bảng CPM (ES/EF/LS/LF, dự trữ, đường găng, ngày lịch) CSV; cảnh báo ngày trong file vi phạm quan hệ logic | CPM với quan hệ FS/SS/FF/SF + lag, lịch nghỉ (Chủ nhật, ngày lễ) |
| **Quản lý Chất lượng**| Phiếu thí nghiệm nén R7/R28, kéo thép, PDA | 22 Biên bản nghiệm thu KCS in ấn A4 chuẩn | Excel A4 Form (`MAU_BIEN_BAN_KCS`, thay thế hoàn toàn Word) |
| **Biện pháp Thi công** | Yêu cầu KTXD, điều kiện địa chất, thủy văn | Thuyết minh BPTC 8 chương TCVN | Markdown chuẩn kỹ thuật, tích hợp RAG Hugging Face |
| **Đối soát Hiện trường**| Nhật ký thi công hàng ngày `DailySiteLog` | Báo cáo chênh lệch tiến độ & Chi phí phát sinh | As-Built Closed Loop, tự động cập nhật mạng CPM |

---

## 🗺️ 4. Lộ trình Phát triển (Roadmap 3 Phase)

```text
  Phase 1 (v1.x) [100% HOÀN THÀNH]
  ├── Tự động hóa tác vụ kỹ thuật cốt lõi (Core Engines)
  ├── Bóc tách hình học Takeoff 100% công thức động (0 số chết)
  ├── Cắt thép 1D Cutting Stock: tối ưu số cây theo từng Ø + mác thép, có cận dưới chứng minh
  ├── Bảng phân tích định mức & Tổng hợp vật tư toàn cầu BOM
  ├── Dự toán G_xd Thông tư 11/2021 & Thanh toán kỳ Phụ lục 03a
  └── Bộ 14 Sheet Master Excel đạt 100/100 điểm Audit Verifier

  Phase 2 (v2.x) [100% HOÀN THÀNH - State Graph v3.0]
  ├── Tái cấu trúc State Graph & AI Supervisor điều phối tập trung
  ├── Cách ly 100% toán học số học khỏi LLM (Google OR-Tools, Pure Python CPM)
  ├── Phản biện chéo Inter-Agent Negotiation: Thẩm tra vị trí nối thép TCVN 5574
  ├── Vòng lặp đối soát hiện trường (Site Feedback & As-Built Loop)
  ├── QA/QC Lab Link: Tích hợp kết quả nén R7/R28, siêu âm cọc, PDA vào KCS
  └── Cổng phê duyệt Kỹ sư trưởng Human-in-the-loop (Awaiting Approval)

  Phase 3 (v3.x) [ĐANG TRIỂN KHAI & MỞ RỘNG]
  ├── Động cơ so sánh phiên bản CAD/BIM Versioning (Incremental Diff Rev00 vs Rev01)
  ├── Tích hợp Ký số điện tử (E-Signatures / PKI) trực tiếp trên Web Dashboard
  ├── Mở rộng Multi-Project State Graph quản trị đồng thời nhiều phân đoạn cao tốc
  └── Giao diện WebUI / Mobile App cho Kỹ sư hiện trường nhập Daily Log trực tiếp
```

---

## ⚡ 5. Hướng dẫn Cài đặt & Bắt đầu Nhanh (Quick Start)

### 1. Yêu cầu Hệ thống
- Hệ điều hành: Windows 10/11 (hỗ trợ tốt nhất cho AutoCAD COM Interop) hoặc Linux/macOS.
- Python: Phiên bản **3.10** trở lên.
- RAM: Tối thiểu 8 GB (khuyến nghị 16 GB khi xử lý file DWG lớn).

### 2. Cài đặt Môi trường
```powershell
# 1. Clone kho lưu trữ mã nguồn
git clone https://github.com/baotuhg/23HG-multiagent-system.git
cd 23HG-multiagent-system

# 2. Tạo và kích hoạt môi trường ảo Python (Virtual Environment)
python -m venv venv
.\venv\Scripts\activate       # Trên Windows PowerShell
# source venv/bin/activate    # Trên Linux/macOS

# 3. Cài đặt các gói phụ thuộc kỹ thuật
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Các Lệnh Thực thi Chính

> ⚠ **Dữ liệu thật và dữ liệu mẫu:** mặc định hệ thống **chỉ dùng dữ liệu thật**. Pha nào thiếu dữ liệu sẽ dừng ngay và báo rõ cần cung cấp gì, không tự thay bằng số liệu mẫu. Dữ liệu mẫu (Cầu Km19+529.080) chỉ được dùng khi có cờ `--demo`, và mọi chỗ dùng đều được đánh dấu trong log, Quality Gate, Human Gate và báo cáo cuối.
>
> Hiện tại có **pha cắt thép** (BBS), **pha tiến độ CPM** (MS Project XML / Excel / CSV) và **pha dự toán G_XD** (bảng QS) chạy được bằng dữ liệu thật. Các pha CAD, QA/QC và hoàn công vẫn chỉ có dữ liệu mẫu trong code, nên sẽ dừng nếu không có `--demo`.
>
> File Excel chỉ có công thức mà chưa từng được Excel tính (ví dụ file do phần mềm tạo ra) vẫn đọc được: `tools/excel_eval.py` tự tính các hàm thông dụng (SUM, SUMIF(S), COUNTIF(S), ROUND/ROUNDUP, tham chiếu sang sheet khác…). Gặp hàm chưa hỗ trợ thì hệ thống báo rõ, không đoán.

#### a. Tối ưu cắt thép từ BBS thật (dùng được cho dự án):
```powershell
python run_state_graph.py --phase rebar --bbs "BBS_du_an.xlsx" --cut-plan-out phieu_cat_thep.csv
```
> - Đọc được Excel (tự tìm sheet có bảng BBS, hoặc chỉ định `--bbs-sheet`), CSV (`,` `;` hoặc tab) và JSON. Cột nhận diện theo tiêu đề tiếng Việt hoặc tiếng Anh: *Ký hiệu thanh*, *Đường kính Ø (mm)*, *Mác thép*, *Chiều dài 1 thanh (m)* hoặc `length_mm`, *Tổng số thanh* hoặc *Số thanh / cấu kiện* × *Số cấu kiện*. Cột chiều dài bắt buộc ghi đơn vị (m hoặc mm).
> - Dòng BBS sai dữ liệu (đường kính không tiêu chuẩn, chiều dài quá ngắn, số lượng lẻ...) làm hệ thống **dừng và liệt kê từng dòng**. Nếu muốn loại các dòng đó và tiếp tục, thêm `--bbs-skip-invalid`; các dòng bị loại vẫn được cảnh báo.
> - Solver chỉ ghép các đoạn **cùng đường kính và cùng mác thép**, trừ 3mm lưỡi cắt mỗi nhát, rồi báo **cận dưới** số cây. `OPTIMAL` nghĩa là đã chứng minh không thể dùng ít cây hơn. Nếu đề-xê vẫn > 1.5% thì đó là do chiều dài thanh trong BBS, không phải do cách ghép.
> - **Thanh dài hơn 11.7m** được tự tách thành k đoạn nối (vd 39.85m thành 4 đoạn, 3 mối nối), nếu có vùng cho phép nối (cột *Vùng cho phép nối* hoặc `--splice-zone`):
>   - Mỗi đoạn ≤ cây thép và ≥ max(L nối, 20D). Mọi vùng chồng nối nằm trong vùng cho phép.
>   - **Mối nối so le theo từng dòng BBS:** tâm các mối nối cách nhau < 1.3 × L_nối coi là cùng mặt cắt, và mỗi mặt cắt có tối đa `--max-splice-ratio` (mặc định 50%) số thanh có mối nối.
>   - Cách tách được tối ưu cùng lúc với phần cắt các thanh khác, ưu tiên đoạn dài đúng 11.7m để dùng trọn cây.
>   - Thanh không có vùng nối, hoặc không bố trí được mối nối trong vùng cho phép, được liệt kê là **CHƯA có trong kế hoạch cắt**.
> - Cáp DƯL không cắt từ cây thép và được cảnh báo riêng. Vị trí nối phải đối chiếu với bản vẽ (`NOT_RUN` ở Gate-2).
>
> - **Giới hạn cho tổ cắt:** `--max-pieces-per-bar 4 --max-marks-per-bar 2`. **Cắt đầu cây:** `--end-trim-mm 50`. **Lưỡi cắt:** `--kerf-mm 3`. **Đầu thừa** được phân loại *Tái sử dụng* (≥ 100D, đổi bằng `--reuse-xd`), *Đầu thừa ngắn* (≥ 20D) hoặc *Phế*.
> - **Phương án nối thép tận dụng đầu thừa** (`--splice`), đưa phép nối vào ngay mô hình tối ưu OR-Tools, chặt hơn PA4 của RebarCut:
>   - Chỉ nối Bar Mark có *Cho nối? = Có* **và** có *Vùng cho phép nối* dạng số (vd `0-0.25; 0.75-1`, `0-25%`, hoặc `0-2.5m`). Có thể đặt vùng mặc định cho mọi thanh bằng `--splice-zone`.
>   - Vùng chồng nối phải nằm trọn trong vùng cho phép. Mỗi đoạn ≥ max(L nối, 20D). Số thanh được nối không vượt *Max nối %* của từng Bar Mark.
>   - Phiếu cắt chính thức vẫn là phương án không nối. Phương án nối được xuất riêng (`*_PA_noi.csv`, sheet `PA_NOI`, `MOI_NOI`) để kỹ thuật duyệt.
> - **Xuất theo bố cục RebarCut Pro Excel:** `--rebarcut-out ket_qua.xlsx`, gồm các sheet INPUT, SO_SANH, PA_TOI_UU, PA_NOI, MOI_NOI, REMAIN, CHI_TIET. Sheet INPUT copy sang RebarCut được, và hệ thống cũng đọc được sheet INPUT của file RebarCut (`--bbs RebarCut_Pro_Excel_V5.xlsm`).
>
> Ví dụ với BBS trong workbook mẫu: `python run_state_graph.py --phase rebar --bbs "templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx" --bbs-skip-invalid`

#### a2. Tính tiến độ CPM từ file tiến độ thật:
```powershell
python run_state_graph.py --phase schedule --schedule "TienDo.xml" --non-working-days cn --holidays 2027-02-05:2027-02-12 --schedule-out tien_do_cpm.csv
```
> - Đọc được **MS Project XML** (File → Save As → XML trong MS Project), Excel, CSV hoặc JSON. Cột nhận diện theo tiêu đề: *Mã WBS*, *Danh mục công tác*, *Thời gian (ngày)*, *Quan hệ logic* (vd `1.2FS; 1.3SS+3d`), *Ngày bắt đầu/hoàn thành* (tùy chọn, dùng để đối chiếu).
> - Hỗ trợ quan hệ **FS / SS / FF / SF** có độ trễ (âm hoặc dương), ngày nghỉ trong tuần (`--non-working-days t7,cn`) và ngày lễ (`--holidays`). Ngày khởi công lấy từ file, hoặc chỉ định bằng `--start-date`.
> - Liên kết tới công việc không tồn tại, vòng lặp logic, thời lượng sai: **dừng và liệt kê từng lỗi**. Ô Excel là công thức chưa được tính thì phải mở và lưu lại file bằng Excel trước.
> - **Đối chiếu ngày ghi trong file:** cảnh báo khi ngày trong file vi phạm chính quan hệ logic của nó. Ví dụ file mẫu `Tien_Do_Thi_Cong_Cau_Km19+529.080.xml` có 4 công việc khai báo FS nhưng ngày lại chồng lấn, nên tính lại ra 187 ngày (hoàn thành 05/04/2027) thay vì 28/03/2027 như ghi trong file.

#### a3. Tính dự toán G_XD từ bảng QS thật:
```powershell
python run_state_graph.py --phase qs --qs "Du_toan.xlsx" --qs-out du_toan_gxd.xlsx
```
> - Đọc bảng QS / BOQ (Excel, CSV hoặc JSON) theo các cột *STT*, *Mã hiệu*, *Nội dung công tác*, *ĐVT*, *Khối lượng*, *Đơn giá* (hoặc *Đơn giá vật liệu / nhân công / máy*), *Thành tiền*. Dòng bắt đầu bằng `-` hoặc `+` là dòng diễn giải.
> - `T = Σ khối lượng × đơn giá` (tính lại, không lấy cột thành tiền). `GT = T × (chi phí chung + nhà tạm + công việc không xác định KL)`, `TL = (T + GT) × tỷ lệ`, `G = T + GT + TL`, `G_XD = G + VAT` (TT 11/2021/TT-BXD).
> - **Tỷ lệ** được đọc từ sheet tổng hợp G_XD trong file, hoặc truyền bằng `--rate-chung --rate-nha-tam --rate-kxd --rate-tl --vat` (đơn vị %). **Không có tỷ lệ mặc định**: thiếu tỷ lệ nào hệ thống sẽ dừng và báo tên tỷ lệ đó.
> - **Đối chiếu tự động:** công tác thiếu đơn giá hoặc khối lượng (dừng); thành tiền trong file ≠ KL × ĐG; khối lượng công tác khác tổng các dòng diễn giải; T và G_XD ghi trong file ≠ kết quả tính lại.

#### b. Chạy thử toàn bộ 7 pha bằng dữ liệu mẫu (demo):
```powershell
python run_state_graph.py --demo
```
> Tuần tự: `CAD Takeoff` $\rightarrow$ `OR-Tools Rebar Cut` $\rightarrow$ `QS G_xd` $\rightarrow$ `QA/QC Lab Link` $\rightarrow$ `Human Gate` $\rightarrow$ `CPM Schedule` $\rightarrow$ `As-Built Loop`. Ở chế độ demo, Human Gate mặc định tự phê duyệt (`auto`). Khi chạy không có `--demo`, Human Gate mặc định là `cli`: hệ thống dừng lại chờ Kỹ sư trưởng nhập `[A] Approve` / `[R] Reject` / `[S] Skip`.

#### c. Kiểm tra Solver Cắt thép OR-Tools & Bộ tính Tiến độ CPM:
```powershell
python run_state_graph.py --solver-test
```

#### c2. Chạy bộ test tự động:
```powershell
python -m unittest discover tests
```

#### d. Thử nghiệm So sánh Phiên bản Bản vẽ CAD (Incremental Diff Rev00 vs Rev01):
```powershell
python examples/run_cad_diff_demo.py
```
> Phân tích biến động hình học giữa 2 lần sửa đổi thiết kế và chỉ đưa các cấu kiện thay đổi vào hàng đợi tính toán lại.

#### e. Chạy Kiểm toán Độc lập trên Workbook 14 Sheet Master (Audit Score 100/100):
```powershell
python examples/run_pipeline.py
```

---

## 🏆 6. Chi tiết 14 Sheet Bảng tính Mẫu Hoàn thiện (Master Workbook)

Tệp Excel Master: **[`templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx`](templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx)** gồm **14 Sheet** liên thông 100% công thức động (0 số chết, 0 link gãy, Audit Score **100/100**):

| STT | Tên Sheet Excel | Vai trò Kỹ thuật | Quy chuẩn / Cơ chế Công thức |
|:---:|---|---|---|
| **01** | `TO_HOP_CAT_THEP_11M7` | Tổ hợp cắt thép thanh 11.7m bài toán 1D Cutting Stock | Google OR-Tools CP-SAT, đề-xê hao hụt **1.44%** ($< 1.5\%$) |
| **02** | `KHOI_LUONG_DAO_DAP` | Thể tích đào đắp mặt cắt ngang $V = \frac{F_1 + F_2}{2} \times L$ | Link trực tiếp sang `QS_DIEN_GIAI_CHI_TIET!J33` |
| **03** | `QS_DIEN_GIAI_CHI_TIET` | Bóc tách hình học Dài x Rộng x Cao x Số lượng x Hệ số | 100% công thức động, SUMIFS trực tiếp từ Sheet BBS cốt thép |
| **04** | `THONG_KE_THEP_CHI_TIET` | Bar Bending Schedule (BBS) 396 thanh thép chi tiết | TCVN 1651:2018, bóc tách toàn bộ mố, trụ, dầm Super-T, mặt cầu |
| **05** | `CAP_PHOI_1M3_VA_TAN_SUAT`| Cấp phối 1m³ bê tông & Ma trận 809 phép thử KCS | Tự động tính số tổ mẫu thí nghiệm bằng `=ROUNDUP(G/H,0)` |
| **06** | `PHAN_TICH_VAT_TU_WBS` | Phân tích định mức chi tiết vật liệu công tác WBS | Thông tư 12/2021/TT-BXD, chi tiết xi măng, cát, đá, sắt thép, cáp DƯL |
| **07** | `TONG_HOP_VAT_TU_TOAN_BO`| Bảng tổng hợp toàn bộ nhu cầu vật tư (BOM toàn cầu) | Công thức `=SUMIF()` động, tính hao hụt & 4 giai đoạn cấp hàng |
| **08** | `TONG_HOP_DU_TOAN_GXD` | Tổng hợp kinh phí xây dựng Thông tư 11/2021/TT-BXD | `G_xd = T + GT(7.3%) + TL(5.5%) + VAT(10%)` theo Luật XD 135/2025 |
| **09** | `THANH_TOAN_KY_PHU_LUC_03A` | Xác định giá trị khối lượng hoàn thành đề nghị thanh toán | Nghị định 99/2021/NĐ-CP, trỏ trực tiếp đơn giá và khối lượng lũy kế |
| **10** | `TIEN_DO_THI_CONG_WBS` | 36 công tác WBS, định mức nhân công TT 12, Gantt Chart | Mạng công việc CPM (Critical Path), xuất file MS Project XML/MPP |
| **11** | `HOSO_KCS_NGHIEM_THU` | Danh mục 22 Biên bản nghiệm thu KCS theo phân đoạn | Nghị định 207/2026/NĐ-CP, đồng bộ logic ngày tháng với tiến độ CPM |
| **12** | `MAU_BIEN_BAN_KCS` | Biểu mẫu nghiệm thu công việc in ấn A4 chuẩn | Chọn mã tại ô `C2` (1-22) tự động nhảy toàn bộ nội dung (thay thế Word) |
| **13** | `MAU_BB_NGHIEM_THU_VAT_LIEU`| Biểu mẫu nghiệm thu vật liệu đầu vào A4 | Chọn mã tại ô `C2` (1-16) tự động cập nhật tiêu chuẩn và chứng chỉ |
| **14** | `MAU_BB_LAY_MAU_HIEN_TRUONG`| Biểu mẫu lấy mẫu thí nghiệm hiện trường A4 | Chọn mã `C2` và ngày đúc `C3` $\rightarrow$ tự động tính ngày nén $R_7, R_{28}$ |

---

## 📁 7. Cấu trúc Cây Thư mục Dự án

```text
23HG-multiagent-system/
│
├── core/                         # 🧠 BỘ ĐIỀU PHỐI ĐỒ THỊ TRẠNG THÁI (STATE GRAPH v3.0)
│   ├── state/
│   │   ├── shared_state.py       # Pydantic/Dataclass SharedState (SSOT 8 miền dữ liệu)
│   │   └── state_bus.py          # State Bus thread-safe (RLock, read/write gateway)
│   ├── supervisor/
│   │   ├── supervisor_agent.py   # AI Supervisor (Chỉ huy trưởng ảo điều phối 7 pha)
│   │   └── base_agent.py         # Lớp cơ sở trừu tượng BaseAgent
│   ├── agents/
│   │   ├── rebar_agent.py        # Sub-Agent Cắt thép OR-Tools & Phản biện TCVN 5574
│   │   ├── asbuilt_agent.py      # Sub-Agent Vòng lặp Hiện trường & Phụ lục 03a
│   │   └── sub_agents.py         # CADAgent, QSAgent, BPTCKCSAgent (Lab Link), SchedulerAgent
│   └── gates/
│       ├── quality_gate.py       # 4 Cổng kiểm soát kỹ thuật số học xác định
│       └── human_gate.py         # Human-in-the-loop Gate (Ký duyệt Kỹ sư trưởng)
│
├── tools/                        # ⚙️ CÔNG CỤ TÍNH TOÁN XÁC ĐỊNH (PURE PYTHON, ZERO LLM)
│   ├── cutting_stock_solver.py   # Solver cắt thép 1D (Column Generation GLOP + CP-SAT, tách nhóm Ø + mác thép)
│   ├── bbs_loader.py             # Đọc BBS thật từ Excel / CSV / JSON (cả sheet INPUT của RebarCut), cột nối thép
│   ├── rebarcut_export.py        # Xuất kết quả cắt thép theo bố cục RebarCut Pro Excel (.xlsx)
│   ├── schedule_loader.py        # Đọc tiến độ thật từ MS Project XML / Excel / CSV, đối chiếu ngày trong file
│   ├── qs_loader.py              # Đọc bảng QS thật, tỷ lệ chi phí, tính G_XD TT 11/2021 + đối chiếu
│   ├── qs_export.py              # Xuất bảng tổng hợp G_XD + chi tiết công tác (.xlsx)
│   ├── excel_eval.py             # Tính công thức Excel chưa có kết quả lưu sẵn (Pure Python)
│   ├── cpm_calculator.py         # Bộ tính CPM: FS/SS/FF/SF + lag, lịch nghỉ, Forward/Backward Pass
│   └── cad_diff_engine.py        # Động cơ so sánh phiên bản bản vẽ CAD Rev00 vs Rev01
│
├── schemas/                      # 📋 ĐẶC TẢ SCHEMA DỮ LIỆU CHUYÊN NGÀNH
│   ├── site_log_schema.py        # Schema Nhật ký hiện trường & Khối lượng hoàn công As-Built
│   └── lab_result_schema.py      # Schema Phiếu thí nghiệm phòng LAS-XD (R7/R28, kéo thép, PDA)
│
├── aec_core/                     # 🔍 BỘ CÔNG CỤ KIỂM TOÁN VÀ XÁC THỰC ĐỘC LẬP
│   ├── audit_verifier.py         # AECAuditVerifier: Quét 1,234 công thức, 0 số chết, 100/100
│   └── project_state.py          # Trình quản lý trạng thái dự án cơ sở
│
├── agents/                       # 🤖 CỤM TÁC TỬ THU NHẬN & HỢP NHẤT DỮ LIỆU ĐA PHƯƠNG THỨC
│   ├── aec_cad_extractor.py      # Bộ trích xuất bản vẽ CAD (COM Interop, TCVN3/VNI decoder)
│   ├── aec_office_extractor.py   # Bộ đọc dữ liệu bảng tính Excel dự toán / BoQ
│   ├── aec_markdown_ingestor.py  # Bộ đọc hồ sơ thuyết minh kỹ thuật Markdown/Text
│   ├── aec_data_aggregator.py    # Bộ hợp nhất đa luồng vào Single Source of Truth
│   └── PROJECT_STATE.json        # Dữ liệu trạng thái dự án Cầu Km19+529.080
│
├── workflows/                    # 📚 QUY TRÌNH KỸ THUẬT & TIÊU CHUẨN THI CÔNG
│   ├── 01_QUY_TRINH_BOC_TACH_HINH_HOC_TAKEOFF.md
│   ├── 02_QUY_TRINH_TO_HOP_CAT_THEP_1D.md
│   ├── 03_QUY_TRINH_DU_TOAN_GXD_TT11.md
│   ├── 04_QUY_TRINH_THANH_TOAN_03A_ND99.md
│   ├── 05_QUY_TRINH_TIEN_DO_WBS_CPM_MS_PROJECT.md
│   ├── 06_QUY_TRINH_HO_SO_KCS_WORD_ND207.md
│   ├── 07_QUY_TRINH_THUYET_MINH_BIEN_PHAP_HUGGINGFACE.md
│   ├── 08_QUY_TRINH_PHAN_TICH_TONG_HOP_VAT_TU_DINH_MUC.md
│   ├── 09_QUY_TRINH_THONG_KE_THEP_BBS_VA_TAN_SUAT_THI_NGHIEM.md
│   ├── 10_QUY_TRINH_THU_NHAN_VA_HOP_NHAT_DU_LIEU_DA_PHUONG_THUC.md
│   ├── 11_QUY_TRINH_VALIDATION_KIEM_TRA_CHEO.md
│   ├── 12_SO_DO_DIEU_PHOI_MULTI_AGENT_TOAN_HE_THONG.md
│   └── 13_KIEN_TRUC_STATE_GRAPH_V3_SUPERVISOR_PATTERN.md
│
├── templates/                    # 📦 SẢN PHẨM MẪU SỐ HÓA HOÀN THIỆN
│   ├── Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx  # Workbook Master 14 Sheet liên kết động
│   ├── Tien_Do_Thi_Cong_Cau_Km19+529.080.xml      # Tiến độ MS Project XML (CPM, 36 công tác)
│   ├── Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp      # File Microsoft Project Native binary
│   ├── Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md # Thuyết minh BPTC 8 chương TCVN
│   └── BAO_CAO_THAM_TRA_AEC_AUDIT.md              # Báo cáo thẩm tra độc lập Audit 100/100
│
├── examples/                     # 🚀 SCRIPT THỰC THI & CHẠY THỬ NGHIỆM
│   ├── run_pipeline.py                            # Runner kiểm tra toàn diện 14 Sheet Master
│   ├── run_cad_diff_demo.py                       # Demo so sánh bản vẽ CAD Rev00 vs Rev01
│   ├── run_data_ingestion_pipeline.py             # Pipeline thu nhận đa phương thức
│   ├── update_full_cross_linked_workbook.py       # Script tái tạo 14 sheet liên kết động
│   └── add_rebar_bbs_and_mix_sheets.py            # Trích xuất BBS & Tần suất thí nghiệm
│
├── tests/                        # 🧪 Test tự động (python -m unittest discover tests)
├── run_state_graph.py            # 🌟 ENTRY POINT MỚI: State Graph & Supervisor Runner v3.0
├── requirements.txt              # Danh mục thư viện phụ thuộc (ortools, openpyxl, pandas...)
├── pyproject.toml                # Cấu hình đóng gói hệ thống chuẩn PEP 621
├── LICENSE                       # Giấy phép phần mềm mã nguồn mở MIT
└── README.md                     # Tài liệu hướng dẫn chính thức của dự án
```

---

## ⚖️ 8. Giấy phép Bản quyền (License)

Dự án được phân phối dưới giấy phép mã nguồn mở **[MIT License](LICENSE)**.

*Hệ thống được nghiên cứu, phát triển và đóng gói bởi Cộng đồng Kỹ sư Xây dựng Số hóa Việt Nam.*
