# 🏗️ 23HG-AEC-MultiAgent-System
### Nền tảng Đa tác tử Kỹ thuật số hóa & Quản lý Dự án Xây dựng (Closed-Loop ConTech System)
**Kỹ sư Trưởng Số hóa: Bóc tách Hình học • Cắt thép 1D OR-Tools • Dự toán G_xd • Thanh toán 03a • Tiến độ CPM MS Project • Ca xe Ca máy & Dầu Diezel • Phân quyền Hub & Spoke 5 Gói • Hồ sơ KCS & BPTC • Vòng lặp Hiện trường As-Built**

> [!IMPORTANT]
> **THÔNG CÁO BẢN QUYỀN VÀ NGUỒN GỐC CHÍNH THỨC (OFFICIAL AUTHORSHIP & COPYRIGHT NOTICE)**
> - **Tác giả sáng lập & Duy trì**: **Nguyễn Bảo Tú** ([@baotuhg](https://github.com/baotuhg))
> - **Kho lưu trữ chính thức duy nhất**: [https://github.com/baotuhg/23HG-multiagent-system](https://github.com/baotuhg/23HG-multiagent-system)
> - **Cảnh báo bản quyền & Đạo nhái**: Toàn bộ kiến trúc Đa tác tử AEC, thuật toán tối ưu hóa cắt thép 1D Column Generation (Google OR-Tools CP-SAT), công cụ giải mã bản vẽ CAD, engine thanh toán Phụ lục 03a và bộ biểu mẫu Master Excel liên kết động thuộc bản quyền trí tuệ của **Nguyễn Bảo Tú (@baotuhg)**. Mọi hành vi clone/tải về re-upload dưới tên tổ chức/cá nhân khác, xóa lịch sử commit (commit history), nhận vơ sản phẩm mà không Fork chính thức từ repo gốc đều là hành vi xâm phạm quyền tác giả và bị xử lý theo quy định bảo vệ bản quyền phần mềm (DMCA Takedown).

---

[![CI](https://github.com/baotuhg/23HG-multiagent-system/actions/workflows/ci.yml/badge.svg)](https://github.com/baotuhg/23HG-multiagent-system/actions/workflows/ci.yml)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Standards: TCVN & BXD](https://img.shields.io/badge/Standards-TCVN%20%7C%20Lu%E1%BA%ADt%20XD%20135%2F2025-brightgreen)](workflows/)
[![Optimization: Google OR-Tools](https://img.shields.io/badge/Optimization-OR--Tools%20Column%20Generation%20%2B%20CP--SAT-blue)](tools/cutting_stock_solver.py)
[![Fleet Scheduling: Vincons & TT13](https://img.shields.io/badge/Fleet%20Scheduling-Vincons%20%7C%20TT13%2F2021-orange)](tools/equipment_fleet_scheduler.py)
[![Packaging: Hub & Spoke 5 Packages](https://img.shields.io/badge/Packaging-Hub%20%26%20Spoke%20Role--Based-purple)](workflows/15_QUY_TRINH_DONG_GOI_HUB_AND_SPOKE_PHAN_QUYEN_THUC_CHIEN.md)
[![Zero Dead Numbers](https://img.shields.io/badge/Math-100%25%20Dynamic%20Formulas-red.svg)](templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx)

---

## 📖 1. Giới thiệu Tổng quan (Overview)

**23HG-AEC-MultiAgent-System** là nền tảng ConTech mã nguồn mở chuyên sâu cho quản lý kỹ thuật, dự toán, kế hoạch cơ giới và điều hành thi công xây dựng tại Việt Nam. Hệ thống chuyển đổi mô hình tự động hóa từ chuỗi tuyến tính (Linear Pipeline) sang **Hệ thống Đa tác tử Khép kín (Closed-Loop Multi-Agent System)** vận hành trên **Đồ thị Trạng thái (State Graph v3.0)** với sự điều phối tập trung của **AI Supervisor (Chỉ huy trưởng ảo)**.

Nền tảng tích hợp trọn gói chu trình vòng đời dự án: từ bóc tách bản vẽ CAD/BIM, giải bài toán tổ hợp cắt thép 1D Column Generation, lập dự toán chi phí `G_xd` và hồ sơ thanh toán Phụ lục 03a, quản lý tiến độ CPM đường găng, **điều phối ca xe ca máy & định mức cấp phát nhiên liệu dầu Diezel**, đến xuất hồ sơ chất lượng KCS in ấn A4 chuẩn chỉ và **đóng gói phân quyền thực chiến theo mô hình Hub & Spoke 5 gói vệ tinh**.

### Cơ sở Pháp lý & Tiêu chuẩn Kỹ thuật:
- **Luật Xây dựng số 135/2025/QH15** & **Nghị định số 207/2026/NĐ-CP**: Quản lý chất lượng thi công, nhật ký thi công, giám sát và nghiệm thu KCS.
- **Nghị định số 254/2025/NĐ-CP**: Quản lý, thanh toán, quyết toán dự án vốn đầu tư công (Bảng thanh toán khối lượng hoàn thành Phụ lục 03a).
- **Thông tư số 36/2026/TT-BXD**: Phương pháp xác định và quản lý chi phí đầu tư xây dựng (Dự toán chi phí xây dựng `G_xd = T + GT + TL + VAT 10%`).
- **Thông tư số 38/2026/TT-BXD**: Định mức dự toán xây dựng công trình, định mức hao phí vật tư (xi măng, cát, đá, cốt thép, cáp DƯL), nhân công và ca máy.
- **Thông tư số 37/2026/TT-BXD**: Phương pháp xác định các chỉ tiêu kinh tế kỹ thuật và đo bóc khối lượng xây dựng (Đơn giá nhân công, giá ca máy và thiết bị thi công).
- **Tiêu chuẩn Định mức Cơ giới Thực chiến**: Bộ định mức ca máy, năng suất thiết bị thi công và định mức tiêu hao nhiên liệu dầu Diezel theo chuẩn Vincons / Vinhomes và các Tổng công ty xây dựng hạ tầng lớn.
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
                  | - Kiểm soát các Cổng Chất lượng (Quality Gates) |
                  +-----------------------+-------------------------+
                                          |
        +---------------------------------+---------------------------------+
        |                                 |                                 |
        v                                 v                                 v
+------------------+             +------------------+             +------------------+
| CAD/BIM PARSER   |             | KỸ THUẬT BPTC    |             | QS DỰ TOÁN G_XD  |
| (Trắc đạc CAD)   |             | & KCS LAB LINK   |             | & PHỤ LỤC 03A    |
+--------+---------+             +--------+---------+             +--------+---------+
| * Đọc DWG/DXF    |             | * Kiểm soát BPTC |             | * Đơn giá TT 38/2026  |
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
| * Nhận diện Đường găng (Critical Path Chain) & Xuất MS Project XML/MPP    |
+------------------------------------+--------------------------------------+
                                     |
                                     v
+---------------------------------------------------------------------------+
|             ĐIỀU PHỐI CA XE, CA MÁY & NHIÊN LIỆU DẦU DIEZEL               |
| * Tác tử Equipment Fleet Engine: Tính ca máy từ định mức Vincons / TT 37/2026   |
| * Phân bổ máy theo ngày/tuần, biểu đồ phụ tải & kế hoạch cấp phát dầu (L) |
+------------------------------------+--------------------------------------+
                                     |
                                     v
                  +-------------------------------------------------+
                  |       SHARED STATE BUS (SINGLE SOURCE OF TRUTH) |
                  |  9 Miền: Meta • CAD • Rebar • QS • QAQC • CPM   |
                  |     Fleet • As-Built • Human Approvals (RLock)  |
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
                  +------------------+------------------------------+
                                     |
                                     v
                  +-------------------------------------------------+
                  |    ĐÓNG GÓI PHÂN QUYỀN THỰC CHIẾN HUB & SPOKE   |
                  |  Gói A: Cơ giới & Dầu  | Gói B: Xưởng Thép CNC  |
                  |  Gói C: Hiện trường KCS| Gói D: QS & Dự toán    |
                  |  Gói E: Executive Hub  | DISPATCH_MANIFEST.json |
                  +-------------------------------------------------+
```

---

## 📥 3. Chuẩn I/O (Input/Output Specifications)

| Thành phần | Định dạng Đầu vào (Input) | Định dạng Đầu ra (Output) | Công cụ & Đặc tả Kỹ thuật |
|---|---|---|---|
| **Bóc tách CAD/BIM** | Bản vẽ CAD `.dwg`, `.dxf`, mô hình `.ifc` | Bảng khối lượng hình học Bê tông, Ván khuôn, Đào đắp | `AutoCAD COM Interop`, `ezdxf`, Shoelace & Average-End-Area |
| **Gia công Cốt thép** | File BBS thật `.xlsx` / `.csv` / `.json` (`--bbs`) | Phiếu cắt từng phương án cây 11.7m (CSV), số cây, cận dưới, đề-xê, mẩu thừa tận dụng | OR-Tools Column Generation (GLOP) + CP-SAT, tách nhóm Ø + mác thép, tính lưỡi cắt 3mm |
| **Dự toán Chi phí** | Khối lượng trích xuất, Đơn giá định mức | Bảng dự toán tổng hợp chi phí xây dựng `G_xd` | Excel 100% công thức động (`G_xd = T + GT + TL + VAT 10%`) |
| **Thanh toán Hợp đồng**| Khối lượng thiết kế vs Khối lượng hoàn công | Bảng xác định khối lượng hoàn thành Phụ lục 03a | Nghị định 254/2025/NĐ-CP, tính phát sinh tự động |
| **Quản lý Tiến độ** | File tiến độ thật `MS Project .xml` / `.xlsx` / `.csv` (`--schedule`) | Bảng CPM (ES/EF/LS/LF, dự trữ, đường găng, ngày lịch) CSV; cảnh báo ngày trong file vi phạm quan hệ logic | CPM với quan hệ FS/SS/FF/SF + lag, lịch nghỉ (Chủ nhật, ngày lễ) |
| **Ca xe & Dầu Diezel** | Tiến độ CPM (`.xml`/`.xlsx`), Khối lượng hình học, Định mức ca máy | Bảng tiến độ ca máy theo ngày/tuần (`.xlsx` + `.xml`), Biểu đồ phụ tải, Kế hoạch cấp dầu Diezel (Lít) | `EquipmentFleetScheduler`, định mức Vincons / TT 37/2026, tính ca/ngày và nhiên liệu chi tiết |
| **Quản lý Chất lượng**| Phiếu thí nghiệm nén R7/R28, kéo thép, PDA | 22 Biên bản nghiệm thu KCS in ấn A4 chuẩn | Excel A4 Form (`MAU_BIEN_BAN_KCS`, thay thế hoàn toàn Word) |
| **Biện pháp Thi công** | Yêu cầu KTXD, điều kiện địa chất, thủy văn | Thuyết minh BPTC 8 chương TCVN | Markdown mẫu viết sẵn (Cầu Km19+529.080) — **RAG Hugging Face chưa triển khai**, xem lộ trình |
| **Đối soát Hiện trường**| Nhật ký thi công hàng ngày `DailySiteLog` | Báo cáo chênh lệch tiến độ & Chi phí phát sinh | As-Built Closed Loop, tự động cập nhật mạng CPM |
| **Đóng gói Hub & Spoke**| Toàn bộ dữ liệu & sản phẩm đầu ra dự án | 5 Gói vệ tinh độc lập (`GOI_A` đến `GOI_E`) kèm `DISPATCH_MANIFEST.json` | `AECPackageDispatcher`, phân quyền theo vai trò (RBAC), bảo mật giá thầu, chống khóa file |

---

## 🗺️ 4. Lộ trình Phát triển (Roadmap 3 Phase)

```text
  Phase 1 (v1.x) [100% HOÀN THÀNH]
  ├── Tự động hóa tác vụ kỹ thuật cốt lõi (Core Engines)
  ├── Bóc tách hình học Takeoff 100% công thức động (0 số chết)
  ├── Cắt thép 1D Cutting Stock: tối ưu số cây theo từng Ø + mác thép, có cận dưới chứng minh
  ├── Bảng phân tích định mức & Tổng hợp vật tư toàn cầu BOM
  ├── Dự toán G_xd Thông tư 36/2026 & Thanh toán kỳ Phụ lục 03a
  └── Bộ 14 Sheet Master Excel đạt 100/100 điểm Audit Verifier

  Phase 2 (v2.x) [100% HOÀN THÀNH - State Graph v3.0]
  ├── Tái cấu trúc State Graph & AI Supervisor điều phối tập trung
  ├── Cách ly 100% toán học số học khỏi LLM (Google OR-Tools, Pure Python CPM)
  ├── Phản biện chéo Inter-Agent Negotiation: Thẩm tra vị trí nối thép TCVN 5574
  ├── Vòng lặp đối soát hiện trường (Site Feedback & As-Built Loop)
  ├── QA/QC Lab Link: Tích hợp kết quả nén R7/R28, siêu âm cọc, PDA vào KCS
  └── Cổng phê duyệt Kỹ sư trưởng Human-in-the-loop (Awaiting Approval)

  Phase 3 (v3.x) [ĐANG TRIỂN KHAI & MỞ RỘNG]
  ├── [100% HOÀN THÀNH] Tác tử Điều phối Ca xe, Ca máy & Kế hoạch Nhiên liệu Dầu Diezel (AEC Equipment & Fleet Engine)
  ├── [100% HOÀN THÀNH] Mô hình Đóng gói Phân quyền Thực chiến Hub & Spoke 5 Gói Vệ tinh (Role-Based Dispatcher)
  ├── [100% HOÀN THÀNH] Mở rộng bộ hồ sơ mẫu thực chiến: Cống hộp Tuyến A5 & Cụm B9 San lấp Olympic Thường Tín
  ├── Động cơ so sánh phiên bản CAD/BIM Versioning (Incremental Diff Rev00 vs Rev01) — đã so sánh được dữ liệu cấu kiện;
  │   còn thiếu bước tự bóc khối lượng từ bản vẽ (hiện đọc được diện tích đa tuyến khép kín trong DXF)
  ├── RAG thật cho Thuyết minh BPTC (embedding + truy xuất TCVN + sinh nội dung) — hiện là bản mẫu viết sẵn
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
> Hệ thống hỗ trợ thực thi độc lập từng pha chuyên sâu hoặc chạy toàn diện khép kín **8 pha**: CAD Takeoff $\rightarrow$ OR-Tools Rebar Cut $\rightarrow$ QS G_xd $\rightarrow$ QA/QC Lab Link $\rightarrow$ Human Gate $\rightarrow$ CPM Schedule $\rightarrow$ **Fleet Dispatch** $\rightarrow$ As-Built Loop.
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
> - **Giới hạn cho tổ cắt:** `--max-pieces-per-bar 4 --max-marks-per-bar 2`. **Cắt đầu cây:** `--end-trim-mm 50`. **Lưỡi cắt:** `--kerf-mm 3`. **Đầu thừa** được phân loại *Tái sử dụng* (≥ 100D, đổi bằng `--reuse-xd`), *Đầu thừa ngắn* (≥ 20D) hoặc *Phế*.
> - **Phương án nối thép tận dụng đầu thừa** (`--splice`), đưa phép nối vào ngay mô hình tối ưu OR-Tools, chặt hơn PA4 của RebarCut.
> - **Xuất theo bố cục RebarCut Pro Excel:** `--rebarcut-out ket_qua.xlsx`, gồm các sheet INPUT, SO_SANH, PA_TOI_UU, PA_NOI, MOI_NOI, REMAIN, CHI_TIET.

#### a2. Tính tiến độ CPM từ file tiến độ thật:
```powershell
python run_state_graph.py --phase schedule --schedule "TienDo.xml" --non-working-days cn --holidays 2027-02-05:2027-02-12 --schedule-out tien_do_cpm.csv
```
> - Đọc được **MS Project XML** (File → Save As → XML trong MS Project), Excel, CSV hoặc JSON. Cột nhận diện theo tiêu đề: *Mã WBS*, *Danh mục công tác*, *Thời gian (ngày)*, *Quan hệ logic* (vd `1.2FS; 1.3SS+3d`), *Ngày bắt đầu/hoàn thành* (tùy chọn, dùng để đối chiếu).
> - Hỗ trợ quan hệ **FS / SS / FF / SF** có độ trễ (âm hoặc dương), ngày nghỉ trong tuần (`--non-working-days t7,cn`) và ngày lễ (`--holidays`). Ngày khởi công lấy từ file, hoặc chỉ định bằng `--start-date`.
> - Liên kết tới công việc không tồn tại, vòng lặp logic, thời lượng sai: **dừng và liệt kê từng lỗi**.

#### a3. Tính dự toán G_XD từ bảng QS thật:
```powershell
python run_state_graph.py --phase qs --qs "Du_toan.xlsx" --qs-out du_toan_gxd.xlsx
```
> - Đọc bảng QS / BOQ (Excel, CSV hoặc JSON) theo các cột *STT*, *Mã hiệu*, *Nội dung công tác*, *ĐVT*, *Khối lượng*, *Đơn giá* (hoặc *Đơn giá vật liệu / nhân công / máy*), *Thành tiền*.
> - `T = Σ khối lượng × đơn giá`. `GT = T × (chi phí chung + nhà tạm + công việc không xác định KL)`, `TL = (T + GT) × tỷ lệ`, `G = T + GT + TL`, `G_XD = G + VAT` (TT 36/2026/TT-BXD).
> - **Tỷ lệ** được đọc từ sheet tổng hợp G_XD trong file, hoặc truyền bằng `--rate-chung --rate-nha-tam --rate-kxd --rate-tl --vat` (đơn vị %).

#### a4. Lập Mẫu 03a — giá trị khối lượng hoàn thành đề nghị thanh toán (NĐ 254/2025):
```powershell
python run_state_graph.py --phase payment --qs "Du_toan.xlsx" --progress "KL_ky_01.xlsx" --price-basis direct --advance-recovery-pct 20 --retention-pct 5 --period 01 --payment-out Mau_03a_ky01.xlsx
```
> - **Hợp đồng** (khối lượng, đơn giá) lấy từ bảng QS. **Khối lượng thực hiện** lấy từ file `--progress`.
> - Tự động tính khấu trừ thu hồi tạm ứng, khấu trừ bảo hành, phát hiện và cảnh báo khối lượng vượt hợp đồng tại sheet `PHAT_SINH_CANH_BAO`.

#### a4b. Đánh giá phiếu thí nghiệm & điểm dừng kỹ thuật (Hold Point) từ file thật:
```powershell
python run_state_graph.py --phase qaqc --lab "Phieu_thi_nghiem.xlsx" --lab-out danh_gia_thi_nghiem.xlsx
```
> - Đọc phiếu thí nghiệm (Excel, CSV hoặc JSON; mẫu cột: `templates/Phieu_thi_nghiem_mau.csv`): nén bê tông R7/R28, kéo thép, PDA, siêu âm cọc, độ sụt...
> - Mỗi phiếu được đánh giá **ĐẠT / KHÔNG ĐẠT / CHỜ**; mỗi biên bản nghiệm thu (BBNT) liên kết được xếp **GIẢI TỎA / CHỜ / CHẶN**. Có phiếu KHÔNG ĐẠT thì phase dừng.
> - Kiểm toán Excel master (GATE-4, 100/100) chỉ chạy khi có `--excel`; không có thì Human Gate ghi rõ "chưa kiểm toán".

#### a4c. Kiểm tra file đầu vào trước khi chạy (không tính toán, không ghi file):
```powershell
python run_state_graph.py --check-inputs --bbs "BBS.xlsx" --qs "Du_toan.xlsx" --schedule "TienDo.xml" --lab "Phieu_thi_nghiem.xlsx"
```
> Đọc từng file bằng đúng bộ đọc của hệ thống, liệt kê dòng lỗi, trả mã thoát 1 nếu có lỗi — dùng được trong script / CI.

#### a5. Điều phối Ca xe, Ca máy & Kế hoạch Nhiên liệu Dầu Diezel:
```powershell
python run_state_graph.py --phase fleet --fleet-out ca_xe_ca_may.xlsx --shifts 2
```
> - Tự động bóc tách ca máy từ khối lượng công tác và tiến độ CPM theo định mức ca máy Vincons / Thông tư 37/2026/TT-BXD.
> - Xuất bảng tiến độ ca máy chi tiết theo ngày/tuần, biểu đồ phụ tải máy móc và bảng dự trù cấp phát nhiên liệu dầu Diezel (Lít) theo từng ca làm việc.

#### a6. Đóng gói & Phân quyền Công trường Hub & Spoke (5 Gói vệ tinh):
```powershell
# Chạy trực tiếp qua module Package Dispatcher:
python -m tools.package_dispatcher --source ./examples/HO_SO_CONG_HOP_TUYEN_A5 --target ./HO_SO_HUB_AND_SPOKE

# Hoặc kích hoạt qua State Graph CLI:
python run_state_graph.py --phase dispatch --dispatch-out ./HO_SO_HUB_AND_SPOKE
```
> - Tự động quét và phân loại toàn bộ hồ sơ dự án thành 5 gói vệ tinh độc lập theo vai trò (`GOI_A` đến `GOI_E`).
> - Tạo tệp kê khai `DISPATCH_MANIFEST.json` ghi nhận định danh nguồn, ngày giờ xuất xưởng và mã băm MD5 xác thực tính toàn vẹn của từng tệp tin.

#### b. Chạy thử toàn bộ 8 pha bằng dữ liệu mẫu (demo):
```powershell
python run_state_graph.py --demo
```
> Vòng lặp tự động tuần tự qua 8 pha: `CAD Takeoff` $\rightarrow$ `OR-Tools Rebar Cut` $\rightarrow$ `QS G_xd` $\rightarrow$ `QA/QC Lab Link` $\rightarrow$ `Human Gate` $\rightarrow$ `CPM Schedule` $\rightarrow$ `Fleet Dispatch` $\rightarrow$ `As-Built Loop`. Ở chế độ demo, Human Gate mặc định tự phê duyệt (`auto`). Khi chạy không có `--demo`, Human Gate dừng lại chờ Kỹ sư trưởng phê duyệt `[A] Approve` / `[R] Reject`.

#### c. Kiểm tra Solver Cắt thép OR-Tools & Bộ tính Tiến độ CPM:
```powershell
python run_state_graph.py --solver-test
```

> **Trạng thái runtime** (khôi phục sau khi dừng giữa chừng) được ghi vào `.aec_state/RUNTIME_STATE.json` — thư mục này không đưa vào git. Đổi vị trí bằng biến môi trường `AEC_STATE_DIR`.

#### c2. Chạy bộ kiểm thử tự động toàn diện (Unit Tests):
```powershell
# Chạy toàn bộ test suite
python -m unittest discover tests

# Hoặc kiểm tra riêng module ca máy và phân gói Hub & Spoke
python -m unittest tests/test_equipment_fleet_scheduler.py tests/test_package_dispatcher.py
```

#### d. Thử nghiệm So sánh Phiên bản Bản vẽ CAD (Incremental Diff Rev00 vs Rev01):
```powershell
python examples/run_cad_diff_demo.py
```

#### e. Chạy Kiểm toán Độc lập trên Workbook 14 Sheet Master (Audit Score 100/100):
```powershell
python examples/run_pipeline.py
```

---

## 🏆 6. Chuẩn Đóng Gói Hồ Sơ: Từ Macro/Micro đến Mô hình Thực chiến "Hub & Spoke" Phân Quyền 5 Gói

Hệ thống thiết lập **3 Tầng Đóng Gói Hồ Sơ Linh Hoạt**, đáp ứng trọn vẹn từ yêu cầu thẩm định vĩ mô của Chủ đầu tư đến nhu cầu thi công, gia công thực chiến tại hiện trường:

```
                            +-------------------------------------+
                            |        CENTRAL STATE BUS / DB       |
                            |       "Single Source of Truth"      |
                            +------------------+------------------+
                                               |
             +---------------------------------+---------------------------------+
             |                                 |                                 |
             v                                 v                                 v
+------------------------+        +------------------------+        +------------------------+
| GÓI A: CƠ GIỚI & DẦU   |        | GÓI B: XƯỞNG CỐT THÉP  |        | GÓI C: HIỆN TRƯỜNG KCS |
| • Quản lý máy móc      |        | • Quản đốc xưởng       |        | • Cán bộ QA/QC, TVGS   |
| • Cấp phát Dầu Diezel  |        | • Lệnh cắt CNC (.csv)  |        | • BBNT Word A4 (.docx) |
| • Biểu đồ phụ tải máy  |        | • Sơ đồ 1D CSP 11.7m   |        | • Ma trận nén R7/R28   |
+------------------------+        +------------------------+        +------------------------+
             |                                 |                                 |
             +---------------------------------+---------------------------------+
                                               |
                                               v
                            +-------------------------------------+
                            | GÓI D: QS, DỰ TOÁN & THANH TOÁN     |
                            | • Kỹ sư QS, Kế toán dự án           |
                            | • Bảo mật đơn giá dự toán G_XD      |
                            | • Bảng thanh toán Phụ lục 03a       |
                            +------------------+------------------+
                                               |
                                               v
                            +-------------------------------------+
                            | GÓI E: EXECUTIVE DASHBOARD (HUB)    |
                            | • Giám đốc Dự án, Ban QLDA, CĐT     |
                            | • Tiến độ CPM, Báo cáo Audit 100/100|
                            +-------------------------------------+
```

### 1. TẦNG 1: VĨ MÔ / MASTER 14 SHEET LIÊN KẾT ĐỘNG TOÀN DIỆN (Macro Tier)
*Mục đích: Phục vụ Hội đồng Thẩm định, Lưu trữ Pháp lý, Kiểm toán Độc lập.*
- **Tệp Excel Master:** [`templates/Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2_Km14+363.65.xlsx`](templates/Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2_Km14+363.65.xlsx) gồm **14 Sheet** liên thông 100% công thức động (0 số chết, 0 link gãy, Audit Score **100/100**).
- **Tệp Tiến độ MS Project:** [`templates/Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml`](templates/Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml) & [`.mpp`](templates/Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.mpp).
- **Báo cáo Thẩm tra Độc lập:** [`templates/BAO_CAO_THAM_TRA_AEC_AUDIT_KHAI_HOANG_2.md`](templates/BAO_CAO_THAM_TRA_AEC_AUDIT_KHAI_HOANG_2.md).

### 2. TẦNG 2: VI MÔ / 14 BỘ HỒ SƠ CHUYÊN SÂU ĐỘC LẬP (Micro Tier 1-to-1)
*Mục đích: Cho phép chuyên viên kỹ thuật tra cứu chuyên sâu từng hạng mục mà không làm gãy công thức `#REF!`.*
- Mỗi Sheet trong Master được tách thành 1 file độc lập (`01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`, `02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx`... đến `14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx`), chứa sẵn các tab dữ liệu nội bộ (`DATA_CONG_TAC`, `DATA_VAT_LIEU`, `DATA_NEN_MAU`).
- Lệnh tự động dựng: `python examples/build_14_micro_standalone_dossiers.py`.

### 3. TẦNG 3: MÔ HÌNH THỰC CHIẾN "HUB & SPOKE" PHÂN QUYỀN 5 GÓI (Role-Based Field Dispatching)
*Mục đích: KHUYẾN NGHỊ HÀNG ĐẦU CHO ĐIỀU HÀNH CÔNG TRƯỜNG THỰC TẾ.*
Khắc phục triệt để 3 nhược điểm lớn khi dùng 1 file 14 sheet trên hiện trường:
1. **Xung đột file khóa (Read-Only Lock):** Nhiều bộ phận (QS, Đội xe, Thợ sắt, QA/QC) cùng mở và chỉnh sửa không bị tranh chấp tệp.
2. **Bảo mật tuyệt đối dữ liệu tài chính:** Thợ sắt, lái máy và thầu phụ chỉ nhận đúng thông số kỹ thuật, **không thấy đơn giá thầu, chi phí gián tiếp hay lợi nhuận định mức** của Tổng thầu.
3. **Tối ưu hóa thiết bị di động:** Dung lượng file nhẹ, mở tức thì trên điện thoại ngoài công trường, không lo lag hay gãy công thức.

#### Bảng Phân Quyền & Bàn Giao 5 Gói Vệ Tinh:
| Gói hồ sơ | Đối tượng sử dụng | Quyền hạn dữ liệu | Tệp bàn giao chính |
| :--- | :--- | :--- | :--- |
| **Gói A: Cơ giới & Dầu** | Đội trưởng xe máy, Thủ kho dầu, Lái xe máy | Xem tiến độ ca máy, phụ tải, ký nhận dầu. **Không xem đơn giá tiền.** | `260920_TDTC_CaXe_CaMay_...xlsx`<br>`260920_Tien_Do_CaMay_...xml` |
| **Gói B: Xưởng cốt thép** | Quản đốc xưởng, Thợ uốn cắt, Nhà cấp thép | Xem BBS, sơ đồ cắt thép, file CSV nạp máy CNC. **Không xem đơn giá tiền.** | `01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`<br>`01_Phieu_Cat_Thep_Xuong_CNC_...csv`<br>`04_Thong_Ke_Thep_Chi_Tiet_BBS_...xlsx` |
| **Gói C: Hiện trường KCS** | Kỹ sư QA/QC, Tư vấn giám sát, Thí nghiệm | Xem ngày nghiệm thu, nén mẫu, in ấn biên bản. **Không xem đơn giá tiền.** | `11_Danh_Muc_KCS_Bien_Ban_Nghiem_Thu.xlsx`<br>`Ho_So_Bien_Ban_Nghiem_Thu_KCS_...docx`<br>`Mau_A4_Bien_Ban_...xlsx` |
| **Gói D: QS & Dự toán** | Kỹ sư QS, Trưởng phòng Kế hoạch, Kế toán | **Toàn quyền xem đơn giá, doanh thu, thanh toán.** | `03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx`<br>`08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx`<br>`09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx` |
| **Gói E: Executive Hub** | Giám đốc Dự án, Ban Giám đốc, Chủ đầu tư | Xem KPI tổng thể, tiến độ đường găng CPM, báo cáo thẩm tra Audit 100/100. | `02_Tien_Do_Thi_Cong_Master_...xml`<br>`03_BAO_CAO_THAM_TRA_AEC_AUDIT_...md` |

> Chi tiết quy trình đóng gói: Xem [`workflows/15_QUY_TRINH_DONG_GOI_HUB_AND_SPOKE_PHAN_QUYEN_THUC_CHIEN.md`](workflows/15_QUY_TRINH_DONG_GOI_HUB_AND_SPOKE_PHAN_QUYEN_THUC_CHIEN.md).  
> Bộ hồ sơ mẫu thực chiến đã đóng gói hoàn chỉnh: Xem thư mục [`examples/HO_SO_CONG_HOP_TUYEN_A5/HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5/`](examples/HO_SO_CONG_HOP_TUYEN_A5/HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5/).

---

## 📁 7. Cấu trúc Cây Thư mục Dự án

```text
23HG-multiagent-system/
│
├── .github/                      # 🤖 CI/CD WORKFLOWS
│   └── workflows/ci.yml          # GitHub Actions (Python 3.10, 3.11, 3.12 trên Ubuntu & Windows)
│
├── core/                         # 🧠 BỘ ĐIỀU PHỐI ĐỒ THỊ TRẠNG THÁI (STATE GRAPH v3.0)
│   ├── state/
│   │   ├── shared_state.py       # Pydantic/Dataclass SharedState (SSOT 9 miền, bổ sung FLEET_DISPATCH)
│   │   └── state_bus.py          # State Bus thread-safe (RLock, read/write gateway)
│   ├── supervisor/
│   │   ├── supervisor_agent.py   # AI Supervisor (Chỉ huy trưởng ảo điều phối các pha)
│   │   └── base_agent.py         # Lớp cơ sở trừu tượng BaseAgent
│   ├── agents/
│   │   ├── rebar_agent.py        # Sub-Agent Cắt thép OR-Tools & Phản biện TCVN 5574
│   │   ├── asbuilt_agent.py      # Sub-Agent Vòng lặp Hiện trường (demo)
│   │   ├── payment_agent.py      # Sub-Agent Thanh toán Mẫu 03a từ bảng QS + khối lượng thực hiện
│   │   └── sub_agents.py         # CADAgent, QSAgent, BPTCKCSAgent (Lab Link), SchedulerAgent
│   └── gates/
│       ├── quality_gate.py       # Các Cổng kiểm soát kỹ thuật số học xác định
│       └── human_gate.py         # Human-in-the-loop Gate (Ký duyệt Kỹ sư trưởng)
│
├── tools/                        # ⚙️ CÔNG CỤ TÍNH TOÁN XÁC ĐỊNH (PURE PYTHON, ZERO LLM)
│   ├── cutting_stock_solver.py   # Solver cắt thép 1D (Column Generation GLOP + CP-SAT, tách nhóm Ø + mác)
│   ├── dynamic_schedule_builder.py # Khởi tạo Tiến độ CPM 100% công thức sống & 4 Biểu đồ Native Excel
│   ├── equipment_fleet_scheduler.py # Động cơ lập tiến độ Ca xe, Ca máy & Kế hoạch Dầu Diezel Vincons/TT 37/2026
│   ├── package_dispatcher.py     # Bộ điều phối đóng gói phân quyền Hub & Spoke 5 gói vệ tinh công trường
│   ├── bbs_loader.py             # Đọc BBS thật từ Excel / CSV / JSON, nhận diện mối nối
│   ├── rebarcut_export.py        # Xuất kết quả cắt thép theo bố cục RebarCut Pro Excel (.xlsx)
│   ├── schedule_loader.py        # Đọc tiến độ thật từ MS Project XML / Excel / CSV, đối chiếu ngày logic
│   ├── qs_loader.py              # Đọc bảng QS thật, tỷ lệ chi phí, tính G_XD TT 36/2026 + đối chiếu
│   ├── qs_export.py              # Xuất bảng tổng hợp G_XD + chi tiết công tác (.xlsx)
│   ├── payment.py                # Mẫu 03a NĐ 254/2025: KL thực hiện × đơn giá HĐ, tạm ứng, giữ lại
│   ├── lab_qaqc.py               # Xử lý kết quả thí nghiệm nén mẫu R7/R28 & liên kết QLCL
│   ├── excel_eval.py             # Tính công thức Excel chưa có kết quả lưu sẵn (Pure Python)
│   ├── cpm_calculator.py         # Bộ tính CPM: FS/SS/FF/SF + lag, lịch nghỉ, Forward/Backward Pass
│   └── cad_diff_engine.py        # Động cơ so sánh phiên bản bản vẽ CAD Rev00 vs Rev01
│
├── schemas/                      # 📋 ĐẶC TẢ SCHEMA DỮ LIỆU CHUYÊN NGÀNH
│   ├── site_log_schema.py        # Schema Nhật ký hiện trường & Khối lượng hoàn công As-Built
│   └── lab_result_schema.py      # Schema Phiếu thí nghiệm phòng LAS-XD (R7/R28, kéo thép, PDA)
│
├── aec_core/                     # 🔍 BỘ CÔNG CỤ KIỂM TOÁN VÀ XÁC THỰC ĐỘC LẬP
│   ├── audit_verifier.py         # AECAuditVerifier: Quét toàn diện, 0 số chết, điểm 100/100
│   └── project_state.py          # Trình quản lý trạng thái dự án cơ sở
│
├── workflows/                    # 📚 QUY TRÌNH KỸ THUẬT & TIÊU CHUẨN THI CÔNG
│   ├── 00_TONG_QUAN_QUY_TRINH_KHEP_KIN_AEC.md
│   ├── 01_QUY_TRINH_BOC_TACH_HINH_HOC_TAKEOFF.md
│   ├── 02_QUY_TRINH_TO_HOP_CAT_THEP_1D.md
│   ├── 03_QUY_TRINH_DU_TOAN_GXD_TT11.md
│   ├── 04_QUY_TRINH_THANH_TOAN_PHU_LUC_03A.md
│   ├── 05_QUY_TRINH_TIEN_DO_CPM_MS_PROJECT.md
│   ├── 06_QUY_TRINH_KCS_LOGIC_CHEO_XUAT_WORD.md
│   ├── 07_QUY_TRINH_THUYET_MINH_BIEN_PHAP_HUGGINGFACE.md
│   ├── 08_QUY_TRINH_PHAN_TICH_TONG_HOP_VAT_TU_DINH_MUC.md
│   ├── 09_QUY_TRINH_THONG_KE_THEP_BBS_VA_TAN_SUAT_THI_NGHIEM.md
│   ├── 10_QUY_TRINH_THU_NHAN_VA_HOP_NHAT_DU_LIEU_DA_PHUONG_THUC.md
│   ├── 11_QUY_TRINH_VALIDATION_KIEM_TRA_CHEO.md
│   ├── 12_SO_DO_DIEU_PHOI_MULTI_AGENT_TOAN_HE_THONG.md
│   ├── 13_KIEN_TRUC_STATE_GRAPH_V3_SUPERVISOR_PATTERN.md
│   ├── 14_QUY_TRINH_DONG_GOI_2_GOI_HO_SO_MACRO_VA_MICRO_14_BO.md
│   └── 15_QUY_TRINH_DONG_GOI_HUB_AND_SPOKE_PHAN_QUYEN_THUC_CHIEN.md
│
├── templates/                    # 📦 SẢN PHẨM MẪU SỐ HÓA HOÀN THIỆN
│   ├── Ho_So_KCS_QS_TienDo_Cau_Khai_Hoang_2_Km14+363.65.xlsx # Master 14 Sheet liên kết động
│   ├── Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Khai_Hoang_2.docx  # 43 Biên bản KCS Word chuẩn NĐ 207
│   ├── Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.xml               # Tiến độ MS Project XML
│   ├── Tien_Do_Thi_Cong_Cau_Khai_Hoang_2.mpp               # Tiến độ MS Project MPP
│   └── BAO_CAO_THAM_TRA_AEC_AUDIT_KHAI_HOANG_2.md          # Báo cáo thẩm tra Audit Score 100/100
│
├── examples/                     # 🚀 SCRIPT THỰC THI & HỒ SƠ MẪU DỰ ÁN THỰC CHIẾN
│   ├── HO_SO_CONG_HOP_TUYEN_A5/  # 📁 Trọn bộ Hồ sơ Cống hộp Tuyến A5 + 5 Gói Hub & Spoke
│   │   ├── BANG_BOC_TACH_KHOI_LUONG_CONG_HOP_TUYEN_A5.xlsx
│   │   └── HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5/ (GOI_A đến GOI_E + MANIFEST)
│   ├── TIEN_DO_THI_CONG_CUM_B9_OLYMPIC/ # 📁 Hồ sơ mẫu Tiến độ & Ca máy Cụm B9 Olympic
│   │   ├── 260820_TDTC_Cum_B9_TINH_GIAN_CHUAN_CPM.xlsx # Master Excel 100% công thức sống + 4 biểu đồ Native
│   │   ├── 260920_TDTC_Cum_B9_SanLap_Va_DuongNoiBo_Olympic_ThuongTin.xlsx
│   │   └── 260920_Tien_Do_Thi_Cong_Cum_B9_Olympic_ThuongTin.xml
│   ├── generate_clean_streamlined_schedule_b9.py # Generator tiến độ Cụm B9 tinh giản chuẩn CPM
│   ├── build_14_micro_standalone_dossiers.py     # Generator 14 bộ hồ sơ vi mô chuyên sâu độc lập
│   ├── apply_khai_hoang_2_full.py                # Áp giá TT 38/2026, TT 36/2026 và liên kết động 14 Sheet
│   ├── build_khai_hoang_2_dossier.py             # Dựng hồ sơ từ dữ liệu gốc 166 dòng BBS
│   ├── run_pipeline.py                           # Runner kiểm tra toàn diện 14 Sheet Master
│   └── run_cad_diff_demo.py                      # Demo so sánh bản vẽ CAD Rev00 vs Rev01
│
├── tests/                        # 🧪 TEST TỰ ĐỘNG (python -m unittest discover tests)
│   ├── test_dynamic_cpm_schedule.py              # Test tiến độ CPM 100% công thức sống & biểu đồ native
│   ├── test_cad_and_state.py                     # Test đọc DXF hình học & lưu/khôi phục State
│   ├── test_equipment_fleet_scheduler.py         # Test động cơ ca xe, ca máy & nhiên liệu dầu
│   ├── test_package_dispatcher.py                # Test bộ đóng gói phân quyền Hub & Spoke
│   ├── test_cutting_stock_solver.py              # Test solver cắt thép 1D CP-SAT
│   ├── test_schedule.py                          # Test tính toán tiến độ đường găng CPM
│   └── test_payment.py                           # Test biểu thanh toán Phụ lục 03a
│
├── run_state_graph.py            # 🌟 ENTRY POINT: State Graph & Supervisor Runner v3.0
├── requirements.txt              # Danh mục thư viện phụ thuộc (ortools, openpyxl, pandas...)
├── pyproject.toml                # Cấu hình đóng gói hệ thống chuẩn PEP 621
├── CONTRIBUTING.md               # Quy chuẩn đóng góp mã nguồn (Zero LLM Math, Real Data)
├── LICENSE                       # Giấy phép phần mềm mã nguồn mở MIT
└── README.md                     # Tài liệu hướng dẫn chính thức của dự án
```

---

## ⚖️ 8. Giấy phép Bản quyền (License & Authorship)

- **Tác giả & Bản quyền trí tuệ**: **Nguyễn Bảo Tú** ([@baotuhg](https://github.com/baotuhg))
- **Kho lưu trữ chính thức**: [https://github.com/baotuhg/23HG-multiagent-system](https://github.com/baotuhg/23HG-multiagent-system)
- Dự án được phân phối dưới giấy phép mã nguồn mở **[MIT License](LICENSE)**.
- **Quy định bắt buộc**: Mọi cá nhân, tổ chức sử dụng, sao chép, trích xuất mã nguồn hoặc kế thừa hệ thống **BẮT BUỘC** phải giữ nguyên thông báo bản quyền của tác giả **Nguyễn Bảo Tú** và dẫn liên kết đầy đủ về kho lưu trữ gốc theo đúng điều khoản pháp lý của MIT License. Nghiêm cấm mọi hành vi re-upload xóa nguồn hoặc mạo danh tác giả gốc.
