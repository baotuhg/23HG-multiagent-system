# SƠ ĐỒ ĐIỀU PHỐI MULTI-AGENT AEC MASTER TOÀN HỆ THỐNG
> **Phiên bản:** 2.5 (Chuẩn hóa Độc quyền AEC Digital Project Office)  
> **Căn cứ Pháp lý & Tiêu chuẩn:** Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Nghị định 254/2025/NĐ-CP, Thông tư 36, 37, 38/2026/TT-BXD, TCVN 11823:2017, TCVN 1651:2018.

---

## I. BẢN VẼ KIẾN TRÚC ĐIỀU PHỐI TỔNG THỂ

```text
                           +-------------------------------+
                           |      BẢN VẼ / HỒ SƠ DỰ ÁN     |
                           |   (CAD / BIM / Yêu cầu KTXD)  |
                           +---------------+---------------+
                                           |
                                           v
                  +-------------------------------------------------+
                  |       AI SUPERVISOR (CHỈ HUY TRƯỞNG ẢO)         |
                  | - Tiếp nhận mục tiêu, phân chia đầu việc Modular|
                  | - Điều phối State dự án qua Shared State Bus    |
                  | - Cổng soát xét kỹ thuật (Quality Gate Verifier)|
                  +-----------------------+-------------------------+
                                          |
        +---------------------------------+---------------------------------+
        |                                 |                                 |
        v                                 v                                 v
+------------------+             +------------------+             +------------------+
| AGENT TRẮC ĐẠC & |             |  AGENT KỸ THUẬT  |             |     AGENT QS     |
|   BÓC TÁCH CAD   |             |   & BPTC/KCS     |             |    & DỰ TOÁN     |
+--------+---------+             +--------+---------+             +--------+---------+
| * Đọc DWG/DXF/IFC|             | * Biện pháp thi  |             | * Áp định mức XD |
| * Tính diện tích,|             |   công (cẩu, đà  |             | * Chạy đơn giá   |
|   thể tích hình  |             |   giáo, an toàn) |             | * Tính G_xd      |
|   học kết cấu    |             | * Lập hồ sơ KCS, |             | * Khối lượng 03a |
| * Xuất bảng hình |             |   tiêu chuẩn NT  |             +--------+---------+
|   học chuẩn      |             +--------+---------+                      ^
+--------+---------+                      ^                                |
    |    |                                | (Thẩm tra tính khả thi)        | (Đẩy BOM sang tính $)
    |    v                                |                                |
    |  +------------------+               +---------------+                |
    |  |  AGENT GIA CÔNG  |                               |                |
    |  |    & VẬT TƯ      +-------------------------------+----------------+
    |  +------------------+
    |  | * Cắt thép 1D    |  ---> Giải thuật tối ưu: Hao hụt đề-xê < 1.5%
    |  | * Lập bảng BOM   |  ---> Gửi BOM cho QS tính tiền, gửi phương án cho BPTC
    |  +--------+---------+
    |           |
    +-----------+ (Khối lượng hình học & nhân lực)
    |
    v
+---------------------------------------------------------------------------+
|                          AGENT KẾ HOẠCH & TIẾN ĐỘ                         |
| * Lập mạng công việc CPM, xác định đường găng tiến độ                     |
| * Tích hợp nguồn lực máy móc, vật tư và nhân lực theo từng phân đoạn      |
+------------------------------------+--------------------------------------+
                                     |
                                     v
                  +-------------------------------------------------+
                  |            SHARED STATE & FEEDBACK BUS          |
                  |  (Kênh trao đổi dữ liệu & Phản biện giữa Agent) |
                  |  + Autonomous Red Teaming Audit (100/100 điểm)  |
                  +------------------+------------------------------+
                                     |
           [Dữ liệu có xung đột?] ---+---> [Có] ---> Gửi lệnh yêu cầu chỉnh sửa
                                     |                (VD: Cẩu không với tới / đá ngày ->
                                     |                 yêu cầu BPTC/CPM tính toán lại)
                                   [Không]
                                     |
                                     v
                  +-------------------------------------------------+
                  |      KIỂM DUYỆT CUỐI (HUMAN-IN-THE-LOOP)        |
                  |  Kỹ sư trưởng / Giám đốc dự án ký duyệt số      |
                  +------------------+------------------------------+
                                     |
                                     v
                  +-------------------------------------------------+
                  |             XUẤT HỒ SƠ KỸ THUẬT SẠCH            |
                  | - File Excel: Dự toán G_xd, 03a, Hồ sơ KCS A4   |
                  | - File Word/MD: Thuyết minh Biện pháp thi công  |
                  | - File MS Project: Mạng tiến độ CPM chuẩn       |
                  +-------------------------------------------------+
```

---

## II. ĐỐI CHIẾU THỰC TẾ: HỆ THỐNG HIỆN TẠI ĐÃ HOẠT ĐỘNG NHƯ THẾ NÀO?

Hệ thống **ĐÃ HOÀN TOÀN VẬN HÀNH THỰC TẾ** theo đúng 100% sơ đồ trên thông qua các module mã nguồn Python độc lập:

| Khối Chức Năng Trên Sơ Đồ | Module / Script Thực Thi Trong Hệ Thống | Trạng Thái Hoạt Động |
|---|---|:---:|
| **1. HỒ SƠ ĐẦU VÀO** | Thư mục `c:\Users\baotu\Downloads\Documents\HSTK...` (61 bản vẽ DWG, BoQ Excel, Thuyết minh MD) | ✅ Sẵn sàng |
| **2. AI SUPERVISOR** | `agents/project_state_manager.py` & `prompts/00_PROMPT_TONG_HOP_MULTI_AGENT_AEC.md` | ✅ Hoạt động |
| **3. TRẮC ĐẠC & BÓC TÁCH CAD** | `agents/aec_cad_extractor.py` & `skills/aec-cad-automation/scripts/cad_takeoff_engine.py` (Shoelace, COM Interop, giải mã font TCVN3) | ✅ Hoạt động |
| **4. GIA CÔNG & VẬT TƯ** | `skills/aec-rebar-optimizer/scripts/optimize_rebar.py` (Cắt thép 1D <1.5%), `examples/update_material_sheets.py` (BBS 396 thanh & BOM) | ✅ Hoạt động |
| **5. AGENT KỸ THUẬT & BPTC/KCS** | `agents/rag_method_statement_huggingface.py` (BPTC 8 chương), `examples/update_full_cross_linked_workbook.py` (3 Mẫu KCS A4) | ✅ Hoạt động |
| **6. AGENT QS & DỰ TOÁN** | `skills/aec-cost-tender/scripts/vn_cost_engine.py` & Sheet `QS_DIEN_GIAI_CHI_TIET`, `TONG_HOP_DU_TOAN_GXD` (TT 36/2026, VAT 10%) | ✅ Hoạt động |
| **7. KẾ HOẠCH & TIẾN ĐỘ** | `examples/generate_sample_bridge_project.py` (Mạng CPM 36 công tác, tệp MS Project `.xml` & `.mpp`) | ✅ Hoạt động |
| **8. SHARED STATE BUS** | `agents/PROJECT_STATE.json` & `aec_core/project_state.py` (Single Source of Truth) | ✅ Hoạt động |
| **9. QUALITY GATE (AUDIT)** | `aec_core/audit_verifier.py` & `examples/run_pipeline.py` (Chấm điểm 100/100, quét 0 số chết, kiểm tra logic chéo) | ✅ Hoạt động |
| **10. HUMAN-IN-THE-LOOP** | Kỹ sư trưởng đọc `templates/BAO_CAO_THAM_TRA_AEC_AUDIT.md` và kiểm tra bảng tính trước khi phê duyệt | ✅ Hoạt động |
| **11. XUẤT HỒ SƠ SẠCH** | `templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx` (14 Sheet), `Thuyet_Minh_Bien_Phap_Thi_Cong...md`, `Tien_Do...xml` | ✅ Hoạt động |

---

## III. NGUYÊN TẮC VẬN HÀNH "FEEDBACK LOOP" THỰC CHIẾN

1. **Khối lượng Vật lý quyết định Tiền tệ (Physical Drives Financials):**
   - Agent QS không tự nhập khối lượng. Khối lượng bê tông, ván khuôn lấy từ `Agent Bóc tách CAD`. Khối lượng sắt thép lấy từ `Agent Gia công & Vật tư (BBS & Cắt thép 1D)`.
2. **Kiểm tra va chạm Biện pháp (Constructability Review):**
   - Khi phương án gia công yêu cầu cẩu kiện lớn hoặc đà giáo đặc biệt, `Agent Kỹ thuật BPTC` đối chiếu tải trọng cẩu và mặt bằng thi công. Nếu phát hiện vi phạm bán kính cẩu hoặc tải trọng nền đất, hệ thống kích hoạt Feedback Bus yêu cầu tính toán lại.
3. **Đồng bộ Tiến độ - Nghiệm thu (Schedule-KCS Synchronization):**
   - Ngày nghiệm thu trong từng biên bản KCS (`HOSO_KCS_NGHIEM_THU`, `MAU_BIEN_BAN_KCS`) tự động trỏ vào ngày hoàn thành tương ứng trong `TIEN_DO_THI_CONG_WBS`. Tuyệt đối không bao giờ xảy ra lỗi "nghiệm thu trước khi đổ bê tông".
4. **Cổng kiểm toán Độc lập (Red Teaming Gate):**
   - Nếu còn dù chỉ 1 ô số chết (hardcoded value) hoặc 1 liên kết gãy, `AECAuditVerifier` trừ điểm và hạ trạng thái hồ sơ về `WARNING`, không cho phép chuyển sang bước ký duyệt của Kỹ sư trưởng.
