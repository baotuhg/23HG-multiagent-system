# QUY TRÌNH 01: BÓC TÁCH HÌNH HỌC TAKEOFF & PHÂN RÃ MODULAR WBS
## NGUYÊN TẮC BẮT BUỘC: 100% CÔNG THỨC ĐỘNG - CẤM SỐ CHẾT

---


> **Trạng thái:** ✅ **Đã hiện thực** — đo bóc từ bảng cấu kiện: `tools/takeoff_loader.py`,
> `tools/takeoff_rules.py`, `CADAgent` (pha CAD_TAKEOFF), xuất Bảng 6.1 / 6.2.
> 🟡 Bóc tách hình học trực tiếp từ DXF/DWG (`core/agents/aec_cad_extractor.py`) còn là nguyên mẫu
> (cần `ezdxf`, chưa phải đường chạy chính).

### I. NGUYÊN TẮC CỐT LÕI
1. **Tuyệt đối không điền số chết (Hard-coded numbers):** Mọi con số khối lượng phải được giải trình rõ nguồn gốc kích thước hình học từ bản vẽ.
2. **Cấu trúc 12 cột chuẩn:**
   - Cột A (1): TT
   - Cột B (2): Mã hiệu định mức (Mã định mức Bộ Xây dựng: AB, AC, AF, AK, AL...)
   - Cột C (3): Nội dung công tác & Diễn giải hình học cấu kiện
   - Cột D (4): Đơn vị tính (m3, m2, m, Tấn, Cái, Bộ, Gói)
   - Cột E (5): Số lượng cấu kiện (N)
   - Cột F (6): Chiều dài L (m)
   - Cột G (7): Chiều rộng W (m)
   - Cột H (8): Chiều cao / Chiều dày H (m)
   - Cột I (9): Hệ số hình học (ví dụ: vát taluy $= 0.5$, diện tích hình tròn $\pi R^2$, hệ số nở thành cọc $= 1.12$)
   - Cột J (10): Khối lượng công thức
   - Cột K (11): Đơn giá dự toán (VNĐ)
   - Cột L (12): Thành tiền công thức

### II. CÔNG THỨC SỐNG BẮT BUỘC
- **Dòng con (Diễn giải cấu kiện chi tiết):**
  ```excel
  =E{row} * F{row} * G{row} * H{row} * I{row}
  ```
- **Dòng cha (Công tác chính tổng hợp):**
  ```excel
  =SUM(J{dòng_con_đầu} : J{dòng_con_cuối})
  ```
- **Cột Thành tiền (L):**
  ```excel
  =J{row} * K{row}
  ```

### III. TRÌNH TỰ PHÂN RÃ MODULAR WBS GIAO THÔNG
1. **Phần I: Móng cọc & Thí nghiệm:** Khoan cọc nhồi, ống vách, bê tông cọc, cốt thép cọc, đập đầu cọc, siêu âm, PDA.
2. **Phần II: Kết cấu phần dưới:** Hố móng, bê tông lót C10, bệ mố trụ, thân mố trụ, tường cánh, tường ngực, xà mũ, đá kê gối C30.
3. **Phần III: Kết cấu phần trên:** Dầm chủ Super-T (C45), cáp DUL 15.2mm, ván khuôn đúc sẵn, dầm ngang, bản mặt cầu C35, liên tục nhiệt, gối chậu, khe co giãn.
4. **Phần IV: Kết cấu phụ trợ & Đường đầu cầu:** Bản quá độ C25, đắp hạt chọn lọc K98 sau mố, lan can thép mạ kẽm, thoát nước D150, màng chống thấm & thảm BTN C16 dày 7cm.
