---
name: aec-cost-tender
description: Quản lý Tiên lượng, Dự toán, Đấu thầu & So sánh chi phí xây dựng (BoQ & Tender Cost Engineering) theo tiêu chuẩn Việt Nam (Thông tư 36/2026/TT-BXD, Thông tư 38/2026/TT-BXD, Thông tư 37/2026/TT-BXD, Luật Xây dựng 135/2025/QH15, Nghị định 206/2026/NĐ-CP và NĐ 207/2026/NĐ-CP). Hỗ trợ bóc tách bảng khối lượng mời thầu (BoQ), tính tổng mức chi phí xây dựng G_XD (T + GT + TL + VAT), so sánh 2 phiên bản dự toán (Document Diff), phân tích giá dự thầu và xuất file Excel dự toán phân cấp WBS chuẩn mẫu Bộ Xây dựng.
---

# Tiên lượng, Dự toán & Đấu thầu Xây dựng Việt Nam (AEC-Cost-Tender)

## 1. Mục tiêu
Chuẩn hóa và tự động hóa quy trình quản lý chi phí, lập tiên lượng (BoQ - Bill of Quantities), thẩm tra dự toán và chấm thầu theo pháp luật và tiêu chuẩn xây dựng Việt Nam:
- **Quản lý Tiên lượng & Dự toán chi phí:** Tự động tổng hợp chi phí trực tiếp ($T = VL + NC + MTC$), chi phí gián tiếp ($GT = C_C + C_{NT} + C_{KXD}$), thu nhập chịu thuế tính trước ($TL$) và thuế giá trị gia tăng ($VAT$) để ra tổng chi phí xây dựng sau thuế ($G_{XD}$).
- **So sánh 2 phiên bản hồ sơ (Document Diff):** Tự động phát hiện biến động khối lượng giữa Thiết kế BVTC vs Thực tế thi công, chỉ rõ các đầu việc phát sinh, cắt giảm và tác động tài chính.
- **Phân tích hồ sơ dự thầu (Bid Evaluation):** So sánh bảng đơn giá của các nhà thầu tham gia đấu thầu để phát hiện đơn giá bất thường (quá cao hoặc phá giá).
- **Xuất bảng dự toán chuẩn mẫu Bộ Xây dựng:** Xuất file Excel (.xlsx) 2 sheet chuẩn gồm: *Bảng tổng hợp kinh phí dự toán chi phí xây dựng* và *Bảng khối lượng dự toán chi tiết*.

## 2. Trục pháp lý & Định mức cốt lõi

1. **Trục pháp lý quản lý chi phí:**
   - **Luật Xây dựng số 135/2025/QH15** (thay thế Luật số 50/2014, 62/2020/QH14 đã hết hiệu lực).
   - **Nghị định số 206/2026/NĐ-CP** (và Nghị định 207/2026/NĐ-CP) về quản lý chi phí đầu tư xây dựng.
   - **Thông tư số 36/2026/TT-BXD** của Bộ Xây dựng hướng dẫn một số nội dung xác định và quản lý chi phí đầu tư xây dựng.
   - **Thông tư số 38/2026/TT-BXD** ban hành định mức xây dựng.
   - **Thông tư số 37/2026/TT-BXD** hướng dẫn phương pháp xác định các chỉ tiêu kinh tế kỹ thuật và đo bóc khối lượng.

2. **Hệ số định mức tỷ lệ % theo Thông tư 36/2026/TT-BXD:**
   - **Công trình Dân dụng (Trường học, Nhà ở):** Chi phí chung $C_C = 7.3\% \times T$; Nhà tạm $C_{NT} = 1.0\% \times T$; Chi phí KXD $C_{KXD} = 1.5\% \times T$; Thu nhập tính trước $TL = 5.5\% \times (T + GT)$.
   - **Công trình Giao thông (Cầu, đường theo tuyến):** $C_C = 6.2\% \times T$; $C_{NT} = 1.2\% \times T$; $C_{KXD} = 1.5\% \times T$; $TL = 5.5\% \times (T + GT)$.
   - **Công trình Hạ tầng kỹ thuật (San nền, Kè đá, Thoát nước):** $C_C = 5.8\% \times T$; $C_{NT} = 1.0\% \times T$; $C_{KXD} = 1.5\% \times T$; $TL = 5.5\% \times (T + GT)$.
   - **Công trình Nông nghiệp & PTNT (Kè suối, Thủy lợi):** $C_C = 6.0\% \times T$; $C_{NT} = 1.2\% \times T$; $C_{KXD} = 1.5\% \times T$; $TL = 5.5\% \times (T + GT)$.
   - **Thuế VAT:** $10\%$ (hoặc $8\%$ theo chính sách giảm thuế áp dụng từng thời kỳ).

3. **Mã hiệu định mức công tác theo Thông tư 38/2026/TT-BXD:**
   - `AB`: Công tác đào, đắp đất, đá, cát (móng, san nền, nền đường K95, K98).
   - `AC`: Công tác cọc (ép, đóng, khoan nhồi).
   - `AD`: Công tác làm đường (cấp phối đá dăm loại 1, bê tông nhựa).
   - `AE`: Công tác xây gạch, xây kè đá hộc, rãnh đá.
   - `AF`: Công tác bê tông tại chỗ (mác 150, 200, 250, 300).
   - `AG`: Công tác sản xuất, lắp dựng cốt thép ($\Phi \le 10$, $\Phi \le 18$, $\Phi > 18$).
   - `AH`: Công tác bê tông đúc sẵn (ống cống ly tâm, cọc đúc sẵn).
   - `AK`: Công tác hoàn thiện (trát, láng, ốp, lát, sơn).

## 3. Công cụ & Lệnh thực thi

### A. Kiểm tra và Tổng hợp Dự toán chuẩn Thông tư 36/2026
```bash
# Tự động tính toán tổng mức chi phí xây dựng G_XD (VNĐ)
python .openspace/skills/aec-cost-tender/scripts/boq_tool.py <duong_dan_file_du_toan.json>

# Chỉ định rõ loại công trình (DAN_DUNG, GIAO_THONG, HA_TANG_KY_THUAT, NONG_NGHIEP_PTNT)
python .openspace/skills/aec-cost-tender/scripts/boq_tool.py <file.json> --type GIAO_THONG
```

### B. So sánh 2 phiên bản Dự toán / Bảng Tiên lượng (Document Diff)
```bash
python .openspace/skills/aec-cost-tender/scripts/boq_tool.py <file_goc.json> --compare <file_moi.json>
```
*Kết quả xuất ra bảng chi tiết:*
- Các hạng mục phát sinh mới (kèm khối lượng, đơn giá, tổng tiền phát sinh tăng).
- Các hạng mục bị cắt giảm bỏ.
- Các hạng mục bị thay đổi khối lượng hoặc điều chỉnh đơn giá.
- Tổng chênh lệch chi phí trực tiếp phát sinh (+/- VNĐ).

### C. Xuất file Excel chuẩn mẫu Bộ Xây dựng (2 Sheets)
```bash
python .openspace/skills/aec-cost-tender/scripts/boq_tool.py <file.json> --excel Du_Toan_Cong_Trinh.xlsx
```
- **Sheet 1 (`TONG_HOP_DU_TOAN`):** Bảng tổng hợp chi phí xây dựng theo đúng mẫu Biểu Thông tư 36/2026/TT-BXD ($T, GT, TL, G, VAT, G_{XD}$).
- **Sheet 2 (`DU_TOAN_CHI_TIET`):** Bảng khối lượng dự toán chi tiết với đầy đủ Mã hiệu ĐM, ĐVT, Khối lượng, Đơn giá VL, NC, MTC và Thành tiền.

## 4. Phối hợp với Hệ sinh thái Kỹ thuật (BIM 5D)
1. **CAD $\rightarrow$ Dự toán:** Nhận diện tích mặt cắt đào đắp từ `aec-cad-automation` $\rightarrow$ nạp khối lượng vào mã `AB.11111` hoặc `AB.24111`.
2. **Dự toán $\rightarrow$ Thép:** Nhận khối lượng thép từ mã `AG.11111` $\rightarrow$ chuyển bảng thanh sang `aec-rebar-optimizer` để cắt tối ưu trên cây 11.7m.
3. **Dự toán $\rightarrow$ Nghiệm thu:** Chuyển toàn bộ danh mục công tác sang `aec-qlcl` để lập kế hoạch nghiệm thu công việc, ngày kiểm tra vật liệu và ngày ép mẫu bê tông R7, R28.
