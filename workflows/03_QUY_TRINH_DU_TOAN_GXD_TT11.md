# QUY TRÌNH 03: TỔNG HỢP KINH PHÍ DỰ TOÁN XÂY DỰNG (G_XD)
## CĂN CỨ PHÁP LÝ: THÔNG TƯ 36/2026/TT-BXD & LUẬT XÂY DỰNG SỐ 135/2025/QH15

---


> **Trạng thái:** ✅ **Đã hiện thực** — `tools/qs_loader.py` và `tools/survey_estimate.py`;
> có golden test đối chiếu dự toán thật đã thẩm định (cả khảo sát và xây lắp), khớp đến từng đồng.

### I. CÔNG THỨC TÍNH TOÁN KINH PHÍ XÂY DỰNG CHUẨN

$$G_{XD} = T + GT + TL + VAT$$

Trong đó:
1. **Chi phí trực tiếp ($T$):** Gồm Chi phí Vật liệu ($VL$) + Chi phí Nhân công ($NC$) + Chi phí Máy thi công ($M$).
   - Link trực tiếp từ ô Tổng cộng của bảng Tiên lượng chi tiết:
     ```excel
     =QS_DIEN_GIAI_CHI_TIET!L{dòng_tổng_cộng}
     ```
2. **Chi phí gián tiếp ($GT$):**
   - Định mức công trình Giao thông theo tuyến: **$GT = 7.30\% \times T$**
     + Chi phí chung: $5.10\% \times T$
     + Chi phí nhà tạm để ở và điều hành thi công: $1.20\% \times T$
     + Chi phí một số công tác không xác định được khối lượng từ thiết kế: $1.00\% \times T$
     ```excel
     =E6 * 0.073
     ```
3. **Thu nhập chịu thuế tính trước ($TL$):**
   - Định mức công trình giao thông: **$TL = 5.50\% \times (T + GT)$**
     ```excel
     =(E6 + E7) * 0.055
     ```
4. **Chi phí xây dựng trước thuế ($G$):**
   - $G = T + GT + TL$
     ```excel
     =E6 + E7 + E11
     ```
5. **Thuế giá trị gia tăng ($VAT$):**
   - Theo quy định thuế hiện hành: **$VAT = 8.00\% \times G$** (hoặc $10\%$ tùy thời điểm chính sách thuế)
     ```excel
     =E12 * 0.08
     ```
6. **Tổng cộng kinh phí dự toán xây dựng ($G_{XD}$):**
   - $G_{XD} = G + VAT$
     ```excel
     =E12 + E13
     ```

### II. NGUYÊN TẮC LIÊN KẾT ĐỘNG
- Bảng dự toán tổng hợp không được chứa bất kỳ giá trị nhập tay nào ngoài tỷ lệ phần trăm định mức.
- Mọi biến động từ bản vẽ hình học bóc tách sẽ ngay lập tức được phản ánh vào tổng giá trị gói thầu $G_{XD}$.
