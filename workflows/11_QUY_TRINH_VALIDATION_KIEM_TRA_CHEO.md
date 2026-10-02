# QUY TRÌNH 11: KIỂM TRA CHÉO & VALIDATION TỰ ĐỘNG HỆ THỐNG 14 SHEET
> **Phiên bản:** 2.0 | **Cập nhật:** 2026-09-26 | **Căn cứ:** NĐ 207/2026/NĐ-CP, TT 32/2026/TT-BXD

---

## I. MỤC TIÊU

Đảm bảo 100% tính nhất quán dữ liệu trong toàn bộ hệ thống 14 Sheet Excel Master:
- **Không số chết (Zero Dead Numbers):** Mọi ô tính toán phải là công thức
- **Chuỗi liên kết khép kín:** BBS → QS → Cấp phối → Tần suất → Biên bản
- **Audit tự động:** Chạy `run_pipeline.py` để báo cáo 100/100

---

## II. CHUỖI LIÊN KẾT CHUẨN

```
THONG_KE_THEP_CHI_TIET (BBS 396 thanh)
    ↓ SUMIFS(O, B/D/E) theo kết cấu & đường kính
QS_DIEN_GIAI_CHI_TIET (Cột J - KL từng hạng mục)
    ├─→ TONG_HOP_DU_TOAN_GXD (E6=L101 QS → G_XD với VAT 10%, TL 5.5%)
    ├─→ THANH_TOAN_KY_PHU_LUC_03A (KL thanh toán từng kỳ)
    ├─→ PHAN_TICH_VAT_TU_WBS → TONG_HOP_VAT_TU_TOAN_BO (SUMIF)
    └─→ CAP_PHOI_1M3_VA_TAN_SUAT Bảng 1 (J9:J17 thể tích từng mác BT)

KHOI_LUONG_DAO_DAP (F21 Tổng đào, I21 Tổng đắp)
    ↓ link trực tiếp
QS_DIEN_GIAI_CHI_TIET (J33=F21, J34=I21)

TO_HOP_CAT_THEP_11M7 (D cột số lượng)
    ↑ COUNTIFS ← THONG_KE_THEP_CHI_TIET (theo đường kính & kết cấu)
    → F (Tổng dài), H (Số cây), J (Đề-xê), K (Trọng lượng), L (% hao hụt)

CAP_PHOI_1M3_VA_TAN_SUAT Bảng 1 → Bảng 2
    K18,L18,M18,N18,O18 (Tổng vật liệu) → G35:G39 (KL tần suất KCS)

TIEN_DO_THI_CONG_WBS (Finish date cột J)
    ↓ trỏ thẳng
HOSO_KCS_NGHIEM_THU (Cột H - Ngày nghiệm thu)
    ↓ VLOOKUP
MAU_BIEN_BAN_KCS (Auto: BB số, tên CV, KL, TC, ngày)
MAU_BB_NGHIEM_THU_VAT_LIEU (Auto: loại VL, TC, KL, tần suất)
MAU_BB_LAY_MAU_HIEN_TRUONG (Auto: R7=ngày+7, R28=ngày+28)
```

---

## III. CHECKLIST KIỂM TRA TRƯỚC KHI PHÁT HỒ SƠ

### A. Kiểm tra Excel Master (tự động)
```powershell
cd 23HG-multiagent-system
python examples/update_full_cross_linked_workbook.py  # Tái tạo công thức
python examples/run_pipeline.py                        # Chạy audit → phải 100/100
```

### B. Kiểm tra thủ công trong Excel
| STT | Ô cần kiểm tra | Kết quả kỳ vọng |
|-----|----------------|----------------|
| 1 | `QS!J22` | `=SUMIFS(BBS!$O..., BBS!$B..., "*Cọc khoan nhồi*")` |
| 2 | `CAP_PHOI!J17` | `=QS!J86` (vữa chèn khe C40) |
| 3 | `CAP_PHOI!G38` | `=N18` (Nước tổng Bảng 1) |
| 4 | `CAP_PHOI!G39` | `=O18` (Phụ gia tổng Bảng 1) |
| 5 | `TO_HOP!D6` | `=COUNTIFS(BBS!$E..., 25, BBS!$B..., "*Cọc khoan*")` |
| 6 | `QS!J33` | `=KHOI_LUONG_DAO_DAP!F21` |
| 7 | `HOSO_KCS!H6:H27` | Tất cả là `=TIEN_DO!J...` |
| 8 | `MAU_BB!C22` | `=TEXT(VLOOKUP(..., 8, FALSE), "dd/mm/yyyy")` |

### C. Kiểm tra kết nối CAD (nếu cập nhật bản vẽ)
```powershell
python examples/run_data_ingestion_pipeline.py
# Kiểm tra data/PROJECT_STATE.json được cập nhật
```

---

## IV. QUY TẮC "SỐ KHÔNG CHẾT"

| Loại ô | Quy tắc |
|--------|---------|
| Khối lượng tính toán | PHẢI là công thức `=E*F*G*H*I` hoặc SUMIFS |
| Khối lượng đặc biệt | Số chứng minh được: Đơn vị Điểm/Lỗ/Gói (BBNT-01, 06, 10, 22) |
| Đơn giá | SỐ CHẾT hợp lệ (lấy từ bảng đơn giá ban hành) |
| Hệ số định mức | SỐ CHẾT hợp lệ (lấy từ ĐM 1776/2007, ĐM 2016) |
| Ngày tháng | Phải là công thức hoặc nhập liệu theo MAU_BB |

---

## V. ĐIỂM CHUẨN KIỂM TRA (BENCHMARK)

| Chỉ số | Ngưỡng OK | Ngưỡng CẢNH BÁO |
|--------|-----------|----------------|
| Điểm audit | ≥ 98/100 | < 95/100 → DỪNG phát hồ sơ |
| Số công thức quét được | ≥ 750 | < 600 → Kiểm tra sheet rỗng |
| Số cảnh báo số chết | ≤ 4 | > 8 → Rà soát ngay |
| % hao hụt thép TO_HOP | < 1.5% | > 3% → Tối ưu lại sơ đồ cắt |
| Số biên bản KCS | = 22 | ≠ 22 → Kiểm tra HOSO_KCS |

---

## VI. XỬ LÝ KHI PHÁT HIỆN LỖI

1. **Lỗi số chết:** Chạy `update_full_cross_linked_workbook.py` để tái tạo
2. **Lỗi VLOOKUP #N/A:** Kiểm tra ô C2 (số BB) có trong bảng tra không
3. **Lỗi SUMIFS = 0:** Kiểm tra wildcard trong công thức có khớp với tên cấu kiện trong BBS
4. **Lỗi ngày #VALUE!:** Ô C3 trong MAU_BB_LAY_MAU phải nhập đúng định dạng dd/mm/yyyy
5. **Lỗi CAD extractor:** Kiểm tra AutoCAD đang mở file hoặc chạy MCP server

---

## VII. PHỤ LỤC: SCRIPT AUDIT NHANH

```python
# Chạy nhanh toàn bộ audit
from aec_core.audit_verifier import AECAuditVerifier
av = AECAuditVerifier()
av.audit_excel_workbook("templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")
av.export_audit_report("templates/BAO_CAO_THAM_TRA_AEC_AUDIT.md")
```
