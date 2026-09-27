# QUY TRÌNH 14: ĐÓNG GÓI 2 GÓI HỒ SƠ SONG HÀNH — MACRO (MASTER 14 SHEET) VÀ MICRO (14 BỘ CHUYÊN SÂU ĐỘC LẬP)
## CHUẨN ĐÓNG GÓI QUẢN TRỊ ĐIỀU HÀNH VÀ SẢN XUẤT CÔNG TRƯỜNG HỆ THỐNG AEC MULTI-AGENT
### Tuân thủ: Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Nghị định 99/2021/NĐ-CP, Thông tư 11/2021 & 12/2021/TT-BXD

---

## 1. BỐI CẢNH VÀ NGUYÊN LÝ THIẾT KẾ HAI TẦNG HỒ SƠ (DUAL-TIER DELIVERABLES)

Trong quản lý dự án xây dựng và công trình giao thông (cầu đường, hạ tầng kỹ thuật), nhu cầu sử dụng hồ sơ tại các cấp quản lý có sự phân hóa sâu sắc:

1. **Cấp Quản trị Điều hành & Chủ đầu tư / Ban QLDA (Vĩ mô - Macro Tier):**
   - Cần một **Bản đồ dữ liệu duy nhất (Single Source of Truth)** tích hợp toàn bộ các giai đoạn từ Bóc tách, Cắt thép, Định mức, Cấp phối, Vật tư, Dự toán $G_{XD}$, Thanh toán Phụ lục 03a, Tiến độ CPM và Danh mục KCS.
   - Toàn bộ 14 Sheet nằm trong 1 Workbook duy nhất, các công thức liên kết động chéo giữa các Sheet thông suốt, **tuyệt đối 0 số chết**, khi thay đổi một thông số kích thước hình học thì toàn bộ dự toán và tiến độ tự nhảy.
   - Điểm kiểm toán `aec_audit_verifier` phải đạt tối đa **100/100 điểm**.

2. **Cấp Công trường, Nhà xưởng gia công & Kỹ sư QA/QC hiện trường (Vi mô - Micro Tier):**
   - Kỹ sư công trường và đội trưởng tổ thợ không thể dùng một file Excel tổng hợp cồng kềnh chứa cả chục sheet để in ấn hoặc giao việc tại hiện trường.
   - Công trường cần **14 BỘ HỒ SƠ RIÊNG BIỆT (14 Standalone Dossiers)**, tương ứng chính xác 1-to-1 với 14 Sheet của file Master, nhưng ở **độ sâu chi tiết tối đa**:
     - *Xưởng cắt thép:* Cần file cắt thép riêng với sơ đồ ra phôi chi tiết từng thanh cắt từ cây 11.7m, bảng quản lý đề-xê, file CSV nhập máy cắt CNC.
     - *Đội thi công móng:* Cần bảng bóc tách diễn giải đào đắp hố móng, cao độ ngàm đá gốc, phân lớp đầm nén $\le 20cm$.
     - *Đội bê tông & Phòng Las-XD:* Cần định mức cấp phối 1m3, kế hoạch lấy mẫu nén R7/R28 theo từng cấu kiện.
     - *Cán bộ KCS:* Cần từng file biểu mẫu A4 độc lập in trực tiếp tại hiện trường (Biên bản công việc, vật liệu, ép mẫu) mà **không bị lỗi `#REF!`** khi mở riêng lẻ.

---

## 2. CẤU TRÚC THƯ MỤC BÀN GIAO 2 GÓI CHUẨN

```
📁 DU_AN_CAU_THI_CONG/
│
├── 📁 BO_HO_SO_01_MACRO_MASTER_14_SHEET/       <-- GÓI 01: VĨ MÔ / ĐIỀU HÀNH TỔNG HỢP
│   ├── Ho_So_KCS_QS_TienDo_Master_14_Sheets.xlsx (Master Workbook 14 Sheet liên kết động)
│   ├── Tien_Do_Thi_Cong_Master.xml               (File tiến độ XML chuẩn import MS Project)
│   ├── Tien_Do_Thi_Cong_Master.mpp               (File tiến độ MS Project)
│   └── BAO_CAO_THAM_TRA_AEC_AUDIT.md             (Báo cáo Audit độc lập Gate-3: Điểm 100/100)
│
└── 📁 BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO/      <-- GÓI 02: VI MÔ / SẢN XUẤT CÔNG TRƯỜNG
    ├── 01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx     (kèm file CSV sơ đồ cắt máy CNC)
    ├── 02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx     (bóc tách hình học đào đắp đất đá móng mố)
    ├── 03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx     (bóc tách tiên lượng bê tông, ván khuôn)
    ├── 04_Thong_Ke_Thep_Chi_Tiet_BBS_166_Dong.xlsx (BBS đầy đủ 166 dòng, đối chiếu THKL)
    ├── 05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx (cấp phối C10..C40 & ma trận 72 tổ mẫu KCS)
    ├── 06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx     (phân tích hao phí xi măng, cát, đá theo WBS)
    ├── 07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx (BOM đặt hàng & kế hoạch cung ứng 4 đợt)
    ├── 08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx      (dự toán chi tiết 102 công tác, G_XD)
    ├── 09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx (thanh toán kỳ mẫu 03.a NĐ 99/2021)
    ├── 10_Tien_Do_Thi_Cong_CPM_Gantt_Chart.xlsx  (tiến độ CPM, kèm XML, MPP và CSV)
    ├── 11_Danh_Muc_KCS_43_Bien_Ban_Nghiem_Thu.xlsx (danh mục 43 BBNT, kèm Word .docx trọn bộ)
    ├── 12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx (file Excel độc lập in A4 nghiệm thu CV)
    ├── 13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx (file Excel độc lập in A4 nghiệm thu VL)
    └── 14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx (file Excel độc lập in A4 nén mẫu BT)
```

---

## 3. BẢNG ĐỐI CHIẾU 1-1 GIỮA 14 SHEET MASTER VÀ 14 BỘ HỒ SƠ VI MÔ CHUYÊN SÂU

| Mã | Tên Sheet trong Master Workbook | Tên File Hồ sơ Vi mô Chuyên sâu | Chi tiết kỹ thuật & Tính năng độc lập |
| :---: | :--- | :--- | :--- |
| **01** | `TO_HOP_CAT_THEP_11M7` | `01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx`<br>+ `01_Phieu_Cat_Thep_...csv` | 5 tab: INPUT, SO_SANH, PA_TOI_UU, REMAIN, CHI_TIET. OR-Tools giải bài toán CSP, tối ưu cắt 3.236 cây 11.7m, hao hụt 1.85%, quản lý đề-xê tái sử dụng. |
| **02** | `KHOI_LUONG_DAO_DAP` | `02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx` | 3 tab: TONG_HOP_DAO_DAP, CHI_TIET_HO_MONG_M1_M2, CHI_TIET_DAP_DAU_CAU. Diễn giải kích thước đào móng ngàm đá gốc $\ge 0.5m$, mái taluy, đắp đầm cóc sau mố $\le 20cm$. |
| **03** | `QS_DIEN_GIAI_CHI_TIET` | `03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx` | 3 tab: QS_TONG_HOP, BE_MONG_MO_M1_M2, DAM_T_VA_DAM_NGANG. Bóc tách chi tiết hình học bê tông, ván khuôn từng bộ phận cấu kiện mố, dầm T, dầm ngang, bản mặt cầu, tường chắn. |
| **04** | `THONG_KE_THEP_CHI_TIET` | `04_Thong_Ke_Thep_Chi_Tiet_BBS_166_Dong.xlsx` | 4 tab: BBS, TONG_HOP_THEO_D, DOI_CHIEU_THKL, CANH_BAO. 166 dòng cốt thép đầy đủ 12.361 thanh (68.012 tấn), kiểm tra chéo THKL $\pm 2\%$, nhật ký OCR. |
| **05** | `CAP_PHOI_1M3_VA_TAN_SUAT` | `05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx` | 2 tab: CAP_PHOI_1M3_BE_TONG, TAN_SUAT_THI_NGHIEM_KCS. Định mức cấp phối C10-C40 & kế hoạch thí nghiệm 72 tổ mẫu nén R7/R28, kéo uốn cơ lý thép theo lô 50T. |
| **06** | `PHAN_TICH_VAT_TU_WBS` | `06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx` | Phân tích chi tiết nhu cầu xi măng (tấn), cát vàng (m3), đá dăm 1x2 (m3), thép tròn (kg), ván khuôn phủ phim (m2) theo từng gói công việc WBS. |
| **07** | `TONG_HOP_VAT_TU_TOAN_BO` | `07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx` | 2 tab: BOM_TONG_THE, CUNG_UNG_4_GIAI_DOAN. Tổng hợp BOM toàn dự án có xét hao hụt đặt hàng & phân kỳ 4 đợt cung ứng kèm định mức tồn kho an toàn vùng cao. |
| **08** | `TONG_HOP_DU_TOAN_GXD` | `08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx` | 2 tab: BOQ_CHI_TIET, TONG_HOP_GXD. Dự toán chi tiết 102 công tác đầy đủ mã định mức TT 12, chi phí trực tiếp $T$, chi phí gián tiếp $GT = 7.3\%$, thu nhập chịu thuế $TL = 5.5\%$, thuế VAT $10\%$, Tổng $G_{XD} = 4.769$ tỷ đồng. |
| **09** | `THANH_TOAN_KY_PHU_LUC_03A` | `09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx` | Hồ sơ đề nghị thanh toán Đợt 1 chuẩn Mẫu 03.a NĐ 99/2021/NĐ-CP: Lũy kế hoàn thành, giảm trừ thu hồi tạm ứng 20%, giảm trừ bảo hành 5%, giá trị thanh toán thực nhận. |
| **10** | `TIEN_DO_THI_CONG_WBS` | `10_Tien_Do_Thi_Cong_CPM_Gantt_Chart.xlsx`<br>+ `.xml` + `.mpp` + `.csv` | 46 công tác WBS, tính toán đường găng Critical Path Method chi tiết: ES, EF, LS, LF, Total Float TF, nhận diện 23 công tác găng. Kèm file MS Project XML/MPP. |
| **11** | `HOSO_KCS_NGHIEM_THU` | `11_Danh_Muc_KCS_43_Bien_Ban_Nghiem_Thu.xlsx`<br>+ `Ho_So_...docx` | Danh mục 43 biên bản nghiệm thu KCS theo NĐ 207/2026/NĐ-CP & TT 32/2026/TT-BXD, kèm file Word đầy đủ văn bản ký tá pháp lý. |
| **12** | `MAU_BIEN_BAN_KCS` | `12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx` | File Excel in A4 Portrait: Ô chọn số biên bản (1 - 43) để hàm VLOOKUP tự động lấy dữ liệu từ tab `DATA_CONG_TAC` ngay trong file, không lỗi link ngoài. |
| **13** | `MAU_BB_NGHIEM_THU_VAT_LIEU` | `13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx` | File Excel in A4 Portrait: Ô chọn mã vật liệu (1 - 10) để VLOOKUP tự lấy chứng chỉ CO/CQ, khối lượng lô từ tab `DATA_VAT_LIEU` nội bộ. |
| **14** | `MAU_BB_LAY_MAU_HIEN_TRUONG` | `14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx` | File Excel in A4 Portrait: Ô chọn tổ mẫu (1 - 10) để tự động nhảy lực nén phá hoại, cường độ R7, R28 và kết luận đạt chuẩn TCVN 3118 từ tab `DATA_NEN_MAU`. |

---

## 4. NGUYÊN TẮC THIẾT KẾ "ZERO BROKEN LINKS" CHO BỘ HỒ SƠ VI MÔ

Đối với các file mẫu in ấn độc lập (File 12, 13, 14):
- **Vấn đề thường gặp:** Nếu công thức trong mẫu in A4 trỏ sang file Master bằng đường dẫn tuyệt đối (ví dụ `='[Master.xlsx]Sheet'!A1`), khi người dùng gửi riêng file mẫu in A4 cho kỹ sư công trường hoặc tư vấn giám sát, file sẽ bị lỗi `#REF!` hoặc popup đòi update link ngoài.
- **Giải pháp chuẩn hóa:** Mỗi file vi mô chuyên sâu được thiết kế theo cấu trúc **Self-Contained (Tự chủ dữ liệu)**:
  1. Tab 1: `BIEN_BAN_A4` (Giao diện in ấn chuẩn khổ A4, căn lề 2cm, fit to 1 page).
  2. Tab 2: `DATA_NOI_BO` (`DATA_CONG_TAC`, `DATA_VAT_LIEU`, `DATA_NEN_MAU`): Chứa toàn bộ cơ sở dữ liệu dự án.
  3. Hàm `=VLOOKUP(...)` trỏ trực tiếp vào Tab 2 nội bộ trong cùng file $\rightarrow$ **100% không bao giờ bị lỗi gãy liên kết**.

---

## 5. KỊCH BẢN TỰ ĐỘNG HÓA THỰC THI (AUTOMATION PIPELINE)

Để tạo lập hoặc cập nhật đồng bộ toàn bộ 2 gói hồ sơ từ dữ liệu gốc, chỉ cần chạy một lệnh duy nhất:

```bash
python examples/build_14_micro_standalone_dossiers.py
```

### Các bước kịch bản tự động thực hiện:
1. Nạp và giải mã số liệu từ bản vẽ (`bang_so_lieu.json`) và bảng tổng hợp khối lượng (`THKL G9 - CHUẨN.xlsx`).
2. Thiết lập Gói 01 Macro: Kiểm tra 14 Sheet liên kết động, sao chép file tiến độ MS Project XML/MPP, cập nhật báo cáo thẩm định Audit 100/100.
3. Sinh tuần tự 14 bộ hồ sơ chuyên sâu vi mô vào thư mục `BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO/`.
4. Nhúng cấu trúc dữ liệu tự chủ và thiết lập vùng in ấn A4 chuẩn cho các file biểu mẫu.
5. Đồng bộ song song kết quả sang thư mục gốc dự án.
6. Chạy bộ kiểm toán `aec_audit_verifier.py` để xác nhận điểm 100/100 tuyệt đối trước khi bàn giao.
