# 📘 Cẩm nang viết prompt và chạy lệnh cho người mới

> Mọi lệnh trong tài liệu này đã được chạy thử trên Windows với các file mẫu có sẵn trong repo (`templates/`, `examples/`) ngày 03/10/2026. Nếu một lệnh báo lỗi, hãy chạy `python run_state_graph.py --help` để xem tham số mới nhất.

---

## 1. Hiểu đúng hệ thống trước khi ra lệnh

- **Mọi phép tính là code Python xác định**, không do AI "tính nhẩm": OR-Tools cho cắt thép, CPM cho tiến độ, `Decimal` cho tiền. AI trợ lý (ví dụ Antigravity) chỉ nên **gọi đúng công cụ** rồi đọc kết quả.
- **Mặc định chỉ dùng dữ liệu thật.** Thiếu file đầu vào thì hệ thống dừng và báo cần gì. Cờ `--demo` chạy bằng dữ liệu mẫu, kết quả luôn có cảnh báo *"KHÔNG DÙNG CHO HỒ SƠ THẬT"*.
- **Human Gate:** mặc định (`--human-gate cli`) hệ thống dừng hỏi kỹ sư duyệt. `--human-gate auto` tự duyệt và **chỉ dùng để chạy thử**.

---

## 2. Công thức prompt 4 phần (C-I-R-O)

```text
[C] VAI TRÒ     : Kỹ sư QS / kỹ sư cầu / kỹ sư KCS / kỹ sư tiến độ...
[I] ĐẦU VÀO     : đường dẫn file thật (BBS, bảng QS, tiến độ XML...) hoặc thông số kích thước có đơn vị
[R] RÀNG BUỘC   : tiêu chuẩn áp dụng, mác thép, tỷ lệ chi phí, lưỡi cắt...; "chỉ dùng công cụ của repo, không tự tính"
[O] ĐẦU RA      : file xuất (.xlsx/.csv), báo cáo cần đọc, những gì phải cảnh báo
```

Câu nên luôn thêm vào cuối prompt:

> *"Chỉ dùng lệnh/hàm có trong repo 23HG. Chạy thật và dán kết quả. Nếu thiếu dữ liệu thì hỏi lại, không tự điền số. Ghi rõ mọi cảnh báo hệ thống in ra."*

---

## 3. Mẫu prompt theo nghiệp vụ (kèm lệnh đã chạy thử)

### 3.1. Tối ưu cắt thép từ BBS thật

```text
[VAI TRÒ] Kỹ sư quản lý cốt thép.
[ĐẦU VÀO] BBS: examples/HO_SO_CONG_HOP_TUYEN_A5/BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO/04_Thong_Ke_Thep_Chi_Tiet_BBS_Cong_A5.xlsx
[RÀNG BUỘC] Cây 11,7 m, lưỡi cắt 3 mm, tách theo đường kính và mác thép. Không tự nối thép nếu tôi chưa cho phép.
[ĐẦU RA] Số cây cần mua, cận dưới, đề-xê từng Ø, phiếu cắt CSV cho xưởng.
Chỉ dùng lệnh của repo, chạy thật và dán kết quả.
```

```bash
python run_state_graph.py --phase rebar --bbs "duong_dan/BBS.xlsx" --kerf-mm 3 --cut-plan-out phieu_cat.csv
# Thêm --splice nếu muốn hệ thống đề xuất phương án nối tận dụng đầu thừa (kỹ thuật phải duyệt)
# Thêm --rebarcut-out ket_qua.xlsx để xuất theo bố cục RebarCut
```

Đọc kết quả: nếu đề-xê > 1,5% nhưng **số cây bằng cận dưới** thì hệ thống đã chứng minh không thể dùng ít cây hơn. Hao hụt khi đó là do chiều dài thanh trong BBS, không phải do solver.

### 3.2. Dự toán G_XD từ bảng QS/BOQ

```text
[VAI TRÒ] Kỹ sư định giá.
[ĐẦU VÀO] Bảng QS: templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx
[RÀNG BUỘC] Chi phí chung 5,1% T; nhà tạm 1,2% T; KXĐ 1,0% T; TL 5,5% (T+GT); VAT 8%.
            Nếu file đã có sheet tổng hợp thì đối chiếu và báo mọi chênh lệch.
[ĐẦU RA] File Excel tổng hợp G_XD + chi tiết công tác.
```

```bash
# Bước 1: chỉ kiểm tra file đọc được không (không tính, không ghi file)
python run_state_graph.py --check-inputs --qs "bang_qs.xlsx"

# Bước 2: tính và xuất
python run_state_graph.py --phase qs --qs "bang_qs.xlsx" --rate-chung 5.1 --rate-nha-tam 1.2 --rate-kxd 1.0 --rate-tl 5.5 --vat 8 --qs-out G_XD.xlsx
```

Với file mẫu, hệ thống báo: *"G_XD ghi trong file 33.974.615.853 ≠ tính lại 33.356.895.564"*. Đây là chức năng đối chiếu đang hoạt động, không phải lỗi. Kỹ sư cần kiểm tra chênh lệch này.

### 3.3. Thanh toán Mẫu 03a (NĐ 254/2025/NĐ-CP)

```text
[VAI TRÒ] Kỹ sư thanh toán hợp đồng.
[ĐẦU VÀO] Bảng QS hợp đồng: <file>; khối lượng thực hiện (lũy kế kỳ trước, kỳ này): <file>
[RÀNG BUỘC] Đơn giá QS là đơn giá hợp đồng trước thuế; thu hồi tạm ứng 20%; giữ lại 5%; kỳ 01.
[ĐẦU RA] File Mẫu 03a.
```

```bash
python run_state_graph.py --phase payment --qs "hop_dong.xlsx" --progress "kl_thuc_hien.xlsx" --price-basis contract --advance-recovery-pct 20 --retention-pct 5 --period 01 --payment-out 03a_ky01.xlsx

# Chạy thử bằng dữ liệu mẫu (không dùng cho hồ sơ thật):
python run_state_graph.py --demo --phase payment --payment-out 03a_thu.xlsx
```

### 3.4. Tiến độ CPM từ MS Project / Excel / CSV

```text
[VAI TRÒ] Kỹ sư tiến độ.
[ĐẦU VÀO] File tiến độ: templates/Tien_Do_Thi_Cong_Cau_Km19+529.080.xml
[RÀNG BUỘC] Nghỉ Chủ nhật; nghỉ Tết 2027-02-05 đến 2027-02-12.
[ĐẦU RA] Bảng ES/EF/LS/LF/dự trữ, danh sách công việc găng.
```

```bash
python run_state_graph.py --phase schedule --schedule "tien_do.xml" --non-working-days cn --holidays "2027-02-05:2027-02-12" --schedule-out cpm.csv
```

### 3.5. Đo bóc khối lượng từ bảng cấu kiện (có diễn giải)

```bash
# Mẫu cột đầu vào: templates/Mau_dau_vao_do_boc.csv
python run_state_graph.py --takeoff "bang_cau_kien.csv" --takeoff-out do_boc.xlsx
```

### 3.6. Đánh giá phiếu thí nghiệm và điểm dừng kỹ thuật

```bash
# Mẫu: templates/Phieu_thi_nghiem_mau.csv
python run_state_graph.py --phase qaqc --lab "phieu_thi_nghiem.csv" --lab-out danh_gia_TN.xlsx
```

### 3.7. Xuất trọn bộ hồ sơ 3 tầng (Master, hồ sơ vi mô, gói Hub & Spoke)

```bash
python run_state_graph.py --export-all --excel "Master.xlsx" --export-dir HO_SO_XUAT --project-name "Tên dự án"
```

Cuối quá trình xuất, Quality Gate tự quét lỗi công thức (`#REF!`, `#VALUE!`...) và ghi `DISPATCH_MANIFEST.json` kèm MD5.

### 3.10. Kiểm toán & Rà soát Hồ sơ QLCL theo 6 Bất biến Lõi

```text
[VAI TRÒ] Kỹ sư Trưởng QLCL & Kiểm toán KCS.
[ĐẦU VÀO] File Excel hồ sơ nghiệm thu thực tế: <đường dẫn file .xlsx hoặc .xlsm>
[RÀNG BUỘC] Áp dụng 6 Bất biến Lõi TCVN:
            1. Khối lượng bảo toàn (A x B x C = V).
            2. Chu kỳ bê tông: Tháo cốp pha >= 2 ngày, Nghiệm thu hoàn thành cấu kiện >= 28 ngày (R28), PYC trước NT >= 24h.
            3. Tuyệt đối không nghiệm thu ngoài trời vào ngày Tết Nguyên Đán, Tết DL, Lễ 30/4-1/5, Quốc khánh.
            4. Phải kẹp đủ chùm checklist (CL, PLKL, KTĐBT, TVKM).
[ĐẦU RA] Báo cáo chấm điểm chất lượng (thang 100) và danh sách các lỗi đá ngày / lỗi hình học cần sửa.
```

```bash
python run_state_graph.py --audit-kcs "duong_dan_file_ho_so.xlsx"
# Hoặc chạy qua skill AEC-QLCL:
python skills/aec-qlcl/scripts/audit_qa_register.py "duong_dan_file_ho_so.xlsx"
```

### 3.11. Lập trọn bộ Hồ sơ Nghiệm thu QLCL / KCS Thực chiến (1 Phôi - N Biên bản)

```text
[VAI TRÒ] Kỹ sư QLCL / KCS công trường.
[ĐẦU VÀO] Hạng mục: <Tên công trình/hạng mục, ví dụ: Cầu Khai Hoang Km14+363.65 hoặc Kè bờ tả H0+00 - H0+20>;
          Kích thước hình học thực tế (Dài A, Rộng B, Cao C); Mác bê tông, Nhóm cốt thép;
          Thời gian khởi công: YYYY-MM-DD đến YYYY-MM-DD.
[RÀNG BUỘC] 1. CSDL gốc: Kích thước A x B x C liên kết bằng CÔNG THỨC SỐNG sang thể tích, diện tích ván khuôn, số chuyến xe bê tông. ZERO SỐ CHẾT.
            2. Đồ thị thời gian kết cấu (DAG): Đào móng -> Cốt thép chờ -> Cốp pha + Thép -> Đổ BT -> Tháo cốp pha (+2 ngày) -> Nghiệm thu hoàn thành (+28 ngày có R28).
            3. Né 100% ngày nghỉ Tết Nguyên Đán và các ngày nghỉ lễ quốc gia.
            4. Xuất theo cơ chế "1 Phôi in - N Biên bản": Sheet phôi A4 nhảy dữ liệu theo con trỏ Pointer; kẹp đủ chùm Checklist hình học và Sheet Nhật ký thi công đồng bộ.
[ĐẦU RA] File Excel Master QLCL hoàn chỉnh, in ấn vừa vặn trang A4 Portrait, sẵn sàng ký duyệt.
```

### 3.12. Bóc tách sơ bộ cầu (nguyên mẫu)


```text
[VAI TRÒ] Kỹ sư cầu.
[NHIỆM VỤ] Dùng BridgePierParams + BridgePierEngine.compute_takeoff trong tools/civil_and_bridge_takeoff_engine.py
           để ước tính bê tông, ván khuôn trụ: bệ 7,0 x 3,6 x 1,8 m; 2 cột D1,5 m cao 8,5 m; xà mũ 9,6 x 1,8 m,
           cao 1,6 m tại tim, 1,0 m tại mút.
[LƯU Ý] Đây là ước tính sơ bộ, ghi rõ các giả định trong code (xem skills/skill-civil-to-bridge-takeoff/SKILL.md).
```

```bash
python examples/demo_bridge_takeoff_from_qs_logic.py
```

### 3.9. Kiểm tra hệ thống

```bash
python run_state_graph.py --solver-test               # OR-Tools và CPM hoạt động không
python run_state_graph.py --demo                      # chạy thử đủ các pha bằng dữ liệu mẫu
python run_state_graph.py --level                     # báo cáo kho kinh nghiệm (Level/XP)
python -m tools.audit_excels_static "thu_muc_excel"   # quét lỗi công thức Excel (không cần Excel)
python -m unittest discover -s tests -t .             # toàn bộ unit test
```

---

## 4. Bảng tra nhanh

| Việc cần làm | Lệnh |
|---|---|
| Kiểm tra file đầu vào | `python run_state_graph.py --check-inputs --qs <file> --bbs <file> ...` |
| Cắt thép | `--phase rebar --bbs <file> --cut-plan-out <csv>` |
| Dự toán G_XD | `--phase qs --qs <file> --rate-chung .. --rate-nha-tam .. --rate-kxd .. --rate-tl .. --vat .. --qs-out <xlsx>` |
| Thanh toán 03a | `--phase payment --qs <file> --progress <file> --price-basis contract --advance-recovery-pct .. --retention-pct .. --payment-out <xlsx>` |
| Tiến độ CPM | `--phase schedule --schedule <xml/xlsx/csv> --schedule-out <csv>` |
| Đo bóc bảng cấu kiện | `--takeoff <csv/xlsx> --takeoff-out <xlsx>` |
| Phiếu thí nghiệm | `--phase qaqc --lab <file> --lab-out <xlsx>` |
| Xuất hồ sơ 3 tầng | `--export-all --excel <Master.xlsx> --export-dir <thư mục>` |

> ⚠️ **Lỗi đã biết:** `--phase fleet` / `--phase dispatch` có trong `--help` nhưng Supervisor chưa xử lý pha này nên dừng ngay với thông báo *"Phase không xác định"*. Kế hoạch ca máy hiện chỉ có trong các script dựng hồ sơ ở `examples/`.

---

## 5. Sai lầm thường gặp

1. **Để AI tự tính hoặc tự điền số.** Hãy yêu cầu chạy công cụ của repo và dán kết quả thật.
2. **Thiếu đơn vị.** Luôn ghi rõ mm hay m, kg hay tấn, % của T hay của (T+GT).
3. **Dùng `--demo` hoặc `--human-gate auto` cho hồ sơ thật.** Hai cờ này chỉ để chạy thử.
4. **Gom chung mác thép.** Solver tách theo Ø và mác; BBS phải ghi mác thép đúng từng thanh.
5. **Coi kết quả là hồ sơ pháp lý.** Kết quả dự toán, thanh toán, nghiệm thu phải được kỹ sư rà soát trước khi ký.
