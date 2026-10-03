# 📘 CẨM NANG HƯỚNG DẪN VIẾT PROMPT & CÂU LỆNH CHO HỆ THỐNG 23HG-AEC
### Dành cho Người mới bắt đầu (Quickstart Prompting & Command Guide)

---

## 🧭 1. Triết lý Vận hành: Công thức Prompt 4 Thành Tố (C-I-R-O)

Hệ thống **23HG-AEC-MultiAgent-System** vận hành theo nguyên lý **Toán học Xác định (Zero LLM Math)**: Mọi kết quả bóc tách hình học, tính tiền dự toán, tối ưu cắt thép, tiến độ CPM đều được giải bằng code Python thuần và thuật toán tối ưu (Google OR-Tools, NetworkX CPM), **không để AI đoán mò số liệu**.

Do đó, một Prompt hiệu quả cần tuân thủ cấu trúc **C-I-R-O**:
```text
┌────────────────────────────────────────────────────────────────────────┐
│ [C] CONTEXT (Vai trò)    : Bạn là Kỹ sư QS / Kỹ sư Cầu đường / KCS... │
│ [I] INPUT (Dữ liệu vào)  : File Excel/CAD/JSON hoặc thông số kích thước│
│ [R] RULES (Ràng buộc)    : Tiêu chuẩn kỹ thuật (TCVN, TT BXD, 40d...)  │
│ [O] OUTPUT (Đầu ra)      : File Excel công thức sống / Báo cáo / Sơ đồ │
└────────────────────────────────────────────────────────────────────────┘
```

---

## 🎯 2. Các Mẫu Prompt Thực Chiến Theo Từng Nghiệp Vụ

### 🔹 MẪU 1: Bóc tách Khối lượng Cầu đường & Kết cấu thép
> **Áp dụng khi:** Cần tính bê tông, ván khuôn mố trụ, dầm Super-T hoặc dầm cầu thép tấm chữ I.

```text
[VAI TRÒ]: Hãy đóng vai Kỹ sư Cầu đường và QS chuyên nghiệp của hệ thống 23HG-AEC.
[NHIỆM VỤ]: Bóc tách khối lượng hình học cho hạng mục: Trụ cầu xẻ nước mũi thuyền và Xà mũ vươn hẫng Hammerhead.
[THÔNG SỐ ĐẦU VÀO]:
- Bệ trụ: Dài 10.0m, Rộng 4.5m, Cao 2.0m, Mũi xẻ nước dài 2.25m (góc vát 45 độ).
- Thân trụ: 2 cột tròn đường kính D1.5m, Chiều cao 8.5m.
- Xà mũ Hammerhead: Dài 12.5m, Rộng 2.2m, Chiều cao tâm 1.8m, Chiều cao mút thừa 1.0m, Cánh hẫng vươn 3.5m mỗi bên.
[RÀNG BUỘC KỸ THUẬT]: 
- Tuân thủ TCVN 11823:2017 và Thông tư 38/2026/TT-BXD.
- Trừ giao thể tích chính xác giữa cột và xà mũ.
- Ván khuôn xà mũ tính riêng đáy, thành và vát hẫng đáy.
[ĐẦU RA MONG MUỐN]:
- Bảng tổng hợp thể tích bê tông (m3) và diện tích ván khuôn (m2).
- Diễn giải công thức toán học chi tiết từng cấu kiện.
```
*Lệnh CLI tương đương:*
```bash
python examples/demo_bridge_takeoff_from_qs_logic.py
```

---

### 🔹 MẪU 2: Tối ưu hóa Cắt thép 1D (Rebar 1D Cutting Stock)
> **Áp dụng khi:** Có danh mục thép cần cắt, muốn giảm phế liệu đề-xê $< 1.5\%$ bằng cây thép 11.7m.

```text
[VAI TRÒ]: Hãy đóng vai Kỹ sư Quản lý Cốt thép và Tối ưu hóa Vật tư.
[NHIỆM VỤ]: Giải bài toán tổ hợp cắt thép 1D (Cutting Stock) từ danh mục thanh sau:
- Thép D25, Mác CB400-V, Cây thép nguyên thương mại: 11.7m (11700mm).
- Thanh Mark 1: Chiều dài 11.700mm - Số lượng: 24 thanh.
- Thanh Mark 2: Chiều dài 7.500mm  - Số lượng: 48 thanh.
- Thanh Mark 3: Chiều dài 4.200mm  - Số lượng: 48 thanh.
[RÀNG BUỘC KỸ THUẬT]:
- Lưỡi cắt (kerf): 5mm.
- Mục tiêu: Tỷ lệ đề-xê phế thải <= 1.5% (đạt tiêu chuẩn Golden Rebar Pattern).
- Tự động bổ sung định mức dây thép buộc 1.5% theo Thông tư 38/2026/TT-BXD.
[ĐẦU RA MONG MUỐN]:
- Số cây thép 11.7m nguyên cần xuất kho.
- Sơ đồ cắt chi tiết từng cây (cây số mấy cắt những thanh nào, thừa bao nhiêu mm).
- Tỷ lệ phế thải tổng thể (%) và khối lượng dây thép buộc cần mua.
```
*Lệnh CLI tương đương:*
```bash
python -m tools.cutting_stock_solver --stock 11700 --demands "11700:24,7500:48,4200:48"
```

---

### 🔹 MẪU 3: Lập Dự toán Xây dựng & Tính Chi phí G_XD
> **Áp dụng khi:** Có bảng khối lượng BoQ, cần lập bảng dự toán chi phí xây dựng $G_{XD} = T + GT + TL + VAT$.

```text
[VAI TRÒ]: Hãy đóng vai Kỹ sư Định giá Xây dựng & Lập Dự toán (Cost Engineer).
[NHIỆM VỤ]: Đọc bảng khối lượng BoQ từ file [ten_file_boq.xlsx] và tính toán Tổng chi phí xây dựng G_XD.
[RÀNG BUỘC PHÁP LÝ & KỸ THUẬT]:
- Áp dụng Nghị định 207/2026/NĐ-CP, Thông tư 36/2026/TT-BXD và Thông tư 38/2026/TT-BXD.
- Công thức: G_XD = T + GT + TL + VAT.
  + Chi phí trực tiếp: T = VL + NC + M.
  + Chi phí gián tiếp: GT = Chi phí chung + Chi phí nhà tạm + Chi phí không xác định.
  + Thu nhập chịu thuế tính trước: TL = (T + GT) * Định mức %.
  + Thuế VAT: 8% hoặc 10% theo quy định hiện hành.
- BẢO TOÀN CÔNG THỨC SỐNG: 100% ô tính phải dùng công thức Excel (=C*D, =SUM), nghiêm cấm điền số tĩnh (số chết).
[ĐẦU RA MONG MUỐN]:
- Xuất file Excel dự toán phân cấp WBS chuẩn Bộ Xây dựng.
```
*Lệnh CLI tương đương:*
```bash
python run_state_graph.py --phase qs --excel "du_lieu_du_an/bang_boq.xlsx"
```

---

### 🔹 MẪU 4: Lập Tiến độ Thi công CPM & Kế hoạch Ca máy - Dầu Diesel
> **Áp dụng khi:** Cần tính toán đường găng Critical Path, nhân công, ca máy và tiêu hao nhiên liệu.

```text
[VAI TRÒ]: Hãy đóng vai Kỹ sư Điều độ Hiện trường & Quản trị Thiết bị Thi công.
[NHIỆM VỤ]: Lập tiến độ thi công mạng CPM và kế hoạch huy động máy thi công - dầu diesel cho dự án.
[DỮ LIỆU ĐẦU VÀO]: File tiến độ hoặc bảng tác vụ tại [duong_dan_file_tien_do.xlsx] (gồm Mã CV, Tên CV, Thời gian, Quan hệ FS/SS/FF và Ca máy dự kiến).
[RÀNG BUỘC KỸ THUẬT]:
- Tính toán đúng đường găng CPM (ES, EF, LS, LF, Total Float TF).
- Áp dụng định mức tiêu hao dầu diesel theo bảng định mức Vincons trong thư mục data/.
- Cân bằng tài nguyên máy đào, máy ủi, lu rung chống xung đột đỉnh tải.
[ĐẦU RA MONG MUỐN]:
- Bảng tiến độ Gantt Chart Native Excel với công tắc tương tác tại ô C11 (Nhập 1: hiện vạch tiến độ; Nhập 2: hiện nhân công; Nhập 3: hiện ca máy).
- Bảng cân bằng dầu Diesel từng ngày và cả đợt thi công.
```
*Lệnh CLI tương đương:*
```bash
python run_state_graph.py --phase cpm --excel "examples/HO_SO_CAU_KM19_529/260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"
```

---

### 🔹 MẪU 5: Lập Bảng Thanh toán Khối lượng Hoàn thành Mẫu 03a
> **Áp dụng khi:** Nghiệm thu giai đoạn, lập hồ sơ thanh toán A-B theo Nghị định 254/2025/NĐ-CP.

```text
[VAI TRÒ]: Hãy đóng vai Kỹ sư Quản lý Hợp đồng & Thanh toán (Payment Engineer).
[NHIỆM VỤ]: Lập bảng thanh toán khối lượng hoàn thành Mẫu 03a (theo Nghị định 254/2025/NĐ-CP) cho Đợt thanh toán số [X].
[DỮ LIỆU ĐẦU VÀO]: 
- Bảng khối lượng hợp đồng đã duyệt.
- Bảng khối lượng hoàn thành thực tế lũy kế đến kỳ này.
- Số tiền tạm ứng hợp đồng và mức thu hồi tạm ứng theo điều khoản thanh toán.
[RÀNG BUỘC KỸ THUẬT]:
- Tính toán đầy đủ: Khối lượng thực hiện kỳ này, lũy kế đến hết kỳ này, giá trị thanh toán, giảm trừ tạm ứng, giữ lại bảo hành (thường 5%).
- Không làm tròn tiền tệ trung gian, sử dụng Decimal chính xác từng đồng.
[ĐẦU RA MONG MUỐN]:
- Xuất bảng tính Mẫu 03a chuẩn mẫu Bộ Tài chính bằng file Excel công thức sống.
```
*Lệnh CLI tương đương:*
```bash
python run_state_graph.py --phase payment --excel "du_lieu_du_an/hop_dong_thanh_toan.xlsx"
```

---

### 🔹 MẪU 6: Xuất Trọn Bộ Hồ Sơ Công Nghiệp 3 Tầng (Hub & Spoke)
> **Áp dụng khi:** Muốn xuất trọn gói 14 hồ sơ vi mô chuyên sâu và 5 gói bàn giao phân quyền.

```text
[VAI TRÒ]: Hãy đóng vai Giám đốc Kỹ thuật & Quản trị Hệ thống 23HG-AEC.
[NHIỆM VỤ]: Kích hoạt Quy trình Xuất Hồ Sơ Công Nghiệp 3 Tầng (Industrial End-to-End Export Pipeline) cho dự án [Tên Dự Án].
[DỮ LIỆU ĐẦU VÀO]: Tệp Master Excel tại [duong_dan_master.xlsx].
[QUY CHUẨN XUẤT XƯỞNG]:
- TẦNG 1: Macro Master File (Bộ tính tổng hợp gốc).
- TẦNG 2: 14 Bộ hồ sơ vi mô độc lập (QS, BBS, Đơn giá, 03a, KCS, Ca máy, Cấp phối, Biểu đồ...).
- TẦNG 3: Hub & Spoke 5 Gói vệ tinh phân quyền (Gói Chủ đầu tư, Gói Tư vấn Giám sát, Gói Ban Chỉ huy, Gói Thầu phụ, Gói Đội thi công).
- QUALITY GATE: Bắt buộc chạy kiểm toán tự động trước khi xuất xưởng, đảm bảo 0 lỗi công thức (#REF!, #VALUE!, #NAME!).
[ĐẦU RA MONG MUỐN]:
- Đóng gói đầy đủ vào thư mục xuất xưởng và xuất báo cáo Quality Gate 100% ĐẠT.
```
*Lệnh CLI tương đương:*
```bash
python run_state_graph.py --export-all --excel "du_lieu_du_an/Master.xlsx" --export-dir "./HO_SO_XUAT_XUONG"
```

---

### 🔹 MẪU 7: Kiểm tra Cấp độ Tiến hóa (Level-Up) & Nạp Kinh Nghiệm Mới
> **Áp dụng khi:** Muốn xem hệ thống đã lên Level mấy, đạt bao nhiêu XP, hoặc nạp kinh nghiệm dự án mới.

```text
[VAI TRÒ]: Quản trị viên Kho Kinh nghiệm Hệ thống (AEC Experience Agent).
[NHIỆM VỤ]: 
1. Hiển thị Báo cáo Cấp độ hiện tại (Level), Điểm kinh nghiệm (XP), Danh hiệu nghề nghiệp và Tiến độ lên cấp.
2. Liệt kê các Mẫu cắt thép vàng, Quy tắc miễn dịch lỗi (Immunity Rules) và Kỹ năng mới đã được phê duyệt.
```
*Lệnh CLI tương đương:*
```bash
# Xem báo cáo Cấp độ (Level & XP)
python run_state_graph.py --level

# Chạy kiểm thử toàn bộ hệ thống (227 tests)
python -m unittest discover -s tests -t .
```

---

## ⚡ 3. Bảng Tra Cứu Lệnh Dòng Lệnh Nhanh (CLI Cheat Sheet)

| Nhu cầu nghiệp vụ | Câu lệnh Terminal chạy ngay |
| :--- | :--- |
| **Chạy demo toàn bộ 8 pha** | `python run_state_graph.py --demo` |
| **Xem Level-Up & Cấp bậc AI** | `python run_state_graph.py --level` |
| **Bóc tách dầm cầu & Trụ xẻ nước** | `python examples/demo_bridge_takeoff_from_qs_logic.py` |
| **Tối ưu hóa cắt thép 1D** | `python -m tools.cutting_stock_solver --help` |
| **Tính tiến độ CPM & Ca máy** | `python run_state_graph.py --phase cpm --excel "<file.xlsx>"` |
| **Lập dự toán & Tính G_XD** | `python run_state_graph.py --phase qs --excel "<file.xlsx>"` |
| **Lập bảng thanh toán 03a** | `python run_state_graph.py --phase payment --excel "<file.xlsx>"` |
| **Xuất trọn bộ hồ sơ 3 tầng** | `python run_state_graph.py --export-all --excel "<file.xlsx>"` |
| **Kiểm toán lỗi file Excel tĩnh** | `python -m tools.audit_excels_static "<thư_mục_excel>"` |
| **Chạy kiểm thử 227 unit tests** | `python -m unittest discover -s tests -t .` |

---

## ⚠️ 4. Những Sai Lầm Phổ Biến Cần Tránh Khi Viết Prompt

1. ❌ **Không yêu cầu AI "tính nhẩm":** 
   - *Sai:* "Hãy tính xem dầm cầu này hết bao nhiêu tiền?" $\to$ Dễ dính ảo giác LLM.
   - *Đúng:* "Hãy dùng module `tools.civil_and_bridge_takeoff_engine` và `tools.qs_export` để tính khối lượng và xuất bảng dự toán G_XD bằng công thức sống."
2. ❌ **Không bỏ sót đơn vị đo:**
   - Khi nhập thép tấm hay kích thước hình học, luôn nói rõ: Chiều dài (mm hay m), Khối lượng (kg hay tấn).
3. ❌ **Quên mác thép khi cắt thép:**
   - Hệ thống tối ưu hóa cắt thép phân tách nghiêm ngặt theo từng đường kính ($\Phi$) và mác thép (CB240-T, CB300-V, CB400-V, CB500-V). Không gom chung các loại mác thép khác nhau vào cùng một mẻ cắt.
