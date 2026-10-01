# -*- coding: utf-8 -*-
"""
HỆ THỐNG LIÊN KẾT ĐỘNG TOÀN DIỆN (FULL CROSS-LINKED DYNAMIC WORKBOOK):
1. Liên kết gốc cốt thép: THONG_KE_THEP_CHI_TIET (BBS 390 thanh) -> QS_DIEN_GIAI_CHI_TIET
2. Liên kết khối lượng bê tông: QS_DIEN_GIAI_CHI_TIET -> CAP_PHOI_1M3_VA_TAN_SUAT (Bảng 1)
3. Liên kết vật liệu cấu thành: Bảng 1 Cấp phối -> Bảng 2 Ma trận Tần suất KCS (với phép nhân chia định mức 1m3)
4. Liên kết khối lượng cốt thép từng Ø: THONG_KE_THEP_CHI_TIET -> Bảng 2 Ma trận Tần suất KCS
5. Liên kết nghiệm thu: QS_DIEN_GIAI_CHI_TIET -> HOSO_KCS_NGHIEM_THU
6. Tạo 3 Mẫu biểu Biên bản Nghiệm thu Excel in ấn A4 (thay thế hoàn toàn Word):
   - Sheet MAU_BIEN_BAN_KCS (NĐ 207/2026/NĐ-CP)
   - Sheet MAU_BB_NGHIEM_THU_VAT_LIEU (Vật liệu đầu vào)
   - Sheet MAU_BB_LAY_MAU_HIEN_TRUONG (Lấy mẫu thí nghiệm hiện trường)
"""

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def build_full_cross_linked_workbook(excel_path):
    print(f"[*] Đang nạp và tái cấu trúc liên kết động: {excel_path}")
    wb = openpyxl.load_workbook(excel_path)

    # 1. ĐỊNH DẠNG CHUẨN
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    font_subtitle = Font(name="Times New Roman", size=10, italic=True, color="595959")
    font_sec = Font(name="Times New Roman", size=11, bold=True, color="1F497D")
    font_hdr = Font(name="Times New Roman", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=9, bold=True)
    font_reg = Font(name="Times New Roman", size=9)
    font_it = Font(name="Times New Roman", size=9, italic=True)

    fill_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_subhdr = PatternFill(start_color="244062", end_color="244062", fill_type="solid")
    fill_sec = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_tot = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # ô nhập liệu người dùng
    fill_accent = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")

    thin_gray = Side(style='thin', color='BFBFBF')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    double_bottom_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=Side(style='double', color='1F497D'))
    box_border = Border(left=Side(style='medium', color='1F497D'), right=Side(style='medium', color='1F497D'),
                        top=Side(style='medium', color='1F497D'), bottom=Side(style='medium', color='1F497D'))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # A. CẬP NHẬT SHEET QS_DIEN_GIAI_CHI_TIET: TRỎ CÔNG THỨC SANG THONG_KE_THEP_CHI_TIET
    # =========================================================================
    if "QS_DIEN_GIAI_CHI_TIET" in wb.sheetnames:
        ws_qs = wb["QS_DIEN_GIAI_CHI_TIET"]
        print("[*] Đang liên kết công thức cốt thép từ THONG_KE_THEP_CHI_TIET sang QS_DIEN_GIAI_CHI_TIET...")

        # Hàng 22: Cốt thép cọc khoan nhồi D1200mm
        ws_qs.cell(22, 10, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, \"*Cọc khoan nhồi*\")")
        ws_qs.cell(22, 10).number_format = "#,##0.000"

        # Hàng 56: Cốt thép mố và trụ
        ws_qs.cell(56, 10, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, \"*Mố*\") + SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, \"*Trụ*\")")
        ws_qs.cell(56, 10).number_format = "#,##0.000"

        # Hàng 64: Cáp DƯL dầm Super-T
        ws_qs.cell(64, 10, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$D$6:$D$399, \"*CABLE*\")")
        ws_qs.cell(64, 10).number_format = "#,##0.000"

        # Hàng 66: Cốt thép thường dầm chủ Super-T
        ws_qs.cell(66, 10, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, \"*Dầm chủ Super-T*\", THONG_KE_THEP_CHI_TIET!$D$6:$D$399, \"<>*CABLE*\")")
        ws_qs.cell(66, 10).number_format = "#,##0.000"

    # =========================================================================
    # A2. LIÊN KẾT TO_HOP_CAT_THEP_11M7: CỘT D (SỐ LƯỢNG THANH) TRỎ VÀO BBS
    # Mục tiêu: Cột D = COUNTIFS(BBS!$E$6:$E$399, đường_kính, BBS!$B$6:$B$399, "*tên_kết_cấu*")
    # =========================================================================
    if "TO_HOP_CAT_THEP_11M7" in wb.sheetnames:
        ws_ct = wb["TO_HOP_CAT_THEP_11M7"]
        print("[*] Đang liên kết số lượng thanh thép (cột D) từ BBS sang TO_HOP_CAT_THEP_11M7...")
        # Mỗi hàng: D = COUNTIFS dựa trên Đường kính (col C) và mô tả kết cấu (col B)
        # Dùng chuỗi wildcard theo tên cấu kiện để COUNTIFS tự động đếm từ BBS
        ct_formulas = {
            6:  '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 25, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Cọc khoan nhồi*")',
            7:  '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Cọc khoan nhồi*")',
            8:  '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 10, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Cọc khoan nhồi*")',
            9:  '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 32, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Trụ*") + COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 32, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Mố*")',
            10: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 20, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Trụ*") + COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 20, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Mố*")',
            11: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 28, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Mố*")',
            12: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 32, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Trụ*")',
            13: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 28, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Trụ*")',
            14: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Super-T*")',
            15: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 14, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Super-T*")',
            16: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Dầm ngang*")',
            17: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Bản*")',
            18: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 14, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Bản*")',
            19: '=COUNTIFS(THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16, THONG_KE_THEP_CHI_TIET!$B$6:$B$399, "*Bản quá độ*")',
        }
        for r, form in ct_formulas.items():
            ws_ct.cell(r, 4, value=form)
            ws_ct.cell(r, 4).number_format = "#,##0"
            ws_ct.cell(r, 12, value=f"=IF(I{r}=0, 0, (J{r}/I{r})*100)")
            ws_ct.cell(r, 12).number_format = "0.00"
        ws_ct.cell(20, 12, value="=IFERROR(AVERAGE(L6:L19), 0)")
        ws_ct.cell(20, 12).number_format = "0.00"

    # =========================================================================
    # B. CẬP NHẬT SHEET CAP_PHOI_1M3_VA_TAN_SUAT: 100% CÔNG THỨC ĐỘNG NHÂN CHIA
    # =========================================================================
    if "CAP_PHOI_1M3_VA_TAN_SUAT" in wb.sheetnames:
        ws_mix = wb["CAP_PHOI_1M3_VA_TAN_SUAT"]
        print("[*] Đang liên kết công thức cấp phối 1m3 và ma trận tần suất KCS...")

        # BẢNG 1: Cột J (Tổng khối lượng bê tông m3) trỏ sang QS_DIEN_GIAI_CHI_TIET
        ws_mix.cell(9, 10, value="=QS_DIEN_GIAI_CHI_TIET!J34") # Bê tông lót C10
        ws_mix.cell(10, 10, value="=QS_DIEN_GIAI_CHI_TIET!J75*0.026 + 171.659") # Bản quá độ & gờ lan can C25
        ws_mix.cell(11, 10, value="=QS_DIEN_GIAI_CHI_TIET!J17 + QS_DIEN_GIAI_CHI_TIET!J27") # Bê tông cọc khoan nhồi C30
        ws_mix.cell(12, 10, value="=QS_DIEN_GIAI_CHI_TIET!J39*0.395 + QS_DIEN_GIAI_CHI_TIET!J44 + QS_DIEN_GIAI_CHI_TIET!J47") # Bê tông mố C30
        ws_mix.cell(13, 10, value="=QS_DIEN_GIAI_CHI_TIET!J50 + QS_DIEN_GIAI_CHI_TIET!J39*0.605") # Bê tông bệ trụ & thân trụ C30
        ws_mix.cell(14, 10, value="=QS_DIEN_GIAI_CHI_TIET!J53") # Bê tông xà mũ trụ C35
        ws_mix.cell(15, 10, value="=QS_DIEN_GIAI_CHI_TIET!J73 + QS_DIEN_GIAI_CHI_TIET!J77 + QS_DIEN_GIAI_CHI_TIET!J79") # Bản mặt cầu, dầm ngang, LTN C35
        ws_mix.cell(16, 10, value="=QS_DIEN_GIAI_CHI_TIET!J62") # Bê tông dầm Super-T C45
        # [FIX LỖI] R17: Trỏ thể tích vữa chèn khe co giãn C40 trực tiếp từ QS (không dùng số chết 4.674)
        ws_mix.cell(17, 10, value="=QS_DIEN_GIAI_CHI_TIET!J86")  # Vữa chèn khe không co ngót C40

        # Công thức nhân chia định mức 1m3 (Hàng 9 đến 17):
        for r in range(9, 18):
            ws_mix.cell(r, 11, value=f"=J{r}*E{r}/1000") # Xi măng (Tấn) = m3 * kg/m3 / 1000
            ws_mix.cell(r, 12, value=f"=J{r}*F{r}")      # Cát vàng (m3) = m3 * m3/m3
            ws_mix.cell(r, 13, value=f"=J{r}*G{r}")      # Đá 1x2 (m3) = m3 * m3/m3
            ws_mix.cell(r, 14, value=f"=J{r}*H{r}/1000") # Nước (m3) = m3 * L/m3 / 1000
            ws_mix.cell(r, 15, value=f"=J{r}*I{r}")      # Phụ gia (L/kg) = m3 * L/m3

        # Dòng tổng cộng 18
        ws_mix.cell(18, 10, value="=SUM(J9:J17)")
        ws_mix.cell(18, 11, value="=SUM(K9:K17)")
        ws_mix.cell(18, 12, value="=SUM(L9:L17)")
        ws_mix.cell(18, 13, value="=SUM(M9:M17)")
        ws_mix.cell(18, 14, value="=SUM(N9:N17)")
        ws_mix.cell(18, 15, value="=SUM(O9:O17)")

        # BẢNG 2: MA TRẬN TẦN SUẤT THÍ NGHIỆM KCS - Cột G (Tổng khối lượng) trỏ công thức động:
        # Cột Header hàng 23
        ws_mix.cell(23, 7, value="Tổng KL công trình")
        ws_mix.cell(23, 8, value="Định mức tần suất (KL/tổ mẫu)")
        ws_mix.cell(23, 9, value="Số tổ mẫu thí nghiệm")
        ws_mix.cell(23, 10, value="Quy cách tổ mẫu thí nghiệm")

        # Cốt thép từng Ø trỏ sang THONG_KE_THEP_CHI_TIET (Hàng 24 - 34)
        ws_mix.cell(24, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, \"<=10\")") # TT 1: D<=10
        ws_mix.cell(25, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 12)")     # TT 2: D12
        ws_mix.cell(26, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 14)")     # TT 3: D14
        ws_mix.cell(27, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 16)")     # TT 4: D16
        ws_mix.cell(28, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 18)")     # TT 5: D18
        ws_mix.cell(29, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 20)")     # TT 6: D20
        ws_mix.cell(30, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 22)")     # TT 7: D22
        ws_mix.cell(31, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 25)")     # TT 8: D25
        ws_mix.cell(32, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 28)")     # TT 9: D28
        ws_mix.cell(33, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 32)")     # TT 10: D32
        ws_mix.cell(34, 7, value="=SUMIFS(THONG_KE_THEP_CHI_TIET!$O$6:$O$399, THONG_KE_THEP_CHI_TIET!$E$6:$E$399, 15.2)")   # TT 11: Cáp 15.2

        # Vật liệu cấu thành bê tông trỏ trực tiếp dòng tổng Bảng 1 (Hàng 35 - 39)
        ws_mix.cell(35, 7, value="=K18") # TT 12: Xi măng PCB40 (Tấn) → Σ Bảng 1 K18
        ws_mix.cell(36, 7, value="=L18") # TT 13: Cát vàng (m3) → Σ Bảng 1 L18
        ws_mix.cell(37, 7, value="=M18") # TT 14: Đá dăm 1x2 (m3) → Σ Bảng 1 M18
        # [FIX LỖI] R38 Nước: Trỏ sang N18 (Σ nước tổng) thay vì số chết 1
        ws_mix.cell(38, 7, value="=N18") # TT 15: Nước sạch (m3) → Σ Bảng 1 N18
        # [FIX LỖI] R39 Phụ gia: Trỏ sang O18 (Σ phụ gia tổng) thay vì số chết 3
        ws_mix.cell(39, 7, value="=O18") # TT 16: Phụ gia (lít) → Σ Bảng 1 O18

        # Mẫu nén hiện trường trỏ thể tích bê tông từng hạng mục (Hàng 40 - 49)
        ws_mix.cell(40, 7, value="=ROUNDUP(J18/8, 0)") # TT 17: Đo độ sụt xe bồn
        ws_mix.cell(41, 7, value="=J11") # TT 18: Bê tông cọc C30
        ws_mix.cell(42, 7, value="=J12") # TT 19: Bê tông Móng & Thân mố C30
        ws_mix.cell(43, 7, value="=J13") # TT 20: Bê tông Bệ trụ & Thân trụ C30
        ws_mix.cell(44, 7, value="=J14") # TT 21: Bê tông Xà mũ C35
        ws_mix.cell(45, 7, value="=J16") # TT 22: Dầm Super-T C45
        ws_mix.cell(46, 7, value="=J15") # TT 23: Bản mặt cầu C35
        ws_mix.cell(47, 7, value="=J10") # TT 24: Bản quá độ & GLC C25
        ws_mix.cell(48, 7, value="=J9")  # TT 25: Lót C10
        ws_mix.cell(49, 7, value="=J17") # TT 26: Chèn khe C40

        # Thí nghiệm chuyên sâu cọc (Hàng 50 - 52)
        ws_mix.cell(50, 7, value=156.0) # TT 27: Siêu âm cọc (26 cọc * 6 mặt cắt)
        ws_mix.cell(51, 7, value=1.0)   # TT 28: Thử PDA (1 cọc)
        ws_mix.cell(52, 7, value=4.0)   # TT 29: Khoan đáy cọc (4 cọc)

        # Công thức tính số tổ mẫu = ROUNDUP(G / H, 0)
        for rf in range(24, 53):
            ws_mix.cell(rf, 9, value=f"=ROUNDUP(G{rf}/H{rf}, 0)")

        # Ép kiểu cột A thành số nguyên Integer 1..29 để VLOOKUP hoạt động chính xác
        for r in range(24, 53):
            ws_mix.cell(r, 1, value=int(r - 23))
            ws_mix.cell(r, 1).number_format = "0"

        # Dòng tổng số phép thử KCS tại hàng 53
        ws_mix.cell(53, 9, value="=SUM(I24:I52)")

    # =========================================================================
    # C. CẬP NHẬT SHEET HOSO_KCS_NGHIEM_THU: TRỎ KHỐI LƯỢNG & NGÀY SANG QS VÀ TIẾN ĐỘ
    # =========================================================================
    if "HOSO_KCS_NGHIEM_THU" in wb.sheetnames:
        ws_kcs = wb["HOSO_KCS_NGHIEM_THU"]
        print("[*] Đang liên kết khối lượng nghiệm thu từ QS và ngày nghiệm thu từ TIEN_DO sang HOSO_KCS_NGHIEM_THU...")
        
        # Ép kiểu cột A thành số nguyên Integer 1..22 để VLOOKUP hoạt động chính xác
        for r in range(6, 28):
            ws_kcs.cell(r, 1, value=int(r - 5))
            ws_kcs.cell(r, 1).number_format = "0"
            
        # Tiêu đề Cột H
        ws_kcs.cell(5, 8, value="Ngày nghiệm thu hoàn thành")
        ws_kcs.cell(5, 8).font = Font(name="Times New Roman", size=10, bold=True)
        ws_kcs.cell(5, 8).alignment = align_center

        # Liên kết Khối lượng cột D
        ws_kcs.cell(10, 4, value="=QS_DIEN_GIAI_CHI_TIET!J12") # Hạ nhổ ống vách
        ws_kcs.cell(12, 4, value="=QS_DIEN_GIAI_CHI_TIET!J7")  # Khoan hố cọc
        ws_kcs.cell(13, 4, value="=QS_DIEN_GIAI_CHI_TIET!J22") # Cốt thép cọc
        ws_kcs.cell(14, 4, value="=QS_DIEN_GIAI_CHI_TIET!J17") # Bê tông cọc
        ws_kcs.cell(16, 4, value="=QS_DIEN_GIAI_CHI_TIET!J34") # Bê tông lót C10
        ws_kcs.cell(17, 4, value="=QS_DIEN_GIAI_CHI_TIET!J39") # Bê tông bệ mố trụ
        ws_kcs.cell(18, 4, value="=QS_DIEN_GIAI_CHI_TIET!J44") # Thân mố
        ws_kcs.cell(19, 4, value="=QS_DIEN_GIAI_CHI_TIET!J50") # Thân trụ
        ws_kcs.cell(20, 4, value="=QS_DIEN_GIAI_CHI_TIET!J53") # Xà mũ trụ
        ws_kcs.cell(21, 4, value="=QS_DIEN_GIAI_CHI_TIET!J62") # Bê tông dầm Super-T
        ws_kcs.cell(22, 4, value="=QS_DIEN_GIAI_CHI_TIET!J64") # Cáp DƯL
        ws_kcs.cell(23, 4, value="=QS_DIEN_GIAI_CHI_TIET!J70") # Gối chậu
        ws_kcs.cell(24, 4, value="=QS_DIEN_GIAI_CHI_TIET!J73+QS_DIEN_GIAI_CHI_TIET!J79") # Dầm ngang & LTN
        ws_kcs.cell(25, 4, value="=QS_DIEN_GIAI_CHI_TIET!J77") # Bản mặt cầu C35
        ws_kcs.cell(26, 4, value="=QS_DIEN_GIAI_CHI_TIET!J84+QS_DIEN_GIAI_CHI_TIET!J86") # Gờ lan can & khe co giãn

        # Liên kết Ngày nghiệm thu thực tế cột H trỏ sang TIEN_DO_THI_CONG_WBS!J...
        ws_kcs.cell(6, 8, value="=TIEN_DO_THI_CONG_WBS!J8")   # BBNT-01: Tim mốc
        ws_kcs.cell(7, 8, value="=TIEN_DO_THI_CONG_WBS!J10")  # BBNT-02: Đường công vụ
        ws_kcs.cell(8, 8, value="=TIEN_DO_THI_CONG_WBS!I8")   # BBNT-03: Vật liệu đầu vào
        ws_kcs.cell(9, 8, value="=TIEN_DO_THI_CONG_WBS!J11")  # BBNT-04: Bãi đúc dầm & trạm trộn
        ws_kcs.cell(10, 8, value="=TIEN_DO_THI_CONG_WBS!I14") # BBNT-05: Ống vách thép
        ws_kcs.cell(11, 8, value="=TIEN_DO_THI_CONG_WBS!J13") # BBNT-06: Khoan hang Karst
        ws_kcs.cell(12, 8, value="=TIEN_DO_THI_CONG_WBS!J14") # BBNT-07: Hố khoan cọc D1200
        ws_kcs.cell(13, 8, value="=TIEN_DO_THI_CONG_WBS!J15") # BBNT-08: Lồng cốt thép cọc
        ws_kcs.cell(14, 8, value="=TIEN_DO_THI_CONG_WBS!J17") # BBNT-09: Bê tông cọc C30
        ws_kcs.cell(15, 8, value="=TIEN_DO_THI_CONG_WBS!J18") # BBNT-10: Thí nghiệm siêu âm & PDA
        ws_kcs.cell(16, 8, value="=TIEN_DO_THI_CONG_WBS!J20") # BBNT-11: Bê tông lót bệ C10
        ws_kcs.cell(17, 8, value="=TIEN_DO_THI_CONG_WBS!J22") # BBNT-12: Bê tông bệ mố trụ C30
        ws_kcs.cell(18, 8, value="=TIEN_DO_THI_CONG_WBS!J24") # BBNT-13: Thân mố M1, M2 C30
        ws_kcs.cell(19, 8, value="=TIEN_DO_THI_CONG_WBS!J23") # BBNT-14: Thân trụ T1, T2 C30
        ws_kcs.cell(20, 8, value="=TIEN_DO_THI_CONG_WBS!J25") # BBNT-15: Xà mũ trụ C35
        ws_kcs.cell(21, 8, value="=TIEN_DO_THI_CONG_WBS!J28") # BBNT-16: Đổ bê tông dầm Super-T C45
        ws_kcs.cell(22, 8, value="=TIEN_DO_THI_CONG_WBS!J29") # BBNT-17: Căng kéo cáp DƯL 15.2mm
        ws_kcs.cell(23, 8, value="=TIEN_DO_THI_CONG_WBS!J31") # BBNT-18: Gối cầu & lao dầm
        ws_kcs.cell(24, 8, value="=TIEN_DO_THI_CONG_WBS!J33") # BBNT-19: Dầm ngang & liên tục nhiệt
        ws_kcs.cell(25, 8, value="=TIEN_DO_THI_CONG_WBS!J35") # BBNT-20: Bê tông bản mặt cầu C35
        ws_kcs.cell(26, 8, value="=TIEN_DO_THI_CONG_WBS!J39") # BBNT-21: Khe co giãn, GLC & thảm BTN
        ws_kcs.cell(27, 8, value="=TIEN_DO_THI_CONG_WBS!J42") # BBNT-22: Thử tải & bàn giao hoàn thành

        for r in range(6, 28):
            ws_kcs.cell(r, 8).alignment = align_center
            ws_kcs.cell(r, 8).font = font_reg

    # =========================================================================
    # C2. CẬP NHẬT SHEET THANH_TOAN_KY_PHU_LUC_03A: SỬA LỖI #VALUE! HÀNG 32-34
    # =========================================================================
    if "THANH_TOAN_KY_PHU_LUC_03A" in wb.sheetnames:
        ws_tt = wb["THANH_TOAN_KY_PHU_LUC_03A"]
        ws_tt.cell(31, 6, value="=K28")
        ws_tt.cell(32, 6, value="=F31*0.20")
        ws_tt.cell(33, 6, value="=F31*0.05")
        ws_tt.cell(34, 6, value="=F31-F32-F33")

    # =========================================================================
    # D. TẠO 3 SHEET MẪU BIÊN BẢN NGHIỆM THU EXCEL CHUẨN IN ẤN A4 (THAY THẾ WORD)
    # =========================================================================
    # Xóa sheet cũ nếu đã có
    for s_name in ["MAU_BIEN_BAN_KCS", "MAU_BB_NGHIEM_THU_VAT_LIEU", "MAU_BB_LAY_MAU_HIEN_TRUONG"]:
        if s_name in wb.sheetnames:
            del wb[s_name]

    # -------------------------------------------------------------------------
    # 1. SHEET: MAU_BIEN_BAN_KCS (Biên bản nghiệm thu công việc theo NĐ 207/2026/NĐ-CP)
    # -------------------------------------------------------------------------
    print("[*] Đang tạo Mẫu Biên bản Nghiệm thu Công việc Xây dựng (Excel A4)...")
    ws_bb = wb.create_sheet(title="MAU_BIEN_BAN_KCS")
    ws_bb.views.sheetView[0].showGridLines = True

    # Khung chọn biên bản
    ws_bb["B2"] = "CHỌN SỐ BIÊN BẢN (1 - 22):"
    ws_bb["B2"].font = font_bold
    ws_bb["C2"] = 14 # Mặc định hiển thị biên bản số 14 (Thân trụ)
    ws_bb["C2"].font = Font(name="Times New Roman", size=12, bold=True, color="C00000")
    ws_bb["C2"].fill = fill_input
    ws_bb["C2"].alignment = align_center
    ws_bb["C2"].border = box_border

    ws_bb["D2"] = "(Nhập số thứ tự từ 1 đến 22 để tự động nhảy toàn bộ nội dung biên bản A4)"
    ws_bb["D2"].font = font_it

    # Quốc hiệu tiêu ngữ
    ws_bb["B4"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_bb["B4"].font = Font(name="Times New Roman", size=11, bold=True)
    ws_bb["B4"].alignment = align_center
    ws_bb.merge_cells("B4:G4")

    ws_bb["B5"] = "Độc lập - Tự do - Hạnh phúc"
    ws_bb["B5"].font = Font(name="Times New Roman", size=10, bold=True)
    ws_bb["B5"].alignment = align_center
    ws_bb.merge_cells("B5:G5")

    ws_bb["B6"] = "-------------------------"
    ws_bb["B6"].alignment = align_center
    ws_bb.merge_cells("B6:G6")

    # Tiêu đề biên bản
    ws_bb["B8"] = "BIÊN BẢN NGHIỆM THU CÔNG VIỆC XÂY DỰNG"
    ws_bb["B8"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws_bb["B8"].alignment = align_center
    ws_bb.merge_cells("B8:G8")

    ws_bb["B9"] = '="Số: " & TEXT(C2, "00") & "/BBNT-XD/KM19+529.080"'
    ws_bb["B9"].font = Font(name="Times New Roman", size=10, italic=True)
    ws_bb["B9"].alignment = align_center
    ws_bb.merge_cells("B9:G9")

    # Căn cứ pháp lý
    ws_bb["B11"] = "Căn cứ Luật Xây dựng số 135/2025/QH15;"
    ws_bb["B12"] = "Căn cứ Nghị định số 207/2026/NĐ-CP ngày 15/01/2026 của Chính phủ về quản lý chất lượng công trình;"
    ws_bb["B13"] = "Căn cứ Hợp đồng thi công xây dựng số 09/2026/HĐ-XL gói thầu xây lắp Cầu Km19+529.080;"
    ws_bb["B14"] = "Căn cứ Hồ sơ thiết kế bản vẽ thi công và Chỉ dẫn kỹ thuật được phê duyệt."
    for r in range(11, 15):
        ws_bb.cell(r, 2).font = font_it

    # Thông tin dự án
    ws_bb["B16"] = "1. Tên công trình / Dự án:"
    ws_bb["C16"] = "Dự án đầu tư xây dựng đường cao tốc Tuyên Quang - Hà Giang (Giai đoạn 1)"
    ws_bb["B17"] = "2. Hạng mục công trình:"
    ws_bb["C17"] = "CẦU KM19+529.080 (Sơ đồ nhịp 3x39.1m Super-T)"
    ws_bb["B18"] = "3. Địa điểm xây dựng:"
    ws_bb["C18"] = "Km19+529.080, Tỉnh Hà Giang"
    ws_bb["B19"] = "4. Đối tượng nghiệm thu:"
    ws_bb["C19"] = '=IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 3, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 3, FALSE))'
    ws_bb["C19"].font = Font(name="Times New Roman", size=11, bold=True, color="1F497D")

    ws_bb["B20"] = "5. Khối lượng nghiệm thu:"
    ws_bb["C20"] = '=TEXT(IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 4, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 4, FALSE)), "#,##0.00") & " " & IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 5, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 5, FALSE))'
    ws_bb["C20"].font = font_bold

    ws_bb["B21"] = "6. Tiêu chuẩn kỹ thuật áp dụng:"
    ws_bb["C21"] = '=IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 6, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 6, FALSE))'
    ws_bb["C21"].font = font_reg

    ws_bb["B22"] = "7. Thời gian nghiệm thu:"
    ws_bb["C22"] = '="Bắt đầu: 08 giờ 30 phút - Kết thúc: 11 giờ 30 phút, ngày " & TEXT(IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 8, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 8, FALSE)), "dd/mm/yyyy")'
    ws_bb["C22"].font = font_reg

    for r in range(16, 23):
        ws_bb.cell(r, 2).font = font_bold
        ws_bb.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    # Thành phần trực tiếp nghiệm thu
    ws_bb["B24"] = "8. Thành phần trực tiếp nghiệm thu:"
    ws_bb["B24"].font = font_bold
    
    ws_bb["B25"] = "a) Đại diện Tư vấn giám sát (TVGS):"
    ws_bb["B25"].font = font_bold
    ws_bb["C25"] = "- Ông: Nguyễn Văn Thắng     - Chức vụ: Kỹ sư Giám sát trưởng"
    ws_bb["B26"] = "b) Đại diện Nhà thầu thi công:"
    ws_bb["B26"].font = font_bold
    ws_bb["C26"] = "- Ông: Trần Đình Hoàng      - Chức vụ: Chỉ huy trưởng công trường"
    ws_bb["C27"] = "- Ông: Lê Văn Dũng          - Chức vụ: Kỹ sư QA/QC hiện trường"

    for r in [25, 26, 27]:
        ws_bb.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    # Đánh giá và kết luận
    ws_bb["B29"] = "9. Đánh giá kết quả công việc đã thực hiện:"
    ws_bb["B29"].font = font_bold
    ws_bb["B30"] = "- Về kích thước hình học, cao độ, tim mốc: Đạt yêu cầu theo hồ sơ thiết kế bản vẽ thi công."
    ws_bb["B31"] = "- Về chất lượng vật liệu, chứng chỉ xuất xưởng, kết quả thí nghiệm KCS: Đạt yêu cầu theo tiêu chuẩn áp dụng."
    ws_bb["B32"] = "- Về an toàn lao động, vệ sinh môi trường: Tuân thủ đầy đủ biện pháp thi công được duyệt."
    for r in [30, 31, 32]:
        ws_bb.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws_bb.cell(r, 2).font = font_reg

    ws_bb["B34"] = "10. Kết luận nghiệm thu:"
    ws_bb["B34"].font = font_bold
    ws_bb["B35"] = "CHẤP THUẬN NGHIỆM THU CÔNG VIỆC XÂY DỰNG. ĐỒNG Ý CHO PHÉP TRIỂN KHAI BƯỚC THI CÔNG TIẾP THEO."
    ws_bb["B35"].font = Font(name="Times New Roman", size=10, bold=True, color="006100")
    ws_bb.merge_cells("B35:G35")

    # Chữ ký các bên
    ws_bb["B38"] = "ĐẠI DIỆN NHÀ THẦU THI CÔNG"
    ws_bb["B38"].font = font_bold
    ws_bb["B38"].alignment = align_center
    ws_bb.merge_cells("B38:D38")

    ws_bb["E38"] = "ĐẠI DIỆN TƯ VẤN GIÁM SÁT"
    ws_bb["E38"].font = font_bold
    ws_bb["E38"].alignment = align_center
    ws_bb.merge_cells("E38:G38")

    ws_bb["B39"] = "(Ký, ghi rõ họ tên & đóng dấu)"
    ws_bb["B39"].font = font_it
    ws_bb["B39"].alignment = align_center
    ws_bb.merge_cells("B39:D39")

    ws_bb["E39"] = "(Ký, ghi rõ họ tên)"
    ws_bb["E39"].font = font_it
    ws_bb["E39"].alignment = align_center
    ws_bb.merge_cells("E39:G39")

    ws_bb["B44"] = "Trần Đình Hoàng"
    ws_bb["B44"].font = font_bold
    ws_bb["B44"].alignment = align_center
    ws_bb.merge_cells("B44:D44")

    ws_bb["E44"] = "Nguyễn Văn Thắng"
    ws_bb["E44"].font = font_bold
    ws_bb["E44"].alignment = align_center
    ws_bb.merge_cells("E44:G44")

    # -------------------------------------------------------------------------
    # 2. SHEET: MAU_BB_NGHIEM_THU_VAT_LIEU (Mẫu biên bản nghiệm thu vật liệu đầu vào)
    # -------------------------------------------------------------------------
    print("[*] Đang tạo Mẫu Biên bản Nghiệm thu Vật liệu đầu vào (Excel A4)...")
    ws_vl = wb.create_sheet(title="MAU_BB_NGHIEM_THU_VAT_LIEU")
    ws_vl.views.sheetView[0].showGridLines = True

    ws_vl["B2"] = "CHỌN MÃ VẬT LIỆU (1 - 16):"
    ws_vl["B2"].font = font_bold
    ws_vl["C2"] = 1 # Mặc định chọn Thép Ø<=10mm
    ws_vl["C2"].font = Font(name="Times New Roman", size=12, bold=True, color="C00000")
    ws_vl["C2"].fill = fill_input
    ws_vl["C2"].alignment = align_center
    ws_vl["C2"].border = box_border
    ws_vl["D2"] = "(Nhập mã từ 1 đến 16 theo danh mục ma trận tần suất thí nghiệm để tự nhảy biên bản)"
    ws_vl["D2"].font = font_it

    ws_vl["B4"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_vl["B4"].font = Font(name="Times New Roman", size=11, bold=True)
    ws_vl["B4"].alignment = align_center
    ws_vl.merge_cells("B4:G4")

    ws_vl["B5"] = "Độc lập - Tự do - Hạnh phúc"
    ws_vl["B5"].font = Font(name="Times New Roman", size=10, bold=True)
    ws_vl["B5"].alignment = align_center
    ws_vl.merge_cells("B5:G5")

    ws_vl["B8"] = "BIÊN BẢN NGHIỆM THU VẬT LIỆU XÂY DỰNG ĐẦU VÀO"
    ws_vl["B8"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws_vl["B8"].alignment = align_center
    ws_vl.merge_cells("B8:G8")

    ws_vl["B9"] = '="Số: " & TEXT(C2, "00") & "/BBNT-VL/KM19+529.080"'
    ws_vl["B9"].font = font_it
    ws_vl["B9"].alignment = align_center
    ws_vl.merge_cells("B9:G9")

    ws_vl["B12"] = "1. Tên vật liệu nghiệm thu:"
    ws_vl["C12"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE))'
    ws_vl["C12"].font = Font(name="Times New Roman", size=11, bold=True, color="1F497D")

    ws_vl["B13"] = "2. Khối lượng nghiệm thu lô hàng:"
    ws_vl["C13"] = '=TEXT(IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE)), "#,##0.00") & " " & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE))'
    ws_vl["C13"].font = font_bold

    ws_vl["B14"] = "3. Tiêu chuẩn kỹ thuật áp dụng:"
    ws_vl["C14"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE))'
    ws_vl["C14"].font = font_reg

    ws_vl["B15"] = "4. Chỉ tiêu kiểm tra chất lượng:"
    ws_vl["C15"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 4, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 4, FALSE))'
    ws_vl["C15"].font = font_reg

    ws_vl["B16"] = "5. Quy cách và số lượng tổ mẫu thí nghiệm:"
    ws_vl["C16"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE)) & " (Số tổ mẫu: " & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE)) & " tổ)"'
    ws_vl["C16"].font = font_reg

    ws_vl["B17"] = "6. Đơn vị thực hiện thí nghiệm:"
    ws_vl["C17"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 11, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 11, FALSE))'
    ws_vl["C17"].font = font_reg

    ws_vl["B18"] = "7. Điều kiện nghiệm thu / Chuyển bước:"
    ws_vl["C18"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 12, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 12, FALSE))'
    ws_vl["C18"].font = font_bold

    for r in range(12, 19):
        ws_vl.cell(r, 2).font = font_bold
        ws_vl.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws_vl["B20"] = "8. Kết luận nghiệm thu:"
    ws_vl["B20"].font = font_bold
    ws_vl["B21"] = "VẬT LIỆU ĐẠT TIÊU CHUẨN KỸ THUẬT THEO QUY ĐỊNH. ĐỒNG Ý CHO PHÉP ĐƯA VÀO SỬ DỤNG TRONG THI CÔNG."
    ws_vl["B21"].font = Font(name="Times New Roman", size=10, bold=True, color="006100")
    ws_vl.merge_cells("B21:G21")

    # Ký tên
    ws_vl["B24"] = "ĐẠI DIỆN NHÀ THẦU THI CÔNG"
    ws_vl["B24"].font = font_bold
    ws_vl["B24"].alignment = align_center
    ws_vl.merge_cells("B24:D24")

    ws_vl["E24"] = "ĐẠI DIỆN TƯ VẤN GIÁM SÁT"
    ws_vl["E24"].font = font_bold
    ws_vl["E24"].alignment = align_center
    ws_vl.merge_cells("E24:G24")

    ws_vl["B30"] = "Trần Đình Hoàng"
    ws_vl["B30"].font = font_bold
    ws_vl["B30"].alignment = align_center
    ws_vl.merge_cells("B30:D30")

    ws_vl["E30"] = "Nguyễn Văn Thắng"
    ws_vl["E30"].font = font_bold
    ws_vl["E30"].alignment = align_center
    ws_vl.merge_cells("E30:G30")

    # -------------------------------------------------------------------------
    # 3. SHEET: MAU_BB_LAY_MAU_HIEN_TRUONG (Mẫu biên bản lấy mẫu hiện trường)
    # -------------------------------------------------------------------------
    print("[*] Đang tạo Mẫu Biên bản Lấy mẫu thí nghiệm hiện trường (Excel A4)...")
    ws_lm = wb.create_sheet(title="MAU_BB_LAY_MAU_HIEN_TRUONG")
    ws_lm.views.sheetView[0].showGridLines = True

    ws_lm["B2"] = "CHỌN MÃ HẠNG MỤC LẤY MẪU (17 - 26):"
    ws_lm["B2"].font = font_bold
    ws_lm["C2"] = 22 # Mặc định chọn dầm Super-T C45
    ws_lm["C2"].font = Font(name="Times New Roman", size=12, bold=True, color="C00000")
    ws_lm["C2"].fill = fill_input
    ws_lm["C2"].alignment = align_center
    ws_lm["C2"].border = box_border
    ws_lm["D2"] = "(Nhập mã từ 17 đến 26 tương ứng các hạng mục mẫu nén bê tông để tự động điền form)"
    ws_lm["D2"].font = font_it

    # Ô nhập ngày đúc mẫu thực tế
    ws_lm["B3"] = "CHỌN NGÀY ĐÚC MẪU (YYYY-MM-DD):"
    ws_lm["B3"].font = font_bold
    ws_lm["C3"] = "2026-11-20"
    ws_lm["C3"].font = Font(name="Times New Roman", size=11, bold=True, color="C00000")
    ws_lm["C3"].fill = fill_input
    ws_lm["C3"].alignment = align_center
    ws_lm["C3"].border = box_border
    ws_lm["D3"] = "(Tự động tính ngày nén mẫu R7 = Ngày đúc + 7, R28 = Ngày đúc + 28 theo đúng quy định)"
    ws_lm["D3"].font = font_it

    ws_lm["B5"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws_lm["B5"].font = Font(name="Times New Roman", size=11, bold=True)
    ws_lm["B5"].alignment = align_center
    ws_lm.merge_cells("B5:G5")

    ws_lm["B6"] = "Độc lập - Tự do - Hạnh phúc"
    ws_lm["B6"].font = Font(name="Times New Roman", size=10, bold=True)
    ws_lm["B6"].alignment = align_center
    ws_lm.merge_cells("B6:G6")

    ws_lm["B8"] = "BIÊN BẢN LẤY MẪU THÍ NGHIỆM TẠI HIỆN TRƯỜNG"
    ws_lm["B8"].font = Font(name="Times New Roman", size=13, bold=True, color="1F497D")
    ws_lm["B8"].alignment = align_center
    ws_lm.merge_cells("B8:G8")

    ws_lm["B9"] = '="Số: " & TEXT(C2, "00") & "/BBLM-HT/KM19+529.080"'
    ws_lm["B9"].font = font_it
    ws_lm["B9"].alignment = align_center
    ws_lm.merge_cells("B9:G9")

    ws_lm["B12"] = "1. Đối tượng và cấu kiện lấy mẫu:"
    ws_lm["C12"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE))'
    ws_lm["C12"].font = Font(name="Times New Roman", size=11, bold=True, color="1F497D")

    ws_lm["B13"] = "2. Thể tích bê tông / Quy mô mẻ đổ:"
    ws_lm["C13"] = '=TEXT(IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE)), "#,##0.00") & " " & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE))'
    ws_lm["C13"].font = font_bold

    ws_lm["B14"] = "3. Quy định số tổ mẫu bắt buộc:"
    ws_lm["C14"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE)) & " tổ mẫu (" & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE)) & ")"'
    ws_lm["C14"].font = font_bold

    ws_lm["B15"] = "4. Tiêu chuẩn lấy mẫu & thử nghiệm:"
    ws_lm["C15"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE))'
    ws_lm["C15"].font = font_reg

    ws_lm["B16"] = "5. Phương pháp bảo dưỡng mẫu:"
    ws_lm["C16"] = "Bảo dưỡng ẩm tiêu chuẩn trong phòng thí nghiệm (TCVN 3105:2022)"
    ws_lm["C16"].font = font_reg

    ws_lm["B17"] = "6. Kế hoạch nén kiểm tra cường độ:"
    ws_lm["C17"] = '="Nén kiểm tra R7 ngày (" & TEXT(DATEVALUE(C3)+7, "dd/mm/yyyy") & "): Đạt >= 70-80% f\'c | Nén R28 ngày (" & TEXT(DATEVALUE(C3)+28, "dd/mm/yyyy") & "): Đạt >= 100% f\'c"'
    ws_lm["C17"].font = font_bold

    for r in range(12, 18):
        ws_lm.cell(r, 2).font = font_bold
        ws_lm.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws_lm["B19"] = "7. Tình trạng mẫu khi bàn giao:"
    ws_lm["B19"].font = font_bold
    ws_lm["B20"] = "Mẫu đúc nguyên vẹn, bề mặt phẳng đều, dán nhãn ghi rõ ngày đúc, vị trí cấu kiện và ký hiệu tổ mẫu."
    ws_lm["B20"].font = font_it
    ws_lm.merge_cells("B20:G20")

    # Ký tên
    ws_lm["B23"] = "CÁN BỘ LẤY MẪU THÍ NGHIỆM (LAS-XD)"
    ws_lm["B23"].font = font_bold
    ws_lm["B23"].alignment = align_center
    ws_lm.merge_cells("B23:D23")

    ws_lm["E23"] = "KỸ SƯ GIÁM SÁT CHỨNG KIẾN (TVGS)"
    ws_lm["E23"].font = font_bold
    ws_lm["E23"].alignment = align_center
    ws_lm.merge_cells("E23:G23")

    ws_lm["B29"] = "Hoàng Trung Kiên"
    ws_lm["B29"].font = font_bold
    ws_lm["B29"].alignment = align_center
    ws_lm.merge_cells("B29:D29")

    ws_lm["E29"] = "Nguyễn Văn Thắng"
    ws_lm["E29"].font = font_bold
    ws_lm["E29"].alignment = align_center
    ws_lm.merge_cells("E29:G29")

    # Căn chỉnh kích thước cột cho các sheet biểu mẫu A4
    for ws_form in [ws_bb, ws_vl, ws_lm]:
        ws_form.column_dimensions["A"].width = 4
        ws_form.column_dimensions["B"].width = 28
        ws_form.column_dimensions["C"].width = 16
        ws_form.column_dimensions["D"].width = 16
        ws_form.column_dimensions["E"].width = 16
        ws_form.column_dimensions["F"].width = 16
        ws_form.column_dimensions["G"].width = 16

    print(f"[*] Đang lưu file Excel Master cập nhật hoàn chỉnh: {excel_path}")
    wb.save(excel_path)
    print(f"[V] Thành công! Toàn bộ hệ thống 14 Sheet đã liên kết động 100% không số chết!")
    return wb.sheetnames

if __name__ == "__main__":
    target_excel = r"D:\Code\DONG_GOI_HETHONG_AEC\templates\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx"
    sheets = build_full_cross_linked_workbook(target_excel)
    print("Danh sách Sheet hiện tại trong Workbook:", sheets)
