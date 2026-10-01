# -*- coding: utf-8 -*-
"""
Script khắc phục triệt để các lỗi công thức trong Master Workbook:
Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx
1. TO_HOP_CAT_THEP_11M7: Sửa công thức COUNTIFS khớp BBS, chống #DIV/0! bằng IF(I=0, 0, J/I*100)
2. THANH_TOAN_KY_PHU_LUC_03A: Sửa F32, F33, F34 trỏ đúng F31 (số tiền) thay vì C32 (chữ) chống #VALUE!
3. HOSO_KCS_NGHIEM_THU: Chuẩn hóa cột A thành số Integer (1..22)
4. CAP_PHOI_1M3_VA_TAN_SUAT: Chuẩn hóa cột A thành số Integer (1..29)
5. MAU_BIEN_BAN_KCS, MAU_BB_NGHIEM_THU_VAT_LIEU, MAU_BB_LAY_MAU_HIEN_TRUONG:
   Bọc IFERROR đa tầng hỗ trợ cả chuỗi và số nguyên trong VLOOKUP.
"""

import os
import openpyxl

MASTER_PATH = r"D:\Code\23HG-multiagent-system\23HG-multiagent-system\templates\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx"

def fix_master_workbook(fpath=MASTER_PATH):
    print(f"[*] Đang tải và sửa lỗi Master Workbook: {fpath}")
    wb = openpyxl.load_workbook(fpath)
    
    # -------------------------------------------------------------------------
    # 1. FIX HOSO_KCS_NGHIEM_THU: Cột A thành số Integer
    # -------------------------------------------------------------------------
    if "HOSO_KCS_NGHIEM_THU" in wb.sheetnames:
        ws = wb["HOSO_KCS_NGHIEM_THU"]
        print("  - Sửa HOSO_KCS_NGHIEM_THU: Ép kiểu cột A thành Integer 1..22...")
        for r in range(6, 28):
            ws.cell(r, 1, value=int(r - 5))
            ws.cell(r, 1).number_format = "0"

    # -------------------------------------------------------------------------
    # 2. FIX CAP_PHOI_1M3_VA_TAN_SUAT: Cột A thành số Integer
    # -------------------------------------------------------------------------
    if "CAP_PHOI_1M3_VA_TAN_SUAT" in wb.sheetnames:
        ws = wb["CAP_PHOI_1M3_VA_TAN_SUAT"]
        print("  - Sửa CAP_PHOI_1M3_VA_TAN_SUAT: Ép kiểu cột A thành Integer 1..29...")
        for r in range(24, 53):
            ws.cell(r, 1, value=int(r - 23))
            ws.cell(r, 1).number_format = "0"

    # -------------------------------------------------------------------------
    # 3. FIX TO_HOP_CAT_THEP_11M7: COUNTIFS chuẩn BBS + chống #DIV/0!
    # -------------------------------------------------------------------------
    if "TO_HOP_CAT_THEP_11M7" in wb.sheetnames:
        ws = wb["TO_HOP_CAT_THEP_11M7"]
        print("  - Sửa TO_HOP_CAT_THEP_11M7: Cột D COUNTIFS & Cột L an toàn chống chia 0...")
        
        # Cập nhật công thức cột D khớp với cấu trúc thực của sheet THONG_KE_THEP_CHI_TIET
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
            ws.cell(r, 4, value=form)
            ws.cell(r, 4).number_format = "#,##0"
            # Cột L: IF(I=0, 0, (J/I)*100)
            ws.cell(r, 12, value=f"=IF(I{r}=0, 0, (J{r}/I{r})*100)")
            ws.cell(r, 12).number_format = "0.00"
            
        ws.cell(20, 12, value="=IFERROR(AVERAGE(L6:L19), 0)")
        ws.cell(20, 12).number_format = "0.00"

    # -------------------------------------------------------------------------
    # 4. FIX THANH_TOAN_KY_PHU_LUC_03A: Sửa F32..F34 trỏ F31 thay vì C32
    # -------------------------------------------------------------------------
    if "THANH_TOAN_KY_PHU_LUC_03A" in wb.sheetnames:
        ws = wb["THANH_TOAN_KY_PHU_LUC_03A"]
        print("  - Sửa THANH_TOAN_KY_PHU_LUC_03A: Sửa công thức khấu trừ F32, F33, F34...")
        ws.cell(31, 6, value="=K28")
        ws.cell(32, 6, value="=F31*0.20")
        ws.cell(33, 6, value="=F31*0.05")
        ws.cell(34, 6, value="=F31-F32-F33")

    # -------------------------------------------------------------------------
    # 5. FIX MAU_BIEN_BAN_KCS: VLOOKUP an toàn tuyệt đối
    # -------------------------------------------------------------------------
    if "MAU_BIEN_BAN_KCS" in wb.sheetnames:
        ws = wb["MAU_BIEN_BAN_KCS"]
        print("  - Sửa MAU_BIEN_BAN_KCS: Bọc IFERROR VLOOKUP...")
        ws["C19"] = '=IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 3, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 3, FALSE))'
        ws["C20"] = '=TEXT(IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 4, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 4, FALSE)), "#,##0.00") & " " & IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 5, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 5, FALSE))'
        ws["C21"] = '=IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 6, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 6, FALSE))'
        ws["C22"] = '="Bắt đầu: 08 giờ 30 phút - Kết thúc: 11 giờ 30 phút, ngày " & TEXT(IFERROR(VLOOKUP(C2, HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 8, FALSE), VLOOKUP(TEXT(C2, "@"), HOSO_KCS_NGHIEM_THU!$A$6:$H$27, 8, FALSE)), "dd/mm/yyyy")'

    # -------------------------------------------------------------------------
    # 6. FIX MAU_BB_NGHIEM_THU_VAT_LIEU: VLOOKUP an toàn tuyệt đối
    # -------------------------------------------------------------------------
    if "MAU_BB_NGHIEM_THU_VAT_LIEU" in wb.sheetnames:
        ws = wb["MAU_BB_NGHIEM_THU_VAT_LIEU"]
        print("  - Sửa MAU_BB_NGHIEM_THU_VAT_LIEU: Bọc IFERROR VLOOKUP phạm vi A24:L52...")
        ws["C12"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE))'
        ws["C13"] = '=TEXT(IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE)), "#,##0.00") & " " & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE))'
        ws["C14"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE))'
        ws["C15"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 4, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 4, FALSE))'
        ws["C16"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE)) & " (Số tổ mẫu: " & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE)) & " tổ)"'
        ws["C17"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 11, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 11, FALSE))'
        ws["C18"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 12, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 12, FALSE))'

    # -------------------------------------------------------------------------
    # 7. FIX MAU_BB_LAY_MAU_HIEN_TRUONG: VLOOKUP an toàn tuyệt đối
    # -------------------------------------------------------------------------
    if "MAU_BB_LAY_MAU_HIEN_TRUONG" in wb.sheetnames:
        ws = wb["MAU_BB_LAY_MAU_HIEN_TRUONG"]
        print("  - Sửa MAU_BB_LAY_MAU_HIEN_TRUONG: Bọc IFERROR VLOOKUP phạm vi A24:L52...")
        ws["C12"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 2, FALSE))'
        ws["C13"] = '=TEXT(IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 7, FALSE)), "#,##0.00") & " " & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 6, FALSE))'
        ws["C14"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 9, FALSE)) & " tổ mẫu (" & IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 10, FALSE)) & ")"'
        ws["C15"] = '=IFERROR(VLOOKUP(C2, CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE), VLOOKUP(TEXT(C2, "@"), CAP_PHOI_1M3_VA_TAN_SUAT!$A$24:$L$52, 3, FALSE))'

    wb.save(fpath)
    print(f"[V] Đã sửa và lưu thành công Master Workbook tại: {fpath}")

if __name__ == "__main__":
    fix_master_workbook()
