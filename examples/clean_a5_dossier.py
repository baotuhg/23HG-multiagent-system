# -*- coding: utf-8 -*-
"""
DỌN HỒ SƠ MẪU CỐNG HỘP TUYẾN A5 (chạy lặp được — idempotent)

1. Đắp lưng cống: tách theo từng loại cống, trừ phần cống chiếm chỗ (Phụ lục VI TT 13/2021, mục II.5.2);
   khoảng thao tác và lớp phủ là ô đầu vào có nhãn. Số cũ được GIỮ để đối chiếu, không bị ghi đè.
2. Mẫu 03a: thêm 3 cột đối chiếu khối lượng hợp đồng với khối lượng bóc tách (hợp đồng vẫn là số nhập).
3. Dự toán G_XD: chi phí trực tiếp T và các tỷ lệ là ô đầu vào có nhãn (không còn "=42500000000" giả công thức).
4. Sheet Master chưa có dữ liệu được ghi "CHƯA LẬP"; file vi mô rỗng bị xóa; Gói C không còn chứa bản Master có đơn giá.

Chạy:  python examples/clean_a5_dossier.py
"""
import json
import os
import shutil
import sys
from copy import copy

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from tools.package_dispatcher import clone_worksheet

A5 = os.path.join(ROOT, "examples", "HO_SO_CONG_HOP_TUYEN_A5")
T1 = os.path.join(A5, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
T2 = os.path.join(A5, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")
HUB = os.path.join(A5, "HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5")
MASTER = os.path.join(T1, "01_Ho_So_KCS_QS_TienDo_Master_14_Sheets_Cong_Hop_A5.xlsx")

QS = "QS_DIEN_GIAI_CHI_TIET"
DAO_DAP_MASTER = "KHOI_LUONG_DAO_DAP"
THEP = "THONG_KE_THEP_CHI_TIET"

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(wrap_text=True, vertical="top")


def _hdr(ws, coord, text):
    c = ws[coord]
    c.value, c.font, c.fill, c.alignment, c.border = text, HDR_FONT, HDR_FILL, CENTER, BORDER


def _input(ws, coord, value, fmt=None):
    c = ws[coord]
    c.value, c.fill, c.border = value, INPUT_FILL, BORDER
    if fmt:
        c.number_format = fmt


# ─────────────────────────────────────────────────────────────────────────────
# 1. ĐẮP LƯNG CỐNG
# ─────────────────────────────────────────────────────────────────────────────

def patch_dao_dap(ws):
    """Cần sheet QS_DIEN_GIAI_CHI_TIET (cột D số đốt, E dài, F rộng ngoài, G cao ngoài ở dòng 8-10) trong cùng workbook."""
    if ws["A10"].value:                      # đã vá
        return False
    ws["I8"].value = f"=D8*E8*F8*G8*H8-{QS}!I8"      # khôi phục liên kết sống (bản vi mô trước đây dán số cứng)
    ws["B8"].value = "Đắp cát/đất đầm chặt K95 mang cống và lưng cống — SỐ CŨ (chỉ để đối chiếu, xem khối tính lại bên dưới)"
    ws["J8"].value = ("Số cũ: một mặt cắt hào trung bình, chỉ trừ thể tích BÊ TÔNG cống 2.0x2.0m; không đúng với 3 loại "
                      "cống và không trừ phần bao ngoài chiếm chỗ")
    ws["A10"].value = ("ĐẮP LƯNG CỐNG TÍNH LẠI THEO TỪNG LOẠI CỐNG — trừ phần cống chiếm chỗ "
                       "(Phụ lục VI TT 13/2021, mục II.5.2; cần đối chiếu bản gốc)")
    ws["A10"].font = Font(bold=True)
    ws["B11"].value = "Khoảng thao tác mỗi bên cống (m)"
    _input(ws, "D11", 0.75, "0.00")
    ws["J11"].value = "GIẢ ĐỊNH suy từ số cũ: hào rộng 4.0 m quanh cống 2.5 m → 0.75 m mỗi bên. QS xác nhận theo biện pháp thi công"
    ws["B12"].value = "Chiều dày lớp đắp phủ trên nắp cống (m)"
    _input(ws, "D12", 0, "0.00")
    ws["J12"].value = "0 = chỉ đắp lưng cống đến đỉnh cống (số cũ dùng sâu 2.5 m = cao ngoài cống 2.0x2.0m)"
    for col, text in zip("ABCDEFGHIJ", ["STT", "LOẠI CỐNG", "ĐƠN VỊ", "SỐ ĐỐT", "DÀI (m)", "RỘNG HÀO (m)", "SÂU HÀO (m)",
                                        "TIẾT DIỆN CỐNG CHIẾM CHỖ (m2)", "KHỐI LƯỢNG ĐẮP (m3)", "GHI CHÚ"]):
        _hdr(ws, f"{col}14", text)
    for k, (qrow, name) in enumerate(((8, "Cống đơn 2.0x2.0m"), (9, "Cống đơn 3.0x3.0m"), (10, "Cống đôi 2x(3.0x3.0m)"))):
        r = 15 + k
        ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"C{r}"].value = k + 1, name, "m3"
        ws[f"D{r}"].value = f"={QS}!D{qrow}"
        ws[f"E{r}"].value = f"={QS}!E{qrow}"
        ws[f"F{r}"].value = f"={QS}!F{qrow}+2*$D$11"
        ws[f"G{r}"].value = f"={QS}!G{qrow}+$D$12"
        ws[f"H{r}"].value = f"={QS}!F{qrow}*{QS}!G{qrow}"
        ws[f"I{r}"].value = f"=D{r}*E{r}*(F{r}*G{r}-H{r})"
        ws[f"J{r}"].value = "Hào − phần bao ngoài cống chiếm chỗ"
        for col in "ABCDEFGHIJ":
            ws[f"{col}{r}"].border = BORDER
        for col in "EFGH":
            ws[f"{col}{r}"].number_format = "#,##0.00"
        ws[f"I{r}"].number_format = "#,##0.000"
    ws["B18"].value = "TỔNG ĐẮP LƯNG CỐNG (tính lại)"
    ws["I18"].value = "=SUM(I15:I17)"
    ws["B19"].value = "Số cũ (dòng 4)"
    ws["I19"].value = "=I8"
    ws["B20"].value = "Chênh lệch tính lại so với số cũ"
    ws["I20"].value = "=I18/I19-1"
    ws["J20"].value = '=IF(ABS(I20)>0.05,"LỆCH >5% - QS xác nhận","Khớp (<=5%)")'
    for r in (18, 19, 20):
        ws[f"B{r}"].font = Font(bold=True)
        ws[f"I{r}"].number_format = "0.0%" if r == 20 else "#,##0.000"
        ws[f"I{r}"].border = BORDER
    return True


# ─────────────────────────────────────────────────────────────────────────────
# 2. MẪU 03a — ĐỐI CHIẾU KHỐI LƯỢNG HỢP ĐỒNG VỚI BÓC TÁCH
# ─────────────────────────────────────────────────────────────────────────────

def patch_03a(ws, dao_dap_sheet, qs=QS, thep=THEP):
    if ws["J4"].value:
        return False
    for col, text in zip("JKL", ["KL THEO BÓC TÁCH", "CHÊNH LỆCH SO VỚI HỢP ĐỒNG", "KẾT LUẬN (ngưỡng 0.5%)"]):
        _hdr(ws, f"{col}4", text)
        ws.column_dimensions[col].width = 24
    sources = {
        5: f"={dao_dap_sheet}!I5",
        6: f"={dao_dap_sheet}!I6",
        7: f"=SUM({qs}!I5:I7)",
        8: f"=SUM({qs}!I8:I10)",
        9: f"=SUM({qs}!I11:I13)",
        10: f"={thep}!J24/1000",
    }
    for r, f in sources.items():
        ws[f"J{r}"].value = f
        ws[f"K{r}"].value = f"=J{r}/D{r}-1"
        ws[f"L{r}"].value = f'=IF(ABS(K{r})<=0.005,"Khớp","LỆCH - QS xác nhận")'
        ws[f"J{r}"].number_format = "#,##0.000"
        ws[f"K{r}"].number_format = "0.00%"
        for col in "JKL":
            ws[f"{col}{r}"].border = BORDER
    ws["B17"].value = ("Cột D (khối lượng hợp đồng) là SỐ NHẬP từ hợp đồng; cột J..L chỉ đối chiếu với bóc tách, không ghi đè. "
                       "Ván khuôn dùng số đang áp dụng trong bảng diễn giải (xem cột T của sheet diễn giải).")
    ws["B17"].alignment = WRAP
    ws.merge_cells("B17:I17")
    ws.row_dimensions[17].height = 48
    return True


# ─────────────────────────────────────────────────────────────────────────────
# 3. DỰ TOÁN G_XD — ĐẦU VÀO CÓ NHÃN
# ─────────────────────────────────────────────────────────────────────────────

def patch_gxd(ws):
    if ws["A12"].value:
        return False
    ws["F5"].value = ("ĐẦU VÀO: nhập từ dự toán chi tiết đã duyệt (Σ khối lượng × đơn giá theo định mức). "
                      "CHƯA liên kết với khối lượng/đơn giá trong hồ sơ này")
    _input(ws, "D5", 42500000000, "#,##0")
    ws["A12"].value = "TỶ LỆ ĐẦU VÀO (đổi ô vàng, các dòng II-VI tự tính lại)"
    ws["A12"].font = Font(bold=True)
    for r, (label, val) in zip((13, 14, 15, 16), (("Chi phí chung (% T)", 0.073), ("Nhà tạm để ở và điều hành thi công (% T)", 0.012),
                                                  ("Thu nhập chịu thuế tính trước (% (T + GT))", 0.055),
                                                  ("Thuế giá trị gia tăng (% G)", 0.10))):
        ws[f"B{r}"].value = label
        _input(ws, f"D{r}", val, "0.0%")
    ws["D6"].value = "=D5*(D13+D14)"
    ws["D7"].value = "=(D5+D6)*D15"
    ws["D9"].value = "=D8*D16"
    ws["C6"].value = "GT = T × (chi phí chung + nhà tạm)"
    ws["C7"].value = "TL = (T + GT) × tỷ lệ TL"
    ws["C9"].value = "VAT = G × thuế suất"
    ws["F6"].value, ws["F7"].value, ws["F9"].value = "Theo ô tỷ lệ D13 + D14", "Theo ô tỷ lệ D15", "Theo ô tỷ lệ D16"
    return True


# ─────────────────────────────────────────────────────────────────────────────
# 4. SHEET CHƯA LẬP, FILE RỖNG, GÓI C
# ─────────────────────────────────────────────────────────────────────────────

PLACEHOLDER_SHEETS = ["CAP_PHOI_1M3_VA_TAN_SUAT", "PHAN_TICH_VAT_TU_WBS", "TONG_HOP_VAT_TU_TOAN_BO",
                      "MAU_BIEN_BAN_KCS", "MAU_BB_NGHIEM_THU_VAT_LIEU", "MAU_BB_LAY_MAU_HIEN_TRUONG"]
PLACEHOLDER_NOTE = "CHƯA LẬP — chưa có dữ liệu cho dự án này. Không xuất thành hồ sơ riêng cho tới khi có nội dung."

EMPTY_FILES = [
    os.path.join(T2, "05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem_Cong_A5.xlsx"),
    os.path.join(T2, "06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS_Cong_A5.xlsx"),
    os.path.join(T2, "07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan_Cong_A5.xlsx"),
    os.path.join(T2, "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec_Cong_A5.xlsx"),
    os.path.join(T2, "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu_Cong_A5.xlsx"),
    os.path.join(T2, "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28_Cong_A5.xlsx"),
    os.path.join(HUB, "GOI_C_HIEN_TRUONG_QLCL_KCS", "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec_Cong_A5.xlsx"),
    os.path.join(HUB, "GOI_C_HIEN_TRUONG_QLCL_KCS", "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu_Cong_A5.xlsx"),
    os.path.join(HUB, "GOI_C_HIEN_TRUONG_QLCL_KCS", "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28_Cong_A5.xlsx"),
]
# Gói C (hiện trường KCS) không được xem đơn giá: bản Master có G_XD và 03a không thuộc gói này
MASTER_COPY_IN_GOI_C = os.path.join(HUB, "GOI_C_HIEN_TRUONG_QLCL_KCS", os.path.basename(MASTER))


def _support(wb_dst, wb_master, names):
    for n in names:
        if n not in wb_dst.sheetnames:
            clone_worksheet(wb_master[n], wb_dst.create_sheet(title=n))


def main() -> int:
    print("1. Master A5")
    wb = openpyxl.load_workbook(MASTER)
    changed = [patch_dao_dap(wb[DAO_DAP_MASTER]), patch_03a(wb["THANH_TOAN_KY_PHU_LUC_03A"], DAO_DAP_MASTER),
               patch_gxd(wb["TONG_HOP_DU_TOAN_GXD"])]
    for n in PLACEHOLDER_SHEETS:
        ws = wb[n]
        if not ws["A3"].value:
            ws["A3"].value = PLACEHOLDER_NOTE
            ws["A3"].font = Font(bold=True, color="C00000")
            changed.append(True)
    if any(changed):
        wb.save(MASTER)
        print("   đã vá Master")
    wb_master = openpyxl.load_workbook(MASTER)

    print("2. File vi mô (Tầng 2) và bản sao ở Gói D")
    p02 = os.path.join(T2, "02_Khoi_Luong_Dao_Dap_Mong_Cong_A5.xlsx")
    p08 = os.path.join(T2, "08_Du_Toan_GXD_Thong_Tu_11_2021_Cong_A5.xlsx")
    p09 = os.path.join(T2, "09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a_Cong_A5.xlsx")
    w = openpyxl.load_workbook(p02)
    _support(w, wb_master, [QS])
    if patch_dao_dap(w["DAO_DAP"]):
        w.save(p02)
    w = openpyxl.load_workbook(p08)
    if patch_gxd(w["DU_TOAN_GXD"]):
        w.save(p08)
    w = openpyxl.load_workbook(p09)
    _support(w, wb_master, [DAO_DAP_MASTER, QS, THEP])
    if patch_03a(w["PHU_LUC_03A"], DAO_DAP_MASTER):
        w.save(p09)
    goi_d = os.path.join(HUB, "GOI_D_QS_DU_TOAN_THANH_TOAN")
    for p in (p08, p09):
        shutil.copyfile(p, os.path.join(goi_d, os.path.basename(p)))      # bản sao Gói D luôn trùng Tầng 2

    print("3. Xóa file rỗng và bản Master có đơn giá trong Gói C")
    removed = 0
    for p in EMPTY_FILES + [MASTER_COPY_IN_GOI_C]:
        if os.path.exists(p):
            os.remove(p)
            removed += 1
            print("   xóa:", os.path.relpath(p, A5))

    manifest_path = os.path.join(HUB, "DISPATCH_MANIFEST.json")
    with open(manifest_path, encoding="utf-8") as f:
        manifest = json.load(f)
    manifest["total_files"] = sum(len(fs) for _, _, fs in os.walk(HUB)) - 1      # trừ chính file manifest
    manifest["not_generated"] = {
        "reason": "Sheet tương ứng trong Master chưa có dữ liệu (CHƯA LẬP) nên không xuất thành hồ sơ riêng",
        "dossiers": ["05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem", "06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS",
                     "07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan", "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec",
                     "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu", "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28"],
    }
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"Xong. Đã xóa {removed} file.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
