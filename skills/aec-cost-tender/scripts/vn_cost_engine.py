#!/usr/bin/env python3
"""Vietnamese Construction Cost Estimation Engine (Dự toán Xây dựng Việt Nam).

Conforms to:
- Luật Xây dựng số 135/2025/QH15 & Nghị định 206/2026/NĐ-CP (NĐ 207/2026/NĐ-CP).
- Thông tư 36/2026/TT-BXD: Hướng dẫn xác định và quản lý chi phí đầu tư xây dựng.
- Thông tư 38/2026/TT-BXD: Ban hành định mức xây dựng (Mã hiệu AB, AC, AD, AE, AF, AG, AK...).
- Thông tư 37/2026/TT-BXD: Phương pháp xác định chỉ tiêu kinh tế kỹ thuật và đo bóc khối lượng.

Calculates:
G_XD = T + GT + TL + VAT
Where:
- T = Chi phí trực tiếp (Vật liệu VL + Nhân công NC + Máy thi công MTC)
- GT = Chi phí gián tiếp (Chi phí chung C_C + Nhà tạm C_NT + Chi phí KXD C_KXD)
- TL = Thu nhập chịu thuế tính trước
- VAT = Thuế giá trị gia tăng (mặc định 10% hoặc 8%)
"""
from dataclasses import dataclass
from decimal import Decimal
from typing import Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Định mức tỷ lệ % theo Thông tư 36/2026/TT-BXD cho các loại công trình
# Tỷ lệ chi phí chung (C_C) tính trên chi phí trực tiếp (T <= 100 tỷ)
PROJECT_RATES = {
    "DAN_DUNG": {
        "name": "Công trình Dân dụng (Trường học, Nhà ở, Trụ sở)",
        "rate_chi_phi_chung": Decimal("0.073"),      # 7.3%
        "rate_nha_tam": Decimal("0.010"),            # 1.0%
        "rate_khong_xac_dinh": Decimal("0.015"),     # 1.5%
        "rate_thu_nhap_tinh_truoc": Decimal("0.055"),# 5.5%
        "vat": Decimal("0.10"),                      # 10%
    },
    "GIAO_THONG": {
        "name": "Công trình Giao thông (Đường, Cầu, Cống theo tuyến)",
        "rate_chi_phi_chung": Decimal("0.062"),      # 6.2%
        "rate_nha_tam": Decimal("0.012"),            # 1.2% (công trình theo tuyến)
        "rate_khong_xac_dinh": Decimal("0.015"),     # 1.5%
        "rate_thu_nhap_tinh_truoc": Decimal("0.055"),# 5.5%
        "vat": Decimal("0.10"),                      # 10%
    },
    "HA_TANG_KY_THUAT": {
        "name": "Công trình Hạ tầng kỹ thuật (San nền, Kè đá, Thoát nước)",
        "rate_chi_phi_chung": Decimal("0.058"),      # 5.8%
        "rate_nha_tam": Decimal("0.010"),            # 1.0%
        "rate_khong_xac_dinh": Decimal("0.015"),     # 1.5%
        "rate_thu_nhap_tinh_truoc": Decimal("0.055"),# 5.5%
        "vat": Decimal("0.10"),                      # 10%
    },
    "NONG_NGHIEP_PTNT": {
        "name": "Công trình Nông nghiệp & PTNT (Kè suối, Đê điều, Thủy lợi)",
        "rate_chi_phi_chung": Decimal("0.060"),      # 6.0%
        "rate_nha_tam": Decimal("0.012"),            # 1.2%
        "rate_khong_xac_dinh": Decimal("0.015"),     # 1.5%
        "rate_thu_nhap_tinh_truoc": Decimal("0.055"),# 5.5%
        "vat": Decimal("0.10"),                      # 10%
    }
}

# Tiền tố mã định mức Thông tư 38/2026/TT-BXD
NORM_PREFIX_MAP = {
    "AA": "Công tác chuẩn bị mặt bằng",
    "AB": "Công tác đào, đắp đất, đá, cát (San nền, móng, nền đường)",
    "AC": "Công tác cọc (Đóng, ép, cọc khoan nhồi)",
    "AD": "Công tác làm đường (Cấp phối đá dăm, thảm nhựa)",
    "AE": "Công tác xây gạch, xây đá hộc kè, lát đá",
    "AF": "Công tác bê tông tại chỗ (Móng, cột, dầm, sàn, tường chắn)",
    "AG": "Công tác sản xuất, lắp dựng cốt thép",
    "AH": "Công tác sản xuất, lắp dựng cấu kiện bê tông đúc sẵn",
    "AI": "Công tác gia công, lắp dựng kết cấu gỗ",
    "AK": "Công tác hoàn thiện (Trát, láng, ốp, sơn)",
    "AL": "Công tác kết cấu thép, gia công kim loại",
    "BA": "Công tác lắp đặt thiết bị điện, nước, MEP",
}


@dataclass
class CostItem:
    stt: int
    ma_hieu: str
    ten_cong_tac: str
    don_vi: str
    khoi_luong: Decimal
    don_gia_vat_lieu: Decimal
    don_gia_nhan_cong: Decimal
    don_gia_may: Decimal
    ghi_chu: str = ""

    @property
    def don_gia_tong_hop(self) -> Decimal:
        return self.don_gia_vat_lieu + self.don_gia_nhan_cong + self.don_gia_may

    @property
    def thanh_tien_vat_lieu(self) -> Decimal:
        return self.khoi_luong * self.don_gia_vat_lieu

    @property
    def thanh_tien_nhan_cong(self) -> Decimal:
        return self.khoi_luong * self.don_gia_nhan_cong

    @property
    def thanh_tien_may(self) -> Decimal:
        return self.khoi_luong * self.don_gia_may

    @property
    def thanh_tien_truc_tiep(self) -> Decimal:
        return self.khoi_luong * self.don_gia_tong_hop


@dataclass
class CostSummary:
    project_type: str
    project_name: str
    chi_phi_vat_lieu: Decimal
    chi_phi_nhan_cong: Decimal
    chi_phi_may: Decimal
    chi_phi_truc_tiep: Decimal  # T
    chi_phi_chung: Decimal      # C_C
    chi_phi_nha_tam: Decimal    # C_NT
    chi_phi_kxd: Decimal        # C_KXD
    tong_chi_phi_gian_tiep: Decimal # GT
    thu_nhap_chiu_thue_tinh_truoc: Decimal # TL
    chi_phi_xay_dung_truoc_thue: Decimal # G = T + GT + TL
    thue_vat: Decimal           # VAT
    tong_chi_phi_xay_dung_sau_thue: Decimal # G_XD = G + VAT


def calculate_vietnam_estimate(
    items: List[CostItem],
    project_type: str = "GIAO_THONG",
    project_name: str = "Dự án Công trình Xây dựng",
    vat_rate: Optional[Decimal] = None
) -> Tuple[CostSummary, List[CostItem]]:
    rates = PROJECT_RATES.get(project_type, PROJECT_RATES["GIAO_THONG"])
    vat_used = vat_rate if vat_rate is not None else rates["vat"]

    vl = sum((it.thanh_tien_vat_lieu for it in items), Decimal("0"))
    nc = sum((it.thanh_tien_nhan_cong for it in items), Decimal("0"))
    mtc = sum((it.thanh_tien_may for it in items), Decimal("0"))
    t = vl + nc + mtc

    c_c = t * rates["rate_chi_phi_chung"]
    c_nt = t * rates["rate_nha_tam"]
    c_kxd = t * rates["rate_khong_xac_dinh"]
    gt = c_c + c_nt + c_kxd

    tl = (t + gt) * rates["rate_thu_nhap_tinh_truoc"]
    g = t + gt + tl
    vat = g * vat_used
    g_xd = g + vat

    summary = CostSummary(
        project_type=project_type,
        project_name=project_name,
        chi_phi_vat_lieu=round(vl, 0),
        chi_phi_nhan_cong=round(nc, 0),
        chi_phi_may=round(mtc, 0),
        chi_phi_truc_tiep=round(t, 0),
        chi_phi_chung=round(c_c, 0),
        chi_phi_nha_tam=round(c_nt, 0),
        chi_phi_kxd=round(c_kxd, 0),
        tong_chi_phi_gian_tiep=round(gt, 0),
        thu_nhap_chiu_thue_tinh_truoc=round(tl, 0),
        chi_phi_xay_dung_truoc_thue=round(g, 0),
        thue_vat=round(vat, 0),
        tong_chi_phi_xay_dung_sau_thue=round(g_xd, 0)
    )
    return summary, items


def export_vietnam_estimate_excel(
    summary: CostSummary,
    items: List[CostItem],
    output_path: str
):
    """Xuất file Excel Dự toán chuẩn Thông tư 36/2026/TT-BXD (2 Sheets)."""
    wb = openpyxl.Workbook()

    # Sheet 1: Tổng hợp chi phí xây dựng
    ws1 = wb.active
    ws1.title = "TONG_HOP_DU_TOAN"

    # Sheet 2: Bảng khối lượng dự toán chi tiết
    ws2 = wb.create_sheet(title="DU_TOAN_CHI_TIET")

    font_header = Font(name="Times New Roman", size=11, bold=True)
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_normal = Font(name="Times New Roman", size=11)
    font_title = Font(name="Times New Roman", size=14, bold=True)

    fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # --- ĐIỀN SHEET 1: TỔNG HỢP DỰ TOÁN ---
    ws1.merge_cells("A1:G1")
    ws1["A1"] = "BẢNG TỔNG HỢP DỰ TOÁN CHI PHÍ XÂY DỰNG"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = Alignment(horizontal="center")

    ws1.merge_cells("A2:G2")
    ws1["A2"] = f"Dự án: {summary.project_name} | Loại công trình: {PROJECT_RATES[summary.project_type]['name']}"
    ws1["A2"].font = font_bold
    ws1["A2"].alignment = Alignment(horizontal="center")

    ws1.merge_cells("A3:G3")
    ws1["A3"] = "Căn cứ Thông tư số 36/2026/TT-BXD của Bộ Xây dựng (Đơn vị tính: VNĐ)"
    ws1["A3"].font = Font(name="Times New Roman", size=10, italic=True)
    ws1["A3"].alignment = Alignment(horizontal="center")

    headers_s1 = ["STT", "Khoản mục chi phí", "Cách tính", "Hệ số %", "Ký hiệu", "Giá trị (VNĐ)", "Ghi chú"]
    ws1.append([]) # row 4 trống
    ws1.append(headers_s1) # row 5

    for col in range(1, 8):
        cell = ws1.cell(row=5, column=col)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    rows_s1 = [
        ("I", "Chi phí trực tiếp", "", "", "T", summary.chi_phi_truc_tiep, "VL + NC + M"),
        ("1", "- Chi phí vật liệu", "Bảng dự toán chi tiết", "", "VL", summary.chi_phi_vat_lieu, ""),
        ("2", "- Chi phí nhân công", "Bảng dự toán chi tiết", "", "NC", summary.chi_phi_nhan_cong, ""),
        ("3", "- Chi phí máy thi công", "Bảng dự toán chi tiết", "", "M", summary.chi_phi_may, ""),
        ("II", "Chi phí gián tiếp", "", "", "GT", summary.tong_chi_phi_gian_tiep, "C_C + C_NT + C_KXD"),
        ("1", "- Chi phí chung", "T x Tỷ lệ %", f"{PROJECT_RATES[summary.project_type]['rate_chi_phi_chung']*100:.1f}%", "C_C", summary.chi_phi_chung, "Thông tư 36/2026/TT-BXD"),
        ("2", "- Chi phí nhà tạm để ở và điều hành", "T x Tỷ lệ %", f"{PROJECT_RATES[summary.project_type]['rate_nha_tam']*100:.1f}%", "C_NT", summary.chi_phi_nha_tam, "Công trình xây dựng"),
        ("3", "- Chi phí một số công việc KXD từ thiết kế", "T x Tỷ lệ %", f"{PROJECT_RATES[summary.project_type]['rate_khong_xac_dinh']*100:.1f}%", "C_KXD", summary.chi_phi_kxd, ""),
        ("III", "Thu nhập chịu thuế tính trước", "(T + GT) x Tỷ lệ %", f"{PROJECT_RATES[summary.project_type]['rate_thu_nhap_tinh_truoc']*100:.1f}%", "TL", summary.thu_nhap_chiu_thue_tinh_truoc, "5.5% x (T + GT)"),
        ("", "CHI PHÍ XÂY DỰNG TRƯỚC THUẾ", "T + GT + TL", "", "G", summary.chi_phi_xay_dung_truoc_thue, ""),
        ("IV", "Thuế giá trị gia tăng (VAT)", "G x 10%", "10.0%", "VAT", summary.thue_vat, "Thuế VAT"),
        ("", "TỔNG CỘNG CHI PHÍ XÂY DỰNG SAU THUẾ", "G + VAT", "", "G_XD", summary.tong_chi_phi_xay_dung_sau_thue, "Giá trị dự toán phê duyệt"),
    ]

    for r_idx, r_data in enumerate(rows_s1, start=6):
        ws1.append(list(r_data))
        is_major = r_data[0] in ["I", "II", "III", "IV", ""]
        is_grand = "TỔNG CỘNG" in str(r_data[1])
        for c_idx in range(1, 8):
            cell = ws1.cell(row=r_idx, column=c_idx)
            cell.font = font_bold if is_major else font_normal
            cell.border = thin_border
            if is_grand:
                cell.fill = fill_highlight
            if c_idx == 6 and isinstance(cell.value, (int, float, Decimal)):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")
            elif c_idx in [1, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center")

    # --- ĐIỀN SHEET 2: DỰ TOÁN CHI TIẾT ---
    ws2.merge_cells("A1:J1")
    ws2["A1"] = f"BẢNG DỰ TOÁN CHI TIẾT: {summary.project_name}"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(horizontal="center")

    headers_s2 = [
        "STT", "Mã hiệu ĐM", "Tên công tác xây dựng", "ĐVT", "Khối lượng",
        "Đơn giá VL", "Đơn giá NC", "Đơn giá Máy", "Thành tiền (VNĐ)", "Ghi chú"
    ]
    ws2.append([]) # row 2
    ws2.append(headers_s2) # row 3

    for col in range(1, 11):
        cell = ws2.cell(row=3, column=col)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for it in items:
        row_vals = [
            it.stt, it.ma_hieu, it.ten_cong_tac, it.don_vi, float(it.khoi_luong),
            float(it.don_gia_vat_lieu), float(it.don_gia_nhan_cong), float(it.don_gia_may),
            float(it.thanh_tien_truc_tiep), it.ghi_chu
        ]
        ws2.append(row_vals)
        cur_row = ws2.max_row
        for c_idx in range(1, 11):
            cell = ws2.cell(row=cur_row, column=c_idx)
            cell.font = font_normal
            cell.border = thin_border
            if c_idx == 1:
                cell.alignment = Alignment(horizontal="center")
            elif c_idx in [2, 4]:
                cell.alignment = Alignment(horizontal="center")
            elif c_idx in [5, 6, 7, 8, 9]:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "#,##0.00" if c_idx == 5 else "#,##0"

    # Dòng tổng cộng Sheet 2
    total_row = ws2.max_row + 1
    ws2.cell(row=total_row, column=1, value="")
    ws2.cell(row=total_row, column=2, value="")
    ws2.cell(row=total_row, column=3, value="TỔNG CHI PHÍ TRỰC TIẾP (T)").font = font_bold
    ws2.cell(row=total_row, column=9, value=float(summary.chi_phi_truc_tiep)).font = font_bold
    ws2.cell(row=total_row, column=9).number_format = "#,##0"
    for c in range(1, 11):
        ws2.cell(row=total_row, column=c).border = thin_border
        ws2.cell(row=total_row, column=c).fill = fill_highlight

    # Auto-adjust column widths
    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_path)
    print(f"[V] Đã xuất file Dự toán Xây dựng Việt Nam thành công: {output_path}")
