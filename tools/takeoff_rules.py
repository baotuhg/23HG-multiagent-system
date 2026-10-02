# -*- coding: utf-8 -*-
"""
TAKEOFF RULES — diễn giải khối lượng bê tông & ván khuôn cống hộp từ kích thước hình học.

Mục tiêu: không còn "số chết" trong bảng diễn giải. Kích thước lòng cống, số khoang, cạnh vút và
các quy ước tính ván khuôn là ô đầu vào có nhãn; khối lượng là công thức của các ô đó.

Quy ước đo bóc (cần đối chiếu điều khoản cụ thể với văn bản hiện hành khi lập hồ sơ):
  - Khối lượng đo theo kích thước trong bản vẽ thiết kế; bê tông và ván khuôn tách riêng theo
    chủng loại / cấu kiện (nguyên tắc chung của Thông tư 17/2019/TT-BXD về đo bóc khối lượng).
  - Thông tư 12/2021/TT-BXD là Định mức xây dựng (hao phí), KHÔNG phải quy định đo bóc.
  - Ván khuôn tính theo diện tích bề mặt bê tông tiếp xúc ván khuôn (nguyên tắc thông dụng).
    Mặt nào được coi là có ván khuôn là quy ước của biện pháp thi công → để thành công tắc 1/0.

Mặt cắt cống hộp (đơn vị mét), tính trên 1 m dài:
  diện tích BT   = B_ngoài × H_ngoài − n × b_lòng × h_lòng + n × 4 × ½ × c²       (c: cạnh vút)
  chu vi lòng    = 2(b + h) − 4c(2 − √2)       (mỗi góc vút thay 2c bằng cạnh huyền c√2)
  ván khuôn/m    = k_trong × n × chu_vi_lòng + k_hông × 2H + k_trên × B + k_đáy × B
                   + k_đầu × 2 × diện_tích_BT / L_đốt

Pure Python, Zero LLM.
"""

from __future__ import annotations
import math
from dataclasses import dataclass
from typing import Dict, Iterable, List

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SQRT2 = math.sqrt(2.0)


@dataclass(frozen=True)
class BoxCulvert:
    outer_w: float      # bề rộng ngoài (m)
    outer_h: float      # chiều cao ngoài (m)
    inner_w: float      # bề rộng lòng một khoang (m)
    inner_h: float      # chiều cao lòng một khoang (m)
    cells: int = 1      # số khoang
    haunch: float = 0.0  # cạnh vút trong lòng (m), tam giác vuông cân

    def concrete_area(self) -> float:
        """Diện tích mặt cắt bê tông (m2) = thể tích trên 1 m dài (m3/m)."""
        return (self.outer_w * self.outer_h - self.cells * self.inner_w * self.inner_h
                + self.cells * 4 * 0.5 * self.haunch ** 2)

    def inner_perimeter(self) -> float:
        """Chu vi lòng một khoang (m) có trừ phần góc vút."""
        return 2 * (self.inner_w + self.inner_h) - 4 * self.haunch * (2 - SQRT2)

    def formwork_per_m(self, segment_length: float, inner: bool = True, sides: bool = True,
                       top: bool = False, bottom: bool = False, ends: bool = True) -> Dict[str, float]:
        """Ván khuôn tiếp xúc (m2) trên 1 m dài theo từng mặt và tổng."""
        parts = {
            "trong": self.cells * self.inner_perimeter() if inner else 0.0,
            "hong_ngoai": 2 * self.outer_h if sides else 0.0,
            "mat_tren": self.outer_w if top else 0.0,
            "mat_day": self.outer_w if bottom else 0.0,
            "dau_dot": 2 * self.concrete_area() / segment_length if ends else 0.0,
        }
        parts["tong"] = sum(parts.values())
        return parts


# ─────────────────────────────────────────────────────────────────────────────
# ÁP VÀO SHEET DIỄN GIẢI (QS_DIEN_GIAI_CHI_TIET / QS_TAKEOFF)
# ─────────────────────────────────────────────────────────────────────────────

# Dòng bê tông thân → dòng ván khuôn tương ứng, và đầu vào của từng loại cống (K,L,M,N)
CULVERT_ROWS = {
    8: {"formwork_row": 11, "inputs": (2.0, 2.0, 1, 0.20)},
    9: {"formwork_row": 12, "inputs": (3.0, 3.0, 1, 0.25)},
    10: {"formwork_row": 13, "inputs": (3.0, 3.0, 2, 0.25)},
}
FLAG_FIRST_ROW = 17      # C17:C21 = công tắc mặt ván khuôn
FLAGS = [
    ("Ván khuôn mặt trong (lòng cống)", 1),
    ("Ván khuôn mặt hông ngoài (2 bên)", 1),
    ("Ván khuôn mặt trên (bản nắp)", 0),
    ("Ván khuôn mặt đáy (bản đáy)", 0),
    ("Ván khuôn đầu đốt (2 đầu mỗi đốt)", 1),
]
NOTE = ("Quy ước: ván khuôn tính theo diện tích bề mặt bê tông tiếp xúc ván khuôn; mặt nào có ván khuôn do biện pháp "
        "thi công quyết định (công tắc 1/0 bên dưới). Nguyên tắc đo bóc theo Thông tư 17/2019/TT-BXD (cần đối chiếu "
        "điều khoản cụ thể); Thông tư 12/2021/TT-BXD là định mức, không phải đo bóc.")


def apply_culvert_derivation(ws) -> None:
    """Chuyển bảng diễn giải cống hộp A5 sang đầu vào có tên (sửa tại chỗ, giữ nguyên vị trí cột I).

    - Dòng bê tông thân (8-10): thêm K (rộng lòng), L (cao lòng), M (số khoang), N (cạnh vút);
      cột I trở thành công thức của các ô đó (giá trị không đổi so với công thức cũ).
    - Dòng ván khuôn (11-13): GIỮ NGUYÊN số đang dùng ở cột F (không tự đổi khối lượng hợp đồng);
      thêm P..S tính lại theo hình học, T = % chênh lệch, U = kết luận để kỹ sư QS quyết định.
    - Khối công tắc quy ước ván khuôn ở A15:C21 và ghi chú căn cứ.
    """
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    inp_fill = PatternFill("solid", fgColor="FFF2CC")   # ô đầu vào
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = {
        "K": "RỘNG LÒNG b (m)", "L": "CAO LÒNG h (m)", "M": "SỐ KHOANG n", "N": "CẠNH VÚT c (m)",
        "P": "VK TRONG (m2/m)", "Q": "VK HÔNG NGOÀI (m2/m)", "R": "VK ĐẦU ĐỐT (m2/m)",
        "S": "VK TÍNH LẠI (m2/m)", "T": "CHÊNH LỆCH SO VỚI SỐ ĐANG DÙNG (cột F)", "U": "KẾT LUẬN",
    }
    for col, text in headers.items():
        c = ws[f"{col}4"]
        c.value, c.font, c.fill, c.alignment, c.border = text, hdr_font, hdr_fill, center, border
        ws.column_dimensions[col].width = 18

    f_in, f_side, f_top, f_bot, f_end = (f"$C${FLAG_FIRST_ROW + k}" for k in range(5))
    for crow, spec in CULVERT_ROWS.items():
        for col, val in zip("KLMN", spec["inputs"]):
            c = ws[f"{col}{crow}"]
            c.value, c.fill, c.border = val, inp_fill, border
            c.number_format = "0" if col == "M" else "0.00"
        area = f"(F{crow}*G{crow}-M{crow}*K{crow}*L{crow}+M{crow}*4*0.5*N{crow}^2)"
        ws[f"I{crow}"].value = f"=D{crow}*E{crow}*{area}"

        r = spec["formwork_row"]
        ws[f"P{r}"].value = f"={f_in}*M{crow}*(2*(K{crow}+L{crow})-4*N{crow}*(2-SQRT(2)))"
        ws[f"Q{r}"].value = f"={f_side}*2*G{crow}+{f_top}*F{crow}+{f_bot}*F{crow}"
        ws[f"R{r}"].value = f"={f_end}*2*{area}/E{crow}"
        ws[f"S{r}"].value = f"=P{r}+Q{r}+R{r}"
        ws[f"T{r}"].value = f"=S{r}/F{r}-1"
        ws[f"U{r}"].value = f'=IF(ABS(T{r})>0.05,"LỆCH >5% - QS xác nhận","Khớp (<=5%)")'
        for col in "PQRSTU":
            c = ws[f"{col}{r}"]
            c.border = border
            c.number_format = "0.0%" if col == "T" else "0.000"
        ws[f"J{r}"].value = "Cột F = VK m2/m ĐANG DÙNG (không phải bề rộng); P..U là tính lại theo hình học"

    ws["A15"].value = "QUY ƯỚC VÁN KHUÔN (1 = có ván khuôn, 0 = không) - đổi ô vàng, cột P..U của dòng 11-13 tự tính lại"
    ws["A15"].font = Font(bold=True)
    for k, (label, default) in enumerate(FLAGS):
        r = FLAG_FIRST_ROW + k
        ws[f"B{r}"].value = label
        c = ws[f"C{r}"]
        c.value, c.fill, c.border, c.alignment = default, inp_fill, border, center
    note_row = FLAG_FIRST_ROW + len(FLAGS) + 1
    ws[f"A{note_row}"].value = NOTE
    ws[f"A{note_row}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=10)
    ws.row_dimensions[note_row].height = 62
