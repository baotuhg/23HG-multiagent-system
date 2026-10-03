"""Master cống hộp tuyến A5 (Vincons - OLP_SD_KC_HT_CH_06..22) -> Hub & Spoke 5 gói.

Nguồn: 20 trang bản vẽ shop drawing đã OCR (Marker). KÍCH THƯỚC lấy từ OCR (ngoại 2.5/trong 2.0 cho cống 2x2;
ngoại 3.6/trong 3.0 cho cống 3x3; đốt L=11.3 m). KHÔNG có trong bản vẽ: tổng chiều dài tuyến, số đốt, số ga.
=> Mặc định 1 đốt mỗi loại (MẪU). Sửa ô vàng C6, C7 (số đốt) thì QS + tiến độ tự cập nhật.
Thép: chỉ liệt kê ký hiệu đọc từ OCR (CẦN ĐỐI CHIẾU), KHÔNG tự suy số thanh.

Dùng: python examples/build_cong_hop_a5_master.py --out <Master.xlsx> [--start 2026-09-05 --deadline 2026-11-12]
"""
import argparse
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_km19_vina_hub_schedule import build  # noqa: E402

BOLD = Font(bold=True)
YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="1F497D")
WARN = "⚠ MẪU: 1 đốt/loại, kích thước từ OCR bản vẽ — cần đối chiếu bản gốc và nhập số đốt thật vào ô vàng."


def header(ws, row, names):
    for i, n in enumerate(names, 1):
        c = ws.cell(row, i, n)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), HEAD
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def make_qs(wb):
    ws = wb.create_sheet("QS_DIEN_GIAI_CHI_TIET")
    ws["A1"] = "BÓC TÁCH KHỐI LƯỢNG CỐNG HỘP TUYẾN A5 (OLP_SD_KC_HT_CH_06..22)"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = WARN
    ws["A3"] = "Ô vàng = đầu vào. Giả định (không có trên bản vẽ OCR) ghi rõ ở cột Ghi chú."
    header(ws, 5, ["Hạng mục", "Công thức", "Giá trị", "ĐV", "Ghi chú"])
    rows = [
        (6, "Số đốt cống 2.0x2.0", "nhập", 1, "đốt", "MẪU — chưa có tổng chiều dài tuyến", True),
        (7, "Số đốt cống 3.0x3.0", "nhập", 1, "đốt", "MẪU — chưa có tổng chiều dài tuyến", True),
        (8, "Chiều dài 1 đốt", "bản vẽ", 11.3, "m", "OCR: L=11300 (khe co giãn mỗi 11,7 m)", True),
        (9, "Kích thước ngoài 2x2", "bản vẽ", 2.5, "m", "OCR: 2500", True),
        (10, "Kích thước trong 2x2", "bản vẽ", 2.0, "m", "OCR: 2000", True),
        (11, "Kích thước ngoài 3x3", "bản vẽ", 3.6, "m", "OCR: 3600 / 3520 — cần đối chiếu", True),
        (12, "Kích thước trong 3x3", "bản vẽ", 3.0, "m", "OCR: 3000", True),
        (13, "Dày bê tông lót M100", "bản vẽ", 0.05, "m", "Ghi chú bản vẽ: lót dày 50 mm", True),
        (14, "Lót mở rộng mỗi bên", "GIẢ ĐỊNH", 0.1, "m", "Không đọc được từ OCR", True),
        (15, "Dày lớp thay cát đầm chặt", "bản vẽ", 0.3, "m", "Ghi chú: thay 30 cm bằng cát đầm chặt K0.9", True),
        (16, "Mở rộng nền mỗi bên", "bản vẽ", 0.3, "m", "Ghi chú: phạm vi = bề rộng móng + 300 mm mỗi phía", True),
        (18, "BT B20 (M250) cống 2x2", "(B²−b²)·L·N", "=(C9*C9-C10*C10)*C8*C6", "m3", "", False),
        (19, "BT B20 (M250) cống 3x3", "(B²−b²)·L·N", "=(C11*C11-C12*C12)*C8*C7", "m3", "", False),
        (20, "TỔNG BT B20", "", "=C18+C19", "m3", "", False),
        (21, "Ván khuôn cống 2x2", "(4B+4b)·L·N", "=(4*C9+4*C10)*C8*C6", "m2", "chưa tính mặt đầu đốt", False),
        (22, "Ván khuôn cống 3x3", "(4B+4b)·L·N", "=(4*C11+4*C12)*C8*C7", "m2", "chưa tính mặt đầu đốt", False),
        (23, "TỔNG VÁN KHUÔN", "", "=C21+C22", "m2", "", False),
        (24, "BT lót M100 cống 2x2", "(B+2e)·t·L·N", "=(C9+2*C14)*C13*C8*C6", "m3", "", False),
        (25, "BT lót M100 cống 3x3", "(B+2e)·t·L·N", "=(C11+2*C14)*C13*C8*C7", "m3", "", False),
        (26, "TỔNG BT LÓT", "", "=C24+C25", "m3", "", False),
        (27, "Cát thay nền cống 2x2", "(B+2w)·h·L·N", "=(C9+2*C16)*C15*C8*C6", "m3", "", False),
        (28, "Cát thay nền cống 3x3", "(B+2w)·h·L·N", "=(C11+2*C16)*C15*C8*C7", "m3", "", False),
        (29, "TỔNG CÁT THAY NỀN", "", "=C27+C28", "m3", "", False),
        (30, "Tổng chiều dài cống", "L·(N2+N3)", "=C8*(C6+C7)", "m", "", False),
        (31, "Số khe co giãn (ước)", "N2+N3", "=C6+C7", "khe", "mỗi đốt 1 khe", False),
    ]
    for r, name, f, v, u, note, inp in rows:
        ws.cell(r, 1, name)
        ws.cell(r, 2, f)
        ws.cell(r, 3, v)
        ws.cell(r, 4, u)
        ws.cell(r, 5, note)
        if inp:
            ws.cell(r, 3).fill = YELLOW
        if name.startswith("TỔNG"):
            for c in range(1, 4):
                ws.cell(r, c).font = BOLD
    ws["A17"] = "KẾT QUẢ (công thức sống)"
    ws["A17"].font = BOLD
    for col, w in zip("ABCDE", (34, 20, 14, 8, 55)):
        ws.column_dimensions[col].width = w


BARS = [
    # (trang, vị trí, Φ, a, L m, ghi chú)
    (2, "Cống 2x2 - thép ôm ngoài", 12, 250, 2.92, "2x450Φ12a250 (OCR)"),
    (2, "Cống 2x2 - thép đáy/nắp", 12, 100, 2.82, "2x1130Φ12a100 (OCR)"),
    (2, "Cống 2x2 - thép thành", 12, 250, 2.84, "2x450Φ12a250 (OCR)"),
    (2, "Cống 2x2 - thép thành dày", 12, 125, 2.84, "2x900Φ12a125 (OCR)"),
    (2, "Cống 2x2 - thép gia cường", 12, 250, 1.32, "4x450Φ12a250 (OCR)"),
    (2, "Cống 2x2 - thép phân bố dọc", 10, 200, 11.7, "Φ10a200 L=11700 (OCR)"),
    (3, "Hố lắng cặn 2x2 - thép sàn", 12, 100, 2.94, "200Φ12a100 (OCR)"),
    (3, "Hố lắng cặn 2x2 - thép sàn", 12, 250, 2.94, "160Φ12a250 (OCR)"),
    (3, "Hố lắng cặn 2x2 - phân bố", 10, 200, 3.9, "100Φ10a200 L=3900 (OCR)"),
    (5, "Cống 3x3 - thép chính", 14, 250, 3.9, "2x450Φ14a250 (OCR)"),
    (5, "Cống 3x3 - thép chính dày", 14, 125, 3.9, "2x900Φ14a125 (OCR)"),
    (5, "Cống 3x3 - thép gia cường", 14, 250, 1.95, "4x450Φ14a250 (OCR)"),
    (5, "Cống 3x3 - thép phân bố dọc", 10, 200, 11.7, "Φ10a200 L=11700 (OCR)"),
    (6, "Hố lắng cặn 3x3 - thép chính", 14, 250, 4.31, "2x450Φ14a250 (OCR)"),
    (6, "Hố lắng cặn 3x3 - thép chính dày", 14, 125, 4.31, "2x900Φ14a125 (OCR)"),
    (6, "Hố lắng cặn 3x3 - gia cường", 14, 250, 1.17, "450Φ14a250 (OCR)"),
]


def make_bbs(wb):
    ws = wb.create_sheet("THONG_KE_THEP_CHI_TIET")
    ws["A1"] = "THỐNG KÊ THÉP CỐNG HỘP A5 — KÝ HIỆU ĐỌC TỪ OCR"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("⚠ OCR sai nhiều ký hiệu thép. Cột 'Số thanh' ĐỂ TRỐNG (nhập sau khi đối chiếu bản vẽ gốc); "
                "hệ thống không tự suy số thanh. Thép <Φ10: CB240-T; ≥Φ10: CB500-V (ghi chú bản vẽ).")
    header(ws, 5, ["Mã", "Trang BV", "Vị trí", "Φ (mm)", "a (mm)", "L 1 thanh (m)", "Số thanh",
                   "KL 1m (kg/m)", "KL 1 thanh (kg)", "Tổng KL (kg)", "Nguồn OCR", "Trạng thái"])
    for i, (pg, pos, d, a, L, note) in enumerate(BARS):
        r = 6 + i
        ws.cell(r, 1, f"T{i + 1:02d}")
        ws.cell(r, 2, pg)
        ws.cell(r, 3, pos)
        ws.cell(r, 4, d)
        ws.cell(r, 5, a)
        ws.cell(r, 6, L)
        ws.cell(r, 7).fill = YELLOW
        ws.cell(r, 8, f"=0.00617*D{r}*D{r}")
        ws.cell(r, 9, f"=F{r}*H{r}")
        ws.cell(r, 10, f'=IF(G{r}="",0,G{r}*I{r})')
        ws.cell(r, 11, note)
        ws.cell(r, 12, "CẦN ĐỐI CHIẾU")
        ws.cell(r, 8).number_format = "0.000"
        ws.cell(r, 9).number_format = ws.cell(r, 10).number_format = "0.00"
    last = 5 + len(BARS)
    ws.cell(last + 1, 3, "TỔNG (kg)").font = BOLD
    ws.cell(last + 1, 10, f"=SUM(J6:J{last})").font = BOLD
    for col, w in zip("ABCDEFGHIJKL", (7, 9, 36, 8, 8, 12, 10, 12, 14, 14, 30, 16)):
        ws.column_dimensions[col].width = w


# (WBS, tên, qty, đv, định mức công/ĐV, tổ đội, tiền nhiệm, group)
TASKS = [
    ("1", "CHUẨN BỊ", None),
    ("1.1", "Định vị tuyến, rào chắn, lắp dựng ĐD", ("=QS_DIEN_GIAI_CHI_TIET!C30", "m", 0.15, 4, "-")),
    ("2", "NỀN MÓNG", None),
    ("2.1", "Đào bóc 30 cm nền hiện trạng", ("=QS_DIEN_GIAI_CHI_TIET!C29", "m3", 0.3, 6, "1.1FS")),
    ("2.2", "Đệm cát đầm chặt K0.9", ("=QS_DIEN_GIAI_CHI_TIET!C29", "m3", 0.4, 6, "2.1FS")),
    ("2.3", "Bê tông lót M100 dày 50", ("=QS_DIEN_GIAI_CHI_TIET!C26", "m3", 1.5, 6, "2.2FS")),
    ("3", "THÂN CỐNG", None),
    ("3.1", "Gia công, lắp dựng cốt thép", ("=QS_DIEN_GIAI_CHI_TIET!C20", "m3 BT", 2.5, 8, "2.3FS")),
    ("3.2", "Lắp dựng ván khuôn", ("=QS_DIEN_GIAI_CHI_TIET!C23", "m2", 0.3, 8, "3.1SS+2d")),
    ("3.3", "Đổ bê tông B20 thân cống", ("=QS_DIEN_GIAI_CHI_TIET!C20", "m3", 0.6, 10, "3.2FS")),
    ("3.4", "Bảo dưỡng, tháo ván khuôn (7 ngày)", (7, "ngày", 1, 1, "3.3FS")),
    ("4", "HOÀN THIỆN", None),
    ("4.1", "Khe co giãn: xốp, bao tải nhựa đường", ("=QS_DIEN_GIAI_CHI_TIET!C31", "khe", 1.0, 3, "3.4FS")),
    ("4.2", "Đắp cát hai bên, hoàn trả", ("=QS_DIEN_GIAI_CHI_TIET!C29", "m3", 0.3, 6, "4.1FS")),
    ("5", "NGHIỆM THU", None),
    ("5.1", "Nghiệm thu KCS, hoàn công", (1, "hồ sơ", 2, 1, "4.2FS")),
]


def make_schedule(wb):
    ws = wb.create_sheet("TIEN_DO_THI_CONG_WBS")
    ws["B2"] = "TIẾN ĐỘ THI CÔNG CỐNG HỘP TUYẾN A5 — chuẩn Vina Alpha (công thức sống)"
    ws["B2"].font = Font(bold=True, size=13)
    ws["B3"] = "⚠ Định mức công/ĐV và tổ đội là GIẢ ĐỊNH để chạy thử — thay bằng định mức thực tế của nhà thầu."
    header(ws, 6, ["WBS", "Hạng mục", "Khối lượng", "ĐV", "Định mức", "Tổng công", "Tổ đội", "Số ngày",
                   "Bắt đầu", "Kết thúc", "Tiền nhiệm", "Găng"])
    r = 7
    for wbs, name, t in TASKS:
        ws.cell(r, 1, wbs)
        ws.cell(r, 2, name)
        if t is None:
            ws.cell(r, 2).font = BOLD
        else:
            qty, unit, norm, crew, pred = t
            ws.cell(r, 3, qty)
            ws.cell(r, 4, unit)
            ws.cell(r, 5, norm)
            ws.cell(r, 6, f"=C{r}*E{r}")
            ws.cell(r, 7, crew)
            ws.cell(r, 8, f"=MAX(1,ROUNDUP(F{r}/G{r},0))")
            ws.cell(r, 11, pred)
        r += 1
    for col, w in zip("ABCDEFGHIJKL", (6, 40, 11, 8, 9, 10, 8, 9, 12, 12, 12, 8)):
        ws.column_dimensions[col].width = w


def make_kcs(wb):
    ws = wb.create_sheet("HOSO_KCS_NGHIEM_THU")
    ws["A1"] = "DANH MỤC NGHIỆM THU KCS — CỐNG HỘP A5 (ngày lấy từ tiến độ)"
    ws["A1"].font = Font(bold=True, size=13)
    header(ws, 5, ["STT", "Công việc nghiệm thu", "Biên bản", "Trạng thái", "Ngày bắt đầu", "Ngày kết thúc"])
    items = [("Nghiệm thu nền móng, đệm cát", 9, 10), ("Nghiệm thu BT lót", 10, 11),
             ("Nghiệm thu cốt thép", 12, 13), ("Nghiệm thu ván khuôn", 13, 14),
             ("Nghiệm thu đổ bê tông thân cống", 14, 15), ("Nghiệm thu khe co giãn", 17, 18),
             ("Nghiệm thu hoàn thành hạng mục", 20, 20)]
    sched_rows = {"Nghiệm thu nền móng, đệm cát": 11, "Nghiệm thu BT lót": 12, "Nghiệm thu cốt thép": 14,
                  "Nghiệm thu ván khuôn": 15, "Nghiệm thu đổ bê tông thân cống": 16,
                  "Nghiệm thu khe co giãn": 19, "Nghiệm thu hoàn thành hạng mục": 22}
    for i, (n, _, _) in enumerate(items):
        r = 6 + i
        sr = sched_rows[n]
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, n)
        ws.cell(r, 3, f"BB-{i + 1:02d}")
        ws.cell(r, 4, "Chờ lập biên bản")
        ws.cell(r, 5, f"=TIEN_DO_THI_CONG_WBS!I{sr}")
        ws.cell(r, 6, f"=TIEN_DO_THI_CONG_WBS!J{sr}")
    for col, w in zip("ABCDEF", (6, 40, 12, 14, 14, 14)):
        ws.column_dimensions[col].width = w


def main(out, start, deadline):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    make_qs(wb)
    make_bbs(wb)
    make_schedule(wb)
    make_kcs(wb)
    base = str(out) + ".base.xlsx"
    wb.save(base)
    build(str(out), start, deadline, template=base)
    Path(base).unlink()
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", default="2026-09-05")
    ap.add_argument("--deadline", default="2026-11-12")
    a = ap.parse_args()
    print(main(a.out, date.fromisoformat(a.start), date.fromisoformat(a.deadline)))
