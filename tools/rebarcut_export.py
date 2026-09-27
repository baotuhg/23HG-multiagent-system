# -*- coding: utf-8 -*-
"""
REBARCUT EXPORT — Xuất kết quả cắt thép theo bố cục RebarCut Pro Excel (.xlsx)

Sheet:
  INPUT     — tham số (B6:B15) và BBS từ dòng 19 đúng cột của RebarCut, có thể
              copy vùng B19:K sang RebarCut Pro hoặc đọc lại bằng --bbs
  SO_SANH   — so sánh phương án tối ưu (không nối) và phương án nối thép
  PA_TOI_UU — sơ đồ cắt phương án tối ưu OR-Tools (không nối)
  PA_NOI    — sơ đồ cắt phương án nối thép (nếu có)
  MOI_NOI   — chi tiết mối nối đề xuất (cần kỹ thuật duyệt)
  REMAIN    — đầu thừa theo từng phương án
  CHI_TIET  — từng đoạn cắt của từng cây
"""

from __future__ import annotations
from collections import defaultdict
from typing import Dict, List, Optional

from tools.cutting_stock_solver import CuttingStockSolution, CutDemand, _grouped_cuts

HEADER_FILL = "1F3A5E"


def _zones_text(zones) -> str:
    if not zones:
        return ""
    return "; ".join(f"{a:g}-{b:g}" for a, b in zones)


def _style_header(ws, row: int, ncols: int) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def _plan_sheet(wb, name: str, title: str, sol: CuttingStockSolution) -> None:
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A3"] = f"{sol.status} — {sol.solver_name}"
    ws["A4"], ws["B4"] = "Cây", sol.total_bars_needed
    ws["C4"], ws["D4"] = "Cận dưới", sol.lower_bound_bars
    ws["E4"], ws["F4"] = "Đề-xê (%)", round(sol.waste_ratio_pct, 2)
    ws["A5"], ws["B5"] = "Mua kg", sol.total_weight_kg
    ws["C5"], ws["D5"] = "Mẫu cắt", len(sol.patterns)
    ws["E5"], ws["F5"] = "Số nối", sol.total_splices
    headers = ["STT", "Ø", "Lần cắt", "SƠ ĐỒ TỐI ƯU HÓA CẮT THÉP", "Số cây cần mua",
               "Tổng chiều dài/cây (m)", "Thép thừa/cây (m)", "Tổng thừa (m)", "Phân loại đầu thừa", "Ghi chú"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=7, column=c, value=h)
    _style_header(ws, 7, len(headers))
    usable = sol.bar_length_mm - 2 * sol.end_trim_mm
    for r, p in enumerate(sol.patterns, start=1):
        scheme = " + ".join(f"{n}*[{mark}] {length / 1000:.3f}m" for length, n, mark in _grouped_cuts(p))
        ws.append([r, p.diameter_mm, r, scheme, p.bars_used, (usable - p.waste_mm) / 1000,
                   p.waste_mm / 1000, p.waste_mm * p.bars_used / 1000, p.offcut_class,
                   f"Kerf {sol.kerf_mm}mm; cắt đầu {2 * sol.end_trim_mm}mm/cây"
                   + (f"; mác {p.grade}" if p.grade else "")])
    ws.column_dimensions["D"].width = 70


def write_rebarcut_workbook(
    path: str,
    demands: List[CutDemand],
    base: CuttingStockSolution,
    spliced: Optional[CuttingStockSolution] = None,
    params: Optional[Dict] = None,
) -> None:
    import openpyxl

    params = params or {}
    wb = openpyxl.Workbook()

    # ── INPUT ────────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "INPUT"
    ws["A1"] = "TỐI ƯU CẮT THÉP — xuất từ 23HG MultiAgent System (bố cục RebarCut Pro Excel)"
    ws["A5"] = "THÔNG SỐ CÂY THÉP / HIỆN TRƯỜNG"
    rows = [
        ("Chiều dài cây nguyên (m)", base.bar_length_mm / 1000),
        ("Hao hụt nhát cắt / kerf (mm)", base.kerf_mm),
        ("Cắt bỏ đầu cây mỗi đầu (mm)", base.end_trim_mm),
        ("Đầu thừa tái sử dụng (xD)", params.get("reuse_xd", 100)),
        ("Chiều dài kiểm soát tối thiểu (xD)", params.get("short_xd", 20)),
        ("Chiều dài nối chồng mặc định (xD)", params.get("lap_xd", 40)),
        ("Đoạn nối tối thiểu mỗi phía (xD)", params.get("min_splice_segment_xd", 20)),
        ("Tỷ lệ tối đa thanh được nối (%)", params.get("max_splice_ratio", 0.5)),
        ("PA3: tối đa đoạn cắt / cây", params.get("max_pieces_per_bar") or 10000),
        ("PA3: tối đa Bar Mark / cây", params.get("max_marks_per_bar") or 300),
    ]
    for offset, (label, value) in enumerate(rows):
        ws.cell(row=6 + offset, column=1, value=label)
        ws.cell(row=6 + offset, column=2, value=value)
    ws["A17"] = "INPUT – BẢNG THỐNG KÊ THÉP / BBS"
    headers = ["STT", "Số hiệu / Bar Mark", "Ø (mm)", "Số lượng", "Chiều dài (m)", "Cho nối?",
               "Kiểu nối", "L nối xD", "Max nối %", "Nhóm / vùng nối", "Ghi chú", "Mác thép"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=18, column=c, value=h)
    _style_header(ws, 18, len(headers))
    for r, d in enumerate(demands, start=1):
        ws.append([
            r, d.mark, d.diameter_mm, d.quantity, d.length_mm / 1000,
            "Có" if d.splice_allowed else "Không",
            d.splice_kind or ("Nối chồng" if d.splice_allowed else ""),
            d.lap_xd, d.max_splice_ratio, _zones_text(d.splice_zones), "", d.grade,
        ])
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["A"].width = 36

    # ── SO_SANH ──────────────────────────────────────────────────────────────
    ws = wb.create_sheet("SO_SANH")
    ws["A1"] = "SO SÁNH PHƯƠNG ÁN TỐI ƯU CẮT THÉP"
    headers = ["Phương án", "Số cây", "Cận dưới", "Trạng thái", "Giảm cây vs PA tối ưu", "Tận dụng BBS",
               "Phế + hao hụt (m)", "Đầu thừa tái dùng (m)", "Khối lượng mua (kg)", "Số pattern",
               "Số mối nối", "Thép tăng do nối (m)", "Nhận xét"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=5, column=c, value=h)
    _style_header(ws, 5, len(headers))
    plans = [("PA tối ưu – OR-Tools (không nối)", base,
              "Column Generation + CP-SAT; OPTIMAL = đã chứng minh ít cây nhất")]
    if spliced is not None:
        plans.append(("PA nối thép – đề xuất", spliced,
                      "Chỉ nối Bar Mark có vùng cho phép nối; CẦN KỸ THUẬT DUYỆT"))
    for label, sol, note in plans:
        capacity = sol.total_bars_needed * sol.bar_length_mm
        net = sum(length for a in sol.assignment for length in a["cuts_mm"]) - sol.splice_extra_mm
        reusable = sol.offcut_summary.get("Tái sử dụng", [0, 0])[1]
        ws.append([label, sol.total_bars_needed, sol.lower_bound_bars, sol.status,
                   base.total_bars_needed - sol.total_bars_needed,
                   round(net / capacity, 4) if capacity else 0,
                   (sol.total_waste_mm - reusable) / 1000, reusable / 1000, sol.total_weight_kg,
                   len(sol.patterns), sol.total_splices, sol.splice_extra_mm / 1000, note])
    ws["A9"] = "LƯU Ý KỸ THUẬT CHO PHƯƠNG ÁN NỐI"
    ws["A10"] = ("Chỉ là đề xuất vật tư. Phải kiểm tra vị trí nối, vùng cấm nối, tỷ lệ nối trên cùng mặt cắt, "
                 "chiều dài nối/neo, khoảng cách so le mối nối và được TVGS/thiết kế chấp thuận.")
    ws.column_dimensions["A"].width = 38

    # ── PA sheets ─────────────────────────────────────────────────────────────
    _plan_sheet(wb, "PA_TOI_UU", "PA TỐI ƯU – OR-TOOLS (KHÔNG NỐI)", base)
    if spliced is not None:
        _plan_sheet(wb, "PA_NOI", "PA NỐI THÉP – ĐỀ XUẤT, CẦN KỸ THUẬT DUYỆT", spliced)

        ws = wb.create_sheet("MOI_NOI")
        ws["A1"] = "CHI TIẾT MỐI NỐI ĐỀ XUẤT"
        ws["A2"] = "Vùng nối đo từ đầu thanh (đầu có đoạn ghi ở cột 'Đoạn ở đầu thanh'). Cần duyệt theo hồ sơ dự án."
        headers = ["Mã nối", "Ø (mm)", "Bar Mark", "Chiều dài thanh (m)", "Đoạn A (m)", "Đoạn B (m)",
                   "Bù nối (m)", "Kiểm tra A+B−L (m)", "Kiểu nối", "Cây A / B", "Vùng nối từ (m)",
                   "Vùng nối đến (m)", "Đoạn ở đầu thanh", "Kỹ thuật duyệt"]
        for c, h in enumerate(headers, start=1):
            ws.cell(row=4, column=c, value=h)
        _style_header(ws, 4, len(headers))
        for s in spliced.splices:
            ws.append([s["splice_id"], s["diameter_mm"], s["mark"], s["bar_length_mm"] / 1000,
                       s["a_mm"] / 1000, s["b_mm"] / 1000, s["lap_mm"] / 1000,
                       (s["a_mm"] + s["b_mm"] - s["lap_mm"] - s["bar_length_mm"]) / 1000, s["kind"],
                       f"C{s['bar_a']} / C{s['bar_b']}", s["lap_from_mm"] / 1000, s["lap_to_mm"] / 1000,
                       s["first_segment"], ""])

    # ── REMAIN ───────────────────────────────────────────────────────────────
    ws = wb.create_sheet("REMAIN")
    ws["A1"] = "KHO ĐẦU THỪA / REMAIN"
    headers = ["Phương án", "Ø", "Chiều dài đầu thừa (m)", "Số lượng", "Tổng dài (m)", "Phân loại", "Nguồn"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=4, column=c, value=h)
    _style_header(ws, 4, len(headers))
    for label, sol in [("PA_TOI_UU", base)] + ([("PA_NOI", spliced)] if spliced is not None else []):
        for r, p in enumerate(sol.patterns, start=1):
            if p.waste_mm > 0:
                ws.append([label, p.diameter_mm, p.waste_mm / 1000, p.bars_used,
                           p.waste_mm * p.bars_used / 1000, p.offcut_class, f"{label} - mẫu {r}"])

    # ── CHI_TIET ─────────────────────────────────────────────────────────────
    ws = wb.create_sheet("CHI_TIET")
    ws["A1"] = "CHI TIẾT TỪNG ĐOẠN CẮT"
    headers = ["Phương án", "Mã cây", "Bar Mark", "Ø (mm)", "Dài cắt (m)", "Mã nối / phía",
               "Dư cuối cây (m)", "Cây gốc (m)", "Cắt 2 đầu (m)"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=4, column=c, value=h)
    _style_header(ws, 4, len(headers))
    for label, sol in [("PA_TOI_UU", base)] + ([("PA_NOI", spliced)] if spliced is not None else []):
        seg_ids: Dict[int, List[List]] = defaultdict(list)   # cây → [(dài, phía, mã nối)]
        for s in sol.splices:
            seg_ids[s["bar_a"]].append([s["a_mm"], "A", f"{s['splice_id']}-A"])
            seg_ids[s["bar_b"]].append([s["b_mm"], "B", f"{s['splice_id']}-B"])
        for a in sol.assignment:
            pending = seg_ids.get(a["bar_id"], [])
            for length, mark in zip(a["cuts_mm"], a["marks"]):
                tag = ""
                if "(nối-" in mark:
                    side = mark[-2]
                    match = next((e for e in pending if e[0] == length and e[1] == side), None)
                    if match:
                        pending.remove(match)
                        tag = match[2]
                ws.append([label, f"C{a['bar_id']}", mark, a["diameter_mm"], length / 1000, tag,
                           a["waste_mm"] / 1000, sol.bar_length_mm / 1000, 2 * sol.end_trim_mm / 1000])

    wb.save(path)
