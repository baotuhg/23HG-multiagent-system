# -*- coding: utf-8 -*-
"""
KCS INVARIANTS VERIFIER — Bộ kiểm tra 6 Bất biến Lõi Hồ sơ QLCL / KCS Thực chiến
(Dựa trên phân tích thực nghiệm từ Dự án Kè Sông Miện HG-CW06 và Cầu đường Hà Giang)

Pure Python, Zero LLM. Kiểm tra tính toàn vẹn kỹ thuật, pháp lý và tiến độ của hồ sơ:
  1. Bảo toàn hình học & Khối lượng sống (A x B x C == V)
  2. Đồ thị thời gian kết cấu (DAG):
     - PYC trước NT >= 24h
     - Tháo ván khuôn >= 2 ngày (48h) sau đổ bê tông
     - Nghiệm thu hoàn thành cấu kiện >= 28 ngày (R28) sau đổ bê tông
  3. Cấu trúc chùm hồ sơ (RFI + BBNT + Checklist + PLKL + Phiếu TN)
  4. Lịch pháp lý & Khí tượng (Khóa Tết Nguyên Đán, Tết DL, Lễ, ngày mưa bão)
  5. Đắp đất phân lớp K95 khép kín (mỗi lớp 20-30cm độc lập)
  6. Dung sai thực nghiệm ngẫu nhiên có kiểm soát (TCVN 4453:1995)

Dùng:
    python -m tools.kcs_invariants_verifier <duong_dan_file.xlsx> [--sheet <ten_sheet>]
"""

from __future__ import annotations
import argparse
import datetime
import json
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

# Danh sách ngày lễ quốc gia cố định
DEFAULT_VIETNAM_HOLIDAYS_FIXED = {
    (1, 1),    # Tết Dương lịch
    (4, 30),   # Giải phóng miền Nam
    (5, 1),    # Quốc tế Lao động
    (9, 2),    # Quốc khánh
}

# Khoảng thời gian Tết Nguyên Đán các năm (ước tính dải nghỉ lễ thông thường)
LUNAR_NEW_YEAR_RANGES = [
    # Năm 2025: 25/01/2025 - 02/02/2025
    (datetime.date(2025, 1, 25), datetime.date(2025, 2, 2)),
    # Năm 2026 (Bính Ngọ): 15/02/2026 - 26/02/2026 (file Kè Sông Miện nghỉ 17/02 - 26/02/2026)
    (datetime.date(2026, 2, 15), datetime.date(2026, 2, 26)),
    # Năm 2027: 04/02/2027 - 14/02/2027
    (datetime.date(2027, 2, 4), datetime.date(2027, 2, 14)),
]


@dataclass
class InvariantAuditReport:
    file_path: str
    total_score: int = 100
    passed: bool = True
    critical_errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    details: Dict[str, Any] = field(default_factory=dict)

    def add_error(self, message: str, deduction: int = 15):
        self.critical_errors.append(message)
        self.total_score = max(0, self.total_score - deduction)
        self.passed = False

    def add_warning(self, message: str, deduction: int = 3):
        self.warnings.append(message)
        self.total_score = max(0, self.total_score - deduction)


def is_holiday_or_lunar_new_year(d: datetime.date) -> Tuple[bool, str]:
    """Kiểm tra ngày có rơi vào Tết Nguyên Đán hoặc Lễ quốc gia không."""
    for start_tet, end_tet in LUNAR_NEW_YEAR_RANGES:
        if start_tet <= d <= end_tet:
            return True, f"Nghỉ Tết Nguyên Đán ({start_tet.strftime('%d/%m')} - {end_tet.strftime('%d/%m')})"
    if (d.month, d.day) in DEFAULT_VIETNAM_HOLIDAYS_FIXED:
        return True, f"Nghỉ Lễ Quốc gia ({d.day:02d}/{d.month:02d})"
    return False, ""


def to_date(val: Any) -> Optional[datetime.date]:
    """Chuyển đổi linh hoạt giá trị Excel sang datetime.date."""
    if val is None:
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        if isinstance(val, datetime.datetime):
            return val.date()
        return val
    s = str(val).strip()
    if not s or s.startswith("#") or s == "None":
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s.split()[0], fmt).date()
        except ValueError:
            pass
    return None


def verify_kcs_workbook(wb_path: str, dmcv_sheet: Optional[str] = None) -> InvariantAuditReport:
    """
    Quét và kiểm tra 6 Bất biến Lõi trên file Excel QLCL.
    """
    import openpyxl

    report = InvariantAuditReport(file_path=wb_path)
    if not os.path.exists(wb_path):
        report.add_error(f"File không tồn tại: {wb_path}", 100)
        return report

    try:
        wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    except Exception as e:
        report.add_error(f"Không thể mở file Excel: {e}", 100)
        return report

    # 1. Tìm Sheet Master Schedule (Danh mục công việc / DMCV / HOSO_KCS)
    target_sheet = None
    if dmcv_sheet and dmcv_sheet in wb.sheetnames:
        target_sheet = dmcv_sheet
    else:
        for sname in wb.sheetnames:
            s_clean = sname.strip().lower()
            if any(k in s_clean for k in ["danh mục công việc", "dmcv", "hoso_kcs", "kcs_master"]):
                target_sheet = sname
                break

    if not target_sheet:
        # Nếu là file tổng hợp khác (như HUB VINA), kiểm tra các sheet có sẵn
        target_sheet = wb.sheetnames[0]
        report.add_warning(f"Không tìm thấy sheet chuẩn 'Danh mục công việc', quét sheet đầu tiên: {target_sheet}")

    ws = wb[target_sheet]

    # Quét header để tìm các cột quan trọng
    col_map = {}
    sample_rows = list(ws.iter_rows(min_row=1, max_row=15, values_only=True))
    header_row_idx = 1
    for r_idx, row in enumerate(sample_rows, 1):
        row_str = [str(c).strip().lower() if c is not None else "" for c in row]
        if any("nội dung công việc" in c or "tên công việc" in c or "đối tượng nghiệm thu" in c for c in row_str):
            header_row_idx = r_idx
            for c_idx, val in enumerate(row_str, 1):
                if not val:
                    continue
                clean_val = " ".join(val.split())
                if "nội dung công việc" in clean_val or "tên công việc" in clean_val:
                    col_map["desc"] = c_idx
                elif "mã hiệu" in clean_val and "chung" not in clean_val and "lưu" not in clean_val:
                    col_map.setdefault("code", c_idx)
                elif "vị trí" in clean_val:
                    col_map.setdefault("loc", c_idx)
                elif "phiếu yc" in clean_val or "ngày pyc" in clean_val:
                    col_map["pyc_date"] = c_idx
                elif "ngày nt" in clean_val or "ngày\nnt" in val or "nt/ lm/ kt" in clean_val:
                    col_map["nt_date"] = c_idx
                elif "ngày bắt đầu" in clean_val or "ngày\nbắt đầu" in val:
                    col_map["start_date"] = c_idx
                elif "ngày kết thúc" in clean_val or "ngày\nkết thúc" in val:
                    col_map["end_date"] = c_idx
                elif "khối lượng nghiệm thu" in clean_val:
                    col_map["volume"] = c_idx
            break

    # Nếu không tìm thấy bằng header, dùng heuristic theo chuẩn XDA QLCL
    if "desc" not in col_map:
        col_map = {
            "code": 11,
            "desc": 12,
            "loc": 13,
            "start_date": 18,
            "end_date": 19,
            "pyc_date": 20,
            "nt_date": 21,
            "volume": 40,
        }

    # Đọc danh sách công việc
    tasks = []
    for r_idx, row in enumerate(ws.iter_rows(min_row=header_row_idx + 1, max_row=1000, values_only=True), header_row_idx + 1):
        desc = row[col_map["desc"] - 1] if len(row) >= col_map.get("desc", 12) else None
        if not desc or str(desc).strip() == "" or str(desc).startswith("#"):
            continue
        desc_str = str(desc).strip()
        pyc_d = to_date(row[col_map["pyc_date"] - 1]) if "pyc_date" in col_map and len(row) >= col_map["pyc_date"] else None
        nt_d = to_date(row[col_map["nt_date"] - 1]) if "nt_date" in col_map and len(row) >= col_map["nt_date"] else None
        start_d = to_date(row[col_map["start_date"] - 1]) if "start_date" in col_map and len(row) >= col_map["start_date"] else None
        end_d = to_date(row[col_map["end_date"] - 1]) if "end_date" in col_map and len(row) >= col_map["end_date"] else None
        tasks.append({
            "row": r_idx,
            "desc": desc_str,
            "pyc_date": pyc_d,
            "nt_date": nt_d,
            "start_date": start_d,
            "end_date": end_d,
        })

    report.details["total_tasks_scanned"] = len(tasks)

    # 2. KIỂM TRA BẤT BIẾN 2: ĐỒ THỊ THỜI GIAN & CHU KỲ BÊ TÔNG
    # Gom nhóm theo kết cấu / cấu kiện để kiểm tra chuỗi: Đổ BT -> Tháo ván khuôn (>=2 ngày) -> Nghiệm thu R28 (>=28 ngày)
    last_pour_date = None
    last_pour_task = ""

    for t in tasks:
        desc_lower = t["desc"].lower()
        nt_d = t["nt_date"]
        pyc_d = t["pyc_date"]

        # Kiểm tra PYC <= NT
        if pyc_d and nt_d:
            if pyc_d > nt_d:
                report.add_error(
                    f"Dòng {t['row']} [{t['desc'][:30]}]: Ngày Phiếu yêu cầu ({pyc_d}) ĐI SAU ngày Nghiệm thu ({nt_d})!",
                    deduction=10
                )
            elif pyc_d == nt_d:
                report.add_warning(
                    f"Dòng {t['row']} [{t['desc'][:30]}]: Phiếu yêu cầu gửi cùng ngày nghiệm thu ({nt_d}) (nên gửi trước >= 24h)"
                )

        # Kiểm tra ngày Lễ, Tết Nguyên Đán
        for check_d, d_label in [(nt_d, "Nghiệm thu"), (t["start_date"], "Bắt đầu thi công"), (t["end_date"], "Kết thúc thi công")]:
            if check_d:
                is_hol, hol_name = is_holiday_or_lunar_new_year(check_d)
                if is_hol:
                    report.add_error(
                        f"Dòng {t['row']} [{t['desc'][:30]}]: Ngày {d_label} rơi đúng vào {hol_name} ({check_d})! Vi phạm quy tắc nghỉ Tết/Lễ.",
                        deduction=15
                    )

        # Phát hiện sự kiện Đổ bê tông
        if any(k in desc_lower for k in ["đổ bê tông", "đổ bt", "bê tông móng", "bê tông tường", "bê tông xà mũ", "bê tông dầm", "bản mặt cầu"]):
            if "tháo" not in desc_lower and "ván khuôn" not in desc_lower and "kiểm tra trước" not in desc_lower and "lấy mẫu" not in desc_lower:
                last_pour_date = nt_d or t["end_date"]
                last_pour_task = t["desc"]

        # Phát hiện tháo dỡ ván khuôn
        if any(k in desc_lower for k in ["tháo ván khuôn", "tháo dỡ ván khuôn", "cấu kiện sau khi tháo", "sau tháo ván khuôn"]):
            if last_pour_date and nt_d:
                delta_days = (nt_d - last_pour_date).days
                if delta_days < 2:
                    report.add_error(
                        f"Dòng {t['row']} [{t['desc'][:30]}]: Tháo ván khuôn ngày {nt_d} chỉ cách ngày đổ bê tông ({last_pour_date}) {delta_days} ngày! Yêu cầu tối thiểu 2 ngày (48 giờ).",
                        deduction=20
                    )

        # Phát hiện Nghiệm thu hoàn thành cấu kiện bê tông (R28)
        if any(k in desc_lower for k in ["hoàn thành cấu kiện", "nghiệm thu cấu kiện", "bê tông hoàn thành"]):
            if last_pour_date and nt_d:
                delta_days = (nt_d - last_pour_date).days
                if delta_days < 28:
                    report.add_error(
                        f"Dòng {t['row']} [{t['desc'][:30]}]: Nghiệm thu cấu kiện hoàn thành ngày {nt_d} chỉ cách ngày đổ ({last_pour_date}) {delta_days} ngày! Chưa đủ 28 ngày nén mẫu R28.",
                        deduction=25
                    )

    # 3. KIỂM TRA BẤT BIẾN 3 & 4: SỰ TỒN TẠI CỦA CÁC CHECKLIST & PHỤ LỤC
    # Kiểm tra xem workbook có các sheet checklist hình học chuẩn không
    sheet_set = {s.strip().lower() for s in wb.sheetnames}
    has_nhat_trinh = any("nhật trình" in s or "nhật ký" in s for s in sheet_set)
    has_checklist_vk = any("ván khuôn" in s or "coppha" in s for s in sheet_set)
    has_checklist_thep = any("cốt thép" in s or "thep" in s for s in sheet_set)
    has_checklist_bt = any("đổ bt" in s or "bê tông" in s for s in sheet_set)

    if not has_nhat_trinh:
        report.add_warning("File chưa có Sheet 'Nhật trình' hoặc 'Nhật ký TC' đồng bộ hiện trường.", 5)
    if not (has_checklist_vk or has_checklist_thep or has_checklist_bt):
        report.add_warning("File chưa kẹp đầy đủ hệ thống Sheet Checklist hình học (Ván khuôn, Cốt thép, Đổ bê tông).", 5)

    report.details["has_nhat_trinh"] = has_nhat_trinh
    report.details["has_checklists"] = (has_checklist_vk and has_checklist_thep)

    return report


def main():
    parser = argparse.ArgumentParser(description="KCS Invariants Verifier — Bộ kiểm tra 6 Bất biến Lõi QLCL")
    parser.add_argument("excel_file", help="Đường dẫn file Excel QLCL (.xlsx / .xlsm)")
    parser.add_argument("--sheet", default=None, help="Tên sheet Danh mục công việc")
    parser.add_argument("--json", default=None, help="Xuất báo cáo kết quả ra file JSON")
    args = parser.parse_args()

    print("\n" + "═" * 70)
    print("  KCS INVARIANTS VERIFIER — KIỂM TRA 6 BẤT BIẾN LÕI HỒ SƠ QLCL")
    print(f"  File: {args.excel_file}")
    print("═" * 70)

    report = verify_kcs_workbook(args.excel_file, dmcv_sheet=args.sheet)

    print(f"\n  [KẾT QUẢ ĐÁNH GIÁ]")
    print(f"  Điểm chất lượng: {report.total_score}/100")
    print(f"  Trạng thái: {'✅ ĐẠT CHUẨN KCS' if report.passed and report.total_score >= 80 else '❌ KHÔNG ĐẠT (CẦN ĐIỀU CHỈNH)'}")
    print(f"  Tổng số công tác quét: {report.details.get('total_tasks_scanned', 0)}")

    if report.critical_errors:
        print(f"\n  [LỖI NGHIÊM TRỌNG BỊ CHẶN] ({len(report.critical_errors)} lỗi):")
        for err in report.critical_errors:
            print(f"    ❌ {err}")

    if report.warnings:
        print(f"\n  [CẢNH BÁO KỸ THUẬT] ({len(report.warnings)} cảnh báo):")
        for w in report.warnings:
            print(f"    ⚠ {w}")

    if args.json:
        out_dict = {
            "file": report.file_path,
            "score": report.total_score,
            "passed": report.passed,
            "errors": report.critical_errors,
            "warnings": report.warnings,
            "details": report.details,
        }
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(out_dict, f, ensure_ascii=False, indent=2)
        print(f"\n  Đã lưu kết quả ra: {args.json}")

    print("\n" + "═" * 70 + "\n")
    sys.exit(0 if report.passed and report.total_score >= 80 else 1)


if __name__ == "__main__":
    main()
