#!/usr/bin/env python3
"""Audit QA/QC Master Register & Construction Dossier Timelines.

Validates:
1. Document ID uniqueness and cross-references (Material -> Test -> Work Acceptance -> Stage Acceptance).
2. Date formats and chronological integrity (Material intake <= Test <= Request <= Inspection/Acceptance).
3. Concrete curing time logic (R7 ~ +7d, R28 ~ +28d).
4. Signature and participant completeness based on roles.
5. Closed-loop status for Non-Conformance Reports (NCR).
"""
import argparse
import csv
import os
import sys
from datetime import datetime, timedelta
from collections import Counter
from typing import Dict, List, Optional, Tuple

DATE_FMT = "%Y-%m-%d"


def norm(v: Optional[str]) -> str:
    return (v or "").strip()


def parse_date(v: Optional[str]):
    v = norm(v)
    if not v:
        return None
    try:
        return datetime.strptime(v, DATE_FMT).date()
    except ValueError:
        # Try DD/MM/YYYY fallback
        try:
            return datetime.strptime(v, "%d/%m/%Y").date()
        except ValueError:
            return "INVALID"


def main():
    ap = argparse.ArgumentParser(description="Audit QA/QC Dossier Master Register")
    ap.add_argument("file_path", help="Path to register CSV or Excel (.xlsx/.xlsm) file")
    ap.add_argument("--strict", action="store_true", help="Fail on warnings as well as errors")
    args = ap.parse_args()

    # Nếu là file Excel, chạy bộ kiểm tra 6 Bất biến Lõi KCS Invariants Verifier
    if args.file_path.endswith((".xlsx", ".xlsm")):
        import sys
        sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../")))
        from tools.kcs_invariants_verifier import verify_kcs_workbook
        report = verify_kcs_workbook(args.file_path)
        print("\n" + "═" * 70)
        print("  AEC-QLCL: AUDIT MASTER REGISTER THEO 6 BẤT BIẾN LÕI")
        print(f"  File: {args.file_path}")
        print(f"  Điểm: {report.total_score}/100 | Trạng thái: {'ĐẠT' if report.passed else 'KHÔNG ĐẠT'}")
        print("═" * 70)
        for e in report.critical_errors:
            print(f"  ❌ {e}")
        for w in report.warnings:
            print(f"  ⚠ {w}")
        sys.exit(0 if report.passed and (not args.strict or not report.warnings) else 1)

    with open(args.file_path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    errors: List[str] = []
    warnings: List[str] = []

    # 1. ID Uniqueness
    ids = [norm(r.get("doc_id")) for r in rows if norm(r.get("doc_id"))]
    dup = [x for x, n in Counter(ids).items() if n > 1]
    for x in dup:
        errors.append(f"CRITICAL: Trùng lặp Document ID '{x}'")

    by_id: Dict[str, dict] = {norm(r.get("doc_id")): r for r in rows if norm(r.get("doc_id"))}

    # 2. Row by row audit
    for i, r in enumerate(rows, start=2):
        did = norm(r.get("doc_id")) or f"ROW-{i}"
        dtype = norm(r.get("doc_type")).upper()
        status = norm(r.get("status")).upper()
        date = parse_date(r.get("date"))
        pour_date = parse_date(r.get("pour_date"))

        # Date validation
        if date == "INVALID":
            errors.append(f"{did}: Sai định dạng ngày (yêu cầu YYYY-MM-DD hoặc DD/MM/YYYY): '{r.get('date')}'")

        # Signature validation
        if norm(r.get("signer_required")).upper() in {"YES", "Y", "1", "TRUE"}:
            if norm(r.get("signer_present")).upper() not in {"YES", "Y", "1", "TRUE"}:
                errors.append(f"{did}: Thiếu chữ ký/thành phần nghiệm thu bắt buộc")

        # Acceptance item checks
        if dtype in {"ACCEPTANCE", "WORK_ACCEPTANCE", "NGHIEM_THU_CONG_VIEC"}:
            if not norm(r.get("drawing_ref")):
                warnings.append(f"{did}: Biên bản nghiệm thu chưa ghi số hiệu bản vẽ thiết kế (drawing_ref)")
            if not norm(r.get("location")):
                warnings.append(f"{did}: Biên bản nghiệm thu chưa ghi vị trí/lý trình (location/chainage)")
            if not norm(r.get("material_ref")) and norm(r.get("has_material", "YES")).upper() in {"YES", "1", "TRUE"}:
                warnings.append(f"{did}: Công việc có vật liệu nhưng chưa liên kết biên bản vật liệu đầu vào (material_ref)")
            if not norm(r.get("test_ref")) and norm(r.get("has_test", "NO")).upper() in {"YES", "1", "TRUE"}:
                warnings.append(f"{did}: Công việc yêu cầu thí nghiệm nhưng chưa liên kết phiếu kết quả thí nghiệm (test_ref)")

        # Concrete age check (R7, R28)
        if dtype in {"CONCRETE_TEST", "THI_NGHIEM_BE_TONG"}:
            age_type = norm(r.get("age_type")).upper()
            if pour_date and date and date != "INVALID" and pour_date != "INVALID":
                delta_days = (date - pour_date).days
                if "R7" in age_type or "7" in age_type:
                    if not (6 <= delta_days <= 8):
                        warnings.append(f"{did}: Thí nghiệm R7 ngày nén cách ngày đổ {delta_days} ngày (chuẩn: 7 ngày)")
                elif "R28" in age_type or "28" in age_type:
                    if not (27 <= delta_days <= 30):
                        warnings.append(f"{did}: Thí nghiệm R28 ngày nén cách ngày đổ {delta_days} ngày (chuẩn: 28 ngày)")

        # NCR check
        if dtype in {"NCR", "SU_CO"} and status not in {"CLOSED", "CLOSE", "HOAN_THANH", "VOID"}:
            warnings.append(f"{did}: Phiếu không phù hợp/Sự cố chưa được đóng trạng thái (status={status or 'Trống'})")

        # Cross-reference timeline check
        for field in ("material_ref", "test_ref", "request_ref", "acceptance_ref"):
            ref = norm(r.get(field))
            if ref:
                for token in [x.strip() for x in ref.split(";") if x.strip()]:
                    if token not in by_id:
                        warnings.append(f"{did}: {field} trỏ tới mã không tồn tại trong danh mục: '{token}'")
                    elif date not in (None, "INVALID"):
                        ref_date = parse_date(by_id[token].get("date"))
                        if ref_date not in (None, "INVALID"):
                            # Material & Test must precede or coincide with acceptance
                            if field in {"material_ref", "test_ref", "request_ref"} and ref_date > date:
                                errors.append(f"{did}: LỖI LOGIC NGÀY THÁNG - {field} [{token}] ngày {ref_date} sau ngày nghiệm thu {date}")

    print("==================================================")
    print("   BÁO CÁO AUDIT HỒ SƠ QUẢN LÝ CHẤT LƯỢNG (QLCL)   ")
    print("==================================================")
    print(f"Tổng số bản ghi: {len(rows)}")
    print(f"Tổng số lỗi (CRITICAL/ERROR): {len(errors)}")
    print(f"Tổng số cảnh báo (WARNING):   {len(warnings)}")
    print("--------------------------------------------------")

    if errors:
        print("\n[!] DANH SÁCH LỖI NGHIÊM TRỌNG:")
        for idx, e in enumerate(errors, 1):
            print(f"  {idx}. {e}")

    if warnings:
        print("\n[*] DANH SÁCH CẢNH BÁO CẦN LƯU Ý:")
        for idx, w in enumerate(warnings, 1):
            print(f"  {idx}. {w}")

    if not errors and not warnings:
        print("\n[V] XÁC NHẬN: Toàn bộ danh mục hồ sơ đảm bảo tính toàn vẹn và logic pháp lý!")

    print("==================================================")
    raise SystemExit(1 if (errors or (args.strict and warnings)) else 0)


if __name__ == "__main__":
    main()
