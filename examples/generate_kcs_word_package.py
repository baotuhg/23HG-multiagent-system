# -*- coding: utf-8 -*-
"""
Sinh sổ nghiệm thu KCS (Word) cho Cầu Km19+529.080 từ sổ nghiệm thu CSV.

Logic nằm trong tools/kcs_register.py (trạng thái được TÍNH từ ngày và tham chiếu, không gõ tay).
Dữ liệu nằm trong examples/data/:
  - kcs_register_km19.csv      danh mục bản ghi (mã, công việc, ngày, điều kiện, tham chiếu)
  - kcs_nguoi_ky_km19.json     người ký biên bản

Chạy:  python examples/generate_kcs_word_package.py [--register ...] [--signers ...] [--out ...]
"""

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tools.kcs_register import build_docx, check_register, load_register, load_signers  # noqa: E402

DATA = os.path.join(ROOT, "examples", "data")
DEFAULT_REGISTER = os.path.join(DATA, "kcs_register_km19.csv")
DEFAULT_SIGNERS = os.path.join(DATA, "kcs_nguoi_ky_km19.json")
DEFAULT_OUT = os.path.join(ROOT, "templates", "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx")
PROJECT = "CẦU KM19+529.080 - DỰ ÁN CAO TỐC TUYÊN QUANG - HÀ GIANG"
LOCATION = "Cầu Km19+529.080 (Phân đoạn Km19+120 -:- Km27+480) - Cao tốc Tuyên Quang - Hà Giang"


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sinh sổ nghiệm thu KCS Word từ sổ CSV")
    ap.add_argument("--register", default=DEFAULT_REGISTER)
    ap.add_argument("--signers", default=DEFAULT_SIGNERS)
    ap.add_argument("--out", default=DEFAULT_OUT)
    a = ap.parse_args(argv)

    records = load_register(a.register)
    checks = check_register(records)
    counts = {}
    for c in checks.values():
        counts[c.status] = counts.get(c.status, 0) + 1
    path = build_docx(records, checks, a.out, project_name=PROJECT, location=LOCATION,
                      signers=load_signers(a.signers))
    print(f"-> Đã xuất sổ KCS: {path}")
    print(f"   {len(records)} bản ghi: " + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    return 0


if __name__ == "__main__":
    sys.exit(main())
