# -*- coding: utf-8 -*-
"""
Đường dẫn dùng chung cho các script ví dụ — KHÔNG ghi đường dẫn máy cá nhân vào code.

- repo_path("templates/x.xlsx")      → <thư mục repo>/templates/x.xlsx (tự tính, chạy được trên mọi máy)
- project_path("HSTK Cầu .../a.xlsx") → <AEC_PROJECTS_DIR>/HSTK Cầu .../a.xlsx
  AEC_PROJECTS_DIR là thư mục chứa hồ sơ dự án gốc (bản vẽ, file nguồn) trên máy bạn; mặc định
  <repo>/du_lieu_du_an (không đưa vào git). Ví dụ:  set AEC_PROJECTS_DIR=C:\\HoSo   (Windows)
                                                    export AEC_PROJECTS_DIR=~/HoSo  (Linux/macOS)
"""
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJECTS_DIR = os.path.expanduser(os.environ.get("AEC_PROJECTS_DIR", os.path.join(REPO_ROOT, "du_lieu_du_an")))


def _join(base: str, rel: str) -> str:
    parts = [p for p in rel.replace("\\", "/").split("/") if p]
    return os.path.join(base, *parts)


def repo_path(rel: str = "") -> str:
    return _join(REPO_ROOT, rel)


def project_path(rel: str = "") -> str:
    return _join(PROJECTS_DIR, rel)
