# -*- coding: utf-8 -*-
"""
Test vệ sinh repo — chặn hai lỗi đã từng xảy ra:

  1. Commit nhầm sản phẩm sinh ra rất lớn (một lần đưa vào ≈13 MB: CSV 138.045 dòng + xlsx 4,2 MB), làm
     repo phình gấp 3 và làm bộ test duyệt Excel chậm thêm vài chục giây.
  2. Các bản sao CÙNG TÊN của một file (mỗi gói hub-and-spoke giữ một bản tự đủ) bị lệch nhau vì chỉ
     cập nhật một bản.

Khi test (1) đỏ: đừng nâng ngưỡng. Thêm file vào .gitignore và ghi cách sinh lại trong README.
Khi test (2) đỏ: đồng bộ các bản sao (xem sync_same_named_copies trong tools/apply_a5_schedule_0509_1211.py).
"""

import collections
import hashlib
import os
import subprocess
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAX_TRACKED_BYTES = 1024 * 1024            # 1 MB
# File hợp lệ khác nhau theo từng hub/gói (đường dẫn bên trong khác nhau) nên không bắt buộc giống nhau.
PER_PACKAGE_FILES = {"DISPATCH_MANIFEST.json"}


def tracked_files():
    try:
        out = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, capture_output=True, check=True).stdout
    except (OSError, subprocess.CalledProcessError):
        return None
    return [p for p in out.decode("utf-8").split("\0") if p]


def oversized(sizes, limit=MAX_TRACKED_BYTES):
    """sizes: {đường dẫn: số byte} → danh sách (đường dẫn, byte) vượt ngưỡng, lớn nhất trước."""
    return sorted(((p, n) for p, n in sizes.items() if n > limit), key=lambda x: -x[1])


def diverged_copies(files, md5_of, per_package=PER_PACKAGE_FILES):
    """
    Nhóm theo (dự án, tên file) trong examples/ — dự án là thư mục cấp 1 dưới examples/ — và trả các nhóm có
    >1 bản nhưng nội dung khác nhau. Hai dự án khác nhau được phép có file trùng tên khác nội dung.
    """
    groups = collections.defaultdict(list)
    for f in files:
        parts = f.replace("\\", "/").split("/")
        if parts[0] == "examples" and len(parts) > 2 and parts[-1] not in per_package:
            groups[(parts[1], parts[-1])].append(f)
    return {k: v for k, v in groups.items() if len(v) > 1 and len({md5_of(p) for p in v}) > 1}


def _md5(rel_path):
    with open(os.path.join(ROOT, rel_path), "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


class HelperLogicTest(unittest.TestCase):
    """Kiểm tra chính các hàm bắt lỗi — để biết chúng có bắt được lỗi thật."""

    def test_oversized_detects_and_sorts(self):
        sizes = {"a.py": 10, "big.csv": 9_000_000, "mid.xlsx": 4_300_000, "edge.bin": MAX_TRACKED_BYTES}
        self.assertEqual([p for p, _ in oversized(sizes)], ["big.csv", "mid.xlsx"])   # đúng ngưỡng thì chưa vượt

    def test_diverged_copies_found_within_project(self):
        files = ["examples/P1/a/x.xlsx", "examples/P1/b/x.xlsx", "examples/P1/c/y.xlsx"]
        content = {"examples/P1/a/x.xlsx": "v1", "examples/P1/b/x.xlsx": "v2", "examples/P1/c/y.xlsx": "v1"}
        found = diverged_copies(files, content.get)
        self.assertEqual(list(found), [("P1", "x.xlsx")])

    def test_same_name_in_different_projects_is_allowed(self):
        files = ["examples/P1/x.xlsx", "examples/P2/x.xlsx"]
        self.assertEqual(diverged_copies(files, {"examples/P1/x.xlsx": "v1", "examples/P2/x.xlsx": "v2"}.get), {})

    def test_identical_copies_and_per_package_files_pass(self):
        files = ["examples/P1/a/x.xlsx", "examples/P1/b/x.xlsx",
                 "examples/P1/a/DISPATCH_MANIFEST.json", "examples/P1/b/DISPATCH_MANIFEST.json"]
        content = {files[0]: "v1", files[1]: "v1", files[2]: "m1", files[3]: "m2"}
        self.assertEqual(diverged_copies(files, content.get), {})


class RepoHygieneTest(unittest.TestCase):
    def setUp(self):
        self.files = tracked_files()
        if self.files is None:
            self.skipTest("không phải bản checkout git")

    def test_no_oversized_tracked_files(self):
        sizes = {p: os.path.getsize(os.path.join(ROOT, p)) for p in self.files
                 if os.path.isfile(os.path.join(ROOT, p))}
        big = oversized(sizes)
        self.assertEqual(big, [], "file theo dõi vượt 1 MB — thêm vào .gitignore, đừng commit: "
                                  + ", ".join(f"{p} ({n // 1024} KB)" for p, n in big))

    def test_same_named_copies_in_a_project_are_identical(self):
        existing = [p for p in self.files if os.path.isfile(os.path.join(ROOT, p))]
        bad = diverged_copies(existing, _md5)
        self.assertEqual(bad, {}, "bản sao cùng tên bị lệch nhau: "
                                  + "; ".join(f"{proj}/{name} ({len(v)} bản)" for (proj, name), v in bad.items()))


if __name__ == "__main__":
    unittest.main()
