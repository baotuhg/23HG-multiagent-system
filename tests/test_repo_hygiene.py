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
# File lớn được CHỦ Ý giữ trong repo (sản phẩm giao xưởng cắt thép theo từng loại thép của Gói B Km19).
# Chỉ thêm vào đây khi chủ dự án xác nhận cần giữ, kèm lý do; mọi file lớn khác vẫn bị chặn.
GOI_B = ("examples/HO_SO_CAU_KM19_529/HUB/03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH/"
         "GOI_B_XUONG_TIEN_CHE_COT_THEP/")
# Cả thư mục (tiền tố đường dẫn) được phép chứa file lớn, kèm lý do.
ALLOWED_LARGE_PREFIXES = {
    GOI_B + "01_HE_THONG_CAT_THEP_REBARCUT/":
        "bộ cắt thép giao xưởng theo từng Ø: một RebarCut + một lệnh cắt CNC cho mỗi Ø, Master và bảng tổng hợp "
        "(sinh bằng examples/generate_rebarcut_dedicated_package.py --out)",
}
ALLOWED_LARGE = {
    GOI_B + "01_Phieu_Cat_Thep_Cau_Km19+529.080.csv": "lệnh cắt CNC từng đoạn cắt (138.045 dòng) cho xưởng",
    GOI_B + "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx": "bảng tổ hợp cắt thép RebarCut theo từng loại thép",
}
# File hợp lệ khác nhau theo từng hub/gói (đường dẫn bên trong khác nhau) nên không bắt buộc giống nhau.
PER_PACKAGE_FILES = {"DISPATCH_MANIFEST.json"}
# Cặp (dự án, tên file) được phép khác nhau giữa các bản, kèm lý do. Chỉ thêm khi chủ dự án xác nhận.
EXPECTED_DIVERGENT = {
    ("HO_SO_CAU_KM19_529", "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx"):
        "Gói B giữ bản đầy đủ do bộ giải xuất (≈4,2 MB, 138.048 dòng chi tiết); bộ vi mô 14 file giữ bản nhẹ (8 KB)",
}


def tracked_files():
    try:
        out = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, capture_output=True, check=True).stdout
    except (OSError, subprocess.CalledProcessError):
        return None
    return [p for p in out.decode("utf-8").split("\0") if p]


def oversized(sizes, limit=MAX_TRACKED_BYTES, allowed=ALLOWED_LARGE, allowed_prefixes=ALLOWED_LARGE_PREFIXES):
    """sizes: {đường dẫn: số byte} → danh sách (đường dẫn, byte) vượt ngưỡng và chưa được cho phép, lớn nhất trước."""
    return sorted(((p, n) for p, n in sizes.items()
                   if n > limit and p not in allowed and not p.startswith(tuple(allowed_prefixes))),
                  key=lambda x: -x[1])


def diverged_copies(files, md5_of, per_package=PER_PACKAGE_FILES, expected=EXPECTED_DIVERGENT):
    """
    Nhóm theo (dự án, tên file) trong examples/ — dự án là thư mục cấp 1 dưới examples/ — và trả các nhóm có
    >1 bản nhưng nội dung khác nhau. Hai dự án khác nhau được phép có file trùng tên khác nội dung.
    """
    groups = collections.defaultdict(list)
    for f in files:
        parts = f.replace("\\", "/").split("/")
        if parts[0] == "examples" and len(parts) > 2 and parts[-1] not in per_package:
            groups[(parts[1], parts[-1])].append(f)
    return {k: v for k, v in groups.items()
            if k not in expected and len(v) > 1 and len({md5_of(p) for p in v}) > 1}


def _md5(rel_path):
    with open(os.path.join(ROOT, rel_path), "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


class HelperLogicTest(unittest.TestCase):
    """Kiểm tra chính các hàm bắt lỗi — để biết chúng có bắt được lỗi thật."""

    def test_oversized_detects_and_sorts(self):
        sizes = {"a.py": 10, "big.csv": 9_000_000, "mid.xlsx": 4_300_000, "edge.bin": MAX_TRACKED_BYTES}
        self.assertEqual([p for p, _ in oversized(sizes)], ["big.csv", "mid.xlsx"])   # đúng ngưỡng thì chưa vượt

    def test_allowlisted_large_file_is_not_flagged_but_others_are(self):
        sizes = {GOI_B + "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx": 4_300_000, "other.xlsx": 4_300_000}
        self.assertEqual([p for p, _ in oversized(sizes)], ["other.xlsx"])

    def test_allowlisted_prefix_exempts_the_folder_only(self):
        sizes = {GOI_B + "01_HE_THONG_CAT_THEP_REBARCUT/THEO_TUNG_DUONG_KINH_PHI/x.xlsx": 5_000_000,
                 GOI_B + "khac/x.xlsx": 5_000_000}
        self.assertEqual([p for p, _ in oversized(sizes)], [GOI_B + "khac/x.xlsx"])

    def test_allowlist_has_no_stale_entries(self):
        files = tracked_files()
        if files is None:
            self.skipTest("không phải bản checkout git")
        self.assertEqual([p for p in ALLOWED_LARGE if p not in files], [], "allowlist trỏ tới file không còn tồn tại")
        self.assertEqual([pre for pre in ALLOWED_LARGE_PREFIXES if not any(f.startswith(pre) for f in files)], [],
                         "tiền tố trong allowlist không còn file nào")

    def test_diverged_copies_found_within_project(self):
        files = ["examples/P1/a/x.xlsx", "examples/P1/b/x.xlsx", "examples/P1/c/y.xlsx"]
        content = {"examples/P1/a/x.xlsx": "v1", "examples/P1/b/x.xlsx": "v2", "examples/P1/c/y.xlsx": "v1"}
        found = diverged_copies(files, content.get)
        self.assertEqual(list(found), [("P1", "x.xlsx")])

    def test_same_name_in_different_projects_is_allowed(self):
        files = ["examples/P1/x.xlsx", "examples/P2/x.xlsx"]
        self.assertEqual(diverged_copies(files, {"examples/P1/x.xlsx": "v1", "examples/P2/x.xlsx": "v2"}.get), {})

    def test_expected_divergence_is_exempt_only_for_that_file(self):
        files = ["examples/P1/a/x.xlsx", "examples/P1/b/x.xlsx", "examples/P1/a/y.xlsx", "examples/P1/b/y.xlsx"]
        content = {"examples/P1/a/x.xlsx": "v1", "examples/P1/b/x.xlsx": "v2",
                   "examples/P1/a/y.xlsx": "v1", "examples/P1/b/y.xlsx": "v2"}
        found = diverged_copies(files, content.get, expected={("P1", "x.xlsx"): "lý do"})
        self.assertEqual(list(found), [("P1", "y.xlsx")])

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
