# -*- coding: utf-8 -*-
"""
Test bảng trạng thái pháp lý (workflows/LEGAL_STATUS.md).

Bảo đảm: mọi văn bản được trích dẫn trong quy tắc đo bóc và cảnh báo đều có dòng trong bảng; quy tắc
TT 13/2021 vẫn ở trạng thái chưa xác minh; và không dòng nào được ghi là đã đối chiếu bản gốc khi chưa có
nguồn chính thức được lưu kèm.
"""

import os
import re
import unittest

from tools.takeoff_rules import PROFILE_TT13_2021_PL_VI

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATUS = os.path.join(ROOT, "workflows", "LEGAL_STATUS.md")
RULES = os.path.join(ROOT, "tools", "takeoff_rules.py")


def _status_text():
    with open(STATUS, encoding="utf-8") as f:
        return f.read()


class LegalStatusTest(unittest.TestCase):
    def test_key_documents_listed(self):
        text = _status_text()
        for doc in ("TT 13/2021", "TT 37/2026", "TT 38/2026", "TT 36/2026", "NĐ 207/2026", "NĐ 254/2025",
                    "QĐ 1041/QĐ-BXD"):
            self.assertIn(doc, text, doc)

    def test_documents_cited_in_takeoff_rules_are_listed(self):
        with open(RULES, encoding="utf-8") as f:
            src = f.read()
        cited = set(re.findall(r"(TT \d+/\d{4}/TT-BXD|TT \d+/\d{4}|QĐ \d+/QĐ-BXD|QĐ \d+/QĐ-BXD)", src))
        text = _status_text()
        missing = sorted(c for c in cited if c not in text and c.replace("/TT-BXD", "") not in text)
        self.assertEqual(missing, [], "văn bản được trích trong takeoff_rules nhưng thiếu trong bảng trạng thái")

    def test_tt13_profile_is_not_verified(self):
        self.assertFalse(PROFILE_TT13_2021_PL_VI.verified)
        warnings = " ".join(PROFILE_TT13_2021_PL_VI.warnings())
        self.assertIn("CHƯA đối chiếu", warnings)

    def test_no_row_claims_verified_without_source(self):
        text = _status_text()
        self.assertNotIn("ĐÃ ĐỐI CHIẾU BẢN GỐC", text.split("## Trạng thái")[1].split("## Hệ quả")[0].replace(
            "*ĐÃ ĐỐI CHIẾU BẢN GỐC*", ""))

    def test_table_rows_have_five_or_six_cells(self):
        rows = [l for l in _status_text().splitlines() if l.startswith("| **") or l.startswith("| TT") ]
        self.assertGreaterEqual(len(rows), 10)
        for r in rows:
            self.assertEqual(r.count("|"), 6, r[:60])


if __name__ == "__main__":
    unittest.main()
