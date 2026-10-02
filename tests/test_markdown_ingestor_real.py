# -*- coding: utf-8 -*-
"""Bộ nạp hồ sơ Markdown trên hồ sơ KTTC THẬT (nhà ký túc xá 3 tầng, bản OCR).

tests/fixtures/ho_so_ktx_trich_ocr.md là bản trích kỹ thuật (không có tên người/đơn vị) cho ra cùng bảng cấu kiện
như toàn bộ hồ sơ. Muốn chạy trên toàn bộ hồ sơ: đặt AEC_TEST_DOSSIER=<đường dẫn noi_dung.md>.
"""

import contextlib
import io
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core.agents.aec_markdown_ingestor import AECMarkdownIngestor

FIXTURE = os.path.join(ROOT, "tests", "fixtures", "ho_so_ktx_trich_ocr.md")


def ingest(path):
    with contextlib.redirect_stdout(io.StringIO()):
        return AECMarkdownIngestor().process_file(path)


class RealDossierTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.r = ingest(FIXTURE)
        cls.checks = cls.r["schedule_checks"]
        cls.specs = cls.r["technical_specs"]

    def test_footings_equal_columns(self):
        # MT1 18 + MT1A 2 + MT2 15 + MT5 4 + MT4 32 + MT3 4 + MT3A 4 = 79 ; C-01 19 + C-02 8 + C-03 32 + C-04 20 = 79
        self.assertEqual(self.checks["totals"]["MT"], 79)
        self.assertEqual(self.checks["totals"]["C"], 79)
        self.assertIn({"check": "Số móng = số cột", "MT": 79, "C": 79, "result": "KHỚP"}, self.checks["findings"])

    def test_typical_floors_have_identical_beams(self):
        # 18 loại dầm tầng 2, cộng tay: 39 cái, Σ SL×L = 657 680 mm ; tầng 3 giống hệt ; mái (D4) khác
        self.assertEqual((self.checks["totals"]["D2"], self.checks["lengths_m"]["D2"]), (39, 657.68))
        self.assertEqual((self.checks["totals"]["D3"], self.checks["lengths_m"]["D3"]), (39, 657.68))
        results = {f["check"]: f["result"] for f in self.checks["findings"]}
        self.assertEqual(results["Dầm D2 so với D3"], "GIỐNG")
        self.assertEqual(results["Dầm D3 so với D4"], "KHÁC")

    def test_foundation_beams_and_ties(self):
        self.assertEqual(self.checks["lengths_m"]["GM"], 530.8)           # GM (SL=1; L=530800)
        self.assertEqual(self.checks["totals"]["DM"], 31)
        row = next(r for r in self.r["element_schedule"] if r["symbol"] == "DM-X-A-01")
        self.assertEqual((row["count"], row["length_mm"], row["page"]), (1, 43070, 20))

    def test_materials_from_ocr_notes(self):
        self.assertEqual(self.specs["concrete_class_grade_pairs"], {"B20": "M250", "B10": "M150"})
        self.assertEqual(self.specs["concrete_classes"], ["B10", "B20"])
        self.assertEqual(self.specs["steel_strengths"], ["RS=2800DAN/CM2"])

    def test_standards_complete_with_years(self):
        stds = self.specs["standards_cited"]
        self.assertIn("TCVN 1651:2008", stds)                              # số hiệu bị OCR tách sang dòng sau
        self.assertIn("TCVN 46:2007", stds)
        self.assertNotIn("TCVN 46-", stds)                                 # trước đây bị cụt do OCR "2OO7"
        self.assertNotIn("TCVN 46", stds)

    def test_no_bridge_span_on_building_dossier(self):
        self.assertNotIn("span_schema", self.specs["key_parameters"])     # trước đây nhận nhầm "2020+1020"

    def test_bridge_span_still_detected(self):
        specs = AECMarkdownIngestor().extract_technical_specifications(
            "Sơ đồ nhịp: 39.1m + 40.0m + 39.1m. Kích thước 2020+1020 mm.")
        self.assertEqual(specs["key_parameters"]["span_schema"], "39.1m + 40.0m + 39.1m")


@unittest.skipUnless(os.environ.get("AEC_TEST_DOSSIER"), "đặt AEC_TEST_DOSSIER để chạy trên toàn bộ hồ sơ")
class FullDossierTest(unittest.TestCase):

    def test_full_dossier_matches_excerpt(self):
        full, excerpt = ingest(os.environ["AEC_TEST_DOSSIER"]), ingest(FIXTURE)
        self.assertEqual(full["schedule_checks"], excerpt["schedule_checks"])


if __name__ == "__main__":
    unittest.main()
