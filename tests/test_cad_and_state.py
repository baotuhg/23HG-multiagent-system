# -*- coding: utf-8 -*-
"""Kiểm thử đọc DXF, đối soát CAD ↔ hồ sơ, lưu / khôi phục state — python -m unittest discover tests"""

import os
import sys
import tempfile
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import core.agents.aec_cad_extractor as cad
from core.state.shared_state import ProjectSharedState, QAQCData
from core.state.state_bus import StateBus


def _pairs(*items):
    return "\n".join(str(x) for x in items) + "\n"


DXF = _pairs(
    0, "SECTION", 2, "HEADER", 9, "$INSUNITS", 70, 6, 0, "ENDSEC",
    0, "SECTION", 2, "ENTITIES",
    # Hình chữ nhật 10 × 5 khép kín, layer BT_BE
    0, "LWPOLYLINE", 8, "BT_BE", 90, 4, 70, 1,
    10, 0, 20, 0, 10, 10, 20, 0, 10, 10, 20, 5, 10, 0, 20, 5,
    # Đa tuyến hở — bỏ qua
    0, "LWPOLYLINE", 8, "BT_BE", 90, 2, 70, 0, 10, 0, 20, 0, 10, 3, 20, 3,
    # POLYLINE kiểu cũ: tam giác vuông 4 × 3 khép kín, layer DAO_DAT
    0, "POLYLINE", 8, "DAO_DAT", 66, 1, 70, 1,
    0, "VERTEX", 8, "DAO_DAT", 10, 0, 20, 0,
    0, "VERTEX", 8, "DAO_DAT", 10, 4, 20, 0,
    0, "VERTEX", 8, "DAO_DAT", 10, 0, 20, 3,
    0, "SEQEND",
    0, "ENDSEC", 0, "EOF",
)


class DxfReaderTest(unittest.TestCase):

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.path = os.path.join(tmp.name, "mat_cat.dxf")
        if cad.HAS_EZDXF:
            doc = cad.ezdxf.new("R2000")
            doc.header["$INSUNITS"] = 6
            msp = doc.modelspace()
            msp.add_lwpolyline([(0, 0), (10, 0), (10, 5), (0, 5)], close=True, dxfattribs={"layer": "BT_BE"})
            msp.add_lwpolyline([(0, 0), (3, 3)], close=False, dxfattribs={"layer": "BT_BE"})
            msp.add_polyline2d([(0, 0), (4, 0), (0, 3)], close=True, dxfattribs={"layer": "DAO_DAT"})
            doc.saveas(self.path)
        else:
            with open(self.path, "w", encoding="utf-8") as f:
                f.write(DXF)

    def test_ascii_reader(self):
        with mock.patch.object(cad, "HAS_EZDXF", False):
            res = cad.read_dxf_closed_areas(self.path)
        self.assertEqual(res["reader"], "ascii")
        self.assertEqual(res["units"], "m")
        self.assertEqual(res["closed_polylines"], 2)
        self.assertEqual(res["open_polylines_skipped"], 1)
        self.assertAlmostEqual(res["layers"]["BT_BE"]["total_area"], 50.0)
        self.assertAlmostEqual(res["layers"]["DAO_DAT"]["total_area"], 6.0)

    @unittest.skipUnless(cad.HAS_EZDXF, "ezdxf chưa cài")
    def test_ezdxf_reader_matches_ascii(self):
        res = cad.read_dxf_closed_areas(self.path)
        self.assertEqual(res["reader"], "ezdxf")
        self.assertAlmostEqual(res["layers"]["BT_BE"]["total_area"], 50.0)
        self.assertAlmostEqual(res["layers"]["DAO_DAT"]["total_area"], 6.0)

    def test_process_drawing_includes_areas(self):
        with mock.patch.object(cad, "HAS_EZDXF", False), mock.patch("builtins.print"):
            info = cad.AECCadExtractor().process_drawing(self.path)
        self.assertEqual(info["dxf_areas"]["closed_polylines"], 2)

    def test_rejects_dwg(self):
        with self.assertRaises(ValueError):
            cad.read_dxf_closed_areas("ban_ve.dwg")


class ReconcileTest(unittest.TestCase):

    def test_no_data_is_not_a_match(self):
        recs = cad.AECCadExtractor().reconcile_with_design_documents({}, {}, {})
        self.assertTrue(recs)
        self.assertTrue(all(r["status"] == "NOT_CHECKED" for r in recs))

    def test_match_and_mismatch(self):
        recs = cad.AECCadExtractor().reconcile_with_design_documents(
            {"total_drawings": 12, "concrete_grades": ["C30", "C45"]},
            {"drawing_list_count": 12},
            {"technical_specs": {"concrete_grades": ["c30", "C 40"]}},
        )
        status = {r["check_item"]: r["status"] for r in recs}
        self.assertEqual(status["Số lượng bản vẽ thiết kế"], "MATCHED")
        self.assertEqual(status["Mác bê tông"], "MISMATCH")
        self.assertEqual(status["Mác thép"], "NOT_CHECKED")


class StatePersistTest(unittest.TestCase):

    def test_roundtrip_restores_dataclasses(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "sub", "state.json")
            bus = StateBus(ProjectSharedState(), persist_path=path)
            bus.set_qaqc_data({"audit_score": 87, "lab_summary": {"total": 3}})
            self.assertTrue(os.path.exists(path))
            self.assertEqual([f for f in os.listdir(os.path.dirname(path)) if f.endswith(".tmp")], [])
            with mock.patch("builtins.print"):
                restored = StateBus.load_from_file(path)
        qaqc = restored.get_qaqc_data()
        self.assertIsInstance(qaqc, QAQCData)
        self.assertEqual(qaqc.audit_score, 87)
        self.assertEqual(qaqc.lab_summary, {"total": 3})


class CheckInputsCliTest(unittest.TestCase):

    def run_cli(self, *args):
        import subprocess
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
        return subprocess.run([sys.executable, os.path.join(ROOT, "run_state_graph.py"), "--check-inputs", *args],
                              capture_output=True, text=True, encoding="utf-8", env=env, cwd=ROOT)

    def test_valid_and_invalid_inputs(self):
        lab = os.path.join(ROOT, "templates", "Phieu_thi_nghiem_mau.csv")
        ok = self.run_cli("--lab", lab)
        self.assertEqual(ok.returncode, 0, ok.stdout + ok.stderr)
        self.assertIn("8 phiếu", ok.stdout)
        bad = self.run_cli("--lab", lab, "--bbs", os.path.join(ROOT, "khong_co_file.xlsx"))
        self.assertEqual(bad.returncode, 1)
        self.assertEqual(self.run_cli().returncode, 1)


if __name__ == "__main__":
    unittest.main()
