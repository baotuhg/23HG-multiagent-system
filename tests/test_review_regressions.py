"""Regression coverage for missing input, unit conversion and export approval."""
import contextlib
import io
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from core.agents.aec_data_aggregator import AECDataAggregator
from core.agents.sub_agents import QSAgent
from core.gates.human_gate import ApprovalDecision
from core.state.shared_state import ProjectPhase
from core.supervisor.supervisor_agent import AECSupervisor
from tools.excel_eval import WorkbookEvaluator, _Parser, FormulaError
from tools.ifc_loader import IFCLoader, HAS_IFCOPENSHELL
from run_state_graph import build_parser, export_dossier, PHASE_MAP

ROOT = Path(__file__).resolve().parents[1]
MASTER = ROOT / 'templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx'


class ReviewRegressionTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        env = patch.dict(os.environ, {'AEC_STATE_DIR': self.tmp.name})
        env.start()
        self.addCleanup(env.stop)
        stdout = contextlib.redirect_stdout(io.StringIO())
        stdout.__enter__()
        self.addCleanup(stdout.__exit__, None, None, None)

    def test_missing_ifc_runtime_never_fabricates_quantities(self):
        p = Path(self.tmp.name) / 'empty.ifc'
        p.write_text("#1=IFCBEAM('abc',$,'Beam',$,$,$,$,$);", encoding='utf-8')
        loader = IFCLoader()
        loader.has_native_ifc = False
        result = loader.load_ifc(str(p))
        self.assertEqual(result.total_concrete_volume_m3, 0)
        self.assertTrue(result.errors)
        with self.assertRaises(ValueError):
            result.to_cutting_stock_demands()

    @unittest.skipUnless(HAS_IFCOPENSHELL, 'requires IfcOpenShell')
    def test_ifc_metre_and_millimetre_models_agree(self):
        import ifcopenshell
        results = []
        for prefix, scale in [(None, 1), ('MILLI', 1000)]:
            model = ifcopenshell.file(schema='IFC4')
            unit = model.create_entity('IfcSIUnit', UnitType='LENGTHUNIT', Name='METRE', Prefix=prefix)
            units = model.create_entity('IfcUnitAssignment', Units=[unit])
            model.create_entity('IfcProject', GlobalId=ifcopenshell.guid.new(), UnitsInContext=units)
            model.create_entity('IfcReinforcingBar', GlobalId=ifcopenshell.guid.new(),
                                Name='Bar', NominalDiameter=.025 * scale, BarLength=6.5 * scale)
            p = Path(self.tmp.name) / f'{scale}.ifc'
            model.write(str(p))
            result = IFCLoader().load_ifc(str(p))
            self.assertEqual(result.errors, [])
            bar = result.rebar_elements[0]
            self.assertAlmostEqual(bar.diameter_mm, 25)
            self.assertAlmostEqual(bar.length_mm, 6500)
            self.assertGreater(bar.total_weight_kg, 25)
            results.append(bar.total_weight_kg)
        self.assertEqual(results[0], results[1])

    def test_aggregation_missing_and_zero_inputs(self):
        agent = AECDataAggregator()
        with self.assertRaises(ValueError):
            agent.synthesize_to_project_state(self.tmp.name)
        state = agent.synthesize_to_project_state(self.tmp.name, office_data={'sheets_count': 1})
        self.assertEqual(state['meta']['cross_check_status'], 'NOT_CHECKED')
        self.assertNotIn('detailed_rebar_bbs_summary', state)
        state = agent.synthesize_to_project_state(self.tmp.name,
            cad_data={'summary_quantities': {'piles_d1200_count': 0}},
            md_data={'technical_specs': {'key_parameters': {'piles_count': 26}}})
        self.assertEqual(state['meta']['cross_check_status'], 'MISMATCH')

    def test_reject_does_not_overwrite_existing_financial_output(self):
        dest = Path(self.tmp.name) / 'gxd.xlsx'
        dest.write_bytes(b'existing approved document')
        sup = AECSupervisor(project_root=str(ROOT), demo_mode=True,
                            persist_path=str(Path(self.tmp.name) / 'state.json'))
        sup.register_agent(QSAgent(qs_out=str(dest)))
        sup.human_gate.request_approval = lambda req: ApprovalDecision(req.gate_id, False)
        self.assertFalse(sup.run([ProjectPhase.QS_ESTIMATE, ProjectPhase.HUMAN_GATE]))
        self.assertEqual(dest.read_bytes(), b'existing approved document')
        self.assertEqual(sup.bus.pending_legal_exports(), [])

    def test_single_phase_export_requires_approval_of_computed_result(self):
        dest = Path(self.tmp.name) / 'gxd.xlsx'
        sup = AECSupervisor(project_root=str(ROOT), demo_mode=True,
                            persist_path=str(Path(self.tmp.name) / 'state.json'))
        sup.register_agent(QSAgent(qs_out=str(dest)))
        def approve(req):
            self.assertFalse(dest.exists())
            self.assertGreater(sup.bus.get_qs_data().total_G_XD_vnd, 0)
            return ApprovalDecision(req.gate_id, True)
        sup.human_gate.request_approval = approve
        self.assertTrue(sup.run([ProjectPhase.QS_ESTIMATE]))
        self.assertTrue(dest.exists())

    def test_export_no_implicit_companions_and_propagates_failed_audit(self):
        args = build_parser().parse_args(['--demo', '--export-all', '--project-name', 'Unrelated',
                                         '--export-dir', self.tmp.name])
        with patch('tools.package_dispatcher.AECPackageDispatcher.dispatch_full_industrial_dossier',
                   return_value=SimpleNamespace(audit_zero_errors=False, summary_report='failed')) as dispatch:
            self.assertFalse(export_dossier(args, str(MASTER)))
        self.assertEqual(dispatch.call_args.kwargs['companion_files'], {})

    def test_export_rejection_never_calls_dispatcher(self):
        args = build_parser().parse_args(['--demo', '--export-all'])
        with patch('core.gates.human_gate.HumanGate.request_approval',
                   return_value=ApprovalDecision('export', False)), \
             patch('tools.package_dispatcher.AECPackageDispatcher.dispatch_full_industrial_dossier') as dispatch:
            self.assertFalse(export_dossier(args, str(MASTER)))
            dispatch.assert_not_called()

    def test_only_implemented_phases_advertised(self):
        self.assertTrue(set(PHASE_MAP.values()) <= set(AECSupervisor._PHASE_HANDLERS))


class LetEvaluationTest(unittest.TestCase):
    def evaluate(self, text):
        return _Parser(None, 'Sheet1', text).parse()

    def test_let_scopes_dependencies_and_case(self):
        self.assertEqual(self.evaluate('LET(qty,3,width,4,QTY*width)'), 12)
        self.assertEqual(self.evaluate('_xlfn.LET(_xlpm.n,2,_xlpm.w,_xlpm.n+3,_xlpm.n*_xlpm.w)'), 10)
        self.assertEqual(self.evaluate('LET(x,2,LET(x,7,x)+x)'), 9)

    def test_let_binding_does_not_leak(self):
        with self.assertRaises(FormulaError):
            self.evaluate('LET(x,2,x)+x')
        with self.assertRaises(FormulaError):
            self.evaluate('LET(x,y+1,x)')


class BridgeWorkbookLinksTest(unittest.TestCase):
    def test_super_t_subtotals_use_current_detail_rows(self):
        path = ROOT / 'examples/HO_SO_CAU_KM19_529/BANG_BOC_TACH_CHI_TIET_DAM_SUPER_T_38.2M_KM19.xlsx'
        evaluator = WorkbookEvaluator(str(path))
        evaluator._cached = {sheet: {} for sheet in evaluator.sheetnames}
        sheet = '00_DIEN_GIAI_HINH_HOC_CAD'
        self.assertAlmostEqual(evaluator.value(sheet, 48, 9), 89668.81)
        self.assertAlmostEqual(evaluator.value(sheet, 49, 9), 89.66881)
        self.assertAlmostEqual(evaluator.value(sheet, 42, 9), 77.25)
        self.assertAlmostEqual(evaluator.value(sheet, 51, 9), 19.38)
        evaluator._formulas[sheet][(51, 4)] = 4
        evaluator._computed.clear()
        self.assertAlmostEqual(evaluator.value(sheet, 51, 9), 25.84)

    def test_dashboard_tracks_detail_quantities_and_units(self):
        folder = ROOT / 'examples/HO_SO_CAU_KM19_529'
        names = ('BANG_BOC_TACH_TOAN_BO_BAN_VE_CAU_KM19.xlsx',
                 'BANG_BOC_TACH_CHI_TIET_TOAN_BO_BAN_VE_CAU_KM19.xlsx')
        for name in names:
            with self.subTest(workbook=name):
                evaluator = WorkbookEvaluator(str(folder / name))
                # Recompute rather than trust previously saved Excel caches.
                evaluator._cached = {sheet: {} for sheet in evaluator.sheetnames}
                dashboard = '00_DASHBOARD_TOAN_CAU'
                self.assertEqual(evaluator.value(dashboard, 33, 7), 30)
                self.assertEqual(evaluator.value(dashboard, 34, 7), 24.5)
                self.assertEqual(evaluator.value(dashboard, 35, 7), 24)
                self.assertEqual(evaluator.value('02_DIEN_GIAI_HA_BO', 52, 10), 872)
                master = '06_MASTER_BOQ_TIEN_LUONG'
                self.assertAlmostEqual(evaluator.value(dashboard, 32, 7),
                                       evaluator.value(master, 26, 7) * 1000)
                # A changed input must flow through detail -> BOQ -> dashboard.
                detail = '03_DIEN_GIAI_PHU_TRO'
                evaluator._formulas[detail][(31, 6)] = evaluator.value(detail, 31, 6) + 1
                evaluator._computed.clear()
                self.assertEqual(evaluator.value(dashboard, 33, 7), 31)


if __name__ == '__main__':
    unittest.main()
