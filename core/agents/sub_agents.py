# -*- coding: utf-8 -*-
"""
CAD AGENT — Sub-Agent Trắc đạc & Bóc tách CAD
Quét folder DWG → bóc tách khối lượng → ghi vào StateBus

Tích hợp với: aec_cad_extractor.py (COM Interop AutoCAD) + cad_takeoff_engine.py
"""

from __future__ import annotations
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from datetime import date
from typing import Iterable, Optional

from core.supervisor.base_agent import BaseAgent, DataInputError, MissingDataError
from core.state.state_bus import StateBus

DEMO_HINT = "Muốn chạy thử với dữ liệu mẫu Cầu Km19+529.080 thì thêm cờ --demo."


class CADAgent(BaseAgent):
    """
    Sub-Agent Trắc đạc & Bóc tách CAD.
    Input  (từ StateBus): drawings_folder path
    Output (vào StateBus): cad_data (concrete, formwork, excavation, drawings_processed)
    """

    def __init__(self, drawings_folder: str = ""):
        super().__init__(
            agent_id="cad_agent",
            description="Trắc đạc CAD — Shoelace + Average-End-Area + COM Interop"
        )
        self.drawings_folder = drawings_folder

    def run(self, bus: StateBus) -> bool:
        print("  [CADAgent] Bắt đầu quét bản vẽ CAD...")

        folder = self.drawings_folder or bus._state.drawings_folder
        problem = self._takeoff_from_drawings(bus, folder)
        if problem is None:
            return True

        if bus.is_demo_mode():
            print(f"  [CADAgent] {problem}")
            return self._use_sample_data(bus)
        raise MissingDataError(f"{problem} {DEMO_HINT}")

    def _takeoff_from_drawings(self, bus: StateBus, folder: str):
        """Bóc tách từ thư mục bản vẽ thật. Trả về None nếu thành công, ngược lại là lý do thiếu dữ liệu."""
        if not folder:
            return "Chưa chỉ định thư mục bản vẽ CAD (--drawings)."
        if not os.path.isdir(folder):
            return f"Không tìm thấy thư mục bản vẽ: {folder}"

        from agents.aec_cad_extractor import AECCadExtractor
        scan = AECCadExtractor().scan_drawings_folder(folder)
        drawings = scan.get("drawings", [])
        components = [d for d in drawings if d.get("volume_m3", 0) > 0]
        if not components:
            return (
                f"Đã quét {len(drawings)} bản vẽ trong {folder} nhưng bộ trích xuất CAD hiện chỉ "
                f"phân loại hạng mục theo tên file, chưa bóc được khối lượng bê tông/ván khuôn."
            )

        total_concrete = sum(c["volume_m3"] for c in components)
        bus.set_cad_data({
            "concrete_components": components,
            "total_concrete_m3": total_concrete,
            "drawings_scanned": len(drawings),
            "drawings_processed": len(components),
            "source_dwg_files": [d.get("file_path", "") for d in drawings],
        })
        print(f"  [CADAgent] Đã xử lý {len(components)} cấu kiện, tổng BT: {total_concrete:.2f} m³")
        return None

    def _use_sample_data(self, bus: StateBus) -> bool:
        """Dữ liệu mẫu kỹ thuật cho Cầu Km19+529.080."""
        sample_components = [
            {"id": "COC-T1-01", "name": "Cọc Ø1200 T1-01", "wbs": "KẾT CẤU MÓNG CỌC",
             "volume_m3": 45.24, "formwork_m2": 0.0, "rebar_kg": 1250.0},
            {"id": "COC-T2-01", "name": "Cọc Ø1200 T2-01", "wbs": "KẾT CẤU MÓNG CỌC",
             "volume_m3": 34.06, "formwork_m2": 0.0, "rebar_kg": 940.0},
            {"id": "BE-T1", "name": "Bệ trụ T1", "wbs": "KẾT CẤU TRỤ CẦU",
             "volume_m3": 112.5, "formwork_m2": 185.0, "rebar_kg": 8200.0},
            {"id": "THAN-T1", "name": "Thân đặc T1", "wbs": "KẾT CẤU TRỤ CẦU",
             "volume_m3": 65.8, "formwork_m2": 210.0, "rebar_kg": 4100.0},
            {"id": "XA-MU-T1", "name": "Xà mũ T1", "wbs": "KẾT CẤU TRỤ CẦU",
             "volume_m3": 28.4, "formwork_m2": 95.0, "rebar_kg": 3200.0},
            {"id": "DAM-ST-01", "name": "Dầm Super-T nhịp 1 D1", "wbs": "KẾT CẤU NHỊP",
             "volume_m3": 28.99, "formwork_m2": 0.0, "rebar_kg": 2100.0},
            {"id": "MAT-CAU", "name": "Bản mặt cầu giai đoạn 1", "wbs": "KẾT CẤU MẶT CẦU",
             "volume_m3": 305.47, "formwork_m2": 1560.0, "rebar_kg": 42000.0},
            {"id": "MO-M1", "name": "Mố M1 tổng thể", "wbs": "KẾT CẤU MỐ CẦU",
             "volume_m3": 185.6, "formwork_m2": 420.0, "rebar_kg": 15800.0},
        ]
        total = sum(c["volume_m3"] for c in sample_components)
        bus.mark_sample_data(self.agent_id, "8 cấu kiện mẫu Cầu Km19+529.080 viết sẵn trong code")
        bus.set_cad_data({
            "concrete_components": sample_components,
            "total_concrete_m3": total,
            "drawings_scanned": 61,
            "drawings_processed": len(sample_components),
            "source_dwg_files": [],
            "revision_tag": "Rev00-Sample",
        })
        print(f"  [CADAgent] Dữ liệu mẫu: {len(sample_components)} cấu kiện, tổng BT: {total:.2f} m³")
        return True


class QSAgent(BaseAgent):
    """
    Sub-Agent Dự toán — tính G_XD theo TT 11/2021/TT-BXD.
    Input  (từ StateBus): cad_data (khối lượng BT, VK, thép)
    Output (vào StateBus): qs_data (T, GT, TL, VAT, G_XD)
    """

    def __init__(self):
        super().__init__(
            agent_id="qs_agent",
            description="Dự toán G_XD — TT 11/2021/TT-BXD, VAT 10%"
        )

    def run(self, bus: StateBus) -> bool:
        print("  [QSAgent] Tính dự toán G_XD...")

        # Chưa có nguồn chi phí trực tiếp T thật (đơn giá × khối lượng) — chỉ có số mẫu
        if not bus.is_demo_mode():
            raise MissingDataError(
                "Chưa có nguồn chi phí trực tiếp T thật (khối lượng × đơn giá). "
                "T = 61,28 tỷ và tỷ lệ GT 7,3% / TL 5,5% trong code chỉ là số mẫu. " + DEMO_HINT
            )

        T = 61_280_000_000  # VNĐ — số mẫu Cầu Km19+529.080
        bus.mark_sample_data(self.agent_id, "chi phí trực tiếp T = 61,28 tỷ và tỷ lệ GT/TL mẫu")

        GT = round(T * 0.073)            # Chi phí gián tiếp 7.3%
        TL = round((T + GT) * 0.055)     # Lợi nhuận 5.5%
        subtotal = T + GT + TL
        VAT = round(subtotal * 0.10)     # VAT 10% (Luật XD 135/2025)
        G_XD = subtotal + VAT

        bus.set_qs_data({
            "direct_cost_T_vnd": T,
            "indirect_cost_GT_vnd": GT,
            "tax_TL_vnd": TL,
            "subtotal_vnd": subtotal,
            "vat_vnd": VAT,
            "total_G_XD_vnd": G_XD,
        })

        print(f"  [QSAgent] G_XD = {G_XD:,.0f} VNĐ "
              f"(T={T/1e9:.2f}B + GT={GT/1e9:.2f}B + TL={TL/1e9:.2f}B + VAT={VAT/1e9:.2f}B)")
        return True


class BPTCKCSAgent(BaseAgent):
    """
    Sub-Agent BPTC + KCS — lập biên bản nghiệm thu và kiểm soát chất lượng.
    Tích hợp QA/QC Lab Link: Ánh xạ phiếu thí nghiệm R7/R28, kéo thép, siêu âm cọc, PDA
    vào 22 biên bản nghiệm thu KCS (NĐ 207/2026/NĐ-CP, TT 32/2026/TT-BXD).
    """

    def __init__(self):
        super().__init__(
            agent_id="bptc_kcs_agent",
            description="BPTC + KCS & QA/QC Lab Link — NĐ 207/2026, TT 32/2026"
        )

    def run(self, bus: StateBus) -> bool:
        print("  [BPTCKCSAgent] Kiểm tra QA/QC, liên kết phiếu thí nghiệm Lab và lập biên bản...")

        if not bus.is_demo_mode():
            raise MissingDataError(
                "Chưa có kết quả thí nghiệm thật (phiếu nén R7/R28, kéo thép, siêu âm cọc, PDA). "
                "6 phiếu thí nghiệm và 4 điểm dừng kỹ thuật trong code chỉ là dữ liệu mẫu. " + DEMO_HINT
            )
        bus.mark_sample_data(self.agent_id, "6 phiếu thí nghiệm + 4 điểm dừng kỹ thuật mẫu (luôn PASS)")

        # ── BƯỚC 1: QA/QC Lab Link & Hold Points Check ────────────────────────
        lab_results, hold_points_cleared, lab_clashes = self._verify_lab_results_and_hold_points()
        print(f"  [BPTCKCSAgent] Đã liên kết {len(lab_results)} phiếu thí nghiệm vào hệ thống KCS")
        for hp in hold_points_cleared:
            print(f"  [BPTCKCSAgent] ✓ Giải tỏa điểm dừng kỹ thuật (Hold Point): {hp}")

        # ── BƯỚC 2: Thẩm tra logic chéo ngày tháng & kiểm toán file Excel ───────
        excel_path = bus._state.excel_master_path
        if not excel_path or not os.path.exists(excel_path):
            raise MissingDataError(
                f"Không tìm thấy Excel master để kiểm toán: {excel_path or '(chưa chỉ định --excel)'}"
            )
        audit_score = 0
        clashes = list(lab_clashes)

        try:
            from aec_core.audit_verifier import AECAuditVerifier
            auditor = AECAuditVerifier(excel_path)
            auditor.audit_excel_workbook()
            audit_score = auditor.score
            if audit_score < 100:
                clashes.append(f"Audit score {audit_score}/100 — chưa đạt 100/100")
        except Exception as e:
            clashes.append(f"Lỗi audit: {e}")

        # ── BƯỚC 3: Đồng bộ trạng thái vào StateBus ───────────────────────────
        bus.set_qaqc_data({
            "audit_score": audit_score,
            "clashes_detected": clashes,
            "date_cross_check_status": "PASSED" if not clashes else "FAILED",
            "total_inspection_records": 22,
            "hold_points": hold_points_cleared,
            "lab_results": [
                {
                    "test_id": r["test_id"],
                    "material_type": r["material_type"],
                    "sample_code": r["sample_code"],
                    "r7_mpa": r.get("r7_mpa", 0.0),
                    "r28_mpa": r.get("r28_mpa", 0.0),
                    "required_mpa": r.get("required_mpa", 0.0),
                    "status": r["status"],
                    "certificate_ref": r["certificate_ref"],
                    "kcs_record_linked": r["kcs_record_linked"],
                } for r in lab_results
            ]
        })

        return audit_score >= 100 and len(lab_clashes) == 0

    def _verify_lab_results_and_hold_points(self) -> tuple[list, list, list]:
        """Xác thực kết quả thí nghiệm phòng LAS-XD và giải tỏa Hold Points."""
        lab_results = [
            {
                "test_id": "LAS188-BT-01", "material_type": "CONCRETE", "sample_code": "M-COC-T1-01",
                "r7_mpa": 25.2, "r28_mpa": 33.5, "required_mpa": 30.0, "status": "PASS",
                "certificate_ref": "PTN-2026/088", "kcs_record_linked": "BBNT-04",
                "desc": "Bê tông C30 cọc khoan nhồi trụ T1"
            },
            {
                "test_id": "LAS188-BT-02", "material_type": "CONCRETE", "sample_code": "M-DAM-ST-01",
                "r7_mpa": 42.0, "r28_mpa": 52.8, "required_mpa": 45.0, "status": "PASS",
                "certificate_ref": "PTN-2026/102", "kcs_record_linked": "BBNT-14",
                "desc": "Bê tông C45/55 dầm Super-T (R7 đạt 93.3% R28)"
            },
            {
                "test_id": "LAS188-BT-03", "material_type": "CONCRETE", "sample_code": "M-BAN-MC-01",
                "r7_mpa": 29.8, "r28_mpa": 38.6, "required_mpa": 35.0, "status": "PASS",
                "certificate_ref": "PTN-2026/115", "kcs_record_linked": "BBNT-18",
                "desc": "Bê tông C35 bản mặt cầu"
            },
            {
                "test_id": "LAS188-STEEL-01", "material_type": "REBAR", "sample_code": "ST-D25-CB500",
                "r7_mpa": 0.0, "r28_mpa": 0.0, "required_mpa": 500.0, "status": "PASS",
                "certificate_ref": "CCXX-HP-2026-991", "kcs_record_linked": "BBNT-03",
                "desc": "Chứng chỉ kéo uốn thép Ø25 CB500-V (fy=542MPa, fu=668MPa)"
            },
            {
                "test_id": "LAS188-SONIC-01", "material_type": "PILE_INTEGRITY", "sample_code": "SONIC-156-SECTIONS",
                "r7_mpa": 0.0, "r28_mpa": 0.0, "required_mpa": 1.0, "status": "PASS",
                "certificate_ref": "BC-SA-2026/01", "kcs_record_linked": "BBNT-07",
                "desc": "Siêu âm cọc khoan nhồi 156 mặt cắt: 100% đạt Loại 1"
            },
            {
                "test_id": "LAS188-PDA-01", "material_type": "PDA_TEST", "sample_code": "PDA-COC-T1-02",
                "r7_mpa": 0.0, "r28_mpa": 0.0, "required_mpa": 7800.0, "status": "PASS",
                "certificate_ref": "BC-PDA-2026/02", "kcs_record_linked": "BBNT-08",
                "desc": "Nén động PDA cọc T1-02 đạt 9,434 kN (Sức chịu tải thiết kế 7,800 kN)"
            }
        ]

        hold_points = [
            "Đã nghiệm thu dò Karst 26 lỗ đạt 5m vào đá liền khối (BBNT-06)",
            "Đã siêu âm 156 mặt cắt cọc nhồi đạt 100% Loại 1 (BBNT-07)",
            "Thí nghiệm nén động PDA cọc đạt 9,434 kN vượt tải thiết kế (BBNT-08)",
            "Bê tông dầm Super-T đạt R28 = 52.8 MPa > 45 MPa, đủ điều kiện căng kéo cáp DƯL (BBNT-14)"
        ]

        clashes = []
        for r in lab_results:
            if r["status"] != "PASS":
                clashes.append(f"Phiếu thí nghiệm {r['test_id']} ({r['desc']}) KHÔNG ĐẠT chuẩn!")

        return lab_results, hold_points, clashes


class SchedulerAgent(BaseAgent):
    """
    Sub-Agent Tiến độ CPM — tính đường găng từ danh mục công việc THẬT.
    Nguồn: MS Project XML / Excel / CSV / JSON (schedule_path). Không có file → dừng,
    trừ chế độ --demo (dùng 11 công việc mẫu).
    """

    MAX_LISTED = 10

    def __init__(
        self,
        schedule_path: Optional[str] = None,
        schedule_sheet: Optional[str] = None,
        start_date: Optional[str] = None,
        non_working_weekdays: Iterable[int] = (),
        holidays: Iterable[date] = (),
        schedule_out: Optional[str] = None,
    ):
        super().__init__(
            agent_id="scheduler_agent",
            description="Tiến độ CPM — đường găng và As-Built tracking"
        )
        self.schedule_path = schedule_path
        self.schedule_sheet = schedule_sheet
        self.start_date = start_date
        self.non_working_weekdays = set(non_working_weekdays)
        self.holidays = set(holidays)
        self.schedule_out = schedule_out

    def run(self, bus: StateBus) -> bool:
        print("  [SchedulerAgent] Tính CPM tiến độ...")

        from tools.cpm_calculator import CPMCalculator

        tasks, source, file_start, warnings = self._load_tasks(bus)
        start_date = self.start_date or file_start
        if not start_date:
            warnings.append("Chưa có ngày khởi công (--start-date) — chỉ tính theo số ngày, không quy đổi ngày lịch")

        result = CPMCalculator().calculate(
            tasks, start_date_str=start_date or "",
            non_working_weekdays=self.non_working_weekdays, holidays=self.holidays,
        )
        if result.status == "ERROR":
            raise DataInputError(
                f"Tiến độ {source} có lỗi logic:\n    " + "\n    ".join(result.warnings[:self.MAX_LISTED])
            )
        warnings.extend(result.warnings)

        # Đối chiếu ngày ghi trong file với kết quả tính
        file_dates = {t["id"]: (t.get("file_start", ""), t.get("file_finish", "")) for t in tasks}
        differ = [
            t for t in result.tasks
            if all(file_dates.get(t.task_id, ("", ""))) and t.start_date
            and file_dates[t.task_id] != (t.start_date, t.finish_date)
        ]
        if differ:
            warnings.append(
                f"{len(differ)}/{len(result.tasks)} công việc có ngày trong file khác kết quả tính CPM "
                f"(với lịch nghỉ đang chọn)"
            )

        self._codes = {t["id"]: str(t.get("code") or t["id"]) for t in tasks}
        critical_codes = [self._codes[c] for c in result.critical_path]
        bus.set_schedule_data({
            "start_date": result.project_start,
            "finish_date": result.project_finish,
            "total_duration_days": result.total_duration_days,
            "critical_path": result.critical_path,
            "tasks": [self._task_dict(t, file_dates.get(t.task_id, ("", ""))) for t in result.tasks],
            "overall_progress_pct": result.overall_progress_pct,
            "delay_days": result.delay_days,
        })

        print(f"  [SchedulerAgent] Nguồn tiến độ: {source} — {len(result.tasks)} công việc")
        print(f"  [SchedulerAgent] Tổng thời gian: {result.total_duration_days} ngày làm việc")
        if result.project_finish:
            print(f"  [SchedulerAgent] Khởi công {result.project_start} → hoàn thành {result.project_finish}")
        print(f"  [SchedulerAgent] Đường găng ({len(critical_codes)} việc): " + " → ".join(critical_codes))
        for w in warnings[:self.MAX_LISTED]:
            print(f"    ⚠ {w}")
        if len(warnings) > self.MAX_LISTED:
            print(f"    ... và {len(warnings) - self.MAX_LISTED} cảnh báo khác")

        if self.schedule_out:
            self._write_csv(result, file_dates)
            print(f"  [SchedulerAgent] Đã xuất bảng tiến độ CPM: {self.schedule_out}")
        return True

    def _load_tasks(self, bus: StateBus):
        """Trả về (tasks, nguồn, ngày khởi công trong file, cảnh báo)."""
        if self.schedule_path:
            from tools.schedule_loader import ScheduleLoadError, find_date_violations, load_schedule
            try:
                loaded = load_schedule(self.schedule_path, sheet=self.schedule_sheet)
            except ScheduleLoadError as e:
                raise DataInputError(str(e)) from e
            if loaded.errors:
                listed = "\n    ".join(loaded.errors[:self.MAX_LISTED])
                more = f"\n    ... và {len(loaded.errors) - self.MAX_LISTED} dòng khác" \
                    if len(loaded.errors) > self.MAX_LISTED else ""
                raise DataInputError(
                    f"Tiến độ {loaded.source} có {len(loaded.errors)} dòng sai dữ liệu — sửa file:\n    "
                    f"{listed}{more}"
                )
            violations = find_date_violations(loaded.tasks)
            warnings = [f"Ngày trong file vi phạm quan hệ logic: {v}" for v in violations]
            return loaded.tasks, loaded.source, loaded.project_start, warnings

        if not bus.is_demo_mode():
            raise MissingDataError(
                "Chưa có danh mục công việc tiến độ thật — chỉ định file bằng "
                "--schedule <file.xml|.xlsx|.csv> (MS Project XML hoặc bảng Excel). " + DEMO_HINT
            )
        bus.mark_sample_data(self.agent_id, "11 công việc tiến độ mẫu viết sẵn trong code")
        return self._sample_tasks(), "11 công việc mẫu Cầu Km19 (demo)", "2026-10-01", []

    def _task_dict(self, t, file_dates) -> dict:
        codes = getattr(self, "_codes", {})
        return {
            "task_id": t.task_id, "code": codes.get(t.task_id, t.task_id),
            "name": t.name, "duration_days": t.duration_days,
            "predecessors": [
                f"{codes.get(l.pred_id, l.pred_id)}{l.type}" + (f"{l.lag_days:+g}d" if l.lag_days else "")
                for l in t.links
            ],
            "early_start": t.start_date or t.es, "early_finish": t.finish_date or t.ef,
            "float_days": t.tf, "is_critical": t.is_critical,
            "file_start": file_dates[0], "file_finish": file_dates[1],
        }

    def _write_csv(self, result, file_dates) -> None:
        import csv
        with open(self.schedule_out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Mã", "Công việc", "Thời gian (ngày)", "Quan hệ", "ES", "EF", "LS", "LF",
                        "Dự trữ TF (ngày)", "Găng", "Bắt đầu (tính)", "Kết thúc (tính)",
                        "Bắt đầu (file)", "Kết thúc (file)"])
            for t in result.tasks:
                d = self._task_dict(t, file_dates.get(t.task_id, ("", "")))
                w.writerow([d["code"], t.name, t.duration_days, "; ".join(d["predecessors"]),
                            t.es, t.ef, t.ls, t.lf, t.tf, "X" if t.is_critical else "",
                            t.start_date, t.finish_date, d["file_start"], d["file_finish"]])

    @staticmethod
    def _sample_tasks() -> list:
        """11 công việc mẫu Cầu Km19+529.080 — CHỈ dùng ở chế độ --demo."""
        return [
            {"id": "T01", "name": "Tim mốc định vị", "duration": 3, "predecessors": []},
            {"id": "T02", "name": "Đường công vụ", "duration": 7, "predecessors": ["T01"]},
            {"id": "T03", "name": "Khoan dò Karst", "duration": 14, "predecessors": ["T02"]},
            {"id": "T04", "name": "Cọc nhồi T1/T2", "duration": 30, "predecessors": ["T03"]},
            {"id": "T05", "name": "Bệ trụ T1/T2", "duration": 20, "predecessors": ["T04"]},
            {"id": "T06", "name": "Thân đặc T1/T2", "duration": 25, "predecessors": ["T05"]},
            {"id": "T07", "name": "Xà mũ T1/T2", "duration": 15, "predecessors": ["T06"]},
            {"id": "T08", "name": "Lao dầm Super-T", "duration": 10, "predecessors": ["T07"]},
            {"id": "T09", "name": "Mặt cầu C35", "duration": 30, "predecessors": ["T08"]},
            {"id": "T10", "name": "Thảm BTN C16", "duration": 7, "predecessors": ["T09"]},
            {"id": "T11", "name": "Thử tải", "duration": 5, "predecessors": ["T10"]},
        ]
