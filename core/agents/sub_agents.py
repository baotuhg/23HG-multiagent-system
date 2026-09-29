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
from typing import Dict, Iterable, Optional

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
    Sub-Agent Dự toán — tính G_XD theo TT 11/2021/TT-BXD từ bảng QS THẬT.
    Input : file QS/BOQ (qs_path): khối lượng × đơn giá từng công tác; tỷ lệ chi phí lấy từ
            sheet tổng hợp G_XD trong file hoặc tham số (rate_overrides, đơn vị %)
    Output (vào StateBus): qs_data (T, GT, TL, G, VAT, G_XD, tỷ lệ, số công tác)
    Không có file → dừng, trừ chế độ --demo (dùng bảng QS của workbook mẫu Cầu Km19).
    """

    MAX_LISTED = 10

    def __init__(
        self,
        qs_path: Optional[str] = None,
        qs_sheet: Optional[str] = None,
        rate_overrides: Optional[Dict[str, Optional[float]]] = None,
        qs_out: Optional[str] = None,
        sample_path: Optional[str] = None,
    ):
        super().__init__(
            agent_id="qs_agent",
            description="Dự toán G_XD — TT 11/2021/TT-BXD từ bảng QS"
        )
        self.qs_path = qs_path
        self.qs_sheet = qs_sheet
        self.rate_overrides = rate_overrides or {}
        self.qs_out = qs_out
        self.sample_path = sample_path or os.path.join(
            _ROOT, "templates", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx")

    def run(self, bus: StateBus) -> bool:
        print("  [QSAgent] Tính dự toán G_XD...")
        from tools.qs_loader import QSLoadError, RATE_LABELS, apply_rate_overrides, load_qs

        path = self.qs_path
        if not path:
            if not bus.is_demo_mode():
                raise MissingDataError(
                    "Chưa có bảng QS thật — chỉ định bằng --qs <file.xlsx|.csv|.json> "
                    "(bảng khối lượng × đơn giá, kèm sheet tổng hợp G_XD hoặc tham số --rate-*). " + DEMO_HINT
                )
            path = self.sample_path
            bus.mark_sample_data(self.agent_id, "bảng QS và tỷ lệ chi phí của workbook mẫu Cầu Km19")

        try:
            est = load_qs(path, sheet=self.qs_sheet)
        except QSLoadError as e:
            raise DataInputError(str(e)) from e
        if est.errors:
            listed = "\n    ".join(est.errors[:self.MAX_LISTED])
            raise DataInputError(f"Bảng QS {est.source} có {len(est.errors)} dòng sai dữ liệu:\n    {listed}")
        apply_rate_overrides(est, self.rate_overrides)
        if est.missing_rates:
            flags = {"chung": "--rate-chung", "nha_tam": "--rate-nha-tam", "kxd": "--rate-kxd",
                     "tl": "--rate-tl", "vat": "--vat"}
            raise MissingDataError(
                "Thiếu tỷ lệ tính G_XD (không có trong sheet tổng hợp của file, không tự điền mặc định): "
                + "; ".join(f"{RATE_LABELS[k]} → {flags[k]}" for k in est.missing_rates)
            )
        est.compute()

        print(f"  [QSAgent] Nguồn QS: {est.source} — {len(est.items)} công tác"
              + (f" (tự tính {est.evaluated_cells:,} ô công thức chưa có kết quả)" if est.evaluated_cells else ""))
        for key, rate in est.rates.items():
            print(f"    {RATE_LABELS[key]}: {rate * 100:g}%  ← {est.rate_sources.get(key, '')}")
        print(f"  [QSAgent] T = {est.T:,} | GT = {est.GT:,} | TL = {est.TL:,} | G = {est.G:,} | "
              f"VAT = {est.VAT:,}")
        print(f"  [QSAgent] G_XD = {est.G_XD:,} VNĐ")
        for w in est.warnings[:self.MAX_LISTED]:
            print(f"    ⚠ {w}")
        if len(est.warnings) > self.MAX_LISTED:
            print(f"    ... và {len(est.warnings) - self.MAX_LISTED} cảnh báo khác")

        bus.set_qs_data({
            "direct_cost_T_vnd": est.T,
            "indirect_cost_GT_vnd": est.GT,
            "gt_breakdown_vnd": dict(est.GT_components),
            "tax_TL_vnd": est.TL,
            "subtotal_vnd": est.G,
            "vat_vnd": est.VAT,
            "total_G_XD_vnd": est.G_XD,
            "rates": dict(est.rates),
            "rate_sources": dict(est.rate_sources),
            "items_count": len(est.items),
            "data_source": est.source,
            "warnings": list(est.warnings),
        })

        if self.qs_out:
            from tools.qs_export import write_gxd_workbook
            write_gxd_workbook(self.qs_out, est)
            print(f"  [QSAgent] Đã xuất bảng dự toán G_XD: {self.qs_out}")
        return True


class BPTCKCSAgent(BaseAgent):
    """
    Sub-Agent BPTC + KCS — đánh giá phiếu thí nghiệm và điểm dừng kỹ thuật (Hold Point).

    Input : file phiếu thí nghiệm thật (--lab: .xlsx/.csv/.json) — nén bê tông R7/R28,
            kéo thép, siêu âm cọc, PDA, độ sụt... (đọc bằng tools/lab_qaqc.py)
    Output: qaqc_data.lab_results, lab_summary, hold_point_status, clashes_detected;
            báo cáo Excel (--lab-out)
    Phiếu KHÔNG ĐẠT → điểm dừng liên quan bị CHẶN → phase dừng (trả về False).
    Kiểm toán Excel master (GATE-4) do Supervisor chạy sau agent này, chỉ khi có --excel.
    Thiếu file → dừng; --demo dùng templates/Phieu_thi_nghiem_mau.csv (đánh dấu mẫu).
    """

    MAX_LISTED = 10

    def __init__(
        self,
        lab_path: Optional[str] = None,
        lab_sheet: Optional[str] = None,
        lab_out: Optional[str] = None,
        criteria=None,
        sample_path: Optional[str] = None,
    ):
        super().__init__(
            agent_id="bptc_kcs_agent",
            description="BPTC + KCS & QA/QC Lab Link — đánh giá phiếu thí nghiệm & điểm dừng kỹ thuật"
        )
        self.lab_path = lab_path
        self.lab_sheet = lab_sheet
        self.lab_out = lab_out
        self.criteria = criteria
        self.sample_path = sample_path or os.path.join(_ROOT, "templates", "Phieu_thi_nghiem_mau.csv")

    def run(self, bus: StateBus) -> bool:
        from tools.lab_qaqc import LabLoadError, evaluate, load_lab_results, write_lab_report

        print("  [BPTCKCSAgent] Đánh giá phiếu thí nghiệm và điểm dừng kỹ thuật...")
        path = self.lab_path
        if not path:
            if not bus.is_demo_mode():
                raise MissingDataError(
                    "Chưa có file phiếu thí nghiệm (--lab .xlsx/.csv/.json: nén R7/R28, kéo thép, "
                    "siêu âm cọc, PDA...). " + DEMO_HINT
                )
            path = self.sample_path
            bus.mark_sample_data(self.agent_id, f"Phiếu thí nghiệm mẫu {os.path.basename(path)}")
        elif not os.path.exists(path):
            raise MissingDataError(f"Không tìm thấy file phiếu thí nghiệm: {path}")

        try:
            rows = load_lab_results(path, sheet=self.lab_sheet)
        except LabLoadError as e:
            raise DataInputError(str(e)) from e
        ev = evaluate(rows, path, self.criteria)
        if ev.errors:
            raise DataInputError(
                f"File phiếu thí nghiệm {path} có {len(ev.errors)} lỗi dữ liệu:\n    "
                + "\n    ".join(ev.errors[:self.MAX_LISTED])
            )

        summary = {"total": len(ev.records), "pass": ev.count("PASS"),
                   "fail": ev.count("FAIL"), "pending": ev.count("PENDING"), "source": path}
        hold_points = [{"bbnt": h.bbnt, "status": h.status, "tests": list(h.tests),
                        "reasons": list(h.reasons)} for h in ev.hold_points]
        failed = [r for r in ev.records if r.status == "FAIL"]
        clashes = [f"Phiếu {r.test_id} ({r.component}) KHÔNG ĐẠT: {r.detail}" for r in failed]

        bus.set_qaqc_data({
            "lab_results": [{
                "test_id": r.test_id, "kind": r.kind, "component": r.component, "bbnt": r.bbnt,
                "status": r.status, "value": r.value, "required": r.required, "ratio": r.ratio,
                "age_days": r.age_days, "detail": r.detail,
            } for r in ev.records],
            "lab_summary": summary,
            "hold_point_status": hold_points,
            "hold_points": [f"{h['bbnt']}: {h['status']}" for h in hold_points],
            "clashes_detected": clashes,
            "date_cross_check_status": "FAILED" if failed else "PASSED",
            "total_inspection_records": len({h["bbnt"] for h in hold_points}),
            "warnings": list(ev.warnings),
        })

        print(f"  [BPTCKCSAgent] {summary['total']} phiếu: {summary['pass']} đạt, "
              f"{summary['fail']} không đạt, {summary['pending']} chờ kết quả")
        for h in hold_points:
            mark = {"GIẢI TỎA": "✓", "CHỜ": "…", "CHẶN": "✗"}.get(h["status"], "?")
            print(f"  [BPTCKCSAgent] {mark} Điểm dừng {h['bbnt']}: {h['status']}")
        for w in ev.warnings[:self.MAX_LISTED]:
            print(f"  [BPTCKCSAgent] ⚠ {w}")

        if self.lab_out:
            write_lab_report(self.lab_out, ev, self.criteria)
            print(f"  [BPTCKCSAgent] Đã xuất báo cáo thí nghiệm: {self.lab_out}")
        return not failed


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
