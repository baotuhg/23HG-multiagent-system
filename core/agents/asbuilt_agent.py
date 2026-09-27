# -*- coding: utf-8 -*-
"""
AS-BUILT AGENT — Sub-Agent Vòng lặp Đối soát Hiện trường (Site Feedback Loop)
Nhiệm vụ:
  1. Thu nhận dữ liệu nhật ký thi công hàng ngày (Daily Site Logs)
  2. Đối chiếu Khối lượng Hoàn công vs Khối lượng Thiết kế (Actual vs Planned)
  3. Cập nhật lại đường găng tiến độ CPM và độ trễ công trường (Schedule Delay Tracking)
  4. Lập bảng xác định khối lượng phát sinh phục vụ Phụ lục 03a (Nghị định 99/2021/NĐ-CP)
  5. Đồng bộ trạng thái vào Shared State Bus
"""

from __future__ import annotations
import os
import sys
from typing import Any, Dict, List, Optional
from datetime import date, timedelta

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.supervisor.base_agent import BaseAgent, MissingDataError
from core.state.state_bus import StateBus
from tools.cpm_calculator import CPMCalculator
from schemas.site_log_schema import (
    DailySiteLogEntry, SiteTaskExecution, AsBuiltVarianceItem
)


class AsBuiltAgent(BaseAgent):
    """
    As-Built Agent thực hiện vòng lặp đóng (closed-loop) giữa hiện trường và văn phòng.
    """

    def __init__(self, site_logs_input: Optional[List[Dict[str, Any]]] = None):
        super().__init__(
            agent_id="asbuilt_agent",
            description="Vòng lặp Đối soát Hiện trường — As-Built vs Planned, CPM Update & Phụ lục 03a"
        )
        self.site_logs_input = site_logs_input or []

    def run(self, bus: StateBus) -> bool:
        print("  [AsBuiltAgent] Khởi động vòng lặp đối soát hiện trường...")

        # Khối lượng hợp đồng, đơn giá và tiến độ gốc để đối soát hiện đều viết sẵn trong code
        if not bus.is_demo_mode():
            raise MissingDataError(
                "Chưa có dữ liệu đối soát hoàn công thật (nhật ký hiện trường, khối lượng hợp đồng, "
                "đơn giá, tiến độ gốc). Các số liệu này trong code chỉ là mẫu Cầu Km19+529.080. "
                "Muốn chạy thử với dữ liệu mẫu thì thêm cờ --demo."
            )
        bus.mark_sample_data(
            self.agent_id,
            "khối lượng hợp đồng, đơn giá, tiến độ gốc"
            + ("" if self.site_logs_input else " và 4 nhật ký hiện trường")
            + " mẫu viết sẵn trong code",
        )

        # ── BƯỚC 1: Thu thập / giả lập nhật ký công trường thực tế ─────────────
        logs = self._collect_site_logs(bus)
        print(f"  [AsBuiltAgent] Tiếp nhận {len(logs)} ngày nhật ký hiện trường")

        # ── BƯỚC 2: Tổng hợp khối lượng thực tế và đối chiếu thiết kế ───────────
        actual_progress, variance_items = self._reconcile_quantities(logs, bus)
        print(f"  [AsBuiltAgent] Đã đối soát {len(variance_items)} hạng mục hoàn công")

        # In tóm tắt chênh lệch
        total_delta_cost = sum(v.delta_amount_vnd for v in variance_items)
        print(f"  [AsBuiltAgent] Tổng giá trị biến động phát sinh: {total_delta_cost:+,.0f} VNĐ")

        # ── BƯỚC 3: Cập nhật lại tiến độ CPM dựa trên tiến độ thực tế ──────────
        cpm_result = self._update_cpm_schedule(actual_progress, bus)
        print(f"  [AsBuiltAgent] Tiến độ thi công thực tế: {cpm_result.overall_progress_pct:.1f}%")
        print(f"  [AsBuiltAgent] Độ lệch tiến độ (Delay): {cpm_result.delay_days} ngày so với kế hoạch")
        if cpm_result.delay_days > 0:
            print(f"  [AsBuiltAgent] ⚠ Cảnh báo: Công trình đang trễ {cpm_result.delay_days} ngày trên đường găng!")

        # ── BƯỚC 4: Cập nhật StateBus ──────────────────────────────────────────
        # Ghi nhật ký vào StateBus
        for log in logs:
            bus.push_site_log({
                "log_date": log.log_date,
                "weather": log.weather,
                "tasks_executed": [
                    {
                        "task_id": t.task_id,
                        "component_id": t.component_id,
                        "actual_volume": t.actual_volume,
                        "unit": t.unit,
                        "quality_status": t.quality_status
                    } for t in log.tasks_executed
                ],
                "issues": log.incidents_or_delays,
                "logged_by": log.site_engineer or "KS_HIEN_TRUONG"
            })

        # Ghi các hạng mục phát sinh Phụ lục 03a vào qs_data
        supplement_dicts = [
            {
                "task_id": v.task_id,
                "component_id": v.component_id,
                "description": v.description,
                "unit": v.unit,
                "contract_volume": v.contract_volume,
                "as_built_volume": v.as_built_volume,
                "delta_volume": v.delta_volume,
                "unit_price_vnd": v.unit_price_vnd,
                "delta_amount_vnd": v.delta_amount_vnd,
                "category": v.appendix_03a_category,
                "justification": v.justification
            } for v in variance_items
        ]

        bus.set_qs_data({
            "supplement_items": supplement_dicts,
            "payment_period_03a_vnd": total_delta_cost
        })

        # Cập nhật schedule_data
        bus.set_schedule_data({
            "overall_progress_pct": cpm_result.overall_progress_pct,
            "delay_days": cpm_result.delay_days,
            "critical_path": cpm_result.critical_path
        })

        print("  [AsBuiltAgent] ✅ Hoàn thành đồng bộ vòng lặp hiện trường vào StateBus")
        return True

    def _collect_site_logs(self, bus: StateBus) -> List[DailySiteLogEntry]:
        """Thu thập dữ liệu nhật ký hiện trường từ file/API hoặc tạo mẫu thực tế."""
        if self.site_logs_input:
            entries = []
            for d in self.site_logs_input:
                tasks = [SiteTaskExecution(**t) for t in d.get("tasks_executed", [])]
                entries.append(DailySiteLogEntry(
                    log_id=d.get("log_id", ""),
                    log_date=d.get("log_date", ""),
                    weather=d.get("weather", "Nắng"),
                    tasks_executed=tasks,
                    incidents_or_delays=d.get("incidents_or_delays", [])
                ))
            return entries

        # Mẫu hiện trường giai đoạn thi công móng cọc & bệ trụ Cầu Km19+529.080
        sample_logs = [
            DailySiteLogEntry(
                log_id="LOG-2026-10-05",
                log_date="2026-10-05",
                weather="Nắng ráo",
                tasks_executed=[
                    SiteTaskExecution(
                        task_id="T01", wbs_code="KẾT CẤU CHUNG", component_id="TIM-MOC",
                        description="Bàn giao tim mốc định vị mố M1, M2 và trụ T1, T2",
                        unit="Điểm", planned_volume=4.0, actual_volume=4.0,
                        quality_status="PASS", notes="Bàn giao mốc cao tọa độ chuẩn xác"
                    )
                ]
            ),
            DailySiteLogEntry(
                log_id="LOG-2026-10-15",
                log_date="2026-10-15",
                weather="Mưa nhẹ",
                tasks_executed=[
                    SiteTaskExecution(
                        task_id="T02", wbs_code="KẾT CẤU CHUNG", component_id="DUONG-CONG-VU",
                        description="Đắp đường công vụ tiếp cận trụ T1 và bãi đúc dầm",
                        unit="m", planned_volume=350.0, actual_volume=365.0, # Phát sinh thêm 15m do sình lầy
                        quality_status="PASS", notes="Mở rộng đường tránh xe cẩu 50 tấn"
                    )
                ],
                incidents_or_delays=["Mưa trưa 2h, làm chậm lu lèn nền đường"]
            ),
            DailySiteLogEntry(
                log_id="LOG-2026-10-28",
                log_date="2026-10-28",
                weather="Nắng đẹp",
                tasks_executed=[
                    SiteTaskExecution(
                        task_id="T03", wbs_code="KẾT CẤU MÓNG CỌC", component_id="DO-KARST",
                        description="Khoan dò Karst 26 vị trí cọc nhồi trụ T1 và T2",
                        unit="Lỗ", planned_volume=26.0, actual_volume=26.0,
                        quality_status="PASS", notes="Chiều sâu khoan dò đạt 5.0m vào đá liền khối"
                    )
                ]
            ),
            DailySiteLogEntry(
                log_id="LOG-2026-11-20",
                log_date="2026-11-20",
                weather="Nắng nhẹ",
                tasks_executed=[
                    SiteTaskExecution(
                        task_id="T04", wbs_code="KẾT CẤU MÓNG CỌC", component_id="COC-T1",
                        description="Khoan và đổ bê tông cọc nhồi D1200 trụ T1 (8 cọc)",
                        unit="m", planned_volume=320.0, actual_volume=328.0, # Phát sinh thêm 8m do cao độ đá thực tế sâu hơn
                        quality_status="PASS", notes="Bê tông C30 đạt độ sụt 18±2cm"
                    )
                ],
                incidents_or_delays=["Gặp hang Karst nhỏ tại cọc T1-03, phải xử lý trám xi măng mất 1.5 ngày"]
            )
        ]
        return sample_logs

    def _reconcile_quantities(
        self, logs: List[DailySiteLogEntry], bus: StateBus
    ) -> tuple[Dict[str, float], List[AsBuiltVarianceItem]]:
        """Đối chiếu khối lượng thực hiện so với hợp đồng/thiết kế."""
        # Tổng hợp khối lượng thực tế lũy kế theo task_id
        actual_progress: Dict[str, float] = {}
        for log in logs:
            for t in log.tasks_executed:
                actual_progress[t.task_id] = actual_progress.get(t.task_id, 0.0) + t.actual_volume

        # Bảng đơn giá định mức tham chiếu (VNĐ)
        unit_prices = {
            "T01": 25_000_000,     # Định vị tim mốc (VNĐ/điểm)
            "T02": 1_200_000,      # Đường công vụ (VNĐ/m)
            "T03": 12_500_000,     # Khoan dò Karst (VNĐ/lỗ)
            "T04": 6_850_000,      # Cọc khoan nhồi Ø1200 (VNĐ/m)
            "T05": 3_450_000,      # Bê tông bệ trụ C30 (VNĐ/m3)
        }

        # Kế hoạch thiết kế tham chiếu
        planned_contract = {
            "T01": {"desc": "Tim mốc định vị", "unit": "Điểm", "vol": 4.0},
            "T02": {"desc": "Đường công vụ tiếp cận", "unit": "m", "vol": 350.0},
            "T03": {"desc": "Khoan dò Karst đáy móng cọc", "unit": "Lỗ", "vol": 26.0},
            "T04": {"desc": "Khoan và đổ bê tông cọc nhồi D1200", "unit": "m", "vol": 320.0},
            "T05": {"desc": "Bê tông bệ trụ T1, T2", "unit": "m3", "vol": 225.0},
        }

        variance_items: List[AsBuiltVarianceItem] = []
        for tid, actual_vol in actual_progress.items():
            if tid in planned_contract:
                p = planned_contract[tid]
                contract_vol = p["vol"]
                delta_vol = round(actual_vol - contract_vol, 2)
                price = unit_prices.get(tid, 1_000_000)
                delta_amt = round(delta_vol * price)

                justification = "Khối lượng theo đúng hồ sơ thiết kế"
                category = "TRONG_HOP_DONG"
                if delta_vol > 0:
                    justification = f"Khối lượng phát sinh tăng thực tế tại hiện trường ({delta_vol:+} {p['unit']})"
                    category = "PHAT_SINH_NGOAI_HD"

                variance_items.append(AsBuiltVarianceItem(
                    task_id=tid,
                    component_id=f"COMP-{tid}",
                    description=p["desc"],
                    unit=p["unit"],
                    contract_volume=contract_vol,
                    as_built_volume=actual_vol,
                    delta_volume=delta_vol,
                    unit_price_vnd=price,
                    delta_amount_vnd=delta_amt,
                    appendix_03a_category=category,
                    justification=justification
                ))

        return actual_progress, variance_items

    def _update_cpm_schedule(self, actual_progress: Dict[str, float], bus: StateBus):
        """Tính toán lại CPM với khối lượng thực tế và thời gian thực hiện."""
        calc = CPMCalculator()

        # Danh mục công việc chuẩn của Cầu Km19+529.080
        base_tasks = [
            {"id": "T01", "name": "Tim mốc định vị", "duration": 3, "predecessors": [], "planned_volume": 4.0},
            {"id": "T02", "name": "Đường công vụ", "duration": 7, "predecessors": ["T01"], "planned_volume": 350.0},
            {"id": "T03", "name": "Khoan dò Karst", "duration": 14, "predecessors": ["T02"], "planned_volume": 26.0},
            {"id": "T04", "name": "Cọc nhồi T1/T2", "duration": 30, "predecessors": ["T03"], "planned_volume": 320.0},
            {"id": "T05", "name": "Bệ trụ T1/T2", "duration": 20, "predecessors": ["T04"], "planned_volume": 225.0},
            {"id": "T06", "name": "Thân đặc T1/T2", "duration": 25, "predecessors": ["T05"], "planned_volume": 131.6},
            {"id": "T07", "name": "Xà mũ T1/T2", "duration": 15, "predecessors": ["T06"], "planned_volume": 56.8},
            {"id": "T08", "name": "Lao dầm Super-T", "duration": 10, "predecessors": ["T07"], "planned_volume": 15.0},
            {"id": "T09", "name": "Mặt cầu C35", "duration": 30, "predecessors": ["T08"], "planned_volume": 305.5},
            {"id": "T10", "name": "Thảm BTN C16", "duration": 7, "predecessors": ["T09"], "planned_volume": 1450.0},
            {"id": "T11", "name": "Thử tải toàn cầu", "duration": 5, "predecessors": ["T10"], "planned_volume": 1.0},
        ]

        # Điền khối lượng thực tế và cập nhật thời gian
        for t in base_tasks:
            tid = t["id"]
            if tid in actual_progress:
                t["actual_volume"] = actual_progress[tid]
                # Nếu cọc nhồi phát sinh thêm địa chất đá sâu -> kéo dài thêm 2 ngày
                if tid == "T04":
                    t["duration"] = 32 # Delay 2 ngày
                    t["actual_finish_day"] = 56

        return calc.calculate(base_tasks, start_date_str="2026-10-01")
