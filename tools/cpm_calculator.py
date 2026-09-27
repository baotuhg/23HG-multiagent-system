# -*- coding: utf-8 -*-
"""
CPM CALCULATOR — Critical Path Method (Phương pháp đường găng)
Pure Python, Zero LLM — tính toán xác định 100% bằng code.

Tính:
  - Early Start (ES), Early Finish (EF) — Forward Pass
  - Late Start (LS), Late Finish (LF) — Backward Pass
  - Total Float (TF) = LS - ES
  - Critical Path (TF = 0)
  - Quy đổi ngày làm việc → ngày lịch (bỏ ngày nghỉ trong tuần và ngày lễ)
  - So sánh Planned vs Actual — As-Built Loop

Quan hệ logic (như MS Project), độ trễ (lag) tính bằng ngày làm việc, có thể âm:
  FS: sau bắt đầu ≥ trước kết thúc + lag     SS: sau bắt đầu ≥ trước bắt đầu + lag
  FF: sau kết thúc ≥ trước kết thúc + lag     SF: sau kết thúc ≥ trước bắt đầu + lag

Input: tasks = [{"id":..., "name":..., "duration":..., "predecessors":[...]}]
  predecessors: ["T01"] (FS), ["1.3SS+3d"], hoặc [{"id": "1.3", "type": "SS", "lag": 3}]
Output: CPMResult
"""

from __future__ import annotations
import math
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Any, Dict, Iterable, List, Optional, Set

EPS = 1e-9
LINK_TYPES = ("FS", "SS", "FF", "SF")

_PRED_RE = re.compile(
    r"^\s*(?P<id>.+?)\s*(?P<type>FS|SS|FF|SF)?\s*"
    r"(?:(?P<lag>[+-]\s*\d+(?:[.,]\d+)?)\s*(?:d|day|days|ngay|ngày)?)?\s*$",
    re.IGNORECASE,
)


@dataclass
class CPMLink:
    pred_id: str
    type: str = "FS"
    lag_days: float = 0.0


@dataclass
class CPMTask:
    task_id: str
    name: str
    duration_days: float
    predecessors: List[str] = field(default_factory=list)   # chỉ ID (tương thích cũ)
    links: List[CPMLink] = field(default_factory=list)
    # Calculated (đơn vị: ngày làm việc tính từ ngày khởi công)
    es: float = 0    # Early Start
    ef: float = 0    # Early Finish
    ls: float = 0    # Late Start
    lf: float = 0    # Late Finish
    tf: float = 0    # Total Float
    is_critical: bool = False
    start_date: str = ""
    finish_date: str = ""
    # As-Built
    planned_volume: float = 0.0
    actual_volume: float = 0.0
    actual_start_day: Optional[int] = None
    actual_finish_day: Optional[int] = None


@dataclass
class CPMResult:
    tasks: List[CPMTask] = field(default_factory=list)
    critical_path: List[str] = field(default_factory=list)
    total_duration_days: float = 0
    project_start: str = ""
    project_finish: str = ""
    overall_progress_pct: float = 0.0
    delay_days: float = 0
    status: str = "OK"
    warnings: List[str] = field(default_factory=list)


def parse_predecessor(item: Any) -> CPMLink:
    """'1.3SS+3d' / 'T01' / {'id','type','lag'} → CPMLink. Sai cú pháp → ValueError."""
    if isinstance(item, CPMLink):
        return item
    if isinstance(item, dict):
        link_type = str(item.get("type", "FS")).upper()
        if link_type not in LINK_TYPES:
            raise ValueError(f"Loại quan hệ không hợp lệ: {item}")
        return CPMLink(str(item["id"]), link_type, float(item.get("lag", 0) or 0))
    m = _PRED_RE.match(str(item))
    if not m or not m.group("id").strip():
        raise ValueError(f"Không đọc được quan hệ logic '{item}'")
    lag = m.group("lag")
    return CPMLink(
        pred_id=m.group("id").strip(),
        type=(m.group("type") or "FS").upper(),
        lag_days=float(lag.replace(" ", "").replace(",", ".")) if lag else 0.0,
    )


class WorkCalendar:
    """Lịch làm việc: bỏ các thứ nghỉ trong tuần (0=Thứ 2 … 6=Chủ nhật) và ngày lễ."""

    def __init__(self, start: date, non_working_weekdays: Iterable[int] = (),
                 holidays: Iterable[date] = ()):
        self.non_working = set(non_working_weekdays)
        if len(self.non_working) >= 7:
            raise ValueError("Lịch không có ngày làm việc nào")
        self.holidays = set(holidays)
        self._days: List[date] = []
        d = start
        while not self._is_working(d):
            d += timedelta(days=1)
        self._next = d

    def _is_working(self, d: date) -> bool:
        return d.weekday() not in self.non_working and d not in self.holidays

    def date_at(self, index: int) -> date:
        """Ngày lịch của ngày làm việc thứ `index` (0 = ngày khởi công)."""
        while len(self._days) <= index:
            d = self._next
            while not self._is_working(d):
                d += timedelta(days=1)
            self._days.append(d)
            self._next = d + timedelta(days=1)
        return self._days[index]


class CPMCalculator:
    """
    Tính Critical Path Method theo thuật toán Topological Sort + Forward/Backward Pass.
    """

    def calculate(
        self,
        tasks_input: List[Dict[str, Any]],
        start_date_str: str = "",
        non_working_weekdays: Iterable[int] = (),
        holidays: Iterable[date] = (),
    ) -> CPMResult:
        """
        tasks_input: [
            {"id": "T01", "name": "Tim mốc định vị", "duration": 3, "predecessors": []},
            {"id": "T02", "name": "Đường công vụ", "duration": 7, "predecessors": ["T01"]},
            {"id": "T03", "name": "Bãi đúc", "duration": 10, "predecessors": ["T02SS+3d"]},
            ...
        ]
        """
        result = CPMResult()

        # Build task dict
        task_dict: Dict[str, CPMTask] = {}
        errors: List[str] = []
        for t in tasks_input:
            tid = str(t["id"])
            if tid in task_dict:
                errors.append(f"Trùng mã công việc: {tid}")
                continue
            duration = float(t.get("duration", 1))
            if duration < 0:
                errors.append(f"{tid}: thời lượng âm ({duration})")
            links = []
            for p in t.get("predecessors", []) or []:
                try:
                    links.append(parse_predecessor(p))
                except (ValueError, KeyError) as e:
                    errors.append(f"{tid}: {e}")
            task_dict[tid] = CPMTask(
                task_id=tid,
                name=t.get("name", tid),
                duration_days=_clean(duration),
                predecessors=[l.pred_id for l in links],
                links=links,
                planned_volume=t.get("planned_volume", 0.0),
                actual_volume=t.get("actual_volume", 0.0),
                actual_start_day=t.get("actual_start_day"),
                actual_finish_day=t.get("actual_finish_day"),
            )

        if not task_dict:
            errors.append("Không có công việc nào được cung cấp")

        for t in task_dict.values():
            for l in t.links:
                if l.pred_id not in task_dict:
                    errors.append(f"{t.task_id}: công việc trước '{l.pred_id}' không tồn tại")
                elif l.pred_id == t.task_id:
                    errors.append(f"{t.task_id}: tự liên kết với chính nó")

        if errors:
            result.status = "ERROR"
            result.warnings.extend(errors)
            return result

        # Topological Sort (Kahn's algorithm)
        order = self._topological_sort(task_dict)
        if order is None:
            result.status = "ERROR"
            result.warnings.append("Phát hiện vòng lặp (circular dependency) trong biểu đồ CPM!")
            return result

        successors: Dict[str, List[tuple]] = {tid: [] for tid in task_dict}
        for t in task_dict.values():
            for l in t.links:
                successors[l.pred_id].append((t, l))

        # Forward Pass — tính ES, EF
        for tid in order:
            t = task_dict[tid]
            es = 0.0
            for l in t.links:
                p = task_dict[l.pred_id]
                es = max(es, _earliest_start(l, p.es, p.ef, t.duration_days))
            t.es = es
            t.ef = es + t.duration_days

        # Project duration
        total_dur = max(t.ef for t in task_dict.values())
        result.total_duration_days = _clean(total_dur)

        # Backward Pass — tính LS, LF
        for tid in reversed(order):
            t = task_dict[tid]
            lf = total_dur
            for s, l in successors[tid]:
                lf = min(lf, _latest_finish(l, s.ls, s.lf, t.duration_days))
            t.lf = lf
            t.ls = lf - t.duration_days
            t.tf = t.ls - t.es
            t.is_critical = t.tf <= EPS

        for t in task_dict.values():
            for attr in ("es", "ef", "ls", "lf", "tf"):
                setattr(t, attr, _clean(getattr(t, attr)))

        # Critical Path
        result.critical_path = [tid for tid in order if task_dict[tid].is_critical]
        result.tasks = [task_dict[tid] for tid in order]

        # Date mapping
        if start_date_str:
            try:
                start = date.fromisoformat(start_date_str)
                cal = WorkCalendar(start, non_working_weekdays, holidays)
                for t in result.tasks:
                    s_idx = math.floor(t.es + EPS)
                    f_idx = max(math.ceil(t.ef - EPS) - 1, s_idx)
                    t.start_date = cal.date_at(s_idx).isoformat()
                    t.finish_date = cal.date_at(f_idx).isoformat()
                result.project_start = cal.date_at(0).isoformat()
                result.project_finish = max(t.finish_date for t in result.tasks)
            except ValueError as e:
                result.warnings.append(f"Không thể quy đổi ngày (ngày bắt đầu '{start_date_str}'): {e}")

        # As-Built: tính % hoàn thành và delay
        result = self._calculate_asbuilt(result, task_dict)

        return result

    def _topological_sort(self, task_dict: Dict[str, CPMTask]) -> Optional[List[str]]:
        """Kahn's algorithm — trả về None nếu có vòng. Giữ thứ tự nhập để dễ đọc."""
        position = {tid: i for i, tid in enumerate(task_dict)}
        in_degree: Dict[str, int] = {tid: len(t.links) for tid, t in task_dict.items()}
        succ: Dict[str, List[str]] = {tid: [] for tid in task_dict}
        for t in task_dict.values():
            for l in t.links:
                succ[l.pred_id].append(t.task_id)

        ready = sorted((tid for tid, d in in_degree.items() if d == 0), key=position.get)
        order: List[str] = []
        while ready:
            tid = ready.pop(0)
            order.append(tid)
            for s in succ[tid]:
                in_degree[s] -= 1
                if in_degree[s] == 0:
                    ready.append(s)
            ready.sort(key=position.get)

        if len(order) != len(task_dict):
            return None  # Circular dependency
        return order

    def _calculate_asbuilt(self, result: CPMResult, task_dict: Dict[str, CPMTask]) -> CPMResult:
        """So sánh Planned vs Actual để tính % hoàn thành và delay."""
        total_planned = sum(t.planned_volume for t in task_dict.values())
        total_actual = sum(t.actual_volume for t in task_dict.values())

        if total_planned > 0:
            result.overall_progress_pct = round(total_actual / total_planned * 100, 1)

        # Delay: dựa trên ngày thực tế của công việc găng cuối
        critical_tasks = [t for t in task_dict.values() if t.is_critical]
        if critical_tasks and any(t.actual_finish_day is not None for t in critical_tasks):
            last_critical = max(critical_tasks, key=lambda t: t.ef)
            if last_critical.actual_finish_day is not None:
                result.delay_days = max(0, last_critical.actual_finish_day - last_critical.ef)

        return result


def _earliest_start(link: CPMLink, pred_es: float, pred_ef: float, duration: float) -> float:
    if link.type == "FS":
        return pred_ef + link.lag_days
    if link.type == "SS":
        return pred_es + link.lag_days
    if link.type == "FF":
        return pred_ef + link.lag_days - duration
    return pred_es + link.lag_days - duration  # SF


def _latest_finish(link: CPMLink, succ_ls: float, succ_lf: float, duration: float) -> float:
    if link.type == "FS":
        return succ_ls - link.lag_days
    if link.type == "SS":
        return succ_ls - link.lag_days + duration
    if link.type == "FF":
        return succ_lf - link.lag_days
    return succ_lf - link.lag_days + duration  # SF


def _clean(x: float) -> float:
    """Làm tròn nhiễu dấu phẩy động; số nguyên trả về int."""
    r = round(x, 6)
    return int(r) if r == int(r) else r


# ─────────────────────────────────────────────────────────────────────────────
# QUICK TEST
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    tasks = [
        {"id": "T01", "name": "Tim mốc định vị", "duration": 3, "predecessors": []},
        {"id": "T02", "name": "Đường công vụ", "duration": 7, "predecessors": ["T01"]},
        {"id": "T03", "name": "Khoan dò Karst", "duration": 14, "predecessors": ["T02"]},
        {"id": "T04", "name": "Cọc nhồi T1/T2", "duration": 30, "predecessors": ["T03"]},
        {"id": "T05", "name": "Bệ trụ T1/T2", "duration": 20, "predecessors": ["T04"]},
        {"id": "T05b", "name": "Bãi đúc dầm", "duration": 12, "predecessors": ["T02SS+3d"]},
        {"id": "T06", "name": "Thân đặc T1/T2", "duration": 25, "predecessors": ["T05"]},
        {"id": "T07", "name": "Xà mũ T1/T2", "duration": 15, "predecessors": ["T06"]},
        {"id": "T08", "name": "Lao dầm Super-T", "duration": 10, "predecessors": ["T07", "T05b"]},
    ]

    calc = CPMCalculator()
    result = calc.calculate(tasks, start_date_str="2026-10-01", non_working_weekdays={6})

    print(f"\n{'='*70}")
    print(f"  CPM RESULT — nghỉ Chủ nhật")
    print(f"{'='*70}")
    print(f"  Tổng thời gian: {result.total_duration_days} ngày làm việc")
    print(f"  Khởi công → hoàn thành: {result.project_start} → {result.project_finish}")
    print(f"  Đường găng: {' → '.join(result.critical_path)}")
    print(f"\n  {'ID':<6} {'Tên':<22} {'ES':>4} {'EF':>4} {'TF':>4}  {'Bắt đầu':<10} {'Kết thúc':<10}")
    print(f"  {'─'*68}")
    for t in result.tasks:
        star = "★" if t.is_critical else " "
        print(f"  {t.task_id:<6} {t.name:<22} {t.es:>4} {t.ef:>4} {t.tf:>4}  "
              f"{t.start_date:<10} {t.finish_date:<10} {star}")
