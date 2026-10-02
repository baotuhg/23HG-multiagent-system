# -*- coding: utf-8 -*-
"""
AGENT REGISTRY — tự động phát hiện Sub-Agent trong core/agents/

Thêm một Sub-Agent mới KHÔNG cần sửa run_state_graph.py: chỉ cần, trong module agent của mình,
khai báo một factory nhận `args` (Namespace dòng lệnh) và trả về một BaseAgent, gắn decorator:

    from core.agents.registry import register_agent

    @register_agent("my_agent", order=80)
    def _make_my_agent(args):
        return MyAgent(path=args.my_path)

`discover_agents()` nạp mọi module trong core/agents/ (kích hoạt các decorator), `build_agents(args)`
dựng tất cả agent đã đăng ký theo thứ tự `order`. Thuần Python, không phụ thuộc ngoài.
"""

from __future__ import annotations

import importlib
import pkgutil
from typing import Callable, Dict, List, NamedTuple

from core.supervisor.base_agent import BaseAgent

_PACKAGE = "core.agents"


class AgentSpec(NamedTuple):
    agent_id: str
    order: int                       # thứ tự dựng/đăng ký (nhỏ chạy trước); không thay thế thứ tự phase
    factory: Callable[[object], BaseAgent]
    module: str


_REGISTRY: Dict[str, AgentSpec] = {}
_discovered = False


def register_agent(agent_id: str, order: int = 100) -> Callable[[Callable], Callable]:
    """Decorator đánh dấu một factory `(args) -> BaseAgent` là một Sub-Agent của hệ thống."""
    def decorator(factory: Callable[[object], BaseAgent]) -> Callable[[object], BaseAgent]:
        existing = _REGISTRY.get(agent_id)
        if existing is not None and existing.factory is not factory:
            raise ValueError(
                f"Agent '{agent_id}' đã được đăng ký ở module '{existing.module}', "
                f"không thể đăng ký lại ở '{factory.__module__}'")
        _REGISTRY[agent_id] = AgentSpec(agent_id, order, factory, factory.__module__)
        return factory
    return decorator


def discover_agents(package: str = _PACKAGE, force: bool = False) -> List[str]:
    """Nạp mọi module con của `package` để kích hoạt các @register_agent. Trả danh sách agent_id."""
    global _discovered
    if _discovered and not force:
        return sorted(_REGISTRY)
    pkg = importlib.import_module(package)
    for mod in pkgutil.iter_modules(pkg.__path__):
        if mod.name.startswith("_") or mod.name == "registry":
            continue
        try:
            importlib.import_module(f"{package}.{mod.name}")
        except ImportError as e:        # agent tùy chọn thiếu thư viện → bỏ qua, không làm hỏng cả hệ thống
            print(f"  [registry] Bỏ qua agent trong '{mod.name}': thiếu thư viện ({e})")
    _discovered = True
    return sorted(_REGISTRY)


def registered_specs() -> List[AgentSpec]:
    """Các agent đã đăng ký, sắp theo (order, agent_id)."""
    discover_agents()
    return sorted(_REGISTRY.values(), key=lambda s: (s.order, s.agent_id))


def build_agents(args) -> List[BaseAgent]:
    """Dựng tất cả Sub-Agent đã đăng ký từ tham số dòng lệnh, theo thứ tự order."""
    return [spec.factory(args) for spec in registered_specs()]
