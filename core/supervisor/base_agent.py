# -*- coding: utf-8 -*-
"""
BASE AGENT — Lớp cơ sở cho tất cả Sub-Agent trong State Graph
Sub-Agent KHÔNG giao tiếp trực tiếp với nhau.
Sub-Agent chỉ: đọc context từ StateBus + gọi Tools + ghi kết quả vào StateBus.

Pattern sử dụng:
    class MyAgent(BaseAgent):
        def run(self, bus: StateBus) -> bool:
            ctx = bus.get_state_snapshot()
            # ... xử lý ...
            bus.set_xxx_data({...})
            return True
"""

from __future__ import annotations
import sys
import os
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any, Dict, Optional

# Thêm project root vào path
_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.state.state_bus import StateBus
from core.state.shared_state import NodeStatus


class DataInputError(Exception):
    """
    Lỗi dữ liệu đầu vào — chạy lại cũng không khắc phục được.
    Supervisor KHÔNG retry khi gặp lỗi này mà dừng và báo người dùng sửa dữ liệu.
    """


class MissingDataError(DataInputError):
    """
    Thiếu dữ liệu thật của dự án. Không được tự thay bằng dữ liệu mẫu
    trừ khi chạy chế độ demo (--demo).
    """


class BaseAgent(ABC):
    """
    Lớp cơ sở trừu tượng cho mọi Sub-Agent.
    Mỗi Sub-Agent implement phương thức `run()`.
    """

    def __init__(self, agent_id: str, description: str = ""):
        self.agent_id = agent_id
        self.description = description
        self.last_error_fatal = False  # True nếu lần chạy gần nhất lỗi dữ liệu đầu vào

    @abstractmethod
    def run(self, bus: StateBus) -> bool:
        """
        Logic nghiệp vụ chính.
        Returns True nếu thành công, False nếu cần retry.
        Raise Exception nếu lỗi không thể khắc phục.
        """
        ...

    def execute(self, bus: StateBus) -> bool:
        """
        Wrapper an toàn cho run() — ghi status vào StateBus, bắt exception.
        Supervisor gọi execute(), không gọi run() trực tiếp.
        """
        self.last_error_fatal = False
        bus.update_node_status(self.agent_id, NodeStatus.RUNNING,
                               output_summary=f"Bắt đầu {self.agent_id}")
        try:
            success = self.run(bus)
            if success:
                bus.update_node_status(self.agent_id, NodeStatus.SUCCESS,
                                       output_summary="Hoàn thành thành công")
            else:
                bus.update_node_status(self.agent_id, NodeStatus.FAILED,
                                       error_message="run() trả về False — cần retry")
            return success
        except DataInputError as e:
            self.last_error_fatal = True
            label = "THIẾU DỮ LIỆU THẬT" if isinstance(e, MissingDataError) else "LỖI DỮ LIỆU ĐẦU VÀO"
            bus.update_node_status(self.agent_id, NodeStatus.FAILED, error_message=f"{label}: {e}")
            bus.push_error(f"[{self.agent_id}] {label}: {e}")
            return False
        except Exception as e:
            msg = f"{type(e).__name__}: {e}"
            bus.update_node_status(self.agent_id, NodeStatus.FAILED, error_message=msg)
            bus.push_error(f"[{self.agent_id}] {msg}")
            return False
