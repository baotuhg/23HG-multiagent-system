# -*- coding: utf-8 -*-
"""
Test cơ chế tự động phát hiện Sub-Agent (core/agents/registry.py).

Bảo đảm: thêm agent mới chỉ cần một module kèm factory @register_agent, không phải sửa
run_state_graph.py; và tất cả agent của pipeline đều được dựng đúng từ tham số dòng lệnh.
"""

import unittest

from core.agents.registry import (AgentSpec, build_agents, discover_agents,
                                   register_agent, registered_specs)
from core.supervisor.base_agent import BaseAgent
from run_state_graph import build_parser

EXPECTED = {"cad_agent", "rebar_agent", "qs_agent", "bptc_kcs_agent",
            "scheduler_agent", "asbuilt_agent", "payment_agent"}


class AgentRegistryTest(unittest.TestCase):
    def test_discovers_all_pipeline_agents(self):
        self.assertEqual(set(discover_agents()), EXPECTED)

    def test_specs_sorted_by_order(self):
        specs = registered_specs()
        self.assertEqual([s.agent_id for s in specs],
                         ["cad_agent", "rebar_agent", "qs_agent", "bptc_kcs_agent",
                          "scheduler_agent", "asbuilt_agent", "payment_agent"])
        self.assertTrue(all(isinstance(s, AgentSpec) for s in specs))

    def test_build_agents_from_cli_defaults(self):
        args = build_parser().parse_args([])
        agents = build_agents(args)
        self.assertEqual([a.agent_id for a in agents],
                         [s.agent_id for s in registered_specs()])
        self.assertTrue(all(isinstance(a, BaseAgent) for a in agents))
        # mỗi agent_id là duy nhất
        self.assertEqual(len(agents), len(EXPECTED))

    def test_duplicate_registration_rejected(self):
        """Đăng ký agent_id đã tồn tại với factory khác → báo lỗi, không ghi đè ngầm."""
        with self.assertRaises(ValueError):
            @register_agent("qs_agent")       # đã do sub_agents đăng ký với factory khác
            def _dup(args):
                return None

    def test_same_factory_reregister_is_idempotent(self):
        before = dict((s.agent_id, s.factory) for s in registered_specs())
        factory = before["payment_agent"]
        # đăng ký lại đúng factory cũ → không lỗi
        register_agent("payment_agent", order=70)(factory)
        self.assertIs(dict((s.agent_id, s.factory) for s in registered_specs())["payment_agent"], factory)


if __name__ == "__main__":
    unittest.main()
