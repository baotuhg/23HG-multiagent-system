# 🏗 KIẾN TRÚC STATE GRAPH v3.0 — 23HG MultiAgent System

> Tài liệu thiết kế kiến trúc sau Refactor từ Linear Pipeline → State Graph + Supervisor Pattern

---

## I. Sơ đồ kiến trúc tổng thể

```
┌─────────────────────────────────────────────────────────────────┐
│              BẢN VẼ DWG / HỒ SƠ THIẾT KẾ / YÊU CẦU KTXD        │
└───────────────────────────┬─────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│         AI SUPERVISOR — Chỉ huy trưởng ảo (supervisor_agent.py) │
│  • Tiếp nhận mục tiêu → phân chia đầu việc                      │
│  • Quản lý Shared State Bus (thread-safe RLock)                 │
│  • Chạy Quality Gate sau mỗi phase                              │
│  • Xử lý REJECT/RETRY (Inter-Agent Negotiation)                 │
│  • Kích hoạt Human Gate trước tài liệu pháp lý                  │
└──────┬────────┬────────┬────────┬────────┬──────────────────────┘
       │        │        │        │        │
       ▼        ▼        ▼        ▼        ▼
  [CAD]    [Rebar]   [QS]    [BPTC]  [Schedule]
  Agent    Agent     Agent   KCS     Agent
                             Agent

       ↑────────────── SHARED STATE BUS ──────────────────↑
                    (Single Source of Truth)
```

---

## II. State Graph — Các Phase

```
INIT
  │
  ▼ ─── cad_agent.execute()
CAD_TAKEOFF ──→ [Quality Gate 1: CAD Sanity Check]
  │               ✓ ≥1 cấu kiện, tổng BT > 0, WBS hợp lệ, không volume âm
  │               ✗ → retry cad_agent (1 lần)
  ▼
REBAR_CUT ──→ rebar_agent.execute()
  │             [OR-Tools CP-SAT 1D Cutting Stock]
  │             [SpliceZoneValidator — TCVN 5574:2018]
  │           → [Quality Gate 2: Rebar + Splice Check]
  │               ✓ status OPTIMAL/FEASIBLE, waste < 1.5%, splice PASS
  │               ✗ REJECT → Supervisor gửi REJECTED về rebar_agent → retry
  │               ✗ Hết retry → escalate lên Human Gate
  ▼
QS_ESTIMATE ──→ qs_agent.execute()
  │              [G_XD = (T + GT + TL) × 1.10 — VAT 10%]
  │            → [Quality Gate 3: Formula Verification]
  │               ✓ GT = T×7.3%, TL = (T+GT)×5.5%, VAT = subtotal×10%
  ▼
QAQC_REVIEW ──→ bptc_kcs_agent.execute()
  │              [AECAuditVerifier — yêu cầu 100/100]
  │              [Lab Result Link — R7/R28 bê tông, chứng chỉ thép]
  │            → [Quality Gate 4: Excel Audit 100/100]
  ▼
HUMAN_GATE ──→ [Hiển thị tất cả clash + anomaly]
  │             [Kỹ sư trưởng: Approve / Reject / Comment]
  │             ✓ APPROVED → tiếp tục
  │             ✗ REJECTED → dừng, yêu cầu xem xét lại
  ▼
SCHEDULE_CPM ──→ scheduler_agent.execute()
  │               [CPM Calculator: ES/EF/LS/LF/TF, đường găng]
  │               [Planned vs Actual — As-Built tracking]
  ▼
ASBUILT_LOOP ──→ asbuilt_agent.execute() [vòng lặp hàng ngày]
  │               [DailySiteLog: daily_site_log + as_built_quantity]
  │               [Actual vs Planned → Cập nhật CPM + Phụ lục 03a phát sinh]
  ▼
COMPLETED ✅
```

---

## III. Cấu trúc thư mục mới (song song với cấu trúc cũ)

```
23HG-multiagent-system/
│
├── run_state_graph.py          ← Entry Point mới (State Graph v3.0)
├── examples/run_pipeline.py    ← Entry Point cũ (giữ nguyên, backward compat)
│
├── core/                       ← MỚI: State Graph Architecture
│   ├── state/
│   │   ├── shared_state.py     # ProjectSharedState schema (dataclasses)
│   │   └── state_bus.py        # Thread-safe R/W bus (RLock)
│   ├── supervisor/
│   │   ├── supervisor_agent.py # AI Supervisor — điều phối State Graph
│   │   └── base_agent.py       # Abstract BaseAgent
│   ├── agents/
│   │   ├── rebar_agent.py      # Sub-Agent mẫu: OR-Tools + Splice Validator
│   │   └── sub_agents.py       # CADAgent, QSAgent, BPTCKCSAgent, SchedulerAgent
│   └── gates/
│       ├── quality_gate.py     # 4 Quality Gates (deterministic, zero LLM)
│       └── human_gate.py       # Human-in-the-loop (CLI/auto/file/callback)
│
├── tools/                      ← MỚI: Pure Python Tools (ZERO LLM)
│   ├── cutting_stock_solver.py # OR-Tools CP-SAT 1D Cutting Stock
│   ├── cpm_calculator.py       # CPM Critical Path Calculator
│   └── cad_diff_engine.py      # CAD Rev01 vs Rev02 Incremental Diff
│
├── schemas/                    ← MỚI: Schema definitions
│   └── __init__.py
│
├── agents/                     ← CŨ: Giữ nguyên (backward compat)
│   ├── aec_cad_extractor.py
│   ├── aec_data_aggregator.py
│   ├── project_state_manager.py
│   └── PROJECT_STATE.json
│
├── aec_core/                   ← CŨ: Giữ nguyên
│   ├── audit_verifier.py       # AECAuditVerifier (100/100)
│   └── project_state.py
│
├── templates/
│   └── Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx  # Excel 14 Sheet Master
│
└── .aec_state/RUNTIME_STATE.json   ← State Graph runtime persistence (không đưa vào git; đổi bằng AEC_STATE_DIR)
```

---

## IV. SharedState Schema — 8 Domain

| Domain | Dataclass | Nội dung |
|---|---|---|
| Meta | `ProjectSharedState` | session_id, current_phase, node_status map |
| CAD | `CADTakeoffData` | concrete_components, total_m3, formwork_m2, excavation, WBS mapping |
| Rebar | `RebarData` + `CuttingStockResult` | BBS items, OR-Tools cutting patterns, splice violations |
| QS | `QSData` | T, GT, TL, VAT, G_XD, Phụ lục 03a phát sinh |
| QA/QC | `QAQCData` | audit_score, lab results R7/R28, hold points, clashes |
| Schedule | `ScheduleData` | CPM tasks, critical path, planned vs actual % |
| As-Built | `DailySiteLog[]` | Nhật ký hiện trường hàng ngày |
| Approvals | `ApprovalGate[]` | Human-in-the-loop PENDING/APPROVED/REJECTED |

---

## V. Nguyên tắc thiết kế bất biến

> [!IMPORTANT]
> **TÁCH BIỆT TUYỆT ĐỐI: LLM ≠ Tính toán số học**
> - LLM: Intent recognition, Tool Calling, trình bày ngôn ngữ tự nhiên
> - Code thuần: Mọi phép toán (cộng, nhân, diện tích, khối lượng, tối ưu hóa)

> [!NOTE]
> **Inter-Agent Negotiation — REJECT/RETRY Pattern**
> - BPTC_KCS Agent kiểm tra vùng nối thép của Rebar Agent (TCVN 5574:2018 §8.6)
> - Nếu vi phạm → Supervisor gửi `REJECTED` → Rebar Agent chạy lại Solver
> - Hết `max_retries` → tự động escalate lên Human Gate

> [!TIP]
> **Backward Compatibility**
> - `examples/run_pipeline.py` vẫn hoạt động 100% — không bị sửa đổi
> - `run_state_graph.py` là layer mới chạy song song

---

## VI. Lệnh thực thi

```bash
# Chạy toàn bộ pipeline (auto-approve human gate — CI/CD)
python run_state_graph.py

# Chạy chọn phase
python run_state_graph.py --phase cad rebar qs qaqc schedule

# Bật chế độ CLI approval (production — KS trưởng tự phê duyệt)
python run_state_graph.py --human-gate cli

# Chỉ test OR-Tools + CPM Calculator
python run_state_graph.py --solver-test

# Pipeline cũ (backward compat — không thay đổi)
python examples/run_pipeline.py
```

---

## VII. Inter-Agent Negotiation — Luồng phản biện chéo

```
Supervisor
    │─ gọi ──→ rebar_agent.execute()
    │              │ Gọi CuttingStockSolver (OR-Tools)
    │              │ Gọi SpliceZoneValidator
    │              └─ ghi kết quả vào StateBus
    │
    │─ đọc ←── StateBus.get_rebar_data()
    │              .splice_zone_check = "REJECT"
    │              .splice_violations = ["Thanh T1: nối tại 45% nhịp — vùng kéo!"]
    │
    │─ ra quyết định:
    │     if REJECT and attempt < max_retries:
    │         → update_node_status("rebar_agent", REJECTED)
    │         → rebar_agent.execute() lần 2 (solver tìm sơ đồ cắt khác)
    │     elif hết retry:
    │         → create_approval_gate("REBAR-ESCALATE")  ← escalate lên KS trưởng
```

---

*Phiên bản: 3.0.0 — 2026-09-26*
*Tiêu chuẩn áp dụng: Luật XD 135/2025/QH15, NĐ 207/2026/NĐ-CP, TCVN 5574:2018, TCVN 1651:2018*
