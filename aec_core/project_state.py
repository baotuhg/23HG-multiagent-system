# -*- coding: utf-8 -*-
"""
PROJECT STATE MANAGER — bộ nhớ trạng thái dự án dùng chung (PROJECT_STATE.json).

Mặc định đọc/ghi data/PROJECT_STATE.json. load_state() KHÔNG tự tạo file; muốn khởi tạo bằng trạng thái
mẫu Cầu Km19+529.080 thì gọi get_default_state() hoặc chạy:
    python -m aec_core.project_state [đường_dẫn ...]
"""

import os
import json

DEFAULT_STATE_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "PROJECT_STATE.json")


class ProjectStateManager:
    def __init__(self, state_file_path=None):
        self.state_file_path = state_file_path or DEFAULT_STATE_FILE
        self.state = self.load_state()

    def get_default_state(self):
        return {
            "meta": {
                "version": "2.0.0",
                "system": "AEC Master Multi-Agent Architecture",
                "last_updated": "2026-09-25T19:45:00+07:00",
                "governance_standard": "Luat Xay dung 135/2025/QH15, ND 207/2026/ND-CP, TT 36/2026/TT-BXD"
            },
            "project_identity": {
                "project_name": "Cao tốc Tuyên Quang - Hà Giang (Giai đoạn 1) - Đoạn qua tỉnh Hà Giang",
                "package_name": "Phân đoạn Km19+120 -:- Km27+480",
                "structure_name": "CẦU KM19+529.080",
                "chainage_start": "Km19+453.880",
                "chainage_end": "Km19+604.280",
                "total_bridge_length_m": 130.40,
                "span_schema": "39.1m + 40.0m + 39.1m (Tổng nhịp 118.2m)",
                "design_speed_kmh": 100,
                "live_load": "HL-93 (TCVN 11823:2017)"
            },
            "geometry_and_structure": {
                "cross_section": {
                    "stage_1_lanes": 2,
                    "stage_1_width_m": 12.615,
                    "stage_2_lanes": 4,
                    "stage_2_width_m": 25.25
                },
                "superstructure": {
                    "girder_type": "Dầm Super-T BTCT DUL căng trước kết hợp đúc sẵn",
                    "girder_length_m": 38.20,
                    "girder_height_m": 1.75,
                    "girders_per_span": 5,
                    "total_girders": 15,
                    "concrete_grade": "C45/55",
                    "total_concrete_m3": 434.80,
                    "prestressing_strand": "15.2mm Grade 270 (ASTM A416)",
                    "total_strand_tons": 29.91,
                    "deck_slab_concrete_m3": 305.47,
                    "continuity_slab_concrete_m3": 21.87,
                    "precast_panels_count": 510,
                    "pot_bearings_count": 30,
                    "expansion_joints_m": 24.50
                },
                "substructure": {
                    "abutment_M1": {
                        "type": "Mố chân dê BTCT C30",
                        "piles_count": 3,
                        "pile_diameter_mm": 1200,
                        "pile_length_m": 20.0
                    },
                    "abutment_M2": {
                        "type": "Mố chữ U BTCT C30",
                        "piles_count": 7,
                        "pile_diameter_mm": 1200,
                        "pile_length_m": 36.0
                    },
                    "pier_T1": {
                        "type": "Trụ thân đặc BTCT C30 (H=14.5m)",
                        "piles_count": 8,
                        "pile_diameter_mm": 1200,
                        "pile_length_m": 40.0
                    },
                    "pier_T2": {
                        "type": "Trụ thân đặc BTCT C30 (H=11.4m)",
                        "piles_count": 8,
                        "pile_diameter_mm": 1200,
                        "pile_length_m": 30.0
                    },
                    "total_bored_piles": 26,
                    "total_bored_pile_meters": 872.0,
                    "karst_investigation_depth_m": 5.0
                }
            },
            "commercial_and_cost": {
                "direct_cost_T_vnd": 61280000000,
                "indirect_cost_GT_ratio": 0.073,
                "tax_TL_ratio": 0.055,
                "vat_ratio": 0.08,
                "cost_formula": "G_XD = (T + GT + TL) * 1.08",
                "payment_cycle_01_vnd": 32540000000
            },
            "schedule_control": {
                "start_date": "2026-10-01",
                "finish_date": "2027-03-28",
                "total_duration_days": 178,
                "critical_path_chain": "Tim mốc -> Đường công vụ -> Dò Karst -> Cọc T1/T2 -> Bệ trụ -> Thân đặc -> Xà mũ -> Lao dầm Super-T -> Mặt cầu C35 -> Thảm BTN C16 -> Thử tải"
            },
            "qaqc_compliance": {
                "total_inspection_records": 22,
                "hold_points": [
                    "Khoan dò Karst 5m vào đá liền khối",
                    "Siêu âm 156 mặt cắt cọc nhồi (100%)",
                    "Thí nghiệm nén động PDA trụ T1 (9434 kN)",
                    "Căng kéo cáp DUL dầm Super-T đạt R28",
                    "Thử tải tĩnh động toàn cầu"
                ],
                "date_cross_check_status": "PASSED (Zero conflicts)"
            }
        }

    def load_state(self):
        if os.path.exists(self.state_file_path):
            try:
                with open(self.state_file_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[!] Lỗi đọc state: {e}. Tạo mới.")
        return {}

    def save_state(self, state_data=None):
        if state_data is not None:
            self.state = state_data
        os.makedirs(os.path.dirname(os.path.abspath(self.state_file_path)), exist_ok=True)
        with open(self.state_file_path, "w", encoding="utf-8") as f:
            json.dump(self.state, f, ensure_ascii=False, indent=2)
        print(f"[*] Đã lưu trạng thái dự án: {self.state_file_path}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Ghi PROJECT_STATE.json mẫu (Cầu Km19+529.080) ra các đường dẫn chỉ định")
    parser.add_argument("targets", nargs="*", help=f"Đường dẫn cần ghi (mặc định: {DEFAULT_STATE_FILE})")
    args = parser.parse_args()
    mgr = ProjectStateManager()
    mgr.state = mgr.get_default_state()
    for target in args.targets or [mgr.state_file_path]:
        mgr.state_file_path = target
        mgr.save_state()
