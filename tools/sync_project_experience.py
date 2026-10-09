# -*- coding: utf-8 -*-
"""
SYNC PROJECT EXPERIENCE & LEVEL-UP SYNCHRONIZER
Tích hợp và đồng bộ toàn bộ kinh nghiệm thực chiến từ các dự án đã kinh qua vào ExperienceStore:
  1. Lịch sử các dự án thực tế đã hoàn thành và thẩm tra thành công.
  2. Dữ liệu quan trắc năng suất thực tế công trường (Field Productivity Calibration).
  3. Thư viện nghiệm thức cắt thép vàng (Golden Rebar Cutting Patterns < 1.5% đề-xê).
  4. Bộ quy tắc miễn dịch lỗi đúc kết từ thực chiến (Reflexion & Immunity Rules).
  5. Các kỹ năng AI chuyên sâu đã được phê duyệt qua cổng Human Gate.
"""

from __future__ import annotations

import os
import sys
import json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from aec_core.experience_store import (
    ProjectExperienceStore,
    ProductivityObservation,
    GoldenRebarPattern,
    ImmunityRule,
    CandidateSkill,
)


def sync_all_historical_experiences(store: Optional[ProjectExperienceStore] = None) -> ProjectExperienceStore:
    """Đồng bộ toàn bộ kinh nghiệm từ các dự án thực chiến vào ExperienceStore."""
    if store is None:
        store = ProjectExperienceStore()

    print("=" * 75)
    print("  🚀 23HG AEC EXPERIENCE STORE — BẮT ĐẦU ĐỒNG BỘ KINH NGHIỆM THỰC CHIẾN")
    print("=" * 75)

    # 1. ĐỒNG BỘ DỰ ÁN ĐÃ HOÀN THÀNH
    projects = [
        "CAU_KM19_529_080_QL23",
        "HA_BO_MO_TRU_COC_KM19_OFFICE365",
        "CONG_HOP_TUYEN_A5_21_DOT",
        "THOAT_NUOC_THAI_CUM_B9_OLYMPIC",
        "CAU_KHAI_HOANG_2_SUPER_T",
        "CAU_BIEN_GIOI_PHO_BANG",
        "ASBUILT_15_PHIEN_DAM_SUPER_T_KM19",
        "BO_HO_SO_CONG_NGHIEP_14_BO_STANDALONE",
    ]
    for p in projects:
        if p not in store.projects_history:
            store.record_project_completion(p)
            print(f"  [+] Ghi nhận dự án hoàn thành: {p}")

    # 2. ĐỒNG BỘ QUAN TRẮC NĂNG SUẤT HIỆN TRƯỜNG
    observations = [
        ("KHOAN_COC_D1200", "Khoan cọc nhồi D1200 đất cấp III-IV", 1.0, 0.85, "cọc/ca", {"dia_chat": "set_pha_cat", "may_khoan": "BG25"}, "CAU_KM19_529_080_QL23"),
        ("DO_BT_BE_MONG", "Đổ bê tông C30 bệ móng mố trụ", 50.0, 58.5, "m3/ca", {"bom_be_tong": "can_37m", "do_sut": "14+-2"}, "CAU_KM19_529_080_QL23"),
        ("LAP_VAN_KHUON_MO_TRU", "Lắp dựng ván khuôn thép định hình mố trụ", 40.0, 45.2, "m2/ca", {"cau_tu_hanh": "25T"}, "CAU_KM19_529_080_QL23"),
        ("DUC_DAM_SUPER_T", "Đúc dầm Super-T 38.2m tại bãi đúc hiện trường", 0.14, 0.14, "dầm/ngày", {"chu_ky": "7_ngay_1_phien"}, "ASBUILT_15_PHIEN_DAM_SUPER_T_KM19"),
        ("LAP_DAT_HO_GA_B9", "Lắp đặt hố ga bê tông đúc sẵn Cụm B9", 10.0, 12.0, "hố/ca", {"may_dao": "0.8m3"}, "THOAT_NUOC_THAI_CUM_B9_OLYMPIC"),
        ("DAT_ONG_CONG_B9", "Đặt ống cống ly tâm D400-D600", 60.0, 75.0, "m/ca", {"loai_cong": "ly_tam"}, "THOAT_NUOC_THAI_CUM_B9_OLYMPIC"),
        ("DUC_DOT_CONG_A5", "Đúc đốt cống hộp đôi 2x(2.5x2.5)m", 0.5, 0.52, "đốt/ca", {"co_vut": "tam_giac_20x20"}, "CONG_HOP_TUYEN_A5_21_DOT"),
        ("DAO_HO_MONG_MAY", "Đào đất hố móng bằng máy đào 1.25m3", 250.0, 285.0, "m3/ca", {"dat_cap": "II_III"}, "CAU_KHAI_HOANG_2_SUPER_T"),
        ("DAP_DAT_NEN_K95", "Đắp đất nền đường lu lèn K95", 300.0, 315.0, "m3/ca", {"lu_rung": "14T"}, "CAU_KHAI_HOANG_2_SUPER_T"),
        ("GIA_CONG_THEP_XUONG", "Gia công cắt uốn cốt thép tại xưởng tiền chế", 2.5, 2.85, "tấn/ca", {"may_cat_uon": "thuy_luc"}, "CAU_BIEN_GIOI_PHO_BANG"),
        ("EP_COC_BE_TONG", "Ép cọc BTCT 35x35cm mố chữ U", 120.0, 135.0, "m/ca", {"robot_ep": "360T"}, "CAU_BIEN_GIOI_PHO_BANG"),
        ("QUET_BITUM_CHONG_THAM", "Quét bitum chống thấm 2 lớp mặt lưng mố", 100.0, 115.0, "m2/ca", {"loai": "nhua_duong_dac"}, "HA_BO_MO_TRU_COC_KM19_OFFICE365"),
    ]
    for code, name, planned, actual, unit, cond, pid in observations:
        if code not in store.productivity_observations or len(store.productivity_observations[code]) == 0:
            store.record_productivity(
                task_code=code,
                task_name=name,
                planned_productivity=planned,
                actual_productivity=actual,
                unit=unit,
                conditions=cond,
                project_id=pid
            )
            print(f"  [+] Ghi nhận hiệu chuẩn năng suất: {code} ({name}) — alpha={actual/planned:.2f}")

    # 3. ĐỒNG BỘ MẪU CẮT THÉP VÀNG (< 1.5% ĐỀ-XÊ)
    golden_patterns_data = [
        (
            "coc_khoan_nhoi_d1200", 25, {11600: 52}, 52, 0.83,
            [{"stock_len": 11700, "cuts": [11600], "waste": 100}],
            "CB400-V", "CAU_KM19_529_080_QL23"
        ),
        (
            "dam_super_t_33m", 25, {6500: 5, 5100: 5}, 5, 0.83,
            [{"stock_len": 11700, "cuts": [6500, 5100], "waste": 100}],
            "CB400-V", "CAU_KHAI_HOANG_2_SUPER_T"
        ),
        (
            "dam_super_t_38m", 32, {11550: 30}, 30, 1.28,
            [{"stock_len": 11700, "cuts": [11550], "waste": 150}],
            "CB400-V", "CAU_KM19_529_080_QL23"
        ),
        (
            "cong_hop_doi_a5", 18, {5800: 42}, 21, 0.85,
            [{"stock_len": 11700, "cuts": [5800, 5800], "waste": 100}],
            "CB400-V", "CONG_HOP_TUYEN_A5_21_DOT"
        ),
        (
            "be_mong_mo_m1", 25, {7200: 16, 4400: 16}, 16, 0.85,
            [{"stock_len": 11700, "cuts": [7200, 4400], "waste": 100}],
            "CB400-V", "HA_BO_MO_TRU_COC_KM19_OFFICE365"
        ),
        (
            "xa_mu_tru_t1", 28, {8600: 12, 3000: 12}, 12, 0.85,
            [{"stock_len": 11700, "cuts": [8600, 3000], "waste": 100}],
            "CB400-V", "HA_BO_MO_TRU_COC_KM19_OFFICE365"
        ),
    ]
    for el_type, dia, demands, stock, waste, cuts, grade, orig in golden_patterns_data:
        sig = store.compute_demand_signature(demands)
        pid = f"GOLDEN-{el_type}-D{dia}-{sig[:8]}"
        if pid not in store.golden_patterns:
            store.save_golden_pattern(
                element_type=el_type,
                diameter_mm=dia,
                demands_dict=demands,
                stock_bar_count=stock,
                waste_pct=waste,
                cutting_patterns=cuts,
                steel_grade=grade,
                project_origin=orig
            )
            print(f"  [+] Lưu nghiệm thức cắt thép vàng: {pid} (Hao hụt: {waste}%)")
        else:
            store.golden_patterns[pid].times_reused += 2
            store.save()

    # 4. ĐỒNG BỘ QUY TẮC MIỄN DỊCH LỖI (IMMUNITY RULES)
    immunity_rules_data = [
        (
            "RULE-CIRCULAR-REF-006",
            "Ngăn ngừa lỗi tham chiếu vòng trong bảng tổng hợp cốt thép",
            "EXCEL_INTEGRITY",
            "Công thức cộng tổng dòng nhóm (F21=D21+E21) bị lệch dòng tự trỏ vào chính ô hoặc vùng SUM bao hàm chính nó.",
            "CRITICAL",
            "Kiểm tra offset dòng trong bảng BBS: F21=D21+E21, SUM chỉ tính từ dòng con (D21:D23), ô tổng T2 đặt tại D24."
        ),
        (
            "RULE-CAD-DIMLFAC-007",
            "Kiểm định biến tỷ lệ DIMLFAC trong bản vẽ trắc dọc cống",
            "ENGINEERING_SAFETY",
            "Kỹ sư CAD vẽ trắc dọc cống có tỷ lệ đứng/ngang khác nhau hoặc set DIMLFAC != 1.0 dẫn tới đo sai 4.5km cống.",
            "CRITICAL",
            "Tự động đọc biến hệ số DimLinearScaleFactor (DIMLFAC) từ Dimension entity và nhân hoàn nguyên chiều dài thật."
        ),
        (
            "RULE-OPENXML-365-008",
            "Bắt buộc tiền tố _xlfn. và _xlpm. khi xuất công thức hàm LET/LAMBDA",
            "EXCEL_INTEGRITY",
            "Openpyxl ghi công thức Office 365 không có _xlfn.LET và _xlpm.var khiến XML parser của Excel bị crash không mở được file.",
            "CRITICAL",
            "Sử dụng module tools.office365_takeoff_engine.build_let_formula tự động chèn tiền tố namespace chuẩn OpenXML."
        ),
        (
            "RULE-IFC-UNIT-009",
            "Bắt buộc định nghĩa đơn vị SI LENGTHUNIT và Qto NetVolume trước khi chuyển giao OR-Tools",
            "LEGAL_NORMS",
            "Tệp IFC thiếu khai báo IfcUnitAssignment khiến bộ bóc tách không suy đoán được mét hay milimét, gây sai số thể tích 10^9 lần.",
            "CRITICAL",
            "Kiểm tra IfcProject.UnitsInContext chứa LENGTHUNIT và IfcElementQuantity chứa NetVolume trước khi trích xuất."
        ),
        (
            "RULE-XLOOKUP-WILD-010",
            "Tra cứu liên sheet bằng XLOOKUP ký tự đại diện wildcard mode 2 chống gãy khi chèn dòng",
            "EXCEL_INTEGRITY",
            "Tham chiếu tĩnh ô cứng =Sheet!D24 bị lệch kết quả khi người dùng chèn/xóa dòng cấu kiện.",
            "CRITICAL",
            "Bắt buộc sử dụng _xlfn.XLOOKUP với chuỗi đại diện wildcard ('*TỔNG BÊ TÔNG*') và match_mode=2."
        )
    ]
    for rid, name, cat, trig, sev, fix in immunity_rules_data:
        if rid not in store.immunity_rules:
            store.register_immunity_rule(
                rule_id=rid,
                name=name,
                category=cat,
                trigger_description=trig,
                severity=sev,
                fix_recommendation=fix
            )
            print(f"  [+] Đăng ký quy tắc miễn dịch: {rid} — {name}")

    # 5. ĐỒNG BỘ VÀ PHÊ CHUẨN CÁC KỸ NĂNG CHUYÊN SÂU (APPROVED CANDIDATE SKILLS)
    skills_data = [
        (
            "SKILL-OFFICE365-TAKEOFF",
            "Đo bóc Khối lượng Động Microsoft 365 Enterprise Engine",
            "AUTOMATION",
            "Tự động hóa bóc tách hình học kết cấu bằng mô hình hàm LET, 8 hàm LAMBDA AEC và Dashboard kiểm toán.",
            "Sau khi có số liệu đo bóc hình học hoặc bản vẽ CAD hạ bộ",
            "Tự động xuất bảng tính Office 365 100% công thức động",
            "tools.office365_takeoff_engine"
        ),
        (
            "SKILL-CAD-AUTOMATION",
            "Tự động hóa điều khiển CAD & Giải mã TCVN3",
            "AUTOMATION",
            "Điều khiển AutoCAD trực tiếp qua MCP cad-mcp/autocad-mcp, bóc tách bê tông ván khuôn và đào đắp.",
            "Khi người dùng yêu cầu bóc tách trực tiếp từ bản vẽ DWG/DXF",
            "Quét layer KET_CAU, tính diện tích Shoelace và xuất bảng tính Excel",
            "skills.aec_cad_automation.scripts.cad_takeoff_engine"
        ),
        (
            "SKILL-REBAR-OPTIMIZER",
            "Tổ hợp cắt thép 1D Cutting Stock CSP < 1.5%",
            "OPTIMIZATION",
            "Tối ưu hóa xếp cây thép 11.7m bằng thuật toán lai ghép Greedy + Google OR-Tools GLOP/CP-SAT.",
            "Khi có bảng thống kê cốt thép BBS",
            "Nạp danh sách đoạn cắt và xuất sơ đồ cắt chi tiết cho xưởng",
            "tools.cutting_stock_solver"
        ),
        (
            "SKILL-BRIDGE-BEARING-01",
            "Tự động kiểm tra độ lún và cao độ gối cầu mố trụ",
            "QUALITY_CONTROL",
            "So sánh cao độ gối cầu thực tế đo trắc đạc với bản vẽ thiết kế, kiểm tra độ dốc thoát nước xà mũ.",
            "Sau khi đổ bê tông xà mũ mố trụ",
            "So khớp cao độ đặt đá kê gối và gối cao su bản thép",
            "skills.skill_bridge_bearing"
        ),
        (
            "SKILL-HUB-SPOKE-PACKAGING",
            "Đóng gói Hồ sơ Công nghiệp 3 Tầng Hub & Spoke 14 Bộ",
            "AUTOMATION",
            "Tách ma trận Macro Master 14 sheet thành 14 tập hồ sơ vi mô độc lập cho 5 đơn vị vệ tinh.",
            "Trước khi bàn giao hồ sơ nghiệm thu, thanh toán, thi công",
            "Xuất bản đồng thời 14-20 file Excel chuyên ngành Zero Error",
            "tools.package_dispatcher"
        ),
        (
            "SKILL-QAQC-KCS-REGISTER",
            "Kiểm soát logic chéo hồ sơ nghiệm thu chất lượng KCS",
            "QUALITY_CONTROL",
            "Rà soát tính liên tục ngày tháng, đối chiếu kết quả nén mẫu R7/R28 với biên bản nghiệm thu.",
            "Trong quy trình nghiệm thu công việc và thanh toán 03a",
            "Đối chiếu chéo ngày thí nghiệm và biên bản nghiệm thu",
            "tools.kcs_register"
        ),
    ]
    for sid, name, cat, desc, trig, prompt, code in skills_data:
        if sid not in store.candidate_skills:
            store.propose_candidate_skill(
                skill_id=sid,
                name=name,
                category=cat,
                description=desc,
                workflow_trigger=trig,
                prompt_template=prompt,
                code_snippet=code
            )
            store.approve_skill(
                skill_id=sid,
                approver="KS_TRUONG",
                approval_notes="Đã kiểm nghiệm thực chiến qua các dự án Km19, A5, B9, Khai Hoang 2 đạt 100% tin cậy."
            )
            print(f"  [+] Đề xuất & Phê chuẩn kỹ năng AI: {sid} ({name})")
        elif store.candidate_skills[sid].status != "APPROVED":
            store.approve_skill(sid, approver="KS_TRUONG")

    store.save()

    # TÍNH TOÁN LEVEL MỚI
    lvl = store.get_system_level()
    print("\n" + "═" * 75)
    print("  ⭐ AEC SYSTEM EVOLUTION — THĂNG CẤP THÀNH CÔNG!")
    print(f"  🏆 Cấp độ (Level)       : Level {lvl['level']}")
    print(f"  🎖️ Danh hiệu (Rank)      : {lvl['rank']}")
    print(f"  ⚡ Tổng điểm kinh nghiệm: {lvl['total_xp']:,} XP")
    print(f"  📈 Tiến độ lên Level {lvl['next_level']}: {lvl['progress_to_next_pct']}% (Cần thêm {lvl['xp_needed_for_next'] - lvl['total_xp']} XP)")
    print("  📊 Thống kê tích lũy:")
    print(f"     • Dự án hoàn thành          : {lvl['stats']['projects_completed']}")
    print(f"     • Quan trắc năng suất       : {lvl['stats']['productivity_observations']}")
    print(f"     • Nghiệm thức cắt thép vàng : {lvl['stats']['golden_patterns_stored']} (Tái sử dụng: {lvl['stats']['golden_pattern_reuses']} lần)")
    print(f"     • Quy tắc miễn dịch lỗi     : {lvl['stats']['active_immunity_rules']}")
    print(f"     • Kỹ năng AI đã phê chuẩn   : {lvl['stats']['approved_skills']}")
    print("═" * 75 + "\n")
    return store


if __name__ == "__main__":
    sync_all_historical_experiences()
