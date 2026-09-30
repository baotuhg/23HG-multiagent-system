# -*- coding: utf-8 -*-
"""
MÔ-ĐUN AEC-CORE: PHÂN TÍCH ĐỊNH MỨC CẤP PHỐI 1M3 VÀ MA TRẬN TẦN SUẤT THÍ NGHIỆM VẬT LIỆU (QA/QC)
Căn cứ pháp lý & tiêu chuẩn áp dụng:
- Luật Xây dựng số 135/2025/QH15
- Nghị định số 207/2026/NĐ-CP về Quản lý chất lượng và thi công xây dựng
- Thông tư 38/2026/TT-BXD về Định mức dự toán xây dựng công trình
- TCVN 4453:1995: Kết cấu bê tông và bê tông cốt thép toàn khối - Quy phạm thi công và nghiệm thu
- TCVN 1651:2018: Thép cốt bê tông (Thép tròn trơn & thép thanh vằn)
- TCVN 6260:2020: Xi măng poóc lăng hỗn hợp
- TCVN 7570:2006: Cốt liệu cho bê tông và vữa
- TCVN 9396:2012 / ASTM D6760: Thử nghiệm không phá hủy - Phương pháp xung siêu âm qua cọc khoan nhồi
- ASTM D4945 / TCVN 11321:2016: Thí nghiệm thử tải động biến dạng lớn cọc khoan nhồi (PDA)
"""

from typing import Dict, List, Any
import math

class ConcreteMixDesigner:
    """Quản lý định mức cấp phối vật liệu cho 1 m3 bê tông từng kết cấu."""
    
    # Định mức chuẩn cho 1 m3 bê tông theo TCVN & TT 38/2026/TT-BXD
    MIX_STANDARDS = {
        "C10": {
            "name": "Bê tông đệm móng mố, bệ trụ, bản quá độ (Mác 150)",
            "slump": "6 ± 2 cm",
            "cement_kg": 225.0,
            "sand_m3": 0.510,
            "stone_m3": 0.880,
            "water_l": 185.0,
            "admixture_l": 0.0,
            "standard": "TCVN 4453:1995"
        },
        "C25": {
            "name": "Bản quá độ sau mố, gờ lan can, tấm đúc sẵn (Mác 300)",
            "slump": "10 ± 2 cm",
            "cement_kg": 350.0,
            "sand_m3": 0.460,
            "stone_m3": 0.860,
            "water_l": 175.0,
            "admixture_l": 2.8,
            "standard": "TCVN 4453:1995 / ASTM C494"
        },
        "C30_PILE": {
            "name": "Bê tông cọc khoan nhồi D1200 thi công ống Tremie (Mác 350)",
            "slump": "18 ± 2 cm",
            "cement_kg": 410.0,
            "sand_m3": 0.470,
            "stone_m3": 0.840,
            "water_l": 180.0,
            "admixture_l": 4.1,
            "standard": "TCVN 9395:2012 / ASTM C494 Type G"
        },
        "C30_SUB": {
            "name": "Bê tông bệ mố, thân mố, tường cánh, bệ trụ, thân trụ (Mác 350)",
            "slump": "12 ± 2 cm",
            "cement_kg": 385.0,
            "sand_m3": 0.450,
            "stone_m3": 0.850,
            "water_l": 175.0,
            "admixture_l": 3.5,
            "standard": "TCVN 4453:1995"
        },
        "C35_SUPER": {
            "name": "Xà mũ trụ, bản mặt cầu, dầm ngang, liên tục nhiệt (Mác 400)",
            "slump": "14 ± 2 cm",
            "cement_kg": 430.0,
            "sand_m3": 0.440,
            "stone_m3": 0.850,
            "water_l": 165.0,
            "admixture_l": 4.3,
            "standard": "TCVN 4453:1995 / HRWR"
        },
        "C45_GIRDER": {
            "name": "Bê tông dầm Super-T DƯL đúc sẵn L=38.2m (Mác 550)",
            "slump": "16 ± 2 cm",
            "cement_kg": 475.0,
            "sand_m3": 0.420,
            "stone_m3": 0.830,
            "water_l": 150.0,
            "admixture_l": 5.7,
            "standard": "AASHTO LRFD / TCVN 11823 / R kéo >= 85%"
        },
        "C40_NON_SHRINK": {
            "name": "Bê tông hạt nhỏ không co ngót chèn khe co giãn (Mác 450)",
            "slump": "12 ± 2 cm",
            "cement_kg": 450.0,
            "sand_m3": 0.460,
            "stone_m3": 0.800,
            "water_l": 160.0,
            "admixture_l": 4.5,
            "standard": "TCVN 9204:2012 / Phụ gia bù co ngót"
        }
    }

    @classmethod
    def calculate_total_materials(cls, concrete_volumes: Dict[str, float]) -> Dict[str, Any]:
        """Tính toán tổng nhu cầu vật tư cấu thành cho toàn bộ thể tích bê tông."""
        totals = {
            "cement_ton": 0.0,
            "sand_m3": 0.0,
            "stone_m3": 0.0,
            "water_m3": 0.0,
            "admixture_liter": 0.0,
            "breakdown": []
        }
        
        for key, vol in concrete_volumes.items():
            if key in cls.MIX_STANDARDS and vol > 0:
                spec = cls.MIX_STANDARDS[key]
                c_ton = (vol * spec["cement_kg"]) / 1000.0
                s_m3 = vol * spec["sand_m3"]
                st_m3 = vol * spec["stone_m3"]
                w_m3 = (vol * spec["water_l"]) / 1000.0
                adm_l = vol * spec["admixture_l"]
                
                totals["cement_ton"] += c_ton
                totals["sand_m3"] += s_m3
                totals["stone_m3"] += st_m3
                totals["water_m3"] += w_m3
                totals["admixture_liter"] += adm_l
                
                totals["breakdown"].append({
                    "type": key,
                    "name": spec["name"],
                    "volume_m3": vol,
                    "cement_ton": round(c_ton, 2),
                    "sand_m3": round(s_m3, 2),
                    "stone_m3": round(st_m3, 2),
                    "water_m3": round(w_m3, 2),
                    "admixture_liter": round(adm_l, 1)
                })
        
        totals["cement_ton"] = round(totals["cement_ton"], 2)
        totals["sand_m3"] = round(totals["sand_m3"], 2)
        totals["stone_m3"] = round(totals["stone_m3"], 2)
        totals["water_m3"] = round(totals["water_m3"], 2)
        totals["admixture_liter"] = round(totals["admixture_liter"], 1)
        return totals


class QualityTestingFrequencyMatrix:
    """Quản lý ma trận tần suất kiểm tra thí nghiệm vật liệu và nghiệm thu KCS."""
    
    @staticmethod
    def calculate_testing_plan(materials_summary: Dict[str, Any], structure_counts: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Tính toán số lượng tổ mẫu thí nghiệm bắt buộc theo quy chuẩn hiện hành."""
        plan = []
        
        # 1. Cốt thép xây dựng theo TCVN 1651:2018 (<= 50 tấn / tổ mẫu / đường kính)
        steel_by_dia = materials_summary.get("steel_by_dia", {})
        for dia, tonnage in steel_by_dia.items():
            tests_req = math.ceil(tonnage / 50.0) if tonnage > 0 else 1
            plan.append({
                "category": "Cốt thép xây dựng",
                "item": f"Thép cốt bê tông Ø{dia}mm",
                "standard": "TCVN 1651:2018",
                "volume": tonnage,
                "unit": "Tấn",
                "frequency_rule": "<= 50 Tấn / lô / đường kính",
                "required_tests": tests_req,
                "sample_spec": "3 thanh kéo + 3 thanh uốn",
                "hold_point": "Có KQTN đạt trước khi cắt uốn"
            })
            
        # 2. Cáp DƯL theo ASTM A416
        cable_ton = materials_summary.get("cable_ton", 29.91)
        plan.append({
            "category": "Cáp dự ứng lực",
            "item": "Cáp DƯL 7 sợi Ø15.2mm",
            "standard": "ASTM A416 Gr270",
            "volume": cable_ton,
            "unit": "Tấn",
            "frequency_rule": "<= 30-50 Tấn / lô",
            "required_tests": math.ceil(cable_ton / 30.0),
            "sample_spec": "3 tao cáp L=1.5m",
            "hold_point": "Nghiệm thu trước khi xỏ cáp dầm"
        })

        # 3. Xi măng PCB40 (<= 100-200 tấn / tổ mẫu)
        cement_ton = materials_summary.get("cement_ton", 1406.26)
        plan.append({
            "category": "Xi măng",
            "item": "Xi măng PCB40",
            "standard": "TCVN 6260:2020",
            "volume": cement_ton,
            "unit": "Tấn",
            "frequency_rule": "<= 100 Tấn / lô",
            "required_tests": math.ceil(cement_ton / 100.0),
            "sample_spec": "1 tổ mẫu 10kg",
            "hold_point": "Đạt R3, R7, R28 theo lô xuất xưởng"
        })

        # 4. Cát vàng bê tông (<= 200 m3 / tổ mẫu)
        sand_m3 = materials_summary.get("sand_m3", 1565.06)
        plan.append({
            "category": "Cát bê tông",
            "item": "Cát vàng đổ bê tông (Mk >= 2.5)",
            "standard": "TCVN 7570:2006",
            "volume": sand_m3,
            "unit": "m3",
            "frequency_rule": "<= 200 m3 / lô",
            "required_tests": math.ceil(sand_m3 / 200.0),
            "sample_spec": "1 tổ mẫu 20kg",
            "hold_point": "Thành phần hạt, hàm lượng bùn bụi sét"
        })

        # 5. Đá dăm 1x2 (<= 200 m3 / tổ mẫu)
        stone_m3 = materials_summary.get("stone_m3", 2925.37)
        plan.append({
            "category": "Đá dăm bê tông",
            "item": "Đá dăm 1x2 chọn lọc",
            "standard": "TCVN 7570:2006",
            "volume": stone_m3,
            "unit": "m3",
            "frequency_rule": "<= 200 m3 / lô",
            "required_tests": math.ceil(stone_m3 / 200.0),
            "sample_spec": "1 tổ mẫu 50kg",
            "hold_point": "Độ nén dập, thoi dẹt, độ sạch"
        })

        # 6. Mẫu nén cọc khoan nhồi
        piles_count = structure_counts.get("piles", 26)
        pile_vol = materials_summary.get("pile_concrete_m3", 1051.33)
        plan.append({
            "category": "Mẫu nén bê tông",
            "item": "Bê tông Cọc khoan nhồi C30 (R7, R28)",
            "standard": "TCVN 3118:1993",
            "volume": pile_vol,
            "unit": "m3",
            "frequency_rule": "Mỗi cọc 2 tổ mẫu (<= 50m3 / tổ)",
            "required_tests": piles_count * 2,
            "sample_spec": "3 viên R7 + 3 viên R28",
            "hold_point": "Đạt R28 nghiệm thu cấu kiện cọc"
        })

        # 7. Mẫu nén Dầm Super-T (3 tổ mẫu / phiến dầm)
        girders_count = structure_counts.get("girders", 15)
        plan.append({
            "category": "Mẫu nén bê tông",
            "item": "Bê tông Dầm Super-T C45 (R kéo, R7, R28)",
            "standard": "TCVN 3118:1993",
            "volume": girders_count * 28.987,
            "unit": "m3",
            "frequency_rule": "Mỗi phiến dầm 3 tổ mẫu",
            "required_tests": girders_count * 3,
            "sample_spec": "3 tổ (9 viên) / phiến dầm",
            "hold_point": "Đạt R kéo >= 85% mới cho kích cáp"
        })

        # 8. Siêu âm cọc khoan nhồi (100% cọc)
        plan.append({
            "category": "Thí nghiệm chuyên sâu",
            "item": "Siêu âm độ đồng nhất cọc D1200",
            "standard": "TCVN 9396:2012 / ASTM D6760",
            "volume": piles_count * 6,
            "unit": "Mặt cắt",
            "frequency_rule": "100% số cọc (6 mặt cắt / cọc)",
            "required_tests": piles_count * 6,
            "sample_spec": "Đo xung siêu âm 4 ống sonic",
            "hold_point": "100% cọc đạt khuyết tật loại 1"
        })

        # 9. Thử tải động PDA
        plan.append({
            "category": "Thí nghiệm chuyên sâu",
            "item": "Thí nghiệm thử tải động biến dạng lớn PDA",
            "standard": "ASTM D4945 / TCVN 11321",
            "volume": 1.0,
            "unit": "Cọc",
            "frequency_rule": "Tối thiểu 1% số cọc (1 cọc thử)",
            "required_tests": 1,
            "sample_spec": "Gắn cảm biến gia tốc & biến dạng",
            "hold_point": "Sức chịu tải giới hạn đạt thiết kế"
        })

        return plan
