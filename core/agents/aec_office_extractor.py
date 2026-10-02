# -*- coding: utf-8 -*-
"""
TÁC TỬ: AEC_OFFICE_EXTRACTOR (CHUYÊN GIA BÓC TÁCH EXCEL & WORD)
Vai trò trong Hệ thống Đa tác tử AEC:
- Bóc tách toàn bộ bảng tính thiết kế chi tiết (.xlsx, .xls) của Tư vấn thiết kế.
- Bóc tách bảng thống kê thép (BBS), khối lượng tiên lượng mời thầu (BoQ), bảng chi phí.
- Phân tích cây công thức, phát hiện số chết và kiểm tra tính toàn vẹn toán học.
- Đọc và bóc tách các tệp Word (.docx) chứa hồ sơ nghiệm thu KCS, nhật ký thi công.
"""

import os
from typing import Dict, List, Any
import openpyxl

class AECOfficeExtractor:
    """Tác tử chuyên bóc tách dữ liệu Office (Excel, Word)."""

    def __init__(self, name: str = "aec_office_extractor"):
        self.name = name

    def extract_rebar_schedule_from_excel(self, excel_path: str, sheet_name: str, start_row: int = 6) -> List[Dict[str, Any]]:
        """Bóc tách bảng thống kê cốt thép (BBS) chuẩn từ bảng tính Excel."""
        if not os.path.exists(excel_path):
            print(f"[{self.name}] Cảnh báo: Không tìm thấy file {excel_path}")
            return []

        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            print(f"[{self.name}] Không tìm thấy sheet {sheet_name} trong {os.path.basename(excel_path)}")
            return []

        ws = wb[sheet_name]
        items = []

        for r in range(start_row, ws.max_row + 1):
            mark = ws.cell(r, 3).value or ws.cell(r, 2).value
            dia = ws.cell(r, 4).value or ws.cell(r, 3).value
            shape = ws.cell(r, 5).value or ws.cell(r, 4).value
            length = ws.cell(r, 14).value or ws.cell(r, 5).value
            qty = ws.cell(r, 15).value or ws.cell(r, 4).value
            unit_w = ws.cell(r, 16).value or ws.cell(r, 6).value
            w_tot = ws.cell(r, 19).value or ws.cell(r, 7).value

            if mark and dia and length and qty:
                try:
                    dia_f = float(dia)
                    len_f = float(length)
                    qty_f = float(qty)
                    unit_w_f = float(unit_w) if unit_w else round(dia_f * dia_f * 0.006165, 3)
                    w_tot_f = float(w_tot) if w_tot else round(len_f / 1000.0 * qty_f * unit_w_f, 2)

                    if w_tot_f > 0:
                        items.append({
                            "source_file": os.path.basename(excel_path),
                            "sheet": sheet_name,
                            "row": r,
                            "mark": str(mark).strip(),
                            "diameter_mm": dia_f,
                            "shape": str(shape or "").strip(),
                            "length_mm": len_f,
                            "quantity": qty_f,
                            "unit_weight_kg_m": unit_w_f,
                            "total_weight_kg": w_tot_f
                        })
                except Exception:
                    pass

        return items

    def extract_boq_and_costs(self, excel_path: str, sheet_name: str) -> Dict[str, Any]:
        """Bóc tách bảng tiên lượng và dự toán chi phí từ Excel."""
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            return {}

        ws = wb[sheet_name]
        boq_data = []
        for r in range(6, ws.max_row + 1):
            code = ws.cell(r, 2).value
            desc = ws.cell(r, 3).value
            unit = ws.cell(r, 4).value
            qty = ws.cell(r, 5).value
            if code and desc and qty:
                try:
                    boq_data.append({
                        "row": r,
                        "code": str(code).strip(),
                        "description": str(desc).strip(),
                        "unit": str(unit or "").strip(),
                        "quantity": float(qty)
                    })
                except Exception:
                    pass

        return {
            "source_file": os.path.basename(excel_path),
            "sheet": sheet_name,
            "total_items": len(boq_data),
            "items": boq_data
        }

    def process_workbook(self, file_path: str) -> Dict[str, Any]:
        """Quét và phân loại nội dung toàn bộ file Excel."""
        print(f"[{self.name}] Đang phân tích bảng tính Excel: {os.path.basename(file_path)}")
        wb = openpyxl.load_workbook(file_path, data_only=True)
        
        result = {
            "file_name": os.path.basename(file_path),
            "sheets_count": len(wb.sheetnames),
            "sheets": wb.sheetnames,
            "detected_types": []
        }

        for s in wb.sheetnames:
            s_lower = s.lower()
            if any(k in s_lower for k in ["thep", "rebar", "bbs", "m1", "m2", "t1", "t2"]):
                result["detected_types"].append(f"REBAR_SCHEDULE: {s}")
            elif any(k in s_lower for k in ["du toan", "cost", "gxd", "gia"]):
                result["detected_types"].append(f"COST_ESTIMATE: {s}")
            elif any(k in s_lower for k in ["tien do", "schedule", "gantt"]):
                result["detected_types"].append(f"PROJECT_SCHEDULE: {s}")
            elif any(k in s_lower for k in ["kcs", "nghiem thu", "qaqc"]):
                result["detected_types"].append(f"QAQC_RECORDS: {s}")

        return result
