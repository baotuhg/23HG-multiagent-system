# -*- coding: utf-8 -*-
"""
VÍ DỤ THỰC THI CHUỖI TÁC TỬ THU NHẬN & HỢP NHẤT DỮ LIỆU ĐA PHƯƠNG THỨC
(MULTI-MODAL DATA INGESTION & FUSION PIPELINE)
Tuân thủ Kiến trúc Hệ thống MCP 3 Thành Phần:
[ Giao diện AI: Claude Desktop / Windsurf / Cursor / Antigravity ]
                               |
                               ▼ (Giao thức MCP: cad-mcp / autocad-mcp)
               [ MCP Server điều khiển AutoCAD (Local) ]
                               |
                               ▼ (API / COM Interop: win32com / ezdxf)
      [ Bản vẽ DWG trong AutoCAD ] <---> [ Hồ sơ thiết kế (Excel/PDF) ]
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)

import os
import sys

# Thêm đường dẫn gốc vào sys.path
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from core.agents.aec_cad_extractor import AECCadExtractor
from core.agents.aec_office_extractor import AECOfficeExtractor
from core.agents.aec_markdown_ingestor import AECMarkdownIngestor
from core.agents.aec_data_aggregator import AECDataAggregator

def main():
    print("=" * 80)
    print("  AEC MULTI-MODAL DATA INGESTION & RECONCILIATION PIPELINE")
    print("  Kiến trúc Hệ thống MCP 3 Thành phần:")
    print("  [ AI Interface ] -> [ MCP Server Local ] -> [ AutoCAD DWG ] <-> [ Excel/MD ]")
    print("=" * 80)

    cad_agent = AECCadExtractor()
    office_agent = AECOfficeExtractor()
    md_agent = AECMarkdownIngestor()
    aggregator = AECDataAggregator()

    # Thư mục hồ sơ dự án
    drawing_folder = project_path(r"HSTK Cầu Km19+529.080_Marker")
    design_base = os.path.join(drawing_folder, r"HSTK Cầu Km19+529.080_Marker\01.CAU KM19+529.08")
    md_file = os.path.join(drawing_folder, r"PHAN_TICH_KHOI_LUONG_QS_CAU_KM19+529.080.md")
    excel_file = os.path.join(design_base, r"06. QUANTITY\06.THKL_KM19.5.xlsx")

    # --- BƯỚC 1: KIỂM TRA KẾT NỐI AUTOCAD COM INTEROP & QUÉT BẢN VẼ ---
    print("\n--- BƯỚC 1: TÁC TỬ AEC_CAD_EXTRACTOR (KIẾN TRÚC MCP TẦNG 2 & 3) ---")
    acad_connected = cad_agent.check_autocad_connection()
    if acad_connected:
        print("  -> Kết nối trực tiếp AutoCAD COM Interop: ĐÃ KẾT NỐI (Active Session)")
        open_docs = cad_agent.com_connector.get_open_drawings()
        print(f"  -> Bản vẽ đang mở trong AutoCAD: {open_docs}")
    else:
        print("  -> Chế độ AutoCAD: Standalone File Ingestion (Sẵn sàng nhận lệnh qua MCP Server)")

    # Quét đệ quy toàn bộ thư mục bản vẽ khi có thêm file
    if os.path.exists(drawing_folder):
        cad_scan = cad_agent.scan_drawings_folder(drawing_folder)
        print(f"  -> Tổng số bản vẽ DWG/DXF đã quét: {cad_scan['total_drawings']} bản vẽ")
        for cat, count in cad_scan['categories_summary'].items():
            print(f"     + {cat}: {count} bản vẽ")
    else:
        cad_scan = {"total_drawings": 0, "categories_summary": {}}

    # --- BƯỚC 2: TÁC TỬ AEC_OFFICE_EXTRACTOR BÓC TÁCH BẢNG TÍNH EXCEL ---
    print("\n--- BƯỚC 2: TÁC TỬ AEC_OFFICE_EXTRACTOR BÓC TÁCH BẢNG TÍNH EXCEL ---")
    if os.path.exists(excel_file):
        wb_res = office_agent.process_workbook(excel_file)
        print(f"  -> Tổng số sheet trong tệp Excel: {wb_res.get('sheets_count')}")
        print(f"  -> Phân loại nội dung: {wb_res.get('detected_types')[:3]}...")
    else:
        wb_res = {}

    # --- BƯỚC 3: TÁC TỬ AEC_MARKDOWN_INGESTOR ĐỌC HỒ SƠ THUYẾT MINH ---
    print("\n--- BƯỚC 3: TÁC TỬ AEC_MARKDOWN_INGESTOR ĐỌC HỒ SƠ THUYẾT MINH ---")
    if os.path.exists(md_file):
        md_res = md_agent.process_file(md_file)
        print(f"  -> Dung lượng văn bản: {md_res.get('file_size_chars')} ký tự, {md_res.get('tables_count')} bảng")
        print(f"  -> Mác bê tông phát hiện: {md_res.get('technical_specs', {}).get('concrete_grades')}")
        print(f"  -> Mác thép phát hiện: {md_res.get('technical_specs', {}).get('steel_grades')}")
    else:
        md_res = {}

    # --- BƯỚC 4: ĐỐI SOÁT 2 CHIỀU: [BẢN VẼ DWG] <---> [HỒ SƠ THIẾT KẾ (EXCEL/MD)] ---
    print("\n--- BƯỚC 4: ĐỐI SOÁT 2 CHIỀU [BẢN VẼ DWG] <---> [HỒ SƠ THIẾT KẾ] ---")
    reconciliations = cad_agent.reconcile_with_design_documents(cad_scan, wb_res, md_res)
    for rec in reconciliations:
        mark = {"MATCHED": "[V]", "MISMATCH": "[X]"}.get(rec["status"], "[?]")
        note = f" — {rec['note']}" if rec.get("note") else ""
        print(f"  {mark} {rec['check_item']}: CAD={rec['cad_value']} <---> Hồ sơ={rec['excel_value']} ({rec['status']}){note}")

    # --- BƯỚC 5: TỔNG HỢP VÀO BLACKBOARD STATE (PROJECT_STATE.JSON) ---
    print("\n--- BƯỚC 5: TỔNG HỢP VÀO BLACKBOARD ARCHITECTURE (PROJECT_STATE.JSON) ---")
    target_state = os.path.join(os.environ.get("AEC_STATE_DIR") or os.path.join(root_dir, ".aec_state"), "PROJECT_STATE.json")
    state = aggregator.synthesize_to_project_state(design_base, target_state,
        cad_data=cad_scan, office_data=wb_res, md_data=md_res)
    print(f"  -> Đã lưu nguồn trích xuất: {target_state}")
    print(f"  -> Đối chiếu: {state['meta']['cross_check_status']} ({state['meta']['checks_performed']} chỉ tiêu)")

    print("\n" + "=" * 80)
    print("  [V] HOÀN TẤT THU NHẬN & ĐỐI SOÁT ĐA PHƯƠNG THỨC 3 THÀNH PHẦN!")
    print("=" * 80)

if __name__ == "__main__":
    main()
