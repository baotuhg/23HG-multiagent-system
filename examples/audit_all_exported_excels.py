# -*- coding: utf-8 -*-
"""
Full-scale Excel COM Verification Script:
Kiểm tra toàn diện 100% tất cả các tệp Excel đã xuất của dự án Cầu Km19+529.080
bằng Microsoft Excel COM Engine chính quy.
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path, repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)

import os
import glob

try:
    import win32com.client
except ImportError:   # chỉ có trên Windows cài Microsoft Excel
    raise SystemExit("Script này cần Windows + Microsoft Excel (pywin32). Trên máy khác dùng kiểm tra tĩnh:\n"
                     "    python -m tools.audit_excels_static <thư_mục>")

def audit_exported_excels():
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    # Danh sách các thư mục chứa hồ sơ đã xuất
    base_dirs = [
        repo_path(r"templates"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\BO_HO_SO_01_MACRO_MASTER_14_SHEET"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\BO_HO_SO_01_MACRO_MASTER_14_SHEET"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\HSTK Cầu Km19+529.080_Marker\03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH"),
    ]
    
    # Thêm các file root đã xuất
    root_files = [
        project_path(r"HSTK Cầu Km19+529.080_Marker\BANG_DIEN_GIAI_KHOI_LUONG_PHAN_DUOI_CAU_KM19+529.080.xlsx"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\Bang_Dien_Giai_Khoi_Luong_Thi_Cong_Km19+529.080.xlsx"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\Bang_Tinh_Khoi_Luong_Cau_Km19+529.080.xlsx"),
        project_path(r"HSTK Cầu Km19+529.080_Marker\Bang_Tong_Hop_QLCL_va_Doi_Chieu_QS_Km19+529.080.xlsx"),
    ]

    target_files = set()
    for b in base_dirs:
        for f in glob.glob(os.path.join(b, "**", "*.xlsx"), recursive=True):
            if not os.path.basename(f).startswith("~$"):
                target_files.add(os.path.abspath(f))
                
    for f in root_files:
        if os.path.exists(f):
            target_files.add(os.path.abspath(f))
            
    print(f"================================================================================")
    print(f"TỔNG SỐ TỆP EXCEL ĐÃ XUẤT CẦN AUDIT: {len(target_files)}")
    print(f"================================================================================")
    
    total_errors = 0
    clean_files = 0
    error_details = []

    for idx, fpath in enumerate(sorted(target_files), start=1):
        fname = os.path.basename(fpath)
        rel_folder = os.path.basename(os.path.dirname(fpath))
        file_errs = 0
        try:
            wb = excel.Workbooks.Open(fpath, UpdateLinks=0, ReadOnly=True)
            excel.CalculateFullRebuild()
            
            for ws in wb.Worksheets:
                try:
                    # xlCellTypeFormulas = -4123, xlErrors = 16
                    err_cells = ws.Cells.SpecialCells(-4123, 16)
                    cnt = err_cells.Count
                    file_errs += cnt
                    for cell in err_cells:
                        error_details.append({
                            'file': fname,
                            'folder': rel_folder,
                            'sheet': ws.Name,
                            'cell': cell.Address,
                            'val': cell.Text,
                            'formula': cell.Formula
                        })
                except Exception:
                    pass
            wb.Close(False)
            
            if file_errs == 0:
                clean_files += 1
                print(f"  [{idx:02d}/{len(target_files):02d}] [CLEAN - 0 ERRORS] {rel_folder}/{fname}")
            else:
                total_errors += file_errs
                print(f"  [{idx:02d}/{len(target_files):02d}] [FAIL - {file_errs} ERRORS] {rel_folder}/{fname}")
        except Exception as e:
            print(f"  [{idx:02d}/{len(target_files):02d}] [OPEN ERROR] {fname}: {e}")

    excel.Quit()
    
    print(f"\n================================================================================")
    print(f"KẾT QUẢ KIỂM TRA TOÀN BỘ HỒ SƠ ĐÃ XUẤT:")
    print(f"- Tổng số file kiểm tra: {len(target_files)}")
    print(f"- Số file ĐẠT 100% (0 lỗi): {clean_files} / {len(target_files)}")
    print(f"- TỔNG SỐ LỖI CÔNG THỨC TOÀN HỆ THỐNG: {total_errors}")
    print(f"================================================================================")
    
    if error_details:
        print("\nCHI TIẾT CÁC LỖI CÒN SÓT LẠI:")
        for err in error_details:
            print(f"  + [{err['folder']}/{err['file']}] Sheet: {err['sheet']} | Ô: {err['cell']} | Lỗi: {err['val']} | Hàm: {err['formula']}")
    else:
        print("\n>>> TUYỆT VỜI! 100% CÔNG THỨC TRÊN TOÀN BỘ CÁC TỆP EXCEL ĐỀU SẠCH HOÀN TOÀN (ZERO ERROR)! <<<")

    return total_errors

if __name__ == "__main__":
    audit_exported_excels()
