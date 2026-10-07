# -*- coding: utf-8 -*-
"""
KỊCH BẢN THIẾT LẬP ĐỘI HÌNH "DREAM TEAM" CHO DỰ ÁN CẦU KM19+529.080
Tập hợp các tệp tin xuất sắc nhất vào 5 gói chuyên môn thực chiến + 1 kho lưu trữ dữ liệu gốc.
Xóa bỏ hoàn toàn các thư mục lồng đúp và các file copy trùng lặp.
"""
import os
import shutil
import glob

BASE_DIR = r"C:\Users\baotu\Downloads\HSTK Cầu Km19+529.080_Marker"

DREAM_TEAM_STRUCTURE = {
    "00_BAN_CHI_HUY_MASTER": [
        (r"BO_HO_SO_01_MACRO_MASTER_14_SHEET\Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx", "Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx"),
        (r"BO_HO_SO_01_MACRO_MASTER_14_SHEET\Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp", "Tien_Do_Thi_Cong_Cau_Km19+529.080.mpp"),
        (r"BO_HO_SO_01_MACRO_MASTER_14_SHEET\Tien_Do_Thi_Cong_Cau_Km19+529.080.xml", "Tien_Do_Thi_Cong_Cau_Km19+529.080.xml"),
        (r"BO_HO_SO_01_MACRO_MASTER_14_SHEET\Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md", "Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md"),
        (r"BO_HO_SO_01_MACRO_MASTER_14_SHEET\BAO_CAO_THAM_TRA_AEC_AUDIT.md", "BAO_CAO_THAM_TRA_AEC_AUDIT.md"),
        (r"03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\BANG_PHAN_QUYEN_VA_BIEN_BAN_BAN_GIAO_5_GOI_VE_TINH.xlsx", "BANG_PHAN_QUYEN_VA_BIEN_BAN_BAN_GIAO_5_GOI_VE_TINH.xlsx"),
        (r"03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\BANG_PHAN_QUYEN_VA_BIEN_BAN_BAN_GIAO_5_GOI_VE_TINH.md", "BANG_PHAN_QUYEN_VA_BIEN_BAN_BAN_GIAO_5_GOI_VE_TINH.md"),
    ],
    "01_HIEN_TRUONG_QLCL_KCS": [
        (r"Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Km19.xlsx", "Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Km19.xlsx"),
        (r"BO_HO_SO_01_MACRO_MASTER_14_SHEET\Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx", "Ho_So_Bien_Ban_Nghiem_Thu_KCS_Cau_Km19+529.080.docx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx", "11_Danh_Muc_KCS_22_Bien_Ban_Nghiem_Thu.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx", "12_Mau_A4_Bien_Ban_Nghiem_Thu_Cong_Viec.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx", "13_Mau_A4_Bien_Ban_Nghiem_Thu_Vat_Lieu.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx", "14_Mau_A4_Bien_Ban_Lay_Mau_Thi_Nghiem_R7_R28.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx", "05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx"),
    ],
    "02_XUONG_TIEN_CHE_COT_THEP": [
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx", "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx", "04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx"),
        (r"01_HE_THONG_CAT_THEP_REBARCUT\00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx", "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx"),
        (r"01_HE_THONG_CAT_THEP_REBARCUT\README_QUY_TRINH_VAN_HANH_BAI_THEP.md", "README_QUY_TRINH_VAN_HANH_BAI_THEP.md"),
    ],
    "03_KINH_TE_QS_DU_TOAN_THANH_TOAN": [
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx", "03_QS_Dien_Giai_Chi_Tiet_Takeoff.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx", "02_Khoi_Luong_Dao_Dap_Trinh_Dien.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx", "08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx", "06_Phan_Tich_Vat_Tu_Chi_Tiet_WBS.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx", "07_Tong_Hop_Nhu_Cau_Vat_Tu_BOM_4_Giai_Doan.xlsx"),
        (r"BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO\09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx", "09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx"),
    ],
    "04_CO_GIOI_THIET_BI_VA_DAU_DIEZEL": [
        (r"03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx", "260920_TDTC_CaXe_CaMay_DauDiezel_Cau_Km19+529.080.xlsx"),
        (r"03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_Tien_Do_CaMay_Cau_Km19+529.080.xml", "260920_Tien_Do_CaMay_Cau_Km19+529.080.xml"),
    ],
    "05_DU_LIEU_GOC_SCAN_MARKER": [
        ("bang_so_lieu.json", "bang_so_lieu.json"),
        ("thep_cho_to_hop_cat.json", "thep_cho_to_hop_cat.json"),
        ("tien_luong_du_toan_boq.json", "tien_luong_du_toan_boq.json"),
        ("du_lieu.json", "du_lieu.json"),
        ("noi_dung.md", "noi_dung.md"),
        ("noi_dung.ai.md", "noi_dung.ai.md"),
        ("noi_dung.txt", "noi_dung.txt"),
        ("chia_doan.jsonl", "chia_doan.jsonl"),
        ("kiem_tra_so_lieu.md", "kiem_tra_so_lieu.md"),
        ("can_kiem_tra.md", "can_kiem_tra.md"),
        ("goi_y_cho_AI.txt", "goi_y_cho_AI.txt"),
        ("bang_so_lieu.xlsx", "bang_so_lieu.xlsx"),
        ("HUONG_DAN_KET_QUA.txt", "HUONG_DAN_KET_QUA.txt"),
    ]
}

def execute_reorganization():
    print(f"=== BẮT ĐẦU TÁI THIẾT ĐỘI HÌNH DREAM TEAM DỰ ÁN CẦU KM19 ===")
    
    # 1. Tạo các folder mới
    for folder_name in DREAM_TEAM_STRUCTURE.keys():
        fpath = os.path.join(BASE_DIR, folder_name)
        os.makedirs(fpath, exist_ok=True)
        print(f"[OK] Đã tạo thư mục chuyên môn: {folder_name}")

    # 2. Sao chép các tệp tin xuất sắc nhất vào từng thư mục
    for folder_name, file_tuples in DREAM_TEAM_STRUCTURE.items():
        dest_dir = os.path.join(BASE_DIR, folder_name)
        for src_rel, dest_name in file_tuples:
            src_full = os.path.join(BASE_DIR, src_rel)
            dest_full = os.path.join(dest_dir, dest_name)
            if os.path.exists(src_full):
                shutil.copy2(src_full, dest_full)
                print(f"  -> [{folder_name}] Đã nạp cầu thủ: {dest_name}")
            else:
                print(f"  [!] CẢNH BÁO: Không tìm thấy file nguồn: {src_rel}")

    # 2b. Sao chép các thư mục con chuyên biệt cho Xưởng thép
    # LENH_CAT_CNC_CSV
    src_cnc = os.path.join(BASE_DIR, r"01_HE_THONG_CAT_THEP_REBARCUT\LENH_CAT_CNC_CSV")
    dest_cnc = os.path.join(BASE_DIR, r"02_XUONG_TIEN_CHE_COT_THEP\LENH_CAT_CNC_CSV")
    if os.path.exists(src_cnc):
        if os.path.exists(dest_cnc): shutil.rmtree(dest_cnc)
        shutil.copytree(src_cnc, dest_cnc)
        print(f"  -> [02_XUONG_TIEN_CHE_COT_THEP] Đã sao chép 11 file CSV lệnh cắt CNC.")

    # THEO_TUNG_DUONG_KINH_PHI
    src_phi = os.path.join(BASE_DIR, r"01_HE_THONG_CAT_THEP_REBARCUT\THEO_TUNG_DUONG_KINH_PHI")
    dest_phi = os.path.join(BASE_DIR, r"02_XUONG_TIEN_CHE_COT_THEP\THEO_TUNG_DUONG_KINH_PHI")
    if os.path.exists(src_phi):
        if os.path.exists(dest_phi): shutil.rmtree(dest_phi)
        shutil.copytree(src_phi, dest_phi)
        print(f"  -> [02_XUONG_TIEN_CHE_COT_THEP] Đã sao chép 12 file bóc tách theo từng phi.")

    # 3. Dọn dẹp các thư mục đúp lồng nhau và thư mục hệ thống cũ
    old_folders = [
        os.path.join(BASE_DIR, "HSTK Cầu Km19+529.080_Marker"),
        os.path.join(BASE_DIR, "BO_HO_SO_01_MACRO_MASTER_14_SHEET"),
        os.path.join(BASE_DIR, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO"),
        os.path.join(BASE_DIR, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH"),
        os.path.join(BASE_DIR, "01_HE_THONG_CAT_THEP_REBARCUT"),
    ]
    for of in old_folders:
        if os.path.exists(of):
            try:
                shutil.rmtree(of)
                print(f"[REMOVED] Đã dọn dẹp sạch thư mục cũ/đúp: {os.path.basename(of)}")
            except Exception as e:
                print(f"[NOTE] Không thể xóa {of}: {e}")

    # 4. Dọn dẹp các file lẻ ở thư mục gốc (đã được nạp vào 01 hoặc 05)
    root_redundant_files = [
        "Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Cau_Km19.xlsx",
        "Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Cot_Thep_Cau_Km19.xlsx",
        "Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Tong_Hop_Cau_Km19.xlsx",
        "Bang_Theo_Doi_Tan_Suat_Thi_Nghiem_Be_Tong_Va_Thep_Cau_Km19.xlsx",
        "bang_so_lieu.json",
        "thep_cho_to_hop_cat.json",
        "tien_luong_du_toan_boq.json",
        "du_lieu.json",
        "noi_dung.md",
        "noi_dung.ai.md",
        "noi_dung.txt",
        "chia_doan.jsonl",
        "kiem_tra_so_lieu.md",
        "can_kiem_tra.md",
        "goi_y_cho_AI.txt",
        "bang_so_lieu.xlsx",
        "HUONG_DAN_KET_QUA.txt",
    ]
    for rf in root_redundant_files:
        rf_path = os.path.join(BASE_DIR, rf)
        if os.path.exists(rf_path):
            try:
                os.remove(rf_path)
                print(f"[CLEANED ROOT] Đã dọn tệp lẻ khỏi thư mục gốc: {rf}")
            except Exception as e:
                pass

    # Xóa lock files ~$*.xlsx
    for lock_f in glob.glob(os.path.join(BASE_DIR, "~$*")):
        try: os.remove(lock_f)
        except: pass

    print("\n=== HOÀN TẤT THIẾT LẬP ĐỘI HÌNH DREAM TEAM ===")

if __name__ == "__main__":
    execute_reorganization()
