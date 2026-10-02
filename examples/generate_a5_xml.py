# -*- coding: utf-8 -*-
"""
XUẤT TIẾN ĐỘ MS PROJECT XML CHO CỐNG HỘP TUYẾN A5
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path, repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)
import xml.etree.ElementTree as ET

def generate_a5_xml():
    project = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(project, "Name").text = "TIEN_DO_CA_MAY_CONG_HOP_TUYEN_A5"
    ET.SubElement(project, "Title").text = "Tiến độ thi công Ca xe ca máy Cống Hộp Tuyến A5 (L = 2,170m)"
    ET.SubElement(project, "StartDate").text = "2026-09-20T07:00:00"
    ET.SubElement(project, "FinishDate").text = "2026-10-25T17:00:00"

    tasks_data = [
        (1, "THI CÔNG CỐNG HỘP TUYẾN A5 (L = 2,170M - 63 HỐ GA)", "2026-09-20T07:00:00", "2026-10-25T17:00:00", 36, 0, True),
        (2, "Đào đất hố móng cống hộp 3 mũi (73,800 m3)", "2026-09-20T07:00:00", "2026-10-05T17:00:00", 16, 1, True),
        (3, "Đào hố móng 63 hố ga thu nước & đấu nối (12,650 m3)", "2026-09-22T07:00:00", "2026-10-03T17:00:00", 12, 1, False),
        (4, "Vận chuyển đất đào cự ly <5km (86,450 m3)", "2026-09-20T07:00:00", "2026-10-05T17:00:00", 16, 1, False),
        (5, "Bơm hạ mực nước ngầm hố móng cống 24/7", "2026-09-20T07:00:00", "2026-10-25T17:00:00", 36, 1, False),
        (6, "Đệm cát / đá dăm 4x6 lót móng cống (1,250 m3)", "2026-09-22T07:00:00", "2026-10-05T17:00:00", 14, 1, False),
        (7, "Bê tông lót móng M100# dày 50mm (592.4 m3)", "2026-09-23T07:00:00", "2026-10-08T17:00:00", 16, 1, True),
        (8, "Gia công & lắp dựng cốt thép CB300/CB400 (753.36 tấn)", "2026-09-24T07:00:00", "2026-10-15T17:00:00", 22, 1, True),
        (9, "Lắp dựng ván khuôn thép/phủ phim cống (44,940.7 m2)", "2026-09-25T07:00:00", "2026-10-16T17:00:00", 22, 1, True),
        (10, "Đổ bê tông thân cống B20 (11,870.0 m3)", "2026-09-26T07:00:00", "2026-10-17T17:00:00", 22, 1, True),
        (11, "Vận chuyển bê tông thương phẩm B20 trạm trộn (264 ca xe)", "2026-09-26T07:00:00", "2026-10-17T17:00:00", 22, 1, False),
        (12, "Bê tông & cốt thép 63 hố ga BTCT (409.5 m3)", "2026-09-28T07:00:00", "2026-10-12T17:00:00", 15, 1, False),
        (13, "Máy xúc lốp PC140 cẩu lắp & đầm cóc mang cống (2,170 md)", "2026-09-26T07:00:00", "2026-10-17T17:00:00", 22, 1, False),
        (14, "Đắp cát/đất K95 hoàn trả mang cống (54,200 m3)", "2026-10-05T07:00:00", "2026-10-22T17:00:00", 18, 1, True),
        (15, "Máy ủi D3-D5 san gạt đỉnh móng hoàn trả", "2026-10-05T07:00:00", "2026-10-22T17:00:00", 18, 1, False),
        (16, "Máy lu rung 12-16T đầm nén K95 hoàn trả", "2026-10-05T07:00:00", "2026-10-22T17:00:00", 18, 1, True),
        (17, "Xe téc cấp dầu lưu động & máy phát điện phục vụ", "2026-09-20T07:00:00", "2026-10-25T17:00:00", 36, 1, False),
        (18, "HOÀN THÀNH TOÀN DIỆN HẠNG MỤC CỐNG HỘP TUYẾN A5", "2026-10-25T17:00:00", "2026-10-25T17:00:00", 0, 0, True)
    ]

    tasks_elem = ET.SubElement(project, "Tasks")
    for uid, name, start, finish, dur_days, outline_lvl, is_crit in tasks_data:
        t = ET.SubElement(tasks_elem, "Task")
        ET.SubElement(t, "UID").text = str(uid)
        ET.SubElement(t, "ID").text = str(uid)
        ET.SubElement(t, "Name").text = name
        ET.SubElement(t, "OutlineLevel").text = str(outline_lvl)
        ET.SubElement(t, "Start").text = start
        ET.SubElement(t, "Finish").text = finish
        ET.SubElement(t, "Duration").text = f"PT{dur_days * 8}H0M0S"
        ET.SubElement(t, "Milestone").text = "1" if dur_days == 0 else "0"
        ET.SubElement(t, "Critical").text = "1" if is_crit else "0"

    tree = ET.ElementTree(project)
    paths = [
        project_path(r"CỐNG HỘP TUYẾN A5\260920_Tien_Do_CaMay_Cong_Hop_Tuyen_A5.xml"),
        project_path(r"TIEN_DO_THI_CONG_CUM_B9_OLYMPIC\260920_Tien_Do_CaMay_Cong_Hop_Tuyen_A5.xml"),
        repo_path(r"examples\HO_SO_CONG_HOP_TUYEN_A5\HO_SO_THUC_CHIEN_HUB_AND_SPOKE_CONG_A5\GOI_A_CO_GIOI_VA_DAU_DIEZEL\260920_Tien_Do_CaMay_Cong_Hop_Tuyen_A5.xml")
    ]
    for p in paths:
        tree.write(p, encoding="utf-8", xml_declaration=True)
        print(f"[OK] XML đã lưu: {p}")

if __name__ == "__main__":
    generate_a5_xml()
