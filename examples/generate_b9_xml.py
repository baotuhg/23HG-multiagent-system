# -*- coding: utf-8 -*-
"""
XUẤT TIẾN ĐỘ THI CÔNG MS PROJECT XML CHO CỤM B9 - SÂN VẬN ĐỘNG OLYMPIC THƯỜNG TÍN
"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
from _paths import project_path, repo_path  # noqa: E402 — đường dẫn repo / thư mục dự án (AEC_PROJECTS_DIR)
import datetime
import xml.etree.ElementTree as ET

def generate_b9_xml():
    project = ET.Element("Project", xmlns="http://schemas.microsoft.com/project")
    ET.SubElement(project, "Name").text = "TIEN_DO_THI_CONG_CUM_B9_OLYMPIC_THUONG_TIN"
    ET.SubElement(project, "Title").text = "Tiến độ thi công HTKT San lấp Cụm B9 và Đường nội bộ (Mốc 25/10 xong thảm thô)"
    ET.SubElement(project, "StartDate").text = "2026-09-20T07:00:00"
    ET.SubElement(project, "FinishDate").text = "2026-10-25T17:00:00"

    tasks_data = [
        (1, "HTKT THI CÔNG SAN LẤP CỤM B9 VÀ ĐƯỜNG NỘI BỘ", "2026-09-20T07:00:00", "2026-10-25T17:00:00", 36, 0, True),
        # San lấp
        (2, "I. SAN LẤP CỤM B9 (88.7 HA)", "2026-09-20T07:00:00", "2026-10-23T17:00:00", 34, 1, False),
        (3, "Phát quang mặt bằng san lấp", "2026-09-20T07:00:00", "2026-10-04T17:00:00", 15, 2, False),
        (4, "Đào khuôn, bóc đất hữu cơ san lấp dày 10cm", "2026-09-20T07:00:00", "2026-10-04T17:00:00", 15, 2, False),
        (5, "Vận chuyển nội bộ đất hữu cơ san lấp <5km", "2026-09-20T07:00:00", "2026-10-04T17:00:00", 15, 2, False),
        (6, "Thi công đắp đất/cát K90 san nền cụm B9", "2026-09-22T07:00:00", "2026-10-23T17:00:00", 32, 2, True),
        # Đường nội bộ
        (7, "II. ĐƯỜNG NỘI BỘ CỤM B9 (L = 18,034M)", "2026-09-20T07:00:00", "2026-10-25T17:00:00", 36, 1, False),
        (8, "Phát quang mặt bằng tuyến đường nội bộ", "2026-09-20T07:00:00", "2026-10-02T17:00:00", 13, 2, False),
        (9, "Đào khuôn đường vét hữu cơ nền đường", "2026-09-20T07:00:00", "2026-10-03T17:00:00", 14, 2, True),
        (10, "Vận chuyển nội bộ đất khuôn đường <5km", "2026-09-20T07:00:00", "2026-10-03T17:00:00", 14, 2, False),
        (11, "Thi công đắp cát K90 nền đường", "2026-09-22T07:00:00", "2026-10-07T17:00:00", 16, 2, True),
        (12, "Thi công đắp cát K95 (Lớp 1)", "2026-09-25T07:00:00", "2026-10-09T17:00:00", 15, 2, True),
        (13, "Thi công đắp cát K98 (Lớp 1)", "2026-09-28T07:00:00", "2026-10-11T17:00:00", 14, 2, True),
        (14, "Thi công Hệ thống thoát nước mưa TNM (D300-D1800)", "2026-09-26T07:00:00", "2026-10-11T17:00:00", 16, 2, False),
        (15, "Thi công Hệ thống thoát nước thải TNT (Ống HDPE)", "2026-09-28T07:00:00", "2026-10-13T17:00:00", 16, 2, False),
        (16, "Thi công lớp móng Base B (CPĐD loại 2)", "2026-10-06T07:00:00", "2026-10-17T17:00:00", 12, 2, True),
        (17, "Thi công lớp móng Base A (CPĐD loại 1)", "2026-10-10T07:00:00", "2026-10-20T17:00:00", 11, 2, True),
        (18, "Vận chuyển nội bộ đá Base A & B <5km", "2026-10-06T07:00:00", "2026-10-20T17:00:00", 15, 2, False),
        (19, "Tưới nhựa thấm bám 1.0 kg/m2 & Vệ sinh mặt móng", "2026-10-16T07:00:00", "2026-10-24T17:00:00", 9, 2, False),
        (20, "THI CÔNG LỚP THẢM BTN C19 (THẢM THÔ 6.5CM) - VỀ ĐÍCH", "2026-10-17T07:00:00", "2026-10-25T17:00:00", 9, 2, True),
        (21, "MỐC HOÀN THÀNH TOÀN DIỆN THẢM THÔ CỤM B9", "2026-10-25T17:00:00", "2026-10-25T17:00:00", 0, 1, True)
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
    xml_paths = [
        project_path(r"TIEN_DO_THI_CONG_CUM_B9_OLYMPIC\260920_Tien_Do_Thi_Cong_Cum_B9_Olympic_ThuongTin.xml"),
        project_path(r"260920_Tien_Do_Thi_Cong_Cum_B9_Olympic_ThuongTin.xml"),
        repo_path(r"examples\TIEN_DO_THI_CONG_CUM_B9_OLYMPIC\260920_Tien_Do_Thi_Cong_Cum_B9_Olympic_ThuongTin.xml")
    ]
    for xp in xml_paths:
        tree.write(xp, encoding="utf-8", xml_declaration=True)
        print(f"[OK] XML đã lưu: {xp}")

if __name__ == "__main__":
    generate_b9_xml()
