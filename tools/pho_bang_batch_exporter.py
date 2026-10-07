# -*- coding: utf-8 -*-
"""
BỘ ĐIỀU PHỐI VÀ XUẤT XƯỞNG HÀNG LOẠT 25 DỰ ÁN TRƯỜNG PHỐ BẢNG
Hệ thống Multi-Agent AEC (23HG-multiagent-system).
Pure Python, Zero LLM — Chuẩn Công Nghiệp 3 Tầng (Zero Formula Errors).
"""
from __future__ import annotations

import os
import sys
import gc
import json
import time
import argparse
import datetime

# Đảm bảo import được tools
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, ".."))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from typing import List, Dict, Any, Optional

from tools.pho_bang_project_definitions import PROJECTS_PHO_BANG, ProjectDefinition
from tools.pho_bang_master_builder import (
    load_project_data,
    build_project_master_workbook,
    build_project_fleet_workbook,
    build_project_companion_files,
)
from tools.package_dispatcher import AECPackageDispatcher


BASE_MARKDOWN_DIR = r"D:\Tú\Trường PTTHNT LCTH&THCS Phố Bảng\1. Hồ sơ KCS PDF\Hồ sơ markdown"
DL_DIR = os.environ.get("AEC_DL_DIR", os.path.join(os.path.expanduser("~"), "Downloads", "Documents", "Trường liên cấp phố bảng_Marker"))
MASTER_XLSM_PATH = r"D:\Tú\Trường PTTHNT LCTH&THCS Phố Bảng\1. Trường liên cấp Phố Bảng 02.08.26.xlsm"


def export_single_project(proj: ProjectDefinition, base_dir: str = BASE_MARKDOWN_DIR) -> Dict[str, Any]:
    """Xuất trọn bộ hệ thống công nghiệp 3 tầng cho 1 dự án."""
    t0 = time.time()
    proj_dir = os.path.join(base_dir, proj.folder_name)
    if not os.path.exists(proj_dir):
        os.makedirs(proj_dir, exist_ok=True)

    print("=" * 80)
    print(f"[*] BẮT ĐẦU XUẤT DỰ ÁN: {proj.full_name}")
    print(f"    - Thư mục đích: {proj_dir}")
    print(f"    - Short name:   {proj.short_name}")
    print(f"    - Tiến độ:      {proj.start_date.strftime('%d/%m/%Y')} -> {proj.finish_date.strftime('%d/%m/%Y')}")
    print("=" * 80)

    # 1. Nạp dữ liệu đa nguồn
    tasks, materials, rebar_items = load_project_data(proj, base_dir, DL_DIR, MASTER_XLSM_PATH)
    print(f"  [+] Đã nạp dữ liệu: {len(tasks)} công tác, {len(materials)} vật tư, {len(rebar_items)} cấu kiện thép.")

    # 2. Tạo Master Workbook 14 Sheet
    master_xlsx = os.path.join(proj_dir, f"Ho_So_KCS_QS_TienDo_{proj.short_name}.xlsx")
    build_project_master_workbook(proj, tasks, materials, rebar_items, master_xlsx)

    # 3. Tạo Gói A Ca máy Dầu diezel chuẩn 5 sheets Vincons (Sheet 01 layout 3 tầng hợp nhất)
    fleet_xlsx = os.path.join(proj_dir, f"TDTC_CaXe_CaMay_DauDiezel_{proj.short_name}.xlsx")
    build_project_fleet_workbook(proj, tasks, fleet_xlsx)

    # 4. Tạo companion files
    companion = build_project_companion_files(proj, tasks, materials, proj_dir)

    # 5. Kích hoạt AECPackageDispatcher đóng gói 3 Tầng
    print("  [*] Đang kích hoạt AECPackageDispatcher đóng gói 3 tầng...")
    dispatcher = AECPackageDispatcher(base_output_dir=proj_dir)
    manifest = dispatcher.dispatch_full_industrial_dossier(
        master_excel_path=master_xlsx,
        project_name=proj.full_name,
        target_dir=proj_dir,
        companion_files={
            "fleet_template": fleet_xlsx,
            "fleet_xml": companion["master_xml"],
            "mpp": companion["mpp"],
            "docx": companion["docx"],
            "bptc": companion["bptc"],
            "audit": companion["audit"]
        }
    )

    # 6. Đồng bộ DISPATCH_MANIFEST.json ra thư mục gốc dự án & dọn dẹp file tạm
    manifest_in_hub = os.path.join(proj_dir, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH", "DISPATCH_MANIFEST.json")
    manifest_at_root = os.path.join(proj_dir, "DISPATCH_MANIFEST.json")
    if os.path.exists(manifest_in_hub):
        import shutil
        shutil.copyfile(manifest_in_hub, manifest_at_root)

    # Dọn dẹp các tệp tạm ở root vì đã được phân phối đầy đủ vào 3 tầng
    temp_files = [
        master_xlsx, fleet_xlsx, companion["master_xml"], companion["mpp"],
        companion["docx"], companion["bptc"], companion["audit"]
    ]
    for tf in temp_files:
        if os.path.exists(tf):
            try:
                os.remove(tf)
            except OSError:
                pass

    elapsed = time.time() - t0
    print(f"  [DONE] Hoàn tất xuất xưởng '{proj.short_name}' trong {elapsed:.1f}s | Quality Gate Zero Error: {manifest.audit_zero_errors}")
    gc.collect()

    return {
        "project": proj.full_name,
        "short_name": proj.short_name,
        "dir": proj_dir,
        "elapsed_sec": elapsed,
        "zero_errors": manifest.audit_zero_errors,
        "total_files": manifest.total_files_count
    }


def run_batch_export(indices: Optional[List[int]] = None, base_dir: str = BASE_MARKDOWN_DIR):
    """Chạy hàng loạt danh sách các dự án."""
    total_projects = len(PROJECTS_PHO_BANG)
    to_run = []
    if indices:
        for idx in indices:
            if 1 <= idx <= total_projects:
                to_run.append((idx, PROJECTS_PHO_BANG[idx - 1]))
    else:
        to_run = list(enumerate(PROJECTS_PHO_BANG, 1))

    print("#" * 80)
    print(f"# KÍCH HOẠT QUY TRÌNH XUẤT TRỌN BỘ CÔNG NGHIỆP 3 TẦNG CHO {len(to_run)}/{total_projects} DỰ ÁN")
    print(f"# Căn cứ: Luật Xây dựng 135/2025/QH15, Nghị định 207/2026/NĐ-CP, Thông tư 38/2026/TT-BXD")
    print(f"# Thư mục gốc: {base_dir}")
    print("#" * 80)

    results = []
    overall_start = time.time()
    for seq, (proj_idx, proj) in enumerate(to_run, 1):
        print(f"\n>>> [{seq}/{len(to_run)}] XỬ LÝ DỰ ÁN SỐ {proj_idx}: {proj.folder_name}")
        try:
            res = export_single_project(proj, base_dir)
            results.append(res)
        except Exception as e:
            print(f"  [ERROR] Lỗi xử lý dự án {proj.folder_name}: {e}")
            import traceback
            traceback.print_exc()
            results.append({
                "project": proj.full_name,
                "short_name": proj.short_name,
                "dir": os.path.join(base_dir, proj.folder_name),
                "elapsed_sec": 0,
                "zero_errors": False,
                "total_files": 0,
                "error": str(e)
            })

    total_time = time.time() - overall_start
    print("\n" + "=" * 80)
    print("                     BÁO CÁO TỔNG HỢP XUẤT XƯỞNG HÀNG LOẠT")
    print("=" * 80)
    print(f"Tổng số dự án đã chạy: {len(results)}")
    pass_cnt = sum(1 for r in results if r.get("zero_errors"))
    fail_cnt = len(results) - pass_cnt
    print(f"Trạng thái Quality Gate: PASS {pass_cnt}/{len(results)} ({pass_cnt/len(results)*100:.1f}%) | FAIL: {fail_cnt}")
    print(f"Tổng thời gian thực hiện: {total_time:.1f} giây ({total_time/60:.2f} phút)")
    print("-" * 80)
    for idx, r in enumerate(results, 1):
        status = "PASS 100%" if r.get("zero_errors") else "FAIL / ERROR"
        print(f"{idx:2d}. {r['short_name']:<30} | {r['total_files']:2d} tệp | {r['elapsed_sec']:4.1f}s | {status}")
    print("=" * 80)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Batch exporter cho 25 dự án trường Phố Bảng")
    parser.add_argument("--index", "-i", type=int, help="Chỉ định STT dự án cần chạy (1-25)")
    parser.add_argument("--all", "-a", action="store_true", help="Chạy toàn bộ 25 dự án")
    parser.add_argument("--range", "-r", type=str, help="Dải STT cần chạy (vd: 1-5, 6-10)")
    args = parser.parse_args()

    if args.index:
        run_batch_export(indices=[args.index])
    elif args.range:
        parts = args.range.split("-")
        start_i, end_i = int(parts[0]), int(parts[1])
        run_batch_export(indices=list(range(start_i, end_i + 1)))
    elif args.all:
        run_batch_export()
    else:
        # Mặc định chạy thử dự án số 1
        print("Chạy thử nghiệm dự án số 1:")
        run_batch_export(indices=[1])
