#!/usr/bin/env python3
"""AEC BoQ & Tender Engineering Tool adapted for Vietnamese Construction Standards.

Conforms to:
- Thông tư 36/2026/TT-BXD (Quản lý chi phí xây dựng: Chi phí trực tiếp T, Gián tiếp GT, Thu nhập tính trước TL, Thuế VAT).
- Thông tư 38/2026/TT-BXD (Hệ thống Định mức dự toán xây dựng: AB, AC, AD, AE, AF, AG...).
- GAEB DA XML (International / European BoQ exchange standard via pyGAEB).
"""
import argparse
import json
import sys
from decimal import Decimal
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

from pygaeb.models.document import GAEBDocument
from pygaeb.diff import BoQDiff, DiffMode
from pygaeb.convert.to_excel import to_excel

# Import engine dự toán Việt Nam
try:
    from vn_cost_engine import (
        CostItem, CostSummary, PROJECT_RATES, NORM_PREFIX_MAP,
        calculate_vietnam_estimate, export_vietnam_estimate_excel
    )
except ImportError:
    from .vn_cost_engine import (
        CostItem, CostSummary, PROJECT_RATES, NORM_PREFIX_MAP,
        calculate_vietnam_estimate, export_vietnam_estimate_excel
    )


def load_any_doc(file_path: str) -> Tuple[Optional[GAEBDocument], Optional[Dict[str, Any]]]:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Tệp không tồn tại: {file_path}")
    raw = path.read_text(encoding="utf-8")
    data = json.loads(raw)
    if "award" in data or "source_version" in data:
        return GAEBDocument.model_validate_json(raw), None
    return None, data


def parse_vietnam_items(data: Dict[str, Any]) -> Tuple[str, str, List[CostItem]]:
    pname = data.get("project_name", "Công trình Xây dựng")
    ptype = data.get("project_type", "GIAO_THONG")
    items_raw = data.get("items", [])

    cost_items: List[CostItem] = []
    for idx, r in enumerate(items_raw, start=1):
        stt = r.get("stt", idx)
        mh = r.get("ma_hieu", f"CV.{idx:03d}")
        ten = r.get("ten_cong_tac", r.get("short_text", "Công tác xây dựng"))
        dvt = r.get("don_vi", r.get("unit", "m3"))
        kl = Decimal(str(r.get("khoi_luong", r.get("qty", 0))))
        dg_vl = Decimal(str(r.get("don_gia_vat_lieu", r.get("unit_price", 0))))
        dg_nc = Decimal(str(r.get("don_gia_nhan_cong", 0)))
        dg_m = Decimal(str(r.get("don_gia_may", 0)))
        note = r.get("ghi_chu", "")

        cost_items.append(CostItem(
            stt=stt,
            ma_hieu=mh,
            ten_cong_tac=ten,
            don_vi=dvt,
            khoi_luong=kl,
            don_gia_vat_lieu=dg_vl,
            don_gia_nhan_cong=dg_nc,
            don_gia_may=dg_m,
            ghi_chu=note
        ))
    return pname, ptype, cost_items


def inspect_vietnam_estimate(file_path: str, project_type_override: Optional[str] = None):
    gaeb_doc, vn_data = load_any_doc(file_path)

    if vn_data:
        pname, ptype, items = parse_vietnam_items(vn_data)
        if project_type_override and project_type_override in PROJECT_RATES:
            ptype = project_type_override

        summary, cost_items = calculate_vietnam_estimate(items, project_type=ptype, project_name=pname)

        print("=========================================================================================")
        print("          BẢNG TỔNG HỢP DỰ TOÁN CHI PHÍ XÂY DỰNG (THEO THÔNG TƯ 36/2026/TT-BXD)          ")
        print("=========================================================================================")
        print(f"Dự án: {summary.project_name}")
        print(f"Loại công trình: {PROJECT_RATES[summary.project_type]['name']}")
        print(f"Số lượng đầu việc: {len(cost_items)} công tác")
        print("-----------------------------------------------------------------------------------------")
        print(f"{'Ký hiệu':<8} | {'Khoản mục chi phí':<40} | {'Tỷ lệ %':<10} | {'Giá trị (VNĐ)'}")
        print("-" * 80)
        print(f"{'T':<8} | {'1. Chi phí trực tiếp (VL + NC + M)':<40} | {'':<10} | {summary.chi_phi_truc_tiep:>15,f} đ")
        print(f"{'VL':<8} | {'   - Chi phí vật liệu':<40} | {'':<10} | {summary.chi_phi_vat_lieu:>15,f} đ")
        print(f"{'NC':<8} | {'   - Chi phí nhân công':<40} | {'':<10} | {summary.chi_phi_nhan_cong:>15,f} đ")
        print(f"{'M':<8} | {'   - Chi phí máy thi công':<40} | {'':<10} | {summary.chi_phi_may:>15,f} đ")
        print(f"{'GT':<8} | {'2. Chi phí gián tiếp':<40} | {'':<10} | {summary.tong_chi_phi_gian_tiep:>15,f} đ")
        print(f"{'C_C':<8} | {'   - Chi phí chung (T x %)':<40} | {PROJECT_RATES[summary.project_type]['rate_chi_phi_chung']*100:<9.1f}% | {summary.chi_phi_chung:>15,f} đ")
        print(f"{'C_NT':<8} | {'   - Chi phí nhà tạm điều hành':<40} | {PROJECT_RATES[summary.project_type]['rate_nha_tam']*100:<9.1f}% | {summary.chi_phi_nha_tam:>15,f} đ")
        print(f"{'C_KXD':<8} | {'   - Chi phí KXD từ thiết kế':<40} | {PROJECT_RATES[summary.project_type]['rate_khong_xac_dinh']*100:<9.1f}% | {summary.chi_phi_kxd:>15,f} đ")
        print(f"{'TL':<8} | {'3. Thu nhập chịu thuế tính trước':<40} | {PROJECT_RATES[summary.project_type]['rate_thu_nhap_tinh_truoc']*100:<9.1f}% | {summary.thu_nhap_chiu_thue_tinh_truoc:>15,f} đ")
        print(f"{'G':<8} | {'CHI PHÍ XÂY DỰNG TRƯỚC THUẾ':<40} | {'':<10} | {summary.chi_phi_xay_dung_truoc_thue:>15,f} đ")
        print(f"{'VAT':<8} | {'4. Thuế giá trị gia tăng':<40} | {summary.thue_vat/summary.chi_phi_xay_dung_truoc_thue*100:<9.1f}% | {summary.thue_vat:>15,f} đ")
        print("=" * 80)
        print(f"{'G_XD':<8} | {'TỔNG CỘNG CHI PHÍ XÂY DỰNG SAU THUẾ':<40} | {'':<10} | {summary.tong_chi_phi_xay_dung_sau_thue:>15,f} VNĐ")
        print("=========================================================================================\n")

        print("DANH SÁCH CHI TIẾT CÁC ĐẦU VIỆC (MÃ HIỆU THÔNG TƯ 38/2026/TT-BXD):")
        print(f"{'STT':<4} | {'Mã ĐM':<10} | {'Tên công tác xây dựng':<38} | {'ĐVT':<6} | {'Khối lượng':<10} | {'Đơn giá tổng':<13} | {'Thành tiền (VNĐ)'}")
        print("-" * 105)
        for it in cost_items:
            stext = it.ten_cong_tac[:36]
            print(f"{it.stt:<4} | {it.ma_hieu:<10} | {stext:<38} | {it.don_vi:<6} | {float(it.khoi_luong):<10.2f} | {int(it.don_gia_tong_hop):<13,d} | {int(it.thanh_tien_truc_tiep):>16,d}")
        print("-" * 105)

    elif gaeb_doc:
        # Standard GAEB display
        items = list(gaeb_doc.iter_items())
        print(f"Hồ sơ chuẩn GAEB quốc tế: {gaeb_doc.award.project_name or Path(file_path).name}")
        print(f"Tổng số đầu việc: {len(items)} | Tiền tệ: {gaeb_doc.award.currency or 'EUR'}")
        for it in items:
            print(f"  {it.oz}: {it.short_text} | {it.qty} {it.unit} | {it.total_price}")


def compare_vietnam_boqs(file_a: str, file_b: str):
    gaeb_a, vn_a = load_any_doc(file_a)
    gaeb_b, vn_b = load_any_doc(file_b)

    if vn_a and vn_b:
        _, _, items_a = parse_vietnam_items(vn_a)
        _, _, items_b = parse_vietnam_items(vn_b)

        map_a = {it.ma_hieu: it for it in items_a}
        map_b = {it.ma_hieu: it for it in items_b}

        all_keys = list(dict.fromkeys(list(map_a.keys()) + list(map_b.keys())))

        added = [map_b[k] for k in all_keys if k not in map_a]
        removed = [map_a[k] for k in all_keys if k not in map_b]
        modified = []
        for k in all_keys:
            if k in map_a and k in map_b:
                a, b = map_a[k], map_b[k]
                if a.khoi_luong != b.khoi_luong or a.don_gia_tong_hop != b.don_gia_tong_hop:
                    modified.append((a, b))

        print("=========================================================================================")
        print("          KẾT QUẢ SO SÁNH 2 PHIÊN BẢN DỰ TOÁN / BẢNG TIÊN LƯỢNG (VIỆT NAM)               ")
        print("=========================================================================================")
        print(f"Hồ sơ gốc (A): {Path(file_a).name}")
        print(f"Hồ sơ điều chỉnh (B): {Path(file_b).name}")
        print(f"Hạng mục phát sinh mới: {len(added)} mục")
        print(f"Hạng mục cắt giảm bỏ:   {len(removed)} mục")
        print(f"Hạng mục thay đổi KL/ĐG: {len(modified)} mục")
        print("-" * 80)

        delta_total = Decimal("0")

        if added:
            print("\n[+] CÁC HẠNG MỤC PHÁT SINH MỚI (ADDITIONAL ITEMS):")
            for it in added:
                delta_total += it.thanh_tien_truc_tiep
                print(f"  + [{it.ma_hieu}] {it.ten_cong_tac}: KL={it.khoi_luong} {it.don_vi} x ĐG={it.don_gia_tong_hop:,} = +{it.thanh_tien_truc_tiep:,} VNĐ")

        if removed:
            print("\n[-] CÁC HẠNG MỤC BỊ CẮT GIẢM (REMOVED ITEMS):")
            for it in removed:
                delta_total -= it.thanh_tien_truc_tiep
                print(f"  - [{it.ma_hieu}] {it.ten_cong_tac}: -{it.thanh_tien_truc_tiep:,} VNĐ")

        if modified:
            print("\n[*] CÁC HẠNG MỤC BIẾN ĐỘNG KHỐI LƯỢNG / ĐƠN GIÁ (VARIATIONS):")
            for a, b in modified:
                d_kl = b.khoi_luong - a.khoi_luong
                d_cost = b.thanh_tien_truc_tiep - a.thanh_tien_truc_tiep
                delta_total += d_cost
                print(f"  * [{b.ma_hieu}] {b.ten_cong_tac}:")
                print(f"      Khối lượng: {a.khoi_luong} -> {b.khoi_luong} (Lệch: {d_kl:+} {b.don_vi})")
                print(f"      Đơn giá:    {a.don_gia_tong_hop:,} -> {b.don_gia_tong_hop:,} VNĐ")
                print(f"      Chênh lệch thành tiền: {d_cost:+,} VNĐ")

        print("=" * 80)
        print(f"TỔNG CHÊNH LỆCH CHI PHÍ TRỰC TIẾP PHÁT SINH: {delta_total:+,} VNĐ")
        print("=========================================================================================")

    elif gaeb_a and gaeb_b:
        result = BoQDiff.compare(gaeb_a, gaeb_b, mode=DiffMode.DEFAULT)
        print(f"GAEB Diff: Added={len(result.items_added)}, Removed={len(result.items_removed)}, Modified={len(result.items_modified)}")


def main():
    parser = argparse.ArgumentParser(description="AEC BoQ & Construction Cost Tool for Vietnam")
    parser.add_argument("file", help="Đường dẫn file hồ sơ dự toán (JSON hoặc GAEB XML)")
    parser.add_argument("--type", choices=["GIAO_THONG", "DAN_DUNG", "HA_TANG_KY_THUAT", "NONG_NGHIEP_PTNT"],
                        help="Loại công trình theo Thông tư 36/2026/TT-BXD", default=None)
    parser.add_argument("--compare", help="File dự toán thứ hai để so sánh phát sinh", default=None)
    parser.add_argument("--excel", help="Xuất ra file Excel (.xlsx) chuẩn dự toán Thông tư 36/2026", default=None)
    args = parser.parse_args()

    if args.compare:
        compare_vietnam_boqs(args.file, args.compare)
    elif args.excel:
        gaeb_doc, vn_data = load_any_doc(args.file)
        if vn_data:
            pname, ptype, items = parse_vietnam_items(vn_data)
            if args.type:
                ptype = args.type
            summary, cost_items = calculate_vietnam_estimate(items, project_type=ptype, project_name=pname)
            export_vietnam_estimate_excel(summary, cost_items, args.excel)
        elif gaeb_doc:
            to_excel(gaeb_doc, args.excel)
            print(f"Đã xuất file GAEB Excel: {args.excel}")
    else:
        inspect_vietnam_estimate(args.file, project_type_override=args.type)


if __name__ == "__main__":
    main()
