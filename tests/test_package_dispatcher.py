# -*- coding: utf-8 -*-
"""
UNIT TESTS — PACKAGE DISPATCHER (HUB & SPOKE PACKAGING)
Kiểm thử tính năng đóng gói và phân quyền hồ sơ thực chiến công trường.
"""

import os
import unittest
import tempfile

import openpyxl

from tools.package_dispatcher import (
    AECPackageDispatcher, DispatchManifest, audit_all_exported_excels, find_missing_sheet_refs, sheet_is_placeholder,
)


class TestPackageDispatcher(unittest.TestCase):

    def test_dispatch_site_operation(self):
        """Kiểm tra đóng gói theo mô hình Hub & Spoke phân quyền 5 gói."""
        with tempfile.TemporaryDirectory() as tmp_src, tempfile.TemporaryDirectory() as tmp_dst:
            # Tạo các tệp mẫu giả lập
            f_camay = os.path.join(tmp_src, "TienDo_CaMay_CongA5.xlsx")
            f_rebar = os.path.join(tmp_src, "01_To_Hop_Cat_Thep_11m7.xlsx")
            f_kcs = os.path.join(tmp_src, "Ho_So_Bien_Ban_Nghiem_Thu.docx")
            f_qs = os.path.join(tmp_src, "Du_Toan_GXD_TT11.xlsx")
            f_dash = os.path.join(tmp_src, "Master_Dashboard.xlsx")

            for f_path in [f_camay, f_rebar, f_kcs, f_qs, f_dash]:
                with open(f_path, "w", encoding="utf-8") as f:
                    f.write("mock content")

            dispatcher = AECPackageDispatcher(base_output_dir=tmp_dst)
            manifest = dispatcher.dispatch_site_operation_packages(
                project_name="Du_An_Test",
                artifacts_source_dir=tmp_src,
                custom_subfolder="TEST_HUB_SPOKE"
            )

            self.assertEqual(manifest.mode, "site_operation")
            self.assertEqual(manifest.total_files_count, 5)
            self.assertEqual(len(manifest.generated_packages), 5)

            # Kiểm tra các thư mục gói con
            target_root = os.path.join(tmp_dst, "TEST_HUB_SPOKE")
            pkg_a = os.path.join(target_root, "GOI_A_CO_GIOI_VA_DAU_DIEZEL")
            pkg_b = os.path.join(target_root, "GOI_B_XUONG_TIEN_CHE_COT_THEP")
            pkg_c = os.path.join(target_root, "GOI_C_HIEN_TRUONG_QLCL_KCS")
            pkg_d = os.path.join(target_root, "GOI_D_QS_DU_TOAN_THANH_TOAN")
            pkg_e = os.path.join(target_root, "GOI_E_EXECUTIVE_DASHBOARD")

            self.assertTrue(os.path.exists(os.path.join(pkg_a, "TienDo_CaMay_CongA5.xlsx")))
            self.assertTrue(os.path.exists(os.path.join(pkg_b, "01_To_Hop_Cat_Thep_11m7.xlsx")))
            self.assertTrue(os.path.exists(os.path.join(pkg_c, "Ho_So_Bien_Ban_Nghiem_Thu.docx")))
            self.assertTrue(os.path.exists(os.path.join(pkg_d, "Du_Toan_GXD_TT11.xlsx")))
            self.assertTrue(os.path.exists(os.path.join(pkg_e, "Master_Dashboard.xlsx")))
            self.assertTrue(os.path.exists(os.path.join(target_root, "DISPATCH_MANIFEST.json")))

    def test_dispatch_full_industrial_dossier(self):
        """Kiểm tra quy trình công nghiệp xuất 3 tầng hồ sơ chuẩn xác 100%."""
        import openpyxl
        with tempfile.TemporaryDirectory() as tmp_dir:
            # Tạo Master Excel giả lập với 3 sheet
            master_file = os.path.join(tmp_dir, "Master_Test.xlsx")
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "QS_DIEN_GIAI_CHI_TIET"
            ws1["A1"] = "Hạng mục"
            ws1["B1"] = 100.0

            ws2 = wb.create_sheet(title="KHOI_LUONG_DAO_DAP")
            ws2["A1"] = "Đào đất"
            ws2["B1"] = 50.0

            wb.save(master_file)
            wb.close()

            out_dir = os.path.join(tmp_dir, "OUTPUT_DOSSIER")
            dispatcher = AECPackageDispatcher(base_output_dir=out_dir)

            manifest = dispatcher.dispatch_full_industrial_dossier(
                master_excel_path=master_file,
                project_name="Cầu Thử Nghiệm",
                target_dir=out_dir
            )

            # 1. Kiểm tra đủ 3 tầng
            macro_dir = os.path.join(out_dir, "BO_HO_SO_01_MACRO_MASTER_14_SHEET")
            micro_dir = os.path.join(out_dir, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")
            hub_dir = os.path.join(out_dir, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH")

            self.assertTrue(os.path.exists(macro_dir))
            self.assertTrue(os.path.exists(micro_dir))
            self.assertTrue(os.path.exists(hub_dir))

            # 2. Gói A: thiếu dữ liệu ca máy thật => KHÔNG dựng mẫu cầu (cọc khoan nhồi...), chỉ có ghi chú
            pkg_a_dir = os.path.join(hub_dir, "GOI_A_CO_GIOI_VA_DAU_DIEZEL")
            self.assertTrue(os.path.exists(pkg_a_dir))
            self.assertEqual([f for f in os.listdir(pkg_a_dir) if f.endswith(".xlsx")], [])
            self.assertTrue(os.path.exists(os.path.join(pkg_a_dir, "CHUA_CO_DU_LIEU_CA_MAY.md")))

            # 2b. Có dữ liệu ca máy thật (fleet_template) => chép đúng 5 sheet Vincons
            from tools.package_dispatcher import build_vincons_5_sheets_fleet_workbook
            fleet_src = os.path.join(tmp_dir, "CaMay_That.xlsx")
            build_vincons_5_sheets_fleet_workbook(dest_path=fleet_src, project_name="Cầu Thử Nghiệm")
            out_dir2 = os.path.join(tmp_dir, "OUTPUT_DOSSIER_2")
            AECPackageDispatcher(base_output_dir=out_dir2).dispatch_full_industrial_dossier(
                master_excel_path=master_file,
                project_name="Cầu Thử Nghiệm",
                target_dir=out_dir2,
                companion_files={"fleet_template": fleet_src},
            )
            pkg_a2 = os.path.join(out_dir2, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH",
                                  "GOI_A_CO_GIOI_VA_DAU_DIEZEL")
            camay_files = [f for f in os.listdir(pkg_a2) if f.endswith(".xlsx")]
            self.assertTrue(len(camay_files) >= 1)
            wb_camay = openpyxl.load_workbook(os.path.join(pkg_a2, camay_files[0]))
            expected_vincons_sheets = [
                "01_TienDo_CaMay_Master",
                "02_TongHop_CaXe_CaMay_MMTB",
                "03_KeHoach_Dau_Diezel",
                "04_KeHoach_NhanLuc",
                "05_DoiChieu_BocTach"
            ]
            for s in expected_vincons_sheets:
                self.assertIn(s, wb_camay.sheetnames, f"Thiếu sheet Vincons: {s}")
            wb_camay.close()

            # 3. Kiểm tra Manifest MD5
            manifest_json = os.path.join(hub_dir, "DISPATCH_MANIFEST.json")
            self.assertTrue(os.path.exists(manifest_json))
            self.assertTrue(manifest.audit_zero_errors)


class TestMissingSheetRefGate(unittest.TestCase):
    """Quality Gate phải bắt công thức trỏ tới sheet không tồn tại (không sinh ra mã lỗi #REF!)."""

    def test_find_missing_sheet_refs(self):
        sheets = ["DAO_DAP", "Bảng 1"]
        self.assertEqual(find_missing_sheet_refs("=A1+DAO_DAP!B2", sheets), [])
        self.assertEqual(find_missing_sheet_refs("='Bảng 1'!A1*2", sheets), [])
        self.assertEqual(find_missing_sheet_refs("=A1-QS_DIEN_GIAI_CHI_TIET!I8", sheets),
                         ["QS_DIEN_GIAI_CHI_TIET"])
        self.assertEqual(find_missing_sheet_refs("='Sheet Khác'!A1", sheets), ["Sheet Khác"])
        # chuỗi chứa dấu ! và liên kết ngoài không bị coi là tham chiếu sheet
        self.assertEqual(find_missing_sheet_refs('=IF(A1>0,"Xong!","")', sheets), [])
        self.assertEqual(find_missing_sheet_refs("=[1]Ngoai!A1", sheets), [])

    def test_audit_flags_missing_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = openpyxl.Workbook()
            bad.active.title = "DAO_DAP"
            bad.active["A1"] = "=1-QS_DIEN_GIAI_CHI_TIET!I8"
            bad.save(os.path.join(tmp, "bad.xlsx"))
            good = openpyxl.Workbook()
            good.active.title = "DAO_DAP"
            good.active["A1"] = "=1-2"
            good.active["B1"] = "Đã loại bỏ hoàn toàn lỗi #REF!"   # chữ mô tả, không phải lỗi
            good.save(os.path.join(tmp, "good.xlsx"))
            total, err_files, details = audit_all_exported_excels([tmp])
            self.assertEqual((total, err_files), (2, 1))
            self.assertTrue(any("QS_DIEN_GIAI_CHI_TIET" in d and "bad.xlsx" in d for d in details))


class TestEmptyDossiersAreNotExported(unittest.TestCase):

    def test_sheet_is_placeholder(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        self.assertTrue(sheet_is_placeholder(ws))                         # rỗng
        ws["A1"], ws["A2"] = "TIÊU ĐỀ", "Phụ đề"
        self.assertTrue(sheet_is_placeholder(ws))                         # chỉ tiêu đề ở cột A
        ws["B1"] = 5
        self.assertFalse(sheet_is_placeholder(ws))                        # có dữ liệu
        ws2 = wb.create_sheet("X")
        ws2["A1"], ws2["B1"] = "CHƯA LẬP — chưa có dữ liệu", 1
        self.assertTrue(sheet_is_placeholder(ws2))                        # đánh dấu CHƯA LẬP

    def test_full_dossier_skips_placeholder_sheets_and_records_real_audit(self):
        import json
        with tempfile.TemporaryDirectory() as tmp:
            master = os.path.join(tmp, "Master.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "QS_DIEN_GIAI_CHI_TIET"
            ws["A1"], ws["B1"] = "Hạng mục", 100.0
            ph = wb.create_sheet("CAP_PHOI_1M3_VA_TAN_SUAT")             # chỉ có tiêu đề
            ph["A1"], ph["A2"] = "CẤP PHỐI", "chưa có dữ liệu"
            wb.save(master)
            out = os.path.join(tmp, "OUT")
            AECPackageDispatcher(base_output_dir=out).dispatch_full_industrial_dossier(
                master_excel_path=master, project_name="Thử", target_dir=out)
            micro = os.path.join(out, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO")
            self.assertFalse(os.path.exists(os.path.join(micro, "05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx")))
            hub = os.path.join(out, "03_HO_SO_THUC_CHIEN_HUB_AND_SPOKE_5_GOI_VE_TINH", "DISPATCH_MANIFEST.json")
            with open(hub, encoding="utf-8") as f:
                mf = json.load(f)
            self.assertIn("05_Cap_Phoi_1m3_Va_Tan_Suat_Thi_Nghiem.xlsx", mf["skipped_empty_dossiers"])
            self.assertIs(mf["quality_gate"]["zero_formula_errors"], True)    # kết quả kiểm toán thật, không ghi cứng
            self.assertGreater(mf["quality_gate"]["files_audited"], 0)
            self.assertNotIn("zero_dead_numbers", mf["quality_gate"])         # không có phép đo → không khẳng định


if __name__ == "__main__":
    unittest.main()
