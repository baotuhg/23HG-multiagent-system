# -*- coding: utf-8 -*-
"""
UNIT TESTS — PACKAGE DISPATCHER (HUB & SPOKE PACKAGING)
Kiểm thử tính năng đóng gói và phân quyền hồ sơ thực chiến công trường.
"""

import os
import unittest
import tempfile

from tools.package_dispatcher import AECPackageDispatcher, DispatchManifest


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


if __name__ == "__main__":
    unittest.main()
