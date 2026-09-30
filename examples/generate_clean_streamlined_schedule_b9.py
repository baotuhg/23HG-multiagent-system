# -*- coding: utf-8 -*-
"""
GENERATE 100% DYNAMIC FORMULA-DRIVEN SCHEDULE FOR CỤM B9 WITH NATIVE EXCEL CHARTS & COMPARISONS
Chắt lọc và tinh giản tiến độ Vina Alpha:
- 100% CÔNG THỨC SỐNG LIÊN KẾT ĐỘNG TOÀN DIỆN (ZERO SỐ CHẾT)
- TÍCH HỢP ĐẦY ĐỦ CÁC BIỂU ĐỒ TRỰC QUAN NATIVE EXCEL CHARTS:
  1. Biểu đồ đường cong phụ tải thiết bị & nhân công 97 ngày (Sheet 03).
  2. Biểu đồ cột phân bổ dầu Diesel theo 4 tháng (Sheet 04).
  3. Biểu đồ cơ cấu khối lượng vật tư chính (Sheet 05).
  4. Biểu đồ cột đôi SO SÁNH THIẾT BỊ: ĐỊNH MỨC VS ĐỀ XUẤT THỰC TẾ (Sheet 06).
- SHEET 06 SO SÁNH PHỤC HỒI & NÂNG CẤP TỪ SHEET 'SS' CỦA BẢN GỐC:
  * Không còn bất kỳ lỗi #REF! nào.
  * Phân tích đối chiếu: Nhu cầu theo định mức vs Đề xuất BĐH vs Chênh lệch vs Lượng dầu thực tế.
"""

from __future__ import annotations
import datetime
import os
import sys

# Đảm bảo đường dẫn gốc hệ thống có trong sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tools.dynamic_schedule_builder import DynamicScheduleBuilder

OUTPUT_DIR = r"C:\Users\baotu\Downloads\TĐTC vina alpha"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "260820_TDTC_Cum_B9_TINH_GIAN_CHUAN_CPM.xlsx")

REPO_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "TIEN_DO_THI_CONG_CUM_B9_OLYMPIC"))
REPO_FILE = os.path.join(REPO_DIR, "260820_TDTC_Cum_B9_TINH_GIAN_CHUAN_CPM.xlsx")


def build_dynamic_schedule_with_charts():
    builder = DynamicScheduleBuilder()
    
    # 1. Xuất file chính vào thư mục người dùng yêu cầu (Thư mục 1)
    builder.build_cum_b9_workbook(OUTPUT_FILE)
    print(f"[OK] Đã xuất file 100% CÔNG THỨC SỐNG + BIỂU ĐỒ NATIVE tại: {OUTPUT_FILE}")
    print(f"     Kích thước file: {os.path.getsize(OUTPUT_FILE):,} bytes")

    # 2. Đồng bộ bản sao nội bộ vào repository hệ thống
    os.makedirs(REPO_DIR, exist_ok=True)
    builder.build_cum_b9_workbook(REPO_FILE)
    print(f"[OK] Đã lưu bản sao vào kho tài nguyên hệ thống: {REPO_FILE}")


if __name__ == "__main__":
    build_dynamic_schedule_with_charts()
