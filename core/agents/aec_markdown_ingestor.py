# -*- coding: utf-8 -*-
"""
TÁC TỬ: AEC_MARKDOWN_INGESTOR (CHUYÊN GIA BÓC TÁCH HỒ SƠ THIẾT KẾ MARKDOWN & VĂN BẢN)
Vai trò trong Hệ thống Đa tác tử AEC:
- Đọc hiểu hồ sơ thiết kế kỹ thuật, thuyết minh biện pháp thi công, chỉ dẫn kỹ thuật dạng Markdown.
- Bóc tách cấu trúc đề mục, bảng biểu Markdown, danh mục công tác WBS.
- Trích xuất tự động các tham số kỹ thuật then chốt: mác bê tông, tiêu chuẩn áp dụng (TCVN, QCVN, ASTM), điều kiện nghiệm thu.
- Cung cấp dữ liệu ngữ cảnh phong phú cho Agent Tổng hợp (Aggregator).
"""

import os
import re
from typing import Dict, List, Any

class AECMarkdownIngestor:
    """Tác tử chuyên bóc tách hồ sơ thiết kế và thuyết minh kỹ thuật định dạng Markdown."""

    def __init__(self, name: str = "aec_markdown_ingestor"):
        self.name = name

    def parse_markdown_tables(self, md_content: str) -> List[Dict[str, Any]]:
        """Bóc tách tất cả các bảng biểu Markdown thành dữ liệu dạng danh sách dòng."""
        tables = []
        table_blocks = re.findall(r'(\|[^\n]+\|\n\|[-:| ]+\|\n(?:\|[^\n]+\|\n?)+)', md_content)

        for blk in table_blocks:
            lines = [ln.strip() for ln in blk.strip().split('\n') if ln.strip()]
            if len(lines) < 3:
                continue

            # Header
            headers = [c.strip() for c in lines[0].split('|')[1:-1]]
            rows = []
            for r_line in lines[2:]:
                cells = [c.strip() for c in r_line.split('|')[1:-1]]
                if len(cells) == len(headers):
                    rows.append(dict(zip(headers, cells)))

            if rows:
                tables.append({
                    "headers": headers,
                    "row_count": len(rows),
                    "rows": rows
                })

        return tables

    def extract_technical_specifications(self, md_content: str) -> Dict[str, Any]:
        """Trích xuất các chỉ tiêu kỹ thuật cốt lõi từ văn bản thiết kế."""
        specs = {
            "concrete_grades": [],
            "steel_grades": [],
            "standards_cited": [],
            "key_parameters": {}
        }

        # 1. Tìm mác bê tông
        concrete_matches = re.findall(r'\b(C10|C15|C20|C25|C30|C35|C40|C45|C50|M150|M200|M250|M300|M350|M400|M450|M500)\b', md_content)
        specs["concrete_grades"] = sorted(list(set(concrete_matches)))

        # 2. Tìm mác thép
        steel_matches = re.findall(r'\b(CB240-T|CB300-V|CB400-V|CB500-V|ASTM A416|Gr270|Gr 270)\b', md_content)
        specs["steel_grades"] = sorted(list(set(steel_matches)))

        # 3. Tìm các tiêu chuẩn viện dẫn
        std_matches = re.findall(r'\b(TCVN\s+\d+[:\-\d]*|QCVN\s+\d+[:\-\d]*\/[A-Z]+|ASTM\s+[A-Z0-9]+|AASHTO\s+[A-Z0-9\-]+)\b', md_content)
        specs["standards_cited"] = sorted(list(set(std_matches)))

        # 4. Tìm các thông số nhịp và cọc
        span_match = re.search(r'(\d+[\.\d]*\s*\+\s*\d+[\.\d]*\s*\+\s*\d+[\.\d]*)', md_content)
        if span_match:
            specs["key_parameters"]["span_schema"] = span_match.group(1)

        pile_match = re.search(r'(\d+)\s*cọc\s*khoan\s*nhồi\s*D\s*=?\s*(\d+)', md_content, re.IGNORECASE)
        if pile_match:
            specs["key_parameters"]["piles_count"] = int(pile_match.group(1))
            specs["key_parameters"]["pile_diameter_mm"] = int(pile_match.group(2))

        return specs

    def process_file(self, file_path: str) -> Dict[str, Any]:
        """Xử lý toàn diện một tệp Markdown."""
        print(f"[{self.name}] Đang phân tích hồ sơ Markdown: {os.path.basename(file_path)}")
        if not os.path.exists(file_path):
            return {"status": "FILE_NOT_FOUND", "file_path": file_path}

        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        tables = self.parse_markdown_tables(content)
        tech_specs = self.extract_technical_specifications(content)

        return {
            "file_name": os.path.basename(file_path),
            "file_size_chars": len(content),
            "tables_count": len(tables),
            "tables": tables,
            "technical_specs": tech_specs,
            "status": "PROCESSED"
        }
