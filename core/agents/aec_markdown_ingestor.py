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

    # Ký hiệu cấu kiện + số lượng (+ tổng chiều dài) trên bản vẽ: "MT1A (SL=2)", "C-03 (SL=32)",
    # "D2-X-B*-01 (SL=2; L=20220)", "GM (SL=1; L=530800)". Ký hiệu phải bắt đầu bằng chữ in, ≥ 2 ký tự.
    _SCHEDULE = re.compile(
        r'(?<![\w*\-])([A-Z]{1,3}\d{0,2}[A-Z]?(?:\*?-[A-Z0-9]{1,3}\*?){0,4})'
        r'\s*\(\s*SL\s*[=:]\s*(\d+)\s*(?:;\s*L\s*=\s*(\d+))?')
    _PAGE = re.compile(r'^## \[Trang (\d+)\]', re.M)

    def extract_technical_specifications(self, md_content: str) -> Dict[str, Any]:
        """Trích xuất các chỉ tiêu kỹ thuật cốt lõi từ văn bản thiết kế (cầu hoặc nhà)."""
        specs = {
            "concrete_grades": [],      # mác M (vd M250)
            "concrete_classes": [],     # cấp độ bền B (vd B20)
            "concrete_class_grade_pairs": {},   # "B20" -> "M250" khi ghi dạng "B20 (MÁC M250#)"
            "steel_grades": [],
            "steel_strengths": [],      # "RS=2800 DAN/CM2"
            "standards_cited": [],
            "key_parameters": {}
        }

        # 1. Bê tông: mác M, cấp độ bền B (chỉ nhận B đứng sau "cấp ... bền" hoặc ngay trước "(MÁC M...)")
        specs["concrete_grades"] = sorted(set(re.findall(
            r'\b(C(?:10|15|20|25|30|35|40|45|50)|M(?:100|150|200|250|300|350|400|450|500))\b', md_content)))
        classes = set(re.findall(r'C[ẤA]P\s+(?:Đ[ỘO]\s+)?B\w{1,2}N(?:\s+N\w{1,2}N)?\s+B(\d{1,2}(?:[,.]5)?)\b',
                                 md_content, re.I))
        for b_cls, m_grade in re.findall(r'\bB(\d{1,2}(?:[,.]5)?)\s*\(\s*M[ÁA]C\s*M(\d{3})', md_content, re.I):
            classes.add(b_cls)
            specs["concrete_class_grade_pairs"][f"B{b_cls}"] = f"M{m_grade}"
        specs["concrete_classes"] = sorted(f"B{c}" for c in classes)

        # 2. Thép: mác theo TCVN 1651 mới (CB...), JIS (SD...), ASTM; cường độ ghi trực tiếp (Rs = ...)
        specs["steel_grades"] = sorted(set(re.findall(
            r'\b(CB\s?\d{3}-[TV]|SD\s?\d{3}|ASTM A416|Gr\s?270)\b', md_content)))
        specs["steel_strengths"] = sorted(set(
            re.sub(r'\s+', '', m).upper() for m in re.findall(
                r'\bR[sS]\s*=\s*\d{3,4}\s*(?:daN/cm2|DAN/CM2|MPa|N/mm2)', md_content, re.I)))

        # 3. Tiêu chuẩn: số hiệu + năm (nếu có). Bỏ bản không năm khi đã có bản có năm của cùng số hiệu.
        found = {}
        # cho phép số hiệu bị tách sang dòng sau ("TCVN\n- 1651:2008" — thường gặp ở bản OCR)
        for org, num, year in re.findall(r'\b(TCVN|QCVN|TCXDVN)(?:\s|-\s)*(\d{2,5})(?:\s*[:\-]\s*(\d{2,4})(?!\d))?', md_content):
            found.setdefault((org, num), set()).add(year)
        stds = []
        for (org, num), years in found.items():
            with_year = sorted(y for y in years if y)
            stds += [f"{org} {num}:{y}" if len(y) == 4 else f"{org} {num}-{y}" for y in with_year] or [f"{org} {num}"]
        stds += sorted(set(re.findall(r'\b(ASTM\s+[A-Z]\d{1,4}|AASHTO\s+[A-Z]{1,2}\s?\d{1,4})\b', md_content)))
        specs["standards_cited"] = sorted(set(stds))

        # 4. Sơ đồ nhịp (cầu): chỉ nhận khi có chữ "nhịp" ngay trước — tránh bắt nhầm kích thước kiến trúc
        span_match = re.search(r'nh[ịi]p[^\n\d]{0,40}(\d+(?:[.,]\d+)?\s*m?\s*\+\s*\d+(?:[.,]\d+)?\s*m?'
                               r'(?:\s*\+\s*\d+(?:[.,]\d+)?\s*m?)+)', md_content, re.I)
        if span_match:
            specs["key_parameters"]["span_schema"] = span_match.group(1).strip()

        pile_match = re.search(r'(\d+)\s*cọc\s*khoan\s*nhồi\s*D\s*=?\s*(\d+)', md_content, re.IGNORECASE)
        if pile_match:
            specs["key_parameters"]["piles_count"] = int(pile_match.group(1))
            specs["key_parameters"]["pile_diameter_mm"] = int(pile_match.group(2))

        return specs

    def extract_element_schedule(self, md_content: str) -> List[Dict[str, Any]]:
        """Ký hiệu cấu kiện, số lượng (SL), tổng chiều dài L (mm, nếu có) và trang chứa."""
        pages = [(m.start(), int(m.group(1))) for m in self._PAGE.finditer(md_content)]
        out = []
        for m in self._SCHEDULE.finditer(md_content):
            sym = m.group(1)
            if len(sym) < 2:
                continue
            page = max((n for pos, n in pages if pos <= m.start()), default=None)
            out.append({"symbol": sym, "group": self._group(sym), "count": int(m.group(2)),
                        "length_mm": int(m.group(3)) if m.group(3) else None, "page": page})
        return out

    @staticmethod
    def _group(symbol: str) -> str:
        """Nhóm cấu kiện: phần trước dấu '-' ("D2-X-A-01" → "D2", "C-03" → "C");
        không có '-' thì bỏ số/chữ đuôi ("MT1A" → "MT", "GM" → "GM")."""
        if "-" in symbol:
            return symbol.split("-")[0]
        return re.sub(r"\d+[A-Z]?$", "", symbol) or symbol

    @staticmethod
    def cross_check_schedule(schedule: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Đối chiếu nội bộ bảng cấu kiện: số móng vs số cột; các tầng dầm giống nhau (tầng điển hình)."""
        totals: Dict[str, int] = {}
        lengths: Dict[str, float] = {}
        for r in schedule:
            totals[r["group"]] = totals.get(r["group"], 0) + r["count"]
            if r["length_mm"]:
                lengths[r["group"]] = lengths.get(r["group"], 0.0) + r["count"] * r["length_mm"] / 1000
        checks = {"totals": totals, "lengths_m": {k: round(v, 3) for k, v in lengths.items()}, "findings": []}
        if "MT" in totals and "C" in totals:
            same = totals["MT"] == totals["C"]
            checks["findings"].append({"check": "Số móng = số cột", "MT": totals["MT"], "C": totals["C"],
                                       "result": "KHỚP" if same else "LỆCH"})
        floors = sorted(g for g in totals if re.fullmatch(r"D\d", g))
        sig = {g: sorted((r["symbol"].split("-", 1)[1], r["count"], r["length_mm"])
                         for r in schedule if r["group"] == g) for g in floors}
        for a, b in zip(floors, floors[1:]):
            checks["findings"].append({"check": f"Dầm {a} so với {b}", "result": "GIỐNG" if sig[a] == sig[b] else "KHÁC",
                                       a: totals[a], b: totals[b]})
        return checks

    def process_file(self, file_path: str) -> Dict[str, Any]:
        """Xử lý toàn diện một tệp Markdown."""
        print(f"[{self.name}] Đang phân tích hồ sơ Markdown: {os.path.basename(file_path)}")
        if not os.path.exists(file_path):
            return {"status": "FILE_NOT_FOUND", "file_path": file_path}

        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        tables = self.parse_markdown_tables(content)
        tech_specs = self.extract_technical_specifications(content)
        schedule = self.extract_element_schedule(content)

        return {
            "file_name": os.path.basename(file_path),
            "file_size_chars": len(content),
            "tables_count": len(tables),
            "tables": tables,
            "technical_specs": tech_specs,
            "element_schedule": schedule,
            "schedule_checks": self.cross_check_schedule(schedule),
            "status": "PROCESSED"
        }
