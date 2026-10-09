# -*- coding: utf-8 -*-
"""
TÁC TỬ: AEC_DATA_AGGREGATOR (TRỌNG TÀI HỢP NHẤT & ĐỐI CHIẾU DỮ LIỆU ĐA PHƯƠNG THỨC)
Vai trò trong Hệ thống Đa tác tử AEC:
- Hạt nhân hợp nhất dữ liệu (Data Fusion & Master Synthesizer).
- Thu nhận dữ liệu đã bóc tách từ 3 tác tử đầu vào:
  1. aec_cad_extractor (Bản vẽ CAD DWG/DXF)
  2. aec_office_extractor (Bảng tính Excel & Word)
  3. aec_markdown_ingestor (Hồ sơ thiết kế Markdown & Thuyết minh)
- Thực hiện ĐỐI CHIẾU CHÉO ĐA PHƯƠNG THỨC (Cross-Modal Reconciliation):
  * Phát hiện sự lệch pha giữa Bản vẽ CAD vs Bảng tính Excel vs Thuyết minh Markdown.
  * Gắn cờ cảnh báo nếu có sự sai khác về số lượng cọc, chiều dài nhịp hoặc khối lượng cốt thép.
- Chuẩn hóa dữ liệu thành định dạng Canonical Data (Single Source of Truth).
- Đồng bộ trực tiếp vào Blackboard `PROJECT_STATE.json` để bàn giao cho các Agent nghiệp vụ tiếp theo.
"""

import os
import json
from typing import Dict, List, Any

from core.agents.aec_cad_extractor import AECCadExtractor
from core.agents.aec_office_extractor import AECOfficeExtractor
from core.agents.aec_markdown_ingestor import AECMarkdownIngestor

class AECDataAggregator:
    """Tác tử hợp nhất, đối chiếu chéo và chuẩn hóa dữ liệu toàn hệ thống."""

    def __init__(self, name: str = "aec_data_aggregator"):
        self.name = name
        self.cad_extractor = AECCadExtractor()
        self.office_extractor = AECOfficeExtractor()
        self.md_ingestor = AECMarkdownIngestor()
        self.discrepancies = []
        self.canonical_state = {}

    def reconcile_sources(self, cad_data: Dict[str, Any], office_data: Dict[str, Any], md_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Đối chiếu chéo số liệu giữa CAD, Office và Markdown để phát hiện xung đột."""
        print(f"[{self.name}] Bắt đầu đối chiếu chéo dữ liệu đa phương thức (Cross-modal Reconciliation)...")
        discrepancies = []

        cad = cad_data.get("summary_quantities", {})
        md = md_data.get("technical_specs", {}).get("key_parameters", {})
        self.checks_performed = 0
        piles_md, piles_cad = md.get("piles_count"), cad.get("piles_d1200_count")
        if piles_md is not None and piles_cad is not None:
            self.checks_performed += 1
            if piles_md != piles_cad:
                discrepancies.append({"parameter": "Số lượng cọc khoan nhồi D1200",
                    "severity": "WARNING", "source_markdown": piles_md, "source_cad": piles_cad,
                    "resolution": "Cần kỹ sư đối chiếu nguồn đã duyệt"})

        self.discrepancies = discrepancies
        if not discrepancies:
            print(f"[{self.name}] Đã đối chiếu {self.checks_performed} chỉ tiêu; các chỉ tiêu khác chưa kiểm tra.")
        else:
            print(f"[{self.name}] [!] Phát hiện {len(discrepancies)} điểm cần rà soát!")

        return discrepancies

    def synthesize_to_project_state(self, project_dir: str, output_state_path: str = None,
                                    *, cad_data=None, office_data=None, md_data=None) -> Dict[str, Any]:
        """Lưu dữ liệu đã trích xuất và phạm vi đối chiếu; không tạo số liệu thay thế."""
        from copy import deepcopy
        if not os.path.isdir(project_dir):
            raise ValueError(f"Không tìm thấy thư mục dự án: {project_dir}")
        sources = {"cad": cad_data, "office": office_data, "markdown": md_data}
        if not any(sources.values()):
            raise ValueError("Chưa có dữ liệu đã trích xuất để hợp nhất; cần CAD/Office/Markdown thật")
        if any(v is not None and not isinstance(v, dict) for v in sources.values()):
            raise ValueError("Dữ liệu trích xuất phải là dictionary")
        discrepancies = self.reconcile_sources(cad_data or {}, office_data or {}, md_data or {})
        status = "MISMATCH" if discrepancies else ("MATCHED_CHECKED_FIELDS" if self.checks_performed else "NOT_CHECKED")
        state = {
            "meta": {"version": "3.0.0", "source_dir": os.path.abspath(project_dir),
                     "cross_check_status": status, "checks_performed": self.checks_performed,
                     "unchecked": "Chỉ đối chiếu số lượng cọc khi cả CAD và Markdown có dữ liệu"},
            "sources": deepcopy({k: v for k, v in sources.items() if v is not None}),
            "discrepancies": deepcopy(discrepancies),
        }
        if output_state_path:
            target = os.path.abspath(output_state_path)
            os.makedirs(os.path.dirname(target), exist_ok=True)
            tmp = f"{target}.{os.getpid()}.tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
            os.replace(tmp, target)
        self.canonical_state = state
        return state
