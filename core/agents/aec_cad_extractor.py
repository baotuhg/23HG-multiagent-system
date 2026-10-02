# -*- coding: utf-8 -*-
"""
TÁC TỬ: AEC_CAD_EXTRACTOR (CHUYÊN GIA BÓC TÁCH BẢN VẼ CAD DWG/DXF & GIAO THỨC MCP)
==================================================================================
KIẾN TRÚC HỆ THỐNG MCP 3 THÀNH PHẦN (3-TIER MCP ARCHITECTURE):

[ Giao diện AI: Claude Desktop / Windsurf / Cursor / Antigravity ]
                               |
                               ▼ (Giao thức MCP: cad-mcp / autocad-mcp)
               [ MCP Server điều khiển AutoCAD (Local) ]
                               |
                               ▼ (API / COM Interop: win32com / ezdxf)
      [ Bản vẽ DWG trong AutoCAD ] <---> [ Hồ sơ thiết kế (Excel/PDF) ]

Chức năng cốt lõi:
1. Tầng Giao diện AI: Tiếp nhận yêu cầu bóc tách hình học, diện tích, cốt thép từ User.
2. Tầng MCP Server: Điều phối qua giao thức MCP (smart_cad_command, batch_execute).
3. Tầng COM Interop & Tệp bản vẽ:
   - Kết nối trực tiếp phiên AutoCAD đang chạy qua COM Automation (win32com.client).
   - Quét đệ quy toàn bộ thư mục hồ sơ bản vẽ (.dwg, .dxf) khi người dùng bổ sung thêm bản vẽ.
   - Giải mã TCVN3 / VNI sang Unicode UTF-8 chuẩn.
   - Đo bóc diện tích mặt cắt đào đắp (Shoelace algorithm).
   - Trích xuất bảng thống kê thép BBS và đối soát 2 chiều với Hồ sơ thiết kế (Excel BoQ / Markdown).
"""

import os
import re
import glob
from typing import Dict, List, Any, Tuple, Optional

# Thử nạp win32com cho COM Automation với AutoCAD
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# Thử nạp ezdxf để đọc trực tiếp file DXF nếu có
try:
    import ezdxf
    HAS_EZDXF = True
except ImportError:
    HAS_EZDXF = False


class TCVN3Decoder:
    """Bộ giải mã font tiếng Việt TCVN3 (.VnTime, .VnArial) trong bản vẽ AutoCAD sang Unicode UTF-8."""
    
    CHAR_MAP = {
        'µ': 'à', '¸': 'á', '¶': 'ả', '·': 'ã', '¹': 'ạ',
        '¨': 'ă', '¾': 'ắ', '»': 'ằ', '¼': 'ẳ', '½': 'ẵ', 'Æ': 'ặ',
        '©': 'â', 'Ê': 'ấ', 'Ç': 'ầ', 'È': 'ẩ', 'É': 'ẫ', 'Ë': 'ậ',
        'e': 'e', 'Ì': 'è', 'Ð': 'é', 'Î': 'ẻ', 'Ï': 'ẽ', 'Ñ': 'ẹ',
        'ª': 'ê', 'Õ': 'ế', 'Ò': 'ề', 'Ó': 'ể', 'Ô': 'ễ', 'Ö': 'ệ',
        'i': 'i', '×': 'ì', 'Ý': 'í', 'Ø': 'ỉ', 'Ü': 'ĩ', 'Þ': 'ị',
        'o': 'o', 'ß': 'ò', 'ã': 'ó', 'á': 'ỏ', 'â': 'õ', 'ä': 'ọ',
        '«': 'ô', 'è': 'ố', 'å': 'ồ', 'æ': 'ổ', 'ç': 'ỗ', 'é': 'ộ',
        '¬': 'ơ', 'í': 'ớ', 'ê': 'ờ', 'ë': 'ở', 'ì': 'ỡ', 'î': 'ợ',
        'u': 'u', 'ï': 'ù', 'ó': 'ú', 'ñ': 'ủ', 'ò': 'ũ', 'ô': 'ụ',
        '®': 'ư', 'ứ': 'ứ', 'ừ': 'ừ', 'ử': 'ử', 'ữ': 'ữ', 'ự': 'ự',
        'y': 'y', 'ú': 'ỳ', 'ý': 'ý', 'û': 'ỷ', 'ü': 'ỹ', 'þ': 'ỵ',
        '®': 'đ', '§': 'Đ'
    }

    @classmethod
    def decode(cls, text: str) -> str:
        """Chuyển đổi xâu ký tự TCVN3 sang Unicode tiếng Việt chuẩn."""
        if not text:
            return ""
        result = text
        for k, v in cls.CHAR_MAP.items():
            result = result.replace(k, v)
        return result


class AutoCADCOMConnector:
    """Tầng 3: Điều khiển AutoCAD trực tiếp qua COM Interop API (AutoCAD.Application)."""

    def __init__(self):
        self.acad = None
        self.doc = None
        self.is_connected = False

    def connect(self) -> bool:
        """Kết nối tới phiên làm việc AutoCAD đang mở trên máy tính."""
        if not HAS_WIN32COM:
            return False
        try:
            # Thử kết nối phiên AutoCAD đang chạy
            self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            self.is_connected = True
            print(f"[*] [AutoCAD COM] Đã kết nối thành công tới AutoCAD: {self.acad.Caption}")
            if self.doc:
                print(f"[*] [AutoCAD COM] Bản vẽ đang mở: {self.doc.Name}")
            return True
        except Exception:
            self.is_connected = False
            return False

    def get_open_drawings(self) -> List[str]:
        """Lấy danh sách tất cả các bản vẽ DWG đang mở trong AutoCAD."""
        if not self.is_connected or not self.acad:
            return []
        try:
            drawings = []
            for doc in self.acad.Documents:
                drawings.append(doc.Name)
            return drawings
        except Exception as e:
            print(f"[!] [AutoCAD COM] Lỗi truy vấn Documents: {e}")
            return []

    def read_modelspace_entities(self) -> Dict[str, Any]:
        """Đọc danh sách Text, Block, Polyline từ ModelSpace của bản vẽ đang mở."""
        if not self.is_connected or not self.doc:
            return {"status": "DISCONNECTED"}
        
        entities_summary = {
            "drawing_name": self.doc.Name,
            "texts": [],
            "blocks": [],
            "polylines_count": 0,
            "layers": []
        }

        try:
            # Đọc danh sách Layer
            for layer in self.doc.Layers:
                entities_summary["layers"].append(layer.Name)

            # Đọc các đối tượng trong ModelSpace
            ms = self.doc.ModelSpace
            for entity in ms:
                entity_name = entity.EntityName
                if entity_name in ["AcDbText", "AcDbMText"]:
                    text_str = entity.TextString
                    decoded = TCVN3Decoder.decode(text_str)
                    entities_summary["texts"].append(decoded)
                elif entity_name == "AcDbBlockReference":
                    entities_summary["blocks"].append(entity.Name)
                elif entity_name in ["AcDbPolyline", "AcDb2dPolyline"]:
                    entities_summary["polylines_count"] += 1

        except Exception as e:
            print(f"[!] [AutoCAD COM] Lỗi đọc ModelSpace: {e}")

        return entities_summary


# ─────────────────────────────────────────────────────────────────────────────
# ĐỌC ĐA TUYẾN KHÉP KÍN TỪ FILE DXF (không cần AutoCAD)
# ─────────────────────────────────────────────────────────────────────────────

# Mã $INSUNITS của DXF → tên đơn vị (chỉ các đơn vị hay gặp trong hồ sơ xây dựng)
DXF_UNITS = {0: "không khai báo", 1: "inch", 2: "feet", 4: "mm", 5: "cm", 6: "m"}


def _shoelace(vertices: List[Tuple[float, float]]) -> float:
    n = len(vertices)
    if n < 3:
        return 0.0
    total = 0.0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        total += x1 * y2 - x2 * y1
    return abs(total) / 2.0


def _read_dxf_ascii(path: str) -> Tuple[int, List[Dict[str, Any]]]:
    """Bộ đọc DXF ASCII tối giản: LWPOLYLINE và POLYLINE/VERTEX (2D) trong mục ENTITIES."""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        lines = [ln.rstrip("\r\n") for ln in f]
    if len(lines) % 2:
        lines.append("")
    pairs = [(lines[i].strip(), lines[i + 1]) for i in range(0, len(lines), 2)]

    insunits, polylines = 0, []
    section, current, in_polyline = None, None, None
    for idx, (code, value) in enumerate(pairs):
        if code == "0" and value.strip() == "SECTION" and idx + 1 < len(pairs):
            section = pairs[idx + 1][1].strip()
            continue
        if section == "HEADER" and code == "9" and value.strip() == "$INSUNITS" and idx + 1 < len(pairs):
            try:
                insunits = int(pairs[idx + 1][1].strip())
            except ValueError:
                pass
            continue
        if section != "ENTITIES":
            continue
        if code == "0":
            kind = value.strip()
            if current is not None and current["type"] == "LWPOLYLINE":
                polylines.append(current)
            current = None
            if kind == "LWPOLYLINE":
                current = {"type": "LWPOLYLINE", "layer": "0", "closed": False, "vertices": []}
            elif kind == "POLYLINE":
                in_polyline = {"type": "POLYLINE", "layer": "0", "closed": False, "vertices": []}
            elif kind == "VERTEX" and in_polyline is not None:
                current = {"type": "VERTEX", "target": in_polyline}
            elif kind == "SEQEND" and in_polyline is not None:
                polylines.append(in_polyline)
                in_polyline = None
            elif kind == "ENDSEC":
                section = None
            continue
        target = current if current is not None and current["type"] == "LWPOLYLINE" else None
        if current is None and in_polyline is not None:
            target = in_polyline
        if current is not None and current["type"] == "VERTEX":
            v = current["target"]["vertices"]
            if code == "10":
                v.append([float(value), 0.0])
            elif code == "20" and v:
                v[-1][1] = float(value)
            continue
        if target is None:
            continue
        if code == "8":
            target["layer"] = value.strip()
        elif code == "70":
            target["closed"] = bool(int(value.strip() or 0) & 1)
        elif code == "10" and target["type"] == "LWPOLYLINE":
            target["vertices"].append([float(value), 0.0])
        elif code == "20" and target["type"] == "LWPOLYLINE" and target["vertices"]:
            target["vertices"][-1][1] = float(value)
    if current is not None and current["type"] == "LWPOLYLINE":
        polylines.append(current)
    return insunits, [{"layer": p["layer"], "closed": p["closed"],
                       "vertices": [tuple(v) for v in p["vertices"]]} for p in polylines]


def _read_dxf_ezdxf(path: str) -> Tuple[int, List[Dict[str, Any]]]:
    doc = ezdxf.readfile(path)
    insunits = int(doc.header.get("$INSUNITS", 0) or 0)
    polylines = []
    for e in doc.modelspace().query("LWPOLYLINE POLYLINE"):
        if e.dxftype() == "LWPOLYLINE":
            pts = [(float(x), float(y)) for x, y in e.get_points("xy")]
        else:
            pts = [(float(v.dxf.location.x), float(v.dxf.location.y)) for v in e.vertices]
        polylines.append({"layer": e.dxf.layer, "closed": bool(e.is_closed), "vertices": pts})
    return insunits, polylines


def read_dxf_closed_areas(path: str) -> Dict[str, Any]:
    """
    Đọc mọi đa tuyến KHÉP KÍN trong file DXF và tính diện tích (Shoelace), gom theo layer.
    Dùng ezdxf nếu đã cài, ngược lại dùng bộ đọc DXF ASCII tích hợp (không đọc được DXF nhị phân / DWG).

    Diện tích tính theo ĐƠN VỊ BẢN VẼ bình phương — xem 'units' ($INSUNITS) để quy đổi.
    Đa tuyến có cung tròn (bulge) được tính theo dây cung → sai số nhỏ; đa tuyến hở bị bỏ qua.
    """
    if not path.lower().endswith(".dxf"):
        raise ValueError(f"Chỉ đọc được file .dxf (DWG cần AutoCAD COM hoặc chuyển sang DXF): {path}")
    reader = "ezdxf" if HAS_EZDXF else "ascii"
    insunits, polylines = (_read_dxf_ezdxf if HAS_EZDXF else _read_dxf_ascii)(path)

    by_layer: Dict[str, Dict[str, Any]] = {}
    open_count = 0
    for pl in polylines:
        if not pl["closed"]:
            open_count += 1
            continue
        area = _shoelace(pl["vertices"])
        entry = by_layer.setdefault(pl["layer"], {"count": 0, "total_area": 0.0, "areas": []})
        entry["count"] += 1
        entry["total_area"] += area
        entry["areas"].append(round(area, 6))
    for entry in by_layer.values():
        entry["total_area"] = round(entry["total_area"], 6)
    return {
        "file_path": path,
        "reader": reader,
        "units": DXF_UNITS.get(insunits, f"mã {insunits}"),
        "closed_polylines": sum(e["count"] for e in by_layer.values()),
        "open_polylines_skipped": open_count,
        "layers": by_layer,
    }


class AECCadExtractor:
    """
    TÁC TỬ AEC CAD EXTRACTOR
    Hỗ trợ Kiến trúc Hệ thống MCP 3 Thành Phần:
    - Giao tiếp MCP Server (cad-mcp, autocad-mcp)
    - Tương tác trực tiếp AutoCAD qua COM Interop
    - Quét đệ quy thư mục hồ sơ bản vẽ (hỗ trợ hàng chục đến hàng trăm file DWG/DXF)
    - Đối soát 2 chiều với Hồ sơ thiết kế (Excel BoQ / Markdown)
    """

    def __init__(self, name: str = "aec_cad_extractor"):
        self.name = name
        self.supported_extensions = [".dwg", ".dxf"]
        self.com_connector = AutoCADCOMConnector()

    def check_autocad_connection(self) -> bool:
        """Kiểm tra và kết nối với phiên AutoCAD đang mở."""
        return self.com_connector.connect()

    def extract_cross_section_area(self, vertices: List[Tuple[float, float]]) -> float:
        """
        Tính diện tích đa giác khép kín từ danh sách đỉnh theo thuật toán Shoelace.
        A = 0.5 * |sum(x_i * y_{i+1} - x_{i+1} * y_i)|
        """
        n = len(vertices)
        if n < 3:
            return 0.0
        
        area = 0.0
        for i in range(n):
            j = (i + 1) % n
            area += vertices[i][0] * vertices[j][1]
            area -= vertices[j][0] * vertices[i][1]
        
        return abs(area) / 2.0

    def parse_rebar_table_from_text(self, text_lines: List[str]) -> List[Dict[str, Any]]:
        """Bóc tách các dòng bảng thống kê thép (BBS) trích từ CAD."""
        rebars = []
        pattern = re.compile(r'([A-Za-z0-9\-_]+)\s+(\d{1,2})\s+(\d+)\s+(\d+)')
        
        for line in text_lines:
            line_decoded = TCVN3Decoder.decode(line.strip())
            match = pattern.search(line_decoded)
            if match:
                mark, dia, length, qty = match.groups()
                rebars.append({
                    "mark": mark,
                    "diameter_mm": int(dia),
                    "length_mm": float(length),
                    "quantity": int(qty),
                    "raw_text": line_decoded
                })
        return rebars

    def classify_drawing_component(self, file_path: str) -> Dict[str, Any]:
        """Phân loại hạng mục kết cấu từ tên tệp và đường dẫn bản vẽ DWG."""
        base_name = os.path.basename(file_path).lower()
        full_path = file_path.lower()

        classification = {
            "file_name": os.path.basename(file_path),
            "file_path": file_path,
            "category": "KẾT CẤU CHUNG",
            "component": "Chưa phân loại",
            "concrete_grade": "C30",
            "steel_types": ["CB400-V", "CB240-T"],
            "has_rebar_bbs": False
        }

        # 1. Cọc khoan nhồi
        if "1200" in base_name or "coc" in base_name or "cọc" in base_name:
            classification["category"] = "KẾT CẤU MÓNG CỌC"
            classification["component"] = "Cọc khoan nhồi D1200mm"
            classification["concrete_grade"] = "C30"
            classification["has_rebar_bbs"] = True
            classification["diameter_mm"] = 1200
            classification["sonic_tubes"] = 4

        # 2. Dầm Super-T / Dầm chủ
        elif "super" in base_name or "dam" in base_name or "dầm" in base_name:
            classification["category"] = "KẾT CẤU NHỊP"
            classification["component"] = "Dầm chủ Super-T L=38.2m"
            classification["concrete_grade"] = "C45"
            classification["has_rebar_bbs"] = True
            classification["span_length_m"] = 38.2
            classification["prestress_cable_strands"] = 44

        # 3. Mố cầu M1 / M2
        elif "mo " in base_name or "mố" in base_name or base_name.startswith("mo") or "\\01. mo" in full_path:
            classification["category"] = "KẾT CẤU MỐ CẦU"
            classification["component"] = "Mố M1 / M2 chân dê & chữ U"
            classification["concrete_grade"] = "C30"
            classification["has_rebar_bbs"] = True

        # 4. Trụ cầu T1 / T2
        elif "tru" in base_name or "trụ" in base_name or "\\02. tru" in full_path:
            classification["category"] = "KẾT CẤU TRỤ CẦU"
            classification["component"] = "Trụ T1 / T2 thân đặc & xà mũ C35"
            classification["concrete_grade"] = "C30/C35"
            classification["has_rebar_bbs"] = True

        # 5. Bản mặt cầu, Dầm ngang, Gờ lan can, Bản quá độ
        elif "bmc" in base_name or "ban mat cau" in base_name or "mặt cầu" in base_name:
            classification["category"] = "KẾT CẤU MẶT CẦU"
            classification["component"] = "Bản mặt cầu liên tục nhiệt"
            classification["concrete_grade"] = "C35"
            classification["has_rebar_bbs"] = True
        elif "dam ngang" in base_name or "dầm ngang" in base_name:
            classification["category"] = "KẾT CẤU MẶT CẦU"
            classification["component"] = "Dầm ngang mố & trụ"
            classification["concrete_grade"] = "C35"
            classification["has_rebar_bbs"] = True
        elif "lan can" in base_name:
            classification["category"] = "PHỤ TRỢ MẶT CẦU"
            classification["component"] = "Gờ lan can bê tông & tay vịn thép mạ kẽm"
            classification["concrete_grade"] = "C25"
            classification["has_rebar_bbs"] = True
        elif "qua do" in base_name or "quá độ" in base_name:
            classification["category"] = "ĐẦU CẦU"
            classification["component"] = "Bản quá độ 2 đầu mố L=8.0m"
            classification["concrete_grade"] = "C25"
            classification["has_rebar_bbs"] = True

        # 6. Đào đắp, trắc dọc, trắc ngang
        elif "dao" in base_name or "dap" in base_name or "trac doc" in base_name or "trac ngang" in base_name:
            classification["category"] = "NỀN ĐƯỜNG & ĐÀO ĐẮP"
            classification["component"] = "Trắc dọc / Trắc ngang đào đắp nền móng"
            classification["has_rebar_bbs"] = False

        return classification

    def scan_drawings_folder(self, folder_path: str) -> Dict[str, Any]:
        """
        Quét đệ quy toàn bộ thư mục hồ sơ bản vẽ khi người dùng bổ sung thêm bản vẽ.
        Tự động nhận diện cấu trúc, phân loại WBS và trích xuất danh mục.
        """
        print(f"[{self.name}] Đang quét đệ quy thư mục bản vẽ CAD: {folder_path}")
        if not os.path.exists(folder_path):
            return {
                "status": "NOT_FOUND",
                "folder_path": folder_path,
                "total_drawings": 0,
                "drawings": []
            }

        # Tìm tất cả file .dwg và .dxf
        dwg_files = []
        for root, _, files in os.walk(folder_path):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in self.supported_extensions:
                    dwg_files.append(os.path.join(root, f))

        dwg_files.sort()
        classified_drawings = []
        category_summary = {}

        for fp in dwg_files:
            cls_info = self.classify_drawing_component(fp)
            classified_drawings.append(cls_info)
            cat = cls_info["category"]
            category_summary[cat] = category_summary.get(cat, 0) + 1

        result = {
            "status": "SUCCESS",
            "folder_path": folder_path,
            "total_drawings": len(dwg_files),
            "categories_summary": category_summary,
            "drawings": classified_drawings
        }

        print(f"[{self.name}] Hoàn tất quét: Phát hiện {len(dwg_files)} bản vẽ DWG/DXF thuộc {len(category_summary)} phân loại WBS:")
        for cat, count in category_summary.items():
            print(f"   • {cat}: {count} bản vẽ")

        return result

    def reconcile_with_design_documents(self, cad_summary: Dict[str, Any], excel_data: Dict[str, Any], md_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Đối soát 2 chiều [ Bản vẽ CAD ] <---> [ Hồ sơ thiết kế (Excel / Markdown) ].

        Chỉ so sánh những chỉ tiêu mà CẢ HAI phía đều có số liệu thật:
          MATCHED     — hai phía khớp
          MISMATCH    — hai phía lệch nhau (cần kỹ sư xem lại)
          NOT_CHECKED — một phía chưa có dữ liệu → KHÔNG được coi là khớp
        """
        specs = (md_data or {}).get("technical_specs", {}) or {}
        excel_data = excel_data or {}
        cad_summary = cad_summary or {}

        def as_set(values):
            if not values:
                return None
            return {str(v).strip().upper().replace(" ", "") for v in values if str(v).strip()}

        def check(item, cad_value, doc_value, doc_label):
            if cad_value is None or doc_value is None:
                missing = "bản vẽ CAD" if cad_value is None else doc_label
                return {"check_item": item,
                        "cad_value": "—" if cad_value is None else self._fmt(cad_value),
                        "excel_value": "—" if doc_value is None else self._fmt(doc_value),
                        "status": "NOT_CHECKED",
                        "note": f"Chưa có dữ liệu từ {missing} — chưa đối soát"}
            return {"check_item": item, "cad_value": self._fmt(cad_value), "excel_value": self._fmt(doc_value),
                    "status": "MATCHED" if cad_value == doc_value else "MISMATCH", "note": ""}

        return [
            check("Số lượng bản vẽ thiết kế",
                  cad_summary.get("total_drawings") if cad_summary.get("total_drawings") else None,
                  excel_data.get("drawing_list_count"), "danh mục bản vẽ trong Excel"),
            check("Mác bê tông", as_set(cad_summary.get("concrete_grades")),
                  as_set(specs.get("concrete_grades")), "thuyết minh (Markdown)"),
            check("Mác thép", as_set(cad_summary.get("steel_grades")),
                  as_set(specs.get("steel_grades")), "thuyết minh (Markdown)"),
        ]

    @staticmethod
    def _fmt(value) -> str:
        if isinstance(value, (set, list, tuple)):
            return ", ".join(sorted(str(v) for v in value))
        return str(value)

    def process_drawing(self, file_path: str) -> Dict[str, Any]:
        """
        Xử lý một tệp bản vẽ: phân loại WBS theo tên file; với .dxf đọc thêm diện tích
        các đa tuyến khép kín theo layer (read_dxf_closed_areas). DWG cần AutoCAD (COM).
        """
        print(f"[{self.name}] Đang phân tích chi tiết bản vẽ CAD: {os.path.basename(file_path)}")
        cls_info = self.classify_drawing_component(file_path)
        if file_path.lower().endswith(".dxf"):
            try:
                cls_info["dxf_areas"] = read_dxf_closed_areas(file_path)
            except Exception as e:  # file hỏng / DXF nhị phân — vẫn trả phân loại
                cls_info["dxf_error"] = str(e)
        cls_info["status"] = "PROCESSED"
        return cls_info
