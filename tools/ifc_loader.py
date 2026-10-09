# -*- coding: utf-8 -*-
"""
IFC / BIM LOADER — BÓC TÁCH HÌNH HỌC BÊ TÔNG & CỐT THÉP 3D TỪ MÔ HÌNH OPENBIM (IFC)
Dành cho Hệ thống Multi-Agent AEC (23HG-multiagent-system).
Pure Python, Zero LLM — Tuân thủ chuẩn OpenBIM ISO 16739 (IFC2X3, IFC4, IFC4X3).

Khả năng:
1. Đọc và phân tích cấu trúc mô hình OpenBIM IFC: Dự án (IfcProject), Công trình (IfcBuilding),
   Tầng/Phân đoạn (IfcBuildingStorey).
2. Bóc tách khối lượng cấu kiện bê tông cốt thép:
   - Dầm (IfcBeam), Cột (IfcColumn), Móng (IfcFooting), Cọc khoan nhồi / Cọc ép (IfcPile),
     Bản mặt cầu / Sàn (IfcSlab), Mố trụ / Tường (IfcWall), Thanh kết cấu (IfcMember).
   - Trích xuất: Thể tích Bê tông (m3), Diện tích ván khuôn (m2), Kích thước Dài x Rộng x Cao (m),
     Mác bê tông thiết kế (C20/25, C30/37, C35/45, M300, M400, M500).
3. Bóc tách cốt thép 3D (3D Rebar Takeoff):
   - Đọc toàn bộ thực thể IfcReinforcingBar, IfcReinforcingElement, IfcReinforcingMesh.
   - Trích xuất: Ký hiệu thanh (Mark), Đường kính danh định Ø (mm), Chiều dài thanh L (mm),
     Số lượng thanh (Quantity), Mác thép (CB240-T, CB300-V, CB400-V, CB500-V), Khối lượng (kg).
4. Tích hợp trực tiếp vào Solver Cắt thép 1D:
   - Chuyển đổi thẳng thành danh sách `CutDemand` nạp vào `tools/cutting_stock_solver.py`
     để giải bài toán tối ưu cắt cây 11.7m, khép kín quy trình từ BIM sang Xưởng gia công!
"""

from __future__ import annotations
import math
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

try:
    import ifcopenshell
    import ifcopenshell.util.element
    import ifcopenshell.util.unit
    HAS_IFCOPENSHELL = True
except ImportError:
    HAS_IFCOPENSHELL = False

from tools.cutting_stock_solver import CutDemand


STANDARD_DIAMETERS = {6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40}
STEEL_DENSITY_KG_M3 = 7850.0  # Khối lượng riêng thép xây dựng (TCVN 1651:2018)


def nominal_weight_kg_per_m(diameter_mm: float) -> float:
    """Tính trọng lượng danh định 1 mét thép thanh theo TCVN 1651:2018: q = pi * (d/2)^2 * 7850."""
    radius_m = (diameter_mm / 2.0) / 1000.0
    area_m2 = math.pi * (radius_m ** 2)
    return round(area_m2 * STEEL_DENSITY_KG_M3, 3)


@dataclass
class IFCConcreteElement:
    """Cấu kiện bê tông trích xuất từ IFC."""
    element_id: str
    global_id: str
    name: str
    ifc_type: str
    volume_m3: float = 0.0
    formwork_area_m2: float = 0.0
    length_m: float = 0.0
    width_m: float = 0.0
    height_m: float = 0.0
    material: str = "Bê tông C30/37"
    storey_or_section: str = ""
    properties: Dict[str, Any] = field(default_factory=dict)


@dataclass
class IFCRebarElement:
    """Thanh cốt thép trích xuất từ IFC."""
    element_id: str
    global_id: str
    mark: str
    diameter_mm: float
    length_mm: float
    quantity: int = 1
    grade: str = "CB400-V"
    bending_shape: str = "Thẳng"
    host_element: str = ""
    total_length_m: float = 0.0
    unit_weight_kg_m: float = 0.0
    total_weight_kg: float = 0.0

    def __post_init__(self):
        if self.unit_weight_kg_m == 0.0 and self.diameter_mm > 0:
            self.unit_weight_kg_m = nominal_weight_kg_per_m(self.diameter_mm)
        if self.total_length_m == 0.0:
            self.total_length_m = round((self.length_mm / 1000.0) * self.quantity, 2)
        if self.total_weight_kg == 0.0:
            self.total_weight_kg = round(self.total_length_m * self.unit_weight_kg_m, 2)


@dataclass
class IFCTakeoffResult:
    """Kết quả bóc tách tổng hợp toàn bộ file IFC."""
    source_file: str
    schema_version: str = "IFC4"
    project_name: str = ""
    concrete_elements: List[IFCConcreteElement] = field(default_factory=list)
    rebar_elements: List[IFCRebarElement] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def total_concrete_volume_m3(self) -> float:
        return round(sum(e.volume_m3 for e in self.concrete_elements), 3)

    @property
    def total_formwork_area_m2(self) -> float:
        return round(sum(e.formwork_area_m2 for e in self.concrete_elements), 2)

    @property
    def total_rebar_weight_kg(self) -> float:
        return round(sum(r.total_weight_kg for r in self.rebar_elements), 2)

    @property
    def total_rebar_weight_ton(self) -> float:
        return round(self.total_rebar_weight_kg / 1000.0, 3)

    def summary_concrete_by_type(self) -> Dict[str, Dict[str, float]]:
        """Tổng hợp khối lượng bê tông và ván khuôn phân theo từng loại cấu kiện."""
        res: Dict[str, Dict[str, float]] = {}
        for e in self.concrete_elements:
            t = e.ifc_type
            if t not in res:
                res[t] = {"count": 0, "volume_m3": 0.0, "formwork_m2": 0.0}
            res[t]["count"] += 1
            res[t]["volume_m3"] = round(res[t]["volume_m3"] + e.volume_m3, 3)
            res[t]["formwork_m2"] = round(res[t]["formwork_m2"] + e.formwork_area_m2, 2)
        return res

    def summary_rebar_by_diameter(self) -> Dict[float, Dict[str, Any]]:
        """Tổng hợp cốt thép phân theo từng đường kính Ø (mm)."""
        res: Dict[float, Dict[str, Any]] = {}
        for r in self.rebar_elements:
            d = r.diameter_mm
            if d not in res:
                res[d] = {"total_pieces": 0, "total_length_m": 0.0, "total_weight_kg": 0.0, "grades": set()}
            res[d]["total_pieces"] += r.quantity
            res[d]["total_length_m"] = round(res[d]["total_length_m"] + r.total_length_m, 2)
            res[d]["total_weight_kg"] = round(res[d]["total_weight_kg"] + r.total_weight_kg, 2)
            res[d]["grades"].add(r.grade)
        return res

    def to_cutting_stock_demands(self) -> List[CutDemand]:
        """
        Chuyển đổi trực tiếp toàn bộ cốt thép trong mô hình IFC thành danh sách CutDemand
        để nạp thẳng vào Google OR-Tools Cutting Stock Solver (cắt cây nguyên 11.7m).
        """
        if self.errors:
            raise ValueError("IFC chưa đủ dữ liệu để cắt thép: " + "; ".join(self.errors))
        demands: List[CutDemand] = []
        for r in self.rebar_elements:
            if r.length_mm > 0 and r.quantity > 0 and r.diameter_mm > 0:
                demands.append(
                    CutDemand(
                        mark=r.mark,
                        length_mm=int(round(r.length_mm)),
                        quantity=r.quantity,
                        diameter_mm=r.diameter_mm,
                        grade=r.grade,
                    )
                )
        return demands


class IFCLoader:
    """
    Trình nạp và bóc tách dữ liệu OpenBIM IFC chuyên nghiệp cho ngành Xây dựng Cầu đường & Hạ tầng.
    """

    CONCRETE_TYPES = {
        "IfcBeam", "IfcColumn", "IfcFooting", "IfcPile", "IfcSlab",
        "IfcWall", "IfcWallStandardCase", "IfcMember", "IfcPlate",
        "IfcBuildingElementProxy", "IfcCivilElement"
    }

    REBAR_TYPES = {
        "IfcReinforcingBar", "IfcReinforcingElement", "IfcReinforcingMesh", "IfcTendon"
    }

    def __init__(self):
        self.has_native_ifc = HAS_IFCOPENSHELL

    def load_ifc(self, file_path: str) -> IFCTakeoffResult:
        """
        Đọc và trích xuất toàn bộ khối lượng hình học và cốt thép từ file IFC.
        Tự động sử dụng IfcOpenShell C++ Engine (hoặc fallback ISO 10303-21 parser).
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Không tìm thấy file IFC: {file_path}")

        if self.has_native_ifc:
            try:
                return self._load_with_ifcopenshell(file_path)
            except Exception as e:
                # Nếu IfcOpenShell gặp lỗi ngoại lệ, chuyển sang fallback
                fallback_res = self._load_with_step_parser(file_path)
                fallback_res.errors.append(f"IfcOpenShell exception: {str(e)} — Đã kích hoạt Fallback Parser.")
                return fallback_res
        else:
            return self._load_with_step_parser(file_path)

    def _load_with_ifcopenshell(self, file_path: str) -> IFCTakeoffResult:
        """Trích xuất khối lượng bằng thư viện chuẩn IfcOpenShell."""
        model = ifcopenshell.open(file_path)
        schema_version = getattr(model, "schema", "IFC4")
        if ifcopenshell.util.unit.get_project_unit(model, "LENGTHUNIT") is None:
            raise ValueError("IFC thiếu đơn vị LENGTHUNIT; không suy đoán mét/mm")
        length_scale = ifcopenshell.util.unit.calculate_unit_scale(model, "LENGTHUNIT")


        # Lấy tên dự án
        project_name = "Dự án Hạ tầng OpenBIM"
        projects = model.by_type("IfcProject")
        if projects and projects[0].Name:
            project_name = str(projects[0].Name)

        result = IFCTakeoffResult(
            source_file=file_path,
            schema_version=schema_version,
            project_name=project_name,
        )

        # 1. BÓC TÁCH CẤU KIỆN BÊ TÔNG
        seen_concrete_ids = set()
        for c_type in self.CONCRETE_TYPES:
            try:
                elements = model.by_type(c_type)
            except RuntimeError:  # Entity không tồn tại trong schema này (vd IFC2X3).
                continue
            for elem in elements:
                elem_id = str(elem.id())
                if elem_id in seen_concrete_ids:
                    continue
                seen_concrete_ids.add(elem_id)
                global_id = str(elem.GlobalId) if hasattr(elem, "GlobalId") else elem_id
                name = str(elem.Name or f"{c_type}_{elem_id}")

                volume_m3 = self._extract_quantity(elem, model, "IfcQuantityVolume", "VolumeValue", "VOLUMEUNIT")
                formwork_area_m2 = self._extract_quantity(elem, model, "IfcQuantityArea", "AreaValue", "AREAUNIT")
                if volume_m3 is None:
                    result.errors.append(f"{name}: thiếu thể tích hoặc đơn vị thể tích đã khai báo")
                    volume_m3 = 0.0
                length_m, width_m, height_m = self._extract_element_dimensions(elem)
                length_m, width_m, height_m = (v * length_scale for v in (length_m, width_m, height_m))
                material = self._extract_material_name(elem)

                # Xác định tầng / phân đoạn
                storey_name = ""
                try:
                    for rel in getattr(elem, "ContainedInStructure", []):
                        if hasattr(rel, "RelatingStructure") and rel.RelatingStructure:
                            storey_name = str(rel.RelatingStructure.Name or "")
                except Exception:
                    pass

                result.concrete_elements.append(
                    IFCConcreteElement(
                        element_id=elem_id,
                        global_id=global_id,
                        name=name,
                        ifc_type=c_type,
                        volume_m3=volume_m3,
                        formwork_area_m2=formwork_area_m2 or 0.0,
                        length_m=length_m,
                        width_m=width_m,
                        height_m=height_m,
                        material=material,
                        storey_or_section=storey_name,
                    )
                )

        # 2. BÓC TÁCH CỐT THÉP 3D
        seen_rebar_ids = set()
        for r_type in self.REBAR_TYPES:
            try:
                elements = model.by_type(r_type)
            except RuntimeError:
                continue
            for r_elem in elements:
                elem_id = str(r_elem.id())
                if elem_id in seen_rebar_ids:
                    continue
                seen_rebar_ids.add(elem_id)
                global_id = str(r_elem.GlobalId) if hasattr(r_elem, "GlobalId") else elem_id
                mark = str(r_elem.Name or r_elem.Tag or f"RB-{elem_id}")

                # Đường kính danh định Ø (mm)
                dia = 0.0
                if hasattr(r_elem, "NominalDiameter") and r_elem.NominalDiameter:
                    dia = float(r_elem.NominalDiameter) * length_scale * 1000
                elif hasattr(r_elem, "NominalBarDiameter") and r_elem.NominalBarDiameter:
                    dia = float(r_elem.NominalBarDiameter) * length_scale * 1000
                # Không suy luận kích thước từ tên hay độ lớn của con số.
                length_mm = float(getattr(r_elem, "BarLength", 0) or 0) * length_scale * 1000
                if not all(math.isfinite(v) and v > 0 for v in (dia, length_mm)):
                    result.errors.append(f"{mark}: thiếu/sai đường kính hoặc chiều dài thanh")
                    continue

                # Số lượng thanh
                quantity = 1
                if hasattr(r_elem, "NumberOfBars") and r_elem.NumberOfBars:
                    quantity = int(r_elem.NumberOfBars)

                # Mác thép
                grade = "CB400-V"
                if hasattr(r_elem, "SteelGrade") and r_elem.SteelGrade:
                    grade = str(r_elem.SteelGrade)

                if dia > 0 and length_mm > 0:
                    result.rebar_elements.append(
                        IFCRebarElement(
                            element_id=elem_id,
                            global_id=global_id,
                            mark=mark,
                            diameter_mm=dia,
                            length_mm=length_mm,
                            quantity=quantity,
                            grade=grade,
                        )
                    )

        return result

    @staticmethod
    def _extract_quantity(elem, model, entity_type, value_attr, unit_type):
        """Đổi Qto sang SI, tôn trọng cả đơn vị riêng của quantity."""
        candidates = []
        for rel in getattr(elem, "IsDefinedBy", ()):
            if not rel.is_a("IfcRelDefinesByProperties"):
                continue
            pset = rel.RelatingPropertyDefinition
            if not pset or not pset.is_a("IfcElementQuantity"):
                continue
            for q in pset.Quantities:
                if not q.is_a(entity_type):
                    continue
                # Diện tích mặt bằng/diện tích sàn không phải diện tích ván khuôn.
                name = str(q.Name or "").lower()
                if unit_type == "AREAUNIT" and "formwork" not in name:
                    continue
                unit = q.Unit or ifcopenshell.util.unit.get_project_unit(model, unit_type)
                if unit is None:
                    continue
                value = float(getattr(q, value_attr))
                target = model.create_entity("IfcSIUnit", UnitType=unit_type,
                    Name={"VOLUMEUNIT": "CUBIC_METRE", "AREAUNIT": "SQUARE_METRE"}[unit_type])
                value = ifcopenshell.util.unit.convert_unit(value, unit, target)
                if math.isfinite(value) and value >= 0:
                    candidates.append((0 if name == "netvolume" else 1, value))
        return min(candidates, key=lambda v: v[0])[1] if candidates else None

    def _extract_element_dimensions(self, elem) -> Tuple[float, float, float]:
        """Trích xuất kích thước Dài, Rộng, Cao."""
        length, width, height = 0.0, 0.0, 0.0
        try:
            psets = ifcopenshell.util.element.get_psets(elem)
            for pset_name, pset_data in psets.items():
                if "Length" in pset_data:
                    length = float(pset_data["Length"])
                if "Width" in pset_data:
                    width = float(pset_data["Width"])
                if "Height" in pset_data:
                    height = float(pset_data["Height"])
        except Exception:
            pass
        return length, width, height

    def _extract_material_name(self, elem) -> str:
        """Trích xuất tên mác bê tông từ IfcMaterial / IfcMaterialSelect."""
        try:
            material = ifcopenshell.util.element.get_material(elem)
            if material:
                if hasattr(material, "Name") and material.Name:
                    return str(material.Name)
                return str(material)
        except Exception:
            pass
        return "Bê tông C30/37"

    def _load_with_step_parser(self, file_path: str) -> IFCTakeoffResult:
        """
        Bộ đọc ISO 10303-21 STEP thuần Python (Zero Dependency Fallback).
        Phân tích dòng văn bản tệp IFC để trích xuất cấu kiện và cốt thép khi không có C++ runtime.
        """
        # Parser thô không đọc được UnitsInContext/Qto và quan hệ STEP một cách tin cậy.
        # Trả lỗi có cấu trúc, tuyệt đối không tạo khối lượng hay thanh thép giả.
        return IFCTakeoffResult(source_file=file_path, errors=[
            "Không thể bóc IFC bằng parser dự phòng. Cần IfcOpenShell và mô hình hợp lệ có đơn vị/Qto."
        ])
