import os
import sys
import re
import math
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

print("=== STARTING MASTER CAD TAKEOFF EXCEL GENERATOR (WITH FULL X-Y AXES) ===", flush=True)

# 1. LOAD CAD DATA
df_blocks = pd.read_csv(r'C:\Users\baotu\Downloads\cad_blocks.csv', on_bad_lines='skip')
tengas = df_blocks[df_blocks['BlockName'].isin(['TengaK', 'XR TNT$0$Xr-TNT-Khai$0$xr COT$0$TN-TENGA5'])].copy().reset_index(drop=True)

def parse_attrs(s):
    if not isinstance(s, str): return {}
    parts = s.split('|')
    d = {}
    for p in parts:
        if '=' in p:
            k, v = p.split('=', 1)
            d[k.strip()] = v.strip()
    return d

def clean_float(val):
    if not isinstance(val, str) and not isinstance(val, (int, float)): return None
    try: return float(val)
    except: return None

def clean_coord(val):
    if not isinstance(val, str): return None
    m = re.search(r'[XY]=([0-9]+\.[0-9]+|[0-9]+)', val)
    if m:
        return float(m.group(1))
    m2 = re.search(r'([0-9]{6,7}(?:\.[0-9]+)?)', val)
    if m2:
        return float(m2.group(1))
    return None

manholes = []
for idx, r in tengas.iterrows():
    attrs = parse_attrs(r['Attributes'])
    t_name = attrs.get('T', f'GA_{idx+1}')
    sh = attrs.get('SH', 'TB41')
    mg = clean_float(attrs.get('MG'))
    dc = clean_float(attrs.get('DC'))
    
    # Coordinate extraction:
    # Surveyor attribute has X=North (~2.305M) and Y=East (~585k)
    # CAD native insertion point has r['Y'] = North and r['X'] = East
    x_vn2000 = clean_coord(attrs.get('X'))
    if x_vn2000 is None or x_vn2000 < 1000000:
        x_vn2000 = round(float(r['Y']), 3)
    else:
        x_vn2000 = round(x_vn2000, 3)
        
    y_vn2000 = clean_coord(attrs.get('Y'))
    if y_vn2000 is None or y_vn2000 > 1000000:
        y_vn2000 = round(float(r['X']), 3)
    else:
        y_vn2000 = round(y_vn2000, 3)
    
    cad_x = r['X']
    cad_y = r['Y']
    
    h = (mg - dc) if (mg is not None and dc is not None) else 1.2
    h = round(h, 3)
    
    if h <= 1.5:
        cap_sau = "Cấp 1 (H <= 1.5m)"
        nhom = 1
    elif h <= 2.5:
        cap_sau = "Cấp 2 (1.5m < H <= 2.5m)"
        nhom = 2
    else:
        cap_sau = "Cấp 3 (H > 2.5m)"
        nhom = 3
        
    qa_status = "ĐẠT CHUẨN"
    if h <= 0: qa_status = "LỖI: H <= 0"
    elif h < 0.6: qa_status = "CẢNH BÁO: H < 0.6m"
    elif h > 3.5: qa_status = "LƯU Ý: H > 3.5m (Cần cừ ván thép)"
    
    manholes.append({
        'Handle': r['Handle'],
        'TenGa': t_name,
        'LoaiGa': sh,
        'MG': mg,
        'DC': dc,
        'H': h,
        'X_VN2000': x_vn2000,
        'Y_VN2000': y_vn2000,
        'CAD_X': cad_x,
        'CAD_Y': cad_y,
        'CapSau': cap_sau,
        'Nhom': nhom,
        'QA_Status': qa_status
    })

df_manholes = pd.DataFrame(manholes)
print(f"Total Manholes: {len(df_manholes)}")

# 2. LOAD CULVERTS
df_p = pd.read_csv(r'C:\Users\baotu\Downloads\cad_pline_coords.csv')
d800_cross = df_p[df_p['Layer'] == 'dim 800 ngang duong'].copy().reset_index(drop=True)
d300_cross = df_p[df_p['Layer'] == 'dim 300 ngang duong'].copy().reset_index(drop=True)
d300_he = df_p[df_p['Layer'].str.contains('HTKT_TN_D300', regex=False)].copy().reset_index(drop=True)

ga_cad_x = df_manholes['CAD_X'].values
ga_cad_y = df_manholes['CAD_Y'].values

cell = 10.0
grid = {}
for idx in range(len(df_manholes)):
    cx = int(ga_cad_x[idx] // cell)
    cy = int(ga_cad_y[idx] // cell)
    grid.setdefault((cx, cy), []).append(idx)

def find_nearest_ga(px, py):
    cx = int(px // cell)
    cy = int(py // cell)
    best_dist = 999999.0
    best_idx = -1
    for dx in range(-2, 3):
        for dy in range(-2, 3):
            cand = grid.get((cx+dx, cy+dy), [])
            for c_idx in cand:
                dist = math.hypot(px - ga_cad_x[c_idx], py - ga_cad_y[c_idx])
                if dist < best_dist:
                    best_dist = dist
                    best_idx = c_idx
    return best_dist, best_idx

culverts = []

# Process D800
for idx, r in d800_cross.iterrows():
    d1, i1 = find_nearest_ga(r['StartX'], r['StartY'])
    d2, i2 = find_nearest_ga(r['EndX'], r['EndY'])
    g1 = df_manholes.iloc[i1] if i1 >= 0 else None
    g2 = df_manholes.iloc[i2] if i2 >= 0 else None
    name1 = g1['TenGa'] if g1 is not None else "GA_DAU"
    name2 = g2['TenGa'] if g2 is not None else "GA_CUOI"
    dc1 = g1['DC'] if g1 is not None else None
    dc2 = g2['DC'] if g2 is not None else None
    h1 = g1['H'] if g1 is not None else 1.8
    h2 = g2['H'] if g2 is not None else 1.8
    h_tb = round((h1 + h2) / 2.0, 3)
    L = round(r['Length'], 3)
    dh = abs(dc1 - dc2) if (dc1 is not None and dc2 is not None) else 0.0
    slope = round((dh / L) * 100, 3) if L > 0 else 0.0
    
    culverts.append({
        'Handle': r['Handle'],
        'PhanLoai': 'Cống D800 BTCT qua đường',
        'DuongKinh': 800,
        'ViTri': 'Băng ngang lòng đường',
        'GaDau': name1,
        'GaCuoi': name2,
        'X_Dau': round(r['StartY'], 3),
        'Y_Dau': round(r['StartX'], 3),
        'X_Cuoi': round(r['EndY'], 3),
        'Y_Cuoi': round(r['EndX'], 3),
        'ChieuDai_L': L,
        'DC1': dc1,
        'DC2': dc2,
        'dH': round(dh, 3),
        'Slope_pct': slope,
        'H_Dao_TB': h_tb,
    })

# Process D300 road cross
for idx, r in d300_cross.iterrows():
    d1, i1 = find_nearest_ga(r['StartX'], r['StartY'])
    d2, i2 = find_nearest_ga(r['EndX'], r['EndY'])
    g1 = df_manholes.iloc[i1] if i1 >= 0 else None
    g2 = df_manholes.iloc[i2] if i2 >= 0 else None
    name1 = g1['TenGa'] if g1 is not None else "GA_DAU"
    name2 = g2['TenGa'] if g2 is not None else "GA_CUOI"
    dc1 = g1['DC'] if g1 is not None else None
    dc2 = g2['DC'] if g2 is not None else None
    h1 = g1['H'] if g1 is not None else 1.35
    h2 = g2['H'] if g2 is not None else 1.35
    h_tb = round((h1 + h2) / 2.0, 3)
    L = round(r['Length'], 3)
    dh = abs(dc1 - dc2) if (dc1 is not None and dc2 is not None) else 0.0
    slope = round((dh / L) * 100, 3) if L > 0 else 0.0
    
    culverts.append({
        'Handle': r['Handle'],
        'PhanLoai': 'Cống D300 HDPE qua đường',
        'DuongKinh': 300,
        'ViTri': 'Băng ngang lòng đường',
        'GaDau': name1,
        'GaCuoi': name2,
        'X_Dau': round(r['StartY'], 3),
        'Y_Dau': round(r['StartX'], 3),
        'X_Cuoi': round(r['EndY'], 3),
        'Y_Cuoi': round(r['EndX'], 3),
        'ChieuDai_L': L,
        'DC1': dc1,
        'DC2': dc2,
        'dH': round(dh, 3),
        'Slope_pct': slope,
        'H_Dao_TB': h_tb,
    })

# Process D300 sidewalk
for idx, r in d300_he.iterrows():
    d1, i1 = find_nearest_ga(r['StartX'], r['StartY'])
    d2, i2 = find_nearest_ga(r['EndX'], r['EndY'])
    g1 = df_manholes.iloc[i1] if i1 >= 0 else None
    g2 = df_manholes.iloc[i2] if i2 >= 0 else None
    name1 = g1['TenGa'] if g1 is not None else "GA_DAU"
    name2 = g2['TenGa'] if g2 is not None else "GA_CUOI"
    dc1 = g1['DC'] if g1 is not None else None
    dc2 = g2['DC'] if g2 is not None else None
    h1 = g1['H'] if g1 is not None else 1.25
    h2 = g2['H'] if g2 is not None else 1.25
    h_tb = round((h1 + h2) / 2.0, 3)
    L = round(r['Length'], 3)
    dh = abs(dc1 - dc2) if (dc1 is not None and dc2 is not None) else 0.0
    slope = round((dh / L) * 100, 3) if L > 0 else 0.0
    
    culverts.append({
        'Handle': r['Handle'],
        'PhanLoai': 'Cống D300 HDPE vỉa hè',
        'DuongKinh': 300,
        'ViTri': 'Dọc vỉa hè / hào kỹ thuật',
        'GaDau': name1,
        'GaCuoi': name2,
        'X_Dau': round(r['StartY'], 3),
        'Y_Dau': round(r['StartX'], 3),
        'X_Cuoi': round(r['EndY'], 3),
        'Y_Cuoi': round(r['EndX'], 3),
        'ChieuDai_L': L,
        'DC1': dc1,
        'DC2': dc2,
        'dH': round(dh, 3),
        'Slope_pct': slope,
        'H_Dao_TB': h_tb,
    })

df_culverts = pd.DataFrame(culverts)
print(f"Total Culvert Segments: {len(df_culverts)}")
print(f"Total Length: {df_culverts['ChieuDai_L'].sum():.2f} m")

# 3. BUILD WORKBOOK
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Styles
FONT_TITLE = Font(name="Segoe UI", size=13, bold=True, color="1B365D")
FONT_SUBTITLE = Font(name="Segoe UI", size=10, italic=True, color="444444")
FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_SECTION = Font(name="Segoe UI", size=10, bold=True, color="1B365D")
FONT_BOLD = Font(name="Segoe UI", size=10, bold=True)
FONT_REGULAR = Font(name="Segoe UI", size=10)
FONT_NOTE = Font(name="Segoe UI", size=9, italic=True, color="555555")

FILL_NAVY = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
FILL_SECTION = PatternFill(start_color="E6EEF8", end_color="E6EEF8", fill_type="solid")
FILL_SUBTOTAL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

BORDER_THIN = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF')
)
BORDER_TOTAL = Border(
    top=Side(style='thin', color='1B365D'),
    bottom=Side(style='double', color='1B365D'),
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF')
)

def style_cell(cell, font=None, fill=None, alignment=None, border=BORDER_THIN, num_format=None):
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if border: cell.border = border
    if num_format: cell.number_format = num_format

# ==========================================
# SHEET 1: 📑 Tổng Hợp BoQ
# ==========================================
ws1 = wb.create_sheet(title="📑 Tổng Hợp BoQ")
ws1.views.sheetView[0].showGridLines = True

ws1.merge_cells("B2:G2")
ws1["B2"] = "BẢNG TỔNG HỢP TIÊN LƯỢNG KHỐI LƯỢNG (BoQ) HỆ THỐNG THOÁT NƯỚC THẢI"
ws1["B2"].font = FONT_TITLE

ws1.merge_cells("B3:G3")
ws1["B3"] = "DỰ ÁN: KHU ĐÔ THỊ VINHOMES — LÔ B9.2, B9.3, B9.4 (MẶT BẰNG THOÁT NƯỚC THẢI TOÀN KHU)"
ws1["B3"].font = FONT_SUBTITLE

ws1.merge_cells("B4:G4")
ws1["B4"] = "Chủ đầu tư: Vinhomes PMU | Nhà thầu: Vinalpha | Số liệu trích xuất CAD 100% | Đầy đủ Tọa độ Trục X - Y"
ws1["B4"].font = FONT_NOTE

boq_headers = ["STT", "Mã hiệu ĐM", "Nội dung công việc", "ĐVT", "Khối lượng", "Liên kết bảng chi tiết", "Ghi chú kỹ thuật"]
for col_idx, h in enumerate(boq_headers, start=2):
    cell = ws1.cell(row=6, column=col_idx, value=h)
    style_cell(cell, font=FONT_HEADER, fill=FILL_NAVY, alignment=ALIGN_CENTER)
ws1.row_dimensions[6].height = 28

boq_rows = [
    {"type": "sec", "stt": "A", "code": "", "name": "HẠNG MỤC I: CỐNG THOÁT NƯỚC THẢI (D300 & D800)", "unit": ""},
    {"type": "item", "stt": "1", "code": "AB.24111", "name": "Đào đất rãnh đặt cống thoát nước thải bằng máy đào kết hợp thủ công (đất cấp II)", "unit": "m3", "ref": "='📊 Bóc Tách Cống'!L13", "note": "Đào hào hình thang taluy m=0.33"},
    {"type": "item", "stt": "2", "code": "AB.65111", "name": "Đắp cát đệm bảo vệ cống & đắp đất rãnh cống hoàn trả đầm chặt K95", "unit": "m3", "ref": "='📊 Bóc Tách Cống'!L18", "note": "Đắp hoàn trả trừ thể tích ống"},
    {"type": "item", "stt": "3", "code": "AB.71112", "name": "Vận chuyển đất đào thừa đổ đi bãi tập kết quy định bằng ô tô tự đổ 10T", "unit": "m3", "ref": "='📊 Bóc Tách Cống'!L22", "note": "Đất thừa = Đào - Đắp"},
    {"type": "item", "stt": "4", "code": "BB.11101", "name": "Cung cấp và lắp đặt cống HDPE 2 lớp SN4 D300 trên vỉa hè / rãnh kỹ thuật", "unit": "m", "ref": "='📊 Bóc Tách Cống'!L26", "note": "852 đoạn cống dọc vỉa hè"},
    {"type": "item", "stt": "5", "code": "BB.11102", "name": "Cung cấp và lắp đặt cống HDPE 2 vách SN8 D300 chịu lực dưới lòng đường", "unit": "m", "ref": "='📊 Bóc Tách Cống'!L30", "note": "106 đoạn cống băng ngang đường"},
    {"type": "item", "stt": "6", "code": "BB.12104", "name": "Cung cấp và lắp đặt cống bê tông cốt thép đúc sẵn D800 chịu tải trọng TC", "unit": "m", "ref": "='📊 Bóc Tách Cống'!L34", "note": "9 đoạn cống BTCT đúc sẵn qua đường"},
    {"type": "sec", "stt": "B", "code": "", "name": "HẠNG MỤC II: HỐ GA THOÁT NƯỚC THẢI (TB41 & GA THĂM TOÀN KHU)", "unit": ""},
    {"type": "item", "stt": "7", "code": "AB.25111", "name": "Đào đất hố móng ga thoát nước thải bằng máy đào kết hợp thủ công (đất cấp II)", "unit": "m3", "ref": "='📊 Bóc Tách Hố Ga'!L14", "note": "Đào móng 990 hố ga theo 3 cấp độ sâu"},
    {"type": "item", "stt": "8", "code": "AF.11111", "name": "Bê tông lót móng hố ga M100 đá 4x6 dày 100mm", "unit": "m3", "ref": "='📊 Bóc Tách Hố Ga'!L18", "note": "Bê tông lót đáy hố móng 1.5x1.5m"},
    {"type": "item", "stt": "9", "code": "AF.81111", "name": "Ván khuôn bê tông lót móng hố ga", "unit": "m2", "ref": "='📊 Bóc Tách Hố Ga'!L22", "note": "Ván khuôn thành đệm lót dày 100mm"},
    {"type": "item", "stt": "10", "code": "AF.21111", "name": "Bê tông thành và đáy hố ga C20/M200 (đúc sẵn hoặc đổ tại chỗ)", "unit": "m3", "ref": "='📊 Bóc Tách Hố Ga'!L27", "note": "Đáy dày 15cm, thành dày 15cm"},
    {"type": "item", "stt": "11", "code": "AF.82111", "name": "Ván khuôn thân và thành hố ga (2 mặt trong và ngoài)", "unit": "m2", "ref": "='📊 Bóc Tách Hố Ga'!L31", "note": "Ván khuôn thành hố ga"},
    {"type": "item", "stt": "12", "code": "AF.61111", "name": "Cốt thép hố ga CB300-V (đường kính <= 18mm)", "unit": "kg", "ref": "='📊 Bóc Tách Hố Ga'!L35", "note": "Định mức 65 kg/m3 bê tông"},
    {"type": "item", "stt": "13", "code": "BB.81111", "name": "Sản xuất và lắp đặt nắp hố ga gang cầu / composite ngăn mùi / song chắn rác", "unit": "bộ", "ref": "='📊 Bóc Tách Hố Ga'!L39", "note": "990 bộ nắp chuẩn TB41 và ga D800"},
    {"type": "item", "stt": "14", "code": "AB.65112", "name": "Đắp đất hố móng xung quanh thành ga hoàn trả đầm chặt K95", "unit": "m3", "ref": "='📊 Bóc Tách Hố Ga'!L43", "note": "Đắp hoàn trả mang hố ga"},
    {"type": "item", "stt": "15", "code": "AB.71113", "name": "Vận chuyển đất đào thừa hố ga đổ đi bãi tập kết quy định bằng ô tô tự đổ 10T", "unit": "m3", "ref": "='📊 Bóc Tách Hố Ga'!L47", "note": "Đất thừa hố móng ga = Đào - Đắp"}
]


curr_row = 7
for r in boq_rows:
    ws1.row_dimensions[curr_row].height = 22
    if r["type"] == "sec":
        ws1.cell(row=curr_row, column=2, value=r["stt"])
        ws1.cell(row=curr_row, column=3, value=r["code"])
        ws1.cell(row=curr_row, column=4, value=r["name"])
        for c in range(2, 9):
            style_cell(ws1.cell(row=curr_row, column=c), font=FONT_SECTION, fill=FILL_SECTION)
        ws1.cell(row=curr_row, column=2).alignment = ALIGN_CENTER
    else:
        ws1.cell(row=curr_row, column=2, value=r["stt"]).alignment = ALIGN_CENTER
        ws1.cell(row=curr_row, column=3, value=r["code"]).alignment = ALIGN_CENTER
        ws1.cell(row=curr_row, column=4, value=r["name"]).alignment = ALIGN_LEFT
        ws1.cell(row=curr_row, column=5, value=r["unit"]).alignment = ALIGN_CENTER
        
        c_kl = ws1.cell(row=curr_row, column=6, value=r["ref"])
        c_kl.alignment = ALIGN_RIGHT
        c_kl.number_format = '#,##0.00' if r["unit"] != "bộ" else '#,##0'
        c_kl.font = FONT_BOLD
        
        ws1.cell(row=curr_row, column=7, value=r["ref"].split("!")[0].replace("='", "").replace("'", "")).alignment = ALIGN_LEFT
        ws1.cell(row=curr_row, column=8, value=r["note"]).alignment = ALIGN_LEFT
        
        for c in [2, 3, 4, 5, 7, 8]:
            style_cell(ws1.cell(row=curr_row, column=c), font=FONT_REGULAR)
        style_cell(c_kl, font=FONT_BOLD)
    curr_row += 1

# ==========================================
# SHEET 2: 📊 Bóc Tách Cống (WBS 12 Cột)
# ==========================================
ws2 = wb.create_sheet(title="📊 Bóc Tách Cống")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("B2:L2")
ws2["B2"] = "BẢNG TÍNH BÓC TÁCH CHI TIẾT KHỐI LƯỢNG HỆ THỐNG CỐNG THOÁT NƯỚC THẢI"
ws2["B2"].font = FONT_TITLE

ws2.merge_cells("B3:L3")
ws2["B3"] = "Căn cứ hình học CAD 100% (852 đoạn D300 hè, 106 đoạn D300 qua đường, 9 đoạn D800 BTCT) | 100% CÔNG THỨC SỐNG"
ws2["B3"].font = FONT_SUBTITLE

wbs_headers = ["STT", "Mã hiệu ĐM", "Nội dung công tác & Diễn giải khối lượng", "ĐVT", "Số lượng", "Dài L (m)", "Rộng B (m)", "Cao H (m)", "Hệ số", "Công thức diễn giải", "Khối lượng", "Ghi chú"]
for c_idx, h in enumerate(wbs_headers, start=2):
    cell = ws2.cell(row=5, column=c_idx, value=h)
    style_cell(cell, font=FONT_HEADER, fill=FILL_NAVY, alignment=ALIGN_CENTER)
ws2.row_dimensions[5].height = 28

ws2.cell(row=6, column=2, value="A").alignment = ALIGN_CENTER
ws2.cell(row=6, column=4, value="CỐNG THOÁT NƯỚC THẢI KHU B9.2, B9.3, B9.4 (CAD 100%)")
for c in range(2, 14):
    style_cell(ws2.cell(row=6, column=c), font=FONT_SECTION, fill=FILL_SECTION)

# 1. Đào rãnh cống
ws2.cell(row=7, column=2, value="1").alignment = ALIGN_CENTER
ws2.cell(row=7, column=3, value="AB.24111").alignment = ALIGN_CENTER
ws2.cell(row=7, column=4, value="Đào đất rãnh đặt cống thoát nước thải bằng máy đào kết hợp thủ công (đất cấp II)")
ws2.cell(row=7, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws2.cell(row=7, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

details_culvert_excav = [
    ("Đào rãnh cống HDPE D300 trên vỉa hè (852 đoạn cống dọc vỉa hè)", "m3", 1, "='📋 CSDL Tuyến Cống'!G4", 0.85, 1.28, 1.0, "=F8*G8*(H8*I8+0.33*I8^2)*J8", "=F8*G8*(H8*I8+0.33*I8^2)*J8", "Hào đào hình thang m=0.33"),
    ("Đào rãnh cống HDPE D300 băng ngang lòng đường (106 đoạn)", "m3", 1, "='📋 CSDL Tuyến Cống'!G2", 0.90, 1.35, 1.0, "=F9*G9*(H9*I9+0.33*I9^2)*J9", "=F9*G9*(H9*I9+0.33*I9^2)*J9", "Hào đào qua đường m=0.33"),
    ("Đào rãnh cống BTCT D800 đúc sẵn qua đường (9 đoạn)", "m3", 1, "='📋 CSDL Tuyến Cống'!G3", 1.50, 1.85, 1.0, "=F10*G10*(H10*I10+0.33*I10^2)*J10", "=F10*G10*(H10*I10+0.33*I10^2)*J10", "Cống D800 m=0.33"),
    ("Đào mở rộng tại vị trí nối ga / mở rộng hố đấu nối (5%)", "m3", 1, "=SUM(L8:L10)", 1.0, 1.0, 0.05, "=G11*J11", "=G11*J11", "5% đào mở rộng kỹ thuật"),
]

r_idx = 8
for item in details_culvert_excav:
    ws2.cell(row=r_idx, column=4, value=item[0])
    ws2.cell(row=r_idx, column=5, value=item[1]).alignment = ALIGN_CENTER
    ws2.cell(row=r_idx, column=6, value=item[2]).alignment = ALIGN_CENTER
    ws2.cell(row=r_idx, column=7, value=item[3]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=7).number_format = '#,##0.00'
    ws2.cell(row=r_idx, column=8, value=item[4]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=9, value=item[5]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=10, value=item[6]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=11, value=item[7]).alignment = ALIGN_LEFT
    ws2.cell(row=r_idx, column=12, value=item[8]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=12).font = FONT_BOLD
    ws2.cell(row=r_idx, column=12).number_format = '#,##0.00'
    ws2.cell(row=r_idx, column=13, value=item[9])
    for c in range(2, 14):
        style_cell(ws2.cell(row=r_idx, column=c), font=FONT_REGULAR)
    style_cell(ws2.cell(row=r_idx, column=12), font=FONT_BOLD)
    r_idx += 1

# Subtotal row for Item 1
ws2.cell(row=13, column=4, value="CỘNG CÔNG TÁC ĐÀO ĐẤT RÃNH CỐNG (AB.24111)").font = FONT_BOLD
ws2.cell(row=13, column=5, value="m3").alignment = ALIGN_CENTER
ws2.cell(row=13, column=12, value="=SUM(L8:L11)").font = FONT_BOLD
ws2.cell(row=13, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws2.cell(row=13, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 2. Đắp cát & đất rãnh cống
ws2.cell(row=14, column=2, value="2").alignment = ALIGN_CENTER
ws2.cell(row=14, column=3, value="AB.65111").alignment = ALIGN_CENTER
ws2.cell(row=14, column=4, value="Đắp cát đệm bảo vệ cống & đắp đất rãnh cống hoàn trả đầm chặt K95")
ws2.cell(row=14, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws2.cell(row=14, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

details_culvert_backfill = [
    ("Đệm cát lót đáy cống dày 100mm (D300 hè, đường & D800)", "m3", 1, "='📋 CSDL Tuyến Cống'!G4+'📋 CSDL Tuyến Cống'!G2", 0.85, 0.10, 1.0, "=F15*G15*H15*I15*J15", "=F15*G15*H15*I15*J15", "Đệm cát đáy hào dày 10cm"),
    ("Đắp cát bao bọc bảo vệ ống cống đến lưng cống +20cm (K95)", "m3", 1, "='📋 CSDL Tuyến Cống'!G4+'📋 CSDL Tuyến Cống'!G2", 0.85, 0.45, 1.0, "=F16*G16*(H16*I16-3.1416*0.35^2/4)*J16", "=F16*G16*(H16*I16-3.1416*0.35^2/4)*J16", "Cát bọc thân ống"),
    ("Đắp đất hoàn trả rãnh cống đầm chặt K95 (trừ kết cấu áo đường & cát)", "m3", 1, "=L13-L15-L16-('📋 CSDL Tuyến Cống'!G2*0.9*0.5)-('📋 CSDL Tuyến Cống'!G3*1.5*0.6)", 1.0, 1.0, 1.0, "=G17*J17", "=G17*J17", "Đất đắp hoàn trả K95"),
]
r_idx = 15
for item in details_culvert_backfill:
    ws2.cell(row=r_idx, column=4, value=item[0])
    ws2.cell(row=r_idx, column=5, value=item[1]).alignment = ALIGN_CENTER
    ws2.cell(row=r_idx, column=6, value=item[2]).alignment = ALIGN_CENTER
    ws2.cell(row=r_idx, column=7, value=item[3]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=7).number_format = '#,##0.00'
    ws2.cell(row=r_idx, column=8, value=item[4]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=9, value=item[5]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=10, value=item[6]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=11, value=item[7]).alignment = ALIGN_LEFT
    ws2.cell(row=r_idx, column=12, value=item[8]).alignment = ALIGN_RIGHT
    ws2.cell(row=r_idx, column=12).font = FONT_BOLD
    ws2.cell(row=r_idx, column=12).number_format = '#,##0.00'
    ws2.cell(row=r_idx, column=13, value=item[9])
    for c in range(2, 14):
        style_cell(ws2.cell(row=r_idx, column=c), font=FONT_REGULAR)
    style_cell(ws2.cell(row=r_idx, column=12), font=FONT_BOLD)
    r_idx += 1

# Subtotal row for Item 2
ws2.cell(row=18, column=4, value="CỘNG CÔNG TÁC ĐẮP CÁT & ĐẤT HOÀN TRẢ CỐNG (AB.65111)").font = FONT_BOLD
ws2.cell(row=18, column=5, value="m3").alignment = ALIGN_CENTER
ws2.cell(row=18, column=12, value="=SUM(L15:L17)").font = FONT_BOLD
ws2.cell(row=18, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws2.cell(row=18, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 3. Vận chuyển đất thừa
ws2.cell(row=19, column=2, value="3").alignment = ALIGN_CENTER
ws2.cell(row=19, column=3, value="AB.71112").alignment = ALIGN_CENTER
ws2.cell(row=19, column=4, value="Vận chuyển đất đào thừa đổ đi bãi tập kết quy định bằng ô tô tự đổ 10T")
ws2.cell(row=19, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws2.cell(row=19, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws2.cell(row=20, column=4, value="Đất thừa rãnh cống đào ra không dùng đắp (bằng Đào - Đắp đất)")
ws2.cell(row=20, column=5, value="m3").alignment = ALIGN_CENTER
ws2.cell(row=20, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=20, column=7, value="=L13-L17").alignment = ALIGN_RIGHT
ws2.cell(row=20, column=7).number_format = '#,##0.00'
ws2.cell(row=20, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=20, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=20, column=10, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=20, column=11, value="=G20*J20").alignment = ALIGN_LEFT
ws2.cell(row=20, column=12, value="=G20*J20").alignment = ALIGN_RIGHT
ws2.cell(row=20, column=12).font = FONT_BOLD
ws2.cell(row=20, column=12).number_format = '#,##0.00'
ws2.cell(row=20, column=13, value="Đất thừa = Đào hào - Đắp hoàn trả")
for c in range(2, 14):
    style_cell(ws2.cell(row=20, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=20, column=12), font=FONT_BOLD)

ws2.cell(row=21, column=4, value="Đất thừa nở rời vận chuyển ô tô (hệ số nở rời k=1.15)")
ws2.cell(row=21, column=5, value="m3").alignment = ALIGN_CENTER
ws2.cell(row=21, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=21, column=7, value="=L20").alignment = ALIGN_RIGHT
ws2.cell(row=21, column=7).number_format = '#,##0.00'
ws2.cell(row=21, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=21, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=21, column=10, value=0.15).alignment = ALIGN_RIGHT
ws2.cell(row=21, column=11, value="=G21*J21").alignment = ALIGN_LEFT
ws2.cell(row=21, column=12, value="=G21*J21").alignment = ALIGN_RIGHT
ws2.cell(row=21, column=12).font = FONT_BOLD
ws2.cell(row=21, column=12).number_format = '#,##0.00'
ws2.cell(row=21, column=13, value="Độ nở rời 15%")
for c in range(2, 14):
    style_cell(ws2.cell(row=21, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=21, column=12), font=FONT_BOLD)

# Subtotal row for Item 3
ws2.cell(row=22, column=4, value="CỘNG CÔNG TÁC VẬN CHUYỂN ĐẤT THỪA (AB.71112)").font = FONT_BOLD
ws2.cell(row=22, column=5, value="m3").alignment = ALIGN_CENTER
ws2.cell(row=22, column=12, value="=SUM(L20:L21)").font = FONT_BOLD
ws2.cell(row=22, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws2.cell(row=22, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 4. Lắp đặt cống HDPE D300 vỉa hè
ws2.cell(row=23, column=2, value="4").alignment = ALIGN_CENTER
ws2.cell(row=23, column=3, value="BB.11101").alignment = ALIGN_CENTER
ws2.cell(row=23, column=4, value="Cung cấp và lắp đặt cống HDPE 2 lớp SN4 D300 trên vỉa hè / rãnh kỹ thuật")
ws2.cell(row=23, column=5, value="m").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws2.cell(row=23, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws2.cell(row=24, column=4, value="Tuyến cống HDPE D300 chạy dọc vỉa hè (852 đoạn cống - trích xuất CAD)")
ws2.cell(row=24, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=24, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=24, column=7, value="='📋 CSDL Tuyến Cống'!G4").alignment = ALIGN_RIGHT
ws2.cell(row=24, column=7).number_format = '#,##0.00'
ws2.cell(row=24, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=24, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=24, column=10, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=24, column=11, value="=F24*G24*J24").alignment = ALIGN_LEFT
ws2.cell(row=24, column=12, value="=F24*G24*J24").alignment = ALIGN_RIGHT
ws2.cell(row=24, column=12).font = FONT_BOLD
ws2.cell(row=24, column=12).number_format = '#,##0.00'
ws2.cell(row=24, column=13, value="HDPE gân sóng 2 lớp SN4")
for c in range(2, 14):
    style_cell(ws2.cell(row=24, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=24, column=12), font=FONT_BOLD)

ws2.cell(row=25, column=4, value="Hao hụt cắt nối ống tại các vị trí đầu ga (1.0%)")
ws2.cell(row=25, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=25, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=25, column=7, value="=L24").alignment = ALIGN_RIGHT
ws2.cell(row=25, column=7).number_format = '#,##0.00'
ws2.cell(row=25, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=25, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=25, column=10, value=0.01).alignment = ALIGN_RIGHT
ws2.cell(row=25, column=11, value="=G25*J25").alignment = ALIGN_LEFT
ws2.cell(row=25, column=12, value="=G25*J25").alignment = ALIGN_RIGHT
ws2.cell(row=25, column=12).font = FONT_BOLD
ws2.cell(row=25, column=12).number_format = '#,##0.00'
ws2.cell(row=25, column=13, value="Định mức hao hụt cống BXD")
for c in range(2, 14):
    style_cell(ws2.cell(row=25, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=25, column=12), font=FONT_BOLD)

# Subtotal row for Item 4
ws2.cell(row=26, column=4, value="CỘNG CÔNG TÁC CỐNG HDPE D300 VỈA HÈ (BB.11101)").font = FONT_BOLD
ws2.cell(row=26, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=26, column=12, value="=SUM(L24:L25)").font = FONT_BOLD
ws2.cell(row=26, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws2.cell(row=26, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 5. Lắp đặt cống HDPE D300 qua đường
ws2.cell(row=27, column=2, value="5").alignment = ALIGN_CENTER
ws2.cell(row=27, column=3, value="BB.11102").alignment = ALIGN_CENTER
ws2.cell(row=27, column=4, value="Cung cấp và lắp đặt cống HDPE 2 vách SN8 D300 chịu lực dưới lòng đường")
ws2.cell(row=27, column=5, value="m").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws2.cell(row=27, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws2.cell(row=28, column=4, value="Đoạn cống HDPE D300 băng ngang qua lòng đường (106 đoạn - trích xuất CAD)")
ws2.cell(row=28, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=28, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=28, column=7, value="='📋 CSDL Tuyến Cống'!G2").alignment = ALIGN_RIGHT
ws2.cell(row=28, column=7).number_format = '#,##0.00'
ws2.cell(row=28, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=28, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=28, column=10, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=28, column=11, value="=F28*G28*J28").alignment = ALIGN_LEFT
ws2.cell(row=28, column=12, value="=F28*G28*J28").alignment = ALIGN_RIGHT
ws2.cell(row=28, column=12).font = FONT_BOLD
ws2.cell(row=28, column=12).number_format = '#,##0.00'
ws2.cell(row=28, column=13, value="HDPE 2 vách gân xoắn SN8")
for c in range(2, 14):
    style_cell(ws2.cell(row=28, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=28, column=12), font=FONT_BOLD)

ws2.cell(row=29, column=4, value="Hao hụt cắt nối cống qua đường (1.0%)")
ws2.cell(row=29, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=29, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=29, column=7, value="=L28").alignment = ALIGN_RIGHT
ws2.cell(row=29, column=7).number_format = '#,##0.00'
ws2.cell(row=29, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=29, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=29, column=10, value=0.01).alignment = ALIGN_RIGHT
ws2.cell(row=29, column=11, value="=G29*J29").alignment = ALIGN_LEFT
ws2.cell(row=29, column=12, value="=G29*J29").alignment = ALIGN_RIGHT
ws2.cell(row=29, column=12).font = FONT_BOLD
ws2.cell(row=29, column=12).number_format = '#,##0.00'
ws2.cell(row=29, column=13, value="Định mức hao hụt cống BXD")
for c in range(2, 14):
    style_cell(ws2.cell(row=29, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=29, column=12), font=FONT_BOLD)

# Subtotal row for Item 5
ws2.cell(row=30, column=4, value="CỘNG CÔNG TÁC CỐNG HDPE D300 QUA ĐƯỜNG (BB.11102)").font = FONT_BOLD
ws2.cell(row=30, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=30, column=12, value="=SUM(L28:L29)").font = FONT_BOLD
ws2.cell(row=30, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws2.cell(row=30, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 6. Lắp đặt cống BTCT D800
ws2.cell(row=31, column=2, value="6").alignment = ALIGN_CENTER
ws2.cell(row=31, column=3, value="BB.12104").alignment = ALIGN_CENTER
ws2.cell(row=31, column=4, value="Cung cấp và lắp đặt cống bê tông cốt thép đúc sẵn D800 chịu tải trọng TC")
ws2.cell(row=31, column=5, value="m").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws2.cell(row=31, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws2.cell(row=32, column=4, value="Đoạn cống BTCT D800 đúc sẵn qua đường (9 đoạn - trích xuất CAD)")
ws2.cell(row=32, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=32, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=32, column=7, value="='📋 CSDL Tuyến Cống'!G3").alignment = ALIGN_RIGHT
ws2.cell(row=32, column=7).number_format = '#,##0.00'
ws2.cell(row=32, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=32, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=32, column=10, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=32, column=11, value="=F32*G32*J32").alignment = ALIGN_LEFT
ws2.cell(row=32, column=12, value="=F32*G32*J32").alignment = ALIGN_RIGHT
ws2.cell(row=32, column=12).font = FONT_BOLD
ws2.cell(row=32, column=12).number_format = '#,##0.00'
ws2.cell(row=32, column=13, value="Cống BTCT D800 đúc sẵn tải H30")
for c in range(2, 14):
    style_cell(ws2.cell(row=32, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=32, column=12), font=FONT_BOLD)

ws2.cell(row=33, column=4, value="Hao hụt cống BTCT mối nối (1.0%)")
ws2.cell(row=33, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=33, column=6, value=1).alignment = ALIGN_CENTER
ws2.cell(row=33, column=7, value="=L32").alignment = ALIGN_RIGHT
ws2.cell(row=33, column=7).number_format = '#,##0.00'
ws2.cell(row=33, column=8, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=33, column=9, value=1.0).alignment = ALIGN_RIGHT
ws2.cell(row=33, column=10, value=0.01).alignment = ALIGN_RIGHT
ws2.cell(row=33, column=11, value="=G33*J33").alignment = ALIGN_LEFT
ws2.cell(row=33, column=12, value="=G33*J33").alignment = ALIGN_RIGHT
ws2.cell(row=33, column=12).font = FONT_BOLD
ws2.cell(row=33, column=12).number_format = '#,##0.00'
ws2.cell(row=33, column=13, value="Định mức hao hụt cống BXD")
for c in range(2, 14):
    style_cell(ws2.cell(row=33, column=c), font=FONT_REGULAR)
style_cell(ws2.cell(row=33, column=12), font=FONT_BOLD)

# Subtotal row for Item 6
ws2.cell(row=34, column=4, value="CỘNG CÔNG TÁC CỐNG BTCT D800 (BB.12104)").font = FONT_BOLD
ws2.cell(row=34, column=5, value="m").alignment = ALIGN_CENTER
ws2.cell(row=34, column=12, value="=SUM(L32:L33)").font = FONT_BOLD
ws2.cell(row=34, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws2.cell(row=34, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)


# ==========================================
# SHEET 3: 📊 Bóc Tách Hố Ga (WBS 12 Cột)
# ==========================================
ws3 = wb.create_sheet(title="📊 Bóc Tách Hố Ga")
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells("B2:L2")
ws3["B2"] = "BẢNG TÍNH BÓC TÁCH CHI TIẾT KHỐI LƯỢNG HỐ GA THOÁT NƯỚC THẢI (TB41)"
ws3["B2"].font = FONT_TITLE

ws3.merge_cells("B3:L3")
ws3["B3"] = "Bóc tách 990 hố ga thực tế trích xuất từ CAD theo 3 cấp độ sâu tiêu chuẩn Bộ Xây dựng | 100% CÔNG THỨC SỐNG"
ws3["B3"].font = FONT_SUBTITLE

for c_idx, h in enumerate(wbs_headers, start=2):
    cell = ws3.cell(row=5, column=c_idx, value=h)
    style_cell(cell, font=FONT_HEADER, fill=FILL_NAVY, alignment=ALIGN_CENTER)
ws3.row_dimensions[5].height = 28

ws3.cell(row=6, column=2, value="B").alignment = ALIGN_CENTER
ws3.cell(row=6, column=4, value="HỐ GA THOÁT NƯỚC THẢI (990 HỐ GA - TOÀN KHU B9.2, B9.3, B9.4)")
for c in range(2, 14):
    style_cell(ws3.cell(row=6, column=c), font=FONT_SECTION, fill=FILL_SECTION)

# 7. Đào móng hố ga
ws3.cell(row=7, column=2, value="7").alignment = ALIGN_CENTER
ws3.cell(row=7, column=3, value="AB.25111").alignment = ALIGN_CENTER
ws3.cell(row=7, column=4, value="Đào đất hố móng ga thoát nước thải bằng máy đào kết hợp thủ công (đất cấp II)")
ws3.cell(row=7, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=7, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

details_ga_excav = [
    ("Đào hố móng ga Cấp 1 (H <= 1.5m)", "m3", "=COUNTIF('📋 CSDL Hố Ga'!L7:L996, \"*Cấp 1*\")", 1.70, 1.70, "=AVERAGEIFS('📋 CSDL Hố Ga'!J7:J996, '📋 CSDL Hố Ga'!L7:L996, \"*Cấp 1*\")+0.1", 1.0, "=F8*G8*H8*I8*(1+0.25*I8/G8)*J8", "=F8*G8*H8*I8*(1+0.25*I8/G8)*J8", "Đáy móng 1.7x1.7m taluy 0.25"),
    ("Đào hố móng ga Cấp 2 (1.5m < H <= 2.5m)", "m3", "=COUNTIF('📋 CSDL Hố Ga'!L7:L996, \"*Cấp 2*\")", 1.80, 1.80, "=AVERAGEIFS('📋 CSDL Hố Ga'!J7:J996, '📋 CSDL Hố Ga'!L7:L996, \"*Cấp 2*\")+0.1", 1.0, "=F9*G9*H9*I9*(1+0.25*I9/G9)*J9", "=F9*G9*H9*I9*(1+0.25*I9/G9)*J9", "Đáy móng 1.8x1.8m taluy 0.25"),
    ("Đào hố móng ga Cấp 3 (H > 2.5m)", "m3", "=COUNTIF('📋 CSDL Hố Ga'!L7:L996, \"*Cấp 3*\")", 2.00, 2.00, "=AVERAGEIFS('📋 CSDL Hố Ga'!J7:J996, '📋 CSDL Hố Ga'!L7:L996, \"*Cấp 3*\")+0.1", 1.0, "=F10*G10*H10*I10*(1+0.25*I10/G10)*J10", "=F10*G10*H10*I10*(1+0.25*I10/G10)*J10", "Đáy móng 2.0x2.0m taluy 0.25"),
]

r_idx = 8
for item in details_ga_excav:
    ws3.cell(row=r_idx, column=4, value=item[0])
    ws3.cell(row=r_idx, column=5, value=item[1]).alignment = ALIGN_CENTER
    ws3.cell(row=r_idx, column=6, value=item[2]).alignment = ALIGN_CENTER
    ws3.cell(row=r_idx, column=6).number_format = '#,##0'
    ws3.cell(row=r_idx, column=7, value=item[3]).alignment = ALIGN_RIGHT
    ws3.cell(row=r_idx, column=8, value=item[4]).alignment = ALIGN_RIGHT
    ws3.cell(row=r_idx, column=9, value=item[5]).alignment = ALIGN_RIGHT
    ws3.cell(row=r_idx, column=9).number_format = '#,##0.00'
    ws3.cell(row=r_idx, column=10, value=item[6]).alignment = ALIGN_RIGHT
    ws3.cell(row=r_idx, column=11, value=item[7]).alignment = ALIGN_LEFT
    ws3.cell(row=r_idx, column=12, value=item[8]).alignment = ALIGN_RIGHT
    ws3.cell(row=r_idx, column=12).font = FONT_BOLD
    ws3.cell(row=r_idx, column=12).number_format = '#,##0.00'
    ws3.cell(row=r_idx, column=13, value=item[9])
    for c in range(2, 14):
        style_cell(ws3.cell(row=r_idx, column=c), font=FONT_REGULAR)
    style_cell(ws3.cell(row=r_idx, column=12), font=FONT_BOLD)
    r_idx += 1

# Subtotal row for Item 7
ws3.cell(row=14, column=4, value="CỘNG CÔNG TÁC ĐÀO ĐẤT HỐ MÓNG GA (AB.25111)").font = FONT_BOLD
ws3.cell(row=14, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=14, column=12, value="=SUM(L8:L10)").font = FONT_BOLD
ws3.cell(row=14, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=14, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 8. Bê tông lót móng hố ga M100
ws3.cell(row=15, column=2, value="8").alignment = ALIGN_CENTER
ws3.cell(row=15, column=3, value="AF.11111").alignment = ALIGN_CENTER
ws3.cell(row=15, column=4, value="Bê tông lót móng hố ga M100 đá 4x6 dày 100mm")
ws3.cell(row=15, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=15, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=16, column=4, value="Bê tông lót đáy hố móng 990 hố ga (kích thước lót 1.5m x 1.5m x 0.1m)")
ws3.cell(row=16, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=16, column=6, value="=COUNTA('📋 CSDL Hố Ga'!C7:C996)").alignment = ALIGN_CENTER
ws3.cell(row=16, column=7, value=1.50).alignment = ALIGN_RIGHT
ws3.cell(row=16, column=8, value=1.50).alignment = ALIGN_RIGHT
ws3.cell(row=16, column=9, value=0.10).alignment = ALIGN_RIGHT
ws3.cell(row=16, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=16, column=11, value="=F16*G16*H16*I16*J16").alignment = ALIGN_LEFT
ws3.cell(row=16, column=12, value="=F16*G16*H16*I16*J16").alignment = ALIGN_RIGHT
ws3.cell(row=16, column=12).font = FONT_BOLD
ws3.cell(row=16, column=12).number_format = '#,##0.00'
ws3.cell(row=16, column=13, value="Lót đá 4x6 M100 dày 10cm")
for c in range(2, 14):
    style_cell(ws3.cell(row=16, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=16, column=12), font=FONT_BOLD)

# Subtotal row for Item 8
ws3.cell(row=18, column=4, value="CỘNG BÊ TÔNG LÓT MÓNG HỐ GA (AF.11111)").font = FONT_BOLD
ws3.cell(row=18, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=18, column=12, value="=L16").font = FONT_BOLD
ws3.cell(row=18, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=18, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 9. Ván khuôn lót móng
ws3.cell(row=19, column=2, value="9").alignment = ALIGN_CENTER
ws3.cell(row=19, column=3, value="AF.81111").alignment = ALIGN_CENTER
ws3.cell(row=19, column=4, value="Ván khuôn bê tông lót móng hố ga")
ws3.cell(row=19, column=5, value="m2").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=19, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=20, column=4, value="Ván khuôn thành lớp bê tông lót đáy móng (chu vi 4x1.5m, cao 0.1m)")
ws3.cell(row=20, column=5, value="m2").alignment = ALIGN_CENTER
ws3.cell(row=20, column=6, value="=COUNTA('📋 CSDL Hố Ga'!C7:C996)").alignment = ALIGN_CENTER
ws3.cell(row=20, column=7, value=6.00).alignment = ALIGN_RIGHT
ws3.cell(row=20, column=8, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=20, column=9, value=0.10).alignment = ALIGN_RIGHT
ws3.cell(row=20, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=20, column=11, value="=F20*G20*I20*J20").alignment = ALIGN_LEFT
ws3.cell(row=20, column=12, value="=F20*G20*I20*J20").alignment = ALIGN_RIGHT
ws3.cell(row=20, column=12).font = FONT_BOLD
ws3.cell(row=20, column=12).number_format = '#,##0.00'
ws3.cell(row=20, column=13, value="Chu vi lót móng 4 x 1.5m")
for c in range(2, 14):
    style_cell(ws3.cell(row=20, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=20, column=12), font=FONT_BOLD)

# Subtotal row for Item 9
ws3.cell(row=22, column=4, value="CỘNG VÁN KHUÔN LÓT MÓNG HỐ GA (AF.81111)").font = FONT_BOLD
ws3.cell(row=22, column=5, value="m2").alignment = ALIGN_CENTER
ws3.cell(row=22, column=12, value="=L20").font = FONT_BOLD
ws3.cell(row=22, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=22, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 10. Bê tông thành và đáy hố ga C20
ws3.cell(row=23, column=2, value="10").alignment = ALIGN_CENTER
ws3.cell(row=23, column=3, value="AF.21111").alignment = ALIGN_CENTER
ws3.cell(row=23, column=4, value="Bê tông thành và đáy hố ga C20/M200 (đúc sẵn hoặc đổ tại chỗ)")
ws3.cell(row=23, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=23, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=24, column=4, value="Bê tông bản đáy hố ga dày 150mm (phủ bì 1.3m x 1.3m)")
ws3.cell(row=24, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=24, column=6, value="=COUNTA('📋 CSDL Hố Ga'!C7:C996)").alignment = ALIGN_CENTER
ws3.cell(row=24, column=7, value=1.30).alignment = ALIGN_RIGHT
ws3.cell(row=24, column=8, value=1.30).alignment = ALIGN_RIGHT
ws3.cell(row=24, column=9, value=0.15).alignment = ALIGN_RIGHT
ws3.cell(row=24, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=24, column=11, value="=F24*G24*H24*I24*J24").alignment = ALIGN_LEFT
ws3.cell(row=24, column=12, value="=F24*G24*H24*I24*J24").alignment = ALIGN_RIGHT
ws3.cell(row=24, column=12).font = FONT_BOLD
ws3.cell(row=24, column=12).number_format = '#,##0.00'
ws3.cell(row=24, column=13, value="Bản đáy dày 15cm")
for c in range(2, 14):
    style_cell(ws3.cell(row=24, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=24, column=12), font=FONT_BOLD)

ws3.cell(row=25, column=4, value="Bê tông thành hố ga C20 (thành dày 150mm, tiết diện F=0.69m2)")
ws3.cell(row=25, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=25, column=6, value="=COUNTA('📋 CSDL Hố Ga'!C7:C996)").alignment = ALIGN_CENTER
ws3.cell(row=25, column=7, value=0.69).alignment = ALIGN_RIGHT
ws3.cell(row=25, column=8, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=25, column=9, value="=AVERAGE('📋 CSDL Hố Ga'!J7:J996)-0.15").alignment = ALIGN_RIGHT
ws3.cell(row=25, column=9).number_format = '#,##0.00'
ws3.cell(row=25, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=25, column=11, value="=F25*G25*I25*J25").alignment = ALIGN_LEFT
ws3.cell(row=25, column=12, value="=F25*G25*I25*J25").alignment = ALIGN_RIGHT
ws3.cell(row=25, column=12).font = FONT_BOLD
ws3.cell(row=25, column=12).number_format = '#,##0.00'
ws3.cell(row=25, column=13, value="Thành dày 15cm, F=1.3^2-1.0^2")
for c in range(2, 14):
    style_cell(ws3.cell(row=25, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=25, column=12), font=FONT_BOLD)

# Subtotal row for Item 10
ws3.cell(row=27, column=4, value="CỘNG BÊ TÔNG THÀNH VÀ ĐÁY HỐ GA C20 (AF.21111)").font = FONT_BOLD
ws3.cell(row=27, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=27, column=12, value="=SUM(L24:L25)").font = FONT_BOLD
ws3.cell(row=27, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=27, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 11. Ván khuôn thân hố ga
ws3.cell(row=28, column=2, value="11").alignment = ALIGN_CENTER
ws3.cell(row=28, column=3, value="AF.82111").alignment = ALIGN_CENTER
ws3.cell(row=28, column=4, value="Ván khuôn thân và thành hố ga (2 mặt trong và ngoài)")
ws3.cell(row=28, column=5, value="m2").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=28, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=29, column=4, value="Ván khuôn thành trong và ngoài hố ga (tổng chu vi 4x1.3m + 4x1.0m = 9.2m)")
ws3.cell(row=29, column=5, value="m2").alignment = ALIGN_CENTER
ws3.cell(row=29, column=6, value="=COUNTA('📋 CSDL Hố Ga'!C7:C996)").alignment = ALIGN_CENTER
ws3.cell(row=29, column=7, value=9.20).alignment = ALIGN_RIGHT
ws3.cell(row=29, column=8, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=29, column=9, value="=AVERAGE('📋 CSDL Hố Ga'!J7:J996)-0.15").alignment = ALIGN_RIGHT
ws3.cell(row=29, column=9).number_format = '#,##0.00'
ws3.cell(row=29, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=29, column=11, value="=F29*G29*I29*J29").alignment = ALIGN_LEFT
ws3.cell(row=29, column=12, value="=F29*G29*I29*J29").alignment = ALIGN_RIGHT
ws3.cell(row=29, column=12).font = FONT_BOLD
ws3.cell(row=29, column=12).number_format = '#,##0.00'
ws3.cell(row=29, column=13, value="2 mặt ván khuôn chu vi 9.2m")
for c in range(2, 14):
    style_cell(ws3.cell(row=29, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=29, column=12), font=FONT_BOLD)

# Subtotal row for Item 11
ws3.cell(row=31, column=4, value="CỘNG VÁN KHUÔN THÀNH HỐ GA (AF.82111)").font = FONT_BOLD
ws3.cell(row=31, column=5, value="m2").alignment = ALIGN_CENTER
ws3.cell(row=31, column=12, value="=L29").font = FONT_BOLD
ws3.cell(row=31, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=31, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 12. Cốt thép hố ga
ws3.cell(row=32, column=2, value="12").alignment = ALIGN_CENTER
ws3.cell(row=32, column=3, value="AF.61111").alignment = ALIGN_CENTER
ws3.cell(row=32, column=4, value="Cốt thép hố ga CB300-V (đường kính <= 18mm)")
ws3.cell(row=32, column=5, value="kg").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=32, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=33, column=4, value="Gia công lắp dựng cốt thép hố ga (hàm lượng thép 65 kg/m3 bê tông C20)")
ws3.cell(row=33, column=5, value="kg").alignment = ALIGN_CENTER
ws3.cell(row=33, column=6, value=1).alignment = ALIGN_CENTER
ws3.cell(row=33, column=7, value="=L27").alignment = ALIGN_RIGHT
ws3.cell(row=33, column=7).number_format = '#,##0.00'
ws3.cell(row=33, column=8, value=65.0).alignment = ALIGN_RIGHT
ws3.cell(row=33, column=9, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=33, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=33, column=11, value="=G33*H33*J33").alignment = ALIGN_LEFT
ws3.cell(row=33, column=12, value="=G33*H33*J33").alignment = ALIGN_RIGHT
ws3.cell(row=33, column=12).font = FONT_BOLD
ws3.cell(row=33, column=12).number_format = '#,##0.00'
ws3.cell(row=33, column=13, value="65 kg/m3 bê tông C20")
for c in range(2, 14):
    style_cell(ws3.cell(row=33, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=33, column=12), font=FONT_BOLD)

# Subtotal row for Item 12
ws3.cell(row=35, column=4, value="CỘNG CỐT THÉP HỐ GA (AF.61111)").font = FONT_BOLD
ws3.cell(row=35, column=5, value="kg").alignment = ALIGN_CENTER
ws3.cell(row=35, column=12, value="=L33").font = FONT_BOLD
ws3.cell(row=35, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=35, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 13. Nắp hố ga
ws3.cell(row=36, column=2, value="13").alignment = ALIGN_CENTER
ws3.cell(row=36, column=3, value="BB.81111").alignment = ALIGN_CENTER
ws3.cell(row=36, column=4, value="Sản xuất và lắp đặt nắp hố ga gang cầu / composite ngăn mùi / song chắn rác")
ws3.cell(row=36, column=5, value="bộ").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=36, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=37, column=4, value="Bộ nắp ga gang cầu / composite ngăn mùi chuẩn TB41 và ga D800")
ws3.cell(row=37, column=5, value="bộ").alignment = ALIGN_CENTER
ws3.cell(row=37, column=6, value="=COUNTA('📋 CSDL Hố Ga'!C7:C996)").alignment = ALIGN_CENTER
ws3.cell(row=37, column=7, value=1).alignment = ALIGN_RIGHT
ws3.cell(row=37, column=8, value=1).alignment = ALIGN_RIGHT
ws3.cell(row=37, column=9, value=1).alignment = ALIGN_RIGHT
ws3.cell(row=37, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=37, column=11, value="=F37*J37").alignment = ALIGN_LEFT
ws3.cell(row=37, column=12, value="=F37*J37").alignment = ALIGN_RIGHT
ws3.cell(row=37, column=12).font = FONT_BOLD
ws3.cell(row=37, column=12).number_format = '#,##0'
ws3.cell(row=37, column=13, value="1 nắp / 1 hố ga")
for c in range(2, 14):
    style_cell(ws3.cell(row=37, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=37, column=12), font=FONT_BOLD)

# Subtotal row for Item 13
ws3.cell(row=39, column=4, value="CỘNG NẮP HỐ GA (BB.81111)").font = FONT_BOLD
ws3.cell(row=39, column=5, value="bộ").alignment = ALIGN_CENTER
ws3.cell(row=39, column=12, value="=L37").font = FONT_BOLD
ws3.cell(row=39, column=12).number_format = '#,##0'
for c in range(2, 14):
    style_cell(ws3.cell(row=39, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 14. Đắp đất hố móng ga K95
ws3.cell(row=40, column=2, value="14").alignment = ALIGN_CENTER
ws3.cell(row=40, column=3, value="AB.65112").alignment = ALIGN_CENTER
ws3.cell(row=40, column=4, value="Đắp đất hố móng xung quanh thành ga hoàn trả đầm chặt K95")
ws3.cell(row=40, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=40, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=41, column=4, value="Đắp đất hoàn trả hố ga (Đào hố ga - Thể tích bê tông ga - Bê tông lót)")
ws3.cell(row=41, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=41, column=6, value=1).alignment = ALIGN_CENTER
ws3.cell(row=41, column=7, value="=L14-(COUNTA('📋 CSDL Hố Ga'!C7:C996)*1.3*1.3*AVERAGE('📋 CSDL Hố Ga'!J7:J996))-L18").alignment = ALIGN_RIGHT
ws3.cell(row=41, column=7).number_format = '#,##0.00'
ws3.cell(row=41, column=8, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=41, column=9, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=41, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=41, column=11, value="=G41*J41").alignment = ALIGN_LEFT
ws3.cell(row=41, column=12, value="=G41*J41").alignment = ALIGN_RIGHT
ws3.cell(row=41, column=12).font = FONT_BOLD
ws3.cell(row=41, column=12).number_format = '#,##0.00'
ws3.cell(row=41, column=13, value="Đắp đất xung quanh mang hố ga")
for c in range(2, 14):
    style_cell(ws3.cell(row=41, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=41, column=12), font=FONT_BOLD)

# Subtotal row for Item 14
ws3.cell(row=43, column=4, value="CỘNG CÔNG TÁC ĐẮP ĐẤT HOÀN TRẢ HỐ GA (AB.65112)").font = FONT_BOLD
ws3.cell(row=43, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=43, column=12, value="=L41").font = FONT_BOLD
ws3.cell(row=43, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=43, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# 15. Vận chuyển đất thừa hố ga
ws3.cell(row=44, column=2, value="15").alignment = ALIGN_CENTER
ws3.cell(row=44, column=3, value="AB.71113").alignment = ALIGN_CENTER
ws3.cell(row=44, column=4, value="Vận chuyển đất đào thừa hố ga đổ đi bãi tập kết quy định bằng ô tô tự đổ 10T")
ws3.cell(row=44, column=5, value="m3").alignment = ALIGN_CENTER
for c in range(2, 14):
    style_cell(ws3.cell(row=44, column=c), font=FONT_BOLD, fill=FILL_SUBTOTAL)

ws3.cell(row=45, column=4, value="Đất thừa hố ga do chiếm chỗ của bê tông ga và bê tông lót")
ws3.cell(row=45, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=45, column=6, value=1).alignment = ALIGN_CENTER
ws3.cell(row=45, column=7, value="=L14-L43").alignment = ALIGN_RIGHT
ws3.cell(row=45, column=7).number_format = '#,##0.00'
ws3.cell(row=45, column=8, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=45, column=9, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=45, column=10, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=45, column=11, value="=G45*J45").alignment = ALIGN_LEFT
ws3.cell(row=45, column=12, value="=G45*J45").alignment = ALIGN_RIGHT
ws3.cell(row=45, column=12).font = FONT_BOLD
ws3.cell(row=45, column=12).number_format = '#,##0.00'
ws3.cell(row=45, column=13, value="Đất thừa hố ga = Đào ga - Đắp ga")
for c in range(2, 14):
    style_cell(ws3.cell(row=45, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=45, column=12), font=FONT_BOLD)

ws3.cell(row=46, column=4, value="Đất thừa hố ga nở rời vận chuyển ô tô (hệ số k=1.15)")
ws3.cell(row=46, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=46, column=6, value=1).alignment = ALIGN_CENTER
ws3.cell(row=46, column=7, value="=L45").alignment = ALIGN_RIGHT
ws3.cell(row=46, column=7).number_format = '#,##0.00'
ws3.cell(row=46, column=8, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=46, column=9, value=1.0).alignment = ALIGN_RIGHT
ws3.cell(row=46, column=10, value=0.15).alignment = ALIGN_RIGHT
ws3.cell(row=46, column=11, value="=G46*J46").alignment = ALIGN_LEFT
ws3.cell(row=46, column=12, value="=G46*J46").alignment = ALIGN_RIGHT
ws3.cell(row=46, column=12).font = FONT_BOLD
ws3.cell(row=46, column=12).number_format = '#,##0.00'
ws3.cell(row=46, column=13, value="Độ nở rời 15%")
for c in range(2, 14):
    style_cell(ws3.cell(row=46, column=c), font=FONT_REGULAR)
style_cell(ws3.cell(row=46, column=12), font=FONT_BOLD)

# Subtotal row for Item 15
ws3.cell(row=47, column=4, value="CỘNG VẬN CHUYỂN ĐẤT THỪA HỐ GA (AB.71113)").font = FONT_BOLD
ws3.cell(row=47, column=5, value="m3").alignment = ALIGN_CENTER
ws3.cell(row=47, column=12, value="=SUM(L45:L46)").font = FONT_BOLD
ws3.cell(row=47, column=12).number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws3.cell(row=47, column=c), font=FONT_BOLD, fill=FILL_TOTAL, border=BORDER_TOTAL)

# ==========================================
# SHEET 4: 📋 CSDL Hố Ga (ĐẦY ĐỦ TRỤC X - Y VN-2000)
# ==========================================
ws4 = wb.create_sheet(title="📋 CSDL Hố Ga")
ws4.views.sheetView[0].showGridLines = True

ws4.merge_cells("B2:L2")
ws4["B2"] = "CƠ SỞ DỮ LIỆU TOÀN BỘ 990 HỐ GA THOÁT NƯỚC THẢI (TRÍCH XUẤT CAD 100%)"
ws4["B2"].font = FONT_TITLE

ws4.merge_cells("B3:L3")
ws4["B3"] = "Bao gồm Handle CAD, Tên ga, Loại ga, Tọa độ Trục X (Bắc), Trục Y (Đông) VN-2000, Cốt đỉnh, Cốt đáy, Chiều sâu H và Phân cấp BXD"
ws4["B3"].font = FONT_SUBTITLE

ga_table_headers = [
    "STT", "Handle CAD", "Tên Hố Ga", "Loại Ga", 
    "Tọa độ Trục X (VN-2000 / Bắc)", "Tọa độ Trục Y (VN-2000 / Đông)", 
    "Cốt đỉnh MG (m)", "Cốt đáy DC (m)", "Chiều sâu H (m)", "Công thức H", "Phân cấp BXD", "Trạng thái QA/QC"
]
for c_idx, h in enumerate(ga_table_headers, start=2):
    cell = ws4.cell(row=5, column=c_idx, value=h)
    style_cell(cell, font=FONT_HEADER, fill=FILL_NAVY, alignment=ALIGN_CENTER)
ws4.row_dimensions[5].height = 28

# Summary Row (Row 6)
ws4.cell(row=6, column=2, value="TỔNG").alignment = ALIGN_CENTER
ws4.cell(row=6, column=4, value="990 HỐ GA")
ws4.cell(row=6, column=10, value="TB Chiều sâu =")
ws4.cell(row=6, column=10).alignment = ALIGN_RIGHT
ws4.cell(row=6, column=10).font = FONT_BOLD
c_avg_h = ws4.cell(row=6, column=11, value="=AVERAGE(J7:J996)")
c_avg_h.font = FONT_BOLD
c_avg_h.number_format = '#,##0.00'
for c in range(2, 14):
    style_cell(ws4.cell(row=6, column=c), font=FONT_BOLD, fill=FILL_SECTION)

r_ga = 7
for idx, r in df_manholes.iterrows():
    ws4.cell(row=r_ga, column=2, value=idx+1).alignment = ALIGN_CENTER
    ws4.cell(row=r_ga, column=3, value=r['Handle']).alignment = ALIGN_CENTER
    ws4.cell(row=r_ga, column=4, value=r['TenGa']).alignment = ALIGN_CENTER
    ws4.cell(row=r_ga, column=5, value=r['LoaiGa']).alignment = ALIGN_CENTER
    
    # Exact VN-2000 Coordinates
    cX = ws4.cell(row=r_ga, column=6, value=r['X_VN2000'])
    cX.alignment = ALIGN_RIGHT
    cX.number_format = '#,##0.000'
    
    cY = ws4.cell(row=r_ga, column=7, value=r['Y_VN2000'])
    cY.alignment = ALIGN_RIGHT
    cY.number_format = '#,##0.000'
    
    cMG = ws4.cell(row=r_ga, column=8, value=r['MG'])
    cMG.alignment = ALIGN_RIGHT
    cMG.number_format = '#,##0.00'
    
    cDC = ws4.cell(row=r_ga, column=9, value=r['DC'])
    cDC.alignment = ALIGN_RIGHT
    cDC.number_format = '#,##0.00'
    
    # 100% LIVE FORMULA FOR H = MG - DC!
    cH = ws4.cell(row=r_ga, column=10, value=f"=H{r_ga}-I{r_ga}")
    cH.alignment = ALIGN_RIGHT
    cH.font = FONT_BOLD
    cH.number_format = '#,##0.00'
    
    ws4.cell(row=r_ga, column=11, value=f"H{r_ga}-I{r_ga}").alignment = ALIGN_CENTER
    ws4.cell(row=r_ga, column=12, value=r['CapSau']).alignment = ALIGN_LEFT
    
    cQA = ws4.cell(row=r_ga, column=13, value=r['QA_Status'])
    cQA.alignment = ALIGN_LEFT
    if "LỖI" in r['QA_Status']:
        cQA.font = Font(name="Segoe UI", size=9, bold=True, color="FF0000")
    elif "LƯU Ý" in r['QA_Status']:
        cQA.font = Font(name="Segoe UI", size=9, bold=True, color="FF8C00")
    else:
        cQA.font = Font(name="Segoe UI", size=9, color="008000")
        
    for c in range(2, 14):
        if c not in [10, 13]:
            style_cell(ws4.cell(row=r_ga, column=c), font=FONT_REGULAR)
        else:
            style_cell(ws4.cell(row=r_ga, column=c))
    r_ga += 1

# ==========================================
# SHEET 5: 📋 CSDL Tuyến Cống (ĐẦY ĐỦ TRỤC X - Y ĐIỂM ĐẦU & ĐIỂM CUỐI)
# ==========================================
ws5 = wb.create_sheet(title="📋 CSDL Tuyến Cống")
ws5.views.sheetView[0].showGridLines = True

ws5.merge_cells("B2:E2")
ws5["B2"] = "CƠ SỞ DỮ LIỆU TOÀN BỘ 967 ĐOẠN CỐNG THOÁT NƯỚC THẢI (CAD 100%)"
ws5["B2"].font = FONT_TITLE

ws5.merge_cells("B3:E3")
ws5["B3"] = "Bao gồm Handle CAD, Phân loại tuyến, Ga đầu, Ga cuối, Tọa độ Trục X - Y điểm đầu & cuối, Chiều dài L, Cốt đáy DC1, DC2, Độ dốc i"
ws5["B3"].font = FONT_SUBTITLE

# Summary stats on rows 2-4
ws5.cell(row=2, column=6, value="TỔNG D300 QUA ĐƯỜNG (m):")
ws5.cell(row=2, column=7, value=f"=SUMIFS(M7:M{len(df_culverts)+6}, D7:D{len(df_culverts)+6}, \"*qua đường*\", E7:E{len(df_culverts)+6}, 300)").font = FONT_BOLD
ws5.cell(row=2, column=7).number_format = '#,##0.00'

ws5.cell(row=3, column=6, value="TỔNG D800 QUA ĐƯỜNG (m):")
ws5.cell(row=3, column=7, value=f"=SUMIFS(M7:M{len(df_culverts)+6}, E7:E{len(df_culverts)+6}, 800)").font = FONT_BOLD
ws5.cell(row=3, column=7).number_format = '#,##0.00'

ws5.cell(row=4, column=6, value="TỔNG D300 VỈA HÈ (m):")
ws5.cell(row=4, column=7, value=f"=SUMIFS(M7:M{len(df_culverts)+6}, D7:D{len(df_culverts)+6}, \"*vỉa hè*\")").font = FONT_BOLD
ws5.cell(row=4, column=7).number_format = '#,##0.00'

culvert_table_headers = [
    "STT", "Handle CAD", "Phân loại tuyến", "Đường kính D (mm)", "Vị trí lắp đặt",
    "Ga đầu (Upstream)", "Ga cuối (Downstream)", 
    "Trục X Điểm Đầu (VN-2000 / Bắc)", "Trục Y Điểm Đầu (VN-2000 / Đông)", 
    "Trục X Điểm Cuối (VN-2000 / Bắc)", "Trục Y Điểm Cuối (VN-2000 / Đông)", 
    "Chiều dài L (m)", "Cốt đáy DC1 (m)", "Cốt đáy DC2 (m)", "Chênh cao dH (m)", "Độ dốc i (%)", "H_đào trung bình (m)"
]
for c_idx, h in enumerate(culvert_table_headers, start=2):
    cell = ws5.cell(row=6, column=c_idx, value=h)
    style_cell(cell, font=FONT_HEADER, fill=FILL_NAVY, alignment=ALIGN_CENTER)
ws5.row_dimensions[6].height = 28

r_c = 7
for idx, r in df_culverts.iterrows():
    ws5.cell(row=r_c, column=2, value=idx+1).alignment = ALIGN_CENTER
    ws5.cell(row=r_c, column=3, value=r['Handle']).alignment = ALIGN_CENTER
    ws5.cell(row=r_c, column=4, value=r['PhanLoai']).alignment = ALIGN_LEFT
    ws5.cell(row=r_c, column=5, value=r['DuongKinh']).alignment = ALIGN_CENTER
    ws5.cell(row=r_c, column=6, value=r['ViTri']).alignment = ALIGN_LEFT
    ws5.cell(row=r_c, column=7, value=r['GaDau']).alignment = ALIGN_CENTER
    ws5.cell(row=r_c, column=8, value=r['GaCuoi']).alignment = ALIGN_CENTER
    
    # 4 New Coordinate Columns:
    cX1 = ws5.cell(row=r_c, column=9, value=r['X_Dau'])
    cX1.alignment = ALIGN_RIGHT
    cX1.number_format = '#,##0.000'
    
    cY1 = ws5.cell(row=r_c, column=10, value=r['Y_Dau'])
    cY1.alignment = ALIGN_RIGHT
    cY1.number_format = '#,##0.000'
    
    cX2 = ws5.cell(row=r_c, column=11, value=r['X_Cuoi'])
    cX2.alignment = ALIGN_RIGHT
    cX2.number_format = '#,##0.000'
    
    cY2 = ws5.cell(row=r_c, column=12, value=r['Y_Cuoi'])
    cY2.alignment = ALIGN_RIGHT
    cY2.number_format = '#,##0.000'
    
    # Chiều dài L (m)
    cL = ws5.cell(row=r_c, column=13, value=r['ChieuDai_L'])
    cL.alignment = ALIGN_RIGHT
    cL.number_format = '#,##0.00'
    cL.font = FONT_BOLD
    
    # Cốt đáy DC1 (m)
    cDC1 = ws5.cell(row=r_c, column=14, value=r['DC1'])
    cDC1.alignment = ALIGN_RIGHT
    cDC1.number_format = '#,##0.00' if r['DC1'] is not None else '@'
    
    # Cốt đáy DC2 (m)
    cDC2 = ws5.cell(row=r_c, column=15, value=r['DC2'])
    cDC2.alignment = ALIGN_RIGHT
    cDC2.number_format = '#,##0.00' if r['DC2'] is not None else '@'
    
    # 100% LIVE FORMULA FOR dH = ABS(DC1 - DC2)
    cdH = ws5.cell(row=r_c, column=16, value=f"=IF(AND(ISNUMBER(N{r_c}),ISNUMBER(O{r_c})),ABS(N{r_c}-O{r_c}),0)")
    cdH.alignment = ALIGN_RIGHT
    cdH.number_format = '#,##0.00'
    
    # 100% LIVE FORMULA FOR SLOPE i = dH / L * 100
    cI = ws5.cell(row=r_c, column=17, value=f"=IF(M{r_c}>0,ROUND(P{r_c}/M{r_c}*100,3),0)")
    cI.alignment = ALIGN_RIGHT
    cI.number_format = '#,##0.00'
    
    # H_đào trung bình
    cHdao = ws5.cell(row=r_c, column=18, value=r['H_Dao_TB'])
    cHdao.alignment = ALIGN_RIGHT
    cHdao.number_format = '#,##0.00'
    
    for c in range(2, 19):
        if c not in [13]:
            style_cell(ws5.cell(row=r_c, column=c), font=FONT_REGULAR)
        else:
            style_cell(ws5.cell(row=r_c, column=c), font=FONT_BOLD)
    r_c += 1

# Adjust Column Widths
for ws in [ws1, ws2, ws3, ws4, ws5]:
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col[0].column == 1:
            ws.column_dimensions[col_letter].width = 3
            continue
        max_len = 0
        for cell in col:
            val = str(cell.value or '')
            if cell.row in [2, 3, 4] and ws.title in ["📑 Tổng Hợp BoQ", "📊 Bóc Tách Cống", "📊 Bóc Tách Hố Ga", "📋 CSDL Hố Ga", "📋 CSDL Tuyến Cống"]:
                continue
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

# Specific column widths
ws1.column_dimensions["B"].width = 6
ws1.column_dimensions["C"].width = 14
ws1.column_dimensions["D"].width = 55
ws1.column_dimensions["E"].width = 8
ws1.column_dimensions["F"].width = 16
ws1.column_dimensions["G"].width = 24
ws1.column_dimensions["H"].width = 40

ws2.column_dimensions["B"].width = 6
ws2.column_dimensions["C"].width = 14
ws2.column_dimensions["D"].width = 58
ws2.column_dimensions["E"].width = 8
ws2.column_dimensions["F"].width = 12
ws2.column_dimensions["G"].width = 16
ws2.column_dimensions["H"].width = 12
ws2.column_dimensions["I"].width = 12
ws2.column_dimensions["J"].width = 10
ws2.column_dimensions["K"].width = 25
ws2.column_dimensions["L"].width = 18
ws2.column_dimensions["M"].width = 32

ws3.column_dimensions["B"].width = 6
ws3.column_dimensions["C"].width = 14
ws3.column_dimensions["D"].width = 58
ws3.column_dimensions["E"].width = 8
ws3.column_dimensions["F"].width = 12
ws3.column_dimensions["G"].width = 14
ws3.column_dimensions["H"].width = 14
ws3.column_dimensions["I"].width = 14
ws3.column_dimensions["J"].width = 10
ws3.column_dimensions["K"].width = 25
ws3.column_dimensions["L"].width = 18
ws3.column_dimensions["M"].width = 32

ws4.column_dimensions["B"].width = 6
ws4.column_dimensions["C"].width = 12
ws4.column_dimensions["D"].width = 14
ws4.column_dimensions["E"].width = 10
ws4.column_dimensions["F"].width = 20
ws4.column_dimensions["G"].width = 20
ws4.column_dimensions["H"].width = 14
ws4.column_dimensions["I"].width = 14
ws4.column_dimensions["J"].width = 14
ws4.column_dimensions["K"].width = 14
ws4.column_dimensions["L"].width = 24
ws4.column_dimensions["M"].width = 30

ws5.column_dimensions["B"].width = 6
ws5.column_dimensions["C"].width = 12
ws5.column_dimensions["D"].width = 28
ws5.column_dimensions["E"].width = 12
ws5.column_dimensions["F"].width = 26
ws5.column_dimensions["G"].width = 14
ws5.column_dimensions["H"].width = 14
ws5.column_dimensions["I"].width = 20
ws5.column_dimensions["J"].width = 20
ws5.column_dimensions["K"].width = 20
ws5.column_dimensions["L"].width = 20
ws5.column_dimensions["M"].width = 16
ws5.column_dimensions["N"].width = 14
ws5.column_dimensions["O"].width = 14
ws5.column_dimensions["P"].width = 14
ws5.column_dimensions["Q"].width = 14
ws5.column_dimensions["R"].width = 16

# Save in folder
target_folder = r"C:\Users\baotu\Downloads\261008.MB TNT LÔ B9.2-3-4 Layout1 (12)_Marker"
out_excel = os.path.join(target_folder, "Boc_Tach_Khoi_Luong_Cong_HoGa_B9_CAD_Master.xlsx")
wb.save(out_excel)

print(f"SUCCESS: Saved Master Excel with Full Trục X-Y to:\n{out_excel}", flush=True)
