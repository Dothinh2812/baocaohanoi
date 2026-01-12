# -*- coding: utf-8 -*-
"""
Module tạo báo cáo Word tự động cho KPI/BSC NVKT
Tạo file Word với bảng biểu và biểu đồ

Author: Auto-generated
Date: 2026-01-08
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import os

# Thiết lập matplotlib để hỗ trợ tiếng Việt
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.use('Agg')  # Use non-interactive backend

# =============================================================================
# CẤU HÌNH
# =============================================================================
DEFAULT_KPI_FOLDER = "downloads/KPI"
DEFAULT_OUTPUT_FOLDER = "downloads/reports"

# Mapping tên đơn vị ngắn gọn
TEAM_SHORT_NAMES = {
    "Tổ Kỹ thuật Địa bàn Phúc Thọ": "Phúc Thọ",
    "Tổ Kỹ thuật Địa bàn Quảng Oai": "Quảng Oai", 
    "Tổ Kỹ thuật Địa bàn Suối hai": "Suối Hai",
    "Tổ Kỹ thuật Địa bàn Sơn Tây": "Sơn Tây"
}

# Màu sắc cho biểu đồ
CHART_COLORS = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6B5B95']
BAR_COLORS = ['#4CAF50', '#2196F3', '#FF9800', '#E91E63', '#9C27B0']  # Xanh lá, Xanh dương, Cam, Hồng, Tím


# =============================================================================
# HÀM TIỆN ÍCH
# =============================================================================
def set_cell_shading(cell, color):
    """Đặt màu nền cho ô trong bảng"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_border(table):
    """Đặt viền cho bảng"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def format_number(value, decimal_places=2):
    """Format số với xử lý NaN"""
    if pd.isna(value):
        return "N/A"
    return f"{value:.{decimal_places}f}"


def format_percent(value, decimal_places=2):
    """Format tỷ lệ phần trăm"""
    if pd.isna(value):
        return "N/A"
    return f"{value:.{decimal_places}f}%"


# =============================================================================
# HÀM ĐỌC DỮ LIỆU
# =============================================================================
def load_kpi_data(kpi_folder=DEFAULT_KPI_FOLDER):
    """
    Đọc dữ liệu KPI từ các file Excel
    
    Returns:
        tuple: (df_summary, df_detail) - DataFrame tóm tắt và chi tiết
    """
    kpi_path = Path(kpi_folder)
    
    # Đọc file tóm tắt
    summary_file = kpi_path / "KPI_NVKT_TomTat.xlsx"
    df_summary = pd.read_excel(summary_file)
    
    # Đọc file chi tiết
    detail_file = kpi_path / "KPI_NVKT_ChiTiet.xlsx"
    df_detail = pd.read_excel(detail_file)
    
    return df_summary, df_detail


def load_c1x_reports(data_folder="downloads/baocao_hanoi"):
    """
    Đọc dữ liệu chi tiết từ các file báo cáo C1.x
    
    Returns:
        dict: Dictionary chứa các DataFrame từ các sheet tổng hợp
    """
    data_path = Path(data_folder)
    reports = {}
    
    # C1.1 Report
    try:
        c11_file = data_path / "c1.1 report.xlsx"
        if c11_file.exists():
            reports['c11'] = pd.read_excel(c11_file, sheet_name='TH_C1.1')
            print("   ✅ Đọc C1.1 report thành công")
    except Exception as e:
        print(f"   ⚠️ Không thể đọc C1.1 report: {e}")
    
    # C1.2 Report
    try:
        c12_file = data_path / "c1.2 report.xlsx"
        if c12_file.exists():
            reports['c12'] = pd.read_excel(c12_file, sheet_name='TH_C1.2')
            print("   ✅ Đọc C1.2 report thành công")
    except Exception as e:
        print(f"   ⚠️ Không thể đọc C1.2 report: {e}")
    
    # C1.3 Report
    try:
        c13_file = data_path / "c1.3 report.xlsx"
        if c13_file.exists():
            reports['c13'] = pd.read_excel(c13_file, sheet_name='TH_C1.3')
            print("   ✅ Đọc C1.3 report thành công")
    except Exception as e:
        print(f"   ⚠️ Không thể đọc C1.3 report: {e}")
    
    # C1.4 Report
    try:
        c14_file = data_path / "c1.4 report.xlsx"
        if c14_file.exists():
            reports['c14'] = pd.read_excel(c14_file, sheet_name='TH_C1.4')
            print("   ✅ Đọc C1.4 report thành công")
    except Exception as e:
        print(f"   ⚠️ Không thể đọc C1.4 report: {e}")
    
    # C1.5 Chi tiết Report - Sheet TH_TTVTST
    try:
        c15_file = data_path / "c1.5_chitiet_report.xlsx"
        if c15_file.exists():
            reports['c15_ttvtst'] = pd.read_excel(c15_file, sheet_name='TH_TTVTST')
            print("   ✅ Đọc C1.5 report (TH_TTVTST) thành công")
    except Exception as e:
        print(f"   ⚠️ Không thể đọc C1.5 report: {e}")
    
    return reports


def add_c1x_overview_table(doc, c1x_reports):
    """
    Thêm bảng tổng quan chi tiết từ các báo cáo C1.x vào document
    
    Args:
        doc: Document Word
        c1x_reports: Dictionary chứa các DataFrame từ load_c1x_reports()
    """
    doc.add_heading('1.3. Số liệu chi tiết các chỉ tiêu BSC theo Đội/TTVT', level=2)
    
    # =========================================================================
    # Bảng C1.1 - Tỷ lệ sửa chữa
    # =========================================================================
    if 'c11' in c1x_reports:
        doc.add_heading('C1.1 - Tỷ lệ sửa chữa phiếu chất lượng & báo hỏng', level=3)
        df = c1x_reports['c11']
        
        headers = ['Đơn vị', 'SC Chủ động (SM1)', 'Đạt (SM2)', 'TL SC CĐ (%)', 
                   'Báo hỏng (SM3)', 'Đạt ĐH (SM4)', 'TL SCBH (%)', 'Điểm BSC']
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        set_table_border(table)
        
        # Header
        for i, cell in enumerate(table.rows[0].cells):
            cell.text = headers[i]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell, '2E7D32')
            run = cell.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
        
        # Dữ liệu
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            cells = table.add_row().cells
            short_name = TEAM_SHORT_NAMES.get(row['Đơn vị'], row['Đơn vị'])
            if short_name == 'Tổng':
                short_name = 'TTVT Sơn Tây'
            data = [
                short_name,
                str(int(row.get('SM1', 0))),
                str(int(row.get('SM2', 0))),
                format_number(row.get('Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu', 0)),
                str(int(row.get('SM3', 0))),
                str(int(row.get('SM4', 0))),
                format_number(row.get('Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', 0)),
                format_number(row.get('Chỉ tiêu BSC', 0))
            ]
            for i, value in enumerate(data):
                cells[i].text = value
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cells[i].paragraphs[0].runs[0]
                run.font.size = Pt(9)
                if idx % 2 == 0:
                    set_cell_shading(cells[i], 'E8F5E9')
                # Tô đậm dòng tổng
                if short_name == 'TTVT Sơn Tây':
                    run.font.bold = True
                    set_cell_shading(cells[i], 'C8E6C9')
        
        doc.add_paragraph()
    
    # =========================================================================
    # Bảng C1.2 - Tỷ lệ báo hỏng lặp lại & sự cố
    # =========================================================================
    if 'c12' in c1x_reports:
        doc.add_heading('C1.2 - Tỷ lệ báo hỏng lặp lại & Tỷ lệ sự cố dịch vụ', level=3)
        df = c1x_reports['c12']
        
        headers = ['Đơn vị', 'HLL (SM1)', 'BH (SM2)', 'TL HLL (%)', 
                   'BH SC (SM3)', 'TB (SM4)', 'TL SC (%)', 'Điểm BSC']
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        set_table_border(table)
        
        for i, cell in enumerate(table.rows[0].cells):
            cell.text = headers[i]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell, '1565C0')
            run = cell.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
        
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            cells = table.add_row().cells
            short_name = TEAM_SHORT_NAMES.get(row['Đơn vị'], row['Đơn vị'])
            if short_name == 'Tổng':
                short_name = 'TTVT Sơn Tây'
            data = [
                short_name,
                str(int(row.get('SM1', 0))),
                str(int(row.get('SM2', 0))),
                format_number(row.get('Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại', 0)),
                str(int(row.get('SM3', 0))),
                str(int(row.get('SM4', 0))),
                format_number(row.get('Tỷ lệ sự cố dịch vụ BRCĐ', 0)),
                format_number(row.get('Chỉ tiêu BSC', 0))
            ]
            for i, value in enumerate(data):
                cells[i].text = value
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cells[i].paragraphs[0].runs[0]
                run.font.size = Pt(9)
                if idx % 2 == 0:
                    set_cell_shading(cells[i], 'E3F2FD')
                if short_name == 'TTVT Sơn Tây':
                    run.font.bold = True
                    set_cell_shading(cells[i], 'BBDEFB')
        
        doc.add_paragraph()
    
    # =========================================================================
    # Bảng C1.3 - Kênh TSL
    # =========================================================================
    if 'c13' in c1x_reports:
        doc.add_heading('C1.3 - Chỉ tiêu kênh thuê leased line (TSL)', level=3)
        df = c1x_reports['c13']
        
        headers = ['Đơn vị', 'SC TSL (SM1)', 'Đạt (SM2)', 'TL SC (%)', 
                   'HLL (SM3)', 'BH (SM4)', 'TL HLL (%)', 'Số TB (SM6)', 'TL SC (%)', 'Điểm BSC']
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        set_table_border(table)
        
        for i, cell in enumerate(table.rows[0].cells):
            cell.text = headers[i]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell, '6A1B9A')
            run = cell.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(8)
        
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            cells = table.add_row().cells
            short_name = TEAM_SHORT_NAMES.get(row['Đơn vị'], row['Đơn vị'])
            if short_name == 'Tổng':
                short_name = 'TTVT Sơn Tây'
            data = [
                short_name,
                str(int(row.get('SM1', 0))),
                str(int(row.get('SM2', 0))),
                format_number(row.get('Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định', 0)),
                str(int(row.get('SM3', 0))),
                str(int(row.get('SM4', 0))),
                format_number(row.get('Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại', 0)),
                str(int(row.get('SM6', 0))),
                format_number(row.get('Tỷ lệ sự cố dịch vụ kênh TSL', 0)),
                format_number(row.get('Chỉ tiêu BSC', 0))
            ]
            for i, value in enumerate(data):
                cells[i].text = value
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cells[i].paragraphs[0].runs[0]
                run.font.size = Pt(8)
                if idx % 2 == 0:
                    set_cell_shading(cells[i], 'F3E5F5')
                if short_name == 'TTVT Sơn Tây':
                    run.font.bold = True
                    set_cell_shading(cells[i], 'E1BEE7')
        
        doc.add_paragraph()
    
    # =========================================================================
    # Bảng C1.4 - Hài lòng khách hàng
    # =========================================================================
    if 'c14' in c1x_reports:
        doc.add_heading('C1.4 - Độ hài lòng khách hàng sau sửa chữa', level=3)
        df = c1x_reports['c14']
        
        headers = ['Đơn vị', 'Tổng phiếu', 'Đã KS', 'KS TC', 'KH HL', 
                   'KHL KT PV', 'TL HL PV (%)', 'TL KH HL (%)', 'Điểm BSC']
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        set_table_border(table)
        
        for i, cell in enumerate(table.rows[0].cells):
            cell.text = headers[i]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell, 'F57C00')
            run = cell.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(8)
        
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            cells = table.add_row().cells
            short_name = TEAM_SHORT_NAMES.get(row['Đơn vị'], row['Đơn vị'])
            if short_name == 'Tổng':
                short_name = 'TTVT Sơn Tây'
            data = [
                short_name,
                str(int(row.get('Tổng phiếu', 0))),
                str(int(row.get('SL đã KS', 0))),
                str(int(row.get('SL KS thành công', 0))),
                str(int(row.get('SL KH hài lòng', 0))),
                str(int(row.get('Không HL KT phục vụ', 0))),
                format_number(row.get('Tỷ lệ HL KT phục vụ', 0)),
                format_number(row.get('Tỷ lệ KH hài lòng', 0)),
                format_number(row.get('Điểm BSC', 0))
            ]
            for i, value in enumerate(data):
                cells[i].text = value
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cells[i].paragraphs[0].runs[0]
                run.font.size = Pt(8)
                if idx % 2 == 0:
                    set_cell_shading(cells[i], 'FFF3E0')
                if short_name == 'TTVT Sơn Tây':
                    run.font.bold = True
                    set_cell_shading(cells[i], 'FFE0B2')
        
        doc.add_paragraph()
    
    # =========================================================================
    # Bảng C1.5 - Tỷ lệ thiết lập dịch vụ đạt
    # =========================================================================
    if 'c15_ttvtst' in c1x_reports:
        doc.add_heading('C1.5 - Tỷ lệ thiết lập dịch vụ đạt thời gian quy định', level=3)
        df = c1x_reports['c15_ttvtst']
        
        headers = ['Đơn vị', 'Phiếu đạt', 'Phiếu không đạt', 'Tổng HC', 'Tỉ lệ đạt (%)']
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        set_table_border(table)
        
        for i, cell in enumerate(table.rows[0].cells):
            cell.text = headers[i]
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell, '00796B')
            run = cell.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(10)
        
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            cells = table.add_row().cells
            don_vi = row.get('DOIVT', '')
            short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
            data = [
                short_name,
                str(int(row.get('Phiếu đạt', 0))),
                str(int(row.get('Phiếu không đạt', 0))),
                str(int(row.get('Tổng Hoàn công', 0))),
                format_number(row.get('Tỉ lệ đạt (%)', 0))
            ]
            for i, value in enumerate(data):
                cells[i].text = value
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cells[i].paragraphs[0].runs[0]
                run.font.size = Pt(10)
                if idx % 2 == 0:
                    set_cell_shading(cells[i], 'E0F2F1')
                # Tô đậm dòng TTVT Sơn Tây 
                if 'TTVT' in don_vi:
                    run.font.bold = True
                    set_cell_shading(cells[i], 'B2DFDB')
        
        doc.add_paragraph()


# =============================================================================
# HÀM TẠO BIỂU ĐỒ
# =============================================================================
def create_team_comparison_chart(c1x_reports, output_path=None):
    """
    Tạo biểu đồ so sánh điểm BSC thực tế giữa 4 tổ
    Lấy điểm BSC từ các báo cáo C1.x (không tính dòng Tổng)
    
    Args:
        c1x_reports: Dictionary chứa các DataFrame từ load_c1x_reports()
        output_path: Đường dẫn lưu file ảnh (None = trả về bytes)
    
    Returns:
        bytes hoặc str: Dữ liệu ảnh hoặc đường dẫn file
    """
    # Chuẩn bị dữ liệu từ các báo cáo C1.x
    teams_order = ['Phúc Thọ', 'Quảng Oai', 'Suối Hai', 'Sơn Tây']
    
    # Khởi tạo dict chứa điểm BSC (bao gồm C1.5)
    bsc_scores = {team: {'C1.1': 0, 'C1.2': 0, 'C1.3': 0, 'C1.4': 0, 'C1.5': 0} for team in teams_order}
    
    # Lấy điểm BSC từ C1.1
    if 'c11' in c1x_reports:
        df = c1x_reports['c11']
        for _, row in df.iterrows():
            don_vi = row.get('Đơn vị', '')
            short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
            if short_name in teams_order:
                bsc_scores[short_name]['C1.1'] = row.get('Chỉ tiêu BSC', 0)
    
    # Lấy điểm BSC từ C1.2
    if 'c12' in c1x_reports:
        df = c1x_reports['c12']
        for _, row in df.iterrows():
            don_vi = row.get('Đơn vị', '')
            short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
            if short_name in teams_order:
                bsc_scores[short_name]['C1.2'] = row.get('Chỉ tiêu BSC', 0)
    
    # Lấy điểm BSC từ C1.3
    if 'c13' in c1x_reports:
        df = c1x_reports['c13']
        for _, row in df.iterrows():
            don_vi = row.get('Đơn vị', '')
            short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
            if short_name in teams_order:
                bsc_scores[short_name]['C1.3'] = row.get('Chỉ tiêu BSC', 0)
    
    # Lấy điểm BSC từ C1.4
    if 'c14' in c1x_reports:
        df = c1x_reports['c14']
        for _, row in df.iterrows():
            don_vi = row.get('Đơn vị', '')
            short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
            if short_name in teams_order:
                bsc_scores[short_name]['C1.4'] = row.get('Điểm BSC', 0)
    
    # Lấy điểm BSC từ C1.5 (tính từ tỷ lệ đạt)
    if 'c15_ttvtst' in c1x_reports:
        df = c1x_reports['c15_ttvtst']
        for _, row in df.iterrows():
            don_vi = row.get('DOIVT', '')
            short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
            if short_name in teams_order:
                ty_le = row.get('Tỉ lệ đạt (%)', 0)
                if ty_le >= 99.5:
                    diem_bsc = 5.0
                elif ty_le <= 89.5:
                    diem_bsc = 1.0
                else:
                    diem_bsc = 1 + 4 * (ty_le - 89.5) / 10
                bsc_scores[short_name]['C1.5'] = round(diem_bsc, 2)
    
    # Tạo DataFrame từ dữ liệu
    chart_data = pd.DataFrame(bsc_scores).T
    chart_data = chart_data.reindex(teams_order)  # Đảm bảo thứ tự
    
    # Tạo biểu đồ
    fig, ax = plt.subplots(figsize=(14, 6))
    
    x = np.arange(len(teams_order))
    width = 0.15  # Thu hẹp để có chỗ cho 5 cột
    
    metrics = ['C1.1', 'C1.2', 'C1.3', 'C1.4', 'C1.5']
    
    for i, metric in enumerate(metrics):
        values = chart_data[metric].fillna(0).values
        bars = ax.bar(x + i*width, values, width, label=metric, color=BAR_COLORS[i])
        # Thêm giá trị lên cột
        for bar, val in zip(bars, values):
            if val > 0:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.05,
                       f'{val:.2f}', ha='center', va='bottom', fontsize=8)
    
    ax.set_xlabel('Tổ Kỹ thuật', fontsize=12)
    ax.set_ylabel('Điểm BSC', fontsize=12)
    ax.set_title('SO SÁNH ĐIỂM BSC THỰC TẾ GIỮA CÁC TỔ', fontsize=14, fontweight='bold')
    ax.set_xticks(x + width * 2)  # Điều chỉnh vị trí label
    ax.set_xticklabels(teams_order, fontsize=11)
    ax.set_ylim(0, 6)
    ax.legend(loc='upper right')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    
    plt.tight_layout()
    
    # Lưu hoặc trả về bytes
    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        return output_path
    else:
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf


def create_nvkt_bar_chart(df_summary, team_name, output_path=None):
    """
    Tạo biểu đồ cột so sánh điểm KPI theo NVKT trong 1 tổ
    
    Args:
        df_summary: DataFrame chứa dữ liệu KPI
        team_name: Tên tổ cần tạo biểu đồ
        output_path: Đường dẫn lưu file (None = trả về bytes)
    
    Returns:
        bytes hoặc str
    """
    # Lọc dữ liệu theo tổ
    df_team = df_summary[df_summary['don_vi'] == team_name].copy()
    
    if df_team.empty:
        return None
    
    # Sắp xếp theo tên NVKT
    df_team = df_team.sort_values('nvkt')
    
    # Tạo biểu đồ
    fig, ax = plt.subplots(figsize=(14, 6))
    
    x = np.arange(len(df_team))
    width = 0.2
    
    metrics = ['Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4', 'Diem_C1.5']
    labels = ['C1.1', 'C1.2', 'C1.4', 'C1.5']
    
    for i, (metric, label) in enumerate(zip(metrics, labels)):
        values = df_team[metric].fillna(0).values
        bars = ax.bar(x + i*width, values, width, label=label, color=BAR_COLORS[i])
        # Thêm giá trị lên cột
        for bar, val in zip(bars, values):
            if val > 0:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.03,
                       f'{val:.1f}', ha='center', va='bottom', fontsize=7, rotation=90)
    
    short_name = TEAM_SHORT_NAMES.get(team_name, team_name)
    ax.set_xlabel('NVKT', fontsize=11)
    ax.set_ylabel('Điểm KPI', fontsize=11)
    ax.set_title(f'ĐIỂM KPI THEO NVKT - {short_name.upper()}', fontsize=13, fontweight='bold')
    ax.set_xticks(x + width * 1.5)
    ax.set_xticklabels(df_team['nvkt'].values, fontsize=8, rotation=45, ha='right')
    ax.set_ylim(0, 6)
    ax.legend(loc='upper right')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    
    plt.tight_layout()
    
    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        return output_path
    else:
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf


# =============================================================================
# HÀM TẠO BẢNG TRONG WORD
# =============================================================================
def add_kpi_summary_table(doc, df_summary, team_name=None):
    """
    Thêm bảng tổng hợp KPI vào document
    
    Args:
        doc: Document Word
        df_summary: DataFrame dữ liệu
        team_name: Lọc theo tổ (None = tất cả)
    """
    if team_name:
        df = df_summary[df_summary['don_vi'] == team_name].copy()
    else:
        df = df_summary.copy()
    
    # Sắp xếp
    df = df.sort_values(['don_vi', 'nvkt'])
    
    # Tạo bảng
    headers = ['STT', 'Đơn vị', 'NVKT', 'C1.1', 'C1.2', 'C1.4', 'C1.5']
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        header_cells[i].text = header
        header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(header_cells[i], '1F4E79')
        run = header_cells[i].paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    # Dữ liệu
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        cells = table.add_row().cells
        
        short_name = TEAM_SHORT_NAMES.get(row['don_vi'], row['don_vi'])
        
        data = [
            str(idx),
            short_name,
            row['nvkt'],
            format_number(row.get('Diem_C1.1', np.nan)),
            format_number(row.get('Diem_C1.2', np.nan)),
            format_number(row.get('Diem_C1.4', np.nan)),
            format_number(row.get('Diem_C1.5', np.nan))
        ]
        
        for i, value in enumerate(data):
            cells[i].text = value
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[i].paragraphs[0].runs[0]
            run.font.size = Pt(9)
            
            # Tô màu dòng xen kẽ
            if idx % 2 == 0:
                set_cell_shading(cells[i], 'E8F4FD')


def add_c11_detail_table(doc, df_detail, team_name=None):
    """
    Thêm bảng chi tiết C1.1 (Thành phần 1 + Thành phần 2)
    """
    if team_name:
        df = df_detail[df_detail['don_vi'] == team_name].copy()
    else:
        df = df_detail.copy()
    
    df = df.sort_values(['don_vi', 'nvkt'])
    
    # Tiêu đề
    doc.add_heading('Chi tiết chỉ tiêu C1.1 - Chất lượng sửa chữa thuê bao BRCĐ', level=3)
    
    # Chú thích
    p = doc.add_paragraph()
    p.add_run('📋 Chú thích: ').bold = True
    p.add_run('TP1 = Sửa chữa chủ động (SCCD ≤72h) | TP2 = Sửa chữa theo báo hỏng (SC BH)')
    
    headers = ['STT', 'NVKT', 'Tổng SCCD', 'Đạt ≤72h', 'TL(%)', 'Điểm TP1',
               'Tổng SC BH', 'Đúng hạn', 'TL(%)', 'Điểm TP2', 'Điểm C1.1']
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    for i, header in enumerate(table.rows[0].cells):
        header.text = headers[i]
        header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(header, '2E7D32')
        run = header.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(8)
    
    # Dữ liệu
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        cells = table.add_row().cells
        data = [
            str(idx),
            row['nvkt'],
            format_number(row.get('c11_tp1_tong_phieu', np.nan), 0),
            format_number(row.get('c11_tp1_phieu_dat', np.nan), 0),
            format_number(row.get('c11_tp1_ty_le', np.nan)),
            format_number(row.get('diem_c11_tp1', np.nan)),
            format_number(row.get('c11_tp2_tong_phieu', np.nan), 0),
            format_number(row.get('c11_tp2_phieu_dat', np.nan), 0),
            format_number(row.get('c11_tp2_ty_le', np.nan)),
            format_number(row.get('diem_c11_tp2', np.nan)),
            format_number(row.get('Diem_C1.1', np.nan))
        ]
        
        for i, value in enumerate(data):
            cells[i].text = value
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[i].paragraphs[0].runs[0]
            run.font.size = Pt(8)
            
            if idx % 2 == 0:
                set_cell_shading(cells[i], 'E8F5E9')


def add_c12_detail_table(doc, df_detail, team_name=None):
    """
    Thêm bảng chi tiết C1.2 (Thành phần 1 + Thành phần 2)
    """
    if team_name:
        df = df_detail[df_detail['don_vi'] == team_name].copy()
    else:
        df = df_detail.copy()
    
    df = df.sort_values(['don_vi', 'nvkt'])
    
    doc.add_heading('Chi tiết chỉ tiêu C1.2 - Tỷ lệ thuê bao báo hỏng', level=3)
    
    # Chú thích
    p = doc.add_paragraph()
    p.add_run('📋 Chú thích: ').bold = True
    p.add_run('TP1 = Hỏng lặp (≥2 lần/7 ngày) | TP2 = Tỷ lệ BH/TB quản lý | BH = Báo hỏng | TB QL = Thuê bao quản lý')
    
    headers = ['STT', 'NVKT', 'Hỏng lặp', 'Tổng BH', 'TL(%)', 'Điểm TP1',
               'Phiếu BH', 'TB QL', 'TL(‰)', 'Điểm TP2', 'Điểm C1.2']
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    for i, header in enumerate(table.rows[0].cells):
        header.text = headers[i]
        header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(header, '1565C0')
        run = header.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(8)
    
    # Dữ liệu
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        cells = table.add_row().cells
        data = [
            str(idx),
            row['nvkt'],
            format_number(row.get('c12_tp1_phieu_hll', np.nan), 0),
            format_number(row.get('c12_tp1_phieu_bh', np.nan), 0),
            format_number(row.get('c12_tp1_ty_le', np.nan)),
            format_number(row.get('diem_c12_tp1', np.nan)),
            format_number(row.get('c12_tp2_phieu_bh', np.nan), 0),
            format_number(row.get('c12_tp2_tong_tb', np.nan), 0),
            format_number(row.get('c12_tp2_ty_le', np.nan)),
            format_number(row.get('diem_c12_tp2', np.nan)),
            format_number(row.get('Diem_C1.2', np.nan))
        ]
        
        for i, value in enumerate(data):
            cells[i].text = value
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[i].paragraphs[0].runs[0]
            run.font.size = Pt(8)
            
            if idx % 2 == 0:
                set_cell_shading(cells[i], 'E3F2FD')


def add_c14_detail_table(doc, df_detail, team_name=None):
    """
    Thêm bảng chi tiết C1.4 - Độ hài lòng khách hàng
    """
    if team_name:
        df = df_detail[df_detail['don_vi'] == team_name].copy()
    else:
        df = df_detail.copy()
    
    df = df.sort_values(['don_vi', 'nvkt'])
    
    doc.add_heading('Chi tiết chỉ tiêu C1.4 - Độ hài lòng khách hàng', level=3)
    
    # Chú thích
    p = doc.add_paragraph()
    p.add_run('📋 Chú thích: ').bold = True
    p.add_run('KS = Khảo sát | Không HL = Không hài lòng | HL = Hài lòng')
    
    headers = ['STT', 'NVKT', 'Tổng KS', 'Không HL', 'Tỷ lệ HL (%)', 'Điểm C1.4']
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    for i, header in enumerate(table.rows[0].cells):
        header.text = headers[i]
        header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(header, 'F57C00')
        run = header.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)
    
    # Dữ liệu
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        cells = table.add_row().cells
        data = [
            str(idx),
            row['nvkt'],
            format_number(row.get('c14_phieu_ks', np.nan), 0),
            format_number(row.get('c14_phieu_khl', np.nan), 0),
            format_number(row.get('c14_ty_le', np.nan)),
            format_number(row.get('Diem_C1.4', np.nan))
        ]
        
        for i, value in enumerate(data):
            cells[i].text = value
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[i].paragraphs[0].runs[0]
            run.font.size = Pt(9)
            
            if idx % 2 == 0:
                set_cell_shading(cells[i], 'FFF3E0')


def add_c15_detail_table(doc, df_detail, team_name=None):
    """
    Thêm bảng chi tiết C1.5 - Tỷ lệ thiết lập dịch vụ
    """
    if team_name:
        df = df_detail[df_detail['don_vi'] == team_name].copy()
    else:
        df = df_detail.copy()
    
    df = df.sort_values(['don_vi', 'nvkt'])
    
    doc.add_heading('Chi tiết chỉ tiêu C1.5 - Thiết lập dịch vụ BRCĐ đạt thời gian quy định', level=3)
    
    # Chú thích
    p = doc.add_paragraph()
    p.add_run('📋 Chú thích: ').bold = True
    p.add_run('Đạt TG = Hoàn thành đúng thời gian (ngoài CCCO: ≤24h, trong CCCO: theo quy định)')
    
    headers = ['STT', 'NVKT', 'Đạt TG', 'Không đạt', 'Tổng phiếu', 'Tỷ lệ (%)', 'Điểm C1.5']
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    for i, header in enumerate(table.rows[0].cells):
        header.text = headers[i]
        header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(header, '7B1FA2')
        run = header.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)
    
    # Dữ liệu
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        cells = table.add_row().cells
        data = [
            str(idx),
            row['nvkt'],
            format_number(row.get('c15_phieu_dat', np.nan), 0),
            format_number(row.get('c15_phieu_khong_dat', np.nan), 0),
            format_number(row.get('c15_tong_phieu', np.nan), 0),
            format_number(row.get('c15_ty_le', np.nan)),
            format_number(row.get('Diem_C1.5', np.nan))
        ]
        
        for i, value in enumerate(data):
            cells[i].text = value
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[i].paragraphs[0].runs[0]
            run.font.size = Pt(9)
            
            if idx % 2 == 0:
                set_cell_shading(cells[i], 'F3E5F5')


# =============================================================================
# HÀM SHC CHO BÁO CÁO TỔNG HỢP
# =============================================================================

def load_shc_summary_by_unit(data_folder="downloads/baocao_hanoi"):
    """
    Đọc dữ liệu SHC tổng hợp theo đơn vị từ sheet Xu_huong_theo_don_vi
    
    Returns:
        dict: {'units': {...}, 'dates': [...]} hoặc None
    """
    import glob
    
    pattern = os.path.join(data_folder, "Bao_cao_xu_huong_SHC_*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        return None
    
    latest_file = max(files, key=os.path.getmtime)
    
    try:
        df = pd.read_excel(latest_file, sheet_name='Xu_huong_theo_don_vi')
        
        # Cột đầu tiên là Đơn vị, các cột còn lại là ngày
        date_columns = [col for col in df.columns if col != 'Đơn vị']
        
        result = {'units': {}, 'dates': date_columns}
        
        for _, row in df.iterrows():
            unit_name = row['Đơn vị']
            values = [int(row[col]) if pd.notna(row[col]) else 0 for col in date_columns]
            result['units'][unit_name] = values
        
        return result
    except Exception as e:
        print(f"   ⚠️ Lỗi đọc SHC summary: {e}")
        return None


def load_shc_by_nvkt_for_unit(unit_name, data_folder="downloads/baocao_hanoi"):
    """
    Đọc dữ liệu SHC theo từng NVKT cho 1 đơn vị từ sheet Xu_huong_theo_NVKT
    
    Returns:
        dict: {'nvkt_list': [...], 'dates': [...], 'data': {nvkt: [values]}} hoặc None
    """
    import glob
    
    pattern = os.path.join(data_folder, "Bao_cao_xu_huong_SHC_*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        return None
    
    latest_file = max(files, key=os.path.getmtime)
    
    try:
        df = pd.read_excel(latest_file, sheet_name='Xu_huong_theo_NVKT')
        
        # Lọc theo đơn vị
        df_unit = df[df['Đơn vị'] == unit_name]
        
        if df_unit.empty:
            return None
        
        date_columns = [col for col in df.columns if col not in ['Đơn vị', 'NVKT']]
        
        result = {
            'nvkt_list': [],
            'dates': [str(d) for d in date_columns],
            'data': {}
        }
        
        for _, row in df_unit.iterrows():
            nvkt = row['NVKT']
            values = [int(row[col]) if pd.notna(row[col]) else 0 for col in date_columns]
            result['nvkt_list'].append(nvkt)
            result['data'][nvkt] = values
        
        return result
    except Exception as e:
        print(f"   ⚠️ Lỗi đọc SHC by NVKT: {e}")
        return None


def create_nvkt_shc_grouped_chart(nvkt_data, unit_name, output_path=None):
    """
    Tạo biểu đồ nhóm cột SHC theo NVKT, mỗi ngày 1 màu khác nhau
    """
    if not nvkt_data or not nvkt_data['data']:
        return None
    
    nvkt_list = nvkt_data['nvkt_list']
    dates = nvkt_data['dates']
    data = nvkt_data['data']
    
    # Sử dụng họ tên đầy đủ
    nvkt_labels = nvkt_list
    
    # Setup figure
    fig, ax = plt.subplots(figsize=(14, 6))
    
    x = np.arange(len(nvkt_list))
    n_dates = len(dates)
    width = 0.8 / n_dates  # Chiều rộng mỗi cột
    
    # Màu sắc cho từng ngày
    colors = plt.cm.tab10(np.linspace(0, 1, n_dates))
    
    # Vẽ từng ngày
    for i, date in enumerate(dates):
        values = [data[nvkt][i] for nvkt in nvkt_list]
        offset = (i - n_dates/2 + 0.5) * width
        bars = ax.bar(x + offset, values, width, label=date, color=colors[i])
        
        # Thêm giá trị lên cột (chỉ nếu > 0)
        for bar, val in zip(bars, values):
            if val > 0:
                ax.annotate(f'{val}',
                           xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                           xytext=(0, 1),
                           textcoords="offset points",
                           ha='center', va='bottom',
                           fontsize=7, fontweight='bold')
    
    short_name = TEAM_SHORT_NAMES.get(unit_name, unit_name)
    ax.set_xlabel('NVKT', fontsize=11)
    ax.set_ylabel('Số TB suy hao cao', fontsize=11)
    ax.set_title(f'KẾT QUẢ XỬ LÝ SUY HAO CAO - {short_name}', fontsize=14, fontweight='bold', pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(nvkt_labels, rotation=45, ha='right', fontsize=9)
    ax.legend(title='Ngày', loc='upper right', fontsize=8, ncol=2)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.set_axisbelow(True)
    
    plt.tight_layout()
    
    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        return output_path
    else:
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf


def create_shc_overview_chart(shc_data, output_path=None):
    """
    Tạo biểu đồ tổng hợp SHC theo ngày cho tất cả đơn vị (stacked bar)
    """
    if not shc_data:
        return None
    
    dates = [str(d) for d in shc_data['dates']]
    units = shc_data['units']
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    x = np.arange(len(dates))
    width = 0.2
    colors = ['#2196F3', '#4CAF50', '#FF9800', '#E91E63']
    
    unit_names = list(units.keys())
    for i, unit_name in enumerate(unit_names):
        short_name = TEAM_SHORT_NAMES.get(unit_name, unit_name)
        values = units[unit_name]
        bars = ax.bar(x + i * width, values, width, label=short_name, color=colors[i % len(colors)])
        
        # Thêm giá trị lên cột
        for bar, val in zip(bars, values):
            if val > 0:
                ax.annotate(f'{val}', xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                           xytext=(0, 2), textcoords="offset points",
                           ha='center', va='bottom', fontsize=8)
    
    ax.set_xlabel('Ngày', fontsize=11)
    ax.set_ylabel('Số TB suy hao cao', fontsize=11)
    ax.set_title('XU HƯỚNG SUY HAO CAO THEO ĐƠN VỊ', fontsize=14, fontweight='bold', pad=15)
    ax.set_xticks(x + width * (len(unit_names) - 1) / 2)
    ax.set_xticklabels(dates, rotation=45, ha='right', fontsize=9)
    ax.legend(loc='upper right')
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    
    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        return output_path
    else:
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf


def add_shc_overview_section(doc, data_folder="downloads/baocao_hanoi"):
    """
    Thêm phần tổng quan SHC vào báo cáo (PHẦN 1)
    """
    shc_data = load_shc_summary_by_unit(data_folder)
    
    if not shc_data:
        return
    
    doc.add_heading('1.4. Tổng quan Suy Hao Cao', level=2)
    
    dates = shc_data['dates']
    units = shc_data['units']
    
    # Bảng tổng hợp
    table = doc.add_table(rows=1, cols=len(dates) + 2)
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    headers = ['Đơn vị'] + [str(d) for d in dates] + ['Tổng']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '1565C0')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)
    
    # Data rows
    grand_total = 0
    for unit_name, values in units.items():
        cells = table.add_row().cells
        short_name = TEAM_SHORT_NAMES.get(unit_name, unit_name)
        cells[0].text = short_name
        
        total = sum(values)
        grand_total += total
        
        for j, val in enumerate(values):
            cells[j + 1].text = str(val)
            cells[j + 1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[j + 1].paragraphs[0].runs[0]
            run.font.size = Pt(9)
            if val > 5:
                run.font.bold = True
                run.font.color.rgb = RGBColor(200, 0, 0)
        
        cells[-1].text = str(total)
        cells[-1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cells[-1].paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(9)
    
    # Dòng tổng
    cells = table.add_row().cells
    cells[0].text = 'TỔNG CỘNG'
    cells[0].paragraphs[0].runs[0].font.bold = True
    set_cell_shading(cells[0], 'E3F2FD')
    
    totals_by_date = [sum(units[u][i] for u in units) for i in range(len(dates))]
    for j, total in enumerate(totals_by_date):
        cells[j + 1].text = str(total)
        cells[j + 1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cells[j + 1].paragraphs[0].runs[0].font.bold = True
        set_cell_shading(cells[j + 1], 'E3F2FD')
    
    cells[-1].text = str(grand_total)
    cells[-1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cells[-1].paragraphs[0].runs[0].font.bold = True
    set_cell_shading(cells[-1], 'E3F2FD')
    
    doc.add_paragraph()
    
    # Biểu đồ
    try:
        chart = create_shc_overview_chart(shc_data)
        if chart:
            doc.add_picture(chart, width=Inches(6.5))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception as e:
        print(f"   ⚠️ Không thể tạo biểu đồ SHC: {e}")
    
    doc.add_paragraph()


def add_shc_unit_section(doc, unit_name, data_folder="downloads/baocao_hanoi"):
    """
    Thêm phần SHC chi tiết cho 1 đơn vị (trong PHẦN 2)
    """
    shc_data = load_shc_summary_by_unit(data_folder)
    
    if not shc_data or unit_name not in shc_data['units']:
        return
    
    dates = shc_data['dates']
    values = shc_data['units'][unit_name]
    
    short_name = TEAM_SHORT_NAMES.get(unit_name, unit_name)
    doc.add_heading(f'Số liệu Suy Hao Cao - {short_name}', level=3)
    
    # Bảng dữ liệu - cải thiện format
    table = doc.add_table(rows=2, cols=len(dates) + 1)
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header row - bao gồm cột Chỉ tiêu
    headers = ['Ngày'] + [str(d) for d in dates]
    for j, header in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = header
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '1565C0')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)
    
    # Data row
    table.rows[1].cells[0].text = 'Số TB SHC'
    table.rows[1].cells[0].paragraphs[0].runs[0].font.bold = True
    table.rows[1].cells[0].paragraphs[0].runs[0].font.size = Pt(9)
    set_cell_shading(table.rows[1].cells[0], 'E3F2FD')
    
    for j, val in enumerate(values):
        cell = table.rows[1].cells[j + 1]
        cell.text = str(val)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
        run.font.bold = True
        
        # Tô màu theo mức độ (dựa trên ngưỡng phù hợp cho đơn vị)
        avg_val = sum(values) / len(values) if values else 0
        if val <= avg_val * 0.5:
            set_cell_shading(cell, 'C8E6C9')  # Xanh nhạt - tốt
            run.font.color.rgb = RGBColor(0, 128, 0)
        elif val <= avg_val * 1.2:
            set_cell_shading(cell, 'FFF9C4')  # Vàng nhạt - trung bình
        else:
            set_cell_shading(cell, 'FFCDD2')  # Đỏ nhạt - cao
            run.font.color.rgb = RGBColor(200, 0, 0)
    
    doc.add_paragraph()
    
    # Thống kê chi tiết
    total = sum(values)
    avg = total / len(values) if values else 0
    max_val = max(values) if values else 0
    min_val = min(values) if values else 0
    max_day = dates[values.index(max_val)] if values else 'N/A'
    min_day = dates[values.index(min_val)] if values else 'N/A'
    
    p = doc.add_paragraph()
    p.add_run(f'📊 Tổng: {total} | Trung bình: {avg:.1f}/ngày | ').bold = True
    p.add_run(f'Cao nhất: {max_val} ({max_day}) | Thấp nhất: {min_val} ({min_day})')
    
    doc.add_paragraph()
    
    # Biểu đồ nhóm theo NVKT (như hình mẫu)
    try:
        nvkt_data = load_shc_by_nvkt_for_unit(unit_name, data_folder)
        if nvkt_data:
            chart = create_nvkt_shc_grouped_chart(nvkt_data, unit_name)
            if chart:
                doc.add_picture(chart, width=Inches(6.5))
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception as e:
        print(f"   ⚠️ Không thể tạo biểu đồ SHC cho {short_name}: {e}")
    
    doc.add_paragraph()


# =============================================================================
# HÀM TẠO BÁO CÁO WORD HOÀN CHỈNH
# =============================================================================

def generate_kpi_report(kpi_folder=DEFAULT_KPI_FOLDER, output_folder=DEFAULT_OUTPUT_FOLDER, 
                         report_month=None, report_title=None):
    """
    Tạo báo cáo Word hoàn chỉnh với bảng biểu và biểu đồ KPI
    
    Args:
        kpi_folder: Thư mục chứa file KPI Excel
        output_folder: Thư mục xuất báo cáo Word
        report_month: Tháng báo cáo (vd: "01/2026"), mặc định là tháng hiện tại
        report_title: Tiêu đề tùy chỉnh
        
    Returns:
        str: Đường dẫn file Word đã tạo
    """
    print("="*60)
    print("📝 BẮT ĐẦU TẠO BÁO CÁO WORD KPI")
    print("="*60)
    
    # Xác định tháng báo cáo
    if report_month is None:
        report_month = datetime.now().strftime("%m/%Y")
    
    # Tạo thư mục output nếu chưa có
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Đọc dữ liệu
    print("📊 Đọc dữ liệu KPI...")
    df_summary, df_detail = load_kpi_data(kpi_folder)
    
    # Lấy danh sách các tổ
    teams = df_summary['don_vi'].unique()
    print(f"   Tìm thấy {len(teams)} tổ kỹ thuật")
    
    # Tạo document
    doc = Document()
    
    # Thiết lập style
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # =========================================================================
    # TIÊU ĐỀ
    # =========================================================================
    created_time = datetime.now().strftime('%d/%m/%Y %H:%M')
    title = doc.add_heading(level=0)
    title_run = title.add_run(report_title or f'BÁO CÁO KẾT QUẢ BSC/KPI\nTHÁNG {report_month}')
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Ngày tạo: {created_time}")
    doc.add_paragraph()
    
    # =========================================================================
    # PHẦN 1: TỔNG QUAN
    # =========================================================================
    print("📈 Tạo phần Tổng quan...")
    doc.add_heading('PHẦN 1: TỔNG QUAN', level=1)
    
    # Đọc dữ liệu chi tiết từ các báo cáo C1.x (cần cho biểu đồ BSC)
    print("📊 Đọc dữ liệu chi tiết từ các báo cáo C1.x...")
    c1x_reports = load_c1x_reports()
    
    # 1.1 Biểu đồ so sánh điểm BSC thực tế 4 tổ
    doc.add_heading('1.1. So sánh điểm BSC thực tế giữa các tổ', level=2)
    if c1x_reports:
        team_chart = create_team_comparison_chart(c1x_reports)
        doc.add_picture(team_chart, width=Inches(6.5))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph("(Không có dữ liệu C1.x để tạo biểu đồ)")
    doc.add_paragraph()
    
    # 1.2 Thêm bảng thống kê tổng hợp theo tổ - sử dụng điểm BSC thực tế
    doc.add_heading('1.2. Thống kê điểm BSC theo đơn vị', level=2)
    
    # Chú thích giải thích các chỉ tiêu
    legend = doc.add_paragraph()
    legend.add_run('📋 CHÚ THÍCH CÁC CHỈ TIÊU BSC - VIỄN CẢNH KHÁCH HÀNG (C)').bold = True
    
    # C1.1
    p11 = doc.add_paragraph()
    p11.add_run('C1.1 - Chất lượng sửa chữa thuê bao BRCĐ: ').bold = True
    p11.add_run('Gồm 2 thành phần:\n')
    p11.add_run('   • TP1 (30%): Sửa chữa chủ động - Tỷ lệ phiếu SCCD hoàn thành ≤72h\n')
    p11.add_run('   • TP2 (70%): Sửa chữa theo báo hỏng - Tỷ lệ phiếu BH hoàn thành đúng hạn')
    
    # C1.2
    p12 = doc.add_paragraph()
    p12.add_run('C1.2 - Tỷ lệ thuê bao báo hỏng: ').bold = True
    p12.add_run('Gồm 2 thành phần:\n')
    p12.add_run('   • TP1 (50%): Hỏng lặp lại - Tỷ lệ TB báo hỏng ≥2 lần/7 ngày\n')
    p12.add_run('   • TP2 (50%): Tỷ lệ sự cố - Tỷ lệ phiếu BH / Tổng TB quản lý (‰)')
    
    # C1.3
    p13 = doc.add_paragraph()
    p13.add_run('C1.3 - Chất lượng sửa chữa kênh TSL (Leased Line): ').bold = True
    p13.add_run('Áp dụng cho các dịch vụ Internet trực tiếp, kênh thuê riêng, MegaWan, Metronet, Siptrunking')
    
    # C1.4
    p14 = doc.add_paragraph()
    p14.add_run('C1.4 - Độ hài lòng khách hàng: ').bold = True
    p14.add_run('Tỷ lệ khách hàng hài lòng sau khi được sửa chữa (qua khảo sát)')
    
    # C1.5
    p15 = doc.add_paragraph()
    p15.add_run('C1.5 - Thiết lập dịch vụ BRCĐ đạt thời gian quy định: ').bold = True
    p15.add_run('Tỷ lệ phiếu lắp đặt hoàn thành đúng hạn\n')
    p15.add_run('   • Ngoài CCCO: ≤24h | Trong CCCO: Phiếu trước 17h xong trong ngày')
    
    doc.add_paragraph()  # Spacing
    
    # Chuẩn bị dữ liệu BSC từ các báo cáo C1.x
    teams_order = ['Phúc Thọ', 'Quảng Oai', 'Suối Hai', 'Sơn Tây']
    bsc_data = {team: {'C1.1': 0, 'C1.2': 0, 'C1.3': 0, 'C1.4': 0, 'C1.5': 0} for team in teams_order}
    
    # Lấy điểm BSC từ các báo cáo
    if c1x_reports:
        if 'c11' in c1x_reports:
            for _, row in c1x_reports['c11'].iterrows():
                short_name = TEAM_SHORT_NAMES.get(row.get('Đơn vị', ''), row.get('Đơn vị', ''))
                if short_name in teams_order:
                    bsc_data[short_name]['C1.1'] = row.get('Chỉ tiêu BSC', 0)
        
        if 'c12' in c1x_reports:
            for _, row in c1x_reports['c12'].iterrows():
                short_name = TEAM_SHORT_NAMES.get(row.get('Đơn vị', ''), row.get('Đơn vị', ''))
                if short_name in teams_order:
                    bsc_data[short_name]['C1.2'] = row.get('Chỉ tiêu BSC', 0)
        
        if 'c13' in c1x_reports:
            for _, row in c1x_reports['c13'].iterrows():
                short_name = TEAM_SHORT_NAMES.get(row.get('Đơn vị', ''), row.get('Đơn vị', ''))
                if short_name in teams_order:
                    bsc_data[short_name]['C1.3'] = row.get('Chỉ tiêu BSC', 0)
        
        if 'c14' in c1x_reports:
            for _, row in c1x_reports['c14'].iterrows():
                short_name = TEAM_SHORT_NAMES.get(row.get('Đơn vị', ''), row.get('Đơn vị', ''))
                if short_name in teams_order:
                    bsc_data[short_name]['C1.4'] = row.get('Điểm BSC', 0)
        
        # Lấy điểm C1.5 từ TH_TTVTST (tính từ tỷ lệ đạt)
        if 'c15_ttvtst' in c1x_reports:
            for _, row in c1x_reports['c15_ttvtst'].iterrows():
                don_vi = row.get('DOIVT', '')
                short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
                if short_name in teams_order:
                    # Tính điểm BSC từ tỷ lệ đạt: >= 99.5% = 5, <= 89.5% = 1, giữa = nội suy
                    ty_le = row.get('Tỉ lệ đạt (%)', 0)
                    if ty_le >= 99.5:
                        diem_bsc = 5.0
                    elif ty_le <= 89.5:
                        diem_bsc = 1.0
                    else:
                        diem_bsc = 1 + 4 * (ty_le - 89.5) / 10
                    bsc_data[short_name]['C1.5'] = round(diem_bsc, 2)
    
    # Lấy điểm BSC tổng hợp cho TTVT Sơn Tây (dòng "Tổng" trong các báo cáo)
    ttvt_scores = {'C1.1': 0, 'C1.2': 0, 'C1.3': 0, 'C1.4': 0, 'C1.5': 0}
    if c1x_reports:
        if 'c11' in c1x_reports:
            for _, row in c1x_reports['c11'].iterrows():
                if row.get('Đơn vị', '') == 'Tổng':
                    ttvt_scores['C1.1'] = row.get('Chỉ tiêu BSC', 0)
        if 'c12' in c1x_reports:
            for _, row in c1x_reports['c12'].iterrows():
                if row.get('Đơn vị', '') == 'Tổng':
                    ttvt_scores['C1.2'] = row.get('Chỉ tiêu BSC', 0)
        if 'c13' in c1x_reports:
            for _, row in c1x_reports['c13'].iterrows():
                if row.get('Đơn vị', '') == 'Tổng':
                    ttvt_scores['C1.3'] = row.get('Chỉ tiêu BSC', 0)
        if 'c14' in c1x_reports:
            for _, row in c1x_reports['c14'].iterrows():
                if row.get('Đơn vị', '') == 'Tổng':
                    ttvt_scores['C1.4'] = row.get('Điểm BSC', 0)
        if 'c15_ttvtst' in c1x_reports:
            for _, row in c1x_reports['c15_ttvtst'].iterrows():
                if 'TTVT' in row.get('DOIVT', ''):
                    ty_le = row.get('Tỉ lệ đạt (%)', 0)
                    if ty_le >= 99.5:
                        ttvt_scores['C1.5'] = 5.0
                    elif ty_le <= 89.5:
                        ttvt_scores['C1.5'] = 1.0
                    else:
                        ttvt_scores['C1.5'] = round(1 + 4 * (ty_le - 89.5) / 10, 2)
    
    headers = ['Đơn vị', 'C1.1', 'C1.2', 'C1.3', 'C1.4', 'C1.5']
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    set_table_border(table)
    
    for i, header in enumerate(table.rows[0].cells):
        header.text = headers[i]
        header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(header, '1F4E79')
        run = header.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    for idx, team in enumerate(teams_order, 1):
        cells = table.add_row().cells
        scores = bsc_data[team]
        data = [
            team,
            format_number(scores['C1.1']),
            format_number(scores['C1.2']),
            format_number(scores['C1.3']),
            format_number(scores['C1.4']),
            format_number(scores['C1.5'])
        ]
        for i, value in enumerate(data):
            cells[i].text = value
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cells[i].paragraphs[0].runs[0]
            run.font.size = Pt(10)
            if idx % 2 == 0:
                set_cell_shading(cells[i], 'E8F4FD')
    
    # Thêm dòng TTVT Sơn Tây (tổng hợp)
    cells = table.add_row().cells
    ttvt_data = [
        'TTVT Sơn Tây',
        format_number(ttvt_scores['C1.1']),
        format_number(ttvt_scores['C1.2']),
        format_number(ttvt_scores['C1.3']),
        format_number(ttvt_scores['C1.4']),
        format_number(ttvt_scores['C1.5'])
    ]
    for i, value in enumerate(ttvt_data):
        cells[i].text = value
        cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cells[i].paragraphs[0].runs[0]
        run.font.size = Pt(10)
        run.font.bold = True
        set_cell_shading(cells[i], 'B2DFDB')  # Màu xanh lá nhạt để nổi bật
    
    doc.add_paragraph()
    
    # 1.4 Số liệu chi tiết các chỉ tiêu BSC theo Đội/TTVT (sử dụng c1x_reports đã load)
    if c1x_reports:
        add_c1x_overview_table(doc, c1x_reports)
    
    # 1.5 Tổng quan Suy Hao Cao
    print("📊 Thêm phần Suy Hao Cao...")
    add_shc_overview_section(doc, data_folder="downloads/baocao_hanoi")
    
    doc.add_page_break()
    
    # =========================================================================
    # PHẦN 2: CHI TIẾT THEO TỔ
    # =========================================================================
    print("📋 Tạo phần Chi tiết theo tổ...")
    doc.add_heading('PHẦN 2: CHI TIẾT THEO TỪNG TỔ', level=1)
    
    for team_idx, team_name in enumerate(teams, 1):
        short_name = TEAM_SHORT_NAMES.get(team_name, team_name)
        print(f"   📁 Tổ {team_idx}: {short_name}")
        
        doc.add_heading(f'2.{team_idx}. {short_name}', level=2)
        
        # Bảng KPI tổng hợp của tổ
        doc.add_heading(f'Bảng điểm KPI tổng hợp', level=3)
        add_kpi_summary_table(doc, df_summary, team_name)
        doc.add_paragraph()
        
        # Biểu đồ cột so sánh NVKT
        doc.add_heading(f'Biểu đồ so sánh điểm KPI theo NVKT', level=3)
        nvkt_chart = create_nvkt_bar_chart(df_summary, team_name)
        if nvkt_chart:
            doc.add_picture(nvkt_chart, width=Inches(6.5))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        
        # Bảng chi tiết từng chỉ tiêu
        add_c11_detail_table(doc, df_detail, team_name)
        doc.add_paragraph()
        
        add_c12_detail_table(doc, df_detail, team_name)
        doc.add_paragraph()
        
        add_c14_detail_table(doc, df_detail, team_name)
        doc.add_paragraph()
        
        add_c15_detail_table(doc, df_detail, team_name)
        
        # Số liệu Suy Hao Cao cho tổ
        add_shc_unit_section(doc, team_name, data_folder="downloads/baocao_hanoi")
        
        # Thêm page break sau mỗi tổ (trừ tổ cuối)
        if team_idx < len(teams):
            doc.add_page_break()
    
    # =========================================================================
    # PHẦN 3: KẾT LUẬN
    # =========================================================================
    doc.add_page_break()
    doc.add_heading('PHẦN 3: KẾT LUẬN VÀ KIẾN NGHỊ', level=1)
    
    # Thống kê tổng quan
    total_nvkt = len(df_summary)
    avg_c11 = df_summary['Diem_C1.1'].mean()
    avg_c12 = df_summary['Diem_C1.2'].mean()
    avg_c14 = df_summary['Diem_C1.4'].mean()
    avg_c15 = df_summary['Diem_C1.5'].mean()
    
    doc.add_heading('3.1. Thống kê tổng quan', level=2)
    doc.add_paragraph(f"• Tổng số NVKT: {total_nvkt}")
    doc.add_paragraph(f"• Điểm trung bình C1.1: {format_number(avg_c11)}")
    doc.add_paragraph(f"• Điểm trung bình C1.2: {format_number(avg_c12)}")
    doc.add_paragraph(f"• Điểm trung bình C1.4: {format_number(avg_c14)}")
    doc.add_paragraph(f"• Điểm trung bình C1.5: {format_number(avg_c15)}")
    
    # Đánh giá
    doc.add_heading('3.2. Đánh giá chung', level=2)
    doc.add_paragraph("(Phần này cần bổ sung nội dung đánh giá theo thực tế)")
    
    doc.add_heading('3.3. Kiến nghị', level=2)
    doc.add_paragraph("(Phần này cần bổ sung nội dung kiến nghị theo thực tế)")
    
    # =========================================================================
    # PHỤ LỤC: BẢNG ĐIỂM KPI CHI TIẾT THEO NVKT
    # =========================================================================
    doc.add_page_break()
    doc.add_heading('PHỤ LỤC: BẢNG ĐIỂM KPI CHI TIẾT THEO NVKT', level=1)
    
    # Bảng KPI tổng hợp theo NVKT
    doc.add_heading('Bảng điểm KPI tổng hợp theo NVKT', level=2)
    add_kpi_summary_table(doc, df_summary)
    doc.add_paragraph()
    
    # Tạo biểu đồ bar cho từng tổ
    doc.add_heading('Biểu đồ điểm KPI theo NVKT (theo từng tổ)', level=2)
    
    for team_name in teams:
        short_name = TEAM_SHORT_NAMES.get(team_name, team_name)
        doc.add_heading(f'{short_name}', level=3)
        
        nvkt_chart = create_nvkt_bar_chart(df_summary, team_name)
        if nvkt_chart:
            doc.add_picture(nvkt_chart, width=Inches(6.5))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
    
    # =========================================================================
    # LƯU FILE
    # =========================================================================
    output_file = output_path / f"Bao_cao_KPI_NVKT_{report_month.replace('/', '_')}.docx"
    doc.save(output_file)
    
    print("="*60)
    print(f"✅ ĐÃ TẠO BÁO CÁO WORD THÀNH CÔNG!")
    print(f"   📄 File: {output_file}")
    print("="*60)
    
    return str(output_file)


# =============================================================================
# HÀM TẠO BÁO CÁO CÁ NHÂN CHO TỪNG NVKT
# =============================================================================

def sanitize_filename(name):
    """
    Chuẩn hóa tên file - loại bỏ ký tự đặc biệt và thay khoảng trắng bằng _
    """
    import re
    # Thay khoảng trắng bằng _
    name = name.replace(' ', '_')
    # Loại bỏ ký tự đặc biệt (giữ lại chữ cái Unicode, số và _)
    name = re.sub(r'[^\w\s-]', '', name, flags=re.UNICODE)
    return name


def create_individual_radar_chart(nvkt_data, output_path=None):
    """
    Tạo biểu đồ radar so sánh điểm KPI của 1 NVKT
    
    Args:
        nvkt_data: Dictionary chứa điểm KPI của NVKT
        output_path: Đường dẫn lưu file (None = trả về bytes)
    
    Returns:
        bytes hoặc str
    """
    # Lấy điểm các chỉ tiêu
    metrics = ['Diem_C1.1', 'Diem_C1.2', 'Diem_C1.4', 'Diem_C1.5']
    labels = ['C1.1\nSửa chữa', 'C1.2\nBáo hỏng', 'C1.4\nHài lòng', 'C1.5\nThiết lập DV']
    
    values = []
    for m in metrics:
        val = nvkt_data.get(m, 0)
        if pd.isna(val):
            val = 0
        values.append(val)
    
    # Số lượng biến
    num_vars = len(labels)
    
    # Tính góc cho mỗi trục
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    # Đóng vòng radar
    values += values[:1]
    angles += angles[:1]
    
    # Tạo figure
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    
    # Vẽ radar
    ax.fill(angles, values, color='#2E86AB', alpha=0.25)
    ax.plot(angles, values, color='#2E86AB', linewidth=2, marker='o', markersize=8)
    
    # Vẽ đường chuẩn 5 điểm
    target_values = [5] * (num_vars + 1)
    ax.plot(angles, target_values, color='#C73E1D', linewidth=1.5, linestyle='--', alpha=0.7, label='Mục tiêu (5 điểm)')
    
    # Thiết lập các trục
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=11)
    ax.set_ylim(0, 5.5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(['1', '2', '3', '4', '5'], fontsize=9)
    
    # Thêm giá trị lên điểm
    for angle, value, label in zip(angles[:-1], values[:-1], labels):
        ax.annotate(f'{value:.2f}', 
                   xy=(angle, value), 
                   xytext=(angle, value + 0.3),
                   ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1.1))
    ax.set_title('BIỂU ĐỒ ĐIỂM KPI', fontsize=14, fontweight='bold', pad=20)
    
    plt.tight_layout()
    
    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        return output_path
    else:
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf


def add_individual_summary_table(doc, nvkt_data):
    """
    Thêm bảng tổng hợp điểm KPI cho 1 NVKT vào document
    
    Args:
        doc: Document Word
        nvkt_data: Dictionary hoặc Series chứa dữ liệu NVKT
    """
    # Tạo bảng 2 cột: Chỉ tiêu - Điểm
    headers = ['Chỉ tiêu', 'Điểm']
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    set_table_border(table)
    
    # Header
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = headers[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '1F4E79')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(11)
    
    # Dữ liệu các chỉ tiêu
    kpi_info = [
        ('C1.1 - Tỷ lệ sửa chữa phiếu chất lượng & báo hỏng', 'Diem_C1.1'),
        ('C1.2 - Tỷ lệ báo hỏng lặp lại & sự cố dịch vụ', 'Diem_C1.2'),
        ('C1.4 - Độ hài lòng khách hàng', 'Diem_C1.4'),
        ('C1.5 - Tỷ lệ thiết lập dịch vụ đạt thời gian quy định', 'Diem_C1.5'),
    ]
    
    for idx, (label, col) in enumerate(kpi_info, 1):
        cells = table.add_row().cells
        cells[0].text = label
        cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        value = nvkt_data.get(col, np.nan)
        cells[1].text = format_number(value)
        cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for i, cell in enumerate(cells):
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(11)
            if idx % 2 == 0:
                set_cell_shading(cell, 'E8F4FD')
            
            # Tô màu điểm theo mức
            if i == 1 and not pd.isna(value):
                if value >= 4.5:
                    run.font.color.rgb = RGBColor(0, 128, 0)  # Xanh lá
                    run.font.bold = True
                elif value < 3:
                    run.font.color.rgb = RGBColor(255, 0, 0)  # Đỏ
                    run.font.bold = True


def add_individual_c11_detail(doc, nvkt_data):
    """
    Thêm chi tiết chỉ tiêu C1.1 cho 1 NVKT
    """
    doc.add_heading('2. CHI TIẾT CHỈ TIÊU C1.1 - CHẤT LƯỢNG SỬA CHỮA BRCĐ', level=2)
    
    # Thành phần 1
    doc.add_heading('2.1. Thành phần 1: Sửa chữa chủ động (SCCD ≤72h) - 30%', level=3)
    
    table1 = doc.add_table(rows=2, cols=4)
    table1.style = 'Table Grid'
    set_table_border(table1)
    
    headers1 = ['Tổng SCCD', 'Đạt ≤72h', 'Tỷ lệ (%)', 'Điểm']
    for i, cell in enumerate(table1.rows[0].cells):
        cell.text = headers1[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '2E7D32')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    data1 = [
        format_number(nvkt_data.get('c11_tp1_tong_phieu', np.nan), 0),
        format_number(nvkt_data.get('c11_tp1_phieu_dat', np.nan), 0),
        format_number(nvkt_data.get('c11_tp1_ty_le', np.nan)),
        format_number(nvkt_data.get('diem_c11_tp1', np.nan))
    ]
    for i, cell in enumerate(table1.rows[1].cells):
        cell.text = data1[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Thành phần 2
    doc.add_heading('2.2. Thành phần 2: Sửa chữa theo báo hỏng (đúng hạn) - 70%', level=3)
    
    table2 = doc.add_table(rows=2, cols=4)
    table2.style = 'Table Grid'
    set_table_border(table2)
    
    headers2 = ['Tổng SC BH', 'Đúng hạn', 'Tỷ lệ (%)', 'Điểm']
    for i, cell in enumerate(table2.rows[0].cells):
        cell.text = headers2[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '388E3C')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    data2 = [
        format_number(nvkt_data.get('c11_tp2_tong_phieu', np.nan), 0),
        format_number(nvkt_data.get('c11_tp2_phieu_dat', np.nan), 0),
        format_number(nvkt_data.get('c11_tp2_ty_le', np.nan)),
        format_number(nvkt_data.get('diem_c11_tp2', np.nan))
    ]
    for i, cell in enumerate(table2.rows[1].cells):
        cell.text = data2[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Điểm tổng hợp
    p = doc.add_paragraph()
    p.add_run('➤ ĐIỂM TỔNG HỢP C1.1: ').bold = True
    diem_c11 = nvkt_data.get('Diem_C1.1', np.nan)
    run = p.add_run(format_number(diem_c11))
    run.bold = True
    run.font.size = Pt(12)
    if not pd.isna(diem_c11):
        if diem_c11 >= 4.5:
            run.font.color.rgb = RGBColor(0, 128, 0)
        elif diem_c11 < 3:
            run.font.color.rgb = RGBColor(255, 0, 0)
    
    p.add_run(' (= TP1 × 30% + TP2 × 70%)')


def add_individual_c12_detail(doc, nvkt_data):
    """
    Thêm chi tiết chỉ tiêu C1.2 cho 1 NVKT
    """
    doc.add_heading('3. CHI TIẾT CHỈ TIÊU C1.2 - TỶ LỆ THUÊ BAO BÁO HỎNG', level=2)
    
    # Thành phần 1
    doc.add_heading('3.1. Thành phần 1: Hỏng lặp (≥2 lần/7 ngày) - 50%', level=3)
    
    table1 = doc.add_table(rows=2, cols=4)
    table1.style = 'Table Grid'
    set_table_border(table1)
    
    headers1 = ['TB hỏng lặp', 'Tổng BH', 'Tỷ lệ (%)', 'Điểm']
    for i, cell in enumerate(table1.rows[0].cells):
        cell.text = headers1[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '1565C0')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    data1 = [
        format_number(nvkt_data.get('c12_tp1_phieu_hll', np.nan), 0),
        format_number(nvkt_data.get('c12_tp1_phieu_bh', np.nan), 0),
        format_number(nvkt_data.get('c12_tp1_ty_le', np.nan)),
        format_number(nvkt_data.get('diem_c12_tp1', np.nan))
    ]
    for i, cell in enumerate(table1.rows[1].cells):
        cell.text = data1[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Thành phần 2
    doc.add_heading('3.2. Thành phần 2: Tỷ lệ BH/TB quản lý (‰) - 50%', level=3)
    
    table2 = doc.add_table(rows=2, cols=4)
    table2.style = 'Table Grid'
    set_table_border(table2)
    
    headers2 = ['Phiếu BH', 'TB quản lý', 'Tỷ lệ (‰)', 'Điểm']
    for i, cell in enumerate(table2.rows[0].cells):
        cell.text = headers2[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '1976D2')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    data2 = [
        format_number(nvkt_data.get('c12_tp2_phieu_bh', np.nan), 0),
        format_number(nvkt_data.get('c12_tp2_tong_tb', np.nan), 0),
        format_number(nvkt_data.get('c12_tp2_ty_le', np.nan)),
        format_number(nvkt_data.get('diem_c12_tp2', np.nan))
    ]
    for i, cell in enumerate(table2.rows[1].cells):
        cell.text = data2[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Điểm tổng hợp
    p = doc.add_paragraph()
    p.add_run('➤ ĐIỂM TỔNG HỢP C1.2: ').bold = True
    diem_c12 = nvkt_data.get('Diem_C1.2', np.nan)
    run = p.add_run(format_number(diem_c12))
    run.bold = True
    run.font.size = Pt(12)
    if not pd.isna(diem_c12):
        if diem_c12 >= 4.5:
            run.font.color.rgb = RGBColor(0, 128, 0)
        elif diem_c12 < 3:
            run.font.color.rgb = RGBColor(255, 0, 0)
    
    p.add_run(' (= TP1 × 50% + TP2 × 50%)')


def add_individual_c14_detail(doc, nvkt_data):
    """
    Thêm chi tiết chỉ tiêu C1.4 cho 1 NVKT
    """
    doc.add_heading('4. CHI TIẾT CHỈ TIÊU C1.4 - ĐỘ HÀI LÒNG KHÁCH HÀNG', level=2)
    
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    set_table_border(table)
    
    headers = ['Phiếu KS thành công', 'Phiếu KH không HL', 'Tỷ lệ HL (%)', 'Điểm']
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = headers[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, 'F57C00')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    data = [
        format_number(nvkt_data.get('c14_phieu_ks', np.nan), 0),
        format_number(nvkt_data.get('c14_phieu_khl', np.nan), 0),
        format_number(nvkt_data.get('c14_ty_le', np.nan)),
        format_number(nvkt_data.get('Diem_C1.4', np.nan))
    ]
    for i, cell in enumerate(table.rows[1].cells):
        cell.text = data[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Công thức tính điểm
    p = doc.add_paragraph()
    p.add_run('📌 Công thức tính điểm: ').bold = True
    p.add_run('≥ 99.5% = 5 điểm, > 95% = nội suy, ≤ 95% = 1 điểm')


def add_individual_c15_detail(doc, nvkt_data):
    """
    Thêm chi tiết chỉ tiêu C1.5 cho 1 NVKT
    """
    doc.add_heading('5. CHI TIẾT CHỈ TIÊU C1.5 - TỶ LỆ THIẾT LẬP DỊCH VỤ ĐẠT', level=2)
    
    table = doc.add_table(rows=2, cols=5)
    table.style = 'Table Grid'
    set_table_border(table)
    
    headers = ['Phiếu đạt', 'Phiếu không đạt', 'Tổng phiếu', 'Tỷ lệ đạt (%)', 'Điểm']
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = headers[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, '7B1FA2')
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(10)
    
    data = [
        format_number(nvkt_data.get('c15_phieu_dat', np.nan), 0),
        format_number(nvkt_data.get('c15_phieu_khong_dat', np.nan), 0),
        format_number(nvkt_data.get('c15_tong_phieu', np.nan), 0),
        format_number(nvkt_data.get('c15_ty_le', np.nan)),
        format_number(nvkt_data.get('Diem_C1.5', np.nan))
    ]
    for i, cell in enumerate(table.rows[1].cells):
        cell.text = data[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Công thức tính điểm
    p = doc.add_paragraph()
    p.add_run('📌 Công thức tính điểm: ').bold = True
    p.add_run('≥ 99.5% = 5 điểm, 89.5% < KQ < 99.5% = nội suy, ≤ 89.5% = 1 điểm')


def load_shc_trend_data(nvkt_name, data_folder="downloads/baocao_hanoi"):
    """
    Đọc dữ liệu xu hướng SHC cho 1 NVKT từ file Bao_cao_xu_huong_*.xlsx
    
    Args:
        nvkt_name: Tên NVKT cần tìm
        data_folder: Thư mục chứa file báo cáo
    
    Returns:
        dict: {'dates': [...], 'values': [...], 'don_vi': '...'} hoặc None
    """
    import glob
    
    # Tìm file Bao_cao_xu_huong mới nhất
    pattern = os.path.join(data_folder, "Bao_cao_xu_huong_SHC_*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        return None
    
    # Sắp xếp theo thời gian chỉnh sửa file (mới nhất cuối cùng)
    latest_file = max(files, key=os.path.getmtime)
    print(f"   📊 Sử dụng file SHC: {os.path.basename(latest_file)}")
    
    try:
        # Đọc sheet Xu_huong_theo_NVKT
        df = pd.read_excel(latest_file, sheet_name='Xu_huong_theo_NVKT')
        
        # Tìm NVKT trong cột 'NVKT'
        nvkt_row = df[df['NVKT'] == nvkt_name]
        
        if nvkt_row.empty:
            return None
        
        row = nvkt_row.iloc[0]
        don_vi = row.get('Đơn vị', '')
        
        # Lấy các cột ngày (không phải 'Đơn vị' và 'NVKT')
        date_columns = [col for col in df.columns if col not in ['Đơn vị', 'NVKT']]
        
        dates = []
        values = []
        for col in date_columns:
            dates.append(str(col))
            val = row[col]
            values.append(int(val) if pd.notna(val) else 0)
        
        return {
            'dates': dates,
            'values': values,
            'don_vi': don_vi
        }
    except Exception as e:
        print(f"   ⚠️ Lỗi đọc dữ liệu SHC: {e}")
        return None


def create_shc_trend_bar_chart(shc_data, nvkt_name, output_path=None):
    """
    Tạo biểu đồ cột thể hiện xu hướng số TB suy hao cao theo ngày
    
    Args:
        shc_data: dict với keys 'dates' và 'values'
        nvkt_name: Tên NVKT
        output_path: Đường dẫn lưu file (None = trả về bytes)
    
    Returns:
        bytes hoặc str: Chart image
    """
    if not shc_data or not shc_data.get('dates') or not shc_data.get('values'):
        return None
    
    dates = shc_data['dates']
    values = shc_data['values']
    
    # Tạo figure
    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Vẽ biểu đồ cột
    x_pos = range(len(dates))
    bars = ax.bar(x_pos, values, color='#2196F3', edgecolor='#1565C0', linewidth=1)
    
    # Thêm giá trị lên cột
    for bar, val in zip(bars, values):
        height = bar.get_height()
        ax.annotate(f'{val}',
                   xy=(bar.get_x() + bar.get_width() / 2, height),
                   xytext=(0, 3),
                   textcoords="offset points",
                   ha='center', va='bottom',
                   fontsize=9, fontweight='bold')
    
    # Thiết lập trục
    ax.set_xticks(x_pos)
    ax.set_xticklabels(dates, rotation=45, ha='right', fontsize=9)
    ax.set_ylabel('Số TB suy hao cao', fontsize=11)
    ax.set_xlabel('Ngày', fontsize=11)
    ax.set_title(f'XU HƯỚNG SỐ TB SUY HAO CAO - {nvkt_name}', fontsize=12, fontweight='bold', pad=15)
    
    # Grid và layout
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.set_axisbelow(True)
    
    # Đường xu hướng trung bình
    if len(values) > 1:
        avg = sum(values) / len(values)
        ax.axhline(y=avg, color='#E91E63', linestyle='--', linewidth=1.5, alpha=0.7, 
                   label=f'TB: {avg:.1f}')
        ax.legend(loc='upper right')
    
    plt.tight_layout()
    
    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        return output_path
    else:
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        return buf


def add_individual_shc_section(doc, nvkt_name, data_folder="downloads/baocao_hanoi"):
    """
    Thêm phần Số liệu Suy Hao Cao vào báo cáo cá nhân
    Bao gồm: Bảng dữ liệu + Biểu đồ bar
    
    Args:
        doc: Document Word
        nvkt_name: Tên NVKT
        data_folder: Thư mục chứa file báo cáo
    """
    # Load dữ liệu SHC
    shc_data = load_shc_trend_data(nvkt_name, data_folder)
    
    if not shc_data:
        return  # Không có dữ liệu SHC
    
    doc.add_page_break()
    doc.add_heading('6. SỐ LIỆU SUY HAO CAO', level=2)
    
    dates = shc_data['dates']
    values = shc_data['values']
    
    # Mô tả
    p = doc.add_paragraph()
    p.add_run('📊 Xu hướng số thuê bao suy hao cao theo ngày:').bold = True
    
    doc.add_paragraph()
    
    # Tạo bảng dữ liệu (chia thành các nhóm nếu nhiều ngày)
    max_cols = 10  # Số cột tối đa mỗi bảng
    
    for i in range(0, len(dates), max_cols):
        chunk_dates = dates[i:i+max_cols]
        chunk_values = values[i:i+max_cols]
        
        table = doc.add_table(rows=2, cols=len(chunk_dates))
        table.style = 'Table Grid'
        set_table_border(table)
        
        # Header row - Ngày
        for j, date in enumerate(chunk_dates):
            cell = table.rows[0].cells[j]
            cell.text = str(date)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell, '1E88E5')
            run = cell.paragraphs[0].runs[0]
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
        
        # Data row - Số lượng
        for j, val in enumerate(chunk_values):
            cell = table.rows[1].cells[j]
            cell.text = str(val)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(10)
            run.font.bold = True
            
            # Tô màu theo mức độ
            if val == 0:
                set_cell_shading(cell, 'C8E6C9')  # Xanh lá nhạt
                run.font.color.rgb = RGBColor(0, 128, 0)
            elif val <= 3:
                set_cell_shading(cell, 'FFF9C4')  # Vàng nhạt
            else:
                set_cell_shading(cell, 'FFCDD2')  # Đỏ nhạt
                run.font.color.rgb = RGBColor(200, 0, 0)
        
        doc.add_paragraph()
    
    # Thống kê tổng quan
    total = sum(values)
    avg = total / len(values) if values else 0
    max_val = max(values) if values else 0
    min_val = min(values) if values else 0
    
    p = doc.add_paragraph()
    p.add_run(f'📈 Tổng số TB SHC trong kỳ: ').bold = True
    p.add_run(f'{total}')
    
    p = doc.add_paragraph()
    p.add_run(f'📊 Trung bình/ngày: ').bold = True
    p.add_run(f'{avg:.1f}')
    
    p = doc.add_paragraph()
    p.add_run(f'⬆️ Cao nhất: ').bold = True
    p.add_run(f'{max_val}')
    p.add_run(f'  |  ')
    p.add_run(f'⬇️ Thấp nhất: ').bold = True
    p.add_run(f'{min_val}')
    
    doc.add_paragraph()
    
    # Biểu đồ bar
    try:
        chart = create_shc_trend_bar_chart(shc_data, nvkt_name)
        if chart:
            doc.add_picture(chart, width=Inches(6))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception as e:
        print(f"   ⚠️ Không thể tạo biểu đồ SHC: {e}")


def generate_individual_kpi_report(nvkt_name, don_vi, kpi_folder=DEFAULT_KPI_FOLDER, 
                                    output_folder=DEFAULT_OUTPUT_FOLDER, report_month=None):
    """
    Tạo báo cáo Word cho 1 NVKT cụ thể
    
    Args:
        nvkt_name: Tên NVKT (vd: "Bùi Văn Duẩn")
        don_vi: Tên đơn vị/tổ (vd: "Tổ Kỹ thuật Địa bàn Phúc Thọ")
        kpi_folder: Thư mục chứa file KPI
        output_folder: Thư mục xuất báo cáo
        report_month: Tháng báo cáo (vd: "01/2026")
    
    Returns:
        str: Đường dẫn file Word đã tạo
    """
    # Xác định tháng báo cáo
    if report_month is None:
        report_month = datetime.now().strftime("%m/%Y")
    
    # Đọc dữ liệu KPI chi tiết
    kpi_path = Path(kpi_folder)
    detail_file = kpi_path / "KPI_NVKT_ChiTiet.xlsx"
    df_detail = pd.read_excel(detail_file)
    
    # Lọc dữ liệu cho NVKT cụ thể
    nvkt_df = df_detail[(df_detail['nvkt'] == nvkt_name) & (df_detail['don_vi'] == don_vi)]
    
    if nvkt_df.empty:
        print(f"⚠️ Không tìm thấy dữ liệu cho NVKT: {nvkt_name} - {don_vi}")
        return None
    
    nvkt_data = nvkt_df.iloc[0].to_dict()
    
    # Tạo thư mục output theo tổ
    short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
    folder_name = sanitize_filename(short_name)
    output_path = Path(output_folder) / "individual_reports" / folder_name
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Tạo document
    doc = Document()
    
    # Thiết lập style
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # =========================================================================
    # TIÊU ĐỀ
    # =========================================================================
    created_time = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    title = doc.add_heading(level=0)
    title_run = title.add_run('BÁO CÁO KẾT QUẢ BSC/KPI CÁ NHÂN')
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle = doc.add_heading(level=1)
    subtitle_run = subtitle.add_run(f'THÁNG {report_month}')
    subtitle_run.font.size = Pt(16)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Thông tin cá nhân
    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = 'Table Grid'
    
    info_data = [
        ('Họ và tên:', nvkt_name),
        ('Đơn vị:', short_name),
        ('Ngày tạo báo cáo:', created_time)
    ]
    
    for i, (label, value) in enumerate(info_data):
        info_table.rows[i].cells[0].text = label
        info_table.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
        info_table.rows[i].cells[1].text = value
    
    doc.add_paragraph()
    
    # =========================================================================
    # PHẦN 1: TỔNG QUAN
    # =========================================================================
    doc.add_heading('1. TỔNG QUAN ĐIỂM KPI', level=2)
    
    # Bảng tổng hợp
    add_individual_summary_table(doc, nvkt_data)
    doc.add_paragraph()
    
    # Biểu đồ radar
    try:
        radar_chart = create_individual_radar_chart(nvkt_data)
        doc.add_picture(radar_chart, width=Inches(5))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception as e:
        print(f"   ⚠️ Không thể tạo biểu đồ radar: {e}")
    
    doc.add_page_break()
    
    # =========================================================================
    # CHI TIẾT TỪNG CHỈ TIÊU
    # =========================================================================
    add_individual_c11_detail(doc, nvkt_data)
    doc.add_paragraph()
    
    add_individual_c12_detail(doc, nvkt_data)
    doc.add_paragraph()
    
    add_individual_c14_detail(doc, nvkt_data)
    doc.add_paragraph()
    
    add_individual_c15_detail(doc, nvkt_data)
    
    # =========================================================================
    # PHẦN 6: SỐ LIỆU SUY HAO CAO
    # =========================================================================
    add_individual_shc_section(doc, nvkt_name, data_folder="downloads/baocao_hanoi")
    
    # =========================================================================
    # LƯU FILE
    # =========================================================================
    safe_name = sanitize_filename(nvkt_name)
    output_file = output_path / f"Bao_cao_KPI_{safe_name}_{report_month.replace('/', '_')}.docx"
    doc.save(output_file)
    
    return str(output_file)



def generate_all_individual_reports(kpi_folder=DEFAULT_KPI_FOLDER, output_folder=DEFAULT_OUTPUT_FOLDER, 
                                     report_month=None):
    """
    Tạo báo cáo cho TẤT CẢ NVKT
    
    Args:
        kpi_folder: Thư mục chứa file KPI
        output_folder: Thư mục xuất báo cáo
        report_month: Tháng báo cáo
    
    Returns:
        list: Danh sách đường dẫn các file đã tạo
    """
    print("="*60)
    print("📝 BẮT ĐẦU TẠO BÁO CÁO KPI CÁ NHÂN CHO TẤT CẢ NVKT")
    print("="*60)
    
    if report_month is None:
        report_month = datetime.now().strftime("%m/%Y")
    
    # Đọc dữ liệu KPI
    kpi_path = Path(kpi_folder)
    detail_file = kpi_path / "KPI_NVKT_ChiTiet.xlsx"
    df_detail = pd.read_excel(detail_file)
    
    # Lấy danh sách NVKT
    nvkt_list = df_detail[['don_vi', 'nvkt']].drop_duplicates()
    total = len(nvkt_list)
    
    print(f"📊 Tìm thấy {total} NVKT")
    print()
    
    success_files = []
    failed_count = 0
    
    for idx, row in nvkt_list.iterrows():
        don_vi = row['don_vi']
        nvkt_name = row['nvkt']
        short_name = TEAM_SHORT_NAMES.get(don_vi, don_vi)
        
        current = len(success_files) + failed_count + 1
        print(f"   [{current}/{total}] {nvkt_name} ({short_name})...", end=" ")
        
        try:
            result = generate_individual_kpi_report(
                nvkt_name=nvkt_name,
                don_vi=don_vi,
                kpi_folder=kpi_folder,
                output_folder=output_folder,
                report_month=report_month
            )
            if result:
                success_files.append(result)
                print("✅")
            else:
                failed_count += 1
                print("❌ (không có dữ liệu)")
        except Exception as e:
            failed_count += 1
            print(f"❌ ({str(e)[:30]})")
    
    print()
    print("="*60)
    print(f"✅ HOÀN THÀNH!")
    print(f"   📄 Đã tạo: {len(success_files)} báo cáo")
    print(f"   ❌ Thất bại: {failed_count}")
    print(f"   📁 Thư mục: {Path(output_folder) / 'individual_reports'}")
    print("="*60)
    
    return success_files


def generate_all_individual_reports_after_exclusion(kpi_folder, output_root, report_month=None):
    """
    Tạo báo cáo cá nhân sau giảm trừ, phân loại theo thư mục Tổ kỹ thuật
    Lưu tại: {output_root}/ca_nhan/{tên tổ kỹ thuật}/
    """
    print("="*60)
    print("📝 BẮT ĐẦU TẠO BÁO CÁO KPI CÁ NHÂN SAU GIẢM TRỪ")
    print("="*60)
    
    if report_month is None:
        report_month = datetime.now().strftime("%m/%Y")
        
    kpi_path = Path(kpi_folder)
    detail_file = kpi_path / "KPI_NVKT_SauGiamTru_ChiTiet.xlsx"
    summary_file = kpi_path / "KPI_NVKT_SauGiamTru_TomTat.xlsx"
    
    if not detail_file.exists():
        print(f"❌ Không tìm thấy file: {detail_file}")
        return 0
        
    df_detail = pd.read_excel(detail_file)
    nvkt_list = df_detail[['don_vi', 'nvkt']].drop_duplicates()
    total = len(nvkt_list)
    
    print(f"📊 Tìm thấy {total} NVKT sau giảm trừ")
    
    success_count = 0
    for idx, row in nvkt_list.iterrows():
        don_vi = row['don_vi']
        nvkt_name = row['nvkt']
        
        # Đảm bảo don_vi là chuỗi
        don_vi_str = str(don_vi) if pd.notna(don_vi) else "Unknown"
        
        # Tạo thư mục cho từng Đội (Tổ)
        team_folder_name = sanitize_filename(don_vi_str)
        team_output_path = Path(output_root) / "ca_nhan" / team_folder_name
        team_output_path.mkdir(parents=True, exist_ok=True)
        
        current = success_count + 1
        print(f"   [{current}/{total}] {nvkt_name} ({don_vi})...", end=" ")
        
        try:
            # Tạo document mới
            doc = Document()
            
            # Thiết lập style mặc định cho doc
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
            
            # Lấy data NVKT
            # Lấy data NVKT - sử dụng logic lọc an toàn với NaN
            if pd.isna(don_vi):
                mask = (df_detail['nvkt'] == nvkt_name) & (df_detail['don_vi'].isna())
            else:
                mask = (df_detail['nvkt'] == nvkt_name) & (df_detail['don_vi'] == don_vi)
            
            nvkt_df_match = df_detail[mask]
            if nvkt_df_match.empty:
                print(f"❌ (Không tìm thấy data)")
                continue
                
            nvkt_data = nvkt_df_match.iloc[0].to_dict()
            
            short_name = TEAM_SHORT_NAMES.get(don_vi_str, don_vi_str)
            created_time = datetime.now().strftime("%d/%m/%Y %H:%M")
            
            # Header
            header = doc.sections[0].header
            p = header.paragraphs[0]
            p.text = f"BÁO CÁO KẾT QUẢ KPI CÁ NHÂN - THÁNG {report_month} (SAU GIẢM TRỪ)"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Title
            title = doc.add_heading(level=0)
            title_run = title.add_run('BÁO CÁO KẾT QUẢ BSC/KPI CÁ NHÂN (SAU GIẢM TRỪ)')
            title_run.font.size = Pt(18)
            title_run.font.bold = True
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            subtitle = doc.add_heading(level=1)
            subtitle_run = subtitle.add_run(f'THÁNG {report_month}')
            subtitle_run.font.size = Pt(16)
            subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph()
            
            # Thông tin cá nhân
            info_table = doc.add_table(rows=3, cols=2)
            info_table.style = 'Table Grid'
            info_data = [
                ('Họ và tên:', nvkt_name),
                ('Đơn vị:', short_name),
                ('Ngày tạo báo cáo:', created_time)
            ]
            for i, (label, value) in enumerate(info_data):
                info_table.rows[i].cells[0].text = label
                info_table.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
                info_table.rows[i].cells[1].text = value
            
            doc.add_paragraph()
            
            # Phần 1: Tổng quan
            doc.add_heading('1. TỔNG QUAN ĐIỂM KPI', level=2)
            add_individual_summary_table(doc, nvkt_data)
            
            # Biểu đồ radar
            try:
                radar_chart = create_individual_radar_chart(nvkt_data)
                doc.add_picture(radar_chart, width=Inches(5))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except:
                pass
                
            doc.add_page_break()
            
            # Chi tiết từng chỉ tiêu
            add_individual_c11_detail(doc, nvkt_data)
            doc.add_paragraph()
            add_individual_c12_detail(doc, nvkt_data)
            doc.add_paragraph()
            add_individual_c14_detail(doc, nvkt_data)
            doc.add_paragraph()
            add_individual_c15_detail(doc, nvkt_data)
            
            # Suy hao cao - truyền data_folder mặc định
            add_individual_shc_section(doc, nvkt_name, data_folder="downloads/baocao_hanoi")
            
            # Lưu file
            safe_name = sanitize_filename(nvkt_name)
            filename = f"Bao_cao_KPI_{safe_name}_SauGT_{report_month.replace('/', '_')}.docx"
            output_file = team_output_path / filename
            doc.save(output_file)
            
            success_count += 1
            print("✅")
        except Exception as e:
            print(f"❌ (Lỗi: {e})")
            
    print(f"\n✅ Hoàn thành: Đã tạo {success_count}/{total} báo cáo cá nhân sau giảm trừ.")
    print(f"📁 Thư mục xuất: {output_root}/ca_nhan/")
    
    return success_count


# =============================================================================
# MAIN - Chạy trực tiếp module
# =============================================================================
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Tạo báo cáo KPI/BSC cho NVKT')
    parser.add_argument('--individual', action='store_true', 
                       help='Tạo báo cáo cá nhân cho từng NVKT')
    parser.add_argument('--all', action='store_true',
                       help='Tạo báo cáo cho tất cả NVKT (dùng với --individual)')
    parser.add_argument('--nvkt', type=str, default=None,
                       help='Tên NVKT cụ thể (dùng với --individual)')
    parser.add_argument('--donvi', type=str, default=None,
                       help='Tên đơn vị của NVKT (dùng với --individual --nvkt)')
    parser.add_argument('--month', type=str, default="01/2026",
                       help='Tháng báo cáo (vd: 01/2026)')
    parser.add_argument('--kpi-folder', type=str, default="downloads/KPI",
                       help='Thư mục chứa file KPI')
    parser.add_argument('--output-folder', type=str, default="downloads/reports",
                       help='Thư mục xuất báo cáo')
    
    args = parser.parse_args()
    
    if args.individual:
        if args.all:
            # Tạo báo cáo cho tất cả NVKT
            generate_all_individual_reports(
                kpi_folder=args.kpi_folder,
                output_folder=args.output_folder,
                report_month=args.month
            )
        elif args.nvkt and args.donvi:
            # Tạo báo cáo cho 1 NVKT cụ thể
            result = generate_individual_kpi_report(
                nvkt_name=args.nvkt,
                don_vi=args.donvi,
                kpi_folder=args.kpi_folder,
                output_folder=args.output_folder,
                report_month=args.month
            )
            if result:
                print(f"✅ Đã tạo báo cáo: {result}")
            else:
                print("❌ Không thể tạo báo cáo")
        else:
            print("❌ Lỗi: Cần chỉ định --all hoặc --nvkt và --donvi")
            print("   Ví dụ: python report_generator.py --individual --all")
            print("   Hoặc:  python report_generator.py --individual --nvkt 'Bùi Văn Duẩn' --donvi 'Tổ Kỹ thuật Địa bàn Phúc Thọ'")
    else:
        # Tạo cả báo cáo tổng hợp và báo cáo cá nhân (mặc định)
        print("=" * 60)
        print("📊 TẠO BÁO CÁO TỔNG HỢP")
        print("=" * 60)
        report_path = generate_kpi_report(
            kpi_folder=args.kpi_folder,
            output_folder=args.output_folder,
            report_month=args.month
        )
        print(f"\n📁 Đường dẫn báo cáo tổng hợp: {report_path}")
        
        print("\n")
        print("=" * 60)
        print("📝 TẠO BÁO CÁO CÁ NHÂN CHO TẤT CẢ NVKT")
        print("=" * 60)
        generate_all_individual_reports(
            kpi_folder=args.kpi_folder,
            output_folder=args.output_folder,
            report_month=args.month
        )

