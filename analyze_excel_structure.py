#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import os
from pathlib import Path
import json

def analyze_excel_file(file_path):
    """Phân tích cấu trúc của một file Excel"""
    try:
        # Đọc tất cả các sheet trong file
        excel_file = pd.ExcelFile(file_path)

        result = {
            'file_name': os.path.basename(file_path),
            'file_path': file_path,
            'sheets': []
        }

        for sheet_name in excel_file.sheet_names:
            # Đọc sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            sheet_info = {
                'sheet_name': sheet_name,
                'num_rows': len(df),
                'num_cols': len(df.columns),
                'columns': list(df.columns),
                'sample_data': df.head(3).to_dict(orient='records') if len(df) > 0 else []
            }

            result['sheets'].append(sheet_info)

        return result
    except Exception as e:
        return {
            'file_name': os.path.basename(file_path),
            'file_path': file_path,
            'error': str(e)
        }

def main():
    base_dir = '/home/vtst/baocaohanoi'

    # Các loại file cần phân tích (chọn mẫu đại diện)
    sample_files = [
        'downloads/baocao_hanoi/hoan_cong_20112025.xlsx',
        'downloads/baocao_hanoi/ngung_psc_20112025.xlsx',
        'downloads/baocao_hanoi/thuc_tang_20112025.xlsx',
        'downloads/baocao_hanoi/mytv_hoan_cong_20112025.xlsx',
        'downloads/baocao_hanoi/Bao_cao_tuan_47_2025.xlsx',
        'downloads/baocao_hanoi/Bao_cao_thang_11_2025.xlsx',
        'downloads/baocao_hanoi/Bao_cao_xu_huong_20251101_20251119.xlsx',
        'downloads/baocao_hanoi/c1.1 report.xlsx',
        'downloads/baocao_hanoi/c1.2 report.xlsx',
        'downloads/baocao_hanoi/c1.5 report.xlsx',
        'downloads/baocao_hanoi/I1.5 report.xlsx',
        'downloads/baocao_hanoi/download_KR6_report_tong_hop.xlsx',
        'downloads/baocao_hanoi/download_KR7_report_tong_hop.xlsx',
        'downloads/baocao_hanoi/bc_thu_hoi_vat_tu.xlsx',
        'downloads/baocao_hanoi/SM1-C12.xlsx',
        'downloads/baocao_hanoi/SM2-C11.xlsx',
        'dsnv.xlsx',
    ]

    results = []

    for file_path in sample_files:
        full_path = os.path.join(base_dir, file_path)
        if os.path.exists(full_path):
            print(f"\n{'='*80}")
            print(f"Đang phân tích: {file_path}")
            print('='*80)

            result = analyze_excel_file(full_path)
            results.append(result)

            if 'error' in result:
                print(f"LỖI: {result['error']}")
            else:
                for sheet in result['sheets']:
                    print(f"\nSheet: {sheet['sheet_name']}")
                    print(f"  - Số dòng: {sheet['num_rows']}")
                    print(f"  - Số cột: {sheet['num_cols']}")
                    print(f"  - Các cột: {sheet['columns']}")
                    if sheet['sample_data']:
                        print(f"  - Dữ liệu mẫu (3 dòng đầu):")
                        for idx, row in enumerate(sheet['sample_data'], 1):
                            print(f"    Dòng {idx}: {row}")

    # Lưu kết quả vào file JSON
    output_file = os.path.join(base_dir, 'excel_structure_analysis.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n\nĐã lưu kết quả phân tích vào: {output_file}")

if __name__ == '__main__':
    main()
