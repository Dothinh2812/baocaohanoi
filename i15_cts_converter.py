# -*- coding: utf-8 -*-
"""
Module chuyển đổi file suy_hao_cts.xlsx sang định dạng I1.5 report.xlsx
Sử dụng như fallback khi không tải được báo cáo I1.5 gốc.

Workflow:
1. Đọc suy_hao_cts.xlsx
2. Parse cột Cổng để trích xuất IP, card, slot, port, ONT_ID
3. Tra cứu danhba.db bằng SYSTEM + CARD + PORT + ONT_ID
4. Tạo file I1.5 report.xlsx với định dạng tương thích

Matching logic:
- suy_hao_cts['Thiết bị'] = danhba['SYSTEM']
- Cổng format: 10.31.24.131/0/1/5:8 = IP/card/slot/port:ONT_ID
- Mapping: slot → CARD, port → PORT, ont_id → ONT_ID
- Match: SYSTEM + CARD + PORT + ONT_ID
"""

import pandas as pd
import os
import sqlite3
import re
from datetime import datetime


def parse_cong_column(cong_str):
    """
    Parse cột Cổng thành các thành phần.
    
    Input format: 10.31.24.131/0/1/5:8
    - IP: 10.31.24.131
    - Card: 0 (không dùng)
    - Slot: 1 → tương ứng với CARD trong danhba.db
    - Port: 5 → tương ứng với PORT trong danhba.db
    - ONT_ID: 8 → tương ứng với ONT_ID trong danhba.db
    
    Returns:
        dict: {'ip': str, 'card': int, 'slot': int, 'port': int, 'ont_id': int}
        hoặc None nếu không parse được
    """
    if not isinstance(cong_str, str) or not cong_str.strip():
        return None
    
    try:
        # Format: IP/card/slot/port:ont_id
        # Example: 10.31.24.131/0/1/5:8
        
        parts = cong_str.split('/')
        if len(parts) < 4:
            return None
        
        ip = parts[0]
        card = int(parts[1])
        slot = int(parts[2])
        
        # Phần cuối có format port:ont_id
        port_ont = parts[3]
        if ':' in port_ont:
            port_str, ont_str = port_ont.split(':')
            port = int(port_str)
            ont_id = int(ont_str)
        else:
            port = int(port_ont)
            ont_id = None
        
        return {
            'ip': ip,
            'card': card,
            'slot': slot,
            'port': port,
            'ont_id': ont_id
        }
    except Exception as e:
        return None


def convert_cts_to_i15_report(force_update=False):
    """
    Chuyển đổi file suy_hao_cts.xlsx sang định dạng I1.5 report.xlsx
    
    Matching logic:
    - Thiết bị (CTS) = SYSTEM (danhba.db)
    - Slot → CARD, Port → PORT, ONT_ID từ cột Cổng
    
    Args:
        force_update: Nếu True, ghi đè file output nếu đã tồn tại
    
    Returns:
        bool: True nếu thành công, False nếu thất bại
    """
    try:
        print("\n" + "="*80)
        print("CHUYỂN ĐỔI SỮY HAO CTS → I1.5 REPORT")
        print("="*80)
        
        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "suy_hao_cts.xlsx")
        output_file = os.path.join("downloads", "baocao_hanoi", "I1.5 report.xlsx")
        db_file = "danhba.db"
        
        # Kiểm tra file input
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False
        
        # Kiểm tra file output đã tồn tại
        if os.path.exists(output_file) and not force_update:
            print(f"⚠️ File {output_file} đã tồn tại")
            print("   Sử dụng force_update=True để ghi đè")
            return False
        
        print(f"\n✓ Đang đọc file: {input_file}")
        df_cts = pd.read_excel(input_file)
        print(f"✅ Đã đọc {len(df_cts)} dòng từ suy_hao_cts.xlsx")
        
        # Lấy ngày báo cáo
        # LƯU Ý: File CTS sử dụng format MM/DD/YYYY (02/06/2026 = ngày 6 tháng 2)
        # Nhưng pandas/Excel có thể tự động parse sai thành DD/MM/YYYY
        # Cần hoán đổi ngày và tháng nếu đã bị parse sai
        if 'Ngày' in df_cts.columns and len(df_cts) > 0:
            ngay_val = df_cts['Ngày'].iloc[0]
            try:
                if isinstance(ngay_val, str):
                    # Nếu là string, parse theo format MM/DD/YYYY (format CTS)
                    parsed_date = pd.to_datetime(ngay_val, format='%m/%d/%Y')
                else:
                    # Nếu là Timestamp/datetime (pandas đã tự parse)
                    # pandas có thể đã parse sai (đọc 02/06 thành ngày 2 tháng 6 thay vì ngày 6 tháng 2)
                    # Cần hoán đổi ngày và tháng
                    parsed_date = pd.to_datetime(ngay_val)
                    # Hoán đổi: ngày -> tháng, tháng -> ngày
                    parsed_date = parsed_date.replace(month=parsed_date.day, day=parsed_date.month)
                
                report_date = parsed_date.strftime('%d/%m/%Y')
                print(f"✓ Ngày báo cáo: {report_date}")
            except Exception as e:
                report_date = datetime.now().strftime('%d/%m/%Y')
                print(f"⚠️ Không parse được ngày ({e}), dùng ngày hiện tại: {report_date}")
        else:
            report_date = datetime.now().strftime('%d/%m/%Y')
        
        # Tra cứu danhba.db
        print(f"\n✓ Đang đọc {db_file}...")
        if not os.path.exists(db_file):
            print(f"❌ Không tìm thấy {db_file}")
            return False
        
        conn = sqlite3.connect(db_file)
        
        # Đọc danhba với TẤT CẢ các cột cần thiết, CHỈ LẤY FIBER
        df_danhba = pd.read_sql_query("""
            SELECT SYSTEM, THIETBI, CARD, PORT, ONT_ID, SA, DOI_VT, NVKT, KETCUOI, 
                   MA_TB, TEN_TB, DIACHI_TB, SO_DT, TRANGTHAI_TB,
                   TRANGTHAI_PORT, MA_LT, MATB_TN, MADOICAP, VATTU_THUEBAO,
                   HDTB_ID, TRANGTHAI_HD, TEN_DVVT, LOAIHINH_TB, TOCDO, TOCDOTHUC,
                   KHACH_HANG, LOAI, CAPGOC
            FROM danhba
            WHERE LOAIHINH_TB = 'Fiber'
        """, conn)
        print(f"✅ Đã đọc {len(df_danhba)} bản ghi từ danhba.db (chỉ Fiber)")
        conn.close()
        
        # Tạo DataFrame output
        print("\n✓ Đang match dữ liệu theo SYSTEM + CARD + PORT + ONT_ID...")
        print("  (Chỉ lấy SYSTEM/CARD/PORT/ONT_ID/Ngày từ CTS, còn lại từ danhba.db)")
        
        output_rows = []
        matched = 0
        unmatched = 0
        
        for idx, row in df_cts.iterrows():
            thiet_bi = row.get('Thiết bị')  # = SYSTEM trong danhba
            cong = row.get('Cổng')
            parsed = parse_cong_column(cong)
            
            # Tra cứu danhba.db bằng SYSTEM + CARD + PORT + ONT_ID
            if thiet_bi and parsed and parsed['ont_id'] is not None:
                card_db = parsed['slot']  # slot trong CTS → CARD trong danhba.db
                port = parsed['port']
                ont_id = parsed['ont_id']
                
                # Match: SYSTEM = Thiết bị, CARD = slot, PORT = port, ONT_ID = ont_id
                mask = (
                    (df_danhba['SYSTEM'] == thiet_bi) &
                    (df_danhba['CARD'] == card_db) &
                    (df_danhba['PORT'] == port) &
                    (df_danhba['ONT_ID'] == ont_id)
                )
                
                matches = df_danhba[mask]
                
                if len(matches) > 0:
                    match = matches.iloc[0]
                    
                    # Chỉ lấy từ CTS: SYSTEM, CARD, PORT, ONT_ID (qua Cổng) và Ngày suy hao
                    # Tất cả thông tin khác lấy từ danhba.db
                    out_row = {
                        # Từ CTS (chỉ key matching và ngày)
                        'NGAY_SUYHAO': report_date,
                        'OLT_CTS': thiet_bi,
                        'PORT_CTS': cong,
                        
                        # Từ danhba.db (tất cả thông tin thuê bao)
                        'THIETBI': match['THIETBI'],
                        'SYSTEM': match['SYSTEM'],
                        'CARD': match['CARD'],
                        'PORT': match['PORT'],
                        'ONT_ID': match['ONT_ID'],
                        'SA': match['SA'],
                        'DOI_ONE': match['DOI_VT'],
                        'NVKT_DB': match['NVKT'],
                        'KETCUOI': match['KETCUOI'],
                        'MA_TB': match['MA_TB'],
                        'TEN_TB_ONE': match['TEN_TB'],
                        'DIACHI_ONE': match['DIACHI_TB'],
                        'DT_ONE': match['SO_DT'],
                        'TRANGTHAI_TB': match['TRANGTHAI_TB'],
                        'TRANGTHAI_PORT': match['TRANGTHAI_PORT'],
                        'MA_LT': match['MA_LT'],
                        'MATB_TN': match['MATB_TN'],
                        'MADOICAP': match['MADOICAP'],
                        'VATTU_THUEBAO': match['VATTU_THUEBAO'],
                        'HDTB_ID': match['HDTB_ID'],
                        'TRANGTHAI_HD': match['TRANGTHAI_HD'],
                        'TEN_DVVT': match['TEN_DVVT'],
                        'LOAIHINH_TB': match['LOAIHINH_TB'],
                        'TOCDO': match['TOCDO'],
                        'TOCDOTHUC': match['TOCDOTHUC'],
                        'KHACH_HANG': match['KHACH_HANG'],
                        'LOAI': match['LOAI'],
                        'CAPGOC': match['CAPGOC'],
                    }
                    
                    matched += 1
                    output_rows.append(out_row)
                else:
                    unmatched += 1
            else:
                unmatched += 1
        
        print(f"\n✓ Kết quả tra cứu danhba.db:")
        print(f"  ✅ Khớp (TTVT Sơn Tây): {matched} thuê bao")
        print(f"  ❌ Không khớp (các TT khác): {unmatched} thuê bao")
        
        if matched == 0:
            print("⚠️ Không có bản ghi nào khớp, không tạo file output")
            return False
        
        # Tạo DataFrame output
        df_output = pd.DataFrame(output_rows)
        
        # Thêm cột TT
        df_output.insert(0, 'TT', range(1, len(df_output) + 1))
        
        # Sắp xếp lại các cột theo thứ tự của I1.5 report
        column_order = [
            'TT', 'NGAY_SUYHAO', 'OLT_CTS', 'PORT_CTS',
            'MA_TB', 'TEN_TB_ONE', 'DT_ONE', 'DIACHI_ONE', 
            'TRANGTHAI_TB', 'TRANGTHAI_HD',
            'DOI_ONE', 'NVKT_DB', 'THIETBI', 'SA', 'KETCUOI',
            'SYSTEM', 'CARD', 'PORT', 'ONT_ID',
            'TRANGTHAI_PORT', 'MA_LT', 'MATB_TN', 'MADOICAP', 
            'VATTU_THUEBAO', 'HDTB_ID', 'TEN_DVVT', 
            'LOAIHINH_TB', 'TOCDO', 'TOCDOTHUC',
            'KHACH_HANG', 'LOAI', 'CAPGOC'
        ]
        
        # Chỉ giữ các cột tồn tại
        existing_cols = [c for c in column_order if c in df_output.columns]
        other_cols = [c for c in df_output.columns if c not in column_order]
        df_output = df_output[existing_cols + other_cols]
        
        # Ghi file output
        print(f"\n✓ Đang ghi file: {output_file}")
        df_output.to_excel(output_file, index=False, sheet_name='Sheet1')
        print(f"✅ Đã ghi {len(df_output)} dòng vào {output_file}")
        
        print("\n" + "="*80)
        print("✅ HOÀN THÀNH CHUYỂN ĐỔI CTS → I1.5 REPORT")
        print("="*80)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Lỗi khi chuyển đổi: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_cts_fallback(force_update=False):
    """
    Hàm wrapper để sử dụng như fallback trong workflow chính.
    
    1. Chuyển đổi suy_hao_cts.xlsx → I1.5 report.xlsx
    2. Gọi process_I15_report_with_tracking() để xử lý tiếp
    
    Args:
        force_update: Ghi đè dữ liệu đã tồn tại
    
    Returns:
        bool: True nếu thành công
    """
    try:
        print("\n" + "="*80)
        print("FALLBACK: XỬ LÝ TỪ SỮY HAO CTS")
        print("="*80)
        
        # Bước 1: Chuyển đổi CTS → I1.5
        success = convert_cts_to_i15_report(force_update=True)
        if not success:
            print("❌ Không thể chuyển đổi CTS → I1.5")
            return False
        
        # Bước 2: Gọi i15_process để xử lý tiếp
        print("\n✓ Đang gọi process_I15_report_with_tracking()...")
        from i15_process import process_I15_report_with_tracking
        success = process_I15_report_with_tracking(force_update=force_update)
        
        if success:
            print("\n✅ FALLBACK HOÀN THÀNH THÀNH CÔNG")
        else:
            print("\n❌ FALLBACK THẤT BẠI")
        
        return success
        
    except Exception as e:
        print(f"\n❌ Lỗi trong fallback: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Chuyển đổi suy_hao_cts.xlsx → I1.5 report.xlsx")
    parser.add_argument("--convert-only", action="store_true", 
                       help="Chỉ chuyển đổi, không xử lý tiếp")
    parser.add_argument("--force", action="store_true", 
                       help="Ghi đè file/dữ liệu đã tồn tại")
    
    args = parser.parse_args()
    
    if args.convert_only:
        convert_cts_to_i15_report(force_update=args.force)
    else:
        process_cts_fallback(force_update=args.force)
