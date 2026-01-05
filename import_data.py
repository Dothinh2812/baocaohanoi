#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script import dữ liệu từ Excel vào SQLite database
"""

import pandas as pd
import sqlite3
import os
import re
from datetime import datetime
from pathlib import Path
import glob

class BaoCaoImporter:
    def __init__(self, db_path='baocao_hanoi.db'):
        self.db_path = db_path
        self.conn = None
        self.cursor = None

    def connect(self):
        """Kết nối database"""
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        print(f"✓ Đã kết nối database: {self.db_path}")

    def close(self):
        """Đóng kết nối"""
        if self.conn:
            self.conn.close()
            print("✓ Đã đóng kết nối database")

    def create_schema(self):
        """Tạo schema database"""
        schema_file = 'database_schema.sql'
        if not os.path.exists(schema_file):
            print(f"✗ Không tìm thấy file schema: {schema_file}")
            return False

        with open(schema_file, 'r', encoding='utf-8') as f:
            schema_sql = f.read()

        # Thực thi từng câu lệnh SQL
        for statement in schema_sql.split(';'):
            if statement.strip():
                try:
                    self.cursor.execute(statement)
                except Exception as e:
                    print(f"Lỗi khi thực thi SQL: {e}")
                    print(f"Statement: {statement[:100]}...")

        self.conn.commit()
        print("✓ Đã tạo schema database")
        return True

    def extract_date_from_filename(self, filename):
        """Trích xuất ngày từ tên file"""
        # Pattern: ddmmyyyy
        pattern = r'(\d{2})(\d{2})(\d{4})'
        match = re.search(pattern, filename)
        if match:
            day, month, year = match.groups()
            try:
                date_str = f"{year}-{month}-{day}"
                return date_str
            except:
                pass

        # Pattern: dd/mm/yyyy hoặc dd-mm-yyyy
        pattern2 = r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})'
        match = re.search(pattern2, filename)
        if match:
            day, month, year = match.groups()
            try:
                date_str = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                return date_str
            except:
                pass

        return None

    def parse_datetime(self, value):
        """Parse datetime từ nhiều định dạng khác nhau"""
        if pd.isna(value):
            return None

        if isinstance(value, str):
            # Thử nhiều định dạng
            formats = [
                '%m/%d/%Y %I:%M:%S %p',  # 11/13/2025 8:09:24 AM
                '%d/%m/%Y %H:%M:%S',      # 13/11/2025 08:09:24
                '%d/%m/%Y %H:%M',         # 13/11/2025 08:09
                '%d/%m/%Y',               # 13/11/2025
                '%Y-%m-%d %H:%M:%S',      # 2025-11-13 08:09:24
                '%Y-%m-%d',               # 2025-11-13
            ]

            for fmt in formats:
                try:
                    return datetime.strptime(value, fmt).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    continue

        return str(value)

    def import_hoan_cong(self, file_path, ngay_bao_cao, loai_dv='FIBER'):
        """Import dữ liệu hoàn công"""
        try:
            df = pd.read_excel(file_path, sheet_name='Data')

            records = []
            for _, row in df.iterrows():
                record = {
                    'ngay_bao_cao': ngay_bao_cao,
                    'loai_dv': loai_dv,
                    'stt': int(row.get('STT', 0)) if pd.notna(row.get('STT')) else None,
                    'ma_tb': str(row.get('Mã TB', '')),
                    'ngay_nghiem_thu': self.parse_datetime(row.get('Ngày nghiệm thu')),
                    'ngay_yeu_cau': self.parse_datetime(row.get('Ngày yêu cầu')),
                    'doi': str(row.get('Đội', '')) if pd.notna(row.get('Đội')) else None,
                    'nhom_dia_ban': str(row.get('Nhóm địa bàn', '')),
                    'ten_ttvt': str(row.get('Tên TTVT', '')),
                    'trang_thai_phieu': str(row.get('Trạng thái phiếu', '')),
                    'hdtb_id': int(row.get('HDTB_ID', 0)) if pd.notna(row.get('HDTB_ID')) else None,
                    'nhan_vien_kt': str(row.get('Nhân viên KT', '')),
                    'ma_gd': str(row.get('Mã GD', '')),
                    'nvkt': str(row.get('NVKT', '')),
                    'don_vi': str(row.get('Đơn vị', '')),
                }
                records.append(record)

            # Insert vào database
            sql = """
                INSERT OR IGNORE INTO hoan_cong (
                    ngay_bao_cao, loai_dv, stt, ma_tb, ngay_nghiem_thu, ngay_yeu_cau,
                    doi, nhom_dia_ban, ten_ttvt, trang_thai_phieu, hdtb_id,
                    nhan_vien_kt, ma_gd, nvkt, don_vi
                ) VALUES (
                    :ngay_bao_cao, :loai_dv, :stt, :ma_tb, :ngay_nghiem_thu, :ngay_yeu_cau,
                    :doi, :nhom_dia_ban, :ten_ttvt, :trang_thai_phieu, :hdtb_id,
                    :nhan_vien_kt, :ma_gd, :nvkt, :don_vi
                )
            """

            self.cursor.executemany(sql, records)
            self.conn.commit()

            return len(records)

        except Exception as e:
            print(f"✗ Lỗi import hoàn công: {e}")
            return 0

    def import_ngung_psc(self, file_path, ngay_bao_cao, loai_dv='FIBER'):
        """Import dữ liệu ngừng PSC"""
        try:
            df = pd.read_excel(file_path, sheet_name='Data')

            records = []
            for _, row in df.iterrows():
                record = {
                    'ngay_bao_cao': ngay_bao_cao,
                    'loai_dv': loai_dv,
                    'stt': int(row.get('STT', 0)) if pd.notna(row.get('STT')) else None,
                    'ma_tb': str(row.get('Mã TB', '')),
                    'so_may': str(row.get('Số máy', '')) if pd.notna(row.get('Số máy')) else None,
                    'ten_tb': str(row.get('Tên TB', '')) if pd.notna(row.get('Tên TB')) else None,
                    'loai_dich_vu': str(row.get('Loại dịch vụ', '')),
                    'dia_chi_ld': str(row.get('Địa chỉ LĐ', '')),
                    'ngay_tam_dung': self.parse_datetime(row.get('Ngày tạm dừng')),
                    'ngay_khoi_phuc': self.parse_datetime(row.get('Ngày khôi phục')),
                    'ngay_huy': self.parse_datetime(row.get('Ngày hủy')),
                    'nhom_dia_ban': str(row.get('Nhóm địa bàn', '')),
                    'ten_to': str(row.get('Tên Tổ', '')) if pd.notna(row.get('Tên Tổ')) else str(row.get('Tên TTVT', '')),
                    'ten_ttvt': str(row.get('Tên TTVT', '')),
                    'ten_kh': str(row.get('Tên KH', '')) if pd.notna(row.get('Tên KH')) else None,
                    'dien_thoai_lh': str(row.get('Điện thoại LH', '')) if pd.notna(row.get('Điện thoại LH')) else None,
                    'trang_thai_tb': str(row.get('Trạng thái TB', '')),
                    'ly_do_huy_tam_dung': str(row.get('Lý do hủy tạm dừng', '')) if pd.notna(row.get('Lý do hủy tạm dừng')) else None,
                    'ttvt_xac_minh_huy': str(row.get('TTVT xác minh huỷ', '')) if pd.notna(row.get('TTVT xác minh huỷ')) else None,
                    'doi_tuong': str(row.get('Đối tượng', '')) if pd.notna(row.get('Đối tượng')) else None,
                    'nvkt': str(row.get('NVKT', '')),
                    'don_vi': str(row.get('Đơn vị', '')),
                }
                records.append(record)

            sql = """
                INSERT OR IGNORE INTO ngung_psc (
                    ngay_bao_cao, loai_dv, stt, ma_tb, so_may, ten_tb, loai_dich_vu,
                    dia_chi_ld, ngay_tam_dung, ngay_khoi_phuc, ngay_huy, nhom_dia_ban,
                    ten_to, ten_ttvt, ten_kh, dien_thoai_lh, trang_thai_tb,
                    ly_do_huy_tam_dung, ttvt_xac_minh_huy, doi_tuong, nvkt, don_vi
                ) VALUES (
                    :ngay_bao_cao, :loai_dv, :stt, :ma_tb, :so_may, :ten_tb, :loai_dich_vu,
                    :dia_chi_ld, :ngay_tam_dung, :ngay_khoi_phuc, :ngay_huy, :nhom_dia_ban,
                    :ten_to, :ten_ttvt, :ten_kh, :dien_thoai_lh, :trang_thai_tb,
                    :ly_do_huy_tam_dung, :ttvt_xac_minh_huy, :doi_tuong, :nvkt, :don_vi
                )
            """

            self.cursor.executemany(sql, records)
            self.conn.commit()

            return len(records)

        except Exception as e:
            print(f"✗ Lỗi import ngừng PSC: {e}")
            return 0

    def import_thuc_tang(self, file_path, ngay_bao_cao, loai_dv='FIBER'):
        """Import dữ liệu thực tăng"""
        try:
            records = []

            # Import thực tăng theo tổ
            try:
                df_to = pd.read_excel(file_path, sheet_name='thuc_tang_theo_to')
                for _, row in df_to.iterrows():
                    record = {
                        'ngay_bao_cao': ngay_bao_cao,
                        'loai_dv': loai_dv,
                        'cap_do': 'to',
                        'don_vi': str(row.get('Đơn vị', '')),
                        'nvkt': None,
                        'hoan_cong': int(row.get('Hoàn công', 0)) if pd.notna(row.get('Hoàn công')) else 0,
                        'ngung_psc': int(row.get('Ngưng PSC', 0)) if pd.notna(row.get('Ngưng PSC')) else 0,
                        'thuc_tang': int(row.get('Thực tăng', 0)) if pd.notna(row.get('Thực tăng')) else 0,
                        'ty_le_ngung_psc': float(row.get('Tỷ lệ ngưng/psc', 0)) if pd.notna(row.get('Tỷ lệ ngưng/psc')) else 0,
                    }
                    records.append(record)
            except:
                pass

            # Import thực tăng theo NVKT
            try:
                df_nvkt = pd.read_excel(file_path, sheet_name='thuc_tang_theo_NVKT')
                for _, row in df_nvkt.iterrows():
                    record = {
                        'ngay_bao_cao': ngay_bao_cao,
                        'loai_dv': loai_dv,
                        'cap_do': 'nvkt',
                        'don_vi': str(row.get('Đơn vị', '')),
                        'nvkt': str(row.get('NVKT', '')),
                        'hoan_cong': int(row.get('Hoàn công', 0)) if pd.notna(row.get('Hoàn công')) else 0,
                        'ngung_psc': int(row.get('Ngưng PSC', 0)) if pd.notna(row.get('Ngưng PSC')) else 0,
                        'thuc_tang': int(row.get('Thực tăng', 0)) if pd.notna(row.get('Thực tăng')) else 0,
                        'ty_le_ngung_psc': float(row.get('Tỷ lệ ngưng/psc', 0)) if pd.notna(row.get('Tỷ lệ ngưng/psc')) else 0,
                    }
                    records.append(record)
            except:
                pass

            if records:
                sql = """
                    INSERT OR REPLACE INTO thuc_tang (
                        ngay_bao_cao, loai_dv, cap_do, don_vi, nvkt,
                        hoan_cong, ngung_psc, thuc_tang, ty_le_ngung_psc
                    ) VALUES (
                        :ngay_bao_cao, :loai_dv, :cap_do, :don_vi, :nvkt,
                        :hoan_cong, :ngung_psc, :thuc_tang, :ty_le_ngung_psc
                    )
                """

                self.cursor.executemany(sql, records)
                self.conn.commit()

            return len(records)

        except Exception as e:
            print(f"✗ Lỗi import thực tăng: {e}")
            return 0

    def import_suy_hao_cao(self, file_path, ngay_bao_cao):
        """Import dữ liệu suy hao cao I1.5"""
        try:
            df = pd.read_excel(file_path, sheet_name='Sheet1')

            records = []
            for _, row in df.iterrows():
                record = {
                    'ngay_bao_cao': ngay_bao_cao,
                    'ttvt_cts': str(row.get('TTVT_CTS', '')) if pd.notna(row.get('TTVT_CTS')) else None,
                    'ttvt_one': str(row.get('TTVT_ONE', '')) if pd.notna(row.get('TTVT_ONE')) else None,
                    'doi_one': str(row.get('DOI_ONE', '')) if pd.notna(row.get('DOI_ONE')) else None,
                    'ten_kv': str(row.get('TEN_KV', '')) if pd.notna(row.get('TEN_KV')) else None,
                    'olt_cts': str(row.get('OLT_CTS', '')) if pd.notna(row.get('OLT_CTS')) else None,
                    'port_cts': str(row.get('PORT_CTS', '')) if pd.notna(row.get('PORT_CTS')) else None,
                    'account_cts': str(row.get('ACCOUNT_CTS', '')),
                    'ten_tb_one': str(row.get('TEN_TB_ONE', '')) if pd.notna(row.get('TEN_TB_ONE')) else None,
                    'dt_one': str(row.get('DT_ONE', '')) if pd.notna(row.get('DT_ONE')) else None,
                    'diachi_one': str(row.get('DIACHI_ONE', '')) if pd.notna(row.get('DIACHI_ONE')) else None,
                    'ngay_suyhao': self.parse_datetime(row.get('NGAY_SUYHAO')),
                    'trangthai_tb': str(row.get('TRANGTHAI_TB', '')) if pd.notna(row.get('TRANGTHAI_TB')) else None,
                    'ma_module_quang_olt': str(row.get('Mã module quang OLT', '')) if pd.notna(row.get('Mã module quang OLT')) else None,
                    'chi_so_olt_rx': float(row.get('Chỉ số OLT RX', 0)) if pd.notna(row.get('Chỉ số OLT RX')) else None,
                    'chi_so_onu_rx': float(row.get('Chỉ số ONU RX', 0)) if pd.notna(row.get('Chỉ số ONU RX')) else None,
                    'nvkt_db_normalized': str(row.get('NVKT_DB_NORMALIZED', '')) if pd.notna(row.get('NVKT_DB_NORMALIZED')) else None,
                }
                records.append(record)

            sql = """
                INSERT OR REPLACE INTO suy_hao_cao (
                    ngay_bao_cao, ttvt_cts, ttvt_one, doi_one, ten_kv, olt_cts, port_cts,
                    account_cts, ten_tb_one, dt_one, diachi_one, ngay_suyhao, trangthai_tb,
                    ma_module_quang_olt, chi_so_olt_rx, chi_so_onu_rx, nvkt_db_normalized
                ) VALUES (
                    :ngay_bao_cao, :ttvt_cts, :ttvt_one, :doi_one, :ten_kv, :olt_cts, :port_cts,
                    :account_cts, :ten_tb_one, :dt_one, :diachi_one, :ngay_suyhao, :trangthai_tb,
                    :ma_module_quang_olt, :chi_so_olt_rx, :chi_so_onu_rx, :nvkt_db_normalized
                )
            """

            self.cursor.executemany(sql, records)
            self.conn.commit()

            return len(records)

        except Exception as e:
            print(f"✗ Lỗi import suy hao cao: {e}")
            return 0

    def log_import(self, file_name, file_path, loai_bao_cao, ngay_bao_cao, so_ban_ghi, trang_thai, thong_bao=''):
        """Ghi log import"""
        sql = """
            INSERT INTO import_log (
                file_name, file_path, loai_bao_cao, ngay_bao_cao,
                so_ban_ghi, trang_thai, thong_bao
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        self.cursor.execute(sql, (file_name, file_path, loai_bao_cao, ngay_bao_cao, so_ban_ghi, trang_thai, thong_bao))
        self.conn.commit()

    def import_all_files(self, directory='downloads/baocao_hanoi'):
        """Import tất cả các file Excel"""
        base_dir = Path(directory)

        # Tìm tất cả file Excel
        excel_files = list(base_dir.glob('*.xlsx')) + list(base_dir.glob('*.xls'))

        print(f"\n{'='*80}")
        print(f"Tìm thấy {len(excel_files)} file Excel")
        print(f"{'='*80}\n")

        for file_path in excel_files:
            file_name = file_path.name

            # Bỏ qua các file đã xử lý (processed)
            if '_processed' in file_name:
                continue

            print(f"\nĐang xử lý: {file_name}")

            # Trích xuất ngày từ tên file
            ngay_bao_cao = self.extract_date_from_filename(file_name)
            if not ngay_bao_cao:
                ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')
                print(f"  ⚠ Không tìm thấy ngày trong tên file, dùng ngày hiện tại: {ngay_bao_cao}")

            try:
                # Xác định loại báo cáo và import
                if 'hoan_cong' in file_name.lower():
                    loai_dv = 'MYTV' if 'mytv' in file_name.lower() else 'FIBER'
                    so_ban_ghi = self.import_hoan_cong(str(file_path), ngay_bao_cao, loai_dv)
                    loai_bao_cao = f'hoan_cong_{loai_dv}'
                    print(f"  ✓ Import {so_ban_ghi} bản ghi hoàn công {loai_dv}")

                elif 'ngung_psc' in file_name.lower():
                    loai_dv = 'MYTV' if 'mytv' in file_name.lower() else 'FIBER'
                    so_ban_ghi = self.import_ngung_psc(str(file_path), ngay_bao_cao, loai_dv)
                    loai_bao_cao = f'ngung_psc_{loai_dv}'
                    print(f"  ✓ Import {so_ban_ghi} bản ghi ngừng PSC {loai_dv}")

                elif 'thuc_tang' in file_name.lower():
                    loai_dv = 'MYTV' if 'mytv' in file_name.lower() else 'FIBER'
                    so_ban_ghi = self.import_thuc_tang(str(file_path), ngay_bao_cao, loai_dv)
                    loai_bao_cao = f'thuc_tang_{loai_dv}'
                    print(f"  ✓ Import {so_ban_ghi} bản ghi thực tăng {loai_dv}")

                elif 'i1.5' in file_name.lower() or 'I1.5' in file_name:
                    so_ban_ghi = self.import_suy_hao_cao(str(file_path), ngay_bao_cao)
                    loai_bao_cao = 'suy_hao_cao'
                    print(f"  ✓ Import {so_ban_ghi} bản ghi suy hao cao")

                else:
                    print(f"  ⊘ Bỏ qua (chưa hỗ trợ loại file này)")
                    continue

                # Log import thành công
                self.log_import(file_name, str(file_path), loai_bao_cao, ngay_bao_cao, so_ban_ghi, 'thanh_cong', '')

            except Exception as e:
                print(f"  ✗ Lỗi: {e}")
                self.log_import(file_name, str(file_path), 'unknown', ngay_bao_cao, 0, 'loi', str(e))


def main():
    """Hàm chính"""
    print("\n" + "="*80)
    print("IMPORT DỮ LIỆU BÁO CÁO HÀ NỘI VÀO DATABASE")
    print("="*80 + "\n")

    importer = BaoCaoImporter('baocao_hanoi.db')

    try:
        # Kết nối database
        importer.connect()

        # Tạo schema
        print("\nBước 1: Tạo schema database...")
        importer.create_schema()

        # Import dữ liệu
        print("\nBước 2: Import dữ liệu từ các file Excel...")
        importer.import_all_files()

        print("\n" + "="*80)
        print("HOÀN THÀNH IMPORT DỮ LIỆU")
        print("="*80 + "\n")

        # Thống kê
        cursor = importer.cursor
        cursor.execute("SELECT COUNT(*) FROM hoan_cong")
        count_hc = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM ngung_psc")
        count_np = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM thuc_tang")
        count_tt = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM suy_hao_cao")
        count_shc = cursor.fetchone()[0]

        print(f"Tổng kết:")
        print(f"  - Hoàn công: {count_hc} bản ghi")
        print(f"  - Ngừng PSC: {count_np} bản ghi")
        print(f"  - Thực tăng: {count_tt} bản ghi")
        print(f"  - Suy hao cao: {count_shc} bản ghi")
        print()

    finally:
        importer.close()


if __name__ == '__main__':
    main()
