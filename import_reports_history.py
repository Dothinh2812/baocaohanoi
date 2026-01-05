# -*- coding: utf-8 -*-
"""
import_reports_history.py
Module import dữ liệu từ các file Excel đã xử lý vào database reports_history.db
"""

import sqlite3
import os
import re
import json
import time
from datetime import datetime
import pandas as pd
from init_reports_history_db import (
    get_connection, check_date_exists, log_import, DB_PATH
)


class ReportsHistoryImporter:
    """Class để import dữ liệu từ Excel vào database"""

    def __init__(self):
        self.download_dir = os.path.join(os.path.dirname(__file__), "downloads", "baocao_hanoi")
        self.dsnv_file = os.path.join(os.path.dirname(__file__), "dsnv.xlsx")
        self._dsnv_dict = None

    def _get_dsnv_dict(self):
        """Load và cache dictionary từ dsnv.xlsx"""
        if self._dsnv_dict is None:
            try:
                df = pd.read_excel(self.dsnv_file)
                df.columns = df.columns.str.strip()
                # Tìm cột họ tên và đơn vị
                ho_ten_col = None
                don_vi_col = None
                for col in df.columns:
                    col_lower = col.lower()
                    if 'họ tên' in col_lower or 'ho ten' in col_lower:
                        ho_ten_col = col
                    if 'đơn vị' in col_lower or 'don vi' in col_lower:
                        don_vi_col = col
                if ho_ten_col and don_vi_col:
                    self._dsnv_dict = dict(zip(
                        df[ho_ten_col].str.strip().str.lower(),
                        df[don_vi_col]
                    ))
            except Exception as e:
                print(f"   Lỗi đọc dsnv.xlsx: {e}")
                self._dsnv_dict = {}
        return self._dsnv_dict

    def _lookup_don_vi(self, nvkt: str) -> str:
        """Tra cứu đơn vị từ tên NVKT"""
        if pd.isna(nvkt) or nvkt == '':
            return None
        dsnv = self._get_dsnv_dict()
        return dsnv.get(nvkt.lower().strip())

    def _normalize_nvkt(self, name: str) -> str:
        """Chuẩn hóa tên NVKT"""
        if pd.isna(name) or name == '':
            return ''
        name = str(name).strip()
        name = re.sub(r'\([^)]*\)', '', name)
        if '-' in name:
            name = name.split('-')[-1]
        return name.strip()

    def _parse_percentage(self, value) -> float:
        """Chuyển đổi giá trị percentage (có thể có ký tự %) sang float"""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        # Nếu là string, loại bỏ ký tự % và convert
        value_str = str(value).strip().replace('%', '').replace(',', '.')
        try:
            return float(value_str)
        except ValueError:
            return 0.0

    def _safe_int(self, value) -> int:
        """Chuyển đổi an toàn sang int"""
        if pd.isna(value):
            return 0
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return 0

    def _extract_date_from_filename(self, filename: str) -> str:
        """Trích xuất ngày từ tên file (định dạng DDMMYYYY)"""
        match = re.search(r'(\d{8})', filename)
        if match:
            date_str = match.group(1)
            # Convert DDMMYYYY to YYYY-MM-DD
            return f"{date_str[4:8]}-{date_str[2:4]}-{date_str[0:2]}"
        return datetime.now().strftime('%Y-%m-%d')

    # ========================================
    # IMPORT GROWTH REPORTS (Thực tăng, Hoàn công, Ngưng PSC)
    # ========================================

    def import_growth_pttb(self, ngay_bao_cao: str = None, force: bool = False):
        """
        Import báo cáo Thực tăng PTTB

        Args:
            ngay_bao_cao: Ngày báo cáo (YYYY-MM-DD), mặc định là hôm nay
            force: True để ghi đè dữ liệu đã có
        """
        report_code = 'THUC_TANG_PTTB'
        loai_dich_vu = 'PTTB'
        start_time = time.time()

        print(f"\n{'=' * 60}")
        print(f"IMPORT BÁO CÁO THỰC TĂNG PTTB")
        print(f"{'=' * 60}")

        # Xác định ngày
        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        date_ddmmyyyy = datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%d%m%Y')

        # Kiểm tra đã import chưa
        if not force and check_date_exists(report_code, ngay_bao_cao):
            print(f"   Ngày {ngay_bao_cao} đã được import. Bỏ qua.")
            return

        # Tìm file thực tăng
        thuc_tang_file = os.path.join(self.download_dir, f"thuc_tang_{date_ddmmyyyy}.xlsx")
        if not os.path.exists(thuc_tang_file):
            print(f"   Không tìm thấy file: {thuc_tang_file}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message='File not found')
            return

        print(f"   Đang đọc file: {thuc_tang_file}")

        try:
            conn = get_connection()
            cursor = conn.cursor()

            # Đọc sheet thuc_tang_theo_to
            df_to = pd.read_excel(thuc_tang_file, sheet_name='thuc_tang_theo_to')
            df_to = df_to[df_to['Đơn vị'] != 'TỔNG CỘNG']  # Loại bỏ dòng tổng

            # Import vào growth_snapshots_donvi
            print("   Đang import growth_snapshots_donvi...")
            for _, row in df_to.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO growth_snapshots_donvi
                    (ngay_bao_cao, loai_dich_vu, don_vi, so_hoan_cong, so_ngung_psc, thuc_tang, ty_le_ngung_psc)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    ngay_bao_cao,
                    loai_dich_vu,
                    row.get('Đơn vị', ''),
                    self._safe_int(row.get('Hoàn công', 0)),
                    self._safe_int(row.get('Ngưng PSC', 0)),
                    self._safe_int(row.get('Thực tăng', 0)),
                    self._parse_percentage(row.get('Tỷ lệ ngưng/psc', 0))
                ))

            # Đọc sheet thuc_tang_theo_NVKT
            df_nvkt = pd.read_excel(thuc_tang_file, sheet_name='thuc_tang_theo_NVKT')
            df_nvkt = df_nvkt[df_nvkt['Đơn vị'] != 'TỔNG CỘNG']

            # Import vào growth_snapshots_nvkt
            print("   Đang import growth_snapshots_nvkt...")
            for _, row in df_nvkt.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO growth_snapshots_nvkt
                    (ngay_bao_cao, loai_dich_vu, nvkt, don_vi, so_hoan_cong, so_ngung_psc, thuc_tang, ty_le_ngung_psc)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    ngay_bao_cao,
                    loai_dich_vu,
                    row.get('NVKT', ''),
                    row.get('Đơn vị', ''),
                    self._safe_int(row.get('Hoàn công', 0)),
                    self._safe_int(row.get('Ngưng PSC', 0)),
                    self._safe_int(row.get('Thực tăng', 0)),
                    self._parse_percentage(row.get('Tỷ lệ ngưng/psc', 0))
                ))

            conn.commit()
            conn.close()

            # Log import
            duration_ms = int((time.time() - start_time) * 1000)
            log_import(
                report_code, ngay_bao_cao,
                file_name=os.path.basename(thuc_tang_file),
                file_path=thuc_tang_file,
                record_count=len(df_to) + len(df_nvkt),
                duration_ms=duration_ms
            )

            print(f"   Import {len(df_to)} đơn vị, {len(df_nvkt)} NVKT")
            print(f"   Hoàn thành trong {duration_ms}ms")

        except Exception as e:
            print(f"   Lỗi: {e}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))
            import traceback
            traceback.print_exc()

    def import_growth_mytv(self, ngay_bao_cao: str = None, force: bool = False):
        """Import báo cáo Thực tăng MyTV"""
        report_code = 'THUC_TANG_MYTV'
        loai_dich_vu = 'MYTV'
        start_time = time.time()

        print(f"\n{'=' * 60}")
        print(f"IMPORT BÁO CÁO THỰC TĂNG MYTV")
        print(f"{'=' * 60}")

        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        date_ddmmyyyy = datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%d%m%Y')

        if not force and check_date_exists(report_code, ngay_bao_cao):
            print(f"   Ngày {ngay_bao_cao} đã được import. Bỏ qua.")
            return

        thuc_tang_file = os.path.join(self.download_dir, f"mytv_thuc_tang_{date_ddmmyyyy}.xlsx")
        if not os.path.exists(thuc_tang_file):
            print(f"   Không tìm thấy file: {thuc_tang_file}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message='File not found')
            return

        print(f"   Đang đọc file: {thuc_tang_file}")

        try:
            conn = get_connection()
            cursor = conn.cursor()

            # Đọc sheet thuc_tang_theo_to
            df_to = pd.read_excel(thuc_tang_file, sheet_name='thuc_tang_theo_to')
            df_to = df_to[df_to['Đơn vị'] != 'TỔNG CỘNG']

            print("   Đang import growth_snapshots_donvi...")
            for _, row in df_to.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO growth_snapshots_donvi
                    (ngay_bao_cao, loai_dich_vu, don_vi, so_hoan_cong, so_ngung_psc, thuc_tang, ty_le_ngung_psc)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    ngay_bao_cao,
                    loai_dich_vu,
                    row.get('Đơn vị', ''),
                    self._safe_int(row.get('Hoàn công', 0)),
                    self._safe_int(row.get('Ngưng PSC', 0)),
                    self._safe_int(row.get('Thực tăng', 0)),
                    self._parse_percentage(row.get('Tỷ lệ ngưng/psc', 0))
                ))

            # Đọc sheet thuc_tang_theo_NVKT
            df_nvkt = pd.read_excel(thuc_tang_file, sheet_name='thuc_tang_theo_NVKT')
            df_nvkt = df_nvkt[df_nvkt['Đơn vị'] != 'TỔNG CỘNG']

            print("   Đang import growth_snapshots_nvkt...")
            for _, row in df_nvkt.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO growth_snapshots_nvkt
                    (ngay_bao_cao, loai_dich_vu, nvkt, don_vi, so_hoan_cong, so_ngung_psc, thuc_tang, ty_le_ngung_psc)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    ngay_bao_cao,
                    loai_dich_vu,
                    row.get('NVKT', ''),
                    row.get('Đơn vị', ''),
                    self._safe_int(row.get('Hoàn công', 0)),
                    self._safe_int(row.get('Ngưng PSC', 0)),
                    self._safe_int(row.get('Thực tăng', 0)),
                    self._parse_percentage(row.get('Tỷ lệ ngưng/psc', 0))
                ))

            conn.commit()
            conn.close()

            duration_ms = int((time.time() - start_time) * 1000)
            log_import(
                report_code, ngay_bao_cao,
                file_name=os.path.basename(thuc_tang_file),
                file_path=thuc_tang_file,
                record_count=len(df_to) + len(df_nvkt),
                duration_ms=duration_ms
            )

            print(f"   Import {len(df_to)} đơn vị, {len(df_nvkt)} NVKT")
            print(f"   Hoàn thành trong {duration_ms}ms")

        except Exception as e:
            print(f"   Lỗi: {e}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))
            import traceback
            traceback.print_exc()

    # ========================================
    # IMPORT KR REPORTS
    # ========================================

    def import_kr6(self, ngay_bao_cao: str = None, force: bool = False):
        """Import báo cáo KR6"""
        report_code = 'KR6'
        start_time = time.time()

        print(f"\n{'=' * 60}")
        print(f"IMPORT BÁO CÁO KR6")
        print(f"{'=' * 60}")

        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        if not force and check_date_exists(report_code, ngay_bao_cao):
            print(f"   Ngày {ngay_bao_cao} đã được import. Bỏ qua.")
            return

        # Tìm file KR6
        kr6_tonghop_file = os.path.join(self.download_dir, "download_KR6_report_tong_hop_processed.xlsx")
        kr6_nvkt_file = os.path.join(self.download_dir, "download_KR6_report_NVKT_processed.xlsx")

        record_count = 0

        try:
            conn = get_connection()
            cursor = conn.cursor()
            thang = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%m'))
            nam = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%Y'))

            # Import tổng hợp
            if os.path.exists(kr6_tonghop_file):
                print(f"   Đang đọc file: {kr6_tonghop_file}")
                df = pd.read_excel(kr6_tonghop_file)
                df = df[df['Đơn vị'] != 'Tổng']

                for _, row in df.iterrows():
                    cursor.execute('''
                        INSERT OR REPLACE INTO kr_snapshots_tonghop
                        (ngay_bao_cao, thang, nam, loai_bao_cao, don_vi,
                         so_giao, so_hoan_thanh, ty_le_hoan_thanh, ke_hoach_giao, diem_okr, ton_chua_nghiem_thu)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        ngay_bao_cao, thang, nam, report_code,
                        row.get('Đơn vị', ''),
                        self._safe_int(row.get('Giao tháng T', 0)),
                        self._safe_int(row.get('Hoàn thành tháng T', 0)),
                        self._parse_percentage(row.get('Tỷ lệ', 0)),
                        self._safe_int(row.get('Kế hoạch giao', 0)),
                        self._parse_percentage(row.get('Điểm OKR', 0)),
                        self._safe_int(row.get('Tồn chưa nghiệm thu', 0))
                    ))
                record_count += len(df)
                print(f"   Import {len(df)} đơn vị từ tổng hợp")

            # Import NVKT
            if os.path.exists(kr6_nvkt_file):
                print(f"   Đang đọc file: {kr6_nvkt_file}")
                df = pd.read_excel(kr6_nvkt_file, sheet_name='Tất cả')

                for _, row in df.iterrows():
                    cursor.execute('''
                        INSERT OR REPLACE INTO kr_snapshots_nvkt
                        (ngay_bao_cao, thang, nam, loai_bao_cao, nvkt, don_vi,
                         so_giao, so_hoan_thanh, ty_le_hoan_thanh, ke_hoach_giao, diem_okr, ton_chua_nghiem_thu)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        ngay_bao_cao, thang, nam, report_code,
                        row.get('NVKT', ''),
                        row.get('Đơn vị', ''),
                        self._safe_int(row.get('Giao tháng T', 0)),
                        self._safe_int(row.get('Hoàn thành tháng T', 0)),
                        self._parse_percentage(row.get('Tỉ lệ hoàn thành', 0)),
                        self._safe_int(row.get('Kế hoạch giao', 0)),
                        self._parse_percentage(row.get('Điểm OKR', 0)),
                        self._safe_int(row.get('Tồn chưa nghiệm thu', 0))
                    ))
                record_count += len(df)
                print(f"   Import {len(df)} NVKT")

            conn.commit()
            conn.close()

            duration_ms = int((time.time() - start_time) * 1000)
            log_import(report_code, ngay_bao_cao, record_count=record_count, duration_ms=duration_ms)
            print(f"   Hoàn thành trong {duration_ms}ms")

        except Exception as e:
            print(f"   Lỗi: {e}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))
            import traceback
            traceback.print_exc()

    def import_kr7(self, ngay_bao_cao: str = None, force: bool = False):
        """Import báo cáo KR7"""
        report_code = 'KR7'
        start_time = time.time()

        print(f"\n{'=' * 60}")
        print(f"IMPORT BÁO CÁO KR7")
        print(f"{'=' * 60}")

        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        if not force and check_date_exists(report_code, ngay_bao_cao):
            print(f"   Ngày {ngay_bao_cao} đã được import. Bỏ qua.")
            return

        kr7_tonghop_file = os.path.join(self.download_dir, "download_KR7_report_tong_hop_processed.xlsx")
        kr7_nvkt_file = os.path.join(self.download_dir, "download_KR7_report_NVKT_processed.xlsx")

        record_count = 0

        try:
            conn = get_connection()
            cursor = conn.cursor()
            thang = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%m'))
            nam = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%Y'))

            if os.path.exists(kr7_tonghop_file):
                print(f"   Đang đọc file: {kr7_tonghop_file}")
                df = pd.read_excel(kr7_tonghop_file)
                df = df[df['Đơn vị'] != 'Tổng']

                for _, row in df.iterrows():
                    cursor.execute('''
                        INSERT OR REPLACE INTO kr_snapshots_tonghop
                        (ngay_bao_cao, thang, nam, loai_bao_cao, don_vi,
                         so_giao, so_hoan_thanh, ty_le_hoan_thanh, ke_hoach_giao, diem_okr, ton_chua_nghiem_thu)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        ngay_bao_cao, thang, nam, report_code,
                        row.get('Đơn vị', ''),
                        self._safe_int(row.get('Giao tháng T+1', 0)),
                        self._safe_int(row.get('Hoàn thành tháng T+1', 0)),
                        self._parse_percentage(row.get('Tỷ lệ', 0)),
                        self._safe_int(row.get('Kế hoạch giao', 0)),
                        self._parse_percentage(row.get('Điểm OKR', 0)),
                        self._safe_int(row.get('Tồn chưa nghiệm thu', 0))
                    ))
                record_count += len(df)
                print(f"   Import {len(df)} đơn vị từ tổng hợp")

            if os.path.exists(kr7_nvkt_file):
                print(f"   Đang đọc file: {kr7_nvkt_file}")
                df = pd.read_excel(kr7_nvkt_file, sheet_name='Tất cả')

                for _, row in df.iterrows():
                    cursor.execute('''
                        INSERT OR REPLACE INTO kr_snapshots_nvkt
                        (ngay_bao_cao, thang, nam, loai_bao_cao, nvkt, don_vi,
                         so_giao, so_hoan_thanh, ty_le_hoan_thanh, ke_hoach_giao, diem_okr, ton_chua_nghiem_thu)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        ngay_bao_cao, thang, nam, report_code,
                        row.get('NVKT', ''),
                        row.get('Đơn vị', ''),
                        self._safe_int(row.get('Giao tháng T', 0)),
                        self._safe_int(row.get('Hoàn thành tháng T+1', 0)),
                        self._parse_percentage(row.get('Tỉ lệ hoàn thành', 0)),
                        self._safe_int(row.get('Kế hoạch giao', 0)),
                        self._parse_percentage(row.get('Điểm OKR', 0)),
                        self._safe_int(row.get('Tồn chưa nghiệm thu', 0))
                    ))
                record_count += len(df)
                print(f"   Import {len(df)} NVKT")

            conn.commit()
            conn.close()

            duration_ms = int((time.time() - start_time) * 1000)
            log_import(report_code, ngay_bao_cao, record_count=record_count, duration_ms=duration_ms)
            print(f"   Hoàn thành trong {duration_ms}ms")

        except Exception as e:
            print(f"   Lỗi: {e}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))
            import traceback
            traceback.print_exc()

    # ========================================
    # IMPORT VẬT TƯ THU HỒI
    # ========================================

    def import_vat_tu_thu_hoi(self, ngay_bao_cao: str = None, force: bool = False):
        """Import báo cáo Vật tư thu hồi"""
        report_code = 'VAT_TU_THU_HOI'
        start_time = time.time()

        print(f"\n{'=' * 60}")
        print(f"IMPORT BÁO CÁO VẬT TƯ THU HỒI")
        print(f"{'=' * 60}")

        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        if not force and check_date_exists(report_code, ngay_bao_cao):
            print(f"   Ngày {ngay_bao_cao} đã được import. Bỏ qua.")
            return

        vat_tu_file = os.path.join(self.download_dir, "bc_thu_hoi_vat_tu_processed.xlsx")
        if not os.path.exists(vat_tu_file):
            print(f"   Không tìm thấy file: {vat_tu_file}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message='File not found')
            return

        print(f"   Đang đọc file: {vat_tu_file}")

        try:
            conn = get_connection()
            cursor = conn.cursor()

            # Đọc sheet Tổng hợp
            df_tonghop = pd.read_excel(vat_tu_file, sheet_name='Tổng hợp')

            print("   Đang import vat_tu_snapshots...")
            for _, row in df_tonghop.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO vat_tu_snapshots
                    (ngay_bao_cao, don_vi, nvkt, diem_chia, so_luong_chua_thu_hoi)
                    VALUES (?, ?, ?, ?, ?)
                ''', (
                    ngay_bao_cao,
                    None,  # Sẽ cập nhật sau nếu cần
                    row.get('NVKT_DIABAN_GIAO', ''),
                    row.get('DIEMCHIA', ''),
                    self._safe_int(row.get('Số lượng', 0))
                ))

            # Đọc sheet Chi tiết vật tư
            df_chitiet = pd.read_excel(vat_tu_file, sheet_name='Chi tiết vật tư')

            print("   Đang import vat_tu_detail...")
            for _, row in df_chitiet.iterrows():
                cursor.execute('''
                    INSERT INTO vat_tu_detail
                    (ngay_bao_cao, nvkt_diaban_giao, ma_tb, ten_tb, ten_tbi,
                     ngay_giao, ten_loaihd, ten_kieuld, so_dt, ngay_sd_tb, diem_chia, nvkt)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    ngay_bao_cao,
                    row.get('NVKT_DIABAN_GIAO', ''),
                    row.get('MA_TB', ''),
                    row.get('TEN_TB', ''),
                    row.get('TEN_TBI', ''),
                    str(row.get('NGAY_GIAO', '')),
                    row.get('TEN_LOAIHD', ''),
                    row.get('TEN_KIEULD', ''),
                    row.get('SO_DT', ''),
                    str(row.get('NGAY_SD_TB', '')),
                    row.get('DIEMCHIA', ''),
                    row.get('NVKT_DIABAN_GIAO', '')
                ))

            conn.commit()
            conn.close()

            duration_ms = int((time.time() - start_time) * 1000)
            log_import(
                report_code, ngay_bao_cao,
                file_name=os.path.basename(vat_tu_file),
                record_count=len(df_tonghop) + len(df_chitiet),
                duration_ms=duration_ms
            )

            print(f"   Import {len(df_tonghop)} tổng hợp, {len(df_chitiet)} chi tiết")
            print(f"   Hoàn thành trong {duration_ms}ms")

        except Exception as e:
            print(f"   Lỗi: {e}")
            log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))
            import traceback
            traceback.print_exc()

    # ========================================
    # IMPORT C1 REPORTS
    # ========================================

    def import_c1_reports(self, ngay_bao_cao: str = None, force: bool = False):
        """Import tất cả báo cáo C1.1 - C1.5"""
        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        print(f"\n{'=' * 60}")
        print(f"IMPORT BÁO CÁO C1.x")
        print(f"{'=' * 60}")

        thang = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%m'))
        nam = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%Y'))

        # Import từng báo cáo C1
        c1_reports = [
            ('C1.1', 'c1.1 report.xlsx', 'TH_C1.1'),
            ('C1.2', 'c1.2 report.xlsx', 'TH_C1.2'),
            ('C1.3', 'c1.3 report.xlsx', 'TH_C1.3'),
            ('C1.4', 'c1.4 report.xlsx', 'TH_C1.4'),
            ('C1.5', 'c1.5 report.xlsx', 'TH_C1.5'),
        ]

        total_records = 0
        start_time = time.time()

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for report_code, file_name, sheet_name in c1_reports:
                file_path = os.path.join(self.download_dir, file_name)

                if not os.path.exists(file_path):
                    print(f"   ⚠️ Không tìm thấy file: {file_name}")
                    continue

                if not force and check_date_exists(report_code, ngay_bao_cao):
                    print(f"   {report_code}: Đã import, bỏ qua")
                    continue

                try:
                    # Đọc sheet đã xử lý
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    print(f"   Đang import {report_code}...")

                    for _, row in df.iterrows():
                        don_vi = row.get('Đơn vị', '')
                        if pd.isna(don_vi) or don_vi == '':
                            continue

                        # Lấy các giá trị SM và tỷ lệ
                        sm1 = self._safe_int(row.get('SM1', 0))
                        sm2 = self._safe_int(row.get('SM2', 0))
                        sm3 = self._safe_int(row.get('SM3', 0))
                        sm4 = self._safe_int(row.get('SM4', 0))

                        # Tìm cột tỷ lệ chính (cột thứ 4 hoặc cột có chứa 'Tỷ lệ')
                        ty_le_chinh = 0.0
                        ty_le_phu = 0.0
                        for col in df.columns:
                            if 'Tỷ lệ' in col or 'tỷ lệ' in col.lower():
                                val = self._parse_percentage(row.get(col, 0))
                                if ty_le_chinh == 0:
                                    ty_le_chinh = val
                                else:
                                    ty_le_phu = val

                        diem_bsc = self._parse_percentage(row.get('Chỉ tiêu BSC', 0))

                        cursor.execute('''
                            INSERT OR REPLACE INTO c1_snapshots
                            (ngay_bao_cao, thang, nam, loai_bao_cao, don_vi,
                             sm1, sm2, sm3, sm4, ty_le_chinh, ty_le_phu, diem_bsc)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            ngay_bao_cao, thang, nam, report_code, don_vi,
                            sm1, sm2, sm3, sm4, ty_le_chinh, ty_le_phu, diem_bsc
                        ))
                        total_records += 1

                    # Log import cho từng report
                    log_import(report_code, ngay_bao_cao, file_name=file_name, record_count=len(df))
                    print(f"   ✅ {report_code}: {len(df)} đơn vị")

                except Exception as e:
                    print(f"   ❌ Lỗi import {report_code}: {e}")
                    log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))

            conn.commit()
            conn.close()

            duration_ms = int((time.time() - start_time) * 1000)
            print(f"   Tổng: {total_records} bản ghi trong {duration_ms}ms")

        except Exception as e:
            print(f"   Lỗi: {e}")
            import traceback
            traceback.print_exc()

    def import_c1_chitiet(self, ngay_bao_cao: str = None, force: bool = False):
        """Import chi tiết C1.4 và C1.5"""
        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        print(f"\n   Đang import C1 chi tiết...")

        thang = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%m'))
        nam = int(datetime.strptime(ngay_bao_cao, '%Y-%m-%d').strftime('%Y'))

        chitiet_reports = [
            ('C1.4_CHITIET', 'c1.4_chitiet_report.xlsx'),
            ('C1.5_CHITIET', 'c1.5_chitiet_report.xlsx'),
        ]

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for report_code, file_name in chitiet_reports:
                file_path = os.path.join(self.download_dir, file_name)

                if not os.path.exists(file_path):
                    print(f"   ⚠️ Không tìm thấy file chi tiết: {file_name}")
                    continue

                if not force and check_date_exists(report_code, ngay_bao_cao):
                    print(f"   {report_code}: Đã import, bỏ qua")
                    continue

                try:
                    df = pd.read_excel(file_path)
                    print(f"   Đang import {report_code}...")

                    # Tìm cột NVKT
                    nvkt_col = None
                    for col in df.columns:
                        if 'NVKT' in col.upper() or 'TEN_NVKT' in col.upper():
                            nvkt_col = col
                            break

                    if nvkt_col is None:
                        print(f"   ⚠️ Không tìm thấy cột NVKT trong {file_name}")
                        continue

                    # Group by NVKT
                    df_grouped = df.groupby(nvkt_col).size().reset_index(name='tong_phieu')

                    for _, row in df_grouped.iterrows():
                        nvkt = row[nvkt_col]
                        if pd.isna(nvkt) or nvkt == '':
                            continue

                        # Lookup đơn vị
                        don_vi = self._lookup_don_vi(str(nvkt))

                        cursor.execute('''
                            INSERT OR REPLACE INTO c1_snapshots_nvkt
                            (ngay_bao_cao, thang, nam, loai_bao_cao, don_vi, nvkt, tong_phieu)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            ngay_bao_cao, thang, nam, report_code,
                            don_vi, nvkt, int(row['tong_phieu'])
                        ))

                    log_import(report_code, ngay_bao_cao, file_name=file_name, record_count=len(df_grouped))
                    print(f"   ✅ {report_code}: {len(df_grouped)} NVKT")

                except Exception as e:
                    print(f"   ❌ Lỗi import {report_code}: {e}")
                    log_import(report_code, ngay_bao_cao, status='ERROR', error_message=str(e))

            conn.commit()
            conn.close()

        except Exception as e:
            print(f"   Lỗi: {e}")
            import traceback
            traceback.print_exc()

    # ========================================
    # UPDATE SUMMARY TABLES
    # ========================================

    def update_daily_summary(self, ngay_bao_cao: str = None):
        """Cập nhật bảng daily_summary từ các bảng chi tiết"""
        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        print(f"\n   Đang cập nhật daily_summary cho ngày {ngay_bao_cao}...")

        conn = get_connection()
        cursor = conn.cursor()

        # Lấy tổng PTTB
        cursor.execute('''
            SELECT
                COALESCE(SUM(so_hoan_cong), 0) as hc,
                COALESCE(SUM(so_ngung_psc), 0) as np,
                COALESCE(SUM(thuc_tang), 0) as tt
            FROM growth_snapshots_donvi
            WHERE ngay_bao_cao = ? AND loai_dich_vu = 'PTTB'
        ''', (ngay_bao_cao,))
        pttb = cursor.fetchone()

        # Lấy tổng MyTV
        cursor.execute('''
            SELECT
                COALESCE(SUM(so_hoan_cong), 0) as hc,
                COALESCE(SUM(so_ngung_psc), 0) as np,
                COALESCE(SUM(thuc_tang), 0) as tt
            FROM growth_snapshots_donvi
            WHERE ngay_bao_cao = ? AND loai_dich_vu = 'MYTV'
        ''', (ngay_bao_cao,))
        mytv = cursor.fetchone()

        # Lấy tổng vật tư chưa thu hồi
        cursor.execute('''
            SELECT COALESCE(SUM(so_luong_chua_thu_hoi), 0)
            FROM vat_tu_snapshots
            WHERE ngay_bao_cao = ?
        ''', (ngay_bao_cao,))
        vat_tu = cursor.fetchone()[0]

        # Insert/Update daily_summary
        cursor.execute('''
            INSERT OR REPLACE INTO daily_summary
            (ngay_bao_cao, pttb_hoan_cong, pttb_ngung_psc, pttb_thuc_tang,
             mytv_hoan_cong, mytv_ngung_psc, mytv_thuc_tang, vat_tu_chua_thu_hoi, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ''', (
            ngay_bao_cao,
            pttb[0] if pttb else 0,
            pttb[1] if pttb else 0,
            pttb[2] if pttb else 0,
            mytv[0] if mytv else 0,
            mytv[1] if mytv else 0,
            mytv[2] if mytv else 0,
            vat_tu
        ))

        conn.commit()
        conn.close()

        print(f"   PTTB: HC={pttb[0] if pttb else 0}, NP={pttb[1] if pttb else 0}, TT={pttb[2] if pttb else 0}")
        print(f"   MyTV: HC={mytv[0] if mytv else 0}, NP={mytv[1] if mytv else 0}, TT={mytv[2] if mytv else 0}")
        print(f"   Vật tư chưa thu hồi: {vat_tu}")

    # ========================================
    # IMPORT ALL
    # ========================================

    def import_all(self, ngay_bao_cao: str = None, force: bool = False):
        """Import tất cả báo cáo cho một ngày"""
        if ngay_bao_cao is None:
            ngay_bao_cao = datetime.now().strftime('%Y-%m-%d')

        print(f"\n{'#' * 60}")
        print(f"IMPORT TẤT CẢ BÁO CÁO CHO NGÀY {ngay_bao_cao}")
        print(f"{'#' * 60}")

        # Import từng loại báo cáo
        self.import_growth_pttb(ngay_bao_cao, force)
        self.import_growth_mytv(ngay_bao_cao, force)
        self.import_kr6(ngay_bao_cao, force)
        self.import_kr7(ngay_bao_cao, force)
        self.import_vat_tu_thu_hoi(ngay_bao_cao, force)
        self.import_c1_reports(ngay_bao_cao, force)
        self.import_c1_chitiet(ngay_bao_cao, force)

        # Cập nhật summary
        self.update_daily_summary(ngay_bao_cao)

        print(f"\n{'#' * 60}")
        print(f"HOÀN THÀNH IMPORT CHO NGÀY {ngay_bao_cao}")
        print(f"{'#' * 60}")


def import_today():
    """Hàm tiện ích để import dữ liệu hôm nay"""
    from init_reports_history_db import init_database

    # Đảm bảo database đã được khởi tạo
    if not os.path.exists(DB_PATH):
        init_database()

    importer = ReportsHistoryImporter()
    importer.import_all()


if __name__ == '__main__':
    import_today()
