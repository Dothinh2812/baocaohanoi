# -*- coding: utf-8 -*-
"""
Module import dữ liệu C1.x vào database c1_history.db
"""

import pandas as pd
import os
import sqlite3
from datetime import datetime, date
from c1_history_db import get_connection, DB_PATH


class C1HistoryImporter:
    """Class để import dữ liệu C1 vào database"""

    def __init__(self):
        self.download_dir = os.path.join(os.path.dirname(__file__), "downloads", "baocao_hanoi")
        self.today = date.today().isoformat()

    def _parse_percentage(self, value) -> float:
        """Chuyển đổi giá trị percentage (có thể có ký tự %) sang float"""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
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

    def _safe_str(self, value) -> str:
        """Chuyển đổi an toàn sang string"""
        if pd.isna(value):
            return ""
        return str(value).strip()

    def _log_import(self, conn, loai_du_lieu: str, so_ban_ghi: int, trang_thai: str = "success", ghi_chu: str = None):
        """Ghi log import"""
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO c1_import_log (loai_du_lieu, so_ban_ghi, trang_thai, ghi_chu)
            VALUES (?, ?, ?, ?)
        ''', (loai_du_lieu, so_ban_ghi, trang_thai, ghi_chu))

    def import_c1_tong_hop(self):
        """
        Import dữ liệu tổng hợp từ các file C1.1-C1.5 vào bảng c1_tong_hop và c1_theo_to
        """
        print("📊 Đang import dữ liệu C1 tổng hợp...")

        # Đọc tất cả các file C1
        c1_data = {}
        for i in ['1', '2', '3', '4', '5']:
            filename = os.path.join(self.download_dir, f"c1.{i} report.xlsx")
            sheet_name = f"TH_C1.{i}"

            if os.path.exists(filename):
                try:
                    xl = pd.ExcelFile(filename)
                    if sheet_name in xl.sheet_names:
                        df = pd.read_excel(xl, sheet_name=sheet_name)
                        c1_data[f"c1{i}"] = df
                        print(f"   ✓ Đọc được C1.{i}: {len(df)} dòng")
                    else:
                        print(f"   ⚠️ Không tìm thấy sheet {sheet_name} trong {filename}")
                except Exception as e:
                    print(f"   ❌ Lỗi đọc C1.{i}: {e}")
            else:
                print(f"   ⚠️ Không tìm thấy file: {filename}")

        if not c1_data:
            print("   ❌ Không có dữ liệu C1 để import")
            return False

        conn = get_connection()
        cursor = conn.cursor()

        try:
            # Xóa dữ liệu cũ của ngày hôm nay
            cursor.execute("DELETE FROM c1_tong_hop WHERE ngay_cap_nhat = ?", (self.today,))
            cursor.execute("DELETE FROM c1_theo_to WHERE ngay_cap_nhat = ?", (self.today,))

            # Lấy danh sách các tổ (từ C1.1 làm chuẩn)
            if 'c11' in c1_data:
                df_base = c1_data['c11']
                don_vi_col = 'Đơn vị'

                for _, row in df_base.iterrows():
                    ten_to = self._safe_str(row.get(don_vi_col, ''))
                    if not ten_to:
                        continue

                    # Tạo dict để lưu tất cả các chỉ tiêu
                    data = {
                        'ngay_cap_nhat': self.today,
                        'ten_to': ten_to,
                    }

                    # C1.1
                    if 'c11' in c1_data:
                        df = c1_data['c11']
                        row_data = df[df[don_vi_col] == ten_to]
                        if len(row_data) > 0:
                            r = row_data.iloc[0]
                            data['c11_sm1'] = self._safe_int(r.get('SM1', 0))
                            data['c11_sm2'] = self._safe_int(r.get('SM2', 0))
                            # Tìm cột tỷ lệ sửa chữa CLCD
                            for col in df.columns:
                                if 'sửa chữa' in col.lower() and 'clcd' in col.lower():
                                    data['c11_ty_le_sua_chua_clcd'] = self._parse_percentage(r.get(col, 0))
                                    break
                                elif 'chất lượng chủ động' in col.lower():
                                    data['c11_ty_le_sua_chua_clcd'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c11_sm3'] = self._safe_int(r.get('SM3', 0))
                            data['c11_sm4'] = self._safe_int(r.get('SM4', 0))
                            # Tìm cột tỷ lệ BRCD
                            for col in df.columns:
                                if 'brcd' in col.lower() and 'tỷ lệ' in col.lower():
                                    data['c11_ty_le_brcd'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c11_bsc'] = self._parse_percentage(r.get('Chỉ tiêu BSC', 0))

                    # C1.2
                    if 'c12' in c1_data:
                        df = c1_data['c12']
                        row_data = df[df[don_vi_col] == ten_to]
                        if len(row_data) > 0:
                            r = row_data.iloc[0]
                            data['c12_sm1'] = self._safe_int(r.get('SM1', 0))
                            data['c12_sm2'] = self._safe_int(r.get('SM2', 0))
                            for col in df.columns:
                                if 'lặp lại' in col.lower():
                                    data['c12_ty_le_bao_hong_lap_lai'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c12_sm3'] = self._safe_int(r.get('SM3', 0))
                            data['c12_sm4'] = self._safe_int(r.get('SM4', 0))
                            for col in df.columns:
                                if 'sự cố' in col.lower() and 'brcđ' in col.lower():
                                    data['c12_ty_le_su_co'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c12_bsc'] = self._parse_percentage(r.get('Chỉ tiêu BSC', 0))

                    # C1.3
                    if 'c13' in c1_data:
                        df = c1_data['c13']
                        row_data = df[df[don_vi_col] == ten_to]
                        if len(row_data) > 0:
                            r = row_data.iloc[0]
                            data['c13_sm1'] = self._safe_int(r.get('SM1', 0))
                            data['c13_sm2'] = self._safe_int(r.get('SM2', 0))
                            for col in df.columns:
                                if 'sửa chữa' in col.lower() and 'tsl' in col.lower():
                                    data['c13_ty_le_sua_chua_tsl'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c13_sm3'] = self._safe_int(r.get('SM3', 0))
                            data['c13_sm4'] = self._safe_int(r.get('SM4', 0))
                            for col in df.columns:
                                if 'lặp lại' in col.lower() and 'tsl' in col.lower():
                                    data['c13_ty_le_bao_hong_tsl'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c13_sm5'] = self._safe_int(r.get('SM5', 0))
                            data['c13_sm6'] = self._safe_int(r.get('SM6', 0))
                            for col in df.columns:
                                if 'sự cố' in col.lower() and 'tsl' in col.lower():
                                    data['c13_ty_le_su_co_tsl'] = self._parse_percentage(r.get(col, 0))
                                    break
                            data['c13_bsc'] = self._parse_percentage(r.get('Chỉ tiêu BSC', 0))

                    # C1.4
                    if 'c14' in c1_data:
                        df = c1_data['c14']
                        row_data = df[df[don_vi_col] == ten_to]
                        if len(row_data) > 0:
                            r = row_data.iloc[0]
                            data['c14_tong_phieu'] = self._safe_int(r.get('Tổng phiếu', 0))
                            data['c14_sl_da_ks'] = self._safe_int(r.get('SL đã KS', 0))
                            data['c14_sl_ks_thanh_cong'] = self._safe_int(r.get('SL KS thành công', 0))
                            data['c14_sl_kh_hai_long'] = self._safe_int(r.get('SL KH hài lòng', 0))
                            data['c14_khong_hl_phuc_vu'] = self._safe_int(r.get('Không HL KT phục vụ', 0))
                            data['c14_ty_le_hl_phuc_vu'] = self._parse_percentage(r.get('Tỷ lệ HL KT phục vụ', 0))
                            data['c14_khong_hl_dich_vu'] = self._safe_int(r.get('Không HL KT dịch vụ', 0))
                            data['c14_ty_le_hl_dich_vu'] = self._parse_percentage(r.get('Tỷ lệ HL KT dịch vụ', 0))
                            data['c14_tong_phieu_hai_long'] = self._safe_int(r.get('Tổng phiếu hài lòng KT', 0))
                            data['c14_ty_le_kh_hai_long'] = self._parse_percentage(r.get('Tỷ lệ KH hài lòng', 0))
                            data['c14_bsc'] = self._parse_percentage(r.get('Điểm BSC', 0))

                    # C1.5
                    if 'c15' in c1_data:
                        df = c1_data['c15']
                        row_data = df[df[don_vi_col] == ten_to]
                        if len(row_data) > 0:
                            r = row_data.iloc[0]
                            data['c15_sm1'] = self._safe_int(r.get('SM1', 0))
                            data['c15_sm2'] = self._safe_int(r.get('SM2', 0))
                            data['c15_kq_thuc_hien'] = self._parse_percentage(r.get('KQ thực hiện chỉ tiêu', 0))
                            data['c15_bsc'] = self._parse_percentage(r.get('Điểm BSC', 0))

                    # Insert vào bảng c1_theo_to
                    columns = ', '.join(data.keys())
                    placeholders = ', '.join(['?' for _ in data])
                    cursor.execute(f'''
                        INSERT INTO c1_theo_to ({columns}) VALUES ({placeholders})
                    ''', list(data.values()))

                    # Nếu là dòng "Tổng", cũng insert vào c1_tong_hop
                    if ten_to == 'Tổng':
                        data_tong_hop = {k: v for k, v in data.items() if k != 'ten_to'}
                        columns_th = ', '.join(data_tong_hop.keys())
                        placeholders_th = ', '.join(['?' for _ in data_tong_hop])
                        cursor.execute(f'''
                            INSERT INTO c1_tong_hop ({columns_th}) VALUES ({placeholders_th})
                        ''', list(data_tong_hop.values()))

            conn.commit()

            # Đếm số bản ghi
            cursor.execute("SELECT COUNT(*) FROM c1_theo_to WHERE ngay_cap_nhat = ?", (self.today,))
            count_to = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM c1_tong_hop WHERE ngay_cap_nhat = ?", (self.today,))
            count_tong = cursor.fetchone()[0]

            self._log_import(conn, 'c1_tong_hop', count_tong)
            self._log_import(conn, 'c1_theo_to', count_to)
            conn.commit()

            print(f"   ✅ Đã import c1_tong_hop: {count_tong} bản ghi")
            print(f"   ✅ Đã import c1_theo_to: {count_to} bản ghi")
            return True

        except Exception as e:
            conn.rollback()
            print(f"   ❌ Lỗi import: {e}")
            import traceback
            traceback.print_exc()
            self._log_import(conn, 'c1_tong_hop', 0, 'error', str(e))
            conn.commit()
            return False
        finally:
            conn.close()

    def import_c1_theo_nvkt(self):
        """
        Import dữ liệu theo NVKT từ C1.4 và C1.5 chi tiết
        """
        print("📊 Đang import dữ liệu C1 theo NVKT...")

        conn = get_connection()
        cursor = conn.cursor()

        try:
            # Xóa dữ liệu cũ của ngày hôm nay
            cursor.execute("DELETE FROM c1_theo_nvkt WHERE ngay_cap_nhat = ?", (self.today,))

            nvkt_data = {}  # {(ten_to, ten_nvkt): data}

            # Đọc C1.4 chi tiết - TH_HL_NVKT
            c14_file = os.path.join(self.download_dir, "c1.4_chitiet_report.xlsx")
            if os.path.exists(c14_file):
                try:
                    df = pd.read_excel(c14_file, sheet_name='TH_HL_NVKT')
                    print(f"   ✓ Đọc được C1.4 NVKT: {len(df)} dòng")

                    for _, row in df.iterrows():
                        ten_to = self._safe_str(row.get('DOIVT', ''))
                        ten_nvkt = self._safe_str(row.get('NVKT', ''))
                        if not ten_to or not ten_nvkt:
                            continue

                        key = (ten_to, ten_nvkt)
                        if key not in nvkt_data:
                            nvkt_data[key] = {
                                'ngay_cap_nhat': self.today,
                                'ten_to': ten_to,
                                'ten_nvkt': ten_nvkt,
                            }

                        nvkt_data[key]['c14_tong_phieu_ks_thanh_cong'] = self._safe_int(row.get('Tổng phiếu KS thành công', 0))
                        nvkt_data[key]['c14_tong_phieu_hai_long'] = self._safe_int(row.get('Tổng phiếu KHL', 0))
                        nvkt_data[key]['c14_ty_le_hai_long'] = self._parse_percentage(row.get('Tỉ lệ HL NVKT (%)', 0))

                except Exception as e:
                    print(f"   ⚠️ Lỗi đọc C1.4 NVKT: {e}")

            # Đọc C1.5 chi tiết - KQ_C15_chitiet
            c15_file = os.path.join(self.download_dir, "c1.5_chitiet_report.xlsx")
            if os.path.exists(c15_file):
                try:
                    df = pd.read_excel(c15_file, sheet_name='KQ_C15_chitiet')
                    print(f"   ✓ Đọc được C1.5 NVKT: {len(df)} dòng")

                    for _, row in df.iterrows():
                        ten_to = self._safe_str(row.get('DOIVT', ''))
                        ten_nvkt = self._safe_str(row.get('NVKT', ''))
                        if not ten_to or not ten_nvkt:
                            continue

                        key = (ten_to, ten_nvkt)
                        if key not in nvkt_data:
                            nvkt_data[key] = {
                                'ngay_cap_nhat': self.today,
                                'ten_to': ten_to,
                                'ten_nvkt': ten_nvkt,
                            }

                        nvkt_data[key]['c15_phieu_dat'] = self._safe_int(row.get('Phiếu đạt', 0))
                        nvkt_data[key]['c15_tong_hoan_cong'] = self._safe_int(row.get('Tổng Hoàn công', 0))
                        nvkt_data[key]['c15_ty_le_dat'] = self._parse_percentage(row.get('Tỉ lệ đạt', 0))

                except Exception as e:
                    print(f"   ⚠️ Lỗi đọc C1.5 NVKT: {e}")

            # Insert dữ liệu NVKT
            count = 0
            for key, data in nvkt_data.items():
                # Đảm bảo tất cả các trường đều có giá trị mặc định
                data.setdefault('c14_tong_phieu_ks_thanh_cong', 0)
                data.setdefault('c14_tong_phieu_hai_long', 0)
                data.setdefault('c14_ty_le_hai_long', 0.0)
                data.setdefault('c15_phieu_dat', 0)
                data.setdefault('c15_tong_hoan_cong', 0)
                data.setdefault('c15_ty_le_dat', 0.0)

                columns = ', '.join(data.keys())
                placeholders = ', '.join(['?' for _ in data])
                cursor.execute(f'''
                    INSERT INTO c1_theo_nvkt ({columns}) VALUES ({placeholders})
                ''', list(data.values()))
                count += 1

            conn.commit()
            self._log_import(conn, 'c1_theo_nvkt', count)
            conn.commit()

            print(f"   ✅ Đã import c1_theo_nvkt: {count} bản ghi")
            return True

        except Exception as e:
            conn.rollback()
            print(f"   ❌ Lỗi import NVKT: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            conn.close()

    def import_c1_4_chi_tiet(self):
        """Import chi tiết C1.4 (phiếu hài lòng)"""
        print("📊 Đang import chi tiết C1.4...")

        filename = os.path.join(self.download_dir, "c1.4_chitiet_report.xlsx")
        if not os.path.exists(filename):
            print(f"   ⚠️ Không tìm thấy file: {filename}")
            return False

        conn = get_connection()
        cursor = conn.cursor()

        try:
            # Xóa dữ liệu cũ
            cursor.execute("DELETE FROM c1_4_chi_tiet WHERE ngay_cap_nhat = ?", (self.today,))

            df = pd.read_excel(filename, sheet_name='Sheet1')
            print(f"   ✓ Đọc được C1.4 chi tiết: {len(df)} dòng")

            count = 0
            for _, row in df.iterrows():
                cursor.execute('''
                    INSERT INTO c1_4_chi_tiet (
                        ngay_cap_nhat, ma_tb, baohong_id, hdtb_id, nguoi_tl,
                        dia_chi_ld, dien_thoai_ks, dien_thoai_lh, ghi_chu, nguoi_cn,
                        do_hl, ma_tl, hl, ktc, ktm, khl_kt, khl_kd, nd_ktc_ktm,
                        ten_dv_hni, ngay_hoi, ngay_hc, ten_kv, doi_vt, ttvt,
                        nguoi_khoa, ten_nvkt_db
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    self.today,
                    self._safe_str(row.get('MA_TB', '')),
                    self._safe_str(row.get('BAOHONG_ID', '')),
                    self._safe_str(row.get('HDTB_ID', '')),
                    self._safe_str(row.get('NGUOI_TL', '')),
                    self._safe_str(row.get('DIACHI_LD', '')),
                    self._safe_str(row.get('DIENTHOAI_KS', '')),
                    self._safe_str(row.get('DIENTHOAI_LH', '')),
                    self._safe_str(row.get('GHICHU', '')),
                    self._safe_str(row.get('NGUOI_CN', '')),
                    self._safe_str(row.get('DO_HL', '')),
                    self._safe_str(row.get('MA_TL', '')),
                    self._safe_str(row.get('HL', '')),
                    self._safe_str(row.get('KTC', '')),
                    self._safe_str(row.get('KTM', '')),
                    self._safe_str(row.get('KHL_KT', '')),
                    self._safe_str(row.get('KHL_KD', '')),
                    self._safe_str(row.get('ND_KTC_KTM', '')),
                    self._safe_str(row.get('TEN_DV_HNI', '')),
                    self._safe_str(row.get('NGAY_HOI', '')),
                    self._safe_str(row.get('NGAY_HC', '')),
                    self._safe_str(row.get('TEN_KV', '')),
                    self._safe_str(row.get('DOIVT', '')),
                    self._safe_str(row.get('TTVT', '')),
                    self._safe_str(row.get('NGUOI_KHOA', '')),
                    self._safe_str(row.get('TEN_NVKT_DB', '')),
                ))
                count += 1

            conn.commit()
            self._log_import(conn, 'c1_4_chi_tiet', count)
            conn.commit()

            print(f"   ✅ Đã import c1_4_chi_tiet: {count} bản ghi")
            return True

        except Exception as e:
            conn.rollback()
            print(f"   ❌ Lỗi import C1.4 chi tiết: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            conn.close()

    def import_c1_5_chi_tiet(self):
        """Import chi tiết C1.5 (hoàn công)"""
        print("📊 Đang import chi tiết C1.5...")

        filename = os.path.join(self.download_dir, "c1.5_chitiet_report.xlsx")
        if not os.path.exists(filename):
            print(f"   ⚠️ Không tìm thấy file: {filename}")
            return False

        conn = get_connection()
        cursor = conn.cursor()

        try:
            # Xóa dữ liệu cũ
            cursor.execute("DELETE FROM c1_5_chi_tiet WHERE ngay_cap_nhat = ?", (self.today,))

            df = pd.read_excel(filename, sheet_name='Sheet1')
            print(f"   ✓ Đọc được C1.5 chi tiết: {len(df)} dòng")

            count = 0
            for _, row in df.iterrows():
                cursor.execute('''
                    INSERT INTO c1_5_chi_tiet (
                        ngay_cap_nhat, ma_tb, hdtb_id, ma_gd, ten_dvvt_hni,
                        so_ngay_hoan_thanh, so_gio_hoan_thanh, toanha_id, ten_kieu_ld,
                        ngay_giao_phieu, ngay_hc, dat_chi_tieu, ten_kv, nguoi_khoa,
                        nvkt_dia_ban, diem_chia, doi_vt
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    self.today,
                    self._safe_str(row.get('MA_TB', '')),
                    self._safe_str(row.get('HDTB_ID', '')),
                    self._safe_str(row.get('MA_GD', '')),
                    self._safe_str(row.get('TEN_DVVT_HNI', '')),
                    self._safe_int(row.get('SONGAY_HOANTHANH', 0)),
                    self._parse_percentage(row.get('SOGIO_HOANTHANH', 0)),
                    self._safe_str(row.get('TOANHA_ID', '')),
                    self._safe_str(row.get('TEN_KIEULD', '')),
                    self._safe_str(row.get('NGAY_GIAOPHIEU', '')),
                    self._safe_str(row.get('NGAY_HC', '')),
                    self._safe_str(row.get('Đạt chỉ tiêu', '')),
                    self._safe_str(row.get('TEN_KV', '')),
                    self._safe_str(row.get('Người khoá', '')),
                    self._safe_str(row.get('NVKT địa bàn', '')),
                    self._safe_str(row.get('DIEMCHIA', '')),
                    self._safe_str(row.get('DOIVT', '')),
                ))
                count += 1

            conn.commit()
            self._log_import(conn, 'c1_5_chi_tiet', count)
            conn.commit()

            print(f"   ✅ Đã import c1_5_chi_tiet: {count} bản ghi")
            return True

        except Exception as e:
            conn.rollback()
            print(f"   ❌ Lỗi import C1.5 chi tiết: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            conn.close()

    def import_all(self):
        """Import tất cả dữ liệu C1"""
        print("\n" + "=" * 50)
        print("🚀 BẮT ĐẦU IMPORT DỮ LIỆU C1")
        print("=" * 50)
        print(f"📅 Ngày: {self.today}")

        results = {
            'tong_hop': self.import_c1_tong_hop(),
            'theo_nvkt': self.import_c1_theo_nvkt(),
            'c14_chi_tiet': self.import_c1_4_chi_tiet(),
            'c15_chi_tiet': self.import_c1_5_chi_tiet(),
        }

        print("\n" + "=" * 50)
        print("📊 KẾT QUẢ IMPORT:")
        for key, success in results.items():
            status = "✅" if success else "❌"
            print(f"   {status} {key}")

        success_count = sum(results.values())
        total_count = len(results)
        print(f"\n   Thành công: {success_count}/{total_count}")
        print("=" * 50)

        return all(results.values())


def show_summary():
    """Hiển thị tóm tắt dữ liệu trong database"""
    conn = get_connection()
    cursor = conn.cursor()

    print("\n📊 TÓM TẮT DỮ LIỆU TRONG DATABASE C1")
    print("=" * 50)

    # Đếm số bản ghi trong các bảng
    tables = ['c1_tong_hop', 'c1_theo_to', 'c1_theo_nvkt', 'c1_4_chi_tiet', 'c1_5_chi_tiet']
    for table in tables:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        count = cursor.fetchone()[0]
        print(f"   {table}: {count} bản ghi")

    # Hiển thị các ngày có dữ liệu
    cursor.execute("SELECT DISTINCT ngay_cap_nhat FROM c1_tong_hop ORDER BY ngay_cap_nhat DESC LIMIT 10")
    dates = cursor.fetchall()
    if dates:
        print(f"\n📅 Các ngày có dữ liệu (10 gần nhất):")
        for d in dates:
            print(f"   - {d[0]}")

    conn.close()


if __name__ == "__main__":
    importer = C1HistoryImporter()
    importer.import_all()
    show_summary()
