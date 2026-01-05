#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script truy vấn và tạo báo cáo từ database
"""

import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import argparse

class BaoCaoQuery:
    def __init__(self, db_path='baocao_hanoi.db'):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)

    def close(self):
        """Đóng kết nối"""
        if self.conn:
            self.conn.close()

    def bao_cao_theo_ngay(self, ngay, loai_dv='FIBER'):
        """Báo cáo tổng hợp theo ngày"""
        sql = """
            SELECT
                don_vi,
                nvkt,
                COUNT(CASE WHEN loai_dv = ? THEN 1 END) as so_hoan_cong,
                (SELECT COUNT(*) FROM ngung_psc WHERE ngay_bao_cao = ? AND loai_dv = ? AND ngung_psc.nvkt = hoan_cong.nvkt) as so_ngung_psc
            FROM hoan_cong
            WHERE ngay_bao_cao = ? AND loai_dv = ?
            GROUP BY don_vi, nvkt
            ORDER BY don_vi, so_hoan_cong DESC
        """

        df = pd.read_sql_query(sql, self.conn, params=[loai_dv, ngay, loai_dv, ngay, loai_dv])
        df['thuc_tang'] = df['so_hoan_cong'] - df['so_ngung_psc']
        df['ty_le_ngung_psc'] = (df['so_ngung_psc'] / df['so_hoan_cong'] * 100).round(2)

        return df

    def bao_cao_theo_thang(self, thang, nam, loai_dv='FIBER'):
        """Báo cáo tổng hợp theo tháng"""
        sql = """
            SELECT
                don_vi,
                nvkt,
                COUNT(DISTINCT ma_tb) as so_hoan_cong,
                (SELECT COUNT(DISTINCT ma_tb)
                 FROM ngung_psc
                 WHERE strftime('%Y-%m', ngay_bao_cao) = ?
                   AND loai_dv = ?
                   AND ngung_psc.nvkt = hoan_cong.nvkt) as so_ngung_psc
            FROM hoan_cong
            WHERE strftime('%Y-%m', ngay_bao_cao) = ? AND loai_dv = ?
            GROUP BY don_vi, nvkt
            ORDER BY don_vi, so_hoan_cong DESC
        """

        thang_str = f"{nam}-{str(thang).zfill(2)}"
        df = pd.read_sql_query(sql, self.conn, params=[thang_str, loai_dv, thang_str, loai_dv])
        df['thuc_tang'] = df['so_hoan_cong'] - df['so_ngung_psc']
        df['ty_le_ngung_psc'] = (df['so_ngung_psc'] / df['so_hoan_cong'] * 100).round(2)

        return df

    def bao_cao_xu_huong(self, tu_ngay, den_ngay, loai_dv='FIBER'):
        """Báo cáo xu hướng theo khoảng thời gian"""
        sql = """
            SELECT
                ngay_bao_cao,
                don_vi,
                COUNT(*) as so_hoan_cong
            FROM hoan_cong
            WHERE ngay_bao_cao BETWEEN ? AND ?
              AND loai_dv = ?
            GROUP BY ngay_bao_cao, don_vi
            ORDER BY ngay_bao_cao, don_vi
        """

        df = pd.read_sql_query(sql, self.conn, params=[tu_ngay, den_ngay, loai_dv])

        # Pivot để có dạng ngày x đơn vị
        df_pivot = df.pivot(index='ngay_bao_cao', columns='don_vi', values='so_hoan_cong')
        df_pivot = df_pivot.fillna(0).astype(int)

        return df_pivot

    def bao_cao_suy_hao_cao(self, ngay):
        """Báo cáo suy hao cao theo ngày"""
        sql = """
            SELECT
                doi_one as don_vi,
                nvkt_db_normalized as nvkt,
                COUNT(*) as so_tb_suy_hao
            FROM suy_hao_cao
            WHERE ngay_bao_cao = ?
            GROUP BY doi_one, nvkt_db_normalized
            ORDER BY doi_one, so_tb_suy_hao DESC
        """

        df = pd.read_sql_query(sql, self.conn, params=[ngay])
        return df

    def bao_cao_bien_dong_suy_hao(self, ngay_hien_tai, ngay_truoc):
        """Báo cáo biến động suy hao cao so với ngày trước"""
        sql = """
            SELECT
                COALESCE(h.doi_one, t.doi_one) as don_vi,
                COALESCE(h.nvkt_db_normalized, t.nvkt_db_normalized) as nvkt,
                COALESCE(h.count, 0) as hien_tai,
                COALESCE(t.count, 0) as truoc_do,
                COALESCE(h.count, 0) - COALESCE(t.count, 0) as bien_dong
            FROM (
                SELECT doi_one, nvkt_db_normalized, COUNT(*) as count
                FROM suy_hao_cao
                WHERE ngay_bao_cao = ?
                GROUP BY doi_one, nvkt_db_normalized
            ) h
            FULL OUTER JOIN (
                SELECT doi_one, nvkt_db_normalized, COUNT(*) as count
                FROM suy_hao_cao
                WHERE ngay_bao_cao = ?
                GROUP BY doi_one, nvkt_db_normalized
            ) t ON h.doi_one = t.doi_one AND h.nvkt_db_normalized = t.nvkt_db_normalized
            WHERE COALESCE(h.count, 0) != COALESCE(t.count, 0)
            ORDER BY bien_dong DESC
        """

        df = pd.read_sql_query(sql, self.conn, params=[ngay_hien_tai, ngay_truoc])
        return df

    def bao_cao_top_nvkt(self, tu_ngay, den_ngay, loai_dv='FIBER', top=10, sap_xep='hoan_cong'):
        """Báo cáo top NVKT theo chỉ tiêu"""
        sql = """
            SELECT
                h.nvkt,
                h.don_vi,
                COUNT(DISTINCT h.ma_tb) as so_hoan_cong,
                COUNT(DISTINCT n.ma_tb) as so_ngung_psc,
                COUNT(DISTINCT h.ma_tb) - COUNT(DISTINCT n.ma_tb) as thuc_tang
            FROM hoan_cong h
            LEFT JOIN ngung_psc n ON h.nvkt = n.nvkt
                AND h.ngay_bao_cao = n.ngay_bao_cao
                AND h.loai_dv = n.loai_dv
            WHERE h.ngay_bao_cao BETWEEN ? AND ?
              AND h.loai_dv = ?
            GROUP BY h.nvkt, h.don_vi
            ORDER BY {} DESC
            LIMIT ?
        """.format(sap_xep)

        df = pd.read_sql_query(sql, self.conn, params=[tu_ngay, den_ngay, loai_dv, top])
        df['ty_le_ngung_psc'] = (df['so_ngung_psc'] / df['so_hoan_cong'] * 100).round(2)

        return df

    def bao_cao_theo_tuan(self, tuan, nam, loai_dv='FIBER'):
        """Báo cáo theo tuần"""
        # Tính ngày đầu và cuối của tuần
        first_day_of_year = datetime(nam, 1, 1)
        first_monday = first_day_of_year - timedelta(days=first_day_of_year.weekday())
        start_of_week = first_monday + timedelta(weeks=tuan - 1)
        end_of_week = start_of_week + timedelta(days=6)

        tu_ngay = start_of_week.strftime('%Y-%m-%d')
        den_ngay = end_of_week.strftime('%Y-%m-%d')

        sql = """
            SELECT
                don_vi,
                nvkt,
                COUNT(DISTINCT ma_tb) as so_hoan_cong
            FROM hoan_cong
            WHERE ngay_bao_cao BETWEEN ? AND ?
              AND loai_dv = ?
            GROUP BY don_vi, nvkt
            ORDER BY don_vi, so_hoan_cong DESC
        """

        df = pd.read_sql_query(sql, self.conn, params=[tu_ngay, den_ngay, loai_dv])

        return df

    def thong_ke_tong_quan(self):
        """Thống kê tổng quan database"""
        stats = {}

        # Số lượng bản ghi
        tables = ['hoan_cong', 'ngung_psc', 'thuc_tang', 'suy_hao_cao']
        for table in tables:
            sql = f"SELECT COUNT(*) as count FROM {table}"
            result = pd.read_sql_query(sql, self.conn)
            stats[table] = result['count'][0]

        # Khoảng thời gian dữ liệu
        sql = """
            SELECT
                MIN(ngay_bao_cao) as tu_ngay,
                MAX(ngay_bao_cao) as den_ngay
            FROM (
                SELECT ngay_bao_cao FROM hoan_cong
                UNION
                SELECT ngay_bao_cao FROM ngung_psc
                UNION
                SELECT ngay_bao_cao FROM suy_hao_cao
            )
        """
        result = pd.read_sql_query(sql, self.conn)
        stats['tu_ngay'] = result['tu_ngay'][0]
        stats['den_ngay'] = result['den_ngay'][0]

        # Số lượng NVKT
        sql = "SELECT COUNT(DISTINCT nvkt) as count FROM hoan_cong"
        result = pd.read_sql_query(sql, self.conn)
        stats['so_nvkt'] = result['count'][0]

        # Số lượng đơn vị
        sql = "SELECT COUNT(DISTINCT don_vi) as count FROM hoan_cong"
        result = pd.read_sql_query(sql, self.conn)
        stats['so_don_vi'] = result['count'][0]

        return stats

    def export_to_excel(self, df, filename):
        """Export DataFrame ra Excel"""
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"✓ Đã xuất báo cáo ra file: {filename}")


def main():
    """Hàm chính"""
    parser = argparse.ArgumentParser(description='Truy vấn và tạo báo cáo từ database')
    parser.add_argument('--db', default='baocao_hanoi.db', help='Đường dẫn file database')
    parser.add_argument('--loai', choices=['ngay', 'tuan', 'thang', 'xu-huong', 'suy-hao', 'top-nvkt', 'thong-ke'],
                        required=True, help='Loại báo cáo')
    parser.add_argument('--ngay', help='Ngày báo cáo (YYYY-MM-DD)')
    parser.add_argument('--tuan', type=int, help='Tuần (1-52)')
    parser.add_argument('--thang', type=int, help='Tháng (1-12)')
    parser.add_argument('--nam', type=int, help='Năm')
    parser.add_argument('--tu-ngay', help='Từ ngày (YYYY-MM-DD)')
    parser.add_argument('--den-ngay', help='Đến ngày (YYYY-MM-DD)')
    parser.add_argument('--loai-dv', choices=['FIBER', 'MYTV'], default='FIBER', help='Loại dịch vụ')
    parser.add_argument('--top', type=int, default=10, help='Top N (mặc định 10)')
    parser.add_argument('--sap-xep', choices=['hoan_cong', 'thuc_tang', 'so_ngung_psc'], default='hoan_cong',
                        help='Sắp xếp theo')
    parser.add_argument('--export', help='Xuất ra file Excel')

    args = parser.parse_args()

    query = BaoCaoQuery(args.db)

    try:
        if args.loai == 'ngay':
            if not args.ngay:
                print("✗ Vui lòng cung cấp --ngay")
                return
            df = query.bao_cao_theo_ngay(args.ngay, args.loai_dv)
            print(f"\n📊 BÁO CÁO THEO NGÀY {args.ngay} - {args.loai_dv}")
            print("="*80)
            print(df.to_string(index=False))

        elif args.loai == 'tuan':
            if not args.tuan or not args.nam:
                print("✗ Vui lòng cung cấp --tuan và --nam")
                return
            df = query.bao_cao_theo_tuan(args.tuan, args.nam, args.loai_dv)
            print(f"\n📊 BÁO CÁO TUẦN {args.tuan}/{args.nam} - {args.loai_dv}")
            print("="*80)
            print(df.to_string(index=False))

        elif args.loai == 'thang':
            if not args.thang or not args.nam:
                print("✗ Vui lòng cung cấp --thang và --nam")
                return
            df = query.bao_cao_theo_thang(args.thang, args.nam, args.loai_dv)
            print(f"\n📊 BÁO CÁO THÁNG {args.thang}/{args.nam} - {args.loai_dv}")
            print("="*80)
            print(df.to_string(index=False))

        elif args.loai == 'xu-huong':
            if not args.tu_ngay or not args.den_ngay:
                print("✗ Vui lòng cung cấp --tu-ngay và --den-ngay")
                return
            df = query.bao_cao_xu_huong(args.tu_ngay, args.den_ngay, args.loai_dv)
            print(f"\n📊 BÁO CÁO XU HƯỚNG TỪ {args.tu_ngay} ĐẾN {args.den_ngay} - {args.loai_dv}")
            print("="*80)
            print(df.to_string())

        elif args.loai == 'suy-hao':
            if not args.ngay:
                print("✗ Vui lòng cung cấp --ngay")
                return
            df = query.bao_cao_suy_hao_cao(args.ngay)
            print(f"\n📊 BÁO CÁO SUY HAO CAO NGÀY {args.ngay}")
            print("="*80)
            print(df.to_string(index=False))

        elif args.loai == 'top-nvkt':
            if not args.tu_ngay or not args.den_ngay:
                print("✗ Vui lòng cung cấp --tu-ngay và --den-ngay")
                return
            df = query.bao_cao_top_nvkt(args.tu_ngay, args.den_ngay, args.loai_dv, args.top, args.sap_xep)
            print(f"\n📊 TOP {args.top} NVKT TỪ {args.tu_ngay} ĐẾN {args.den_ngay} - {args.loai_dv}")
            print(f"Sắp xếp theo: {args.sap_xep}")
            print("="*80)
            print(df.to_string(index=False))

        elif args.loai == 'thong-ke':
            stats = query.thong_ke_tong_quan()
            print("\n📊 THỐNG KÊ TỔNG QUAN DATABASE")
            print("="*80)
            print(f"Khoảng thời gian: {stats['tu_ngay']} đến {stats['den_ngay']}")
            print(f"Số lượng bản ghi:")
            print(f"  - Hoàn công: {stats['hoan_cong']:,}")
            print(f"  - Ngừng PSC: {stats['ngung_psc']:,}")
            print(f"  - Thực tăng: {stats['thuc_tang']:,}")
            print(f"  - Suy hao cao: {stats['suy_hao_cao']:,}")
            print(f"Số lượng NVKT: {stats['so_nvkt']}")
            print(f"Số lượng đơn vị: {stats['so_don_vi']}")
            df = None

        # Export nếu có yêu cầu
        if args.export and df is not None:
            query.export_to_excel(df, args.export)

    finally:
        query.close()


if __name__ == '__main__':
    main()
