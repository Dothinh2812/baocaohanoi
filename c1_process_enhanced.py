# -*- coding: utf-8 -*-
"""
Enhanced version of process_I15_report() with historical tracking
Thay thế hàm process_I15_report() trong c1_process.py
"""

import pandas as pd
import os
import sqlite3
import re
from datetime import datetime, timedelta


def normalize_nvkt(x):
    """Chuẩn hóa tên NVKT_DB - giữ phần sau dấu '-'"""
    if not isinstance(x, str):
        return x
    if '-' in x:
        x = x.split('-')[1].strip()
    x = re.sub(r'\([^)]*\)', '', x).strip()
    return x


def process_I15_report_with_tracking():
    """
    Xử lý báo cáo I1.5 với tracking lịch sử:
    1. Đọc file I1.5 report.xlsx
    2. Tra cứu thông tin từ danhba.db
    3. Chuẩn hóa cột NVKT_DB
    4. So sánh với dữ liệu ngày hôm qua
    5. Tạo các sheet: TH_SHC_I15, Tang_moi, Giam_het, Van_con, Bien_dong_tong_hop
    6. Lưu vào database để tracking lịch sử
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO I1.5 (VỚI TRACKING LỊCH SỬ)")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "I1.5 report.xlsx")
        db_file = "danhba.db"
        history_db = "suy_hao_history.db"

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}, tổng số cột: {df.shape[1]}")

        # Lấy ngày báo cáo từ cột NGAY_SUYHAO
        if 'NGAY_SUYHAO' in df.columns and len(df) > 0:
            ngay_str = df['NGAY_SUYHAO'].iloc[0]
            try:
                report_date = pd.to_datetime(ngay_str, format='%d/%m/%Y').strftime('%Y-%m-%d')
                print(f"✓ Ngày báo cáo: {report_date}")
            except:
                report_date = datetime.now().strftime('%Y-%m-%d')
                print(f"⚠️  Không parse được ngày, dùng ngày hiện tại: {report_date}")
        else:
            report_date = datetime.now().strftime('%Y-%m-%d')
            print(f"⚠️  Không tìm thấy NGAY_SUYHAO, dùng ngày hiện tại: {report_date}")

        # Tra cứu thông tin từ danhba.db
        print("\n✓ Đang tra cứu thông tin từ danhba.db...")
        if os.path.exists(db_file):
            try:
                conn = sqlite3.connect(db_file)
                query = "SELECT MA_TB, THIETBI, SA, KETCUOI FROM danhba"
                df_danhba = pd.read_sql_query(query, conn)
                conn.close()

                print(f"✅ Đã đọc {len(df_danhba)} bản ghi từ danhba.db")

                if 'ACCOUNT_CTS' in df.columns:
                    cols_to_remove = ['MA_TB', 'THIETBI', 'SA', 'KETCUOI']
                    for col in cols_to_remove:
                        if col in df.columns:
                            df = df.drop(columns=[col])

                    df = df.merge(df_danhba, left_on='ACCOUNT_CTS', right_on='MA_TB', how='left')
                    if 'MA_TB' in df.columns:
                        df = df.drop(columns=['MA_TB'])
                    print(f"✅ Đã tra cứu và thêm các cột: THIETBI, SA, KETCUOI")
            except Exception as e:
                print(f"⚠️ Lỗi khi tra cứu danhba.db: {e}")
        else:
            print(f"⚠️ Không tìm thấy file {db_file}")

        # Chuẩn hóa cột NVKT_DB
        print("\n✓ Đang chuẩn hóa cột NVKT_DB...")
        if 'NVKT_DB' in df.columns:
            df['NVKT_DB_NORMALIZED'] = df['NVKT_DB'].apply(normalize_nvkt)
            print("✅ Đã chuẩn hóa cột NVKT_DB")
        else:
            print("⚠️ Không tìm thấy cột NVKT_DB")
            df['NVKT_DB_NORMALIZED'] = None

        # ==================================================================
        # SO SÁNH VỚI NGÀY HÔM QUA VÀ LƯU VÀO DATABASE
        # ==================================================================
        print("\n" + "="*80)
        print("TRACKING LỊCH SỬ VÀ SO SÁNH VỚI NGÀY HÔM QUA")
        print("="*80)

        if not os.path.exists(history_db):
            print(f"⚠️ Không tìm thấy {history_db}, bỏ qua tracking lịch sử")
            df_tang_moi = pd.DataFrame()
            df_giam_het = pd.DataFrame()
            df_van_con = pd.DataFrame()
            df_bien_dong = pd.DataFrame()
        else:
            hist_conn = sqlite3.connect(history_db)
            cursor = hist_conn.cursor()

            # Tính ngày hôm qua
            yesterday = (datetime.strptime(report_date, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')

            # Đọc dữ liệu ngày hôm qua
            print(f"\n✓ Đang đọc dữ liệu ngày {yesterday}...")
            df_yesterday = pd.read_sql_query(f"""
                SELECT account_cts FROM suy_hao_snapshots
                WHERE ngay_bao_cao = '{yesterday}'
            """, hist_conn)

            print(f"  Ngày {yesterday}: {len(df_yesterday)} thuê bao")
            print(f"  Ngày {report_date}: {len(df)} thuê bao")

            # Phân loại (loại bỏ NaN/None)
            if 'ACCOUNT_CTS' in df.columns:
                accounts_today = set([x for x in df['ACCOUNT_CTS'].tolist() if pd.notna(x) and str(x).strip() != ''])
            else:
                accounts_today = set()

            if len(df_yesterday) > 0:
                accounts_yesterday = set([x for x in df_yesterday['account_cts'].tolist() if pd.notna(x) and str(x).strip() != ''])
            else:
                accounts_yesterday = set()

            tang_moi_set = accounts_today - accounts_yesterday
            giam_het_set = accounts_yesterday - accounts_today
            van_con_set = accounts_today & accounts_yesterday

            print(f"\n✓ Phân tích biến động:")
            print(f"  🆕 TĂNG MỚI: {len(tang_moi_set)} thuê bao")
            print(f"  ⬇️  GIẢM/HẾT: {len(giam_het_set)} thuê bao")
            print(f"  ↔️  VẪN CÒN: {len(van_con_set)} thuê bao")

            # Tạo DataFrame cho từng loại
            df_tang_moi = df[df['ACCOUNT_CTS'].isin(tang_moi_set)].copy() if len(tang_moi_set) > 0 else pd.DataFrame()
            df_van_con = df[df['ACCOUNT_CTS'].isin(van_con_set)].copy() if len(van_con_set) > 0 else pd.DataFrame()

            # Lấy thông tin GIẢM/HẾT từ database
            if len(giam_het_set) > 0:
                accounts_str = ','.join([f"'{x}'" for x in list(giam_het_set)[:1000]])  # Limit 1000 để tránh query quá dài
                df_giam_het = pd.read_sql_query(f"""
                    SELECT s.*, t.so_ngay_lien_tuc
                    FROM suy_hao_snapshots s
                    LEFT JOIN suy_hao_tracking t ON s.account_cts = t.account_cts
                    WHERE s.ngay_bao_cao = '{yesterday}'
                      AND s.account_cts IN ({accounts_str})
                """, hist_conn)
            else:
                df_giam_het = pd.DataFrame()

            # Thêm số ngày liên tục cho VẪN CÒN
            if len(van_con_set) > 0 and len(df_van_con) > 0:
                print("\n✓ Đang lấy số ngày liên tục cho thuê bao VẪN CÒN...")
                tracking_data = pd.read_sql_query(f"""
                    SELECT account_cts, so_ngay_lien_tuc
                    FROM suy_hao_tracking
                    WHERE account_cts IN ({','.join([f"'{x}'" for x in list(van_con_set)[:1000]])})
                """, hist_conn)

                df_van_con = df_van_con.merge(
                    tracking_data,
                    left_on='ACCOUNT_CTS',
                    right_on='account_cts',
                    how='left'
                )
                if 'account_cts' in df_van_con.columns:
                    df_van_con = df_van_con.drop(columns=['account_cts'])

                # Cộng thêm 1 ngày (vì hôm nay tiếp tục suy hao)
                df_van_con['so_ngay_lien_tuc'] = df_van_con['so_ngay_lien_tuc'].fillna(1) + 1
            else:
                if len(df_van_con) > 0:
                    df_van_con['so_ngay_lien_tuc'] = 2  # Ít nhất 2 ngày (hôm qua + hôm nay)

            # Lưu snapshot hôm nay vào database
            print(f"\n✓ Đang lưu snapshot ngày {report_date} vào database...")

            # Xóa dữ liệu cũ nếu có (để tránh trùng)
            cursor.execute("DELETE FROM suy_hao_snapshots WHERE ngay_bao_cao = ?", (report_date,))

            inserted = 0
            skipped = 0
            for idx, row in df.iterrows():
                account = row.get('ACCOUNT_CTS')
                # Bỏ qua nếu account là null hoặc NaN
                if pd.isna(account) or account is None or str(account).strip() == '':
                    skipped += 1
                    continue

                try:
                    cursor.execute("""
                        INSERT INTO suy_hao_snapshots (
                            ngay_bao_cao, account_cts, ten_tb_one, dt_onediachi_one,
                            doi_one, nvkt_db, nvkt_db_normalized, sa,
                            olt_cts, port_cts, thietbi, ketcuoi, trangthai_tb
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        report_date,
                        account,
                        row.get('TEN_TB_ONE'),
                        row.get('DT_ONEDIACHI_ONE'),
                        row.get('DOI_ONE'),
                        row.get('NVKT_DB'),
                        row.get('NVKT_DB_NORMALIZED'),
                        row.get('SA'),
                        row.get('OLT_CTS'),
                        row.get('PORT_CTS'),
                        row.get('THIETBI'),
                        row.get('KETCUOI'),
                        row.get('TRANGTHAI_TB')
                    ))
                    inserted += 1
                except Exception as e:
                    print(f"  ⚠️  Lỗi insert account '{account}': {e}")
                    skipped += 1

            if skipped > 0:
                print(f"  ✅ Đã lưu {inserted} bản ghi vào snapshots (bỏ qua {skipped} dòng)")
            else:
                print(f"  ✅ Đã lưu {inserted} bản ghi vào snapshots")

            # Cập nhật tracking table
            print(f"\n✓ Đang cập nhật bảng tracking...")

            # TĂNG MỚI: Insert mới
            for account in tang_moi_set:
                # Kiểm tra account có trong DataFrame không
                df_filtered = df[df['ACCOUNT_CTS'] == account]
                if len(df_filtered) > 0:
                    row_data = df_filtered.iloc[0]
                    cursor.execute("""
                        INSERT OR REPLACE INTO suy_hao_tracking (
                            account_cts, ngay_xuat_hien_dau_tien, ngay_thay_cuoi_cung,
                            so_ngay_lien_tuc, doi_one, nvkt_db, sa, trang_thai
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        account, report_date, report_date, 1,
                        row_data.get('DOI_ONE'), row_data.get('NVKT_DB_NORMALIZED'),
                        row_data.get('SA'), 'DANG_SUY_HAO'
                    ))
                else:
                    print(f"  ⚠️  Bỏ qua account không tìm thấy trong DataFrame: {account}")

            # VẪN CÒN: Cập nhật số ngày và ngày cuối
            for account in van_con_set:
                cursor.execute("""
                    UPDATE suy_hao_tracking
                    SET ngay_thay_cuoi_cung = ?,
                        so_ngay_lien_tuc = so_ngay_lien_tuc + 1,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE account_cts = ?
                """, (report_date, account))

            # GIẢM/HẾT: Cập nhật trạng thái
            for account in giam_het_set:
                cursor.execute("""
                    UPDATE suy_hao_tracking
                    SET trang_thai = 'DA_HET_SUY_HAO',
                        updated_at = CURRENT_TIMESTAMP
                    WHERE account_cts = ?
                """, (account,))

            # Lưu daily changes
            print(f"\n✓ Đang lưu daily changes...")
            cursor.execute("DELETE FROM suy_hao_daily_changes WHERE ngay_bao_cao = ?", (report_date,))

            def save_changes(df_changes, loai):
                for idx, row in df_changes.iterrows():
                    so_ngay = row.get('so_ngay_lien_tuc', 1) if loai != 'TANG_MOI' else 1
                    cursor.execute("""
                        INSERT INTO suy_hao_daily_changes (
                            ngay_bao_cao, account_cts, loai_bien_dong,
                            doi_one, nvkt_db, nvkt_db_normalized, sa, so_ngay_lien_tuc,
                            ten_tb_one, dt_onediachi_one, olt_cts, port_cts, thietbi, ketcuoi
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        report_date, row.get('ACCOUNT_CTS') or row.get('account_cts'), loai,
                        row.get('DOI_ONE') or row.get('doi_one'),
                        row.get('NVKT_DB') or row.get('nvkt_db'),
                        row.get('NVKT_DB_NORMALIZED') or row.get('nvkt_db_normalized'),
                        row.get('SA') or row.get('sa'),
                        so_ngay,
                        row.get('TEN_TB_ONE') or row.get('ten_tb_one'),
                        row.get('DT_ONEDIACHI_ONE') or row.get('dt_onediachi_one'),
                        row.get('OLT_CTS') or row.get('olt_cts'),
                        row.get('PORT_CTS') or row.get('port_cts'),
                        row.get('THIETBI') or row.get('thietbi'),
                        row.get('KETCUOI') or row.get('ketcuoi')
                    ))

            if len(df_tang_moi) > 0:
                save_changes(df_tang_moi, 'TANG_MOI')
            if len(df_giam_het) > 0:
                save_changes(df_giam_het, 'GIAM_HET')
            if len(df_van_con) > 0:
                save_changes(df_van_con, 'VAN_CON')

            # Tạo daily summary
            print(f"\n✓ Đang tạo daily summary...")
            cursor.execute("DELETE FROM suy_hao_daily_summary WHERE ngay_bao_cao = ?", (report_date,))

            cursor.execute(f"""
                INSERT INTO suy_hao_daily_summary (
                    ngay_bao_cao, doi_one, nvkt_db_normalized,
                    tong_so_hien_tai, so_tang_moi, so_giam_het, so_van_con
                )
                SELECT
                    '{report_date}',
                    doi_one,
                    nvkt_db_normalized,
                    SUM(CASE WHEN loai_bien_dong IN ('TANG_MOI', 'VAN_CON') THEN 1 ELSE 0 END),
                    SUM(CASE WHEN loai_bien_dong = 'TANG_MOI' THEN 1 ELSE 0 END),
                    SUM(CASE WHEN loai_bien_dong = 'GIAM_HET' THEN 1 ELSE 0 END),
                    SUM(CASE WHEN loai_bien_dong = 'VAN_CON' THEN 1 ELSE 0 END)
                FROM suy_hao_daily_changes
                WHERE ngay_bao_cao = '{report_date}'
                GROUP BY doi_one, nvkt_db_normalized
            """)

            hist_conn.commit()

            # Đọc summary để tạo báo cáo
            df_bien_dong = pd.read_sql_query(f"""
                SELECT
                    doi_one as "Đơn vị",
                    nvkt_db_normalized as "NVKT_DB",
                    tong_so_hien_tai as "Tổng số hiện tại",
                    so_tang_moi as "Tăng mới",
                    so_giam_het as "Giảm/Hết",
                    so_van_con as "Vẫn còn"
                FROM suy_hao_daily_summary
                WHERE ngay_bao_cao = '{report_date}'
                ORDER BY doi_one, nvkt_db_normalized
            """, hist_conn)

            hist_conn.close()

            print(f"  ✅ Đã lưu toàn bộ vào database lịch sử")

        # ==================================================================
        # TẠO CÁC SHEET THỐNG KÊ (LOGIC CŨ)
        # ==================================================================
        print("\n" + "="*80)
        print("TẠO CÁC SHEET THỐNG KÊ")
        print("="*80)

        # Sheet tổng hợp theo NVKT_DB và DOI_ONE
        print("\n✓ Đang đếm theo NVKT_DB và DOI_ONE...")
        if 'NVKT_DB_NORMALIZED' in df.columns and 'DOI_ONE' in df.columns:
            df_result = df.groupby(['NVKT_DB_NORMALIZED', 'DOI_ONE']).size().reset_index(name='Count')
            df_result = df_result[['DOI_ONE', 'NVKT_DB_NORMALIZED', 'Count']]
            df_result.columns = ['Đơn vị', 'NVKT_DB', 'Số TB Suy hao cao K1']
            df_result = df_result.sort_values(by='Đơn vị').reset_index(drop=True)
            print(f"✅ Đã đếm xong, tổng số nhóm: {len(df_result)}")
        else:
            print("❌ Không tìm thấy cột NVKT_DB_NORMALIZED hoặc DOI_ONE")
            return False

        # Tổng hợp theo tổ
        print("\n✓ Đang tạo tổng hợp theo tổ...")
        df_by_to = df_result.groupby('Đơn vị')['Số TB Suy hao cao K1'].sum().reset_index()
        df_by_to = df_by_to.sort_values(by='Đơn vị').reset_index(drop=True)
        total_row = pd.DataFrame({'Đơn vị': ['Tổng'], 'Số TB Suy hao cao K1': [df_by_to['Số TB Suy hao cao K1'].sum()]})
        df_by_to = pd.concat([df_by_to, total_row], ignore_index=True)

        # Thống kê theo SA
        print("\n✓ Đang tạo thống kê theo SA...")
        if 'SA' in df.columns:
            df_by_sa = df.groupby('SA').size().reset_index(name='Số lượng')
            df_by_sa = df_by_sa.sort_values(by='Số lượng', ascending=False).reset_index(drop=True)
            total_sa_row = pd.DataFrame({'SA': ['Tổng'], 'Số lượng': [df_by_sa['Số lượng'].sum()]})
            df_by_sa = pd.concat([df_by_sa, total_sa_row], ignore_index=True)
        else:
            df_by_sa = None

        # Danh sách chi tiết cho từng NVKT_DB
        print("\n✓ Đang tạo danh sách chi tiết cho từng NVKT_DB...")
        columns_to_keep = ['ACCOUNT_CTS', 'TEN_TB_ONE', 'DT_ONEDIACHI_ONE', 'NGAY_SUYHAO',
                          'OLT_CTS', 'PORT_CTS', 'THIETBI', 'SA', 'KETCUOI', 'NVKT_DB_NORMALIZED']
        missing_cols = [col for col in columns_to_keep if col not in df.columns]
        if missing_cols:
            print(f"⚠️ Không tìm thấy các cột: {', '.join(missing_cols)}")
            columns_to_keep = [col for col in columns_to_keep if col in df.columns]

        df_detail = df[columns_to_keep].copy()
        nvkt_list = df_detail['NVKT_DB_NORMALIZED'].unique()
        print(f"✅ Tìm thấy {len(nvkt_list)} NVKT_DB cần tạo sheet chi tiết")

        # ==================================================================
        # GHI VÀO FILE EXCEL
        # ==================================================================
        print("\n✓ Đang ghi vào các sheet...")

        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # Sheet1: Dữ liệu gốc đầy đủ
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            print(f"  ✅ Sheet1 (dữ liệu gốc): {len(df)} dòng")

            # Các sheet thống kê cũ
            df_result.to_excel(writer, sheet_name='TH_SHC_I15', index=False)
            df_by_to.to_excel(writer, sheet_name='TH_SHC_theo_to', index=False)
            if df_by_sa is not None:
                df_by_sa.to_excel(writer, sheet_name='shc_theo_SA', index=False)

            # SHEET MỚI: Biến động tổng hợp
            if len(df_bien_dong) > 0:
                df_bien_dong.to_excel(writer, sheet_name='Bien_dong_tong_hop', index=False)
                print(f"  ✅ Bien_dong_tong_hop: {len(df_bien_dong)} dòng")

            # SHEET MỚI: TĂNG MỚI
            if len(df_tang_moi) > 0:
                cols_tang = ['ACCOUNT_CTS', 'TEN_TB_ONE', 'DT_ONEDIACHI_ONE', 'DOI_ONE',
                            'NVKT_DB_NORMALIZED', 'SA', 'OLT_CTS', 'PORT_CTS', 'THIETBI', 'KETCUOI']
                cols_tang = [c for c in cols_tang if c in df_tang_moi.columns]
                df_tang_moi[cols_tang].to_excel(writer, sheet_name='Tang_moi', index=False)
                print(f"  ✅ Tang_moi: {len(df_tang_moi)} dòng")

            # SHEET MỚI: GIẢM/HẾT
            if len(df_giam_het) > 0:
                cols_giam = ['account_cts', 'ten_tb_one', 'dt_onediachi_one', 'doi_one',
                            'nvkt_db_normalized', 'sa', 'so_ngay_lien_tuc', 'olt_cts', 'port_cts', 'thietbi', 'ketcuoi']
                cols_giam = [c for c in cols_giam if c in df_giam_het.columns]
                df_giam_out = df_giam_het[cols_giam].copy()
                df_giam_out.columns = [c.upper() if c != 'so_ngay_lien_tuc' else 'Số ngày suy hao' for c in df_giam_out.columns]
                df_giam_out.to_excel(writer, sheet_name='Giam_het', index=False)
                print(f"  ✅ Giam_het: {len(df_giam_het)} dòng")

            # SHEET MỚI: VẪN CÒN
            if len(df_van_con) > 0:
                cols_van = ['ACCOUNT_CTS', 'TEN_TB_ONE', 'DT_ONEDIACHI_ONE', 'DOI_ONE',
                           'NVKT_DB_NORMALIZED', 'SA', 'so_ngay_lien_tuc', 'OLT_CTS', 'PORT_CTS', 'THIETBI', 'KETCUOI']
                cols_van = [c for c in cols_van if c in df_van_con.columns]
                df_van_out = df_van_con[cols_van].copy()
                if 'so_ngay_lien_tuc' in df_van_out.columns:
                    df_van_out = df_van_out.rename(columns={'so_ngay_lien_tuc': 'Số ngày liên tục'})
                df_van_out.to_excel(writer, sheet_name='Van_con', index=False)
                print(f"  ✅ Van_con: {len(df_van_con)} dòng")

            # Các sheet chi tiết NVKT_DB
            for nvkt in nvkt_list:
                df_nvkt = df_detail[df_detail['NVKT_DB_NORMALIZED'] == nvkt].copy()
                if 'SA' in df_nvkt.columns:
                    df_nvkt = df_nvkt.sort_values(by='SA').reset_index(drop=True)
                df_nvkt = df_nvkt.drop(columns=['NVKT_DB_NORMALIZED'])
                sheet_name = str(nvkt)[:31]
                df_nvkt.to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"  ✅ Đã tạo {len(nvkt_list)} sheet chi tiết NVKT_DB")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO I1.5")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo I1.5: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # Test hàm
    process_I15_report_with_tracking()
