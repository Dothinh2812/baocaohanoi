import sqlite3
import os
import pandas as pd
from datetime import datetime, date, timedelta
from config import BASE_DATA_PATH


class TicketTracker:
    """
    Hệ thống tracking phiếu nhận và xử lý
    So sánh snapshot giữa các lần tải báo cáo để xác định phiếu mới nhận và phiếu đã xử lý xong
    """

    def __init__(self, db_path=None):
        if db_path is None:
            db_path = os.path.join(BASE_DATA_PATH, 'database', 'brcd.db')
        self.db_path = db_path
        self._ensure_database_exists()
        self._create_tables()

    def _ensure_database_exists(self):
        """Đảm bảo thư mục database tồn tại"""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)

    def _create_tables(self):
        """Tạo các bảng cần thiết nếu chưa tồn tại"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Bảng lưu snapshot mỗi lần tải báo cáo
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS ticket_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            snapshot_time DATETIME NOT NULL,
            ticket_id TEXT NOT NULL,
            ticket_type TEXT NOT NULL,
            ma_tb TEXT,
            doi_vt TEXT,
            nhan_vien TEXT,
            ngay_bh DATETIME,
            chitieu_tg REAL,
            gio_ton_thuc REAL,
            gio_con_lai REAL,
            trang_thai_cong TEXT,
            lydoton TEXT,
            UNIQUE(snapshot_time, ticket_id, ticket_type)
        )
        ''')

        # Bảng lưu lịch sử vòng đời phiếu
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS ticket_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ticket_id TEXT NOT NULL,
            ticket_type TEXT NOT NULL,
            ma_tb TEXT,
            doi_vt TEXT,
            nhan_vien TEXT,
            ngay_bh DATETIME,
            ngay_nhan DATETIME NOT NULL,
            ngay_xu_ly_xong DATETIME,
            so_gio_ton_thuc_te REAL,
            trang_thai TEXT NOT NULL,
            UNIQUE(ticket_id, ticket_type, ngay_nhan)
        )
        ''')

        # Bảng thống kê theo ngày
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_statistics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ngay DATE NOT NULL,
            ticket_type TEXT NOT NULL,
            doi_vt TEXT,
            nhan_vien TEXT,
            so_phieu_nhan_trong_ngay INTEGER DEFAULT 0,
            so_phieu_xu_ly_xong_trong_ngay INTEGER DEFAULT 0,
            so_phieu_ton_dau_ngay INTEGER DEFAULT 0,
            so_phieu_ton_cuoi_ngay INTEGER DEFAULT 0,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ngay, ticket_type, doi_vt, nhan_vien)
        )
        ''')

        # Tạo indexes để tối ưu query
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_snapshots_time ON ticket_snapshots(snapshot_time)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_snapshots_ticket ON ticket_snapshots(ticket_id, ticket_type)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_ticket ON ticket_history(ticket_id, ticket_type)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_dates ON ticket_history(ngay_nhan, ngay_xu_ly_xong)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_daily_stats_date ON daily_statistics(ngay)')

        conn.commit()
        conn.close()
        print("✅ Database tables created successfully")

    def capture_snapshot(self):
        """
        Lưu snapshot của tất cả phiếu hiện tại (BRCD và PTTB)
        Đọc từ file Excel đã xử lý và lưu vào database
        """
        snapshot_time = datetime.now()
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        total_tickets = 0

        # 1. Capture BRCD tickets
        brcd_file = os.path.join(BASE_DATA_PATH, 'downloads', 'kq_dhsc', 'bc_BRCD.xlsx')
        if os.path.exists(brcd_file):
            try:
                df_brcd = pd.read_excel(brcd_file, sheet_name='chi_tiet_ton_brcd')

                for _, row in df_brcd.iterrows():
                    ticket_id = str(row.get('baohong_id', ''))
                    if not ticket_id or ticket_id == 'nan':
                        continue

                    # Parse ngay_bh with proper format
                    ngay_bh = row.get('ngay_bh')
                    if pd.notna(ngay_bh):
                        if isinstance(ngay_bh, str):
                            try:
                                ngay_bh_dt = datetime.strptime(ngay_bh, '%d/%m/%Y %H:%M')
                                ngay_bh = ngay_bh_dt.strftime('%Y-%m-%d %H:%M:%S')
                            except:
                                try:
                                    ngay_bh_dt = datetime.strptime(ngay_bh, '%d/%m/%Y %H:%M:%S')
                                    ngay_bh = ngay_bh_dt.strftime('%Y-%m-%d %H:%M:%S')
                                except:
                                    ngay_bh = None
                        elif isinstance(ngay_bh, (datetime, pd.Timestamp)):
                            # Convert pandas Timestamp or datetime to string
                            ngay_bh = pd.to_datetime(ngay_bh).strftime('%Y-%m-%d %H:%M:%S')
                        else:
                            ngay_bh = None
                    else:
                        ngay_bh = None

                    # Helper function to safely get numeric value
                    def safe_numeric(val):
                        if pd.isna(val):
                            return None
                        try:
                            return float(val)
                        except:
                            return None

                    cursor.execute('''
                    INSERT OR REPLACE INTO ticket_snapshots
                    (snapshot_time, ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien,
                     ngay_bh, chitieu_tg, gio_ton_thuc, gio_con_lai, trang_thai_cong, lydoton)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        snapshot_time,
                        ticket_id,
                        'BRCD',
                        str(row.get('ma_tb', '')) if pd.notna(row.get('ma_tb')) else '',
                        str(row.get('DOI_VT', '')) if pd.notna(row.get('DOI_VT')) else '',
                        str(row.get('ds_nhanvien_th', '')) if pd.notna(row.get('ds_nhanvien_th')) else '',
                        ngay_bh,
                        safe_numeric(row.get('chitieu_tg')),
                        safe_numeric(row.get('thời gian tồn thực')),
                        safe_numeric(row.get('giờ còn lại thực')),
                        str(row.get('Trạng thái cổng', '')) if pd.notna(row.get('Trạng thái cổng')) else '',
                        str(row.get('lydoton', '')) if pd.notna(row.get('lydoton')) else ''
                    ))
                    total_tickets += 1

                print(f"✅ Captured {len(df_brcd)} BRCD tickets")
            except Exception as e:
                print(f"⚠️ Error capturing BRCD snapshot: {e}")

        # 2. Capture PTTB tickets
        pttb_file = os.path.join(BASE_DATA_PATH, 'downloads', 'ton pttb', 'baoCaoPTTB.xlsx')
        if os.path.exists(pttb_file):
            try:
                df_pttb = pd.read_excel(pttb_file, sheet_name='chi_tiet_pttb')

                # Helper function to safely get numeric value
                def safe_numeric(val):
                    if pd.isna(val):
                        return None
                    try:
                        return float(val)
                    except:
                        return None

                for _, row in df_pttb.iterrows():
                    ticket_id = str(row.get('MA_THUE_BAO', ''))
                    if not ticket_id or ticket_id == 'nan':
                        continue

                    # PTTB không có ngay_bh, có thể dùng timestamp hiện tại hoặc field khác
                    ngay_bh = snapshot_time.strftime('%Y-%m-%d %H:%M:%S')

                    cursor.execute('''
                    INSERT OR REPLACE INTO ticket_snapshots
                    (snapshot_time, ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien,
                     ngay_bh, chitieu_tg, gio_ton_thuc, gio_con_lai, lydoton)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        snapshot_time,
                        ticket_id,
                        'PTTB',
                        ticket_id,
                        str(row.get('DOI_VT', '')) if pd.notna(row.get('DOI_VT')) else '',
                        str(row.get('NHANVIEN_TIEPTHI', '')) if pd.notna(row.get('NHANVIEN_TIEPTHI')) else '',
                        ngay_bh,
                        safe_numeric(row.get('chitieu_tg')),
                        safe_numeric(row.get('GIO_TON')),
                        safe_numeric(row.get('gio_conlai')),
                        str(row.get('LYDOTON', '')) if pd.notna(row.get('LYDOTON')) else ''
                    ))
                    total_tickets += 1

                print(f"✅ Captured {len(df_pttb)} PTTB tickets")
            except Exception as e:
                print(f"⚠️ Error capturing PTTB snapshot: {e}")

        conn.commit()
        conn.close()

        print(f"📸 Snapshot completed at {snapshot_time.strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"   Total tickets: {total_tickets}")

        return snapshot_time

    def compare_snapshots_and_update_history(self):
        """
        So sánh snapshot hiện tại với snapshot trước đó
        Cập nhật ticket_history:
        - Phiếu mới xuất hiện → Thêm record mới (ngay_nhan = snapshot_time)
        - Phiếu biến mất → Cập nhật ngay_xu_ly_xong
        - Phiếu vẫn tồn → Không làm gì
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Lấy 2 snapshot gần nhất
        cursor.execute('''
        SELECT DISTINCT snapshot_time
        FROM ticket_snapshots
        ORDER BY snapshot_time DESC
        LIMIT 2
        ''')
        snapshots = cursor.fetchall()

        if len(snapshots) < 2:
            print("⚠️ Chưa đủ dữ liệu để so sánh (cần ít nhất 2 snapshots)")
            conn.close()
            return

        current_snapshot = snapshots[0][0]
        previous_snapshot = snapshots[1][0]

        print(f"🔍 Comparing snapshots:")
        print(f"   Previous: {previous_snapshot}")
        print(f"   Current:  {current_snapshot}")

        # Lấy danh sách phiếu từ 2 snapshots
        cursor.execute('''
        SELECT ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien, ngay_bh, gio_ton_thuc
        FROM ticket_snapshots
        WHERE snapshot_time = ?
        ''', (current_snapshot,))
        current_tickets = {(row[0], row[1]): row for row in cursor.fetchall()}

        cursor.execute('''
        SELECT ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien, ngay_bh, gio_ton_thuc
        FROM ticket_snapshots
        WHERE snapshot_time = ?
        ''', (previous_snapshot,))
        previous_tickets = {(row[0], row[1]): row for row in cursor.fetchall()}

        new_tickets = 0
        completed_tickets = 0

        # 1. Tìm phiếu MỚI (có trong current nhưng không có trong previous)
        for ticket_key, ticket_data in current_tickets.items():
            if ticket_key not in previous_tickets:
                ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien, ngay_bh, gio_ton_thuc = ticket_data

                cursor.execute('''
                INSERT OR IGNORE INTO ticket_history
                (ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien, ngay_bh, ngay_nhan, trang_thai)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien, ngay_bh, current_snapshot, 'DANG_XU_LY'))

                new_tickets += 1

        # 2. Tìm phiếu ĐÃ XỬ LÝ XONG (có trong previous nhưng không có trong current)
        for ticket_key, ticket_data in previous_tickets.items():
            if ticket_key not in current_tickets:
                ticket_id, ticket_type, ma_tb, doi_vt, nhan_vien, ngay_bh, gio_ton_thuc = ticket_data

                # Tìm record trong history chưa có ngay_xu_ly_xong
                cursor.execute('''
                SELECT id, ngay_nhan FROM ticket_history
                WHERE ticket_id = ? AND ticket_type = ? AND ngay_xu_ly_xong IS NULL
                ORDER BY ngay_nhan DESC LIMIT 1
                ''', (ticket_id, ticket_type))

                history_record = cursor.fetchone()
                if history_record:
                    history_id, ngay_nhan = history_record

                    # Tính số giờ tồn thực tế
                    if ngay_nhan:
                        ngay_nhan_dt = datetime.fromisoformat(ngay_nhan) if isinstance(ngay_nhan, str) else ngay_nhan
                        current_snapshot_dt = datetime.fromisoformat(current_snapshot) if isinstance(current_snapshot, str) else current_snapshot
                        so_gio_ton = (current_snapshot_dt - ngay_nhan_dt).total_seconds() / 3600
                    else:
                        so_gio_ton = None

                    cursor.execute('''
                    UPDATE ticket_history
                    SET ngay_xu_ly_xong = ?,
                        so_gio_ton_thuc_te = ?,
                        trang_thai = 'DA_XONG'
                    WHERE id = ?
                    ''', (current_snapshot, so_gio_ton, history_id))

                    completed_tickets += 1

        conn.commit()
        conn.close()

        print(f"📊 Comparison results:")
        print(f"   ➕ New tickets: {new_tickets}")
        print(f"   ✅ Completed tickets: {completed_tickets}")
        print(f"   🔄 Continuing tickets: {len(current_tickets) - new_tickets}")

        return new_tickets, completed_tickets

    def calculate_daily_statistics(self, target_date=None):
        """
        Tính toán thống kê cho một ngày cụ thể
        Nếu không truyền target_date, mặc định là hôm nay
        """
        if target_date is None:
            target_date = date.today()
        elif isinstance(target_date, str):
            target_date = datetime.strptime(target_date, '%Y-%m-%d').date()

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Xóa thống kê cũ cho ngày này (để tính lại)
        cursor.execute('DELETE FROM daily_statistics WHERE ngay = ?', (target_date,))

        # Query để lấy tất cả combination của ticket_type, doi_vt, nhan_vien có data
        cursor.execute('''
        SELECT DISTINCT ticket_type, doi_vt, nhan_vien
        FROM ticket_history
        WHERE DATE(ngay_nhan) <= ? OR DATE(ngay_xu_ly_xong) = ?
        ''', (target_date, target_date))

        combinations = cursor.fetchall()

        for ticket_type, doi_vt, nhan_vien in combinations:
            # 1. Số phiếu nhận trong ngày
            cursor.execute('''
            SELECT COUNT(*) FROM ticket_history
            WHERE ticket_type = ? AND doi_vt = ? AND nhan_vien = ?
            AND DATE(ngay_nhan) = ?
            ''', (ticket_type, doi_vt, nhan_vien, target_date))
            so_phieu_nhan = cursor.fetchone()[0]

            # 2. Số phiếu xử lý xong trong ngày
            cursor.execute('''
            SELECT COUNT(*) FROM ticket_history
            WHERE ticket_type = ? AND doi_vt = ? AND nhan_vien = ?
            AND DATE(ngay_xu_ly_xong) = ?
            ''', (ticket_type, doi_vt, nhan_vien, target_date))
            so_phieu_xu_ly_xong = cursor.fetchone()[0]

            # 3. Tồn đầu ngày (đã nhận trước ngày này và chưa xử lý xong hoặc xử lý xong sau ngày này)
            cursor.execute('''
            SELECT COUNT(*) FROM ticket_history
            WHERE ticket_type = ? AND doi_vt = ? AND nhan_vien = ?
            AND DATE(ngay_nhan) < ?
            AND (ngay_xu_ly_xong IS NULL OR DATE(ngay_xu_ly_xong) >= ?)
            ''', (ticket_type, doi_vt, nhan_vien, target_date, target_date))
            so_phieu_ton_dau = cursor.fetchone()[0]

            # 4. Tồn cuối ngày
            so_phieu_ton_cuoi = so_phieu_ton_dau + so_phieu_nhan - so_phieu_xu_ly_xong

            # Insert vào bảng statistics
            cursor.execute('''
            INSERT OR REPLACE INTO daily_statistics
            (ngay, ticket_type, doi_vt, nhan_vien,
             so_phieu_nhan_trong_ngay, so_phieu_xu_ly_xong_trong_ngay,
             so_phieu_ton_dau_ngay, so_phieu_ton_cuoi_ngay, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (target_date, ticket_type, doi_vt, nhan_vien,
                  so_phieu_nhan, so_phieu_xu_ly_xong,
                  so_phieu_ton_dau, so_phieu_ton_cuoi,
                  datetime.now()))

        conn.commit()
        conn.close()

        print(f"📈 Daily statistics calculated for {target_date}")
        print(f"   Processed {len(combinations)} combinations")

        return True

    def get_statistics_report(self, start_date=None, end_date=None,
                             ticket_type=None, doi_vt=None, nhan_vien=None,
                             group_by='day'):
        """
        Lấy báo cáo thống kê với các filter

        Parameters:
        - start_date, end_date: Khoảng thời gian (string 'YYYY-MM-DD' hoặc date object)
        - ticket_type: 'BRCD', 'PTTB', hoặc None (tất cả)
        - doi_vt: Tên đội, hoặc None (tất cả)
        - nhan_vien: Tên nhân viên, hoặc None (tất cả)
        - group_by: 'day', 'week', 'month'

        Returns: pandas DataFrame
        """
        conn = sqlite3.connect(self.db_path)

        # Build query dynamically
        where_clauses = []
        params = []

        if start_date:
            where_clauses.append('ngay >= ?')
            params.append(start_date)

        if end_date:
            where_clauses.append('ngay <= ?')
            params.append(end_date)

        if ticket_type:
            where_clauses.append('ticket_type = ?')
            params.append(ticket_type)

        if doi_vt:
            where_clauses.append('doi_vt = ?')
            params.append(doi_vt)

        if nhan_vien:
            where_clauses.append('nhan_vien = ?')
            params.append(nhan_vien)

        where_sql = ' AND '.join(where_clauses) if where_clauses else '1=1'

        query = f'''
        SELECT
            ngay,
            ticket_type,
            doi_vt,
            nhan_vien,
            SUM(so_phieu_nhan_trong_ngay) as so_phieu_nhan,
            SUM(so_phieu_xu_ly_xong_trong_ngay) as so_phieu_xu_ly_xong,
            AVG(so_phieu_ton_dau_ngay) as so_phieu_ton_dau,
            AVG(so_phieu_ton_cuoi_ngay) as so_phieu_ton_cuoi
        FROM daily_statistics
        WHERE {where_sql}
        GROUP BY ngay, ticket_type, doi_vt, nhan_vien
        ORDER BY ngay DESC
        '''

        df = pd.read_sql_query(query, conn, params=params)
        conn.close()

        # Group by week or month if requested
        if group_by == 'week' and not df.empty:
            df['ngay'] = pd.to_datetime(df['ngay'])
            df['week'] = df['ngay'].dt.to_period('W')
            df = df.groupby(['week', 'ticket_type', 'doi_vt', 'nhan_vien']).agg({
                'so_phieu_nhan': 'sum',
                'so_phieu_xu_ly_xong': 'sum',
                'so_phieu_ton_dau': 'mean',
                'so_phieu_ton_cuoi': 'mean'
            }).reset_index()

        elif group_by == 'month' and not df.empty:
            df['ngay'] = pd.to_datetime(df['ngay'])
            df['month'] = df['ngay'].dt.to_period('M')
            df = df.groupby(['month', 'ticket_type', 'doi_vt', 'nhan_vien']).agg({
                'so_phieu_nhan': 'sum',
                'so_phieu_xu_ly_xong': 'sum',
                'so_phieu_ton_dau': 'mean',
                'so_phieu_ton_cuoi': 'mean'
            }).reset_index()

        return df

    def cleanup_old_snapshots(self, days_to_keep=90):
        """
        Xóa các snapshot cũ hơn X ngày để tiết kiệm dung lượng
        Giữ lại ticket_history và daily_statistics
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cutoff_date = datetime.now() - timedelta(days=days_to_keep)

        cursor.execute('DELETE FROM ticket_snapshots WHERE snapshot_time < ?', (cutoff_date,))
        deleted_count = cursor.rowcount

        conn.commit()
        conn.close()

        print(f"🧹 Cleaned up {deleted_count} old snapshot records (older than {days_to_keep} days)")

        return deleted_count


# Convenience functions for easy import
def capture_snapshot():
    """Wrapper function để dễ import"""
    tracker = TicketTracker()
    return tracker.capture_snapshot()


def compare_snapshots_and_update_history():
    """Wrapper function để dễ import"""
    tracker = TicketTracker()
    return tracker.compare_snapshots_and_update_history()


def calculate_daily_statistics(target_date=None):
    """Wrapper function để dễ import"""
    tracker = TicketTracker()
    return tracker.calculate_daily_statistics(target_date)


def get_statistics_report(**kwargs):
    """Wrapper function để dễ import"""
    tracker = TicketTracker()
    return tracker.get_statistics_report(**kwargs)


if __name__ == "__main__":
    # Test run
    print("=== Testing Ticket Tracking System ===")

    tracker = TicketTracker()

    # Capture snapshot
    print("\n1. Capturing snapshot...")
    tracker.capture_snapshot()

    # Compare (will only work if there's a previous snapshot)
    print("\n2. Comparing snapshots...")
    try:
        tracker.compare_snapshots_and_update_history()
    except Exception as e:
        print(f"Note: {e}")

    # Calculate statistics
    print("\n3. Calculating daily statistics...")
    tracker.calculate_daily_statistics()

    # Get report
    print("\n4. Getting statistics report...")
    df = tracker.get_statistics_report(
        start_date=(date.today() - timedelta(days=7)).isoformat(),
        end_date=date.today().isoformat()
    )
    print(df.head() if not df.empty else "No data yet")

    print("\n✅ Test completed!")
