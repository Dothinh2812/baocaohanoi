# -*- coding: utf-8 -*-
"""
Module chứa các hàm xử lý báo cáo thực tăng PTTB và MyTV
"""
import os
import re
from datetime import datetime
import pandas as pd

# Import module để lưu vào database history
try:
    from import_reports_history import ReportsHistoryImporter
    HISTORY_IMPORT_AVAILABLE = True
except ImportError:
    HISTORY_IMPORT_AVAILABLE = False


def _extract_son_tay_ngung_psc_subset(df_raw):
    """
    Trích xuất 4 dòng dữ liệu Sơn Tây từ báo cáo nhiều tầng header.
    Dò cột theo nhãn thực tế để tránh phụ thuộc vào chỉ số cột cố định.
    """
    data_row_indices = [3, 4, 5, 6]
    if len(df_raw) <= max(data_row_indices):
        raise ValueError(
            f"File Sơn Tây không đủ số dòng dữ liệu. Cần ít nhất {max(data_row_indices) + 1} dòng, "
            f"nhưng chỉ có {len(df_raw)} dòng."
        )

    header_row_count = min(3, len(df_raw))
    header_map = {}
    for col_idx in range(df_raw.shape[1]):
        header_values = [
            str(df_raw.iat[row_idx, col_idx]).strip()
            for row_idx in range(header_row_count)
            if pd.notna(df_raw.iat[row_idx, col_idx]) and str(df_raw.iat[row_idx, col_idx]).strip()
        ]
        for value in header_values:
            header_map.setdefault(value, col_idx)

    required_columns = {
        "Đơn vị/Nhân viên KT": "Đơn vị/Nhân viên KT",
        "Hoàn công(*) (1.5)": "Hoàn công(*) (1.5)",
        "Lũy kế tháng(1.6)": "Lũy kế tháng(1.6)",
        "Lũy kế năm(1.7)": "Lũy kế năm(1.7)",
        "Ngưng PSC tạm tính tháng T(5.1)": "Ngưng PSC tạm tính tháng T(5.1)",
        "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)": "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)",
    }
    missing_headers = [header for header in required_columns if header not in header_map]
    if missing_headers:
        raise ValueError(
            "Không tìm thấy các cột bắt buộc trong báo cáo Sơn Tây: "
            + ", ".join(missing_headers)
        )

    selected_col_indices = [header_map[header] for header in required_columns]
    df_subset = df_raw.iloc[data_row_indices, selected_col_indices].copy()
    df_subset.columns = list(required_columns.values())
    return df_subset

def process_ngung_psc_report():
    """
    Xử lý báo cáo ngung_psc:
    1. Đọc file ngung_psc_DDMMYYYY.xlsx
    2. Chuẩn hóa cột TEN_KV (giữ phần sau dấu -, bỏ ngoặc đơn)
    3. Tạo sheet thống kê fiber_ngung_psc_thang và fiber_ngung_psc_thang_theo_to
    """
    print("\n=== Bắt đầu xử lý báo cáo Ngưng PSC ===")

    try:
        # Lấy ngày hiện tại để tìm file
        date_str = datetime.now().strftime("%d%m%Y")
        download_dir = "PTTB-PSC"
        ngung_psc_file = os.path.join(download_dir, f"ngung_psc_{date_str}.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(ngung_psc_file):
            print(f"❌ Không tìm thấy file: {ngung_psc_file}")
            return

        print(f"Đang đọc file: {ngung_psc_file}")
        df_ngung_psc = pd.read_excel(ngung_psc_file)
        print(f"✅ Đọc thành công {len(df_ngung_psc)} dòng dữ liệu")
        print(f"   Các cột: {', '.join(df_ngung_psc.columns)}")

        # Kiểm tra cột TEN_KV tồn tại
        if 'TEN_KV' not in df_ngung_psc.columns:
            print(f"❌ Không tìm thấy cột 'TEN_KV' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_ngung_psc.columns)}")
            return

        # Kiểm tra cột TEN_DOI tồn tại
        if 'TEN_DOI' not in df_ngung_psc.columns:
            print(f"❌ Không tìm thấy cột 'TEN_DOI' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_ngung_psc.columns)}")
            return

        # Chuẩn hóa cột TEN_KV: giữ phần sau dấu -, bỏ ngoặc đơn
        # VD: PCT1-Nguyễn Mạnh Hùng(PTO) -> Nguyễn Mạnh Hùng
        # VD: Tây Đằng 05 -Lê Quyết Tiến -> Lê Quyết Tiến
        print("Đang chuẩn hóa cột TEN_KV...")

        def normalize_ten_kv(name):
            if pd.isna(name) or str(name).strip() == '':
                return ''
            name = str(name).strip()
            # Lấy phần sau dấu -
            if '-' in name:
                name = name.split('-', 1)[1].strip()
            # Bỏ phần trong ngoặc đơn
            name = re.sub(r'\([^)]*\)', '', name).strip()
            return name

        df_ngung_psc['TEN_KV'] = df_ngung_psc['TEN_KV'].apply(normalize_ten_kv)
        print(f"✅ Đã chuẩn hóa {len(df_ngung_psc)} tên nhân viên")

        # Tạo sheet thống kê theo NVKT: TEN_DOI, TEN_KV, SỐ LƯỢNG NGƯNG PSC THÁNG
        print("\n📊 Đang tạo thống kê ngưng PSC theo NVKT...")
        df_stats = df_ngung_psc.groupby(['TEN_DOI', 'TEN_KV'], dropna=False).size().reset_index(name='SỐ LƯỢNG NGƯNG PSC THÁNG')
        df_stats = df_stats.fillna('(Chưa xác định)')
        df_stats = df_stats.sort_values(['TEN_DOI', 'SỐ LƯỢNG NGƯNG PSC THÁNG'], ascending=[True, False])

        # Thêm dòng TỔNG CỘNG
        total_row = pd.DataFrame([{
            'TEN_DOI': 'TỔNG CỘNG',
            'TEN_KV': '',
            'SỐ LƯỢNG NGƯNG PSC THÁNG': int(df_stats['SỐ LƯỢNG NGƯNG PSC THÁNG'].sum())
        }])
        df_stats = pd.concat([df_stats, total_row], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats) - 1} NVKT")

        # Tạo sheet thống kê theo Đội VT: TEN_DOI, SỐ LƯỢNG NGƯNG PSC THÁNG
        print("📊 Đang tạo thống kê ngưng PSC theo Đội VT...")
        df_stats_to = df_ngung_psc.groupby('TEN_DOI', dropna=False).size().reset_index(name='SỐ LƯỢNG NGƯNG PSC THÁNG')
        df_stats_to = df_stats_to.fillna('(Chưa xác định)')
        df_stats_to = df_stats_to.sort_values('SỐ LƯỢNG NGƯNG PSC THÁNG', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_row_to = pd.DataFrame([{
            'TEN_DOI': 'TỔNG CỘNG',
            'SỐ LƯỢNG NGƯNG PSC THÁNG': int(df_stats_to['SỐ LƯỢNG NGƯNG PSC THÁNG'].sum())
        }])
        df_stats_to = pd.concat([df_stats_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats_to) - 1} Đội VT")

        # Lưu file
        print(f"\n💾 Đang lưu file...")
        with pd.ExcelWriter(ngung_psc_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu gốc
            df_ngung_psc.to_excel(writer, sheet_name='Data', index=False)

            # Sheet 2: Thống kê ngưng PSC tháng theo NVKT
            df_stats.to_excel(writer, sheet_name='fiber_ngung_psc_thang', index=False)

            # Sheet 3: Thống kê ngưng PSC tháng theo Đội VT
            df_stats_to.to_excel(writer, sheet_name='fiber_ngung_psc_thang_theo_to', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Đã lưu file: {ngung_psc_file}")
        print(f"   - Sheet 'Data': Dữ liệu đầy đủ ({len(df_ngung_psc)} dòng)")
        print(f"   - Sheet 'fiber_ngung_psc_thang': Thống kê theo NVKT ({len(df_stats)} dòng)")
        print(f"   - Sheet 'fiber_ngung_psc_thang_theo_to': Thống kê theo Đội VT ({len(df_stats_to)} dòng)")

        # Hiển thị top 5 NVKT
        print("\n📊 Top 5 NVKT có nhiều TB ngưng PSC nhất:")
        top5 = df_stats[df_stats['TEN_DOI'] != 'TỔNG CỘNG'].nlargest(5, 'SỐ LƯỢNG NGƯNG PSC THÁNG')
        for idx, (_, row) in enumerate(top5.iterrows(), 1):
            print(f"   {idx}. {row['TEN_KV']} ({row['TEN_DOI']}): {row['SỐ LƯỢNG NGƯNG PSC THÁNG']} TB")

        print("\n✅ Hoàn thành xử lý báo cáo Ngưng PSC!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo: {e}")
        import traceback
        traceback.print_exc()

def process_hoan_cong_report():
    """
    Xử lý báo cáo hoan_cong:
    1. Đọc file hoan_cong_DDMMYYYY.xlsx
    2. Chuẩn hóa cột NHANVIEN_KT (giữ phần sau dấu -)
    3. Tạo sheet thống kê fiber_hoan_cong_thang gồm: DOIVT, NHANVIEN_KT, SỐ LƯỢNG HOÀN CÔNG THÁNG
    """
    print("\n=== Bắt đầu xử lý báo cáo Hoàn công ===")

    try:
        # Lấy ngày hiện tại để tìm file
        date_str = datetime.now().strftime("%d%m%Y")
        download_dir = "PTTB-PSC"
        hoan_cong_file = os.path.join(download_dir, f"hoan_cong_{date_str}.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(hoan_cong_file):
            print(f"❌ Không tìm thấy file: {hoan_cong_file}")
            return

        print(f"Đang đọc file: {hoan_cong_file}")
        df_hoan_cong = pd.read_excel(hoan_cong_file)
        print(f"✅ Đọc thành công {len(df_hoan_cong)} dòng dữ liệu")
        print(f"   Các cột: {', '.join(df_hoan_cong.columns)}")

        # Kiểm tra cột NHANVIEN_KT tồn tại
        if 'NHANVIEN_KT' not in df_hoan_cong.columns:
            print(f"❌ Không tìm thấy cột 'NHANVIEN_KT' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_hoan_cong.columns)}")
            return

        # Kiểm tra cột DOIVT tồn tại
        if 'DOIVT' not in df_hoan_cong.columns:
            print(f"❌ Không tìm thấy cột 'DOIVT' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_hoan_cong.columns)}")
            return

        # Chuẩn hóa cột NHANVIEN_KT: giữ phần sau dấu -
        # VD: VNPT016768-Trịnh Thanh Quang -> Trịnh Thanh Quang
        print("Đang chuẩn hóa cột NHANVIEN_KT...")
        df_hoan_cong['NHANVIEN_KT'] = df_hoan_cong['NHANVIEN_KT'].apply(
            lambda x: str(x).split('-', 1)[1].strip() if pd.notna(x) and '-' in str(x) else (str(x).strip() if pd.notna(x) else '')
        )
        print(f"✅ Đã chuẩn hóa {len(df_hoan_cong)} tên nhân viên")

        # Tạo sheet thống kê: DOIVT, NHANVIEN_KT, SỐ LƯỢNG HOÀN CÔNG THÁNG
        print("\n📊 Đang tạo thống kê hoàn công theo NVKT...")
        df_stats = df_hoan_cong.groupby(['DOIVT', 'NHANVIEN_KT'], dropna=False).size().reset_index(name='SỐ LƯỢNG HOÀN CÔNG THÁNG')
        df_stats = df_stats.fillna('(Chưa xác định)')
        df_stats = df_stats.sort_values(['DOIVT', 'SỐ LƯỢNG HOÀN CÔNG THÁNG'], ascending=[True, False])

        # Thêm dòng TỔNG CỘNG
        total_row = pd.DataFrame([{
            'DOIVT': 'TỔNG CỘNG',
            'NHANVIEN_KT': '',
            'SỐ LƯỢNG HOÀN CÔNG THÁNG': int(df_stats['SỐ LƯỢNG HOÀN CÔNG THÁNG'].sum())
        }])
        df_stats = pd.concat([df_stats, total_row], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats) - 1} NVKT")

        # Tạo sheet thống kê theo Đội VT: DOIVT, SỐ LƯỢNG HOÀN CÔNG THÁNG
        print("📊 Đang tạo thống kê hoàn công theo Đội VT...")
        df_stats_to = df_hoan_cong.groupby('DOIVT', dropna=False).size().reset_index(name='SỐ LƯỢNG HOÀN CÔNG THÁNG')
        df_stats_to = df_stats_to.fillna('(Chưa xác định)')
        df_stats_to = df_stats_to.sort_values('SỐ LƯỢNG HOÀN CÔNG THÁNG', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_row_to = pd.DataFrame([{
            'DOIVT': 'TỔNG CỘNG',
            'SỐ LƯỢNG HOÀN CÔNG THÁNG': int(df_stats_to['SỐ LƯỢNG HOÀN CÔNG THÁNG'].sum())
        }])
        df_stats_to = pd.concat([df_stats_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats_to) - 1} Đội VT")

        # Lưu file
        print(f"\n💾 Đang lưu file...")
        with pd.ExcelWriter(hoan_cong_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu gốc
            df_hoan_cong.to_excel(writer, sheet_name='Data', index=False)

            # Sheet 2: Thống kê hoàn công tháng theo NVKT
            df_stats.to_excel(writer, sheet_name='fiber_hoan_cong_thang', index=False)

            # Sheet 3: Thống kê hoàn công tháng theo Đội VT
            df_stats_to.to_excel(writer, sheet_name='fiber_hoan_cong_thang_theo_to', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Đã lưu file: {hoan_cong_file}")
        print(f"   - Sheet 'Data': Dữ liệu đầy đủ ({len(df_hoan_cong)} dòng)")
        print(f"   - Sheet 'fiber_hoan_cong_thang': Thống kê theo NVKT ({len(df_stats)} dòng)")
        print(f"   - Sheet 'fiber_hoan_cong_thang_theo_to': Thống kê theo Đội VT ({len(df_stats_to)} dòng)")

        # Hiển thị top 5 NVKT
        print("\n📊 Top 5 NVKT có nhiều TB hoàn công nhất:")
        top5 = df_stats[df_stats['DOIVT'] != 'TỔNG CỘNG'].nlargest(5, 'SỐ LƯỢNG HOÀN CÔNG THÁNG')
        for idx, (_, row) in enumerate(top5.iterrows(), 1):
            print(f"   {idx}. {row['NHANVIEN_KT']} ({row['DOIVT']}): {row['SỐ LƯỢNG HOÀN CÔNG THÁNG']} TB")

        print("\n✅ Hoàn thành xử lý báo cáo Hoàn công!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo: {e}")
        import traceback
        traceback.print_exc()

def create_thuc_tang_report():
    """
    Tạo báo cáo thực tăng từ 2 báo cáo đã xử lý (Hoàn công và Ngưng PSC)
    Thực tăng = Hoàn công - Ngưng phát sinh cước

    Tạo 2 sheet:
    1. thuc_tang_theo_to: Thống kê theo Đội VT (TEN_DOI)
    2. thuc_tang_theo_NVKT: Thống kê theo NVKT
    """
    print("\n=== Bắt đầu tạo báo cáo Thực tăng ===")

    try:
        # Lấy ngày hiện tại để tìm file
        date_str = datetime.now().strftime("%d%m%Y")
        download_dir = "PTTB-PSC"

        ngung_psc_file = os.path.join(download_dir, f"ngung_psc_{date_str}.xlsx")
        hoan_cong_file = os.path.join(download_dir, f"hoan_cong_{date_str}.xlsx")
        thuc_tang_file = os.path.join(download_dir, f"fiber_thuc_tang_{date_str}.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(ngung_psc_file):
            print(f"❌ Không tìm thấy file: {ngung_psc_file}")
            return

        if not os.path.exists(hoan_cong_file):
            print(f"❌ Không tìm thấy file: {hoan_cong_file}")
            return

        print(f"Đang đọc dữ liệu từ file Ngưng PSC...")
        # Đọc sheet thống kê từ file Ngưng PSC
        df_ngung_psc_to = pd.read_excel(ngung_psc_file, sheet_name='fiber_ngung_psc_thang_theo_to')
        df_ngung_psc_nvkt = pd.read_excel(ngung_psc_file, sheet_name='fiber_ngung_psc_thang')

        print(f"Đang đọc dữ liệu từ file Hoàn công...")
        # Đọc sheet thống kê từ file Hoàn công
        df_hoan_cong_to = pd.read_excel(hoan_cong_file, sheet_name='fiber_hoan_cong_thang_theo_to')
        df_hoan_cong_nvkt = pd.read_excel(hoan_cong_file, sheet_name='fiber_hoan_cong_thang')

        # === XỬ LÝ SHEET 1: THỰC TĂNG THEO ĐỘI VT ===
        print("\n📊 Đang tạo báo cáo Thực tăng theo Đội VT...")

        # Loại bỏ dòng TỔNG CỘNG trước khi merge
        df_ngung_psc_to_clean = df_ngung_psc_to[df_ngung_psc_to['TEN_DOI'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_to_clean = df_hoan_cong_to[df_hoan_cong_to['DOIVT'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột để thống nhất và phân biệt
        df_ngung_psc_to_clean.rename(columns={'TEN_DOI': 'Đội VT', 'SỐ LƯỢNG NGƯNG PSC THÁNG': 'Ngưng phát sinh cước'}, inplace=True)
        df_hoan_cong_to_clean.rename(columns={'DOIVT': 'Đội VT', 'SỐ LƯỢNG HOÀN CÔNG THÁNG': 'Hoàn công'}, inplace=True)

        # Merge 2 dataframe theo Đội VT
        df_thuc_tang_to = pd.merge(
            df_hoan_cong_to_clean,
            df_ngung_psc_to_clean,
            on='Đội VT',
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Hoàn công'] - df_thuc_tang_to['Ngưng phát sinh cước']

        # Tính Tỷ lệ ngưng/psc (%)
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to.apply(
            lambda row: (row['Ngưng phát sinh cước'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int cho các cột số lượng
        df_thuc_tang_to['Hoàn công'] = df_thuc_tang_to['Hoàn công'].astype(int)
        df_thuc_tang_to['Ngưng phát sinh cước'] = df_thuc_tang_to['Ngưng phát sinh cước'].astype(int)
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Thực tăng'].astype(int)
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần
        df_thuc_tang_to = df_thuc_tang_to.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc = int(df_thuc_tang_to['Hoàn công'].sum())
        total_np = int(df_thuc_tang_to['Ngưng phát sinh cước'].sum())
        total_tt = int(df_thuc_tang_to['Thực tăng'].sum())
        total_ty_le = (total_np / total_hc * 100) if total_hc != 0 else 0

        total_row_to = pd.DataFrame([{
            'Đội VT': 'TỔNG CỘNG',
            'Hoàn công': total_hc,
            'Ngưng phát sinh cước': total_np,
            'Thực tăng': total_tt,
            'Tỷ lệ ngưng/psc': round(total_ty_le, 2)
        }])
        df_thuc_tang_to = pd.concat([df_thuc_tang_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_to) - 1} Đội VT")

        # === XỬ LÝ SHEET 2: THỰC TĂNG THEO NVKT ===
        print("📊 Đang tạo báo cáo Thực tăng theo NVKT...")

        # Loại bỏ dòng TỔNG CỘNG
        df_ngung_psc_nvkt_clean = df_ngung_psc_nvkt[df_ngung_psc_nvkt['TEN_DOI'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_nvkt_clean = df_hoan_cong_nvkt[df_hoan_cong_nvkt['DOIVT'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột
        df_ngung_psc_nvkt_clean.rename(columns={
            'TEN_DOI': 'Đội VT',
            'TEN_KV': 'NVKT',
            'SỐ LƯỢNG NGƯNG PSC THÁNG': 'Ngưng phát sinh cước'
        }, inplace=True)
        df_hoan_cong_nvkt_clean.rename(columns={
            'DOIVT': 'Đội VT',
            'NHANVIEN_KT': 'NVKT',
            'SỐ LƯỢNG HOÀN CÔNG THÁNG': 'Hoàn công'
        }, inplace=True)

        # Merge theo Đội VT và NVKT
        df_thuc_tang_nvkt = pd.merge(
            df_hoan_cong_nvkt_clean,
            df_ngung_psc_nvkt_clean,
            on=['Đội VT', 'NVKT'],
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Hoàn công'] - df_thuc_tang_nvkt['Ngưng phát sinh cước']

        # Tính Tỷ lệ ngưng/psc (%)
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt.apply(
            lambda row: (row['Ngưng phát sinh cước'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int
        df_thuc_tang_nvkt['Hoàn công'] = df_thuc_tang_nvkt['Hoàn công'].astype(int)
        df_thuc_tang_nvkt['Ngưng phát sinh cước'] = df_thuc_tang_nvkt['Ngưng phát sinh cước'].astype(int)
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Thực tăng'].astype(int)
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần
        df_thuc_tang_nvkt = df_thuc_tang_nvkt.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc_nvkt = int(df_thuc_tang_nvkt['Hoàn công'].sum())
        total_np_nvkt = int(df_thuc_tang_nvkt['Ngưng phát sinh cước'].sum())
        total_tt_nvkt = int(df_thuc_tang_nvkt['Thực tăng'].sum())
        total_ty_le_nvkt = (total_np_nvkt / total_hc_nvkt * 100) if total_hc_nvkt != 0 else 0

        total_row_nvkt = pd.DataFrame([{
            'Đội VT': 'TỔNG CỘNG',
            'NVKT': '',
            'Hoàn công': total_hc_nvkt,
            'Ngưng phát sinh cước': total_np_nvkt,
            'Thực tăng': total_tt_nvkt,
            'Tỷ lệ ngưng/psc': round(total_ty_le_nvkt, 2)
        }])
        df_thuc_tang_nvkt = pd.concat([df_thuc_tang_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_nvkt) - 1} NVKT")

        # === LƯU FILE ===
        print(f"\n💾 Đang lưu file báo cáo Thực tăng...")
        with pd.ExcelWriter(thuc_tang_file, engine='openpyxl') as writer:
            # Sheet 1: Thống kê theo Đội VT
            df_thuc_tang_to.to_excel(writer, sheet_name='thuc_tang_theo_to', index=False)

            # Sheet 2: Thống kê theo NVKT
            df_thuc_tang_nvkt.to_excel(writer, sheet_name='thuc_tang_theo_NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Đã lưu file: {thuc_tang_file}")
        print(f"   - Sheet 'thuc_tang_theo_to': Thống kê theo Đội VT ({len(df_thuc_tang_to)} dòng)")
        print(f"   - Sheet 'thuc_tang_theo_NVKT': Thống kê theo NVKT ({len(df_thuc_tang_nvkt)} dòng)")

        # Hiển thị thống kê tổng quan
        total_hoan_cong = int(df_thuc_tang_to[df_thuc_tang_to['Đội VT'] == 'TỔNG CỘNG']['Hoàn công'].iloc[0])
        total_ngung_psc = int(df_thuc_tang_to[df_thuc_tang_to['Đội VT'] == 'TỔNG CỘNG']['Ngưng phát sinh cước'].iloc[0])
        total_thuc_tang = int(df_thuc_tang_to[df_thuc_tang_to['Đội VT'] == 'TỔNG CỘNG']['Thực tăng'].iloc[0])

        print(f"\n📊 Tổng quan:")
        print(f"   - Tổng Hoàn công: {total_hoan_cong} TB")
        print(f"   - Tổng Ngưng phát sinh cước: {total_ngung_psc} TB")
        print(f"   - Thực tăng: {total_thuc_tang} TB")

        # Top 5 Đội VT có thực tăng cao nhất
        print(f"\n📊 Top 5 Đội VT có Thực tăng cao nhất:")
        top5_to = df_thuc_tang_to[df_thuc_tang_to['Đội VT'] != 'TỔNG CỘNG'].head(5)
        for idx, (_, row) in enumerate(top5_to.iterrows(), 1):
            doi_vt = row['Đội VT'] if pd.notna(row['Đội VT']) else '(Chưa xác định)'
            print(f"   {idx}. {doi_vt}: {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng phát sinh cước']})")

        # Top 5 NVKT có thực tăng cao nhất
        print(f"\n📊 Top 5 NVKT có Thực tăng cao nhất:")
        top5_nvkt = df_thuc_tang_nvkt[df_thuc_tang_nvkt['NVKT'] != ''].nlargest(5, 'Thực tăng')
        for idx, (_, row) in enumerate(top5_nvkt.iterrows(), 1):
            nvkt = row['NVKT'] if pd.notna(row['NVKT']) else '(Chưa xác định)'
            doi_vt = row['Đội VT'] if pd.notna(row['Đội VT']) else '(Chưa xác định)'
            print(f"   {idx}. {nvkt} ({doi_vt}): {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng phát sinh cước']})")

        # Import vào database history
        if HISTORY_IMPORT_AVAILABLE:
            try:
                print(f"\n💾 Đang lưu vào database history...")
                importer = ReportsHistoryImporter()
                importer.import_growth_pttb()
                print(f"✅ Đã lưu vào database history")
            except Exception as e:
                print(f"⚠️  Không thể lưu vào database history: {e}")

        print("\n✅ Hoàn thành tạo báo cáo Thực tăng!")

    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo Thực tăng: {e}")
        import traceback
        traceback.print_exc()



def process_mytv_ngung_psc_report():
    """
    Xử lý báo cáo mytv_ngung_psc:
    1. Đọc file mytv_ngung_psc_DDMMYYYY.xlsx
    2. Chuẩn hóa cột TEN_KV (giữ phần sau dấu -, bỏ ngoặc đơn)
    3. Tạo sheet thống kê mytv_ngung_psc_thang và mytv_ngung_psc_thang_theo_to
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO MYTV NGƯNG PSC")
        print("="*80)

        # === ĐỌC FILE ===
        download_dir = "PTTB-PSC"
        date_str = datetime.now().strftime("%d%m%Y")
        ngung_psc_file = os.path.join(download_dir, f"mytv_ngung_psc_{date_str}.xlsx")

        if not os.path.exists(ngung_psc_file):
            print(f"❌ Không tìm thấy file: {ngung_psc_file}")
            return

        print(f"📂 Đang đọc file: {ngung_psc_file}")
        df_ngung_psc = pd.read_excel(ngung_psc_file)
        print(f"✅ Đọc thành công {len(df_ngung_psc)} dòng dữ liệu")
        print(f"   Các cột: {', '.join(df_ngung_psc.columns)}")

        # Kiểm tra cột TEN_KV tồn tại
        if 'TEN_KV' not in df_ngung_psc.columns:
            print(f"❌ Không tìm thấy cột 'TEN_KV' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_ngung_psc.columns)}")
            return

        # Kiểm tra cột TEN_DOI tồn tại
        if 'TEN_DOI' not in df_ngung_psc.columns:
            print(f"❌ Không tìm thấy cột 'TEN_DOI' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_ngung_psc.columns)}")
            return

        # Chuẩn hóa cột TEN_KV: giữ phần sau dấu -, bỏ ngoặc đơn
        # VD: PCT1-Nguyễn Mạnh Hùng(PTO) -> Nguyễn Mạnh Hùng
        # VD: Tây Đằng 05 -Lê Quyết Tiến -> Lê Quyết Tiến
        print("Đang chuẩn hóa cột TEN_KV...")

        def normalize_ten_kv(name):
            if pd.isna(name) or str(name).strip() == '':
                return ''
            name = str(name).strip()
            # Lấy phần sau dấu -
            if '-' in name:
                name = name.split('-', 1)[1].strip()
            # Bỏ phần trong ngoặc đơn
            name = re.sub(r'\([^)]*\)', '', name).strip()
            return name

        df_ngung_psc['TEN_KV'] = df_ngung_psc['TEN_KV'].apply(normalize_ten_kv)
        print(f"✅ Đã chuẩn hóa {len(df_ngung_psc)} tên nhân viên")

        # Tạo sheet thống kê theo NVKT: TEN_DOI, TEN_KV, SỐ LƯỢNG NGƯNG PSC THÁNG
        print("\n📊 Đang tạo thống kê ngưng PSC theo NVKT...")
        df_stats = df_ngung_psc.groupby(['TEN_DOI', 'TEN_KV'], dropna=False).size().reset_index(name='SỐ LƯỢNG NGƯNG PSC THÁNG')
        df_stats = df_stats.fillna('(Chưa xác định)')
        df_stats = df_stats.sort_values(['TEN_DOI', 'SỐ LƯỢNG NGƯNG PSC THÁNG'], ascending=[True, False])

        # Thêm dòng TỔNG CỘNG
        total_row = pd.DataFrame([{
            'TEN_DOI': 'TỔNG CỘNG',
            'TEN_KV': '',
            'SỐ LƯỢNG NGƯNG PSC THÁNG': int(df_stats['SỐ LƯỢNG NGƯNG PSC THÁNG'].sum())
        }])
        df_stats = pd.concat([df_stats, total_row], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats) - 1} NVKT")

        # Tạo sheet thống kê theo Đội VT: TEN_DOI, SỐ LƯỢNG NGƯNG PSC THÁNG
        print("📊 Đang tạo thống kê ngưng PSC theo Đội VT...")
        df_stats_to = df_ngung_psc.groupby('TEN_DOI', dropna=False).size().reset_index(name='SỐ LƯỢNG NGƯNG PSC THÁNG')
        df_stats_to = df_stats_to.fillna('(Chưa xác định)')
        df_stats_to = df_stats_to.sort_values('SỐ LƯỢNG NGƯNG PSC THÁNG', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_row_to = pd.DataFrame([{
            'TEN_DOI': 'TỔNG CỘNG',
            'SỐ LƯỢNG NGƯNG PSC THÁNG': int(df_stats_to['SỐ LƯỢNG NGƯNG PSC THÁNG'].sum())
        }])
        df_stats_to = pd.concat([df_stats_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats_to) - 1} Đội VT")

        # Lưu file
        print(f"\n💾 Đang lưu file...")
        with pd.ExcelWriter(ngung_psc_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu gốc
            df_ngung_psc.to_excel(writer, sheet_name='Data', index=False)

            # Sheet 2: Thống kê ngưng PSC tháng theo NVKT
            df_stats.to_excel(writer, sheet_name='mytv_ngung_psc_thang', index=False)

            # Sheet 3: Thống kê ngưng PSC tháng theo Đội VT
            df_stats_to.to_excel(writer, sheet_name='mytv_ngung_psc_thang_theo_to', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Đã lưu file: {ngung_psc_file}")
        print(f"   - Sheet 'Data': Dữ liệu đầy đủ ({len(df_ngung_psc)} dòng)")
        print(f"   - Sheet 'mytv_ngung_psc_thang': Thống kê theo NVKT ({len(df_stats)} dòng)")
        print(f"   - Sheet 'mytv_ngung_psc_thang_theo_to': Thống kê theo Đội VT ({len(df_stats_to)} dòng)")

        # Hiển thị top 5 NVKT
        print("\n📊 Top 5 NVKT có nhiều TB ngưng PSC nhất:")
        top5 = df_stats[df_stats['TEN_DOI'] != 'TỔNG CỘNG'].nlargest(5, 'SỐ LƯỢNG NGƯNG PSC THÁNG')
        for idx, (_, row) in enumerate(top5.iterrows(), 1):
            print(f"   {idx}. {row['TEN_KV']} ({row['TEN_DOI']}): {row['SỐ LƯỢNG NGƯNG PSC THÁNG']} TB")

        print("\n✅ Hoàn thành xử lý báo cáo MyTV Ngưng PSC!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo MyTV Ngưng PSC: {e}")
        import traceback
        traceback.print_exc()

def process_mytv_hoan_cong_report():
    """
    Xử lý báo cáo mytv_hoan_cong:
    1. Đọc file mytv_hoan_cong_DDMMYYYY.xlsx
    2. Chuẩn hóa cột NHANVIEN_KT (giữ phần sau dấu -)
    3. Tạo sheet thống kê mytv_hoan_cong_thang và mytv_hoan_cong_thang_theo_to
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO MYTV HOÀN CÔNG")
        print("="*80)

        # === ĐỌC FILE ===
        download_dir = "PTTB-PSC"
        date_str = datetime.now().strftime("%d%m%Y")
        hoan_cong_file = os.path.join(download_dir, f"mytv_hoan_cong_{date_str}.xlsx")

        if not os.path.exists(hoan_cong_file):
            print(f"❌ Không tìm thấy file: {hoan_cong_file}")
            return

        print(f"📂 Đang đọc file: {hoan_cong_file}")
        df_hoan_cong = pd.read_excel(hoan_cong_file)
        print(f"✅ Đọc thành công {len(df_hoan_cong)} dòng dữ liệu")
        print(f"   Các cột: {', '.join(df_hoan_cong.columns)}")

        # Kiểm tra cột NHANVIEN_KT tồn tại
        if 'NHANVIEN_KT' not in df_hoan_cong.columns:
            print(f"❌ Không tìm thấy cột 'NHANVIEN_KT' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_hoan_cong.columns)}")
            return

        # Kiểm tra cột DOIVT tồn tại
        if 'DOIVT' not in df_hoan_cong.columns:
            print(f"❌ Không tìm thấy cột 'DOIVT' trong file")
            print(f"   Các cột có sẵn: {', '.join(df_hoan_cong.columns)}")
            return

        # Chuẩn hóa cột NHANVIEN_KT: giữ phần sau dấu -
        # VD: CTV072872-Phạm Anh Tuấn -> Phạm Anh Tuấn
        print("Đang chuẩn hóa cột NHANVIEN_KT...")
        df_hoan_cong['NHANVIEN_KT'] = df_hoan_cong['NHANVIEN_KT'].apply(
            lambda x: str(x).split('-', 1)[1].strip() if pd.notna(x) and '-' in str(x) else (str(x).strip() if pd.notna(x) else '')
        )
        print(f"✅ Đã chuẩn hóa {len(df_hoan_cong)} tên nhân viên")

        # Tạo sheet thống kê theo NVKT: DOIVT, NHANVIEN_KT, SỐ LƯỢNG HOÀN CÔNG THÁNG
        print("\n📊 Đang tạo thống kê hoàn công theo NVKT...")
        df_stats = df_hoan_cong.groupby(['DOIVT', 'NHANVIEN_KT'], dropna=False).size().reset_index(name='SỐ LƯỢNG HOÀN CÔNG THÁNG')
        df_stats = df_stats.fillna('(Chưa xác định)')
        df_stats = df_stats.sort_values(['DOIVT', 'SỐ LƯỢNG HOÀN CÔNG THÁNG'], ascending=[True, False])

        # Thêm dòng TỔNG CỘNG
        total_row = pd.DataFrame([{
            'DOIVT': 'TỔNG CỘNG',
            'NHANVIEN_KT': '',
            'SỐ LƯỢNG HOÀN CÔNG THÁNG': int(df_stats['SỐ LƯỢNG HOÀN CÔNG THÁNG'].sum())
        }])
        df_stats = pd.concat([df_stats, total_row], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats) - 1} NVKT")

        # Tạo sheet thống kê theo Đội VT: DOIVT, SỐ LƯỢNG HOÀN CÔNG THÁNG
        print("📊 Đang tạo thống kê hoàn công theo Đội VT...")
        df_stats_to = df_hoan_cong.groupby('DOIVT', dropna=False).size().reset_index(name='SỐ LƯỢNG HOÀN CÔNG THÁNG')
        df_stats_to = df_stats_to.fillna('(Chưa xác định)')
        df_stats_to = df_stats_to.sort_values('SỐ LƯỢNG HOÀN CÔNG THÁNG', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_row_to = pd.DataFrame([{
            'DOIVT': 'TỔNG CỘNG',
            'SỐ LƯỢNG HOÀN CÔNG THÁNG': int(df_stats_to['SỐ LƯỢNG HOÀN CÔNG THÁNG'].sum())
        }])
        df_stats_to = pd.concat([df_stats_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_stats_to) - 1} Đội VT")

        # Lưu file
        print(f"\n💾 Đang lưu file...")
        with pd.ExcelWriter(hoan_cong_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu gốc
            df_hoan_cong.to_excel(writer, sheet_name='Data', index=False)

            # Sheet 2: Thống kê hoàn công tháng theo NVKT
            df_stats.to_excel(writer, sheet_name='mytv_hoan_cong_thang', index=False)

            # Sheet 3: Thống kê hoàn công tháng theo Đội VT
            df_stats_to.to_excel(writer, sheet_name='mytv_hoan_cong_thang_theo_to', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Đã lưu file: {hoan_cong_file}")
        print(f"   - Sheet 'Data': Dữ liệu đầy đủ ({len(df_hoan_cong)} dòng)")
        print(f"   - Sheet 'mytv_hoan_cong_thang': Thống kê theo NVKT ({len(df_stats)} dòng)")
        print(f"   - Sheet 'mytv_hoan_cong_thang_theo_to': Thống kê theo Đội VT ({len(df_stats_to)} dòng)")

        # Hiển thị top 5 NVKT
        print("\n📊 Top 5 NVKT có nhiều TB hoàn công nhất:")
        top5 = df_stats[df_stats['DOIVT'] != 'TỔNG CỘNG'].nlargest(5, 'SỐ LƯỢNG HOÀN CÔNG THÁNG')
        for idx, (_, row) in enumerate(top5.iterrows(), 1):
            print(f"   {idx}. {row['NHANVIEN_KT']} ({row['DOIVT']}): {row['SỐ LƯỢNG HOÀN CÔNG THÁNG']} TB")

        print("\n✅ Hoàn thành xử lý báo cáo MyTV Hoàn công!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo MyTV Hoàn công: {e}")
        import traceback
        traceback.print_exc()

def create_mytv_thuc_tang_report():
    """
    Tạo báo cáo MyTV Thực tăng = Hoàn công - Ngưng phát sinh cước
    Sử dụng dữ liệu từ 2 file mytv đã xử lý
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TẠO BÁO CÁO MYTV THỰC TĂNG")
        print("="*80)

        download_dir = "PTTB-PSC"
        date_str = datetime.now().strftime("%d%m%Y")

        ngung_psc_file = os.path.join(download_dir, f"mytv_ngung_psc_{date_str}.xlsx")
        hoan_cong_file = os.path.join(download_dir, f"mytv_hoan_cong_{date_str}.xlsx")
        thuc_tang_file = os.path.join(download_dir, f"mytv_thuc_tang_{date_str}.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(ngung_psc_file):
            print(f"❌ Không tìm thấy file: {ngung_psc_file}")
            return

        if not os.path.exists(hoan_cong_file):
            print(f"❌ Không tìm thấy file: {hoan_cong_file}")
            return

        # === ĐỌC DỮ LIỆU ===
        print(f"\n📂 Đang đọc dữ liệu...")

        # Đọc sheet thống kê từ file Ngưng PSC
        df_ngung_psc_to = pd.read_excel(ngung_psc_file, sheet_name='mytv_ngung_psc_thang_theo_to')
        df_ngung_psc_nvkt = pd.read_excel(ngung_psc_file, sheet_name='mytv_ngung_psc_thang')

        # Đọc sheet thống kê từ file Hoàn công
        df_hoan_cong_to = pd.read_excel(hoan_cong_file, sheet_name='mytv_hoan_cong_thang_theo_to')
        df_hoan_cong_nvkt = pd.read_excel(hoan_cong_file, sheet_name='mytv_hoan_cong_thang')

        print(f"✅ Đã đọc dữ liệu từ các file")

        # === XỬ LÝ SHEET 1: THỰC TĂNG THEO ĐỘI VT ===
        print(f"\n📊 Đang tạo báo cáo Thực tăng theo Đội VT...")

        # Loại bỏ dòng TỔNG CỘNG
        df_ngung_psc_to_clean = df_ngung_psc_to[df_ngung_psc_to['TEN_DOI'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_to_clean = df_hoan_cong_to[df_hoan_cong_to['DOIVT'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột
        df_ngung_psc_to_clean.rename(columns={'TEN_DOI': 'Đội VT', 'SỐ LƯỢNG NGƯNG PSC THÁNG': 'Ngưng phát sinh cước'}, inplace=True)
        df_hoan_cong_to_clean.rename(columns={'DOIVT': 'Đội VT', 'SỐ LƯỢNG HOÀN CÔNG THÁNG': 'Hoàn công'}, inplace=True)

        # Merge 2 dataframe
        df_thuc_tang_to = pd.merge(
            df_hoan_cong_to_clean,
            df_ngung_psc_to_clean,
            on='Đội VT',
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Hoàn công'] - df_thuc_tang_to['Ngưng phát sinh cước']

        # Tính Tỷ lệ ngưng/psc (%)
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to.apply(
            lambda row: (row['Ngưng phát sinh cước'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int
        df_thuc_tang_to['Hoàn công'] = df_thuc_tang_to['Hoàn công'].astype(int)
        df_thuc_tang_to['Ngưng phát sinh cước'] = df_thuc_tang_to['Ngưng phát sinh cước'].astype(int)
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Thực tăng'].astype(int)
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần
        df_thuc_tang_to = df_thuc_tang_to.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc = int(df_thuc_tang_to['Hoàn công'].sum())
        total_np = int(df_thuc_tang_to['Ngưng phát sinh cước'].sum())
        total_tt = int(df_thuc_tang_to['Thực tăng'].sum())
        total_ty_le = (total_np / total_hc * 100) if total_hc != 0 else 0

        total_row_to = pd.DataFrame([{
            'Đội VT': 'TỔNG CỘNG',
            'Hoàn công': total_hc,
            'Ngưng phát sinh cước': total_np,
            'Thực tăng': total_tt,
            'Tỷ lệ ngưng/psc': round(total_ty_le, 2)
        }])
        df_thuc_tang_to = pd.concat([df_thuc_tang_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_to) - 1} Đội VT")

        # === XỬ LÝ SHEET 2: THỰC TĂNG THEO NVKT ===
        print(f"📊 Đang tạo báo cáo Thực tăng theo NVKT...")

        # Loại bỏ dòng TỔNG CỘNG
        df_ngung_psc_nvkt_clean = df_ngung_psc_nvkt[df_ngung_psc_nvkt['TEN_DOI'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_nvkt_clean = df_hoan_cong_nvkt[df_hoan_cong_nvkt['DOIVT'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột
        df_ngung_psc_nvkt_clean.rename(columns={
            'TEN_DOI': 'Đội VT',
            'TEN_KV': 'NVKT',
            'SỐ LƯỢNG NGƯNG PSC THÁNG': 'Ngưng phát sinh cước'
        }, inplace=True)
        df_hoan_cong_nvkt_clean.rename(columns={
            'DOIVT': 'Đội VT',
            'NHANVIEN_KT': 'NVKT',
            'SỐ LƯỢNG HOÀN CÔNG THÁNG': 'Hoàn công'
        }, inplace=True)

        # Merge theo Đội VT và NVKT
        df_thuc_tang_nvkt = pd.merge(
            df_hoan_cong_nvkt_clean,
            df_ngung_psc_nvkt_clean,
            on=['Đội VT', 'NVKT'],
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Hoàn công'] - df_thuc_tang_nvkt['Ngưng phát sinh cước']

        # Tính Tỷ lệ ngưng/psc (%)
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt.apply(
            lambda row: (row['Ngưng phát sinh cước'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int
        df_thuc_tang_nvkt['Hoàn công'] = df_thuc_tang_nvkt['Hoàn công'].astype(int)
        df_thuc_tang_nvkt['Ngưng phát sinh cước'] = df_thuc_tang_nvkt['Ngưng phát sinh cước'].astype(int)
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Thực tăng'].astype(int)
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần
        df_thuc_tang_nvkt = df_thuc_tang_nvkt.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc_nvkt = int(df_thuc_tang_nvkt['Hoàn công'].sum())
        total_np_nvkt = int(df_thuc_tang_nvkt['Ngưng phát sinh cước'].sum())
        total_tt_nvkt = int(df_thuc_tang_nvkt['Thực tăng'].sum())
        total_ty_le_nvkt = (total_np_nvkt / total_hc_nvkt * 100) if total_hc_nvkt != 0 else 0

        total_row_nvkt = pd.DataFrame([{
            'Đội VT': 'TỔNG CỘNG',
            'NVKT': '',
            'Hoàn công': total_hc_nvkt,
            'Ngưng phát sinh cước': total_np_nvkt,
            'Thực tăng': total_tt_nvkt,
            'Tỷ lệ ngưng/psc': round(total_ty_le_nvkt, 2)
        }])
        df_thuc_tang_nvkt = pd.concat([df_thuc_tang_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_nvkt) - 1} NVKT")

        # === LƯU FILE ===
        print(f"\n💾 Đang lưu file báo cáo Thực tăng...")
        with pd.ExcelWriter(thuc_tang_file, engine='openpyxl') as writer:
            # Sheet 1: Thống kê theo Đội VT
            df_thuc_tang_to.to_excel(writer, sheet_name='thuc_tang_theo_to', index=False)

            # Sheet 2: Thống kê theo NVKT
            df_thuc_tang_nvkt.to_excel(writer, sheet_name='thuc_tang_theo_NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Đã lưu file: {thuc_tang_file}")
        print(f"   - Sheet 'thuc_tang_theo_to': Thống kê theo Đội VT ({len(df_thuc_tang_to)} dòng)")
        print(f"   - Sheet 'thuc_tang_theo_NVKT': Thống kê theo NVKT ({len(df_thuc_tang_nvkt)} dòng)")

        # Hiển thị thống kê tổng quan
        total_hoan_cong = int(df_thuc_tang_to[df_thuc_tang_to['Đội VT'] == 'TỔNG CỘNG']['Hoàn công'].iloc[0])
        total_ngung_psc = int(df_thuc_tang_to[df_thuc_tang_to['Đội VT'] == 'TỔNG CỘNG']['Ngưng phát sinh cước'].iloc[0])
        total_thuc_tang = int(df_thuc_tang_to[df_thuc_tang_to['Đội VT'] == 'TỔNG CỘNG']['Thực tăng'].iloc[0])

        print(f"\n📊 Tổng quan:")
        print(f"   - Tổng Hoàn công: {total_hoan_cong} TB")
        print(f"   - Tổng Ngưng phát sinh cước: {total_ngung_psc} TB")
        print(f"   - Thực tăng: {total_thuc_tang} TB")

        # Top 5 Đội VT có thực tăng cao nhất
        print(f"\n📊 Top 5 Đội VT có Thực tăng cao nhất:")
        top5_to = df_thuc_tang_to[df_thuc_tang_to['Đội VT'] != 'TỔNG CỘNG'].head(5)
        for idx, (_, row) in enumerate(top5_to.iterrows(), 1):
            doi_vt = row['Đội VT'] if pd.notna(row['Đội VT']) else '(Chưa xác định)'
            print(f"   {idx}. {doi_vt}: {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng phát sinh cước']})")

        # Top 5 NVKT có thực tăng cao nhất
        print(f"\n📊 Top 5 NVKT có Thực tăng cao nhất:")
        top5_nvkt = df_thuc_tang_nvkt[df_thuc_tang_nvkt['NVKT'] != ''].nlargest(5, 'Thực tăng')
        for idx, (_, row) in enumerate(top5_nvkt.iterrows(), 1):
            nvkt = row['NVKT'] if pd.notna(row['NVKT']) else '(Chưa xác định)'
            doi_vt = row['Đội VT'] if pd.notna(row['Đội VT']) else '(Chưa xác định)'
            print(f"   {idx}. {nvkt} ({doi_vt}): {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng phát sinh cước']})")

        # Import vào database history
        if HISTORY_IMPORT_AVAILABLE:
            try:
                print(f"\n💾 Đang lưu vào database history...")
                importer = ReportsHistoryImporter()
                importer.import_growth_mytv()
                print(f"✅ Đã lưu vào database history")
            except Exception as e:
                print(f"⚠️  Không thể lưu vào database history: {e}")

        print("\n✅ Hoàn thành tạo báo cáo MyTV Thực tăng!")

    except Exception as e:
        print(f"❌ Lỗi khi tạo báo cáo MyTV Thực tăng: {e}")
        import traceback
        traceback.print_exc()


def process_son_tay_ngung_psc_report():
    """
    Xử lý báo cáo Sơn Tây Ngưng PSC (Tháng T-1):
    1. Tìm file ngung_psc_thang_t-1_sontay_*.xlsx mới nhất
    2. Trích xuất dữ liệu từ các row/column cụ thể
    3. Thêm dòng Tổng
    4. Lưu vào sheet TH_ngung_PSC-Thang T-1
    """
    import glob
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO SƠN TÂY NGƯNG PSC (THÁNG T-1)")
        print("="*80)

        download_dir = "PTTB-PSC"
        pattern = os.path.join(download_dir, "ngung_psc_fiber_thang_t-1_sontay_*.xlsx")
        files = glob.glob(pattern)

        if not files:
            print(f"❌ Không tìm thấy file dạng: {pattern}")
            return

        latest_file = max(files, key=os.path.getmtime)
        print(f"📂 Đang xử lý file mới nhất: {latest_file}")

        # Đọc file với header=None vì header nhiều tầng
        df_raw = pd.read_excel(latest_file, header=None)
        df_subset = _extract_son_tay_ngung_psc_subset(df_raw)

        # Ép kiểu dữ liệu số
        numeric_cols = df_subset.columns[1:]
        for col in numeric_cols:
            df_subset[col] = pd.to_numeric(df_subset[col], errors='coerce').fillna(0).astype(int)

        # Trình bày dữ liệu
        print("\n📊 Dữ liệu trích xuất:")
        print(df_subset.to_string(index=False))

        # Thêm dòng Tổng
        totals = df_subset[numeric_cols].sum()
        total_row = pd.DataFrame([{
            "Đơn vị/Nhân viên KT": "Tổng",
            "Hoàn công(*) (1.5)": totals["Hoàn công(*) (1.5)"],
            "Lũy kế tháng(1.6)": totals["Lũy kế tháng(1.6)"],
            "Lũy kế năm(1.7)": totals["Lũy kế năm(1.7)"],
            "Ngưng PSC tạm tính tháng T(5.1)": totals["Ngưng PSC tạm tính tháng T(5.1)"],
            "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)": totals["TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)"],
        }])
        df_final = pd.concat([df_subset, total_row], ignore_index=True)

        # Lưu vào file gốc, sheet mới
        print(f"\n💾 Đang lưu vào sheet 'TH_ngung_PSC-Thang T-1'...")
        with pd.ExcelWriter(latest_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_final.to_excel(writer, sheet_name='TH_ngung_PSC-Thang T-1', index=False)

            # Định dạng cột
            worksheet = writer.sheets['TH_ngung_PSC-Thang T-1']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[column_letter].width = max_length + 5

        print(f"✅ Đã lưu sheet mới vào: {latest_file}")
        print("✅ Hoàn thành xử lý báo cáo Sơn Tây!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo Sơn Tây: {e}")
        import traceback
        traceback.print_exc()


def process_son_tay_mytv_ngung_psc_report():
    """
    Xử lý báo cáo MyTV Sơn Tây Ngưng PSC (Tháng T-1):
    1. Tìm file ngung_psc_mytv_thang_t-1_sontay_*.xlsx mới nhất
    2. Trích xuất dữ liệu từ các row/column cụ thể
    3. Thêm dòng Tổng
    4. Lưu vào sheet TH_ngung_PSC-Thang T-1
    """
    import glob
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO MYTV SƠN TÂY NGƯNG PSC (THÁNG T-1)")
        print("="*80)

        download_dir = "PTTB-PSC"
        pattern = os.path.join(download_dir, "ngung_psc_mytv_thang_t-1_sontay_*.xlsx")
        files = glob.glob(pattern)

        if not files:
            print(f"❌ Không tìm thấy file dạng: {pattern}")
            return

        latest_file = max(files, key=os.path.getmtime)
        print(f"📂 Đang xử lý file mới nhất: {latest_file}")

        # Đọc file với header=None vì header nhiều tầng
        df_raw = pd.read_excel(latest_file, header=None)
        df_subset = _extract_son_tay_ngung_psc_subset(df_raw)

        # Ép kiểu dữ liệu số
        numeric_cols = df_subset.columns[1:]
        for col in numeric_cols:
            df_subset[col] = pd.to_numeric(df_subset[col], errors='coerce').fillna(0).astype(int)

        # Trình bày dữ liệu
        print("\n📊 Dữ liệu trích xuất:")
        print(df_subset.to_string(index=False))

        # Thêm dòng Tổng
        totals = df_subset[numeric_cols].sum()
        total_row = pd.DataFrame([{
            "Đơn vị/Nhân viên KT": "Tổng",
            "Hoàn công(*) (1.5)": totals["Hoàn công(*) (1.5)"],
            "Lũy kế tháng(1.6)": totals["Lũy kế tháng(1.6)"],
            "Lũy kế năm(1.7)": totals["Lũy kế năm(1.7)"],
            "Ngưng PSC tạm tính tháng T(5.1)": totals["Ngưng PSC tạm tính tháng T(5.1)"],
            "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)": totals["TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)"],
        }])
        df_final = pd.concat([df_subset, total_row], ignore_index=True)

        # Lưu vào file gốc, sheet mới
        print(f"\n💾 Đang lưu vào sheet 'TH_ngung_PSC-Thang T-1'...")
        with pd.ExcelWriter(latest_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_final.to_excel(writer, sheet_name='TH_ngung_PSC-Thang T-1', index=False)

            # Định dạng cột
            worksheet = writer.sheets['TH_ngung_PSC-Thang T-1']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[column_letter].width = max_length + 5

        print(f"✅ Đã lưu sheet mới vào: {latest_file}")
        print("✅ Hoàn thành xử lý báo cáo MyTV Sơn Tây!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo MyTV Sơn Tây: {e}")
        import traceback
        traceback.print_exc()


def main():
    """
    Hàm main để chạy standalone tất cả các hàm xử lý báo cáo thực tăng
    """
    print("=" * 80)
    print("BẮT ĐẦU XỬ LÝ BÁO CÁO THỰC TĂNG")
    print("=" * 80)

    # === PTTB ===
    print("\n📌 [1/7] Xử lý báo cáo PTTB Ngưng PSC...")
    process_ngung_psc_report()

    print("\n📌 [2/7] Xử lý báo cáo PTTB Hoàn công...")
    process_hoan_cong_report()

    print("\n📌 [3/7] Tạo báo cáo PTTB Thực tăng...")
    create_thuc_tang_report()

    # === MyTV ===
    print("\n📌 [4/7] Xử lý báo cáo MyTV Ngưng PSC...")
    process_mytv_ngung_psc_report()

    print("\n📌 [5/7] Xử lý báo cáo MyTV Hoàn công...")
    process_mytv_hoan_cong_report()

    print("\n📌 [6/7] Tạo báo cáo MyTV Thực tăng...")
    create_mytv_thuc_tang_report()

    # === Sơn Tây ===
    print("\n📌 [7/8] Xử lý báo cáo Fiber Sơn Tây Ngưng PSC (Tháng T-1)...")
    process_son_tay_ngung_psc_report()

    print("\n📌 [8/8] Xử lý báo cáo MyTV Sơn Tây Ngưng PSC (Tháng T-1)...")
    process_son_tay_mytv_ngung_psc_report()

    print("\n" + "=" * 80)
    print("✅ HOÀN THÀNH XỬ LÝ TẤT CẢ BÁO CÁO THỰC TĂNG!")
    print("=" * 80)


if __name__ == "__main__":
    main()
