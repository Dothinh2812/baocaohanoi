# -*- coding: utf-8 -*-
import pandas as pd
import os
import re
from datetime import datetime

# Import module để lưu vào database history
try:
    from import_reports_history import ReportsHistoryImporter
    HISTORY_IMPORT_AVAILABLE = True
except ImportError:
    HISTORY_IMPORT_AVAILABLE = False


def process_GHTT_report_NVKT():
    """
    Xử lý file GHTT/tong_hop_ghtt_nvktdb.xlsx:
    1. Đọc file với 2 hàng header (merged cells)
    2. Chuẩn hóa cột 'Đơn vị' - giữ lại phần sau dấu '-'
    3. Đổi tên các cột
    4. Bổ sung cột 'Đơn vị' từ file dsnv.xlsx bằng cách tìm 'Họ tên' = 'NVKT'
    5. Lưu vào sheet kq_nvktdb
    """
    try:
        # Đường dẫn file
        input_file = os.path.join("GHTT", "tong_hop_ghtt_nvktdb.xlsx")
        output_file = os.path.join("GHTT", "tong_hop_ghtt_nvktdb.xlsx")
        dsnv_file = "dsnv.xlsx"

        # Kiểm tra file tồn tại
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"=== Bắt đầu xử lý GHTT NVKT ĐB ===")
        print(f"Đang đọc file: {input_file}")

        # Đọc file Excel với header=None để xử lý 2 hàng header merged
        df_raw = pd.read_excel(input_file, header=None)
        print(f"Dữ liệu raw - Hình dạng: {df_raw.shape}")

        # Lấy 2 hàng header
        header_row1 = df_raw.iloc[0].ffill()  # Header nhóm (merged) - forward fill cho merged cells
        header_row2 = df_raw.iloc[1]  # Header con


        # Tạo tên cột kết hợp từ 2 hàng header
        new_columns = []
        for i in range(len(header_row1)):
            h1 = str(header_row1[i]).strip() if pd.notna(header_row1[i]) else ""
            h2 = str(header_row2[i]).strip() if pd.notna(header_row2[i]) else ""

            if h1 == h2 or h2 == "":
                # Cột không có sub-header (vd: Đơn vị, Tồn quá hạn...)
                new_columns.append(h1)
            elif h1 == "":
                new_columns.append(h2)
            else:
                # Cột có sub-header -> kết hợp: "nhóm - con"
                new_columns.append(f"{h1} - {h2}")

        # Lấy dữ liệu từ hàng thứ 3 trở đi (bỏ 2 hàng header)
        df = df_raw.iloc[2:].reset_index(drop=True)
        df.columns = new_columns

        print(f"Dữ liệu sau xử lý header - Hình dạng: {df.shape}")
        print(f"Cột:")
        for i, c in enumerate(df.columns):
            print(f"  {i}: {c}")

        # 0. Đọc file dsnv.xlsx để lấy thông tin đơn vị
        print(f"\nĐang đọc file dsnv: {dsnv_file}")
        if os.path.exists(dsnv_file):
            df_dsnv = pd.read_excel(dsnv_file)
            print(f"✅ Đã đọc file dsnv, hình dạng: {df_dsnv.shape}")
            df_dsnv.columns = df_dsnv.columns.str.strip()
        else:
            print(f"⚠️ Không tìm thấy file {dsnv_file}, sẽ bỏ qua bước bổ sung đơn vị")
            df_dsnv = None

        # 1. Tìm cột 'Đơn vị'
        nvkt_col_name = None
        for col in df.columns:
            if col.lower() == 'đơn vị':
                nvkt_col_name = col
                break

        if nvkt_col_name is None:
            print(f"⚠️ Không tìm thấy cột 'Đơn vị' trong file gốc")
            print(f"Các cột có sẵn: {df.columns.tolist()}")
            return False

        # 2. Chuẩn hóa cột 'Đơn vị' - giữ lại phần sau dấu '-'
        print(f"\n✓ Đang chuẩn hóa cột '{nvkt_col_name}'...")
        df[nvkt_col_name] = df[nvkt_col_name].apply(
            lambda x: x.split('-', 1)[1].strip() if isinstance(x, str) and '-' in x else x
        )
        print(f"✅ Đã chuẩn hóa cột '{nvkt_col_name}'")

        # 3. Đổi tên cột 'Đơn vị' -> 'NVKT'
        df = df.rename(columns={nvkt_col_name: 'NVKT'})
        print(f"✅ Đã đổi tên '{nvkt_col_name}' -> 'NVKT'")

        # 4. Đổi tên các cột khác (rút gọn tên cột dài)
        column_rename = {}
        for col in df.columns:
            # Lấy phần sub-header (sau dấu ' - ') nếu có
            sub = col.split(' - ')[-1].strip() if ' - ' in col else col

            # Tháng T columns
            if 'tháng T đạt 80%' in col:
                if 'Kế hoạch giao' in sub:
                    column_rename[col] = 'KH giao T'
                elif 'gia hạn thành công' in sub:
                    column_rename[col] = 'Hoàn thành T'
                elif 'giao về NVKT' in sub:
                    column_rename[col] = 'Giao NVKT T'
                elif sub == 'Tỷ lệ':
                    column_rename[col] = 'Tỷ lệ T'
                elif sub == 'Điểm':
                    column_rename[col] = 'Điểm T'
            # Tháng T+1 columns
            elif 'tháng T+1 đạt 90%' in col:
                if 'Kế hoạch giao' in sub:
                    column_rename[col] = 'KH giao T+1'
                elif 'gia hạn thành công' in sub or 'SL thuê bao' in sub:
                    column_rename[col] = 'Hoàn thành T+1'
                elif 'Số giao NVKT' in sub or 'giao NVKT' in sub:
                    column_rename[col] = 'Giao NVKT T+1'
                elif sub == 'Tỷ lệ':
                    column_rename[col] = 'Tỷ lệ T+1'
                elif sub == 'Điểm':
                    column_rename[col] = 'Điểm T+1'

        if column_rename:
            df = df.rename(columns=column_rename)
            print(f"\n✅ Đã đổi tên các cột:")
            for old_name, new_name in column_rename.items():
                print(f"  {old_name} -> {new_name}")

        # Loại bỏ cột không có dữ liệu
        drop_cols = [c for c in ['KH giao T', 'KH giao T+1'] if c in df.columns]
        if drop_cols:
            df = df.drop(columns=drop_cols)
            print(f"✅ Đã loại bỏ cột: {drop_cols}")


        # 5. Bổ sung cột 'Đơn vị' từ dsnv.xlsx (tìm NVKT trong 'Họ tên')
        if df_dsnv is not None and 'Họ tên' in df_dsnv.columns:
            dsnv_unit_col = None
            for col in df_dsnv.columns:
                if col.lower() == 'đơn vị':
                    dsnv_unit_col = col
                    break

            if dsnv_unit_col is None:
                print(f"\n⚠️ Không tìm thấy cột 'đơn vị' trong dsnv.xlsx")
            else:
                print(f"\n✓ Đang bổ sung cột 'Đơn vị' từ dsnv.xlsx...")

                def find_unit_fuzzy(nvkt_name):
                    """Tìm đơn vị trong dsnv (case-insensitive)"""
                    if pd.isna(nvkt_name):
                        return None
                    nvkt_name_lower = str(nvkt_name).strip().lower()
                    for idx, row in df_dsnv.iterrows():
                        ho_ten = str(row['Họ tên']).strip().lower() if pd.notna(row['Họ tên']) else ""
                        if ho_ten == nvkt_name_lower:
                            return row[dsnv_unit_col]
                    return None

                df['Đơn vị'] = df['NVKT'].apply(find_unit_fuzzy)
                # Di chuyển cột 'Đơn vị' ra sau cột 'NVKT'
                cols = df.columns.tolist()
                cols.remove('Đơn vị')
                nvkt_idx = cols.index('NVKT')
                cols.insert(nvkt_idx + 1, 'Đơn vị')
                df = df[cols]

                print("✅ Đã bổ sung cột 'Đơn vị'")
                print(f"   Số bản ghi tìm được đơn vị: {df['Đơn vị'].notna().sum()}/{len(df)}")
        else:
            print("\n⚠️ Không thể bổ sung cột 'Đơn vị' - thiếu dữ liệu từ dsnv.xlsx")

        # 6. Reset index
        df = df.reset_index(drop=True)

        # 7. Lưu vào sheet kq_nvktdb
        print(f"\n✓ Đang lưu vào sheet 'kq_nvktdb' trong file: {output_file}")

        # Đọc các sheet hiện có (nếu có) để giữ lại
        from openpyxl import load_workbook
        if os.path.exists(output_file):
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='kq_nvktdb', index=False)
        else:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='kq_nvktdb', index=False)

        print(f"✅ Đã lưu vào sheet 'kq_nvktdb'")

        print(f"\n=== Tóm tắt ===")
        print(f"Dữ liệu xử lý - Hình dạng: {df.shape}")
        print(f"Cột sau xử lý: {df.columns.tolist()}")
        if 'Đơn vị' in df.columns:
            print(f"Các đơn vị: {sorted([str(u) for u in df['Đơn vị'].dropna().unique()])}")

        return True

    except Exception as e:
        print(f"❌ Lỗi khi xử lý file GHTT NVKT ĐB: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_KR6_report_tong_hop():
    """
    Xử lý file download_KR6_report_tong_hop.xlsx:
    1. Đổi tên cột "Số lượng KH hoàn thành gia hạn TTTC thành công tháng T" -> "Hoàn thành tháng T"
    2. Đổi tên cột "Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng T" -> "Giao tháng T"
    3. Lưu lại thành file download_KR6_report_tong_hop_processed.xlsx
    """
    try:
        print("\n=== Bắt đầu xử lý KR6 tổng hợp ===")
        # Đường dẫn file
        download_dir = os.path.join("downloads", "baocao_hanoi")
        input_file = os.path.join(download_dir, "download_KR6_report_tong_hop.xlsx")
        output_file = os.path.join(download_dir, "download_KR6_report_tong_hop_processed.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"Đang đọc file: {input_file}")
        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"Dữ liệu gốc - Hình dạng: {df.shape}")

        # Đổi tên cột
        column_rename = {
            "Số lượng KH hoàn thành gia hạn TTTC thành công tháng T": "Hoàn thành tháng T",
            "Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng T": "Giao tháng T"
        }

        df = df.rename(columns=column_rename)
        print("\n✅ Đã đổi tên các cột:")
        for old, new in column_rename.items():
            print(f"  '{old}' -> '{new}'")

        # Thêm hàng tổng
        print("\n✓ Đang thêm hàng tổng...")

        # Tính tổng cho các cột số
        tong_hoan_thanh = 0
        tong_giao = 0

        if "Hoàn thành tháng T" in df.columns:
            tong_hoan_thanh = pd.to_numeric(df["Hoàn thành tháng T"], errors='coerce').sum()

        if "Giao tháng T" in df.columns:
            tong_giao = pd.to_numeric(df["Giao tháng T"], errors='coerce').sum()

        # Tính tỷ lệ
        ty_le_tong = (tong_hoan_thanh / tong_giao * 100) if tong_giao > 0 else 0

        # Format cột Tỷ lệ cho tất cả các hàng (làm tròn 2 chữ số thập phân và thêm %)
        for col in df.columns:
            if col == "Tỷ lệ" or "Tỷ lệ" in col or "tỷ lệ" in col.lower():
                # Hàm format giá trị Tỷ lệ (xử lý cả text có % và số)
                def format_ty_le(val):
                    if pd.isna(val):
                        return ""
                    # Nếu là chuỗi, loại bỏ % và convert sang số
                    if isinstance(val, str):
                        val = val.strip().replace('%', '').replace(',', '.')
                        try:
                            val = float(val)
                        except:
                            return ""
                    # Format với 2 chữ số thập phân
                    try:
                        return f"{float(val):.2f}%"
                    except:
                        return ""

                df[col] = df[col].apply(format_ty_le)

        # Tạo hàng tổng
        tong_row = {}
        for col in df.columns:
            if col == "Hoàn thành tháng T":
                tong_row[col] = tong_hoan_thanh
            elif col == "Giao tháng T":
                tong_row[col] = tong_giao
            elif col == "Tỷ lệ" or "Tỷ lệ" in col or "tỷ lệ" in col.lower():
                tong_row[col] = f"{ty_le_tong:.2f}%"
            elif col == df.columns[0]:  # Cột đầu tiên (thường là tên đơn vị)
                tong_row[col] = "Tổng"
            else:
                tong_row[col] = ""

        # Thêm hàng tổng vào DataFrame
        df_tong = pd.DataFrame([tong_row])
        df = pd.concat([df, df_tong], ignore_index=True)

        print(f"✅ Đã thêm hàng tổng:")
        print(f"   - Hoàn thành tháng T: {tong_hoan_thanh}")
        print(f"   - Giao tháng T: {tong_giao}")
        print(f"   - Tỷ lệ: {ty_le_tong:.2f}%")

        # Lưu file đã xử lý
        print(f"\n✓ Đang lưu file: {output_file}")
        df.to_excel(output_file, index=False)
        print(f"✅ Đã lưu file xử lý: {output_file}")

        print(f"\n=== Tóm tắt KR6 tổng hợp ===")
        print(f"Dữ liệu xử lý - Hình dạng: {df.shape}")
        print(f"Cột sau xử lý: {df.columns.tolist()}")

        return True

    except Exception as e:
        print(f"❌ Lỗi khi xử lý file KR6 tổng hợp: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_KR7_report_NVKT():
    """
    Xử lý file download_KR7_report_NVKT.xlsx:
    1. Chuẩn hóa cột 'đơn vị' - giữ lại phần sau dấu '-'
    2. Đổi tên các cột
    3. Bổ sung cột 'đơn vị' từ file dsnv.xlsx bằng cách tìm 'Họ tên' = 'NVKT'
    4. Sắp xếp theo 'Tỉ lệ hoàn thành' từ thấp -> cao
    5. Lưu file đã xử lý
    6. Tạo các tab riêng cho mỗi đơn vị
    """
    try:
        # Đường dẫn file
        download_dir = os.path.join("downloads", "baocao_hanoi")
        input_file = os.path.join(download_dir, "download_KR7_report_NVKT.xlsx")
        output_file = os.path.join(download_dir, "download_KR7_report_NVKT_processed.xlsx")
        dsnv_file = "dsnv.xlsx"

        # Kiểm tra file tồn tại
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"Đang đọc file: {input_file}")
        # Đọc file Excel
        df = pd.read_excel(input_file)

        print(f"Dữ liệu gốc - Hình dạng: {df.shape}")
        print(f"Cột: {df.columns.tolist()}")

        # 0. Đọc file dsnv.xlsx để lấy thông tin đơn vị
        print(f"\nĐang đọc file dsnv: {dsnv_file}")
        if os.path.exists(dsnv_file):
            df_dsnv = pd.read_excel(dsnv_file)
            print(f"✅ Đã đọc file dsnv, hình dạng: {df_dsnv.shape}")
            # Chuẩn hóa tên cột
            df_dsnv.columns = df_dsnv.columns.str.strip()
            print(f"Cột dsnv: {df_dsnv.columns.tolist()}")
        else:
            print(f"⚠️ Không tìm thấy file {dsnv_file}, sẽ bỏ qua bước bổ sung đơn vị")
            df_dsnv = None

        # 1. Tìm cột NVKT (có thể là 'đơn vị', 'Đơn vị', hay tên khác)
        nvkt_col_name = None
        for col in df.columns:
            if col.lower() == 'đơn vị':
                nvkt_col_name = col
                break

        if nvkt_col_name is None:
            print(f"⚠️ Không tìm thấy cột 'đơn vị' trong file gốc")
            print(f"Các cột có sẵn: {df.columns.tolist()}")
            return False

        # 2. Chuẩn hóa cột 'đơn vị' - giữ lại phần sau dấu '-'
        print(f"\n✓ Đang chuẩn hóa cột '{nvkt_col_name}'...")
        df[nvkt_col_name] = df[nvkt_col_name].apply(
            lambda x: x.split('-')[1].strip() if isinstance(x, str) and '-' in x else x
        )
        print(f"✅ Đã chuẩn hóa cột '{nvkt_col_name}'")

        # 3. Đổi tên cột từ 'đơn vị' (hoặc biến thể) -> 'NVKT'
        df = df.rename(columns={nvkt_col_name: 'NVKT'})
        print(f"\n✅ Đã đổi tên '{nvkt_col_name}' -> 'NVKT'")

        # 4. Bổ sung cột 'Đơn vị' từ dsnv.xlsx (tìm NVKT trong Họ tên)
        if df_dsnv is not None and 'Họ tên' in df_dsnv.columns:
            # Chuẩn hóa tên cột dsnv (có thể là 'Đơn vị' hoặc 'đơn vị')
            dsnv_unit_col = None
            for col in df_dsnv.columns:
                if col.lower() == 'đơn vị':
                    dsnv_unit_col = col
                    break

            if dsnv_unit_col is None:
                print(f"\n⚠️ Không tìm thấy cột 'đơn vị' trong dsnv.xlsx")
                print(f"Các cột dsnv: {df_dsnv.columns.tolist()}")
            else:
                print(f"\n✓ Đang bổ sung cột 'Đơn vị' từ dsnv.xlsx (tìm từ cột '{dsnv_unit_col}')...")

                # Tạo hàm tìm đơn vị với fuzzy matching (không phân biệt hoa/thường)
                def find_unit_fuzzy(nvkt_name):
                    """Tìm đơn vị trong dsnv, hỗ trợ các biến thể tên (vd: Bùi Văn Cường vs Bùi văn Cường)"""
                    if pd.isna(nvkt_name):
                        return None

                    nvkt_name_lower = str(nvkt_name).strip().lower()

                    # Tìm kiếm chính xác (case-insensitive)
                    for idx, row in df_dsnv.iterrows():
                        ho_ten = str(row['Họ tên']).strip().lower() if pd.notna(row['Họ tên']) else ""
                        if ho_ten == nvkt_name_lower:
                            return row[dsnv_unit_col]

                    return None

                # Áp dụng hàm tìm kiếm
                df['Đơn vị'] = df['NVKT'].apply(find_unit_fuzzy)
                print("✅ Đã bổ sung cột 'Đơn vị'")
                print(f"   Số bản ghi tìm được đơn vị: {df['Đơn vị'].notna().sum()}/{len(df)}")
        else:
            print("\n⚠️ Không thể bổ sung cột 'Đơn vị' - thiếu dữ liệu từ dsnv.xlsx")

        # 3. Đổi tên các cột khác
        column_rename = {}

        # Đổi 'Tỷ lệ' -> 'Tỉ lệ hoàn thành'
        if 'Tỷ lệ' in df.columns:
            column_rename['Tỷ lệ'] = 'Tỉ lệ hoàn thành'

        # Xử lý các cột tháng
        for col in df.columns:
            # Tìm cột "Lũy kế SL KH KTĐC tháng T hoàn thành gia hạn TTTC đến tháng T+1" -> "Hoàn thành tháng T+1"
            if 'Lũy kế SL KH KTĐC tháng' in col and 'hoàn thành gia hạn TTTC đến tháng' in col:
                match = re.search(r'đến tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Hoàn thành tháng {thang}'

            # Tìm cột "Tổng SL KH KTĐC tháng T giao về TTVT" -> "Giao tháng T"
            elif 'Tổng SL KH KTĐC tháng' in col and 'giao về TTVT' in col:
                match = re.search(r'tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Giao tháng {thang}'

            # Tìm cột "Số lượng KH hoàn thành gia hạn TTTC thành công tháng T" -> "Hoàn thành tháng T+1"
            elif 'Số lượng KH hoàn thành gia hạn TTTC thành công tháng' in col:
                match = re.search(r'tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Hoàn thành tháng {thang}+1'

            # Tìm cột "Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng T" -> "Giao tháng T+1"
            elif 'Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng' in col:
                match = re.search(r'tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Giao tháng {thang}+1'

        # Đổi tên cột
        if column_rename:
            df = df.rename(columns=column_rename)
            print(f"\n✅ Đã đổi tên các cột:")
            for old_name, new_name in column_rename.items():
                print(f"  {old_name} -> {new_name}")

        # 3. Sắp xếp theo 'Tỉ lệ hoàn thành' từ thấp -> cao
        if 'Tỉ lệ hoàn thành' in df.columns:
            print("\n✓ Đang sắp xếp theo 'Tỉ lệ hoàn thành'...")
            # Chuyển đổi thành số (xóa % nếu có)
            df['Tỉ lệ hoàn thành'] = pd.to_numeric(
                df['Tỉ lệ hoàn thành'].astype(str).str.replace('%', ''),
                errors='coerce'
            )
            # Sắp xếp từ thấp -> cao
            df = df.sort_values('Tỉ lệ hoàn thành', ascending=True)
            print("✅ Đã sắp xếp theo 'Tỉ lệ hoàn thành' từ thấp -> cao")
        else:
            print("⚠️ Không tìm thấy cột 'Tỉ lệ hoàn thành'")

        # 4. Reset index
        df = df.reset_index(drop=True)

        # 5. Lưu file đã xử lý (với các tab theo đơn vị)
        print(f"\n✓ Đang lưu file: {output_file}")

        # Lưu file với ExcelWriter để tạo nhiều sheet
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet 1: Tất cả dữ liệu
            df.to_excel(writer, sheet_name='Tất cả', index=False)
            print(f"   ✅ Đã tạo sheet 'Tất cả'")

            # Tạo các sheet riêng cho từng đơn vị
            if 'Đơn vị' in df.columns:
                unique_units = df['Đơn vị'].dropna().unique()
                print(f"   Tìm thấy {len(unique_units)} đơn vị:")

                for unit in sorted(unique_units):
                    if pd.notna(unit):
                        # Lọc dữ liệu theo đơn vị
                        df_unit = df[df['Đơn vị'] == unit].reset_index(drop=True)

                        # Tạo tên sheet (tối đa 31 ký tự cho Excel)
                        sheet_name = str(unit)[:31]

                        # Lưu vào sheet
                        df_unit.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"   ✅ Đã tạo sheet '{sheet_name}' ({len(df_unit)} bản ghi)")

        print(f"✅ Đã lưu file xử lý: {output_file}")

        print(f"\n=== Tóm tắt ===")
        print(f"Dữ liệu xử lý - Hình dạng: {df.shape}")
        print(f"Cột sau xử lý: {df.columns.tolist()}")
        if 'Đơn vị' in df.columns:
            print(f"Các đơn vị: {sorted([str(u) for u in df['Đơn vị'].dropna().unique()])}")

        return True

    except Exception as e:
        print(f"❌ Lỗi khi xử lý file KR7: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_KR7_report_tong_hop():
    """
    Xử lý file download_KR7_report_tong_hop.xlsx:
    1. Đổi tên cột "Lũy kế SL KH KTĐC tháng T hoàn thành gia hạn TTTC đến tháng T+1" -> "Hoàn thành tháng T+1"
    2. Đổi tên cột "Tổng SL KH KTĐC tháng T giao về TTVT" -> "Giao tháng T+1"
    3. Đổi tên cột "Số lượng KH hoàn thành gia hạn TTTC thành công tháng T" -> "Hoàn thành tháng T+1"
    4. Đổi tên cột "Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng T" -> "Giao tháng T+1"
    5. Lưu lại thành file download_KR7_report_tong_hop_processed.xlsx
    """
    try:
        print("\n=== Bắt đầu xử lý KR7 tổng hợp ===")
        # Đường dẫn file
        download_dir = os.path.join("downloads", "baocao_hanoi")
        input_file = os.path.join(download_dir, "download_KR7_report_tong_hop.xlsx")
        output_file = os.path.join(download_dir, "download_KR7_report_tong_hop_processed.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"Đang đọc file: {input_file}")
        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"Dữ liệu gốc - Hình dạng: {df.shape}")

        # Đổi tên cột
        column_rename = {}

        # Xử lý các cột tháng
        for col in df.columns:
            # Tìm cột "Lũy kế SL KH KTĐC tháng T hoàn thành gia hạn TTTC đến tháng T+1" -> "Hoàn thành tháng T+1"
            if 'Lũy kế SL KH KTĐC tháng' in col and 'hoàn thành gia hạn TTTC đến tháng' in col:
                match = re.search(r'đến tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Hoàn thành tháng {thang}'

            # Tìm cột "Tổng SL KH KTĐC tháng T giao về TTVT" -> "Giao tháng T+1"
            elif 'Tổng SL KH KTĐC tháng' in col and 'giao về TTVT' in col:
                match = re.search(r'tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Giao tháng {thang}+1'

            # Tìm cột "Số lượng KH hoàn thành gia hạn TTTC thành công tháng T" -> "Hoàn thành tháng T+1"
            elif 'Số lượng KH hoàn thành gia hạn TTTC thành công tháng' in col:
                match = re.search(r'tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Hoàn thành tháng {thang}+1'

            # Tìm cột "Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng T" -> "Giao tháng T+1"
            elif 'Số lượng KH giao về NVKT qua kênh OB CSKH có thời gian kết thúc đặt cọc tháng' in col:
                match = re.search(r'tháng\s+(\S+)', col)
                if match:
                    thang = match.group(1)
                    column_rename[col] = f'Giao tháng {thang}+1'

        df = df.rename(columns=column_rename)
        if column_rename:
            print("\n✅ Đã đổi tên các cột:")
            for old, new in column_rename.items():
                print(f"  '{old}' -> '{new}'")
        else:
            print("\n⚠️ Không tìm thấy cột tháng để đổi tên")

        # Thêm hàng tổng
        print("\n✓ Đang thêm hàng tổng...")

        # Tìm các cột "Hoàn thành tháng" và "Giao tháng" (có thể có nhiều cột với các tháng khác nhau)
        hoan_thanh_cols = [col for col in df.columns if 'Hoàn thành tháng' in col]
        giao_cols = [col for col in df.columns if 'Giao tháng' in col]

        # Tính tổng cho các cột
        tong_values = {}
        for col in hoan_thanh_cols + giao_cols:
            tong_values[col] = pd.to_numeric(df[col], errors='coerce').sum()

        # Tính tỷ lệ tổng chung từ tất cả các cột Hoàn thành và Giao
        tong_hoan_thanh_all = sum([tong_values.get(col, 0) for col in hoan_thanh_cols])
        tong_giao_all = sum([tong_values.get(col, 0) for col in giao_cols])
        ty_le_tong_chung = (tong_hoan_thanh_all / tong_giao_all * 100) if tong_giao_all > 0 else 0

        # Tính tỷ lệ cho các cặp cột Hoàn thành/Giao (nếu có cùng tháng)
        ty_le_values = {}
        for hoan_thanh_col in hoan_thanh_cols:
            # Tìm cột Giao tương ứng (có cùng tháng)
            thang_match = re.search(r'tháng\s+(\S+)', hoan_thanh_col)
            if thang_match:
                thang = thang_match.group(1)
                giao_col = f'Giao tháng {thang}'
                if giao_col in giao_cols:
                    hoan_thanh_val = tong_values.get(hoan_thanh_col, 0)
                    giao_val = tong_values.get(giao_col, 0)
                    ty_le_values[thang] = (hoan_thanh_val / giao_val * 100) if giao_val > 0 else 0

        # Format cột Tỷ lệ cho tất cả các hàng (làm tròn 2 chữ số thập phân và thêm %)
        for col in df.columns:
            if col == "Tỷ lệ" or "Tỷ lệ" in col or "tỷ lệ" in col.lower():
                # Hàm format giá trị Tỷ lệ (xử lý cả text có % và số)
                def format_ty_le(val):
                    if pd.isna(val):
                        return ""
                    # Nếu là chuỗi, loại bỏ % và convert sang số
                    if isinstance(val, str):
                        val = val.strip().replace('%', '').replace(',', '.')
                        try:
                            val = float(val)
                        except:
                            return ""
                    # Format với 2 chữ số thập phân
                    try:
                        return f"{float(val):.2f}%"
                    except:
                        return ""

                df[col] = df[col].apply(format_ty_le)

        # Tạo hàng tổng
        tong_row = {}
        for col in df.columns:
            if col in tong_values:
                # Cột Hoàn thành hoặc Giao - điền giá trị tổng
                tong_row[col] = tong_values[col]
            elif col == "Tỷ lệ" or "Tỷ lệ" in col or "tỷ lệ" in col.lower():
                # Cột Tỷ lệ - tính từ ty_le_values hoặc ty_le_tong_chung
                # Tìm tháng tương ứng trong tên cột
                thang_match = re.search(r'tháng\s+(\S+)', col)
                if thang_match:
                    thang = thang_match.group(1)
                    if thang in ty_le_values:
                        # Có tỷ lệ cho tháng cụ thể
                        tong_row[col] = f"{ty_le_values[thang]:.2f}%"
                    else:
                        # Không tìm thấy tỷ lệ cho tháng cụ thể, dùng tỷ lệ tổng chung
                        tong_row[col] = f"{ty_le_tong_chung:.2f}%"
                else:
                    # Tỷ lệ không có tháng cụ thể - dùng tỷ lệ tổng chung
                    tong_row[col] = f"{ty_le_tong_chung:.2f}%"
            elif col == df.columns[0]:  # Cột đầu tiên (thường là tên đơn vị)
                tong_row[col] = "Tổng"
            else:
                tong_row[col] = ""

        # Thêm hàng tổng vào DataFrame
        df_tong = pd.DataFrame([tong_row])
        df = pd.concat([df, df_tong], ignore_index=True)

        print(f"✅ Đã thêm hàng tổng:")
        for col, val in tong_values.items():
            print(f"   - {col}: {val}")
        if ty_le_values:
            for thang, ty_le in ty_le_values.items():
                print(f"   - Tỷ lệ tháng {thang}: {ty_le:.2f}%")
        else:
            print(f"   - Tỷ lệ: {ty_le_tong_chung:.2f}%")

        # Lưu file đã xử lý
        print(f"\n✓ Đang lưu file: {output_file}")
        df.to_excel(output_file, index=False)
        print(f"✅ Đã lưu file xử lý: {output_file}")

        print(f"\n=== Tóm tắt KR7 tổng hợp ===")
        print(f"Dữ liệu xử lý - Hình dạng: {df.shape}")
        print(f"Cột sau xử lý: {df.columns.tolist()}")

        return True

    except Exception as e:
        print(f"❌ Lỗi khi xử lý file KR7 tổng hợp: {e}")
        import traceback
        traceback.print_exc()
        return False


def import_kr_to_history():
    """Import dữ liệu KR vào database history sau khi xử lý xong"""
    if HISTORY_IMPORT_AVAILABLE:
        try:
            print(f"\n💾 Đang lưu KR6 và KR7 vào database history...")
            importer = ReportsHistoryImporter()
            importer.import_kr6()
            importer.import_kr7()
            print(f"✅ Đã lưu KR reports vào database history")
        except Exception as e:
            print(f"⚠️  Không thể lưu vào database history: {e}")
    else:
        print("⚠️  Module import_reports_history không khả dụng")


if __name__ == "__main__":
    # Test hàm xử lý KR6 NVKT
    process_GHTT_report_NVKT()
    # process_KR6_report_tong_hop()
    # process_KR7_report_NVKT()
    # process_KR7_report_tong_hop()

    # Import vào database history
    # import_kr_to_history()
