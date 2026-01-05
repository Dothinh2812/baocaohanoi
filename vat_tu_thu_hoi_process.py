# -*- coding: utf-8 -*-
import pandas as pd
import os
import re

# Import module để lưu vào database history
try:
    from import_reports_history import ReportsHistoryImporter
    HISTORY_IMPORT_AVAILABLE = True
except ImportError:
    HISTORY_IMPORT_AVAILABLE = False


def vat_tu_thu_hoi_process():
    """
    Xử lý file bc_thu_hoi_vat_tu.xlsx:
    1. Đọc file bc_thu_hoi_vat_tu.xlsx
    2. Chuẩn hóa cột KHUVUC - giữ lại phần sau dấu '-'
    3. Tạo các cột mới với giá trị cố định:
       - TRANGTHAI_THUHOI = CHUA_THU_HOI
       - LOAI_VT = ONT (hoặc MESH WIFI nếu có)
       - LOAI_PHIEU = PTTB
       - ONT_TBDC_CHAT_LUONG_CAO = 1
       - LOAI_GIAO_GIAMTRU_LAN_HAI = 0
    4. Lưu file xử lý thành bc_thu_hoi_vat_tu_processed.xlsx
    """
    try:
        # Đường dẫn file
        download_dir = os.path.join("downloads", "baocao_hanoi")
        input_file = os.path.join(download_dir, "bc_thu_hoi_vat_tu.xlsx")
        output_file = os.path.join(download_dir, "bc_thu_hoi_vat_tu_processed.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"Đang đọc file: {input_file}")
        # Đọc file Excel
        df = pd.read_excel(input_file)

        print(f"Dữ liệu gốc - Hình dạng: {df.shape}")
        print(f"Cột: {df.columns.tolist()}")

        # 1. Tìm cột NVKT_DIABAN_GIAO (có thể là 'NVKT_DIABAN_GIAO', 'Khu vực', 'KHUVUC', hay tên khác)
        nvkt_col_name = None
        for col in df.columns:
            col_lower = col.lower().strip()
            if col_lower in ['nvkt_diaban_giao', 'nvkt diaban giao', 'khuvuc', 'khu vuc', 'khu_vuc']:
                nvkt_col_name = col
                break

        if nvkt_col_name is None:
            print(f"⚠️ Không tìm thấy cột NVKT_DIABAN_GIAO trong file gốc")
            print(f"Các cột có sẵn: {df.columns.tolist()}")
            # Tiếp tục xử lý ngay cả khi không tìm thấy cột
        else:
            # 2. Chuẩn hóa cột NVKT_DIABAN_GIAO
            print(f"\n✓ Đang chuẩn hóa cột '{nvkt_col_name}'...")

            def normalize_nvkt(x):
                """
                Chuẩn hóa cột NVKT_DIABAN_GIAO:
                1. Giữ lại phần sau dấu '-' (nếu có)
                2. Bỏ các ký tự trong dấu () (nếu có)
                Ví dụ: "Vạn Thắng 11 - Dương Văn Dũng(XND)" -> "Dương Văn Dũng"
                """
                if not isinstance(x, str):
                    return x

                # Bước 1: Giữ lại phần sau dấu '-'
                if '-' in x:
                    x = x.split('-')[1].strip()

                # Bước 2: Bỏ các ký tự trong dấu () - sử dụng regex
                x = re.sub(r'\([^)]*\)', '', x).strip()

                return x

            df[nvkt_col_name] = df[nvkt_col_name].apply(normalize_nvkt)
            print(f"✅ Đã chuẩn hóa cột '{nvkt_col_name}'")
            print(f"   Quy tắc chuẩn hóa:")
            print(f"   1. Giữ lại phần sau dấu '-'")
            print(f"   2. Bỏ các ký tự trong dấu ()")
            print(f"   Mẫu dữ liệu sau chuẩn hóa:")
            print(f"   {df[nvkt_col_name].head().tolist()}")

        # 3. Xử lý cột NVKT_DIABAN_GIAO
        nvkt_giao_col = None
        if nvkt_col_name is not None and nvkt_col_name != 'NVKT_DIABAN_GIAO':
            # Nếu file đã có cột NVKT_DIABAN_GIAO, xóa nó và rename cột mới
            if 'NVKT_DIABAN_GIAO' in df.columns:
                df = df.drop(columns=['NVKT_DIABAN_GIAO'])
                print(f"   Đã xóa cột NVKT_DIABAN_GIAO cũ")

            df = df.rename(columns={nvkt_col_name: 'NVKT_DIABAN_GIAO'})
            print(f"\n✅ Đã đổi tên '{nvkt_col_name}' -> 'NVKT_DIABAN_GIAO'")
            nvkt_giao_col = 'NVKT_DIABAN_GIAO'
        elif nvkt_col_name == 'NVKT_DIABAN_GIAO':
            nvkt_giao_col = 'NVKT_DIABAN_GIAO'

        # 4. Lọc dữ liệu theo các điều kiện
        print(f"\n✓ Đang lọc dữ liệu theo các điều kiện...")

        # Kiểm tra các cột cần lọc tồn tại
        filter_cols = {
            'TRANGTHAI_THUHOI': 'CHUA_THU_HOI',
            'LOAI_VT': 'ONT',
            'LOAI_PHIEU': 'PTTB',
            'ONT_TBDC_CHAT_LUONG_CAO': 1,
            'LOAI_GIAO_GIAMTRU_LAN_HAI': 0
        }

        # Áp dụng các điều kiện lọc
        original_count = len(df)
        for col, value in filter_cols.items():
            if col in df.columns:
                df = df[df[col] == value]
                print(f"   ✓ Lọc {col} = {value}: {len(df)} bản ghi")
            else:
                print(f"   ⚠️ Không tìm thấy cột '{col}' trong file gốc")

        filtered_count = len(df)
        print(f"\n✅ Đã lọc dữ liệu")
        print(f"   Bản ghi gốc: {original_count}")
        print(f"   Bản ghi sau lọc: {filtered_count}")
        print(f"   Bản ghi loại bỏ: {original_count - filtered_count}")

        # 5. Sắp xếp cột
        print(f"\n✓ Đang sắp xếp cột...")
        # Đặt các cột mới lên đầu (theo thứ tự)
        new_columns = ['NVKT_DIABAN_GIAO', 'TRANGTHAI_THUHOI', 'LOAI_VT', 'LOAI_PHIEU',
                       'ONT_TBDC_CHAT_LUONG_CAO', 'LOAI_GIAO_GIAMTRU_LAN_HAI']

        # Chỉ lấy cột mới nếu tồn tại, cộng với các cột cũ
        existing_new_cols = [col for col in new_columns if col in df.columns]
        other_cols = [col for col in df.columns if col not in new_columns]

        df = df[existing_new_cols + other_cols]
        print(f"✅ Đã sắp xếp cột")

        # 6. Reset index
        df = df.reset_index(drop=True)

        # 7. Tạo sheet tổng hợp (Group by DIEMCHIA và NVKT_DIABAN_GIAO)
        print(f"\n✓ Đang tạo sheet tổng hợp...")

        # Kiểm tra cột DIEMCHIA có tồn tại không
        if 'DIEMCHIA' not in df.columns:
            print(f"   ⚠️ Không tìm thấy cột 'DIEMCHIA', bỏ qua sheet tổng hợp")
            df_summary = None
        elif 'NVKT_DIABAN_GIAO' not in df.columns:
            print(f"   ⚠️ Không tìm thấy cột 'NVKT_DIABAN_GIAO', bỏ qua sheet tổng hợp")
            df_summary = None
        else:
            try:
                # Tạo copy dataframe tạm thời để chuẩn hóa
                df_temp = df.copy()

                # Chuẩn hóa các cột trước khi group by
                if pd.api.types.is_numeric_dtype(df_temp['DIEMCHIA']):
                    df_temp['DIEMCHIA'] = df_temp['DIEMCHIA'].astype(str)
                else:
                    df_temp['DIEMCHIA'] = df_temp['DIEMCHIA'].astype(str).str.strip()

                if pd.api.types.is_numeric_dtype(df_temp['NVKT_DIABAN_GIAO']):
                    df_temp['NVKT_DIABAN_GIAO'] = df_temp['NVKT_DIABAN_GIAO'].astype(str)
                else:
                    df_temp['NVKT_DIABAN_GIAO'] = df_temp['NVKT_DIABAN_GIAO'].astype(str).str.strip()

                # Group by DIEMCHIA và NVKT_DIABAN_GIAO, đếm số lượng
                df_summary = df_temp.groupby(['DIEMCHIA', 'NVKT_DIABAN_GIAO'], as_index=False).size()
                df_summary = df_summary.rename(columns={'size': 'Số lượng'})

                # Sắp xếp theo DIEMCHIA và NVKT_DIABAN_GIAO
                df_summary = df_summary.sort_values(['DIEMCHIA', 'NVKT_DIABAN_GIAO'], ascending=[True, True])
                df_summary = df_summary.reset_index(drop=True)
                print(f"   ✅ Đã tạo sheet tổng hợp: {len(df_summary)} dòng")
            except Exception as e:
                print(f"   ⚠️ Lỗi khi tạo sheet tổng hợp: {e}")
                import traceback
                traceback.print_exc()
                df_summary = None

        # 8. Tạo sheet chi tiết vật tư
        print(f"\n✓ Đang tạo sheet chi tiết vật tư...")
        chi_tiet_vt_cols = ['DIEMCHIA', 'NVKT_DIABAN_GIAO', 'MA_TB', 'TEN_TB', 'TEN_TBI',
                            'NGAY_GIAO', 'TEN_LOAIHD', 'TEN_KIEULD', 'SO_DT', 'NGAY_SD_TB']

        # Kiểm tra các cột tồn tại
        chi_tiet_vt_cols_exist = [col for col in chi_tiet_vt_cols if col in df.columns]

        if chi_tiet_vt_cols_exist:
            df_chi_tiet_vt = df[chi_tiet_vt_cols_exist].copy()
            df_chi_tiet_vt = df_chi_tiet_vt.reset_index(drop=True)
            print(f"   ✅ Đã chuẩn bị sheet chi tiết vật tư: {len(df_chi_tiet_vt)} bản ghi")
            print(f"      Cột: {chi_tiet_vt_cols_exist}")
        else:
            print(f"   ⚠️ Không tìm thấy cột chi tiết vật tư, bỏ qua sheet này")
            df_chi_tiet_vt = None

        # 9. Lưu file đã xử lý (với 3 sheet: Dữ liệu chi tiết + Chi tiết vật tư + Tổng hợp)
        print(f"\n✓ Đang lưu file: {output_file}")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu chi tiết (đã lọc)
            df.to_excel(writer, sheet_name='Chi tiết', index=False)
            print(f"   ✅ Đã tạo sheet 'Chi tiết' ({len(df)} bản ghi)")

            # Sheet 2: Chi tiết vật tư (nếu tồn tại)
            if df_chi_tiet_vt is not None:
                df_chi_tiet_vt.to_excel(writer, sheet_name='Chi tiết vật tư', index=False)
                print(f"   ✅ Đã tạo sheet 'Chi tiết vật tư' ({len(df_chi_tiet_vt)} bản ghi)")

            # Sheet 3: Dữ liệu tổng hợp (nếu tồn tại)
            if df_summary is not None:
                df_summary.to_excel(writer, sheet_name='Tổng hợp', index=False)
                print(f"   ✅ Đã tạo sheet 'Tổng hợp' ({len(df_summary)} bản ghi)")

        print(f"✅ Đã lưu file xử lý: {output_file}")

        print(f"\n=== Tóm tắt ===")
        print(f"Sheet 'Chi tiết' - Hình dạng: {df.shape}")
        print(f"Cột: {df.columns.tolist()}")

        if df_chi_tiet_vt is not None:
            print(f"\nSheet 'Chi tiết vật tư' - Hình dạng: {df_chi_tiet_vt.shape}")
            print(f"Cột: {df_chi_tiet_vt.columns.tolist()}")

        if df_summary is not None:
            print(f"\nSheet 'Tổng hợp' - Hình dạng: {df_summary.shape}")
            print(f"Cột: {df_summary.columns.tolist()}")

        print(f"\n✅ Xử lý file vật tư thu hồi thành công!")

        return True

    except Exception as e:
        print(f"❌ Lỗi khi xử lý file vật tư thu hồi: {e}")
        import traceback
        traceback.print_exc()
        return False


def import_vat_tu_to_history():
    """Import dữ liệu vật tư thu hồi vào database history sau khi xử lý xong"""
    if HISTORY_IMPORT_AVAILABLE:
        try:
            print(f"\n💾 Đang lưu vật tư thu hồi vào database history...")
            importer = ReportsHistoryImporter()
            importer.import_vat_tu_thu_hoi()
            print(f"✅ Đã lưu vật tư thu hồi vào database history")
        except Exception as e:
            print(f"⚠️  Không thể lưu vào database history: {e}")
    else:
        print("⚠️  Module import_reports_history không khả dụng")


if __name__ == "__main__":
    # Test hàm xử lý
    vat_tu_thu_hoi_process()

    # Import vào database history
    import_vat_tu_to_history()
