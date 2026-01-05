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

def process_ngung_psc_report():
    """
    Xử lý báo cáo ngung_psc:
    1. Đọc file ngung_psc_DDMMYYYY.xlsx
    2. Chuẩn hóa tên nhân viên kỹ thuật từ cột 'Nhóm địa bàn'
    3. Lọc thông tin từ dsnv.xlsx để lấy đơn vị
    4. Ghi kết quả vào file gốc
    """
    print("\n=== Bắt đầu xử lý báo cáo Ngưng PSC ===")

    try:
        # Lấy ngày hiện tại để tìm file
        date_str = datetime.now().strftime("%d%m%Y")
        download_dir = os.path.join("downloads", "baocao_hanoi")
        ngung_psc_file = os.path.join(download_dir, f"ngung_psc_{date_str}.xlsx")
        dsnv_file = "dsnv.xlsx"

        # Kiểm tra file tồn tại
        if not os.path.exists(ngung_psc_file):
            print(f"❌ Không tìm thấy file: {ngung_psc_file}")
            return

        if not os.path.exists(dsnv_file):
            print(f"❌ Không tìm thấy file: {dsnv_file}")
            return

        print(f"Đang đọc file: {ngung_psc_file}")
        # Đọc file ngung_psc
        df_ngung_psc = pd.read_excel(ngung_psc_file)

        print(f"Đang đọc file: {dsnv_file}")
        # Đọc file dsnv
        df_dsnv = pd.read_excel(dsnv_file)

        # Chuẩn hóa tên cột (loại bỏ khoảng trắng, chuyển về lowercase để so sánh)
        df_dsnv.columns = df_dsnv.columns.str.strip()
        df_ngung_psc.columns = df_ngung_psc.columns.str.strip()

        # Tạo dictionary mapping tên cột (lowercase) -> tên cột gốc
        dsnv_col_map = {col.lower(): col for col in df_dsnv.columns}
        ngung_psc_col_map = {col.lower(): col for col in df_ngung_psc.columns}

        # Tìm cột "Nhóm địa bàn" (không phân biệt hoa thường)
        nhom_dia_ban_col = None
        for key in ['nhóm địa bàn', 'nhom dia ban']:
            if key in ngung_psc_col_map:
                nhom_dia_ban_col = ngung_psc_col_map[key]
                break

        if not nhom_dia_ban_col:
            print("❌ Không tìm thấy cột 'Nhóm địa bàn' trong file ngung_psc")
            print(f"   Các cột có sẵn: {', '.join(df_ngung_psc.columns)}")
            return

        # Tìm cột "Họ tên" trong dsnv
        ho_ten_col = None
        for key in ['họ tên', 'ho ten', 'họ và tên', 'ho va ten', 'hoten']:
            if key in dsnv_col_map:
                ho_ten_col = dsnv_col_map[key]
                break

        if not ho_ten_col:
            print("❌ Không tìm thấy cột 'Họ tên' trong file dsnv")
            print(f"   Các cột có sẵn: {', '.join(df_dsnv.columns)}")
            return

        # Tìm cột "Đơn vị" trong dsnv
        don_vi_col = None
        for key in ['đơn vị', 'don vi', 'đơn vị công tác', 'donvi']:
            if key in dsnv_col_map:
                don_vi_col = dsnv_col_map[key]
                break

        if not don_vi_col:
            print("❌ Không tìm thấy cột 'Đơn vị' trong file dsnv")
            print(f"   Các cột có sẵn: {', '.join(df_dsnv.columns)}")
            return

        print(f"✅ Tìm thấy cột: '{nhom_dia_ban_col}' trong ngung_psc")
        print(f"✅ Tìm thấy cột: '{ho_ten_col}' và '{don_vi_col}' trong dsnv")

        print("Đang chuẩn hóa tên nhân viên kỹ thuật...")

        def normalize_name(name):
            """
            Chuẩn hóa tên nhân viên:
            - Đồng Mô 4 - Đỗ Minh Thăng -> Đỗ Minh Thăng
            - VNM3-Khuất Anh Chiến( VXN) -> Khuất Anh Chiến
            """
            if pd.isna(name) or name == '':
                return ''

            name = str(name).strip()

            # Bỏ phần trong ngoặc đơn
            name = re.sub(r'\([^)]*\)', '', name)

            # Nếu có dấu '-', lấy phần sau dấu '-'
            if '-' in name:
                name = name.split('-')[-1]

            # Loại bỏ khoảng trắng thừa
            name = name.strip()

            return name

        # Áp dụng chuẩn hóa tên
        df_ngung_psc['NVKT'] = df_ngung_psc[nhom_dia_ban_col].apply(normalize_name)

        print(f"✅ Đã chuẩn hóa {len(df_ngung_psc)} tên nhân viên")

        # Tạo dictionary từ df_dsnv để lookup nhanh
        # Tạo 3 dict: exact, lowercase cho đơn vị, và lowercase cho tên chuẩn
        print("Đang tra cứu đơn vị từ danh sách nhân viên...")
        dsnv_dict_exact = dict(zip(df_dsnv[ho_ten_col].str.strip(), df_dsnv[don_vi_col]))
        dsnv_dict_lower = dict(zip(df_dsnv[ho_ten_col].str.strip().str.lower(), df_dsnv[don_vi_col]))
        # Dict để lấy tên chuẩn từ lowercase
        dsnv_dict_name = dict(zip(df_dsnv[ho_ten_col].str.strip().str.lower(), df_dsnv[ho_ten_col].str.strip()))

        # Hàm lookup thông minh: thử exact match trước, nếu không có thì thử lowercase
        def lookup_donvi(nvkt):
            if pd.isna(nvkt) or nvkt == '':
                return None, nvkt
            # Thử exact match trước
            if nvkt in dsnv_dict_exact:
                return dsnv_dict_exact[nvkt], nvkt
            # Nếu không có, thử lowercase match
            nvkt_lower = nvkt.lower()
            if nvkt_lower in dsnv_dict_lower:
                # Trả về cả đơn vị và tên chuẩn từ dsnv
                return dsnv_dict_lower[nvkt_lower], dsnv_dict_name[nvkt_lower]
            return None, nvkt

        # Lookup đơn vị và chuẩn hóa tên NVKT
        df_ngung_psc[['Đơn vị', 'NVKT']] = df_ngung_psc['NVKT'].apply(
            lambda x: pd.Series(lookup_donvi(x))
        )

        # Đếm số lượng match
        matched_count = df_ngung_psc['Đơn vị'].notna().sum()
        print(f"✅ Đã tra cứu được đơn vị cho {matched_count}/{len(df_ngung_psc)} bản ghi")

        # Hiển thị một số thống kê
        print("\n📊 Thống kê cơ bản:")
        print(f"   - Tổng số bản ghi: {len(df_ngung_psc)}")
        print(f"   - Số bản ghi có đơn vị: {matched_count}")
        print(f"   - Số bản ghi chưa có đơn vị: {len(df_ngung_psc) - matched_count}")

        if matched_count < len(df_ngung_psc):
            print("\n⚠️  Một số nhân viên chưa tìm thấy trong danh sách:")
            missing_nvkt = df_ngung_psc[df_ngung_psc['Đơn vị'].isna()]['NVKT'].unique()
            for nvkt in missing_nvkt[:10]:  # Hiển thị tối đa 10 tên
                if nvkt:
                    print(f"     - {nvkt}")
            if len(missing_nvkt) > 10:
                print(f"     ... và {len(missing_nvkt) - 10} tên khác")

        # Tạo thống kê theo Đơn vị (Tổ)
        print("\n📊 Đang tạo thống kê theo Tổ...")
        df_theo_to = df_ngung_psc.groupby('Đơn vị', dropna=False).size().reset_index(name='Số lượng TB')
        df_theo_to = df_theo_to.sort_values('Số lượng TB', ascending=False)

        # Thêm tổng
        total_row_to = pd.DataFrame([{'Đơn vị': 'TỔNG CỘNG', 'Số lượng TB': df_theo_to['Số lượng TB'].sum()}])
        df_theo_to = pd.concat([df_theo_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_theo_to) - 1} tổ")

        # Tạo thống kê theo NVKT
        print("📊 Đang tạo thống kê theo NVKT...")
        df_theo_nvkt = df_ngung_psc.groupby(['Đơn vị', 'NVKT'], dropna=False).size().reset_index(name='Số lượng TB')
        df_theo_nvkt = df_theo_nvkt.sort_values(['Đơn vị', 'Số lượng TB'], ascending=[True, False])

        # Thêm tổng
        total_row_nvkt = pd.DataFrame([{'Đơn vị': 'TỔNG CỘNG', 'NVKT': '', 'Số lượng TB': df_theo_nvkt['Số lượng TB'].sum()}])
        df_theo_nvkt = pd.concat([df_theo_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_theo_nvkt) - 1} NVKT")

        # Lưu file với nhiều sheet
        print(f"\n💾 Đang lưu file với 3 sheet...")
        with pd.ExcelWriter(ngung_psc_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu gốc
            df_ngung_psc.to_excel(writer, sheet_name='Data', index=False)

            # Sheet 2: Thống kê theo Tổ
            df_theo_to.to_excel(writer, sheet_name='ngung-psc-theo-to', index=False)

            # Sheet 3: Thống kê theo NVKT
            df_theo_nvkt.to_excel(writer, sheet_name='ngung-psc-theo-NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]

                # Auto-fit columns
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
        print(f"   - Sheet 'ngung-psc-theo-to': Thống kê theo Tổ ({len(df_theo_to)} dòng)")
        print(f"   - Sheet 'ngung-psc-theo-NVKT': Thống kê theo NVKT ({len(df_theo_nvkt)} dòng)")

        # Hiển thị top 5 tổ có nhiều TB nhất
        print("\n📊 Top 5 Tổ có nhiều TB ngưng PSC nhất:")
        top5_to = df_theo_to[df_theo_to['Đơn vị'] != 'TỔNG CỘNG'].head(5)
        for idx, row in top5_to.iterrows():
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {don_vi}: {row['Số lượng TB']} TB")

        # Hiển thị top 5 NVKT có nhiều TB nhất
        print("\n📊 Top 5 NVKT có nhiều TB ngưng PSC nhất:")
        top5_nvkt = df_theo_nvkt[df_theo_nvkt['NVKT'] != ''].head(5)
        for idx, row in top5_nvkt.iterrows():
            nvkt = row['NVKT'] if pd.notna(row['NVKT']) else '(Chưa xác định)'
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {nvkt} ({don_vi}): {row['Số lượng TB']} TB")

        print("\n✅ Hoàn thành xử lý báo cáo Ngưng PSC!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo: {e}")
        import traceback
        traceback.print_exc()

def process_hoan_cong_report():
    """
    Xử lý báo cáo hoan_cong:
    1. Đọc file hoan_cong_DDMMYYYY.xlsx
    2. Chuẩn hóa tên nhân viên kỹ thuật từ cột 'Nhân viên KT'
    3. Lọc thông tin từ dsnv.xlsx để lấy đơn vị
    4. Tạo 2 sheet thống kê và ghi vào file gốc
    """
    print("\n=== Bắt đầu xử lý báo cáo Hoàn công ===")

    try:
        # Lấy ngày hiện tại để tìm file
        date_str = datetime.now().strftime("%d%m%Y")
        download_dir = os.path.join("downloads", "baocao_hanoi")
        hoan_cong_file = os.path.join(download_dir, f"hoan_cong_{date_str}.xlsx")
        dsnv_file = "dsnv.xlsx"

        # Kiểm tra file tồn tại
        if not os.path.exists(hoan_cong_file):
            print(f"❌ Không tìm thấy file: {hoan_cong_file}")
            return

        if not os.path.exists(dsnv_file):
            print(f"❌ Không tìm thấy file: {dsnv_file}")
            return

        print(f"Đang đọc file: {hoan_cong_file}")
        # Đọc file hoan_cong
        df_hoan_cong = pd.read_excel(hoan_cong_file)

        print(f"Đang đọc file: {dsnv_file}")
        # Đọc file dsnv
        df_dsnv = pd.read_excel(dsnv_file)

        # Chuẩn hóa tên cột (loại bỏ khoảng trắng, chuyển về lowercase để so sánh)
        df_dsnv.columns = df_dsnv.columns.str.strip()
        df_hoan_cong.columns = df_hoan_cong.columns.str.strip()

        # Tạo dictionary mapping tên cột (lowercase) -> tên cột gốc
        dsnv_col_map = {col.lower(): col for col in df_dsnv.columns}
        hoan_cong_col_map = {col.lower(): col for col in df_hoan_cong.columns}

        # Tìm cột "Nhân viên KT" (không phân biệt hoa thường)
        nhan_vien_kt_col = None
        for key in ['nhân viên kt', 'nhan vien kt', 'nvkt', 'nhân viên kỹ thuật']:
            if key in hoan_cong_col_map:
                nhan_vien_kt_col = hoan_cong_col_map[key]
                break

        if not nhan_vien_kt_col:
            print("❌ Không tìm thấy cột 'Nhân viên KT' trong file hoan_cong")
            print(f"   Các cột có sẵn: {', '.join(df_hoan_cong.columns)}")
            return

        # Tìm cột "Họ tên" trong dsnv
        ho_ten_col = None
        for key in ['họ tên', 'ho ten', 'họ và tên', 'ho va ten', 'hoten']:
            if key in dsnv_col_map:
                ho_ten_col = dsnv_col_map[key]
                break

        if not ho_ten_col:
            print("❌ Không tìm thấy cột 'Họ tên' trong file dsnv")
            print(f"   Các cột có sẵn: {', '.join(df_dsnv.columns)}")
            return

        # Tìm cột "Đơn vị" trong dsnv
        don_vi_col = None
        for key in ['đơn vị', 'don vi', 'đơn vị công tác', 'donvi']:
            if key in dsnv_col_map:
                don_vi_col = dsnv_col_map[key]
                break

        if not don_vi_col:
            print("❌ Không tìm thấy cột 'Đơn vị' trong file dsnv")
            print(f"   Các cột có sẵn: {', '.join(df_dsnv.columns)}")
            return

        print(f"✅ Tìm thấy cột: '{nhan_vien_kt_col}' trong hoan_cong")
        print(f"✅ Tìm thấy cột: '{ho_ten_col}' và '{don_vi_col}' trong dsnv")

        print("Đang chuẩn hóa tên nhân viên kỹ thuật...")

        def normalize_name(name):
            """
            Chuẩn hóa tên nhân viên:
            - VNPT016763-Nguyễn Quảng Ba -> Nguyễn Quảng Ba
            - Bỏ phần trước dấu '-' và phần trong ngoặc đơn
            """
            if pd.isna(name) or name == '':
                return ''

            name = str(name).strip()

            # Bỏ phần trong ngoặc đơn
            name = re.sub(r'\([^)]*\)', '', name)

            # Nếu có dấu '-', lấy phần sau dấu '-'
            if '-' in name:
                name = name.split('-')[-1]

            # Loại bỏ khoảng trắng thừa
            name = name.strip()

            return name

        # Áp dụng chuẩn hóa tên
        df_hoan_cong['NVKT'] = df_hoan_cong[nhan_vien_kt_col].apply(normalize_name)

        print(f"✅ Đã chuẩn hóa {len(df_hoan_cong)} tên nhân viên")

        # Tạo dictionary từ df_dsnv để lookup nhanh
        # Tạo 3 dict: exact, lowercase cho đơn vị, và lowercase cho tên chuẩn
        print("Đang tra cứu đơn vị từ danh sách nhân viên...")
        dsnv_dict_exact = dict(zip(df_dsnv[ho_ten_col].str.strip(), df_dsnv[don_vi_col]))
        dsnv_dict_lower = dict(zip(df_dsnv[ho_ten_col].str.strip().str.lower(), df_dsnv[don_vi_col]))
        # Dict để lấy tên chuẩn từ lowercase
        dsnv_dict_name = dict(zip(df_dsnv[ho_ten_col].str.strip().str.lower(), df_dsnv[ho_ten_col].str.strip()))

        # Hàm lookup thông minh: thử exact match trước, nếu không có thì thử lowercase
        def lookup_donvi(nvkt):
            if pd.isna(nvkt) or nvkt == '':
                return None, nvkt
            # Thử exact match trước
            if nvkt in dsnv_dict_exact:
                return dsnv_dict_exact[nvkt], nvkt
            # Nếu không có, thử lowercase match
            nvkt_lower = nvkt.lower()
            if nvkt_lower in dsnv_dict_lower:
                # Trả về cả đơn vị và tên chuẩn từ dsnv
                return dsnv_dict_lower[nvkt_lower], dsnv_dict_name[nvkt_lower]
            return None, nvkt

        # Lookup đơn vị và chuẩn hóa tên NVKT
        df_hoan_cong[['Đơn vị', 'NVKT']] = df_hoan_cong['NVKT'].apply(
            lambda x: pd.Series(lookup_donvi(x))
        )

        # Đếm số lượng match
        matched_count = df_hoan_cong['Đơn vị'].notna().sum()
        print(f"✅ Đã tra cứu được đơn vị cho {matched_count}/{len(df_hoan_cong)} bản ghi")

        # Hiển thị một số thống kê
        print("\n📊 Thống kê cơ bản:")
        print(f"   - Tổng số bản ghi: {len(df_hoan_cong)}")
        print(f"   - Số bản ghi có đơn vị: {matched_count}")
        print(f"   - Số bản ghi chưa có đơn vị: {len(df_hoan_cong) - matched_count}")

        if matched_count < len(df_hoan_cong):
            print("\n⚠️  Một số nhân viên chưa tìm thấy trong danh sách:")
            missing_nvkt = df_hoan_cong[df_hoan_cong['Đơn vị'].isna()]['NVKT'].unique()
            for nvkt in missing_nvkt[:10]:  # Hiển thị tối đa 10 tên
                if nvkt:
                    print(f"     - {nvkt}")
            if len(missing_nvkt) > 10:
                print(f"     ... và {len(missing_nvkt) - 10} tên khác")

        # Tạo thống kê theo Đơn vị (Tổ)
        print("\n📊 Đang tạo thống kê theo Tổ...")
        df_theo_to = df_hoan_cong.groupby('Đơn vị', dropna=False).size().reset_index(name='Số lượng TB')
        df_theo_to = df_theo_to.sort_values('Số lượng TB', ascending=False)

        # Thêm tổng
        total_row_to = pd.DataFrame([{'Đơn vị': 'TỔNG CỘNG', 'Số lượng TB': df_theo_to['Số lượng TB'].sum()}])
        df_theo_to = pd.concat([df_theo_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_theo_to) - 1} tổ")

        # Tạo thống kê theo NVKT
        print("📊 Đang tạo thống kê theo NVKT...")
        df_theo_nvkt = df_hoan_cong.groupby(['Đơn vị', 'NVKT'], dropna=False).size().reset_index(name='Số lượng TB')
        df_theo_nvkt = df_theo_nvkt.sort_values(['Đơn vị', 'Số lượng TB'], ascending=[True, False])

        # Thêm tổng
        total_row_nvkt = pd.DataFrame([{'Đơn vị': 'TỔNG CỘNG', 'NVKT': '', 'Số lượng TB': df_theo_nvkt['Số lượng TB'].sum()}])
        df_theo_nvkt = pd.concat([df_theo_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_theo_nvkt) - 1} NVKT")

        # Lưu file với nhiều sheet
        print(f"\n💾 Đang lưu file với 3 sheet...")
        with pd.ExcelWriter(hoan_cong_file, engine='openpyxl') as writer:
            # Sheet 1: Dữ liệu gốc
            df_hoan_cong.to_excel(writer, sheet_name='Data', index=False)

            # Sheet 2: Thống kê theo Tổ
            df_theo_to.to_excel(writer, sheet_name='hoan-cong-theo-to', index=False)

            # Sheet 3: Thống kê theo NVKT
            df_theo_nvkt.to_excel(writer, sheet_name='hoan-cong-theo-NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]

                # Auto-fit columns
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
        print(f"   - Sheet 'hoan-cong-theo-to': Thống kê theo Tổ ({len(df_theo_to)} dòng)")
        print(f"   - Sheet 'hoan-cong-theo-NVKT': Thống kê theo NVKT ({len(df_theo_nvkt)} dòng)")

        # Hiển thị top 5 tổ có nhiều TB nhất
        print("\n📊 Top 5 Tổ có nhiều TB hoàn công nhất:")
        top5_to = df_theo_to[df_theo_to['Đơn vị'] != 'TỔNG CỘNG'].head(5)
        for idx, row in top5_to.iterrows():
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {don_vi}: {row['Số lượng TB']} TB")

        # Hiển thị top 5 NVKT có nhiều TB nhất
        print("\n📊 Top 5 NVKT có nhiều TB hoàn công nhất:")
        top5_nvkt = df_theo_nvkt[df_theo_nvkt['NVKT'] != ''].head(5)
        for idx, row in top5_nvkt.iterrows():
            nvkt = row['NVKT'] if pd.notna(row['NVKT']) else '(Chưa xác định)'
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {nvkt} ({don_vi}): {row['Số lượng TB']} TB")

        print("\n✅ Hoàn thành xử lý báo cáo Hoàn công!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo: {e}")
        import traceback
        traceback.print_exc()

def create_thuc_tang_report():
    """
    Tạo báo cáo thực tăng từ 2 báo cáo đã xử lý (Hoàn công và Ngưng PSC)
    Thực tăng = Hoàn công - Ngưng PSC

    Tạo 2 sheet:
    1. thuc_tang_theo_to: Thống kê theo Đơn vị (Tổ)
    2. thuc_tang_theo_NVKT: Thống kê theo NVKT
    """
    print("\n=== Bắt đầu tạo báo cáo Thực tăng ===")

    try:
        # Lấy ngày hiện tại để tìm file
        date_str = datetime.now().strftime("%d%m%Y")
        download_dir = os.path.join("downloads", "baocao_hanoi")

        ngung_psc_file = os.path.join(download_dir, f"ngung_psc_{date_str}.xlsx")
        hoan_cong_file = os.path.join(download_dir, f"hoan_cong_{date_str}.xlsx")
        thuc_tang_file = os.path.join(download_dir, f"thuc_tang_{date_str}.xlsx")

        # Kiểm tra file tồn tại
        if not os.path.exists(ngung_psc_file):
            print(f"❌ Không tìm thấy file: {ngung_psc_file}")
            return

        if not os.path.exists(hoan_cong_file):
            print(f"❌ Không tìm thấy file: {hoan_cong_file}")
            return

        print(f"Đang đọc dữ liệu từ file Ngưng PSC...")
        # Đọc sheet thống kê từ file Ngưng PSC
        df_ngung_psc_to = pd.read_excel(ngung_psc_file, sheet_name='ngung-psc-theo-to')
        df_ngung_psc_nvkt = pd.read_excel(ngung_psc_file, sheet_name='ngung-psc-theo-NVKT')

        print(f"Đang đọc dữ liệu từ file Hoàn công...")
        # Đọc sheet thống kê từ file Hoàn công
        df_hoan_cong_to = pd.read_excel(hoan_cong_file, sheet_name='hoan-cong-theo-to')
        df_hoan_cong_nvkt = pd.read_excel(hoan_cong_file, sheet_name='hoan-cong-theo-NVKT')

        # === XỬ LÝ SHEET 1: THỰC TĂNG THEO TỔ ===
        print("\n📊 Đang tạo báo cáo Thực tăng theo Tổ...")

        # Loại bỏ dòng TỔNG CỘNG trước khi merge
        df_ngung_psc_to_clean = df_ngung_psc_to[df_ngung_psc_to['Đơn vị'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_to_clean = df_hoan_cong_to[df_hoan_cong_to['Đơn vị'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột để phân biệt
        df_ngung_psc_to_clean.rename(columns={'Số lượng TB': 'Ngưng PSC'}, inplace=True)
        df_hoan_cong_to_clean.rename(columns={'Số lượng TB': 'Hoàn công'}, inplace=True)

        # Merge 2 dataframe theo Đơn vị
        df_thuc_tang_to = pd.merge(
            df_hoan_cong_to_clean,
            df_ngung_psc_to_clean,
            on='Đơn vị',
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Hoàn công'] - df_thuc_tang_to['Ngưng PSC']

        # Tính Tỷ lệ ngưng/psc (%)
        # Tỷ lệ = (Ngưng PSC / Hoàn công) * 100
        # Xử lý trường hợp chia cho 0
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to.apply(
            lambda row: (row['Ngưng PSC'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int cho các cột số lượng
        df_thuc_tang_to['Hoàn công'] = df_thuc_tang_to['Hoàn công'].astype(int)
        df_thuc_tang_to['Ngưng PSC'] = df_thuc_tang_to['Ngưng PSC'].astype(int)
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Thực tăng'].astype(int)
        # Làm tròn tỷ lệ 2 chữ số thập phân
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần (cao nhất trước)
        df_thuc_tang_to = df_thuc_tang_to.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc = int(df_thuc_tang_to['Hoàn công'].sum())
        total_np = int(df_thuc_tang_to['Ngưng PSC'].sum())
        total_tt = int(df_thuc_tang_to['Thực tăng'].sum())
        total_ty_le = (total_np / total_hc * 100) if total_hc != 0 else 0

        total_row_to = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'Hoàn công': total_hc,
            'Ngưng PSC': total_np,
            'Thực tăng': total_tt,
            'Tỷ lệ ngưng/psc': round(total_ty_le, 2)
        }])
        df_thuc_tang_to = pd.concat([df_thuc_tang_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_to) - 1} tổ")

        # === XỬ LÝ SHEET 2: THỰC TĂNG THEO NVKT ===
        print("📊 Đang tạo báo cáo Thực tăng theo NVKT...")

        # Loại bỏ dòng TỔNG CỘNG
        df_ngung_psc_nvkt_clean = df_ngung_psc_nvkt[df_ngung_psc_nvkt['Đơn vị'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_nvkt_clean = df_hoan_cong_nvkt[df_hoan_cong_nvkt['Đơn vị'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột
        df_ngung_psc_nvkt_clean.rename(columns={'Số lượng TB': 'Ngưng PSC'}, inplace=True)
        df_hoan_cong_nvkt_clean.rename(columns={'Số lượng TB': 'Hoàn công'}, inplace=True)

        # Merge theo Đơn vị và NVKT
        df_thuc_tang_nvkt = pd.merge(
            df_hoan_cong_nvkt_clean,
            df_ngung_psc_nvkt_clean,
            on=['Đơn vị', 'NVKT'],
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Hoàn công'] - df_thuc_tang_nvkt['Ngưng PSC']

        # Tính Tỷ lệ ngưng/psc (%)
        # Tỷ lệ = (Ngưng PSC / Hoàn công) * 100
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt.apply(
            lambda row: (row['Ngưng PSC'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int cho các cột số lượng
        df_thuc_tang_nvkt['Hoàn công'] = df_thuc_tang_nvkt['Hoàn công'].astype(int)
        df_thuc_tang_nvkt['Ngưng PSC'] = df_thuc_tang_nvkt['Ngưng PSC'].astype(int)
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Thực tăng'].astype(int)
        # Làm tròn tỷ lệ 2 chữ số thập phân
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần (cao nhất trước)
        df_thuc_tang_nvkt = df_thuc_tang_nvkt.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc_nvkt = int(df_thuc_tang_nvkt['Hoàn công'].sum())
        total_np_nvkt = int(df_thuc_tang_nvkt['Ngưng PSC'].sum())
        total_tt_nvkt = int(df_thuc_tang_nvkt['Thực tăng'].sum())
        total_ty_le_nvkt = (total_np_nvkt / total_hc_nvkt * 100) if total_hc_nvkt != 0 else 0

        total_row_nvkt = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'NVKT': '',
            'Hoàn công': total_hc_nvkt,
            'Ngưng PSC': total_np_nvkt,
            'Thực tăng': total_tt_nvkt,
            'Tỷ lệ ngưng/psc': round(total_ty_le_nvkt, 2)
        }])
        df_thuc_tang_nvkt = pd.concat([df_thuc_tang_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_nvkt) - 1} NVKT")

        # === LƯU FILE ===
        print(f"\n💾 Đang lưu file báo cáo Thực tăng...")
        with pd.ExcelWriter(thuc_tang_file, engine='openpyxl') as writer:
            # Sheet 1: Thống kê theo Tổ
            df_thuc_tang_to.to_excel(writer, sheet_name='thuc_tang_theo_to', index=False)

            # Sheet 2: Thống kê theo NVKT
            df_thuc_tang_nvkt.to_excel(writer, sheet_name='thuc_tang_theo_NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]

                # Auto-fit columns
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
        print(f"   - Sheet 'thuc_tang_theo_to': Thống kê theo Tổ ({len(df_thuc_tang_to)} dòng)")
        print(f"   - Sheet 'thuc_tang_theo_NVKT': Thống kê theo NVKT ({len(df_thuc_tang_nvkt)} dòng)")

        # Hiển thị thống kê tổng quan
        total_hoan_cong = int(df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] == 'TỔNG CỘNG']['Hoàn công'].iloc[0])
        total_ngung_psc = int(df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] == 'TỔNG CỘNG']['Ngưng PSC'].iloc[0])
        total_thuc_tang = int(df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] == 'TỔNG CỘNG']['Thực tăng'].iloc[0])

        print(f"\n📊 Tổng quan:")
        print(f"   - Tổng Hoàn công: {total_hoan_cong} TB")
        print(f"   - Tổng Ngưng PSC: {total_ngung_psc} TB")
        print(f"   - Thực tăng: {total_thuc_tang} TB")

        # Top 5 Tổ có thực tăng cao nhất
        print(f"\n📊 Top 5 Tổ có Thực tăng cao nhất:")
        top5_to = df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] != 'TỔNG CỘNG'].head(5)
        for idx, row in top5_to.iterrows():
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {don_vi}: {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng PSC']})")

        # Top 5 NVKT có thực tăng cao nhất
        print(f"\n📊 Top 5 NVKT có Thực tăng cao nhất:")
        top5_nvkt = df_thuc_tang_nvkt[df_thuc_tang_nvkt['NVKT'] != ''].nlargest(5, 'Thực tăng')
        for idx, row in top5_nvkt.iterrows():
            nvkt = row['NVKT'] if pd.notna(row['NVKT']) else '(Chưa xác định)'
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {nvkt} ({don_vi}): {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng PSC']})")

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
    - Chuẩn hóa tên NVKT từ cột "Nhóm địa bàn" (lấy tên sau dấu -)
    - Tra cứu Đơn vị từ file dsnv.xlsx
    - Tạo 2 sheet thống kê: theo Đơn vị và theo NVKT
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO MYTV NGƯNG PSC")
        print("="*80)

        # === ĐỌC FILE ===
        # Tìm file ngung_psc mới nhất
        download_dir = os.path.join("downloads", "baocao_hanoi")
        files = [f for f in os.listdir(download_dir) if f.startswith("mytv_ngung_psc_") and f.endswith(".xlsx")]

        if not files:
            print("❌ Không tìm thấy file mytv_ngung_psc_*.xlsx")
            return

        latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
        ngung_psc_file = os.path.join(download_dir, latest_file)

        print(f"📂 Đang đọc file: {ngung_psc_file}")
        df_ngung_psc = pd.read_excel(ngung_psc_file)
        print(f"✅ Đọc thành công {len(df_ngung_psc)} dòng dữ liệu")

        # Đọc file dsnv.xlsx
        dsnv_file = "dsnv.xlsx"
        print(f"\n📂 Đang đọc file danh sách nhân viên: {dsnv_file}")
        df_dsnv = pd.read_excel(dsnv_file)
        print(f"✅ Đọc thành công {len(df_dsnv)} nhân viên")

        # === CHUẨN HÓA TÊN NVKT ===
        print(f"\n🔧 Đang chuẩn hóa tên NVKT từ cột 'Nhóm địa bàn'...")

        # Tạo dictionary tra cứu từ dsnv.xlsx
        # Tìm tên cột đơn vị (có thể là 'Đơn vị' hoặc 'đơn vị')
        donvi_col = 'Đơn vị' if 'Đơn vị' in df_dsnv.columns else 'đơn vị'

        dsnv_dict_exact = dict(zip(df_dsnv['Họ tên'].str.strip(), df_dsnv[donvi_col]))
        dsnv_dict_lower = dict(zip(df_dsnv['Họ tên'].str.strip().str.lower(), df_dsnv[donvi_col]))
        dsnv_dict_name = dict(zip(df_dsnv['Họ tên'].str.strip().str.lower(), df_dsnv['Họ tên'].str.strip()))

        def extract_and_lookup_nvkt(nhom_dia_ban):
            """
            Trích xuất tên NVKT từ 'Nhóm địa bàn' (lấy phần sau dấu -)
            và tra cứu Đơn vị từ dsnv.xlsx

            Returns: (Đơn vị, NVKT chuẩn hóa)
            """
            if pd.isna(nhom_dia_ban):
                return None, None

            # Tách lấy tên sau dấu "-"
            # VD: "Sơn Tây 9 - Đỗ Mạnh Hùng" -> "Đỗ Mạnh Hùng"
            # VD: "PCT1- Nguyễn Huy Tuyến(TML)" -> "Nguyễn Huy Tuyến(TML)"
            parts = str(nhom_dia_ban).split('-', 1)
            if len(parts) < 2:
                return None, None

            nvkt = parts[1].strip()

            # Loại bỏ phần trong ngoặc nếu có
            # VD: "Nguyễn Huy Tuyến(TML)" -> "Nguyễn Huy Tuyến"
            if '(' in nvkt:
                nvkt = nvkt.split('(')[0].strip()

            # Tra cứu trong dsnv.xlsx
            # Thử exact match trước
            if nvkt in dsnv_dict_exact:
                return dsnv_dict_exact[nvkt], nvkt

            # Thử case-insensitive
            nvkt_lower = nvkt.lower()
            if nvkt_lower in dsnv_dict_lower:
                return dsnv_dict_lower[nvkt_lower], dsnv_dict_name[nvkt_lower]

            # Không tìm thấy
            return None, nvkt

        # Áp dụng hàm chuẩn hóa
        df_ngung_psc[['Đơn vị', 'NVKT']] = df_ngung_psc['Nhóm địa bàn'].apply(
            lambda x: pd.Series(extract_and_lookup_nvkt(x))
        )

        # Thống kê
        so_nvkt_co_don_vi = df_ngung_psc['Đơn vị'].notna().sum()
        so_nvkt_khong_co_don_vi = df_ngung_psc['Đơn vị'].isna().sum()

        print(f"✅ Hoàn thành chuẩn hóa:")
        print(f"   - Có Đơn vị: {so_nvkt_co_don_vi} dòng")
        print(f"   - Không tìm thấy Đơn vị: {so_nvkt_khong_co_don_vi} dòng")

        if so_nvkt_khong_co_don_vi > 0:
            print(f"\n⚠️  Danh sách NVKT không tìm thấy trong dsnv.xlsx:")
            nvkt_not_found = df_ngung_psc[df_ngung_psc['Đơn vị'].isna()]['NVKT'].unique()
            for nvkt in nvkt_not_found[:10]:
                if pd.notna(nvkt):
                    print(f"   - {nvkt}")

        # === TẠO SHEET 1: THỐNG KÊ THEO ĐƠN VỊ ===
        print(f"\n📊 Đang tạo thống kê theo Đơn vị...")

        df_ngung_psc_to = df_ngung_psc.groupby('Đơn vị', dropna=False).size().reset_index(name='Số lượng TB')
        df_ngung_psc_to = df_ngung_psc_to.fillna('(Chưa xác định)')
        df_ngung_psc_to = df_ngung_psc_to.sort_values('Số lượng TB', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_row_to = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'Số lượng TB': int(df_ngung_psc_to['Số lượng TB'].sum())
        }])
        df_ngung_psc_to = pd.concat([df_ngung_psc_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_ngung_psc_to) - 1} đơn vị")

        # === TẠO SHEET 2: THỐNG KÊ THEO NVKT ===
        print(f"📊 Đang tạo thống kê theo NVKT...")

        df_ngung_psc_nvkt = df_ngung_psc.groupby(['Đơn vị', 'NVKT'], dropna=False).size().reset_index(name='Số lượng TB')
        df_ngung_psc_nvkt = df_ngung_psc_nvkt.fillna('(Chưa xác định)')
        df_ngung_psc_nvkt = df_ngung_psc_nvkt.sort_values(['Đơn vị', 'Số lượng TB'], ascending=[True, False])

        # Thêm dòng TỔNG CỘNG
        total_row_nvkt = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'NVKT': '',
            'Số lượng TB': int(df_ngung_psc_nvkt['Số lượng TB'].sum())
        }])
        df_ngung_psc_nvkt = pd.concat([df_ngung_psc_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_ngung_psc_nvkt) - 1} NVKT")

        # === LƯU FILE ===
        processed_file = os.path.join(download_dir, latest_file.replace('.xlsx', '_processed.xlsx'))
        print(f"\n💾 Đang lưu file: {processed_file}")

        with pd.ExcelWriter(processed_file, engine='openpyxl') as writer:
            # Sheet 1: Thống kê theo Đơn vị
            df_ngung_psc_to.to_excel(writer, sheet_name='ngung_psc_theo_to', index=False)

            # Sheet 2: Thống kê theo NVKT
            df_ngung_psc_nvkt.to_excel(writer, sheet_name='ngung_psc_theo_NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]

                # Auto-fit columns
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

        print(f"✅ Đã lưu file: {processed_file}")
        print(f"   - Sheet 'ngung_psc_theo_to': {len(df_ngung_psc_to)} dòng")
        print(f"   - Sheet 'ngung_psc_theo_NVKT': {len(df_ngung_psc_nvkt)} dòng")

        print("\n✅ Hoàn thành xử lý báo cáo MyTV Ngưng PSC!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo MyTV Ngưng PSC: {e}")
        import traceback
        traceback.print_exc()

def process_mytv_hoan_cong_report():
    """
    Xử lý báo cáo mytv_hoan_cong:
    - Chuẩn hóa tên NVKT từ cột "Nhân viên KT" (lấy tên sau dấu -)
    - Tra cứu Đơn vị từ file dsnv.xlsx
    - Tạo 2 sheet thống kê: theo Đơn vị và theo NVKT
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO MYTV HOÀN CÔNG")
        print("="*80)

        # === ĐỌC FILE ===
        # Tìm file hoan_cong mới nhất
        download_dir = os.path.join("downloads", "baocao_hanoi")
        files = [f for f in os.listdir(download_dir) if f.startswith("mytv_hoan_cong_") and f.endswith(".xlsx")]

        if not files:
            print("❌ Không tìm thấy file mytv_hoan_cong_*.xlsx")
            return

        latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
        hoan_cong_file = os.path.join(download_dir, latest_file)

        print(f"📂 Đang đọc file: {hoan_cong_file}")
        df_hoan_cong = pd.read_excel(hoan_cong_file)
        print(f"✅ Đọc thành công {len(df_hoan_cong)} dòng dữ liệu")

        # Đọc file dsnv.xlsx
        dsnv_file = "dsnv.xlsx"
        print(f"\n📂 Đang đọc file danh sách nhân viên: {dsnv_file}")
        df_dsnv = pd.read_excel(dsnv_file)
        print(f"✅ Đọc thành công {len(df_dsnv)} nhân viên")

        # === CHUẨN HÓA TÊN NVKT ===
        print(f"\n🔧 Đang chuẩn hóa tên NVKT từ cột 'Nhân viên KT'...")

        # Tạo dictionary tra cứu từ dsnv.xlsx
        # Tìm tên cột đơn vị (có thể là 'Đơn vị' hoặc 'đơn vị')
        donvi_col = 'Đơn vị' if 'Đơn vị' in df_dsnv.columns else 'đơn vị'

        dsnv_dict_exact = dict(zip(df_dsnv['Họ tên'].str.strip(), df_dsnv[donvi_col]))
        dsnv_dict_lower = dict(zip(df_dsnv['Họ tên'].str.strip().str.lower(), df_dsnv[donvi_col]))
        dsnv_dict_name = dict(zip(df_dsnv['Họ tên'].str.strip().str.lower(), df_dsnv['Họ tên'].str.strip()))

        def extract_and_lookup_nvkt(nhan_vien_kt):
            """
            Trích xuất tên NVKT từ 'Nhân viên KT' (lấy phần sau dấu -)
            và tra cứu Đơn vị từ dsnv.xlsx

            Returns: (Đơn vị, NVKT chuẩn hóa)
            """
            if pd.isna(nhan_vien_kt):
                return None, None

            # Tách lấy tên sau dấu "-"
            # VD: "CTV072872-Phạm Anh Tuấn" -> "Phạm Anh Tuấn"
            # VD: "VNPT016776-Bùi Văn Duẩn" -> "Bùi Văn Duẩn"
            parts = str(nhan_vien_kt).split('-', 1)
            if len(parts) < 2:
                return None, None

            nvkt = parts[1].strip()

            # Tra cứu trong dsnv.xlsx
            # Thử exact match trước
            if nvkt in dsnv_dict_exact:
                return dsnv_dict_exact[nvkt], nvkt

            # Thử case-insensitive
            nvkt_lower = nvkt.lower()
            if nvkt_lower in dsnv_dict_lower:
                return dsnv_dict_lower[nvkt_lower], dsnv_dict_name[nvkt_lower]

            # Không tìm thấy
            return None, nvkt

        # Áp dụng hàm chuẩn hóa
        df_hoan_cong[['Đơn vị', 'NVKT']] = df_hoan_cong['Nhân viên KT'].apply(
            lambda x: pd.Series(extract_and_lookup_nvkt(x))
        )

        # Thống kê
        so_nvkt_co_don_vi = df_hoan_cong['Đơn vị'].notna().sum()
        so_nvkt_khong_co_don_vi = df_hoan_cong['Đơn vị'].isna().sum()

        print(f"✅ Hoàn thành chuẩn hóa:")
        print(f"   - Có Đơn vị: {so_nvkt_co_don_vi} dòng")
        print(f"   - Không tìm thấy Đơn vị: {so_nvkt_khong_co_don_vi} dòng")

        if so_nvkt_khong_co_don_vi > 0:
            print(f"\n⚠️  Danh sách NVKT không tìm thấy trong dsnv.xlsx:")
            nvkt_not_found = df_hoan_cong[df_hoan_cong['Đơn vị'].isna()]['NVKT'].unique()
            for nvkt in nvkt_not_found[:10]:
                if pd.notna(nvkt):
                    print(f"   - {nvkt}")

        # === TẠO SHEET 1: THỐNG KÊ THEO ĐƠN VỊ ===
        print(f"\n📊 Đang tạo thống kê theo Đơn vị...")

        df_hoan_cong_to = df_hoan_cong.groupby('Đơn vị', dropna=False).size().reset_index(name='Số lượng TB')
        df_hoan_cong_to = df_hoan_cong_to.fillna('(Chưa xác định)')
        df_hoan_cong_to = df_hoan_cong_to.sort_values('Số lượng TB', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_row_to = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'Số lượng TB': int(df_hoan_cong_to['Số lượng TB'].sum())
        }])
        df_hoan_cong_to = pd.concat([df_hoan_cong_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_hoan_cong_to) - 1} đơn vị")

        # === TẠO SHEET 2: THỐNG KÊ THEO NVKT ===
        print(f"📊 Đang tạo thống kê theo NVKT...")

        df_hoan_cong_nvkt = df_hoan_cong.groupby(['Đơn vị', 'NVKT'], dropna=False).size().reset_index(name='Số lượng TB')
        df_hoan_cong_nvkt = df_hoan_cong_nvkt.fillna('(Chưa xác định)')
        df_hoan_cong_nvkt = df_hoan_cong_nvkt.sort_values(['Đơn vị', 'Số lượng TB'], ascending=[True, False])

        # Thêm dòng TỔNG CỘNG
        total_row_nvkt = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'NVKT': '',
            'Số lượng TB': int(df_hoan_cong_nvkt['Số lượng TB'].sum())
        }])
        df_hoan_cong_nvkt = pd.concat([df_hoan_cong_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_hoan_cong_nvkt) - 1} NVKT")

        # === LƯU FILE ===
        processed_file = os.path.join(download_dir, latest_file.replace('.xlsx', '_processed.xlsx'))
        print(f"\n💾 Đang lưu file: {processed_file}")

        with pd.ExcelWriter(processed_file, engine='openpyxl') as writer:
            # Sheet 1: Thống kê theo Đơn vị
            df_hoan_cong_to.to_excel(writer, sheet_name='hoan_cong_theo_to', index=False)

            # Sheet 2: Thống kê theo NVKT
            df_hoan_cong_nvkt.to_excel(writer, sheet_name='hoan_cong_theo_NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]

                # Auto-fit columns
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

        print(f"✅ Đã lưu file: {processed_file}")
        print(f"   - Sheet 'hoan_cong_theo_to': {len(df_hoan_cong_to)} dòng")
        print(f"   - Sheet 'hoan_cong_theo_NVKT': {len(df_hoan_cong_nvkt)} dòng")

        print("\n✅ Hoàn thành xử lý báo cáo MyTV Hoàn công!")

    except Exception as e:
        print(f"❌ Lỗi khi xử lý báo cáo MyTV Hoàn công: {e}")
        import traceback
        traceback.print_exc()

def create_mytv_thuc_tang_report():
    """
    Tạo báo cáo MyTV Thực tăng = Hoàn công - Ngưng PSC
    Sử dụng dữ liệu từ 2 file processed
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU TẠO BÁO CÁO MYTV THỰC TĂNG")
        print("="*80)

        download_dir = os.path.join("downloads", "baocao_hanoi")

        # === ĐỌC DỮ LIỆU TỪ 2 FILE PROCESSED ===
        print(f"\n📂 Đang đọc dữ liệu từ các file processed...")

        # Tìm file processed mới nhất
        ngung_files = [f for f in os.listdir(download_dir) if f.startswith("mytv_ngung_psc_") and f.endswith("_processed.xlsx")]
        hoan_files = [f for f in os.listdir(download_dir) if f.startswith("mytv_hoan_cong_") and f.endswith("_processed.xlsx")]

        if not ngung_files or not hoan_files:
            print("❌ Không tìm thấy file processed. Vui lòng chạy process_mytv_ngung_psc_report() và process_mytv_hoan_cong_report() trước.")
            return

        latest_ngung = max(ngung_files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
        latest_hoan = max(hoan_files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))

        # Đọc sheet thống kê theo Đơn vị
        df_ngung_psc_to = pd.read_excel(os.path.join(download_dir, latest_ngung), sheet_name='ngung_psc_theo_to')
        df_hoan_cong_to = pd.read_excel(os.path.join(download_dir, latest_hoan), sheet_name='hoan_cong_theo_to')

        # Đọc sheet thống kê theo NVKT
        df_ngung_psc_nvkt = pd.read_excel(os.path.join(download_dir, latest_ngung), sheet_name='ngung_psc_theo_NVKT')
        df_hoan_cong_nvkt = pd.read_excel(os.path.join(download_dir, latest_hoan), sheet_name='hoan_cong_theo_NVKT')

        print(f"✅ Đã đọc dữ liệu từ các file processed")

        # === XỬ LÝ SHEET 1: THỰC TĂNG THEO ĐƠN VỊ ===
        print(f"\n📊 Đang tạo báo cáo Thực tăng theo Đơn vị...")

        # Loại bỏ dòng TỔNG CỘNG
        df_ngung_psc_to_clean = df_ngung_psc_to[df_ngung_psc_to['Đơn vị'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_to_clean = df_hoan_cong_to[df_hoan_cong_to['Đơn vị'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột
        df_ngung_psc_to_clean.rename(columns={'Số lượng TB': 'Ngưng PSC'}, inplace=True)
        df_hoan_cong_to_clean.rename(columns={'Số lượng TB': 'Hoàn công'}, inplace=True)

        # Merge 2 dataframe
        df_thuc_tang_to = pd.merge(
            df_hoan_cong_to_clean,
            df_ngung_psc_to_clean,
            on='Đơn vị',
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Hoàn công'] - df_thuc_tang_to['Ngưng PSC']

        # Tính Tỷ lệ ngưng/psc (%)
        # Tỷ lệ = (Ngưng PSC / Hoàn công) * 100
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to.apply(
            lambda row: (row['Ngưng PSC'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int cho các cột số lượng
        df_thuc_tang_to['Hoàn công'] = df_thuc_tang_to['Hoàn công'].astype(int)
        df_thuc_tang_to['Ngưng PSC'] = df_thuc_tang_to['Ngưng PSC'].astype(int)
        df_thuc_tang_to['Thực tăng'] = df_thuc_tang_to['Thực tăng'].astype(int)
        # Làm tròn tỷ lệ 2 chữ số thập phân
        df_thuc_tang_to['Tỷ lệ ngưng/psc'] = df_thuc_tang_to['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần (cao nhất trước)
        df_thuc_tang_to = df_thuc_tang_to.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc = int(df_thuc_tang_to['Hoàn công'].sum())
        total_np = int(df_thuc_tang_to['Ngưng PSC'].sum())
        total_tt = int(df_thuc_tang_to['Thực tăng'].sum())
        total_ty_le = (total_np / total_hc * 100) if total_hc != 0 else 0

        total_row_to = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'Hoàn công': total_hc,
            'Ngưng PSC': total_np,
            'Thực tăng': total_tt,
            'Tỷ lệ ngưng/psc': round(total_ty_le, 2)
        }])
        df_thuc_tang_to = pd.concat([df_thuc_tang_to, total_row_to], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_to) - 1} đơn vị")

        # === XỬ LÝ SHEET 2: THỰC TĂNG THEO NVKT ===
        print(f"📊 Đang tạo báo cáo Thực tăng theo NVKT...")

        # Loại bỏ dòng TỔNG CỘNG
        df_ngung_psc_nvkt_clean = df_ngung_psc_nvkt[df_ngung_psc_nvkt['Đơn vị'] != 'TỔNG CỘNG'].copy()
        df_hoan_cong_nvkt_clean = df_hoan_cong_nvkt[df_hoan_cong_nvkt['Đơn vị'] != 'TỔNG CỘNG'].copy()

        # Đổi tên cột
        df_ngung_psc_nvkt_clean.rename(columns={'Số lượng TB': 'Ngưng PSC'}, inplace=True)
        df_hoan_cong_nvkt_clean.rename(columns={'Số lượng TB': 'Hoàn công'}, inplace=True)

        # Merge theo Đơn vị và NVKT
        df_thuc_tang_nvkt = pd.merge(
            df_hoan_cong_nvkt_clean,
            df_ngung_psc_nvkt_clean,
            on=['Đơn vị', 'NVKT'],
            how='outer'
        ).fillna(0)

        # Tính Thực tăng
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Hoàn công'] - df_thuc_tang_nvkt['Ngưng PSC']

        # Tính Tỷ lệ ngưng/psc (%)
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt.apply(
            lambda row: (row['Ngưng PSC'] / row['Hoàn công'] * 100) if row['Hoàn công'] != 0 else 0,
            axis=1
        )

        # Chuyển về kiểu int cho các cột số lượng
        df_thuc_tang_nvkt['Hoàn công'] = df_thuc_tang_nvkt['Hoàn công'].astype(int)
        df_thuc_tang_nvkt['Ngưng PSC'] = df_thuc_tang_nvkt['Ngưng PSC'].astype(int)
        df_thuc_tang_nvkt['Thực tăng'] = df_thuc_tang_nvkt['Thực tăng'].astype(int)
        # Làm tròn tỷ lệ 2 chữ số thập phân
        df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'] = df_thuc_tang_nvkt['Tỷ lệ ngưng/psc'].round(2)

        # Sắp xếp theo Thực tăng giảm dần (cao nhất trước)
        df_thuc_tang_nvkt = df_thuc_tang_nvkt.sort_values('Thực tăng', ascending=False)

        # Thêm dòng TỔNG CỘNG
        total_hc_nvkt = int(df_thuc_tang_nvkt['Hoàn công'].sum())
        total_np_nvkt = int(df_thuc_tang_nvkt['Ngưng PSC'].sum())
        total_tt_nvkt = int(df_thuc_tang_nvkt['Thực tăng'].sum())
        total_ty_le_nvkt = (total_np_nvkt / total_hc_nvkt * 100) if total_hc_nvkt != 0 else 0

        total_row_nvkt = pd.DataFrame([{
            'Đơn vị': 'TỔNG CỘNG',
            'NVKT': '',
            'Hoàn công': total_hc_nvkt,
            'Ngưng PSC': total_np_nvkt,
            'Thực tăng': total_tt_nvkt,
            'Tỷ lệ ngưng/psc': round(total_ty_le_nvkt, 2)
        }])
        df_thuc_tang_nvkt = pd.concat([df_thuc_tang_nvkt, total_row_nvkt], ignore_index=True)

        print(f"✅ Đã tạo thống kê cho {len(df_thuc_tang_nvkt) - 1} NVKT")

        # === LƯU FILE ===
        date_str = datetime.now().strftime("%d%m%Y")
        thuc_tang_file = os.path.join(download_dir, f"mytv_thuc_tang_{date_str}.xlsx")
        print(f"\n💾 Đang lưu file báo cáo Thực tăng...")
        with pd.ExcelWriter(thuc_tang_file, engine='openpyxl') as writer:
            # Sheet 1: Thống kê theo Đơn vị
            df_thuc_tang_to.to_excel(writer, sheet_name='thuc_tang_theo_to', index=False)

            # Sheet 2: Thống kê theo NVKT
            df_thuc_tang_nvkt.to_excel(writer, sheet_name='thuc_tang_theo_NVKT', index=False)

            # Định dạng các sheet
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]

                # Auto-fit columns
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
        print(f"   - Sheet 'thuc_tang_theo_to': Thống kê theo Đơn vị ({len(df_thuc_tang_to)} dòng)")
        print(f"   - Sheet 'thuc_tang_theo_NVKT': Thống kê theo NVKT ({len(df_thuc_tang_nvkt)} dòng)")

        # Hiển thị thống kê tổng quan
        total_hoan_cong = int(df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] == 'TỔNG CỘNG']['Hoàn công'].iloc[0])
        total_ngung_psc = int(df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] == 'TỔNG CỘNG']['Ngưng PSC'].iloc[0])
        total_thuc_tang = int(df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] == 'TỔNG CỘNG']['Thực tăng'].iloc[0])

        print(f"\n📊 Tổng quan:")
        print(f"   - Tổng Hoàn công: {total_hoan_cong} TB")
        print(f"   - Tổng Ngưng PSC: {total_ngung_psc} TB")
        print(f"   - Thực tăng: {total_thuc_tang} TB")

        # Top 5 Đơn vị có thực tăng cao nhất
        print(f"\n📊 Top 5 Đơn vị có Thực tăng cao nhất:")
        top5_to = df_thuc_tang_to[df_thuc_tang_to['Đơn vị'] != 'TỔNG CỘNG'].head(5)
        for idx, row in top5_to.iterrows():
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {don_vi}: {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng PSC']})")

        # Top 5 NVKT có thực tăng cao nhất
        print(f"\n📊 Top 5 NVKT có Thực tăng cao nhất:")
        top5_nvkt = df_thuc_tang_nvkt[df_thuc_tang_nvkt['NVKT'] != ''].nlargest(5, 'Thực tăng')
        for idx, row in top5_nvkt.iterrows():
            nvkt = row['NVKT'] if pd.notna(row['NVKT']) else '(Chưa xác định)'
            don_vi = row['Đơn vị'] if pd.notna(row['Đơn vị']) else '(Chưa xác định)'
            print(f"   {idx + 1}. {nvkt} ({don_vi}): {row['Thực tăng']} TB (HC: {row['Hoàn công']}, NP: {row['Ngưng PSC']})")

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
