# -*- coding: utf-8 -*-
import pandas as pd
import os
import sqlite3
import math


def process_c11_report():
    """
    Xử lý báo cáo C1.1:
    1. Đọc file c1.1 report.xlsx
    2. Xóa hàng đầu tiên
    3. Loại bỏ 2 dòng: Tổ Kỹ thuật Địa bàn Bất Bạt và Tổ Kỹ thuật Địa bàn Tùng Thiện
    4. Loại bỏ từ cột thứ 9 trở đi
    5. Ghi lại vào sheet mới với tên TH_C1.1
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO C1.1")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.1 report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Xóa hàng đầu tiên
        print("\n✓ Đang xóa hàng đầu tiên...")
        df = df.iloc[1:].reset_index(drop=True)
        print(f"✅ Đã xóa hàng đầu tiên, còn lại {len(df)} dòng")

        # Loại bỏ 2 dòng chứa "Tổ Kỹ thuật Địa bàn Bất Bạt" và "Tổ Kỹ thuật Địa bàn Tùng Thiện"
        print("\n✓ Đang loại bỏ dòng 'Tổ Kỹ thuật Địa bàn Bất Bạt' và 'Tổ Kỹ thuật Địa bàn Tùng Thiện'...")

        # Tìm và loại bỏ các dòng
        df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("Tổ Kỹ thuật Địa bàn Bất Bạt|Tổ Kỹ thuật Địa bàn Tùng Thiện", na=False)]

        rows_removed = len(df) - len(df_filtered)
        print(f"✅ Đã loại bỏ {rows_removed} dòng, còn lại {len(df_filtered)} dòng")

        # Loại bỏ từ cột thứ 9 trở đi (giữ lại 8 cột đầu tiên)
        print("\n✓ Đang loại bỏ từ cột thứ 9 trở đi...")
        df_filtered = df_filtered.iloc[:, :8]
        print(f"✅ Đã giữ lại {df_filtered.shape[1]} cột đầu tiên")

        # Đặt tên cột (header)
        df_filtered.columns = [
            "Đơn vị",
            "SM1",
            "SM2",
            "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",
            "SM3",
            "SM4",
            "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn",
            "Chỉ tiêu BSC"
        ]

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_C1.1'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtered.to_excel(writer, sheet_name='TH_C1.1', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_C1.1'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.1")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.1: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c12_report():
    """
    Xử lý báo cáo C1.2:
    1. Đọc file c1.2 report.xlsx
    2. Xóa hàng đầu tiên
    3. Loại bỏ 2 dòng: Tổ Kỹ thuật Địa bàn Bất Bạt và Tổ Kỹ thuật Địa bàn Tùng Thiện
    4. Loại bỏ từ cột thứ 9 trở đi
    5. Ghi lại vào sheet mới với tên TH_C1.2
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO C1.2")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.2 report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Xóa hàng đầu tiên
        print("\n✓ Đang xóa hàng đầu tiên...")
        df = df.iloc[1:].reset_index(drop=True)
        print(f"✅ Đã xóa hàng đầu tiên, còn lại {len(df)} dòng")

        # Loại bỏ 2 dòng chứa "Tổ Kỹ thuật Địa bàn Bất Bạt" và "Tổ Kỹ thuật Địa bàn Tùng Thiện"
        print("\n✓ Đang loại bỏ dòng 'Tổ Kỹ thuật Địa bàn Bất Bạt' và 'Tổ Kỹ thuật Địa bàn Tùng Thiện'...")

        # Tìm và loại bỏ các dòng
        df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("Tổ Kỹ thuật Địa bàn Bất Bạt|Tổ Kỹ thuật Địa bàn Tùng Thiện", na=False)]

        rows_removed = len(df) - len(df_filtered)
        print(f"✅ Đã loại bỏ {rows_removed} dòng, còn lại {len(df_filtered)} dòng")

        # Loại bỏ từ cột thứ 9 trở đi (giữ lại 8 cột đầu tiên)
        print("\n✓ Đang loại bỏ từ cột thứ 9 trở đi...")
        df_filtered = df_filtered.iloc[:, :8]
        print(f"✅ Đã giữ lại {df_filtered.shape[1]} cột đầu tiên")

        # Đặt tên cột (header)
        df_filtered.columns = [
            "Đơn vị",
            "SM1",
            "SM2",
            "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",
            "SM3",
            "SM4",
            "Tỷ lệ sự cố dịch vụ BRCĐ",
            "Chỉ tiêu BSC"
        ]

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_C1.2'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtered.to_excel(writer, sheet_name='TH_C1.2', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_C1.2'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.2")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.2: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c13_report():
    """
    Xử lý báo cáo C1.3:
    1. Đọc file c1.3 report.xlsx
    2. Xóa hàng đầu tiên
    3. Loại bỏ 2 dòng: Tổ Kỹ thuật Địa bàn Bất Bạt và Tổ Kỹ thuật Địa bàn Tùng Thiện
    4. Loại bỏ từ cột thứ 11 trở đi (giữ lại 10 cột đầu tiên)
    5. Ghi lại vào sheet mới với tên TH_C1.3
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO C1.3")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.3 report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Xóa hàng đầu tiên
        print("\n✓ Đang xóa hàng đầu tiên...")
        df = df.iloc[1:].reset_index(drop=True)
        print(f"✅ Đã xóa hàng đầu tiên, còn lại {len(df)} dòng")

        # Loại bỏ 2 dòng chứa "Tổ Kỹ thuật Địa bàn Bất Bạt" và "Tổ Kỹ thuật Địa bàn Tùng Thiện"
        print("\n✓ Đang loại bỏ dòng 'Tổ Kỹ thuật Địa bàn Bất Bạt' và 'Tổ Kỹ thuật Địa bàn Tùng Thiện'...")

        # Tìm và loại bỏ các dòng
        df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("Tổ Kỹ thuật Địa bàn Bất Bạt|Tổ Kỹ thuật Địa bàn Tùng Thiện", na=False)]

        rows_removed = len(df) - len(df_filtered)
        print(f"✅ Đã loại bỏ {rows_removed} dòng, còn lại {len(df_filtered)} dòng")

        # Loại bỏ từ cột thứ 12 trở đi (giữ lại 11 cột đầu tiên)
        print("\n✓ Đang loại bỏ từ cột thứ 12 trở đi...")
        df_filtered = df_filtered.iloc[:, :11]
        print(f"✅ Đã giữ lại {df_filtered.shape[1]} cột đầu tiên")

        # Đặt tên cột (header)
        df_filtered.columns = [
            "Đơn vị",
            "SM1",
            "SM2",
            "Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định",
            "SM3",
            "SM4",
            "Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại",
            "SM5",
            "SM6",
            "Tỷ lệ sự cố dịch vụ kênh TSL",
            "Chỉ tiêu BSC"
        ]

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_C1.3'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtered.to_excel(writer, sheet_name='TH_C1.3', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_C1.3'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.3")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.3: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c14_report():
    """
    Xử lý báo cáo C1.4:
    1. Đọc file c1.4 report.xlsx
    2. Xóa hàng đầu tiên
    3. Xóa 4 cột F, G, H, I
    4. Thêm dòng Tổng với các công thức tính toán
    5. Ghi lại vào sheet mới với tên TH_C1.4
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO C1.4")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.4 report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}, tổng số cột: {df.shape[1]}")

        # Xóa hàng đầu tiên
        print("\n✓ Đang xóa hàng đầu tiên...")
        df = df.iloc[1:].reset_index(drop=True)
        print(f"✅ Đã xóa hàng đầu tiên, còn lại {len(df)} dòng")

        # Xóa 4 cột F, G, H, I (index 5, 6, 7, 8 - vì index bắt đầu từ 0)
        print("\n✓ Đang xóa 4 cột F, G, H, I...")

        # Lấy danh sách tất cả các cột
        all_columns = list(range(df.shape[1]))

        # Loại bỏ các cột F, G, H, I (index 5, 6, 7, 8)
        columns_to_keep = [i for i in all_columns if i not in [5, 6, 7, 8]]

        # Giữ lại các cột còn lại
        df_filtered = df.iloc[:, columns_to_keep]

        print(f"✅ Đã xóa 4 cột F, G, H, I, còn lại {df_filtered.shape[1]} cột")

        # Thêm dòng Tổng
        print("\n✓ Đang thêm dòng Tổng...")

        # Chuyển đổi các cột số về kiểu numeric (bỏ qua lỗi)
        # Cấu trúc cột ĐÚNG (sau khi xóa F, G, H, I):
        # Col 0: Đơn vị
        # Col 1: Tổng phiếu
        # Col 2: SL đã KS
        # Col 3: SL KS thành công
        # Col 4: SL KH hài lòng
        # Col 5: Không HL KT phục vụ
        # Col 6: Tỷ lệ HL KT phục vụ
        # Col 7: Không HL KT dịch vụ
        # Col 8: Tỷ lệ HL KT dịch vụ
        # Col 9: Tổng phiếu hài lòng KT
        # Col 10: Tỷ lệ KH hài lòng
        # Col 11: Điểm BSC

        # Chuyển đổi tất cả các cột số về numeric
        for i in range(1, df_filtered.shape[1]):
            df_filtered.iloc[:, i] = pd.to_numeric(df_filtered.iloc[:, i], errors='coerce').fillna(0)

        # Tính tổng cho các cột số liệu
        tong_phieu = df_filtered.iloc[:, 1].sum()  # Col 1: Tổng phiếu
        sl_da_ks = df_filtered.iloc[:, 2].sum()  # Col 2: SL đã KS
        sl_ks_thanh_cong = df_filtered.iloc[:, 3].sum()  # Col 3: SL KS thành công
        sl_kh_hai_long = df_filtered.iloc[:, 4].sum()  # Col 4: SL KH hài lòng
        khong_hl_kt_phuc_vu = df_filtered.iloc[:, 5].sum()  # Col 5: Không HL KT phục vụ
        khong_hl_kt_dich_vu = df_filtered.iloc[:, 7].sum()  # Col 7: Không HL KT dịch vụ
        tong_phieu_hai_long_kt = df_filtered.iloc[:, 9].sum()  # Col 9: Tổng phiếu hài lòng KT

        # Tính Tỷ lệ HL KT phục vụ = SL KH hài lòng / SL KS thành công
        ty_le_hl_kt_phuc_vu = round((sl_kh_hai_long / sl_ks_thanh_cong * 100), 2) if sl_ks_thanh_cong > 0 else 0

        # Tính Tỷ lệ HL KT dịch vụ = Không HL KT dịch vụ / SL KS thành công
        ty_le_hl_kt_dich_vu = round((khong_hl_kt_dich_vu / sl_ks_thanh_cong * 100), 2) if sl_ks_thanh_cong > 0 else 0

        # Tính tỷ lệ KH hài lòng = Tổng phiếu hài lòng KT / SL KS thành công
        ty_le_kh_hai_long = round((tong_phieu_hai_long_kt / sl_ks_thanh_cong * 100), 2) if sl_ks_thanh_cong > 0 else 0

        # Tính Điểm BSC dựa trên Tỷ lệ KH hài lòng
        if ty_le_kh_hai_long >= 99.5:
            diem_bsc = 5
        elif ty_le_kh_hai_long <= 95:
            diem_bsc = 1
        else:
            # Công thức: 1 + 4 * (TYLE_KH_HAILONG - 95) / 4.5
            diem_bsc = round(1 + 4 * (ty_le_kh_hai_long - 95) / 4.5, 2)

        # Tạo dòng Tổng - sử dụng index số để giữ nguyên cấu trúc cột
        tong_row = pd.Series({
            df_filtered.columns[0]: "Tổng",  # Đơn vị
            df_filtered.columns[1]: tong_phieu,  # Tổng phiếu
            df_filtered.columns[2]: sl_da_ks,  # SL đã KS
            df_filtered.columns[3]: sl_ks_thanh_cong,  # SL KS thành công
            df_filtered.columns[4]: sl_kh_hai_long,  # SL KH hài lòng
            df_filtered.columns[5]: khong_hl_kt_phuc_vu,  # Không HL KT phục vụ
            df_filtered.columns[6]: ty_le_hl_kt_phuc_vu,  # Tỷ lệ HL KT phục vụ
            df_filtered.columns[7]: khong_hl_kt_dich_vu,  # Không HL KT dịch vụ
            df_filtered.columns[8]: ty_le_hl_kt_dich_vu,  # Tỷ lệ HL KT dịch vụ
            df_filtered.columns[9]: tong_phieu_hai_long_kt,  # Tổng phiếu hài lòng KT
            df_filtered.columns[10]: ty_le_kh_hai_long,  # Tỷ lệ KH hài lòng
            df_filtered.columns[11]: diem_bsc  # Điểm BSC
        })

        # Thêm dòng Tổng vào DataFrame
        df_filtered = pd.concat([df_filtered, tong_row.to_frame().T], ignore_index=True)

        print(f"✅ Đã thêm dòng Tổng với các giá trị tính toán")
        print(f"   - Tổng phiếu: {tong_phieu}")
        print(f"   - SL KS thành công: {sl_ks_thanh_cong}")
        print(f"   - Tổng phiếu hài lòng KT: {tong_phieu_hai_long_kt}")
        print(f"   - Tỷ lệ KH hài lòng: {ty_le_kh_hai_long}%")
        print(f"   - Tỷ lệ HL KT phục vụ: {ty_le_hl_kt_phuc_vu}%")
        print(f"   - Tỷ lệ HL KT dịch vụ: {ty_le_hl_kt_dich_vu}%")
        print(f"   - Điểm BSC: {diem_bsc}")

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_C1.4'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtered.to_excel(writer, sheet_name='TH_C1.4', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_C1.4'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.4")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.4: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c14_chitiet_report():
    """
    Xử lý báo cáo chi tiết C1.4:
    1. Đọc file c1.4_chitiet_report.xlsx
    2. Chuẩn hóa cột TEN_KV để lấy tên NVKT:
       - TMH4-Bùi Văn Duẩn(PGT) -> Bùi Văn Duẩn
       - Đồng Mô 1 - Nguyễn Văn Minh -> Nguyễn Văn Minh
    3. Tạo sheet 'TH_HL_NVKT' với các cột:
       - DOIVT: Tên đội (nếu có)
       - NVKT: Tên NVKT sau khi chuẩn hóa
       - Tổng phiếu KS thành công: Đếm số bản ghi theo NVKT với điều kiện DO_HL = 'HL' hoặc 'KHL'
       - Tổng phiếu KHL: Đếm số bản ghi theo NVKT với điều kiện KHL_KT != null
       - Tỉ lệ HL NVKT: (Số bản ghi KHL_KT = null / Số bản ghi DO_HL = 'HL' hoặc 'KHL') * 100
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.4")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.4_chitiet_report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Kiểm tra các cột cần thiết
        required_columns = ['TEN_KV', 'DO_HL', 'KHL_KT']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            print(f"❌ Không tìm thấy các cột: {', '.join(missing_columns)}")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False

        # Kiểm tra cột DOIVT
        has_doivt = 'DOIVT' in df.columns
        if not has_doivt:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'DOIVT' trong file")

        # Chuẩn hóa cột TEN_KV để lấy tên NVKT
        print("\n✓ Đang chuẩn hóa cột TEN_KV để lấy tên NVKT...")

        def extract_nvkt_name(ten_kv):
            """
            Trích xuất tên NVKT từ cột TEN_KV
            Ví dụ:
            - TMH4-Bùi Văn Duẩn(PGT) -> Bùi Văn Duẩn
            - Đồng Mô 1 - Nguyễn Văn Minh -> Nguyễn Văn Minh
            """
            if pd.isna(ten_kv):
                return None

            ten_kv = str(ten_kv).strip()

            # Trường hợp có dấu "-"
            if '-' in ten_kv:
                # Lấy phần sau dấu "-" cuối cùng
                parts = ten_kv.split('-')
                nvkt_name = parts[-1].strip()
            else:
                nvkt_name = ten_kv

            # Loại bỏ phần trong ngoặc đơn (ví dụ: (PGT))
            if '(' in nvkt_name:
                nvkt_name = nvkt_name.split('(')[0].strip()

            return nvkt_name

        # Áp dụng hàm chuẩn hóa
        df['NVKT'] = df['TEN_KV'].apply(extract_nvkt_name)
        print(f"✅ Đã chuẩn hóa cột TEN_KV cho {len(df)} dòng")

        # Tạo báo cáo tổng hợp theo DOIVT và NVKT
        if has_doivt:
            print("\n✓ Đang tạo báo cáo tổng hợp theo DOIVT và NVKT...")
        else:
            print("\n✓ Đang tạo báo cáo tổng hợp theo NVKT...")

        # Nhóm theo DOIVT và NVKT (hoặc chỉ NVKT)
        report_data = []

        if has_doivt:
            # Nhóm theo cả DOIVT và NVKT
            for (doivt, nvkt) in df.groupby(['DOIVT', 'NVKT']).groups.keys():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo DOIVT và NVKT
                df_group = df[(df['DOIVT'] == doivt) & (df['NVKT'] == nvkt)]

                # Tổng phiếu KS thành công: đếm số bản ghi có DO_HL = 'HL' hoặc 'KHL'
                tong_phieu_ks_thanh_cong = len(df_group[df_group['DO_HL'].isin(['HL', 'KHL'])])

                # Tổng phiếu KHL: đếm số bản ghi có KHL_KT != null
                tong_phieu_khl = len(df_group[df_group['KHL_KT'].notna()])

                # Tỉ lệ HL NVKT: (Số bản ghi KHL_KT = null / Số bản ghi DO_HL = 'HL' hoặc 'KHL') * 100
                # KHL_KT = null nghĩa là hài lòng
                so_phieu_hai_long = len(df_group[(df_group['DO_HL'].isin(['HL', 'KHL'])) & (df_group['KHL_KT'].isna())])
                # Nếu chưa có phiếu KS nào, mặc định tỷ lệ hài lòng = 100%
                ty_le_hl = round((so_phieu_hai_long / tong_phieu_ks_thanh_cong * 100), 2) if tong_phieu_ks_thanh_cong > 0 else 100

                report_data.append({
                    'DOIVT': doivt,
                    'NVKT': nvkt,
                    'Tổng phiếu KS thành công': tong_phieu_ks_thanh_cong,
                    'Tổng phiếu KHL': tong_phieu_khl,
                    'Tỉ lệ HL NVKT (%)': ty_le_hl
                })
        else:
            # Chỉ nhóm theo NVKT
            for nvkt in df['NVKT'].unique():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo NVKT
                df_nvkt = df[df['NVKT'] == nvkt]

                # Tổng phiếu KS thành công: đếm số bản ghi có DO_HL = 'HL' hoặc 'KHL'
                tong_phieu_ks_thanh_cong = len(df_nvkt[df_nvkt['DO_HL'].isin(['HL', 'KHL'])])

                # Tổng phiếu KHL: đếm số bản ghi có KHL_KT != null
                tong_phieu_khl = len(df_nvkt[df_nvkt['KHL_KT'].notna()])

                # Tỉ lệ HL NVKT: (Số bản ghi KHL_KT = null / Số bản ghi DO_HL = 'HL' hoặc 'KHL') * 100
                # KHL_KT = null nghĩa là hài lòng
                so_phieu_hai_long = len(df_nvkt[(df_nvkt['DO_HL'].isin(['HL', 'KHL'])) & (df_nvkt['KHL_KT'].isna())])
                # Nếu chưa có phiếu KS nào, mặc định tỷ lệ hài lòng = 100%
                ty_le_hl = round((so_phieu_hai_long / tong_phieu_ks_thanh_cong * 100), 2) if tong_phieu_ks_thanh_cong > 0 else 100

                report_data.append({
                    'NVKT': nvkt,
                    'Tổng phiếu KS thành công': tong_phieu_ks_thanh_cong,
                    'Tổng phiếu KHL': tong_phieu_khl,
                    'Tỉ lệ HL NVKT (%)': ty_le_hl
                })

        # Tạo DataFrame từ dữ liệu tổng hợp
        df_report = pd.DataFrame(report_data)

        # Sắp xếp theo DOIVT và NVKT (hoặc chỉ NVKT)
        if has_doivt:
            df_report = df_report.sort_values(['DOIVT', 'NVKT']).reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} nhóm DOIVT - NVKT")
        else:
            df_report = df_report.sort_values('NVKT').reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} NVKT")

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_HL_NVKT'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_report.to_excel(writer, sheet_name='TH_HL_NVKT', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_HL_NVKT'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN:")
        if has_doivt:
            print(f"  - Tổng số DOIVT: {df_report['DOIVT'].nunique()}")
            print(f"  - Tổng số nhóm DOIVT - NVKT: {len(df_report)}")
        else:
            print(f"  - Tổng số NVKT: {len(df_report)}")
        print(f"  - Tổng số phiếu KS thành công: {df_report['Tổng phiếu KS thành công'].sum()}")
        print(f"  - Tổng số phiếu KHL: {df_report['Tổng phiếu KHL'].sum()}")
        ty_le_hl_chung = round(
            ((df_report['Tổng phiếu KS thành công'].sum() - df_report['Tổng phiếu KHL'].sum()) /
             df_report['Tổng phiếu KS thành công'].sum() * 100), 2
        ) if df_report['Tổng phiếu KS thành công'].sum() > 0 else 0
        print(f"  - Tỷ lệ HL chung: {ty_le_hl_chung}%")
        print("-"*80)

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO CHI TIẾT C1.4")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo chi tiết C1.4: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_c15_report():
    """
    Xử lý báo cáo C1.5:
    1. Đọc file c1.5 report.xlsx
    2. Xóa 1 dòng đầu tiên (dòng header thừa)
    3. Chỉ giữ lại các cột A, B, C, D, E (5 cột đầu tiên)
    4. Đặt header mới
    5. Ghi lại vào sheet mới với tên TH_C1.5
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO C1.5")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.5 report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}, tổng số cột: {df.shape[1]}")

        # Xóa 1 dòng đầu tiên (dòng header thừa)
        print("\n✓ Đang xóa 1 dòng đầu tiên...")
        df = df.iloc[1:].reset_index(drop=True)
        print(f"✅ Đã xóa 1 dòng đầu tiên, còn lại {len(df)} dòng")

        # Chỉ giữ lại 5 cột đầu tiên (A, B, C, D, E - index 0, 1, 2, 3, 4)
        print("\n✓ Đang giữ lại 5 cột A, B, C, D, E...")
        df_filtered = df.iloc[:, :5]
        print(f"✅ Đã giữ lại {df_filtered.shape[1]} cột đầu tiên")

        # Đặt tên cột (header)
        print("\n✓ Đang đặt header mới...")
        df_filtered.columns = [
            "Đơn vị",
            "SM1",
            "SM2",
            "KQ thực hiện chỉ tiêu",
            "Điểm BSC"
        ]
        print("✅ Đã đặt header mới")

        # Thêm hàng Tổng
        print("\n✓ Đang thêm hàng Tổng...")

        # Tính tổng SM1 và SM2
        tong_sm1 = pd.to_numeric(df_filtered["SM1"], errors='coerce').sum()
        tong_sm2 = pd.to_numeric(df_filtered["SM2"], errors='coerce').sum()

        # Tính KQ thực hiện chỉ tiêu (dạng phần trăm, làm tròn 2 chữ số thập phân)
        kq_thuc_hien = ((tong_sm1 / tong_sm2) * 100) if tong_sm2 > 0 else 0
        kq_thuc_hien_rounded = round(kq_thuc_hien, 2)

        # Tính Điểm BSC theo công thức (kq_thuc_hien_rounded đã ở dạng phần trăm)
        if kq_thuc_hien_rounded >= 99.5:
            diem_bsc = 5
        elif kq_thuc_hien_rounded >= 89.5:
            diem_bsc = 1 + 4 * ((kq_thuc_hien_rounded - 89.5) / 10)
        else:
            diem_bsc = 1

        # Làm tròn Điểm BSC về 2 chữ số thập phân
        diem_bsc_rounded = round(diem_bsc, 2)

        # Tạo hàng Tổng
        tong_row = {
            "Đơn vị": "Tổng",
            "SM1": tong_sm1,
            "SM2": tong_sm2,
            "KQ thực hiện chỉ tiêu": kq_thuc_hien_rounded,
            "Điểm BSC": diem_bsc_rounded
        }

        # Thêm hàng Tổng vào DataFrame
        df_tong = pd.DataFrame([tong_row])
        df_filtered = pd.concat([df_filtered, df_tong], ignore_index=True)

        print(f"✅ Đã thêm hàng Tổng:")
        print(f"   - SM1: {tong_sm1}")
        print(f"   - SM2: {tong_sm2}")
        print(f"   - KQ thực hiện: {kq_thuc_hien_rounded}")
        print(f"   - Điểm BSC: {diem_bsc_rounded}")

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_C1.5'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtered.to_excel(writer, sheet_name='TH_C1.5', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_C1.5'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.5")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.5: {e}")
        import traceback
        traceback.print_exc()
        return False


# def process_c15_chitiet_report():
#     """
#       hàm cũ không dùng cho báo cáo 2.Tất cả. hiện không còn dùng
#     # Xử lý báo cáo chi tiết C1.5:
#     # 1. Đọc file c1.5_chitiet_report.xlsx
#     # 2. Chuẩn hóa cột NVKT địa bàn:
#     #    - VNPT016770-Nguyễn Mạnh Hùng -> Nguyễn Mạnh Hùng
#     # 3. Tạo sheet mới 'KQ_C15_chitiet' với các cột:
#     #    - DOIVT: Tên đội
#     #    - NVKT: Tên NVKT sau khi chuẩn hóa
#     #    - Phiếu đạt: Số bản ghi có "Đạt chỉ tiêu" = "ĐẠT"
#     #    - Tổng Hoàn công: Tổng số bản ghi
#     #    - Tỉ lệ đạt: (Phiếu đạt / Tổng Hoàn công) * 100%
#     """
#     try:
#         print("\n" + "="*80)
#         print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.5")
#         print("="*80)

#         # Đường dẫn file
#         input_file = os.path.join("downloads", "baocao_hanoi", "c1.5_chitiet_report.xlsx")

#         if not os.path.exists(input_file):
#             print(f"❌ Không tìm thấy file: {input_file}")
#             return False

#         print(f"\n✓ Đang đọc file: {input_file}")

#         # Đọc file Excel
#         df = pd.read_excel(input_file)
#         print(f"✅ Đã đọc file, tổng số dòng: {len(df)}, tổng số cột: {df.shape[1]}")
#         print(f"Các cột hiện có: {', '.join(df.columns)}")

#         # Kiểm tra các cột cần thiết
#         required_columns = ['Đạt chỉ tiêu']
#         missing_columns = [col for col in required_columns if col not in df.columns]

#         if missing_columns:
#             print(f"❌ Không tìm thấy các cột bắt buộc: {', '.join(missing_columns)}")
#             return False

#         # Tìm cột NVKT địa bàn (có thể có tên khác nhau)
#         nvkt_column = None
#         for col in df.columns:
#             if 'NVKT' in col.upper() or 'địa bàn' in col.lower():
#                 nvkt_column = col
#                 break

#         if nvkt_column is None:
#             print(f"❌ Không tìm thấy cột NVKT địa bàn")
#             return False

#         print(f"\n✓ Sử dụng cột '{nvkt_column}' làm cột NVKT")

#         # Tìm cột DOIVT (có thể có tên khác nhau)
#         doivt_column = None
#         for col in df.columns:
#             if 'DOI' in col.upper() or 'đội' in col.lower():
#                 doivt_column = col
#                 break

#         if doivt_column is None:
#             print(f"⚠️ Cảnh báo: Không tìm thấy cột DOIVT, sẽ bỏ qua cột này")
#             has_doivt = False
#         else:
#             print(f"✓ Sử dụng cột '{doivt_column}' làm cột DOIVT")
#             has_doivt = True

#         # Chuẩn hóa cột NVKT
#         print("\n✓ Đang chuẩn hóa cột NVKT...")

#         def extract_nvkt_name(nvkt_text):
#             """
#             Chuẩn hóa tên NVKT:
#             VNPT016770-Nguyễn Mạnh Hùng -> Nguyễn Mạnh Hùng
#             """
#             if pd.isna(nvkt_text):
#                 return None

#             nvkt_text = str(nvkt_text).strip()

#             # Loại bỏ phần mã VNPT nếu có
#             if '-' in nvkt_text:
#                 parts = nvkt_text.split('-')
#                 # Lấy phần sau dấu "-" cuối cùng
#                 nvkt_name = parts[-1].strip()
#             else:
#                 nvkt_name = nvkt_text

#             return nvkt_name

#         # Áp dụng hàm chuẩn hóa
#         df['NVKT'] = df[nvkt_column].apply(extract_nvkt_name)

#         # Loại bỏ các dòng có NVKT là None/NaN
#         df_filtered = df[df['NVKT'].notna()].copy()
#         print(f"✅ Đã chuẩn hóa cột NVKT, còn lại {len(df_filtered)} dòng hợp lệ")

#         # Tạo báo cáo tổng hợp
#         print("\n✓ Đang tạo báo cáo tổng hợp...")

#         report_data = []

#         # Xác định các cột để group by
#         if has_doivt:
#             group_columns = [doivt_column, 'NVKT']
#             print("✓ Nhóm theo DOIVT và NVKT")
#         else:
#             group_columns = ['NVKT']
#             print("✓ Nhóm theo NVKT")

#         # Nhóm dữ liệu
#         for group_key, group_df in df_filtered.groupby(group_columns):
#             # Tính toán các chỉ số
#             tong_hoan_cong = len(group_df)
#             phieu_dat = len(group_df[group_df['Đạt chỉ tiêu'].str.upper() == 'ĐẠT'])
#             ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0

#             if has_doivt:
#                 doivt, nvkt = group_key
#                 report_data.append({
#                     'DOIVT': doivt,
#                     'NVKT': nvkt,
#                     'Phiếu đạt': phieu_dat,
#                     'Tổng Hoàn công': tong_hoan_cong,
#                     'Tỉ lệ đạt': f"{ti_le_dat:.2f}%"
#                 })
#             else:
#                 nvkt = group_key
#                 report_data.append({
#                     'NVKT': nvkt,
#                     'Phiếu đạt': phieu_dat,
#                     'Tổng Hoàn công': tong_hoan_cong,
#                     'Tỉ lệ đạt': f"{ti_le_dat:.2f}%"
#                 })

#         # Tạo DataFrame từ report_data
#         df_report = pd.DataFrame(report_data)
#         print(f"✅ Đã tạo báo cáo với {len(df_report)} dòng")

#         # Ghi vào sheet mới
#         print("\n✓ Đang ghi vào sheet mới 'KQ_C15_chitiet'...")

#         # Tạo sheet TH_TTVTST - Tổng hợp TTVT Sơn Tây
#         print("\n✓ Đang tạo sheet TH_TTVTST...")

#         # Tính toán các chỉ số tổng
#         tong_phieu_dat = len(df_filtered[df_filtered['Đạt chỉ tiêu'].str.upper() == 'ĐẠT'])
#         tong_phieu_hoan_cong = len(df_filtered)
#         ti_le_dat_phan_tram = (tong_phieu_dat / tong_phieu_hoan_cong * 100) if tong_phieu_hoan_cong > 0 else 0

#         # Tính điểm BSC theo công thức
#         def calculate_bsc(ti_le):
#             """
#             Tính điểm BSC theo công thức:
#             - KQ >= 99.5%: BSC = 5
#             - 89.5% <= KQ < 99.5%: BSC = 1 + 4 * ((KQ - 89.5) / 10)
#             - KQ < 89.5%: BSC = 1
#             """
#             if ti_le >= 99.5:
#                 return 5.0
#             elif ti_le >= 89.5:
#                 return 1 + 4 * ((ti_le - 89.5) / 10)
#             else:
#                 return 1.0

#         diem_bsc = calculate_bsc(ti_le_dat_phan_tram)

#         # Tạo DataFrame cho sheet TH_TTVTST
#         df_tongho = pd.DataFrame([{
#             'Tổng phiếu đạt': tong_phieu_dat,
#             'Tổng phiếu Hoàn công': tong_phieu_hoan_cong,
#             'Tỉ lệ đạt': f"{ti_le_dat_phan_tram:.2f}%",
#             'BSC': f"{diem_bsc:.2f}"
#         }])

#         print(f"✅ Đã tạo sheet TH_TTVTST:")
#         print(f"   - Tổng phiếu đạt: {tong_phieu_dat}")
#         print(f"   - Tổng phiếu Hoàn công: {tong_phieu_hoan_cong}")
#         print(f"   - Tỉ lệ đạt: {ti_le_dat_phan_tram:.2f}%")
#         print(f"   - Điểm BSC: {diem_bsc:.2f}")

#         # Mở file Excel và thêm cả 2 sheet mới
#         with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#             df_report.to_excel(writer, sheet_name='KQ_C15_chitiet', index=False)
#             df_tongho.to_excel(writer, sheet_name='TH_TTVTST', index=False)

#         print(f"\n✅ Đã ghi dữ liệu vào sheet 'KQ_C15_chitiet'")
#         print(f"✅ Đã ghi dữ liệu vào sheet 'TH_TTVTST'")

#         print("\n" + "="*80)
#         print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO CHI TIẾT C1.5")
#         print("="*80)

#         return True

#     except Exception as e:
#         print(f"\n❌ Lỗi khi xử lý báo cáo chi tiết C1.5: {e}")
#         import traceback
#         traceback.print_exc()
#         return False

def process_c15_chitiet_report():
    """
    Xử lý báo cáo chi tiết C1.5:
    1. Đọc file c1.5_chitiet_report.xlsx
    2. Tính thời gian thi công: TG_THI_CONG = (NGAY_HC - NGAY_LHD) * 24 (đơn vị: giờ)
    3. Đánh giá đạt/không đạt: TG_THI_CONG > 24 -> "Không đạt", ngược lại -> "Đạt"
    4. Chuẩn hóa cột TEN_NVKT để lấy tên NVKT:
       - VNPT016770-Nguyễn Mạnh Hùng -> Nguyễn Mạnh Hùng
    5. Tạo sheet mới 'KQ_C15_chitiet' với các cột:
       - DOIVT: Tên đội
       - NVKT: Tên NVKT sau khi chuẩn hóa
       - Phiếu đạt: Số bản ghi có TG_THI_CONG <= 24
       - Phiếu không đạt: Số bản ghi có TG_THI_CONG > 24
       - Tổng Hoàn công: Tổng số bản ghi
       - Tỉ lệ đạt: (Phiếu đạt / Tổng Hoàn công) * 100%
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.5")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "c1.5_chitiet_report.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}, tổng số cột: {df.shape[1]}")
        print(f"Các cột hiện có: {', '.join(df.columns)}")

        # Kiểm tra các cột cần thiết cho tính thời gian thi công
        required_columns = ['NGAY_HC', 'NGAY_LHD']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            print(f"❌ Không tìm thấy các cột bắt buộc: {', '.join(missing_columns)}")
            return False

        # Tìm cột TEN_NVKT
        nvkt_column = None
        for col in df.columns:
            if col == 'TEN_NVKT':
                nvkt_column = col
                break
            elif 'NVKT' in col.upper():
                nvkt_column = col
                break

        if nvkt_column is None:
            print(f"❌ Không tìm thấy cột TEN_NVKT")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False

        print(f"\n✓ Sử dụng cột '{nvkt_column}' làm cột NVKT")

        # Tìm cột DOIVT (có thể có tên khác nhau)
        doivt_column = None
        for col in df.columns:
            if col == 'DOIVT':
                doivt_column = col
                break
            elif 'DOI' in col.upper():
                doivt_column = col
                break

        if doivt_column is None:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột DOIVT, sẽ bỏ qua cột này")
            has_doivt = False
        else:
            print(f"✓ Sử dụng cột '{doivt_column}' làm cột DOIVT")
            has_doivt = True

        # Tính thời gian thi công
        print("\n✓ Đang tính thời gian thi công (TG_THI_CONG)...")

        # Chuyển đổi cột ngày sang datetime
        df['NGAY_HC'] = pd.to_datetime(df['NGAY_HC'], errors='coerce', dayfirst=True)
        df['NGAY_LHD'] = pd.to_datetime(df['NGAY_LHD'], errors='coerce', dayfirst=True)

        # Tính TG_THI_CONG = (NGAY_HC - NGAY_LHD) * 24 (đơn vị: giờ)
        # Chênh lệch ngày * 24 = số giờ
        df['TG_THI_CONG'] = (df['NGAY_HC'] - df['NGAY_LHD']).dt.total_seconds() / 3600

        # Đánh giá đạt/không đạt: TG_THI_CONG > 24 -> "Không đạt", ngược lại -> "Đạt"
        df['KET_QUA'] = df['TG_THI_CONG'].apply(lambda x: 'Không đạt' if pd.notna(x) and x > 24 else 'Đạt')

        # Thống kê ban đầu
        valid_records = df[df['TG_THI_CONG'].notna()]
        print(f"✅ Đã tính TG_THI_CONG cho {len(valid_records)} bản ghi hợp lệ")
        print(f"   - TG_THI_CONG trung bình: {valid_records['TG_THI_CONG'].mean():.2f} giờ")
        print(f"   - TG_THI_CONG min: {valid_records['TG_THI_CONG'].min():.2f} giờ")
        print(f"   - TG_THI_CONG max: {valid_records['TG_THI_CONG'].max():.2f} giờ")

        # Chuẩn hóa cột NVKT
        print("\n✓ Đang chuẩn hóa cột NVKT...")

        def extract_nvkt_name(nvkt_text):
            """
            Chuẩn hóa tên NVKT:
            VNPT016770-Nguyễn Mạnh Hùng -> Nguyễn Mạnh Hùng
            """
            if pd.isna(nvkt_text):
                return None

            nvkt_text = str(nvkt_text).strip()

            # Loại bỏ phần mã VNPT nếu có
            if '-' in nvkt_text:
                parts = nvkt_text.split('-')
                # Lấy phần sau dấu "-" cuối cùng
                nvkt_name = parts[-1].strip()
            else:
                nvkt_name = nvkt_text

            return nvkt_name

        # Áp dụng hàm chuẩn hóa
        df['NVKT'] = df[nvkt_column].apply(extract_nvkt_name)

        # Loại bỏ các dòng có NVKT là None/NaN và TG_THI_CONG hợp lệ
        df_filtered = df[(df['NVKT'].notna()) & (df['TG_THI_CONG'].notna())].copy()
        print(f"✅ Đã chuẩn hóa cột NVKT, còn lại {len(df_filtered)} dòng hợp lệ")

        # Tạo báo cáo tổng hợp
        print("\n✓ Đang tạo báo cáo tổng hợp...")

        report_data = []

        # Xác định các cột để group by
        if has_doivt:
            group_columns = [doivt_column, 'NVKT']
            print("✓ Nhóm theo DOIVT và NVKT")
        else:
            group_columns = ['NVKT']
            print("✓ Nhóm theo NVKT")

        # Nhóm dữ liệu
        for group_key, group_df in df_filtered.groupby(group_columns):
            # Tính toán các chỉ số
            tong_hoan_cong = len(group_df)
            phieu_dat = len(group_df[group_df['KET_QUA'] == 'Đạt'])
            phieu_khong_dat = len(group_df[group_df['KET_QUA'] == 'Không đạt'])
            ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0

            if has_doivt:
                doivt, nvkt = group_key
                report_data.append({
                    'DOIVT': doivt,
                    'NVKT': nvkt,
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })
            else:
                nvkt = group_key
                report_data.append({
                    'NVKT': nvkt,
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })

        # Tạo DataFrame từ report_data
        df_report = pd.DataFrame(report_data)
        print(f"✅ Đã tạo báo cáo với {len(df_report)} dòng")

        # Tạo sheet TH_TTVTST - Tổng hợp TTVT Sơn Tây (theo DOIVT và NVKT)
        print("\n✓ Đang tạo sheet TH_TTVTST...")

        # Tính điểm BSC theo công thức
        def calculate_bsc(ti_le):
            """
            Tính điểm BSC theo công thức:
            - KQ >= 99.5%: BSC = 5
            - 89.5% <= KQ < 99.5%: BSC = 1 + 4 * ((KQ - 89.5) / 10)
            - KQ < 89.5%: BSC = 1
            """
            if ti_le >= 99.5:
                return 5.0
            elif ti_le >= 89.5:
                return 1 + 4 * ((ti_le - 89.5) / 10)
            else:
                return 1.0

        # Tạo DataFrame cho sheet TH_TTVTST với cột DOIVT và NVKT
        th_data = []
        
        if has_doivt:
            # Thống kê theo DOIVT (theo đội)
            doivt_list = df_filtered[doivt_column].dropna().unique()
            doivt_list = sorted(doivt_list)
            
            for doivt in doivt_list:
                df_doi = df_filtered[df_filtered[doivt_column] == doivt]
                
                phieu_dat = len(df_doi[df_doi['KET_QUA'] == 'Đạt'])
                phieu_khong_dat = len(df_doi[df_doi['KET_QUA'] == 'Không đạt'])
                tong_hoan_cong = len(df_doi)
                ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0
                
                th_data.append({
                    'DOIVT': doivt,
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })
        else:
            # Không có cột DOIVT, tạo 1 dòng tổng hợp
            phieu_dat = len(df_filtered[df_filtered['KET_QUA'] == 'Đạt'])
            phieu_khong_dat = len(df_filtered[df_filtered['KET_QUA'] == 'Không đạt'])
            tong_hoan_cong = len(df_filtered)
            ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0
            
            th_data.append({
                'DOIVT': 'Tất cả',
                'Phiếu đạt': phieu_dat,
                'Phiếu không đạt': phieu_khong_dat,
                'Tổng Hoàn công': tong_hoan_cong,
                'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
            })
        
        # Thêm dòng Tổng (TTVT Sơn Tây)
        tong_phieu_dat = len(df_filtered[df_filtered['KET_QUA'] == 'Đạt'])
        tong_phieu_khong_dat = len(df_filtered[df_filtered['KET_QUA'] == 'Không đạt'])
        tong_phieu_hoan_cong = len(df_filtered)
        ti_le_dat_phan_tram = (tong_phieu_dat / tong_phieu_hoan_cong * 100) if tong_phieu_hoan_cong > 0 else 0
        diem_bsc = calculate_bsc(ti_le_dat_phan_tram)
        
        th_data.append({
            'DOIVT': 'TTVT Sơn Tây',
            'Phiếu đạt': tong_phieu_dat,
            'Phiếu không đạt': tong_phieu_khong_dat,
            'Tổng Hoàn công': tong_phieu_hoan_cong,
            'Tỉ lệ đạt (%)': round(ti_le_dat_phan_tram, 2)
        })
        
        # Tạo DataFrame
        df_tongho = pd.DataFrame(th_data)

        print(f"✅ Đã tạo sheet TH_TTVTST với {len(th_data) - 1} dòng DOIVT + 1 dòng Tổng:")
        print(f"   - Tổng phiếu đạt (TG <= 24h): {tong_phieu_dat}")
        print(f"   - Tổng phiếu không đạt (TG > 24h): {tong_phieu_khong_dat}")
        print(f"   - Tổng phiếu Hoàn công: {tong_phieu_hoan_cong}")
        print(f"   - Tỉ lệ đạt: {ti_le_dat_phan_tram:.2f}%")
        print(f"   - Điểm BSC: {diem_bsc:.2f}")

        # Tạo sheet chi tiết với cột TG_THI_CONG và KET_QUA
        print("\n✓ Đang tạo sheet chi tiết dữ liệu...")
        
        # Chọn các cột cần thiết cho sheet chi tiết
        detail_columns = ['MA_TB', 'NGAY_LHD', 'NGAY_HC', 'TG_THI_CONG', 'KET_QUA', 'NVKT']
        if has_doivt:
            detail_columns.insert(0, doivt_column)
        
        # Chỉ lấy các cột tồn tại
        available_detail_cols = [col for col in detail_columns if col in df_filtered.columns]
        df_detail = df_filtered[available_detail_cols].copy()
        
        # Làm tròn TG_THI_CONG
        df_detail['TG_THI_CONG'] = df_detail['TG_THI_CONG'].round(2)

        # Tạo sheet TH_KIEULD - Thống kê theo TEN_KIEULD, TEN_NVKT, DOIVT
        print("\n✓ Đang tạo sheet TH_KIEULD...")
        
        kieuld_data = []
        
        # Kiểm tra cột TEN_KIEULD có tồn tại không
        if 'TEN_KIEULD' in df_filtered.columns:
            # Xác định các cột để group by
            group_cols = []
            if has_doivt:
                group_cols.append(doivt_column)
            group_cols.extend(['NVKT', 'TEN_KIEULD'])
            
            # Nhóm dữ liệu
            for group_key, group_df in df_filtered.groupby(group_cols, dropna=False):
                if has_doivt:
                    doivt, nvkt, kieuld = group_key
                else:
                    nvkt, kieuld = group_key
                    doivt = ''
                
                # Tính toán các chỉ số
                phieu_dat = len(group_df[group_df['KET_QUA'] == 'Đạt'])
                phieu_khong_dat = len(group_df[group_df['KET_QUA'] == 'Không đạt'])
                tong_hoan_cong = len(group_df)
                ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0
                
                kieuld_data.append({
                    'DOIVT': doivt if doivt else '',
                    'NVKT': nvkt if nvkt else '',
                    'TEN_KIEULD': kieuld if kieuld else '',
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })
            
            df_kieuld = pd.DataFrame(kieuld_data)
            
            # Sắp xếp theo DOIVT, NVKT, TEN_KIEULD
            if has_doivt:
                df_kieuld = df_kieuld.sort_values(['DOIVT', 'NVKT', 'TEN_KIEULD']).reset_index(drop=True)
            else:
                df_kieuld = df_kieuld.sort_values(['NVKT', 'TEN_KIEULD']).reset_index(drop=True)
            
            print(f"✅ Đã tạo sheet TH_KIEULD với {len(df_kieuld)} dòng")
        else:
            print("⚠️ Không tìm thấy cột TEN_KIEULD, bỏ qua tạo sheet TH_KIEULD")
            df_kieuld = pd.DataFrame()

        # Tạo sheet TH_DVVT - Thống kê theo TEN_DVVT, TEN_NVKT, DOIVT
        print("\n✓ Đang tạo sheet TH_DVVT...")
        
        dvvt_data = []
        
        # Kiểm tra cột TEN_DVVT có tồn tại không
        if 'TEN_DVVT' in df_filtered.columns:
            # Xác định các cột để group by
            group_cols = []
            if has_doivt:
                group_cols.append(doivt_column)
            group_cols.extend(['NVKT', 'TEN_DVVT'])
            
            # Nhóm dữ liệu
            for group_key, group_df in df_filtered.groupby(group_cols, dropna=False):
                if has_doivt:
                    doivt, nvkt, dvvt = group_key
                else:
                    nvkt, dvvt = group_key
                    doivt = ''
                
                # Tính toán các chỉ số
                phieu_dat = len(group_df[group_df['KET_QUA'] == 'Đạt'])
                phieu_khong_dat = len(group_df[group_df['KET_QUA'] == 'Không đạt'])
                tong_hoan_cong = len(group_df)
                ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0
                
                dvvt_data.append({
                    'DOIVT': doivt if doivt else '',
                    'NVKT': nvkt if nvkt else '',
                    'TEN_DVVT': dvvt if dvvt else '',
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })
            
            df_dvvt = pd.DataFrame(dvvt_data)
            
            # Sắp xếp theo DOIVT, NVKT, TEN_DVVT
            if has_doivt:
                df_dvvt = df_dvvt.sort_values(['DOIVT', 'NVKT', 'TEN_DVVT']).reset_index(drop=True)
            else:
                df_dvvt = df_dvvt.sort_values(['NVKT', 'TEN_DVVT']).reset_index(drop=True)
            
            print(f"✅ Đã tạo sheet TH_DVVT với {len(df_dvvt)} dòng")
        else:
            print("⚠️ Không tìm thấy cột TEN_DVVT, bỏ qua tạo sheet TH_DVVT")
            df_dvvt = pd.DataFrame()

        # Tạo sheet TH_DVVT_DOI - Thống kê theo TEN_DVVT và DOIVT (không có NVKT)
        print("\n✓ Đang tạo sheet TH_DVVT_DOI...")
        
        dvvt_doi_data = []
        
        # Kiểm tra cột TEN_DVVT có tồn tại không
        if 'TEN_DVVT' in df_filtered.columns and has_doivt:
            # Nhóm dữ liệu theo DOIVT và TEN_DVVT
            for (doivt, dvvt), group_df in df_filtered.groupby([doivt_column, 'TEN_DVVT'], dropna=False):
                # Tính toán các chỉ số
                phieu_dat = len(group_df[group_df['KET_QUA'] == 'Đạt'])
                phieu_khong_dat = len(group_df[group_df['KET_QUA'] == 'Không đạt'])
                tong_hoan_cong = len(group_df)
                ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0
                
                dvvt_doi_data.append({
                    'DOIVT': doivt if doivt else '',
                    'TEN_DVVT': dvvt if dvvt else '',
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })
            
            df_dvvt_doi = pd.DataFrame(dvvt_doi_data)
            
            # Sắp xếp theo DOIVT, TEN_DVVT
            df_dvvt_doi = df_dvvt_doi.sort_values(['DOIVT', 'TEN_DVVT']).reset_index(drop=True)
            
            print(f"✅ Đã tạo sheet TH_DVVT_DOI với {len(df_dvvt_doi)} dòng")
        else:
            if 'TEN_DVVT' not in df_filtered.columns:
                print("⚠️ Không tìm thấy cột TEN_DVVT, bỏ qua tạo sheet TH_DVVT_DOI")
            else:
                print("⚠️ Không tìm thấy cột DOIVT, bỏ qua tạo sheet TH_DVVT_DOI")
            df_dvvt_doi = pd.DataFrame()

        # Tạo sheet TH_DVVT_TTVT - Thống kê theo TEN_DVVT cho toàn bộ TTVT Sơn Tây
        print("\n✓ Đang tạo sheet TH_DVVT_TTVT...")
        
        dvvt_ttvt_data = []
        
        # Kiểm tra cột TEN_DVVT có tồn tại không
        if 'TEN_DVVT' in df_filtered.columns:
            # Nhóm dữ liệu chỉ theo TEN_DVVT
            for dvvt, group_df in df_filtered.groupby('TEN_DVVT', dropna=False):
                # Tính toán các chỉ số
                phieu_dat = len(group_df[group_df['KET_QUA'] == 'Đạt'])
                phieu_khong_dat = len(group_df[group_df['KET_QUA'] == 'Không đạt'])
                tong_hoan_cong = len(group_df)
                ti_le_dat = (phieu_dat / tong_hoan_cong * 100) if tong_hoan_cong > 0 else 0
                
                dvvt_ttvt_data.append({
                    'TEN_DVVT': dvvt if dvvt else '',
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tổng Hoàn công': tong_hoan_cong,
                    'Tỉ lệ đạt (%)': round(ti_le_dat, 2)
                })
            
            # Thêm dòng Tổng (TTVT Sơn Tây)
            tong_dat = len(df_filtered[df_filtered['KET_QUA'] == 'Đạt'])
            tong_khong_dat = len(df_filtered[df_filtered['KET_QUA'] == 'Không đạt'])
            tong_hc = len(df_filtered)
            ti_le_tong = (tong_dat / tong_hc * 100) if tong_hc > 0 else 0
            
            dvvt_ttvt_data.append({
                'TEN_DVVT': 'TTVT Sơn Tây (Tổng)',
                'Phiếu đạt': tong_dat,
                'Phiếu không đạt': tong_khong_dat,
                'Tổng Hoàn công': tong_hc,
                'Tỉ lệ đạt (%)': round(ti_le_tong, 2)
            })
            
            df_dvvt_ttvt = pd.DataFrame(dvvt_ttvt_data)
            
            print(f"✅ Đã tạo sheet TH_DVVT_TTVT với {len(df_dvvt_ttvt) - 1} loại dịch vụ + 1 dòng Tổng")
        else:
            print("⚠️ Không tìm thấy cột TEN_DVVT, bỏ qua tạo sheet TH_DVVT_TTVT")
            df_dvvt_ttvt = pd.DataFrame()

        # Mở file Excel và thêm các sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_report.to_excel(writer, sheet_name='KQ_C15_chitiet', index=False)
            df_tongho.to_excel(writer, sheet_name='TH_TTVTST', index=False)
            df_detail.to_excel(writer, sheet_name='Chi_tiet_TG', index=False)
            if not df_kieuld.empty:
                df_kieuld.to_excel(writer, sheet_name='TH_KIEULD', index=False)
            if not df_dvvt.empty:
                df_dvvt.to_excel(writer, sheet_name='TH_DVVT', index=False)
            if not df_dvvt_doi.empty:
                df_dvvt_doi.to_excel(writer, sheet_name='TH_DVVT_DOI', index=False)
            if not df_dvvt_ttvt.empty:
                df_dvvt_ttvt.to_excel(writer, sheet_name='TH_DVVT_TTVT', index=False)

        print(f"\n✅ Đã ghi dữ liệu vào sheet 'KQ_C15_chitiet'")
        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_TTVTST'")
        print(f"✅ Đã ghi dữ liệu vào sheet 'Chi_tiet_TG'")
        if not df_kieuld.empty:
            print(f"✅ Đã ghi dữ liệu vào sheet 'TH_KIEULD'")
        if not df_dvvt.empty:
            print(f"✅ Đã ghi dữ liệu vào sheet 'TH_DVVT'")
        if not df_dvvt_doi.empty:
            print(f"✅ Đã ghi dữ liệu vào sheet 'TH_DVVT_DOI'")
        if not df_dvvt_ttvt.empty:
            print(f"✅ Đã ghi dữ liệu vào sheet 'TH_DVVT_TTVT'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO CHI TIẾT C1.5")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo chi tiết C1.5: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_c11_chitiet_report():
    """
    Xử lý báo cáo chi tiết C1.1:
    1. Đọc file SM4-C11.xlsx
    2. Chuẩn hóa cột TEN_KV để lấy tên NVKT:
       - Sơn Lộc 1 - Nguyễn Thành Sơn -> Nguyễn Thành Sơn
       - VNM3-Khuất Anh Chiến( VXN) -> Khuất Anh Chiến
    3. Tạo sheet mới 'chi_tiet' với các cột:
       - TEN_DOI: Tên đội (nếu có trong file)
       - NVKT: Tên NVKT sau khi chuẩn hóa
       - Tổng phiếu: Tổng số bản ghi theo TEN_DOI và NVKT
       - Số phiếu đạt: Số lượng bản ghi có DAT_TT_KO_HEN = 1
       - Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn: (Số phiếu đạt / Tổng phiếu) * 100%
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.1")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "SM4-C11.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Kiểm tra cột TEN_KV có tồn tại không
        if 'TEN_KV' not in df.columns:
            print(f"❌ Không tìm thấy cột 'TEN_KV' trong file")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False

        # Kiểm tra cột DAT_TT_KO_HEN có tồn tại không
        if 'DAT_TT_KO_HEN' not in df.columns:
            print(f"❌ Không tìm thấy cột 'DAT_TT_KO_HEN' trong file")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False

        # Kiểm tra cột TEN_DOI có tồn tại không
        if 'TEN_DOI' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'TEN_DOI' trong file")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            has_ten_doi = False
        else:
            has_ten_doi = True

        # Chuẩn hóa cột TEN_KV để lấy tên NVKT
        print("\n✓ Đang chuẩn hóa cột TEN_KV để lấy tên NVKT...")

        def extract_nvkt_name(ten_kv):
            """
            Trích xuất tên NVKT từ cột TEN_KV
            Ví dụ:
            - Sơn Lộc 1 - Nguyễn Thành Sơn -> Nguyễn Thành Sơn
            - VNM3-Khuất Anh Chiến( VXN) -> Khuất Anh Chiến
            """
            if pd.isna(ten_kv):
                return None

            ten_kv = str(ten_kv).strip()

            # Trường hợp có dấu "-" (ví dụ: Sơn Lộc 1 - Nguyễn Thành Sơn)
            if '-' in ten_kv:
                # Lấy phần sau dấu "-" cuối cùng
                parts = ten_kv.split('-')
                nvkt_name = parts[-1].strip()
            else:
                nvkt_name = ten_kv

            # Loại bỏ phần trong ngoặc đơn (ví dụ: ( VXN))
            if '(' in nvkt_name:
                nvkt_name = nvkt_name.split('(')[0].strip()

            return nvkt_name

        # Áp dụng hàm chuẩn hóa
        df['NVKT'] = df['TEN_KV'].apply(extract_nvkt_name)

        # Loại bỏ các dòng có NVKT là None/NaN
        df_filtered = df[df['NVKT'].notna()].copy()
        print(f"✅ Đã chuẩn hóa cột TEN_KV, còn lại {len(df_filtered)} dòng hợp lệ")

        # Tạo báo cáo tổng hợp theo TEN_DOI và NVKT
        if has_ten_doi:
            print("\n✓ Đang tạo báo cáo tổng hợp theo TEN_DOI và NVKT...")
        else:
            print("\n✓ Đang tạo báo cáo tổng hợp theo NVKT...")

        # Nhóm theo TEN_DOI và NVKT (hoặc chỉ NVKT nếu không có TEN_DOI)
        report_data = []

        if has_ten_doi:
            # Nhóm theo cả TEN_DOI và NVKT
            for (ten_doi, nvkt) in df_filtered.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                # Lọc dữ liệu theo TEN_DOI và NVKT
                df_group = df_filtered[(df_filtered['TEN_DOI'] == ten_doi) & (df_filtered['NVKT'] == nvkt)]

                # Tổng số phiếu
                tong_phieu = len(df_group)

                # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                so_phieu_dat = len(df_group[df_group['DAT_TT_KO_HEN'] == 1])

                # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                report_data.append({
                    'TEN_DOI': ten_doi,
                    'NVKT': nvkt,
                    'Tổng phiếu': tong_phieu,
                    'Số phiếu đạt': so_phieu_dat,
                    'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                })
        else:
            # Chỉ nhóm theo NVKT
            for nvkt in df_filtered['NVKT'].unique():
                # Lọc dữ liệu theo NVKT
                df_nvkt = df_filtered[df_filtered['NVKT'] == nvkt]

                # Tổng số phiếu
                tong_phieu = len(df_nvkt)

                # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                so_phieu_dat = len(df_nvkt[df_nvkt['DAT_TT_KO_HEN'] == 1])

                # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                report_data.append({
                    'NVKT': nvkt,
                    'Tổng phiếu': tong_phieu,
                    'Số phiếu đạt': so_phieu_dat,
                    'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                })

        # Tạo DataFrame từ dữ liệu tổng hợp
        df_report = pd.DataFrame(report_data)

        # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
        if has_ten_doi:
            df_report = df_report.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} nhóm TEN_DOI - NVKT")
        else:
            df_report = df_report.sort_values('NVKT').reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} NVKT")

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'chi_tiet'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_report.to_excel(writer, sheet_name='chi_tiet', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tiet'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN:")
        if has_ten_doi:
            print(f"  - Tổng số TEN_DOI: {df_report['TEN_DOI'].nunique()}")
            print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report)}")
        else:
            print(f"  - Tổng số NVKT: {len(df_report)}")
        print(f"  - Tổng số phiếu: {df_report['Tổng phiếu'].sum()}")
        print(f"  - Tổng số phiếu đạt: {df_report['Số phiếu đạt'].sum()}")
        total_ty_le = round((df_report['Số phiếu đạt'].sum() / df_report['Tổng phiếu'].sum() * 100), 2) if df_report['Tổng phiếu'].sum() > 0 else 0
        print(f"  - Tỷ lệ đạt chung: {total_ty_le}%")
        print("-"*80)

        # ===== TẠO SHEET chi_tieu_ko_hen_16h (Lọc theo thời gian NGAYGIAO từ 22h đến 16h) =====
        print("\n" + "="*80)
        print("TẠO SHEET chi_tieu_ko_hen_16h (Lọc thời gian NGAYGIAO 22h-16h)")
        print("="*80)

        # Kiểm tra cột NGAYGIAO có tồn tại không
        if 'NGAYGIAO' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'NGAYGIAO' trong file")
            print(f"Bỏ qua việc tạo sheet chi_tieu_ko_hen_16h")
        else:
            # Lọc các bản ghi có thời gian NGAYGIAO trong khoảng 22h đến 16h
            # (từ 22:00:00 đến 23:59:59 hoặc từ 00:00:00 đến 15:59:59)
            print("\n✓ Đang lọc các bản ghi có thời gian NGAYGIAO từ 22h đến 16h...")

            # Parse cột NGAYGIAO sang datetime nếu chưa
            df_16h = df_filtered.copy()

            # Kiểm tra xem NGAYGIAO đã là datetime chưa
            if not pd.api.types.is_datetime64_any_dtype(df_16h['NGAYGIAO']):
                # Parse datetime với định dạng DD/MM/YYYY HH:MM:SS
                df_16h['NGAYGIAO'] = pd.to_datetime(df_16h['NGAYGIAO'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

            # Lọc các bản ghi có NGAYGIAO hợp lệ
            df_16h = df_16h[df_16h['NGAYGIAO'].notna()].copy()

            # Trích xuất giờ từ cột NGAYGIAO
            df_16h['gio_giao'] = df_16h['NGAYGIAO'].dt.hour

            # Lọc các bản ghi có giờ trong khoảng [22, 23] hoặc [0, 15]
            df_16h = df_16h[(df_16h['gio_giao'] >= 22) | (df_16h['gio_giao'] <= 15)]

            print(f"✅ Đã lọc được {len(df_16h)} bản ghi (từ tổng số {len(df_filtered)} bản ghi hợp lệ)")

            # Tạo báo cáo tổng hợp cho df_16h tương tự như chi_tiet
            if has_ten_doi:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_16h theo TEN_DOI và NVKT...")
            else:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_16h theo NVKT...")

            report_data_16h = []

            if has_ten_doi:
                # Nhóm theo cả TEN_DOI và NVKT
                for (ten_doi, nvkt) in df_16h.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                    # Lọc dữ liệu theo TEN_DOI và NVKT
                    df_group = df_16h[(df_16h['TEN_DOI'] == ten_doi) & (df_16h['NVKT'] == nvkt)]

                    # Tổng số phiếu
                    tong_phieu = len(df_group)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_group[df_group['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_16h.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })
            else:
                # Chỉ nhóm theo NVKT
                for nvkt in df_16h['NVKT'].unique():
                    # Lọc dữ liệu theo NVKT
                    df_nvkt = df_16h[df_16h['NVKT'] == nvkt]

                    # Tổng số phiếu
                    tong_phieu = len(df_nvkt)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_nvkt[df_nvkt['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_16h.append({
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })

            # Tạo DataFrame từ dữ liệu tổng hợp
            df_report_16h = pd.DataFrame(report_data_16h)

            # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
            if has_ten_doi:
                df_report_16h = df_report_16h.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_16h cho {len(df_report_16h)} nhóm TEN_DOI - NVKT")
            else:
                df_report_16h = df_report_16h.sort_values('NVKT').reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_16h cho {len(df_report_16h)} NVKT")

            # Ghi vào sheet mới chi_tieu_ko_hen_16h
            print("\n✓ Đang ghi vào sheet mới 'chi_tieu_ko_hen_16h'...")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_report_16h.to_excel(writer, sheet_name='chi_tieu_ko_hen_16h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tieu_ko_hen_16h'")

            # In thống kê cho chi_tieu_ko_hen_16h
            print("\n" + "-"*80)
            print("THỐNG KÊ chi_tieu_ko_hen_16h (22h-16h):")
            if has_ten_doi:
                print(f"  - Tổng số TEN_DOI: {df_report_16h['TEN_DOI'].nunique()}")
                print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report_16h)}")
            else:
                print(f"  - Tổng số NVKT: {len(df_report_16h)}")
            print(f"  - Tổng số phiếu (22h-16h): {df_report_16h['Tổng phiếu'].sum()}")
            print(f"  - Tổng số phiếu đạt: {df_report_16h['Số phiếu đạt'].sum()}")
            total_ty_le_16h = round((df_report_16h['Số phiếu đạt'].sum() / df_report_16h['Tổng phiếu'].sum() * 100), 2) if df_report_16h['Tổng phiếu'].sum() > 0 else 0
            print(f"  - Tỷ lệ đạt chung: {total_ty_le_16h}%")
            print("-"*80)

            # Tạo sheet chi tiết phiếu không đạt cho 16h
            print("\n✓ Đang tạo sheet chi tiết phiếu không đạt 'chi_tiet_khong_dat_16h'...")
            df_16h_khong_dat = df_16h[df_16h['DAT_TT_KO_HEN'] != 1].copy()

            # Xóa cột gio_giao tạm thời (nếu có)
            if 'gio_giao' in df_16h_khong_dat.columns:
                df_16h_khong_dat = df_16h_khong_dat.drop(columns=['gio_giao'])

            print(f"  - Số phiếu không đạt trong khung giờ 22h-16h: {len(df_16h_khong_dat)}")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_16h_khong_dat.to_excel(writer, sheet_name='chi_tiet_khong_dat_16h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tiet_khong_dat_16h'")

        # ===== TẠO SHEET chi_tieu_ko_hen_17h (Lọc theo thời gian NGAYGIAO từ 22h đến 17h) =====
        print("\n" + "="*80)
        print("TẠO SHEET chi_tieu_ko_hen_17h (Lọc thời gian NGAYGIAO 22h-17h)")
        print("="*80)

        # Kiểm tra cột NGAYGIAO có tồn tại không
        if 'NGAYGIAO' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'NGAYGIAO' trong file")
            print(f"Bỏ qua việc tạo sheet chi_tieu_ko_hen_17h")
        else:
            # Lọc các bản ghi có thời gian NGAYGIAO trong khoảng 22h đến 17h
            # (từ 22:00:00 đến 23:59:59 hoặc từ 00:00:00 đến 16:59:59)
            print("\n✓ Đang lọc các bản ghi có thời gian NGAYGIAO từ 22h đến 17h...")

            # Parse cột NGAYGIAO sang datetime nếu chưa
            df_17h = df_filtered.copy()

            # Kiểm tra xem NGAYGIAO đã là datetime chưa
            if not pd.api.types.is_datetime64_any_dtype(df_17h['NGAYGIAO']):
                # Parse datetime với định dạng DD/MM/YYYY HH:MM:SS
                df_17h['NGAYGIAO'] = pd.to_datetime(df_17h['NGAYGIAO'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

            # Lọc các bản ghi có NGAYGIAO hợp lệ
            df_17h = df_17h[df_17h['NGAYGIAO'].notna()].copy()

            # Trích xuất giờ từ cột NGAYGIAO
            df_17h['gio_giao'] = df_17h['NGAYGIAO'].dt.hour

            # Lọc các bản ghi có giờ trong khoảng [22, 23] hoặc [0, 16]
            df_17h = df_17h[(df_17h['gio_giao'] >= 22) | (df_17h['gio_giao'] <= 16)]

            print(f"✅ Đã lọc được {len(df_17h)} bản ghi (từ tổng số {len(df_filtered)} bản ghi hợp lệ)")

            # Tạo báo cáo tổng hợp cho df_17h tương tự như chi_tiet
            if has_ten_doi:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_17h theo TEN_DOI và NVKT...")
            else:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_17h theo NVKT...")

            report_data_17h = []

            if has_ten_doi:
                # Nhóm theo cả TEN_DOI và NVKT
                for (ten_doi, nvkt) in df_17h.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                    # Lọc dữ liệu theo TEN_DOI và NVKT
                    df_group = df_17h[(df_17h['TEN_DOI'] == ten_doi) & (df_17h['NVKT'] == nvkt)]

                    # Tổng số phiếu
                    tong_phieu = len(df_group)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_group[df_group['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_17h.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })
            else:
                # Chỉ nhóm theo NVKT
                for nvkt in df_17h['NVKT'].unique():
                    # Lọc dữ liệu theo NVKT
                    df_nvkt = df_17h[df_17h['NVKT'] == nvkt]

                    # Tổng số phiếu
                    tong_phieu = len(df_nvkt)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_nvkt[df_nvkt['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_17h.append({
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })

            # Tạo DataFrame từ dữ liệu tổng hợp
            df_report_17h = pd.DataFrame(report_data_17h)

            # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
            if has_ten_doi:
                df_report_17h = df_report_17h.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_17h cho {len(df_report_17h)} nhóm TEN_DOI - NVKT")
            else:
                df_report_17h = df_report_17h.sort_values('NVKT').reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_17h cho {len(df_report_17h)} NVKT")

            # Ghi vào sheet mới chi_tieu_ko_hen_17h
            print("\n✓ Đang ghi vào sheet mới 'chi_tieu_ko_hen_17h'...")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_report_17h.to_excel(writer, sheet_name='chi_tieu_ko_hen_17h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tieu_ko_hen_17h'")

            # In thống kê cho chi_tieu_ko_hen_17h
            print("\n" + "-"*80)
            print("THỐNG KÊ chi_tieu_ko_hen_17h (22h-17h):")
            if has_ten_doi:
                print(f"  - Tổng số TEN_DOI: {df_report_17h['TEN_DOI'].nunique()}")
                print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report_17h)}")
            else:
                print(f"  - Tổng số NVKT: {len(df_report_17h)}")
            print(f"  - Tổng số phiếu (22h-17h): {df_report_17h['Tổng phiếu'].sum()}")
            print(f"  - Tổng số phiếu đạt: {df_report_17h['Số phiếu đạt'].sum()}")
            total_ty_le_17h = round((df_report_17h['Số phiếu đạt'].sum() / df_report_17h['Tổng phiếu'].sum() * 100), 2) if df_report_17h['Tổng phiếu'].sum() > 0 else 0
            print(f"  - Tỷ lệ đạt chung: {total_ty_le_17h}%")
            print("-"*80)

            # Tạo sheet chi tiết phiếu không đạt cho 17h
            print("\n✓ Đang tạo sheet chi tiết phiếu không đạt 'chi_tiet_khong_dat_17h'...")
            df_17h_khong_dat = df_17h[df_17h['DAT_TT_KO_HEN'] != 1].copy()

            # Xóa cột gio_giao tạm thời (nếu có)
            if 'gio_giao' in df_17h_khong_dat.columns:
                df_17h_khong_dat = df_17h_khong_dat.drop(columns=['gio_giao'])

            print(f"  - Số phiếu không đạt trong khung giờ 22h-17h: {len(df_17h_khong_dat)}")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_17h_khong_dat.to_excel(writer, sheet_name='chi_tiet_khong_dat_17h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tiet_khong_dat_17h'")

        # ===== TẠO SHEET chi_tieu_ko_hen_18h (Lọc theo thời gian NGAYGIAO từ 22h đến 18h) =====
        print("\n" + "="*80)
        print("TẠO SHEET chi_tieu_ko_hen_18h (Lọc thời gian NGAYGIAO 22h-18h)")
        print("="*80)

        # Kiểm tra cột NGAYGIAO có tồn tại không
        if 'NGAYGIAO' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'NGAYGIAO' trong file")
            print(f"Bỏ qua việc tạo sheet chi_tieu_ko_hen_18h")
        else:
            # Lọc các bản ghi có thời gian NGAYGIAO trong khoảng 22h đến 18h
            # (từ 22:00:00 đến 23:59:59 hoặc từ 00:00:00 đến 17:59:59)
            print("\n✓ Đang lọc các bản ghi có thời gian NGAYGIAO từ 22h đến 18h...")

            # Parse cột NGAYGIAO sang datetime nếu chưa
            df_18h = df_filtered.copy()

            # Kiểm tra xem NGAYGIAO đã là datetime chưa
            if not pd.api.types.is_datetime64_any_dtype(df_18h['NGAYGIAO']):
                # Parse datetime với định dạng DD/MM/YYYY HH:MM:SS
                df_18h['NGAYGIAO'] = pd.to_datetime(df_18h['NGAYGIAO'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

            # Lọc các bản ghi có NGAYGIAO hợp lệ
            df_18h = df_18h[df_18h['NGAYGIAO'].notna()].copy()

            # Trích xuất giờ từ cột NGAYGIAO
            df_18h['gio_giao'] = df_18h['NGAYGIAO'].dt.hour

            # Lọc các bản ghi có giờ trong khoảng [22, 23] hoặc [0, 17]
            df_18h = df_18h[(df_18h['gio_giao'] >= 22) | (df_18h['gio_giao'] <= 17)]

            print(f"✅ Đã lọc được {len(df_18h)} bản ghi (từ tổng số {len(df_filtered)} bản ghi hợp lệ)")

            # Tạo báo cáo tổng hợp cho df_18h tương tự như chi_tiet
            if has_ten_doi:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_18h theo TEN_DOI và NVKT...")
            else:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_18h theo NVKT...")

            report_data_18h = []

            if has_ten_doi:
                # Nhóm theo cả TEN_DOI và NVKT
                for (ten_doi, nvkt) in df_18h.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                    # Lọc dữ liệu theo TEN_DOI và NVKT
                    df_group = df_18h[(df_18h['TEN_DOI'] == ten_doi) & (df_18h['NVKT'] == nvkt)]

                    # Tổng số phiếu
                    tong_phieu = len(df_group)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_group[df_group['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_18h.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })
            else:
                # Chỉ nhóm theo NVKT
                for nvkt in df_18h['NVKT'].unique():
                    # Lọc dữ liệu theo NVKT
                    df_nvkt = df_18h[df_18h['NVKT'] == nvkt]

                    # Tổng số phiếu
                    tong_phieu = len(df_nvkt)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_nvkt[df_nvkt['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_18h.append({
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })

            # Tạo DataFrame từ dữ liệu tổng hợp
            df_report_18h = pd.DataFrame(report_data_18h)

            # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
            if has_ten_doi:
                df_report_18h = df_report_18h.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_18h cho {len(df_report_18h)} nhóm TEN_DOI - NVKT")
            else:
                df_report_18h = df_report_18h.sort_values('NVKT').reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_18h cho {len(df_report_18h)} NVKT")

            # Ghi vào sheet mới chi_tieu_ko_hen_18h
            print("\n✓ Đang ghi vào sheet mới 'chi_tieu_ko_hen_18h'...")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_report_18h.to_excel(writer, sheet_name='chi_tieu_ko_hen_18h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tieu_ko_hen_18h'")

            # In thống kê cho chi_tieu_ko_hen_18h
            print("\n" + "-"*80)
            print("THỐNG KÊ chi_tieu_ko_hen_18h (22h-18h):")
            if has_ten_doi:
                print(f"  - Tổng số TEN_DOI: {df_report_18h['TEN_DOI'].nunique()}")
                print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report_18h)}")
            else:
                print(f"  - Tổng số NVKT: {len(df_report_18h)}")
            print(f"  - Tổng số phiếu (22h-18h): {df_report_18h['Tổng phiếu'].sum()}")
            print(f"  - Tổng số phiếu đạt: {df_report_18h['Số phiếu đạt'].sum()}")
            total_ty_le_18h = round((df_report_18h['Số phiếu đạt'].sum() / df_report_18h['Tổng phiếu'].sum() * 100), 2) if df_report_18h['Tổng phiếu'].sum() > 0 else 0
            print(f"  - Tỷ lệ đạt chung: {total_ty_le_18h}%")
            print("-"*80)

            # Tạo sheet chi tiết phiếu không đạt cho 18h
            print("\n✓ Đang tạo sheet chi tiết phiếu không đạt 'chi_tiet_khong_dat_18h'...")
            df_18h_khong_dat = df_18h[df_18h['DAT_TT_KO_HEN'] != 1].copy()

            # Xóa cột gio_giao tạm thời (nếu có)
            if 'gio_giao' in df_18h_khong_dat.columns:
                df_18h_khong_dat = df_18h_khong_dat.drop(columns=['gio_giao'])

            print(f"  - Số phiếu không đạt trong khung giờ 22h-18h: {len(df_18h_khong_dat)}")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_18h_khong_dat.to_excel(writer, sheet_name='chi_tiet_khong_dat_18h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tiet_khong_dat_18h'")

        # ===== TẠO SHEET chi_tieu_ko_hen_15h (Lọc theo thời gian NGAYGIAO từ 22h đến 15h) =====
        print("\n" + "="*80)
        print("TẠO SHEET chi_tieu_ko_hen_15h (Lọc thời gian NGAYGIAO 22h-15h)")
        print("="*80)

        # Kiểm tra cột NGAYGIAO có tồn tại không
        if 'NGAYGIAO' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'NGAYGIAO' trong file")
            print(f"Bỏ qua việc tạo sheet chi_tieu_ko_hen_15h")
        else:
            # Lọc các bản ghi có thời gian NGAYGIAO trong khoảng 22h đến 15h
            # (từ 22:00:00 đến 23:59:59 hoặc từ 00:00:00 đến 14:59:59)
            print("\n✓ Đang lọc các bản ghi có thời gian NGAYGIAO từ 22h đến 15h...")

            # Parse cột NGAYGIAO sang datetime nếu chưa
            df_15h = df_filtered.copy()

            # Kiểm tra xem NGAYGIAO đã là datetime chưa
            if not pd.api.types.is_datetime64_any_dtype(df_15h['NGAYGIAO']):
                # Parse datetime với định dạng DD/MM/YYYY HH:MM:SS
                df_15h['NGAYGIAO'] = pd.to_datetime(df_15h['NGAYGIAO'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

            # Lọc các bản ghi có NGAYGIAO hợp lệ
            df_15h = df_15h[df_15h['NGAYGIAO'].notna()].copy()

            # Trích xuất giờ từ cột NGAYGIAO
            df_15h['gio_giao'] = df_15h['NGAYGIAO'].dt.hour

            # Lọc các bản ghi có giờ trong khoảng [22, 23] hoặc [0, 14]
            df_15h = df_15h[(df_15h['gio_giao'] >= 22) | (df_15h['gio_giao'] <= 14)]

            print(f"✅ Đã lọc được {len(df_15h)} bản ghi (từ tổng số {len(df_filtered)} bản ghi hợp lệ)")

            # Tạo báo cáo tổng hợp cho df_15h tương tự như chi_tiet
            if has_ten_doi:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_15h theo TEN_DOI và NVKT...")
            else:
                print("\n✓ Đang tạo báo cáo tổng hợp chi_tieu_ko_hen_15h theo NVKT...")

            report_data_15h = []

            if has_ten_doi:
                # Nhóm theo cả TEN_DOI và NVKT
                for (ten_doi, nvkt) in df_15h.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                    # Lọc dữ liệu theo TEN_DOI và NVKT
                    df_group = df_15h[(df_15h['TEN_DOI'] == ten_doi) & (df_15h['NVKT'] == nvkt)]

                    # Tổng số phiếu
                    tong_phieu = len(df_group)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_group[df_group['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_15h.append({
                        'TEN_DOI': ten_doi,
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })
            else:
                # Chỉ nhóm theo NVKT
                for nvkt in df_15h['NVKT'].unique():
                    # Lọc dữ liệu theo NVKT
                    df_nvkt = df_15h[df_15h['NVKT'] == nvkt]

                    # Tổng số phiếu
                    tong_phieu = len(df_nvkt)

                    # Số phiếu đạt (DAT_TT_KO_HEN = 1)
                    so_phieu_dat = len(df_nvkt[df_nvkt['DAT_TT_KO_HEN'] == 1])

                    # Tỷ lệ % (làm tròn 2 chữ số thập phân)
                    ty_le = round((so_phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                    report_data_15h.append({
                        'NVKT': nvkt,
                        'Tổng phiếu': tong_phieu,
                        'Số phiếu đạt': so_phieu_dat,
                        'Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn': ty_le
                    })

            # Tạo DataFrame từ dữ liệu tổng hợp
            df_report_15h = pd.DataFrame(report_data_15h)

            # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
            if has_ten_doi:
                df_report_15h = df_report_15h.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_15h cho {len(df_report_15h)} nhóm TEN_DOI - NVKT")
            else:
                df_report_15h = df_report_15h.sort_values('NVKT').reset_index(drop=True)
                print(f"✅ Đã tạo báo cáo chi_tieu_ko_hen_15h cho {len(df_report_15h)} NVKT")

            # Ghi vào sheet mới chi_tieu_ko_hen_15h
            print("\n✓ Đang ghi vào sheet mới 'chi_tieu_ko_hen_15h'...")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_report_15h.to_excel(writer, sheet_name='chi_tieu_ko_hen_15h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tieu_ko_hen_15h'")

            # In thống kê cho chi_tieu_ko_hen_15h
            print("\n" + "-"*80)
            print("THỐNG KÊ chi_tieu_ko_hen_15h (22h-15h):")
            if has_ten_doi:
                print(f"  - Tổng số TEN_DOI: {df_report_15h['TEN_DOI'].nunique()}")
                print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report_15h)}")
            else:
                print(f"  - Tổng số NVKT: {len(df_report_15h)}")
            print(f"  - Tổng số phiếu (22h-15h): {df_report_15h['Tổng phiếu'].sum()}")
            print(f"  - Tổng số phiếu đạt: {df_report_15h['Số phiếu đạt'].sum()}")
            total_ty_le_15h = round((df_report_15h['Số phiếu đạt'].sum() / df_report_15h['Tổng phiếu'].sum() * 100), 2) if df_report_15h['Tổng phiếu'].sum() > 0 else 0
            print(f"  - Tỷ lệ đạt chung: {total_ty_le_15h}%")
            print("-"*80)

            # Tạo sheet chi tiết phiếu không đạt cho 15h
            print("\n✓ Đang tạo sheet chi tiết phiếu không đạt 'chi_tiet_khong_dat_15h'...")
            df_15h_khong_dat = df_15h[df_15h['DAT_TT_KO_HEN'] != 1].copy()

            # Xóa cột gio_giao tạm thời (nếu có)
            if 'gio_giao' in df_15h_khong_dat.columns:
                df_15h_khong_dat = df_15h_khong_dat.drop(columns=['gio_giao'])

            print(f"  - Số phiếu không đạt trong khung giờ 22h-15h: {len(df_15h_khong_dat)}")

            with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_15h_khong_dat.to_excel(writer, sheet_name='chi_tiet_khong_dat_15h', index=False)

            print(f"✅ Đã ghi dữ liệu vào sheet 'chi_tiet_khong_dat_15h'")

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO CHI TIẾT C1.1")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo chi tiết C1.1: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c11_chitiet_report_SM2():
    """
    Xử lý báo cáo chi tiết C1.1 SM2:
    1. Đọc file SM2-C11.xlsx
    2. Tính cột TG = NGAY_NGHIEM_THU - NGAY_BAO_HONG (đổi ra giờ)
    3. Chuẩn hóa cột TEN_KV để lấy tên NVKT:
       - Sơn Tây 1 - Phạm Anh Tuấn -> Phạm Anh Tuấn
       - PCT1- Nguyễn Huy Tuyến(TML) -> Nguyễn Huy Tuyến
    4. Ghi lại cột TG vào sheet gốc để kiểm tra
    5. Tạo sheet mới 'TH_SM2' với các cột:
       - TEN_DOI: Tên đội (nếu có trong file)
       - NVKT: Tên NVKT sau khi chuẩn hóa
       - Tổng phiếu: Tổng số bản ghi theo TEN_DOI và NVKT
       - Phiếu đạt: Số lượng bản ghi có TG <= 72
       - Phiếu không đạt: Số lượng bản ghi có TG > 72
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.1 SM2")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "SM2-C11.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Kiểm tra các cột cần thiết
        required_columns = ['NGAY_NGHIEM_THU', 'NGAY_BAO_HONG', 'TEN_KV']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            print(f"❌ Không tìm thấy các cột: {', '.join(missing_columns)}")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False

        # Kiểm tra cột TEN_DOI có tồn tại không
        if 'TEN_DOI' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'TEN_DOI' trong file")
            has_ten_doi = False
        else:
            has_ten_doi = True

        # Tính cột TG (chênh lệch thời gian tính bằng giờ)
        print("\n✓ Đang tính cột TG (thời gian xử lý)...")

        # Parse datetime với định dạng DD/MM/YYYY HH:MM:SS
        # Ví dụ: "03/11/2025 11:19:47" = ngày 3 tháng 11 năm 2025
        df['NGAY_NGHIEM_THU'] = pd.to_datetime(df['NGAY_NGHIEM_THU'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        df['NGAY_BAO_HONG'] = pd.to_datetime(df['NGAY_BAO_HONG'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

        # Tính chênh lệch và chuyển sang giờ
        df['TG'] = (df['NGAY_NGHIEM_THU'] - df['NGAY_BAO_HONG']).dt.total_seconds() / 3600

        print(f"✅ Đã tính cột TG cho {len(df)} dòng")

        # Ghi lại cột TG vào sheet gốc để kiểm tra
        print("\n✓ Đang ghi cột TG vào sheet gốc...")
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # Ghi toàn bộ dataframe với cột TG mới
            df.to_excel(writer, index=False)
        print("✅ Đã ghi cột TG vào sheet gốc")

        # Chuẩn hóa cột TEN_KV để lấy tên NVKT (không loại bỏ dòng nào)
        print("\n✓ Đang chuẩn hóa cột TEN_KV để lấy tên NVKT...")

        def extract_nvkt_name(ten_kv):
            """
            Trích xuất tên NVKT từ cột TEN_KV
            Ví dụ:
            - Sơn Tây 1 - Phạm Anh Tuấn -> Phạm Anh Tuấn
            - PCT1- Nguyễn Huy Tuyến(TML) -> Nguyễn Huy Tuyến
            """
            if pd.isna(ten_kv):
                return None

            ten_kv = str(ten_kv).strip()

            # Trường hợp có dấu "-" (ví dụ: Sơn Tây 1 - Phạm Anh Tuấn)
            if '-' in ten_kv:
                # Lấy phần sau dấu "-" cuối cùng
                parts = ten_kv.split('-')
                nvkt_name = parts[-1].strip()
            else:
                nvkt_name = ten_kv

            # Loại bỏ phần trong ngoặc đơn (ví dụ: (TML))
            if '(' in nvkt_name:
                nvkt_name = nvkt_name.split('(')[0].strip()

            return nvkt_name

        # Áp dụng hàm chuẩn hóa (trên toàn bộ df, không loại bỏ dòng)
        df['NVKT'] = df['TEN_KV'].apply(extract_nvkt_name)

        print(f"✅ Đã chuẩn hóa cột TEN_KV cho {len(df)} dòng")

        # Tạo báo cáo tổng hợp theo TEN_DOI và NVKT
        if has_ten_doi:
            print("\n✓ Đang tạo báo cáo tổng hợp theo TEN_DOI và NVKT...")
        else:
            print("\n✓ Đang tạo báo cáo tổng hợp theo NVKT...")

        # Nhóm theo TEN_DOI và NVKT (hoặc chỉ NVKT nếu không có TEN_DOI)
        report_data = []

        if has_ten_doi:
            # Nhóm theo cả TEN_DOI và NVKT
            for (ten_doi, nvkt) in df.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                # Lọc dữ liệu theo TEN_DOI và NVKT
                df_group = df[(df['TEN_DOI'] == ten_doi) & (df['NVKT'] == nvkt)]

                # Tổng số phiếu
                tong_phieu = len(df_group)

                # Số phiếu đạt (TG <= 72 và TG không phải NaN)
                phieu_dat = len(df_group[(df_group['TG'].notna()) & (df_group['TG'] <= 72)])

                # Số phiếu không đạt (TG > 72)
                phieu_khong_dat = len(df_group[(df_group['TG'].notna()) & (df_group['TG'] > 72)])

                # Tỉ lệ đạt
                ty_le_dat = round((phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                report_data.append({
                    'TEN_DOI': ten_doi,
                    'NVKT': nvkt,
                    'Tổng phiếu': tong_phieu,
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tỉ lệ đạt (%)': ty_le_dat
                })
        else:
            # Chỉ nhóm theo NVKT
            for nvkt in df['NVKT'].unique():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo NVKT
                df_nvkt = df[df['NVKT'] == nvkt]

                # Tổng số phiếu
                tong_phieu = len(df_nvkt)

                # Số phiếu đạt (TG <= 72 và TG không phải NaN)
                phieu_dat = len(df_nvkt[(df_nvkt['TG'].notna()) & (df_nvkt['TG'] <= 72)])

                # Số phiếu không đạt (TG > 72)
                phieu_khong_dat = len(df_nvkt[(df_nvkt['TG'].notna()) & (df_nvkt['TG'] > 72)])

                # Tỉ lệ đạt
                ty_le_dat = round((phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

                report_data.append({
                    'NVKT': nvkt,
                    'Tổng phiếu': tong_phieu,
                    'Phiếu đạt': phieu_dat,
                    'Phiếu không đạt': phieu_khong_dat,
                    'Tỉ lệ đạt (%)': ty_le_dat
                })

        # Tạo DataFrame từ dữ liệu tổng hợp
        df_report = pd.DataFrame(report_data)

        # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
        if has_ten_doi:
            df_report = df_report.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} nhóm TEN_DOI - NVKT")
        else:
            df_report = df_report.sort_values('NVKT').reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} NVKT")

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_SM2'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_report.to_excel(writer, sheet_name='TH_SM2', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_SM2'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN:")
        if has_ten_doi:
            print(f"  - Tổng số TEN_DOI: {df_report['TEN_DOI'].nunique()}")
            print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report)}")
        else:
            print(f"  - Tổng số NVKT: {len(df_report)}")
        print(f"  - Tổng số phiếu: {df_report['Tổng phiếu'].sum()}")
        print(f"  - Tổng số phiếu đạt (TG <= 72h): {df_report['Phiếu đạt'].sum()}")
        print(f"  - Tổng số phiếu không đạt (TG > 72h): {df_report['Phiếu không đạt'].sum()}")
        ty_le_dat = round((df_report['Phiếu đạt'].sum() / df_report['Tổng phiếu'].sum() * 100), 2) if df_report['Tổng phiếu'].sum() > 0 else 0
        print(f"  - Tỷ lệ đạt: {ty_le_dat}%")
        print("-"*80)

        # # ===== TẠO SHEET TH_SM2_16h (Lọc theo thời gian báo hỏng từ 22h đến 16h) =====
        # print("\n" + "="*80)
        # print("TẠO SHEET TH_SM2_16h (Lọc thời gian báo hỏng 22h-16h)")
        # print("="*80)

        # # Lọc các bản ghi có thời gian báo hỏng trong khoảng 22h đến 16h
        # # (từ 22:00:00 đến 23:59:59 hoặc từ 00:00:00 đến 15:59:59)
        # print("\n✓ Đang lọc các bản ghi có thời gian báo hỏng từ 22h đến 16h...")

        # # Trích xuất giờ từ cột NGAY_BAO_HONG
        # df_16h = df[df['NGAY_BAO_HONG'].notna()].copy()
        # df_16h['gio_bao_hong'] = df_16h['NGAY_BAO_HONG'].dt.hour

        # # Lọc các bản ghi có giờ trong khoảng [22, 23] hoặc [0, 15]
        # df_16h = df_16h[(df_16h['gio_bao_hong'] >= 22) | (df_16h['gio_bao_hong'] <= 15)]

        # print(f"✅ Đã lọc được {len(df_16h)} bản ghi (từ tổng số {len(df)} bản ghi)")

        # # Tạo báo cáo tổng hợp cho df_16h tương tự như TH_SM2
        # if has_ten_doi:
        #     print("\n✓ Đang tạo báo cáo tổng hợp TH_SM2_16h theo TEN_DOI và NVKT...")
        # else:
        #     print("\n✓ Đang tạo báo cáo tổng hợp TH_SM2_16h theo NVKT...")

        # report_data_16h = []

        # if has_ten_doi:
        #     # Nhóm theo cả TEN_DOI và NVKT
        #     for (ten_doi, nvkt) in df_16h.groupby(['TEN_DOI', 'NVKT']).groups.keys():
        #         # Lọc dữ liệu theo TEN_DOI và NVKT
        #         df_group = df_16h[(df_16h['TEN_DOI'] == ten_doi) & (df_16h['NVKT'] == nvkt)]

        #         # Tổng số phiếu
        #         tong_phieu = len(df_group)

        #         # Số phiếu đạt (TG <= 72 và TG không phải NaN)
        #         phieu_dat = len(df_group[(df_group['TG'].notna()) & (df_group['TG'] <= 72)])

        #         # Số phiếu không đạt (TG > 72)
        #         phieu_khong_dat = len(df_group[(df_group['TG'].notna()) & (df_group['TG'] > 72)])

        #         # Tỉ lệ đạt
        #         ty_le_dat = round((phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

        #         report_data_16h.append({
        #             'TEN_DOI': ten_doi,
        #             'NVKT': nvkt,
        #             'Tổng phiếu': tong_phieu,
        #             'Phiếu đạt': phieu_dat,
        #             'Phiếu không đạt': phieu_khong_dat,
        #             'Tỉ lệ đạt (%)': ty_le_dat
        #         })
        # else:
        #     # Chỉ nhóm theo NVKT
        #     for nvkt in df_16h['NVKT'].unique():
        #         if pd.isna(nvkt):
        #             continue

        #         # Lọc dữ liệu theo NVKT
        #         df_nvkt = df_16h[df_16h['NVKT'] == nvkt]

        #         # Tổng số phiếu
        #         tong_phieu = len(df_nvkt)

        #         # Số phiếu đạt (TG <= 72 và TG không phải NaN)
        #         phieu_dat = len(df_nvkt[(df_nvkt['TG'].notna()) & (df_nvkt['TG'] <= 72)])

        #         # Số phiếu không đạt (TG > 72)
        #         phieu_khong_dat = len(df_nvkt[(df_nvkt['TG'].notna()) & (df_nvkt['TG'] > 72)])

        #         # Tỉ lệ đạt
        #         ty_le_dat = round((phieu_dat / tong_phieu * 100), 2) if tong_phieu > 0 else 0

        #         report_data_16h.append({
        #             'NVKT': nvkt,
        #             'Tổng phiếu': tong_phieu,
        #             'Phiếu đạt': phieu_dat,
        #             'Phiếu không đạt': phieu_khong_dat,
        #             'Tỉ lệ đạt (%)': ty_le_dat
        #         })

        # # Tạo DataFrame từ dữ liệu tổng hợp
        # df_report_16h = pd.DataFrame(report_data_16h)

        # # Sắp xếp theo TEN_DOI và NVKT (hoặc chỉ NVKT)
        # if has_ten_doi:
        #     df_report_16h = df_report_16h.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        #     print(f"✅ Đã tạo báo cáo TH_SM2_16h cho {len(df_report_16h)} nhóm TEN_DOI - NVKT")
        # else:
        #     df_report_16h = df_report_16h.sort_values('NVKT').reset_index(drop=True)
        #     print(f"✅ Đã tạo báo cáo TH_SM2_16h cho {len(df_report_16h)} NVKT")

        # # Ghi vào sheet mới TH_SM2_16h
        # print("\n✓ Đang ghi vào sheet mới 'TH_SM2_16h'...")

        # with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        #     df_report_16h.to_excel(writer, sheet_name='TH_SM2_16h', index=False)

        # print(f"✅ Đã ghi dữ liệu vào sheet 'TH_SM2_16h'")

        # # In thống kê cho TH_SM2_16h
        # print("\n" + "-"*80)
        # print("THỐNG KÊ TH_SM2_16h (22h-16h):")
        # if has_ten_doi:
        #     print(f"  - Tổng số TEN_DOI: {df_report_16h['TEN_DOI'].nunique()}")
        #     print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report_16h)}")
        # else:
        #     print(f"  - Tổng số NVKT: {len(df_report_16h)}")
        # print(f"  - Tổng số phiếu (22h-16h): {df_report_16h['Tổng phiếu'].sum()}")
        # print(f"  - Tổng số phiếu đạt (TG <= 72h): {df_report_16h['Phiếu đạt'].sum()}")
        # print(f"  - Tổng số phiếu không đạt (TG > 72h): {df_report_16h['Phiếu không đạt'].sum()}")
        # ty_le_dat_16h = round((df_report_16h['Phiếu đạt'].sum() / df_report_16h['Tổng phiếu'].sum() * 100), 2) if df_report_16h['Tổng phiếu'].sum() > 0 else 0
        # print(f"  - Tỷ lệ đạt: {ty_le_dat_16h}%")
        # print("-"*80)

        # print("\n" + "="*80)
        # print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO CHI TIẾT C1.1 SM2")
        # print("="*80)

        # return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo chi tiết C1.1 SM2: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c12_chitiet_report_SM1SM2():
    """
    Xử lý báo cáo chi tiết C1.2 SM1+SM2:
    Tính chi tiết "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại"

    Bước 1: Đọc file SM1-C12.xlsx
    - Tạo sheet 'TH_phieu_hong_lai_7_ngay' với các cột:
      - TEN_DOI: Tên đội
      - NVKT: Tên NVKT (chuẩn hóa từ TEN_KV)
      - Số phiếu HLL: (đếm số bản ghi từng NVKT) / 2

    Bước 2: Đọc file SM2-C12.xlsx
    - Tạo sheet 'Tong_phieu_bao_hong_thang' với các cột:
      - TEN_DOI: Tên đội
      - NVKT: Tên NVKT (chuẩn hóa từ TEN_KV)
      - Số phiếu báo hỏng: Tổng số bản ghi theo NVKT

    Bước 3: Tổng hợp và tạo sheet TH_SM1C12_HLL_Thang trong SM1-C12.xlsx
    - Merge dữ liệu từ cả 2 file SM1 và SM2
    - Tạo sheet 'TH_SM1C12_HLL_Thang' với các cột:
      - TEN_DOI: Tên đội
      - NVKT: Tên NVKT
      - Tỉ lệ HLL tháng (%): (Số phiếu HLL / Số phiếu báo hỏng) * 100
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.2 SM1+SM2 - BƯỚC 1")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "SM1-C12.xlsx")

        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        print(f"\n✓ Đang đọc file: {input_file}")

        # Đọc file Excel
        df = pd.read_excel(input_file)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df)}")

        # Kiểm tra cột TEN_KV
        if 'TEN_KV' not in df.columns:
            print(f"❌ Không tìm thấy cột 'TEN_KV' trong file")
            print(f"Các cột hiện có: {', '.join(df.columns)}")
            return False

        # Kiểm tra cột TEN_DOI
        if 'TEN_DOI' not in df.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'TEN_DOI' trong file")
            has_ten_doi = False
        else:
            has_ten_doi = True

        # Chuẩn hóa cột TEN_KV để lấy tên NVKT
        print("\n✓ Đang chuẩn hóa cột TEN_KV để lấy tên NVKT...")

        def extract_nvkt_name(ten_kv):
            """
            Trích xuất tên NVKT từ cột TEN_KV
            Ví dụ:
            - Sơn Lộc 1 - Nguyễn Thành Sơn -> Nguyễn Thành Sơn
            - Tây Đằng 03 - Bùi văn Cường -> Bùi văn Cường
            """
            if pd.isna(ten_kv):
                return None

            ten_kv = str(ten_kv).strip()

            # Trường hợp có dấu "-"
            if '-' in ten_kv:
                # Lấy phần sau dấu "-" cuối cùng
                parts = ten_kv.split('-')
                nvkt_name = parts[-1].strip()
            else:
                nvkt_name = ten_kv

            # Loại bỏ phần trong ngoặc đơn
            if '(' in nvkt_name:
                nvkt_name = nvkt_name.split('(')[0].strip()

            return nvkt_name

        # Áp dụng hàm chuẩn hóa
        df['NVKT'] = df['TEN_KV'].apply(extract_nvkt_name)

        print(f"✅ Đã chuẩn hóa cột TEN_KV cho {len(df)} dòng")

        # Tạo báo cáo tổng hợp theo TEN_DOI và NVKT
        if has_ten_doi:
            print("\n✓ Đang tạo báo cáo tổng hợp theo TEN_DOI và NVKT...")
        else:
            print("\n✓ Đang tạo báo cáo tổng hợp theo NVKT...")

        # Nhóm theo TEN_DOI và NVKT
        report_data = []

        if has_ten_doi:
            # Nhóm theo cả TEN_DOI và NVKT
            for (ten_doi, nvkt) in df.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo TEN_DOI và NVKT
                df_group = df[(df['TEN_DOI'] == ten_doi) & (df['NVKT'] == nvkt)]

                # Số bản ghi
                so_ban_ghi = len(df_group)

                # Số phiếu HLL = (số bản ghi) / 2, làm tròn lên
                so_phieu_hll = math.ceil(so_ban_ghi / 2)

                report_data.append({
                    'TEN_DOI': ten_doi,
                    'NVKT': nvkt,
                    'Số phiếu HLL': so_phieu_hll
                })
        else:
            # Chỉ nhóm theo NVKT
            for nvkt in df['NVKT'].unique():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo NVKT
                df_nvkt = df[df['NVKT'] == nvkt]

                # Số bản ghi
                so_ban_ghi = len(df_nvkt)

                # Số phiếu HLL = (số bản ghi) / 2, làm tròn lên
                so_phieu_hll = math.ceil(so_ban_ghi / 2)

                report_data.append({
                    'NVKT': nvkt,
                    'Số phiếu HLL': so_phieu_hll
                })

        # Tạo DataFrame từ dữ liệu tổng hợp
        df_report = pd.DataFrame(report_data)

        # Sắp xếp theo TEN_DOI và NVKT
        if has_ten_doi:
            df_report = df_report.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} nhóm TEN_DOI - NVKT")
        else:
            df_report = df_report.sort_values('NVKT').reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report)} NVKT")

        # Ghi vào sheet mới
        print("\n✓ Đang ghi vào sheet mới 'TH_phieu_hong_lai_7_ngay'...")

        # Mở file Excel và thêm sheet mới
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_report.to_excel(writer, sheet_name='TH_phieu_hong_lai_7_ngay', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_phieu_hong_lai_7_ngay'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN - BƯỚC 1:")
        if has_ten_doi:
            print(f"  - Tổng số TEN_DOI: {df_report['TEN_DOI'].nunique()}")
            print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report)}")
        else:
            print(f"  - Tổng số NVKT: {len(df_report)}")
        print(f"  - Tổng số bản ghi trong file gốc: {len(df)}")
        print(f"  - Tổng số phiếu HLL: {df_report['Số phiếu HLL'].sum()}")
        print("-"*80)

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH BƯỚC 1 - ĐÃ TẠO SHEET 'TH_phieu_hong_lai_7_ngay'")
        print("="*80)

        # ============================================================
        # BƯỚC 2: Đọc file SM2-C12.xlsx
        # ============================================================
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.2 SM1+SM2 - BƯỚC 2")
        print("="*80)

        # Đường dẫn file SM2-C12
        input_file_sm2 = os.path.join("downloads", "baocao_hanoi", "SM2-C12.xlsx")

        if not os.path.exists(input_file_sm2):
            print(f"❌ Không tìm thấy file: {input_file_sm2}")
            return False

        print(f"\n✓ Đang đọc file: {input_file_sm2}")

        # Đọc file Excel SM2-C12
        df_sm2 = pd.read_excel(input_file_sm2)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df_sm2)}")

        # Kiểm tra cột TEN_KV
        if 'TEN_KV' not in df_sm2.columns:
            print(f"❌ Không tìm thấy cột 'TEN_KV' trong file SM2-C12")
            return False

        # Kiểm tra cột TEN_DOI
        if 'TEN_DOI' not in df_sm2.columns:
            print(f"⚠️ Cảnh báo: Không tìm thấy cột 'TEN_DOI' trong file SM2-C12")
            has_ten_doi_sm2 = False
        else:
            has_ten_doi_sm2 = True

        # Chuẩn hóa cột TEN_KV để lấy tên NVKT
        print("\n✓ Đang chuẩn hóa cột TEN_KV của SM2-C12...")

        # Áp dụng hàm chuẩn hóa (dùng lại hàm extract_nvkt_name đã định nghĩa ở trên)
        df_sm2['NVKT'] = df_sm2['TEN_KV'].apply(extract_nvkt_name)

        print(f"✅ Đã chuẩn hóa cột TEN_KV cho {len(df_sm2)} dòng")

        # Tạo báo cáo tổng hợp theo TEN_DOI và NVKT
        if has_ten_doi_sm2:
            print("\n✓ Đang tạo báo cáo tổng hợp theo TEN_DOI và NVKT...")
        else:
            print("\n✓ Đang tạo báo cáo tổng hợp theo NVKT...")

        # Nhóm theo TEN_DOI và NVKT
        report_data_sm2 = []

        if has_ten_doi_sm2:
            # Nhóm theo cả TEN_DOI và NVKT
            for (ten_doi, nvkt) in df_sm2.groupby(['TEN_DOI', 'NVKT']).groups.keys():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo TEN_DOI và NVKT
                df_group = df_sm2[(df_sm2['TEN_DOI'] == ten_doi) & (df_sm2['NVKT'] == nvkt)]

                # Số phiếu báo hỏng = tổng số bản ghi
                so_phieu_bao_hong = len(df_group)

                report_data_sm2.append({
                    'TEN_DOI': ten_doi,
                    'NVKT': nvkt,
                    'Số phiếu báo hỏng': so_phieu_bao_hong
                })
        else:
            # Chỉ nhóm theo NVKT
            for nvkt in df_sm2['NVKT'].unique():
                if pd.isna(nvkt):
                    continue

                # Lọc dữ liệu theo NVKT
                df_nvkt = df_sm2[df_sm2['NVKT'] == nvkt]

                # Số phiếu báo hỏng = tổng số bản ghi
                so_phieu_bao_hong = len(df_nvkt)

                report_data_sm2.append({
                    'NVKT': nvkt,
                    'Số phiếu báo hỏng': so_phieu_bao_hong
                })

        # Tạo DataFrame từ dữ liệu tổng hợp
        df_report_sm2 = pd.DataFrame(report_data_sm2)

        # Sắp xếp theo TEN_DOI và NVKT
        if has_ten_doi_sm2:
            df_report_sm2 = df_report_sm2.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report_sm2)} nhóm TEN_DOI - NVKT")
        else:
            df_report_sm2 = df_report_sm2.sort_values('NVKT').reset_index(drop=True)
            print(f"✅ Đã tạo báo cáo cho {len(df_report_sm2)} NVKT")

        # Ghi vào sheet mới trong file SM2-C12
        print("\n✓ Đang ghi vào sheet mới 'Tong_phieu_bao_hong_thang'...")

        with pd.ExcelWriter(input_file_sm2, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_report_sm2.to_excel(writer, sheet_name='Tong_phieu_bao_hong_thang', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'Tong_phieu_bao_hong_thang'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN - BƯỚC 2:")
        if has_ten_doi_sm2:
            print(f"  - Tổng số TEN_DOI: {df_report_sm2['TEN_DOI'].nunique()}")
            print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_report_sm2)}")
        else:
            print(f"  - Tổng số NVKT: {len(df_report_sm2)}")
        print(f"  - Tổng số bản ghi trong file SM2-C12: {len(df_sm2)}")
        print(f"  - Tổng số phiếu báo hỏng: {df_report_sm2['Số phiếu báo hỏng'].sum()}")
        print("-"*80)

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH BƯỚC 2 - ĐÃ TẠO SHEET 'Tong_phieu_bao_hong_thang'")
        print("="*80)

        # ============================================================
        # BƯỚC 3: Tạo sheet TH_SM1C12_HLL_Thang trong SM1-C12
        # ============================================================
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO CHI TIẾT C1.2 SM1+SM2 - BƯỚC 3")
        print("="*80)

        print("\n✓ Đang tổng hợp dữ liệu từ cả 2 file SM1 và SM2...")

        # Merge dữ liệu từ df_report (SM1) và df_report_sm2 (SM2)
        # Sử dụng merge để kết hợp dữ liệu theo TEN_DOI và NVKT
        if has_ten_doi and has_ten_doi_sm2:
            # Merge theo cả TEN_DOI và NVKT
            df_merged = pd.merge(
                df_report[['TEN_DOI', 'NVKT', 'Số phiếu HLL']],
                df_report_sm2[['TEN_DOI', 'NVKT', 'Số phiếu báo hỏng']],
                on=['TEN_DOI', 'NVKT'],
                how='outer'  # Sử dụng outer join để giữ tất cả NVKT từ cả 2 file
            )
        elif has_ten_doi:
            # SM1 có TEN_DOI nhưng SM2 không có
            # Merge theo NVKT, lấy TEN_DOI từ SM1
            df_merged = pd.merge(
                df_report[['TEN_DOI', 'NVKT', 'Số phiếu HLL']],
                df_report_sm2[['NVKT', 'Số phiếu báo hỏng']],
                on='NVKT',
                how='outer'
            )
        elif has_ten_doi_sm2:
            # SM2 có TEN_DOI nhưng SM1 không có
            # Merge theo NVKT, lấy TEN_DOI từ SM2
            df_merged = pd.merge(
                df_report[['NVKT', 'Số phiếu HLL']],
                df_report_sm2[['TEN_DOI', 'NVKT', 'Số phiếu báo hỏng']],
                on='NVKT',
                how='outer'
            )
        else:
            # Cả 2 đều không có TEN_DOI
            df_merged = pd.merge(
                df_report[['NVKT', 'Số phiếu HLL']],
                df_report_sm2[['NVKT', 'Số phiếu báo hỏng']],
                on='NVKT',
                how='outer'
            )

        # Điền 0 cho các giá trị NaN (trường hợp NVKT chỉ có trong 1 trong 2 file)
        df_merged['Số phiếu HLL'] = df_merged['Số phiếu HLL'].fillna(0)
        df_merged['Số phiếu báo hỏng'] = df_merged['Số phiếu báo hỏng'].fillna(0)

        print(f"✅ Đã merge dữ liệu, tổng số NVKT: {len(df_merged)}")

        # Tính Tỉ lệ HLL tháng = (Số phiếu HLL / Số phiếu báo hỏng) * 100
        print("\n✓ Đang tính Tỉ lệ HLL tháng...")

        def calculate_ty_le_hll(row):
            """Tính tỉ lệ HLL tháng ở dạng %"""
            so_phieu_hll = row['Số phiếu HLL']
            so_phieu_bao_hong = row['Số phiếu báo hỏng']

            # Tránh chia cho 0
            if so_phieu_bao_hong == 0:
                return 0

            # Tính tỉ lệ và làm tròn 2 chữ số thập phân
            ty_le = round((so_phieu_hll / so_phieu_bao_hong) * 100, 2)
            return ty_le

        df_merged['Tỉ lệ HLL tháng (2.5%)'] = df_merged.apply(calculate_ty_le_hll, axis=1)

        print(f"✅ Đã tính Tỉ lệ HLL tháng cho {len(df_merged)} NVKT")

        # Tạo DataFrame kết quả với cấu trúc: TEN_DOI, NVKT, Số phiếu HLL, Số phiếu báo hỏng, Tỉ lệ HLL tháng (2.5%)
        if 'TEN_DOI' in df_merged.columns:
            df_result = df_merged[['TEN_DOI', 'NVKT', 'Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỉ lệ HLL tháng (2.5%)']].copy()
            # Sắp xếp theo TEN_DOI và NVKT
            df_result = df_result.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
        else:
            df_result = df_merged[['NVKT', 'Số phiếu HLL', 'Số phiếu báo hỏng', 'Tỉ lệ HLL tháng (2.5%)']].copy()
            # Sắp xếp theo NVKT
            df_result = df_result.sort_values('NVKT').reset_index(drop=True)

        # Ghi vào sheet mới trong file SM1-C12
        print("\n✓ Đang ghi vào sheet mới 'TH_SM1C12_HLL_Thang' trong file SM1-C12...")

        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_result.to_excel(writer, sheet_name='TH_SM1C12_HLL_Thang', index=False)

        print(f"✅ Đã ghi dữ liệu vào sheet 'TH_SM1C12_HLL_Thang'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN - BƯỚC 3:")
        if 'TEN_DOI' in df_result.columns:
            print(f"  - Tổng số TEN_DOI: {df_result['TEN_DOI'].nunique()}")
            print(f"  - Tổng số nhóm TEN_DOI - NVKT: {len(df_result)}")
        else:
            print(f"  - Tổng số NVKT: {len(df_result)}")
        print(f"  - Tỉ lệ HLL tháng trung bình: {df_result['Tỉ lệ HLL tháng (2.5%)'].mean():.2f}%")
        print(f"  - Tỉ lệ HLL tháng cao nhất: {df_result['Tỉ lệ HLL tháng (2.5%)'].max():.2f}%")
        print(f"  - Tỉ lệ HLL tháng thấp nhất: {df_result['Tỉ lệ HLL tháng (2.5%)'].min():.2f}%")
        print("-"*80)

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH BƯỚC 3 - ĐÃ TẠO SHEET 'TH_SM1C12_HLL_Thang'")
        print("="*80)

        print("\n" + "="*80)
        print("🎉 HOÀN THÀNH TẤT CẢ 3 BƯỚC XỬ LÝ BÁO CÁO C1.2 SM1+SM2")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo chi tiết C1.2 SM1+SM2: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_c12_ti_le_bao_hong():
    """
    Xử lý báo cáo C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng:
    
    Cách tính:
    - Đọc file SM4-C11.xlsx Sheet 1, đếm số phiếu báo hỏng theo từng NVKT
      cho các dịch vụ: FiberVNN (Internet trên cáp quang), MyTV, ĐTCĐ, ĐTCĐ IMS
    - Chia cho tổng số thuê bao từng NVKT quản lý từ file Tonghop_thuebao_NVKT_DB_C12.xlsx cột "Tổng TB"
    - Lưu kết quả ra file SM4-C12-ti-le-su-co-dv-brcd.xlsx
    """
    try:
        print("\n" + "="*80)
        print("BẮT ĐẦU XỬ LÝ BÁO CÁO C1.2 - TỶ LỆ THUÊ BAO BRCĐ BÁO HỎNG")
        print("="*80)

        # Đường dẫn file
        input_file = os.path.join("downloads", "baocao_hanoi", "SM4-C11.xlsx")
        ref_file = os.path.join("du_lieu_tham_chieu", "Tonghop_thuebao_NVKT_DB_C12.xlsx")
        output_file = os.path.join("downloads", "baocao_hanoi", "SM4-C12-ti-le-su-co-dv-brcd.xlsx")

        # Kiểm tra file đầu vào
        if not os.path.exists(input_file):
            print(f"❌ Không tìm thấy file: {input_file}")
            return False

        if not os.path.exists(ref_file):
            print(f"❌ Không tìm thấy file tham chiếu: {ref_file}")
            return False

        # Đọc file SM4-C11.xlsx Sheet 1
        print(f"\n✓ Đang đọc file: {input_file}")
        df_sm4 = pd.read_excel(input_file, sheet_name=0)  # Sheet 1 (index 0)
        print(f"✅ Đã đọc file, tổng số dòng: {len(df_sm4)}")

        # Kiểm tra cột TEN_DICH_VU và TEN_KV
        if 'TEN_DICH_VU' not in df_sm4.columns:
            print(f"❌ Không tìm thấy cột 'TEN_DICH_VU' trong file")
            return False

        if 'TEN_KV' not in df_sm4.columns:
            print(f"❌ Không tìm thấy cột 'TEN_KV' trong file")
            return False

        # Danh sách các dịch vụ BRCĐ cần lọc
        # Băng rộng cố định (FiberVNN), Cố định (ĐTCĐ), IMS (ĐTCĐ IMS)
        dich_vu_brcd = ['Băng rộng cố định', 'Cố định', 'IMS']
        
        print(f"\n✓ Lọc các dịch vụ BRCĐ: {dich_vu_brcd}")
        
        # Lọc dữ liệu theo các dịch vụ BRCĐ
        df_brcd = df_sm4[df_sm4['TEN_DICH_VU'].isin(dich_vu_brcd)].copy()
        print(f"✅ Đã lọc được {len(df_brcd)} bản ghi dịch vụ BRCĐ")

        # Hàm chuẩn hóa tên NVKT
        def extract_nvkt_name(ten_kv):
            """
            Trích xuất tên NVKT từ cột TEN_KV
            Ví dụ:
            - Sơn Lộc 1 - Nguyễn Thành Sơn -> Nguyễn Thành Sơn
            - TMH4-Bùi Văn Duẩn(PGT) -> Bùi Văn Duẩn
            """
            if pd.isna(ten_kv):
                return None

            ten_kv = str(ten_kv).strip()

            # Trường hợp có dấu "-"
            if '-' in ten_kv:
                parts = ten_kv.split('-')
                nvkt_name = parts[-1].strip()
            else:
                nvkt_name = ten_kv

            # Loại bỏ phần trong ngoặc đơn
            if '(' in nvkt_name:
                nvkt_name = nvkt_name.split('(')[0].strip()

            return nvkt_name

        # Chuẩn hóa tên NVKT
        print("\n✓ Đang chuẩn hóa tên NVKT...")
        df_brcd['NVKT'] = df_brcd['TEN_KV'].apply(extract_nvkt_name)
        # Chuẩn hóa về Title Case để tránh trùng lặp (VD: "Bùi văn Cường" -> "Bùi Văn Cường")
        df_brcd['NVKT'] = df_brcd['NVKT'].str.strip().str.title()
        print(f"✅ Đã chuẩn hóa tên NVKT")

        # Kiểm tra cột TEN_DOI
        has_ten_doi = 'TEN_DOI' in df_brcd.columns

        # Đếm số phiếu báo hỏng theo NVKT (và TEN_DOI nếu có)
        print("\n✓ Đang đếm số phiếu báo hỏng theo NVKT...")
        
        if has_ten_doi:
            df_count = df_brcd.groupby(['TEN_DOI', 'NVKT']).size().reset_index(name='Số phiếu báo hỏng')
        else:
            df_count = df_brcd.groupby('NVKT').size().reset_index(name='Số phiếu báo hỏng')
        
        print(f"✅ Đã đếm được {len(df_count)} NVKT có phiếu báo hỏng")

        # Đọc file tham chiếu thuê bao
        print(f"\n✓ Đang đọc file tham chiếu: {ref_file}")
        df_ref = pd.read_excel(ref_file)
        print(f"✅ Đã đọc file tham chiếu, tổng số dòng: {len(df_ref)}")

        # Kiểm tra các cột cần thiết trong file tham chiếu
        # Cột tên NVKT có thể là 'TÊN NVKT ĐB' hoặc tương tự
        nvkt_col = None
        for col in df_ref.columns:
            if 'NVKT' in col.upper() or 'TÊN' in col.upper():
                nvkt_col = col
                break
        
        if nvkt_col is None:
            print(f"❌ Không tìm thấy cột tên NVKT trong file tham chiếu")
            print(f"Các cột có: {df_ref.columns.tolist()}")
            return False

        if 'Tổng TB' not in df_ref.columns:
            print(f"❌ Không tìm thấy cột 'Tổng TB' trong file tham chiếu")
            return False

        print(f"✓ Sử dụng cột '{nvkt_col}' làm tên NVKT và cột 'Tổng TB' làm tổng thuê bao")

        # Đổi tên cột để merge - lấy thêm cột ĐỘI VT để điền TEN_DOI
        doi_vt_col = 'ĐỘI VT' if 'ĐỘI VT' in df_ref.columns else None
        
        if doi_vt_col:
            df_ref_clean = df_ref[[nvkt_col, 'Tổng TB', doi_vt_col]].copy()
            df_ref_clean.columns = ['NVKT_raw', 'Tổng TB', 'TEN_DOI_ref']
        else:
            df_ref_clean = df_ref[[nvkt_col, 'Tổng TB']].copy()
            df_ref_clean.columns = ['NVKT_raw', 'Tổng TB']
        
        # Chuẩn hóa tên NVKT từ file tham chiếu (loại bỏ mã CTV, VNPT)
        # VD: CTV030837-Khuất Anh Chiến -> Khuất Anh Chiến
        #     VNPT016776-Bùi Văn Duẩn -> Bùi Văn Duẩn
        df_ref_clean['NVKT'] = df_ref_clean['NVKT_raw'].apply(extract_nvkt_name)
        # Chuẩn hóa về Title Case để khớp với dữ liệu phiếu báo hỏng
        df_ref_clean['NVKT'] = df_ref_clean['NVKT'].str.strip().str.title()
        
        if doi_vt_col:
            df_ref_clean = df_ref_clean[['NVKT', 'Tổng TB', 'TEN_DOI_ref']]
        else:
            df_ref_clean = df_ref_clean[['NVKT', 'Tổng TB']]
        
        # Loại bỏ các dòng có giá trị NaN
        df_ref_clean = df_ref_clean.dropna(subset=['NVKT', 'Tổng TB'])

        # Merge dữ liệu
        print("\n✓ Đang merge dữ liệu phiếu báo hỏng với thuê bao...")
        
        if has_ten_doi:
            df_result = pd.merge(
                df_count,
                df_ref_clean,
                on='NVKT',
                how='outer'
            )
            # Điền TEN_DOI từ file tham chiếu nếu thiếu
            if doi_vt_col and 'TEN_DOI_ref' in df_result.columns:
                df_result['TEN_DOI'] = df_result['TEN_DOI'].fillna(df_result['TEN_DOI_ref'])
                df_result = df_result.drop(columns=['TEN_DOI_ref'])
        else:
            df_result = pd.merge(
                df_count,
                df_ref_clean,
                on='NVKT',
                how='outer'
            )
            # Thêm cột TEN_DOI từ file tham chiếu
            if doi_vt_col and 'TEN_DOI_ref' in df_result.columns:
                df_result['TEN_DOI'] = df_result['TEN_DOI_ref']
                df_result = df_result.drop(columns=['TEN_DOI_ref'])
        
        # Điền 0 cho các giá trị NaN ở cột số phiếu báo hỏng
        df_result['Số phiếu báo hỏng'] = df_result['Số phiếu báo hỏng'].fillna(0).astype(int)
        df_result['Tổng TB'] = df_result['Tổng TB'].fillna(0).astype(int)

        print(f"✅ Đã merge được {len(df_result)} NVKT")

        # Tính tỷ lệ báo hỏng
        print("\n✓ Đang tính tỷ lệ báo hỏng...")
        
        def calculate_ty_le_bao_hong(row):
            """Tính tỷ lệ báo hỏng = (Số phiếu báo hỏng / Tổng TB) * 100"""
            so_phieu = row['Số phiếu báo hỏng']
            tong_tb = row['Tổng TB']
            
            if tong_tb == 0:
                return 0
            
            ty_le = round((so_phieu / tong_tb) * 100, 2)
            return ty_le

        df_result['Tỷ lệ báo hỏng (%)'] = df_result.apply(calculate_ty_le_bao_hong, axis=1)

        print(f"✅ Đã tính tỷ lệ báo hỏng cho {len(df_result)} NVKT")

        # Sắp xếp kết quả
        if 'TEN_DOI' in df_result.columns:
            df_result = df_result.sort_values(['TEN_DOI', 'NVKT']).reset_index(drop=True)
            # Sắp xếp lại thứ tự cột
            df_result = df_result[['TEN_DOI', 'NVKT', 'Số phiếu báo hỏng', 'Tổng TB', 'Tỷ lệ báo hỏng (%)']]
        else:
            df_result = df_result.sort_values('NVKT').reset_index(drop=True)
            df_result = df_result[['NVKT', 'Số phiếu báo hỏng', 'Tổng TB', 'Tỷ lệ báo hỏng (%)']]

        # Loại bỏ các dòng có NVKT rỗng
        df_result = df_result[df_result['NVKT'].notna()]

        # Ghi ra file Excel
        print(f"\n✓ Đang ghi kết quả ra file: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='TH_C12_TiLeBaoHong', index=False)

        print(f"✅ Đã ghi dữ liệu vào file '{output_file}'")

        # In thống kê tổng quan
        print("\n" + "-"*80)
        print("THỐNG KÊ TỔNG QUAN:")
        print(f"  - Tổng số NVKT: {len(df_result)}")
        print(f"  - Tổng số phiếu báo hỏng BRCĐ: {df_result['Số phiếu báo hỏng'].sum()}")
        print(f"  - Tổng số thuê bao: {df_result['Tổng TB'].sum()}")
        ty_le_tb = round((df_result['Số phiếu báo hỏng'].sum() / df_result['Tổng TB'].sum() * 100), 2) if df_result['Tổng TB'].sum() > 0 else 0
        print(f"  - Tỷ lệ báo hỏng trung bình: {ty_le_tb}%")
        print(f"  - Tỷ lệ báo hỏng cao nhất: {df_result['Tỷ lệ báo hỏng (%)'].max():.2f}%")
        print(f"  - Tỷ lệ báo hỏng thấp nhất: {df_result['Tỷ lệ báo hỏng (%)'].min():.2f}%")
        print("-"*80)

        print("\n" + "="*80)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.2 - TỶ LỆ THUÊ BAO BRCĐ BÁO HỎNG")
        print("="*80)

        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.2 - Tỷ lệ thuê bao BRCĐ báo hỏng: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_I15_report(force_update=False):
    """
    Xử lý báo cáo I1.5 với tracking lịch sử V2:
    1. Đọc file I1.5 report.xlsx
    2. Tra cứu thông tin từ danhba.db
    3. Chuẩn hóa cột NVKT_DB
    4. Kiểm tra đã xử lý ngày này chưa (nếu đã có -> bỏ qua lưu DB)
    5. So sánh với dữ liệu ngày hôm qua
    6. Tạo các sheet: TH_SHC_I15, Tang_moi, Giam_het, Van_con, Bien_dong_tong_hop
    7. Lưu vào database để tracking lịch sử (chỉ lần đầu trong ngày)

    Args:
        force_update: Nếu True, cho phép ghi đè dữ liệu đã tồn tại (mặc định False)

    NOTE: Sử dụng version V2 từ i15_process.py
          - Lần 1 trong ngày: Lưu đầy đủ vào DB
          - Lần 2+ trong ngày: Chỉ tạo Excel, không động DB (trừ khi force_update=True)
    """
    # Import hàm mới
    import sys
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from i15_process import process_I15_report_with_tracking

    return process_I15_report_with_tracking(force_update=force_update)


def process_I15_k2_report(force_update=False):
    """
    Xử lý báo cáo I1.5 K2 tương tự I1.5 nhưng dùng file I1.5_k2 report.xlsx
    """
    # Import hàm mới
    import sys
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from i15_process import process_I15_k2_report_with_tracking

    return process_I15_k2_report_with_tracking(force_update=force_update)


if __name__ == "__main__":
    #Test các hàm xử lý
    # process_c11_report()
    # process_c12_report()
    # process_c13_report()
    # process_c14_report()
    # process_c14_chitiet_report()
    # process_c15_chitiet_report()
    # process_c15_report()
    # process_I15_report()
    # process_c11_chitiet_report_SM2()
    # process_c12_chitiet_report_SM1SM2()
    # process_c11_chitiet_report()
    # process_c12_ti_le_bao_hong()
    process_I15_report()
    process_I15_k2_report()
