# -*- coding: utf-8 -*-
"""
Module xử lý các báo cáo KPI NVKT từ file đã tải về.
"""

import os
import pandas as pd
from openpyxl import load_workbook


# Đường dẫn mặc định
DEFAULT_C11_INPUT  = os.path.join("KPI-DOWNLOAD", "c11-nvktdb report.xlsx")
DEFAULT_DSNV_FILE  = "dsnv.xlsx"
# Ghi sheet kết quả vào chính file c11 input
DEFAULT_OUTPUT_FILE = DEFAULT_C11_INPUT


def _chuan_hoa_ten_nvkt(raw_value: str) -> str:
    """
    Chuẩn hóa tên NVKT từ chuỗi dạng:
      'CTV030830-38910007_-tamcm.hni-Chu Minh Tám'
    → 'Chu Minh Tám'  (phần cuối sau dấu '-' cuối cùng)
    """
    if not isinstance(raw_value, str):
        return str(raw_value) if raw_value is not None else ""
    parts = raw_value.strip().split("-")
    return parts[-1].strip()


def c11_process_report_nvkt(
    c11_input: str   = DEFAULT_C11_INPUT,
    dsnv_file: str   = DEFAULT_DSNV_FILE,
    output_file: str = None,   # Mặc định ghi vào chính file c11_input
    sheet_name: str  = "c11 kpi nvkt",
):
    """
    Xử lý báo cáo C1.1 NVKT quản lý địa bàn.

    Args:
        c11_input:   Đường dẫn file c11-nvktdb report.xlsx
        dsnv_file:   Đường dẫn file dsnv.xlsx (để tra cứu đơn vị)
        output_file: Đường dẫn file Excel đầu ra
        sheet_name:  Tên sheet ghi kết quả

    Các bước:
    1. Đọc file c11 với header 2 dòng
    2. Lấy cột Đơn vị (cột A) và chuẩn hóa → tên NVKT
    3. Lấy 7 cột Tổng (B-H): SM1, SM2, Tỷ lệ CĐ, SM3, SM4, Tỷ lệ BRCĐ, Chỉ tiêu BSC
    4. Tra cứu đơn vị từ dsnv.xlsx theo Họ tên == NVKT
    5. Ghi kết quả vào sheet_name trong output_file
    """
    if output_file is None:
        output_file = c11_input

    try:
        print("\n" + "=" * 70)
        print(f"XỬ LÝ BÁO CÁO C1.1 KPI NVKT")
        print("=" * 70)

        # ── Bước 1: Đọc file c11 với header 2 dòng ───────────────────────
        print(f"\n✓ Đang đọc file: {c11_input}")
        df_raw = pd.read_excel(c11_input, header=[0, 1])
        print(f"✅ Đã đọc xong – {len(df_raw)} dòng dữ liệu")

        # ── Bước 2: Lấy cột Đơn vị (cột vị trí 0) và chuẩn hóa ─────────
        print("\n✓ Đang chuẩn hóa cột NVKT...")
        col_donvi_raw = df_raw.iloc[:, 0]            # cột A
        nvkt_series   = col_donvi_raw.apply(_chuan_hoa_ten_nvkt)

        # Bỏ các dòng rỗng / tổng cộng cuối file
        mask = nvkt_series.str.strip().ne("") & col_donvi_raw.notna()
        nvkt_series = nvkt_series[mask].reset_index(drop=True)
        print(f"✅ Chuẩn hóa xong – {len(nvkt_series)} NVKT")

        # ── Bước 3: Lấy 7 cột Tổng (vị trí 1-7 = cột B-H) ──────────────
        output_col_names = [
            "SM1",
            "SM2",
            "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",
            "SM3",
            "SM4",
            "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCĐ đúng quy định không tính hẹn",
            "Chỉ tiêu BSC",
        ]
        df_tong = df_raw.iloc[mask.values, 1:8].copy()
        df_tong.columns = output_col_names
        df_tong = df_tong.reset_index(drop=True)

        # ── Bước 4: Tra cứu đơn vị từ dsnv.xlsx ─────────────────────────
        print(f"\n✓ Đang tra cứu đơn vị từ: {dsnv_file}")
        df_dsnv = pd.read_excel(dsnv_file)
        # Chuẩn hóa tên cột (strip whitespace)
        df_dsnv.columns = [str(c).strip() for c in df_dsnv.columns]
        # Tạo dict tra cứu: Họ tên → đơn vị
        lookup = dict(zip(
            df_dsnv["Họ tên"].astype(str).str.strip(),
            df_dsnv["đơn vị"].astype(str).str.strip()
        ))
        donvi_series = nvkt_series.map(lookup).fillna("")
        print(f"✅ Tra cứu xong – {donvi_series.ne('').sum()}/{len(donvi_series)} tìm thấy đơn vị")

        # ── Bước 5: Tạo DataFrame kết quả ────────────────────────────────
        df_result = pd.DataFrame()
        df_result.insert(0, "STT",     range(1, len(nvkt_series) + 1))
        df_result.insert(1, "đơn vị",  donvi_series.values)
        df_result.insert(2, "NVKT",    nvkt_series.values)
        for i, col in enumerate(output_col_names, start=3):
            df_result.insert(i, col, df_tong[col].values)

        # Sắp xếp theo đơn vị
        df_result = df_result.sort_values("đơn vị", na_position="last").reset_index(drop=True)
        df_result["STT"] = range(1, len(df_result) + 1)

        # ── Bước 6: Ghi vào file Excel ────────────────────────────────────
        print(f"\n✓ Đang ghi sheet '{sheet_name}' vào: {output_file}")
        os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)

        if os.path.exists(output_file):
            # Thêm/ghi đè sheet vào file đã có
            with pd.ExcelWriter(output_file, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Đã ghi {len(df_result)} dòng vào sheet '{sheet_name}'")
        print("\n" + "=" * 70)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.1 KPI NVKT")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.1 KPI NVKT: {e}")
        import traceback
        traceback.print_exc()
        return False


def c12_process_report_nvkt(
    c12_input: str   = os.path.join("KPI-DOWNLOAD", "c12-nvktdb report.xlsx"),
    dsnv_file: str   = DEFAULT_DSNV_FILE,
    output_file: str = None,   # Mặc định ghi vào chính file c12_input
    sheet_name: str  = "c12 kpi nvkt",
):
    """
    Xử lý báo cáo C1.2 NVKT quản lý địa bàn.
    Hoàn toàn tương tự c11_process_report_nvkt, chỉ khác file input và sheet name.

    Args:
        c12_input:   Đường dẫn file c12-nvktdb report.xlsx
        dsnv_file:   Đường dẫn file dsnv.xlsx (để tra cứu đơn vị)
        output_file: Đường dẫn file Excel đầu ra (mặc định = c12_input)
        sheet_name:  Tên sheet ghi kết quả
    """
    if output_file is None:
        output_file = c12_input

    try:
        print("\n" + "=" * 70)
        print(f"XỬ LÝ BÁO CÁO C1.2 KPI NVKT")
        print("=" * 70)

        # ── Bước 1: Đọc file c12 với header 2 dòng ───────────────────────
        print(f"\n✓ Đang đọc file: {c12_input}")
        df_raw = pd.read_excel(c12_input, header=[0, 1])
        print(f"✅ Đã đọc xong – {len(df_raw)} dòng dữ liệu")

        # ── Bước 2: Lấy cột Đơn vị (cột vị trí 0) và chuẩn hóa ─────────
        print("\n✓ Đang chuẩn hóa cột NVKT...")
        col_donvi_raw = df_raw.iloc[:, 0]
        nvkt_series   = col_donvi_raw.apply(_chuan_hoa_ten_nvkt)

        # Bỏ các dòng rỗng / tổng cộng cuối file
        mask = nvkt_series.str.strip().ne("") & col_donvi_raw.notna()
        nvkt_series = nvkt_series[mask].reset_index(drop=True)
        print(f"✅ Chuẩn hóa xong – {len(nvkt_series)} NVKT")

        # ── Bước 3: Lấy 7 cột Tổng (vị trí 1-7 = cột B-H) ──────────────
        output_col_names = [
            "SM1",
            "SM2",
            "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",
            "SM3",
            "SM4",
            "Tỷ lệ sự cố dịch vụ BRCĐ",
            "Chỉ tiêu BSC",
        ]
        df_tong = df_raw.iloc[mask.values, 1:8].copy()
        df_tong.columns = output_col_names
        df_tong = df_tong.reset_index(drop=True)

        # ── Bước 4: Tra cứu đơn vị từ dsnv.xlsx ─────────────────────────
        print(f"\n✓ Đang tra cứu đơn vị từ: {dsnv_file}")
        df_dsnv = pd.read_excel(dsnv_file)
        df_dsnv.columns = [str(c).strip() for c in df_dsnv.columns]
        lookup = dict(zip(
            df_dsnv["Họ tên"].astype(str).str.strip(),
            df_dsnv["đơn vị"].astype(str).str.strip()
        ))
        donvi_series = nvkt_series.map(lookup).fillna("")
        print(f"✅ Tra cứu xong – {donvi_series.ne('').sum()}/{len(donvi_series)} tìm thấy đơn vị")

        # ── Bước 5: Tạo DataFrame kết quả ────────────────────────────────
        df_result = pd.DataFrame()
        df_result.insert(0, "STT",     range(1, len(nvkt_series) + 1))
        df_result.insert(1, "đơn vị",  donvi_series.values)
        df_result.insert(2, "NVKT",    nvkt_series.values)
        for i, col in enumerate(output_col_names, start=3):
            df_result.insert(i, col, df_tong[col].values)

        # Sắp xếp theo đơn vị
        df_result = df_result.sort_values("đơn vị", na_position="last").reset_index(drop=True)
        df_result["STT"] = range(1, len(df_result) + 1)

        # ── Bước 6: Ghi vào file Excel ────────────────────────────────────
        print(f"\n✓ Đang ghi sheet '{sheet_name}' vào: {output_file}")
        os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)

        if os.path.exists(output_file):
            with pd.ExcelWriter(output_file, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Đã ghi {len(df_result)} dòng vào sheet '{sheet_name}'")
        print("\n" + "=" * 70)
        print("✅ HOÀN THÀNH XỬ LÝ BÁO CÁO C1.2 KPI NVKT")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n❌ Lỗi khi xử lý báo cáo C1.2 KPI NVKT: {e}")
        import traceback
        traceback.print_exc()
        return False


# ── Standalone test ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    c11_process_report_nvkt()
    c12_process_report_nvkt()
