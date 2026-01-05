import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from team_config import get_id_to_shortname_mapping
import glob



def make_chart_pttb_thuc_tang_fiber():
    """
    Đọc file thuc_tang_*.xlsx mới nhất từ thư mục baocao_hanoi,
    tạo biểu đồ cột thể hiện Hoàn công, Ngưng PSC, Thực tăng theo 4 đơn vị,
    lưu vào thư mục '/home/vtst/baocaohanoi/chart/thuc_tang_fiber'.
    """
    # Set up paths
    download_folder = "/home/vtst/baocaohanoi/downloads/baocao_hanoi"
    chart_folder = "/home/vtst/baocaohanoi/chart/thuc_tang_fiber"

    # Tìm file thuc_tang_*.xlsx mới nhất
    pattern = os.path.join(download_folder, "thuc_tang_*.xlsx")
    files = glob.glob(pattern)

    if not files:
        print(f"Không tìm thấy file thuc_tang_*.xlsx trong {download_folder}")
        return

    # Lấy file mới nhất dựa vào modification time
    excel_path = max(files, key=os.path.getmtime)
    print(f"Đọc file: {excel_path}")

    # Lấy thời gian tạo file Excel
    file_mtime = os.path.getmtime(excel_path)
    file_time = datetime.fromtimestamp(file_mtime).strftime("%d/%m/%Y %H:%M")

    # Tạo thư mục chart nếu chưa có
    os.makedirs(chart_folder, exist_ok=True)

    try:
        # Đọc file Excel (giả sử dữ liệu ở sheet đầu tiên)
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Lỗi khi đọc file '{excel_path}': {str(e)}")
        return

    # Kiểm tra các cột cần thiết
    required_cols = ['Đơn vị', 'Hoàn công', 'Ngưng PSC', 'Thực tăng']
    for col in required_cols:
        if col not in df.columns:
            print(f"Không tìm thấy cột '{col}' trong file Excel")
            return

    # Bỏ qua dòng TỔNG CỘNG nếu có
    df = df[~df['Đơn vị'].str.contains('TỔNG CỘNG', case=False, na=False)]

    # Chuẩn bị dữ liệu
    labels = df['Đơn vị'].tolist()
    hoan_cong = df['Hoàn công'].tolist()
    ngung_psc = df['Ngưng PSC'].tolist()
    thuc_tang = df['Thực tăng'].tolist()

    # Tính tổng cho legend
    tong_hoan_cong = sum(hoan_cong)
    tong_ngung_psc = sum(ngung_psc)
    tong_thuc_tang = sum(thuc_tang)

    # Tạo biểu đồ
    x = np.arange(len(labels))
    width = 0.25  # Độ rộng mỗi cột

    plt.figure(figsize=(14, 7))

    # Vẽ 3 nhóm cột với tổng trong legend
    bars1 = plt.bar(x - width, hoan_cong, width,
                    label=f'Hoàn công ({int(tong_hoan_cong)})', color='#2ecc71')
    bars2 = plt.bar(x, ngung_psc, width,
                    label=f'Ngưng PSC ({int(tong_ngung_psc)})', color='#e74c3c')
    bars3 = plt.bar(x + width, thuc_tang, width,
                    label=f'Thực tăng ({int(tong_thuc_tang)})', color='#3498db')

    # Tùy chỉnh biểu đồ
    plt.ylabel('Số lượng thuê bao', fontsize=12, fontweight='bold')
    plt.xlabel('Đơn vị', fontsize=12, fontweight='bold')
    plt.title(f'Thống kê thực tăng Fiber theo đơn vị\n{file_time}',
              fontsize=14, fontweight='bold', pad=20)
    plt.xticks(x, labels, rotation=20, ha='right', fontsize=11)
    plt.legend(fontsize=11)
    plt.grid(axis='y', alpha=0.3, linestyle='--')

    # Thêm giá trị lên đầu cột
    def autolabel(bars):
        for bar in bars:
            height = bar.get_height()
            # Hiển thị cả giá trị âm
            va = 'bottom' if height >= 0 else 'top'
            plt.text(bar.get_x() + bar.get_width()/2, height,
                    f'{int(height)}',
                    ha='center', va=va, fontsize=9, fontweight='bold')

    autolabel(bars1)
    autolabel(bars2)
    autolabel(bars3)

    # Thêm đường y=0 để làm rõ các giá trị âm
    plt.axhline(y=0, color='black', linestyle='-', linewidth=0.8)

    # Lưu biểu đồ
    plt.tight_layout()
    chart_path = os.path.join(chart_folder, "thuc_tang_fiber_pttb.png")
    plt.savefig(chart_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Biểu đồ đã được lưu tại: {chart_path}")


def make_chart_pttb_mytv_thuc_tang():
    """
    Đọc file mytv_thuc_tang_*.xlsx mới nhất từ thư mục baocao_hanoi,
    tạo biểu đồ cột thể hiện Hoàn công, Ngưng PSC, Thực tăng theo 4 đơn vị,
    lưu vào thư mục '/home/vtst/baocaohanoi/chart/thuc_tang_mytv'.
    """
    # Set up paths
    download_folder = "/home/vtst/baocaohanoi/downloads/baocao_hanoi"
    chart_folder = "/home/vtst/baocaohanoi/chart/thuc_tang_mytv"

    # Tìm file mytv_thuc_tang_*.xlsx mới nhất
    pattern = os.path.join(download_folder, "mytv_thuc_tang_*.xlsx")
    files = glob.glob(pattern)

    if not files:
        print(f"Không tìm thấy file mytv_thuc_tang_*.xlsx trong {download_folder}")
        return

    # Lấy file mới nhất dựa vào modification time
    excel_path = max(files, key=os.path.getmtime)
    print(f"Đọc file: {excel_path}")

    # Lấy thời gian tạo file Excel
    file_mtime = os.path.getmtime(excel_path)
    file_time = datetime.fromtimestamp(file_mtime).strftime("%d/%m/%Y %H:%M")

    # Tạo thư mục chart nếu chưa có
    os.makedirs(chart_folder, exist_ok=True)

    try:
        # Đọc file Excel (giả sử dữ liệu ở sheet đầu tiên)
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Lỗi khi đọc file '{excel_path}': {str(e)}")
        return

    # Kiểm tra các cột cần thiết
    required_cols = ['Đơn vị', 'Hoàn công', 'Ngưng PSC', 'Thực tăng']
    for col in required_cols:
        if col not in df.columns:
            print(f"Không tìm thấy cột '{col}' trong file Excel")
            return

    # Bỏ qua dòng TỔNG CỘNG nếu có
    df = df[~df['Đơn vị'].str.contains('TỔNG CỘNG', case=False, na=False)]

    # Chuẩn bị dữ liệu
    labels = df['Đơn vị'].tolist()
    hoan_cong = df['Hoàn công'].tolist()
    ngung_psc = df['Ngưng PSC'].tolist()
    thuc_tang = df['Thực tăng'].tolist()

    # Tính tổng cho legend
    tong_hoan_cong = sum(hoan_cong)
    tong_ngung_psc = sum(ngung_psc)
    tong_thuc_tang = sum(thuc_tang)

    # Tạo biểu đồ
    x = np.arange(len(labels))
    width = 0.25  # Độ rộng mỗi cột

    plt.figure(figsize=(14, 7))

    # Vẽ 3 nhóm cột với tổng trong legend
    bars1 = plt.bar(x - width, hoan_cong, width,
                    label=f'Hoàn công ({int(tong_hoan_cong)})', color='#2ecc71')
    bars2 = plt.bar(x, ngung_psc, width,
                    label=f'Ngưng PSC ({int(tong_ngung_psc)})', color='#e74c3c')
    bars3 = plt.bar(x + width, thuc_tang, width,
                    label=f'Thực tăng ({int(tong_thuc_tang)})', color='#3498db')

    # Tùy chỉnh biểu đồ
    plt.ylabel('Số lượng thuê bao', fontsize=12, fontweight='bold')
    plt.xlabel('Đơn vị', fontsize=12, fontweight='bold')
    plt.title(f'Thống kê thực tăng MyTV theo đơn vị\n{file_time}',
              fontsize=14, fontweight='bold', pad=20)
    plt.xticks(x, labels, rotation=20, ha='right', fontsize=11)
    plt.legend(fontsize=11)
    plt.grid(axis='y', alpha=0.3, linestyle='--')

    # Thêm giá trị lên đầu cột
    def autolabel(bars):
        for bar in bars:
            height = bar.get_height()
            # Hiển thị cả giá trị âm
            va = 'bottom' if height >= 0 else 'top'
            plt.text(bar.get_x() + bar.get_width()/2, height,
                    f'{int(height)}',
                    ha='center', va=va, fontsize=9, fontweight='bold')

    autolabel(bars1)
    autolabel(bars2)
    autolabel(bars3)

    # Thêm đường y=0 để làm rõ các giá trị âm
    plt.axhline(y=0, color='black', linestyle='-', linewidth=0.8)

    # Lưu biểu đồ
    plt.tight_layout()
    chart_path = os.path.join(chart_folder, "thuc_tang_mytv_pttb.png")
    plt.savefig(chart_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Biểu đồ đã được lưu tại: {chart_path}")


def make_chart_pttb_thuc_tang_nvkt():
    """
    Đọc file thuc_tang_*.xlsx mới nhất từ thư mục baocao_hanoi,
    sheet 'thuc_tang_theo_NVKT',
    tạo biểu đồ cột thể hiện Hoàn công, Ngưng PSC, Thực tăng theo NVKT,
    sắp xếp theo Thực tăng giảm dần,
    lưu vào thư mục 'chart/thuc_tang_fiber'.
    """
    # Set up paths
    download_folder = "/home/vtst/baocaohanoi/downloads/baocao_hanoi"
    chart_folder = "/home/vtst/baocaohanoi/chart/thuc_tang_fiber"

    # Tìm file thuc_tang_*.xlsx mới nhất
    pattern = os.path.join(download_folder, "thuc_tang_*.xlsx")
    files = glob.glob(pattern)

    if not files:
        print(f"Không tìm thấy file thuc_tang_*.xlsx trong {download_folder}")
        return

    # Lấy file mới nhất dựa vào modification time
    excel_path = max(files, key=os.path.getmtime)
    print(f"Đọc file: {excel_path}")

    # Lấy thời gian tạo file Excel
    file_mtime = os.path.getmtime(excel_path)
    file_time = datetime.fromtimestamp(file_mtime).strftime("%d/%m/%Y %H:%M")

    # Tạo thư mục chart nếu chưa có
    os.makedirs(chart_folder, exist_ok=True)

    try:
        # Đọc sheet 'thuc_tang_theo_NVKT'
        df = pd.read_excel(excel_path, sheet_name='thuc_tang_theo_NVKT')
    except Exception as e:
        print(f"Lỗi khi đọc sheet 'thuc_tang_theo_NVKT' từ '{excel_path}': {str(e)}")
        return

    # Kiểm tra các cột cần thiết
    required_cols = ['NVKT', 'Hoàn công', 'Ngưng PSC', 'Thực tăng']
    for col in required_cols:
        if col not in df.columns:
            print(f"Không tìm thấy cột '{col}' trong sheet 'thuc_tang_theo_NVKT'")
            return

    # Bỏ qua dòng TỔNG CỘNG và dòng có NVKT rỗng
    df = df[~df['NVKT'].str.contains('TỔNG CỘNG', case=False, na=False)]
    df = df[df['NVKT'].notna() & (df['NVKT'] != '')]

    # Sắp xếp theo Thực tăng giảm dần
    df = df.sort_values('Thực tăng', ascending=False)

    # Chuẩn bị dữ liệu
    labels = df['NVKT'].tolist()
    hoan_cong = df['Hoàn công'].tolist()
    ngung_psc = df['Ngưng PSC'].tolist()
    thuc_tang = df['Thực tăng'].tolist()

    # Tính tổng cho legend
    tong_hoan_cong = sum(hoan_cong)
    tong_ngung_psc = sum(ngung_psc)
    tong_thuc_tang = sum(thuc_tang)

    # Tạo biểu đồ với kích thước lớn hơn vì có nhiều NVKT
    x = np.arange(len(labels))
    width = 0.25  # Độ rộng mỗi cột

    fig_width = max(16, len(labels) * 0.5)  # Tự động điều chỉnh độ rộng
    plt.figure(figsize=(fig_width, 8))

    # Vẽ 3 nhóm cột với tổng trong legend
    bars1 = plt.bar(x - width, hoan_cong, width,
                    label=f'Hoàn công ({int(tong_hoan_cong)})', color='#2ecc71')
    bars2 = plt.bar(x, ngung_psc, width,
                    label=f'Ngưng PSC ({int(tong_ngung_psc)})', color='#e74c3c')
    bars3 = plt.bar(x + width, thuc_tang, width,
                    label=f'Thực tăng ({int(tong_thuc_tang)})', color='#3498db')

    # Tùy chỉnh biểu đồ
    plt.ylabel('Số lượng thuê bao', fontsize=12, fontweight='bold')
    plt.xlabel('NVKT', fontsize=12, fontweight='bold')
    plt.title(f'Thống kê thực tăng Fiber theo NVKT\n{file_time}',
              fontsize=14, fontweight='bold', pad=20)
    plt.xticks(x, labels, rotation=45, ha='right', fontsize=9)
    plt.legend(fontsize=11, loc='upper right')
    plt.grid(axis='y', alpha=0.3, linestyle='--')

    # Thêm giá trị lên đầu cột
    def autolabel(bars):
        for bar in bars:
            height = bar.get_height()
            # Hiển thị cả giá trị âm
            va = 'bottom' if height >= 0 else 'top'
            plt.text(bar.get_x() + bar.get_width()/2, height,
                    f'{int(height)}',
                    ha='center', va=va, fontsize=7, fontweight='bold')

    autolabel(bars1)
    autolabel(bars2)
    autolabel(bars3)

    # Thêm đường y=0 để làm rõ các giá trị âm
    plt.axhline(y=0, color='black', linestyle='-', linewidth=0.8)

    # Lưu biểu đồ
    plt.tight_layout()
    chart_path = os.path.join(chart_folder, "fiber_thuctang_nvkt.png")
    plt.savefig(chart_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Biểu đồ đã được lưu tại: {chart_path}")


def make_chart_mytv_thuc_tang_nvkt():
    """
    Đọc file mytv_thuc_tang_*.xlsx mới nhất từ thư mục baocao_hanoi,
    sheet 'thuc_tang_theo_NVKT',
    tạo biểu đồ cột thể hiện Hoàn công, Ngưng PSC, Thực tăng theo NVKT,
    sắp xếp theo Thực tăng giảm dần,
    lưu vào thư mục 'chart/thuc_tang_mytv'.
    """
    # Set up paths
    download_folder = "/home/vtst/baocaohanoi/downloads/baocao_hanoi"
    chart_folder = "/home/vtst/baocaohanoi/chart/thuc_tang_mytv"

    # Tìm file mytv_thuc_tang_*.xlsx mới nhất
    pattern = os.path.join(download_folder, "mytv_thuc_tang_*.xlsx")
    files = glob.glob(pattern)

    if not files:
        print(f"Không tìm thấy file mytv_thuc_tang_*.xlsx trong {download_folder}")
        return

    # Lấy file mới nhất dựa vào modification time
    excel_path = max(files, key=os.path.getmtime)
    print(f"Đọc file: {excel_path}")

    # Lấy thời gian tạo file Excel
    file_mtime = os.path.getmtime(excel_path)
    file_time = datetime.fromtimestamp(file_mtime).strftime("%d/%m/%Y %H:%M")

    # Tạo thư mục chart nếu chưa có
    os.makedirs(chart_folder, exist_ok=True)

    try:
        # Đọc sheet 'thuc_tang_theo_NVKT'
        df = pd.read_excel(excel_path, sheet_name='thuc_tang_theo_NVKT')
    except Exception as e:
        print(f"Lỗi khi đọc sheet 'thuc_tang_theo_NVKT' từ '{excel_path}': {str(e)}")
        return

    # Kiểm tra các cột cần thiết
    required_cols = ['NVKT', 'Hoàn công', 'Ngưng PSC', 'Thực tăng']
    for col in required_cols:
        if col not in df.columns:
            print(f"Không tìm thấy cột '{col}' trong sheet 'thuc_tang_theo_NVKT'")
            return

    # Bỏ qua dòng TỔNG CỘNG và dòng có NVKT rỗng
    df = df[~df['NVKT'].str.contains('TỔNG CỘNG', case=False, na=False)]
    df = df[df['NVKT'].notna() & (df['NVKT'] != '')]

    # Sắp xếp theo Thực tăng giảm dần
    df = df.sort_values('Thực tăng', ascending=False)

    # Chuẩn bị dữ liệu
    labels = df['NVKT'].tolist()
    hoan_cong = df['Hoàn công'].tolist()
    ngung_psc = df['Ngưng PSC'].tolist()
    thuc_tang = df['Thực tăng'].tolist()

    # Tính tổng cho legend
    tong_hoan_cong = sum(hoan_cong)
    tong_ngung_psc = sum(ngung_psc)
    tong_thuc_tang = sum(thuc_tang)

    # Tạo biểu đồ với kích thước lớn hơn vì có nhiều NVKT
    x = np.arange(len(labels))
    width = 0.25  # Độ rộng mỗi cột

    fig_width = max(16, len(labels) * 0.5)  # Tự động điều chỉnh độ rộng
    plt.figure(figsize=(fig_width, 8))

    # Vẽ 3 nhóm cột với tổng trong legend
    bars1 = plt.bar(x - width, hoan_cong, width,
                    label=f'Hoàn công ({int(tong_hoan_cong)})', color='#2ecc71')
    bars2 = plt.bar(x, ngung_psc, width,
                    label=f'Ngưng PSC ({int(tong_ngung_psc)})', color='#e74c3c')
    bars3 = plt.bar(x + width, thuc_tang, width,
                    label=f'Thực tăng ({int(tong_thuc_tang)})', color='#3498db')

    # Tùy chỉnh biểu đồ
    plt.ylabel('Số lượng thuê bao', fontsize=12, fontweight='bold')
    plt.xlabel('NVKT', fontsize=12, fontweight='bold')
    plt.title(f'Thống kê thực tăng MyTV theo NVKT\n{file_time}',
              fontsize=14, fontweight='bold', pad=20)
    plt.xticks(x, labels, rotation=45, ha='right', fontsize=9)
    plt.legend(fontsize=11, loc='upper right')
    plt.grid(axis='y', alpha=0.3, linestyle='--')

    # Thêm giá trị lên đầu cột
    def autolabel(bars):
        for bar in bars:
            height = bar.get_height()
            # Hiển thị cả giá trị âm
            va = 'bottom' if height >= 0 else 'top'
            plt.text(bar.get_x() + bar.get_width()/2, height,
                    f'{int(height)}',
                    ha='center', va=va, fontsize=7, fontweight='bold')

    autolabel(bars1)
    autolabel(bars2)
    autolabel(bars3)

    # Thêm đường y=0 để làm rõ các giá trị âm
    plt.axhline(y=0, color='black', linestyle='-', linewidth=0.8)

    # Lưu biểu đồ
    plt.tight_layout()
    chart_path = os.path.join(chart_folder, "mytv_thuctang_nvkt.png")
    plt.savefig(chart_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Biểu đồ đã được lưu tại: {chart_path}")


# Main execution
if __name__ == "__main__":
    print("Tạo biểu đồ PTTB...")

    # Biểu đồ theo đơn vị
    print("\n=== Tạo biểu đồ theo đơn vị ===")
    make_chart_pttb_thuc_tang_fiber()
    make_chart_pttb_mytv_thuc_tang()

    # Biểu đồ theo NVKT
    print("\n=== Tạo biểu đồ theo NVKT ===")
    make_chart_pttb_thuc_tang_nvkt()
    make_chart_mytv_thuc_tang_nvkt()

    print("\n✅ Hoàn thành tạo tất cả biểu đồ!")
