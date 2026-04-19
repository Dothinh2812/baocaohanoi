import pandas as pd
from datetime import datetime
from pathlib import Path
from api_transition.onebss_auth import create_session, close_session
from api_transition.onebss_report_client import (
    http_json_request, 
    build_onebss_headers,
    sanitize_filename
)
from api_transition.onebss_downloaders import group_output_dir

def download_ds_phieu_nghiem_thu_bao_hong(
    loaidvvt_id=1,
    ttbh_id=3,
    nhanvien_id="4581",
    huonggiao_id=1251,
    tungay=None,
    denngay=None,
    ma_tb="0",
    giaoviec=0,
    session=None,
    output_dir=None,
    output_name=None
):
    """
    Tải danh sách phiếu nghiệm thu báo hỏng từ Web-CCDV API.
    Dữ liệu được trả về dưới dạng JSON và chuyển đổi sang Excel.
    """
    # 1. Khởi tạo session và tham số mặc định
    should_close_session = False
    if session is None:
        session = create_session()
        should_close_session = True

    if output_dir is None:
        output_dir = group_output_dir("onebss")
    
    # Định dạng ngày nếu chưa truyền (mặc định là hôm nay)
    today_str = datetime.now().strftime("%d/%m/%Y")
    tungay = tungay or today_str
    denngay = denngay or today_str

    if output_name is None:
        safe_tungay = tungay.replace("/", "")
        safe_denngay = denngay.replace("/", "")
        output_name = f"ds_phieu_bao_hong_{huonggiao_id}_{safe_tungay}_{safe_denngay}.xlsx"

    # 2. Dựng Payload và Header
    api_url = f"{session['api_base_url']}/web-ccdv/xuly_nghiemthubaohong/lay_ds_phieu_hoancong_bh_v5"
    
    headers = build_onebss_headers(session["headers"], include_apikey=True)
    # Ghi đè selectedmenuid từ capture nếu cần
    headers["selectedmenuid"] = "13248" 

    payload = {
        "loaidvvt_id": loaidvvt_id,
        "ttbh_id": ttbh_id,
        "nhanvien_id": str(nhanvien_id),
        "ma_tb": str(ma_tb),
        "huonggiao_id": huonggiao_id,
        "giaoviec": giaoviec,
        "tungay": tungay,
        "denngay": denngay
    }

    try:
        # 3. Gọi API
        print(f"[*] Đang lấy danh sách phiếu báo hỏng từ {tungay} đến {denngay}...")
        response_json = http_json_request(
            method="POST",
            url=api_url,
            payload=payload,
            headers=headers
        )

        # 4. Xử lý dữ liệu JSON (Cấu trúc OneBSS thường trả dữ liệu trong field 'data')
        # Tùy theo API, data có thể là list trực tiếp hoặc nằm trong ['data']
        data = response_json.get("data", [])
        if not data and isinstance(response_json, list):
            data = response_json

        if not data:
            print("[!] Không có dữ liệu phiếu báo hỏng trong khoảng thời gian này.")
            return None

        # 5. Chuyển đổi sang Excel bằng Pandas
        df = pd.DataFrame(data)
        
        # Sắp xếp hoặc lọc cột nếu cần thiết (optional)
        # df = df[['MA_TB', 'TEN_TB', 'DIACHI_LD', 'NGAY_BH', 'TEN_TTVT', 'TRANGTHAI_BH']]

        out_path = Path(output_dir) / output_name
        out_path.parent.mkdir(parents=True, exist_ok=True)
        
        df.to_excel(out_path, index=False)
        print(f"[+] Đã lưu {len(df)} dòng dữ liệu vào: {out_path}")

        return str(out_path)

    except Exception as e:
        print(f"[-] Lỗi khi tải danh sách phiếu báo hỏng: {e}")
        return None
    finally:
        if should_close_session:
            close_session(session)

# --- Cách sử dụng ---
if __name__ == "__main__":
    # Ví dụ chạy độc lập
    path = download_ds_phieu_nghiem_thu_bao_hong(
        tungay="09/04/2026",
        denngay="17/04/2026",
        huonggiao_id=1251 # Trạm VT xử lý sự cố cố định
    )
    print(f"Kết quả: {path}")