import pandas as pd
import time
import os
from datetime import datetime, timedelta
import requests
import zipfile
import cloudinary
import cloudinary.uploader
from cloudinary.utils import cloudinary_url
from typing import Tuple

# Import configuration from config module
from config import (
    get_cloudinary_config,
    SEND_START_HOUR,
    SEND_START_MINUTE,
    SEND_END_HOUR,
    SEND_END_MINUTE,
    WEBHOOK_TEXT_URL,
    LOCATION_THREAD_MAPPING,
    LOCATION_CHAT_MAPPING
)

# Import team_config for dynamic team mappings
from team_config import get_active_teams

# Load environment variables
from dotenv import load_dotenv
load_dotenv()

# Configure Cloudinary using config module
cloudinary.config(**get_cloudinary_config())



def get_khdn_exclusion_list():
    """
    Lấy danh sách mã TB bị loại trừ từ biến môi trường KHDN_EXCLUSION_LIST.

    QUAN TRỌNG: Hàm này LUÔN load lại .env file để đảm bảo lấy được giá trị mới nhất,
    ngay cả khi ứng dụng đã được restart.

    Returns:
        set: Tập hợp các mã TB sẽ bị loại trừ (empty set nếu không có)

    Example:
        Nếu .env có: KHDN_EXCLUSION_LIST=MW001171760,MW001171761,MW001171762
        Hàm sẽ trả về: {'MW001171760', 'MW001171761', 'MW001171762'}
    """
    # Load lại .env file mỗi lần gọi hàm để đảm bảo lấy giá trị mới nhất
    from dotenv import load_dotenv
    load_dotenv(override=True)

    exclusion_str = os.getenv('KHDN_EXCLUSION_LIST', '')

    # Debug: In ra giá trị được đọc
    debug_enabled = os.getenv('DEBUG_EXCLUSION_LIST', 'false').lower() == 'true'
    if debug_enabled:
        print(f"[DEBUG] KHDN_EXCLUSION_LIST value: '{exclusion_str}'")

    if not exclusion_str.strip():
        return set()

    # Tách các mã bằng dấu phẩy và loại trừ khoảng trắng
    exclusion_list = set(code.strip() for code in exclusion_str.split(',') if code.strip())

    if debug_enabled:
        print(f"[DEBUG] Parsed exclusion_list: {exclusion_list}")

    return exclusion_list


def is_allowed_send_time() -> Tuple[bool, str]:
    """
    Kiểm tra xem hiện tại có phải thời gian được phép gửi cảnh báo không.

    Thời gian CHO PHÉP gửi: 06:30 - 21:00
    Thời gian KHÔNG cho phép: 21:00 - 06:30 (ngày hôm sau)

    Returns:
        Tuple[bool, str]: (True nếu được phép gửi, message giải thích)
    """
    now = datetime.now()

    # Tạo thời điểm bắt đầu và kết thúc trong ngày hiện tại
    start_time = now.replace(hour=SEND_START_HOUR, minute=SEND_START_MINUTE, second=0, microsecond=0)
    end_time = now.replace(hour=SEND_END_HOUR, minute=SEND_END_MINUTE, second=0, microsecond=0)

    # Kiểm tra xem hiện tại có nằm trong khoảng thời gian cho phép không
    if start_time <= now <= end_time:
        return True, f"✅ Thời gian hiện tại ({now.strftime('%H:%M')}) nằm trong khung giờ cho phép gửi (06:30-21:00)"
    else:
        if now < start_time:
            # Trước 6:30 sáng
            return False, f"⏰ Thời gian hiện tại ({now.strftime('%H:%M')}) quá sớm. Vui lòng đợi đến 06:30"
        else:
            # Sau 21:00 tối
            next_send_time = (now + timedelta(days=1)).replace(hour=SEND_START_HOUR, minute=SEND_START_MINUTE, second=0, microsecond=0)
            return False, f"🌙 Thời gian hiện tại ({now.strftime('%H:%M')}) quá muộn. Cảnh báo sẽ được gửi vào {next_send_time.strftime('%d/%m/%Y %H:%M')}"


def get_latest_file_in_dir(directory):
    """
    Trả về đường dẫn file mới nhất trong thư mục (dựa trên thời gian chỉnh sửa).
    """
    import os
    files = [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def send_image_to_zalo(receiver='xuanthinh'):
    """
    Gửi tất cả ảnh từ các thư mục image và chart đến Zalo

    Chỉ gửi trong khung giờ 06:30 - 21:00
    """
    print(f"🚀 Bắt đầu gửi ảnh đến {receiver}")

    # Kiểm tra thời gian cho phép gửi
    allowed, time_msg = is_allowed_send_time()
    print(time_msg)

    if not allowed:
        print("⏭️  BỎ QUA gửi ảnh (ngoài khung giờ cho phép)")
        return 0, 0
    
    # Danh sách các thư mục cần gửi ảnh
    directories = [
        ("image", "📸")
    ]
    
    total_sent = 0
    total_failed = 0
    
    for dir_name, icon in directories:
        if os.path.exists(dir_name):
            # Tìm tất cả file PNG trong thư mục
            files = [f for f in os.listdir(dir_name) if f.lower().endswith('.png')]
            
            if files:
                print(f"{icon} Tìm thấy {len(files)} ảnh trong thư mục {dir_name}")
                
                for file_name in files:
                    file_path = os.path.join(dir_name, file_name)
                    # Lấy tên file không có phần mở rộng
                    display_name = os.path.splitext(file_name)[0]
                    
                    print(f"📤 Đang gửi: {display_name}")
                    
                    # Gửi ảnh
                    success = send_image_via_webhook(file_path, receiver=receiver, message=display_name)
                    
                    if success:
                        total_sent += 1
                        print(f"✅ Gửi thành công: {display_name}")
                    else:
                        total_failed += 1
                        print(f"❌ Gửi thất bại: {display_name}")
                    
                    # Tạm dừng 1 giây giữa các lần gửi để tránh spam
                    time.sleep(1)
            else:
                print(f"ℹ️ Không tìm thấy ảnh PNG nào trong thư mục {dir_name}")
        else:
            print(f"⚠️ Thư mục {dir_name} không tồn tại")
    
    # Tổng kết
    print(f"\n📋 TỔNG KẾT:")
    print(f"✅ Gửi thành công: {total_sent} ảnh")
    print(f"❌ Gửi thất bại: {total_failed} ảnh")
    print(f"📊 Tổng cộng: {total_sent + total_failed} ảnh")
    
    return total_sent, total_failed


def send_warning_phieu_ton_brcd():
    """
    Gửi cảnh báo phiếu tồn BRCD sắp quá giờ tới 4 nhóm Zalo + Telegram.
    Đọc từ file chiTietBrcd5Doi.xlsx, từ 4 sheet: thachthat, hoalac, hatmon, odien
    Gửi cảnh báo cho các phiếu sắp quá giờ (0 < giờ còn lại thực < 1.5).
    """
    file_path = os.path.join('chiaTheoDoi', 'chiTietBrcd5Doi.xlsx')
    webhook_url = WEBHOOK_TEXT_URL

    # Import Telegram config
    try:
        from send_tele import (
            TELEGRAM_TOKEN
        )
    except ImportError:
        print("⚠️ Không thể import module send_tele. Chỉ gửi Zalo.")
        TELEGRAM_TOKEN = None

    # Auto-generate sheet mappings from team_config (BRCD - 4 teams)
    brcd_teams = get_active_teams('BRCD')

    sheet_to_thread_id = {
        team.id: LOCATION_THREAD_MAPPING[team.id]
        for team in brcd_teams
    }

    sheet_to_telegram_chat_id = {
        team.id: LOCATION_CHAT_MAPPING.get(team.id)
        for team in brcd_teams
    }

    try:
        print(f"\n🚀 Bắt đầu gửi cảnh báo phiếu tồn BRCD...")

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Tổng số bản tin đã gửi
        total_sent_zalo = 0
        total_failed_zalo = 0
        total_sent_telegram = 0
        total_failed_telegram = 0

        # Đọc từng sheet
        for sheet_name, thread_id in sheet_to_thread_id.items():
            try:
                print(f"\n📋 Đang xử lý sheet: {sheet_name}")

                # Đọc sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                if df.empty:
                    print(f"  ℹ️ Sheet {sheet_name} không có dữ liệu")
                    continue

                # Kiểm tra các cột cần thiết
                required_cols = ['giờ còn lại thực', 'ma_tb', 'ngay_bh', 'ds_nhanvien_th', 'Trạng thái cổng']
                missing_cols = [col for col in required_cols if col not in df.columns]

                if missing_cols:
                    print(f"  ⚠️ Sheet {sheet_name} thiếu các cột: {', '.join(missing_cols)}")
                    continue

                # Kiểm tra xem có cột ghichuton không (không bắt buộc, dùng để hiển thị lý do tồn)
                has_ghichuton = 'ghichuton' in df.columns

                # Lọc phiếu sắp quá giờ: 0 < giờ còn lại thực < 1
                # Chỉ gửi cảnh báo các phiếu sắp quá giờ, không gửi phiếu đã quá giờ
                df_filtered = df[(df['giờ còn lại thực'] > 0) & (df['giờ còn lại thực'] < 1.5)].copy()

                if df_filtered.empty:
                    print(f"  ℹ️ Không có phiếu sắp quá giờ trong sheet {sheet_name}")
                    continue

                print(f"  🔍 Tìm thấy {len(df_filtered)} phiếu sắp quá giờ")

                # Chuyển đổi ngày báo hỏng sang định dạng dễ đọc
                if 'ngay_bh' in df_filtered.columns:
                    df_filtered['ngay_bh_formatted'] = pd.to_datetime(
                        df_filtered['ngay_bh'],
                        errors='coerce'
                    ).dt.strftime('%d/%m/%Y %H:%M')
                else:
                    df_filtered['ngay_bh_formatted'] = 'N/A'

                # Tạo bản tin cho từng phiếu
                messages_to_send = []

                for _, row in df_filtered.iterrows():
                    # Lấy giờ còn lại thực (làm tròn 1 chữ số thập phân)
                    gio_con_lai = row['giờ còn lại thực']

                    # Lấy trạng thái cổng
                    trang_thai_cong = row.get('Trạng thái cổng', 'N/A')
                    if pd.isna(trang_thai_cong):
                        trang_thai_cong = 'N/A'

                    # Lấy lý do tồn (nếu có cột ghichuton)
                    if has_ghichuton:
                        ly_do_ton = row.get('ghichuton', 'N/A')
                    else:
                        ly_do_ton = 'N/A'

                    # Xử lý giá trị null cho ly_do_ton
                    if pd.isna(ly_do_ton):
                        ly_do_ton = 'N/A'

                    # Tạo message mới với format cảnh báo sắp quá giờ
                    message = (
                        f"🔔 Cảnh báo phiếu tồn sắp quá\n"
                        f"  - Mã TB: {row.get('ma_tb', 'N/A')}\n"
                        f"  - Ngày báo: {row.get('ngay_bh_formatted', 'N/A')}\n"
                        f"  - NVKT: {row.get('ds_nhanvien_th', 'N/A')}\n"
                        f"  - Trạng thái cổng: {trang_thai_cong}\n"
                        f"  - Lý do tồn: {ly_do_ton}\n"
                        f"  - Thời gian còn: {gio_con_lai:.1f} giờ"
                    )

                    messages_to_send.append(message)

                # Gửi từng bản tin
                print(f"  📤 Đang gửi {len(messages_to_send)} bản tin cho sheet {sheet_name}...")

                sent_count_zalo = 0
                failed_count_zalo = 0
                sent_count_telegram = 0
                failed_count_telegram = 0

                # Lấy Telegram chat ID cho sheet này
                telegram_chat_id = sheet_to_telegram_chat_id.get(sheet_name)

                for i, message in enumerate(messages_to_send, 1):
                    # Gửi Zalo
                    try:
                        # Chuẩn bị dữ liệu để gửi qua webhook
                        data = {
                            'threadID': thread_id,
                            'message': message
                        }

                        # Gửi request
                        response = requests.get(webhook_url, json=data, timeout=10)

                        if response.status_code == 200:
                            sent_count_zalo += 1
                            print(f"    ✅ Zalo [{i}/{len(messages_to_send)}] Gửi thành công")
                        else:
                            failed_count_zalo += 1
                            print(f"    ❌ Zalo [{i}/{len(messages_to_send)}] Gửi thất bại. Status: {response.status_code}")

                    except Exception as e:
                        failed_count_zalo += 1
                        print(f"    ❌ Zalo [{i}/{len(messages_to_send)}] Lỗi khi gửi: {e}")

                    # Gửi Telegram
                    if telegram_chat_id and TELEGRAM_TOKEN:
                        try:
                            import requests as telegram_requests
                            url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
                            data = {
                                "chat_id": telegram_chat_id,
                                "text": message,
                                "parse_mode": "HTML"
                            }
                            response = telegram_requests.post(url, data=data, timeout=10)

                            if response.status_code == 200:
                                sent_count_telegram += 1
                                print(f"    ✅ Telegram [{i}/{len(messages_to_send)}] Gửi thành công")
                            else:
                                failed_count_telegram += 1
                                print(f"    ❌ Telegram [{i}/{len(messages_to_send)}] Gửi thất bại. Status: {response.status_code}")

                        except Exception as e:
                            failed_count_telegram += 1
                            print(f"    ❌ Telegram [{i}/{len(messages_to_send)}] Lỗi khi gửi: {e}")

                    # Tạm dừng 1 giây để tránh spam
                    if i < len(messages_to_send):
                        time.sleep(1)

                total_sent_zalo += sent_count_zalo
                total_failed_zalo += failed_count_zalo
                total_sent_telegram += sent_count_telegram
                total_failed_telegram += failed_count_telegram

                print(f"  📊 Sheet {sheet_name}: Zalo ✅ {sent_count_zalo} | ❌ {failed_count_zalo} | Telegram ✅ {sent_count_telegram} | ❌ {failed_count_telegram}")

            except ValueError as e:
                if "Worksheet named" in str(e):
                    print(f"  ⚠️ Sheet '{sheet_name}' không tồn tại trong file")
                else:
                    print(f"  ❌ ValueError khi xử lý sheet {sheet_name}: {e}")
                continue
            except Exception as e:
                print(f"  ❌ Lỗi khi xử lý sheet {sheet_name}: {e}")
                continue

        # Tổng kết
        print("\n" + "=" * 60)
        print("📋 TỔNG KẾT:")
        print(f"  [Zalo]     ✅ Gửi thành công: {total_sent_zalo} bản tin")
        print(f"  [Zalo]     ❌ Gửi thất bại: {total_failed_zalo} bản tin")
        print(f"  [Telegram] ✅ Gửi thành công: {total_sent_telegram} bản tin")
        print(f"  [Telegram] ❌ Gửi thất bại: {total_failed_telegram} bản tin")
        print(f"  📊 Tổng cộng: {total_sent_zalo + total_failed_zalo + total_sent_telegram + total_failed_telegram} bản tin")
        print("=" * 60)

        return True

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file {file_path}")
        return False
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_warning_hong_lai_trong_thang():
    """
    Gửi cảnh báo phiếu Hỏng lại trong tháng tới 3 nhóm Zalo và Telegram.
    Đọc từ file bc_BRCD.xlsx, sheet chi_tiet_ton_brcd.
    Gửi cảnh báo cho các phiếu có LAN_HONG > 1 (hỏng lại trong tháng).
    Có cơ chế log để tránh gửi trùng lặp trong vòng 4 tiếng.

    Chỉ gửi trong khung giờ 06:30 - 21:00
    """
    print("\n" + "=" * 70)
    print("🔔 KIỂM TRA GỬI CẢNH BÁO HỎNG LẠI TRONG THÁNG")
    print("=" * 70)

    # Kiểm tra thời gian cho phép gửi
    allowed, time_msg = is_allowed_send_time()
    print(time_msg)

    if not allowed:
        print("⏭️  BỎ QUA gửi cảnh báo (ngoài khung giờ cho phép)")
        print("=" * 70 + "\n")
        return

    # Import send_tele để gửi telegram
    try:
        from send_tele import (
            ID_CHAT_SONTAY, ID_CHAT_SUOIHAI, ID_CHAT_QUANGOAI,
            TELEGRAM_TOKEN, send_message as send_telegram_message
        )
        import requests as telegram_requests
    except ImportError:
        print("⚠️ Không thể import module send_tele. Chỉ gửi Zalo.")
        ID_CHAT_SONTAY = None
        ID_CHAT_SUOIHAI = None
        ID_CHAT_QUANGOAI = None
        TELEGRAM_TOKEN = None

    file_path = os.path.join('downloads', 'kq_dhsc', 'bc_BRCD.xlsx')
    webhook_url_zalo = WEBHOOK_TEXT_URL

    # Cấu hình log
    log_dir = "log_message"
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "hong_lai_trong_thang_log.xlsx")

    # Thời gian chờ giữa các lần gửi (4 tiếng = 240 phút)
    TIME_THRESHOLD = pd.Timedelta(hours=4)

    # Mapping DOI_VT sang thread_id Zalo và Telegram chat_id
    # Dựa vào sheet_to_thread_id trong send_warning_phieu_ton_brcd
    doi_vt_mapping = {
        'Tổ Kỹ thuật Địa bàn Sơn Tây': {
            'thread_id': '6337217534995887511',
            'telegram_chat_id': ID_CHAT_SONTAY if ID_CHAT_SONTAY else None,
            'display_name': 'Sơn Tây'
        },
        'Tổ Kỹ thuật Địa bàn Suối Hai': {
            'thread_id': '6085297980620830486',
            'telegram_chat_id': ID_CHAT_SUOIHAI if ID_CHAT_SUOIHAI else None,
            'display_name': 'Suối Hai'
        },
        'Tổ Kỹ thuật Địa bàn Quảng Oai': {
            'thread_id': '5364152493553904404',
            'telegram_chat_id': ID_CHAT_QUANGOAI if ID_CHAT_QUANGOAI else None,
            'display_name': 'Quảng Oai'
        }
    }

    try:
        print(f"\n🚀 Bắt đầu gửi cảnh báo phiếu Hỏng lại trong tháng...")
        current_time = datetime.now()

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Đọc sheet chi_tiet_ton_brcd
        df = pd.read_excel(file_path, sheet_name='chi_tiet_ton_brcd')

        if df.empty:
            print(f"  ℹ️ Sheet chi_tiet_ton_brcd không có dữ liệu")
            return True

        # Kiểm tra các cột cần thiết
        required_cols = ['DOI_VT', 'NVKT', 'ma_tb', 'LAN_HONG', 'thời gian tồn thực', 'Số lần KHL']

        # Thêm cột ID báo hỏng nếu có (hoặc dùng index)
        if 'ID_BH' not in df.columns and 'SO_PHIEU_BH' not in df.columns:
            # Tạo ID tạm từ ma_tb + ngay_bh nếu không có cột ID
            print("  ℹ️ Không tìm thấy cột ID báo hỏng, sử dụng ma_tb làm ID")
            df['ID_BH'] = df['ma_tb'].astype(str)
        elif 'SO_PHIEU_BH' in df.columns:
            df['ID_BH'] = df['SO_PHIEU_BH'].astype(str)

        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            print(f"  ⚠️ Sheet chi_tiet_ton_brcd thiếu các cột: {', '.join(missing_cols)}")
            return False

        # Lọc phiếu có LAN_HONG > 1 (hỏng lại trong tháng)
        df_filtered = df[df['LAN_HONG'] > 1].copy()

        if df_filtered.empty:
            print(f"  ℹ️ Không có phiếu nào hỏng lại trong tháng (LAN_HONG > 1)")
            return True

        print(f"  🔍 Tìm thấy {len(df_filtered)} phiếu hỏng lại trong tháng")

        # Load existing log
        existing_log = pd.DataFrame()
        if os.path.exists(log_file):
            try:
                existing_log = pd.read_excel(log_file)
                existing_log['Thời gian gửi'] = pd.to_datetime(existing_log['Thời gian gửi'], errors='coerce')
                print(f"  📂 Đã load log từ file: {len(existing_log)} bản ghi")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi đọc log file: {e}")
                existing_log = pd.DataFrame()

        # Tổng số bản tin đã gửi
        total_sent_zalo = 0
        total_failed_zalo = 0
        total_sent_telegram = 0
        total_failed_telegram = 0
        total_skipped = 0

        # Danh sách log mới
        log_data = []

        # Nhóm theo DOI_VT để gửi từng nhóm
        grouped = df_filtered.groupby('DOI_VT')

        for doi_vt, group in grouped:
            try:
                print(f"\n📋 Đang xử lý {doi_vt}: {len(group)} phiếu")

                # Lấy mapping info
                mapping_info = doi_vt_mapping.get(doi_vt)
                if not mapping_info:
                    print(f"  ⚠️ Không tìm thấy mapping cho {doi_vt}, bỏ qua")
                    continue

                thread_id = mapping_info['thread_id']
                telegram_chat_id = mapping_info['telegram_chat_id']
                display_name = mapping_info['display_name']

                print(f"  📤 Thread ID (Zalo): {thread_id}")
                print(f"  📤 Chat ID (Telegram): {telegram_chat_id}")

                # Gửi từng phiếu
                sent_zalo = 0
                failed_zalo = 0
                sent_telegram = 0
                failed_telegram = 0
                skipped = 0

                for idx, row in group.iterrows():
                    # Lấy thông tin từ row
                    id_bh = row.get('ID_BH', 'N/A')
                    nvkt = row.get('NVKT', 'N/A')
                    ma_tb = row.get('ma_tb', 'N/A')
                    lan_hong = row.get('LAN_HONG', 0)
                    thoi_gian_ton = row.get('thời gian tồn thực', 0)
                    so_lan_khl = row.get('Số lần KHL', 0)

                    # Lấy tên thuê bao nếu có
                    ten_tb = row.get('TEN_TB', row.get('TEN_KH', 'N/A'))

                    # Xử lý giá trị null
                    if pd.isna(id_bh):
                        id_bh = 'N/A'
                    if pd.isna(nvkt):
                        nvkt = 'N/A'
                    if pd.isna(ma_tb):
                        ma_tb = 'N/A'
                    if pd.isna(ten_tb):
                        ten_tb = 'N/A'
                    if pd.isna(lan_hong):
                        lan_hong = 0
                    if pd.isna(thoi_gian_ton):
                        thoi_gian_ton = 0
                    if pd.isna(so_lan_khl):
                        so_lan_khl = 0

                    # Kiểm tra log - đã gửi trong vòng 4 tiếng chưa?
                    send_count = 0
                    should_skip = False

                    if not existing_log.empty:
                        # Tìm các bản ghi matching trong log
                        matching_records = existing_log[
                            (existing_log['ID báo hỏng'] == str(id_bh)) &
                            (existing_log['Mã thuê bao'] == str(ma_tb))
                        ]

                        if not matching_records.empty:
                            # Lấy bản ghi gần nhất
                            most_recent = matching_records['Thời gian gửi'].max()
                            time_since_last = current_time - most_recent

                            # Lấy số lần đã gửi
                            send_count = matching_records['Số lần gửi'].max()

                            # Kiểm tra nếu trong vòng 4 tiếng
                            if time_since_last < TIME_THRESHOLD:
                                should_skip = True
                                skipped += 1
                                total_skipped += 1
                                print(f"    ⏭️ Bỏ qua {ma_tb} (đã gửi {send_count} lần, lần cuối: {most_recent.strftime('%Y-%m-%d %H:%M:%S')})")
                                continue

                    # Tăng số lần gửi
                    send_count += 1

                    # Tạo message với số lần gửi (trừ lần đầu)
                    if send_count == 1:
                        message = (
                            f"🔔 Cảnh báo phiếu Hỏng lại trong tháng:\n"
                            f"  - Địa bàn: {nvkt}\n"
                            f"  - Mã TB: {ma_tb}\n"
                            f"  - Tên TB: {ten_tb}\n"
                            f"  - Lần hỏng trong tháng: {int(lan_hong)}\n"
                            f"  - Thời gian tồn: {thoi_gian_ton:.1f} giờ\n"
                            f"  - Số lần Không hài lòng: {int(so_lan_khl)}"
                        )
                    else:
                        message = (
                            f"🔔 Cảnh báo phiếu Hỏng lại trong tháng {{Lần {send_count}}}:\n"
                            f"  - Địa bàn: {nvkt}\n"
                            f"  - Mã TB: {ma_tb}\n"
                            f"  - Tên TB: {ten_tb}\n"
                            f"  - Lần hỏng trong tháng: {int(lan_hong)}\n"
                            f"  - Thời gian tồn: {thoi_gian_ton:.1f} giờ\n"
                            f"  - Số lần Không hài lòng: {int(so_lan_khl)}"
                        )

                    # Gửi Zalo
                    zalo_success = False
                    try:
                        data = {
                            'threadID': thread_id,
                            'message': message
                        }
                        response = requests.get(webhook_url_zalo, json=data, timeout=10)

                        if response.status_code == 200:
                            sent_zalo += 1
                            zalo_success = True
                            print(f"    ✅ [Zalo] Gửi thành công: {ma_tb} (Lần {send_count})")
                        else:
                            failed_zalo += 1
                            print(f"    ❌ [Zalo] Gửi thất bại: {ma_tb} (Status: {response.status_code})")

                    except Exception as e:
                        failed_zalo += 1
                        print(f"    ❌ [Zalo] Lỗi khi gửi {ma_tb}: {e}")

                    # Tạm dừng 0.5 giây giữa Zalo và Telegram
                    time.sleep(0.5)

                    # Gửi Telegram
                    telegram_success = False
                    if telegram_chat_id and TELEGRAM_TOKEN:
                        try:
                            url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
                            data = {
                                "chat_id": telegram_chat_id,
                                "text": message,
                                "parse_mode": "HTML"
                            }
                            response = telegram_requests.post(url, data=data, timeout=10)

                            if response.status_code == 200:
                                sent_telegram += 1
                                telegram_success = True
                                print(f"    ✅ [Telegram] Gửi thành công: {ma_tb} (Lần {send_count})")
                            else:
                                failed_telegram += 1
                                print(f"    ❌ [Telegram] Gửi thất bại: {ma_tb} (Status: {response.status_code})")

                        except Exception as e:
                            failed_telegram += 1
                            print(f"    ❌ [Telegram] Lỗi khi gửi {ma_tb}: {e}")

                    # Ghi log nếu gửi thành công (ít nhất 1 nền tảng)
                    if zalo_success or telegram_success:
                        log_entry = {
                            'Thời gian gửi': current_time,
                            'ID báo hỏng': str(id_bh),
                            'Mã thuê bao': str(ma_tb),
                            'Tên thuê bao': str(ten_tb),
                            'Địa bàn': str(doi_vt),
                            'NVKT': str(nvkt),
                            'Số lần gửi': send_count,
                            'Trạng thái Zalo': 'Thành công' if zalo_success else 'Thất bại',
                            'Trạng thái Telegram': 'Thành công' if telegram_success else 'Thất bại'
                        }
                        log_data.append(log_entry)

                    # Tạm dừng 1 giây giữa các phiếu để tránh spam
                    time.sleep(1)

                total_sent_zalo += sent_zalo
                total_failed_zalo += failed_zalo
                total_sent_telegram += sent_telegram
                total_failed_telegram += failed_telegram

                print(f"  📊 {display_name}: Zalo ✅ {sent_zalo} | ❌ {failed_zalo} | Telegram ✅ {sent_telegram} | ❌ {failed_telegram} | ⏭️ Bỏ qua: {skipped}")

            except Exception as e:
                print(f"  ❌ Lỗi khi xử lý {doi_vt}: {e}")
                continue

        # Lưu log
        if log_data:
            log_df = pd.DataFrame(log_data)
            if not existing_log.empty:
                # Kết hợp log cũ và mới
                log_df = pd.concat([existing_log, log_df], ignore_index=True)

            # Lưu file
            log_df.to_excel(log_file, index=False)
            print(f"\n💾 Đã lưu log: {len(log_data)} bản ghi mới")

        # Tổng kết
        print("\n" + "=" * 60)
        print("📋 TỔNG KẾT:")
        print(f"  [Zalo]     ✅ Gửi thành công: {total_sent_zalo} bản tin")
        print(f"  [Zalo]     ❌ Gửi thất bại: {total_failed_zalo} bản tin")
        print(f"  [Telegram] ✅ Gửi thành công: {total_sent_telegram} bản tin")
        print(f"  [Telegram] ❌ Gửi thất bại: {total_failed_telegram} bản tin")
        print(f"  ⏭️ Bỏ qua (đã gửi trong 4 tiếng): {total_skipped} bản tin")
        print(f"  📊 Tổng cộng: {total_sent_zalo + total_failed_zalo + total_skipped} bản tin")
        print("=" * 60)

        return True

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file {file_path}")
        return False
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_warning_phieu_qua_gio():
    """
    Gửi cảnh báo phiếu quá giờ BRCD tới 3 nhóm Zalo.
    Đọc từ file chiTietBrcd5Doi.xlsx, từ 3 sheet: sontay, suoihai, quangoai
    Gửi cảnh báo cho các phiếu quá giờ (giờ còn lại thực < 0).

    Logic nhắc:
    - "Nhắc phiếu tồn quá {x} ngày" (với x là số ngày tồn làm tròn)
    - "Lần nhắc thứ {n}" (với n là lần thứ 2 trở lên)
    - Có cơ chế log để tránh gửi trùng lặp trong vòng 6 tiếng
    - Đếm số lần đã nhắc và hiển thị

    Chỉ gửi trong khung giờ 06:30 - 21:00
    """
    print("\n" + "=" * 70)
    print("🔔 KIỂM TRA GỬI CẢNH BÁO PHIẾU QUÁ GIỜ")
    print("=" * 70)

    # Kiểm tra thời gian cho phép gửi
    allowed, time_msg = is_allowed_send_time()
    print(time_msg)

    if not allowed:
        print("⏭️  BỎ QUA gửi cảnh báo (ngoài khung giờ cho phép)")
        print("=" * 70 + "\n")
        return

    file_path = os.path.join('chiaTheoDoi', 'chiTietBrcd5Doi.xlsx')
    webhook_url = WEBHOOK_TEXT_URL

    # Cấu hình log
    log_dir = "log_message"
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "phieu_qua_gio_log.xlsx")

    # Thời gian chờ giữa các lần gửi (6 tiếng = 360 phút)
    TIME_THRESHOLD = pd.Timedelta(hours=6)

    # Auto-generate mappings from team_config (BRCD - 4 teams)
    brcd_teams = get_active_teams('BRCD')

    sheet_to_thread_id = {
        team.id: LOCATION_THREAD_MAPPING[team.id]
        for team in brcd_teams
    }

    sheet_to_display_name = {
        team.id: team.short_name
        for team in brcd_teams
    }

    try:
        print(f"\n🚀 Bắt đầu gửi cảnh báo phiếu quá giờ...")
        current_time = datetime.now()

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Load existing log
        existing_log = pd.DataFrame()
        if os.path.exists(log_file):
            try:
                existing_log = pd.read_excel(log_file)
                existing_log['Thời gian gửi'] = pd.to_datetime(existing_log['Thời gian gửi'], errors='coerce')
                print(f"  📂 Đã load log từ file: {len(existing_log)} bản ghi")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi đọc log file: {e}")
                existing_log = pd.DataFrame()

        # Tổng số bản tin đã gửi
        total_sent = 0
        total_failed = 0
        total_skipped = 0

        # Danh sách log mới
        log_data = []

        # Đọc từng sheet
        for sheet_name, thread_id in sheet_to_thread_id.items():
            try:
                print(f"\n📋 Đang xử lý sheet: {sheet_name}")
                display_name = sheet_to_display_name[sheet_name]

                # Đọc sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                if df.empty:
                    print(f"  ℹ️ Sheet {sheet_name} không có dữ liệu")
                    continue

                # Kiểm tra các cột cần thiết
                required_cols = ['thời gian tồn thực', 'giờ còn lại thực', 'ma_tb', 'ngay_bh', 'ds_nhanvien_th', 'Trạng thái cổng']
                missing_cols = [col for col in required_cols if col not in df.columns]

                if missing_cols:
                    print(f"  ⚠️ Sheet {sheet_name} thiếu các cột: {', '.join(missing_cols)}")
                    continue

                # Kiểm tra xem có cột ghichuton không (không bắt buộc, dùng để hiển thị lý do tồn)
                has_ghichuton = 'ghichuton' in df.columns

                # Tính số ngày tồn
                df['so_ngay_ton'] = df['thời gian tồn thực'] / 24

                # Lọc phiếu quá giờ (giờ còn lại thực < 0)
                df_filtered = df[df['giờ còn lại thực'] < 0].copy()

                if df_filtered.empty:
                    print(f"  ℹ️ Không có phiếu quá giờ trong sheet {sheet_name}")
                    continue

                print(f"  🔍 Tìm thấy {len(df_filtered)} phiếu quá giờ")

                # Chuyển đổi ngày báo hỏng sang định dạng dễ đọc
                if 'ngay_bh' in df_filtered.columns:
                    df_filtered['ngay_bh_formatted'] = pd.to_datetime(
                        df_filtered['ngay_bh'],
                        errors='coerce'
                    ).dt.strftime('%d/%m/%Y %H:%M')
                else:
                    df_filtered['ngay_bh_formatted'] = 'N/A'

                sent_count = 0
                failed_count = 0
                skipped = 0

                # Gửi từng phiếu
                for idx, row in df_filtered.iterrows():
                    # Lấy thông tin từ row
                    ma_tb = row.get('ma_tb', 'N/A')
                    nvkt = row.get('ds_nhanvien_th', 'N/A')
                    ngay_bh = row.get('ngay_bh_formatted', 'N/A')
                    so_ngay_ton = row['so_ngay_ton']
                    gio_con_lai = row.get('giờ còn lại thực', 0)
                    trang_thai_cong = row.get('Trạng thái cổng', 'N/A')

                    # Lấy lý do tồn (nếu có cột ghichuton)
                    if has_ghichuton:
                        ly_do_ton = row.get('ghichuton', 'N/A')
                    else:
                        ly_do_ton = 'N/A'

                    # Xử lý giá trị null
                    if pd.isna(ma_tb):
                        ma_tb = 'N/A'
                    if pd.isna(nvkt):
                        nvkt = 'N/A'
                    if pd.isna(trang_thai_cong):
                        trang_thai_cong = 'N/A'
                    if pd.isna(ly_do_ton):
                        ly_do_ton = 'N/A'

                    # Làm tròn số ngày tồn
                    so_ngay_ton_rounded = int(round(so_ngay_ton))

                    # Kiểm tra log - đã gửi trong vòng 6 tiếng chưa?
                    send_count = 0
                    should_skip = False

                    if not existing_log.empty:
                        # Tìm các bản ghi matching trong log
                        matching_records = existing_log[
                            (existing_log['Mã thuê bao'] == str(ma_tb)) &
                            (existing_log['Địa bàn'] == sheet_name)
                        ]

                        if not matching_records.empty:
                            # Lấy bản ghi gần nhất
                            most_recent = matching_records['Thời gian gửi'].max()
                            time_since_last = current_time - most_recent

                            # Lấy số lần đã gửi
                            send_count = matching_records['Số lần gửi'].max()

                            # Kiểm tra nếu trong vòng 6 tiếng
                            if time_since_last < TIME_THRESHOLD:
                                should_skip = True
                                skipped += 1
                                total_skipped += 1
                                print(f"    ⏭️ Bỏ qua {ma_tb} (đã gửi {send_count} lần, lần cuối: {most_recent.strftime('%Y-%m-%d %H:%M:%S')})")
                                continue

                    # Tăng số lần gửi
                    send_count += 1

                    # Tạo message với số lần gửi
                    if send_count == 1:
                        message = (
                            f"🔔 Nhắc phiếu tồn quá {so_ngay_ton_rounded} ngày:\n"
                            f"  - Mã TB: {ma_tb}\n"
                            f"  - Ngày báo: {ngay_bh}\n"
                            f"  - NVKT: {nvkt}\n"
                            f"  - Trạng thái cổng: {trang_thai_cong}\n"
                            f"  - Lý do tồn: {ly_do_ton}\n"
                            f"  - Giờ quá hạn: {abs(gio_con_lai):.1f} giờ"
                        )
                    else:
                        message = (
                            f"🔔 Nhắc phiếu tồn quá {so_ngay_ton_rounded} ngày:\n"
                            f"Lần nhắc thứ {send_count}\n"
                            f"  - Mã TB: {ma_tb}\n"
                            f"  - Ngày báo: {ngay_bh}\n"
                            f"  - NVKT: {nvkt}\n"
                            f"  - Trạng thái cổng: {trang_thai_cong}\n"
                            f"  - Lý do tồn: {ly_do_ton}\n"
                            f"  - Giờ quá hạn: {abs(gio_con_lai):.1f} giờ"
                        )

                    # Gửi qua webhook
                    success = False
                    try:
                        data = {
                            'threadID': thread_id,
                            'message': message
                        }
                        response = requests.get(webhook_url, json=data, timeout=10)

                        if response.status_code == 200:
                            sent_count += 1
                            success = True
                            print(f"    ✅ Gửi thành công: {ma_tb} (Lần {send_count})")
                        else:
                            failed_count += 1
                            print(f"    ❌ Gửi thất bại: {ma_tb} (Status: {response.status_code})")

                    except Exception as e:
                        failed_count += 1
                        print(f"    ❌ Lỗi khi gửi {ma_tb}: {e}")

                    # Ghi log nếu gửi thành công
                    if success:
                        log_entry = {
                            'Thời gian gửi': current_time,
                            'Mã thuê bao': str(ma_tb),
                            'Địa bàn': sheet_name,
                            'NVKT': str(nvkt),
                            'Số ngày tồn': so_ngay_ton_rounded,
                            'Số lần gửi': send_count,
                            'Trạng thái': 'Thành công'
                        }
                        log_data.append(log_entry)

                    # Tạm dừng 1 giây giữa các phiếu để tránh spam
                    time.sleep(1)

                total_sent += sent_count
                total_failed += failed_count

                print(f"  📊 {display_name}: ✅ {sent_count} thành công | ❌ {failed_count} thất bại | ⏭️ Bỏ qua: {skipped}")

            except Exception as e:
                print(f"  ❌ Lỗi khi xử lý sheet {sheet_name}: {e}")
                import traceback
                traceback.print_exc()
                continue

        # Lưu log
        if log_data:
            log_df = pd.DataFrame(log_data)
            if not existing_log.empty:
                # Kết hợp log cũ và mới
                log_df = pd.concat([existing_log, log_df], ignore_index=True)

            # Lưu file
            log_df.to_excel(log_file, index=False)
            print(f"\n💾 Đã lưu log: {len(log_data)} bản ghi mới")

        # Tổng kết
        print("\n" + "=" * 60)
        print("📋 TỔNG KẾT:")
        print(f"  ✅ Gửi thành công: {total_sent} bản tin")
        print(f"  ❌ Gửi thất bại: {total_failed} bản tin")
        print(f"  ⏭️ Bỏ qua (đã gửi trong 6 tiếng): {total_skipped} bản tin")
        print(f"  📊 Tổng cộng: {total_sent + total_failed + total_skipped} bản tin")
        print("=" * 60)

        return True

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file {file_path}")
        return False
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_warning_khdn_uu_tien():
    """
    Gửi cảnh báo phiếu KHDN ưu tiên tới 3 nhóm Zalo và Telegram.
    Đọc từ file bc_BRCD.xlsx, sheet chi_tiet_ton_brcd.
    Gửi cảnh báo NGAY LẬP TỨC cho các phiếu có LOAIHINH_TB thuộc danh sách ưu tiên.
    KHÔNG GHI LOG - gửi mọi lần chạy khi gặp.

    Chỉ gửi trong khung giờ 06:30 - 21:00

    Danh sách loại trừ: Các mã TB trong KHDN_EXCLUSION_LIST sẽ KHÔNG gửi cảnh báo
    """
    print("\n" + "=" * 70)
    print("🔔 KIỂM TRA GỬI CẢNH BÁO KHDN ƯU TIÊN")
    print("=" * 70)

    # Lấy danh sách loại trừ
    exclusion_list = get_khdn_exclusion_list()
    if exclusion_list:
        print(f"⛔ Danh sách loại trừ: {', '.join(sorted(exclusion_list))}")
        print(f"   ({len(exclusion_list)} mã TB sẽ được bỏ qua)")
    else:
        print("ℹ️ Không có danh sách loại trừ")

    # Kiểm tra thời gian cho phép gửi
    allowed, time_msg = is_allowed_send_time()
    print(time_msg)

    if not allowed:
        print("⏭️  BỎ QUA gửi cảnh báo (ngoài khung giờ cho phép)")
        print("=" * 70 + "\n")
        return

    # Import send_tele để gửi telegram
    try:
        from send_tele import (
            ID_CHAT_SONTAY, ID_CHAT_SUOIHAI, ID_CHAT_QUANGOAI,
            TELEGRAM_TOKEN, send_message as send_telegram_message
        )
        import requests as telegram_requests
    except ImportError:
        print("⚠️ Không thể import module send_tele. Chỉ gửi Zalo.")
        ID_CHAT_SONTAY = None
        ID_CHAT_SUOIHAI = None
        ID_CHAT_QUANGOAI = None
        TELEGRAM_TOKEN = None

    file_path = os.path.join('downloads', 'kq_dhsc', 'bc_BRCD.xlsx')
    webhook_url_zalo = WEBHOOK_TEXT_URL

    # Danh sách loại hình dịch vụ ưu tiên
    PRIORITY_SERVICES = [
        'Cáp quang trắng',
        'MetroNet FE',
        'Megawan quang FE',
        'Leasedline E1',
        'Megawan quang GE',
        'Leasedline nx64k',
        'Megawan_NNI',
        'MetroNet GE'
    ]

    # Custom thread ID mapping for KHDN ưu tiên warnings (different from global config)
    # These thread IDs are specific for KHDN ưu tiên messages
    CUSTOM_THREAD_IDS = {
        'ToKT_PhucTho': '6780971089121842303',
        'ToKT_SonTay': '6337217534995887511',
        'ToKT_QuangOai': '5364152493553904404',
        'ToKT_SuoiHai': '6085297980620830486',
    }

    brcd_teams = get_active_teams('BRCD')

    doi_vt_mapping = {
        team.id: {
            'thread_id': CUSTOM_THREAD_IDS.get(team.id, LOCATION_THREAD_MAPPING.get(team.id)),
            'telegram_chat_id': LOCATION_CHAT_MAPPING.get(team.id),
            'display_name': team.short_name
        }
        for team in brcd_teams
    }

    try:
        print(f"\n🚀 Bắt đầu kiểm tra phiếu KHDN ưu tiên...")
        print(f"📋 Exclusion list inside send_warning_khdn_uu_tien: {exclusion_list}")
        current_time = datetime.now()

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Đọc sheet chi_tiet_ton_brcd
        df = pd.read_excel(file_path, sheet_name='chi_tiet_ton_brcd')

        if df.empty:
            print(f"  ℹ️ Sheet chi_tiet_ton_brcd không có dữ liệu")
            return True

        # Kiểm tra cột LOAIHINH_TB có tồn tại không
        if 'LOAIHINH_TB' not in df.columns:
            print(f"  ⚠️ Sheet chi_tiet_ton_brcd không có cột LOAIHINH_TB")
            return False

        # Kiểm tra cột ma_tb có tồn tại không
        if 'ma_tb' not in df.columns:
            print(f"  ⚠️ Sheet chi_tiet_ton_brcd không có cột ma_tb")
            print(f"     Các cột có sẵn: {list(df.columns)}")
            return False

        # Lọc phiếu có LOAIHINH_TB thuộc danh sách ưu tiên
        df_filtered = df[df['LOAIHINH_TB'].isin(PRIORITY_SERVICES)].copy()

        if df_filtered.empty:
            print(f"  ℹ️ Không có phiếu KHDN ưu tiên nào")
            return True

        print(f"  🔍 Tìm thấy {len(df_filtered)} phiếu KHDN ưu tiên")
        print(f"  📋 Các loại hình phát hiện:")
        for loaihinh, count in df_filtered['LOAIHINH_TB'].value_counts().items():
            print(f"     - {loaihinh}: {count} phiếu")

        # Tổng số bản tin đã gửi
        total_sent_zalo = 0
        total_failed_zalo = 0
        total_sent_telegram = 0
        total_failed_telegram = 0

        # Nhóm theo DOI_VT để gửi từng nhóm
        grouped = df_filtered.groupby('DOI_VT')

        for doi_vt, group in grouped:
            try:
                print(f"\n📋 Đang xử lý {doi_vt}: {len(group)} phiếu")

                # Lấy mapping info
                mapping_info = doi_vt_mapping.get(doi_vt)
                if not mapping_info:
                    print(f"  ⚠️ Không tìm thấy mapping cho {doi_vt}, bỏ qua")
                    continue

                thread_id = mapping_info['thread_id']
                telegram_chat_id = mapping_info['telegram_chat_id']
                display_name = mapping_info['display_name']

                print(f"  📤 Thread ID (Zalo): {thread_id}")
                print(f"  📤 Chat ID (Telegram): {telegram_chat_id}")

                # Gửi từng phiếu
                sent_zalo = 0
                failed_zalo = 0
                sent_telegram = 0
                failed_telegram = 0

                for idx, row in group.iterrows():
                    # Lấy thông tin từ row
                    nvkt = row.get('NVKT', 'N/A')
                    ma_tb = row.get('ma_tb', 'N/A')
                    loaihinh_tb = row.get('LOAIHINH_TB', 'N/A')
                    thoi_gian_ton = row.get('thời gian tồn thực', 0)
                    so_lan_khl = row.get('Số lần KHL', 0)
                    lan_hong = row.get('LAN_HONG', 0)

                    # Debug: In ra mã TB và kiểm tra loại trừ
                    debug_enabled = os.getenv('DEBUG_EXCLUSION_LIST', 'false').lower() == 'true'
                    if debug_enabled:
                        print(f"    [DEBUG] Checking ma_tb='{ma_tb}', exclusion_list={exclusion_list}, is_excluded={ma_tb in exclusion_list}")

                    # Kiểm tra xem mã TB có trong danh sách loại trừ không
                    if ma_tb in exclusion_list:
                        print(f"    ⛔ Bỏ qua: {ma_tb} ({loaihinh_tb}) - Nằm trong danh sách loại trừ")
                        continue

                    # Lấy tên thuê bao nếu có
                    ten_tb = row.get('TEN_TB', row.get('TEN_KH', 'N/A'))

                    # Xử lý giá trị null
                    if pd.isna(nvkt):
                        nvkt = 'N/A'
                    if pd.isna(ma_tb):
                        ma_tb = 'N/A'
                    if pd.isna(ten_tb):
                        ten_tb = 'N/A'
                    if pd.isna(loaihinh_tb):
                        loaihinh_tb = 'N/A'
                    if pd.isna(thoi_gian_ton):
                        thoi_gian_ton = 0
                    if pd.isna(so_lan_khl):
                        so_lan_khl = 0
                    if pd.isna(lan_hong):
                        lan_hong = 0

                    # Tạo message cho phiếu KHDN ưu tiên
                    message = (
                        f"🔴🔴 Phiếu KHDN ưu tiên 🔴🔴\n"
                        f"  - Địa bàn: {nvkt}\n"
                        f"  - Mã TB: {ma_tb}\n"
                        f"  - Tên TB: {ten_tb}\n"
                        f"  - Loại hình: {loaihinh_tb}\n"
                        f"  - Thời gian tồn: {thoi_gian_ton:.1f} giờ\n"
                        f"  - Lần hỏng trong tháng: {int(lan_hong)}\n"
                        f"  - Số lần Không hài lòng: {int(so_lan_khl)}"
                    )

                    # Gửi Zalo
                    zalo_success = False
                    try:
                        data = {
                            'threadID': thread_id,
                            'message': message
                        }
                        print(f"    📤 Gọi webhook Zalo cho {ma_tb}...")
                        response = requests.get(webhook_url_zalo, json=data, timeout=10)

                        if response.status_code == 200:
                            sent_zalo += 1
                            zalo_success = True
                            print(f"    ✅ [Zalo] Gửi thành công: {ma_tb} ({loaihinh_tb})")
                        else:
                            failed_zalo += 1
                            print(f"    ❌ [Zalo] Gửi thất bại: {ma_tb} (Status: {response.status_code})")

                    except Exception as e:
                        failed_zalo += 1
                        print(f"    ❌ [Zalo] Lỗi khi gửi {ma_tb}: {e}")

                    # Tạm dừng 0.5 giây giữa Zalo và Telegram
                    time.sleep(0.5)

                    # Gửi Telegram
                    telegram_success = False
                    if telegram_chat_id and TELEGRAM_TOKEN:
                        try:
                            url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
                            data = {
                                "chat_id": telegram_chat_id,
                                "text": message,
                                "parse_mode": "HTML"
                            }
                            response = telegram_requests.post(url, data=data, timeout=10)

                            if response.status_code == 200:
                                sent_telegram += 1
                                telegram_success = True
                                print(f"    ✅ [Telegram] Gửi thành công: {ma_tb} ({loaihinh_tb})")
                            else:
                                failed_telegram += 1
                                print(f"    ❌ [Telegram] Gửi thất bại: {ma_tb} (Status: {response.status_code})")

                        except Exception as e:
                            failed_telegram += 1
                            print(f"    ❌ [Telegram] Lỗi khi gửi {ma_tb}: {e}")

                    # Tạm dừng 1 giây giữa các phiếu để tránh spam
                    time.sleep(1)

                total_sent_zalo += sent_zalo
                total_failed_zalo += failed_zalo
                total_sent_telegram += sent_telegram
                total_failed_telegram += failed_telegram

                print(f"  📊 {display_name}: Zalo ✅ {sent_zalo} | ❌ {failed_zalo} | Telegram ✅ {sent_telegram} | ❌ {failed_telegram}")

            except Exception as e:
                print(f"  ❌ Lỗi khi xử lý {doi_vt}: {e}")
                continue

        # Tổng kết
        print("\n" + "=" * 60)
        print("📋 TỔNG KẾT KHDN ƯU TIÊN:")
        print(f"  [Zalo]     ✅ Gửi thành công: {total_sent_zalo} bản tin")
        print(f"  [Zalo]     ❌ Gửi thất bại: {total_failed_zalo} bản tin")
        print(f"  [Telegram] ✅ Gửi thành công: {total_sent_telegram} bản tin")
        print(f"  [Telegram] ❌ Gửi thất bại: {total_failed_telegram} bản tin")
        print(f"  📊 Tổng cộng: {total_sent_zalo + total_failed_zalo} bản tin")
        print("=" * 60)

        return True

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file {file_path}")
        return False
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_warning_mat_huong_sa(thread_id: str, message: str) -> bool:
    """
    Gửi cảnh báo mất hướng SA đến nhóm Zalo và Telegram

    Args:
        thread_id: Thread ID của nhóm Zalo cần gửi
        message: Nội dung cảnh báo (format: "Cảnh báo sự cố mất hướng: {SA}, mất {off_count}/{total_count}")

    Returns:
        bool: True nếu gửi thành công, False nếu thất bại
    """
    print("\n" + "=" * 70)
    print("🔔 GỬI CẢNH BÁO MẤT HƯỚNG SA")
    print("=" * 70)
    print(f"Thread ID: {thread_id}")
    print(f"Message: {message}")

    # Webhook URL
    webhook_url_zalo = WEBHOOK_TEXT_URL

    success_zalo = False
    success_telegram = False

    # 1. Gửi Zalo
    try:
        print("\n📱 Gửi cảnh báo qua Zalo...")
        data = {
            'threadID': thread_id,
            'message': message
        }
        response = requests.get(webhook_url_zalo, json=data, timeout=10)

        if response.status_code == 200:
            print(f"  ✅ Gửi Zalo thành công (status: {response.status_code})")
            success_zalo = True
        else:
            print(f"  ❌ Gửi Zalo thất bại (status: {response.status_code})")
            print(f"  Response: {response.text}")

    except Exception as e:
        print(f"  ❌ Lỗi khi gửi Zalo: {e}")

    # 2. Gửi Telegram
    try:
        print("\n📲 Gửi cảnh báo qua Telegram...")

        # Import telegram config
        try:
            from send_tele import TELEGRAM_TOKEN

            # Auto-generate thread_id to telegram chat_id mapping (BRCD - 4 teams)
            brcd_teams = get_active_teams('BRCD')
            thread_to_telegram = {
                LOCATION_THREAD_MAPPING[team.id]: LOCATION_CHAT_MAPPING.get(team.id)
                for team in brcd_teams
            }

            telegram_chat_id = thread_to_telegram.get(thread_id)

            if telegram_chat_id and TELEGRAM_TOKEN:
                telegram_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
                telegram_data = {
                    'chat_id': telegram_chat_id,
                    'text': message,
                    'parse_mode': 'HTML'
                }
                telegram_response = requests.post(telegram_url, json=telegram_data, timeout=10)

                if telegram_response.status_code == 200:
                    print(f"  ✅ Gửi Telegram thành công (chat_id: {telegram_chat_id})")
                    success_telegram = True
                else:
                    print(f"  ❌ Gửi Telegram thất bại (status: {telegram_response.status_code})")
            else:
                print(f"  ⚠️  Không tìm thấy Telegram chat_id cho thread_id: {thread_id}")

        except ImportError:
            print("  ⚠️  Không thể import module send_tele. Bỏ qua gửi Telegram.")

    except Exception as e:
        print(f"  ❌ Lỗi khi gửi Telegram: {e}")

    # Tổng kết
    print("\n" + "=" * 60)
    if success_zalo or success_telegram:
        print("✅ Gửi cảnh báo mất hướng SA thành công")
        print(f"  [Zalo]     {'✅ Thành công' if success_zalo else '❌ Thất bại'}")
        print(f"  [Telegram] {'✅ Thành công' if success_telegram else '❌ Thất bại'}")
    else:
        print("❌ Gửi cảnh báo mất hướng SA thất bại trên cả 2 kênh")
    print("=" * 60)

    return success_zalo or success_telegram

def send_warning_phieu_ob_khl():
    """
    Gửi cảnh báo phiếu OB (Observation) KHL tới 4 nhóm Zalo
    Đọc từ file bc_BRCD.xlsx sheet chi_tiet_ton_brcd
    Gửi cảnh báo khi cột GHICHU_HONG = "OBTT Dieu lai phieu tu khao sat bao hong"
    Ghi log để mỗi bản ghi chỉ gửi 1 lần duy nhất
    """
    file_path = os.path.join('downloads', 'kq_dhsc', 'bc_BRCD.xlsx')
    webhook_url = WEBHOOK_TEXT_URL
    log_dir = "log_message"
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "ob_khl_sent.xlsx")

    # Custom thread ID mapping for OB KHL warnings (different from global config)
    # These thread IDs are specific for OB KHL messages
    doi_vt_to_thread_id = {
        'ToKT_PhucTho': '6780971089121842303',
        'ToKT_SonTay': '6337217534995887511',
        'ToKT_QuangOai': '5364152493553904404',
        'ToKT_SuoiHai': '6085297980620830486',
    }

    # Log the available teams
    print(f"  📍 Khả dụng {len(doi_vt_to_thread_id)} nhóm Zalo: {', '.join(doi_vt_to_thread_id.keys())}")

    try:
        print(f"\n🚀 Bắt đầu gửi cảnh báo phiếu OB KHL...")

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Load existing log (danh sách các baohong_id đã gửi)
        sent_baohong_ids = set()
        if os.path.exists(log_file):
            try:
                df_log = pd.read_excel(log_file)
                sent_baohong_ids = set(df_log['baohong_id'].astype(str).unique())
                print(f"  📋 Đã load {len(sent_baohong_ids)} bản ghi đã gửi từ log")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi đọc log file: {e}")

        # Đọc sheet chi_tiet_ton_brcd
        print(f"  📖 Đang đọc sheet chi_tiet_ton_brcd từ {file_path}...")
        try:
            df = pd.read_excel(file_path, sheet_name='chi_tiet_ton_brcd')
        except Exception as e:
            print(f"  ❌ Lỗi khi đọc file: {e}")
            return False

        if df.empty:
            print(f"  ℹ️ Sheet chi_tiet_ton_brcd không có dữ liệu")
            return False

        # Kiểm tra các cột cần thiết
        required_cols = ['baohong_id', 'GHICHU_HONG', 'NVKT', 'ma_tb', 'TEN_TB', 'LOAIHINH_TB', 'thời gian tồn thực', 'DOI_VT']
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            print(f"  ❌ Sheet chi_tiet_ton_brcd thiếu các cột: {', '.join(missing_cols)}")
            return False

        # Lọc phiếu OB: GHICHU_HONG chứa "OBTT Dieu lai phieu tu khao sat bao hong"
        # Sử dụng contains thay vì exact match để xử lý khoảng trắng và ký tự phụ
        df_ob = df[df['GHICHU_HONG'].fillna('').str.contains('OBTT Dieu lai phieu tu khao sat bao hong', case=False, na=False)].copy()

        if df_ob.empty:
            print(f"  ℹ️ Không tìm thấy phiếu OB nào (GHICHU_HONG chứa 'OBTT Dieu lai phieu tu khao sat bao hong')")
            return False

        print(f"  🔍 Tìm thấy {len(df_ob)} phiếu OB")

        # Lọc các phiếu chưa gửi
        df_ob['baohong_id_str'] = df_ob['baohong_id'].astype(str)
        df_ob_new = df_ob[~df_ob['baohong_id_str'].isin(sent_baohong_ids)].copy()

        if df_ob_new.empty:
            print(f"  ℹ️ Tất cả {len(df_ob)} phiếu OB đã được gửi rồi")
            return False

        print(f"  📤 Có {len(df_ob_new)} phiếu OB mới chưa được gửi")

        # Chuẩn bị log data để ghi vào file
        log_data = []

        # Gửi từng bản tin
        sent_count = 0
        failed_count = 0

        for _, row in df_ob_new.iterrows():
            try:
                baohong_id = str(row.get('baohong_id', ''))
                nvkt = row.get('NVKT', '')
                ma_tb = row.get('ma_tb', 'N/A')
                ten_tb = row.get('TEN_TB', 'N/A')
                loaihinh_tb = row.get('LOAIHINH_TB', 'N/A')
                thoi_gian_ton = float(row.get('thời gian tồn thực', 0))
                doi_vt = row.get('DOI_VT', 'N/A')
                ngay_bh = row.get('ngay_bh', '')
                
                # Handle null/NaN values
                if pd.isna(nvkt) or nvkt == '':
                    nvkt = 'N/A'
                if pd.isna(ten_tb):
                    ten_tb = 'N/A'
                if pd.isna(loaihinh_tb):
                    loaihinh_tb = 'N/A'
                    
                # Format ngay_bh to readable date
                ngay_bh_formatted = 'N/A'
                if not pd.isna(ngay_bh) and ngay_bh != '':
                    try:
                        ngay_bh_dt = pd.to_datetime(ngay_bh, errors='coerce')
                        if not pd.isna(ngay_bh_dt):
                            ngay_bh_formatted = ngay_bh_dt.strftime('%d/%m/%Y %H:%M')
                    except:
                        ngay_bh_formatted = str(ngay_bh)

                # Tính số giờ
                thoi_gian_ton_hours = thoi_gian_ton if isinstance(thoi_gian_ton, (int, float)) else 0

                # Lấy thread_id dựa trên DOI_VT
                thread_id = doi_vt_to_thread_id.get(doi_vt)
                if not thread_id:
                    print(f"  ⚠️ Không tìm thấy thread_id cho DOI_VT: {doi_vt}")
                    failed_count += 1
                    continue

                # Tạo message với đầy đủ thông tin
                message = (
                    f"🔴 Phiếu OB KHL 🔴\n"
                    f"  - ID báo hỏng: {baohong_id}\n"
                    f"  - Thời gian báo: {ngay_bh_formatted}\n"
                    f"  - Địa bàn: {nvkt}\n"
                    f"  - Mã TB: {ma_tb}\n"
                    f"  - Tên TB: {ten_tb}\n"
                    f"  - Loại hình: {loaihinh_tb}\n"
                    f"  - Thời gian tồn: {thoi_gian_ton_hours:.1f} giờ"
                )

                # Gửi Zalo
                try:
                    data = {
                        'threadID': thread_id,
                        'message': message
                    }
                    response = requests.get(webhook_url, json=data, timeout=10)

                    if response.status_code == 200:
                        sent_count += 1
                        print(f"  ✅ Gửi OB {baohong_id} tới {doi_vt} thành công")

                        # Ghi log
                        log_data.append({
                            'baohong_id': baohong_id,
                            'ma_tb': ma_tb,
                            'DOI_VT': doi_vt,
                            'Thời gian gửi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'Trạng thái': 'Thành công'
                        })
                    else:
                        failed_count += 1
                        print(f"  ❌ Gửi OB {baohong_id} thất bại. Status: {response.status_code}")
                        log_data.append({
                            'baohong_id': baohong_id,
                            'ma_tb': ma_tb,
                            'DOI_VT': doi_vt,
                            'Thời gian gửi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'Trạng thái': f'Thất bại (Status: {response.status_code})'
                        })

                except Exception as e:
                    failed_count += 1
                    print(f"  ❌ Lỗi khi gửi OB {baohong_id}: {e}")
                    log_data.append({
                        'baohong_id': baohong_id,
                        'ma_tb': ma_tb,
                        'DOI_VT': doi_vt,
                        'Thời gian gửi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'Trạng thái': f'Lỗi: {str(e)}'
                    })

                # Delay 1 giây giữa các bản tin
                time.sleep(1)

            except Exception as e:
                failed_count += 1
                print(f"  ❌ Lỗi xử lý bản ghi: {e}")

        # Ghi log vào file
        if log_data:
            try:
                df_new_log = pd.DataFrame(log_data)

                if os.path.exists(log_file):
                    # Append vào file cũ
                    df_existing = pd.read_excel(log_file)
                    df_new_log = pd.concat([df_existing, df_new_log], ignore_index=True)

                df_new_log.to_excel(log_file, index=False)
                print(f"  💾 Đã ghi log vào {log_file}")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi ghi log: {e}")

        # Tổng kết
        print("\n" + "=" * 60)
        print(f"📊 Kết quả gửi cảnh báo OB KHL:")
        print(f"  ✅ Thành công: {sent_count}")
        print(f"  ❌ Thất bại: {failed_count}")
        print("=" * 60)

        return sent_count > 0

    except Exception as e:
        print(f"❌ Lỗi tổng quát: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_warning_phieu_hll_7_ngay():
    """
    Gửi cảnh báo phiếu Hỏng lại trong vòng 7 ngày.
    So sánh MA_TB từ phiếu đang tồn (bc_BRCD.xlsx) với toàn bộ phiếu báo hỏng trong tháng (SM4-C11.xlsx).
    Nếu cùng MA_TB có phiếu báo hỏng trong vòng 7 ngày trước thời điểm nhận phiếu hiện tại, gửi cảnh báo.
    Ghi log để mỗi bản ghi chỉ gửi 1 lần duy nhất.
    """
    file_brcd = os.path.join('downloads', 'kq_dhsc', 'bc_BRCD.xlsx')
    file_sm4 = '/home/vtst/baocaohanoi/downloads/baocao_hanoi/SM4-C11.xlsx'
    webhook_url = WEBHOOK_TEXT_URL
    log_dir = "log_message"
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "hll_7_ngay_sent.xlsx")

    # Thread ID mapping (giống send_warning_phieu_ob_khl)
    doi_vt_to_thread_id = {
        'ToKT_PhucTho': '6780971089121842303',
        'ToKT_SonTay': '6337217534995887511',
        'ToKT_QuangOai': '5364152493553904404',
        'ToKT_SuoiHai': '6085297980620830486',
    }

    print(f"  📍 Khả dụng {len(doi_vt_to_thread_id)} nhóm Zalo: {', '.join(doi_vt_to_thread_id.keys())}")

    try:
        print(f"\n🚀 Bắt đầu gửi cảnh báo phiếu HLL 7 ngày...")

        if not os.path.exists(file_brcd):
            print(f"❌ Không tìm thấy file: {file_brcd}")
            return False
        if not os.path.exists(file_sm4):
            print(f"❌ Không tìm thấy file SM4-C11: {file_sm4}")
            return False

        # Load log đã gửi
        sent_baohong_ids = set()
        if os.path.exists(log_file):
            try:
                df_log = pd.read_excel(log_file)
                sent_baohong_ids = set(df_log['baohong_id'].astype(str).unique())
                print(f"  📋 Đã load {len(sent_baohong_ids)} bản ghi đã gửi từ log")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi đọc log file: {e}")

        # Đọc phiếu đang tồn từ bc_BRCD
        print(f"  📖 Đang đọc sheet chi_tiet_ton_brcd từ {file_brcd}...")
        try:
            df_brcd = pd.read_excel(file_brcd, sheet_name='chi_tiet_ton_brcd')
        except Exception as e:
            print(f"  ❌ Lỗi khi đọc file bc_BRCD: {e}")
            return False

        if df_brcd.empty:
            print(f"  ℹ️ Sheet chi_tiet_ton_brcd không có dữ liệu")
            return False

        required_cols = ['baohong_id', 'ma_tb', 'TEN_TB', 'LOAIHINH_TB', 'NVKT', 'thời gian tồn thực', 'DOI_VT', 'ngay_bh']
        missing_cols = [col for col in required_cols if col not in df_brcd.columns]
        if missing_cols:
            print(f"  ❌ Sheet chi_tiet_ton_brcd thiếu cột: {', '.join(missing_cols)}")
            return False

        # Đọc lịch sử phiếu báo hỏng từ SM4-C11
        print(f"  📖 Đang đọc file SM4-C11...")
        try:
            df_sm4 = pd.read_excel(file_sm4)
        except Exception as e:
            print(f"  ❌ Lỗi khi đọc file SM4-C11: {e}")
            return False

        if df_sm4.empty:
            print(f"  ℹ️ File SM4-C11 không có dữ liệu")
            return False

        print(f"  📊 SM4-C11: {len(df_sm4)} phiếu lịch sử")

        # Parse ngày báo hỏng trong SM4-C11
        df_sm4['NGAY_BAO_HONG_DT'] = pd.to_datetime(
            df_sm4['NGAY_BAO_HONG'], format='%d/%m/%Y %H:%M:%S', errors='coerce'
        )

        # Tạo lookup dict: MA_TB (lowercase) -> list of {baohong_id, ngay_bao_hong}
        sm4_lookup = {}
        for _, row in df_sm4.iterrows():
            ma_tb_raw = str(row.get('MA_TB', '')).strip().lower()
            if ma_tb_raw and ma_tb_raw != 'nan':
                if ma_tb_raw not in sm4_lookup:
                    sm4_lookup[ma_tb_raw] = []
                sm4_lookup[ma_tb_raw].append({
                    'baohong_id': str(row.get('BAOHONG_ID', '')),
                    'ngay_bao_hong': row['NGAY_BAO_HONG_DT'],
                })

        print(f"  🔍 Đã tạo lookup cho {len(sm4_lookup)} MA_TB từ SM4-C11")

        # Xử lý từng phiếu đang tồn
        log_data = []
        sent_count = 0
        failed_count = 0
        hll_count = 0

        for _, row in df_brcd.iterrows():
            try:
                baohong_id = str(row.get('baohong_id', ''))

                # Bỏ qua nếu đã gửi
                if baohong_id in sent_baohong_ids:
                    continue

                ma_tb = str(row.get('ma_tb', '')).strip()
                ma_tb_norm = ma_tb.lower()

                # Parse ngày báo hỏng của phiếu hiện tại
                ngay_bh = row.get('ngay_bh', '')
                ngay_bh_dt = pd.to_datetime(ngay_bh, dayfirst=True, errors='coerce')
                if pd.isna(ngay_bh_dt):
                    continue

                # Tìm phiếu cùng MA_TB trong SM4-C11
                historical = sm4_lookup.get(ma_tb_norm, [])
                if not historical:
                    continue

                # Lọc phiếu trong vòng 7 ngày trước (không tính phiếu hiện tại)
                seven_days_ago = ngay_bh_dt - pd.Timedelta(days=7)
                previous_faults = []
                for h in historical:
                    if h['baohong_id'] != baohong_id and not pd.isna(h['ngay_bao_hong']):
                        if seven_days_ago <= h['ngay_bao_hong'] < ngay_bh_dt:
                            previous_faults.append(h)

                if not previous_faults:
                    continue

                hll_count += 1

                # Lấy thông tin phiếu
                nvkt = row.get('NVKT', 'N/A')
                ten_tb = row.get('TEN_TB', 'N/A')
                loaihinh_tb = row.get('LOAIHINH_TB', 'N/A')
                thoi_gian_ton = float(row.get('thời gian tồn thực', 0))
                doi_vt = row.get('DOI_VT', 'N/A')

                if pd.isna(nvkt) or nvkt == '':
                    nvkt = 'N/A'
                if pd.isna(ten_tb):
                    ten_tb = 'N/A'
                if pd.isna(loaihinh_tb):
                    loaihinh_tb = 'N/A'

                ngay_bh_formatted = ngay_bh_dt.strftime('%d/%m/%Y %H:%M')

                # Lịch sử hỏng trong 7 ngày (sắp xếp mới nhất trước)
                prev_dates = []
                for pf in sorted(previous_faults, key=lambda x: x['ngay_bao_hong'], reverse=True):
                    prev_dates.append(pf['ngay_bao_hong'].strftime('%d/%m/%Y %H:%M'))
                lich_su_hong = ', '.join(prev_dates)

                # Lấy thread_id
                thread_id = doi_vt_to_thread_id.get(doi_vt)
                if not thread_id:
                    print(f"  ⚠️ Không tìm thấy thread_id cho DOI_VT: {doi_vt}")
                    failed_count += 1
                    continue

                thoi_gian_ton_hours = thoi_gian_ton if isinstance(thoi_gian_ton, (int, float)) else 0

                message = (
                    f"🟠 Phiếu HLL trong vòng 7 ngày 🟠\n"
                    f"  - ID báo hỏng: {baohong_id}\n"
                    f"  - Thời gian báo: {ngay_bh_formatted}\n"
                    f"  - Địa bàn: {nvkt}\n"
                    f"  - Mã TB: {ma_tb}\n"
                    f"  - Tên TB: {ten_tb}\n"
                    f"  - Loại hình: {loaihinh_tb}\n"
                    f"  - Thời gian tồn: {thoi_gian_ton_hours:.1f} giờ\n"
                    f"  - Lịch sử hỏng 7 ngày: {lich_su_hong}"
                )

                # Gửi Zalo
                try:
                    data = {
                        'threadID': thread_id,
                        'message': message
                    }
                    response = requests.get(webhook_url, json=data, timeout=10)

                    if response.status_code == 200:
                        sent_count += 1
                        print(f"  ✅ Gửi HLL {baohong_id} ({ma_tb}) tới {doi_vt} thành công")

                        log_data.append({
                            'baohong_id': baohong_id,
                            'ma_tb': ma_tb,
                            'DOI_VT': doi_vt,
                            'lich_su_hong_7_ngay': lich_su_hong,
                            'Thời gian gửi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'Trạng thái': 'Thành công'
                        })
                    else:
                        failed_count += 1
                        print(f"  ❌ Gửi HLL {baohong_id} thất bại. Status: {response.status_code}")
                        log_data.append({
                            'baohong_id': baohong_id,
                            'ma_tb': ma_tb,
                            'DOI_VT': doi_vt,
                            'lich_su_hong_7_ngay': lich_su_hong,
                            'Thời gian gửi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'Trạng thái': f'Thất bại (Status: {response.status_code})'
                        })

                except Exception as e:
                    failed_count += 1
                    print(f"  ❌ Lỗi khi gửi HLL {baohong_id}: {e}")
                    log_data.append({
                        'baohong_id': baohong_id,
                        'ma_tb': ma_tb,
                        'DOI_VT': doi_vt,
                        'lich_su_hong_7_ngay': lich_su_hong,
                        'Thời gian gửi': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'Trạng thái': f'Lỗi: {str(e)}'
                    })

                # Delay 1 giây giữa các bản tin
                time.sleep(1)

            except Exception as e:
                failed_count += 1
                print(f"  ❌ Lỗi xử lý bản ghi: {e}")

        # Ghi log vào file
        if log_data:
            try:
                df_new_log = pd.DataFrame(log_data)

                if os.path.exists(log_file):
                    df_existing = pd.read_excel(log_file)
                    df_new_log = pd.concat([df_existing, df_new_log], ignore_index=True)

                df_new_log.to_excel(log_file, index=False)
                print(f"  💾 Đã ghi log vào {log_file}")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi ghi log: {e}")

        # Tổng kết
        print("\n" + "=" * 60)
        print(f"📊 Kết quả gửi cảnh báo HLL 7 ngày:")
        print(f"  🔍 Tổng phiếu HLL phát hiện: {hll_count}")
        print(f"  ✅ Thành công: {sent_count}")
        print(f"  ❌ Thất bại: {failed_count}")
        print("=" * 60)

        return sent_count > 0

    except Exception as e:
        print(f"❌ Lỗi tổng quát: {e}")
        import traceback
        traceback.print_exc()
        return False


def should_send_sap_qua_gio_warning(ma_tb: str, gio_con_lai: float, log_df: pd.DataFrame) -> tuple:
    """
    Determine if warning should be sent based on remaining time and last sent time.
    
    Throttle rules:
    - 1-2 hours remaining: Send every 30 minutes
    - < 1 hour remaining: Send every 15 minutes  
    - 10-30 minutes remaining: Send every 10 minutes
    - < 10 minutes remaining: Send continuously (no throttle)
    
    Args:
        ma_tb: Mã thuê bao
        gio_con_lai: Giờ còn lại (hours)
        log_df: DataFrame chứa log đã gửi
        
    Returns:
        (should_send: bool, reason: str)
    """
    from datetime import datetime
    
    # Calculate throttle interval based on remaining time
    if gio_con_lai >= 1.0:  # 1-2 hours
        throttle_minutes = 30
    elif gio_con_lai >= 10/60:  # 10min - 1hour (10/60 = 0.167 hours)
        throttle_minutes = 15
    elif gio_con_lai >= 10/60:  # This condition seems redundant - keeping for now
        throttle_minutes = 10
    else:  # < 10 minutes
        throttle_minutes = 0  # Send every time (no throttle)
    
    # Check if ma_tb exists in log
    if log_df.empty or ma_tb not in log_df['ma_tb'].values:
        return True, "Lần đầu gửi"
    
    # Get last sent time for this ma_tb
    last_record = log_df[log_df['ma_tb'] == ma_tb].iloc[-1]
    last_sent = pd.to_datetime(last_record['last_sent_time'])
    time_diff_minutes = (datetime.now() - last_sent).total_seconds() / 60
    
    if throttle_minutes == 0:  # No throttle for < 10 min
        return True, f"Gấp (<10 phút), gửi liên tục"
    
    if time_diff_minutes >= throttle_minutes:
        return True, f"Đã {time_diff_minutes:.1f} phút từ lần gửi cuối (throttle: {throttle_minutes}m)"
    else:
        return False, f"Chưa đủ {throttle_minutes} phút (mới {time_diff_minutes:.1f} phút)"

def send_warning_phieu_ton_sap_qua_gio():
    """
    Gửi cảnh báo phiếu tồn BRCD sắp quá giờ tới 4 nhóm Zalo + Telegram.
    Đọc từ file chiTietBrcd5Doi.xlsx, từ 4 sheet: thachthat, hoalac, hatmon, odien
    Gửi cảnh báo cho các phiếu sắp quá giờ (0 < giờ còn lại thực < 1.5).
    """
    file_path = os.path.join('chiaTheoDoi', 'chiTietBrcd5Doi.xlsx')
    webhook_url = WEBHOOK_TEXT_URL

    # Import Telegram config
    try:
        from send_tele import (
            TELEGRAM_TOKEN
        )
    except ImportError:
        print("⚠️ Không thể import module send_tele. Chỉ gửi Zalo.")
        TELEGRAM_TOKEN = None

    # Auto-generate sheet mappings from team_config (BRCD - 4 teams)
    brcd_teams = get_active_teams('BRCD')

    sheet_to_thread_id = {
        team.id: LOCATION_THREAD_MAPPING[team.id]
        for team in brcd_teams
    }

    sheet_to_telegram_chat_id = {
        team.id: LOCATION_CHAT_MAPPING.get(team.id)
        for team in brcd_teams
    }

    try:
        print(f"\n🚀 Bắt đầu gửi cảnh báo phiếu tồn BRCD...")

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Load log file for throttle tracking
        log_dir = "log_message"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, "phieu_sap_qua_gio_sent.xlsx")
        
        # Load existing log or create new DataFrame
        if os.path.exists(log_file):
            try:
                log_df = pd.read_excel(log_file)
                print(f"  📋 Đã load {len(log_df)} bản ghi từ log")
            except Exception as e:
                print(f"  ⚠️ Lỗi khi đọc log file: {e}, tạo mới")
                log_df = pd.DataFrame(columns=['ma_tb', 'baohong_id', 'last_sent_time', 'gio_con_lai', 'DOI_VT', 'status'])
        else:
            log_df = pd.DataFrame(columns=['ma_tb', 'baohong_id', 'last_sent_time', 'gio_con_lai', 'DOI_VT', 'status'])
            print(f"  📋 Tạo log file mới")
        
        # Track new entries to append
        new_log_entries = []

        # Tổng số bản tin đã gửi
        total_sent_zalo = 0
        total_failed_zalo = 0
        total_sent_telegram = 0
        total_failed_telegram = 0
        total_skipped = 0

        # Đọc từng sheet
        for sheet_name, thread_id in sheet_to_thread_id.items():
            try:
                print(f"\n📋 Đang xử lý sheet: {sheet_name}")

                # Đọc sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                if df.empty:
                    print(f"  ℹ️ Sheet {sheet_name} không có dữ liệu")
                    continue

                # Kiểm tra các cột cần thiết
                required_cols = ['giờ còn lại thực', 'ma_tb', 'ngay_bh', 'ds_nhanvien_th', 'Trạng thái cổng']
                missing_cols = [col for col in required_cols if col not in df.columns]

                if missing_cols:
                    print(f"  ⚠️ Sheet {sheet_name} thiếu các cột: {', '.join(missing_cols)}")
                    continue

                # Kiểm tra xem có cột ghichuton không (không bắt buộc, dùng để hiển thị lý do tồn)
                has_ghichuton = 'ghichuton' in df.columns

                # Lọc phiếu sắp quá giờ: 0 < giờ còn lại thực < 1
                # Chỉ gửi cảnh báo các phiếu sắp quá giờ, không gửi phiếu đã quá giờ
                df_filtered = df[(df['giờ còn lại thực'] > 0) & (df['giờ còn lại thực'] < 1.5)].copy()

                if df_filtered.empty:
                    print(f"  ℹ️ Không có phiếu sắp quá giờ trong sheet {sheet_name}")
                    continue

                print(f"  🔍 Tìm thấy {len(df_filtered)} phiếu sắp quá giờ")

                # Chuyển đổi ngày báo hỏng sang định dạng dễ đọc
                if 'ngay_bh' in df_filtered.columns:
                    df_filtered['ngay_bh_formatted'] = pd.to_datetime(
                        df_filtered['ngay_bh'],
                        errors='coerce'
                    ).dt.strftime('%d/%m/%Y %H:%M')
                else:
                    df_filtered['ngay_bh_formatted'] = 'N/A'


                # Build messages with throttle check
                messages_to_send = []  # Now stores tuple: (row, message, ma_tb, gio_con_lai, baohong_id)

                for _, row in df_filtered.iterrows():
                    ma_tb = row.get('ma_tb', 'N/A')
                    gio_con_lai = row['giờ còn lại thực']
                    baohong_id = row.get('baohong_id', '') if 'baohong_id' in row else ''
                    
                    # Lấy trạng thái cổng
                    trang_thai_cong = row.get('Trạng thái cổng', 'N/A')
                    if pd.isna(trang_thai_cong):
                        trang_thai_cong = 'N/A'

                    # Lấy lý do tồn (nếu có cột ghichuton)
                    if has_ghichuton:
                        ly_do_ton = row.get('ghichuton', 'N/A')
                    else:
                        ly_do_ton = 'N/A'

                    # Xử lý giá trị null cho ly_do_ton
                    if pd.isna(ly_do_ton):
                        ly_do_ton = 'N/A'

                    # Format time remaining
                    total_minutes = int(round(gio_con_lai * 60))
                    hours = total_minutes // 60
                    minutes = total_minutes % 60
                    
                    if hours > 0:
                        thoi_gian_con_str = f"{hours} giờ {minutes} phút"
                    else:
                        thoi_gian_con_str = f"{minutes} phút"

                    # Tạo message mới với format cảnh báo sắp quá giờ
                    message = (
                        f"🔔Phiếu BÁO HỎNG sắp quá giờ\n"
                        f"  - Mã TB: {ma_tb}\n"
                        f"  - Ngày báo: {row.get('ngay_bh_formatted', 'N/A')}\n"
                        f"  - NVKT: {row.get('ds_nhanvien_th', 'N/A')}\n"
                        f"  - Trạng thái cổng: {trang_thai_cong}\n"
                        f"  - Lý do tồn: {ly_do_ton}\n"
                        f"  - Thời gian còn: {thoi_gian_con_str}"
                    )

                    messages_to_send.append((row, message, ma_tb, gio_con_lai, baohong_id))

                # Gửi từng bản tin
                print(f"  📤 Đang gửi {len(messages_to_send)} bản tin cho sheet {sheet_name}...")

                sent_count_zalo = 0
                failed_count_zalo = 0
                sent_count_telegram = 0
                failed_count_telegram = 0

                # Lấy Telegram chat ID cho sheet này
                telegram_chat_id = sheet_to_telegram_chat_id.get(sheet_name)

                for i, (row, message, ma_tb, gio_con_lai, baohong_id) in enumerate(messages_to_send, 1):
                    # CHECK THROTTLE BEFORE SENDING
                    should_send, reason = should_send_sap_qua_gio_warning(ma_tb, gio_con_lai, log_df)
                    
                    if not should_send:
                        total_skipped += 1
                        print(f"    ⏭️  [{i}/{len(messages_to_send)}] Skip {ma_tb}: {reason}")
                        # Log as SKIP
                        new_log_entries.append({
                            'ma_tb': ma_tb,
                            'baohong_id': baohong_id,
                            'last_sent_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'gio_con_lai': gio_con_lai,
                            'DOI_VT': sheet_name,
                            'status': 'SKIP'
                        })
                        continue  # Skip sending this message
                    
                    print(f"    📨 [{i}/{len(messages_to_send)}] Gửi {ma_tb}: {reason}")
                    
                    # Gửi Zalo
                    zalo_success = False
                    try:
                        # Chuẩn bị dữ liệu để gửi qua webhook
                        data = {
                            'threadID': thread_id,
                            'message': message
                        }

                        # Gửi request
                        response = requests.get(webhook_url, json=data, timeout=10)

                        if response.status_code == 200:
                            sent_count_zalo += 1
                            print(f"    ✅ Zalo [{i}/{len(messages_to_send)}] Gửi thành công")
                        else:
                            failed_count_zalo += 1
                            print(f"    ❌ Zalo [{i}/{len(messages_to_send)}] Gửi thất bại. Status: {response.status_code}")

                    except Exception as e:
                        failed_count_zalo += 1
                        print(f"    ❌ Zalo [{i}/{len(messages_to_send)}] Lỗi khi gửi: {e}")

                    # Gửi Telegram
                    if telegram_chat_id and TELEGRAM_TOKEN:
                        try:
                            import requests as telegram_requests
                            url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
                            data = {
                                "chat_id": telegram_chat_id,
                                "text": message,
                                "parse_mode": "HTML"
                            }
                            response = telegram_requests.post(url, data=data, timeout=10)

                            if response.status_code == 200:
                                sent_count_telegram += 1
                                print(f"    ✅ Telegram [{i}/{len(messages_to_send)}] Gửi thành công")
                            else:
                                failed_count_telegram += 1
                                print(f"    ❌ Telegram [{i}/{len(messages_to_send)}] Gửi thất bại. Status: {response.status_code}")

                        except Exception as e:
                            failed_count_telegram += 1
                            print(f"    ❌ Telegram [{i}/{len(messages_to_send)}] Lỗi khi gửi: {e}")

                    # Log successful send
                    if zalo_success or telegram_success:
                        new_log_entries.append({
                            'ma_tb': ma_tb,
                            'baohong_id': baohong_id,
                            'last_sent_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'gio_con_lai': gio_con_lai,
                            'DOI_VT': sheet_name,
                            'status': 'SENT'
                        })

                    # Tạm dừng 1 giây để tránh spam
                    if i < len(messages_to_send):
                        time.sleep(1)

                total_sent_zalo += sent_count_zalo
                total_failed_zalo += failed_count_zalo
                total_sent_telegram += sent_count_telegram
                total_failed_telegram += failed_count_telegram

                print(f"  📊 Sheet {sheet_name}: Zalo ✅ {sent_count_zalo} | ❌ {failed_count_zalo} | Telegram ✅ {sent_count_telegram} | ❌ {failed_count_telegram}")

            except ValueError as e:
                if "Worksheet named" in str(e):
                    print(f"  ⚠️ Sheet '{sheet_name}' không tồn tại trong file")
                else:
                    print(f"  ❌ ValueError khi xử lý sheet {sheet_name}: {e}")
                continue
            except Exception as e:
                print(f"  ❌ Lỗi khi xử lý sheet {sheet_name}: {e}")
                continue

        # Save log file
        if new_log_entries:
            try:
                df_new_log = pd.DataFrame(new_log_entries)
                
                if os.path.exists(log_file):
                    # Append to existing log
                    df_existing = pd.read_excel(log_file)
                    df_combined = pd.concat([df_existing, df_new_log], ignore_index=True)
                else:
                    df_combined = df_new_log
                
                df_combined.to_excel(log_file, index=False)
                print(f"\n💾 Đã lưu log: {len(new_log_entries)} entries vào {log_file}")
            except Exception as e:
                print(f"\n⚠️ Lỗi khi lưu log: {e}")

        # Tổng kết
        print("\n" + "=" * 60)
        print("📋 TỔNG KẾT:")
        print(f"  [Zalo]     ✅ Gửi thành công: {total_sent_zalo} bản tin")
        print(f"  [Zalo]     ❌ Gửi thất bại: {total_failed_zalo} bản tin")
        print(f"  [Telegram] ✅ Gửi thành công: {total_sent_telegram} bản tin")
        print(f"  [Telegram] ❌ Gửi thất bại: {total_failed_telegram} bản tin")
        print(f"  ⏭️  Đã bỏ qua (throttle): {total_skipped} bản tin")
        print(f"  📊 Tổng cộng: {total_sent_zalo + total_failed_zalo + total_sent_telegram + total_failed_telegram} bản tin")
        print("=" * 60)

        return True

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file {file_path}")
        return False
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()
        return False


def send_thong_ke_phieu_ton_tong_hop():
    """
    Gửi báo cáo thống kê phiếu tồn tổng hợp tới 4 nhóm Zalo + Telegram.
    Đọc từ file chiTietBrcd5Doi.xlsx, từ 4 sheet: ToKT_PhucTho_rut_gon, ToKT_SonTay_rut_gon, 
    ToKT_QuangOai_rut_gon, ToKT_SuoiHai_rut_gon
    Tạo báo cáo thống kê tổng hợp cho mỗi tổ.
    """
    file_path = os.path.join('chiaTheoDoi', 'chiTietBrcd5Doi.xlsx')
    webhook_url = WEBHOOK_TEXT_URL

    # Check allowed time
    allowed, time_msg = is_allowed_send_time()
    if not allowed:
        print(f"⏰ {time_msg}")
        return False

    # Import Telegram config
    try:
        from send_tele import TELEGRAM_TOKEN
    except ImportError:
        print("⚠️ Không thể import module send_tele. Chỉ gửi Zalo.")
        TELEGRAM_TOKEN = None

    # Auto-generate sheet mappings from team_config (BRCD - 4 teams)
    brcd_teams = get_active_teams('BRCD')

    # Map sheets to teams
    sheet_mapping = {
        'ToKT_PhucTho_rut_gon': None,
        'ToKT_SonTay_rut_gon': None,
        'ToKT_QuangOai_rut_gon': None,
        'ToKT_SuoiHai_rut_gon': None
    }
    
    # Match sheets to teams by ID
    for team in brcd_teams:
        for sheet_name in sheet_mapping.keys():
            if team.id.replace('ToKT_', '') in sheet_name:
                sheet_mapping[sheet_name] = team
                break

    try:
        print(f"\n🚀 Bắt đầu gửi báo cáo thống kê phiếu tồn tổng hợp...")

        if not os.path.exists(file_path):
            print(f"❌ Không tìm thấy file: {file_path}")
            return False

        # Tổng số bản tin đã gửi
        total_sent_zalo = 0
        total_failed_zalo = 0
        total_sent_telegram = 0
        total_failed_telegram = 0

        # Đọc từng sheet
        for sheet_name, team in sheet_mapping.items():
            if not team:
                print(f"⚠️ Không tìm thấy team cho sheet: {sheet_name}")
                continue

            try:
                print(f"\n📋 Đang xử lý sheet: {sheet_name} - Tổ {team.short_name}")

                # Đọc sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                if df.empty:
                    print(f"  ℹ️ Sheet {sheet_name} không có dữ liệu")
                    continue

                # Kiểm tra các cột cần thiết
                required_cols = ['giờ còn lại thực', 'ma_tb', 'ngay_bh', 'NVKT', 'LOAIHINH_TB']
                missing_cols = [col for col in required_cols if col not in df.columns]

                if missing_cols:
                    print(f"  ⚠️ Sheet {sheet_name} thiếu các cột: {', '.join(missing_cols)}")
                    continue

                # THỐNG KÊ
                total_tickets = len(df)
                print(f"  📊 Tổng số phiếu tồn: {total_tickets}")

                # 1. Thống kê theo trạng thái cổng
                if 'Trạng thái cổng' in df.columns:
                    status_counts = df['Trạng thái cổng'].fillna('N/A').value_counts()
                    on_count = status_counts.get('ON', 0)
                    off_count = status_counts.get('OFF', 0)
                    na_count = status_counts.get('N/A', 0)
                else:
                    on_count = off_count = na_count = 0

                # 2. Thống kê theo loại hình thuê bao
                loaihinh_counts = df['LOAIHINH_TB'].fillna('Khác').value_counts()
                fiber_count = loaihinh_counts.get('Fiber', 0)
                mytv_count = loaihinh_counts.get('MyTV', 0)
                phone_count = loaihinh_counts.get('Điện thoại cố định', 0)
                sip_count = loaihinh_counts.get('Thuê bao SIP', 0)
                other_count = total_tickets - (fiber_count + mytv_count + phone_count + sip_count)

                # 3. Thống kê theo thời gian còn lại
                qua_gio = len(df[df['giờ còn lại thực'] < 0])
                sap_qua = len(df[(df['giờ còn lại thực'] >= 0) & (df['giờ còn lại thực'] <= 1.5)])
                con_tg = len(df[df['giờ còn lại thực'] > 1.5])

                # 4. Thống kê TẤT CẢ NVKT với chi tiết loại hình
                nvkt_counts = df['NVKT'].fillna('N/A').value_counts()
                all_nvkt = []
                
                for i, (nvkt, count) in enumerate(nvkt_counts.items(), 1):
                    # Lọc dữ liệu của NVKT này
                    nvkt_df = df[df['NVKT'].fillna('N/A') == nvkt]

                    # Đếm từng loại hình cho NVKT này
                    nvkt_loaihinh = nvkt_df['LOAIHINH_TB'].fillna('Khác').value_counts()

                    # Tạo danh sách chi tiết loại hình
                    details = []
                    if nvkt_loaihinh.get('Fiber', 0) > 0:
                        details.append(f"Fiber: {nvkt_loaihinh.get('Fiber', 0)}")
                    if nvkt_loaihinh.get('MyTV', 0) > 0:
                        details.append(f"MyTV: {nvkt_loaihinh.get('MyTV', 0)}")
                    if nvkt_loaihinh.get('Điện thoại cố định', 0) > 0:
                        details.append(f"ĐTCĐ: {nvkt_loaihinh.get('Điện thoại cố định', 0)}")
                    if nvkt_loaihinh.get('Thuê bao SIP', 0) > 0:
                        details.append(f"SIP: {nvkt_loaihinh.get('Thuê bao SIP', 0)}")

                    # Tính số lượng "Khác" (các loại không phải Fiber, MyTV, ĐTCĐ, SIP)
                    known_types = ['Fiber', 'MyTV', 'Điện thoại cố định', 'Thuê bao SIP']
                    other_nvkt = len(nvkt_df[~nvkt_df['LOAIHINH_TB'].isin(known_types)])
                    if other_nvkt > 0:
                        details.append(f"Khác: {other_nvkt}")

                    # Tạo danh sách mã thuê bao với giờ báo và giờ còn lại (sắp xếp theo giờ còn lại tăng dần)
                    nvkt_df_sorted = nvkt_df.sort_values('giờ còn lại thực', ascending=True)
                    ma_tb_list = []
                    for _, row in nvkt_df_sorted.iterrows():
                        ma_tb = str(row['ma_tb']) if pd.notna(row['ma_tb']) else 'N/A'

                        # Format giờ báo (ngay_bh) thành dạng rút gọn HH:MM
                        ngay_bh = row.get('ngay_bh', None)
                        if pd.notna(ngay_bh):
                            try:
                                ngay_bh_dt = pd.to_datetime(ngay_bh, errors='coerce')
                                if pd.notna(ngay_bh_dt):
                                    gio_bao = ngay_bh_dt.strftime('%H:%M')
                                else:
                                    gio_bao = 'N/A'
                            except:
                                gio_bao = 'N/A'
                        else:
                            gio_bao = 'N/A'

                        # Format giờ còn lại
                        gio_con_lai = row['giờ còn lại thực']
                        if pd.notna(gio_con_lai):
                            gio_str = f"{gio_con_lai:.1f}h"
                        else:
                            gio_str = "N/A"

                        ma_tb_list.append(f"{ma_tb} | {gio_bao} | {gio_str}")

                    # Format: "1. Tên NVKT: X phiếu (Chi tiết) (MA_TB | giờ báo | giờ còn lại)"
                    detail_str = f"({', '.join(details)})" if details else ""
                    ma_tb_str = f" ({'; '.join(ma_tb_list)})" if ma_tb_list else ""
                    all_nvkt.append(f"{i}. {nvkt}: {count} phiếu {detail_str}{ma_tb_str}")


                # Tạo phần loại hình thuê bao (bỏ dòng Khác nếu = 0)
                loaihinh_lines = [
                    f"- Fiber: {fiber_count} phiếu",
                    f"- MyTV: {mytv_count} phiếu",
                    f"- Điện thoại cố định: {phone_count} phiếu",
                    f"- Thuê bao SIP: {sip_count} phiếu"
                ]
                if other_count > 0:
                    loaihinh_lines.append(f"- Khác: {other_count} phiếu")

                # Tạo bản tin báo cáo
                current_time = datetime.now().strftime('%d/%m/%Y %H:%M')
                
                message = f"""📊 BÁO CÁO THỐNG KÊ PHIẾU TỒN - TỔ {team.short_name.upper()}
Thời gian: {current_time}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📌 TỔNG QUAN
- Tổng số phiếu tồn: {total_tickets} phiếu

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔌 THEO TRẠNG THÁI CỔNG
- ON: {on_count} phiếu
- OFF: {off_count} phiếu
- N/A: {na_count} phiếu

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📡 THEO LOẠI HÌNH THUÊ BAO
{chr(10).join(loaihinh_lines)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⏰ THEO THỜI GIAN CÒN LẠI
- Quá giờ (< 0h): {qua_gio} phiếu
- Sắp quá (0-1.5h): {sap_qua} phiếu
- Còn thời gian (> 1.5h): {con_tg} phiếu

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
👨‍🔧 THỐNG KÊ THEO NVKT
{chr(10).join(all_nvkt) if all_nvkt else 'Không có dữ liệu'}"""


                # Gửi Zalo
                zalo_success = False
                try:
                    data = {
                        'threadID': team.zalo_thread_id,
                        'message': message
                    }
                    print(f"  📤 Gọi webhook Zalo cho tổ {team.short_name}...")
                    response = requests.get(webhook_url, json=data, timeout=10)

                    if response.status_code == 200:
                        total_sent_zalo += 1
                        zalo_success = True
                        print(f"  ✅ [Zalo] Gửi thành công cho tổ {team.short_name}")
                    else:
                        total_failed_zalo += 1
                        print(f"  ❌ [Zalo] Gửi thất bại (Status: {response.status_code})")

                except Exception as e:
                    total_failed_zalo += 1
                    print(f"  ❌ [Zalo] Lỗi khi gửi: {e}")

                # Tạm dừng 0.5 giây giữa Zalo và Telegram
                time.sleep(0.5)

                # Gửi Telegram
                telegram_success = False
                if team.telegram_chat_id and TELEGRAM_TOKEN:
                    try:
                        url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
                        data = {
                            "chat_id": team.telegram_chat_id,
                            "text": message,
                            "parse_mode": "HTML"
                        }
                        response = requests.post(url, data=data, timeout=10)

                        if response.status_code == 200:
                            total_sent_telegram += 1
                            telegram_success = True
                            print(f"  ✅ [Telegram] Gửi thành công cho tổ {team.short_name}")
                        else:
                            total_failed_telegram += 1
                            print(f"  ❌ [Telegram] Gửi thất bại (Status: {response.status_code})")

                    except Exception as e:
                        total_failed_telegram += 1
                        print(f"  ❌ [Telegram] Lỗi khi gửi: {e}")

                # Tạm dừng 1 giây giữa các tổ để tránh spam
                time.sleep(1)

            except Exception as e:
                print(f"  ❌ Lỗi khi xử lý sheet {sheet_name}: {e}")
                continue

        # Tổng kết
        print("\n" + "=" * 60)
        print("📋 TỔNG KẾT BÁO CÁO THỐNG KÊ:")
        print(f"  [Zalo]     ✅ Gửi thành công: {total_sent_zalo} bản tin")
        print(f"  [Zalo]     ❌ Gửi thất bại: {total_failed_zalo} bản tin")
        print(f"  [Telegram] ✅ Gửi thành công: {total_sent_telegram} bản tin")
        print(f"  [Telegram] ❌ Gửi thất bại: {total_failed_telegram} bản tin")
        print(f"  📊 Tổng cộng: {total_sent_zalo + total_failed_zalo} bản tin")
        print("=" * 60)

        return True

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file {file_path}")
        return False
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    #send_zalo_theo_huyen()
    #send_Screenshot_fms()

    # Gửi tất cả ảnh từ các thư mục đến Zalo
    # send_image_to_zalo(receiver="xuanthinh")

    # Gửi tin nhắc phiếu tồn
    # send_nhac_phieu_ton_cho_nhan_vien()

    # # Gửi cảnh báo phiếu tồn BRCD
    send_warning_phieu_ton_brcd()

    # # Gửi cảnh báo phiếu quá giờ (có log nhắc lại sau 6 tiếng)
    # send_warning_phieu_qua_gio()

    # # Gửi cảnh báo phiếu hỏng lại trong tháng
    # send_warning_hong_lai_trong_thang()

    # Gửi cảnh báo phiếu KHDN ưu tiên
    # send_warning_khdn_uu_tien()
    # send_warning_phieu_ob_khl()

    # Original screenshot code (commented out)
    # latest_screenshot = get_latest_file_in_dir("Screenshot_fms")
    # if latest_screenshot:
    #     send_image_via_webhook(latest_screenshot, receiver="xuanthinh", message="Screenshot FMS")
    # else:
    #     print("Không tìm thấy file screenshot mới nhất trong thư mục Screenshot_fms!")
