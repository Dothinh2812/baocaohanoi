import hashlib
import os
import time
from datetime import timedelta


BASE_DIR = os.path.abspath(os.path.dirname(__file__))


def _env_value(primary_name, fallback_name=None, default=None):
    value = os.getenv(primary_name)
    if value is not None:
        return value
    if fallback_name:
        value = os.getenv(fallback_name)
        if value is not None:
            return value
    return default


APP_TIMEZONE = _env_value('DASHV4_TIMEZONE', 'DASH_TIMEZONE', 'Asia/Ho_Chi_Minh')


def apply_process_timezone():
    os.environ['TZ'] = APP_TIMEZONE
    if hasattr(time, 'tzset'):
        time.tzset()


apply_process_timezone()


def _env_flag(name, default=False):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {'1', 'true', 'yes', 'on'}


def _first_existing_path(*candidates):
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return next((candidate for candidate in candidates if candidate), '')


def _default_secret_key():
    seed = f"{BASE_DIR}:{os.getenv('USER', 'dashboard')}:dashv4"
    return hashlib.sha256(seed.encode('utf-8')).hexdigest()


DEFAULT_BASE_DATA_PATH = os.path.abspath(os.path.join(BASE_DIR, '..', 'one-suachua'))
FALLBACK_BASE_DATA_PATH = os.path.abspath(os.path.join(BASE_DIR, '..', 'onev2'))
UNIT_CODE = _env_value('DASHV4_UNIT_CODE', default='son_tay')
UNIT_NAME = _env_value('DASHV4_UNIT_NAME', default='TTVT Sơn Tây')
APP_DISPLAY_NAME = _env_value('DASHV4_APP_NAME', default='Dashboard V4')
INSTANCE_RUNTIME_DIR = os.path.abspath(
    _env_value(
        'DASHV4_RUNTIME_DIR',
        default=os.path.join(BASE_DIR, 'runtime_app', UNIT_CODE),
    )
)

BASE_DATA_PATH = _first_existing_path(
    os.getenv('DASH_BASE_DATA_PATH'),
    DEFAULT_BASE_DATA_PATH,
    FALLBACK_BASE_DATA_PATH,
)
BAOCAO_HANOI_PATH = _first_existing_path(
    os.getenv('DASH_BAOCAO_HANOI_PATH'),
    '/home/vtst/baocaohanoi',
)
BAOCAO_VATTU_PATH = _first_existing_path(
    os.getenv('DASH_BAOCAO_VATTU_PATH'),
    '/home/vtst/baocao-vattu',
)
BAOCAO_HANOI_DOWNLOADS_DIR = os.path.join(BAOCAO_HANOI_PATH, 'downloads', 'baocao_hanoi')
BAOCAO_HANOI_KPI_DIR = os.path.join(BAOCAO_HANOI_PATH, 'downloads', 'KPI')
CAU_HINH_TU_DONG_DIR = os.path.join(BAOCAO_HANOI_PATH, 'api_transition', 'Processed', 'cau_hinh_tu_dong')
CAU_HINH_TU_DONG_CHI_TIET_FILE = os.path.join(
    CAU_HINH_TU_DONG_DIR,
    'cau_hinh_tu_dong_chi_tiet_processed.xlsx',
)
GHTT_DIR = os.path.join(BAOCAO_HANOI_PATH, 'GHTT')
KQ_TIEP_THI_DIR = os.path.join(BAOCAO_HANOI_PATH, 'KQ-TIEP-THI')
PTTB_PSC_DIR = os.path.join(BAOCAO_HANOI_PATH, 'PTTB-PSC')
KPI_TONGHOP_NVKT_FILE = os.path.join(BAOCAO_HANOI_PATH, 'KPI_TongHop_NVKT.xlsx')
REPORT_HISTORY_DB_PATH = _first_existing_path(
    _env_value('DASHV4_DB_PATH', 'DASH_REPORT_HISTORY_DB'),
    os.path.join(BAOCAO_HANOI_PATH, 'api_transition', 'runtime', 'son_tay', 'sqlite_history', 'report_history.db'),
    os.path.join(BAOCAO_HANOI_PATH, 'api_transition', 'report_history.db'),
    os.path.join(BAOCAO_HANOI_PATH, 'report_history.db'),
)
SHC_PROCESSING_RESULTS_DB_PATH = _first_existing_path(
    os.getenv('DASH_SHC_PROCESSING_RESULTS_DB'),
    os.path.abspath(os.path.join(BASE_DIR, '..', 'thong-ke-xly-shc', 'data', 'nvkt_results.db')),
)
SHC_SOURCE_K1_DB_PATH = _first_existing_path(
    os.getenv('DASH_SHC_SOURCE_K1_DB'),
    os.path.join(BAOCAO_HANOI_PATH, 'suy_hao_history.db'),
)
SHC_SOURCE_K2_DB_PATH = _first_existing_path(
    os.getenv('DASH_SHC_SOURCE_K2_DB'),
    os.path.join(BAOCAO_HANOI_PATH, 'suy_hao_history_k2.db'),
)

SESSION_FILE_DIR = os.path.abspath(
    _env_value('DASHV4_SESSION_FILE_DIR', default=os.path.join(INSTANCE_RUNTIME_DIR, 'flask_session'))
)
CACHE_DIR = os.path.abspath(
    _env_value('DASHV4_CACHE_DIR', default=os.path.join(INSTANCE_RUNTIME_DIR, 'cache'))
)
LOG_DIR = os.path.abspath(_env_value('DASHV4_LOG_DIR', default=os.path.join(BASE_DIR, 'logs', UNIT_CODE)))
PID_FILE = _env_value('DASHV4_PID_FILE', default=f'/tmp/dashv4-{UNIT_CODE}.pid')
USER_FILE = os.path.abspath(_env_value('DASHV4_USER_FILE', default=os.path.join(INSTANCE_RUNTIME_DIR, 'users.xlsx')))
LOGIN_LOG_FILE = os.path.abspath(
    _env_value('DASHV4_LOGIN_LOG_FILE', default=os.path.join(LOG_DIR, 'login_history.csv'))
)
QUANG_CHU_DONG_CACHE_DIR = os.path.join(CACHE_DIR, 'quangchudong')
QUANG_CHU_DONG_CACHE_FILE = os.path.join(QUANG_CHU_DONG_CACHE_DIR, 'snapshot.json')
QUANG_CHU_DONG_CACHE_LOCK_FILE = os.path.join(QUANG_CHU_DONG_CACHE_DIR, 'snapshot.lock')
QUANG_CHU_DONG_CACHE_REFRESH_SECONDS = int(os.getenv('DASH_QUANG_CHU_DONG_CACHE_REFRESH_SECONDS', '60'))
QUANG_CHU_DONG_CACHE_WAIT_SECONDS = int(os.getenv('DASH_QUANG_CHU_DONG_CACHE_WAIT_SECONDS', '10'))

SHC_NVKT_DETAIL_REPORTS = {
    'k1': {
        'label': 'K1',
        'path': os.path.join(BAOCAO_HANOI_PATH, 'downloads', 'baocao_hanoi', 'shc_NVKT_danh_sach_chi_tiet_K1'),
    },
    'k2': {
        'label': 'K2',
        'path': os.path.join(BAOCAO_HANOI_PATH, 'downloads', 'baocao_hanoi', 'shc_NVKT_danh_sach_chi_tiet_K2'),
    },
    'k2-26-27': {
        'label': 'K2_26_27',
        'path': os.path.join(BAOCAO_HANOI_PATH, 'downloads', 'baocao_hanoi', 'shc_NVKT_danh_sach_chi_tiet_K2_26_27'),
    },
}

SHC_NVKT_TEAMS = {
    'phuc-tho': 'Tổ Kỹ thuật Địa bàn Phúc Thọ',
    'quang-oai': 'Tổ Kỹ thuật Địa bàn Quảng Oai',
    'son-tay': 'Tổ Kỹ thuật Địa bàn Sơn Tây',
    'suoi-hai': 'Tổ Kỹ thuật Địa bàn Suối hai',
}

QUANG_CHU_DONG_DB_LIST = [
    {'name': 'STY', 'path': '/home/vtst/do-kiem-chu-dong-sty/subscriber_history.db'},
    {'name': 'SHI', 'path': '/home/vtst/do-kiem-chu-dong-shi/subscriber_history.db'},
]
QUANG_CHU_DONG_DB_PATH = QUANG_CHU_DONG_DB_LIST[0]['path']

INVENTORY_PROCESSED_FILE = os.path.join(BAOCAO_VATTU_PATH, '66 bc ton vat tu_processed.xlsx')
INVENTORY_SOURCE_FILE = os.path.join(BAOCAO_VATTU_PATH, '66 bc ton vat tu.xlsx')
INVENTORY_COMMON_GOOD_FILE = os.path.join(BAOCAO_VATTU_PATH, '66 bc ton vat tu_processed-tot-thuong-dung.xlsx')
XAC_MINH_TAM_DUNG_DIR = os.path.join(BAOCAO_HANOI_PATH, 'XAC MINH TAM DUNG')
TAM_DUNG_KHOI_PHUC_FILE = os.path.join(
    BAOCAO_HANOI_PATH,
    'api_transition',
    'Processed',
    'tam_dung_khoi_phuc_dich_vu',
    'tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_processed.xlsx',
)
SH_PORTAL_URL = os.getenv('DASH_SH_PORTAL_URL', 'https://sh.ttvt8.online/')
SH_PORTAL_USERNAME_PARAM = os.getenv('DASH_SH_PORTAL_USERNAME_PARAM', 'username')
INVENTORY_TEAM_CONFIGS = {
    'qoi': {
        'active_page': 'ton_kho_vat_tu_qoi',
        'template': 'pages/ton_kho_vat_tu_qoi.html',
        'sheet_name': 'Tổ KTĐB Quảng Oai',
        'base_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_604_Quảng Oai.xlsx'),
        'base_download_name': 'vattu_604_Quang_Oai.xlsx',
        'good_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_604_Quảng Oai-tot-thuong-dung.xlsx'),
        'good_download_name': 'vattu_604_Quang_Oai_tot_thuong_dung.xlsx',
    },
    'shi': {
        'active_page': 'ton_kho_vat_tu_shi',
        'template': 'pages/ton_kho_vat_tu_shi.html',
        'sheet_name': 'Tổ KTĐB Suối Hai',
        'base_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_606_Suối Hai.xlsx'),
        'base_download_name': 'vattu_606_Suoi_Hai.xlsx',
        'good_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_606_Suối Hai-tot-thuong-dung.xlsx'),
        'good_download_name': 'vattu_606_Suoi_Hai_tot_thuong_dung.xlsx',
    },
    'pto': {
        'active_page': 'ton_kho_vat_tu_pto',
        'template': 'pages/ton_kho_vat_tu_pto.html',
        'sheet_name': 'Tổ KTĐB Phúc Thọ',
        'base_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_901_Phúc Thọ.xlsx'),
        'base_download_name': 'vattu_901_Phuc_Tho.xlsx',
        'good_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_901_Phúc Thọ-tot-thuong-dung.xlsx'),
        'good_download_name': 'vattu_901_Phuc_Tho_tot_thuong_dung.xlsx',
    },
    'sty': {
        'active_page': 'ton_kho_vat_tu_sty',
        'template': 'pages/ton_kho_vat_tu_sty.html',
        'sheet_name': 'Tổ KTĐB Sơn Tây',
        'base_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_605_Sơn Tây.xlsx'),
        'base_download_name': 'vattu_605_Son_Tay.xlsx',
        'good_file': os.path.join(BAOCAO_VATTU_PATH, 'vattu_605_Sơn Tây-tot-thuong-dung.xlsx'),
        'good_download_name': 'vattu_605_Son_Tay_tot_thuong_dung.xlsx',
    },
}


class DashboardConfig:
    SECRET_KEY = _env_value('DASHV4_SECRET_KEY', 'DASH_SECRET_KEY', _default_secret_key())
    TIMEZONE = APP_TIMEZONE
    SESSION_TYPE = 'filesystem'
    SESSION_FILE_DIR = SESSION_FILE_DIR
    SESSION_PERMANENT = True
    PERMANENT_SESSION_LIFETIME = timedelta(hours=12)
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = _env_value('DASHV4_SESSION_COOKIE_SAMESITE', 'DASH_SESSION_COOKIE_SAMESITE', 'Lax')
    SESSION_COOKIE_SECURE = _env_flag('DASHV4_SESSION_COOKIE_SECURE', default=_env_flag('DASH_SESSION_COOKIE_SECURE', default=False))
    TEMPLATES_AUTO_RELOAD = _env_flag('DASHV4_TEMPLATE_AUTO_RELOAD', default=_env_flag('DASH_TEMPLATE_AUTO_RELOAD', default=False))
    BASE_DATA_PATH = BASE_DATA_PATH
    BAOCAO_HANOI_PATH = BAOCAO_HANOI_PATH
    BAOCAO_VATTU_PATH = BAOCAO_VATTU_PATH
    SERVER_HOST = _env_value('DASHV4_HOST', 'DASH_HOST', '0.0.0.0')
    SERVER_PORT = int(_env_value('DASHV4_PORT', 'DASH_PORT', '5011'))
    DEBUG = _env_flag('DASHV4_DEBUG', default=_env_flag('DASH_DEBUG', default=False))
    SH_PORTAL_URL = SH_PORTAL_URL
    SH_PORTAL_USERNAME_PARAM = SH_PORTAL_USERNAME_PARAM
    REPORT_HISTORY_DB_PATH = REPORT_HISTORY_DB_PATH
    UNIT_CODE = UNIT_CODE
    UNIT_NAME = UNIT_NAME
    APP_DISPLAY_NAME = APP_DISPLAY_NAME
    INSTANCE_RUNTIME_DIR = INSTANCE_RUNTIME_DIR
    CACHE_DIR = CACHE_DIR
    LOG_DIR = LOG_DIR
    PID_FILE = PID_FILE
    USER_FILE = USER_FILE
    LOGIN_LOG_FILE = LOGIN_LOG_FILE
    ENABLE_BACKGROUND_SERVICES = _env_flag('DASHV4_ENABLE_BACKGROUND_SERVICES', default=False)


PUBLIC_ENDPOINTS = {
    'static',
    'auth.login',
    'auth.logout',
    'operations.page_kpi',
    'operations.page_kpi_nvkt_bchn',
    'operations.get_kpi_data',
    'operations.get_kpi_nvkt_bchn_data',
}


DISABLED_PAGE_ENDPOINTS = {
    'operations.page_brcd': {
        'title': 'BRCD',
        'active_page': 'brcd',
        'reason': 'Route này vẫn phụ thuộc dữ liệu ngoài SQLite runtime mới.',
    },
    'operations.page_pttb': {
        'title': 'PTTB',
        'active_page': 'pttb',
        'reason': 'Route này chưa có contract dữ liệu tương ứng trong report_history.db.',
    },
    'operations.page_thuctang': {
        'title': 'Thực tăng ảnh/chart',
        'active_page': 'thuctang',
        'reason': 'Màn hiện tại đọc ảnh/chart sẵn có, chưa có bản thay thế DB-native trong dashv4.',
    },
    'quality.page_shc_processing': {
        'title': 'SHC Processing',
        'active_page': 'shc_processing',
        'reason': 'Màn này vẫn dùng nguồn xử lý SHC riêng ngoài report_history.db.',
    },
    'inventory.page_ton_kho_vat_tu': {
        'title': 'Tồn kho vật tư',
        'active_page': 'ton_kho_vat_tu',
        'reason': 'Nguồn vật tư kho cũ chưa được đưa vào contract SQLite mới.',
    },
    'inventory.page_tong_hop_tien': {
        'title': 'Tổng hợp tiền vật tư',
        'active_page': 'tong_hop_tien',
        'reason': 'Nguồn vật tư kho cũ chưa được đưa vào contract SQLite mới.',
    },
    'inventory.page_tra_cuu_nhanh_vat_tu': {
        'title': 'Tra cứu nhanh vật tư',
        'active_page': 'tra_cuu_nhanh_vat_tu',
        'reason': 'Nguồn vật tư kho cũ chưa được đưa vào contract SQLite mới.',
    },
    'quangchudong.page_quangchudong': {
        'title': 'Quang chủ động',
        'active_page': 'quangchudong',
        'reason': 'Màn này vẫn phụ thuộc DB và cache dịch vụ riêng.',
    },
    'sa_outage.page_su_co_sa': {
        'title': 'Sự cố SA',
        'active_page': 'su_co_sa',
        'reason': 'Màn này vẫn phụ thuộc SQLite riêng ngoài report_history.db.',
    },
}


DISABLED_NONPAGE_ENDPOINTS = {
    'operations.get_excel_data': {
        'title': 'BRCD API',
        'reason': 'API BRCD cũ đọc Excel, chưa có contract tương ứng trong report_history.db.',
        'required_display_contract': {
            'summary': ['ton_dv_theo_to', 'ton_dv_theo_nvkt', 'tong_hop_5doi'],
            'details': ['chi_tiet_pending', 'chi_tiet_main'],
        },
    },
    'operations.get_excel_data_main': {
        'title': 'BRCD API main',
        'reason': 'API BRCD cũ đọc Excel, chưa có contract tương ứng trong report_history.db.',
        'required_display_contract': {
            'summary': ['ton_dv_theo_to', 'ton_dv_theo_nvkt', 'tong_hop_5doi'],
        },
    },
    'operations.get_excel_data_pending': {
        'title': 'BRCD API pending',
        'reason': 'API BRCD pending cũ đọc Excel, chưa có contract tương ứng trong report_history.db.',
        'required_display_contract': {
            'details': ['danh_sach_phieu_ton_khong_ly_do'],
        },
    },
    'operations.download_excel': {
        'title': 'BRCD download',
        'reason': 'Tải file Excel BRCD cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc view consumer mới',
        },
    },
    'operations.download_excel_main': {
        'title': 'BRCD download main',
        'reason': 'Tải file Excel BRCD cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc view consumer mới',
        },
    },
    'operations.download_bc_brcd': {
        'title': 'BRCD file gốc',
        'reason': 'Tải file Excel BRCD cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc view consumer mới',
        },
    },
    'operations.get_pttb_data_summary': {
        'title': 'PTTB summary',
        'reason': 'PTTB chưa có contract dữ liệu trong report_history.db.',
        'required_display_contract': {
            'summary': ['tong_hop_don_gian', 'tong_hop_5doi', 'trang_thai'],
        },
    },
    'operations.get_pttb_data_detail': {
        'title': 'PTTB detail',
        'reason': 'PTTB chưa có contract dữ liệu trong report_history.db.',
        'required_display_contract': {
            'details': ['tong_hop_dia_ban', 'chi_tiet_theo_to'],
        },
    },
    'operations.get_pttb_data_pending': {
        'title': 'PTTB pending',
        'reason': 'PTTB chưa có contract dữ liệu trong report_history.db.',
        'required_display_contract': {
            'details': ['phieu_qua_gio', 'phieu_chua_co_ly_do_ton'],
        },
    },
    'operations.get_pttb_data_chitiet_to': {
        'title': 'PTTB chi tiết tổ',
        'reason': 'PTTB chưa có contract dữ liệu trong report_history.db.',
        'required_display_contract': {
            'details': ['danh_sach_phieu_theo_to'],
        },
    },
    'operations.download_excel_pttb': {
        'title': 'PTTB file gốc',
        'reason': 'Tải file Excel PTTB cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc view consumer mới',
        },
    },
    'retention.download_giahan_ghtt': {
        'title': 'Gia hạn file gốc',
        'reason': 'Tải file Excel gia hạn cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db',
        },
    },
    'retention.get_giahan_data': {
        'title': 'Gia hạn KR6',
        'reason': 'API Excel KR6 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {
            'replacement': ['/api/giahan-ghtt-sty', '/api/giahan-ghtt-nvktdb'],
        },
    },
    'retention.get_giahan_data_to': {
        'title': 'Gia hạn KR6 theo tổ',
        'reason': 'API Excel KR6 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {
            'replacement': ['/api/giahan-ghtt-sty'],
        },
    },
    'retention.get_giahan_kr7_data': {
        'title': 'Gia hạn KR7',
        'reason': 'API Excel KR7 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {
            'replacement': ['/api/giahan-ghtt-nvktdb'],
        },
    },
    'retention.get_giahan_kr7_data_to': {
        'title': 'Gia hạn KR7 theo tổ',
        'reason': 'API Excel KR7 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {
            'replacement': ['/api/giahan-ghtt-sty'],
        },
    },
    'quality.get_c15_chi_tiet_data': {
        'title': 'C1.5 chi tiết',
        'reason': 'C1.5 chưa có contract chi tiết ổn định trong report_history.db.',
        'required_display_contract': {
            'summary': ['ttvtst', 'dvvt_ttvt'],
            'details': ['group_by_doi_vien_thong', 'ty_le_dat'],
        },
    },
    'quality.get_shc_processing_report_data': {
        'title': 'SHC processing report',
        'reason': 'SHC processing vẫn dùng DB riêng, không thuộc runtime report_history.db.',
        'required_display_contract': {
            'summary': ['daily', 'monthly', 'team', 'user'],
        },
    },
    'quality.download_excel_chatluong': {
        'title': 'Chất lượng C1 file gốc',
        'reason': 'Tải file Excel C1 cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc contract mới',
        },
    },
    'inventory.api_tra_cuu_nhanh_vat_tu': {
        'title': 'Tra cứu nhanh vật tư',
        'reason': 'Nguồn vật tư kho cũ chưa có contract trong report_history.db.',
        'required_display_contract': {
            'table_shape': ['ma_vat_tu', 'ten_vat_tu', 'don_vi', 'so_luong'],
        },
    },
    'inventory.download_ton_kho_vat_tu': {
        'title': 'Tồn kho vật tư download',
        'reason': 'Tải file Excel vật tư cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc pipeline vật tư mới',
        },
    },
    'inventory.api_ton_kho_vat_tu': {
        'title': 'Tồn kho vật tư',
        'reason': 'Nguồn vật tư kho cũ chưa có contract trong report_history.db.',
        'required_display_contract': {
            'summary': ['tong_hop', 'thong_ke_theo_vat_tu'],
        },
    },
    'inventory.api_ton_kho_vat_tu_tot_thuong_dung': {
        'title': 'Tồn kho vật tư tốt thường dùng',
        'reason': 'Nguồn vật tư kho cũ chưa có contract trong report_history.db.',
        'required_display_contract': {
            'details': ['ton_theo_don_vi', 'gia_tri_ton'],
        },
    },
    'quangchudong.get_active_outages': {
        'title': 'Quang chủ động active',
        'reason': 'Quang chủ động vẫn phụ thuộc service/cache riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'summary': ['alerts', 'excluded', 'sources'],
        },
    },
    'quangchudong.get_recovered_alerts': {
        'title': 'Quang chủ động recovered',
        'reason': 'Quang chủ động vẫn phụ thuộc service/cache riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'details': ['recovered_alerts'],
        },
    },
    'quangchudong.get_wide_area_outages': {
        'title': 'Quang chủ động wide-area',
        'reason': 'Quang chủ động vẫn phụ thuộc service/cache riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'details': ['wide_area_alerts'],
        },
    },
    'quangchudong.get_exclusion_list': {
        'title': 'Quang chủ động exclusion',
        'reason': 'Quang chủ động vẫn phụ thuộc service/cache riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'details': ['exclusion_list'],
        },
    },
    'quangchudong.get_outage_stats': {
        'title': 'Quang chủ động stats',
        'reason': 'Quang chủ động vẫn phụ thuộc service/cache riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'summary': ['active_by_doi_vt', 'recovered_24h_by_doi_vt'],
        },
    },
    'sa_outage.api_su_co_sa_data': {
        'title': 'Sự cố SA',
        'reason': 'Sự cố SA vẫn phụ thuộc SQLite riêng ngoài report_history.db.',
        'required_display_contract': {
            'summary': ['dang_ton', 'da_clear_hom_nay'],
        },
    },
    'statistics.get_ticket_statistics': {
        'title': 'Ticket statistics',
        'reason': 'Thống kê ticket vẫn phụ thuộc brcd.db riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'summary': ['so_phieu_nhan', 'so_phieu_xu_ly_xong', 'so_phieu_ton'],
            'details': ['group_by_day_week_month'],
        },
    },
    'statistics.get_ticket_statistics_summary': {
        'title': 'Ticket statistics summary',
        'reason': 'Thống kê ticket vẫn phụ thuộc brcd.db riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'summary': ['today', 'this_week', 'this_month'],
        },
    },
    'statistics.get_ticket_history': {
        'title': 'Ticket history',
        'reason': 'Lịch sử ticket vẫn phụ thuộc brcd.db riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'details': ['ticket_history_by_ticket_id'],
        },
    },
    'statistics.get_ticket_trend': {
        'title': 'Ticket trend',
        'reason': 'Xu hướng ticket vẫn phụ thuộc brcd.db riêng, không thuộc report_history.db.',
        'required_display_contract': {
            'chart': ['labels', 'datasets'],
        },
    },
    'statistics.download_excel_statistics': {
        'title': 'Ticket statistics download',
        'reason': 'Xuất Excel thống kê ticket cũ đã bị ngắt khỏi dashv4.',
        'required_display_contract': {
            'source_of_truth': 'report_history.db hoặc contract mới',
        },
    },
}
