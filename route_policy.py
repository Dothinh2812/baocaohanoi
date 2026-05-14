from flask import current_app

from config import DISABLED_NONPAGE_ENDPOINTS, DISABLED_PAGE_ENDPOINTS


GENERIC_DISABLED_REASON = 'Route này đang bị khóa trong cấu hình instance hiện tại.'

PAGE_ACTIVE_KEYS = {
    'auth.redirect_sh_portal': 'sh_portal',
    'growth.page_ngungpsc': 'ngungpsc',
    'growth.page_tiepthi': 'tiepthi',
    'inventory.page_tam_dung_khoi_phuc': 'tam_dung_khoi_phuc',
    'inventory.page_ton_kho_vat_tu': 'ton_kho_vat_tu',
    'inventory.page_ton_kho_vat_tu_pto': 'ton_kho_vat_tu_pto',
    'inventory.page_ton_kho_vat_tu_qoi': 'ton_kho_vat_tu_qoi',
    'inventory.page_ton_kho_vat_tu_shi': 'ton_kho_vat_tu_shi',
    'inventory.page_ton_kho_vat_tu_sty': 'ton_kho_vat_tu_sty',
    'inventory.page_tong_hop_tien': 'tong_hop_tien',
    'inventory.page_tra_cuu_nhanh_vat_tu': 'tra_cuu_nhanh_vat_tu',
    'operations.page_brcd': 'brcd',
    'operations.page_cau_hinh_tu_dong': 'cau_hinh_tu_dong',
    'operations.page_kpi_nvkt_bchn': 'kpi_nvkt_bchn',
    'operations.page_pttb': 'pttb',
    'operations.page_tong_hop_bsc_kpi': 'tong_hop_bsc_kpi',
    'quality.page_chatluong': 'chatluong',
    'quality.page_i15': 'i15',
    'quality.page_i15k2': 'i15k2',
    'quality.page_shc_cts': 'shc_cts',
    'quality.page_shc_processing': 'shc_processing',
    'quangchudong.page_quangchudong': 'quangchudong',
    'quangchudong.page_quangchudong_nvkt': 'quangchudong',
    'retention.page_giahan': 'giahan',
    'sa_outage.page_su_co_sa': 'su_co_sa',
}


def _configured_set(name):
    return set(current_app.config.get(name) or set())


def _generic_feature(endpoint):
    active_page = PAGE_ACTIVE_KEYS.get(endpoint)
    if active_page:
        return {
            'kind': 'page',
            'title': endpoint,
            'active_page': active_page,
            'reason': GENERIC_DISABLED_REASON,
        }
    return {
        'kind': 'api',
        'title': endpoint,
        'reason': GENERIC_DISABLED_REASON,
        'required_display_contract': {},
    }


def get_disabled_feature(endpoint):
    if not endpoint:
        return None

    enabled = _configured_set('INSTANCE_ENABLED_ENDPOINTS')
    if endpoint in enabled:
        return None

    if endpoint in DISABLED_PAGE_ENDPOINTS:
        return {'kind': 'page', **DISABLED_PAGE_ENDPOINTS[endpoint]}
    if endpoint in DISABLED_NONPAGE_ENDPOINTS:
        return {'kind': 'api', **DISABLED_NONPAGE_ENDPOINTS[endpoint]}

    disabled = _configured_set('INSTANCE_DISABLED_ENDPOINTS')
    if endpoint in disabled:
        return _generic_feature(endpoint)

    return None


def is_endpoint_enabled(endpoint):
    return get_disabled_feature(endpoint) is None
