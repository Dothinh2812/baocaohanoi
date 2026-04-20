# -*- coding: utf-8 -*-
"""Danh mục ID đơn vị (TTVT) cho các hệ thống báo cáo."""

from dataclasses import dataclass
from typing import Dict, Optional


@dataclass(frozen=True)
class UnitConfig:
    name: str
    id_14xxx: str  # Dùng cho ptrungtamid (KPI/Chỉ tiêu C)
    id_28xxxx: str  # Dùng cho vdonvi, pdonvi_id, vdv, vdonvi_id (Chi tiết, GHTT, Vật tư)
    onebss_ttvt_id: str  # Thường giống id_14xxx


# Danh sách các TTVT Hà Nội
UNITS: Dict[str, UnitConfig] = {
    "BA_DINH": UnitConfig(
        name="TTVT Ba Đình",
        id_14xxx="14329",
        id_28xxxx="1004430",
        onebss_ttvt_id="14329",
    ),
    "CAU_GIAY": UnitConfig(
        name="TTVT Cầu Giấy",
        id_14xxx="14320",
        id_28xxxx="284652",
        onebss_ttvt_id="14320",
    ),
    "DONG_ANH": UnitConfig(
        name="TTVT Đông Anh",
        id_14xxx="14321",
        id_28xxxx="284653",
        onebss_ttvt_id="14321",
    ),
    "DONG_DA": UnitConfig(
        name="TTVT Đống Đa",
        id_14xxx="14328",
        id_28xxxx="1004431",
        onebss_ttvt_id="14328",
    ),
    "GIA_LAM": UnitConfig(
        name="TTVT Gia Lâm",
        id_14xxx="14326",
        id_28xxxx="1003576",
        onebss_ttvt_id="14326",
    ),
    "GIAI_PHONG": UnitConfig(
        name="TTVT Giải Phóng",
        id_14xxx="14318",
        id_28xxxx="284649",
        onebss_ttvt_id="14318",
    ),
    "HA_DONG": UnitConfig(
        name="TTVT Hà Đông",
        id_14xxx="14323",
        id_28xxxx="284655",
        onebss_ttvt_id="14323",
    ),
    "HOAI_DUC": UnitConfig(
        name="TTVT Hoài Đức",
        id_14xxx="14332",
        id_28xxxx="1003577",
        onebss_ttvt_id="14332",
    ),
    "HOAN_KIEM": UnitConfig(
        name="TTVT Hoàn Kiếm",
        id_14xxx="14327",
        id_28xxxx="1003578",
        onebss_ttvt_id="14327",
    ),
    "HOANG_MAI": UnitConfig(
        name="TTVT Hoàng Mai",
        id_14xxx="14319",
        id_28xxxx="284651",
        onebss_ttvt_id="14319",
    ),
    "LONG_BIEN": UnitConfig(
        name="TTVT Long Biên",
        id_14xxx="14317",
        id_28xxxx="284650",
        onebss_ttvt_id="14317",
    ),
    "PHU_XUYEN": UnitConfig(
        name="TTVT Phú Xuyên",
        id_14xxx="14331",
        id_28xxxx="1003579",
        onebss_ttvt_id="14331",
    ),
    "SOC_SON": UnitConfig(
        name="TTVT Sóc Sơn",
        id_14xxx="14330",
        id_28xxxx="1003580",
        onebss_ttvt_id="14330",
    ),
    "SON_TAY": UnitConfig(
        name="TTVT Sơn Tây",
        id_14xxx="14324",
        id_28xxxx="284656",
        onebss_ttvt_id="14324",
    ),
    "TAY_HO": UnitConfig(
        name="TTVT Tây Hồ",
        id_14xxx="14325",
        id_28xxxx="284657",
        onebss_ttvt_id="14325",
    ),
    "THACH_THAT": UnitConfig(
        name="TTVT Thạch Thất",
        id_14xxx="14333",
        id_28xxxx="1003581",
        onebss_ttvt_id="14333",
    ),
    "THANH_TRI": UnitConfig(
        name="TTVT Thanh Trì",
        id_14xxx="14322",
        id_28xxxx="284654",
        onebss_ttvt_id="14322",
    ),
    "TU_LIEM": UnitConfig(
        name="TTVT Từ Liêm",
        id_14xxx="14334",
        id_28xxxx="1003582",
        onebss_ttvt_id="14334",
    ),
}


def get_unit(query: str) -> Optional[UnitConfig]:
    """Tìm đơn vị theo key hoặc tên gần đúng."""
    if not query:
        return None

    query_upper = query.upper().replace("-", "_").replace(" ", "_")
    if query_upper in UNITS:
        return UNITS[query_upper]

    # Tìm theo tên tiếng Việt
    for unit in UNITS.values():
        if query.lower() in unit.name.lower():
            return unit

    return None
