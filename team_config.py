#!/usr/bin/env python3
"""
Team Configuration Module
Single source of truth for all team metadata
"""

from typing import List, Dict, Optional
from dataclasses import dataclass

@dataclass
class Team:
    """Team metadata - Simplified"""
    id: str                    # 'ToKT_ThachThat', 'ToKT_HoaLac', etc. (onebss_code format)
    short_name: str           # 'Thạch Thất', 'Hòa Lạc' (display name)
    zalo_thread_id: str       # '7835396852590969049'
    telegram_chat_id: str     # '-#'
    active: bool              # True/False
    team_type: str            # 'BRCD' or 'PTTB'
    order: int                # Display order

# ========== TEAM DEFINITIONS ==========

BRCD_TEAMS = [
    Team(
        id='ToKT_PhucTho',
        short_name='Phúc Thọ',
        zalo_thread_id='6780971089121842303',
        telegram_chat_id='-4616062001',
        active=True,
        team_type='BRCD',
        order=1
    ),
    Team(
        id='ToKT_SonTay',
        short_name='Sơn Tây',
        zalo_thread_id='6337217534995887511',
        telegram_chat_id='-4654883926',
        active=True,
        team_type='BRCD',
        order=2
    ),
    Team(
        id='ToKT_QuangOai',
        short_name='Quảng Oai',
        zalo_thread_id='5364152493553904404',
        telegram_chat_id='-4734554771',
        active=True,
        team_type='BRCD',
        order=3
    ),
    Team(
        id='ToKT_SuoiHai',
        short_name='Suối Hai',
        zalo_thread_id='6085297980620830486',
        telegram_chat_id='-4607586268',
        active=True,
        team_type='BRCD',
        order=4
    ),
]

PTTB_TEAMS = BRCD_TEAMS + [
    Team(
        id='ToKT_BaVi',
        short_name='Ba Vì',
        zalo_thread_id='',  # Not available
        telegram_chat_id='-4735594488',
        active=False,
        team_type='PTTB',
        order=5
    ),
]

# ========== HELPER FUNCTIONS ==========

def get_active_teams(team_type: str = 'BRCD') -> List[Team]:
    """Lấy danh sách team đang active"""
    teams = BRCD_TEAMS if team_type == 'BRCD' else PTTB_TEAMS
    return [t for t in teams if t.active]

def get_team_by_id(team_id: str, team_type: str = 'BRCD') -> Optional[Team]:
    """Lấy team theo ID"""
    teams = BRCD_TEAMS if team_type == 'BRCD' else PTTB_TEAMS
    for team in teams:
        if team.id == team_id:
            return team
    return None

def get_team_by_short_name(short_name: str, team_type: str = 'BRCD') -> Optional[Team]:
    """Lấy team theo tên ngắn"""
    teams = BRCD_TEAMS if team_type == 'BRCD' else PTTB_TEAMS
    for team in teams:
        if team.short_name == short_name:
            return team
    return None

# ========== MAPPING GENERATORS ==========

def get_shortname_to_id_mapping(team_type: str = 'BRCD') -> Dict[str, str]:
    """
    Short name → ID
    Example: {'Thạch Thất': 'ToKT_ThachThat'}
    """
    teams = get_active_teams(team_type)
    return {t.short_name: t.id for t in teams}

def get_id_to_shortname_mapping(team_type: str = 'BRCD') -> Dict[str, str]:
    """
    ID → Short name
    Example: {'ToKT_ThachThat': 'Thạch Thất'}
    """
    teams = get_active_teams(team_type)
    return {t.id: t.short_name for t in teams}

# ========== BACKWARD COMPATIBILITY (deprecated) ==========
# These functions are kept for backward compatibility but are deprecated
# Use get_id_to_shortname_mapping instead

def get_id_to_fullname_mapping(team_type: str = 'BRCD') -> Dict[str, str]:
    """
    DEPRECATED: Use get_id_to_shortname_mapping instead
    ID → Short name (was full name)
    Example: {'ToKT_ThachThat': 'Thạch Thất'}
    """
    return get_id_to_shortname_mapping(team_type)

def get_fullname_to_id_mapping(team_type: str = 'BRCD') -> Dict[str, str]:
    """
    DEPRECATED: Use get_shortname_to_id_mapping instead
    Short name → ID (was full name → ID)
    Example: {'Thạch Thất': 'ToKT_ThachThat'}
    """
    return get_shortname_to_id_mapping(team_type)

def get_location_thread_mapping() -> Dict[str, str]:
    """
    Location → Zalo thread ID
    Example: {'ToKT_ThachThat': '7835396852590969049'}
    Maps team ID to Zalo thread ID
    """
    mapping = {}
    for team in BRCD_TEAMS + PTTB_TEAMS:
        if team.zalo_thread_id:
            mapping[team.id] = team.zalo_thread_id

    # Add default thread
    mapping['default'] = '4266181895406444369'
    return mapping

def get_location_chat_mapping() -> Dict[str, str]:
    """
    Location → Telegram chat ID
    Example: {'ToKT_ThachThat': '-#'}
    Maps team ID to Telegram chat ID
    """
    mapping = {}
    for team in BRCD_TEAMS + PTTB_TEAMS:
        if team.telegram_chat_id:
            mapping[team.id] = team.telegram_chat_id
    return mapping

def get_active_team_short_names(team_type: str = 'BRCD') -> List[str]:
    """
    Lấy danh sách tên ngắn của team active
    Example: ['Sơn Tây', 'Suối Hai', 'Quảng Oai', 'Phúc Thọ']
    """
    teams = get_active_teams(team_type)
    return [t.short_name for t in sorted(teams, key=lambda x: x.order)]

def get_active_team_ids(team_type: str = 'BRCD') -> List[str]:
    """
    Lấy danh sách ID của team active
    Example: ['sontay', 'suoihai', 'quangoai', 'phuctho']
    """
    teams = get_active_teams(team_type)
    return [t.id for t in sorted(teams, key=lambda x: x.order)]

# ========== VALIDATION ==========

def validate_teams():
    """Validate team configuration"""
    errors = []

    # Check duplicates
    all_teams = BRCD_TEAMS + [t for t in PTTB_TEAMS if t not in BRCD_TEAMS]
    ids = [t.id for t in all_teams]
    if len(ids) != len(set(ids)):
        errors.append("Duplicate team IDs found")

    # Check required fields
    for team in all_teams:
        if not team.id or not team.short_name:
            errors.append(f"Team {team.id} missing required fields")

    if errors:
        print("❌ Team configuration errors:")
        for err in errors:
            print(f"  - {err}")
        return False
    else:
        print("✅ Team configuration valid")
        return True

# ========== TEST ==========

if __name__ == "__main__":
    print("=" * 60)
    print("Team Configuration Test")
    print("=" * 60)

    validate_teams()

    print("\n📋 BRCD Teams (Active):")
    for team in get_active_teams('BRCD'):
        print(f"  {team.order}. {team.short_name} ({team.id})")

    print("\n📋 PTTB Teams (Active):")
    for team in get_active_teams('PTTB'):
        print(f"  {team.order}. {team.short_name} ({team.id})")

    print("\n🔄 ID → Short Name Mapping (BRCD):")
    for k, v in get_id_to_shortname_mapping('BRCD').items():
        print(f"  {k} → {v}")

    print("\n🔄 Short Name → ID Mapping (PTTB):")
    for k, v in get_shortname_to_id_mapping('PTTB').items():
        print(f"  {k} → {v}")

    print("\n🔄 Location Thread Mapping (Zalo):")
    for k, v in get_location_thread_mapping().items():
        if v:  # Only show non-empty
            print(f"  {k} → {v}")

    print("\n🔄 Location Chat Mapping (Telegram):")
    for k, v in get_location_chat_mapping().items():
        if v:  # Only show non-empty
            print(f"  {k} → {v}")
