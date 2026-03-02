"""
creditor_matcher.py - 엑셀 채권기관명 → 기존 시스템 매칭
issue_manual.json / creditors.json 기준으로 매칭. 실패 시 원본 이름 그대로 사용.
"""
from dataclasses import dataclass
from typing import Optional
from modules.config_loader import (
    load_issue_manual, load_creditors, load_savings_banks,
    load_loan_companies, load_misc_companies, get_issue_info
)


# ── 별칭 매핑 테이블 ──
# 엑셀에 다양한 형태로 입력될 수 있는 이름 → 시스템 표준명
ALIAS_MAP = {
    # 은행
    "KB국민은행": "국민은행",
    "국민": "국민은행",
    "신한": "신한은행",
    "하나": "하나은행",
    "우리": "우리은행",
    "NH농협은행": "농협은행",
    "NH농협": "농협은행",
    "IBK기업은행": "기업은행",
    "IBK": "기업은행",
    "SC제일은행": "에스씨제일은행",
    "SC은행": "에스씨제일은행",
    "KDB산업은행": "케이디비산업은행",
    "산업은행": "케이디비산업은행",
    "BNK경남은행": "경남은행",
    "BNK부산은행": "부산은행",
    "DGB대구은행": "대구은행",
    "수협": "수협은행",
    "카카오": "카카오뱅크",
    "케이뱅크": "케이뱅크",
    "토스": "토스뱅크",
    # 카드
    "KB국민카드": "국민카드",
    "NH농협카드": "농협카드",
    "농축협카드": "농협카드",
    "IBK기업카드": "기업카드",
    "삼성": "삼성카드",
    "현대": "현대카드",
    "롯데": "롯데카드",
    "비씨": "비씨카드",
    "BC카드": "비씨카드",
    "BC": "비씨카드",
    # 저축은행
    "OK저축은행": "오케이저축은행",
    "SBI저축은행": "에스비아이저축은행",
    "JT저축은행": "제이티저축은행",
    "OSB저축은행": "오에스비저축은행",
    "HB저축은행": "에이치비저축은행",
    "DB저축은행": "디비저축은행",
    # 캐피탈
    "KB캐피탈": "케이비캐피탈",
    "현대캐피탈": "현대캐피탈",
    "NH캐피탈": "엔에이치농협캐피탈",
    "BNK캐피탈": "비엔케이캐피탈",
    "DGB캐피탈": "디지비캐피탈",
    # 통신사
    "SKT": "에스케이텔레콤",
    "SK텔레콤": "에스케이텔레콤",
    "KT": "케이티",
    "LGU+": "엘지유플러스",
    "LG유플러스": "엘지유플러스",
}


@dataclass
class MatchResult:
    excel_name: str         # 엑셀 원본 이름
    matched_key: str        # 매칭된 시스템 이름 (없으면 excel_name 그대로)
    matched: bool           # 매칭 성공 여부
    match_type: str         # "exact", "alias", "partial", "none"
    issue_info: dict = None # issue_manual 정보


def _build_all_names():
    """시스템에 등록된 모든 이름 수집"""
    names = set()
    # creditors.json
    creds = load_creditors()
    for key, val in creds.items():
        if isinstance(val, list):
            names.update(val)
    # savings_banks.json
    names.update(load_savings_banks().keys())
    # loan_companies.json
    names.update(load_loan_companies())
    # misc_companies.json
    names.update(load_misc_companies())
    # issue_manual.json (가장 포괄적)
    names.update(load_issue_manual().keys())
    return names


def match_creditor(excel_name: str) -> MatchResult:
    """엑셀 채권기관명을 시스템 이름에 매칭"""
    excel_name = excel_name.strip()
    if not excel_name:
        return MatchResult(excel_name, excel_name, False, "none")

    all_names = _build_all_names()

    # 0) 상호저축은행 → 저축은행 정규화
    normalized = excel_name.replace("상호저축은행", "저축은행")

    # 1) 정확 매칭
    if excel_name in all_names:
        return MatchResult(
            excel_name, excel_name, True, "exact",
            issue_info=get_issue_info(excel_name)
        )
    if normalized != excel_name and normalized in all_names:
        return MatchResult(
            excel_name, normalized, True, "alias",
            issue_info=get_issue_info(normalized)
        )

    # 2) 별칭 매칭
    if excel_name in ALIAS_MAP:
        key = ALIAS_MAP[excel_name]
        return MatchResult(
            excel_name, key, True, "alias",
            issue_info=get_issue_info(key)
        )
    if normalized != excel_name and normalized in ALIAS_MAP:
        key = ALIAS_MAP[normalized]
        return MatchResult(
            excel_name, key, True, "alias",
            issue_info=get_issue_info(key)
        )

    # 3) 부분 매칭 (긴 이름부터)
    for sys_name in sorted(all_names, key=len, reverse=True):
        if sys_name in excel_name or excel_name in sys_name:
            return MatchResult(
                excel_name, sys_name, True, "partial",
                issue_info=get_issue_info(sys_name)
            )

    # 4) 매칭 실패 → 원본 이름 그대로
    return MatchResult(
        excel_name, excel_name, False, "none",
        issue_info=get_issue_info(excel_name)  # 혹시 부분매칭 될 수도 있으니
    )


def match_all(names: list) -> list:
    """여러 이름을 한번에 매칭"""
    return [match_creditor(name) for name in names]
