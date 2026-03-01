"""
config_loader.py - 설정 파일 및 경로 관리
모든 JSON 설정, 경로를 여기서 로드.
새 설정 추가 시 이 파일만 수정.
"""

import json
from pathlib import Path

# ── 기본 경로 ──
BASE_DIR = Path(__file__).parent.parent
CONFIG_DIR = BASE_DIR / "config"
TEMPLATES_DIR = BASE_DIR / "templates"
PDF_DIR = TEMPLATES_DIR / "pdf"
COORDS_DIR = TEMPLATES_DIR / "coords"
ID_CARDS_DIR = BASE_DIR / "id_cards" / "agents"
OUTPUT_DIR = BASE_DIR / "output"

# 폴더 자동 생성
for d in [CONFIG_DIR, PDF_DIR, COORDS_DIR, ID_CARDS_DIR, OUTPUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)


def _load_json(filename):
    """config 폴더에서 JSON 파일 로드"""
    path = CONFIG_DIR / filename
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_settings():
    """settings.json 로드 (앱 설정, 서류 옵션 등)"""
    return _load_json("settings.json")


def load_staff():
    """staff.json 로드 (수임인 목록)"""
    return _load_json("staff.json")


def load_creditors():
    """creditors.json 로드 (채권사 목록)"""
    return _load_json("creditors.json")


def get_all_creditor_names():
    """전체 채권사명 리스트 (모든 유형 합쳐서 정렬)"""
    data = load_creditors()
    names = []
    for key, val in data.items():
        if isinstance(val, list):
            names.extend(val)
    return sorted(set(names))


def load_doc_matrix():
    """doc_matrix.json 로드 (유형별 필요서류)"""
    return _load_json("doc_matrix.json")


def load_form_mapping():
    """form_mapping.json 로드 (채권사별 신청서 양식 매핑)"""
    return _load_json("form_mapping.json")


def load_savings_banks():
    """savings_banks.json 로드 (저축은행 DB)"""
    data = _load_json("savings_banks.json")
    data.pop("_comment", None)
    data.pop("_example", None)
    return data


def load_issue_manual():
    """issue_manual.json 로드 (채권사별 발급 매뉴얼)"""
    return _load_json("issue_manual.json")


def load_insurance_manual():
    """insurance_manual.json 로드 (보험사별 발급 매뉴얼)"""
    return _load_json("insurance_manual.json")


def get_insurance_info(company_name):
    """보험사명으로 발급 매뉴얼 조회 (부분매칭 지원)

    Returns:
        {"유형", "발급방법", "고객센터", "내선번호", "홈페이지경로", "특이사항"} 또는 None
    """
    manual = load_insurance_manual()
    # 1차: 정확히 일치
    if company_name in manual:
        return manual[company_name]
    # 2차: 포함 매칭 (PDF에서 "메리츠화재" → manual "메리츠화재보험")
    for name in sorted(manual.keys(), key=len, reverse=True):
        if name in company_name or company_name in name:
            return manual[name]
    return None


def load_law_firms():
    """law_firms.json 로드 (거래처 법률사무소/법무사 목록)"""
    return _load_json("law_firms.json")


def get_issue_info(creditor_name):
    """채권사명으로 발급 매뉴얼 조회
    
    Returns:
        {"유형", "발급방법", "고객센터", "팩스번호", "주소", "특이사항"} 또는 None
    """
    manual = load_issue_manual()
    # 1차: 정확히 일치
    if creditor_name in manual:
        return manual[creditor_name]
    # 2차: 포함 매칭 (긴 이름부터)
    for name in sorted(manual.keys(), key=len, reverse=True):
        if name in creditor_name or creditor_name in name:
            return manual[name]
    return None


def load_bundle_types():
    """bundle_types.json 로드 (유형별 추가 서류 설정)"""
    data = _load_json("bundle_types.json")
    data.pop("_comment", None)
    return data


def load_loan_companies():
    """loan_companies.json 로드 (대부업체 목록)"""
    data = _load_json("loan_companies.json")
    return data.get("companies", [])


def load_misc_companies():
    """misc_companies.json 로드 (기타 회사 목록: 통신사, 공공기관 등)"""
    data = _load_json("misc_companies.json")
    return data.get("companies", [])


def get_bundle_type(creditor_name):
    """채권사명으로 번들 타입 조회
    
    match_list_file: JSON 파일의 companies 리스트에서 매칭
    match_keywords: 채권사명에 키워드 포함 여부
    
    Returns:
        (type_name, config) 또는 (None, None)
    """
    types = load_bundle_types()
    for type_name, config in types.items():
        # 1차: 목록 파일 매칭
        list_file = config.get("match_list_file")
        if list_file:
            list_data = _load_json(list_file)
            companies = list_data.get("companies", [])
            if creditor_name in companies:
                return type_name, config
        # 2차: 키워드 매칭
        keywords = config.get("match_keywords", [])
        for kw in keywords:
            if kw in creditor_name:
                return type_name, config
    return None, None


def get_savings_bank_info(creditor_name):
    """채권사명으로 저축은행 정보 조회
    
    Returns:
        {"tel": "...", "fax": "...", "branch": "..."} 또는 None
    """
    banks = load_savings_banks()
    # 1차: 정확히 일치
    if creditor_name in banks:
        return banks[creditor_name]
    # 2차: 포함 매칭 (긴 이름부터)
    for name in sorted(banks.keys(), key=len, reverse=True):
        if name in creditor_name:
            return banks[name]
    return None


def get_form_info(creditor_name, doc_type):
    """채권사명 + 서류종류로 신청서 양식 정보 조회
    
    1차: 정확히 일치
    2차: 채권사명에 키워드 포함 여부 (양우신협 → 신협, 전주새마을금고 → 새마을금고)
    
    Returns:
        리스트 [{"template": "...", "coords": "..."}, ...] 또는 None
    """
    mapping = load_form_mapping()
    
    def _to_list(val):
        if isinstance(val, list):
            return val
        elif isinstance(val, dict):
            return [val]
        return None
    
    # 1차: 정확히 일치
    bank_forms = mapping.get(creditor_name, {})
    if doc_type in bank_forms:
        return _to_list(bank_forms[doc_type])
    
    # 2차: 키워드 포함 매칭 (긴 키워드부터 먼저 검사)
    keywords = sorted(mapping.keys(), key=len, reverse=True)
    for keyword in keywords:
        if keyword == "_comment":
            continue
        if keyword in creditor_name and keyword != creditor_name:
            bank_forms = mapping.get(keyword, {})
            if doc_type in bank_forms:
                return _to_list(bank_forms[doc_type])
    
    return None


def load_coords(form_name):
    """좌표 JSON 로드 (양식별, 하위폴더 자동 탐색)"""
    # 루트에서 먼저 찾기
    path = COORDS_DIR / f"{form_name}_coords.json"
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    # 하위폴더 탐색
    for p in COORDS_DIR.rglob(f"{form_name}_coords.json"):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def get_template_path(filename):
    """양식 PDF 경로 반환 (하위폴더 자동 탐색)"""
    # 루트에서 먼저 찾기
    path = PDF_DIR / filename
    if path.exists():
        return path
    # 하위폴더 탐색
    for p in PDF_DIR.rglob(filename):
        return p
    return PDF_DIR / filename


def get_id_card_path(filename):
    """수임인 신분증 경로 반환"""
    return ID_CARDS_DIR / filename


def get_doc_options(settings):
    """settings에서 서류 드롭다운 옵션 리스트 추출"""
    return [opt["label"] for opt in settings.get("doc_options", [])]


def get_doc_value_map(settings):
    """label → value 매핑 딕셔너리"""
    return {opt["label"]: opt["value"] for opt in settings.get("doc_options", [])}


def get_needs_date(settings):
    """기간 설정이 필요한 서류 value 목록"""
    return settings.get("needs_date_range", [])


def get_needs_date_labels(settings):
    """기간 설정이 필요한 서류 label 목록"""
    value_set = set(get_needs_date(settings))
    return [opt["label"] for opt in settings.get("doc_options", []) if opt["value"] in value_set]
