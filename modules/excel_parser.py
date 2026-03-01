"""
excel_parser.py - 2단계 엑셀 파싱 모듈
pandas 기반 (openpyxl load_workbook 스타일 호환 이슈 우회)
읽는 시트: 채권목록, 은행/카드, 보험
기본정보는 각 시트 오른쪽 열에서 추출
"""
import pandas as pd
from dataclasses import dataclass, field
from typing import List
import re


@dataclass
class PersonInfo:
    name: str = ""
    phone: str = ""
    ssn: str = ""
    address: str = ""
    cert_bank: str = ""

    @property
    def ssn_front(self) -> str:
        s = str(self.ssn) if self.ssn else ""
        if "-" in s:
            return s.split("-")[0].strip()
        return s[:6] if len(s) >= 6 else s

    @property
    def ssn_back(self) -> str:
        s = str(self.ssn) if self.ssn else ""
        if "-" in s:
            return s.split("-")[1].strip()
        return s[6:] if len(s) > 6 else ""

    @property
    def phone_clean(self) -> str:
        if not self.phone:
            return ""
        return re.sub(r'\(.*?\)', '', str(self.phone)).strip()


@dataclass
class CreditorEntry:
    name: str
    category: str       # "우선", "담보", "일반"
    seq: int = 0
    date: str = ""
    note: str = ""
    issued: str = ""
    method: str = ""


@dataclass
class BankCardEntry:
    name: str
    entry_type: str     # "은행" or "카드"
    seq: int = 0
    account: str = ""
    issued: str = ""
    method: str = ""


@dataclass
class InsuranceEntry:
    name: str           # 보험사명
    seq: int = 0
    contractor: str = ""
    insured: str = ""
    status: str = ""
    product: str = ""
    policy_no: str = ""
    refund: str = ""
    issued: str = ""
    method: str = ""


@dataclass
class ParsedExcel:
    person: PersonInfo = field(default_factory=PersonInfo)
    creditors: List[CreditorEntry] = field(default_factory=list)
    banks: List[BankCardEntry] = field(default_factory=list)
    cards: List[BankCardEntry] = field(default_factory=list)
    insurances: List[InsuranceEntry] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def delegation_creditors(self) -> list:
        return [c for c in self.creditors if c.category in ("담보", "일반")]

    @property
    def all_banks_cards(self) -> list:
        return self.banks + self.cards


def _s(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _i(val) -> int:
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0


def _parse_person_from_sidebar(df) -> PersonInfo:
    """시트 오른쪽 사이드바에서 기본정보 추출 (라벨-값 쌍 검색)"""
    info = PersonInfo()
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns) - 1):
            label = _s(df.iloc[row_idx, col_idx])
            value = _s(df.iloc[row_idx, col_idx + 1])
            if label == '이름' and value and value != '0':
                info.name = value
            elif label == '전화번호(통신사)' and value and value != '0':
                info.phone = value
            elif label == '주민번호' and value and value != '0':
                info.ssn = value
            elif label == '주소' and value and value != '0':
                info.address = value
            elif label == '인증서(은행)' and value:
                info.cert_bank = value
    return info


def _parse_creditor_sheet(df) -> list:
    """채권목록 시트 파싱 - 카테고리 라벨 기반으로 자동 감지"""
    creditors = []
    category = None

    for row_idx in range(len(df)):
        col_b = _s(df.iloc[row_idx, 1])  # 종류/카테고리
        col_c = _s(df.iloc[row_idx, 2])  # 순번
        col_d = _s(df.iloc[row_idx, 3])  # 채권기관

        # 카테고리 감지
        if '우선채권' in col_b:
            category = '우선'
        elif '담보채권' in col_b:
            category = '담보'
        elif '채권목록' == col_b.strip():
            category = '일반'
        elif '조회' in col_b and '채무' in col_b:
            break

        if category and col_d and col_c:
            creditors.append(CreditorEntry(
                name=col_d,
                category=category,
                seq=_i(col_c),
                date=_s(df.iloc[row_idx, 4]),
                note=_s(df.iloc[row_idx, 5]),
                issued=_s(df.iloc[row_idx, 6]),
                method=_s(df.iloc[row_idx, 7]),
            ))

    return creditors


def _parse_bank_card_sheet(df) -> tuple:
    """은행,카드 시트 파싱 - 헤더 감지 기반"""
    banks = []
    cards = []
    mode = None

    for row_idx in range(len(df)):
        col_b = _s(df.iloc[row_idx, 1])
        col_c = _s(df.iloc[row_idx, 2])

        if col_c == '은행':
            mode = 'bank'
            continue
        elif col_c == '카드사':
            mode = 'card'
            continue

        if mode == 'bank' and col_c:
            banks.append(BankCardEntry(
                name=col_c,
                entry_type="은행",
                seq=_i(col_b),
                account=_s(df.iloc[row_idx, 3]),
                issued=_s(df.iloc[row_idx, 4]),
                method=_s(df.iloc[row_idx, 5]),
            ))
        elif mode == 'card' and col_c:
            cards.append(BankCardEntry(
                name=col_c,
                entry_type="카드",
                seq=_i(col_b),
                issued=_s(df.iloc[row_idx, 4]),
            ))

    return banks, cards


def _parse_insurance_sheet(df) -> list:
    """보험 시트 파싱 - 헤더에서 열 위치 동적 감지"""
    insurances = []

    # 헤더 행 찾기 (순번 있는 행)
    header_row = None
    for row_idx in range(min(5, len(df))):
        for col_idx in range(min(15, len(df.columns))):
            if _s(df.iloc[row_idx, col_idx]) == '순번':
                header_row = row_idx
                break
        if header_row is not None:
            break

    if header_row is None:
        return insurances

    # 헤더에서 열 매핑 구축
    col_map = {}  # {'보험사': 3, '상태': 5, ...}
    for col_idx in range(len(df.columns)):
        val = _s(df.iloc[header_row, col_idx])
        if val:
            col_map[val] = col_idx

    # 필수 열
    c_seq = col_map.get('순번', 1)
    c_name = col_map.get('보험사', 2)
    c_contractor = col_map.get('계약자')
    c_insured = col_map.get('피보험자')
    c_status = col_map.get('상태')
    c_product = col_map.get('상품명')
    c_policyno = col_map.get('증권번호')
    c_refund = col_map.get('해지환급금')
    c_issued = col_map.get('발급유무')
    c_method = col_map.get('발급방법')

    def _safe(row_idx, col_idx):
        if col_idx is None or col_idx >= len(df.columns):
            return ""
        return _s(df.iloc[row_idx, col_idx])

    # 데이터 행 파싱 (첫 번째 순번 헤더 이후 ~ 합계행 전)
    found_first_section = False
    for row_idx in range(header_row + 1, len(df)):
        seq_val = _s(df.iloc[row_idx, c_seq])
        name_val = _safe(row_idx, c_name)

        # 두 번째 '순번' 나오면 배우자 영역 → 중단
        if seq_val == '순번':
            break

        if seq_val in ('', ) or '합계' in seq_val or '정리' in seq_val:
            continue

        if name_val and seq_val:
            insurances.append(InsuranceEntry(
                name=name_val,
                seq=_i(seq_val),
                contractor=_safe(row_idx, c_contractor),
                insured=_safe(row_idx, c_insured),
                status=_safe(row_idx, c_status),
                product=_safe(row_idx, c_product),
                policy_no=_safe(row_idx, c_policyno),
                refund=_safe(row_idx, c_refund),
                issued=_safe(row_idx, c_issued),
                method=_safe(row_idx, c_method),
            ))

    return insurances


def parse_excel(file_path_or_buffer) -> ParsedExcel:
    """엑셀 파일 전체 파싱"""
    result = ParsedExcel()

    try:
        sheets = pd.read_excel(
            file_path_or_buffer,
            sheet_name=None,
            header=None,
            engine='openpyxl'
        )
    except Exception as e:
        result.errors.append(f"엑셀 파일 열기 실패: {str(e)}")
        return result

    # ── 채권목록 ──
    if '채권목록' in sheets:
        df = sheets['채권목록']
        result.person = _parse_person_from_sidebar(df)
        if not result.person.name:
            result.errors.append("⚠️ 기본정보에 이름이 없습니다.")
        result.creditors = _parse_creditor_sheet(df)
        if not result.delegation_creditors:
            result.errors.append("⚠️ 채권목록에 담보/일반 채권자가 없습니다.")
    else:
        result.errors.append("⚠️ '채권목록' 시트를 찾을 수 없습니다.")

    # ── 은행, 카드 ──
    if '은행, 카드' in sheets:
        df = sheets['은행, 카드']
        result.banks, result.cards = _parse_bank_card_sheet(df)
        if not result.person.name:
            result.person = _parse_person_from_sidebar(df)
    else:
        result.errors.append("⚠️ '은행, 카드' 시트를 찾을 수 없습니다.")

    # ── 보험 ──
    if '보험' in sheets:
        df = sheets['보험']
        result.insurances = _parse_insurance_sheet(df)
        if not result.person.name:
            result.person = _parse_person_from_sidebar(df)
    else:
        result.errors.append("⚠️ '보험' 시트를 찾을 수 없습니다.")

    return result
