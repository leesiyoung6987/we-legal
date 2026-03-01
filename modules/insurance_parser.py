"""
insurance_parser.py - 보험조회 PDF 파싱
보험내역조회 PDF에서 보험계약 정보를 추출하여 구조화된 데이터로 반환.

필터 기준:
- 계약자 건만 (피보험자 제외)
- 계약기간 1년 이하 제외
- 서울보증 제외
"""
import re
from datetime import datetime
from dataclasses import dataclass, field


@dataclass
class InsuranceEntry:
    """보험 계약 1건"""
    company: str          # 보험회사명
    product: str          # 상품명
    policy_no: str        # 증권(계좌)번호
    status: str           # 계약상태: 유지(정상), 휴면, 실효, 소멸(해약포함), 만기
    relation: str         # 계약관계: 보험계약자, 피보험자
    start_date: str       # 시작일 YYYY-MM-DD
    end_date: str         # 종료일 YYYY-MM-DD
    tel: str = ""         # 전화번호

    @property
    def duration_years(self):
        """계약기간 (년)"""
        try:
            s = datetime.strptime(self.start_date, "%Y-%m-%d")
            e = datetime.strptime(self.end_date, "%Y-%m-%d")
            return (e - s).days / 365.25
        except:
            return 99  # 파싱 실패 시 포함

    @property
    def is_contractor(self):
        return "계약자" in self.relation

    @property
    def doc_type(self):
        """서류 종류 결정"""
        if self.status in ("유지(정상)", "유지", "휴면"):
            return "예상해지환급금증명서"
        else:  # 실효, 소멸(해약포함), 만기
            return "해지확인서"

    @property
    def status_short(self):
        """상태 짧은 표기"""
        m = {"유지(정상)": "유지", "소멸(해약포함)": "소멸"}
        return m.get(self.status, self.status)


@dataclass
class InsuranceParseResult:
    """파싱 결과"""
    person_name: str = ""
    person_ssn: str = ""
    all_entries: list = field(default_factory=list)    # 전체 파싱 결과
    filtered: list = field(default_factory=list)       # 필터 적용 후
    excluded: list = field(default_factory=list)       # 제외된 건
    errors: list = field(default_factory=list)


# 제외 보험사
EXCLUDE_COMPANIES = {"서울보증"}


def parse_insurance_pdf(pdf_bytes) -> InsuranceParseResult:
    """보험조회 PDF 파싱"""
    import fitz

    result = InsuranceParseResult()

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        result.errors.append(f"PDF 열기 실패: {e}")
        return result

    # 전체 텍스트 추출 (보험가입내역 부분만)
    full_text = ""
    for page in doc:
        text = page.get_text()
        # "안내사항" 이후는 불필요
        if "안내사항" in text:
            text = text[:text.index("안내사항")]
        full_text += text + "\n"
    doc.close()

    # 인적사항 추출
    _extract_person(full_text, result)

    # 보험 계약 추출
    _extract_entries(full_text, result)

    # 필터 적용
    for entry in result.all_entries:
        reasons = []
        if not entry.is_contractor:
            reasons.append("피보험자")
        if entry.duration_years <= 1.0:
            reasons.append(f"기간 {entry.duration_years:.1f}년(1년이하)")
        if any(exc in entry.company for exc in EXCLUDE_COMPANIES):
            reasons.append("서울보증 제외")
        if "휴대폰" in entry.product:
            reasons.append("휴대폰보험 제외")

        if reasons:
            result.excluded.append((entry, ", ".join(reasons)))
        else:
            result.filtered.append(entry)

    return result


def _extract_person(text, result):
    """인적사항 추출"""
    # 주민등록번호 패턴
    ssn_match = re.search(r'(\d{6}-\d{7})', text)
    if ssn_match:
        result.person_ssn = ssn_match.group(1)

    # 이름: 주민번호 앞에 있는 한글 이름
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if re.search(r'\d{6}-\d{7}', line):
            # 이전 줄에서 이름 찾기
            for j in range(max(0, i-3), i):
                name_match = re.match(r'^([가-힣]{2,4})$', lines[j].strip())
                if name_match:
                    result.person_name = name_match.group(1)
                    break
            break


def _extract_entries(text, result):
    """보험 계약 항목 추출"""
    lines = text.split('\n')
    lines = [l.strip() for l in lines if l.strip()]

    # 테이블 헤더 줄 제거 (반복되는 헤더)
    # 헤더 블록 감지 및 제거
    cleaned = []
    i = 0
    while i < len(lines):
        # "보험가입내역 조회결과" 발견 시 헤더 블록 스킵
        if lines[i] in ("보험가입내역 조회결과 내용입니다", "보험가입내역 조회결과"):
            # "종료일" 줄까지 스킵
            j = i + 1
            while j < len(lines) and j < i + 20:
                if lines[j] == "종료일":
                    j += 1  # "종료일" 포함 스킵
                    break
                j += 1
            i = j
            continue
        cleaned.append(lines[i])
        i += 1

    lines = cleaned
    current_company = None
    i = 0

    while i < len(lines):
        line = lines[i]

        # "안내사항" 이후 중단
        if line.startswith("안내사항"):
            break

        # 보험회사명 감지
        if _is_company_name(line):
            current_company = line
            i += 1
            continue

        # "보험계약" 구분 시작
        if line == "보험계약" and current_company:
            entry = _parse_one_entry(lines, i, current_company)
            if entry:
                result.all_entries.append(entry)
                i = entry._end_idx
                continue

        i += 1


def _is_company_name(line):
    """보험회사명인지 판별 - 짧은 회사명만 (상품명 제외)"""
    if not line:
        return False
    # 너무 긴 것은 상품명 (회사명은 보통 10자 이내)
    if len(line) > 12:
        return False
    # 테이블 헤더 키워드 제외
    skip = {"보험회사", "구분", "상품명", "보험가입내역", "보험계약", "총",
            "기본사항", "접수번호", "접수일자", "조회", "안내사항", "보험기간",
            "휴대폰보험"}
    if line in skip:
        return False
    # 날짜 제외
    if re.match(r'^\d{4}-\d{2}-\d{2}$', line):
        return False
    # 숫자만 제외
    if re.match(r'^[\d\-]+$', line):
        return False

    # 보험회사 정확한 매칭
    KNOWN_COMPANIES = {
        "메리츠화재", "삼성화재", "삼성생명", "삼성화재보험",
        "한화손해보험", "한화생명", "한화손해",
        "DB손해보험", "DB생명보험", "DGB생명보험",
        "현대해상", "현대해상보험",
        "교보생명", "교보생명보험",
        "동양생명", "동양생명보험",
        "흥국생명보험", "흥국화재보험",
        "라이나생명보험", "라이나손해보험",
        "롯데손해보험",
        "AIA생명보험", "AIG손해보험", "AXA손해보험",
        "NH농협생명보험", "NH농협손해보험", "농협생명", "농협손해보험",
        "MG손해보험",
        "KB손해보험", "KDB생명보험",
        "푸르덴셜생명보험", "푸본현대생명보험",
        "악사손해보험",
        "에이비엘생명보험", "에이스손해보험",
        "서울보증",
        "우체국보험", "우체국",
        "새마을금고보험",
        "신한라이프", "신한라이프보험",
        "하나생명", "하나손해보험",
        "미래에셋생명보험",
        "IM라이프", "캐롯", "쳐브라이프생명",
        "메리츠화재보험", "메트라이프생명보험",
    }
    if line in KNOWN_COMPANIES:
        return True

    # 짧은 패턴 매칭 (5자 이하이면서 보험사 접미사)
    if len(line) <= 8:
        if re.match(r'^[가-힣A-Z]+화재$', line):
            return True
        if re.match(r'^[가-힣A-Z]+생명$', line):
            return True
        if re.match(r'^[가-힣A-Z]+손해$', line):
            return True
        if re.match(r'^[가-힣A-Z]+보증$', line):
            return True

    return False


def _parse_one_entry(lines, start_idx, company):
    """'보험계약' 줄부터 하나의 항목 파싱"""
    # start_idx는 "보험계약" 줄
    # 이후 줄들에서: 상품명, 증권번호, 상태, 관계, 시작일, 종료일, 담당점포, 전화번호

    DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    STATUS_KEYWORDS = {"유지(정상)", "휴면", "실효", "소멸(해약포함)", "만기"}
    RELATION_KEYWORDS = {"보험계약자", "피보험자"}

    # 다음 "보험계약" 또는 보험회사명이 나올 때까지 수집
    collected = []
    i = start_idx + 1  # "보험계약" 다음 줄부터

    while i < len(lines) and len(collected) < 15:
        line = lines[i]
        if line == "보험계약":
            break
        if _is_company_name(line):
            break
        if line.startswith("안내사항"):
            break
        collected.append(line)
        i += 1

    if len(collected) < 4:
        return None

    # 분류
    product_parts = []
    policy_no = None
    status = None
    relation = None
    dates = []
    tel = ""

    for c in collected:
        if c in STATUS_KEYWORDS:
            status = c
        elif c in RELATION_KEYWORDS:
            relation = c
        elif DATE_RE.match(c):
            dates.append(c)
        elif status is None and relation is None and not dates:
            # 상태/관계/날짜 이전 → 상품명 또는 증권번호
            # 증권번호: 숫자+하이픈 or 영문+숫자 패턴
            if re.match(r'^[A-Z0-9].*[\-].*\d', c) or re.match(r'^[A-Z]{2}\d+', c) or re.match(r'^\d{10,}', c):
                policy_no = c
            else:
                product_parts.append(c)
        elif status and relation and len(dates) >= 2:
            # 날짜 이후 → 담당점포, 전화번호
            if re.match(r'[\d\-\s]{8,}', c):
                tel = c.replace(' ', '-')

    product = " ".join(product_parts).strip()
    if not product:
        product = "(상품명 없음)"

    # 증권번호가 상품명에 섞여있을 수 있음
    if not policy_no and len(collected) > 1:
        # 두 번째 줄이 증권번호일 가능성
        for c in collected:
            if c not in STATUS_KEYWORDS and c not in RELATION_KEYWORDS and not DATE_RE.match(c):
                if re.search(r'\d{5,}', c) and c != product:
                    policy_no = c
                    if c in product_parts:
                        product_parts.remove(c)
                        product = " ".join(product_parts).strip()
                    break

    if not policy_no:
        policy_no = ""
    if not status:
        return None
    if not relation:
        return None

    start_date = dates[0] if len(dates) >= 1 else ""
    end_date = dates[1] if len(dates) >= 2 else ""

    entry = InsuranceEntry(
        company=company,
        product=product,
        policy_no=policy_no,
        status=status,
        relation=relation,
        start_date=start_date,
        end_date=end_date,
        tel=tel,
    )
    entry._end_idx = i  # 파싱이 끝난 위치
    return entry


def filter_by_scope(entries, scope="전체"):
    """법률사무소 보험 설정에 따라 필터

    scope: "유지" → 유지/휴면만, "전체" → 전부
    """
    if scope == "유지":
        return [e for e in entries if e.status in ("유지(정상)", "유지", "휴면")]
    else:  # 전체
        return entries


def group_by_company(entries):
    """보험사별 그룹핑

    Returns:
        {보험사명: [InsuranceEntry, ...]}
    """
    groups = {}
    for e in entries:
        if e.company not in groups:
            groups[e.company] = []
        groups[e.company].append(e)
    return groups
