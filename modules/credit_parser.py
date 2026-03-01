"""
credit_parser.py - 신용조회 PDF → 구조화 데이터 파싱
Claude API를 사용하여 스캔 이미지 PDF에서 채권 정보를 추출합니다.
"""

import json
import base64
import fitz  # PyMuPDF
import anthropic

# ── 담보대출 분류 기준 ──
SECURED_LOAN_KEYWORDS = [
    "전세자금담보", "주택담보", "주택외부동산", "주택외 부동산",
    "부동산담보", "자동차담보", "예적금담보", "퇴직금담보",
    "유가증권담보", "신차할부", "중고차할부", "중고차할부"
]

UNSECURED_KEYWORDS = [
    "지급보증담보", "지급보증(보증서)담보", "기타담보",
    "보험계약대출", "현금서비스", "카드론", "신용대출",
    "지급보증", "기타(대출채권)", "신용대출(종합통장대출)"
]


def pdf_to_images(pdf_path, dpi=200):
    """PDF 파일을 페이지별 PNG 이미지(base64)로 변환"""
    doc = fitz.open(pdf_path)
    images = []
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        b64 = base64.standard_b64encode(img_bytes).decode("utf-8")
        images.append(b64)
    doc.close()
    return images


def parse_credit_pdf(pdf_path, api_key=None):
    """신용조회 PDF를 Claude API로 파싱하여 구조화된 데이터 반환
    
    Returns:
        dict: {
            "name": "홍길동",
            "id_number": "800201-*******",
            "cards": [{"기관명": "...", "발생일자": "..."}],
            "loans": [{"구분": "...", "대출종류": "...", "기관명": "...", "발생일자": "...", "금액": ...}],
            "changes": [{"기관명": "...", "채권구분": "...", "변동사유": "...", ...}]
        }
    """
    images = pdf_to_images(pdf_path)
    
    content = []
    for i, img_b64 in enumerate(images):
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img_b64}
        })
        content.append({
            "type": "text",
            "text": f"(페이지 {i+1})"
        })
    
    content.append({
        "type": "text",
        "text": """위 이미지들은 한국 신용정보조회서 PDF입니다. 다음 정보를 JSON으로 추출해주세요:

1. "name": 채무자 성명
2. "id_number": 주민등록번호
3. "cards": 개설·발급정보(카드) 배열. 각 항목: {"기관명": "...", "발생일자": "YYYY.MM.DD"}
4. "debt_list": 채무현황(채권자변동정보 조회서의 1.채무현황) 배열. 각 항목:
   {"순번": N, "구분": "개인대출정보/개인사업자대출", "대출종류": "신용대출(100) 등 원문 그대로", "기관명": "원문 그대로", "발생일자": "YYYY.MM.DD", "금액": 숫자(천원단위)}
5. "creditor_changes": 채권자변동현황 배열 (양수, 대위변제 등). 없으면 빈 배열.
   {"기관명": "...", "변동사유": "양수/대위변제/채무자변제 등", "변동일자": "..."}

주의사항:
- 채무현황은 '채권자변동정보 조회서'의 '1.채무현황' 테이블 데이터를 사용하세요 (가장 정확한 목록).
- 만약 '채권자변동정보 조회서'가 없으면 '(구)채권자변동정보 조회서'의 채무현황을 사용하세요.
- 기관명은 PDF에 적힌 그대로 적어주세요.
- 연체채권 변동현황에서 '채무자 변제'로 해제된 건은 creditor_changes에 포함하지 마세요.
- 반드시 유효한 JSON만 출력하세요. 설명 텍스트 없이 JSON만 출력하세요."""
    })

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
    
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}]
    )
    
    # JSON 파싱
    text = response.content[0].text
    # ```json ... ``` 감싸진 경우 처리
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0]
    elif "```" in text:
        text = text.split("```")[1].split("```")[0]
    
    return json.loads(text.strip())


def is_secured_loan(loan_type_str):
    """대출종류 문자열이 담보대출에 해당하는지 판별"""
    loan_type_str = loan_type_str.replace(" ", "")
    for kw in SECURED_LOAN_KEYWORDS:
        if kw.replace(" ", "") in loan_type_str:
            return True
    return False


def normalize_creditor_name(raw_name):
    """기관명 정규화
    - 농협 (지역) [xxx] → 농축협(xxx)
    - 농협은행 [NH카드분사 NH] → 농협은행카드
    - 삼성카드[금융팀] → 삼성카드
    등
    """
    name = raw_name.strip()
    
    # 농협 (지역) 패턴 → 농축협
    if "농협" in name and "(지역)" in name:
        # 농협 (지역) [남원 덕과] → 농축협(남원덕과)
        import re
        bracket_match = re.search(r'\[([^\]]+)\]', name)
        if bracket_match:
            branch = bracket_match.group(1).replace(" ", "")
            return f"농축협({branch})"
        return "농축협"
    
    # 농협은행 [NH카드분사 NH] → 농협은행카드
    if "농협은행" in name and "NH카드" in name:
        return "농협은행카드"
    
    # 일반 대괄호 제거: 삼성카드[금융팀] → 삼성카드
    import re
    name = re.sub(r'\[.*?\]', '', name).strip()
    
    # 알씨아이파이낸셜서비스코리아[본점] 등
    name = name.replace("(지역)", "").strip()
    
    return name


def classify_and_merge(parsed_data):
    """파싱된 데이터를 채권목록 규칙에 따라 분류 및 병합
    
    규칙:
    1. 담보대출 / 비담보 분류
    2. 같은 채권자의 비담보 대출은 한 행으로 합침 (내역사유 / 로 나열, 발생일자 가장 오래된 것)
    3. 담보대출끼리는 각각 별도 행
    4. 카드는 신용조회-카드개설정보로 분류
    
    Returns:
        dict: {
            "name": "...",
            "secured": [{"채권자명": "...", "내역사유": "...", "발생일자": "...", "조회방법": "..."}],
            "unsecured": [...],
            "cards": [...],
            "no_debt": [...]
        }
    """
    name = parsed_data.get("name", "")
    debt_list = parsed_data.get("debt_list", [])
    cards_raw = parsed_data.get("cards", [])
    changes = parsed_data.get("creditor_changes", [])
    
    secured = []     # 담보대출
    unsecured = {}   # 비담보: {정규화된 채권자명: {"reasons": set(), "earliest_date": "..."}}
    
    for item in debt_list:
        loan_type = item.get("대출종류", "")
        raw_name = item.get("기관명", "")
        norm_name = normalize_creditor_name(raw_name)
        date = item.get("발생일자", "")
        
        # 대출종류에서 간결한 내역사유 추출
        reason = _extract_reason(loan_type)
        
        if is_secured_loan(loan_type):
            # 담보대출: 각각 별도 행 (같은 채권자라도 내역사유가 다르면)
            # 단, 같은 채권자+같은 내역사유면 발생일자 오래된 건만 유지
            secured.append({
                "채권자명": norm_name,
                "내역사유": reason,
                "발생일자": date,
                "조회방법": "신용조회-대출정보"
            })
        else:
            # 비담보: 같은 채권자끼리 합침
            if norm_name not in unsecured:
                unsecured[norm_name] = {
                    "reasons": [],
                    "earliest_date": date
                }
            if reason not in unsecured[norm_name]["reasons"]:
                unsecured[norm_name]["reasons"].append(reason)
            # 가장 오래된 발생일자
            if date < unsecured[norm_name]["earliest_date"]:
                unsecured[norm_name]["earliest_date"] = date
    
    # 비담보 목록 생성
    unsecured_list = []
    for cred_name, info in unsecured.items():
        unsecured_list.append({
            "채권자명": cred_name,
            "내역사유": " / ".join(info["reasons"]),
            "발생일자": info["earliest_date"],
            "조회방법": "신용조회-대출정보"
        })
    

    
    # 카드 목록
    card_list = []
    for card in cards_raw:
        raw_name = card.get("기관명", "")
        norm_name = _normalize_card_name(raw_name)
        card_list.append({
            "채권자명": norm_name,
            "내역사유": "신용카드",
            "발생일자": card.get("발생일자", ""),
            "조회방법": "신용조회-카드개설정보"
        })
    
    return {
        "name": name,
        "secured": secured,
        "unsecured": unsecured_list,
        "cards": card_list,
        "no_debt": []  # 채권자변동에서 해제된 건 (추후 구현)
    }


def _extract_reason(loan_type_str):
    """대출종류 코드에서 간결한 내역사유 추출
    예: '신용대출(100)' → '신용대출'
        '주택담보대출(220)' → '주택담보대출'
        '지급보증(보증서) 담보대출(240)' → '지급보증담보대출'
    """
    import re
    # 코드번호 제거: (100), (220), (0031) 등
    cleaned = re.sub(r'\(\d+\)', '', loan_type_str).strip()
    # 불필요한 공백 정리
    cleaned = re.sub(r'\s+', '', cleaned)
    # (보증서) 같은 건 유지하되 간결하게
    cleaned = cleaned.replace("(보증서)", "")
    if not cleaned:
        cleaned = loan_type_str
    return cleaned


def _normalize_card_name(raw_name):
    """카드 기관명 정규화
    - 농협은행[NH카드분사 NH] → 농협은행카드
    - 농협 (지역) [NH카드본사조합카드] → 농축협카드
    - KB국민카드[전주] → KB국민카드
    - 하나카드[영업부] → 하나카드
    - 신한카드 (통합) [신한카드 (통합)] → 신한카드
    - 삼성카드[심사부] → 삼성카드
    """
    import re
    name = raw_name.strip()
    
    # 농협(지역) + 카드 → 농축협카드
    if "농협" in name and "(지역)" in name and ("카드" in name or "NH카드" in name):
        return "농축협카드"
    
    # 농협은행 + NH카드 → 농협은행카드
    if "농협은행" in name and "NH카드" in name:
        return "농협은행카드"
    
    # 대괄호 제거
    name = re.sub(r'\[.*?\]', '', name).strip()
    # (통합) 등 제거
    name = re.sub(r'\(통합\)', '', name).strip()
    # 리스 등 제거
    name = name.replace("리스", "").strip()
    # 앞뒤 공백
    name = name.strip()
    
    # KB국민카드 → KB국민카드 유지
    return name
