# 📋 WE Legal Automation - 회생파산 서류 자동 생성기

채권사 여러 개 선택 → 위임장 + 신청서 + 신분증 묶음 PDF 일괄 생성

## 🚀 처음 실행하는 법

### 1단계: Python 설치 (한번만)
1. https://www.python.org/downloads/ 접속
2. "Download Python" 클릭하여 설치
3. ⚠️ 설치할 때 **"Add python.exe to PATH"** 반드시 체크!

### 2단계: 이 폴더에서 실행
1. `시작.bat` 파일을 더블클릭
2. 처음엔 패키지 설치가 진행됨 (1~2분)
3. 브라우저가 자동으로 열리면서 앱이 실행됨

### 또는 수동 실행
```
pip install -r requirements.txt
streamlit run app.py
```

## 📁 폴더 구조
```
we-legal-app/
├── app.py                  ← 메인 앱
├── 시작.bat                ← 더블클릭으로 실행
├── requirements.txt        ← 필요 패키지
├── config/
│   ├── staff.json          ← 수임인(담당자) 목록
│   ├── creditors.json      ← 채권사 목록
│   └── doc_matrix.json     ← 기관유형별 필요서류
├── templates/
│   ├── pdf/                ← 빈 양식 PDF (90개)
│   └── coords/             ← 좌표 JSON (양식별)
├── id_cards/
│   ├── agents/             ← 수임인 신분증 사본
│   └── clients/            ← 위임인 신분증 (앱에서 업로드)
└── output/                 ← 생성된 PDF
```

## ➕ 채권사/양식 추가하는 법

### 수임인 추가
`config/staff.json`에 추가:
```json
{
  "name": "홍길동",
  "id_front": "900101",
  "birth": "900101",
  "phone": "010-0000-0000",
  "address": "서울시 ...",
  "fax": "",
  "id_card_file": "홍길동_신분증.pdf"
}
```

### 채권사 추가
`config/creditors.json`에 추가:
```json
{
  "name": "새 은행",
  "type": "은행",
  "id_format": "front_6",
  "form_template": "새은행_부채.pdf",
  "coord_file": "새은행_부채_coords.json",
  "has_own_form": true
}
```

### 양식 좌표 추가
1. 빈 양식 PDF를 `templates/pdf/`에 넣기
2. Google AI Studio에서 좌표 추출
3. 좌표 JSON을 `templates/coords/`에 저장
