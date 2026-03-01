"""
좌표 미세조정 도구 v14 (coord_tuner.py)
─────────────────────────────────────
사이드바: 좌표 조정 (고정) + 필드 추가/삭제
메인: 미리보기 + 마우스 위치 x%, y% 실시간 표시

실행: streamlit run coord_tuner.py --server.port 8502
"""

import streamlit as st
import streamlit.components.v1 as components
import fitz
import json
import base64
from pathlib import Path

st.set_page_config(page_title="좌표 미세조정", layout="wide", initial_sidebar_state="expanded")

# 기본 CSS - 메인 페이지 스크롤 제거
st.markdown("""
<style>
    header[data-testid="stHeader"] { display: none; }
    .block-container { padding-top: 0.2rem; padding-bottom: 0; max-height: 100vh; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

ALL_FIELDS = {
    "의뢰인": {
        "client_name": "이름", "client_birth": "생년월일 (통째: 891102)",
        "client_birth_year": "생년 (2자리: 89)", "client_birth_year_full": "생년 (4자리: 1989)",
        "client_birth_month": "생월 (11)", "client_birth_day": "생일 (02)",
        "client_phone": "연락처", "client_address": "주소",
        "client_id_front": "주민번호 앞6자리", "client_id_back": "주민번호 뒷7자리", "client_id_full": "주민번호 전체 (XXXXXX-XXXXXXX)",
    },
    "대리인": {
        "agent_name": "이름", "agent_birth": "생년월일 (통째)",
        "agent_birth_year": "생년 (2자리)", "agent_birth_year_full": "생년 (4자리)",
        "agent_birth_month": "생월", "agent_birth_day": "생일",
        "agent_phone": "연락처", "agent_address": "주소", "agent_fax": "팩스번호",
        "agent_id_front": "주민번호 앞6자리", "agent_id_back": "주민번호 뒷7자리", "agent_id_full": "주민번호 전체 (XXXXXX-XXXXXXX)",
        "agent_sign": "싸인 (이미지)",
    },
    "날짜": {
        "date": "전체 (2026.02.24)", "date_year": "년 (2026)",
        "date_year_suffix": "년 뒤2자리 (26)", "date_month": "월 (02)", "date_day": "일 (24)",
        "date1_year": "+1일 년 (2026)", "date1_year_suffix": "+1일 년 뒤2자리 (26)",
        "date1_month": "+1일 월 (02)", "date1_day": "+1일 일 (25)",
        "date7_year": "+7일 년 (2026)", "date7_year_suffix": "+7일 년 뒤2자리 (26)",
        "date7_month": "+7일 월 (03)", "date7_day": "+7일 일 (03)",
    },
    "고정텍스트": {
        "fixed_대리인": "대리인",
        "fixed_위임인": "위임인",
        "fixed_본인": "본인",
        "fixed_신청인": "신청인",
        "fixed_대표자": "대표자",
        "fixed_관계_대리인": "대 리 인",
    },
    "채권자": {
        "creditor_name": "채권자명",
    },
    "저축은행": {
        "bank_name": "저축은행명",
        "bank_tel": "전화번호",
        "bank_fax": "팩스번호",
        "bank_branch": "지점명",
    },
    "거래기간": {
        "bank_period": "통장거래기간 (2024.01.01~2025.01.01)",
        "bank_date_from": "통장 시작일",
        "bank_date_to": "통장 종료일",
        "card_period": "카드거래기간 (2024.01.01~2025.01.01)",
        "card_date_from": "카드 시작일",
        "card_date_to": "카드 종료일",
    },
    "관공서": {
        "case_no": "사건번호",
        "court_name": "법원명",
        "plaintiff": "원고(채권자)",
        "defendant": "피고(채무자)",
        "delegation_content": "위임한 내용",
        "bike_reg_no": "이륜차등록번호",
        "car_reg_no": "자동차등록번호",
        "tax_doc_name": "지방세 체납리스트or완납증명서 (체크시 삽입)",
        "tax_year": "년도",
        "period_from": "기간 시작일",
        "period_to": "기간 종료일",
        "residence_period": "거주기간",
        "landlord_name": "집주인 이름",
        "landlord_relation": "관계",
        "landlord_phone": "집주인 전화번호",
    },
    "인감": {
        "stamp": "인감도장 (이미지)",
        "landlord_stamp": "집주인 도장 (자동생성)",
    },
}

def build_test_data(name, birth, phone, address, agent_name, agent_birth, agent_phone, agent_address, date_str):
    from datetime import datetime, timedelta
    dp = date_str.split(".")
    # +1일, +7일 날짜 계산
    try:
        base = datetime.strptime(date_str, "%Y.%m.%d")
        d1 = base + timedelta(days=1)
        d7 = base + timedelta(days=7)
        d1p = [d1.strftime("%Y"), d1.strftime("%m"), d1.strftime("%d")]
        d7p = [d7.strftime("%Y"), d7.strftime("%m"), d7.strftime("%d")]
    except:
        d1p = d7p = ["", "", ""]
    b = birth
    if len(b) == 6:
        cb_y, cb_m, cb_d = b[:2], b[2:4], b[4:6]
        cb_yf = ("19" if int(b[:2]) > 30 else "20") + b[:2]
    else:
        cb_y = cb_yf = cb_m = cb_d = ""
    ab = agent_birth
    if len(ab) == 6:
        ab_y, ab_m, ab_d = ab[:2], ab[2:4], ab[4:6]
        ab_yf = ("19" if int(ab[:2]) > 30 else "20") + ab[:2]
    else:
        ab_y = ab_yf = ab_m = ab_d = ""
    return {
        "client_name": name, "client_birth": birth,
        "client_birth_year": cb_y, "client_birth_year_full": cb_yf,
        "client_birth_month": cb_m, "client_birth_day": cb_d,
        "client_phone": phone, "client_address": address,
        "client_id_front": birth, "client_id_back": "1504316", "client_id_full": birth + "-1504316",
        "agent_name": agent_name, "agent_birth": agent_birth,
        "agent_birth_year": ab_y, "agent_birth_year_full": ab_yf,
        "agent_birth_month": ab_m, "agent_birth_day": ab_d,
        "agent_phone": agent_phone, "agent_address": agent_address,
        "agent_fax": "063-123-4567",
        "agent_id_front": agent_birth, "agent_id_back": "1481915", "agent_id_full": agent_birth + "-1481915",
        "date": date_str,
        "date_year": dp[0] if len(dp) >= 1 else "",
        "date_year_suffix": dp[0][2:] if len(dp) >= 1 and len(dp[0]) >= 4 else "",
        "date_month": dp[1] if len(dp) >= 2 else "",
        "date_day": dp[2] if len(dp) >= 3 else "",
        "date1_year": d1p[0] if len(d1p) >= 1 else "",
        "date1_year_suffix": d1p[0][2:] if len(d1p) >= 1 and len(d1p[0]) >= 4 else "",
        "date1_month": d1p[1] if len(d1p) >= 2 else "",
        "date1_day": d1p[2] if len(d1p) >= 3 else "",
        "date7_year": d7p[0] if len(d7p) >= 1 else "",
        "date7_year_suffix": d7p[0][2:] if len(d7p) >= 1 and len(d7p[0]) >= 4 else "",
        "date7_month": d7p[1] if len(d7p) >= 2 else "",
        "date7_day": d7p[2] if len(d7p) >= 3 else "",
        # 고정텍스트
        "fixed_대리인": "대리인", "fixed_위임인": "위임인", "fixed_본인": "본인",
        "fixed_신청인": "신청인", "fixed_대표자": "대표자", "fixed_관계_대리인": "대 리 인",
        # 저축은행
        "bank_name": "삼호저축은행", "bank_tel": "02-514-2345",
        "bank_fax": "063-288-2040", "bank_branch": "전주",
        # 채권자
        "creditor_name": "대한채권관리대부",
        # 거래기간
        "bank_period": "2024.01.01~2025.01.01",
        "bank_date_from": "2024.01.01", "bank_date_to": "2025.01.01",
        "card_period": "2024.01.01~2025.01.01",
        "card_date_from": "2024.01.01", "card_date_to": "2025.01.01",
        # 관공서
        "case_no": "2025회단1234", "court_name": "전주지방법원",
        "plaintiff": "주식회사 한국은행", "defendant": "홍길동",
        "delegation_content": "소송위임", "creditor_name": "주식회사 한국은행",
        "bike_reg_no": "전북 가 1234",
        "car_reg_no": "12가 3456",
        "tax_doc_name": "지방세 체납리스트or완납증명서", "tax_year": "2025",
        "period_from": "2020.01.01", "period_to": "2025.01.01",
        "residence_period": "2020.01~2025.01", "landlord_name": "김집주",
        "landlord_relation": "지인", "landlord_phone": "010-1234-5678",
    }

st.markdown("""
<style>
    .block-container { padding-top: 0.5rem; max-width: 100%; }
    header[data-testid="stHeader"] { height: 2.5rem; min-height: 2.5rem; }
    [data-testid="stSidebar"] { min-width: 430px; max-width: 520px; }
    [data-testid="stSidebar"] .stNumberInput > div { margin-bottom: -0.5rem; }
    [data-testid="stSidebar"] .stNumberInput label { font-size: 0.75rem !important; margin-bottom: 0 !important; }
    [data-testid="stSidebar"] .stExpander { margin-bottom: 0.1rem !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] details { padding: 0.2rem 0.4rem !important; }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary p { font-size: 0.82rem !important; }
</style>
""", unsafe_allow_html=True)

TEMPLATES_DIR = Path("templates/pdf")
COORDS_DIR = Path("templates/coords")
FONT_PATH = "C:/Windows/Fonts/malgun.ttf"

def find_all_forms():
    forms = {}
    for p in sorted(COORDS_DIR.rglob("*_coords.json")):
        name = p.stem.replace("_coords", "")
        rel = p.relative_to(COORDS_DIR)
        if len(rel.parts) > 1:
            display = f"[{rel.parts[0]}] {name}"
        else:
            display = name
        forms[display] = {"coords_path": p, "form_name": name,
                          "category": rel.parts[0] if len(rel.parts) > 1 else None}
    return forms

def find_template_pdf(form_name, category=None):
    if category:
        path = TEMPLATES_DIR / category / f"{form_name}.pdf"
        if path.exists(): return path
    path = TEMPLATES_DIR / f"{form_name}.pdf"
    if path.exists(): return path
    for p in TEMPLATES_DIR.rglob(f"{form_name}.pdf"):
        return p
    return None

all_forms = find_all_forms()
if not all_forms:
    st.error("templates/coords/ 폴더에 좌표 JSON 파일이 없습니다.")
    st.stop()

# ━━━━━━━━━━━━━ 사이드바 ━━━━━━━━━━━━━
with st.sidebar:
    st.markdown("### 🎯 좌표 조정")
    selected_display = st.selectbox("양식 선택", list(all_forms.keys()))
    c_save, c_ruler, c_data, c_click = st.columns([1, 1, 1, 1])
    with c_save:
        save_btn = st.button("💾 저장", type="primary", use_container_width=True)
    with c_ruler:
        show_ruler = st.toggle("눈금자", True)
    with c_data:
        show_data = st.toggle("데이터", False)
    with c_click:
        click_mode = st.toggle("📍클릭", False, key="click_mode")

    form_info = all_forms[selected_display]
    coords_path = form_info["coords_path"]
    form_name = form_info["form_name"]
    category = form_info["category"]

    state_key = f"coords_{coords_path}"
    if state_key not in st.session_state:
        with open(coords_path, "r", encoding="utf-8") as f:
            st.session_state[state_key] = json.load(f)
    coords_data = st.session_state[state_key]

    if show_data:
        with st.expander("✏️ 테스트 데이터", expanded=True):
            st.caption("의뢰인")
            tn = st.text_input("의뢰인 이름", "이시영")
            tb = st.text_input("생년월일", "891102")
            tp = st.text_input("연락처", "010-4228-6987")
            ta = st.text_input("주소", "전주시 덕진구 기린대로 418")
            st.caption("대리인")
            tag = st.text_input("대리인 이름", "이성철")
            tab = st.text_input("대리인 생년월일", "850315")
            tap = st.text_input("대리인 연락처", "010-1234-5678")
            taa = st.text_input("대리인 주소", "전주시 완산구 효자동 123")
            td = st.text_input("날짜", "2026.02.24")
    else:
        tn, tb, tp, ta = "이시영", "891102", "010-4228-6987", "전주시 덕진구 기린대로 418"
        tag, tab, tap, taa = "이성철", "850315", "010-1234-5678", "전주시 완산구 효자동 123"
        td = "2026.02.24"
    test_data = build_test_data(tn, tb, tp, ta, tag, tab, tap, taa, td)

    # ── PDF 총 페이지 수 확인 ──
    template_path = find_template_pdf(form_name, category)
    total_pdf_pages = 1
    if template_path and template_path.exists():
        _tmp = fitz.open(str(template_path))
        total_pdf_pages = _tmp.page_count
        _tmp.close()

    # ── 페이지 선택 ──
    if total_pdf_pages > 1:
        st.divider()
        current_page = st.radio("📄 페이지", list(range(1, total_pdf_pages + 1)),
                                 format_func=lambda p: f"{p}페이지", horizontal=True, key="page_select")
    else:
        current_page = 1

    # coords_data에 해당 페이지가 없으면 자동 생성
    existing_page_nums = [pg.get("page", 1) for pg in coords_data.get("pages", [])]
    if current_page not in existing_page_nums:
        coords_data.setdefault("pages", []).append({"page": current_page, "fields": []})
        st.session_state[state_key] = coords_data

    # 현재 페이지에 필드가 없으면 기본 필드 자동 추가 (1페이지만)
    DEFAULT_FIELDS = [
        {"field_id": "client_name", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "client_id_front", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "client_address", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "client_phone", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "agent_name", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "agent_id_front", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "agent_address", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
        {"field_id": "agent_phone", "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "text"},
    ]
    if current_page == 1:
        _cleared_key = f"_cleared_{coords_path}_{current_page}"
        for pg in coords_data.get("pages", []):
            if pg.get("page") == 1 and len(pg.get("fields", [])) == 0 and not st.session_state.get(_cleared_key):
                pg["fields"] = [dict(f) for f in DEFAULT_FIELDS]
                st.session_state[state_key] = coords_data
                break

    st.divider()
    with st.expander("➕ 필드 추가", expanded=False):
        # 모든 필드를 항상 추가 가능 (같은 페이지에 동일 필드 여러 번 허용)
        available = {}
        for group, fields in ALL_FIELDS.items():
            for fid, desc in fields.items():
                available[f"{group} > {desc} ({fid})"] = fid
        if available:
            selected_new = st.selectbox("추가할 필드", ["선택하세요..."] + list(available.keys()), key="add_field_select")
            if selected_new != "선택하세요..." and st.button("추가", use_container_width=True):
                # ★ 현재 위젯 값을 coords_data에 반영 후 추가
                for pg in coords_data.get("pages", []):
                    pg_prefix = form_name + f"_p{pg.get('page',1)}_"
                    for idx, f in enumerate(pg.get("fields", [])):
                        fid = f["field_id"]
                        wk = f"{pg_prefix}{idx}_{fid}"
                        if f"{wk}_x" in st.session_state:
                            f["x_pct"] = st.session_state[f"{wk}_x"]
                        if f"{wk}_y" in st.session_state:
                            f["y_pct"] = st.session_state[f"{wk}_y"]
                        if f"{wk}_fs" in st.session_state:
                            f["font_size"] = st.session_state[f"{wk}_fs"]
                        if f"{wk}_sp" in st.session_state:
                            f["spacing"] = st.session_state[f"{wk}_sp"]
                        if f"{wk}_w" in st.session_state:
                            f["width_pct"] = st.session_state[f"{wk}_w"]
                        if f"{wk}_h" in st.session_state:
                            f["height_pct"] = st.session_state[f"{wk}_h"]
                new_fid = available[selected_new]
                is_img = (new_fid in ("agent_sign", "stamp", "landlord_stamp"))
                nf = {"field_id": new_fid, "x_pct": 30.0, "y_pct": 50.0, "font_size": 11, "type": "image" if is_img else "text"}
                if is_img: nf["width_pct"] = 7; nf["height_pct"] = 2.2
                # ★ 현재 선택된 페이지에 추가
                for pg in coords_data.get("pages", []):
                    if pg.get("page") == current_page:
                        pg["fields"].append(nf)
                        break
                st.session_state[state_key] = coords_data
                st.session_state.pop(f"_cleared_{coords_path}_{current_page}", None)
                # 위젯 키 전부 초기화
                keys_to_del = [k for k in st.session_state if k.startswith(form_name + "_p")]
                for k in keys_to_del:
                    del st.session_state[k]
                st.rerun()

    if st.button("🗑️ 현재 페이지 필드 전체삭제", use_container_width=True):
        for pg in coords_data.get("pages", []):
            if pg.get("page") == current_page:
                pg["fields"] = []
        st.session_state[state_key] = coords_data
        st.session_state[f"_cleared_{coords_path}_{current_page}"] = True
        keys_to_del = [k for k in st.session_state if k.startswith(form_name + "_p")]
        for k in keys_to_del:
            del st.session_state[k]
        st.rerun()

    st.divider()
    st.caption("▲▼ 0.1 단위 | 🗑️ 필드 삭제")

    # ── 빠른 좌표 입력 모드 ──
    click_target_idx = -1
    if click_mode:
        _page_fields = []
        for pg in coords_data.get("pages", []):
            if pg.get("page") == current_page:
                for i, f in enumerate(pg.get("fields", [])):
                    fid = f["field_id"]
                    fdesc = ""
                    for g, fm in ALL_FIELDS.items():
                        if fid in fm: fdesc = fm[fid]; break
                    _page_fields.append((i, fid, fdesc))
        if _page_fields:
            _opts = [f"#{i+1} {fid} ({fdesc})" for i, fid, fdesc in _page_fields]
            _sel = st.selectbox("📍 배치할 필드", _opts, key="click_target_sel")
            click_target_idx = int(_sel.split("#")[1].split(" ")[0]) - 1
            st.caption("미리보기에서 좌표 확인 → 아래에 `x,y` 입력 후 Enter")
            _coord_input = st.text_input("좌표 (x,y)", placeholder="예: 35.2,48.1", key="coord_quick_input")
            if _coord_input and "," in _coord_input:
                try:
                    _parts = _coord_input.split(",")
                    _cx, _cy = float(_parts[0].strip()), float(_parts[1].strip())
                    if 0 <= _cx <= 100 and 0 <= _cy <= 100:
                        for pg in coords_data.get("pages", []):
                            if pg.get("page") == current_page:
                                _fields = pg.get("fields", [])
                                if click_target_idx < len(_fields):
                                    _fields[click_target_idx]["x_pct"] = round(_cx, 1)
                                    _fields[click_target_idx]["y_pct"] = round(_cy, 1)
                        st.session_state[state_key] = coords_data
                        _keys_del = [k for k in st.session_state if k.startswith(form_name + "_p")]
                        for k in _keys_del:
                            del st.session_state[k]
                        st.session_state["coord_quick_input"] = ""
                        st.rerun()
                except:
                    st.error("형식: 35.2,48.1")
        else:
            st.caption("⚠️ 먼저 필드를 추가하세요")

    # 위젯 key에 양식명 포함 (양식 간 값 충돌 방지)
    wk_prefix = form_name + f"_p{current_page}_"

    pages = coords_data.get("pages", [])
    adjusted_pages = []
    for page_info in pages:
        pnum = page_info.get("page", 1)
        fields = page_info.get("fields", [])
        adj_fields = []
        del_idx = None

        # ★ 현재 선택된 페이지만 편집 UI 표시
        if pnum == current_page:
            for i, field in enumerate(fields):
                fid, ftype = field["field_id"], field.get("type", "text")
                dval = test_data.get(fid, "")
                fdesc = ""
                for g, fm in ALL_FIELDS.items():
                    if fid in fm: fdesc = fm[fid]; break
                # 위젯 key: 양식명_페이지_인덱스_필드id (중복 허용)
                wk = f"{wk_prefix}{i}_{fid}"
                if ftype == "image":
                    with st.expander(f"🖼 {fid} ({fdesc})", expanded=False):
                        c1, c2 = st.columns(2)
                        nx = c1.number_input("x", 0.0, 100.0, float(field["x_pct"]), step=0.1, format="%.1f", key=f"{wk}_x")
                        ny = c2.number_input("y", 0.0, 100.0, float(field["y_pct"]), step=0.1, format="%.1f", key=f"{wk}_y")
                        c3, c4 = st.columns(2)
                        nw = c3.number_input("w", 0.0, 20.0, float(field.get("width_pct", 8)), step=0.1, format="%.1f", key=f"{wk}_w")
                        nh = c4.number_input("h", 0.0, 10.0, float(field.get("height_pct", 4)), step=0.1, format="%.1f", key=f"{wk}_h")
                        if st.button("🗑️ 삭제", key=f"{wk}_del", use_container_width=True): del_idx = i
                        nf = {**field, "x_pct": nx, "y_pct": ny, "width_pct": nw, "height_pct": nh}
                else:
                    label = f"#{i+1} {fid} ({fdesc})" if fdesc else f"#{i+1} {fid}"
                    with st.expander(f"📝 {label}", expanded=True):
                        c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
                        nx = c1.number_input("x", 0.0, 100.0, float(field["x_pct"]), step=0.1, format="%.1f", key=f"{wk}_x")
                        ny = c2.number_input("y", 0.0, 100.0, float(field["y_pct"]), step=0.1, format="%.1f", key=f"{wk}_y")
                        nfs = c3.number_input("크기", 6, 20, int(field.get("font_size", 11)), step=1, key=f"{wk}_fs")
                        nsp = c4.number_input("자간", 0.0, 30.0, float(field.get("spacing", 0)), step=0.5, format="%.1f", key=f"{wk}_sp")
                        if st.button("🗑️ 삭제", key=f"{wk}_del", use_container_width=True): del_idx = i
                        nf = {**field, "x_pct": nx, "y_pct": ny, "font_size": nfs, "spacing": nsp}
                adj_fields.append(nf)
        else:
            # 다른 페이지는 원본 그대로 유지
            adj_fields = list(fields)

        if del_idx is not None:
            for idx, f in enumerate(adj_fields):
                fid = f["field_id"]
                wk = f"{wk_prefix}{idx}_{fid}"
                if f"{wk}_x" in st.session_state:
                    f["x_pct"] = st.session_state[f"{wk}_x"]
                if f"{wk}_y" in st.session_state:
                    f["y_pct"] = st.session_state[f"{wk}_y"]
                if f"{wk}_fs" in st.session_state:
                    f["font_size"] = st.session_state[f"{wk}_fs"]
                if f"{wk}_sp" in st.session_state:
                    f["spacing"] = st.session_state[f"{wk}_sp"]
                if f"{wk}_w" in st.session_state:
                    f["width_pct"] = st.session_state[f"{wk}_w"]
                if f"{wk}_h" in st.session_state:
                    f["height_pct"] = st.session_state[f"{wk}_h"]
            adj_fields.pop(del_idx)
            # 해당 페이지 찾아서 업데이트
            for pg in coords_data["pages"]:
                if pg.get("page") == pnum:
                    pg["fields"] = adj_fields
                    break
            st.session_state[state_key] = coords_data
            # 위젯 키 전부 초기화 (인덱스가 바뀌므로)
            keys_to_del = [k for k in st.session_state if k.startswith(wk_prefix)]
            for k in keys_to_del:
                del st.session_state[k]
            st.rerun()
        adjusted_pages.append({"page": pnum, "fields": adj_fields})

# ━━━━━━━━━━━━━ 메인: 미리보기 ━━━━━━━━━━━━━
if not template_path or not template_path.exists():
    st.error(f"템플릿 PDF를 찾을 수 없습니다: {form_name}.pdf")
    st.stop()

def draw_rulers(page):
    """세로+가로 눈금자를 PDF에 직접 그리기"""
    rect = page.rect
    w, h = rect.width, rect.height
    rc, lc, mc = (0.6,0.6,0.6), (0.4,0.4,0.4), (0.75,0.75,0.75)
    # 가로 (상단)
    rh = 14
    page.draw_rect(fitz.Rect(0, 0, w, rh), color=None, fill=(0.95, 0.95, 0.95))
    for p in range(0, 101):
        x = w * (p / 100)
        if p % 10 == 0:
            page.draw_line(fitz.Point(x, 0), fitz.Point(x, rh), color=rc, width=0.8)
            page.draw_line(fitz.Point(x, rh), fitz.Point(x, h), color=mc, width=0.3, dashes="[2] 0")
            if 0 < p < 100:
                page.insert_text(fitz.Point(x+1, rh-3), f"{p}", fontsize=5.5, color=lc)
        elif p % 5 == 0:
            page.draw_line(fitz.Point(x, rh-7), fitz.Point(x, rh), color=rc, width=0.5)
    # 세로 (왼쪽)
    rw = 16
    page.draw_rect(fitz.Rect(0, rh, rw, h), color=None, fill=(0.95, 0.95, 0.95))
    for p in range(0, 101):
        y = h * (p / 100)
        if p % 10 == 0:
            page.draw_line(fitz.Point(0, y), fitz.Point(rw, y), color=rc, width=0.8)
            page.draw_line(fitz.Point(rw, y), fitz.Point(w, y), color=mc, width=0.3, dashes="[2] 0")
            if 0 < p < 100:
                page.insert_text(fitz.Point(1, y+5), f"{p}", fontsize=5.5, color=lc)
        elif p % 5 == 0:
            page.draw_line(fitz.Point(rw-7, y), fitz.Point(rw, y), color=rc, width=0.5)
    page.draw_rect(fitz.Rect(0, 0, rw, rh), color=None, fill=(0.9, 0.9, 0.9))
    page.insert_text(fitz.Point(2, rh-3), "%", fontsize=5.5, color=lc)

def render_preview(tpath, pdata, tdata, ruler=True, target_page=1):
    doc = fitz.open(str(tpath))
    page_idx = target_page - 1
    if page_idx >= doc.page_count:
        page_idx = 0
    for pi in pdata:
        pn = pi["page"] - 1
        if pn >= doc.page_count: continue
        page = doc[pn]
        rect = page.rect
        for f in pi["fields"]:
            fid, ft = f["field_id"], f.get("type", "text")
            x = rect.width * (f["x_pct"] / 100)
            y = rect.height * (f["y_pct"] / 100)
            if ft == "image":
                fw = rect.width * (f.get("width_pct", 8) / 100)
                fh = rect.height * (f.get("height_pct", 4) / 100)
                page.draw_rect(fitz.Rect(x, y-fh, x+fw, y), color=(1,0,0), width=0.5)
                page.insert_text(fitz.Point(x+2, y-2), f"[{fid}]", fontsize=7, color=(1,0,0))
            else:
                val = tdata.get(fid, "")
                if not val: continue
                fs = f.get("font_size", 11)
                sp = f.get("spacing", 0)
                if sp > 0:
                    # 자간 적용: 글자 하나씩 배치
                    cx = x
                    font = fitz.Font(fontfile=FONT_PATH)
                    for char in str(val):
                        try:
                            page.insert_text(fitz.Point(cx, y), char, fontname="malgun", fontfile=FONT_PATH, fontsize=fs, color=(0,0,0.8))
                        except:
                            page.insert_text(fitz.Point(cx, y), char, fontsize=fs, color=(0,0,0.8))
                        cx += font.text_length(char, fontsize=fs) + sp
                else:
                    try:
                        page.insert_text(fitz.Point(x, y), str(val), fontname="malgun", fontfile=FONT_PATH, fontsize=fs, color=(0,0,0.8))
                    except:
                        page.insert_text(fitz.Point(x, y), str(val), fontsize=fs, color=(0,0,0.8))
    if ruler and page_idx < doc.page_count:
        draw_rulers(doc[page_idx])
    pix = doc[page_idx].get_pixmap(dpi=150)
    img = pix.tobytes("png")
    w, h = pix.width, pix.height
    doc.close()
    return img, w, h

st.caption(f"👁️ {form_name} ({current_page}p) — 마우스 위치 = 좌표 %")

try:
    preview_bytes, img_w, img_h = render_preview(template_path, adjusted_pages, test_data, ruler=show_ruler, target_page=current_page)
    img_b64 = base64.b64encode(preview_bytes).decode()

    html = f"""
    <div id="preview-container" style="position:relative;width:100%;cursor:crosshair;">
        <img id="preview-img" src="data:image/png;base64,{img_b64}" style="width:100%;display:block;" />
        <div id="coord-label" style="
            display:none;
            position:fixed;
            background:rgba(0,0,0,0.8);
            color:#fff;
            padding:4px 10px;
            border-radius:4px;
            font-size:14px;
            font-weight:bold;
            font-family:monospace;
            pointer-events:none;
            z-index:9999;
            white-space:nowrap;
        "></div>
        <div id="crosshair-x" style="display:none;position:absolute;top:0;width:1px;height:100%;background:rgba(255,0,0,0.4);pointer-events:none;z-index:10;"></div>
        <div id="crosshair-y" style="display:none;position:absolute;left:0;height:1px;width:100%;background:rgba(255,0,0,0.4);pointer-events:none;z-index:10;"></div>
        <div id="click-dot" style="display:none;position:absolute;width:12px;height:12px;background:rgba(255,30,30,0.9);border:2px solid #fff;border-radius:50%;pointer-events:none;z-index:20;transform:translate(-50%,-50%);box-shadow:0 0 6px rgba(0,0,0,0.5);"></div>
        <div id="pinned-label" style="display:none;position:fixed;top:8px;right:8px;background:rgba(220,40,40,0.95);color:#fff;padding:8px 16px;border-radius:6px;font:bold 16px monospace;z-index:9999;box-shadow:0 2px 8px rgba(0,0,0,0.3);"></div>
    </div>
    <script>
        // 스크롤 위치 복원
        const savedScroll = window.parent.sessionStorage.getItem('preview_scroll');
        if (savedScroll) {{
            document.documentElement.scrollTop = parseInt(savedScroll);
            document.body.scrollTop = parseInt(savedScroll);
        }}
        window.addEventListener('scroll', function() {{
            window.parent.sessionStorage.setItem('preview_scroll', window.scrollY || document.documentElement.scrollTop);
        }});

        const container = document.getElementById('preview-container');
        const img = document.getElementById('preview-img');
        const label = document.getElementById('coord-label');
        const crossX = document.getElementById('crosshair-x');
        const crossY = document.getElementById('crosshair-y');
        const clickDot = document.getElementById('click-dot');
        const pinned = document.getElementById('pinned-label');
        
        container.addEventListener('mousemove', function(e) {{
            const rect = img.getBoundingClientRect();
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;
            
            if (x >= 0 && x <= rect.width && y >= 0 && y <= rect.height) {{
                const xPct = (x / rect.width * 100).toFixed(1);
                const yPct = (y / rect.height * 100).toFixed(1);
                
                label.style.display = 'block';
                label.textContent = 'x: ' + xPct + '%  y: ' + yPct + '%';
                
                label.style.left = (e.clientX + 15) + 'px';
                label.style.top = (e.clientY - 35) + 'px';
                
                crossX.style.display = 'block';
                crossX.style.left = x + 'px';
                crossY.style.display = 'block';
                crossY.style.top = (y) + 'px';
            }}
        }});
        
        container.addEventListener('mouseleave', function() {{
            label.style.display = 'none';
            crossX.style.display = 'none';
            crossY.style.display = 'none';
        }});

        // 클릭하면 빨간 점 + 좌표 고정 표시
        container.addEventListener('click', function(e) {{
            const rect = img.getBoundingClientRect();
            const px = e.clientX - rect.left;
            const py = e.clientY - rect.top;
            if (px < 0 || px > rect.width || py < 0 || py > rect.height) return;
            const xPct = (px / rect.width * 100).toFixed(1);
            const yPct = (py / rect.height * 100).toFixed(1);
            // 빨간 점 표시
            clickDot.style.display = 'block';
            clickDot.style.left = px + 'px';
            clickDot.style.top = py + 'px';
            // 우상단에 좌표 고정 표시
            pinned.style.display = 'block';
            pinned.textContent = xPct + ',' + yPct;
        }});
    </script>
    """

    components.html(html, height=min(img_h + 50, 850), scrolling=True)

except Exception as e:
    st.error(f"미리보기 오류: {e}")

# ── 저장 ──
if save_btn:
    output_data = {**coords_data, "pages": adjusted_pages}
    json_str = json.dumps(output_data, indent=2, ensure_ascii=False)
    with open(coords_path, "w", encoding="utf-8") as f:
        f.write(json_str)
    st.session_state[state_key] = output_data
    st.toast(f"✅ 저장 완료: {coords_path}", icon="💾")

# ── 사이드바 스크롤 위치 복원 (맨 마지막에 실행) ──
components.html("""
<script>
(function() {
    const sbKey = 'sidebar_scroll_pos';
    const mainKey = 'main_scroll_pos';
    function getSidebar() {
        return window.parent.document.querySelector('[data-testid="stSidebarContent"]');
    }
    function getMain() {
        return window.parent.document.querySelector('[data-testid="stAppViewContainer"]');
    }
    function restore() {
        const sb = getSidebar();
        if (sb) {
            const saved = window.parent.sessionStorage.getItem(sbKey);
            if (saved) sb.scrollTop = parseInt(saved);
            sb.onscroll = function() {
                window.parent.sessionStorage.setItem(sbKey, sb.scrollTop);
            };
        }
        const main = getMain();
        if (main) {
            const saved = window.parent.sessionStorage.getItem(mainKey);
            if (saved) main.scrollTop = parseInt(saved);
            main.onscroll = function() {
                window.parent.sessionStorage.setItem(mainKey, main.scrollTop);
            };
        }
    }
    setTimeout(restore, 100);
    setTimeout(restore, 300);
    setTimeout(restore, 600);
})();
</script>
""", height=0)
