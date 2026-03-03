"""
ui_main.py - 메인 영역 UI 컴포넌트
멀티셀렉트 UI + expander 일괄설정 + 줄무늬 행
"""

import streamlit as st
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
from modules.config_loader import load_settings, load_creditors, get_doc_options, get_needs_date_labels, get_all_creditor_names


# ═══════════════════════════════════════
# 유틸
# ═══════════════════════════════════════

def _delete_creditor(del_idx):
    """채권사 삭제: del_idx 행의 데이터를 비우고 위로 당김"""
    count = st.session_state.get("creditor_count", 10)
    prefixes = ["cred_", "docs_", "acct_", "cust_",
                "df_bank_", "dt_bank_", "df_card_", "dt_card_"]

    # 삭제 대상부터 마지막까지: 다음 행 값으로 덮어쓰기
    for i in range(del_idx, count - 1):
        for prefix in prefixes:
            next_key = f"{prefix}{i+1}"
            cur_key = f"{prefix}{i}"
            if next_key in st.session_state:
                try:
                    st.session_state[cur_key] = st.session_state[next_key]
                except Exception:
                    pass  # 위젯 키 충돌 무시
            else:
                try:
                    del st.session_state[cur_key]
                except (KeyError, Exception):
                    pass

    # 마지막 슬롯 삭제
    last = count - 1
    for prefix in prefixes:
        try:
            del st.session_state[f"{prefix}{last}"]
        except (KeyError, Exception):
            pass

    st.session_state["creditor_count"] = max(count - 1, 1)


def _calc_start(end_date, years):
    try:
        return end_date - relativedelta(years=years) + timedelta(days=1)
    except:
        return date(end_date.year - years, end_date.month, end_date.day) + timedelta(days=1)

def _default_dates():
    end = st.session_state.get("_warrant_date", date.today())
    return _calc_start(end, 1), end


# ═══════════════════════════════════════
# on_click 콜백
# ═══════════════════════════════════════

def _cb_bank_year(years):
    end = st.session_state.get("_warrant_date", date.today())
    st.session_state["g_bank_from"] = _calc_start(end, years)
    st.session_state["g_bank_to"] = end

def _cb_card_year(years):
    end = st.session_state.get("_warrant_date", date.today())
    st.session_state["g_card_from"] = _calc_start(end, years)
    st.session_state["g_card_to"] = end

def _cb_same_as_bank():
    fr = st.session_state.get("g_bank_from")
    to = st.session_state.get("g_bank_to")
    if fr and to:
        st.session_state["g_card_from"] = fr
        st.session_state["g_card_to"] = to

def _cb_apply_bank():
    fr = st.session_state.get("g_bank_from")
    to = st.session_state.get("g_bank_to")
    if fr and to:
        for i in range(st.session_state.get("creditor_count", 5)):
            for doc in st.session_state.get(f"docs_{i}", []):
                if doc == "통장거래내역":
                    st.session_state[f"df_bank_{i}"] = fr
                    st.session_state[f"dt_bank_{i}"] = to

def _cb_apply_card():
    fr = st.session_state.get("g_card_from")
    to = st.session_state.get("g_card_to")
    if fr and to:
        for i in range(st.session_state.get("creditor_count", 5)):
            for doc in st.session_state.get(f"docs_{i}", []):
                if doc == "카드거래내역":
                    st.session_state[f"df_card_{i}"] = fr
                    st.session_state[f"dt_card_{i}"] = to


# ═══════════════════════════════════════
# CSS
# ═══════════════════════════════════════

ROW_STYLE = """
<style>
div[data-testid="stVerticalBlock"] > div:has(> div.row-even) {
    background: #f8f9fc;
    border-radius: 4px;
}
</style>
"""


# ═══════════════════════════════════════
# UI
# ═══════════════════════════════════════

def render_main():
    settings = load_settings()
    app_cfg = settings.get("app", {})
    max_cred = app_cfg.get("max_creditors", 20)
    default_cred = 10

    # 멀티셀렉트 옵션 (— 선택 — 제외)
    all_labels = get_doc_options(settings)
    doc_options = [d for d in all_labels if d != "— 선택 —"]
    needs_date = set(get_needs_date_labels(settings))

    if "creditor_count" not in st.session_state:
        st.session_state.creditor_count = default_cred

    # ── 헤더 ──
    col_title, col_btn = st.columns([2, 3])
    with col_title:
        count = st.session_state.creditor_count
        st.markdown(
            f"### 채권사 및 발급서류 "
            f"<small style='color:#6b7280;font-weight:400;font-size:14px;'>{count}개</small>",
            unsafe_allow_html=True,
        )
    with col_btn:
        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button("＋ 채권사 추가", key="add_cred", use_container_width=True):
                if st.session_state.creditor_count < max_cred:
                    st.session_state.creditor_count += 1
                    st.rerun()
        with b2:
            generate_clicked = st.button("⚡ PDF 생성", type="primary", key="gen_pdf", use_container_width=True)
        with b3:
            if "pdf_bytes" in st.session_state and st.session_state.pdf_bytes:
                st.download_button(
                    label="📥 인쇄용", data=st.session_state.pdf_bytes,
                    file_name=st.session_state.get("pdf_filename", "인쇄용.pdf"),
                    mime="application/pdf", use_container_width=True,
                )
            else:
                st.button("📥 인쇄용", disabled=True, key="dl_disabled", use_container_width=True)
        # 팩스용 ZIP 다운로드
        if "fax_zip_bytes" in st.session_state and st.session_state.fax_zip_bytes:
            st.download_button(
                label="📠 팩스용 ZIP 다운", data=st.session_state.fax_zip_bytes,
                file_name=st.session_state.get("fax_zip_filename", "팩스용.zip"),
                mime="application/zip", use_container_width=True,
            )

    # ── 기간 일괄설정 (접기) ──
    progress_placeholder = st.empty()  # PDF 생성 진행바 위치 (상단)
    with st.expander("📅 거래내역 기간 일괄설정", expanded=False):
        _render_global_date()

    st.divider()

    # ── 테이블 헤더 ──
    hc = st.columns([0.3, 1.2, 2.5])
    hc[0].markdown("<small style='color:#6b7280;font-weight:600;'>#</small>", unsafe_allow_html=True)
    hc[1].markdown("<small style='color:#6b7280;font-weight:600;'>채권사명</small>", unsafe_allow_html=True)
    hc[2].markdown("<small style='color:#6b7280;font-weight:600;'>서류</small>", unsafe_allow_html=True)

    # ── 채권사 행 ──
    creditors = []
    for i in range(st.session_state.creditor_count):
        cred = _render_creditor_row(i, doc_options, needs_date)
        if cred["name"]:
            creditors.append(cred)

    st.divider()
    total_docs = sum(len(c["docs"]) for c in creditors)

    # ── 채권사명 자동완성 (datalist 주입) ──
    import json
    import streamlit.components.v1 as components
    # issue_manual.json에서 전체 기관명 로드
    from modules.config_loader import load_issue_manual
    manual = load_issue_manual()
    all_cred = sorted(manual.keys())
    # creditors.json 기존 목록도 병합
    cred_json = load_creditors()
    for key, val in cred_json.items():
        if isinstance(val, list):
            all_cred.extend(val)
    # 저축은행 DB 병합
    from modules.config_loader import load_savings_banks
    savings = load_savings_banks()
    all_cred.extend(savings.keys())
    # 대부업체 병합
    from modules.config_loader import load_loan_companies
    all_cred.extend(load_loan_companies())
    # 기타 (통신사, 공공기관 등) 병합
    from modules.config_loader import load_misc_companies
    all_cred.extend(load_misc_companies())
    # 중복 제거 + 정렬
    all_cred = sorted(set(all_cred))
    opts_js = json.dumps(all_cred, ensure_ascii=False)
    components.html(f"""
    <script>
    (function() {{
        const opts = {opts_js};
        const p = window.parent.document;
        let dl = p.getElementById('cred-datalist');
        if (!dl) {{
            dl = p.createElement('datalist');
            dl.id = 'cred-datalist';
            opts.forEach(o => {{
                const op = p.createElement('option');
                op.value = o;
                dl.appendChild(op);
            }});
            p.body.appendChild(dl);
        }}
        p.querySelectorAll('input[placeholder="채권사명"]').forEach(inp => {{
            inp.setAttribute('list', 'cred-datalist');
            inp.addEventListener('keydown', function(e) {{
                if (e.key === 'Enter') {{
                    inp.removeAttribute('list');
                    inp.blur();
                }}
            }});
            inp.addEventListener('focus', function() {{
                if (!inp.getAttribute('list')) {{
                    inp.setAttribute('list', 'cred-datalist');
                }}
            }});
        }});
    }})();
    </script>
    """, height=0)
    col_info, col_add = st.columns([2, 1])
    with col_info:
        st.markdown(f"선택된 채권사: **{len(creditors)}**개 | 총 서류: **{total_docs}**건")
    with col_add:
        if st.button("＋ 채권사 추가", key="add_cred_bottom", use_container_width=True):
            if st.session_state.creditor_count < max_cred:
                st.session_state.creditor_count += 1
                st.rerun()

    return {"creditors": creditors, "generate_clicked": generate_clicked, "progress_placeholder": progress_placeholder}


def _render_global_date():
    ds, de = _default_dates()

    # ── 통장 ──
    st.caption("통장거래내역")
    b1, b2, b3 = st.columns([2, 2, 1])
    with b1:
        st.date_input("시작", value=ds, key="g_bank_from", label_visibility="collapsed")
    with b2:
        st.date_input("종료", value=de, key="g_bank_to", label_visibility="collapsed")
    with b3:
        st.button("통장 적용", key="apply_bank", on_click=_cb_apply_bank, use_container_width=True)

    by = st.columns([1, 1, 1, 1, 1, 5], gap="small")
    for i, col in enumerate(by[:5], 1):
        with col:
            st.button(f"{i}년", key=f"bank_{i}y", on_click=_cb_bank_year, args=(i,), use_container_width=True)

    # ── 카드 ──
    st.caption("카드거래내역")
    c1, c2, c3, c4 = st.columns([2, 2, 0.7, 0.7])
    with c1:
        st.date_input("시작", value=ds, key="g_card_from", label_visibility="collapsed")
    with c2:
        st.date_input("종료", value=de, key="g_card_to", label_visibility="collapsed")
    with c3:
        st.button("카드 적용", key="apply_card", on_click=_cb_apply_card, use_container_width=True)
    with c4:
        st.button("⬆ 동일", key="same_as_bank", on_click=_cb_same_as_bank, use_container_width=True)

    cy = st.columns([1, 1, 1, 1, 1, 5], gap="small")
    for i, col in enumerate(cy[:5], 1):
        with col:
            st.button(f"{i}년", key=f"card_{i}y", on_click=_cb_card_year, args=(i,), use_container_width=True)


def _render_creditor_row(idx, doc_options, needs_date):
    ds, de = _default_dates()

    # 줄무늬 배경
    if idx % 2 == 1:
        st.markdown(
            f'<div class="row-even" style="background:#f8f9fc;margin:-8px -16px;padding:8px 16px;border-radius:4px;"></div>',
            unsafe_allow_html=True,
        )

    cols = st.columns([0.3, 1.2, 2.5, 0.3])

    with cols[0]:
        st.markdown(
            f"<div style='padding-top:8px;color:#6b7280;font-weight:600;'>{idx+1}</div>",
            unsafe_allow_html=True,
        )

    with cols[1]:
        cred_name = st.text_input(
            "채권사", key=f"cred_{idx}", label_visibility="collapsed",
            placeholder="채권사명"
        )

    with cols[2]:
        selected_docs = st.multiselect(
            "서류", doc_options, key=f"docs_{idx}", label_visibility="collapsed",
            placeholder="서류 선택..."
        )

    with cols[3]:
        st.button("🗑️", key=f"del_{idx}", help="채권사 삭제",
                  on_click=_delete_creditor, args=(idx,))

    # 거래내역 날짜 상세 입력 (컬럼 밖에서 전체 너비 사용)
    has_bank = "통장거래내역" in selected_docs
    has_card = "카드거래내역" in selected_docs
    has_custom = "기타" in selected_docs

    if has_bank or has_card:
        if has_bank:
            bc = st.columns([0.5, 1, 1, 1])
            with bc[0]:
                st.markdown("<div style='padding-top:8px;font-size:12px;color:#6b7280;'>통장</div>", unsafe_allow_html=True)
            with bc[1]:
                _bank_from_kwargs = {"key": f"df_bank_{idx}", "label_visibility": "collapsed"}
                if f"df_bank_{idx}" not in st.session_state:
                    _bank_from_kwargs["value"] = ds
                st.date_input("시작", **_bank_from_kwargs)
            with bc[2]:
                _bank_to_kwargs = {"key": f"dt_bank_{idx}", "label_visibility": "collapsed"}
                if f"dt_bank_{idx}" not in st.session_state:
                    _bank_to_kwargs["value"] = de
                st.date_input("종료", **_bank_to_kwargs)
            with bc[3]:
                st.text_input(
                    "계좌", placeholder="계좌번호",
                    key=f"acct_{idx}", label_visibility="collapsed"
                )

        if has_card:
            cc = st.columns([0.5, 1, 1, 1])
            with cc[0]:
                st.markdown("<div style='padding-top:8px;font-size:12px;color:#6b7280;'>카드</div>", unsafe_allow_html=True)
            with cc[1]:
                _card_from_kwargs = {"key": f"df_card_{idx}", "label_visibility": "collapsed"}
                if f"df_card_{idx}" not in st.session_state:
                    _card_from_kwargs["value"] = ds
                st.date_input("시작", **_card_from_kwargs)
            with cc[2]:
                _card_to_kwargs = {"key": f"dt_card_{idx}", "label_visibility": "collapsed"}
                if f"dt_card_{idx}" not in st.session_state:
                    _card_to_kwargs["value"] = de
                st.date_input("종료", **_card_to_kwargs)

    if has_custom:
        custom_text = st.text_input(
            "기타 서류명", placeholder="직접 입력",
            key=f"cust_{idx}", label_visibility="collapsed"
        )

    # docs 리스트 조립
    docs = []
    for doc_type in selected_docs:
        if doc_type == "통장거래내역" and has_bank:
            entry = {
                "type": doc_type,
                "date_from": str(st.session_state.get(f"df_bank_{idx}", ds)),
                "date_to": str(st.session_state.get(f"dt_bank_{idx}", de)),
                "account": st.session_state.get(f"acct_{idx}", ""),
            }
            docs.append(entry)
        elif doc_type == "카드거래내역" and has_card:
            entry = {
                "type": doc_type,
                "date_from": str(st.session_state.get(f"df_card_{idx}", ds)),
                "date_to": str(st.session_state.get(f"dt_card_{idx}", de)),
            }
            docs.append(entry)
        elif doc_type == "기타":
            docs.append({"type": doc_type, "custom": st.session_state.get(f"cust_{idx}", "")})
        else:
            docs.append({"type": doc_type})

    return {"name": cred_name, "docs": docs}
