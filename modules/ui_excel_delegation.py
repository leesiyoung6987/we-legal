"""
ui_excel_delegation.py - 📥 엑셀 → 자동입력 탭 UI
파싱 → 사무소선택/기간설정 → 기관별 병합 미리보기 → 채권사 서류 탭 전달
+ 발급매뉴얼 PDF + 고객요청 문자(기간 포함)
"""
import streamlit as st
import io
import json
from datetime import date
from dateutil.relativedelta import relativedelta
from modules.excel_parser import parse_excel, ParsedExcel
from modules.creditor_matcher import match_creditor, MatchResult
from modules.config_loader import get_issue_info, load_law_firms, load_staff


# ═══════════════════════════════════════
# 콜백
# ═══════════════════════════════════════

def _on_transfer_click():
    """전달 버튼 콜백"""
    parsed = st.session_state.get("parsed_excel")
    merged = st.session_state.get("merged_creditors")
    if not parsed or not merged:
        return

    p = parsed.person
    if p.name:
        st.session_state["client_name"] = p.name
    if p.ssn_front:
        st.session_state["client_id_front"] = p.ssn_front
    if p.ssn_back:
        st.session_state["client_id_back"] = p.ssn_back
    if p.address:
        st.session_state["client_address"] = p.address
    if p.phone_clean:
        st.session_state["client_phone"] = p.phone_clean

    transfer_items = _get_transfer_items(merged)
    count = len(transfer_items)
    st.session_state["creditor_count"] = max(count, 5)

    # 기간 정보 가져오기
    bank_from = st.session_state.get("_period_bank_from")
    bank_to = st.session_state.get("_period_bank_to")
    card_from = st.session_state.get("_period_card_from")
    card_to = st.session_state.get("_period_card_to")

    for i, item in enumerate(transfer_items):
        st.session_state[f"cred_{i}"] = item["name"]
        st.session_state[f"docs_{i}"] = item["docs"]
        if item.get("accounts"):
            st.session_state[f"acct_{i}"] = item["accounts"]
        else:
            st.session_state.pop(f"acct_{i}", None)
        # 기간 전달
        if "통장거래내역" in item["docs"] and bank_from and bank_to:
            st.session_state[f"df_bank_{i}"] = bank_from
            st.session_state[f"dt_bank_{i}"] = bank_to
        if "카드거래내역" in item["docs"] and card_from and card_to:
            st.session_state[f"df_card_{i}"] = card_from
            st.session_state[f"dt_card_{i}"] = card_to

    for i in range(count, 20):
        st.session_state[f"cred_{i}"] = ""
        st.session_state[f"docs_{i}"] = []
        st.session_state.pop(f"acct_{i}", None)

    st.session_state["_excel_transferred"] = True
    st.session_state["_excel_transfer_count"] = count


# ═══════════════════════════════════════
# 기간 계산
# ═══════════════════════════════════════

def _calc_period(years, base_date=None):
    """최근 n년 기간 계산: n년 전 다음날 ~ 기준일
    예: 기준일 2026.03.02, 1년 → 2025.03.03 ~ 2026.03.02
    """
    if not years or not isinstance(years, int):
        return None, None
    if base_date is None:
        base_date = date.today()
    end = base_date
    start = base_date - relativedelta(years=years) + relativedelta(days=1)
    return start, end


# ═══════════════════════════════════════
# 메인 렌더
# ═══════════════════════════════════════

def render_excel_delegation_tab():
    """📥 엑셀 → 자동입력 탭"""

    if st.session_state.pop("_excel_transferred", False):
        cnt = st.session_state.pop("_excel_transfer_count", 0)
        st.success(f"✅ {cnt}개 채권사 + 위임인 정보 + 기간 전달 완료! '📋 채권사 서류' 탭으로 이동하세요.")

    st.subheader("📥 엑셀 → 자동입력")
    st.caption("자료제출목록 엑셀을 파싱하여 채권사 서류 탭에 자동 입력합니다.")

    uploaded = st.file_uploader(
        "자료제출목록 엑셀 업로드",
        type=["xlsx", "xls"],
        key="excel_upload_phase2",
    )

    if not uploaded:
        st.info("💡 엑셀 파일을 업로드하세요.")
        return

    # ── 파싱 ──
    if "parsed_excel" not in st.session_state or st.session_state.get("_excel_name") != uploaded.name:
        with st.spinner("엑셀 파싱 중..."):
            parsed = parse_excel(io.BytesIO(uploaded.read()))
            st.session_state.parsed_excel = parsed
            st.session_state._excel_name = uploaded.name
            st.session_state.merged_creditors = _merge_by_institution(parsed)

    parsed: ParsedExcel = st.session_state.parsed_excel
    merged = st.session_state.merged_creditors

    for err in parsed.errors:
        st.warning(err)

    # ── 위임인 정보 ──
    with st.expander("👤 위임인 정보", expanded=True):
        p = parsed.person
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("이름", value=p.name, key="ex_p_name", disabled=True)
            st.text_input("주민번호", value=p.ssn, key="ex_p_ssn", disabled=True)
        with c2:
            st.text_input("전화번호", value=p.phone, key="ex_p_phone", disabled=True)
            st.text_input("주소", value=p.address, key="ex_p_addr", disabled=True)

    st.divider()

    # ── 법률사무소 선택 + 기간 설정 ──
    _render_period_settings()

    st.divider()

    if not merged:
        st.warning("전달할 채권사가 없습니다.")
        return

    # ── 분류 ──
    transfer_items = _get_transfer_items(merged)
    customer_items = _get_customer_request_items(merged)
    skip_items = _get_skip_items(merged)

    # ── 전달 대상 미리보기 ──
    st.markdown(f"### 📋 전달 미리보기 ({len(transfer_items)}건)")
    st.caption("고객요청·홈페이지 건 제외. 현장/방문발급 → 등기 순서.")

    hcols = st.columns([0.3, 1.3, 1.8, 1.0, 0.6])
    for col, txt in zip(hcols, ["#", "채권사명", "서류", "계좌번호", "발급방법"]):
        col.markdown(f"<small style='color:#6b7280;font-weight:600;'>{txt}</small>", unsafe_allow_html=True)

    for i, item in enumerate(transfer_items):
        _render_merged_row(i, item)

    # ── 고객요청 ──
    if customer_items:
        st.divider()
        st.markdown(f"### 📱 고객요청 ({len(customer_items)}건)")
        for i, item in enumerate(customer_items):
            _render_merged_row(i, item, badge_override="고객요청")

    # ── 기타 ──
    if skip_items:
        st.divider()
        st.markdown(f"### 🌐 기타 ({len(skip_items)}건)")
        for i, item in enumerate(skip_items):
            _render_merged_row(i, item, badge_override="홈페이지")

    # ── 보험 (엑셀 데이터 기반) ──
    ins_scope = st.session_state.get("ins_type_select", "전체")
    if parsed.insurances:
        # 보험사 이름이 있는 건만
        all_ins = [ins for ins in parsed.insurances if ins.name]

        # 유지/전체 필터
        if ins_scope == "유지":
            filtered_ins = [ins for ins in all_ins if ins.status in ("유지", "휴면")]
        else:
            filtered_ins = all_ins

        if filtered_ins:
            st.divider()
            st.markdown(f"### 🛡️ 보험 ({len(filtered_ins)}건, {ins_scope} 모드)")

            # 보험사별 그룹
            from collections import OrderedDict
            ins_groups = OrderedDict()
            for ins in filtered_ins:
                ins_groups.setdefault(ins.name, []).append(ins)

            from modules.config_loader import get_insurance_info
            for comp, entries in ins_groups.items():
                manual_info = get_insurance_info(comp) or {}
                tel = manual_info.get("고객센터", "")

                maintain = [e for e in entries if e.status in ("유지", "휴면")]
                cancel = [e for e in entries if e.status not in ("유지", "휴면") and e.status]
                no_status = [e for e in entries if not e.status]

                st.markdown(
                    f"<div style='margin:12px 0 4px;font-size:15px;'>"
                    f"<b>{comp}</b>"
                    f"<span style='color:#9ca3af;font-size:12px;margin-left:8px;'>{tel}</span>"
                    f"</div>", unsafe_allow_html=True)

                if maintain:
                    nos = ", ".join(e.policy_no for e in maintain if e.policy_no)
                    st.markdown(
                        f"<div style='margin-left:16px;font-size:13px;'>"
                        f"<span style='background:#3b82f622;color:#3b82f6;padding:1px 6px;"
                        f"border-radius:3px;font-size:11px;'>유지건 {len(maintain)}건</span> "
                        f"예상해지환급금증명서<br>"
                        f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                        f"</div>", unsafe_allow_html=True)

                if cancel:
                    nos = ", ".join(e.policy_no for e in cancel if e.policy_no)
                    statuses = "/".join(sorted(set(e.status for e in cancel if e.status)))
                    st.markdown(
                        f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                        f"<span style='background:#ef444422;color:#ef4444;padding:1px 6px;"
                        f"border-radius:3px;font-size:11px;'>해지건 {len(cancel)}건</span> "
                        f"해지확인서 ({statuses})<br>"
                        f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                        f"</div>", unsafe_allow_html=True)

                if no_status:
                    st.markdown(
                        f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                        f"<span style='background:#f59e0b22;color:#f59e0b;padding:1px 6px;"
                        f"border-radius:3px;font-size:11px;'>상태 미입력 {len(no_status)}건</span>"
                        f"</div>", unsafe_allow_html=True)
        elif all_ins:
            st.divider()
            st.markdown("### 🛡️ 보험")
            st.warning(f"보험사 {len(all_ins)}개 있으나 상태(유지/실효 등)가 비어있습니다. 엑셀에 상태를 채워주세요.")

    st.divider()

    # ── 하단 버튼 ──
    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.button("➡️ 채권사 서류 탭으로 전달", type="primary",
                   key="transfer_to_creditor", on_click=_on_transfer_click,
                   use_container_width=True)

    with col2:
        if merged:
            manual_pdf = _build_manual_pdf(merged, parsed.person.name, parsed.insurances)
            if manual_pdf:
                st.download_button("📋 발급매뉴얼 다운", data=manual_pdf,
                    file_name=f"발급매뉴얼_{parsed.person.name}.pdf",
                    mime="application/pdf", use_container_width=True)
            else:
                st.button("📋 발급매뉴얼 다운", disabled=True, key="manual_dis", use_container_width=True)

    # 보험 고객요청 건 확인 (유지건 중 DB고객요청 + 해지건)
    has_ins_customer = False
    if parsed.insurances:
        ins_scope_btn = st.session_state.get("ins_type_select", "전체")
        from modules.config_loader import get_insurance_info as _get_ins
        for ins in parsed.insurances:
            if not ins.name:
                continue
            if ins_scope_btn == "유지" and ins.status not in ("유지", "휴면"):
                continue
            if ins.status not in ("유지", "휴면"):
                has_ins_customer = True
                break
            m = _get_ins(ins.name) or {}
            if m.get("발급방법", "").strip() == "고객요청":
                has_ins_customer = True
                break

    # 고객요청 문자 (바로 복사 가능)
    if customer_items or has_ins_customer:
        sms_text = _build_customer_sms(customer_items, parsed.person.name, parsed.insurances)
        with st.expander("📱 고객요청 문자", expanded=True):
            st.code(sms_text, language=None)
            # 클립보드 복사 버튼
            import streamlit.components.v1 as components
            escaped = sms_text.replace('\\', '\\\\').replace('`', '\\`').replace("'", "\\'").replace('\n', '\\n')
            components.html(f"""
            <button onclick="navigator.clipboard.writeText('{escaped}').then(()=>{{this.innerText='✅ 복사됨!';setTimeout(()=>this.innerText='📋 문자 복사',1500)}})"
            style="background:#3b82f6;color:white;border:none;padding:10px 24px;border-radius:8px;
            font-size:15px;font-weight:600;cursor:pointer;width:100%;margin-top:4px;">
            📋 문자 복사</button>
            """, height=55)


# ═══════════════════════════════════════
# 기간 설정 UI
# ═══════════════════════════════════════

def _render_period_settings():
    """법률사무소 선택 + 거래내역 기간 설정"""
    firms = load_law_firms()
    firm_names = ["— 직접입력 —"] + sorted(firms.keys())

    st.markdown("### ⚙️ 거래내역 기간 설정")

    c1, c2 = st.columns([1, 1])

    with c1:
        selected_firm = st.selectbox(
            "법률사무소", firm_names, key="law_firm_select",
            help="사무소 선택 시 회생/파산별 기간이 자동 설정됩니다."
        )

    with c2:
        case_type = st.radio("사건 종류", ["회생", "파산"], key="case_type_radio", horizontal=True)

    # 위임일자 기준일
    warrant_date_str = st.session_state.get("_warrant_date", "")
    if warrant_date_str:
        try:
            parts = warrant_date_str.split(".")
            base_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
        except:
            base_date = date.today()
    else:
        base_date = date.today()

    # 사무소 데이터
    bank_years = None
    card_years = None
    ins_type = "유지"

    if selected_firm != "— 직접입력 —" and selected_firm in firms:
        firm_data = firms[selected_firm]
        case_data = firm_data.get(case_type, {})
        bank_years = case_data.get("통장")
        card_years = case_data.get("카드")
        ins_type = case_data.get("보험", "유지")

    # 특수 형식 안내
    if isinstance(bank_years, str) or isinstance(card_years, str):
        note = bank_years if isinstance(bank_years, str) else card_years
        st.info(f"ℹ️ 이 사무소는 특수 기간 형식입니다: **{note}** — 기타(직접입력)으로 조정해주세요.")

    # ── 통장 기간 ──
    st.markdown("<small style='color:#6b7280;font-weight:600;'>통장 거래내역 기간</small>", unsafe_allow_html=True)
    bank_mode_options = ["최근 N년"] + (["기타 (직접입력)"] if True else [])
    bc1, bc2 = st.columns([0.4, 1])
    with bc1:
        bank_mode = st.radio("통장 모드", bank_mode_options, key="bank_mode",
                             horizontal=True, label_visibility="collapsed")
    if bank_mode == "최근 N년":
        default_bank = bank_years if isinstance(bank_years, int) else 1
        with bc2:
            bank_y = st.number_input("최근 N년", min_value=1, max_value=10,
                                      value=default_bank, key="bank_years_input",
                                      label_visibility="collapsed")
        bank_start, bank_end = _calc_period(bank_y, base_date)
    else:
        dc1, dc2 = st.columns(2)
        with dc1:
            bank_start = st.date_input("통장 시작일", value=base_date - relativedelta(years=1) + relativedelta(days=1),
                                        key="bank_custom_from")
        with dc2:
            bank_end = st.date_input("통장 종료일", value=base_date,
                                      key="bank_custom_to")

    # ── 카드 기간 ──
    st.markdown("<small style='color:#6b7280;font-weight:600;'>카드 거래내역 기간</small>", unsafe_allow_html=True)
    cc1, cc2 = st.columns([0.4, 1])
    with cc1:
        card_mode = st.radio("카드 모드", bank_mode_options, key="card_mode",
                             horizontal=True, label_visibility="collapsed")
    if card_mode == "최근 N년":
        default_card = card_years if isinstance(card_years, int) else 1
        with cc2:
            card_y = st.number_input("최근 N년", min_value=1, max_value=10,
                                      value=default_card, key="card_years_input",
                                      label_visibility="collapsed")
        card_start, card_end = _calc_period(card_y, base_date)
    else:
        dc3, dc4 = st.columns(2)
        with dc3:
            card_start = st.date_input("카드 시작일", value=base_date - relativedelta(years=1) + relativedelta(days=1),
                                        key="card_custom_from")
        with dc4:
            card_end = st.date_input("카드 종료일", value=base_date,
                                      key="card_custom_to")

    # ── 보험 ──
    # "유지,해지"는 "전체"로 통일
    if ins_type == "유지,해지":
        ins_type = "전체"
    ins_options = ["유지", "전체"]
    default_ins_idx = ins_options.index(ins_type) if ins_type in ins_options else 0
    st.selectbox("보험", ins_options, index=default_ins_idx, key="ins_type_select")

    # session_state에 저장
    st.session_state["_period_bank_from"] = bank_start
    st.session_state["_period_bank_to"] = bank_end
    st.session_state["_period_card_from"] = card_start
    st.session_state["_period_card_to"] = card_end

    if bank_start and bank_end:
        st.caption(f"📅 통장: {bank_start.strftime('%Y.%m.%d')} ~ {bank_end.strftime('%Y.%m.%d')} "
                   f"/ 카드: {card_start.strftime('%Y.%m.%d')} ~ {card_end.strftime('%Y.%m.%d')} "
                   f"(기준일: {base_date.strftime('%Y.%m.%d')})")


# ═══════════════════════════════════════
# 분류
# ═══════════════════════════════════════

def _get_method(name):
    info = get_issue_info(name)
    return info.get("발급방법", "").strip() if info else ""


def _get_transfer_items(merged):
    """위임장 대상 (고객요청·홈페이지 제외), 방문/현장 → 미등록 → 등기 → 팩스"""
    skip_methods = {"고객요청", "홈페이지"}
    ORDER = {"현장발급": 0, "방문발급": 0, "": 1, "등기": 2, "팩스": 3}
    result = [i for i in merged if _get_method(i["name"]) not in skip_methods]
    result.sort(key=lambda x: ORDER.get(_get_method(x["name"]), 1))
    return result


def _get_customer_request_items(merged):
    return [i for i in merged if _get_method(i["name"]) == "고객요청"]


def _get_skip_items(merged):
    return [i for i in merged if _get_method(i["name"]) == "홈페이지"]


# ═══════════════════════════════════════
# 행 렌더
# ═══════════════════════════════════════

def _render_merged_row(i, item, badge_override=None):
    cols = st.columns([0.3, 1.3, 1.8, 1.0, 0.6])
    cols[0].markdown(f"<div style='padding-top:6px;color:#6b7280;'>{i+1}</div>", unsafe_allow_html=True)
    cols[1].markdown(f"<div style='padding-top:6px;font-weight:600;'>{item['name']}</div>", unsafe_allow_html=True)

    doc_badges = []
    for d in item["docs"]:
        colors = {"부채증명서": "#3b82f6", "카드부채증명서": "#8b5cf6",
                  "통장거래내역": "#10b981", "카드거래내역": "#f59e0b"}
        c = colors.get(d, "#6b7280")
        doc_badges.append(f'<span style="background:{c}22;color:{c};padding:2px 8px;border-radius:4px;font-size:11px;margin-right:4px;">{d}</span>')
    cols[2].markdown(f"<div style='padding-top:4px;'>{''.join(doc_badges)}</div>", unsafe_allow_html=True)

    acct = item.get("accounts", "")
    cols[3].markdown(f"<div style='padding-top:6px;font-size:11px;color:#6b7280;'>{acct or '-'}</div>", unsafe_allow_html=True)

    method = badge_override or _get_method(item["name"]) or "-"
    mc = {"팩스": "#ef4444", "등기": "#10b981", "방문발급": "#3b82f6",
          "현장발급": "#3b82f6", "고객요청": "#f59e0b", "홈페이지": "#8b5cf6"}.get(method, "#6b7280")
    cols[4].markdown(f"<div style='padding-top:6px;color:{mc};font-size:12px;font-weight:600;'>{method}</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════
# 발급매뉴얼 PDF
# ═══════════════════════════════════════

def _build_manual_pdf(merged, client_name, insurances=None):
    try:
        from modules.pdf_engine import build_manual_cover
        all_names = [item["name"] for item in merged]

        # 보험 건 수집 (홈페이지 / 고객요청 분리)
        ins_homepage = []
        ins_customer = []
        if insurances:
            ins_scope = st.session_state.get("ins_type_select", "전체")
            from modules.config_loader import get_insurance_info
            from collections import OrderedDict
            hp_groups = OrderedDict()
            cr_groups = OrderedDict()

            for ins in insurances:
                if not ins.name or not ins.status:
                    continue
                if ins_scope == "유지" and ins.status not in ("유지", "휴면"):
                    continue

                info = get_insurance_info(ins.name) or {}
                db_method = info.get("발급방법", "").strip()

                if ins.status in ("유지", "휴면"):
                    if db_method == "홈페이지":
                        hp_groups.setdefault(ins.name, []).append(ins)
                    else:
                        cr_groups.setdefault(ins.name, []).append(ins)
                else:
                    # 해지건(실효/소멸/만기) → 전부 고객요청
                    cr_groups.setdefault(ins.name, []).append(ins)

            for comp, entries in hp_groups.items():
                info = get_insurance_info(comp) or {}
                ins_homepage.append({
                    "name": comp,
                    "count": len(entries),
                    "tel": info.get("고객센터", ""),
                    "route": info.get("경로", ""),
                    "doc_type": "예상해지환급금증명서",
                    "policy_nos": [e.policy_no for e in entries if e.policy_no],
                })

            for comp, entries in cr_groups.items():
                info = get_insurance_info(comp) or {}
                maintain = [e for e in entries if e.status in ("유지", "휴면")]
                cancel = [e for e in entries if e.status not in ("유지", "휴면")]
                docs = []
                if maintain:
                    docs.append(f"예상해지환급금증명서 {len(maintain)}건")
                if cancel:
                    statuses = "/".join(sorted(set(e.status for e in cancel)))
                    docs.append(f"해지확인서 {len(cancel)}건 ({statuses})")
                ins_customer.append({
                    "name": comp,
                    "count": len(entries),
                    "tel": info.get("고객센터", ""),
                    "doc_type": ", ".join(docs),
                    "policy_nos": [e.policy_no for e in entries if e.policy_no],
                })

        today = date.today().strftime("%Y.%m.%d")
        doc = build_manual_cover(all_names, client_name, today,
                                 ins_homepage=ins_homepage, ins_customer=ins_customer)
        if doc.page_count > 0:
            pdf_bytes = doc.tobytes()
            doc.close()
            return pdf_bytes
        doc.close()
    except Exception as e:
        print(f"[DEBUG] 매뉴얼 PDF 에러: {e}")
    return None


# ═══════════════════════════════════════
# 고객요청 문자 (기간 포함)
# ═══════════════════════════════════════

def _build_customer_sms(customer_items, client_name, insurances=None):
    """고객요청 문자 템플릿 — 기간 + 보험 포함"""
    # 팩스번호
    fax_number = ""
    agent_name = st.session_state.get("agent_select", "")
    if agent_name and agent_name != "— 선택 —":
        staff = load_staff()
        agent_data = staff.get(agent_name, {})
        fax_number = agent_data.get("fax", "")

    # 기간
    bank_from = st.session_state.get("_period_bank_from")
    bank_to = st.session_state.get("_period_bank_to")
    card_from = st.session_state.get("_period_card_from")
    card_to = st.session_state.get("_period_card_to")

    def _fmt(d):
        return d.strftime("%Y.%m.%d") if d else ""

    lines = []
    lines.append(f"[{client_name}님 준비사항]")
    if fax_number:
        lines.append(f"팩스번호 : {fax_number}")
    lines.append("")

    idx = 1
    for item in customer_items:
        name = item["name"]
        info = get_issue_info(name)
        tel = ""
        if info:
            tel_raw = info.get("고객센터", "")
            tel = tel_raw.split("\n")[0].strip() if tel_raw else ""

        if tel:
            lines.append(f"{idx}. {name}({tel})")
        else:
            lines.append(f"{idx}. {name}")

        for doc in item["docs"]:
            if doc == "통장거래내역" and bank_from and bank_to:
                lines.append(f"* {doc} ({_fmt(bank_from)} ~ {_fmt(bank_to)})")
            elif doc == "카드거래내역" and card_from and card_to:
                lines.append(f"* {doc} ({_fmt(card_from)} ~ {_fmt(card_to)})")
            else:
                lines.append(f"* {doc}")

        # 증권사는 잔고증명서 자동 추가
        if "증권" in name:
            lines.append("* 잔고증명서")

        lines.append("")
        idx += 1

    # ── 보험 고객요청 건 추가 (엑셀 데이터 기반) ──
    # 유지건: DB가 "고객요청"인 보험사만 문자에 포함
    # 해지건(실효/소멸/만기): DB 무관 전부 문자에 포함
    if insurances:
        ins_scope = st.session_state.get("ins_type_select", "전체")
        if ins_scope == "유지":
            ins_filtered = [i for i in insurances if i.name and i.status in ("유지", "휴면")]
        else:
            ins_filtered = [i for i in insurances if i.name]

        if ins_filtered:
            from collections import OrderedDict
            from modules.config_loader import get_insurance_info
            ins_groups = OrderedDict()
            for ins in ins_filtered:
                ins_groups.setdefault(ins.name, []).append(ins)

            for comp, entries in ins_groups.items():
                manual_info = get_insurance_info(comp) or {}
                tel = manual_info.get("고객센터", "")
                db_method = manual_info.get("발급방법", "").strip()

                maintain = [e for e in entries if e.status in ("유지", "휴면")]
                cancel = [e for e in entries if e.status not in ("유지", "휴면")]

                # 유지건: DB가 고객요청인 보험사만 포함 (홈페이지는 매뉴얼로)
                maintain_sms = maintain if db_method == "고객요청" else []
                # 해지건: 전부 포함
                cancel_sms = cancel

                if not maintain_sms and not cancel_sms:
                    continue

                if tel:
                    lines.append(f"{idx}. {comp}({tel})")
                else:
                    lines.append(f"{idx}. {comp}")

                if maintain_sms:
                    lines.append(f"* 유지건 {len(maintain_sms)}건 예상해지환급금증명서")
                    for e in maintain_sms:
                        lines.append(f"  ({e.policy_no})")

                if cancel_sms:
                    lines.append(f"* 해지건 {len(cancel_sms)}건 해지확인서(해지환급금 기재)")
                    for e in cancel_sms:
                        lines.append(f"  ({e.policy_no})")

                lines.append("")
                idx += 1

    lines.append("각 고객센터 상담사 연결 후 팩스요청해주시면 됩니다")

    return "\n".join(lines)


# ═══════════════════════════════════════
# 병합 로직
# ═══════════════════════════════════════

def _merge_by_institution(parsed: ParsedExcel) -> list:
    """기관명 기준으로 채권목록 + 은행 + 카드 병합"""
    groups = {}

    CARD_TO_BANK = {
        "하나카드": "하나은행", "전북카드": "전북은행", "우리카드": "우리은행",
        "농협카드": "농협은행", "농축협카드": "농축협", "기업카드": "기업은행",
    }

    def _resolve(name):
        return CARD_TO_BANK.get(name, name)

    def _get_or_create(name):
        r = _resolve(name)
        if r not in groups:
            m = match_creditor(r)
            groups[r] = {"docs": [], "accounts": [], "match": m}
        return r

    def _is_card(name):
        return "카드" in name

    for cred in parsed.delegation_creditors:
        key = _get_or_create(cred.name)
        if _is_card(cred.name):
            if "카드부채증명서" not in groups[key]["docs"]:
                groups[key]["docs"].append("카드부채증명서")
        else:
            if "부채증명서" not in groups[key]["docs"]:
                groups[key]["docs"].append("부채증명서")

    for bank in parsed.banks:
        key = _get_or_create(bank.name)
        if "통장거래내역" not in groups[key]["docs"]:
            groups[key]["docs"].append("통장거래내역")
        if bank.account:
            groups[key]["accounts"].append(bank.account)

    for card in parsed.cards:
        key = _get_or_create(card.name)
        if "카드거래내역" not in groups[key]["docs"]:
            groups[key]["docs"].append("카드거래내역")

    return [{"name": n, "docs": d["docs"], "accounts": ", ".join(d["accounts"]), "match": d["match"]}
            for n, d in groups.items()]


# ═══════════════════════════════════════
# 뱃지
# ═══════════════════════════════════════

def _match_badge(m: MatchResult) -> str:
    if m.matched:
        if m.match_type == "exact":
            return '<span style="color:#10b981;font-size:11px;">✅</span>'
        elif m.match_type == "alias":
            return f'<span style="color:#3b82f6;font-size:11px;">🔄 {m.matched_key}</span>'
        else:
            return f'<span style="color:#f59e0b;font-size:11px;">🔍 {m.matched_key}</span>'
    return '<span style="color:#ef4444;font-size:11px;">⚠️ 기본</span>'



