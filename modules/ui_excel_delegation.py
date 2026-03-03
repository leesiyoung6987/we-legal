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
    # 화면에서 수정한 값 우선 사용
    name = st.session_state.get("ex_p_name", "") or p.name
    ssn = st.session_state.get("ex_p_ssn", "") or p.ssn
    phone = st.session_state.get("ex_p_phone", "") or p.phone
    address = st.session_state.get("ex_p_addr", "") or p.address

    # ssn 분리
    ssn_str = str(ssn) if ssn else ""
    if "-" in ssn_str:
        ssn_front, ssn_back = ssn_str.split("-", 1)
    elif len(ssn_str) > 6:
        ssn_front, ssn_back = ssn_str[:6], ssn_str[6:]
    else:
        ssn_front, ssn_back = ssn_str, ""

    import re
    phone_clean = re.sub(r'\(.*?\)', '', str(phone)).strip() if phone else ""

    if name:
        st.session_state["client_name"] = name
    if ssn_front:
        st.session_state["client_id_front"] = ssn_front.strip()
    if ssn_back:
        st.session_state["client_id_back"] = ssn_back.strip()
    if address:
        st.session_state["client_address"] = address
    if phone_clean:
        st.session_state["client_phone"] = phone_clean

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
    st.session_state["_switch_to_creditor_tab"] = True


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
    _file_key = f"{uploaded.name}_{uploaded.size}"
    if "parsed_excel" not in st.session_state or st.session_state.get("_excel_key") != _file_key:
        with st.spinner("엑셀 파싱 중..."):
            parsed = parse_excel(io.BytesIO(uploaded.read()))
            st.session_state.parsed_excel = parsed
            st.session_state._excel_key = _file_key
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
            _name = st.text_input("이름", value=p.name, key="ex_p_name")
            _ssn = st.text_input("주민번호", value=p.ssn, key="ex_p_ssn")
        with c2:
            _phone = st.text_input("전화번호", value=p.phone, key="ex_p_phone")
            _addr = st.text_input("주소", value=p.address, key="ex_p_addr")
        # 화면 입력값을 parsed에 반영
        if _name: p.name = _name
        if _ssn: p.ssn = _ssn
        if _phone: p.phone = _phone
        if _addr: p.address = _addr

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
            filtered_ins = [ins for ins in all_ins if ins.status == "유지"]
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

                # 상태별 분류
                s_maintain = [e for e in entries if e.status == "유지"]
                s_dormant = [e for e in entries if e.status == "휴면"]
                s_cancel = [e for e in entries if e.status not in ("유지", "휴면") and e.status]
                no_status = [e for e in entries if not e.status]

                st.markdown(
                    f"<div style='margin:12px 0 4px;font-size:15px;'>"
                    f"<b>{comp}</b>"
                    f"<span style='color:#9ca3af;font-size:12px;margin-left:8px;'>{tel}</span>"
                    f"</div>", unsafe_allow_html=True)

                if s_maintain:
                    nos = ", ".join(e.policy_no for e in s_maintain if e.policy_no)
                    st.markdown(
                        f"<div style='margin-left:16px;font-size:13px;'>"
                        f"<span style='background:#3b82f622;color:#3b82f6;padding:1px 6px;"
                        f"border-radius:3px;font-size:11px;'>유지 {len(s_maintain)}건</span> "
                        f"예상해지환급금증명서<br>"
                        f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                        f"</div>", unsafe_allow_html=True)

                if s_dormant:
                    nos = ", ".join(e.policy_no for e in s_dormant if e.policy_no)
                    st.markdown(
                        f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                        f"<span style='background:#f59e0b22;color:#f59e0b;padding:1px 6px;"
                        f"border-radius:3px;font-size:11px;'>휴면 {len(s_dormant)}건</span> "
                        f"예상해지환급금증명서<br>"
                        f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                        f"</div>", unsafe_allow_html=True)

                if s_cancel:
                    # 상태별 분리 표시 (실효/소멸/해약 등)
                    from collections import OrderedDict as OD
                    cancel_by_status = OD()
                    for e in s_cancel:
                        cancel_by_status.setdefault(e.status, []).append(e)
                    for st_name, st_entries in cancel_by_status.items():
                        nos = ", ".join(e.policy_no for e in st_entries if e.policy_no)
                        st.markdown(
                            f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                            f"<span style='background:#ef444422;color:#ef4444;padding:1px 6px;"
                            f"border-radius:3px;font-size:11px;'>{st_name} {len(st_entries)}건</span> "
                            f"해지확인서<br>"
                            f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                            f"</div>", unsafe_allow_html=True)

                if no_status:
                    st.markdown(
                        f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                        f"<span style='background:#6b728022;color:#6b7280;padding:1px 6px;"
                        f"border-radius:3px;font-size:11px;'>상태 미입력 {len(no_status)}건</span>"
                        f"</div>", unsafe_allow_html=True)
        elif all_ins:
            st.divider()
            st.markdown("### 🛡️ 보험")
            st.warning(f"보험사 {len(all_ins)}개 있으나 상태(유지/실효 등)가 비어있습니다. 엑셀에 상태를 채워주세요.")

        # ── 배우자 보험 ──
        all_spouse_ins = [ins for ins in parsed.spouse_insurances if ins.name]
        if all_spouse_ins:
            if ins_scope == "유지":
                spouse_ins = [ins for ins in all_spouse_ins if ins.status == "유지"]
            else:
                spouse_ins = all_spouse_ins

            if spouse_ins:
                st.divider()
                st.markdown(f"### 💑 배우자 보험 ({len(spouse_ins)}건, {ins_scope} 모드)")

                from collections import OrderedDict
                sp_groups = OrderedDict()
                for ins in spouse_ins:
                    sp_groups.setdefault(ins.name, []).append(ins)

                from modules.config_loader import get_insurance_info
                for comp, entries in sp_groups.items():
                    manual_info = get_insurance_info(comp) or {}
                    tel = manual_info.get("고객센터", "")

                    s_maintain = [e for e in entries if e.status == "유지"]
                    s_dormant = [e for e in entries if e.status == "휴면"]
                    s_cancel = [e for e in entries if e.status not in ("유지", "휴면") and e.status]

                    st.markdown(
                        f"<div style='margin:12px 0 4px;font-size:15px;'>"
                        f"<b>{comp}</b>"
                        f"<span style='color:#9ca3af;font-size:12px;margin-left:8px;'>{tel}</span>"
                        f"</div>", unsafe_allow_html=True)

                    if s_maintain:
                        nos = ", ".join(e.policy_no for e in s_maintain if e.policy_no)
                        st.markdown(
                            f"<div style='margin-left:16px;font-size:13px;'>"
                            f"<span style='background:#3b82f622;color:#3b82f6;padding:1px 6px;"
                            f"border-radius:3px;font-size:11px;'>유지 {len(s_maintain)}건</span> "
                            f"예상해지환급금증명서<br>"
                            f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                            f"</div>", unsafe_allow_html=True)

                    if s_dormant:
                        nos = ", ".join(e.policy_no for e in s_dormant if e.policy_no)
                        st.markdown(
                            f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                            f"<span style='background:#f59e0b22;color:#f59e0b;padding:1px 6px;"
                            f"border-radius:3px;font-size:11px;'>휴면 {len(s_dormant)}건</span> "
                            f"예상해지환급금증명서<br>"
                            f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                            f"</div>", unsafe_allow_html=True)

                    if s_cancel:
                        # 상태별 분리 표시 (실효/소멸/해약 등)
                        from collections import OrderedDict as OD2
                        cancel_by_status = OD2()
                        for e in s_cancel:
                            cancel_by_status.setdefault(e.status, []).append(e)
                        for st_name, st_entries in cancel_by_status.items():
                            nos = ", ".join(e.policy_no for e in st_entries if e.policy_no)
                            st.markdown(
                                f"<div style='margin-left:16px;font-size:13px;margin-top:4px;'>"
                                f"<span style='background:#ef444422;color:#ef4444;padding:1px 6px;"
                                f"border-radius:3px;font-size:11px;'>{st_name} {len(st_entries)}건</span> "
                                f"해지확인서<br>"
                                f"<span style='color:#9ca3af;font-size:11px;margin-left:8px;'>{nos}</span>"
                                f"</div>", unsafe_allow_html=True)

    st.divider()

    # ── 하단 버튼 ──
    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.button("➡️ 채권사 서류 탭으로 전달", type="primary",
                   key="transfer_to_creditor", on_click=_on_transfer_click,
                   use_container_width=True)

    with col2:
        if merged:
            manual_pdf = _build_manual_pdf(merged, parsed.person.name, parsed.insurances, parsed.spouse_insurances)
            if manual_pdf:
                st.download_button("📋 발급매뉴얼 다운", data=manual_pdf,
                    file_name=f"발급매뉴얼_{parsed.person.name}.pdf",
                    mime="application/pdf", use_container_width=True)
            else:
                st.button("📋 발급매뉴얼 다운", disabled=True, key="manual_dis", use_container_width=True)

    # 보험 고객요청 건 확인 (유지건 중 DB고객요청 + 해지건) - 본인 + 배우자
    def _has_ins_customer_items(ins_list):
        if not ins_list:
            return False
        ins_scope_btn = st.session_state.get("ins_type_select", "전체")
        from modules.config_loader import get_insurance_info as _chk_ins
        from collections import OrderedDict
        groups = OrderedDict()
        for ins in ins_list:
            if not ins.name:
                continue
            if ins_scope_btn == "유지" and ins.status != "유지":
                continue
            groups.setdefault(ins.name, []).append(ins)
        for comp, entries in groups.items():
            m = _chk_ins(comp) or {}
            db_method = m.get("발급방법", "").strip()
            maintain = [e for e in entries if e.status in ("유지", "휴면")]
            cancel = [e for e in entries if e.status not in ("유지", "휴면")]
            if ins_scope_btn == "전체" and maintain and cancel:
                return True  # 혼합 → 고객요청
            if cancel:
                return True
            if maintain and db_method == "고객요청":
                return True
        return False

    has_ins_customer = _has_ins_customer_items(parsed.insurances)
    has_spouse_ins_customer = _has_ins_customer_items(parsed.spouse_insurances)

    # 고객요청 문자 (바로 복사 가능)
    if customer_items or has_ins_customer or has_spouse_ins_customer:
        sms_text = _build_customer_sms(customer_items, parsed.person.name, parsed.insurances, parsed.spouse_insurances)
        with st.expander("📱 고객요청 문자", expanded=True):
            # 클립보드 복사 버튼 (상단)
            import streamlit.components.v1 as components
            sms_json = json.dumps(sms_text, ensure_ascii=False)
            components.html(f"""
            <script>var _sms={sms_json};</script>
            <button onclick="navigator.clipboard.writeText(_sms).then(()=>{{this.innerText='✅ 복사됨!';setTimeout(()=>this.innerText='📋 문자 복사',1500)}})"
            style="background:#3b82f6;color:white;border:none;padding:10px 24px;border-radius:8px;
            font-size:15px;font-weight:600;cursor:pointer;width:100%;margin-top:4px;">
            📋 문자 복사</button>
            """, height=55)
            st.code(sms_text, language=None)


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

def _build_manual_pdf(merged, client_name, insurances=None, spouse_insurances=None):
    try:
        from modules.pdf_engine import build_manual_cover
        all_names = [item["name"] for item in merged]

        # 보험 건 수집 (홈페이지 / 고객요청 분리)
        ins_homepage = []
        ins_customer = []
        sp_ins_homepage = []
        sp_ins_customer = []

        def _collect_insurance(ins_list, hp_out, cr_out):
            if not ins_list:
                return
            ins_scope = st.session_state.get("ins_type_select", "전체")
            from modules.config_loader import get_insurance_info
            from collections import OrderedDict

            # 보험사별 그룹핑 (필터 전)
            all_by_comp = OrderedDict()
            for ins in ins_list:
                if not ins.name or not ins.status:
                    continue
                all_by_comp.setdefault(ins.name, []).append(ins)

            hp_groups = OrderedDict()
            cr_groups = OrderedDict()

            for comp, all_entries in all_by_comp.items():
                info = get_insurance_info(comp) or {}
                db_method = info.get("발급방법", "").strip()

                if ins_scope == "유지":
                    # 유지모드: 유지건만, DB 발급방법대로
                    entries = [e for e in all_entries if e.status == "유지"]
                    if not entries:
                        continue
                    if db_method == "홈페이지":
                        hp_groups[comp] = entries
                    else:
                        cr_groups[comp] = entries
                else:
                    # 전체모드: 유지만 있으면 DB대로, 유지+기타 혼재면 전부 고객요청
                    maintain = [e for e in all_entries if e.status in ("유지", "휴면")]
                    others = [e for e in all_entries if e.status not in ("유지", "휴면")]

                    if others:
                        # 혼재 → 전부 고객요청
                        cr_groups[comp] = all_entries
                    elif maintain:
                        # 유지만 → DB 발급방법대로
                        if db_method == "홈페이지":
                            hp_groups[comp] = maintain
                        else:
                            cr_groups[comp] = maintain

            for comp, entries in hp_groups.items():
                info = get_insurance_info(comp) or {}
                hp_out.append({
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
                cr_out.append({
                    "name": comp,
                    "count": len(entries),
                    "tel": info.get("고객센터", ""),
                    "doc_type": ", ".join(docs),
                    "policy_nos": [e.policy_no for e in entries if e.policy_no],
                })

        _collect_insurance(insurances, ins_homepage, ins_customer)
        _collect_insurance(spouse_insurances, sp_ins_homepage, sp_ins_customer)

        today = date.today().strftime("%Y.%m.%d")
        doc = build_manual_cover(all_names, client_name, today,
                                 ins_homepage=ins_homepage, ins_customer=ins_customer,
                                 sp_ins_homepage=sp_ins_homepage, sp_ins_customer=sp_ins_customer)
        if doc.page_count > 0:
            pdf_bytes = doc.tobytes()
            doc.close()
            return pdf_bytes
        doc.close()
    except Exception as e:
        import traceback
        st.error(f"매뉴얼 PDF 생성 실패: {e}")
        print(f"[DEBUG] 매뉴얼 PDF 에러: {traceback.format_exc()}")
    return None


# ═══════════════════════════════════════
# 고객요청 문자 (기간 포함)
# ═══════════════════════════════════════

def _build_customer_sms(customer_items, client_name, insurances=None, spouse_insurances=None):
    """고객요청 문자 템플릿 — 기간 + 보험(본인+배우자) 포함"""
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
                if item.get("accounts"):
                    lines.append(f"  ({item['accounts']})")
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
    # 유지 모드: DB가 "고객요청"인 보험사만 문자 포함 (홈페이지는 매뉴얼)
    # 전체 모드: 보험사 내 유지만 있으면 DB 분류, 유지+기타 혼합이면 무조건 고객요청
    def _build_ins_sms_lines(ins_list, start_idx):
        """보험 SMS 라인 생성. (lines_list, next_idx) 반환"""
        if not ins_list:
            return [], start_idx
        ins_scope = st.session_state.get("ins_type_select", "전체")
        if ins_scope == "유지":
            ins_filtered = [i for i in ins_list if i.name and i.status in ("유지", "휴면")]
        else:
            ins_filtered = [i for i in ins_list if i.name]

        if not ins_filtered:
            return [], start_idx

        from collections import OrderedDict
        from modules.config_loader import get_insurance_info
        ins_groups = OrderedDict()
        for ins in ins_filtered:
            ins_groups.setdefault(ins.name, []).append(ins)

        result_lines = []
        cur_idx = start_idx
        for comp, entries in ins_groups.items():
            manual_info = get_insurance_info(comp) or {}
            tel = manual_info.get("고객센터", "")
            db_method = manual_info.get("발급방법", "").strip()

            maintain = [e for e in entries if e.status in ("유지", "휴면")]
            cancel = [e for e in entries if e.status not in ("유지", "휴면")]

            if ins_scope == "유지":
                # 유지 모드: DB 발급방법에 따라 분류
                maintain_sms = maintain if db_method == "고객요청" else []
                cancel_sms = []  # 유지 모드에선 해지건 없음
            else:
                # 전체 모드: 유지+기타 혼합이면 무조건 고객요청
                if maintain and cancel:
                    # 혼합 → 전부 고객요청
                    maintain_sms = maintain
                    cancel_sms = cancel
                elif maintain and not cancel:
                    # 유지만 → DB 분류
                    maintain_sms = maintain if db_method == "고객요청" else []
                    cancel_sms = []
                else:
                    # 해지만
                    maintain_sms = []
                    cancel_sms = cancel

            if not maintain_sms and not cancel_sms:
                continue

            if tel:
                result_lines.append(f"{cur_idx}. {comp}({tel})")
            else:
                result_lines.append(f"{cur_idx}. {comp}")

            if maintain_sms:
                # 상태별 그룹 (유지, 휴면 등)
                from collections import OrderedDict
                status_groups = OrderedDict()
                for e in maintain_sms:
                    status_groups.setdefault(e.status or "유지", []).append(e)
                for st_name, st_entries in status_groups.items():
                    result_lines.append(f"* {st_name}건 {len(st_entries)}건 예상해지환급금증명서")
                    for e in st_entries:
                        if e.policy_no:
                            result_lines.append(f"  ({e.policy_no})")

            if cancel_sms:
                # 상태별 그룹 (실효, 소멸, 해약 등)
                from collections import OrderedDict
                cancel_groups = OrderedDict()
                for e in cancel_sms:
                    cancel_groups.setdefault(e.status or "해지", []).append(e)
                for st_name, st_entries in cancel_groups.items():
                    result_lines.append(f"* {st_name}건 {len(st_entries)}건 해지확인서(해지환급금 기재)")
                    for e in st_entries:
                        if e.policy_no:
                            result_lines.append(f"  ({e.policy_no})")

            result_lines.append("")
            cur_idx += 1

        return result_lines, cur_idx

    # 본인 보험
    ins_lines, idx = _build_ins_sms_lines(insurances, idx)
    lines.extend(ins_lines)

    # 배우자 보험 (번호 1부터 재시작)
    if spouse_insurances:
        sp_lines, _ = _build_ins_sms_lines(spouse_insurances, 1)
        if sp_lines:
            lines.append(f"[{client_name}님의 배우자 요청사항]")
            lines.append("")
            lines.extend(sp_lines)

    lines.append("각 고객센터 상담사 연결 후 팩스요청해주시면 됩니다")

    return "\n".join(lines)


# ═══════════════════════════════════════
# 병합 로직
# ═══════════════════════════════════════

def _merge_by_institution(parsed: ParsedExcel) -> list:
    """기관명 기준으로 채권목록 + 은행 + 카드 병합"""
    groups = {}

    # 카드 → 은행 병합 (통장/카드 같이 발급하는 곳만)
    CARD_TO_BANK = {
        "전북카드": "전북은행", "우리카드": "우리은행",
        "농협카드": "농협은행", "농축협카드": "농축협", "기업카드": "기업은행",
    }

    # 은행 약칭 → 정식명
    BANK_ALIASES = {
        "신한": "신한은행", "국민": "국민은행", "우리": "우리은행", "하나": "하나은행",
        "농협": "농협은행", "기업": "기업은행", "전북": "전북은행", "광주": "광주은행",
        "경남": "경남은행", "대구": "대구은행", "부산": "부산은행", "제주": "제주은행",
        "수협": "수협은행", "산업": "산업은행", "씨티": "씨티은행",
    }

    # 카드 약칭 → 정식명
    CARD_ALIASES = {
        "신한": "신한카드", "삼성": "삼성카드", "현대": "현대카드", "롯데": "롯데카드",
        "비씨": "비씨카드", "하나": "하나카드", "우리": "우리카드", "국민": "국민카드",
        "농협": "농협카드",
    }

    def _normalize_bank(name):
        """은행 시트 이름 정규화"""
        name = name.strip()
        if name in BANK_ALIASES:
            return BANK_ALIASES[name]
        return name

    def _normalize_card(name):
        """카드 시트 이름 정규화"""
        name = name.strip()
        if name in CARD_ALIASES:
            return CARD_ALIASES[name]
        return name

    def _resolve(name):
        # 1차: 정확 매칭
        if name in CARD_TO_BANK:
            return CARD_TO_BANK[name]
        # 2차: 공백 제거 후 매칭
        stripped = name.replace(" ", "")
        if stripped in CARD_TO_BANK:
            return CARD_TO_BANK[stripped]
        # 3차: "OO은행 카드" / "OO은행카드" → "OO은행" (병합 대상만)
        import re
        m = re.match(r'^(.+(?:은행|축협))\s*카드$', name)
        if m:
            bank = m.group(1)
            # 하나은행 카드는 분리
            if bank == "하나은행":
                return "하나카드"
            return bank
        return name

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
        key = _get_or_create(_normalize_bank(bank.name))
        if "통장거래내역" not in groups[key]["docs"]:
            groups[key]["docs"].append("통장거래내역")
        if bank.account:
            groups[key]["accounts"].append(bank.account)

    for card in parsed.cards:
        key = _get_or_create(_normalize_card(card.name))
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



