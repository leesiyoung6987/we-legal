"""
WE Legal Automation - 회생파산 서류 자동화 시스템
v3.1

app.py는 UI 조립 + PDF 생성 트리거만 담당.
로직은 modules/ 에서 처리.
"""

import streamlit as st
import io
from modules.config_loader import load_settings, get_template_path
from modules.ui_sidebar import render_sidebar
from modules.ui_main import render_main
from modules.ui_gov import render_gov_tab
from modules.ui_excel_delegation import render_excel_delegation_tab
from modules.pdf_engine import build_creditor_bundle, build_gov_bundle, merge_documents, build_manual_cover
from modules.config_loader import get_issue_info


# ── 함수 정의 ──

def validate_inputs(sidebar, main):
    """입력값 검증 (채권사 서류). 에러 메시지 리스트 반환."""
    errors = []
    if not sidebar["client"]["name"]:
        errors.append("위임인 성명을 입력해주세요.")
    if not sidebar["client"]["id_front"]:
        errors.append("위임인 주민번호 앞자리를 입력해주세요.")
    if sidebar["agent_name"] == "— 선택 —":
        errors.append("수임인을 선택해주세요.")
    if not main["creditors"]:
        errors.append("채권사를 1개 이상 입력해주세요.")
    if not sidebar["client_id_file"]:
        errors.append("위임인 신분증을 업로드해주세요.")
    template = get_template_path("위임장_기본.pdf")
    if not template.exists():
        errors.append(f"위임장 템플릿이 없습니다: {template}")
    return errors


def validate_gov_inputs(sidebar, gov_data):
    """입력값 검증 (관공서 서류). 에러 메시지 리스트 반환."""
    errors = []
    if not sidebar["client"]["name"]:
        errors.append("위임인 성명을 입력해주세요.")
    if sidebar["agent_name"] == "— 선택 —":
        errors.append("수임인을 선택해주세요.")
    if not gov_data["forms"]:
        errors.append("양식을 1개 이상 선택해주세요.")
    if not sidebar["client_id_file"]:
        errors.append("위임인 신분증을 업로드해주세요.")
    return errors


def _build_client_agent(sidebar, settings):
    """사이드바 데이터에서 client, agent 딕셔너리 구성"""
    agent = sidebar["agent_data"]
    date_fmt = settings.get("date_format", "%Y.%m.%d")
    warrant_date = sidebar["warrant_date"].strftime(date_fmt)

    client = {
        "name": sidebar["client"]["name"],
        "birth": sidebar["client"]["birth"],
        "address": sidebar["client"]["address"],
        "phone": sidebar["client"]["phone"],
        "id_front": sidebar["client"].get("id_front", ""),
        "id_back": sidebar["client"].get("id_back", ""),
    }
    agent_info = {
        "name": agent["name"],
        "birth": agent["birth"],
        "address": agent["address"],
        "phone": agent["phone"],
        "fax": agent["fax"],
        "id_back": agent.get("id_full", "").split("-")[1] if "-" in agent.get("id_full", "") else "",
        "id_full": agent.get("id_full", ""),
        "cert": agent.get("cert", ""),
    }
    
    # 수임인 싸인 경로 설정
    if agent.get("sign"):
        from modules.config_loader import get_id_card_path
        sign_path = get_id_card_path(agent["sign"])
        if sign_path.exists():
            agent_info["sign_path"] = str(sign_path)

    # 신분증 바이트 로드
    client_id_bytes = sidebar["client_id_file"].read() if sidebar["client_id_file"] else None
    agent_id_bytes = sidebar["agent_id_path"].read_bytes() if sidebar["agent_id_path"] else None

    # 인감도장 / 인감증명서 바이트 로드 (선택사항)
    stamp_bytes = sidebar["stamp_file"].read() if sidebar.get("stamp_file") else None
    seal_cert_bytes = sidebar["seal_cert_file"].read() if sidebar.get("seal_cert_file") else None

    return client, agent_info, warrant_date, client_id_bytes, agent_id_bytes, stamp_bytes, seal_cert_bytes


def generate_pdf(sidebar, main, settings):
    """채권사 PDF 생성 및 다운로드 처리"""
    errors = validate_inputs(sidebar, main)
    if errors:
        for e in errors:
            st.error(e)
        return

    client, agent_info, warrant_date, client_id_bytes, agent_id_bytes, stamp_bytes, seal_cert_bytes = _build_client_agent(sidebar, settings)
    template_path = get_template_path("위임장_기본.pdf")

    # 채권사별 PDF 생성 (팩스/인쇄용 분리)
    creditors = main["creditors"]
    print_bundles = []  # 인쇄용 (팩스 외)
    fax_bundles = {}    # 팩스용 {채권사명: doc}
    progress = st.progress(0, text="PDF 생성 중...")

    # 매뉴얼 커버 페이지 (인쇄용 맨 앞)
    creditor_names = [c["name"] for c in creditors]
    try:
        manual_cover = build_manual_cover(creditor_names, client["name"], warrant_date)
        if manual_cover.page_count > 0:
            print_bundles.append(manual_cover)
    except Exception as e:
        print(f"[DEBUG] 매뉴얼 커버 에러: {e}")

    for idx, cred in enumerate(creditors):
        progress.progress((idx + 1) / len(creditors), text=f"{cred['name']} 처리 중...")
        bundle = build_creditor_bundle(
            template_path, client, agent_info, cred,
            warrant_date, client_id_bytes, agent_id_bytes,
            stamp_bytes=stamp_bytes, seal_cert_bytes=seal_cert_bytes
        )
        
        # 팩스 여부 판별
        issue_info = get_issue_info(cred["name"])
        is_fax = issue_info and issue_info.get("발급방법", "").strip() == "팩스"
        
        if is_fax:
            fax_bundles[cred["name"]] = bundle
        else:
            print_bundles.append(bundle)

    progress.empty()

    # 인쇄용 합본
    if print_bundles:
        merged = merge_documents(print_bundles)
        st.session_state.pdf_bytes = merged.tobytes()
        st.session_state.pdf_filename = f"인쇄용_{client['name']}_{warrant_date.replace('.', '')}.pdf"
        merged.close()
    else:
        st.session_state.pdf_bytes = None

    # 팩스용 ZIP
    if fax_bundles:
        import io, zipfile
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for cred_name, doc in fax_bundles.items():
                pdf_data = doc.tobytes()
                zf.writestr(f"{cred_name}.pdf", pdf_data)
                doc.close()
        zip_buffer.seek(0)
        st.session_state.fax_zip_bytes = zip_buffer.getvalue()
        st.session_state.fax_zip_filename = f"팩스용_{client['name']}_{warrant_date.replace('.', '')}.zip"
    else:
        st.session_state.fax_zip_bytes = None

    count_print = len(print_bundles) - (1 if print_bundles and print_bundles[0].page_count <= 2 else 0)  # 매뉴얼 제외
    count_fax = len(fax_bundles)
    st.success(f"✅ 인쇄용 {count_print}개 + 팩스용 {count_fax}개 생성 완료!")
    st.rerun()


def generate_gov_pdf(sidebar, gov_data, settings):
    """관공서 PDF 생성 및 다운로드 처리"""
    errors = validate_gov_inputs(sidebar, gov_data)
    if errors:
        for e in errors:
            st.error(e)
        return

    client, agent_info, warrant_date, client_id_bytes, agent_id_bytes, stamp_bytes, seal_cert_bytes = _build_client_agent(sidebar, settings)

    progress = st.progress(0, text="관공서 서류 생성 중...")
    
    # 무상거주사실확인서 / 나머지 분리
    gov_forms = []
    musang_forms = []
    for form in gov_data["forms"]:
        if "무상거주" in form.get("form_name", ""):
            musang_forms.append(form)
        else:
            gov_forms.append(form)
    
    # 일반 관공서 서류
    if gov_forms:
        result = build_gov_bundle(
            gov_forms, client, agent_info,
            warrant_date, client_id_bytes, agent_id_bytes,
            stamp_bytes=stamp_bytes
        )
        st.session_state.gov_pdf_bytes = result.tobytes()
        st.session_state.gov_pdf_filename = f"관공서_{client['name']}_{warrant_date.replace('.', '')}.pdf"
        result.close()
    else:
        st.session_state.gov_pdf_bytes = None
    
    # 무상거주사실확인서 (별도)
    if musang_forms:
        result = build_gov_bundle(
            musang_forms, client, agent_info,
            warrant_date, client_id_bytes, agent_id_bytes,
            stamp_bytes=stamp_bytes
        )
        st.session_state.musang_pdf_bytes = result.tobytes()
        st.session_state.musang_pdf_filename = f"무상거주사실확인서_{client['name']}_{warrant_date.replace('.', '')}.pdf"
        result.close()
    else:
        st.session_state.musang_pdf_bytes = None

    progress.empty()
    st.success(f"✅ 관공서 {len(gov_forms)}개 + 무상거주 {len(musang_forms)}개 생성 완료!")
    st.rerun()




# ══════════════════════════════════════
# 메인 실행
# ══════════════════════════════════════

settings = load_settings()
app_cfg = settings.get("app", {})

st.set_page_config(
    page_title=app_cfg.get("title", "WE Legal"),
    page_icon=app_cfg.get("icon", "⚖️"),
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    [data-testid="stSidebar"] { width: 360px; background: #e8f0fe; }
    [data-testid="stSidebar"] > div { width: 360px; }
    header[data-testid="stHeader"] { display: none; }
    .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

# UI 렌더링
sidebar_data = render_sidebar()

# 위임일자를 session_state에 저장 (연도 버튼에서 사용)
st.session_state["_warrant_date"] = sidebar_data["warrant_date"]

# ── 탭 구조 ──
tab_excel, tab_creditor, tab_gov = st.tabs(["📥 엑셀 → 자동입력", "📋 채권사 서류", "🏛️ 관공서 서류"])

with tab_excel:
    render_excel_delegation_tab()

with tab_creditor:
    main_data = render_main()
    if main_data["generate_clicked"]:
        generate_pdf(sidebar_data, main_data, settings)

with tab_gov:
    gov_data = render_gov_tab()
    if gov_data["generate_clicked"]:
        generate_gov_pdf(sidebar_data, gov_data, settings)
