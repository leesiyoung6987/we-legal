"""
ui_debt_list.py - 채권목록 자동화 탭 UI
신용조회 PDF 업로드 → Claude API 파싱 → 채권목록 엑셀 생성
"""

import streamlit as st
import tempfile
import os
import io
from modules.credit_parser import parse_credit_pdf, classify_and_merge
from modules.debt_list_builder import build_debt_list_workbook
from modules.config_loader import load_settings


def render_debt_list_tab():
    """채권목록 탭 렌더링"""
    
    st.subheader("📊 신용조회 → 채권목록 자동 생성")
    
    # API 키: settings.json에서 자동 로드
    if "anthropic_api_key" not in st.session_state:
        settings = load_settings()
        st.session_state.anthropic_api_key = settings.get("anthropic_api_key", "")
    
    # PDF 업로드
    uploaded_file = st.file_uploader(
        "신용조회 PDF 업로드",
        type=["pdf"],
        help="신용정보조회서 + 채권자변동정보 조회서가 포함된 PDF 파일",
        key="credit_pdf_upload"
    )
    
    if uploaded_file:
        st.info(f"📄 {uploaded_file.name} ({uploaded_file.size / 1024:.0f}KB)")
        
        col1, col2 = st.columns(2)
        with col1:
            parse_btn = st.button("🔍 PDF 분석", type="primary", use_container_width=True)
        
        # 분석 실행
        if parse_btn:
            if not st.session_state.anthropic_api_key:
                st.error("API Key를 입력해주세요.")
                return
            
            with st.spinner("Claude API로 PDF 분석 중... (30초~1분 소요)"):
                try:
                    # 임시 파일 저장
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name
                    
                    # PDF 파싱
                    parsed = parse_credit_pdf(tmp_path, api_key=st.session_state.anthropic_api_key)
                    st.session_state.parsed_credit = parsed
                    
                    # 분류 및 병합
                    classified = classify_and_merge(parsed)
                    st.session_state.classified_credit = classified
                    
                    # 임시 파일 삭제
                    os.unlink(tmp_path)
                    
                    st.success(f"✅ {classified['name']} - 분석 완료!")
                    
                except Exception as e:
                    st.error(f"분석 실패: {e}")
                    import traceback
                    st.code(traceback.format_exc())
                    return
        
        # 결과 표시
        if "classified_credit" in st.session_state:
            classified = st.session_state.classified_credit
            
            st.divider()
            st.subheader(f"📋 {classified['name']} 채권목록")
            
            # 담보대출
            if classified["secured"]:
                st.markdown("**🏠 담보대출**")
                for i, item in enumerate(classified["secured"], 1):
                    st.text(f"  {i}. {item['채권자명']} - {item['내역사유']} ({item['발생일자']})")
            
            # 비담보 (신용조회)
            if classified["unsecured"]:
                st.markdown("**💳 채권자목록 (신용조회)**")
                for i, item in enumerate(classified["unsecured"], 1):
                    st.text(f"  {i}. {item['채권자명']} - {item['내역사유']} ({item['발생일자']})")
            
            # 카드
            if classified["cards"]:
                st.markdown("**🃏 카드**")
                for i, item in enumerate(classified["cards"], 1):
                    st.text(f"  {i}. {item['채권자명']} ({item['발생일자']})")
            
            st.divider()
            
            # 엑셀 다운로드
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📥 채권목록 엑셀 다운로드", type="primary", use_container_width=True):
                    wb = build_debt_list_workbook(classified)
                    buffer = io.BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)
                    
                    st.download_button(
                        label="💾 다운로드",
                        data=buffer.getvalue(),
                        file_name=f"{classified['name']}_채권목록.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            with col2:
                if st.button("➡️ 위임장 자동화로 전달", use_container_width=True):
                    _transfer_to_creditor_tab(classified)
                    st.success("채권사 서류 탭에 채권사 목록이 전달되었습니다!")
                    st.info("'채권사 서류' 탭으로 이동하세요.")


def _transfer_to_creditor_tab(classified):
    """채권목록 데이터를 채권사 서류 탭에 전달
    
    채권사명을 위임장 자동화에서 사용하는 형태로 변환하여 session_state에 저장
    """
    creditor_names = set()
    
    # 담보대출 채권자
    for item in classified.get("secured", []):
        name = item["채권자명"]
        creditor_names.add(name)
    
    # 비담보 채권자
    for item in classified.get("unsecured", []):
        name = item["채권자명"]
        creditor_names.add(name)
    
    # 카드
    for item in classified.get("cards", []):
        name = item["채권자명"]
        creditor_names.add(name)
    
    # session_state에 저장
    st.session_state.auto_creditors = sorted(creditor_names)
    st.session_state.client_name_from_credit = classified.get("name", "")
