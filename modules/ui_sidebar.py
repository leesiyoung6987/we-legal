"""
ui_sidebar.py - 사이드바 UI 컴포넌트
위임인/수임인 정보 입력, 신분증 업로드, 위임일자.
"""

import streamlit as st
from datetime import date
from modules.config_loader import load_staff, get_id_card_path


def render_sidebar():
    """사이드바 전체 렌더링. 입력된 데이터를 dict로 반환.
    
    Returns:
        dict: {
            "client": {"name", "birth", "id_front", "id_back", "address", "phone"},
            "agent_name": str,
            "agent_data": dict or None,
            "client_id_file": UploadedFile or None,
            "agent_id_path": Path or None,
            "warrant_date": date
        }
    """
    staff_data = load_staff()
    
    with st.sidebar:
        st.markdown("## ⚖️ 리셋 위임장 자동화")
        st.caption("회생파산 서류 자동화")
        
        # ── 위임인 ──
        client = _render_client_section()
        
        # ── 위임인 신분증 (위임인 정보 바로 아래) ──
        client_id_file = st.file_uploader(
            "위임인 신분증",
            type=["pdf", "png", "jpg", "jpeg"],
            key="client_id_upload"
        )
        
        # ── 인감도장 / 인감증명서 (선택) ──
        with st.expander("🔏 인감도장 / 인감증명서 (팩스용)", expanded=False):
            stamp_file = st.file_uploader(
                "인감도장 이미지",
                type=["pdf", "png", "jpg", "jpeg"],
                key="stamp_upload",
                help="팩스 발급 채권사에 자동으로 도장이 찍힙니다"
            )
            seal_cert_file = st.file_uploader(
                "인감증명서",
                type=["pdf", "png", "jpg", "jpeg"],
                key="seal_cert_upload",
                help="팩스 발급 채권사에 자동으로 첨부됩니다"
            )
        
        st.divider()
        
        # ── 수임인 ──
        agent_name, agent_data = _render_agent_section(staff_data)
        
        # ── 수임인 신분증 (자동 로드) ──
        agent_id_path = _render_agent_id(agent_name, staff_data)
        
        st.divider()
        
        # ── 위임일자 ──
        warrant_date = _render_date_section()
    
    return {
        "client": client,
        "agent_name": agent_name,
        "agent_data": agent_data,
        "client_id_file": client_id_file,
        "agent_id_path": agent_id_path,
        "warrant_date": warrant_date,
        "stamp_file": stamp_file,
        "seal_cert_file": seal_cert_file,
    }


def _render_client_section():
    """위임인 정보 입력 섹션"""
    st.markdown(
        '<p style="font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7280;'
        'padding-left:8px;border-left:3px solid #4a7dff;">위임인 정보</p>',
        unsafe_allow_html=True
    )
    
    name = st.text_input("성명", placeholder="홍길동", key="client_name")
    
    col1, col2 = st.columns(2)
    with col1:
        id_front = st.text_input("주민번호 앞자리", placeholder="940317", key="client_id_front")
    with col2:
        id_back = st.text_input("주민번호 뒷자리", type="password", placeholder="1234567", key="client_id_back")
    
    address = st.text_input("주소", placeholder="전주시 덕진구 ...", key="client_address")
    phone = st.text_input("전화번호", placeholder="010-1234-5678", key="client_phone")
    
    return {
        "name": name,
        "birth": id_front,
        "id_front": id_front,
        "id_back": id_back,
        "address": address,
        "phone": phone,
    }


def _render_agent_section(staff_data):
    """수임인 선택 및 정보 표시 섹션"""
    st.markdown(
        '<p style="font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7280;'
        'padding-left:8px;border-left:3px solid #4a7dff;">수임인 정보</p>',
        unsafe_allow_html=True
    )
    
    options = ["— 선택 —"] + list(staff_data.keys())
    agent_name = st.selectbox("수임인 선택", options, key="agent_select")
    
    agent_data = None
    if agent_name != "— 선택 —" and agent_name in staff_data:
        agent_data = staff_data[agent_name]
        st.markdown(f"""
        <div style="background:#f0f1f4;padding:12px 14px;border-radius:8px;
                    font-size:13px;line-height:1.8;margin:8px 0;">
            <b>생년월일:</b> {agent_data["birth"]}<br>
            <b>주소:</b> {agent_data["address"]}<br>
            <b>전화:</b> {agent_data["phone"]}<br>
            <b>FAX:</b> {agent_data["fax"]}
        </div>
        """, unsafe_allow_html=True)
    
    return agent_name, agent_data


def _render_agent_id(agent_name, staff_data):
    """수임인 신분증 자동 로드 표시"""
    agent_id_path = None
    if agent_name != "— 선택 —" and agent_name in staff_data:
        id_filename = staff_data[agent_name].get("id_card", "")
        path = get_id_card_path(id_filename)
        
        if path.exists():
            agent_id_path = path
            st.markdown(
                f'<div style="border:1px solid #10b981;border-radius:8px;padding:10px 14px;'
                f'background:rgba(16,185,129,0.06);font-size:13px;color:#10b981;font-weight:600;">'
                f'✅ 수임인 신분증 자동 로드됨<br>📎 {id_filename}</div>',
                unsafe_allow_html=True
            )
        else:
            st.warning(f"⚠️ {id_filename} 파일이 id_cards/agents/ 에 없습니다.")
    
    return agent_id_path


def _render_date_section():
    """위임일자 선택 섹션"""
    st.markdown(
        '<p style="font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7280;'
        'padding-left:8px;border-left:3px solid #4a7dff;">위임일자</p>',
        unsafe_allow_html=True
    )
    return st.date_input("위임일자", value=date.today(), key="warrant_date")
