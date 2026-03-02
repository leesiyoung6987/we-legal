"""
ui_gov.py - 관공서 서류 탭 UI
번호 행 방식 + 드롭다운 + 고유필드 인라인 + 양면인쇄
"""

import streamlit as st
from modules.config_loader import _load_json


def load_gov_forms():
    """gov_forms.json 로드"""
    data = _load_json("gov_forms.json")
    data.pop("_comment", None)
    return data


def render_gov_tab():
    """관공서 서류 탭 렌더링
    
    Returns:
        dict: {"forms": [...], "generate_clicked": bool}
    """
    gov_forms = load_gov_forms()
    form_names = [""] + list(gov_forms.keys())  # 빈 항목 = 미선택
    max_rows = 15
    default_rows = 5

    if "gov_row_count" not in st.session_state:
        st.session_state.gov_row_count = default_rows

    # ── 헤더 + 버튼 ──
    col_title, col_btn = st.columns([2, 3])
    with col_title:
        count = st.session_state.gov_row_count
        st.markdown(
            f"### 관공서 서류 "
            f"<small style='color:#6b7280;font-weight:400;font-size:14px;'>{count}칸</small>",
            unsafe_allow_html=True,
        )
    with col_btn:
        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button("＋ 서류 추가", key="gov_add", use_container_width=True):
                if st.session_state.gov_row_count < max_rows:
                    st.session_state.gov_row_count += 1
                    st.rerun()
        with b2:
            generate_clicked = st.button("⚡ PDF 생성", type="primary", key="gov_gen_pdf", use_container_width=True)
        with b3:
            if "gov_pdf_bytes" in st.session_state and st.session_state.gov_pdf_bytes:
                st.download_button(
                    label="📥 관공서", data=st.session_state.gov_pdf_bytes,
                    file_name=st.session_state.get("gov_pdf_filename", "관공서서류.pdf"),
                    mime="application/pdf", use_container_width=True,
                )
            else:
                st.button("📥 관공서", disabled=True, key="gov_dl_disabled", use_container_width=True)
        # 무상거주사실확인서 별도 다운로드
        if "musang_pdf_bytes" in st.session_state and st.session_state.musang_pdf_bytes:
            st.download_button(
                label="📄 무상거주사실확인서 다운", data=st.session_state.musang_pdf_bytes,
                file_name=st.session_state.get("musang_pdf_filename", "무상거주사실확인서.pdf"),
                mime="application/pdf", use_container_width=True,
            )

    st.divider()
    gov_progress_placeholder = st.empty()  # 관공서 PDF 생성 진행바 (상단)
    hc = st.columns([0.3, 2.5])
    hc[0].markdown("<small style='color:#6b7280;font-weight:600;'>#</small>", unsafe_allow_html=True)
    hc[1].markdown("<small style='color:#6b7280;font-weight:600;'>서류 선택</small>", unsafe_allow_html=True)

    # ── 서류 행 ──
    form_data_list = []
    for i in range(st.session_state.gov_row_count):
        result = _render_gov_row(i, form_names, gov_forms)
        if result:
            form_data_list.append(result)

    st.divider()

    # ── 하단 정보 ──
    selected_count = len(form_data_list)
    duplex_count = sum(1 for f in form_data_list if f.get("duplex"))
    
    col_info, col_add = st.columns([2, 1])
    with col_info:
        info_text = f"선택된 서류: **{selected_count}**개"
        if duplex_count:
            info_text += f" | 양면인쇄: **{duplex_count}**건"
        st.markdown(info_text)
    with col_add:
        if st.button("＋ 서류 추가", key="gov_add_bottom", use_container_width=True):
            if st.session_state.gov_row_count < max_rows:
                st.session_state.gov_row_count += 1
                st.rerun()

    return {"forms": form_data_list, "generate_clicked": generate_clicked, "progress_placeholder": gov_progress_placeholder}


def _render_gov_row(idx, form_names, gov_forms):
    """관공서 서류 행 1개 렌더링"""
    
    # 줄무늬 배경
    if idx % 2 == 1:
        st.markdown(
            '<div class="row-even" style="background:#f8f9fc;margin:-8px -16px;padding:8px 16px;border-radius:4px;"></div>',
            unsafe_allow_html=True,
        )

    cols = st.columns([0.3, 2.5])

    with cols[0]:
        st.markdown(
            f"<div style='padding-top:8px;color:#6b7280;font-weight:600;'>{idx+1}</div>",
            unsafe_allow_html=True,
        )

    with cols[1]:
        selected = st.selectbox(
            "서류", form_names, key=f"gov_sel_{idx}",
            label_visibility="collapsed",
            index=0,
        )

        if not selected:
            return None

        form_config = gov_forms.get(selected, {})
        unique_fields = form_config.get("unique_fields", [])
        is_duplex = form_config.get("duplex", False)

        # 양면 표시
        if is_duplex:
            st.caption("📄 양면인쇄")

        # ── 고유필드 입력 (인라인) ──
        extra_values = {}
        if unique_fields:
            # 3열로 배치
            field_cols = st.columns(min(len(unique_fields), 3))
            for fi, field in enumerate(unique_fields):
                with field_cols[fi % min(len(unique_fields), 3)]:
                    if field.get("type") == "select":
                        options = field.get("options", [])
                        val = st.selectbox(
                            field["label"],
                            options,
                            key=f"gov_{idx}_{field['key']}",
                        )
                    elif field.get("type") == "checkbox":
                        checked = st.checkbox(
                            field["label"],
                            key=f"gov_{idx}_{field['key']}",
                        )
                        # 체크하면 label 텍스트가 PDF에 삽입됨
                        val = field["label"] if checked else ""
                    else:
                        val = st.text_input(
                            field["label"],
                            key=f"gov_{idx}_{field['key']}",
                            placeholder=field["label"],
                        )
                    extra_values[field["key"]] = val

    return {
        "form_name": selected,
        "template": form_config.get("template", ""),
        "coords": form_config.get("coords", ""),
        "extra": extra_values,
        "duplex": is_duplex,
    }
