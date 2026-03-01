"""
coord_adjuster.py - 좌표 미세조정 도구
PDF 양식 위에 텍스트를 실시간 미리보기하며 좌표를 조절할 수 있습니다.

실행: streamlit run coord_adjuster.py
"""

import streamlit as st
import fitz
import json
import io
import base64
from pathlib import Path

# ── 경로 설정 ──
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
PDF_DIR = TEMPLATES_DIR / "pdf"
COORDS_DIR = TEMPLATES_DIR / "coords"
SETTINGS_PATH = BASE_DIR / "config" / "settings.json"

# ── 설정 로드 ──
def load_settings():
    if SETTINGS_PATH.exists():
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

settings = load_settings()
FONT_PATH = settings.get("font_path", "C:/Windows/Fonts/malgun.ttf")


def get_available_forms():
    """coords 폴더에서 사용 가능한 양식 목록"""
    forms = {}
    for f in COORDS_DIR.glob("*_coords.json"):
        form_name = f.stem.replace("_coords", "")
        # 매칭되는 PDF 찾기
        pdf_path = PDF_DIR / f"{form_name}.pdf"
        if pdf_path.exists():
            forms[form_name] = {"coords": f, "pdf": pdf_path}
    return forms


def load_coords(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def render_preview(pdf_path, fields, sample_data, font_path, page_num=0):
    """PDF 위에 텍스트를 오버레이하여 미리보기 이미지 생성"""
    doc = fitz.open(str(pdf_path))
    page = doc[page_num]
    
    for field in fields:
        field_id = field.get("field_id", "")
        value = sample_data.get(field_id, "")
        if not value:
            continue
        
        field_type = field.get("type", "text")
        if field_type == "image":
            # 이미지는 빨간 박스로 표시
            rect = page.rect
            x = rect.width * (field["x_pct"] / 100)
            y = rect.height * (field["y_pct"] / 100)
            w = rect.width * (field.get("width_pct", 8) / 100)
            h = rect.height * (field.get("height_pct", 4) / 100)
            img_rect = fitz.Rect(x, y - h, x + w, y)
            page.draw_rect(img_rect, color=(1, 0, 0), width=1)
            # 라벨
            page.insert_text(
                fitz.Point(x + 2, y - 2),
                f"[{field_id}]",
                fontsize=7,
                color=(1, 0, 0)
            )
        else:
            font_size = field.get("font_size", 11)
            rect = page.rect
            x = rect.width * (field["x_pct"] / 100)
            y = rect.height * (field["y_pct"] / 100)
            
            try:
                page.insert_text(
                    fitz.Point(x, y),
                    str(value),
                    fontname="malgun",
                    fontfile=font_path,
                    fontsize=font_size,
                    color=(1, 0, 0)  # 빨간색으로 표시
                )
            except:
                page.insert_text(
                    fitz.Point(x, y),
                    str(value),
                    fontsize=font_size,
                    color=(1, 0, 0)
                )
    
    pix = page.get_pixmap(dpi=180)
    img_bytes = pix.tobytes("png")
    doc.close()
    return img_bytes


# ── 기본 샘플 데이터 ──
DEFAULT_SAMPLE = {
    "date_year_suffix": "26",
    "date_month": "02",
    "date_day": "24",
    "date_year": "2026",
    "date": "2026.02.24",
    "client_name": "홍길동",
    "client_phone": "010-1234-5678",
    "client_birth": "901231",
    "client_address": "전주시 덕진구 기린대로 418",
    "client_id_front": "901231",
    "client_id_full": "901231-1234567",
    "agent_name": "이성철",
    "agent_birth": "850101",
    "agent_phone": "010-9876-5432",
    "agent_sign": "[싸인]",
}


# ── 메인 UI ──
st.set_page_config(page_title="좌표 조정 도구", page_icon="🎯", layout="wide")
st.title("🎯 좌표 미세조정 도구")
st.caption("PDF 양식 위에 텍스트를 실시간으로 미리보기하며 좌표(%)를 조절합니다.")

forms = get_available_forms()

if not forms:
    st.error("templates/coords/ 와 templates/pdf/ 에 매칭되는 양식이 없습니다.")
    st.stop()

# ── 사이드바: 양식 선택 + 샘플 데이터 ──
with st.sidebar:
    st.header("📄 양식 선택")
    form_name = st.selectbox("양식", list(forms.keys()), format_func=lambda x: x.replace("_", " "))
    
    form_info = forms[form_name]
    coords_data = load_coords(form_info["coords"])
    
    # 페이지 선택 (멀티페이지 대응)
    pages = coords_data.get("pages", [])
    if len(pages) > 1:
        page_idx = st.selectbox("페이지", range(len(pages)), format_func=lambda i: f"페이지 {pages[i].get('page', i+1)}")
    else:
        page_idx = 0
    
    current_page = pages[page_idx]
    fields = current_page.get("fields", [])
    
    st.divider()
    st.header("📝 샘플 데이터")
    st.caption("미리보기에 표시할 텍스트")
    
    sample_data = {}
    for field in fields:
        fid = field["field_id"]
        default = DEFAULT_SAMPLE.get(fid, f"[{fid}]")
        sample_data[fid] = st.text_input(fid, value=default, key=f"sample_{fid}")

# ── 메인 영역: 좌표 조정 + 미리보기 ──
col_controls, col_preview = st.columns([1, 2])

with col_controls:
    st.subheader("⚙️ 좌표 조정")
    st.caption("0.1 단위로 조절 가능. 수정 후 미리보기에 즉시 반영됩니다.")
    
    updated_fields = []
    
    for i, field in enumerate(fields):
        fid = field["field_id"]
        ftype = field.get("type", "text")
        
        with st.expander(f"{'🖼️' if ftype == 'image' else '📍'} {fid}", expanded=True):
            c1, c2 = st.columns(2)
            new_x = c1.number_input("X %", value=field["x_pct"], step=0.1, format="%.1f", key=f"x_{i}")
            new_y = c2.number_input("Y %", value=field["y_pct"], step=0.1, format="%.1f", key=f"y_{i}")
            
            new_field = {**field, "x_pct": new_x, "y_pct": new_y}
            
            if ftype != "image":
                new_fs = st.number_input("글자 크기", value=field.get("font_size", 11), step=1, key=f"fs_{i}")
                new_field["font_size"] = new_fs
            else:
                c3, c4 = st.columns(2)
                new_w = c3.number_input("너비 %", value=field.get("width_pct", 8), step=0.1, format="%.1f", key=f"w_{i}")
                new_h = c4.number_input("높이 %", value=field.get("height_pct", 4), step=0.1, format="%.1f", key=f"h_{i}")
                new_field["width_pct"] = new_w
                new_field["height_pct"] = new_h
            
            updated_fields.append(new_field)

with col_preview:
    st.subheader("👁️ 미리보기")
    
    # 미리보기 렌더링
    page_num = current_page.get("page", 1) - 1
    try:
        img_bytes = render_preview(
            form_info["pdf"], updated_fields, sample_data, FONT_PATH, page_num
        )
        st.image(img_bytes, use_container_width=True)
    except Exception as e:
        st.error(f"미리보기 오류: {e}")
        # 폰트 없을 경우 기본 폰트로 재시도
        try:
            img_bytes = render_preview(
                form_info["pdf"], updated_fields, sample_data, None, page_num
            )
            st.image(img_bytes, use_container_width=True)
            st.warning("⚠️ 맑은고딕 폰트를 찾을 수 없어 기본 폰트로 표시됩니다. 위치는 동일합니다.")
        except Exception as e2:
            st.error(f"렌더링 실패: {e2}")

# ── 하단: 저장 ──
st.divider()
col_save, col_json = st.columns([1, 2])

with col_save:
    if st.button("💾 좌표 저장", type="primary", use_container_width=True):
        # 업데이트된 좌표로 JSON 저장
        new_coords = {**coords_data}
        new_coords["pages"][page_idx]["fields"] = updated_fields
        
        save_path = form_info["coords"]
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(new_coords, f, indent=2, ensure_ascii=False)
        
        st.success(f"✅ 저장 완료: {save_path.name}")
    
    st.caption("저장하면 coords JSON 파일이 직접 덮어쓰기됩니다.")

with col_json:
    with st.expander("📋 현재 JSON 보기"):
        output = {**coords_data}
        output["pages"][page_idx]["fields"] = updated_fields
        st.code(json.dumps(output, indent=2, ensure_ascii=False), language="json")
