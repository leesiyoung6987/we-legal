"""
pdf_engine.py - PDF 생성 엔진
위임장 채우기, 신분증 변환, PDF 합치기 등 모든 PDF 처리.
새 양식 추가 시 generate 함수를 추가하면 됨.
"""

import fitz  # PyMuPDF
from modules.config_loader import load_coords, load_settings

# 폰트 경로는 settings에서 로드
_settings = load_settings()
FONT_PATH = _settings.get("font_path", "C:/Windows/Fonts/malgun.ttf")


def build_manual_cover(creditor_names, client_name="", warrant_date="", ins_homepage=None, ins_customer=None, sp_ins_homepage=None, sp_ins_customer=None):
    """채권사별 발급 매뉴얼 커버 페이지 PDF 생성 (표 형태, 발급방법별 그룹)
    
    Args:
        creditor_names: 채권사명 리스트
        client_name: 위임인 이름
        warrant_date: 위임일자
        ins_homepage: 보험 홈페이지 발급 건 리스트 (optional)
        ins_customer: 보험 고객요청 건 리스트 (optional)
        sp_ins_homepage: 배우자 보험 홈페이지 발급 건 (optional)
        sp_ins_customer: 배우자 보험 고객요청 건 (optional)
    
    Returns:
        fitz.Document (매뉴얼 커버 페이지들)
    """
    from modules.config_loader import get_issue_info
    
    doc = fitz.open()
    font = fitz.Font(fontfile=FONT_PATH)
    
    # A4 가로 (Landscape)
    W, H = 842, 595
    MARGIN_L = 30
    MARGIN_T = 30
    MARGIN_R = W - 30
    MARGIN_B = H - 25
    
    # 색상
    C_TITLE = (0.15, 0.25, 0.45)
    C_GRAY = (0.45, 0.45, 0.45)
    C_TEXT = (0.1, 0.1, 0.1)
    C_WHITE = (1, 1, 1)
    C_LINE = (0.7, 0.7, 0.7)
    C_HEADER_BG = (0.2, 0.33, 0.55)
    C_ROW_EVEN = (0.95, 0.96, 0.98)
    C_ROW_ODD = (1, 1, 1)
    C_SPECIAL = (0.7, 0.1, 0.1)
    
    # 발급방법별 그룹 색상
    GROUP_COLORS = {
        "팩스": (0.85, 0.15, 0.15),
        "등기": (0.15, 0.45, 0.15),
        "방문발급": (0.15, 0.3, 0.65),
        "현장발급": (0.15, 0.3, 0.65),
        "고객요청": (0.6, 0.4, 0.1),
        "기타": (0.4, 0.4, 0.4),
    }
    
    # 컬럼 정의: [이름, x시작, 너비]
    COL_NUM = (MARGIN_L, 25)
    COL_NAME = (MARGIN_L + 25, 100)
    COL_METHOD = (MARGIN_L + 125, 50)
    COL_TEL = (MARGIN_L + 175, 160)
    COL_FAX = (MARGIN_L + 335, 105)
    COL_ADDR = (MARGIN_L + 440, 160)
    COL_NOTE = (MARGIN_L + 600, MARGIN_R - MARGIN_L - 600)
    
    COLS = [COL_NUM, COL_NAME, COL_METHOD, COL_TEL, COL_FAX, COL_ADDR, COL_NOTE]
    COL_HEADERS = ["#", "업체명", "발급방법", "고객센터", "팩스번호", "주소", "특이사항"]
    
    def _text_height(text, font_size, col_width):
        """텍스트가 차지할 높이 계산"""
        if not text:
            return font_size + 4
        lines = str(text).split('\n')
        total_lines = 0
        for line in lines:
            line_w = font.text_length(line, fontsize=font_size)
            num_lines = max(1, int(line_w / (col_width - 6)) + 1)
            total_lines += num_lines
        return max(total_lines * (font_size + 3), font_size + 4)
    
    def _draw_cell_text(page, x, y, w, h, text, font_size=7, color=C_TEXT, padding=3):
        """셀 안에 줄바꿈 텍스트 그리기"""
        if not text:
            return
        text = str(text).strip()
        lines = text.split('\n')
        cy = y + font_size + padding
        max_w = w - padding * 2
        
        for line in lines:
            if cy + font_size > y + h:
                break
            # 긴 줄 자동 줄바꿈
            if font.text_length(line, fontsize=font_size) > max_w:
                current = ""
                for char in line:
                    test = current + char
                    if font.text_length(test, fontsize=font_size) > max_w:
                        if current.strip():
                            page.insert_text(
                                fitz.Point(x + padding, cy), current,
                                fontname="malgun", fontfile=FONT_PATH,
                                fontsize=font_size, color=color
                            )
                            cy += font_size + 3
                        current = char
                    else:
                        current = test
                if current.strip() and cy + font_size <= y + h:
                    page.insert_text(
                        fitz.Point(x + padding, cy), current,
                        fontname="malgun", fontfile=FONT_PATH,
                        fontsize=font_size, color=color
                    )
                    cy += font_size + 3
            else:
                page.insert_text(
                    fitz.Point(x + padding, cy), line,
                    fontname="malgun", fontfile=FONT_PATH,
                    fontsize=font_size, color=color
                )
                cy += font_size + 3
    
    def _draw_header(page, y):
        """테이블 헤더 그리기"""
        h = 18
        # 헤더 배경
        page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + h), color=None, fill=C_HEADER_BG)
        # 헤더 텍스트
        for i, (cx, cw) in enumerate(COLS):
            page.insert_text(
                fitz.Point(cx + 3, y + 13), COL_HEADERS[i],
                fontname="malgun", fontfile=FONT_PATH,
                fontsize=8, color=C_WHITE
            )
        # 세로 구분선
        for cx, cw in COLS[1:]:
            page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + h), color=(0.3, 0.4, 0.6), width=0.5)
        return y + h
    
    def _draw_group_header(page, y, method, count):
        """발급방법 그룹 헤더"""
        h = 16
        color = GROUP_COLORS.get(method, (0.4, 0.4, 0.4))
        # 배경
        bg = (color[0] * 0.15 + 0.85, color[1] * 0.15 + 0.85, color[2] * 0.15 + 0.85)
        page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + h), color=None, fill=bg)
        # 텍스트
        label = f"■ {method} ({count}개)"
        page.insert_text(
            fitz.Point(MARGIN_L + 5, y + 12), label,
            fontname="malgun", fontfile=FONT_PATH,
            fontsize=9, color=color
        )
        return y + h
    
    # 매뉴얼 정보 수집 & 발급방법별 그룹핑
    groups = {}
    method_order = ["팩스", "등기", "방문발급", "현장발급", "기타", "고객요청", ""]
    
    for name in creditor_names:
        info = get_issue_info(name)
        if not info:
            info = {"발급방법": "기타"}
        method = info.get("발급방법", "기타").strip()
        if method not in groups:
            groups[method] = []
        groups[method].append((name, info))
    
    # 정렬
    sorted_groups = []
    for m in method_order:
        if m in groups:
            sorted_groups.append((m, groups.pop(m)))
    for m, items in groups.items():
        sorted_groups.append((m, items))
    
    # 페이지 생성
    page = doc.new_page(width=W, height=H)
    y = MARGIN_T
    
    # 타이틀
    page.insert_text(
        fitz.Point(MARGIN_L, y + 16), f"발급 매뉴얼",
        fontname="malgun", fontfile=FONT_PATH, fontsize=16, color=C_TITLE
    )
    page.insert_text(
        fitz.Point(MARGIN_L + 130, y + 16),
        f"{client_name}  |  {warrant_date}  |  총 {len(creditor_names)}개 채권사",
        fontname="malgun", fontfile=FONT_PATH, fontsize=9, color=C_GRAY
    )
    y += 25
    
    # 헤더
    y = _draw_header(page, y)
    
    row_num = 0
    
    for method, items in sorted_groups:
        # 그룹 헤더 높이 체크
        if y + 16 > MARGIN_B:
            page = doc.new_page(width=W, height=H)
            y = MARGIN_T
            y = _draw_header(page, y)
        
        # 그룹 헤더
        y = _draw_group_header(page, y, method or "미지정", len(items))
        
        for name, info in items:
            # 행 높이 계산
            tel = info.get("고객센터", "")
            fax_num = info.get("팩스번호", "")
            addr = info.get("주소", "")
            note = info.get("특이사항", "")
            
            heights = [
                _text_height(name, 7, COL_NAME[1]),
                _text_height(tel, 7, COL_TEL[1]),
                _text_height(fax_num, 7, COL_FAX[1]),
                _text_height(addr, 7, COL_ADDR[1]),
                _text_height(note, 7, COL_NOTE[1]),
            ]
            row_h = max(max(heights), 16) + 4
            
            # 페이지 넘김 체크
            if y + row_h > MARGIN_B:
                page = doc.new_page(width=W, height=H)
                y = MARGIN_T
                y = _draw_header(page, y)
            
            # 행 배경
            bg = C_ROW_EVEN if row_num % 2 == 0 else C_ROW_ODD
            page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + row_h), color=None, fill=bg)
            
            # 세로 구분선
            for cx, cw in COLS[1:]:
                page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + row_h), color=C_LINE, width=0.3)
            
            # 하단 구분선
            page.draw_line(fitz.Point(MARGIN_L, y + row_h), fitz.Point(MARGIN_R, y + row_h), color=C_LINE, width=0.3)
            
            # 셀 내용
            row_num += 1
            _draw_cell_text(page, COL_NUM[0], y, COL_NUM[1], row_h, str(row_num), 7, C_GRAY)
            _draw_cell_text(page, COL_NAME[0], y, COL_NAME[1], row_h, name, 7, C_TEXT)
            
            m_color = GROUP_COLORS.get(method, C_GRAY)
            _draw_cell_text(page, COL_METHOD[0], y, COL_METHOD[1], row_h, method, 7, m_color)
            _draw_cell_text(page, COL_TEL[0], y, COL_TEL[1], row_h, tel, 7, C_TEXT)
            _draw_cell_text(page, COL_FAX[0], y, COL_FAX[1], row_h, fax_num, 7, C_TEXT)
            _draw_cell_text(page, COL_ADDR[0], y, COL_ADDR[1], row_h, addr, 7, C_TEXT)
            _draw_cell_text(page, COL_NOTE[0], y, COL_NOTE[1], row_h, note, 7, C_SPECIAL)
            
            y += row_h
    
    # ── 보험 고객요청 섹션 ──
    if ins_customer:
        y += 10
        if y + 40 > MARGIN_B:
            page = doc.new_page(width=W, height=H)
            y = MARGIN_T

        h = 16
        bg_cr = (0.98, 0.94, 0.88)
        color_cr = (0.6, 0.4, 0.1)
        page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + h), color=None, fill=bg_cr)
        page.insert_text(
            fitz.Point(MARGIN_L + 5, y + 12),
            f"■ 보험 고객요청 ({len(ins_customer)}개사)",
            fontname="malgun", fontfile=FONT_PATH,
            fontsize=9, color=color_cr
        )
        y += h

        for ins_item in ins_customer:
            comp = ins_item["name"]
            tel = ins_item.get("tel", "")
            doc_type = ins_item.get("doc_type", "")
            nos = ", ".join(ins_item.get("policy_nos", []))

            note_text = ""

            heights = [
                _text_height(comp, 7, COL_NAME[1]),
                _text_height(tel, 7, COL_TEL[1]),
                _text_height(note_text, 7, COL_NOTE[1]),
            ]
            row_h = max(max(heights), 16) + 4

            if y + row_h > MARGIN_B:
                page = doc.new_page(width=W, height=H)
                y = MARGIN_T
                y = _draw_header(page, y)

            bg = C_ROW_EVEN if row_num % 2 == 0 else C_ROW_ODD
            page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + row_h), color=None, fill=bg)

            for cx, cw in COLS[1:]:
                page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + row_h), color=C_LINE, width=0.3)
            page.draw_line(fitz.Point(MARGIN_L, y + row_h), fitz.Point(MARGIN_R, y + row_h), color=C_LINE, width=0.3)

            row_num += 1
            _draw_cell_text(page, COL_NUM[0], y, COL_NUM[1], row_h, str(row_num), 7, C_GRAY)
            _draw_cell_text(page, COL_NAME[0], y, COL_NAME[1], row_h, comp, 7, C_TEXT)
            _draw_cell_text(page, COL_METHOD[0], y, COL_METHOD[1], row_h, "고객요청", 7, color_cr)
            _draw_cell_text(page, COL_TEL[0], y, COL_TEL[1], row_h, tel, 7, C_TEXT)
            _draw_cell_text(page, COL_NOTE[0], y, COL_NOTE[1], row_h, note_text, 7, C_TEXT)

            y += row_h

    # ── 보험 홈페이지 발급 섹션 ──
    if ins_homepage:
        # 섹션 간격
        y += 10
        if y + 40 > MARGIN_B:
            page = doc.new_page(width=W, height=H)
            y = MARGIN_T

        # 보험 섹션 헤더
        h = 16
        bg_ins = (0.92, 0.90, 0.98)
        color_ins = (0.4, 0.2, 0.7)
        page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + h), color=None, fill=bg_ins)
        page.insert_text(
            fitz.Point(MARGIN_L + 5, y + 12),
            f"■ 보험 홈페이지 발급 ({len(ins_homepage)}개사)",
            fontname="malgun", fontfile=FONT_PATH,
            fontsize=9, color=color_ins
        )
        y += h

        for ins_item in ins_homepage:
            comp = ins_item["name"]
            cnt = ins_item["count"]
            tel = ins_item.get("tel", "")
            route = ins_item.get("route", "")
            nos = ", ".join(ins_item.get("policy_nos", []))

            note_text = ""
            if route:
                note_text = f"경로: {route}"

            heights = [
                _text_height(comp, 7, COL_NAME[1]),
                _text_height(tel, 7, COL_TEL[1]),
                _text_height(note_text, 7, COL_NOTE[1]),
            ]
            row_h = max(max(heights), 16) + 4

            if y + row_h > MARGIN_B:
                page = doc.new_page(width=W, height=H)
                y = MARGIN_T
                y = _draw_header(page, y)

            bg = C_ROW_EVEN if row_num % 2 == 0 else C_ROW_ODD
            page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + row_h), color=None, fill=bg)

            for cx, cw in COLS[1:]:
                page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + row_h), color=C_LINE, width=0.3)
            page.draw_line(fitz.Point(MARGIN_L, y + row_h), fitz.Point(MARGIN_R, y + row_h), color=C_LINE, width=0.3)

            row_num += 1
            _draw_cell_text(page, COL_NUM[0], y, COL_NUM[1], row_h, str(row_num), 7, C_GRAY)
            _draw_cell_text(page, COL_NAME[0], y, COL_NAME[1], row_h, comp, 7, C_TEXT)
            _draw_cell_text(page, COL_METHOD[0], y, COL_METHOD[1], row_h, "홈페이지", 7, color_ins)
            _draw_cell_text(page, COL_TEL[0], y, COL_TEL[1], row_h, tel, 7, C_TEXT)
            _draw_cell_text(page, COL_NOTE[0], y, COL_NOTE[1], row_h, note_text, 7, C_TEXT)

            y += row_h

    # ── 배우자 보험 고객요청 섹션 ──
    if sp_ins_customer:
        y += 10
        if y + 40 > MARGIN_B:
            page = doc.new_page(width=W, height=H)
            y = MARGIN_T

        h = 16
        bg_sp = (0.95, 0.90, 0.85)
        color_sp = (0.5, 0.3, 0.1)
        page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + h), color=None, fill=bg_sp)
        page.insert_text(
            fitz.Point(MARGIN_L + 5, y + 12),
            f"■ 배우자 보험 고객요청 ({len(sp_ins_customer)}개사)",
            fontname="malgun", fontfile=FONT_PATH,
            fontsize=9, color=color_sp
        )
        y += h

        for ins_item in sp_ins_customer:
            comp = ins_item["name"]
            tel = ins_item.get("tel", "")
            note_text = ""

            heights = [
                _text_height(comp, 7, COL_NAME[1]),
                _text_height(tel, 7, COL_TEL[1]),
            ]
            row_h = max(max(heights), 16) + 4

            if y + row_h > MARGIN_B:
                page = doc.new_page(width=W, height=H)
                y = MARGIN_T
                y = _draw_header(page, y)

            bg = C_ROW_EVEN if row_num % 2 == 0 else C_ROW_ODD
            page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + row_h), color=None, fill=bg)

            for cx, cw in COLS[1:]:
                page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + row_h), color=C_LINE, width=0.3)
            page.draw_line(fitz.Point(MARGIN_L, y + row_h), fitz.Point(MARGIN_R, y + row_h), color=C_LINE, width=0.3)

            row_num += 1
            _draw_cell_text(page, COL_NUM[0], y, COL_NUM[1], row_h, str(row_num), 7, C_GRAY)
            _draw_cell_text(page, COL_NAME[0], y, COL_NAME[1], row_h, comp, 7, C_TEXT)
            _draw_cell_text(page, COL_METHOD[0], y, COL_METHOD[1], row_h, "고객요청", 7, color_sp)
            _draw_cell_text(page, COL_TEL[0], y, COL_TEL[1], row_h, tel, 7, C_TEXT)
            _draw_cell_text(page, COL_NOTE[0], y, COL_NOTE[1], row_h, note_text, 7, C_TEXT)

            y += row_h

    # ── 배우자 보험 홈페이지 발급 섹션 ──
    if sp_ins_homepage:
        y += 10
        if y + 40 > MARGIN_B:
            page = doc.new_page(width=W, height=H)
            y = MARGIN_T

        h = 16
        bg_sp2 = (0.90, 0.88, 0.95)
        color_sp2 = (0.35, 0.2, 0.6)
        page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + h), color=None, fill=bg_sp2)
        page.insert_text(
            fitz.Point(MARGIN_L + 5, y + 12),
            f"■ 배우자 보험 홈페이지 발급 ({len(sp_ins_homepage)}개사)",
            fontname="malgun", fontfile=FONT_PATH,
            fontsize=9, color=color_sp2
        )
        y += h

        for ins_item in sp_ins_homepage:
            comp = ins_item["name"]
            tel = ins_item.get("tel", "")
            route = ins_item.get("route", "")

            note_text = ""
            if route:
                note_text = f"경로: {route}"

            heights = [
                _text_height(comp, 7, COL_NAME[1]),
                _text_height(tel, 7, COL_TEL[1]),
                _text_height(note_text, 7, COL_NOTE[1]),
            ]
            row_h = max(max(heights), 16) + 4

            if y + row_h > MARGIN_B:
                page = doc.new_page(width=W, height=H)
                y = MARGIN_T
                y = _draw_header(page, y)

            bg = C_ROW_EVEN if row_num % 2 == 0 else C_ROW_ODD
            page.draw_rect(fitz.Rect(MARGIN_L, y, MARGIN_R, y + row_h), color=None, fill=bg)

            for cx, cw in COLS[1:]:
                page.draw_line(fitz.Point(cx, y), fitz.Point(cx, y + row_h), color=C_LINE, width=0.3)
            page.draw_line(fitz.Point(MARGIN_L, y + row_h), fitz.Point(MARGIN_R, y + row_h), color=C_LINE, width=0.3)

            row_num += 1
            _draw_cell_text(page, COL_NUM[0], y, COL_NUM[1], row_h, str(row_num), 7, C_GRAY)
            _draw_cell_text(page, COL_NAME[0], y, COL_NAME[1], row_h, comp, 7, C_TEXT)
            _draw_cell_text(page, COL_METHOD[0], y, COL_METHOD[1], row_h, "홈페이지", 7, color_sp2)
            _draw_cell_text(page, COL_TEL[0], y, COL_TEL[1], row_h, tel, 7, C_TEXT)
            _draw_cell_text(page, COL_NOTE[0], y, COL_NOTE[1], row_h, note_text, 7, C_TEXT)

            y += row_h

    return doc


def insert_text(page, x_pct, y_pct, text, font_size=11):
    """PDF 페이지에 퍼센트 좌표로 텍스트 삽입
    
    Args:
        page: fitz.Page 객체
        x_pct: X 좌표 (0~100, 페이지 너비 기준 %)
        y_pct: Y 좌표 (0~100, 페이지 높이 기준 %)
        text: 삽입할 텍스트
        font_size: 글자 크기 (기본 11)
    """
    rect = page.rect
    x = rect.width * (x_pct / 100)
    y = rect.height * (y_pct / 100)
    page.insert_text(
        fitz.Point(x, y),
        str(text),
        fontname="malgun",
        fontfile=FONT_PATH,
        fontsize=font_size,
        color=(0, 0, 0)
    )


def insert_text_spaced(page, x_pct, y_pct, text, font_size=11, spacing=0):
    """자간 조정 텍스트 삽입
    
    Args:
        spacing: 글자 간 추가 간격 (pt 단위, 0=기본)
    """
    if spacing <= 0:
        insert_text(page, x_pct, y_pct, text, font_size)
        return
    
    rect = page.rect
    x = rect.width * (x_pct / 100)
    y = rect.height * (y_pct / 100)
    font = fitz.Font(fontfile=FONT_PATH)
    
    for char in str(text):
        page.insert_text(
            fitz.Point(x, y), char,
            fontname="malgun", fontfile=FONT_PATH,
            fontsize=font_size, color=(0, 0, 0)
        )
        char_width = font.text_length(char, fontsize=font_size)
        x += char_width + spacing


def generate_name_stamp(name, size=400):
    """이름으로 세로 타원형 인감도장 이미지 자동 생성
    
    Args:
        name: 이름 (2~4글자)
        size: 최종 이미지 높이 (px)
    
    Returns:
        임시 PNG 파일 경로 (str)
    """
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    from PIL.ImageFilter import MaxFilter
    import numpy as np
    import tempfile
    import random
    
    hi = size * 3
    w_ratio = 0.6
    img_w = int(hi * w_ratio)
    img_h = hi
    img = Image.new("RGBA", (img_w, img_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    
    # 실제 도장 색상
    stamp_color = (239, 63, 118, 255)
    cx, cy = img_w // 2, img_h // 2
    rx = int(img_w * 0.44)
    ry = int(img_h * 0.44)
    
    # 타원형 테두리 (두껍게)
    border_width = max(10, int(hi * 0.025))
    draw.ellipse(
        [cx - rx, cy - ry, cx + rx, cy + ry],
        outline=stamp_color, width=border_width
    )
    
    # 프로젝트 내장 붓글씨 폰트 → 세리프 폰트 → 기본 폰트
    from pathlib import Path as _P
    _base = _P(__file__).parent.parent / "fonts"
    _brush = _base / "GEULSEEDANGGoyoTTF.ttf"
    _serif = _base / "NotoSerifCJK-Black.ttc"
    if _brush.exists():
        font_path = str(_brush)
    elif _serif.exists():
        font_path = str(_serif)
    else:
        font_path = FONT_PATH
    
    # 글자가 타원 안에 꽉 차게
    n = len(name)
    usable_h = ry * 2 * 0.92
    fs = int(usable_h / n)
    fs = min(fs, int(rx * 2.0))
    font = ImageFont.truetype(font_path, fs)
    
    # 실제 글자 높이 측정해서 정확히 중앙 배치
    char_heights = []
    for char in name:
        bbox = draw.textbbox((0, 0), char, font=font)
        char_heights.append(bbox[3] - bbox[1])
    total_text_h = sum(char_heights)
    start_y = cy - total_text_h // 2
    cur_y = start_y
    
    for i, char in enumerate(name):
        bbox = draw.textbbox((0, 0), char, font=font)
        tw = bbox[2] - bbox[0]
        jitter_x = random.randint(-int(fs*0.02), int(fs*0.02))
        x = cx - tw // 2 + jitter_x
        draw.text((x, cur_y), char, fill=stamp_color, font=font)
        cur_y += char_heights[i]
    
    # 획 두껍게 (PIL MaxFilter로 팽창)
    r, g, b, a = img.split()
    a = a.filter(MaxFilter(5))
    img = Image.merge("RGBA", (r, g, b, a))
    
    # 잉크 효과
    arr = np.array(img)
    
    # 뽕뽕 구멍 (잉크 있는 곳에만, 적당히)
    alpha = arr[:,:,3]
    ink_pixels = np.argwhere(alpha > 50)
    n_holes = min(300, len(ink_pixels) // 15)
    for _ in range(n_holes):
        idx = random.randint(0, len(ink_pixels) - 1)
        hy, hx = ink_pixels[idx]
        hr = random.randint(3, 7)
        y1, y2 = max(0, hy - hr), min(img_h, hy + hr)
        x1, x2 = max(0, hx - hr), min(img_w, hx + hr)
        yy, xx = np.mgrid[y1:y2, x1:x2]
        mask = ((xx - hx)**2 + (yy - hy)**2) < hr**2
        arr[y1:y2, x1:x2, 3][mask] = 0
    
    # 미세 노이즈 (약하게 - 진하게 유지)
    noise = np.random.random((img_h, img_w))
    arr[:,:,3] = (arr[:,:,3].astype(float) * (noise * 0.1 + 0.9)).clip(0, 255).astype(np.uint8)
    
    # 중심 진하고 가장자리 연하게
    yy, xx = np.mgrid[0:img_h, 0:img_w]
    dist = np.sqrt(((xx - cx) / rx) ** 2 + ((yy - cy) / ry) ** 2)
    center_fade = np.clip(1.0 - (dist - 0.7) * 1.0, 0.6, 1.0)
    arr[:,:,3] = (arr[:,:,3].astype(float) * center_fade).clip(0, 255).astype(np.uint8)
    
    # 테두리 끊김
    for _ in range(random.randint(1, 3)):
        angle = random.uniform(0, 6.283)
        arc_len = random.uniform(0.08, 0.15)
        for t in np.linspace(angle, angle + arc_len, 20):
            bx = int(cx + rx * np.cos(t))
            by = int(cy + ry * np.sin(t))
            r = border_width + 2
            y1, y2 = max(0, by - r), min(img_h, by + r)
            x1, x2 = max(0, bx - r), min(img_w, bx + r)
            arr[y1:y2, x1:x2, 3] = (arr[y1:y2, x1:x2, 3].astype(float) * 0.2).astype(np.uint8)
    
    img = Image.fromarray(arr)
    img = img.filter(ImageFilter.GaussianBlur(radius=2.0))
    
    # 번짐
    spread = img.filter(ImageFilter.GaussianBlur(radius=5))
    s_arr = np.array(spread)
    s_arr[:,:,3] = (s_arr[:,:,3].astype(float) * 0.15).clip(0, 255).astype(np.uint8)
    img = Image.alpha_composite(Image.fromarray(s_arr), img)
    
    # 축소
    final_w = int(size * w_ratio)
    img = img.resize((final_w, size), Image.LANCZOS)
    img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
    
    # 랜덤 회전
    angle = random.uniform(-10, 10)
    img = img.rotate(angle, expand=True, resample=Image.BICUBIC)
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    img.save(tmp.name)
    return tmp.name


def prepare_stamp_image(stamp_bytes):
    """업로드된 인감도장 이미지에서 도장 부분만 추출하고 배경 투명화
    
    A4 스캔(300dpi) 기준으로 도장의 실제 물리적 크기(mm)를 계산하여 반환.
    
    Args:
        stamp_bytes: 도장 이미지/PDF 바이트
    
    Returns:
        (임시 PNG 파일 경로, 실제 너비mm, 실제 높이mm) 또는 None
    """
    from PIL import Image
    import numpy as np
    import tempfile
    import io
    
    img = None
    dpi = 300  # 스캔 DPI 기본값
    
    # PDF인 경우 첫 페이지를 이미지로 변환
    try:
        doc = fitz.open(stream=stamp_bytes, filetype="pdf")
        if doc.page_count > 0:
            pix = doc[0].get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
    except:
        pass
    
    # 이미지인 경우
    if img is None:
        try:
            img = Image.open(io.BytesIO(stamp_bytes))
            # 이미지 DPI 확인
            img_dpi = img.info.get("dpi", (300, 300))
            dpi = img_dpi[0] if isinstance(img_dpi, tuple) else img_dpi
            if dpi < 72:
                dpi = 300
            if img.mode != "RGB":
                img = img.convert("RGB")
        except:
            return None
    
    arr = np.array(img)
    
    # 빨간색/분홍색 도장 영역 찾기
    # R이 G보다 확실히 크고, 밝은 배경이 아닌 픽셀
    red_mask = (arr[:,:,0] > 80) & \
               (arr[:,:,0].astype(int) - arr[:,:,1].astype(int) > 30) & \
               ~((arr[:,:,0] > 200) & (arr[:,:,1] > 200) & (arr[:,:,2] > 200))
    
    if not np.any(red_mask):
        # 빨간 영역 없으면 원본 그대로 사용
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        img.save(tmp.name)
        # 크기 추정 불가 → 기본 18mm
        return tmp.name, 18.0, 18.0
    
    # 도장 위치 찾기
    rows = np.any(red_mask, axis=1)
    cols = np.any(red_mask, axis=0)
    rmin, rmax = np.where(rows)[0][[0, -1]]
    cmin, cmax = np.where(cols)[0][[0, -1]]
    
    # 도장 실제 크기(mm) 계산: px / dpi * 25.4mm
    stamp_w_px = cmax - cmin
    stamp_h_px = rmax - rmin
    stamp_w_mm = stamp_w_px / dpi * 25.4
    stamp_h_mm = stamp_h_px / dpi * 25.4
    
    # 여유 마진
    margin = 25
    rmin = max(0, rmin - margin)
    rmax = min(arr.shape[0], rmax + margin)
    cmin = max(0, cmin - margin)
    cmax = min(arr.shape[1], cmax + margin)
    
    # 도장 부분 잘라내기
    stamp = arr[rmin:rmax, cmin:cmax]
    stamp_rgba = Image.fromarray(stamp).convert("RGBA")
    data = np.array(stamp_rgba)
    
    # 밝은 배경을 투명으로
    bright = (data[:,:,0] > 190) & (data[:,:,1] > 190) & (data[:,:,2] > 190)
    data[bright] = [255, 255, 255, 0]
    
    result = Image.fromarray(data)
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    result.save(tmp.name)
    return tmp.name, stamp_w_mm, stamp_h_mm


def insert_image(page, x_pct, y_pct, image_path, width_pct=8, height_pct=4, width_mm=None, height_mm=None):
    """PDF 페이지에 퍼센트 좌표로 이미지 삽입 (투명 배경 지원)
    
    Args:
        page: fitz.Page 객체
        x_pct: X 좌표 (0~100, 페이지 너비 기준 %)
        y_pct: Y 좌표 (0~100, 페이지 높이 기준 %)
        image_path: 이미지 파일 경로 (또는 "path|w_mm|h_mm" 형식)
        width_pct: 이미지 너비 (페이지 너비 기준 %, 기본 8%)
        height_pct: 이미지 높이 (페이지 높이 기준 %, 기본 4%)
        width_mm: 실제 너비(mm) - 지정시 width_pct 무시
        height_mm: 실제 높이(mm) - 지정시 height_pct 무시
    """
    from pathlib import Path
    from PIL import Image
    import io
    
    # "path|w_mm|h_mm" 형식 파싱 (stamp용)
    if isinstance(image_path, str) and "|" in image_path:
        parts = image_path.split("|")
        image_path = parts[0]
        if len(parts) >= 3:
            try:
                width_mm = float(parts[1])
                height_mm = float(parts[2])
            except:
                pass
    
    img_path = Path(image_path)
    if not img_path.exists():
        return
    
    img = Image.open(img_path)
    
    rect = page.rect
    x = rect.width * (x_pct / 100)
    y = rect.height * (y_pct / 100)
    
    # mm 기반 크기 (1mm = 72/25.4 pt)
    if width_mm and height_mm:
        w = width_mm * 72 / 25.4
        h = height_mm * 72 / 25.4
    else:
        w = rect.width * (width_pct / 100)
        h = rect.height * (height_pct / 100)
    
    img_rect = fitz.Rect(x, y - h, x + w, y)
    
    if img.mode == "RGBA":
        r, g, b, a = img.split()
        rgb_img = Image.merge("RGB", (r, g, b))
        
        rgb_bytes = io.BytesIO()
        rgb_img.save(rgb_bytes, format="PNG")
        rgb_bytes.seek(0)
        
        mask_bytes = io.BytesIO()
        a.save(mask_bytes, format="PNG")
        mask_bytes.seek(0)
        
        page.insert_image(img_rect, stream=rgb_bytes.read(), mask=mask_bytes.read())
    else:
        if img.mode != "RGB":
            img = img.convert("RGB")
        img_bytes = io.BytesIO()
        img.save(img_bytes, format="PNG")
        img_bytes.seek(0)
        
        page.insert_image(img_rect, stream=img_bytes.read())


def insert_multiline(page, x_pct, y_pct, text, font_size=11, line_spacing=3.0):
    """여러 줄 텍스트 삽입
    
    Args:
        line_spacing: 줄 간격 (y_pct 단위, 기본 3.0%)
    """
    lines = text.split("\n")
    for i, line in enumerate(lines):
        insert_text(page, x_pct, y_pct + (i * line_spacing), line, font_size)


def fill_form_by_coords(page, coords_fields, data_map, page_num=1):
    """좌표 JSON 기반으로 폼 자동 채우기
    
    Args:
        page: fitz.Page 객체
        coords_fields: 좌표 JSON의 fields 리스트
        data_map: {field_id: 값} 딕셔너리
        page_num: 현재 페이지 번호 (1부터)
    """
    for field in coords_fields:
        field_id = field.get("field_id", "")
        if field_id not in data_map or not data_map[field_id]:
            continue
        
        value = data_map[field_id]
        field_type = field.get("type", "text")
        font_size = field.get("font_size", 11)
        
        if field_type == "checkbox":
            insert_text(page, field["x_pct"], field["y_pct"], "✔", font_size)
        elif field_type == "image":
            # 이미지 삽입 (value = 파일 경로)
            width_pct = field.get("width_pct", 8)
            height_pct = field.get("height_pct", 4)
            insert_image(page, field["x_pct"], field["y_pct"], value, width_pct, height_pct)
        else:
            spacing = field.get("spacing", 0)
            if spacing > 0:
                insert_text_spaced(page, field["x_pct"], field["y_pct"], value, font_size, spacing)
            else:
                insert_text(page, field["x_pct"], field["y_pct"], value, font_size)


def generate_warrant(template_path, client, agent, creditor_name, delegation_text, warrant_date, stamp_path=None):
    """기본 위임장 PDF 생성
    
    Args:
        template_path: 위임장 양식 PDF 경로
        client: {"name", "birth", "address", "phone"}
        agent: {"name", "birth", "address", "phone", "fax"}
        creditor_name: 채권사명
        delegation_text: 위임 사항 텍스트
        warrant_date: 날짜 문자열 (예: "2026.02.23")
        stamp_path: 인감도장 이미지 경로 (팩스용, 선택)
    
    Returns:
        fitz.Document 객체
    """
    # 좌표 로드
    coords = load_coords("위임장_기본")
    
    doc = fitz.open(str(template_path))
    page = doc[0]
    
    # 텍스트는 항상 하드코딩으로 채움
    _fill_warrant_fallback(page, client, agent, creditor_name, delegation_text, warrant_date)
    
    # 좌표 파일에 image 필드(stamp 등)가 있으면 추가로 처리
    if coords and stamp_path:
        fields = coords.get("pages", [{}])[0].get("fields", [])
        image_fields = [f for f in fields if f.get("type") == "image"]
        if image_fields:
            data_map = {"stamp": stamp_path}
            fill_form_by_coords(page, image_fields, data_map)
    
    return doc


def _fill_warrant_fallback(page, client, agent, creditor_name, delegation_text, warrant_date):
    """위임장 좌표 JSON이 없을 때 기본 좌표로 채우기 (폴백)
    
    양식 라벨 위치 기준:
      위임인: 성명 y=19.7, 생년월일 y=22.8, 주소 y=25.8, 전화 y=28.8
      수임인: 성명 y=34.1, 생년월일 y=37.1, 주소 y=40.1, 전화 y=43.2, FAX y=46.2
      위임일자 y=84.9, 위임인(인) y=87.5, 귀중 y=94.2 (size=29)
    """
    # 위임인
    insert_text(page, 38.5, 19.7, client["name"], 11)
    insert_text(page, 38.5, 22.8, client["birth"], 11)
    insert_text(page, 38.5, 25.8, client["address"], 10)
    insert_text(page, 38.5, 28.8, client["phone"], 11)
    # 수임인
    insert_text(page, 38.5, 34.1, agent["name"], 11)
    insert_text(page, 38.5, 37.1, agent["birth"], 11)
    insert_text(page, 38.5, 40.1, agent["address"], 10)
    insert_text(page, 38.5, 43.2, agent["phone"], 11)
    insert_text(page, 38.5, 46.2, agent["fax"], 11)
    # 위임사항
    insert_multiline(page, 12.5, 65.0, delegation_text, 10, line_spacing=2.0)
    # 하단
    insert_text(page, 69.0, 84.9, warrant_date, 11)
    insert_text(page, 69.0, 87.5, client["name"], 11)
    # 채권사명 + 귀중 (원본 "귀중" 덮고, 채권사명 뒤에 붙여서 출력)
    # 원본 "귀중" 위치: x=47.9%, y=94.2%, size=29 → 흰색으로 덮기
    rect = page.rect
    gx = rect.width * 0.46
    gy = rect.height * 0.90
    gw = rect.width * 0.60
    gh = rect.height * 0.96
    page.draw_rect(fitz.Rect(gx, gy, gw, gh), color=(1, 1, 1), fill=(1, 1, 1))
    # 채권사명 + 귀중을 하나로
    combined = f"{creditor_name}  귀중"
    # 중앙 정렬: 텍스트 너비 계산 후 x 위치 결정
    font = fitz.Font(fontfile=FONT_PATH)
    text_width = font.text_length(combined, fontsize=26)
    center_x = (rect.width - text_width) / 2
    center_x_pct = center_x / rect.width * 100
    insert_text(page, center_x_pct, 94.2, combined, 26)


def generate_application_form(creditor_name, doc_type, client, agent, warrant_date, **kwargs):
    """신청서 양식 PDF 생성 (채권사별 자동 매핑)
    
    Args:
        creditor_name: 채권사명
        doc_type: 서류종류
        client, agent, warrant_date: 기본 정보
        **kwargs: 저축은행 정보 등
    
    Returns:
        리스트 [fitz.Document, ...] (매핑 없으면 빈 리스트)
    """
    from modules.config_loader import get_form_info, load_coords, get_template_path
    
    form_info_list = kwargs.pop("_form_info_override", None)
    if not form_info_list:
        form_info_list = get_form_info(creditor_name, doc_type)
    if not form_info_list:
        return []
    # 단일 dict인 경우 리스트로 변환
    if isinstance(form_info_list, dict):
        form_info_list = [form_info_list]
    
    # 공통 데이터 매핑 구성
    date_parts = warrant_date.split(".")
    date_year = date_parts[0] if len(date_parts) >= 1 else ""
    date_month = date_parts[1] if len(date_parts) >= 2 else ""
    date_day = date_parts[2] if len(date_parts) >= 3 else ""
    
    # +1일, +7일 날짜 계산
    from datetime import datetime, timedelta
    try:
        base = datetime.strptime(warrant_date, "%Y.%m.%d")
        d1 = base + timedelta(days=1)
        d7 = base + timedelta(days=7)
    except:
        d1 = d7 = None
    
    c_birth = client.get("birth", "")
    if len(c_birth) == 6:
        c_birth_year = c_birth[:2]
        c_birth_month = c_birth[2:4]
        c_birth_day = c_birth[4:6]
        c_birth_year_full = ("19" if int(c_birth[:2]) > 30 else "20") + c_birth[:2]
    elif len(c_birth) >= 8:
        c_birth_year_full = c_birth[:4]
        c_birth_year = c_birth[2:4]
        c_birth_month = c_birth[4:6]
        c_birth_day = c_birth[6:8]
    else:
        c_birth_year = c_birth_year_full = c_birth_month = c_birth_day = ""

    a_birth = agent.get("birth", "")
    if len(a_birth) == 6:
        a_birth_year = a_birth[:2]
        a_birth_month = a_birth[2:4]
        a_birth_day = a_birth[4:6]
        a_birth_year_full = ("19" if int(a_birth[:2]) > 30 else "20") + a_birth[:2]
    elif len(a_birth) >= 8:
        a_birth_year_full = a_birth[:4]
        a_birth_year = a_birth[2:4]
        a_birth_month = a_birth[4:6]
        a_birth_day = a_birth[6:8]
    else:
        a_birth_year = a_birth_year_full = a_birth_month = a_birth_day = ""

    data_map = {
        "client_name": client.get("name", ""),
        "client_birth": c_birth,
        "client_birth_year": c_birth_year,
        "client_birth_year_full": c_birth_year_full,
        "client_birth_month": c_birth_month,
        "client_birth_day": c_birth_day,
        "client_phone": client.get("phone", ""),
        "client_address": client.get("address", ""),
        "client_id_front": client.get("id_front", c_birth),
        "client_id_back": client.get("id_back", ""),
        "client_id_full": (client.get("id_front", c_birth) + "-" + client.get("id_back", "")) if client.get("id_back") else "",
        "agent_name": agent.get("name", ""),
        "agent_birth": a_birth,
        "agent_birth_year": a_birth_year,
        "agent_birth_year_full": a_birth_year_full,
        "agent_birth_month": a_birth_month,
        "agent_birth_day": a_birth_day,
        "agent_phone": agent.get("phone", ""),
        "agent_address": agent.get("address", ""),
        "agent_fax": agent.get("fax", ""),
        "agent_id_front": a_birth,
        "agent_id_back": agent.get("id_back", ""),
        "agent_id_full": agent.get("id_full", ""),
        "agent_sign": agent.get("sign_path", ""),
        "date": warrant_date,
        "date_year": date_year,
        "date_year_suffix": date_year[2:] if len(date_year) >= 4 else date_year,
        "date_month": date_month,
        "date_day": date_day,
        "date1_year": d1.strftime("%Y") if d1 else "",
        "date1_year_suffix": d1.strftime("%y") if d1 else "",
        "date1_month": d1.strftime("%m") if d1 else "",
        "date1_day": d1.strftime("%d") if d1 else "",
        "date7_year": d7.strftime("%Y") if d7 else "",
        "date7_year_suffix": d7.strftime("%y") if d7 else "",
        "date7_month": d7.strftime("%m") if d7 else "",
        "date7_day": d7.strftime("%d") if d7 else "",
        "fixed_대리인": "대리인",
        "fixed_위임인": "위임인",
        "fixed_본인": "본인",
        "fixed_신청인": "신청인",
        "fixed_대표자": "대표자",
        "fixed_관계_대리인": "대 리 인",
        "bank_name": kwargs.get("bank_name", ""),
        "bank_tel": kwargs.get("bank_tel", ""),
        "bank_fax": kwargs.get("bank_fax", ""),
        "bank_branch": kwargs.get("bank_branch", ""),
        "creditor_name": creditor_name,
        "bank_period": kwargs.get("bank_period", ""),
        "bank_date_from": kwargs.get("bank_date_from", ""),
        "bank_date_to": kwargs.get("bank_date_to", ""),
        "card_period": kwargs.get("card_period", ""),
        "card_date_from": kwargs.get("card_date_from", ""),
        "card_date_to": kwargs.get("card_date_to", ""),
    }
    
    # 인감도장 (팩스용)
    if kwargs.get("stamp_path"):
        data_map["stamp"] = kwargs["stamp_path"]
    
    results = []
    for fi in form_info_list:
        template_path = get_template_path(fi["template"])
        if not template_path.exists():
            continue
        coords = load_coords(fi["coords"])
        if not coords:
            continue
        
        doc = fitz.open(str(template_path))
        for page_info in coords.get("pages", []):
            page_num = page_info.get("page", 1) - 1
            if page_num >= doc.page_count:
                continue
            page = doc[page_num]
            fields = page_info.get("fields", [])
            fill_form_by_coords(page, fields, data_map)
        
        results.append(doc)
    
    return results


def mask_id_back_digits(doc):
    """(사용하지 않음 - mask_id_bytes로 대체)"""
    pass


def mask_id_bytes(id_bytes):
    """신분증/인감증명서 이미지에서 주민번호 뒷자리 마스킹 (OCR 기반)
    
    Returns:
        마스킹된 PDF 바이트
    """
    from PIL import Image, ImageDraw
    import numpy as np
    import re
    import io
    
    if not id_bytes:
        return id_bytes
    
    # 바이트 → 이미지 (300dpi)
    try:
        doc = fitz.open(stream=id_bytes, filetype="pdf")
        page = doc[0]
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
    except:
        try:
            img = Image.open(io.BytesIO(id_bytes)).convert("RGB")
        except:
            return id_bytes
    
    try:
        import pytesseract
    except ImportError:
        return id_bytes
    
    # OCR: 숫자+하이픈만 인식
    config = '--psm 6 -c tessedit_char_whitelist=0123456789-'
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, config=config)
    
    draw = ImageDraw.Draw(img)
    masked = False
    
    for i, text in enumerate(data['text']):
        t = text.strip()
        if not t:
            continue
        
        # 주민번호 패턴: 6자리-7자리
        match = re.search(r'(\d{6})-(\d{7})', t)
        if match:
            x = data['left'][i]
            y = data['top'][i]
            w = data['width'][i]
            h = data['height'][i]
            
            # 뒷자리 위치 계산 (하이픈 뒤 첫째자리 이후 6자리)
            hyphen_idx = t.index('-')
            char_width = w / len(t)
            
            # 성별숫자(1자리) 이후부터 마스킹
            mask_start = x + int(char_width * (hyphen_idx + 2))
            mask_end = x + w + int(char_width * 0.3)
            
            margin = int(h * 0.15)
            draw.rectangle([mask_start, y - margin, mask_end, y + h + margin], fill=(0, 0, 0))
            masked = True
    
    if not masked:
        return id_bytes
    
    # 다시 PDF 바이트로
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    
    new_doc = fitz.open()
    page = new_doc.new_page(width=img.width * 72 / 300, height=img.height * 72 / 300)
    page.insert_image(page.rect, stream=buf.getvalue())
    result = new_doc.tobytes()
    new_doc.close()
    
    return result


def bytes_to_pdf(file_bytes):
    """이미지/PDF 바이트를 fitz.Document로 변환
    
    Args:
        file_bytes: 파일 바이트 데이터
    
    Returns:
        fitz.Document 또는 None
    """
    # PDF인 경우
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        if doc.page_count > 0:
            return doc
    except:
        pass
    
    # 이미지인 경우 (PNG, JPG 등)
    for fmt in ["png", "jpeg", "jpg"]:
        try:
            img_doc = fitz.open(stream=file_bytes, filetype=fmt)
            pdf_doc = fitz.open()
            page = img_doc[0]
            new_page = pdf_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, img_doc, 0)
            return pdf_doc
        except:
            continue
    
    return None


def merge_documents(doc_list):
    """여러 fitz.Document를 하나의 PDF로 합치기
    
    Args:
        doc_list: fitz.Document 객체 리스트
    
    Returns:
        합쳐진 fitz.Document
    """
    merged = fitz.open()
    for doc in doc_list:
        if doc and doc.page_count > 0:
            try:
                merged.insert_pdf(doc)
            except (RuntimeError, ValueError):
                for page_num in range(doc.page_count):
                    src_page = doc[page_num]
                    new_page = merged.new_page(width=src_page.rect.width, height=src_page.rect.height)
                    # 이미지로 변환 후 삽입
                    pix = src_page.get_pixmap(dpi=150)
                    img_data = pix.tobytes("png")
                    new_page.insert_image(new_page.rect, stream=img_data)
    return merged


def build_gov_bundle(form_data_list, client, agent, warrant_date, client_id_bytes, agent_id_bytes, stamp_bytes=None):
    """관공서 서류 묶음 생성
    
    Args:
        form_data_list: [{"form_name", "template", "coords", "extra": {고유필드}}]
        client, agent: 기본 정보
        warrant_date: 위임일자
        client_id_bytes, agent_id_bytes: 신분증
        stamp_bytes: 인감도장 이미지 바이트 (선택)
    
    Returns:
        fitz.Document (양식들 + 신분증)
    """
    from modules.config_loader import load_coords, get_template_path
    from datetime import datetime, timedelta
    
    # 날짜 파싱
    date_parts = warrant_date.split(".")
    date_year = date_parts[0] if len(date_parts) >= 1 else ""
    date_month = date_parts[1] if len(date_parts) >= 2 else ""
    date_day = date_parts[2] if len(date_parts) >= 3 else ""
    
    # 생년월일 파싱 (위임인)
    c_birth = client.get("birth", "")
    if len(c_birth) == 6:
        c_birth_year = c_birth[:2]
        c_birth_month = c_birth[2:4]
        c_birth_day = c_birth[4:6]
        c_birth_year_full = ("19" if int(c_birth[:2]) > 30 else "20") + c_birth[:2]
    elif len(c_birth) >= 8:
        c_birth_year_full = c_birth[:4]
        c_birth_year = c_birth[2:4]
        c_birth_month = c_birth[4:6]
        c_birth_day = c_birth[6:8]
    else:
        c_birth_year = c_birth_year_full = c_birth_month = c_birth_day = ""

    # 생년월일 파싱 (대리인)
    a_birth = agent.get("birth", "")
    if len(a_birth) == 6:
        a_birth_year = a_birth[:2]
        a_birth_month = a_birth[2:4]
        a_birth_day = a_birth[4:6]
        a_birth_year_full = ("19" if int(a_birth[:2]) > 30 else "20") + a_birth[:2]
    elif len(a_birth) >= 8:
        a_birth_year_full = a_birth[:4]
        a_birth_year = a_birth[2:4]
        a_birth_month = a_birth[4:6]
        a_birth_day = a_birth[6:8]
    else:
        a_birth_year = a_birth_year_full = a_birth_month = a_birth_day = ""

    # 공통 data_map
    base_map = {
        "client_name": client.get("name", ""),
        "client_birth": c_birth,
        "client_birth_year": c_birth_year,
        "client_birth_year_full": c_birth_year_full,
        "client_birth_month": c_birth_month,
        "client_birth_day": c_birth_day,
        "client_phone": client.get("phone", ""),
        "client_address": client.get("address", ""),
        "client_id_front": client.get("id_front", c_birth),
        "client_id_back": client.get("id_back", ""),
        "client_id_full": (client.get("id_front", c_birth) + "-" + client.get("id_back", "")) if client.get("id_back") else "",
        "agent_name": agent.get("name", ""),
        "agent_birth": a_birth,
        "agent_birth_year": a_birth_year,
        "agent_birth_year_full": a_birth_year_full,
        "agent_birth_month": a_birth_month,
        "agent_birth_day": a_birth_day,
        "agent_phone": agent.get("phone", ""),
        "agent_address": agent.get("address", ""),
        "agent_fax": agent.get("fax", ""),
        "agent_id_front": a_birth,
        "agent_id_back": agent.get("id_back", ""),
        "agent_id_full": agent.get("id_full", ""),
        "agent_sign": agent.get("sign_path", ""),
        "date": warrant_date,
        "date_year": date_year,
        "date_year_suffix": date_year[2:] if len(date_year) >= 4 else date_year,
        "date_month": date_month,
        "date_day": date_day,
    }

    pdfs = []

    # 1) 각 양식 채우기
    for form_data in form_data_list:
        form_name = form_data.get("form_name", "")
        template_path = get_template_path(form_data["template"])
        if not template_path.exists():
            continue
        coords = load_coords(form_data["coords"])
        
        doc = fitz.open(str(template_path))
        
        # 좌표가 있으면 채우기
        if coords:
            # 고유필드를 data_map에 추가
            data_map = {**base_map, **form_data.get("extra", {})}
            
            # 집주인 이름이 있으면 도장 자동 생성
            landlord_name = data_map.get("landlord_name", "")
            if landlord_name:
                stamp_path = generate_name_stamp(landlord_name)
                # 인감도장과 동일한 크기 (약 8x12mm 타원형)
                data_map["landlord_stamp"] = f"{stamp_path}|8|12"
            
            # 위임인 인감도장
            if stamp_bytes:
                result = prepare_stamp_image(stamp_bytes)
                if result:
                    path, w_mm, h_mm = result
                    data_map["stamp"] = f"{path}|{w_mm}|{h_mm}"
            
            for page_info in coords.get("pages", []):
                page_num = page_info.get("page", 1) - 1
                if page_num >= doc.page_count:
                    continue
                page = doc[page_num]
                fill_form_by_coords(page, page_info.get("fields", []), data_map)
        
        # 양면인쇄: 무상거주사실확인서는 빈 페이지 불필요
        if not form_data.get("duplex") and "무상거주" not in form_name:
            page_w = doc[0].rect.width
            page_h = doc[0].rect.height
            for i in range(doc.page_count - 1, -1, -1):
                doc.new_page(width=page_w, height=page_h, pno=i + 1)

        pdfs.append(doc)
        
        # 서류별 신분증 첨부
        # - 무상거주사실확인서: 신분증 없음
        # - 법원 서류: 위임인 신분증만
        # - 나머지: 위임인 + 수임인 신분증
        if "무상거주" not in form_name:
            # 위임인 신분증 (공통)
            if client_id_bytes:
                client_id = bytes_to_pdf(client_id_bytes)
                if client_id:
                    pw = client_id[0].rect.width
                    ph = client_id[0].rect.height
                    client_id.new_page(width=pw, height=ph)
                    pdfs.append(client_id)
            
            # 수임인 신분증 (법원 서류 제외, 단 법원_위임장은 포함)
            is_court_no_agent = "법원" in form_name and "위임장" not in form_name
            if not is_court_no_agent and agent_id_bytes:
                agent_id = bytes_to_pdf(agent_id_bytes)
                if agent_id:
                    pw = agent_id[0].rect.width
                    ph = agent_id[0].rect.height
                    agent_id.new_page(width=pw, height=ph)
                    pdfs.append(agent_id)

    return merge_documents(pdfs)


def build_creditor_bundle(template_path, client, agent, creditor, warrant_date, client_id_bytes, agent_id_bytes, stamp_bytes=None, seal_cert_bytes=None):
    """채권사 1개에 대한 서류 묶음 생성
    
    Returns:
        fitz.Document (위임장 + 신청서 + 신분증 + 추가서류)
    """
    from modules.config_loader import get_needs_date_labels, get_bundle_type, load_coords, get_template_path
    from pathlib import Path
    
    # 팩스 발급 여부 확인
    from modules.config_loader import get_issue_info
    _issue_info = get_issue_info(creditor["name"])
    is_fax = _issue_info and _issue_info.get("발급방법", "").strip() == "팩스"
    
    # 위임사항 텍스트 생성 (번호 매기기, 줄바꿈)
    settings = load_settings()
    needs_date_labels = set(get_needs_date_labels(settings))
    
    lines = []
    for i, d in enumerate(creditor.get("docs", []), 1):
        doc_type = d.get("type", "")
        if doc_type in needs_date_labels:
            date_from = d.get("date_from", "").replace("-", ".")
            date_to = d.get("date_to", "").replace("-", ".")
            account = d.get("account", "")
            lines.append(f"{i}. {doc_type} {date_from}~{date_to}")
            if account:
                lines.append(f"   ({account})")
        elif doc_type == "기타":
            custom = d.get("custom", "기타")
            lines.append(f"{i}. {custom}")
        else:
            lines.append(f"{i}. {doc_type}")
    
    delegation_text = "\n".join(lines) if lines else "서류 발급"

    # 법률사무소 기재 필요 채권사 → 위임사항에 사무소 정보 추가
    if _issue_info and _issue_info.get("법률사무소기재"):
        law_firm = creditor.get("law_firm", {})
        lf_name = law_firm.get("name", "")
        lf_tel = law_firm.get("tel", "")
        if lf_name:
            delegation_text += f"\n\n[담당 법률사무소]\n{lf_name}"
            if lf_tel:
                delegation_text += f"\nTEL: {lf_tel}"
    
    pdfs = []
    
    # 팩스 발급 시 도장 이미지 준비
    stamp_path = None
    if is_fax and stamp_bytes:
        result = prepare_stamp_image(stamp_bytes)
        if result:
            path, w_mm, h_mm = result
            # "경로|너비mm|높이mm" 형식으로 전달
            stamp_path = f"{path}|{w_mm}|{h_mm}"
    
    # 1) 위임장
    warrant = generate_warrant(template_path, client, agent, creditor["name"], delegation_text, warrant_date, stamp_path=stamp_path)
    pdfs.append(warrant)
    
    # 번들 타입 확인 (대부업체 등)
    bundle_type_name, bundle_config = get_bundle_type(creditor["name"])
    
    # 2) 번들 타입 공통 양식 (자료송부청구서 등)
    if bundle_config:
        for cf in bundle_config.get("common_forms", []):
            common_forms = generate_application_form(
                creditor["name"], "__common__", client, agent, warrant_date,
                _form_info_override=[cf], stamp_path=stamp_path
            )
            pdfs.extend(common_forms)
    
    # 3) 신청서 (채권사별 매핑된 양식이 있으면 자동 생성)
    from modules.config_loader import get_savings_bank_info
    bank_info = get_savings_bank_info(creditor["name"])
    bank_kwargs = {}
    if bank_info:
        bank_kwargs = {
            "bank_name": creditor["name"],
            "bank_tel": bank_info.get("tel", ""),
            "bank_fax": bank_info.get("fax", ""),
            "bank_branch": bank_info.get("branch", ""),
        }
    
    # 통장/카드 기간 추출 → 신청서에도 전달
    for d in creditor.get("docs", []):
        dtype = d.get("type", "")
        if dtype == "통장거래내역":
            df = d.get("date_from", "").replace("-", ".")
            dt = d.get("date_to", "").replace("-", ".")
            bank_kwargs["bank_period"] = f"{df}~{dt}" if df and dt else ""
            bank_kwargs["bank_date_from"] = df
            bank_kwargs["bank_date_to"] = dt
        elif dtype == "카드거래내역":
            df = d.get("date_from", "").replace("-", ".")
            dt = d.get("date_to", "").replace("-", ".")
            bank_kwargs["card_period"] = f"{df}~{dt}" if df and dt else ""
            bank_kwargs["card_date_from"] = df
            bank_kwargs["card_date_to"] = dt
    
    # 팩스 발급 시 도장 경로 전달
    if stamp_path:
        bank_kwargs["stamp_path"] = stamp_path
    
    added_templates = set()  # 중복 방지 (동일 양식 1번만)
    for d in creditor.get("docs", []):
        doc_type = d.get("type", "")
        # "기타" 선택 시 custom 텍스트로도 form_mapping 매칭 시도
        if doc_type == "기타" and d.get("custom"):
            app_forms = generate_application_form(creditor["name"], d["custom"], client, agent, warrant_date, **bank_kwargs)
        else:
            app_forms = generate_application_form(creditor["name"], doc_type, client, agent, warrant_date, **bank_kwargs)
        for form in app_forms:
            form_key = form.name if hasattr(form, 'name') else id(form)
            if form_key not in added_templates:
                added_templates.add(form_key)
                pdfs.append(form)
            else:
                form.close()
    
    # 4) issue_manual 추가서류
    from modules.config_loader import get_issue_info, get_id_card_path
    issue_info = get_issue_info(creditor["name"])
    if issue_info and "추가서류" in issue_info:
        extras = issue_info["추가서류"]
        BASE_DIR = Path(__file__).parent.parent
        ID_CARDS_DIR = BASE_DIR / "id_cards"
        
        # 대부업체가 아닌 경우에만: 자료송부청구서, 사업자등록증, 재직증명서
        # (대부업체는 bundle_types로 이미 자동 첨부됨)
        if not bundle_config:
            # 자료송부청구서
            if extras.get("자료송부청구서"):
                cf = {"template": "대부업체/자료송부청구서.pdf", "coords": "대부업체/자료송부청구서"}
                extra_forms = generate_application_form(
                    creditor["name"], "__common__", client, agent, warrant_date,
                    _form_info_override=[cf], stamp_path=stamp_path
                )
                pdfs.extend(extra_forms)
            
            # 사업자등록증
            if extras.get("사업자등록증"):
                biz_path = ID_CARDS_DIR / "company" / "리셋플러스_사업자등록증.pdf"
                if biz_path.exists():
                    pdfs.append(fitz.open(str(biz_path)))
            
            # 재직증명서
            if extras.get("재직증명서"):
                cert_filename = agent.get("cert", "")
                if cert_filename:
                    cert_path = ID_CARDS_DIR / "agents" / cert_filename
                    if cert_path.exists():
                        cert_doc = fitz.open(str(cert_path))
                        # 재직증명서 좌표로 날짜 채우기
                        coords_name_base = "대부업체/재직증명서"
                        agent_name = agent.get("name", "")
                        doc_name = coords_name_base.split("/")[-1]
                        folder = coords_name_base.split("/")[0]
                        cert_coords_name = f"{folder}/{agent_name}_{doc_name}"
                        cert_coords = load_coords(cert_coords_name)
                        if cert_coords:
                            from datetime import datetime
                            try:
                                base = datetime.strptime(warrant_date, "%Y.%m.%d")
                            except:
                                base = datetime.now()
                            date_map = {
                                "date_year": base.strftime("%Y"),
                                "date_year_suffix": base.strftime("%y"),
                                "date_month": base.strftime("%m"),
                                "date_day": base.strftime("%d"),
                            }
                            for page_info in cert_coords.get("pages", []):
                                page_num = page_info.get("page", 1) - 1
                                if page_num < cert_doc.page_count:
                                    fill_form_by_coords(cert_doc[page_num], page_info.get("fields", []), date_map)
                        pdfs.append(cert_doc)
        
        # 대부업체든 아니든 공통 적용: 개인정보열람요구서
        if extras.get("개인정보열람요구서"):
            pi_template = get_template_path("개인정보열람요구서.pdf")
            if pi_template.exists():
                pi_coords = load_coords("개인정보열람요구서")
                if pi_coords:
                    pi_doc = fitz.open(str(pi_template))
                    date_parts = warrant_date.split(".")
                    pi_data = {
                        "client_name": client.get("name", ""),
                        "client_phone": client.get("phone", ""),
                        "client_birth": client.get("birth", ""),
                        "client_address": client.get("address", ""),
                        "agent_name": agent.get("name", ""),
                        "agent_phone": agent.get("phone", ""),
                        "agent_birth": agent.get("id_full", "").split("-")[0] if "-" in agent.get("id_full", "") else "",
                        "agent_address": agent.get("address", ""),
                        "creditor_name": creditor["name"],
                        "date_year": date_parts[0] if len(date_parts) >= 1 else "",
                        "date_month": date_parts[1] if len(date_parts) >= 2 else "",
                        "date_day": date_parts[2] if len(date_parts) >= 3 else "",
                    }
                    for page_info in pi_coords.get("pages", []):
                        page_num = page_info.get("page", 1) - 1
                        if page_num < pi_doc.page_count:
                            fill_form_by_coords(pi_doc[page_num], page_info.get("fields", []), pi_data)
                    pdfs.append(pi_doc)
                else:
                    pdfs.append(fitz.open(str(pi_template)))
    
    # 5) 위임인 신분증
    if client_id_bytes:
        client_id = bytes_to_pdf(client_id_bytes)
        if client_id:
            pdfs.append(client_id)
    
    # 6) 수임인 신분증
    if agent_id_bytes:
        agent_id = bytes_to_pdf(agent_id_bytes)
        if agent_id:
            pdfs.append(agent_id)
    
    # 6-1) 팩스 발급: 인감증명서 첨부
    if is_fax and seal_cert_bytes:
        seal_doc = bytes_to_pdf(seal_cert_bytes)
        if seal_doc:
            pdfs.append(seal_doc)
    
    # 7) 번들 타입 고정 첨부 (사업자등록증 등)
    if bundle_config:
        BASE_DIR = Path(__file__).parent.parent
        ID_CARDS_DIR = BASE_DIR / "id_cards"
        
        for fa in bundle_config.get("fixed_attachments", []):
            fa_path = ID_CARDS_DIR / fa["path"]
            if fa_path.exists():
                pdfs.append(fitz.open(str(fa_path)))
        
        # 8) 수임인 첨부 (재직증명서 등 - 날짜 입력)
        for aa in bundle_config.get("agent_attachments", []):
            cert_filename = agent.get(aa.get("staff_key", "cert"), "")
            if cert_filename:
                cert_path = ID_CARDS_DIR / "agents" / cert_filename
                if cert_path.exists():
                    cert_doc = fitz.open(str(cert_path))
                    # 좌표가 있으면 날짜 채우기
                    coords_name = aa.get("coords", "")
                    if coords_name:
                        coords = load_coords(coords_name)
                        if coords:
                            from datetime import datetime, timedelta
                            try:
                                base = datetime.strptime(warrant_date, "%Y.%m.%d")
                            except:
                                base = datetime.now()
                            date_map = {
                                "date_year": base.strftime("%Y"),
                                "date_year_suffix": base.strftime("%y"),
                                "date_month": base.strftime("%m"),
                                "date_day": base.strftime("%d"),
                            }
                            for page_info in coords.get("pages", []):
                                page_num = page_info.get("page", 1) - 1
                                if page_num < cert_doc.page_count:
                                    fill_form_by_coords(cert_doc[page_num], page_info.get("fields", []), date_map)
                    pdfs.append(cert_doc)
    
    return merge_documents(pdfs)
