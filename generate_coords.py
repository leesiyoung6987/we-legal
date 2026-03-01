"""
좌표 초안 자동생성 (generate_coords.py)
──────────────────────────────────────
templates/pdf/ 하위폴더의 PDF를 스캔하여
아직 coords JSON이 없는 양식에 대해 빈 초안을 생성.
coord_tuner에서 필요한 필드만 추가하여 사용.

실행: python generate_coords.py
"""

import json
import fitz
from pathlib import Path

TEMPLATES_DIR = Path("templates")
PDF_DIR = TEMPLATES_DIR / "pdf"
COORDS_DIR = TEMPLATES_DIR / "coords"


def get_doc_type(filename):
    """파일명에서 서류종류 추출 (예: 국민은행_부채증명서.pdf → 부채증명서)"""
    name = Path(filename).stem
    parts = name.split("_", 1)
    if len(parts) >= 2:
        return parts[1]
    return None


def get_bank_name(filename):
    """파일명에서 은행명 추출 (예: 국민은행_부채증명서.pdf → 국민은행)"""
    name = Path(filename).stem
    parts = name.split("_", 1)
    if len(parts) >= 1:
        return parts[0]
    return name


def generate_coords_for_pdf(pdf_path):
    """PDF에 대한 빈 coords JSON 생성 (페이지 수 자동 감지)"""
    form_name = pdf_path.stem
    bank_name = get_bank_name(pdf_path.name)
    doc_type = get_doc_type(pdf_path.name)

    # PDF 페이지 수 확인
    doc = fitz.open(str(pdf_path))
    total_pages = doc.page_count
    doc.close()

    # 페이지별 빈 필드
    pages = []
    for p in range(1, total_pages + 1):
        pages.append({"page": p, "fields": []})

    return {
        "form_name": form_name,
        "bank_name": bank_name,
        "doc_type": doc_type or "",
        "total_pages": total_pages,
        "pages": pages
    }


def main():
    print("=" * 50)
    print("📄 좌표 초안 자동생성 (빈 필드)")
    print("=" * 50)

    created = 0
    skipped = 0

    for category_dir in sorted(PDF_DIR.iterdir()):
        if not category_dir.is_dir():
            continue

        category = category_dir.name
        coords_cat_dir = COORDS_DIR / category
        coords_cat_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n📁 [{category}]")

        for pdf_file in sorted(category_dir.glob("*.pdf")):
            form_name = pdf_file.stem
            coords_path = coords_cat_dir / f"{form_name}_coords.json"

            if coords_path.exists():
                print(f"  ✅ {form_name} (이미 존재)")
                skipped += 1
                continue

            print(f"  🔧 {form_name}")
            coords = generate_coords_for_pdf(pdf_file)

            with open(coords_path, "w", encoding="utf-8") as f:
                json.dump(coords, f, indent=2, ensure_ascii=False)
            print(f"     → 저장: {coords_path} ({coords['total_pages']}페이지)")
            created += 1

    print(f"\n{'=' * 50}")
    print(f"✅ 완료: {created}개 생성, {skipped}개 건너뜀")
    print(f"👉 coord_tuner에서 필요한 필드만 추가하세요!")
    print(f"   streamlit run coord_tuner.py --server.port 8502")


if __name__ == "__main__":
    main()
