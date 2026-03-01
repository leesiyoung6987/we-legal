"""
form_mapping 자동 생성 (update_form_mapping.py)
──────────────────────────────────────────────
templates/coords/ 폴더의 JSON 파일을 스캔하여
config/form_mapping.json을 자동 업데이트.

실행: python update_form_mapping.py
"""

import json
from pathlib import Path

COORDS_DIR = Path("templates/coords")
PDF_DIR = Path("templates/pdf")
MAPPING_PATH = Path("config/form_mapping.json")

def main():
    # 기존 매핑 로드
    if MAPPING_PATH.exists():
        with open(MAPPING_PATH, "r", encoding="utf-8") as f:
            mapping = json.load(f)
    else:
        mapping = {}

    # _comment 보존
    comment = mapping.pop("_comment", "채권사명 → 서류종류 → {template: PDF파일명, coords: 좌표JSON명}")

    added = 0
    skipped = 0

    # coords 폴더 스캔
    for coords_file in sorted(COORDS_DIR.rglob("*_coords.json")):
        form_name = coords_file.stem.replace("_coords", "")
        
        # 하위폴더 경로 계산 (예: 카드/비씨카드_부채증명서)
        rel_dir = coords_file.parent.relative_to(COORDS_DIR)
        if str(rel_dir) == ".":
            form_path_prefix = ""
        else:
            form_path_prefix = str(rel_dir) + "/"
        
        # 파일명에서 채권사명, 서류종류 분리
        parts = form_name.split("_", 1)
        if len(parts) < 2:
            print(f"  ⚠️ 파일명 형식 오류 (채권사_서류종류 아님): {form_name}")
            continue
        
        bank_name = parts[0]
        doc_type = parts[1]
        
        # PDF 파일 존재 확인
        pdf_found = False
        for pdf in PDF_DIR.rglob(f"{form_name}.pdf"):
            pdf_found = True
            break
        
        if not pdf_found:
            print(f"  ⚠️ PDF 없음: {form_name}.pdf")
            continue
        
        # 이미 등록되어 있으면 건너뛰기
        if bank_name in mapping and doc_type in mapping[bank_name]:
            print(f"  ✅ 이미 등록: {bank_name} > {doc_type}")
            skipped += 1
            continue
        
        # 매핑 추가 (하위폴더 경로 포함)
        if bank_name not in mapping:
            mapping[bank_name] = {}
        
        mapping[bank_name][doc_type] = {
            "template": f"{form_path_prefix}{form_name}.pdf",
            "coords": f"{form_path_prefix}{form_name}"
        }
        print(f"  ➕ 추가: {bank_name} > {doc_type} ({form_path_prefix})")
        added += 1

    # 저장
    output = {"_comment": comment}
    # 가나다순 정렬
    for bank in sorted(mapping.keys()):
        output[bank] = {}
        for doc in sorted(mapping[bank].keys()):
            output[bank][doc] = mapping[bank][doc]

    with open(MAPPING_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*50}")
    print(f"✅ 완료: {added}개 추가, {skipped}개 건너뜀")
    print(f"📄 저장: {MAPPING_PATH}")
    print(f"📋 총 등록: {sum(len(v) for v in mapping.values())}개 양식")

if __name__ == "__main__":
    main()
