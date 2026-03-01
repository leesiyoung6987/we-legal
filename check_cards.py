"""카드사 연동 진단"""
from pathlib import Path
import json

PDF_DIR = Path("templates/pdf")
COORDS_DIR = Path("templates/coords")

# 1. form_mapping 확인
print("=== 1. form_mapping 카드사 등록 ===")
with open("config/form_mapping.json", "r", encoding="utf-8") as f:
    mapping = json.load(f)

cards = ["비씨카드","국민카드","롯데카드","삼성카드","신한카드","하나카드","현대카드"]
for card in cards:
    if card in mapping:
        print(f"✅ {card}: {list(mapping[card].keys())}")
    else:
        print(f"❌ {card}: 미등록")

# 2. PDF 파일 존재 확인
print("\n=== 2. templates/pdf/카드/ 파일 목록 ===")
card_pdf = PDF_DIR / "카드"
if card_pdf.exists():
    for f in sorted(card_pdf.iterdir()):
        print(f"  📄 {f.name}")
else:
    print("  ❌ 카드 폴더 없음")

# 3. coords 파일 존재 확인
print("\n=== 3. templates/coords/카드/ 파일 목록 ===")
card_coords = COORDS_DIR / "카드"
if card_coords.exists():
    for f in sorted(card_coords.iterdir()):
        print(f"  📋 {f.name}")
else:
    print("  ❌ 카드 coords 폴더 없음")

# 4. 경로 매칭 테스트
print("\n=== 4. 경로 매칭 테스트 ===")
for card in cards:
    if card not in mapping: continue
    for doc, val in mapping[card].items():
        items = val if isinstance(val, list) else [val]
        for v in items:
            tpl = PDF_DIR / v["template"]
            crd = COORDS_DIR / f"{v['coords']}_coords.json"
            tpl_ok = "✅" if tpl.exists() else "❌"
            crd_ok = "✅" if crd.exists() else "❌"
            if not tpl.exists() or not crd.exists():
                print(f"  {card} > {doc}")
                print(f"    PDF  {tpl_ok}: {tpl}")
                print(f"    좌표 {crd_ok}: {crd}")

# 5. get_form_info 테스트
print("\n=== 5. get_form_info 매칭 테스트 ===")
import sys
sys.path.insert(0, ".")
from modules.config_loader import get_form_info
tests = [
    ("비씨카드", "카드부채증명서"),
    ("국민카드", "카드부채증명서"),
    ("국민카드", "카드거래내역"),
    ("롯데카드", "카드부채증명서"),
    ("신한카드", "카드거래내역"),
    ("하나카드", "부채증명서"),
]
for name, doc in tests:
    result = get_form_info(name, doc)
    if result:
        print(f"  ✅ {name} + {doc} → {len(result)}개 양식")
    else:
        print(f"  ❌ {name} + {doc} → 매칭 없음")
