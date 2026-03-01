"""은행권 연동 진단"""
import sys, json
from pathlib import Path
sys.path.insert(0, ".")

# 1. form_mapping 확인
print("=== form_mapping ===")
with open("config/form_mapping.json", "r", encoding="utf-8") as f:
    mapping = json.load(f)
for bank in ["에스씨제일은행", "광주은행"]:
    print(f"  {bank}: {mapping.get(bank, '미등록')}")

# 2. 파일 탐색
print("\n=== PDF 탐색 ===")
pdf_dir = Path("templates/pdf")
for p in pdf_dir.rglob("*.pdf"):
    if "에스씨" in str(p) or "광주" in str(p):
        print(f"  발견: {p}")

print("\n=== coords 탐색 ===")
coords_dir = Path("templates/coords")
for p in coords_dir.rglob("*_coords.json"):
    if "에스씨" in str(p) or "광주" in str(p):
        print(f"  발견: {p}")

# 3. 실제 로드 테스트
print("\n=== load_coords / get_template_path 테스트 ===")
from modules.config_loader import load_coords, get_template_path

for template, coords_name in [
    ("은행/에스씨제일은행_통장거래내역.pdf", "은행/에스씨제일은행_통장거래내역"),
    ("은행/광주은행_부채증명서.pdf", "은행/광주은행_부채증명서"),
    ("에스씨제일은행_통장거래내역.pdf", "에스씨제일은행_통장거래내역"),
]:
    tpl = get_template_path(template)
    crd = load_coords(coords_name)
    print(f"  template='{template}' → exists={tpl.exists() if tpl else 'None'}")
    print(f"  coords='{coords_name}' → loaded={'✅' if crd else '❌'}")

# 4. get_form_info 테스트
print("\n=== get_form_info 테스트 ===")
from modules.config_loader import get_form_info
for bank, doc in [("에스씨제일은행","통장거래내역"), ("광주은행","부채증명서")]:
    result = get_form_info(bank, doc)
    print(f"  {bank} + {doc} → {result}")
