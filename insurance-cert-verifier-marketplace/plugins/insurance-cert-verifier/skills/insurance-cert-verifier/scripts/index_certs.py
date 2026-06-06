"""증권번호 → 페이지 범위 1차 인덱싱 (footer 숫자 OCR).

각 페이지 하단의 '(증권번호 : 8025xxxxxx)' / '총 N 중 M 장' 을 OCR로 읽어
증권번호별 페이지 묶음을 제안한다. **OCR은 오인식이 있으므로 1차 후보일 뿐**,
경계 페이지는 클로드가 footer crop 을 직접 보고 확정해야 한다.

증권번호는 8025688xxx 형태(이 상품군 기준)라 정규식으로 후보를 잡는다.

사용법:
  TESSDATA_PREFIX=/path/to/tessdata \
  python index_certs.py --pages /tmp/cert_xxx/pages.json --out /tmp/cert_xxx/index.json
"""
import argparse
import json
import re
import subprocess
from PIL import Image

CERT_RE = re.compile(r"(80256\d{5})")  # 필요시 상품군에 맞게 수정


def ocr_footer(image_path):
    im = Image.open(image_path).convert("RGB")
    w, h = im.size
    crop = im.crop((0, int(h * 0.88), int(w * 0.65), h))
    crop = crop.resize((crop.width * 2, crop.height * 2))
    tmp = "/tmp/_footer_ocr.png"
    crop.save(tmp)
    out = subprocess.run(
        ["tesseract", tmp, "stdout", "--psm", "6", "-l", "eng"],
        capture_output=True, text=True,
    ).stdout
    digits = out.replace(" ", "")
    m = CERT_RE.findall(digits)
    return m[0] if m else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pages", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    pages = json.load(open(args.pages, encoding="utf-8"))
    per_page = []
    for p in pages["pages"]:
        cert = ocr_footer(p["image_path"])
        per_page.append({"page": p["page"], "cert_ocr": cert})

    # 연속 구간으로 증권번호 그룹핑 (빈 값은 직전 값 유지 후보)
    groups = []
    last = None
    for pp in per_page:
        c = pp["cert_ocr"] or last
        if groups and groups[-1]["cert"] == c:
            groups[-1]["pages"].append(pp["page"])
        else:
            groups.append({"cert": c, "pages": [pp["page"]]})
        if pp["cert_ocr"]:
            last = pp["cert_ocr"]

    result = {
        "pdf": pages["pdf"],
        "per_page": per_page,
        "groups": [{"증권번호": g["cert"],
                    "시작": min(g["pages"]), "끝": max(g["pages"]),
                    "pages": g["pages"]} for g in groups if g["cert"]],
        "note": "OCR 1차 후보. 경계 페이지는 footer crop 으로 클로드가 확정할 것.",
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"{result['pdf']} 증권 후보 {len(result['groups'])}건:")
    for g in result["groups"]:
        print(f"  {g['증권번호']}: p{g['시작']}~p{g['끝']}")


if __name__ == "__main__":
    main()
