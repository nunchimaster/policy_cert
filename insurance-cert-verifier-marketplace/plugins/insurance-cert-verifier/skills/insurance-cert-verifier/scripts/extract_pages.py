"""증권 PDF(실제로는 JPEG 페이지 이미지들을 묶은 ZIP) 페이지 추출.

이 프로젝트의 '증권 PDF'는 표준 PDF가 아니라 페이지별 .jpeg 이미지와
manifest.json 을 담은 ZIP 아카이브다. (pdfplumber 등으로 안 열림)
이 스크립트는 모든 페이지 이미지를 작업폴더로 풀고 페이지 목록을 출력한다.

사용법:
  python extract_pages.py --pdf 증권.pdf --workdir /tmp/cert_xxx
출력: workdir/pages.json  (각 페이지의 이미지 경로/크기)
"""
import argparse
import json
import os
import zipfile


def extract(pdf_path, workdir):
    os.makedirs(workdir, exist_ok=True)
    with zipfile.ZipFile(pdf_path) as z:
        names = z.namelist()
        z.extractall(workdir)
    manifest_path = os.path.join(workdir, "manifest.json")
    pages = []
    if os.path.exists(manifest_path):
        m = json.load(open(manifest_path, encoding="utf-8"))
        for p in m.get("pages", []):
            img = p.get("image", {})
            pages.append({
                "page": p.get("page_number"),
                "image_path": os.path.join(workdir, img.get("path", "")),
                "width": img.get("dimensions", {}).get("width"),
                "height": img.get("dimensions", {}).get("height"),
            })
    else:
        # manifest 없으면 *.jpeg 를 숫자순 정렬
        import re
        jpegs = [n for n in names if n.lower().endswith(".jpeg")]
        jpegs.sort(key=lambda n: int(re.search(r"(\d+)", n).group(1)))
        for i, n in enumerate(jpegs, 1):
            pages.append({"page": i, "image_path": os.path.join(workdir, n)})

    out = {
        "pdf": os.path.basename(pdf_path),
        "num_pages": len(pages),
        "pages": pages,
    }
    with open(os.path.join(workdir, "pages.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"{out['pdf']}: {out['num_pages']}페이지 추출 -> {workdir}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--workdir", required=True)
    args = ap.parse_args()
    extract(args.pdf, args.workdir)


if __name__ == "__main__":
    main()
