"""증권 PDF 페이지 추출 (두 가지 입력 형식 자동 감지).

지원 형식:
1) **JPEG-ZIP 아카이브** — 페이지별 `.jpeg` 이미지 + `manifest.json` 을 담은 ZIP.
   (일부 다운로드본. pdfplumber 등으로는 안 열림 → ZIP 으로 푼다.)
2) **표준 PDF** (`%PDF-...`, 스캔 이미지 내장) — PyMuPDF(fitz)로 각 페이지를
   이미지로 렌더링한다. (텍스트 레이어가 없는 스캔본도 이미지로 떨궈 판독 가능.)

두 경로 모두 `workdir/pages.json` 을 **동일 스키마**로 출력하므로
이후 단계(index_certs.py / crop_region.py / 클로드 판독)는 입력 형식과 무관하게 동작한다.

사용법:
  python extract_pages.py --pdf 증권.pdf --workdir /tmp/cert_xxx [--zoom 3.0]

옵션:
  --zoom  표준 PDF 렌더 배율(기본 3.0 ≈ 216DPI). 글자가 흐리면 3.5~4.0 으로 올린다.
출력: workdir/pages.json  (각 페이지의 image_path/width/height, 그리고 format)
"""
import argparse
import json
import os
import re
import zipfile


def _pages_from_zip(pdf_path, workdir):
    """JPEG-ZIP 아카이브에서 페이지 목록을 만든다."""
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
        jpegs = [n for n in names if n.lower().endswith((".jpeg", ".jpg"))]
        jpegs.sort(key=lambda n: int(re.search(r"(\d+)", n).group(1)))
        for i, n in enumerate(jpegs, 1):
            pages.append({"page": i, "image_path": os.path.join(workdir, n)})
    return pages


def _pages_from_pdf(pdf_path, workdir, zoom):
    """표준 PDF를 PyMuPDF로 페이지별 이미지로 렌더링한다."""
    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise SystemExit(
            "표준 PDF 렌더링에는 PyMuPDF가 필요합니다: "
            "pip install pymupdf  (오류: %s)" % e
        )
    doc = fitz.open(pdf_path)
    mat = fitz.Matrix(zoom, zoom)
    pages = []
    for i in range(doc.page_count):
        pix = doc[i].get_pixmap(matrix=mat)
        out = os.path.join(workdir, "p%04d.png" % (i + 1))
        pix.save(out)
        pages.append({
            "page": i + 1,
            "image_path": out,
            "width": pix.width,
            "height": pix.height,
        })
    doc.close()
    return pages


def extract(pdf_path, workdir, zoom=3.0):
    os.makedirs(workdir, exist_ok=True)
    # 형식 자동 감지: 진짜 ZIP이면 ZIP 경로, 아니면(=표준 PDF 등) 렌더 경로.
    if zipfile.is_zipfile(pdf_path):
        pages, fmt = _pages_from_zip(pdf_path, workdir), "zip-jpeg"
    else:
        pages, fmt = _pages_from_pdf(pdf_path, workdir, zoom), "pdf-render"

    out = {
        "pdf": os.path.basename(pdf_path),
        "format": fmt,
        "num_pages": len(pages),
        "pages": pages,
    }
    with open(os.path.join(workdir, "pages.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"{out['pdf']}: {out['num_pages']}페이지 추출({fmt}) -> {workdir}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--workdir", required=True)
    ap.add_argument("--zoom", type=float, default=3.0,
                    help="표준 PDF 렌더 배율(기본 3.0). ZIP 입력에는 무시됨.")
    args = ap.parse_args()
    extract(args.pdf, args.workdir, args.zoom)


if __name__ == "__main__":
    main()
