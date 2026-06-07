"""전 페이지 footer의 증권번호를 빠짐없이 인덱싱하기 위한 몽타주 생성.

왜 필요한가:
  한 PDF에 증권이 여러 개 들어 있고, 단일특약 증권은 3장처럼 매우 짧을 수 있다.
  '매 N페이지' 같은 표본 스캔이나 '이 증권이 끝까지 간다'는 페이지범위 추정은
  사이에 숨은 짧은 증권을 놓친다(실측 사례: 56쪽 PDF에 증권 4개, 그중 2개가 3쪽짜리).
  따라서 **모든 페이지**의 footer 증권번호를 한 장도 빠짐없이 읽어 맵을 만들어야 한다.

동작:
  extract_pages.py가 만든 workdir/pages.json(각 페이지 이미지)을 읽어, 각 페이지의
  하단 footer(증권번호·'총 N 중 M 장') 띠를 잘라 페이지번호 라벨과 함께 세로로 이어붙인
  몽타주 PNG(기본 28쪽/장)를 만든다. 클로드가 이 몽타주를 view로 읽어
  `페이지 → 증권번호` 전수 맵과 각 증권의 '총 N 중 M 장'을 확정한다.

사용법:
  python index_footers.py --workdir /tmp/cert_xxx [--per 28] [--band 0.94,0.99,0.03,0.45]
출력: workdir/footer_00.png, footer_01.png ...  (그리고 개수 출력)

판독 후 반드시:
  - 모든 페이지가 어떤 증권번호에 매핑되는지(연속성) 확인. 빈 footer는 표지/서명장.
  - 각 증권의 '총 N 중 M 장'으로 시작/끝 페이지를 계산(추정 금지).
  - 엑셀 대상요약의 모든 '숫자 증권번호'가 이 맵에 있는지 대조. 없을 때만 'PDF 없음'.
"""
import argparse
import json
import os

from PIL import Image, ImageDraw


def build(workdir, per, band):
    pages = json.load(open(os.path.join(workdir, "pages.json"), encoding="utf-8"))["pages"]
    y0, y1, x0, x1 = band
    crops = []
    for p in pages:
        ip = p.get("image_path")
        if not ip or not os.path.exists(ip):
            continue
        im = Image.open(ip).convert("RGB")
        w, h = im.size
        c = im.crop((int(w * x0), int(h * y0), int(w * x1), int(h * y1)))
        crops.append((p.get("page"), c))

    outs = []
    for i in range(0, len(crops), per):
        chunk = crops[i:i + per]
        W = max(c.width for _, c in chunk)
        H = sum(c.height for _, c in chunk) + len(chunk) * 4
        out = Image.new("RGB", (W + 90, H), "white")
        dr = ImageDraw.Draw(out)
        y = 0
        for pn, c in chunk:
            out.paste(c, (90, y))
            dr.text((5, y + 6), f"p{pn}", fill="black")
            y += c.height + 4
        fn = os.path.join(workdir, f"footer_{i // per:02d}.png")
        out.save(fn)
        outs.append(fn)
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workdir", required=True)
    ap.add_argument("--per", type=int, default=28, help="몽타주 1장당 페이지 수")
    ap.add_argument("--band", default="0.94,0.99,0.03,0.45",
                    help="footer 띠 비율 y0,y1,x0,x1 (증권번호 위치에 맞게 조정)")
    args = ap.parse_args()
    band = [float(v) for v in args.band.split(",")]
    outs = build(args.workdir, args.per, band)
    print(f"footer 몽타주 {len(outs)}장 생성 -> {workdir_disp(outs)}")
    print("각 몽타주를 view로 읽어 '페이지→증권번호' 전수 맵을 만들고,")
    print("엑셀 대상요약의 모든 숫자 증권번호가 맵에 있는지 대조하세요(없을 때만 'PDF 없음').")


def workdir_disp(outs):
    return os.path.dirname(outs[0]) if outs else "(없음)"


if __name__ == "__main__":
    main()
