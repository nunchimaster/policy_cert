"""페이지 이미지의 특정 영역을 잘라 확대 저장 (클로드가 표를 또렷이 보고 판독하기 위함).

tesseract OCR은 이 스캔 품질에서 신뢰도가 낮으므로, 값 판독은 클로드가
crop 이미지를 view 도구로 직접 보고 한다. 이 스크립트는 그 보조 도구.

영역(--region):
  full     전체
  contract 계약사항/보험가입내용 표 (페이지 하단 ~55%~100%) — 가입금액·보종코드
  coverage 가입상품 보장내용 표 (페이지 ~50%~100%) — 보장명칭/보장내역/보장금액
  footer   증권번호·'총 N 중 M 장' (하단 ~88%~100%)
  y0,y1    세로 비율 직접 지정 (예: 0.5,0.8). x0,y0,x1,y1 4값도 가능.

사용법:
  python crop_region.py --image p2.jpeg --region contract --out crop.png
"""
import argparse
from PIL import Image

PRESETS = {
    "full":     (0.0, 0.0, 1.0, 1.0),
    "contract": (0.0, 0.52, 1.0, 1.0),
    "coverage": (0.0, 0.48, 1.0, 1.0),
    "footer":   (0.0, 0.88, 0.65, 1.0),
    "top":      (0.0, 0.0, 1.0, 0.5),
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--image", required=True)
    ap.add_argument("--region", default="full")
    ap.add_argument("--scale", type=float, default=1.8)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    im = Image.open(args.image).convert("RGB")
    w, h = im.size
    if args.region in PRESETS:
        x0, y0, x1, y1 = PRESETS[args.region]
    else:
        vals = [float(v) for v in args.region.split(",")]
        if len(vals) == 2:
            x0, y0, x1, y1 = 0.0, vals[0], 1.0, vals[1]
        elif len(vals) == 4:
            x0, y0, x1, y1 = vals
        else:
            raise SystemExit("--region 은 preset, 'y0,y1', 또는 'x0,y0,x1,y1'")

    box = (int(w * x0), int(h * y0), int(w * x1), int(h * y1))
    crop = im.crop(box)
    if args.scale != 1.0:
        crop = crop.resize((int(crop.width * args.scale), int(crop.height * args.scale)))
    crop.save(args.out)
    print(f"{args.image} {args.region} -> {args.out} ({crop.size[0]}x{crop.size[1]})")


if __name__ == "__main__":
    main()
