"""검증용 엑셀(.xls) 파싱.

두 시트를 읽어 보종코드별 검증 입력 데이터를 JSON으로 출력한다.
- '대상요약'      : 보종코드(B) / 보종명(C) / 증권번호(G) / 담당자(H) / 확인내용(I)
- 상세 시트(PLICD): 보종코드(A) / 라인01(D) / 보장내용01=H(H,idx7) / 보장내용02=I(I,idx8)
                    / 분자(L,idx11) / 분모(M,idx12)

상세 시트는 시트명이 파일마다 다를 수 있으므로 '대상요약'이 아닌 시트를 자동 선택한다.
보장명칭=H열(보장내용01) 라인들을 순서대로 이어붙인 것,
보장내역=I열(보장내용02) 라인들의 목록,
금액라인=I열 텍스트 + 분자/분모(둘 다 0 초과)인 라인.

사용법:
  python parse_excel.py --xls 원본.xls --out parsed.json
"""
import argparse
import json
import xlrd
from verify_lib import to_num


def col(sheet, r, c):
    try:
        return sheet.cell_value(r, c)
    except IndexError:
        return ""


def code_str(v):
    """보종코드 셀값을 정수 문자열로. (예: 84445.0 -> '84445')"""
    if isinstance(v, float):
        return str(int(v))
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def parse(xls_path):
    wb = xlrd.open_workbook(xls_path)
    names = wb.sheet_names()
    detail_name = next(n for n in names if n != "대상요약")

    # --- 대상요약 ---
    s1 = wb.sheet_by_name("대상요약")
    targets = []
    for r in range(4, s1.nrows):  # 데이터는 4행(0-index)부터
        code = code_str(col(s1, r, 1))
        if not code or not code.replace(".", "").isdigit():
            continue
        cert_raw = col(s1, r, 6)
        cert = code_str(cert_raw) if isinstance(cert_raw, float) else str(cert_raw).strip()
        targets.append({
            "보종코드": code,
            "보종명": str(col(s1, r, 2)).strip(),
            "증권번호": cert,
            "담당자": str(col(s1, r, 7)).strip(),
            "확인내용현재": str(col(s1, r, 8)).strip(),
            "xls_row": r,  # 결과 기록 시 매칭용
        })

    # --- 상세 시트 (보종코드별 보장 텍스트/금액) ---
    s2 = wb.sheet_by_name(detail_name)
    coverage = {}
    for r in range(3, s2.nrows):  # 데이터는 3행(0-index)부터
        code = code_str(col(s2, r, 0))
        if not code or not code.isdigit():
            continue
        H = str(col(s2, r, 7)).strip()   # 보장내용01 (보장명칭)
        I = str(col(s2, r, 8)).strip()   # 보장내용02 (보장내역)
        수식 = str(col(s2, r, 9)).strip()  # 보장계산수식구분코드01 (01=비율, 13=절대금액)
        분자 = to_num(col(s2, r, 11))
        분모 = to_num(col(s2, r, 12))
        라인 = col(s2, r, 3)
        cov = coverage.setdefault(code, {
            "상품명": str(col(s2, r, 1)).strip(),
            "명칭라인": [], "내역라인": [], "금액라인": [],
        })
        if H:
            cov["명칭라인"].append(H)
        if I:
            cov["내역라인"].append(I)
        if 분자 > 0 and 분모 > 0:
            cov["금액라인"].append({
                "텍스트": I, "수식구분코드": 수식,
                "분자": int(분자), "분모": int(분모), "라인": str(라인),
            })

    # 명칭/내역 합본 생성
    for code, cov in coverage.items():
        cov["보장명칭"] = "".join(cov["명칭라인"])
        cov["보장내역"] = "\n".join(cov["내역라인"])

    return {"detail_sheet": detail_name, "대상요약": targets, "보장": coverage}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xls", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--code", help="특정 보종코드만 출력(디버그용)")
    args = ap.parse_args()
    data = parse(args.xls)
    if args.code:
        data["대상요약"] = [t for t in data["대상요약"] if t["보종코드"] == args.code]
        data["보장"] = {k: v for k, v in data["보장"].items() if k == args.code}
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"대상요약 {len(data['대상요약'])}행, 보장 {len(data['보장'])}코드 -> {args.out}")


if __name__ == "__main__":
    main()
