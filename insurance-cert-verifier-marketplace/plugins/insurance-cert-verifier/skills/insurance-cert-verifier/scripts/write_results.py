"""검증 결과를 엑셀에 기록.

입력 .xls(원본)를 .xlsx로 복제(값 보존)하고 다음을 추가/갱신한다.
- '대상요약' 시트:
    I열(확인내용)  = 종합판정 (이상없음 / X:... / 검토:... / 해당 증권에 보종코드 없음)
    M~Q열(끝에 추가): 검증PDF, 검증페이지, 텍스트판정(O/X/검토), 금액판정(O/X/검토), 종합사유
  * 매칭은 보종코드(B열) 기준.
- '검증상세' 시트(신규): 라인별 1행
    보종코드 / 보종명 / 증권번호 / PDF파일 / 페이지 / 항목구분 / 엑셀값 / 증권출력값 / 결과 / 사유

결과 JSON 스키마(리스트):
[{
  "보종코드","증권번호","보종명","검증PDF","검증페이지",
  "텍스트판정","금액판정","확인내용","종합사유",
  "lines":[{"항목구분","엑셀값","증권출력값","결과","사유","페이지"}]
}]

사용법:
  python write_results.py --xls 원본.xls --results results.json --out 회신_검증.xlsx
"""
import argparse
import json
import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

RESULT_FILL = {
    "O":  PatternFill("solid", fgColor="C6EFCE"),   # 연두
    "X":  PatternFill("solid", fgColor="FFC7CE"),   # 연빨강
    "검토": PatternFill("solid", fgColor="FFEB9C"),  # 노랑
}
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
HDR_FONT = Font(bold=True)


def copy_sheet(ws, xls_sheet):
    for r in range(xls_sheet.nrows):
        for c in range(xls_sheet.ncols):
            v = xls_sheet.cell_value(r, c)
            if v != "":
                ws.cell(row=r + 1, column=c + 1, value=v)


def code_str(v):
    if isinstance(v, float):
        return str(int(v))
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xls", required=True)
    ap.add_argument("--results", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    wb_in = xlrd.open_workbook(args.xls)
    results = json.load(open(args.results, encoding="utf-8"))
    by_code = {r["보종코드"]: r for r in results}

    wb = Workbook()
    wb.remove(wb.active)

    # --- 원본 시트들 복제 ---
    sheet_ws = {}
    for name in wb_in.sheet_names():
        ws = wb.create_sheet(title=name[:31])
        copy_sheet(ws, wb_in.sheet_by_name(name))
        sheet_ws[name] = ws

    # --- 대상요약 갱신 ---
    s1 = wb_in.sheet_by_name("대상요약")
    ws1 = sheet_ws["대상요약"]
    HDR = 4  # 헤더행(1-index). xlrd 3행 == openpyxl 4행
    # 요약 컬럼 헤더 (M=13 ~ Q=17)
    summary_cols = [(13, "검증PDF"), (14, "검증페이지"), (15, "텍스트판정"),
                    (16, "금액판정"), (17, "종합사유")]
    for c, title in summary_cols:
        cell = ws1.cell(row=HDR, column=c, value=title)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT

    for r in range(4, s1.nrows):  # xlrd 데이터행
        code = code_str(s1.cell_value(r, 1))
        res = by_code.get(code)
        if not res:
            continue
        row = r + 1  # openpyxl
        ws1.cell(row=row, column=9, value=res.get("확인내용", ""))  # I열
        ws1.cell(row=row, column=13, value=res.get("검증PDF", ""))
        ws1.cell(row=row, column=14, value=res.get("검증페이지", ""))
        for c, key in [(15, "텍스트판정"), (16, "금액판정")]:
            v = res.get(key, "")
            cell = ws1.cell(row=row, column=c, value=v)
            if v in RESULT_FILL:
                cell.fill = RESULT_FILL[v]
        ws1.cell(row=row, column=17, value=res.get("종합사유", ""))

    # --- 검증상세 시트 ---
    ws2 = wb.create_sheet(title="검증상세")
    headers = ["보종코드", "보종명", "증권번호", "PDF파일", "페이지",
               "항목구분", "엑셀값", "증권출력값", "결과", "사유"]
    for c, h in enumerate(headers, 1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
    widths = [10, 34, 13, 30, 8, 26, 46, 22, 6, 40]
    for c, wdt in enumerate(widths, 1):
        ws2.column_dimensions[chr(64 + c)].width = wdt

    row = 2
    for res in results:
        for ln in res.get("lines", []):
            vals = [res.get("보종코드"), res.get("보종명"), res.get("증권번호"),
                    res.get("검증PDF"), ln.get("페이지"),
                    ln.get("항목구분"), ln.get("엑셀값"), ln.get("증권출력값"),
                    ln.get("결과"), ln.get("사유")]
            for c, v in enumerate(vals, 1):
                cell = ws2.cell(row=row, column=c, value=v)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            rcell = ws2.cell(row=row, column=9)
            if ln.get("결과") in RESULT_FILL:
                rcell.fill = RESULT_FILL[ln["결과"]]
            row += 1

    wb.save(args.out)
    print(f"기록 완료: {args.out}  (검증상세 {row-2}행)")


if __name__ == "__main__":
    main()
