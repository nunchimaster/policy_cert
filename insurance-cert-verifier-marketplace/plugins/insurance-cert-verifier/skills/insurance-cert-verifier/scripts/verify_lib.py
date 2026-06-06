"""공용 유틸: 텍스트 정규화 + 보장금액 계산.

비교 규칙(사용자 확정):
- 띄어쓰기 무시 (일반 공백 + 전각 공백 \u3000 포함)
- 전각/반각 차이 무시 (NFKC 정규화)
- 괄호/꺾쇠/따옴표/쉼표 등 모든 기호 무시
- 그 외 글자(한글/영문/숫자)가 다르면 '불일치(X)'
"""
import re
import unicodedata


def normalize(s) -> str:
    """비교용 정규화. 한글/영문/숫자만 남기고 공백·기호·전각차이를 제거한다."""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s))
    # 한글/영문/숫자만 유지 (공백·괄호·꺾쇠·따옴표·쉼표·점 등 전부 제거)
    t = re.sub(r"[^0-9A-Za-z\uac00-\ud7a3]", "", t)
    return t


def text_match(excel_val, pdf_val):
    """정규화 후 일치 여부. (일치여부, 정규화엑셀, 정규화PDF) 반환."""
    a, b = normalize(excel_val), normalize(pdf_val)
    return (a == b), a, b


def calc_amount(가입금액, 분자, 분모, 수식구분코드="01"):
    """보장금액 계산. 수식구분코드(SCRT_CACL_FRML_DCD)에 따라 공식이 다르다.

    - '01' (비율): 보장금액 = 가입금액 × 분자 / 분모
        예) 가입금액 1,000,000, 100/100 -> 1,000,000 / 50/100 -> 500,000
    - '13' (절대금액): 보장금액 = 분자 / 분모 (가입금액과 무관, 분자 자체가 원 단위 금액)
        예) 50000/1 -> 50,000 (MRI촬영 급여)

    연산자(SCRT_CACL_OPTR01)는 데이터상 항상 곱셈('*').
    새로운 수식구분코드가 나오면 '검토'로 처리하고 사람이 확인하도록 한다.
    """
    분자 = float(분자)
    분모 = float(분모)
    if 분모 == 0:
        return None
    code = str(수식구분코드).strip()
    if code == "13":
        return int(round(분자 / 분모))
    # '01' 및 기본
    return int(round(float(가입금액) * 분자 / 분모))


KNOWN_FORMULA_CODES = {"01", "13"}


def to_num(x):
    """문자열/실수 셀값을 float으로. 실패 시 0.0."""
    try:
        return float(str(x).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def parse_money(s):
    """'1,000,000' / '1.000.000' / '1000000' 같은 표기를 정수로. 실패 시 None."""
    if s is None:
        return None
    digits = re.sub(r"[^0-9]", "", str(s))
    return int(digits) if digits else None


if __name__ == "__main__":
    # 간단 자가 테스트
    assert normalize("＜비급여（전액본인부담　포함）＞") == normalize("<비급여(전액본인부담 포함)>")
    assert calc_amount(1000000, 100, 100) == 1000000
    assert calc_amount(1000000, 50, 100) == 500000
    assert parse_money("1,000,000") == 1000000
    assert parse_money("1.000.000") == 1000000
    print("verify_lib 자가테스트 통과")
