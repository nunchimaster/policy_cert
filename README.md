# 보험 증권 출력검증 스킬 (insurance-cert-verifier)

검증용 엑셀(`대상요약` + 보장 상세 시트)을 기준(정답)으로 삼아, 보험 증권 PDF에
**보장명칭(H)·보장내역(I)·보장금액**이 올바르게 출력됐는지 점검하고 결과를 엑셀에 기록하는
Claude 스킬입니다. (미래에셋생명 M-케어 3.10.5 / 3N5 간편고지류 증권 검증용으로 제작)

- 매칭 키: **보종코드**(증권 계약사항 표 상품명 앞의 `(84445)` 숫자)
- 금액 공식: 수식구분코드에 따라 **`01`=가입금액×분자/분모(비율)**, **`13`=분자/분모(절대금액)**
- 텍스트 비교: 띄어쓰기·전각/반각·기호 무시 후 비교(`verify_lib.normalize()`) — **보장명칭(H)·보장내역(I) 모두**
- 증권 PDF는 **두 형식 자동 지원** — ① JPEG-ZIP 아카이브, ② 스캔 이미지 표준 PDF(PyMuPDF 렌더)
- 값 판독은 OCR이 아니라 **Claude가 확대 이미지를 직접 보고** 수행(스캔 한글 OCR은 신뢰도 낮음)
- 산출물: **원본 양식을 보존**한 `<원본>_검증.xlsx` — `대상요약` 종합판정/요약컬럼 + **검증상세 시트**(보장명칭·보장내역·보장금액 라인별 O/X/검토 + 사유)

> ⚠️ 본 스킬은 특정 증권 양식에 맞춰져 있습니다. 다른 상품/양식에 쓰려면 `scripts/`의
> 컬럼 인덱스·증권번호 위치(`index_footers.py`의 `--band`)·표 영역 비율(`crop_region.py`)을 조정하세요.

---

## 동작 원리 & 핵심 규칙 (꼭 지킬 것)

1. **증권 PDF는 두 형식.** `extract_pages.py`가 ZIP / 표준 PDF를 자동 감지해 페이지 이미지를 만든다.
2. **증권→페이지 맵은 "전 페이지 footer"로 전수 인덱싱한다.** `index_footers.py`로 모든 페이지의
   footer(증권번호·`총 N 중 M 장`)를 몽타주로 뽑아 **한 장도 빠짐없이** 읽는다.
   - **표본 스캔(매 N페이지)·페이지범위 추정 금지.** 한 PDF에 증권이 여러 개이고, 단일특약 증권은
     **3쪽처럼 짧아** 표본 사이에 숨는다(실측: 56쪽 PDF에 증권 4개, 그중 2개가 3쪽).
   - 각 증권 길이는 footer `총 N 중 M 장`으로 **계산**한다("끝까지 간다"고 단정 금지).
   - **대조 게이트:** 엑셀 대상요약의 모든 숫자 증권번호가 맵에 있는지 확인하고,
     **완전 스캔으로 부재가 확인된 경우에만** `증권 PDF 미제공`(검토)으로 둔다.
3. **매칭은 보종코드.** 계약사항 표에서 `(코드)`로 가입금액과 코드 존재 여부를 판정한다.
4. **금액은 수식구분코드대로.** `01`=가입금액×분자/분모, `13`=분자(절대금액, 가입금액 무관).
   처음 보는 수식코드는 자동계산하지 말고 `검토`.
5. **텍스트는 H와 I를 모두 점검.** 보장명칭(H)만 보고 **보장내역(I, 약관 본문) 점검을 생략하지 않는다.**
6. **결과는 원본 양식을 보존**해 기록한다(셀만 갱신, 서식·열너비·병합 유지).

---

## 스크립트 구성 (`skills/insurance-cert-verifier/scripts/`)

| 스크립트 | 역할 |
|---|---|
| `parse_excel.py` | 검증용 엑셀 파싱(`대상요약` + **여러 상세시트 병합**) → `parsed.json` |
| `extract_pages.py` | 증권 PDF → 페이지 이미지(ZIP 자동해제 / 표준 PDF는 PyMuPDF 렌더) → `pages.json` |
| `index_footers.py` | **전 페이지 footer 증권번호 몽타주** 생성(전수 인덱싱용) |
| `index_certs.py` | (보조) tesseract OCR로 증권번호→페이지 1차 후보 |
| `crop_region.py` | 표 영역(계약사항/보장내용/footer) 확대 크롭 → Claude 시각 판독 |
| `verify_lib.py` | 텍스트 정규화(`normalize`) + 금액 계산(`calc_amount`, 식01/식13) |
| `write_results.py` | 결과를 **원본 양식 보존**하며 기록(.xlsx 인플레이스 / .xls는 변환 후) |

---

## 설치 방법

### 방법 A — Claude Code 플러그인 마켓플레이스 (권장)

```shell
/plugin marketplace add nunchimaster/policy_cert
/plugin install insurance-cert-verifier@policy-cert
```
설치 후 `insurance-cert-verifier:insurance-cert-verifier`로 자동 인식됩니다.
업데이트는 `git push` 후 `/plugin marketplace update`.

### 방법 B — 스킬 폴더 직접 복사

```bash
cp -r plugins/insurance-cert-verifier/skills/insurance-cert-verifier ~/.claude/skills/      # 개인용
cp -r plugins/insurance-cert-verifier/skills/insurance-cert-verifier .claude/skills/         # 프로젝트 공유용
```

### 방법 C — Claude.ai / Cowork (.skill 업로드)

스킬 폴더를 zip으로 묶어 커스텀 스킬 업로드가 지원되는 곳에 올립니다.

---

## 사전 준비 (실행 환경)

```bash
pip install xlrd openpyxl pillow pymupdf
# (선택) .xls 입력의 양식 보존 변환용: LibreOffice(soffice) 또는 Windows의 Excel
# (보조) 한글 OCR — index_certs.py 사용 시에만:
#   apt install tesseract-ocr
#   curl -sL -o kor.traineddata https://raw.githubusercontent.com/tesseract-ocr/tessdata/main/kor.traineddata
```

- 표준 PDF 렌더링에는 **PyMuPDF(`pymupdf`)** 가 필요합니다.
- 원본이 `.xls`면 `write_results.py`가 **LibreOffice/Excel로 .xlsx 변환 후** 양식을 보존해 갱신합니다(변환 수단이 없으면 값만 복제하는 폴백, 서식 손실 경고).
- OCR은 보조일 뿐, 정확한 값 판독은 Claude의 시각 판독으로 이뤄집니다.

---

## 사용 예시 (요청 문구)

- "이 검증용 엑셀 기준으로 증권 PDF에 보장내용/보장금액이 제대로 출력됐는지 확인해줘"
- "대상요약 보종코드별로 증권 대조해서 확인내용 채워줘"

자세한 절차·결과 스키마는 `skills/insurance-cert-verifier/SKILL.md`를 참고하세요.

---

## 산출물 형식

`<원본>_검증.xlsx` (원본 서식 보존):
- **대상요약**: I열=종합 확인내용, 끝에 `검증PDF/검증페이지/텍스트판정/금액판정/종합사유`
- **검증상세(신규 시트)**: 보종코드별 `보장명칭`·`보장내역`·`보장금액` 라인 — `엑셀값/증권출력값/결과/사유`
- 셀 색상: O=연두, X=연빨강, 검토=노랑
- 종합판정: 명칭·내역·금액 모두 O면 `이상없음`, 하나라도 X면 `X: …`, 못 읽으면 `검토: …`

> 결과물은 **사람 검토를 전제**로 합니다. 자동 판정 불가 항목은 위조하지 않고 `검토`로 남깁니다.

---

## 레포 구조

```
.
├── .claude-plugin/marketplace.json     # 마켓플레이스 카탈로그
├── plugins/insurance-cert-verifier/
│   ├── .claude-plugin/plugin.json      # 플러그인 매니페스트
│   └── skills/insurance-cert-verifier/
│       ├── SKILL.md                    # 스킬 본문(절차)
│       └── scripts/                    # 파싱·추출·인덱싱·계산·기록 스크립트
└── README.md
```

## 라이선스

MIT (LICENSE 파일 참조).
