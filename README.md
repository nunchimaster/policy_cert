# 보험 증권 출력검증 스킬 (insurance-cert-verifier)

검증용 엑셀(`대상요약` + 보장 상세 시트)을 기준(정답)으로 삼아, 보험 증권 PDF에
**보장명칭·보장내역·보장금액**이 올바르게 출력됐는지 점검하고 결과를 엑셀에 기록하는
Claude 스킬입니다. (미래에셋생명 M-케어 3.10.5 간편고지류 증권 검증용으로 제작)

- 매칭 키: **보종코드**(증권 계약사항 표 상품명 앞의 `(84445)` 숫자)
- 금액 공식: 수식구분코드에 따라 **`01`=가입금액×분자/분모(비율)**, **`13`=분자/분모(절대금액)**
- 텍스트 비교: 띄어쓰기·전각/반각·기호 무시 후 비교
- 증권 PDF가 텍스트 없는 스캔 이미지(JPEG ZIP)라, 값 판독은 Claude가 페이지를 직접 보고 수행
- 산출물: `대상요약` I열(종합판정) + 요약 컬럼 + **검증상세 시트**(라인별 O/X/검토 + 사유)

> ⚠️ 본 스킬은 특정 증권 양식에 맞춰져 있습니다. 다른 상품/양식에 쓰려면 `scripts/`의
> 컬럼 인덱스·증권번호 정규식(`index_certs.py`의 `CERT_RE`)·표 영역 비율(`crop_region.py`)을
> 조정하세요.

---

## 설치 방법

### 방법 A — Claude Code 플러그인 마켓플레이스 (권장)

이 레포를 GitHub에 올린 뒤, 사용자는 Claude Code에서:

```shell
# 1) 마켓플레이스 등록 (당신의 깃허브 계정/레포명으로)
/plugin marketplace add nunchimaster/policy_cert

# 2) 플러그인 설치
/plugin install insurance-cert-verifier@policy-cert
```

설치 후 스킬은 `insurance-cert-verifier:insurance-cert-verifier` 이름으로 자동 인식되어,
"검증용 엑셀로 증권 출력 점검" 같은 요청에 동작합니다.

업데이트는 `git push` 후 사용자가 `/plugin marketplace update` 하면 됩니다.
(`marketplace.json` / `plugin.json`의 `version`을 올리거나, 생략하면 커밋 SHA가 버전이 됩니다.)

### 방법 B — 스킬 폴더 직접 복사 (마켓플레이스 없이)

```bash
# 개인용 (모든 프로젝트에서 사용)
cp -r plugins/insurance-cert-verifier/skills/insurance-cert-verifier ~/.claude/skills/

# 특정 프로젝트 공유용
cp -r plugins/insurance-cert-verifier/skills/insurance-cert-verifier .claude/skills/
```

### 방법 C — Claude.ai / Cowork (.skill 업로드)

`.skill` 파일(이 스킬 폴더를 zip으로 묶은 것)을 커스텀 스킬 업로드가 지원되는 곳에서
업로드합니다. 지원 여부·위치는 https://support.claude.com 를 참고하세요.

---

## 사전 준비 (실행 환경)

스킬 스크립트는 Python과 tesseract OCR을 사용합니다.

```bash
pip install xlrd openpyxl pillow
# (보조용) 한글 OCR 데이터
mkdir -p ~/tessdata && cd ~/tessdata
curl -sL -o kor.traineddata https://raw.githubusercontent.com/tesseract-ocr/tessdata/main/kor.traineddata
export TESSDATA_PREFIX=~/tessdata
```

- `tesseract` 실행파일이 필요합니다 (`apt install tesseract-ocr` 등).
- OCR은 보조 수단입니다. 정확한 값 판독은 Claude의 시각 판독으로 이뤄집니다.

---

## 사용 예시 (요청 문구)

- "이 검증용 엑셀 기준으로 증권 PDF에 보장내용/보장금액이 제대로 출력됐는지 확인해줘"
- "대상요약 보종코드별로 증권 대조해서 확인내용 채워줘"

자세한 절차·결과 스키마는 스킬 본문(`skills/insurance-cert-verifier/SKILL.md`)을 참고하세요.

---

## 레포 구조

```
.
├── .claude-plugin/
│   └── marketplace.json          # 마켓플레이스 카탈로그
├── plugins/
│   └── insurance-cert-verifier/
│       ├── .claude-plugin/
│       │   └── plugin.json        # 플러그인 매니페스트
│       └── skills/
│           └── insurance-cert-verifier/
│               ├── SKILL.md       # 스킬 본문(절차)
│               └── scripts/       # 파싱·추출·계산·기록 스크립트
└── README.md
```

## 라이선스

MIT (LICENSE 파일 참조). `nunchimaster` /  자리표시자를 본인 정보로 바꾸세요.
