---
name: ss-report-light
description: >
  한국어 내부 보고서를 정형화된 디자인 템플릿으로 생성하는 경량 전용 스킬.
  투자 검토, 포트폴리오 현황, 운영안 보고, 진행 현황, 설득/제안 보고 등 모든 내부 보고서에 적용.
  트리거: "보고서 만들어줘", "보고서 작성해줘", "워드로 뽑아줘", "docx 보고서", "내부 보고서",
           "검토 보고서", "현황 보고서", "기획 보고서", "제안 보고서", "ss_report", "보고서 양식",
           "보고서 템플릿", "보고서 써줘", "보고서로", "보고서 형태로", "보고서 형식으로" 등.
  반드시 이 스킬을 참고할 것. 일반 docx 스킬 대신 이 ss-report-light 스킬을 최우선 사용.
  사용자가 보고서 작성을 명시하거나 요청하는 모든 경우 이 스킬 적용. docx 스킬 대체 절대 금지.
---

# ss-report — 범용 내부 보고서 스킬

스크립트: `scripts/patch_annex.js`, `scripts/verify_report.py`
references: `references/rules-layout.md`, `references/rules-symbol.md`, `references/code-helpers.md`, `references/code-helpers-annex.md`

**2단계 flow**: ① 목차+페이지계획 기획 (사용자 1회 확인) → ② JS 코드 생성 + 검증·레이아웃 정제

**빠른 실행 예외**: 사용자 요청에 "빠른 실행", "바로 작성", "초안", "테스트", "스킵" 등 확인 생략 의도가 명확하면 목차·페이지 확인을 내부 판단으로 대체하고 바로 생성한다.

> **⚠️ 날짜 기준 (STRICT)**: datePara() 작성일, 기사·데이터 검색 기준일은 반드시 시스템 currentDate(오늘 실제 날짜) 확인 후 사용. 임의 날짜·기억 속 날짜·이전 대화 날짜 재사용 금지.

---

## FLOW 1. 목차 + 페이지 계획 기획 (1턴 통합)

쿼리와 데이터를 분석해 목차와 페이지 계획을 **동시에** 제안. 사용자 확인 1회로 진행.

**제안 형식:**
```
[목차 + 페이지 계획]
구성 이유: (1~2줄)

1. 섹션명 / 2. 섹션명 / 3. 섹션명  (별첨. 별첨명 — 필요 시)

| 페이지 | 주요 내용 | 예상 단락 수 | 표 | 3줄 케이스 | 사전 조치 |
|--------|---------|------------|-----|-----------|---------|
| P1     | H1+H2+body | 약 N    | 없음 | 없음     | —       |
| P2     | H1+표(N행) | 약 N+표 | 있음 | A: 잔량↓  | 직전 조정 |

→ 이 구성으로 진행할까요?
```

**원칙:** lv1 기준 3~5개 섹션 / 3줄 케이스·계층 고립 사전 탐지 후 조치 계획 포함
**빠른 실행 예외:** "빠른 실행", "바로", "초안", "스킵" 등 → 내부 확정 후 바로 코드 생성

---

## FLOW 2. JS 코드 생성

### 별첨 유무에 따른 코드 분기 (STRICT)
```
별첨 없음: Packer.toBuffer() + &apos; 복원 후 저장. AdmZip·annexLabelXml·buildAndPatch 생성 금지.
           파이프라인: node report.js → verify
별첨 있음: patch_annex.js 필수. report.js에 AdmZip/buildAndPatch 작성 금지.
           파이프라인: node report.js → patch_annex → verify
```

### 페이지 계획표 (FLOW 1에서 확인 완료 후 JS 시작)

```
3줄 케이스 A: 표 잔량 3줄 이하 → 표 전체 다음 페이지로
3줄 케이스 B: 직전 내용 3줄만 남음 → 직전 분량 조정
계층 고립: lv1·lv2가 페이지 끝에 단독 위치 → 직전 분량 조정
→ 탐지된 케이스 조치 완료 후 JS 코드 작성 시작
```

### 풋노트 필요성 계획 (내부 처리)

페이지 계획 뒤에 풋노트 후보를 내부적으로 정한다. 기준은 위치가 아니라 **필요성**이다. 보충설명이 필요한 줄이면 삽입하고, 맞지 않으면 삽입하지 않는다.
풋노트 후보표는 사용자에게 별도로 보여주지 않는다.

| 후보 | 설명 대상 | 필요성 판단 | 배치 위치 |
|------|----------|------------|----------|
| 1 | 약어·전문용어 | 첫 등장이고 독자 오해 가능 | 해당 문장 바로 아래 |
| 2 | 수치·출처·기준시점 | 출처·기준 없으면 해석이 흔들림 | 해당 수치가 있는 본문/표 바로 아래 |
| 3 | 단위·산식·집계 기준 | 표 또는 문장 해석에 직접 필요 | 해당 표/문장 바로 아래 |
| 4 | 부연 설명 | 본문에 넣으면 줄 길이·흐름을 해침 | 설명 대상 바로 아래 |

금지: 위치를 맞추기 위해 억지로 풋노트를 추가하거나 옮기기. 설명 대상과 풋노트 사이에 빈줄을 넣지 않는다.

풋노트 `*` 위치: 줄 전체 설명은 직전 줄의 **기호 시작점+2칸**에 정렬한다. 특정 단어·수치·표현 근방에 붙여야 할 때만 `_isFn: { anchor:"표현", text:"설명" }`을 사용한다. `anchor`는 위치 계산용이라 최종 출력에는 보이지 않고 `* 설명`만 표시된다.

### 문서 구조

```
lv1 (H1):  1. 제목      lv2 (H2): □ 소제목
lv3 (본문): - 항목      lv4 (소항목): · 세부
번호식:     ① ② ③ 또는 ※ — lv3/lv4 대체 기호로만 사용
별첨:       본문 다음 페이지에 배치 (별첨 필요 시에만 → references/code-helpers-annex.md 참조)
```

### JS 헬퍼 핵심 상수 (함수 전문 → references/code-helpers.md 참조)

```javascript
const F      = { name: "바탕체", cs: "바탕체" };  // "Batang"·"바탕" 절대 금지
const BLACK  = "000000";
const BLUE   = "0000FF";
const TW     = 8900;   // 표 너비 DXA
const L2 = 142; const L3 = 568; const L4 = 710; const HANG = 142;  // 기호 기준 들여쓰기(1칸=0.25cm)
const SP = {
  titleAfter:120, dateAfter:400,
  // toH_N = "다음 단락이 HN일 때" 현재 단락에 적용할 after (buildParagraphs 자동 처리)
  // 예: 어떤 단락 다음이 H1이면 → 그 단락의 after = toH1(480=24pt)
  toH1:480, toH2:400, toH3:240, toH4:180, cont:120,
  fnBefore:20, cellV:60,
};
// 헬퍼 함수: t / tb / tu / tcs / titlePara / datePara / h1 / h2 / body / bodyCont / sub / sub2
//            num3 / num4 / fn / tableGap / annexTitlePara / h1Sep
// → 전문 코드: references/code-helpers.md STEP 3
```

### report.js 필수 섹션 순서 (skeleton)

```javascript
// ── SECTION 1: require + 상수 (F, BLACK, BLUE, TW, L2/L3/L4, SP) ────────
// ── SECTION 2: 단락 헬퍼 (t/tb/tu/tcs, h1/h2/body/sub/fn 등) ─────────────
// ── SECTION 3: 표 헬퍼 (C/CL/CLines/mkRow/mkTable 등) ────────────────────
// ── SECTION 4: defs 배열 (본문 단락 정의) ────────────────────────────────
// ── SECTION 5: 별첨 children (별첨 있을 때만 code-helpers-annex.md 참조) ───
//              별첨 없을 경우 이 섹션 생략 — 별첨 reference 읽지 않음
// ── SECTION 6: Document 생성 + buildAndPatch + main() ────────────────────
```

**규칙**: SECTION 2·3 헬퍼 함수는 여기서만 정의 — SECTION 4 defs 배열 내 인라인 정의 금지

### 표 헬퍼 핵심 시그니처 (전문 → references/code-helpers.md 참조)

```javascript
// C(text, width, isHdr=false, sz=24)   — 가운데 정렬 셀 (기본)
// CL(text, width, isHdr=false, sz=24)  — 왼쪽 정렬 셀 (설명·비고 등 장문 컬럼만)
// CP(text, width)                      — C(text, width, false, 20) 단축
// CLines(lines, width, sz=24)          — 명시적 개행 셀
// mkRow(cells) / mkHdrRow(cells) / mkTable(hdrRow, dataRows, colWidths, tableW=TW)
// → 전문 코드: references/code-helpers.md STEP 4
```

---

## STEP 5. 줄 길이 제한 (STRICT)

Word는 렌더링 시 실제 glyph 너비로 줄 넘김을 결정하므로, 코드 생성 단계에서
보수적 기준을 엄격히 지켜야 실제 출력에서 한 줄 내 표시된다.

### 작성 기준 (레벨별 유효 너비 기준)

```
환산 공식: 한글자수 + 영문자수×0.55 + 숫자×0.5 + 공백×0.25

lv1/H1:    환산 33자 이하 → 정상
lv2/H2:    환산 33자 이하 → 정상
lv3/body:  환산 25자 이하 → 정상
lv4/sub:   환산 24자 이하 → 정상

[목적 문장 작성 — 한 줄 상한 내 공간 활용]
보고서 목적 → 줄의 역할(현황/원인/영향/판단/조치/리스크) → 핵심 판단 작성 → 상한 초과 시 조정 순서.
처음부터 짧은 메모를 만들지 말고, 한 줄 상한 안에서는 목적에 맞는 내용을 충분히 담는다.
본문·제목·표 셀의 나열은 가운데 도트(·) 대신 쉼표(,) 사용.

[작성 리듬 — 하한 없음]
lv3/body:  하한 없음. 환산 25자 이하 상한 안에서 맥락 전개에 필요한 공간을 자유롭게 사용.
같은 맥락의 현황·원인·영향·판단은 짧은 여러 줄로 분해하지 말고 한 줄 안에서 결합한다.
분리는 정보 성격이 달라지거나 상한 초과로 결합이 어려울 때만 사용한다.
lv4/sub:   모든 lv3에 붙이지 말고, 근거·예시·리스크·후속조치가 필요한 lv3에만 선택 사용.
lv2별 lv3 개수: 2~5개 범위에서 맥락에 따라 달리한다. 모든 lv2를 3개씩 맞추지 않는다.

[자간 조절 구간 — 미세 초과]
환산 기준 안이어도 Word 렌더링에서 2~3글자 줄넘김 가능성이 있으면 자간을 선제 적용한다.
lv3/body 환산 26~27자 → tcs(s, -8)  (0.4pt 자간)
lv3/body 환산 28~29자 → tcs(s, -12) (0.6pt 최대)
lv3/body 환산 30자 이상 → 표현 조정 또는 개행
lv4/sub 환산 25~26자 → tcs(s, -8), 27~28자 → tcs(s, -12), 29자 이상 → 표현 조정 또는 개행
lv1/lv2 환산 33자↑도 리스크 구간으로 보고 자간 우선 적용.

[개행 필수 구간 — 레벨별 기준 +5자 이상]
자간으로 해결 불가한 경우에만 의미 손실 없는 표현 압축 후, 그래도 초과하면 의미 단위에서 강제 분리.
defs에 2개 단락으로 기재:
  · 1번 줄: { lv:LV.GAP, fn:()=>body("- 첫 번째 줄", SP.cont) }   ← after=6pt 밀착
  · 2번 줄: { lv:LV.GAP, fn:()=>bodyCont("이어지는 텍스트") }      ← after 정상
  bodyCont() → references/code-helpers.md STEP 3 / 새 줄 최소 8자
```

### 표 셀 내 줄 길이

```
셀 유효 너비 = 셀 width(DXA) - 셀 내부 패딩(좌우 각 60 DXA = 120 DXA)
한글 1자 ≈ 240 DXA 기준으로 허용 글자 수 계산

예) 셀 width 2400 DXA → 유효 2280 DXA → 한글 약 9자
    셀 width 4000 DXA → 유효 3880 DXA → 한글 약 16자

셀 내 한글 기준 허용 글자 수를 초과하면 CLines()로 개행 처리.
단, 표 데이터 셀 정렬은 C() 가운데가 기본이며 설명/비고/개요/주요 내용/근거/리스크 등 문장형 장문 컬럼에만 CL()/CLines()를 쓴다.
```

### 줄 넘침 방지 핵심 원칙

```
① 작성 전: 각 항목을 위 환산 공식으로 계산 후 기준 내 작성
② 기준 초과 시: 자간 조정 → 그래도 초과면 반드시 개행
③ 절대 금지: "아마 괜찮겠지"식 판단으로 기준 초과 작성
④ 표 셀: 짧은 값은 C() 가운데 정렬, 장문 설명 컬럼만 CLines() 활용
```

---

## FLOW 3. 검증 및 레이아웃 정제

```bash
# 별첨 없음
node report.js && python "scripts/verify_report.py" report-final.docx

# 별첨 있음
node report.js && node "scripts/patch_annex.js" report-base.docx 1 report-patched.docx && python "scripts/verify_report.py" report-patched.docx
# 별첨 2개: node scripts/patch_annex.js report-base.docx 2
```
한 줄로 연결 실행 → 첫 실패 시 즉시 중단, 원인 파악 후 코드 수정.
실패 0개 확인 후 출력.

⚠️ 별첨 있을 때 report.js에 AdmZip / buildAndPatch 코드 작성 절대 금지
   → patch_annex.js 가 모든 XML 패치를 담당

**레이아웃 정제 — 조절 수단 및 우선순위:**
```
[줄 단위 조절]
1. 자간 축소: cs=-8 → cs=-12 (최대 0.6pt) — 목적 문장 의미 보존 우선
2. 줄 내 분량 조절: 1~2자 압축 또는 보강 (의미 손실 없는 선)
3. 개행 위치 조정: 의미 단위에서 수동 개행으로 줄 배치 제어

[단락/페이지 단위 조절]
4. 조절 가능 간격 미세 조정 (범위 내에서만):
     제목 after:  120~240 DXA (6~12pt)
     lv2 블록 끝 after: 360~480 DXA (18~24pt)
5. 항목 분량 조절: 내용 1~3줄 압축 또는 보강
6. 표 셀 margins 소폭 조정

[절대 금지]
- cs < -12 (0.6pt 초과 자간 축소)
- 의미를 훼손하는 내용 수정
- 단락 옵션(keepNext 등) 사용
```

---

## STEP 1. 페이지·여백

```javascript
page: { size: { width:11906, height:16838 },
        margin: { top:1304, bottom:1304, left:1134, right:1134, gutter:0 } }
// 콘텐츠 너비: 9638 DXA
// lv1 유효 너비: 9638 DXA (left=0)
// lv2 유효 너비: 9496 DXA (left=142 차감)
// lv3 유효 너비: 9070 DXA (left=568 차감, 기호 시작 426)
// lv4 유효 너비: 8928 DXA (left=710 차감, 기호 시작 568)
// 한 줄 안전 기준: lv1/lv2 환산 33자, lv3 환산 30자, lv4 환산 29자
```

---

## STEP 2. 색상·폰트 (STRICT)

**폰트: `"바탕체"` 고정 — `"Batang"`, `"바탕"` 절대 금지**

| 요소 | half-pt | 속성 |
|------|---------|------|
| 제목(본문) | 40 | Bold + 밑줄 + CENTER |
| 날짜(본문) | 28 | RIGHT |
| H1 | 28 | Bold |
| H2·본문·소항목·번호식 | 28 | — |
| 표 기본 셀 | **24 고정** | — |
| 긴 파이프라인 셀 | 20 | 예외만 |
| 풋노트 | **18 (9pt)** | 파랑 `0000FF` |
| **별첨 레이블 텍스트박스** | **28 (14pt)** | 바탕체, 테두리 있음 |
| **별첨 제목** | **40 (20pt)** | Bold + 밑줄 + CENTER |

색상: `000000` 기본 / `0000FF` 풋노트 전용

---

## STEP 13-PRE. 빈번한 오류 TOP 5 — 코드 작성 전 자체 점검

코드 생성 시작 전 아래 5개를 머릿속으로 확인. 이 중 하나라도 놓치면 재실행 발생.

```
① 별첨 있을 때 report.js에 AdmZip/buildAndPatch 코드 작성
   → 반드시 scripts/patch_annex.js 호출로만 처리
② h2 defs text에 기호 없이 "소제목"만 작성 → 기호 미출력
   → 반드시 "□ 소제목" 형식으로 기호+공백 포함
③ h2 defs text에 "□ 소제목" 형식인데 h2() 헬퍼가 □를 또 붙임 (구버전)
   → 현재 h2() 헬퍼는 기호 자동삽입 없음, text 그대로 출력
④ annexMarker가 bodyChildren에 포함 → 본문 첫 페이지에 텍스트박스 출현
⑤ doc 생성 후 배열 unshift/push → 별첨 섹션 반영 안 됨
   (배열 완성 → doc 생성 순서 준수)
⑥ h1/h2/h3 텍스트가 한 줄을 넘는 길이 → 두 줄 발생
⑦ 짧은 표 데이터 셀에 CL() 사용 → 좌측 정렬 과다
⑧ 표 뒤 tinyGap() 수동 삽입 → buildParagraphs 자동 여백과 중복
⑨ 텍스트 없는 빈 Paragraph 삽입 → 표 뒤 자동 spacer 외 금지
```

## STEP 13. 생성 전 체크리스트

```
□ 스토리라인 확정 (사용자 선택 완료)
□ 페이지 계획표 작성 완료 (표 형식으로 페이지별 단락·표·3줄 케이스 사전 기재 / 빠른 실행 문구가 있으면 내부 처리)
□ 페이지 계획: 단락 수 기준 경계 점검 (표·계층 고립·3줄 예외 사전 탐지)
□ 풋노트 필요성 후보는 내부 판단만 수행 — 사용자에게 별도 표로 노출하지 않음
□ 반복 구조 검토: lv2 5개↑ 연속 시 기업/종목별 패턴인지 확인 후 계층 재편
□ 별첨 필요 여부 판단: 독자가 참조할 만한 상세 데이터를 표로 정리할 수 있을 때만 / 줄 글은 본문에 / 1개당 최대 1페이지 / 주제 추가 시 별첨2·3 분리
□ 본문·별첨 제목: 한글 18자 / 환산 22자 이하, 풋노트 없음, 한 줄
□ lv1(H1) 번호: 반드시 "1." "2." 아라비아 숫자 형식 — 로마자(Ⅰ I)/가나다 금지
□ 날짜: datePara() 날짜가 시스템 currentDate(오늘) 기준인지 확인 — 임의 날짜 금지 / 연도 앞 ' 금지 / after=20pt
□ 연도 apostrophe: ' 직접 입력 — report.js 내부에서 &apos;→' 자동 복원
□ 풋노트: 위치가 아니라 필요성 기준으로만 삽입 / 필요 없으면 본문·표 모두 생략
□ 풋노트 * 위치: 기본은 직전 줄 기호 시작점+2칸 / 특정 표현이면 anchor 선택 사용
□ 목적 문장 작성: 목적·역할·핵심 판단을 먼저 정하고, 한 줄 상한 안에서는 내용을 충분히 담음
□ 줄 길이: lv1/lv2 환산 33자, lv3 환산 25자, lv4 환산 24자 이하 (리스크 구간은 자간 우선)
□ 구성 리듬: lv2별 lv3 개수는 2~5개로 맥락에 따라 변주 / 모든 lv2를 3개씩 맞추지 않음
□ lv4: 모든 lv3에 기계적으로 붙이지 않고, 근거·예시·리스크·후속조치가 필요한 곳에만 선택 사용
□ 원형번호·※: lv2 금지 / 필요 위치에 따라 LV.NUM3 또는 LV.NUM4 사용
□ 표 셀: C() 가운데 정렬 기본 / 설명·비고 등 장문 컬럼만 CL()/CLines()
□ 표 앞 빈줄 없음 / 표 헤더 C(text,w,true) 가운데 / bottom HDR_BT
□ 표 너비 TW=8900, 들여쓰기 L2=142 적용
□ 표 한 페이지 내 배치 (분리 금지, 직전 분량으로 조절)
□ 표 뒤 tinyGap() 수동 삽입 금지 / buildParagraphs가 다음 레벨 기준 자동 처리
□ 빈줄 규칙: 텍스트 없는 빈 Paragraph는 표 뒤 자동 spacer 외 금지 / h1 간격은 after로 처리
□ 기호: lv3/lv4 모든 항목 기호 필수 (업체명·고유명사 포함)
   lv3(-·①※) / lv4(··①※) — 같은 레벨 혼용 금지
□ 풋노트: 해당 문장 바로 아래 / 첫 등장 위치에만 / lv1 전체 금지
□ 계층 고립: buildParagraphs(defs) — 분량 조정으로만 해결
□ 별첨: Document.sections 분리 / annexMarker(n) + annexTitlePara() + 내용
   표 기반: annexTitlePara() + mkTable() / 텍스트 기반: annexTitlePara() + buildParagraphs(annexDefs)
   텍스트 기반 annexDefs: LV.H2부터 시작 (LV.H1 금지) / 별첨 1개 = 최대 1페이지 (초과 금지)
□ 본문 별첨 언급: 본문 이해에 직접 필요한 경우에만 언급 / 불필요하면 언급 생략 / 언급 시 "별첨" 단어 직접 사용 금지, 별첨 제목만 쓰고 _isFn anchor로 "* 별첨N" 처리
□ 문체: 명사형·판단명사 종결 최우선 / ~이다·보인다·필요하다 등 문장완결형 금지 / ~함·있음·됨·해야 함은 차선
   짧은 메모형 단락 남발 금지: 한 줄 상한 안에서는 원인·근거·판단을 함께 담고, 의미 단위가 이어지면 한 단락으로 결합
   주어+조사 단문은 명사구·관형절·상황절 등으로 다양하게 재구성하되, 특정 표현 패턴 반복 금지
□ 나열: 본문·제목·표 셀의 여러 항목은 가운데 도트(·) 금지, 쉼표(,) 사용
□ 날짜: 본문·풋노트 전체 날짜는 'YY.MM.DD 형식 — "2026년 3월 28일" 금지
□ 기호 공백: ~, +는 좌우 공백 1칸 (3.20 ~ 4.20, A + B / 단항 +3% 예외)
□ 별첨 레이블 박스: 상단 const(AX_CX=954000/AX_CY=367200/AX_POSH=-7200/AX_POSV=-381600) 선언 후 wp:extent·a:xfrm 양쪽에 동일 변수 사용
   moveWithText="1" allowOverlap="1" 확인 / cx↔cy swap 절대 금지 (cx=너비가로, cy=높이세로)
□ &apos; 정규화: report.js 내 fixApos() 자동 처리 (별첨 있으면 patch_annex.js도 처리)
□ 머릿글·바닥글 없음 / datePara() 날짜만
□ 날짜·연도·분기·나열 표현 통일
```

**verify_report.py 실행 후**: 실패 0개 확인 → 레이아웃 정제 → 재검증 → 출력

---

## References

| 파일 | 담당 내용 |
|------|---------|
| `references/rules-layout.md` | STEP 6 표 페이지 처리, STEP 7 빈줄 규칙, STEP 8 계층 고립 방지, STEP 12 3줄 기준 예외 |
| `references/rules-symbol.md` | STEP 10 풋노트, STEP 10-B 반복 구조, STEP 11 기호·날짜·표현 통일 |
| `references/code-helpers.md` | STEP 3 JS 헬퍼 전문, STEP 4 표 헬퍼 전문, STEP 8 buildParagraphs()+defs 예시 |
| `references/code-helpers-annex.md` | STEP 9 별첨 규칙+XML — 별첨 필요 시에만 읽음 |
