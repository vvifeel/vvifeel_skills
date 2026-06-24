---
name: portfolio
description: >
  [v1.0] VC 포트폴리오 투자 DB(엑셀) 기반 검색·집계·분석·보고 작업 전용 스킬.
  포트폴리오 관련 질문이 오면 반드시 이 스킬을 사용한다.
  트리거: 포트폴리오, 투자 현황, 투자사 목록, 몇 개사, 섹터별, 지역별,
           Exit, IPO, Alive, Dead, 누적 투자, 분야별, BoD Observer, 투자조건,
           시리즈, 라운드, CB Insights, 뉴스, 현황 보고, 포트폴리오 요약 등
           투자 DB와 관련된 모든 분석 및 보고 요청.
---

# Portfolio 분석 스킬

스킬 로드 시 **반드시 먼저** `references/project_instructions.md`를 읽는다.
컬럼 정의·핵심 조회 원칙(1-A~1-E)·공통 코드 블록이 여기에 있다.

스크립트: `scripts/get_db_meta.py`, `scripts/search_portfolio.py`,
           `scripts/get_company_history.py`, `scripts/get_statistics.py`,
           `scripts/get_portfolio_summary.py`

references: `references/project_instructions.md`, `references/output_rules.md`,
            `references/analysis_principles.md`

---

## Step 1. 질문 의도 파악

스크립트 실행 전, 아래를 판단한다. 상세 기준은 → `references/analysis_principles.md`

| 질문 유형 | 의도 | 답변 방향 |
|---|---|---|
| 조건 필터 + 외부 정보 조합 | 심층 모니터링 | 기업별 상세 서술 + 의미 해석 |
| 집계·통계 | 현황 파악·보고 준비 | 수치 + 시각화 |
| 특정 기업 히스토리 | 투자 경위·의사결정 | 라운드별 맥락 + 현황 |
| 목록 확인 단독 | 빠른 참조 | 표 중심, 간결 |

---

## Step 2. Tool 선택

| 상황 | 스크립트 |
|---|---|
| 기업 목록 검색, 조건 필터 | `search_portfolio` |
| 특정 기업 투자 히스토리 | `get_company_history` |
| 섹터/지역별 집계·통계 | `get_statistics` ← **집계 요청 시 항상** |
| 전체 현황 요약 | `get_portfolio_summary` |
| 기업명 정규화 / 필터값 확인 | `get_db_meta` → 이후 본 스크립트 |

**집계 요청 시 필수 원칙:**
- "몇 개", "얼마", "평균", "합계", "분포" → 반드시 `get_statistics` 실행
- `search_portfolio` 결과를 눈으로 세거나 합산 **금지**
- `get_statistics` 실행 후 `search_portfolio`로 기업 목록도 함께 제공
- 세부 유형이 있으면 1-D 원칙에 따라 분리 집계

---

## Step 3. 스크립트 실행

스크립트는 필요한 것만 그때그때 읽어 실행한다:

| 스크립트 | 경로 | 토큰 |
|---|---|---|
| get_db_meta | `scripts/get_db_meta.py` | ~150tok |
| search_portfolio | `scripts/search_portfolio.py` | ~850tok |
| get_company_history | `scripts/get_company_history.py` | ~400tok |
| get_statistics | `scripts/get_statistics.py` | ~900tok |
| get_portfolio_summary | `scripts/get_portfolio_summary.py` | ~300tok |

**실행 절차:**
1. 해당 `.py` 파일을 읽는다
2. 파라미터 설정 부분만 수정하여 bash_tool로 실행한다
3. `_query_meta`를 [검색 해석]으로 결과 앞에 출력한다

**[검색 해석] 형식:**
```
[검색 해석] "{사용자 표현}" → {매핑된 DB 값} 으로 조회했습니다.
⚠ {_query_meta.주의 내용}  ← 주의가 있을 때만
```
`get_portfolio_summary`는 [검색 해석] 생략.

---

## Step 4. 외부 정보 조회

CB Insights 먼저, 없을 때만 웹검색. 상세 기준 → `references/analysis_principles.md`

| 정보 유형 | CB Insights 도구 |
|---|---|
| 최근 뉴스 | `get_company_news` |
| 현재 기업가치 | `get_company_profile` / `get_company_funding` |
| 펀딩 히스토리 | `get_company_funding` |
| 경쟁사 | `get_company_competitors` |
| 투자자 구성 | `get_company_investors` |

출처 표기: `(출처: CB Insights)` / `(출처: 웹검색)` / `(출처: CB Insights / 웹검색 보완)`

---

## Step 5. 결과 출력

출력 규칙 상세 → `references/output_rules.md`

**핵심 요약:**
- 카드형·대시보드형·인터랙티브 UI 아티팩트 **사용 금지**
- 보고서형 마크다운 본문 중심 (ss-report 연계 고려)
- **출력 순서: 표 → 상세 서술 → 차트 → 후속 제안** (표를 항상 먼저)
- 결과 3건 이상 → 마크다운 표 필수 (상세 서술이 있어도 표 생략 금지)
- 개조식 표현 + 헤더 구조화 (`###` 기업별, `####` 주제별)
- 뉴스 날짜 필수: `(YYYY-MM-DD)` 형식, 없으면 `(날짜 미확인)`

**차트 출력 조건 (별도 아티팩트, 단독 차트):**
- `get_statistics` / `get_portfolio_summary` 결과 → **항상**
- 시계열 데이터 포함 시 → **항상**
- "표만", "간단히" 명시 시 → 생략

---

## Step 6. 후속 제안

매 답변 말미에 평서문으로 안내 (의사 묻지 않음). 해당되는 것만 선별:

- **CB Insights**: 기업 프로필, 펀딩 히스토리, 경쟁사, 뉴스, 현재가치(→ 투자 당시 대비 배수 산출 가능)
- **웹검색**: 최신 기사, 시장 리포트, 글로벌 섹터 트렌드
- **추가 조회**: 연관 섹터·지역·투자 유형 확장
- **보고서 작성**: ss-report 스킬로 내부 보고서 정리 가능

CB Insights 연동 제안 기준 → `references/output_rules.md`

---

## 보고서 작성 요청 시

1. ss-report 스킬 로드 (ss-report-winword 또는 ss-report-full)
2. 스킬의 단계별 flow 그대로 따르기
3. 포트폴리오 데이터 필요 시 스크립트 먼저 조회 후 반영
4. 일반 docx 스킬 대신 ss-report 스킬 우선

## 에러 처리

1. 에러 메시지 사용자에게 간략히 안내
2. 파라미터 수정하여 재실행 시도
3. DB 파일 없으면: "포트폴리오 DB 파일을 첨부하거나 프로젝트 지식에 등록해 주세요."
