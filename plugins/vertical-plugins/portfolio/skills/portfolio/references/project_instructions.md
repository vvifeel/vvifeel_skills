# Portfolio 분석 프로젝트 지침 v3.5

포트폴리오 관련 질문이 오면 `/mnt/skills/user/portfolio/SKILL.md`를 읽고 절차를 따른다.

---

## ⚠️ 컬럼 의미 정의 (필독)

| 원본 컬럼명 | 정규화 후 컬럼명 | 의미 |
|---|---|---|
| `투자금액($M)` | `투자금액(M$)` | 해당 라운드에서 당사가 투자한 금액 (누적값 아님) |
| `Round 총액($M)` | `Round총액(M$)` | 해당 라운드 전체 규모 (타 투자자 포함) |
| `당사 지분율(%)` | `지분율(%)` | 해당 라운드 시점 당사 지분율 |
| `투자 Round` | `Round정보` | Pre-Seed / Seed / Series A / B / C … |
| `투자 당시 기업가치($M)` | `투자당시가치(M$)` | 해당 라운드 투자 시점 밸류에이션 |
| `현재 상태` | `현재 상태` | Alive / Exited / Dead / Pending |
| `투자조건` | `투자조건` | 쉼표 구분 복수 조건 |
| `CEO` / `CFO` / `CTO` / `Co-Founder` | 동일 | 임원 성명 (CFO·CTO·Co-Founder 일부 null) |

**핵심 원칙:**
- **누적 투자액** = 동일 기업 모든 라운드 `투자금액(M$)` **합산(sum)**
- `.groupby("기업명").last()` 절대 사용 금지 → 복수 라운드 누락
- `현재가치(M$)` 컬럼은 DB에 없음 — 필요 시 CB Insights 활용
- `Exit 여부` 컬럼 없음 → `현재 상태` 컬럼으로 판별

**모든 스크립트 공통 코드 블록:**

```python
import glob, pandas as pd

# DB 파일 경로
upload_candidates    = glob.glob("/mnt/user-data/uploads/*.xlsx")
knowledge_candidates = glob.glob("/mnt/user-data/knowledge/*.xlsx")
EXCEL_PATH = (upload_candidates or knowledge_candidates)[0]

df = pd.read_excel(EXCEL_PATH)

# 컬럼 정규화
COL_MAP = {
    "투자금액($M)":          "투자금액(M$)",
    "Round 총액($M)":        "Round총액(M$)",
    "당사 지분율(%)":         "지분율(%)",
    "투자 Round":            "Round정보",
    "투자 당시 기업가치($M)": "투자당시가치(M$)",
}
df.rename(columns={k: v for k, v in COL_MAP.items() if k in df.columns}, inplace=True)

for col in ["투자년도","투자월","투자금액(M$)","Round총액(M$)","지분율(%)","투자당시가치(M$)"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# 현재 상태 분류 함수
def classify_status(v):
    v = str(v).strip(); vl = v.lower()
    if vl.startswith("alive"): return "Alive"
    if "dead" in vl: return "Dead"
    if "pending" in vl: return "Pending"
    if any(x in vl for x in ["acquired","ipo","public","subsidiary","reverse acqui-hire"]):
        return "Exited"
    return "기타"
```

> Claude 자체 지식으로 포트폴리오 데이터를 추측하거나 답변하지 않는다.
> 반드시 Python 스크립트로 실제 DB를 조회한 결과만 사용한다.

---

## 핵심 조회 원칙 (매 턴 적용)

### 1-A. 상대적 기간 표현

"최근 N년", "올해" 등은 Python으로 오늘 날짜 직접 계산:

```python
from datetime import date
today = date.today(); current_year = today.year; current_month = today.month
```

| 표현 | year_from | year_to |
|---|---|---|
| "최근 N년" | current_year - N | current_year |
| "올해" | current_year | current_year |
| "작년" | current_year - 1 | current_year - 1 |
| "최근 N개월" | N개월 전 연도 | current_year |

### 1-B. 기업명 검색 정규화

1. 한글·약어 → 영문 후보 생성 (Claude 직접 추론)
2. get_db_meta로 `companies` 목록 대조
3. 불확실하면 Fuzzy 매칭 추가 실행:

```python
from difflib import get_close_matches
partial = [n for n in all_names if query.lower() in n.lower() or n.lower() in query.lower()]
fuzzy   = get_close_matches(query, all_names, n=5, cutoff=0.5)
candidates = list(dict.fromkeys(partial + fuzzy))
```

### 1-C. 현재 상태 자연어 매핑

| 사용자 표현 | 검색 방식 | 비고 |
|---|---|---|
| "Alive", "보유 중" | `classify_status == "Alive"` | |
| "Exit한", "회수된" | `classify_status == "Exited"` | |
| "망한", "폐업" | `classify_status == "Dead"` | |
| "IPO한" | `현재 상태.str.contains("IPO", case=False)` | Public 별도 |
| "인수된" | `현재 상태.str.contains("Acquired", case=False)` | |
| "[연도] 이후 IPO" | IPO 포함 + 연도 추출 후 >= target_year | |

### 1-D. 세부 유형 분리 집계 ★

> 상위 분류로 묶어 카운트 금지. 세부 유형별로 반드시 분리하여 카운트한다.

- 집계 전 해당 컬럼 고유값 확인 (`value_counts()` 또는 get_db_meta)
- 세부 유형 있으면 → 분리 집계 + 유형별 표 + 합계 행
- 단일 숫자만 제시하고 내역 생략 금지

| 컬럼 | 상위 분류 예 | 실제 세부 유형 |
|---|---|---|
| `현재 상태` | "Exit" | IPO / Public / Acquired / Acqui-Hire / Subsidiary |
| `투자유형` | "전환" | 전환채권 / 전환우선주 / SAFE … |
| `Round정보` | "초기" | Pre-Seed / Seed / Series A … |

### 1-E. 수치 재사용 금지 ★

> 이전 답변·분석 수치는 재검증 없이 인용 불가. 집계가 필요한 시점마다 스크립트 재실행.

- 이전 턴 수치를 현재 답변에 재인용 금지
- "앞서 확인한 바와 같이" 표현으로 재실행 생략 금지
- 보고서·비교 분석·후속 답변 등 모든 맥락에 적용
- search_portfolio 목록을 눈으로 세어 집계 수치로 사용 금지

허용 예외: 동일 답변 내에서 방금 실행한 스크립트 결과를 재언급하는 경우
