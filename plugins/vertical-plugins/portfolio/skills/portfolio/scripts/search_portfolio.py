# search_portfolio — 포트폴리오 검색
# 기업 목록 검색, 조건 필터 (분야·지역·상태·라운드·기간·금액·지분율·인물·투자조건 등).
# 파라미터 설정 부분만 수정 후 실행.

import glob, pandas as pd, json, re
from datetime import date

upload_candidates    = glob.glob("/mnt/user-data/uploads/*.xlsx")
knowledge_candidates = glob.glob("/mnt/user-data/knowledge/*.xlsx")
EXCEL_PATH = (upload_candidates or knowledge_candidates)[0]

df = pd.read_excel(EXCEL_PATH)
COL_MAP = {
    "투자금액($M)":          "투자금액(M$)",
    "Round 총액($M)":        "Round총액(M$)",
    "당사 지분율(%)":         "지분율(%)",
    "투자 Round":            "Round정보",
    "투자 당시 기업가치($M)": "투자당시가치(M$)",
}
df.rename(columns={k: v for k, v in COL_MAP.items() if k in df.columns}, inplace=True)
for col in ["투자년도", "투자월", "투자금액(M$)", "Round총액(M$)", "지분율(%)",
            "투자당시가치(M$)"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

def classify_status(v):
    v = str(v).strip(); vl = v.lower()
    if vl.startswith("alive"): return "Alive"
    if "dead" in vl: return "Dead"
    if "pending" in vl: return "Pending"
    if any(x in vl for x in ["acquired","ipo","public","subsidiary","reverse acqui-hire"]):
        return "Exited"
    return "기타"

# ── 상대적 기간 표현 처리 ────────────────────────────────────────
today         = date.today()
current_year  = today.year
current_month = today.month
# 예) "최근 2년" → year_from = current_year - 2, year_to = current_year
# ─────────────────────────────────────────────────────────────────

# ── 파라미터 설정 ────────────────────────────────────────────────
company_name         = None   # str: DB 기업명 (1-B 정규화 완료 후 입력)
sector               = None   # str or list[str]: 분야 ★ 다중 매핑 시 list
region               = None   # str or list[str]: 지역
investment_type      = None   # str or list[str]: "지분투자" | "전환채권"
status               = None   # str or list[str]: "Alive"|"Exited"|"Dead"|"Pending"
                               # 또는 str: 현재 상태 컬럼 contains 검색용 키워드
                               # 예) "Microsoft" → 현재 상태에 Microsoft 포함
status_use_classify  = True   # True: classify_status 기준 / False: 현재 상태 str.contains
round_info           = None   # str or list[str]: 라운드
year_from            = None   # int
year_to              = None   # int
min_amount           = None   # float (M$): 라운드별 당사 투자금액
max_amount           = None   # float (M$)
min_equity           = None   # float (%)
max_equity           = None   # float (%)
min_valuation        = None   # float (M$): 투자 당시 기업가치
max_valuation        = None   # float (M$): 투자 당시 기업가치
person_name          = None   # str: CEO/CFO/CTO/Co-Founder 통합 검색
investment_condition = None   # str or list[str]: 투자조건 키워드 (str.contains)
latest_only          = False  # bool: 기업별 최신 라운드만
sort_by              = "투자년도"  # "투자금액(M$)"|"Round총액(M$)"|"지분율(%)"|"투자년도"|"기업명"|"투자당시가치(M$)"
sort_asc             = False
limit                = 200

# ── 필터 적용 ────────────────────────────────────────────────────
def apply_list_filter(df, col, val):
    if isinstance(val, list): return df[df[col].isin(val)]
    return df[df[col] == val]

applied  = {}
warnings = []

if company_name:
    df = df[df["기업명"].str.contains(company_name, na=False, case=False)]
    applied["기업명"] = f"'{company_name}' 포함"
if sector:
    df = apply_list_filter(df, "분야", sector)
    applied["분야"] = sector
if region:
    df = apply_list_filter(df, "지역", region)
    applied["지역"] = region
if investment_type:
    valid_types = df["투자유형"].dropna().unique().tolist()
    types = investment_type if isinstance(investment_type, list) else [investment_type]
    invalid = [t for t in types if t not in valid_types]
    if invalid: warnings.append(f"투자유형 {invalid}은 DB에 없는 값입니다.")
    valid_in = [t for t in types if t in valid_types]
    if valid_in:
        df = apply_list_filter(df, "투자유형", valid_in if len(valid_in) > 1 else valid_in[0])
        applied["투자유형"] = valid_in
if status:
    if status_use_classify:
        df["_status"] = df["현재 상태"].apply(classify_status)
        statuses = status if isinstance(status, list) else [status]
        df = df[df["_status"].isin(statuses)]
        applied["현재 상태(분류)"] = statuses
    else:
        keywords = status if isinstance(status, list) else [status]
        mask = df["현재 상태"].str.contains("|".join(keywords), case=False, na=False)
        df = df[mask]
        applied["현재 상태(키워드)"] = keywords
if round_info:
    df = apply_list_filter(df, "Round정보", round_info)
    applied["라운드"] = round_info
if year_from is not None:
    df = df[df["투자년도"] >= year_from]
if year_to is not None:
    df = df[df["투자년도"] <= year_to]
if year_from is not None or year_to is not None:
    applied["투자년도"] = f"{year_from or ''}~{year_to or ''}"
if min_amount is not None:
    df = df[df["투자금액(M$)"] >= min_amount]
if max_amount is not None:
    df = df[df["투자금액(M$)"] <= max_amount]
if min_equity is not None:
    df = df[df["지분율(%)"] >= min_equity]
if max_equity is not None:
    df = df[df["지분율(%)"] <= max_equity]
if min_valuation is not None:
    df = df[df["투자당시가치(M$)"] >= min_valuation]
    applied["투자당시가치 min"] = min_valuation
if max_valuation is not None:
    df = df[df["투자당시가치(M$)"] <= max_valuation]
    applied["투자당시가치 max"] = max_valuation
if person_name:
    mask = (
        df["CEO"].str.contains(person_name, na=False, case=False) |
        df["CFO"].fillna("").str.contains(person_name, na=False, case=False) |
        df["CTO"].fillna("").str.contains(person_name, na=False, case=False) |
        df["Co-Founder"].fillna("").str.contains(person_name, na=False, case=False)
    )
    df = df[mask]
    applied["인물(CEO/CFO/CTO/Co-Founder)"] = person_name
if investment_condition:
    conditions = investment_condition if isinstance(investment_condition, list) else [investment_condition]
    for cond in conditions:
        df = df[df["투자조건"].str.contains(cond, na=False, case=False)]
    applied["투자조건"] = conditions
if latest_only:
    df = df.sort_values(["투자년도", "투자월"]).groupby("기업명").last().reset_index()
    applied["latest_only"] = "기업별 최신 라운드만"

if len(df) == 0 and applied:
    warnings.append("적용된 조건으로 매칭된 결과가 없습니다. 조건을 확인하세요.")

VALID_SORT = {"투자금액(M$)", "Round총액(M$)", "지분율(%)", "투자년도", "기업명",
              "투자당시가치(M$)"}
if sort_by not in VALID_SORT: sort_by = "투자년도"
if sort_by in df.columns: df = df.sort_values(sort_by, ascending=sort_asc)

total = len(df)
if total > limit: warnings.append(f"결과 {total}건 중 {limit}건만 반환.")
df = df.head(limit)

if "_status" in df.columns: df = df.drop(columns=["_status"])

result = {
    "_query_meta": {
        "적용된 조건": applied if applied else "없음 (전체 조회)",
        "매칭된 기업 수": int(df["기업명"].nunique()) if "기업명" in df.columns else None,
        "매칭된 총 건수": total,
        "주의": warnings if warnings else None,
    },
    "returned": len(df),
    "records": df.where(df.notna(), None).to_dict(orient="records"),
}
print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

# ── 별칭 매핑 참고 ─────────────────────────────────────────────
# 현재 상태 매핑:
# | 사용자 표현 | status 파라미터 | status_use_classify |
# |---|---|---|
# | "보유 중", "Alive", "살아있는" | `"Alive"` | `True` |
# | "Exit한", "회수된" | `"Exited"` | `True` |
# | "망한", "폐업" | `"Dead"` | `True` |
# | "IPO한" | `"IPO"` | `False` |
# | "인수된" | `"Acquired"` | `False` |
# | "[특정 기업]에 인수된" | `"[기업명]"` | `False` |
# | "[연도] 이후 IPO" | → 1-C의 연도 추출 패턴 사용 | — |
