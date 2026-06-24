# get_statistics — 통계/집계
# metrics 허용값: count_companies | total_invested | sum_round_amount | avg_amount |
#   median_amount | avg_equity | count_rows | count_exited | count_dead | avg_round_interval
# group_by 허용값: 분야 | 지역 | 투자유형 | Round정보 | 현재 상태

import glob, pandas as pd, json, numpy as np
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

df["_status"] = df["현재 상태"].apply(classify_status)

def apply_list_filter(df, col, val):
    if isinstance(val, list): return df[df[col].isin(val)]
    return df[df[col] == val]

# ── 상대적 기간 처리 ─────────────────────────────────────────────
today        = date.today()
current_year = today.year
# ────────────────────────────────────────────────────────────────

# ── 파라미터 설정 ────────────────────────────────────────────────
group_by        = "분야"
metrics         = ["count_companies", "total_invested"]
sector          = None   # str or list[str]
region          = None
investment_type = None
status_filter   = None   # str or list[str]: "Alive"|"Exited"|"Dead"|"Pending"
year_from       = None
year_to         = None

# ── 필터 적용 ────────────────────────────────────────────────────
df_all = df.copy()
applied_filters = {}

if sector:
    df_all = apply_list_filter(df_all, "분야", sector)
    applied_filters["분야"] = sector
if region:
    df_all = apply_list_filter(df_all, "지역", region)
    applied_filters["지역"] = region
if investment_type:
    valid_types = df_all["투자유형"].dropna().unique().tolist()
    types = investment_type if isinstance(investment_type, list) else [investment_type]
    valid_in = [t for t in types if t in valid_types]
    if valid_in:
        df_all = apply_list_filter(df_all, "투자유형", valid_in if len(valid_in) > 1 else valid_in[0])
        applied_filters["투자유형"] = valid_in
if status_filter:
    statuses = status_filter if isinstance(status_filter, list) else [status_filter]
    df_all = df_all[df_all["_status"].isin(statuses)]
    applied_filters["현재 상태"] = statuses

df_period = df_all.copy()
if year_from is not None:
    df_period = df_period[df_period["투자년도"] >= year_from]
    applied_filters["투자년도 from"] = year_from
if year_to is not None:
    df_period = df_period[df_period["투자년도"] <= year_to]
    applied_filters["투자년도 to"] = year_to

agg_period  = df_period.groupby(group_by)
result_rows = {}
notes       = []

for metric in metrics:
    if metric == "count_companies":
        for k, v in agg_period["기업명"].nunique().items():
            result_rows.setdefault(k, {group_by: k})["포트폴리오사 수"] = int(v)

    elif metric == "count_rows":
        for k, v in agg_period.size().items():
            result_rows.setdefault(k, {group_by: k})["투자 건수"] = int(v)

    elif metric == "total_invested":
        for k, v in agg_period["투자금액(M$)"].sum().round(1).items():
            result_rows.setdefault(k, {group_by: k})["당사 투자금액 합계(M$)"] = float(v)

    elif metric == "sum_round_amount":
        for k, v in agg_period["Round총액(M$)"].sum().round(1).items():
            result_rows.setdefault(k, {group_by: k})["라운드 총액 합계(M$)"] = float(v)

    elif metric == "avg_amount":
        for k, v in agg_period["투자금액(M$)"].mean().round(2).items():
            result_rows.setdefault(k, {group_by: k})["평균 라운드 투자금액(M$)"] = float(v)

    elif metric == "median_amount":
        for k, v in agg_period["투자금액(M$)"].median().round(2).items():
            result_rows.setdefault(k, {group_by: k})["라운드 투자금액 중앙값(M$)"] = float(v)

    elif metric == "avg_equity":
        df_eq = df_period[df_period["지분율(%)"] > 0]
        for k, v in df_eq.groupby(group_by)["지분율(%)"].mean().round(2).items():
            result_rows.setdefault(k, {group_by: k})["평균 지분율(%)"] = float(v)
        notes.append("avg_equity: 지분율 0% 제외 후 계산")

    elif metric == "count_exited":
        exited_df = df_all[df_all["_status"] == "Exited"]
        for k, v in exited_df.groupby(group_by)["기업명"].nunique().items():
            result_rows.setdefault(k, {group_by: k})["Exit 완료 기업 수"] = int(v)

    elif metric == "count_dead":
        dead_df = df_all[df_all["_status"] == "Dead"]
        for k, v in dead_df.groupby(group_by)["기업명"].nunique().items():
            result_rows.setdefault(k, {group_by: k})["Dead 기업 수"] = int(v)

    elif metric == "avg_round_interval":
        gaps = []
        for co, grp in df_period.groupby("기업명"):
            years = grp.sort_values("투자년도")["투자년도"].dropna().tolist()
            if len(years) >= 2:
                gaps.extend([years[i] - years[i-1] for i in range(1, len(years))])
        avg_gap = round(float(np.mean(gaps)), 1) if gaps else None
        result_rows["_avg_round_interval"] = {
            "평균 후속투자 간격(년)": avg_gap,
            "분석 대상 간격 수": len(gaps),
        }
        notes.append("avg_round_interval: 멀티라운드 기업의 라운드 간 연도 차이 평균")

rows = [v for k, v in result_rows.items() if not k.startswith("_")]
special = {k: v for k, v in result_rows.items() if k.startswith("_")}
rows = sorted(rows, key=lambda x: list(x.values())[1] if len(x) > 1 else 0, reverse=True)

result = {
    "_query_meta": {
        "집계 기준": group_by,
        "사용된 metrics": metrics,
        "사전 필터": applied_filters if applied_filters else "없음",
        "주의": notes if notes else None,
    },
    "rows": rows,
}
if special:
    result["special_metrics"] = special

print(json.dumps(result, ensure_ascii=False, indent=2, default=str))