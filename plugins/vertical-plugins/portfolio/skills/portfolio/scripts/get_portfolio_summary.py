# get_portfolio_summary — 전체 현황 요약
# 파라미터 없음. 전체 포트폴리오 현황 집계.

import glob, pandas as pd, json

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

companies  = df["기업명"].nunique()
status_co  = df.groupby("기업명")["_status"].last()
first_inv  = df.groupby("기업명")["투자년도"].min()
yearly_new = {int(k): int(v) for k, v in first_inv.value_counts().sort_index().items() if pd.notna(k)}

result = {
    "총 포트폴리오사":          int(companies),
    "투자 분야 수":             int(df["분야"].nunique()),
    "투자 지역 수":             int(df["지역"].nunique()),
    "총 투자 건수(라운드)":     int(len(df)),
    "현재 상태별 기업 수": {
        "Alive":   int((status_co == "Alive").sum()),
        "Exited":  int((status_co == "Exited").sum()),
        "Dead":    int((status_co == "Dead").sum()),
        "Pending": int((status_co == "Pending").sum()),
    },
    "누적 투자금액(M$)":        float(df["투자금액(M$)"].sum().round(1)),
    "분야별 포트폴리오사 수":   {k: int(v) for k, v in df.groupby("분야")["기업명"].nunique().sort_values(ascending=False).items()},
    "지역별 포트폴리오사 수":   {k: int(v) for k, v in df.groupby("지역")["기업명"].nunique().sort_values(ascending=False).items()},
    "연도별 신규 투자사 수":    yearly_new,
}
print(json.dumps(result, ensure_ascii=False, indent=2))