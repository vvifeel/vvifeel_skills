# get_db_meta — DB 메타 확인
# 분야·지역·라운드·투자유형·투자조건·기업명 목록 조회. 기업명 정규화 및 필터값 확인 시 우선 실행.

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

def classify_status(v):
    v = str(v).strip(); vl = v.lower()
    if vl.startswith("alive"): return "Alive"
    if "dead" in vl: return "Dead"
    if "pending" in vl: return "Pending"
    if any(x in vl for x in ["acquired","ipo","public","subsidiary","reverse acqui-hire"]):
        return "Exited"
    return "기타"

meta = {
    "sectors":            sorted(df["분야"].dropna().unique().tolist()),
    "regions":            sorted(df["지역"].dropna().unique().tolist()),
    "rounds":             sorted(df["Round정보"].dropna().unique().tolist()),
    "investment_types":   sorted(df["투자유형"].dropna().unique().tolist()),
    "status_categories":  ["Alive", "Exited", "Dead", "Pending"],
    "investment_conditions": sorted(set(
        c.strip() for val in df["투자조건"].dropna()
        for c in val.split(",")
    )),
    "companies":          sorted(df["기업명"].dropna().unique().tolist()),
}
print(json.dumps(meta, ensure_ascii=False, indent=2))