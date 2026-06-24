# get_company_history — 기업 투자 히스토리
# company_name은 1-B 정규화를 거쳐 DB에 존재하는 정확한 이름으로 교체.

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

company_name = "Harvey AI"  # ← 1-B 정규화 완료된 기업명으로 교체

matched  = df[df["기업명"].str.contains(company_name, na=False, case=False)]
warnings = []

if matched.empty:
    result = {"error": f"'{company_name}'에 해당하는 기업을 찾을 수 없습니다."}
else:
    found = matched["기업명"].unique().tolist()
    if len(found) > 1:
        warnings.append(f"'{company_name}' 검색에 {len(found)}개 기업 매칭됨.")

    companies_data = {}
    for co in found:
        co_df = matched[matched["기업명"] == co].sort_values(["투자년도", "투자월"])
        latest = co_df.iloc[-1]
        val_at = float(latest["투자당시가치(M$)"]) if pd.notna(latest.get("투자당시가치(M$)")) else None

        companies_data[co] = {
            "총_라운드수":          len(co_df),
            "첫_투자년도":          int(co_df["투자년도"].min()) if not co_df["투자년도"].isna().all() else None,
            "최신_라운드":          latest["Round정보"] if len(co_df) > 0 else None,
            "현재_상태":            latest.get("현재 상태"),
            "누적_투자금액(M$)":    round(float(co_df["투자금액(M$)"].sum()), 2),
            "투자당시가치(M$)":     val_at,
            "CEO":                  latest.get("CEO"),
            "CFO":                  latest.get("CFO"),
            "CTO":                  latest.get("CTO"),
            "Co-Founder":           latest.get("Co-Founder"),
            "history":              co_df.where(co_df.notna(), None).to_dict(orient="records"),
        }

    result = {
        "_query_meta": {
            "검색어": company_name,
            "매칭된 기업": found,
            "주의": warnings if warnings else None,
        },
        "companies": companies_data,
    }
print(json.dumps(result, ensure_ascii=False, indent=2, default=str))