"""
verify_report.py v5
인라인 풋노트 기반 검증
"""
import sys, zipfile, re, html
sys.stdout.reconfigure(encoding='utf-8')  # Windows cp949 환경 UTF-8 강제
from collections import defaultdict

PASS = "✅"; WARN = "⚠️ "; FAIL = "❌"
results = []
FN_COLOR = "0000FF"
FN_SZ    = "18"  # fn sz=18 (9pt) — 신규 기준
CS_LIMIT = -12

def log(lv, cat, msg):
    results.append((lv, cat, msg))
    print(f"  {lv} [{cat}] {msg}")

def load(path):
    with zipfile.ZipFile(path) as z:
        return XmlDoc(z.read("word/document.xml").decode("utf-8"))

class XmlDoc(str):
    def __new__(cls, value):
        obj = str.__new__(cls, value)
        obj._cache = {}
        return obj

    def cached(self, key, builder):
        if key not in self._cache:
            self._cache[key] = builder()
        return self._cache[key]

def paras_full(xml):
    if isinstance(xml, XmlDoc):
        return xml.cached("paras_full", lambda: re.findall(r'<w:p>.*?</w:p>', xml, re.DOTALL))
    return re.findall(r'<w:p>.*?</w:p>', xml, re.DOTALL)

def paras_inner(xml):
    if isinstance(xml, XmlDoc):
        return xml.cached("paras_inner", lambda: re.findall(r'<w:p>(.*?)</w:p>', xml, re.DOTALL))
    return re.findall(r'<w:p>(.*?)</w:p>', xml, re.DOTALL)

def blocks_all(xml):
    if isinstance(xml, XmlDoc):
        return xml.cached("blocks_all", lambda: list(re.finditer(r'<w:(p|tbl)\b.*?</w:\1>', xml, re.DOTALL)))
    return list(re.finditer(r'<w:(p|tbl)\b.*?</w:\1>', xml, re.DOTALL))

def is_fn(p):
    return (f'<w:color w:val="{FN_COLOR}"/>' in p and
            f'<w:sz w:val="{FN_SZ}"/>' in p and
            'wp:anchor' not in p)

_TEXT_CACHE = {}
def text_of(block):
    cached = _TEXT_CACHE.get(block)
    if cached is not None:
        return cached
    text = html.unescape("".join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', block, re.DOTALL))).strip()
    if len(_TEXT_CACHE) < 20000:
        _TEXT_CACHE[block] = text
    return text

def equiv_len(text):
    korean = sum(1 for c in text if '가' <= c <= '힣')
    digit_ = sum(1 for c in text if c.isdigit())
    ascii_ = sum(1 for c in text if ord(c) <= 0x007E and c != ' ' and not c.isdigit())
    space_ = text.count(' ')
    other  = sum(1 for c in text if ord(c) > 0x007E and not ('가' <= c <= '힣'))
    return korean + ascii_*0.55 + digit_*0.5 + space_*0.25 + other*0.5

def check_font(xml):
    bad = xml.count('w:ascii="Batang"') + xml.count('w:ascii="바탕"')
    ok  = xml.count('w:ascii="바탕체"')
    log(FAIL if bad else PASS,"폰트",
        f'"Batang"/"바탕" {bad}건' if bad else f'"바탕체" {ok}건')

def check_colors(xml):
    allowed = {"000000","0000FF","FFFFFF","CCCCCC"}
    bad = [c for c in re.findall(r'<w:color w:val="([0-9A-Fa-f]{6})"', xml)
           if c.upper() not in {a.upper() for a in allowed}]
    if bad:
        cnt = defaultdict(int)
        for c in bad: cnt[c.upper()]+=1
        for c,n in cnt.items(): log(FAIL,"색상",f"금지 색상 #{c} {n}건")
    else: log(PASS,"색상","금지 색상 없음")

def check_blank(xml):
    blocks = blocks_all(xml)
    paras = [m.group(0) for m in blocks if m.group(1) == "p"]
    bad = 0
    h1sep_count = 0
    prev_kind = None
    for m in blocks:
        if m.group(1) == "tbl":
            prev_kind = "tbl"
            continue
        p = m.group(0)
        after_m = re.search(r'w:after="(\d+)"', p)
        sz      = re.search(r'w:val="28"', p)
        bold    = re.search(r'<w:b[/> ]', p)
        is_h1   = sz and bold
        text    = text_of(p)
        if not after_m or text or is_h1 or is_fn(p):
            if text:
                prev_kind = "p"
            continue
        after_val = int(after_m.group(1))
        # 표 뒤 자동 spacer만 허용. 수동 tinyGap(100) 또는 임의 빈 단락은 금지.
        if prev_kind == "tbl" and after_val in (180, 240, 400, 480):
            prev_kind = "p"
            continue
        bad += 1
        prev_kind = "p"
    msg = f"불필요한 빈 줄 {bad}개" if bad else "빈 줄 규칙 정상"
    log(FAIL if bad else PASS, "빈줄", msg)
    # H1 직후 빈 단락 탐지 (defs 배열에서 H1 뒤 GAP 삽입한 경우)
    h1_gap_bad = 0
    for i, p in enumerate(paras):
        sz   = re.search(r'w:val="28"', p)
        bold = re.search(r'<w:b[/> ]', p)
        text = text_of(p)
        if sz and bold and text and i + 1 < len(paras):
            # H1 단락 뒤 다음 단락이 빈 단락인지 확인
            next_p = paras[i + 1]
            next_text = text_of(next_p)
            next_after = re.search(r'w:after="(\d+)"', next_p)
            if not next_text and next_after and int(next_after.group(1)) != 480:
                h1_gap_bad += 1
                log(WARN, "H1중복간격", f"H1 '{text[:20]}' 직후 불필요한 빈 단락 — defs에서 GAP 제거 필요")
    if not h1_gap_bad:
        log(PASS, "H1중복간격", "H1 직후 중복 빈 단락 없음")

def check_fn_marker(xml):
    n = len(re.findall(r'\[FN(?:\|[^\]]+)?\]', xml))
    log(FAIL if n else PASS,"풋노트마커",
        f"[FN] 잔류 {n}개 — inject 미실행" if n else "마커 주입 완료")

def check_fn_period(xml):
    paras = paras_inner(xml)
    bad = 0
    for p in paras:
        if not is_fn(p): continue
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)
        full = "".join(texts).strip()
        if full.endswith("."): bad+=1; log(FAIL,"풋노트마침표",f"마침표: '{full[-20:]}'")
    if not bad: log(PASS,"풋노트마침표","풋노트 마침표 없음")

def check_fn_spacing(xml):
    paras = paras_inner(xml)
    fn_iss = prev_iss = 0
    for i, p in enumerate(paras):
        if not is_fn(p): continue
        after = re.search(r'w:after="(\d+)"', p)
        a = int(after.group(1)) if after else -1
        next_lv = None
        for j in range(i + 1, len(paras)):
            if is_fn(paras[j]):
                continue
            t = text_of(paras[j])
            if not t:
                continue
            next_lv = para_level(paras[j])
            break
        if next_lv is not None:
            expected = fn_after_dxa(next_lv)
            if a != expected:
                fn_iss += 1
                log(WARN, "풋노트간격", f"풋노트 after={a} (기대:{expected}, 다음 lv{next_lv})")
        elif a < 0:
            fn_iss += 1
            log(WARN, "풋노트간격", "풋노트 after 누락")
        if i > 0:
            prev = paras[i-1]
            pa_m = re.search(r'w:after="(\d+)"', prev)
            pa = int(pa_m.group(1)) if pa_m else -1
            if pa != 20:
                prev_iss+=1
                pt = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', prev, re.DOTALL)).strip()
                log(WARN,"풋노트앞단락",f"앞 단락 after={pa} (기대:20): '{pt[:25]}'")
    if not fn_iss and not prev_iss:
        log(PASS,"풋노트간격","풋노트 앞단락 1pt 및 after 보존 정상")

def para_level(p):
    text = text_of(p)
    if re.search(r'<w:sz w:val="28"', p) and re.search(r'<w:b[/> ]', p) and re.match(r'^\d+\.', text):
        return 1
    if text.startswith("□"):
        return 2
    left_m = re.search(r'w:left="(\d+)"', p)
    if not left_m:
        return None
    left = int(left_m.group(1))
    if 540 <= left <= 590:
        return 3
    if 680 <= left <= 740:
        return 4
    return None

def fn_after_dxa(next_lv):
    rule = {1: 24, 2: 20, 3: 12, 4: 9}.get(next_lv, 6)
    remain = rule - 1 - 9
    return (max(0, remain) + 1) * 20

def check_fn_style(xml):
    paras = paras_inner(xml)
    iss = cnt = 0
    for p in paras:
        if not is_fn(p): continue
        cnt += 1
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)
        full = "".join(texts).strip()
        if not full.startswith("*"): iss+=1; log(WARN,"풋노트스타일",f"'*' 없음: '{full[:25]}'")
        if len(full) > 60: log(WARN,"풋노트길이",f"60자↑: '{full[:35]}' ({len(full)}자)")
    if not iss: log(PASS,"풋노트스타일",f"풋노트 {cnt}개 정상" if cnt else "풋노트 없음")

def _block_text(block):
    return "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', block, re.DOTALL)).strip()

def check_pagebreak(xml):
    paras = paras_inner(xml)
    pb  = sum(1 for p in paras if 'pageBreakBefore' in p)
    man = sum(1 for p in paras if '<w:br w:type="page"/>' in p)
    # keepNext: 계층 고립 방지 목적으로 허용 — FAIL 처리 금지
    kwn = sum(1 for p in paras if '<w:keepNext' in p)  # <w:keepNext/> 또는 <w:keepNext w:val="1"/>
    log(FAIL if pb  else PASS, "페이지나눔", f"pageBreakBefore {pb}개" if pb  else "pageBreakBefore 없음")
    log(FAIL if man else PASS, "페이지나눔", f"수동 PageBreak {man}개" if man else "수동 PageBreak 없음")
    log(PASS, "keepNext", f"keepNext {kwn}개 (계층 고립 방지용, 정상)" if kwn else "keepNext 없음")

def check_table(xml):
    tbls = re.findall(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL)
    if not tbls: log(PASS,"표","표 없음"); return
    iss = 0
    for i, tbl in enumerate(tbls):
        rows = re.findall(r'<w:tr>(.*?)</w:tr>', tbl, re.DOTALL)
        col_ws = re.findall(r'<w:gridCol w:w="(\d+)"', tbl)
        no_cant = sum(1 for r in rows if 'cantSplit' not in r)
        if no_cant: iss+=1; log(WARN,"표분리",f"표{i+1}: cantSplit 없는 행 {no_cant}개")
        if col_ws and min(int(w) for w in col_ws) < 500:
            log(WARN,"표너비",f"표{i+1}: 최소 열너비 {min(int(w) for w in col_ws)} DXA")
    if not iss: log(PASS,"표","표 분리 방지 정상")

def check_table_alignment(xml):
    """표 데이터 셀은 가운데 정렬이 기본. 짧은 좌측 정렬 데이터 셀은 WARN."""
    tbls = re.findall(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL)
    warns = 0
    for ti, tbl in enumerate(tbls, 1):
        rows = re.findall(r'<w:tr>.*?</w:tr>', tbl, re.DOTALL)
        for ri, row in enumerate(rows[1:], 2):  # 헤더 행 제외
            cells = re.findall(r'<w:tc>.*?</w:tc>', row, re.DOTALL)
            for ci, cell in enumerate(cells, 1):
                text = text_of(cell)
                if not text:
                    continue
                jc_m = re.search(r'<w:jc w:val="([^"]+)"', cell)
                jc = jc_m.group(1) if jc_m else ""
                has_break = '<w:br' in cell
                if jc == "left" and not has_break and equiv_len(text) <= 10:
                    warns += 1
                    log(WARN, "표정렬", f"표{ti} R{ri}C{ci} 짧은 값 좌측정렬: '{text[:20]}'")
    if not warns:
        log(PASS, "표정렬", "짧은 데이터 셀 가운데 정렬 정상")

def check_h2_indent(xml):
    """lv2(□)는 기호 자체가 1칸 뒤에서 시작해야 하므로 hanging 금지."""
    paras = paras_full(xml)
    bad = 0
    for p in paras:
        if is_fn(p):
            continue
        text = text_of(p)
        if not text.startswith("□"):
            continue
        left_m = re.search(r'w:left="(\d+)"', p)
        hanging_m = re.search(r'w:hanging="(\d+)"', p)
        left = int(left_m.group(1)) if left_m else None
        if left != 142 or hanging_m:
            bad += 1
            log(FAIL, "lv2들여쓰기", f"□ left={left}, hanging={hanging_m.group(1) if hanging_m else '-'} — left=142/hanging 없음 필요: '{text[:30]}'")
    if not bad:
        log(PASS, "lv2들여쓰기", "□ 기호 1칸 들여쓰기 정상")

def check_circle_level(xml):
    """원형번호·※는 lv3 또는 lv4로만 사용하고 lv2 위치에서는 금지."""
    paras = paras_full(xml)
    bad = 0
    specials = tuple("①②③④⑤⑥⑦⑧⑨⑩") + ("※",)
    for p in paras:
        if is_fn(p):
            continue
        text = text_of(p)
        if not text.startswith(specials):
            continue
        left_m = re.search(r'w:left="(\d+)"', p)
        left = int(left_m.group(1)) if left_m else None
        if left is None or not (540 <= left <= 590 or 680 <= left <= 740):
            bad += 1
            log(FAIL, "원형번호레벨", f"원형번호·※는 lv3/lv4만 허용: left={left} '{text[:30]}'")
    if not bad:
        log(PASS, "원형번호레벨", "원형번호·※ lv3/lv4 사용 정상")

def check_symbol(xml):
    paras = paras_inner(xml)
    iss = 0
    # lv2(left≈142): □만 허용. 원형번호·※는 lv3/lv4 전용.
    lv2_ok = ('□',)
    # lv3(left≈568)/lv4(left≈710): 기호 시작점은 각각 3칸/4칸
    lv34_ok = ('-','·','①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩','※','→','←','↑','↓','⇒')
    for p in paras:
        if is_fn(p): continue
        m = re.search(r'w:left="(\d+)"', p)
        if not m: continue
        indent = int(m.group(1))
        full = text_of(p)
        if not full: continue
        if 120 <= indent <= 170:  # lv2
            if not any(full.startswith(s) for s in lv2_ok):
                iss+=1; log(WARN,'기호',f"lv2 기호 오류(□만 허용): '{full[:35]}'")
        elif indent >= 540:       # lv3/lv4
            if not any(full.startswith(s) for s in lv34_ok):
                iss+=1; log(WARN,'기호',f"body/sub 기호 누락: '{full[:35]}'")
    if not iss: log(PASS,'기호','기호 통일 정상')

def check_spacing(xml):
    # SP 기준: h1After=400(20pt)/h1SepAfter=480(24pt), h2After=240(12pt)
    # H1 after 허용 범위: 400(일반) 또는 480(섹션 전환) — 둘 다 정상
    # H2 after 허용 범위: 240(12pt) 기본, 단락 끝 360~480(18~24pt) 조절 가능
    paras = paras_inner(xml)
    bad = 0
    for p in paras:
        if is_fn(p): continue
        after = re.search(r'w:after="(\d+)"', p)
        sz    = re.search(r'<w:sz w:val="(\d+)"', p)
        bold  = re.search(r'<w:b[/> ]', p)
        if not after or not sz: continue
        a = int(after.group(1)); s = int(sz.group(1))
        # H1(sz=28, bold, no indent): after=400 또는 480 — 그 외는 경고
        if s==28 and bold:
            if a not in (400, 480): bad+=1; log(WARN,"간격",f"H1 after={a} (정상:400 또는 480)")
        # H2(sz=28, no bold, left=142): after=240~480 허용
        elif s==28 and not bold and 'left="142"' in p:
            if not (200<=a<=500): bad+=1; log(WARN,"간격",f"H2 after={a} (정상:240~480)")
    if not bad: log(PASS,"간격","H1·H2 간격 정상")

def check_last_hyphen(xml):
    paras = paras_inner(xml)
    iss = 0
    ok_next = ('-','·','①','②','③','④','□')
    for i, p in enumerate(paras):
        if is_fn(p) or 'w:left="568"' not in p: continue  # L3=568 (기호 3칸 기준)
        after = re.search(r'w:after="(\d+)"', p)
        text  = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)).strip()
        if not text.startswith("- "): continue
        if i + 1 < len(paras) and is_fn(paras[i + 1]):
            continue
        nxt = ""
        for j in range(i+1, min(i+4, len(paras))):
            if is_fn(paras[j]): continue
            t = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', paras[j], re.DOTALL)).strip()
            if t: nxt=t; break
        is_last = not any(nxt.startswith(s) for s in ok_next)
        if is_last and after:
            a = int(after.group(1))
            if not (180<=a<=480): iss+=1; log(WARN,"마지막하이픈",f"after={a}: '{text[:30]}'")
    if not iss: log(PASS,"마지막하이픈","마지막 하이픈 간격 정상")

def check_cs_and_len(xml):
    """줄 길이 (환산 글자수 기준) 및 자간, 빈 항목 검증"""
    paras = paras_inner(xml)
    bad_len = 0; bad_cs = 0; empty_items = 0
    ok_s = ('□','-','·','①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩','※')

    for p in paras:
        if is_fn(p): continue
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)
        full = "".join(texts).strip()

        # 빈 항목: 기호만 있고 내용 없음
        if full in ('-', '·', '①','②','③','- ','· '):
            empty_items += 1
            log(FAIL,"빈항목",f"기호만 있고 내용 없음: '{full}' — 삭제 필요")
            continue

        # 줄 길이: 한글+영문 환산 글자수 (들여쓰기 레벨별 기준)
        m_indent = re.search(r'w:left="(\d+)"', p)
        is_h1 = bool(re.search(r'<w:sz w:val="28"', p) and re.search(r'<w:b[/> ]', p) and re.match(r'^\d+\.', full))
        if is_h1 or (any(full.startswith(s) for s in ok_s) and m_indent):
            indent = int(m_indent.group(1)) if m_indent else 0
            if 120 <= indent <= 170:
                limit, risk_at, cs8_at = 37, 32, 33  # lv2: 33자↑ 리스크, +5부터 FAIL
            elif 540 <= indent <= 590:
                limit, risk_at, cs8_at = 29, 25, 27  # lv3: 26~27=-8, 28~29=-12, 30↑ FAIL
            elif 680 <= indent <= 740:
                limit, risk_at, cs8_at = 28, 24, 26  # lv4: 25~26=-8, 27~28=-12, 29↑ FAIL
            elif is_h1:
                limit, risk_at, cs8_at = 37, 32, 33  # lv1: 33자↑ 리스크, +5부터 FAIL
            else:
                continue
            korean = sum(1 for c in full if '가' <= c <= '힣')
            digit_ = sum(1 for c in full if c.isdigit())
            ascii_ = sum(1 for c in full if ord(c) <= 0x007E and c != ' ' and not c.isdigit())
            space_ = full.count(' ')
            other  = sum(1 for c in full if ord(c) > 0x007E and not ('가' <= c <= '힣'))
            equiv  = korean + ascii_*0.55 + digit_*0.5 + space_*0.25 + other*0.5
            if equiv > limit:
                bad_len += 1
                log(FAIL,"줄길이",f"환산 {equiv:.0f}자 > {limit} — 개행 필요: '{full[:40]}'")
            elif equiv > risk_at:
                required_cs = -8 if equiv <= cs8_at else -12
                applied = [int(m.group(1)) for m in re.finditer(r'<w:spacing\s+w:val="(-?\d+)"', p)]
                if not applied or min(applied) > required_cs:
                    log(WARN,"줄길이",f"환산 {equiv:.0f}자 — tcs(..., {required_cs}) 적용 필요: '{full[:40]}'")

    # 자간 한도: -12 (0.6pt)
    for r in re.findall(r'<w:r>(.*?)</w:r>', xml, re.DOTALL):
        for m in re.finditer(r'<w:spacing\s+w:val="(-?\d+)"', r):
            val = int(m.group(1))
            if val < CS_LIMIT:
                bad_cs += 1
                log(FAIL,"자간",f"cs={val} (한도 {CS_LIMIT}={abs(CS_LIMIT)/20}pt)")

    if not bad_len and not empty_items: log(PASS,"줄길이","레벨별 줄 길이 기준 통과, 빈 항목 없음")
    if not bad_cs: log(PASS,"자간",f"자간 {CS_LIMIT} ({abs(CS_LIMIT)/20}pt) 이내")

def check_title_length(xml):
    """본문 제목(sz=40, bold, underline, center)과 별첨 제목 글자 수 검증.
    한글 18자 / 환산 22자 초과 시 FAIL."""
    paras = paras_inner(xml)
    bad = 0
    for p in paras:
        if is_fn(p): continue
        is_sz40  = bool(re.search(r'<w:sz w:val="40"', p))
        is_bold  = bool(re.search(r'<w:b[/> ]', p))
        is_ul    = bool(re.search(r'<w:u ', p))
        is_ctr   = bool(re.search(r'w:val="center"', p))
        if not (is_sz40 and is_bold and is_ul and is_ctr): continue
        text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)).strip()
        if not text: continue
        ko = sum(1 for c in text if '\uAC00' <= c <= '\uD7A3')
        other = len(text) - ko
        equiv = ko + other * 0.6
        if ko > 18 or equiv > 22:
            bad += 1
            log(FAIL, "제목길이",
                f"한글{ko}자 환산{equiv:.0f}자 초과 (한18/환22 이내): '{text[:30]}'")
    if not bad: log(PASS, "제목길이", "모든 제목 한 줄 기준 이내")

def check_date_para(xml):
    """보고서 날짜는 YY.MM.DD(요일), apostrophe 없음, after=400(20pt)."""
    paras = paras_full(xml)
    found = bad = 0
    for p in paras[:8]:
        text = text_of(p)
        if not re.match(r"^'?\d{2}\.\d{2}\.\d{2}\([월화수목금토일]\)$", text):
            continue
        found += 1
        if text.startswith("'"):
            bad += 1
            log(FAIL, "날짜", f"보고서 날짜 연도 앞 apostrophe 금지: '{text}'")
        after_m = re.search(r'w:after="(\d+)"', p)
        after = int(after_m.group(1)) if after_m else -1
        if after != 400:
            bad += 1
            log(FAIL, "날짜간격", f"datePara after={after} (정상:400=20pt)")
    if not found:
        log(WARN, "날짜", "상단 날짜 단락 미탐지")
    elif not bad:
        log(PASS, "날짜", "보고서 날짜 형식 및 after=20pt 정상")

def check_apos(xml):
    """&apos; 노출 여부 검증 — inject/patch 단계 누락 탐지."""
    n = xml.count('&apos;')
    if n:
        log(FAIL, "apostrophe", f"&apos; {n}건 노출 — buildAndPatch() 정규화 확인 필요")
    else:
        log(PASS, "apostrophe", "&apos; 노출 없음")

def check_body_annex_word(xml):
    """본문에는 '별첨' 단어를 직접 쓰지 않고 제목만 언급해야 함."""
    paras = paras_full(xml)
    bad = 0
    for p in paras:
        text = text_of(p)
        if not text or "별첨" not in text:
            continue
        if is_fn(p):
            continue
        # 별첨 레이블 텍스트박스/마커성 단락은 본문 언급이 아님
        if "wp:anchor" in p or re.fullmatch(r"별첨\s*\d+", text):
            continue
        bad += 1
        log(FAIL, "본문별첨언급", f"본문 '별첨' 단어 직접 사용: '{text[:40]}'")
    if not bad:
        log(PASS, "본문별첨언급", "본문 내 '별첨' 직접 언급 없음")

def check_long_date_format(xml):
    """본문·풋노트의 날짜는 'YY.MM.DD 형식을 사용하고 장문 날짜는 금지."""
    paras = paras_full(xml)
    bad = 0
    long_date = re.compile(r"20\d{2}년\s*\d{1,2}월\s*\d{1,2}일")
    dotted_date = re.compile(r"20\d{2}\.\d{1,2}\.\d{1,2}")
    for p in paras:
        text = text_of(p)
        if not text:
            continue
        # 상단 작성일은 check_date_para가 별도 검증
        if re.fullmatch(r"\d{2}\.\d{2}\.\d{2}\([월화수목금토일]\)", text):
            continue
        hits = long_date.findall(text) + dotted_date.findall(text)
        if hits:
            bad += len(hits)
            log(FAIL, "날짜표기", f"본문·풋노트 날짜는 'YY.MM.DD 사용: {hits} — '{text[:45]}'")
    if not bad:
        log(PASS, "날짜표기", "본문·풋노트 장문 날짜 표기 없음")

def check_nominal_style(xml):
    """조사·동사형 종결을 줄이고 명사형 종결을 유도하는 경고성 점검."""
    paras = paras_full(xml)
    sentence_ending_re = re.compile(r"(이다|로 보인다|으로 보인다|가 필요하다|이 필요하다|할 필요가 있다|해야 한다)$")
    fallback_ending_re = re.compile(r"(해야 함|하여야 함|하고 있음|되어 있음|하였음|했음|필요함|중요함|지속함|확인함|검토함|포함함|있음|없음|됨|다룸|커짐|작아짐)$")
    middle_dot_list_re = re.compile(r"\S\s*·\s*\S")
    footnote_particle_re = re.compile(r"^[^ ]{1,30}(?:은|는|이|가|에는|에서는)\s")
    warns = 0
    fails = 0
    for p in paras:
        text = text_of(p)
        if not text:
            continue
        stripped = re.sub(r"^(\*|-|·|□|※|[①②③④⑤⑥⑦⑧⑨⑩])\s*", "", text).strip()
        if not stripped:
            continue
        if middle_dot_list_re.search(stripped):
            fails += 1
            log(FAIL, "나열표현", f"본문·제목·표 셀 나열은 쉼표 사용: '{text[:45]}'")
            continue
        if not (is_fn(p) or text.startswith(("- ", "· ", "※")) or re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩]", text)):
            continue
        if sentence_ending_re.search(stripped):
            fails += 1
            log(FAIL, "문장완결형", f"명사형·판단명사 종결 필요: '{text[:45]}'")
        elif fallback_ending_re.search(stripped):
            warns += 1
            log(WARN, "차선문체", f"가능하면 판단명사 종결 검토: '{text[:45]}'")
        if is_fn(p) and footnote_particle_re.search(stripped):
            warns += 1
            log(WARN, "풋노트문체", f"풋노트 주어·조사 생략 검토: '{text[:45]}'")
    if not warns and not fails:
        log(PASS, "명사형문체", "조사·동사형 종결 경고 없음")

def check_operator_spacing(xml):
    """~, + 기호는 좌우 공백 1칸이어야 함. 단항 + 부호는 예외."""
    paras = paras_full(xml)
    bad = 0
    op_re = re.compile(r"(?<=[0-9A-Za-z가-힣\)'\].])\s*([~+])\s*(?=[0-9A-Za-z가-힣\('\[])")
    for p in paras:
        text = text_of(p)
        if not text:
            continue
        for m in op_re.finditer(text):
            op = m.group(1)
            if m.group(0) == f" {op} ":
                continue
            bad += 1
            log(FAIL, "기호공백", f"'{op}' 좌우 공백 1칸 필요: '{text[:50]}'")
    if not bad:
        log(PASS, "기호공백", "~, + 좌우 공백 정상")

def check_textbox_font(xml):
    """별첨 텍스트박스 내부 폰트가 바탕체인지 검증."""
    # txbxContent 내 rFonts 탐지
    txbx_blocks = re.findall(r'<w:txbxContent>(.*?)</w:txbxContent>', xml, re.DOTALL)
    if not txbx_blocks:
        return  # 텍스트박스 없음 — skip
    bad = 0
    for block in txbx_blocks:
        # lang 태그(w:lang)는 폰트가 아니므로 제외하고 폰트 속성만 검사
        fonts = re.findall(r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"', block)
        non_batang = [f for f in fonts if f not in ("바탕체",)]
        if non_batang:
            bad += 1
            log(FAIL, "텍스트박스폰트", f"바탕체 아닌 폰트 발견: {non_batang}")
        # eastAsia 누락 체크 (한글 렌더링 핵심) — w:lang 과 혼동 주의
        if 'w:eastAsia="바탕체"' not in block:
            bad += 1
            log(FAIL, "텍스트박스폰트", "w:eastAsia 바탕체 누락 — 한글이 기본 폰트로 렌더링됨")
    if not bad:
        log(PASS, "텍스트박스폰트", f"텍스트박스 {len(txbx_blocks)}개 모두 바탕체")

def check_header_footer(path):
    """머릿글·바닥글 존재 여부 확인 — 있으면 FAIL."""
    import zipfile as _zf
    try:
        with _zf.ZipFile(path) as z:
            names = z.namelist()
        has_header = any("header" in n for n in names)
        has_footer = any("footer" in n for n in names)
        if has_header:
            log(FAIL, "머릿글", "머릿글 파일 감지됨 — 생성 금지")
        if has_footer:
            log(FAIL, "바닥글", "바닥글 파일 감지됨 — 생성 금지")
        if not has_header and not has_footer:
            log(PASS, "머릿글/바닥글", "없음 (정상)")
    except Exception as e:
        log(WARN, "머릿글/바닥글", f"확인 실패: {e}")

def check_symbol_mixing(xml):
    """같은 들여쓰기 레벨에서 하이픈(-)/도트(·)와 원형번호(①②)/※ 혼용 탐지.
    원형번호·※ 는 하이픈·도트를 대신할 수 있으나, 같은 블록에서 섞이면 안 됨."""
    paras = paras_inner(xml)
    # 들여쓰기 레벨별로 기호 추적
    # lv3 = left≈568, lv4 = left≈710
    level_symbols = {}   # indent_left → set of symbol types seen
    warns = 0
    prev_indent = None
    for p in paras:
        text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)).strip()
        if not text:
            continue
        # H2 경계에서 초기화
        if text.startswith("□"):
            level_symbols = {}
            prev_indent = None
            continue
        # 들여쓰기 값 추출
        m = re.search(r'w:left="(\d+)"', p)
        if not m:
            continue
        indent = int(m.group(1))
        sym_type = None
        if text.startswith("- "):
            sym_type = "hyphen"
        elif text.startswith("· ") or text.startswith("· "):
            sym_type = "dot"
        elif re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text):
            sym_type = "circle"
        elif text.startswith("※"):
            sym_type = "note"
        if sym_type is None:
            continue
        if indent not in level_symbols:
            level_symbols[indent] = set()
        level_symbols[indent].add(sym_type)
        # 같은 레벨에 hyphen/dot 과 circle/note 가 공존하면 혼용
        syms = level_symbols[indent]
        plain = syms & {"hyphen", "dot"}
        special = syms & {"circle", "note"}
        if plain and special:
            warns += 1
            log(WARN, "기호혼용", f"들여쓰기 {indent}DXA 레벨: {syms} 혼용 감지 — '{text[:30]}'")
            level_symbols[indent] = {sym_type}  # 리셋하여 중복 경고 방지
    if not warns:
        log(PASS, "기호혼용", "기호 혼용 없음")

def check_empty_body(xml):
    """기호 없이 텍스트만 있는 body/sub 레벨 단락 탐지.
    업체명·고유명사 포함 모든 lv3/lv4 항목은 반드시 기호로 시작해야 함."""
    paras = paras_inner(xml)
    bad = 0
    # 허용 기호 목록
    allowed_starts = ("-", "·", "•", "□", "※",
                      "①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩",
                      "→","←","↑","↓","⇒")
    for p in paras:
        if is_fn(p): continue
        text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)).strip()
        if not text: continue
        # lv3(left≈568) 또는 lv4(left≈710) 들여쓰기 확인
        m = re.search(r'w:left="(\d+)"', p)
        if not m: continue
        indent = int(m.group(1))
        if indent < 540: continue  # lv2 이상은 체크 안 함
        # 기호로 시작하지 않으면 FAIL
        if not any(text.startswith(s) for s in allowed_starts):
            bad += 1
            log(WARN, "기호누락", f"기호 없는 body/sub 항목: '{text[:30]}'")
    if not bad:
        log(PASS, "기호누락", "모든 body/sub 항목 기호 포함")

def check_symbols(xml):
    """금지 기호 탐지: 계층 기호와 혼동되는 기호 및 잘 사용되지 않는 특수기호."""
    # 혼동 기호: 계층 기호(□ - · ① ※)와 헷갈릴 수 있는 것들
    banned_pattern = r'[○●▶▷▪•◆◇]'
    # 지양 기호: 보고서에 어울리지 않는 특수기호
    discouraged_pattern = r'[★☆✱☎☞☜♥♣♠♦]'

    paras = paras_inner(xml)
    banned_count = 0
    discouraged_count = 0
    for p in paras:
        if is_fn(p): continue
        text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL))
        banned_hits = re.findall(banned_pattern, text)
        disc_hits   = re.findall(discouraged_pattern, text)
        if banned_hits:
            banned_count += len(banned_hits)
            log(FAIL, "금지기호", f"사용 금지 기호 발견: {banned_hits} — '{text[:40]}'")
        if disc_hits:
            discouraged_count += len(disc_hits)
            log(WARN, "지양기호", f"사용 지양 기호: {disc_hits} — '{text[:40]}'")
    if not banned_count and not discouraged_count:
        log(PASS, "기호검증", "금지/지양 기호 없음")

def check_footnote_position(xml):
    """풋노트 위치 검증: lv1(H1) 바로 다음 fn이 오면 WARN."""
    paras = paras_inner(xml)
    for i, p in enumerate(paras):
        # H1 단락 탐지: sz=28, bold
        is_h1 = bool(re.search(r'<w:sz w:val="28"', p)) and bool(re.search(r'<w:b[/> ]', p))
        if not is_h1: continue
        # 바로 다음 단락이 fn인지 확인
        if i + 1 < len(paras) and is_fn(paras[i + 1]):
            h1_text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)).strip()
            log(WARN, "풋노트위치", f"H1 바로 아래 풋노트 금지: '{h1_text[:30]}'")
    log(PASS, "풋노트위치", "H1 직후 풋노트 없음")

def check_balance(xml):
    paras = paras_inner(xml)
    h1s = []
    for i, p in enumerate(paras):
        if is_fn(p): continue
        sz   = re.search(r'w:val="28"', p)
        bold = re.search(r'<w:b[/> ]', p)
        text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL)).strip()
        if sz and bold and text: h1s.append((i, text[:25]))
    if not h1s: log(WARN,"분량","H1 없음"); return
    h1s.append((len(paras),"END"))
    iss = 0
    for idx in range(len(h1s)-1):
        s, title = h1s[idx]; e, _ = h1s[idx+1]
        sec = paras[s:e]
        tbl_c = sum(1 for p in sec if '<w:tcPr>' in p or 'w:tblCellMar' in p)
        adj = len(sec) - int(tbl_c*0.7)
        if adj<15: iss+=1; log(WARN,"분량",f"'{title}' 추정 {adj}단락 — 여백 가능")
        elif adj>50: iss+=1; log(WARN,"분량",f"'{title}' 추정 {adj}단락 — 넘침 가능")
    if not iss: log(PASS,"분량","H1 섹션 분량 적정")

def main():
    path = sys.argv[1] if len(sys.argv)>1 else "report-final.docx"
    print(f"\n{'='*60}\n  보고서 검증: {path}\n{'='*60}")

    # ── 파일 무결성 검증 (최우선) ──
    print("\n[파일무결성]")
    try:
        from docx import Document as _DocxDoc
        _DocxDoc(path)
        print("  ✅ [파일무결성] python-docx 열기 성공")
    except Exception as e:
        print(f"  ❌ [파일무결성] 파일 손상 — Word에서 열리지 않을 수 있음: {e}")
        print("  → buildAndPatch XML 패치 오류 또는 annexLabelXml 구조 문제 확인 필요")
        sys.exit(1)  # 파일 손상 시 이후 검증 무의미하므로 중단

    try: xml = load(path)
    except Exception as e: print(f"❌ 로드 실패: {e}"); sys.exit(1)

    print("\n[폰트/색상]"); check_font(xml); check_colors(xml)
    print("\n[빈줄]"); check_blank(xml)
    print("\n[풋노트]"); check_fn_marker(xml); check_fn_period(xml); check_fn_spacing(xml); check_fn_style(xml)
    print("\n[페이지 나눔]"); check_pagebreak(xml)
    print("\n[표]"); check_table(xml); check_table_alignment(xml)
    print("\n[기호 통일]"); check_h2_indent(xml); check_circle_level(xml); check_symbol(xml)
    print("\n[단락 간격]"); check_spacing(xml); check_last_hyphen(xml)
    print("\n[자간/한줄]"); check_cs_and_len(xml)
    print("\n[제목/표현]"); check_title_length(xml); check_date_para(xml); check_long_date_format(xml); check_body_annex_word(xml); check_nominal_style(xml); check_operator_spacing(xml); check_apos(xml)
    print("\n[머릿글/바닥글]"); check_header_footer(path)
    print("\n[기호/풋노트]"); check_symbols(xml); check_footnote_position(xml); check_symbol_mixing(xml); check_empty_body(xml)
    print("\n[텍스트박스]"); check_textbox_font(xml)
    print("\n[분량 추정]"); check_balance(xml)
    print("\n[별첨검증]"); check_annex_marker_residue(xml); check_annex_label_box(xml)

def check_annex_marker_residue(xml):
    """__ANNEX_LABEL_N__ 텍스트가 XML에 잔류하면 FAIL (buildAndPatch 미실행 또는 마커 탐지 실패)"""
    found = re.findall(r'__ANNEX_LABEL_\d+__', xml)
    if found:
        log(FAIL, "별첨마커잔류", f"미치환 마커 {len(found)}개 잔류: {found[:3]}")
    else:
        log(PASS, "별첨마커잔류", "별첨 마커 잔류 없음")

def check_annex_label_box(xml):
    """별첨 레이블 텍스트박스 고정값 위반 탐지: cx=792000, cy=363600, posOffset=-363600"""
    EXPECTED_CX = 792000
    EXPECTED_CY = 363600
    EXPECTED_POS = -363600
    anchors = re.findall(r'<wp:anchor.*?</wp:anchor>', xml, re.DOTALL)
    iss = 0
    for a in anchors:
        cx_m = re.search(r'cx="(\d+)"', a)
        cy_m = re.search(r'cy="(\d+)"', a)
        pos_m = re.search(r'<wp:posOffset>(-?\d+)</wp:posOffset>', a)
        if not (cx_m and cy_m): continue
        cx, cy = int(cx_m.group(1)), int(cy_m.group(1))
        if cx == EXPECTED_CX or cy == EXPECTED_CY:
            if cx == EXPECTED_CY and cy == EXPECTED_CX:
                iss += 1; log(FAIL, "별첨레이블박스", f"cx↔cy swap 오류: cx={cx}, cy={cy}")
            if pos_m and int(pos_m.group(1)) > 0:
                iss += 1; log(FAIL, "별첨레이블박스", f"posOffset 양수 오류: {pos_m.group(1)} (음수여야 함)")
    if not iss:
        log(PASS, "별첨레이블박스", "별첨 레이블 박스 고정값 정상")

def main_summary():
    fails  = sum(1 for r in results if r[0]==FAIL)
    warns  = sum(1 for r in results if r[0]==WARN)
    passes = sum(1 for r in results if r[0]==PASS)
    print(f"\n{'='*60}")
    print(f"  ✅ {passes}개 통과 / ⚠️  {warns}개 경고 / ❌ {fails}개 실패")
    if fails:   print("  → 실패 항목 수정 후 재생성")
    elif warns: print("  → 경고 검토 후 필요시 조정")
    else:       print("  → 전 항목 통과 ✅")
    print(f"{'='*60}\n")
    sys.exit(1 if fails else 0)

if __name__ == "__main__":
    main()
    main_summary()

