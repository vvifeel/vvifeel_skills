"""
inject_footnotes.py v6
풋노트를 인라인 파랑 단락으로 구현 (플로팅 텍스트 상자 완전 폐기)

이유:
- wrapTopAndBottom = 텍스트 위/아래 배치 (텍스트 앞 아님)
- wrapNone = 텍스트와 겹침 발생
- 인라인 단락이 가장 안정적: 빈줄 없음, 겹침 없음, 텍스트 앞처럼 보임

동작:
- [FN]텍스트[/FN] 마커를 파랑 인라인 단락으로 교체
- 앵커 단락 after=20 (1pt 단락뒤 여백)
- 풋노트 단락 before=0, after=[FN] 마커 단락 after 보존
- 마침표 자동 제거
"""
import sys, re, zipfile, shutil, os, html
from urllib.parse import unquote
sys.stdout.reconfigure(encoding='utf-8')  # Windows cp949 환경 UTF-8 강제

FONT  = "바탕체"
COLOR = "0000FF"
SZ    = 18   # 9pt — verify FN_SZ="18"과 일치


def strip_period(text):
    return text.rstrip(".")


def make_inline_fn(text, indent_dxa=142, after_dxa=0):
    """
    인라인 파랑 단락 XML 생성
    들여쓰기: 설명 대상 윗줄의 기호 시작점 + 2칸
    """
    text = strip_period(html.unescape(text))
    esc  = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return (
        f'<w:p>'
        f'<w:pPr>'
        f'<w:spacing w:before="0" w:after="{after_dxa}"/>'
        f'<w:ind w:left="{indent_dxa}"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
        f'<w:color w:val="{COLOR}"/>'
        f'<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
        f'<w:color w:val="{COLOR}"/>'
        f'<w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/>'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">* {esc}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )


def para_text(p):
    return ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.DOTALL))


def equiv_width(text):
    width = 0.0
    for ch in text:
        if ch.isspace():
            width += 0.25
        elif ch.isdigit():
            width += 0.5
        elif ('A' <= ch <= 'Z') or ('a' <= ch <= 'z'):
            width += 0.55
        elif ch in ".,;:()[]{}'\"/%+-":
            width += 0.5
        else:
            width += 1.0
    return width


def text_start_indent(prev_para, default=142):
    """
    직전 문단의 기호 시작점 + 2칸 기준 들여쓰기.
    한 칸=142DXA, 2칸=기호 1칸+공백 1칸으로 고정한다.
    """
    left_m = re.search(r'w:left="(\d+)"', prev_para)
    if not left_m:
        return default

    left = int(left_m.group(1))
    hanging_m = re.search(r'w:hanging="(\d+)"', prev_para)
    text = html.unescape(para_text(prev_para)).strip()
    marker_m = re.match(r'^([-·□※]|[①②③④⑤⑥⑦⑧⑨⑩])\s+', text)
    if marker_m:
        symbol_start = left
        if hanging_m:
            symbol_start = left - int(hanging_m.group(1))
        return symbol_start + 284
    if hanging_m:
        return left
    return left


def para_after(p, default=0):
    after_m = re.search(r'w:after="(\d+)"', p)
    return int(after_m.group(1)) if after_m else default


def strip_leading_marker(text):
    return re.sub(r'^\s*(?:[-·□※]|[①②③④⑤⑥⑦⑧⑨⑩])\s+', '', text, count=1)


def anchor_indent(prev_para, anchor, fallback):
    if not anchor:
        return fallback

    raw_text = html.unescape(para_text(prev_para)).strip()
    body_text = strip_leading_marker(raw_text)
    idx = body_text.find(anchor)
    if idx < 0:
        return fallback

    indent = fallback + round(equiv_width(body_text[:idx]) * 142)
    return max(fallback, min(indent, fallback + 6200))


def parse_fn_marker(p):
    fn_m = re.search(r'\[FN(?:\|([^\]]+))?\](.*?)\[/FN\]', p, re.DOTALL)
    if not fn_m:
        return None
    anchor = unquote(html.unescape(fn_m.group(1))) if fn_m.group(1) else None
    return anchor, fn_m.group(2)


def process(src, dst):
    shutil.copy2(src, src + ".tmp")
    with zipfile.ZipFile(src + ".tmp") as z:
        contents = {n: z.read(n) for n in z.namelist()}
    xml = contents["word/document.xml"].decode("utf-8")

    # [FN]...[/FN] 단락을 인라인 파랑 단락으로 교체
    # 동시에 직전 단락의 after도 0으로 조정하여 밀착
    count = len(re.findall(r'\[FN(?:\|[^\]]+)?\]', xml))

    # 1단계: FN 단락 교체 (after=0)
    # 풋노트 들여쓰기는 직전 본문 단락의 텍스트 시작점(left)을 따른다.
    def replace_fn_paras(xml_str):
        paras = list(re.finditer(r'<w:p>.*?</w:p>', xml_str, re.DOTALL))
        modified = {}
        for i, m in enumerate(paras):
            p = m.group(0)
            parsed = parse_fn_marker(p)
            if not parsed:
                continue
            anchor, text = parsed
            indent = 142
            prev = None
            if i > 0:
                prev = paras[i - 1].group(0)
                indent = text_start_indent(prev, indent)
            if prev is not None:
                indent = anchor_indent(prev, anchor, indent)
            modified[i] = make_inline_fn(text, indent, para_after(p, 0))

        if not modified:
            return xml_str

        result = []
        last_end = 0
        for i, m in enumerate(paras):
            result.append(xml_str[last_end:m.start()])
            result.append(modified.get(i, m.group(0)))
            last_end = m.end()
        result.append(xml_str[last_end:])
        return ''.join(result)

    xml = replace_fn_paras(xml)

    # 2단계: 풋노트 앞 단락의 after를 0으로 조정 (원문과 밀착)
    # 정규식으로 안전하게 처리
    def adjust_prev_after(xml_str):
        """
        풋노트 인라인 단락 바로 앞 단락의 w:after="N"을 w:after="20"으로 변경
        풋노트 식별: color=0000FF + sz=18 인 단락
        """
        # 모든 단락을 분리
        para_splits = list(re.finditer(r'<w:p>.*?</w:p>', xml_str, re.DOTALL))
        if not para_splits:
            return xml_str

        # 풋노트 단락 인덱스 파악
        fn_indices = set()
        for i, m in enumerate(para_splits):
            p = m.group(0)
            if f'<w:color w:val="{COLOR}"/>' in p and f'<w:sz w:val="{SZ}"/>' in p:
                fn_indices.add(i)

        if not fn_indices:
            return xml_str

        # 풋노트 앞 단락의 after 값 수정 (1pt)
        modified = {}
        for fn_i in fn_indices:
            if fn_i > 0:
                prev_i = fn_i - 1
                prev_text = para_splits[prev_i].group(0)
                new_text = re.sub(r'(w:after=)"(\d+)"', r'\1"20"', prev_text)
                if 'w:after=' not in prev_text:
                    new_text = new_text.replace('<w:pPr>', '<w:pPr><w:spacing w:after="20"/>', 1)
                modified[prev_i] = new_text

        if not modified:
            return xml_str

        # 원본 XML 재조합
        result = []
        last_end = 0
        for i, m in enumerate(para_splits):
            result.append(xml_str[last_end:m.start()])
            if i in modified:
                result.append(modified[i])
            else:
                result.append(m.group(0))
            last_end = m.end()
        result.append(xml_str[last_end:])
        return ''.join(result)

    xml = adjust_prev_after(xml)

    # &apos; → ' (U+0027) 복원: docx 라이브러리 escape 및 풋노트 이중 escape 정규화
    xml = xml.replace("&amp;apos;", "'").replace("&apos;", "'")
    contents["word/document.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in contents.items():
            z.writestr(n, d)
    os.remove(src + ".tmp")
    print(f"  풋노트 주입: {count}개 (인라인 파랑 단락) → {dst}")


if __name__ == "__main__":
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else src.replace(".docx", "-final.docx")
    process(src, dst)
    print("✅ 완료")
