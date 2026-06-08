# code-helpers.md — JS 헬퍼 전문·defs 예시·별첨 XML

---

## STEP 3. JS 헬퍼 표준 (전문)

> **효율화 원칙**: 아래 헬퍼 블록은 매번 재해석하지 말고 그대로 복사 사용.
> 변경이 허용된 부분은 defs 배열, SP 조절 가능 값(주석에 명시), AX_* 상수뿐.
> 나머지는 수정하지 않아도 모든 보고서에 동작함.

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, VerticalAlign } = require('docx');
// 금지: Footer, PageNumber, Header — 머릿글·바닥글 절대 생성 금지
const fs = require('fs');

const F      = { name: "바탕체", cs: "바탕체" };
const BLACK  = "000000";
const BLUE   = "0000FF";
// ── 한글 문자열 유니코드 이스케이프 금지 규칙 ──────────────────────────
// JS 소스 내에서 한글을 \uXXXX 이스케이프로 쓰면 일부 환경에서 문자열로 노출됨
// 반드시 실제 한글 문자를 소스에 직접 작성할 것
// 금지: "바탕체"  → 바탕체가 아니라 이스케이프 문자열로 삽입됨 ✗
// 허용: "바탕체"                 → 실제 UTF-8 한글 문자로 정상 삽입됨 ✓
//
// apostrophe 상수 (파일 상단 필수 선언):
// 연도 표기: t("'26년") 처럼 ' 를 직접 작성 (inject가 &apos; 자동 복원)
// 표 너비: lv2(142) 기준 콘텐츠 너비(9638)에서 여유분 확보, lv2 우측 끝 라인 이하 유지
const TW = 8900;

const t   = (s, o={}) => new TextRun({ text:s, font:F, color:BLACK, size:28, ...o });
const tb  = (s, o={}) => new TextRun({ text:s, font:F, color:BLACK, size:28, bold:true, ...o });
const tu  = (s, o={}) => new TextRun({ text:s, font:F, color:BLACK, size:28, underline:{}, ...o });
// 자간: 최대 -12 (0.6pt). 목적 문장 의미 보존을 위해 미세 초과 시 먼저 적용
const tcs = (s, cs=-12, o={}) => new TextRun({ text:s, font:F, color:BLACK, size:28, characterSpacing:cs, ...o });

// ── 연도 표기 apostrophe 규칙 ──────────────────────────────────────────
// '26년 에서 앞 작은따옴표는 U+0027 (ASCII 직선 작은따옴표 ' ) 를 직접 사용
// docx 라이브러리가 ' 를 &apos; 로 escape하더라도
// inject_footnotes.py 최종 단계에서 &apos; → ' 로 자동 복원됨
// buildAndPatch() 에서도 동일하게 복원 처리
//
// 올바른 방법: ' 를 텍스트에 직접 작성
//   t("'26년")   ✓   t("'26.3Q")  ✓
//
// 금지: ’, ‘, &apos; 등 우회 표현 — 불필요하고 혼란만 가중

// ── 제목(본문·별첨) 글자 수 제한 ──────────────────────────────────────
// 제목은 반드시 한 줄. 풋노트도 없음. 두 줄 절대 금지.
// titlePara / annexTitlePara 에 전달하는 문자열:
//   한글 전용: 18자 이하  (19자↑ → 내용 압축 후 작성)
//   한영 혼용: 환산 22자 이하
// 긴 제목은 부제·보충 설명을 모두 제거하고 핵심 키워드만 남겨 압축할 것
// 예) "전체 AI 포트폴리오 투자 내역 (CB Insights 정정 반영)"  ← 금지 (너무 김)
//     "AI 포트폴리오 투자 내역"  ← 올바른 방식

// ── 들여쓰기 상수 (INDENT 폐기 → 레벨별 고정값) ──
// 한글 바탕체 12pt 기준 1칸 = 0.25cm ≈ 142 DXA
const L2 = 142;   // lv2: 1칸 = 0.25cm
const L3 = 568;   // lv3: 기호 3칸, 본문 4칸 (hanging 1칸)
const L4 = 710;   // lv4: 기호 4칸, 본문 5칸 (hanging 1칸)
const HANG = 142; // hanging 공통: 1칸

// ── 간격 상수 (DXA = pt × 20) ──
// 규칙 원칙: 각 헤더 레벨 앞 기준, 직전줄 after로 구현
//   h1 직전줄 after = 24pt  (첫 번째 h1 제외)
//   h2 직전줄 after = 20pt  (18~20pt 허용)
//   h3 직전줄 after = 12pt
//   h4 직전줄 after =  9pt
//   h1/h2/h3: 절대 두 줄 금지 — 줄 길이 제한 엄수
//   안전 기준: lv1/lv2 환산 33자, lv3 환산 25자, lv4 환산 24자
//   h4 두 줄 필요 시: SUB(첫째줄, after=6pt) + SUB2(둘째줄, after=다음레벨기준)
const SP = {
  titleAfter:  120,  // 제목 after 6pt
  dateAfter:   400,  // 날짜 after 20pt
  // ── 헤더 직전줄 after (= 다음이 해당 레벨일 때 현재줄에 적용) ──
  toH1:        480,  // →h1 직전: 24pt
  toH2:        400,  // →h2 직전: 20pt  (보고서 구성에 따라 360~400 허용)
  toH3:        240,  // →h3 직전: 12pt
  toH4:        180,  // →h4 직전:  9pt
  cont:        120,  // h4 두 줄 개행 시 첫째 줄 after: 6pt  (SUB2 단락에 적용)
  // ── 풋노트 ──
  fnBefore:     20,  // [FN] 마커용. inject 후 실제 1pt는 윗줄 after=20으로 적용
  // fnAfter: buildParagraphs 내 fnAfterDxa() 함수로 동적 계산
  // ── 기타 ──
  cellV:        60,  // 표 셀 before/after 3pt
};

const titlePara = (s) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: SP.titleAfter },  // 기본 6pt, 레이아웃 정제 시 120~240 조절 가능
  children:[new TextRun({text:s, font:F, size:40, bold:true, color:BLACK, underline:{}})]
});
// datePara: 날짜 문자열만 전달 (예: "26.05.16(금)") — 연도 앞 ' 금지, 부서명·작성자 추가 절대 금지
const datePara = (s) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { after: SP.dateAfter },
  children:[t(s)]
});
// h1: 단락. after는 buildParagraphs가 자동 결정. h1/h2/h3 절대 두 줄 금지
// ⚠️ after 기본값(SP.toH2=20pt)은 buildParagraphs 미사용 시 fallback용
//    buildParagraphs 내에서는 afterDxa(nextLv)가 덮어씀 — 기본값이 실제 after가 아님
const h1 = (s, after=SP.toH2) => new Paragraph({
  spacing:{ after },
  children:[t(s, {size:28, bold:true})]
});
// h1Sep: 사용하지 않음 — buildParagraphs가 after로 흡수. defs에 삽입 절대 금지
const h1Sep = () => new Paragraph({ spacing:{after:SP.toH1}, children:[t("")] });
// h2: 단락. after는 buildParagraphs 자동 결정. 절대 두 줄 금지
// ⚠️ after 기본값(SP.toH3=12pt)은 buildParagraphs 미사용 시 fallback용
// ⚠️ 기호 자동삽입 없음 — defs의 text에 반드시 "□ 소제목" 형식으로 포함할 것
//    기호 없이 "소제목"만 넘기면 기호가 출력되지 않음
//    defs 예: { lv: LV.H2, text: "□ 추진 배경" }  ← □ 와 공백 포함
const h2 = (s, after=SP.toH3) => new Paragraph({
  indent:{ left:L2 },
  spacing:{ after },
  children:[t(s)]
});
// body(h3): 단락. after는 buildParagraphs 자동 결정. 절대 두 줄 금지
const body = (runs, after=SP.toH4) => new Paragraph({
  indent:{ left:L3, hanging:HANG },
  spacing:{ after },
  children: Array.isArray(runs) ? runs : [t(runs)]
});
// bodyCont: body 개행 2번 줄용. 기호 없이 L3 위치에서 시작 (hanging 없음)
// 사용: lv3 환산 30자↑ 강제 개행 시 { lv:LV.GAP, fn:()=>bodyCont("이어지는 텍스트") }
const bodyCont = (runs, after=SP.cont) => new Paragraph({
  indent:{ left:L3 },
  spacing:{ after },
  children: Array.isArray(runs) ? runs : [t(runs)]
});
// sub(h4): 단락. after는 buildParagraphs 자동 결정.
// 두 줄 필요 시 sub(첫째줄) + sub2(둘째줄) 분리 사용
const sub = (runs, after=SP.toH4) => new Paragraph({
  indent:{ left:L4, hanging:HANG },
  spacing:{ after },
  children: Array.isArray(runs) ? runs : [t(runs)]
});
// sub2: h4 두 번째 줄 전용 continuation 단락
// 들여쓰기 left=L4(710), hanging 없음 → sub()의 텍스트 시작 위치(710)와 정렬
// after는 항상 SP.cont(6pt) — buildParagraphs가 자동 적용, 직접 지정 금지
// ⚠️ sub2 다음에는 반드시 SUB2 없는 단락이 와야 함 (세 줄 이상은 금지)
const sub2 = (runs, after=SP.cont) => new Paragraph({
  indent:{ left:L4 },
  spacing:{ after },
  children: Array.isArray(runs) ? runs : [t(runs)]
});
// num3: 원형번호·※ (lv3 위치, body와 동일 들여쓰기/여백)
const num3 = (s, after=SP.toH3) => new Paragraph({
  indent:{ left:L3, hanging:HANG },
  spacing:{ after },
  children: Array.isArray(s) ? s : [t(s)]
});
// num4: 원형번호·※ (lv4 위치, sub와 동일 들여쓰기/여백)
const num4 = (s, after=SP.toH4) => new Paragraph({
  indent:{ left:L4, hanging:HANG },
  spacing:{ after },
  children: Array.isArray(s) ? s : [t(s)]
});
// fn: 풋노트 마커 단락
// after는 buildParagraphs가 fnAfterDxa()로 동적 결정 — 직접 지정 금지
// lv1/lv2/lv3 마지막에 올 때도 해당 레벨 종료 기준 after 적용
// _isFn은 문자열 또는 { anchor:"직전 문장 내 표현", text:"표시할 설명" } 허용
// anchor는 위치 계산용이며 최종 출력에는 보이지 않음
const fnPayload = (note) => {
  if (note && typeof note === "object") {
    const anchor = note.anchor ? `|${encodeURIComponent(note.anchor)}` : "";
    return `[FN${anchor}]${note.text ?? ""}[/FN]`;
  }
  return `[FN]${note}[/FN]`;
};
const fn = (note, after=SP.toH4) => new Paragraph({
  indent:{ left:L2 },
  spacing:{ before:SP.fnBefore, after },
  children:[new TextRun({text:fnPayload(note), font:F, size:18, color:BLUE})]
});
// tableGap: 표 뒤 자동 여백. defs에 직접 쓰지 않음 — buildParagraphs가 필요 시 삽입
const tableGap = (after=SP.toH3) => new Paragraph({ spacing:{after}, children:[t("")] });

// 별첨 제목: 바탕체 20pt, Bold, 밑줄, 가운데, after 24pt
const annexTitlePara = (s) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 480 },
  children:[new TextRun({text:s, font:F, size:40, bold:true, color:BLACK, underline:{}})]
});
```

---

## STEP 4. 표 헬퍼 표준 (전문)

```javascript
const OUTER  = {style:BorderStyle.SINGLE, size:4,  color:BLACK};
const INNER  = {style:BorderStyle.SINGLE, size:4,  color:BLACK};
const HDR_BT = {style:BorderStyle.THICK,  size:12, color:BLACK};
// 헤더 행 테두리 원칙:
//   top·left·right → OUTER (일반 굵기)
//   bottom         → HDR_BT (굵게) ← 헤더 행 아랫부분만 굵게

// 헤더 셀: 항상 가운데 정렬 (isHdr=true → C() 사용)
// 데이터 셀: C() 기본. 비고·개요·주요 내용·근거·리스크 등 문장형 장문 컬럼만 CL()/CLines()
const C = (text, width, isHdr=false, sz=24) => new TableCell({
  width:{size:width, type:WidthType.DXA},
  borders:{ top:OUTER, bottom:isHdr?HDR_BT:OUTER, left:OUTER, right:OUTER,
            insideH:INNER, insideV:INNER },
  verticalAlign: VerticalAlign.CENTER,
  children:[new Paragraph({
    alignment: AlignmentType.CENTER, spacing:{before:SP.cellV, after:SP.cellV},
    children: Array.isArray(text)
      ? text.map((s,i)=> new TextRun({text:s, font:F, color:BLACK, size:sz,
          bold: isHdr && i===0, ...(i>0?{break:1}:{}) }))
      : [new TextRun({text, font:F, color:BLACK, size:sz, bold:isHdr})]
  })]
});

const CL = (text, width, isHdr=false, sz=24) => new TableCell({
  width:{size:width, type:WidthType.DXA},
  borders:{ top:OUTER, bottom:isHdr?HDR_BT:OUTER, left:OUTER, right:OUTER,
            insideH:INNER, insideV:INNER },
  verticalAlign: VerticalAlign.CENTER,
  children:[new Paragraph({
    // 헤더 행(isHdr=true)은 내용 종류 무관 가운데 정렬 고정
    alignment: isHdr ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing:{before:SP.cellV, after:SP.cellV},
    indent: isHdr ? {} : {left:60},  // 헤더는 좌측 여백 없음, 데이터만 적용
    children: Array.isArray(text)
      ? text.map((s,i)=> new TextRun({text:s, font:F, color:BLACK, size:sz,
          bold: isHdr && i===0, ...(i>0?{break:1}:{}) }))
      : [new TextRun({text, font:F, color:BLACK, size:sz, bold:isHdr})]
  })]
});

const CP = (text, width) => C(text, width, false, 20);

// 명시적 개행 셀: 설명·비고 등 장문 컬럼에서만 사용
// 예: CLines(["주요 내용 본문 텍스트", "(보충설명이 길어지는 경우 개행 후 작성)"], w)
const CLines = (lines, width, sz=24) => CL(lines, width, false, sz);

const mkRow    = (cells) => new TableRow({cantSplit:true, children:cells});
const mkHdrRow = (cells) => new TableRow({tableHeader:true, cantSplit:true, children:cells});
const mkTable  = (hdrRow, dataRows, colWidths, tableW=TW) =>
  new Table({
    width:{size:tableW, type:WidthType.DXA}, columnWidths:colWidths,
    indent:{size:L2, type:WidthType.DXA},        // lv2 들여쓰기(142)와 정렬
    rows:[hdrRow, ...dataRows]
  });
```

**표 규칙 요약:**
- **헤더 행**: 모든 셀 `C(text, w, true)` 또는 `CL(text, w, true)` — 내용 종류 무관 가운데 정렬 고정
- **데이터 행**: `C()` 기본. 설명/비고/개요/주요 내용/근거/리스크 등 문장형 장문 컬럼이고 대체로 10자 초과 또는 개행 필요 시에만 `CL()`/`CLines()`
- **표 너비**: TW=8900 DXA (lv2 우측 끝 라인 이내)
- **표 들여쓰기**: `indent:{size:INDENT}` (lv2와 동일 좌측 정렬)
- **줄글+괄호 개행**: 괄호 포함 한 줄 초과 예상 시 배열로 `CLines()` 사용
- **표 내 줄글이 긴 경우**: 의미 단위에서 개행하여 셀 내 여러 줄로 적극 처리
- **표 앞 빈줄**: 절대 금지. 이전 단락(H2 또는 body)에 바로 붙임
- **표 뒤 여백**: defs에 tinyGap 삽입 금지. buildParagraphs가 다음 콘텐츠 레벨 기준으로 자동 처리
- **표 전체 한 페이지**: cantSplit:true + 표 직전 내용 분량으로 조절 (STEP 6 참조)

---

## STEP 8. buildParagraphs() + defs 예시

```javascript
// LV: 레벨 상수
const LV = { H1:1, H2:2, BODY:3, SUB:4, SUB2:5, NUM3:3, NUM4:4, GAP:0, TABLE:0 };
// SUB2: h4 두 번째 줄 continuation. 레벨값 5지만 after는 항상 SP.cont(6pt)
// prevLv를 갱신하지 않아 SUB2 다음 단락의 after 계산은 앞선 SUB 기준으로 유지

// ── 풋노트 after 역산 함수 ──────────────────────────────────────────────────
// 풋노트 before=1pt, 폰트=9pt → 이미 소비된 공간 10pt
// 다음 레벨 앞에 확보해야 할 총 pt에서 10pt 차감 후 +1pt
// 음수면 0처리 → 최소 1pt
// 이 계산이 다른 모든 after 규칙보다 최우선 적용
function fnAfterDxa(nextLv) {
  const rule = { 1: 24, 2: 20, 3: 12, 4: 9 }[nextLv] ?? 6;
  const remain = rule - 1 - 9;            // 1pt(before) + 9pt(폰트)
  return (Math.max(0, remain) + 1) * 20;  // 최소 1pt, DXA 변환
}
// ────────────────────────────────────────────────────────────────────────────

// ── afterDxa: 현재 단락의 after를 결정하는 핵심 함수 ────────────────────────
// 규칙: 오직 nextLv(다음 레벨)만 기준으로 after 결정
//   nextLv=1(h1) 또는 null(끝) → 24pt(SP.toH1)
//   nextLv=2(h2)               → 20pt(SP.toH2)
//   nextLv=3(h3)               → 12pt(SP.toH3)
//   nextLv=4(h4)               → 9pt(SP.toH4)
//
// ⚠️ prevLv 기반 6pt(SP.cont) 적용 금지
//   6pt는 오직 SUB2 단락(h4 개행 둘째줄)에만 명시적으로 적용됨
//   h3→h3, h4→h4 등 같은 레벨 연속이더라도 nextLv 기준으로만 계산
//   단, 바로 다음 항목이 풋노트(_isFn)이면 현재 단락 after=20(1pt) 우선 적용
function afterDxa(nextLv) {
  if (nextLv === null || nextLv === LV.H1) return SP.toH1;
  if (nextLv <= LV.H2)                     return SP.toH2;
  if (nextLv === LV.BODY)                  return SP.toH3;
  if (nextLv === LV.SUB)                   return SP.toH4;
  return SP.toH4; // fallback (SUB2 등)
}
// ────────────────────────────────────────────────────────────────────────────

// buildParagraphs: defs 배열 순회, after 값 자동 결정
// → Codex는 defs 배열 작성에만 집중, after 숫자는 직접 쓰지 않음
function buildParagraphs(defs) {
  const result = [];

  for (let i = 0; i < defs.length; i++) {
    const cur = defs[i];

    // 다음 실제 콘텐츠 레벨 탐색 (GAP/TABLE/SUB2 건너뜀)
    let nextLv = null;
    for (let j = i + 1; j < defs.length; j++) {
      const lv = defs[j].lv;
      if (lv > LV.GAP && lv !== LV.SUB2) { nextLv = lv; break; }
    }

    const nextIsFn = (i + 1 < defs.length && defs[i + 1]._isFn !== undefined);
    const curAfter = nextIsFn ? 20 : afterDxa(nextLv);

    if (cur.lv === LV.H1) {
      // h1→h1: 24pt / h1→h2: 20pt / h1→h3: 12pt / h1→h4: 9pt / h1→끝: 24pt
      result.push(h1(cur.text, curAfter));

    } else if (cur.lv === LV.H2) {
      result.push(h2(cur.text, curAfter));

    } else if (cur.lv === LV.BODY || cur.lv === LV.NUM3) {
      result.push(cur.lv === LV.NUM3 ? num3(cur.text, curAfter) : body(cur.text, curAfter));

    } else if (cur.lv === LV.SUB || cur.lv === LV.NUM4) {
      // 다음이 SUB2(개행 둘째줄)이면 현재 SUB의 after=SP.cont(6pt) 고정
      const nextIsSub2 = (i + 1 < defs.length && defs[i + 1].lv === LV.SUB2);
      const after = nextIsSub2 ? SP.cont : curAfter;
      result.push(cur.lv === LV.NUM4 ? num4(cur.text, after) : sub(cur.text, after));

    } else if (cur.lv === LV.SUB2) {
      // h4 개행 둘째줄: after는 nextLv 기준으로 정상 계산 (이 단락이 마지막 줄이므로)
      result.push(sub2(cur.text, curAfter));

    } else {
      // GAP / TABLE / 풋노트
      if (cur._isFn !== undefined) {
        // 풋노트: after는 fnAfterDxa()로 역산 — 다른 모든 규칙보다 최우선
        result.push(fn(cur._isFn, fnAfterDxa(nextLv)));
      } else if (cur.fn) {
        result.push(cur.fn());
        if (cur.lv === LV.TABLE) {
          const afterTableIsFn = (i + 1 < defs.length && defs[i + 1]._isFn !== undefined);
          if (!afterTableIsFn && nextLv !== null) result.push(tableGap(afterDxa(nextLv)));
        }
      }
    }
  }
  return result;
}

// ── defs 배열 작성법 ──
// text 필드: 단순 문자열 또는 TextRun 배열
// fn 필드: GAP/TABLE 등 헬퍼 직접 호출
// lv 필드: 레벨 상수 (after는 buildParagraphs가 자동 계산)
//
// ⚠️ defs 작성 핵심 규칙:
//   H1 앞뒤로 GAP 삽입 절대 금지 — buildParagraphs가 h1.after로 자동 처리
//   h1Sep() 을 H1 주변에 수동 삽입하면 간격 중복으로 빈줄처럼 보임
//   표 뒤 tinyGap() 수동 삽입 금지 — buildParagraphs가 다음 레벨 기준 자동 처리
//   텍스트 없는 빈 Paragraph는 표 뒤 자동 spacer 외 금지
//   H1 다음에는 바로 H2 또는 콘텐츠가 이어져야 함
//
// ── h1/h2/h3 한 줄 원칙 + 리듬 다양화 ──
//   목적에 맞는 핵심 판단문을 먼저 만들고, 한 줄 상한 안에서는 내용을 충분히 담는다.
//   줄 길이 제한(lv3 환산 25자, 리스크 구간 자간 적용)을 지킨다.
//   글자 수 하한 없음. 특정 짧은 길이에 맞추기 위한 인위적 압축 금지.
//   같은 맥락의 현황·원인·영향·판단은 짧은 여러 줄로 분해하지 않고 한 줄 안에서 결합한다.
//   lv3는 원인·영향·판단을 자연스럽게 포함한다.
//   lv2별 lv3 개수는 2~5개 범위에서 맥락에 따라 달리하며, 모든 lv2를 3개씩 맞추지 않는다.
//   렌더링 리스크 구간은 문장 수정 전에 자간 우선 적용
//   lv3 환산 26~27자: tcs(s, -8) / 28~29자: tcs(s, -12) / 30자↑ 표현 조정
//   너무 짧은 메모형 단락이 이어지면 한 줄 상한 안에서 원인·근거·판단을 결합
//
// ── h4(SUB) 두 줄 허용 시 ──
//   SUB(첫째줄) → SUB2(둘째줄) 분리
//   SUB2: left=L4, hanging 없음 → 기호 다음 텍스트 시작 위치(710 DXA)와 자동 정렬
//   SUB after=SP.cont(6pt) 자동 / SUB2 after=다음레벨기준 자동
//   세 줄 이상 금지 — SUB2는 반드시 1개만
//   SUB2 항목에는 풋노트 붙이지 않음
//   lv4는 모든 lv3에 붙이지 않고 근거·예시·리스크·후속조치가 필요한 곳에만 선택 사용
//
// ── 원형번호/※ ──
//   lv2로 사용 금지. 필요 위치에 따라 LV.NUM3 또는 LV.NUM4로 지정
//   NUM3는 body(-)와 동일한 들여쓰기/여백, NUM4는 sub(·)와 동일한 들여쓰기/여백을 따른다
//
// ── 풋노트 ──
//   after: fnAfterDxa()가 역산하여 자동 결정 (직접 지정 금지)
//   위치: 기본은 직전 문단 기호 시작점 + 2칸. 특정 단어·수치 근방이면 anchor 사용
//   예: { lv:LV.GAP, _isFn:{ anchor:"5.1%", text:"임금 인상률 산식 기준" } }
//   anchor는 위치 계산용이라 최종 출력은 "* 임금 인상률 산식 기준"만 표시
//   길이: 직전 줄(풋노트 윗줄) 글자 수(환산) × 2/3 이내 — 두 줄 방지
//         고정 18자 상한선 없음. 윗줄 28자 기준이면 약 19자 이내
//   필요한 경우에만 삽입 — 0개여도 무방, 억지로 채우지 않음 (STEP 10 참조)
//   배치: 위치가 아니라 필요성 기준. 보충설명이 필요한 줄 바로 아래에 빈줄 없이 배치
//   금지: 위치 균형을 맞추기 위한 억지 추가 또는 같은 설명 대상 중복 배치
//   SUB 두 줄(SUB+SUB2) 항목에는 풋노트 붙이지 않음
//   별첨은 본문 이해에 직접 필요한 경우에만 언급한다.
//   불필요하면 본문에서 언급하지 않는다.
//   언급 시 본문에서 "별첨"이라는 단어를 직접 쓰지 않는다.
//   별첨 제목만 본문에 쓰고 _isFn:{ anchor:"별첨 제목", text:"별첨1" } 형식으로 번호를 풋노트 처리
//   본문·풋노트 문체는 명사형·판단명사 종결 최우선
//   금지: ~이다, ~로 보인다, ~가 필요하다 등 문장완결형
//   차선: ~함, ~있음, ~됨, ~해야 함류는 가능하면 판단명사로 정리
//   주어+조사 단문은 명사구·관형절·상황절 등으로 다양하게 재구성하고 같은 패턴 반복 금지
//   예: "확인해야 함" 대신 "확인 필요", "포함함" 대신 "포함", "수치는 ... 기준" 대신 "... 기준"
const defs = [
  { lv: LV.H1,   text: "1. 개요" },
  { lv: LV.H2,   text: "□ 추진 배경" },
  { lv: LV.BODY, text: "- 주요 내용" },
  { lv: LV.NUM3, text: "① lv3 원형번호 항목" },
  { lv: LV.SUB,  text: "· 세부 내용이 길어서 한 줄을 넘어야 하는 경우" }, // after=6pt 자동
  { lv: LV.SUB2, text: "  이어지는 내용 (기호 없음, L4 위치에서 시작)" }, // after=다음레벨기준 자동
  { lv: LV.NUM4, text: "① lv4 원형번호 항목" },
  { lv: LV.SUB,  text: "· 일반 h4 항목" },
  { lv: LV.TABLE, fn:  () => mkTable(hdrRow, rows, cols) },
  { lv: LV.GAP,  _isFn: "필요한 줄 설명만 삽입" },
  { lv: LV.H1,   text: "2. 분석" },
];

const children = buildParagraphs(defs);
```

절대 금지: pageBreakBefore / keepNext / 수동 PageBreak (별첨 section 전환 제외)

---

---

## STEP 9. 별첨 규칙 + XML

별첨이 필요한 경우에만 `references/code-helpers-annex.md`를 읽고 적용한다.
별첨이 없으면 이 파일을 열지 않는다.
