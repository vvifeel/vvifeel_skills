# code-helpers-annex.md — 별첨 전용 헬퍼·XML

> 별첨이 필요한 경우에만 읽는다. 일반 보고서는 code-helpers.md만 사용한다.

---

## STEP 9. 별첨 규칙 + XML

### 별첨 생성 기준

```
아래 조건 중 하나라도 해당하면 별첨으로 분리:
  ① 표가 본문 반 페이지(약 11줄)를 초과할 정도로 내용이 많은 경우
  ② 본문 주요 흐름과 결이 달라 참고자료로 제공할 의미가 있는 경우
     (상세 내역, 부가 데이터, 비교표 등)

해당하지 않는 경우 별첨 생성 금지 — 본문에 포함
```

### 별첨 구성 원칙

```
- 별첨은 반드시 새 페이지에서 시작, 본문과 완전 분리
- 구현: Document.sections 배열 분리 방식 (PageBreak 단락 방식 금지)
  → 본문 = section[0] / 별첨1 = section[1] / 별첨2 = section[2] ...
  → Word가 section 경계를 항상 새 페이지로 처리 — 라이브러리 버전 무관하게 100% 보장
- 별첨 1개 = 반드시 정확히 1페이지. 내용이 넘치면 표·내용 축소
- 내용: 본문에 담지 못한 상세 표·상세 내역 위주
- 각 별첨 section 첫 단락: annexMarker(n) — adm-zip 패치로 텍스트박스 교체

텍스트 기반 별첨 (표가 아닌 텍스트 내용):
  annexTitlePara()가 h1 역할 → 별첨 본문은 LV.H2(□)부터 시작
  LV.H1 사용 금지 — 별첨 내부에 h1() 삽입 금지
  예)
    const annex1Children = [
      annexMarker(1),
      annexTitlePara("별첨 제목"),           // h1 역할
      ...buildParagraphs(annexDefs),          // annexDefs: LV.H2부터 시작
    ];
    const annexDefs = [
      { lv: LV.H2,   text: "□ 소제목" },    // ← LV.H1 금지
      { lv: LV.BODY, text: "- 내용" },
      ...
    ];
```

### Section 분리 + annexMarker 구현

```javascript
const AdmZip = require('adm-zip');

// ── 별첨 제목 ──
// 반드시 한 줄. 한글 18자 / 환산 22자 초과 금지. 풋노트도 없음.
const annexTitlePara = (s) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 480 },   // 24pt
  children: [new TextRun({ text: s, font: F, size: 40, bold: true, color: BLACK, underline: {} })]
});

// ── 별첨 마커: 흰색 텍스트 — XML 탐색용, 출력물에 보이지 않음 ──
// ⚠️ size: 1은 docx 라이브러리가 제거할 수 있음 → size: 20(10pt)으로 고정
// ⚠️ 반드시 section[1~]의 첫 번째 단락으로만 사용
//    본문 section[0] children에 절대 포함하지 말 것 → 본문 첫 페이지에 텍스트박스 붙는 원인
const annexMarker = (n) => new Paragraph({
  spacing: { before: 0, after: 0 },
  children: [new TextRun({ text: `__ANNEX_LABEL_${n}__`, font: F, size: 20, color: "FFFFFF" })]
});

// ── Document 구성 순서 (STRICT) ──────────────────────────────────────────
// ⚠️ 반드시 아래 순서를 지킬 것. doc 생성 후 배열 수정(unshift 등)은 반영 안 됨.
//
// 1단계: 본문 단락 배열 완성 (titlePara/datePara 포함)
const bodyChildren = [
  titlePara("보고서 제목"),   // ← 반드시 여기서 선언
  datePara("26.05.16(금)"),
  ...buildParagraphs(defs),  // spread로 펼쳐서 추가
];

// 2단계: 별첨 단락 배열 완성 (annexMarker가 반드시 첫 항목)
const annex1Children = [
  annexMarker(1),             // ← 반드시 section[1]의 첫 번째 단락
  annexTitlePara("AI 포트폴리오 투자 내역"),  // 반드시 18자 이내
  mkTable(hdrRow1, dataRows1, cols1),
];
// 별첨 2 필요 시:
// const annex2Children = [ annexMarker(2), annexTitlePara("제목"), mkTable(...) ];

// 3단계: Document 생성 (배열이 완성된 후)
const doc = new Document({
  // 머릿글·바닥글 절대 생성 금지 — headers/footers 키 자체를 넣지 않을 것
  sections: [
    { children: bodyChildren },      // 본문: section[0]
    { children: annex1Children },    // 별첨 1: section[1] → 자동 새 페이지
    // { children: annex2Children }, // 별첨 2: section[2] → 별첨 있을 때만 추가
  ]
});
// ⚠️ doc 생성 이후 bodyChildren/annex1Children 수정(push/unshift 등) 절대 금지
// ────────────────────────────────────────────────────────────────────────────
```

### patch_annex.js — 외부 스크립트 (Codex가 수정 금지)

```
별첨 패치는 buildAndPatch 코드를 report.js에 직접 쓰지 않는다.
반드시 외부 스크립트 scripts/patch_annex.js 를 호출한다.

이유:
  - buildAndPatch 내부 XML 조작은 재생성 시 버그가 반복 발생함
  - 외부 스크립트로 분리하면 Codex가 손댈 수 없어 버그가 고정 방어됨
  - 인수(별첨 개수)만 바꿔서 모든 보고서에 재사용 가능

파이프라인 (별첨 있을 때):
  node report.js                                   → report-base.docx
  node scripts/patch_annex.js report-base.docx 1  → report-patched.docx
  python inject_footnotes.py report-patched.docx report-final.docx
  python verify_report.py report-final.docx

  두 번째 인수: 별첨 개수 (별첨 2개면 → node scripts/patch_annex.js report-base.docx 2)

report.js에서 할 일:
  - Packer.toBuffer(doc) 로 base docx 생성만 담당
  - AdmZip / buildAndPatch / annexLabelXml 코드를 report.js에 작성하지 않음
  - buildAndPatch 함수 생성 절대 금지
```

patch_annex.js 파일 내용 (scripts/ 디렉토리에 고정 배치):

```javascript
#!/usr/bin/env node
// scripts/patch_annex.js — 별첨 레이블 텍스트박스 패치 스크립트
// 사용법: node scripts/patch_annex.js <input.docx> <annexCount> [output.docx]
// Codex가 이 파일을 수정하지 말 것 — 수정 시 별첨 버그 재발

const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

const [,, inputPath, countArg, outputPath] = process.argv;
if (!inputPath || !countArg) {
  console.error('Usage: node patch_annex.js <input.docx> <annexCount> [output.docx]');
  process.exit(1);
}
const annexCount = parseInt(countArg, 10);
const outPath = outputPath || inputPath.replace(/(-base)?\.docx$/, '-patched.docx');

// ── 별첨 레이블 박스 고정 상수 ──────────────────────────────────────────────
const AX_CX    = 954000;
const AX_CY    = 367200;
const AX_POSH  = -7200;
const AX_POSV  = -381600;
const AX_LRINS = 90000;
const AX_TBINS = 46800;
// ─────────────────────────────────────────────────────────────────────────────

function annexLabelXml(n) {
  return `<w:p>  <w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/></w:pPr>  <w:r><w:rPr><w:noProof/></w:rPr>  <mc:AlternateContent>    <mc:Choice Requires="wps">      <w:drawing>        <wp:anchor distT="0" distB="0" distL="0" distR="0"          simplePos="0" relativeHeight="251658240" behindDoc="0"          locked="0" layoutInCell="1" allowOverlap="1" moveWithText="1">          <wp:simplePos x="0" y="0"/>          <wp:positionH relativeFrom="margin"><wp:posOffset>${AX_POSH}</wp:posOffset></wp:positionH>          <wp:positionV relativeFrom="paragraph"><wp:posOffset>${AX_POSV}</wp:posOffset></wp:positionV>          <wp:extent cx="${AX_CX}" cy="${AX_CY}"/>          <wp:wrapNone/>          <wp:docPr id="${10 + n}" name="AnnexLabel${n}"/>          <wp:cNvGraphicFramePr/>          <a:graphic>            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">              <wps:wsp>                <wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>                <wps:spPr>                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="${AX_CX}" cy="${AX_CY}"/></a:xfrm>                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>                  <a:noFill/>                  <a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>                </wps:spPr>                <wps:txbx>                  <w:txbxContent>                    <w:p>                      <w:pPr>                        <w:jc w:val="center"/>                        <w:spacing w:before="0" w:after="0"/>                        <w:rPr>                          <w:rFonts w:ascii="바탕체" w:hAnsi="바탕체" w:eastAsia="바탕체" w:cs="바탕체"/>                          <w:sz w:val="28"/><w:szCs w:val="28"/>                          <w:color w:val="000000"/>                        </w:rPr>                      </w:pPr>                      <w:r>                        <w:rPr>                          <w:rFonts w:ascii="바탕체" w:hAnsi="바탕체" w:eastAsia="바탕체" w:cs="바탕체"/>                          <w:sz w:val="28"/><w:szCs w:val="28"/>                          <w:color w:val="000000"/>                        </w:rPr>                        <w:t>별첨 ${n}</w:t>                      </w:r>                    </w:p>                  </w:txbxContent>                </wps:txbx>                <wps:bodyPr lIns="${AX_LRINS}" rIns="${AX_LRINS}"                            tIns="${AX_TBINS}" bIns="${AX_TBINS}"                            anchor="ctr"><a:normAutofit/></wps:bodyPr>              </wps:wsp>            </a:graphicData>          </a:graphic>        </wp:anchor>      </w:drawing>    </mc:Choice>  </mc:AlternateContent>  </w:r></w:p>`;
}

const zip = new AdmZip(inputPath);
let xml = zip.readAsText('word/document.xml');

// 1. xmlns 주입 — 'xmlns:X=' (등호 포함) 으로 체크하여 부분매칭 방지
const nsMap = {
  'xmlns:mc=':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'xmlns:wp=':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'xmlns:a=':   'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:wps=': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
};
const docTagStart = xml.indexOf('<w:document ');
const docTagEnd   = xml.indexOf('>', docTagStart);
if (docTagStart >= 0 && docTagEnd > docTagStart) {
  const attrs = xml.slice(docTagStart, docTagEnd);
  let extra = '';
  for (const [ns, uri] of Object.entries(nsMap)) {
    if (!attrs.includes(ns)) extra += ` ${ns.slice(0, -1)}="${uri}"`;
  }
  if (extra) xml = xml.slice(0, docTagEnd) + extra + xml.slice(docTagEnd);
}

// 2. &apos; 복원
xml = xml.replace(/&apos;/g, "'");

// 3. 마커 → 텍스트박스 교체
// '<w:p ' 과 '<w:p>' 로 정확히 탐색하여 <w:pPr> 오매칭 방지
for (let n = 1; n <= annexCount; n++) {
  const marker = `__ANNEX_LABEL_${n}__`;
  const markerIdx = xml.indexOf(marker);
  if (markerIdx < 0) {
    console.warn(`[patch_annex] 별첨 마커 ${n} 미발견 — annexMarker(${n}) 누락 확인 필요`);
    continue;
  }
  const pStart1 = xml.lastIndexOf('<w:p ', markerIdx);
  const pStart2 = xml.lastIndexOf('<w:p>', markerIdx);
  const pStart  = Math.max(pStart1, pStart2);
  const pEnd    = xml.indexOf('</w:p>', markerIdx) + '</w:p>'.length;
  if (pStart >= 0 && pEnd > pStart) {
    xml = xml.slice(0, pStart) + annexLabelXml(n) + xml.slice(pEnd);
    console.log(`[patch_annex] 별첨 ${n} 텍스트박스 삽입 완료`);
  } else {
    console.warn(`[patch_annex] 별첨 마커 ${n} 단락 경계 탐색 실패`);
  }
}

zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
zip.writeZip(outPath);
console.log(`[patch_annex] 완료: ${outPath}`);
```

### annexLabelXml + buildAndPatch (레거시 참조용)

```javascript
// ── 별첨 레이블 박스 고정 상수 ─────────────────────────────────────────────
// ⚠️ 아래 6개 상수는 절대 수정 금지. 값을 바꾸면 Word에서 박스 크기/위치가 틀어짐.
// ⚠️ cx와 cy를 절대 swap하지 말 것 — cx=너비(가로), cy=높이(세로)
const AX_CX    = 954000;   // 너비  2.65cm
const AX_CY    = 367200;   // 높이  1.02cm
const AX_POSH  = -7200;    // 가로위치 -0.02cm (기준: margin 오른쪽)
const AX_POSV  = -381600;  // 세로위치 -1.06cm (기준: paragraph 아래쪽)
const AX_LRINS = 90000;    // 좌우 내부여백 0.25cm
const AX_TBINS = 46800;    // 상하 내부여백 0.13cm
// ────────────────────────────────────────────────────────────────────────────

// ── inlaineXmlns 제거 → buildAndPatch에서 루트에 한 번만 주입 ──
// 이유: 인라인 xmlns 중복 선언이 Word XML 파서를 손상시켜 파일 열기 오류 발생
function annexLabelXml(n) {
  // xmlns 선언 없음 — 루트 <w:document>에서 이미 선언됨 (buildAndPatch가 처리)
  // wp:extent와 a:xfrm 양쪽에 반드시 AX_CX/AX_CY를 그대로 사용할 것
  return `<w:p>
  <w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/></w:pPr>
  <w:r><w:rPr><w:noProof/></w:rPr>
  <mc:AlternateContent>
    <mc:Choice Requires="wps">
      <w:drawing>
        <wp:anchor distT="0" distB="0" distL="0" distR="0"
          simplePos="0" relativeHeight="251658240" behindDoc="0"
          locked="0" layoutInCell="1" allowOverlap="1" moveWithText="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="margin"><wp:posOffset>${AX_POSH}</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>${AX_POSV}</wp:posOffset></wp:positionV>
          <wp:extent cx="${AX_CX}" cy="${AX_CY}"/>
          <wp:wrapNone/>
          <wp:docPr id="${10 + n}" name="AnnexLabel${n}"/>
          <wp:cNvGraphicFramePr/>
          <a:graphic>
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp>
                <wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>
                <wps:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="${AX_CX}" cy="${AX_CY}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  <a:noFill/>
                  <a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                </wps:spPr>
                <wps:txbx>
                  <w:txbxContent>
                    <w:p>
                      <w:pPr>
                        <w:jc w:val="center"/>
                        <w:spacing w:before="0" w:after="0"/>
                        <w:rPr>
                          <w:rFonts w:ascii="바탕체" w:hAnsi="바탕체"
                                    w:eastAsia="바탕체" w:cs="바탕체"/>
                          <w:sz w:val="28"/><w:szCs w:val="28"/>
                          <w:color w:val="000000"/>
                        </w:rPr>
                      </w:pPr>
                      <w:r>
                        <w:rPr>
                          <w:rFonts w:ascii="바탕체" w:hAnsi="바탕체"
                                    w:eastAsia="바탕체" w:cs="바탕체"/>
                          <w:sz w:val="28"/><w:szCs w:val="28"/>
                          <w:color w:val="000000"/>
                        </w:rPr>
                        <w:t>별첨 ${n}</w:t>
                      </w:r>
                    </w:p>
                  </w:txbxContent>
                </wps:txbx>
                <wps:bodyPr lIns="${AX_LRINS}" rIns="${AX_LRINS}"
                            tIns="${AX_TBINS}" bIns="${AX_TBINS}"
                            anchor="ctr"><a:normAutofit/></wps:bodyPr>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </mc:Choice>
  </mc:AlternateContent>
  </w:r></w:p>`;
}

// ── buildAndPatch: 네임스페이스 루트 주입 + 마커 교체 ──
// 핵심: <w:document> 루트 태그에 필요한 xmlns를 한 번만 선언
//       annexLabelXml은 인라인 xmlns 없이 이 선언에 의존
async function buildAndPatch(doc, annexCount) {
  const buffer = await Packer.toBuffer(doc);
  const zip = new AdmZip(buffer);
  let xml = zip.readAsText('word/document.xml');

  // 1. 루트 태그에 필요한 xmlns 주입
  // Fix①: includes('xmlns:a') 는 'xmlns:aink' 같은 속성과 부분매칭 → 'xmlns:a=' 로 체크
  const nsMap = {
    'xmlns:mc=':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns:wp=':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'xmlns:a=':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xmlns:wps=': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  };
  // Fix⑤: 정규식 대신 indexOf로 w:document 태그 정확히 탐색
  //        (정규식은 XML 선언부 ?> 와 혼동 가능)
  const docTagStart = xml.indexOf('<w:document ');
  const docTagEnd   = xml.indexOf('>', docTagStart);
  if (docTagStart >= 0 && docTagEnd > docTagStart) {
    const attrs = xml.slice(docTagStart, docTagEnd);
    let extra = '';
    for (const [ns, uri] of Object.entries(nsMap)) {
      if (!attrs.includes(ns)) extra += ` ${ns.slice(0,-1)}="${uri}"`;
    }
    if (extra) xml = xml.slice(0, docTagEnd) + extra + xml.slice(docTagEnd);
  }

  // 2. &apos; → ' 복원
  xml = xml.replace(/&apos;/g, "'");

  // 3. 마커 단락 → 텍스트박스 XML 교체
  // Fix③: markerIdx < 0 이면 경고 출력 후 skip (음수 넘기면 lastIndexOf 오동작)
  // Fix④: '<w:p ' 과 '<w:p>' 둘 다 탐색해 더 가까운 것 사용 (<w:pPr> 오매칭 방지)
  for (let n = 1; n <= annexCount; n++) {
    const marker = `__ANNEX_LABEL_${n}__`;
    const markerIdx = xml.indexOf(marker);
    if (markerIdx < 0) {
      console.warn(`[buildAndPatch] 별첨 마커 ${n} 미발견 — annexMarker(${n}) 누락 또는 size가 너무 작아 라이브러리가 제거했을 가능성`);
      continue;
    }
    const pStart1 = xml.lastIndexOf('<w:p ', markerIdx);
    const pStart2 = xml.lastIndexOf('<w:p>', markerIdx);
    const pStart  = Math.max(pStart1, pStart2);
    const pEnd    = xml.indexOf('</w:p>', markerIdx) + '</w:p>'.length;
    if (pStart >= 0 && pEnd > pStart) {
      xml = xml.slice(0, pStart) + annexLabelXml(n) + xml.slice(pEnd);
    } else {
      console.warn(`[buildAndPatch] 별첨 마커 ${n} 단락 경계 탐색 실패`);
    }
  }

  zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
  return zip.toBuffer();
}

// ── 메인 실행 ──
(async () => {
  const annexCount = 1;   // 별첨 개수에 맞게 설정
  const buffer = await buildAndPatch(doc, annexCount);
  fs.writeFileSync('report-base.docx', buffer);
})();
```
