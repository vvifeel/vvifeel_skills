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

// 1. xmlns 주입
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
