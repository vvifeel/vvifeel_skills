/**
 * fn_align.js — 풋노트 * 위치 자동 정렬 유틸리티
 *
 * 사용법:
 *   const { fnPrefix } = require('./scripts/fn_align');
 *   // defs 내 _isFn 값에 prefix 포함:
 *   { lv: LV.GAP, _isFn: `${fnPrefix(bodyStr, "HBM")}* HBM(High Bandwidth Memory): GPU에 적층하는 고대역폭 메모리` }
 *
 * * 위치 규칙:
 *   1순위: 설명 대상 용어·수치 아래 → fnPrefix(body, "설명대상")
 *   2순위: 앵커 없을 때만 → 기호+공백 1칸 제외 첫 글자 아래
 *   예) 용어: fnPrefix(body, "HBM") / 수치 출처: fnPrefix(body, "$")
 *
 * 원리:
 *   본문 텍스트(14pt 바탕체)에서 주석 대상 앞까지의 DXA를 추정,
 *   풋노트(9pt) 공백 1자 = FN_SPACE_DXA 기준으로 공백 수 산출.
 *   fn() 헬퍼는 note 텍스트가 이미 "* "를 포함하면 중복 삽입 없이 그대로 사용.
 */

// 14pt 바탕체 기준 글자폭 DXA 추정
const BODY_DXA = {
  K: 280,   // 한글 (14pt × 20)
  E: 154,   // 영문 (~0.55 × 280)
  D: 140,   // 숫자
  S: 120,   // 공백
  X: 90,    // 특수문자 (, . ' % $ 등)
};

// 9pt 바탕체 공백 1자 DXA 추정
const FN_SPACE_DXA = 90;

// 최대 공백 수 (지나치게 긴 prefix 방지)
const MAX_SPACES = 28;

function _charDxa(ch) {
  const code = ch.codePointAt(0);
  if (code >= 0xAC00 && code <= 0xD7A3) return BODY_DXA.K;
  if (ch === ' ')                        return BODY_DXA.S;
  if (ch >= '0' && ch <= '9')           return BODY_DXA.D;
  if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')) return BODY_DXA.E;
  return BODY_DXA.X;
}

/**
 * 본문 텍스트에서 term 위치까지의 prefix 공백 문자열 반환
 * @param {string} bodyText - 본문 전체 문자열 (예: "- '26년 HBM 시장 $54.6B...")
 * @param {string} term     - 주석 대상 용어 (예: "HBM")
 * @returns {string} - 공백 문자열 (fn() 앞에 붙여 * 위치 조정용)
 */
function fnPrefix(bodyText, term) {
  if (!bodyText || !term) return '';
  const idx = bodyText.indexOf(term);
  if (idx < 0) return '';

  const prefixDxa = [...bodyText.slice(0, idx)].reduce((sum, ch) => sum + _charDxa(ch), 0);
  const spaces    = Math.min(MAX_SPACES, Math.round(prefixDxa / FN_SPACE_DXA));
  return ' '.repeat(spaces);
}

/**
 * _isFn 텍스트 앞부분에서 주석 대상 영문 약어 추출 (자동 term 탐지용)
 * @param {string} fnText - 예: "HBM(High Bandwidth Memory): ..."
 * @returns {string|null}
 */
function extractFnTerm(fnText) {
  const m = fnText.match(/^([A-Z][A-Za-z0-9]+)/);
  return m ? m[1] : null;
}

module.exports = { fnPrefix, extractFnTerm };
