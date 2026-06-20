#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_word_layout.py v2 — Word COM 기반 실제 페이지 레이아웃 검증
사전 조건: Windows + MS Word 설치 + pip install pywin32

검증 항목:
  [표 분리]    표가 페이지 경계에서 잘리는지 (cantSplit 보완 검증)
  [셀 줄바꿈]  헤더 셀 줄바꿈, 데이터 셀 마지막 줄 단독(≤2자) 탐지
  [고립 탐지]  P2+ 첫 줄 body 항목 시작 / H2+body 1줄만 페이지 말미 고립
"""
import sys
import io
import os
import re

# Windows cp949 터미널에서 이모지/한글 출력 보장
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'buffer'):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def classify_para(text):
    """단락 유형 분류: H1 / H2 / BODY / FN / OTHER"""
    if not text:
        return "EMPTY"
    if text.startswith("*"):
        return "FN"
    if text.startswith("□"):
        return "H2"
    if re.match(r"^\d+\.", text):
        return "H1"
    _body_starts = ("-", "·", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "※")
    if any(text.startswith(s) for s in _body_starts):
        return "BODY"
    return "OTHER"


def main():
    if len(sys.argv) < 2:
        print("사용법: python verify_word_layout.py <docx_경로>")
        sys.exit(1)

    docx_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(docx_path):
        print(f"❌ 파일 없음: {docx_path}")
        sys.exit(1)

    try:
        import win32com.client
    except ImportError:
        print("❌ pywin32 미설치 — pip install pywin32 후 재실행")
        sys.exit(1)

    word = None
    doc  = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible       = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(
            docx_path,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        # 레이아웃 강제 계산 (Print Layout 뷰 전환 후 Repaginate)
        word.ActiveDocument.ActiveWindow.View.Type = 3  # wdPrintView
        doc.Repaginate()

        total_pages = doc.ComputeStatistics(2)  # wdStatisticPages

        print(f"=== Word 레이아웃 검증: {os.path.basename(docx_path)} ===")
        print(f"총 페이지: {total_pages}")
        print()

        # ── 단락 수집 (페이지 + 표 여부 포함) ─────────────────────────────────
        # pages_display: page → [텍스트]  (출력용)
        # pages_meta:    page → [{text, in_table}]  (분석용)
        pages_display = {}
        pages_meta    = {}

        for para in doc.Paragraphs:
            text = (para.Range.Text
                    .strip()
                    .replace("\x07", "")
                    .replace("\r", "")
                    .replace("\n", ""))
            if not text:
                continue
            page = int(para.Range.Information(3))  # wdActiveEndPageNumber
            try:
                in_table = bool(para.Range.Information(12))  # wdWithInTable
            except Exception:
                in_table = False

            pages_display.setdefault(page, []).append(text[:70])
            pages_meta.setdefault(page, []).append({"text": text[:70], "in_table": in_table})

        for pg in sorted(pages_display.keys()):
            print(f"[P{pg}]")
            for item in pages_display[pg]:
                print(f"  {item}")
            print()

        all_issues = []

        # ── 표 분리 검증 ────────────────────────────────────────────────────
        table_issues = []
        print("[표 분리]")
        if doc.Tables.Count == 0:
            print("  (표 없음)")
        for t_idx in range(doc.Tables.Count):
            table = doc.Tables(t_idx + 1)
            try:
                first_page = int(table.Cell(1, 1).Range.Information(3))
                last_row   = table.Rows.Count
                last_col   = table.Columns.Count
                last_page  = int(table.Cell(last_row, last_col).Range.Information(3))
                if first_page != last_page:
                    msg = f"❌ 표{t_idx+1}: P{first_page}~P{last_page} 분리됨"
                    table_issues.append(msg)
                else:
                    msg = f"✅ 표{t_idx+1}: P{first_page} (분리 없음)"
                print(f"  {msg}")
            except Exception as e:
                print(f"  ⚠ 표{t_idx+1}: 확인 실패 ({e})")
        all_issues.extend(table_issues)
        print()

        # ── 표 셀 줄바꿈 탐지 ───────────────────────────────────────────────
        # 헤더 셀 줄바꿈(❌) / 데이터 셀 마지막 줄 단독 ≤2자(⚠)
        cell_issues = []
        print("[셀 줄바꿈]")
        for t_idx in range(doc.Tables.Count):
            table = doc.Tables(t_idx + 1)
            for r_idx in range(table.Rows.Count):
                is_hdr = (r_idx == 0)
                for c_idx in range(table.Columns.Count):
                    try:
                        cell = table.Cell(r_idx + 1, c_idx + 1)
                        cell_text = (cell.Range.Text
                                     .replace("\r", "").replace("\x07", "")
                                     .replace("\n", "").strip())
                        if not cell_text:
                            continue
                        try:
                            lines_col = cell.Range.Lines
                            n_lines   = lines_col.Count
                            if n_lines > 1:
                                if is_hdr:
                                    msg = (f"❌ 표{t_idx+1} {r_idx+1}행{c_idx+1}열 "
                                           f"헤더 줄바꿈({n_lines}줄): '{cell_text[:20]}'")
                                    cell_issues.append(msg)
                                    print(f"  {msg}")
                                else:
                                    last_text = (lines_col(n_lines).Range.Text
                                                 .replace("\r", "").replace("\x07", "")
                                                 .replace("\n", "").strip())
                                    if last_text and len(last_text) <= 2:
                                        msg = (f"⚠ 표{t_idx+1} {r_idx+1}행{c_idx+1}열 "
                                               f"마지막 줄 단독({len(last_text)}자): "
                                               f"'{last_text}' ← '{cell_text[:25]}'")
                                        cell_issues.append(msg)
                                        print(f"  {msg}")
                        except Exception:
                            pass  # Range.Lines 미지원 셀 스킵
                    except Exception:
                        pass
        if not cell_issues:
            print("  ✅ 셀 줄바꿈 이상 없음")
        all_issues.extend(cell_issues)
        print()

        # ── 고립 탐지 ────────────────────────────────────────────────────────
        # 규칙 1: P2+ 첫 비(非)표 단락이 BODY/FN → ❌  (H2 없이 body로 시작)
        # 규칙 2: 페이지 말미 마지막 비표 단락이 H1/H2/OTHER(짧음) → ⚠
        # 규칙 3: 페이지 말미 [H2, BODY]만 남고 다음 페이지도 BODY로 시작 → ❌
        orphan_issues = []
        print("[고립 탐지]")
        sorted_pgs = sorted(pages_meta.keys())

        for idx, pg in enumerate(sorted_pgs):
            body_items = [x for x in pages_meta[pg] if not x["in_table"]]
            next_pg    = sorted_pgs[idx + 1] if idx + 1 < len(sorted_pgs) else None
            next_body  = ([x for x in pages_meta[next_pg] if not x["in_table"]]
                          if next_pg else [])

            # 규칙 1: 페이지 상단 body 금지 (P2 이후)
            if pg > 1 and body_items:
                first_type = classify_para(body_items[0]["text"])
                if first_type in ("BODY", "FN"):
                    msg = (f"❌ P{pg} 첫 줄 body 항목 — "
                           f"이전 H2 아래 항목 부족: '{body_items[0]['text'][:30]}'")
                    orphan_issues.append(msg)
                    print(f"  {msg}")

            if not body_items or not next_body:
                continue

            last      = body_items[-1]
            last_type = classify_para(last["text"])

            # 규칙 2: 단독 헤더성 항목 페이지 말미
            if last_type in ("H1", "H2", "OTHER") and len(last["text"]) < 35:
                msg = f"⚠ P{pg} 말미 단독 항목: '{last['text']}'"
                orphan_issues.append(msg)
                print(f"  {msg}")
                continue  # 규칙 3 중복 탐지 방지

            # 규칙 3: H2 + body 1줄만 → 다음 페이지 body 연속
            if (len(body_items) >= 2
                    and classify_para(body_items[-2]["text"]) == "H2"
                    and last_type == "BODY"
                    and classify_para(next_body[0]["text"]) == "BODY"):
                msg = (f"❌ P{pg} 말미 H2+body 1줄 고립 "
                       f"(다음 페이지 body 연속): '{body_items[-2]['text'][:30]}'")
                orphan_issues.append(msg)
                print(f"  {msg}")

        if not orphan_issues:
            print("  ✅ 고립 항목 없음")
        all_issues.extend(orphan_issues)
        print()

        # ── 이슈 요약 ────────────────────────────────────────────────────────
        if all_issues:
            print(f"[이슈 요약] {len(all_issues)}건 발견 → 수정 후 재생성 필요")
            for iss in all_issues:
                print(f"  {iss}")
            sys.exit(1)
        else:
            print("[이슈 요약] ✅ 없음 — Word 렌더링 레이아웃 정상")

    except Exception as e:
        print(f"❌ Word COM 오류: {e}")
        sys.exit(1)
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
