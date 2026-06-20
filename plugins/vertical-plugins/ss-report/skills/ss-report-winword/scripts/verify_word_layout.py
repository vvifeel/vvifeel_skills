#!/usr/bin/env python3
"""
verify_word_layout.py — Word COM 기반 실제 페이지 레이아웃 검증
사전 조건: Windows + MS Word 설치 + pip install pywin32
"""
import sys
import os


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

        # ── 페이지별 단락 수집 ──────────────────────────────────────────────
        pages     = {}
        para_list = []
        for para in doc.Paragraphs:
            text = (para.Range.Text
                    .strip()
                    .replace('\x07', '')
                    .replace('\r', '')
                    .replace('\n', ''))
            if not text:
                continue
            page = int(para.Range.Information(3))  # wdActiveEndPageNumber
            pages.setdefault(page, []).append(text[:70])
            para_list.append({"text": text[:70], "page": page})

        for pg in sorted(pages.keys()):
            print(f"[P{pg}]")
            for item in pages[pg]:
                print(f"  {item}")
            print()

        _body_starts = ("-", "·", "①", "②", "③", "④", "⑤", "※")

        def is_h1(text):
            return bool(text) and text[0].isdigit() and len(text) > 2 and text[1] == '.'

        def is_h2(text):
            return bool(text) and text.startswith("□")

        def is_body(text):
            return bool(text) and any(text.startswith(s) for s in _body_starts)

        # ── 표 분리 검증 ────────────────────────────────────────────────────
        table_issues = []
        print("[표 검증]")
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
        print()

        # ── H2 섹션 페이지 걸침 탐지 ────────────────────────────────────────
        # 규칙: H2(□)와 그 아래 모든 H3/H4(lv3/lv4)는 같은 페이지 안에서 완결되어야 함
        # 탐지: H2 섹션이 여러 페이지에 걸치는 경우 → 내용 압축 또는 H2 구조 재편 필요
        h2_span_issues = []
        print("[H2 섹션 완결 검증]")

        current_h2_text       = None
        current_h2_start_page = None
        current_h2_end_page   = None

        def flush_h2():
            if current_h2_text and current_h2_end_page != current_h2_start_page:
                return (f"❌ H2 걸침: '{current_h2_text}' "
                        f"P{current_h2_start_page}→P{current_h2_end_page} "
                        f"(lv3/lv4가 다음 페이지로 넘어감)")
            return None

        for item in para_list:
            text = item["text"]
            page = item["page"]

            if is_h1(text):
                # H1 경계 — 현재 H2 마무리
                err = flush_h2()
                if err:
                    h2_span_issues.append(err)
                current_h2_text       = None
                current_h2_start_page = None
                current_h2_end_page   = None

            elif is_h2(text):
                # 새 H2 시작 — 이전 H2 마무리
                err = flush_h2()
                if err:
                    h2_span_issues.append(err)
                current_h2_text       = text[:50]
                current_h2_start_page = page
                current_h2_end_page   = page

            elif current_h2_text:
                # H2 하위 항목(H3/H4) — 페이지 범위 확장 추적
                if page > current_h2_end_page:
                    current_h2_end_page = page

        # 마지막 H2 처리
        err = flush_h2()
        if err:
            h2_span_issues.append(err)

        if h2_span_issues:
            for iss in h2_span_issues:
                print(f"  {iss}")
        else:
            print("  ✅ 모든 H2 섹션 단일 페이지 내 완결")
        print()

        # ── 페이지 시작 H3/H4 탐지 (H2 걸침의 증상 보조 확인) ──────────────
        page_first = {}
        for item in para_list:
            pg = item["page"]
            if pg not in page_first:
                page_first[pg] = item["text"]

        page_start_issues = []
        print("[페이지 시작 항목 검증]")
        for pg in sorted(page_first.keys()):
            if pg == 1:
                continue
            first = page_first[pg]
            if is_body(first):
                msg = f"❌ P{pg} 시작이 H3/H4: '{first[:50]}' (H2 없이 본문으로 시작)"
                page_start_issues.append(msg)
                print(f"  {msg}")
        if not page_start_issues:
            print("  ✅ 모든 페이지 H1 또는 H2로 시작")
        print()

        # ── H1/H2 고립 탐지 (페이지 말미 단독 헤더) ─────────────────────────
        orphan_issues = []
        print("[고립 탐지]")
        for i, item in enumerate(para_list[:-1]):
            next_item = para_list[i + 1]
            text = item["text"]
            if (is_h1(text) or is_h2(text)) and item["page"] != next_item["page"]:
                msg = f"⚠ P{item['page']} 말미 단독 헤더: '{text[:50]}'"
                orphan_issues.append(msg)
                print(f"  {msg}")
        if not orphan_issues:
            print("  ✅ 고립 헤더 없음")
        print()

        # ── 이슈 요약 ──────────────────────────────────────────────────────────
        all_issues = table_issues + h2_span_issues + page_start_issues + orphan_issues
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
