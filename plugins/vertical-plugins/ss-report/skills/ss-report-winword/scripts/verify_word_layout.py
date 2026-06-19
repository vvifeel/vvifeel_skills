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

        # ── H1/H2 고립 탐지 ──────────────────────────────────────────────────
        # 페이지 말미에 단독 위치한 헤더성 단락 (기호 없고 짧음)
        orphan_issues = []
        print("[고립 탐지]")
        _body_starts = ("-", "·", "①", "②", "③", "④", "⑤", "※")
        for i, item in enumerate(para_list[:-1]):
            next_item = para_list[i + 1]
            text = item["text"]
            is_header_like = (
                text
                and not any(text.startswith(s) for s in _body_starts)
                and len(text) < 35
            )
            if is_header_like and item["page"] != next_item["page"]:
                msg = f"⚠ P{item['page']} 말미 단독 항목: '{text}'"
                orphan_issues.append(msg)
                print(f"  {msg}")

        if not orphan_issues:
            print("  ✅ 고립 항목 없음")
        print()

        # ── 이슈 요약 ──────────────────────────────────────────────────────────
        all_issues = table_issues + orphan_issues
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
