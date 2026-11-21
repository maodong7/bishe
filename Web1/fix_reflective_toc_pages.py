# -*- coding: utf-8 -*-
"""
Fix Reflective Report manual TOC page numbers: linear remap from current numeric range
onto 1..PMAX (default 24) so the last entry matches a short thesis, without editing body text.

Also re-applies tab + right dot leader for alignment.
"""
from __future__ import annotations

import re

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Inches

DOC_PATH = (
    r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文\个人部分2\毕业论文全套_生成"
    r"\10356_080901_24259701_BS_001_Reflective_Report.docx"
)

TAB_POS = Inches(6.55)
TOC_START, TOC_END = 34, 75
PMAX = 24  # your stated maximum page count for the whole document


def parse_toc_line(s: str) -> tuple[str, int] | None:
    s = s.strip()
    if "\t" not in s:
        return None
    a, b = s.split("\t", 1)
    try:
        return a.strip(), int(b.strip())
    except ValueError:
        return None


def main() -> None:
    doc = Document(DOC_PATH)
    rows: list[tuple[int, str, int]] = []  # (para_index, title_raw, old_page)

    for i in range(TOC_START, TOC_END + 1):
        parsed = parse_toc_line(doc.paragraphs[i].text)
        if parsed is None:
            continue
        title_raw, old_p = parsed
        rows.append((i, title_raw, old_p))

    old_pages = [r[2] for r in rows]
    old_min, old_max = min(old_pages), max(old_pages)
    denom = max(1, old_max - old_min)

    def remap(p: int) -> int:
        if p <= old_min:
            return old_min
        if p >= old_max:
            return PMAX
        return old_min + round((p - old_min) * (PMAX - old_min) / denom)

    # Ensure non-decreasing down the TOC list (rounding can create dips)
    last = 1
    updates: list[tuple[int, str, int]] = []
    for i, title_raw, old_p in rows:
        new_p = remap(old_p)
        new_p = max(last, min(PMAX, new_p))
        last = new_p
        updates.append((i, title_raw, new_p))

    for i, title_raw, new_p in updates:
        p = doc.paragraphs[i]
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(
            TAB_POS, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        p.text = f"{title_raw}\t{new_p}"

    # Refresh Note line so it does not claim 41-page draft
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("Note: Page numbers"):
            p.text = (
                "Note: Manual TOC page numbers were remapped to a "
                f"{PMAX}-page document span (from prior range {old_min}–{old_max}). "
                "Adjust PMAX in fix_reflective_toc_pages.py if your final PDF differs. "
                "In Word, switch to Print Layout and fine-tune any line after repagination."
            )
            break

    doc.save(DOC_PATH)
    print("Remapped TOC pages", old_min, "–", old_max, "→", old_min, "–", PMAX)
    print("Sample:", updates[0], updates[-1])


if __name__ == "__main__":
    main()
