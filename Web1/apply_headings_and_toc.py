# -*- coding: utf-8 -*-
"""
1) Apply heading hierarchy in body: outline level + bold + font size (this doc's styles
   cannot assign built-in 'Heading 1' by name via python-docx).
2) Rebuild manual TOC with indent by level + tab/dot leader + position-based pages 1..PMAX.

Does not change paragraph text except formatting and TOC lines + Note.
"""
from __future__ import annotations

import unicodedata as ud
from typing import Callable

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

DOC_PATH = (
    r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文\个人部分2\毕业论文全套_生成"
    r"\10356_080901_24259701_BS_001_Reflective_Report.docx"
)

NOTE_PREFIX = "Note:"
TOC_FIRST = 34
PMAX = 24
TAB_POS = Inches(6.55)


def norm(s: str) -> str:
    s = ud.normalize("NFKC", s).strip()
    for a, b in (("\u2013", "-"), ("\u2014", "-"), ("\u2212", "-")):
        s = s.replace(a, b)
    return s


def set_outline_level(paragraph, outline_zero_based: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    for old in p_pr.findall(qn("w:outlineLvl")):
        p_pr.remove(old)
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(outline_zero_based))
    p_pr.append(el)


def apply_visual_heading(paragraph, level: int) -> None:
    """level 1..3 -> outline 0..2, sizes 14/12/11 pt."""
    text = paragraph.text
    paragraph.clear()
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt((14, 12, 11)[level - 1])
    run.font.name = "Times New Roman"
    paragraph.paragraph_format.space_before = Pt((12, 8, 6)[level - 1])
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.keep_with_next = True
    set_outline_level(paragraph, level - 1)


def body_index_for_toc_title(body_norm: list[str], note_idx: int, title_n: str) -> int | None:
    special: dict[str, Callable[[str], bool]] = {
        norm("Strategy and regression focus"): lambda b: b.lower().startswith("e.1"),
        norm("Sample matrix and environment"): lambda b: b.lower().startswith("e.2"),
    }
    if title_n in special:
        pred = special[title_n]
        for j in range(note_idx + 1, len(body_norm)):
            if pred(body_norm[j]):
                return j
        return None
    alias = {
        norm("Inference and quality gate"): norm("D.1 Inference and quality gate"),
        norm("WebSocket batching and overlays"): norm(
            "D.2 WebSocket loop and persistence batching"
        ),
        norm("REST surface (representative)"): norm("D.3 Representative REST surface"),
    }.get(title_n, title_n)
    want = alias
    for j in range(note_idx + 1, len(body_norm)):
        b = body_norm[j]
        if b == want or b == title_n:
            return j
        if len(want) >= 14 and want in b:
            return j
        if len(title_n) >= 14 and title_n in b:
            return j
    return None


def heading_level_for_para_idx(idx: int) -> int | None:
    h1 = {
        19,
        24,
        33,
        78,
        90,
        104,
        118,
        129,
        135,
        168,
        179,
        188,
        197,
        209,
        220,
        222,
    }
    h2 = {
        79,
        82,
        85,
        88,
        91,
        94,
        96,
        98,
        100,
        102,
        105,
        107,
        109,
        111,
        113,
        115,
        119,
        121,
        123,
        125,
        127,
        169,
        172,
        175,
        198,
        201,
        203,
        210,
        212,
    }
    h3 = {116, 117, 183, 186, 192, 195, 207, 215, 218}
    if idx in h1:
        return 1
    if idx in h2:
        return 2
    if idx in h3:
        return 3
    return None


def apply_body_headings(doc: Document, note_idx: int) -> None:
    for i in range(0, note_idx):
        lvl = heading_level_for_para_idx(i)
        if lvl is None:
            continue
        p = doc.paragraphs[i]
        if not p.text.strip():
            continue
        if i in (19, 24, 33):
            apply_visual_heading(p, lvl)
    for i in range(note_idx + 1, len(doc.paragraphs)):
        lvl = heading_level_for_para_idx(i)
        if lvl is None:
            continue
        p = doc.paragraphs[i]
        if not p.text.strip():
            continue
        apply_visual_heading(p, lvl)


def cum_chars_before_each_paragraph(doc: Document) -> tuple[list[int], int]:
    cum = [0]
    for p in doc.paragraphs:
        cum.append(cum[-1] + len(p.text) + 1)
    return cum, cum[-1]


def page_for_index(cum: list[int], total: int, idx: int) -> int:
    if total <= 0:
        return 1
    r = cum[idx] / total
    return 1 + min(PMAX - 1, int(r * (PMAX - 1)))


def rebuild_toc(doc: Document, note_idx: int, cum: list[int], total: int) -> None:
    body_norm = [norm(p.text) for p in doc.paragraphs]
    rows: list[tuple[int, str, int]] = []

    for i in range(TOC_FIRST, note_idx):
        raw = doc.paragraphs[i].text.strip()
        if "\t" not in raw:
            continue
        # strip leading indent spaces from previous run
        raw_main = raw.lstrip(" ")
        if "\t" not in raw_main:
            continue
        title_raw, _ = raw_main.split("\t", 1)
        title_raw = title_raw.strip()
        title_n = norm(title_raw)
        bidx = body_index_for_toc_title(body_norm, note_idx, title_n)
        if bidx is None:
            lvl, pg = 2, 1
        else:
            lvl = heading_level_for_para_idx(bidx) or 2
            pg = page_for_index(cum, total, bidx)
        rows.append((lvl, title_raw, pg))

    last = 1
    fixed: list[tuple[int, str, int]] = []
    for lvl, t, pg in rows:
        pg = max(last, min(PMAX, pg))
        last = pg
        fixed.append((lvl, t, pg))

    for row, (lvl, title_raw, pg) in enumerate(fixed):
        idx = TOC_FIRST + row
        if idx >= note_idx:
            break
        p = doc.paragraphs[idx]
        indent = "    " * (lvl - 1)
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(
            TAB_POS, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        p.text = f"{indent}{title_raw}\t{pg}"
        p.style = doc.styles["List Paragraph"]


def update_note(doc: Document, note_idx: int) -> None:
    p = doc.paragraphs[note_idx]
    if not p.text.strip().startswith(NOTE_PREFIX):
        return
    p.text = (
        f"{NOTE_PREFIX} Heading hierarchy: outline levels 0–2 with bold + 14/12/11 pt on marked "
        f"titles (Abstract through appendices). TOC pages are estimates mapped to 1–{PMAX} from "
        "text position; adjust in Print Layout after Word repaginates."
    )


def main() -> None:
    doc = Document(DOC_PATH)
    note_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith(NOTE_PREFIX):
            note_idx = i
            break
    if note_idx is None:
        raise SystemExit("Note paragraph not found.")

    apply_body_headings(doc, note_idx)
    cum, total = cum_chars_before_each_paragraph(doc)
    rebuild_toc(doc, note_idx, cum, total)
    update_note(doc, note_idx)

    doc.save(DOC_PATH)
    print("OK:", DOC_PATH)


if __name__ == "__main__":
    main()
