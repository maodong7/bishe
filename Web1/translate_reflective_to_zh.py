"""
Translate Reflective.docx (English) to Chinese, save beside the original.
Uses deep-translator (Google); preserves paragraph styles; handles TOC tab+page.
"""
from __future__ import annotations

import re
import time
from pathlib import Path

from deep_translator import GoogleTranslator
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = Path(r"D:\桌面文档文件勿动！\Desktop\毕业论文提交\个人\Dongdong Mao\Reflective.docx")
DST = Path(r"D:\桌面文档文件勿动！\Desktop\毕业论文提交\个人\Dongdong Mao\Reflective_中文版.docx")

MAX_CHUNK = 4200
SLEEP = 0.18
translator = GoogleTranslator(source="en", target="zh-CN")


def translate_text(text: str) -> str:
    text = text.strip()
    if not text:
        return text
    # TOC-style: "Title\t12" — translate title only, keep tab + page number
    if "\t" in text:
        left, right = text.rsplit("\t", 1)
        if right.strip().isdigit():
            t_left = translate_plain(left.strip())
            return f"{t_left}\t{right.strip()}"
    return translate_plain(text)


def translate_plain(text: str) -> str:
    if not text:
        return text
    if len(text) <= MAX_CHUNK:
        return _call(text)
    # split on sentence boundaries to stay under API limits
    parts: list[str] = []
    buf = ""
    for chunk in re.split(r"(?<=[.!?])\s+", text):
        if len(buf) + len(chunk) + 1 > MAX_CHUNK:
            if buf:
                parts.append(_call(buf.strip()))
            buf = chunk
        else:
            buf = (buf + " " + chunk).strip() if buf else chunk
    if buf:
        parts.append(_call(buf.strip()))
    return " ".join(parts)


def _call(text: str) -> str:
    last_err: Exception | None = None
    for attempt in range(4):
        try:
            return translator.translate(text)
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(1.2 * (attempt + 1))
    raise RuntimeError(f"translate failed: {last_err}") from last_err


def translate_paragraph(p: Paragraph) -> None:
    raw = p.text
    if not raw.strip():
        return
    new = translate_text(raw)
    p.text = new
    time.sleep(SLEEP)


def walk_table(table: Table) -> None:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                translate_paragraph(p)


def walk_document(doc: Document) -> None:
    for p in doc.paragraphs:
        translate_paragraph(p)
    for table in doc.tables:
        walk_table(table)


def main() -> None:
    if not SRC.is_file():
        raise SystemExit(f"Missing source: {SRC}")
    print("Loading", SRC)
    doc = Document(str(SRC))
    walk_document(doc)
    doc.save(str(DST))
    print("Saved", DST)


if __name__ == "__main__":
    main()
