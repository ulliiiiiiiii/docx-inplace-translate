"""Translate Word files without touching the layout: tables, headings, headers and footers, images and
footnotes stay exactly as they were, only the texts change.

No rebuilding of the XML: each part is split into tags and text, and only the text inside `<w:t>` is
replaced. A paragraph falls apart into pieces at tabs and line breaks; the translation of a piece goes
into its first text run and the other runs become empty (Word likes to split words over several runs,
for example after a spell check). Highlighting in the middle of a sentence is lost that way, everything
else stays.

A table that keeps one language next to the other (app strings: "English | German") only gets the
translation of the source column written into the target column; the header row, the source column
and any other columns stay as they are.
"""

from __future__ import annotations

import html
import re
import zipfile
from collections import defaultdict
from collections.abc import Callable
from pathlib import Path

PARTS = re.compile(r"word/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml")
TOKEN = re.compile(r"<[^>]+>|[^<]+")
PARAGRAPH = re.compile(r"<w:p(?:\s[^>]*)?>")
EMPTY = re.compile(r"<w:p(?:\s[^>]*)?/>")  # an empty paragraph, for example in an empty cell
TEXT = re.compile(r"<w:t(?:\s[^>]*)?>")
BREAK = re.compile(r"<w:(?:tab|br|cr)(?:\s[^>]*)?/>")  # separates the pieces of a paragraph
LETTER = re.compile(r"[^\W\d_]")  # pieces made only of digits and symbols stay as they are
TABLE = re.compile(r"<w:tbl(?:\s[^>]*)?>")
ROW = re.compile(r"<w:tr(?:\s[^>]*)?>")
CELL = re.compile(r"<w:tc(?:\s[^>]*)?>")
# header row of a table -> (source column, target column) when it keeps one language next to the other
Columns = Callable[[list[str]], "tuple[int, int] | None"]
Parsed = dict[str, tuple[list[str], list[list[int]]]]  # per part of the file: tokens and pieces


def pieces(xml: str) -> tuple[list[str], list[list[int]]]:
    """The tokens of the XML and, per piece, the positions of its texts; nested paragraphs (text boxes)
    count on their own."""
    tokens = TOKEN.findall(xml)
    open_: list[list[list[int]]] = []  # per open paragraph its pieces
    done: list[list[int]] = []
    in_text = False
    for i, tok in enumerate(tokens):
        if not tok.startswith("<"):
            if in_text and open_:
                open_[-1][-1].append(i)
        elif tok.endswith("/>"):
            if BREAK.fullmatch(tok) and open_ and open_[-1][-1]:
                open_[-1].append([])
        elif PARAGRAPH.fullmatch(tok):
            open_.append([[]])
        elif tok == "</w:p>" and open_:
            done += [p for p in open_.pop() if p]
        elif TEXT.fullmatch(tok):
            in_text = True
        elif tok == "</w:t>":
            in_text = False
    return tokens, done


def _text(tokens: list[str], piece: list[int]) -> str:
    """The text of a piece with whitespace collapsed, without line breaks, so pieces can be separated by
    blank lines."""
    return " ".join("".join(html.unescape(tokens[i]) for i in piece).split())


def cells(tokens: list[str]) -> list[tuple[int, int, int] | None]:
    """For every token the table cell it sits in, as (table, row, column), None outside a table; nested
    tables count on their own."""
    out: list[tuple[int, int, int] | None] = []
    open_: list[list[int]] = []  # per open table [number, row, column]
    tables = 0
    for tok in tokens:
        if TABLE.fullmatch(tok):
            open_.append([tables, -1, -1])
            tables += 1
        elif tok == "</w:tbl>" and open_:
            open_.pop()
        elif ROW.fullmatch(tok) and open_:
            open_[-1][1:] = [open_[-1][1] + 1, -1]
        elif CELL.fullmatch(tok) and open_:
            open_[-1][2] += 1
        out.append(tuple(open_[-1]) if open_ and open_[-1][2] >= 0 else None)
    return out


def _read(src: Path) -> tuple[dict[str, bytes], list[zipfile.ZipInfo], Parsed]:
    with zipfile.ZipFile(src) as z:
        parts = {i.filename: z.read(i) for i in z.infolist()}
        infos = z.infolist()
    return parts, infos, {name: pieces(data.decode("utf-8")) for name, data in parts.items() if PARTS.fullmatch(name)}


def _language_tables(parsed: Parsed, columns: Columns) -> tuple[set, list]:
    """The tables whose header row `columns` recognises as language next to language: their pieces (which
    stay as they are) and, per row with text in the source column, (part, source text, pieces of the
    target cell, position for a new text run when the target cell is empty)."""
    taken, rows = set(), []
    for name, (tokens, found) in parsed.items():
        where = cells(tokens)
        fields: dict[int, dict[tuple[int, int], list[list[int]]]] = defaultdict(lambda: defaultdict(list))
        for p in found:
            if (cell := where[p[0]]) is not None:
                fields[cell[0]][cell[1:]].append(p)
        paragraph: dict[tuple[int, int, int], int] = {}  # per cell its first paragraph
        for i, (tok, cell) in enumerate(zip(tokens, where)):
            if cell is not None and cell not in paragraph and (EMPTY.fullmatch(tok) or PARAGRAPH.fullmatch(tok)):
                paragraph[cell] = i
        for t, table in fields.items():
            header = [" ".join(_text(tokens, p) for p in table.get((0, c), []))
                      for c in range(max(c for _, c in table) + 1)]
            if (pair := columns(header)) is None:
                continue
            source, target = pair
            taken |= {(name, p[0]) for cell_pieces in table.values() for p in cell_pieces}
            for r in sorted({r for r, _ in table if r > 0}):
                text = " ".join(_text(tokens, p) for p in table.get((r, source), []))
                if LETTER.search(text):
                    rows.append((name, text, table.get((r, target), []), paragraph.get((t, r, target))))
    return taken, rows


def target_texts(src: Path, columns: Columns) -> list[str]:
    """What the target columns already hold: existing texts in the target language, useful as examples of
    tone and terminology for the translator."""
    _, _, parsed = _read(src)
    return [" ".join(_text(parsed[name][0], p) for p in target) for name, _, target, _ in
            _language_tables(parsed, columns)[1] if target]


def _set(tokens: list[str], piece: list[int], text: str) -> None:
    tokens[piece[0] - 1] = '<w:t xml:space="preserve">'  # spaces at the edges stay
    tokens[piece[0]] = html.escape(text, quote=False)
    for i in piece[1:]:
        tokens[i] = ""


def _fill(tokens: list[str], paragraph: int, text: str) -> None:
    """A text run in the first paragraph of an empty cell."""
    run = f'<w:r><w:t xml:space="preserve">{html.escape(text, quote=False)}</w:t></w:r>'
    if EMPTY.fullmatch(tokens[paragraph]):
        tokens[paragraph] = f"{tokens[paragraph][:-2].rstrip()}>{run}</w:p>"
        return
    end = tokens.index("</w:p>", paragraph)
    tokens[end] = run + tokens[end]


def translate_docx(src: Path, dst: Path, translate: Callable[[list[str]], list[str]],
                   columns: Columns | None = None) -> None:
    """Write the Word file `src` as `dst` with translated texts. `translate` gets all texts of the file at
    once (so it can translate in batches) and returns as many. With `columns`: in tables that keep one
    language next to the other, only the source column is translated and replaces the target cell."""
    parts, infos, parsed = _read(src)
    taken, rows = _language_tables(parsed, columns) if columns else (set(), [])
    places = [(name, p) for name, (tokens, found) in parsed.items() for p in found
              if (name, p[0]) not in taken and LETTER.search(_text(tokens, p))]
    new = translate([_text(parsed[name][0], p) for name, p in places] + [text for _, text, _, _ in rows])
    if len(new) != len(places) + len(rows):
        raise ValueError(f"{len(new)} translations for {len(places) + len(rows)} texts")
    for (name, p), text in zip(places, new):
        _set(parsed[name][0], p, text)
    for (name, _, target_pieces, paragraph), text in zip(rows, new[len(places):]):
        tokens = parsed[name][0]
        if target_pieces:  # the translation replaces what was there
            _set(tokens, target_pieces[0], text)
            for p in target_pieces[1:]:
                _set(tokens, p, "")
        elif paragraph is not None:
            _fill(tokens, paragraph, text)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w") as z:
        for info in infos:
            data = "".join(parsed[info.filename][0]).encode("utf-8") if info.filename in parsed \
                else parts[info.filename]
            z.writestr(info, data)


def language_columns(source: str, target: str) -> Columns:
    """A `columns` function for header rows that name both languages, for example
    language_columns("english", "german") for a table headed "Key | English | German (current)"."""
    def find(header: list[str]) -> tuple[int, int] | None:
        lowered = [h.lower() for h in header]
        src = next((i for i, h in enumerate(lowered) if source.lower() in h), None)
        dst = next((i for i, h in enumerate(lowered) if target.lower() in h), None)
        return (src, dst) if src is not None and dst is not None and src != dst else None
    return find


def main(argv: list[str] | None = None) -> int:
    """docx-translate SRC.docx DST.docx --command CMD [--columns SOURCE TARGET]"""
    import argparse
    import shlex
    import subprocess
    import sys

    parser = argparse.ArgumentParser(prog="docx-translate",
                                     description="Translate the texts of a .docx, keep its layout.")
    parser.add_argument("src", type=Path)
    parser.add_argument("dst", type=Path)
    parser.add_argument("--command", required=True,
                        help="translator command: one text per line on stdin, one translation per line on stdout")
    parser.add_argument("--columns", nargs=2, metavar=("SOURCE", "TARGET"),
                        help="language names in the header row of two-language tables")
    args = parser.parse_args(argv)

    def translate(texts: list[str]) -> list[str]:
        if not texts:
            return []
        result = subprocess.run(shlex.split(args.command), input="\n".join(texts) + "\n", capture_output=True,
                                text=True, check=True)
        return result.stdout.splitlines()

    try:
        translate_docx(args.src, args.dst, translate, language_columns(*args.columns) if args.columns else None)
    except (OSError, ValueError, zipfile.BadZipFile, subprocess.CalledProcessError) as exc:
        print(f"docx-translate: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
