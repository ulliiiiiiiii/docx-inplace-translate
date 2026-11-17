# docx-inplace-translate

Translate a Word file and get the same file back: tables, headings, headers and footers, images,
footnotes and styles stay exactly as they were. Only the texts change, and any translator can be
plugged in, as a Python function or as a command line tool.

Many tools translate a .docx by pulling out the text and writing a new, plain document. This one
never rebuilds the XML. It splits each part into tags and text and replaces only the text inside
`<w:t>`.

## How it works

- A paragraph is split into pieces at tabs and line breaks. Each piece is translated as a whole, even
  when Word has split its words over several runs, as it likes to do after a spell check. The
  translation goes into the first run and the others become empty. Highlighting in the middle of a
  sentence is lost that way; everything else stays.
- Body, headers, footers, footnotes and endnotes are translated. Pieces made only of digits and
  symbols are left alone.
- Two-language tables (app strings, "English | German"): with `columns`, only the source column is
  translated and the result goes into the target column. Empty target cells get a new run.
- The translator gets all texts of a file in one call, so it can work in batches, and returns as many.

## Usage

```python
from pathlib import Path
from docx_translate import translate_docx, language_columns

def translate(texts: list[str]) -> list[str]:
    ...  # DeepL, an LLM, a local model: one translation per text, in the same order

translate_docx(Path("manual.docx"), Path("manual.de.docx"), translate)

# app strings: fill the German column from the English one
translate_docx(Path("strings.docx"), Path("strings.out.docx"), translate, language_columns("english", "german"))
```

On the command line, with any program that reads one text per line and prints one translation per
line:

```sh
docx-translate manual.docx manual.de.docx --command "my-translator --to de"
docx-translate strings.docx out.docx --command "my-translator --to de" --columns english german
```

`target_texts(src, columns)` returns what the target column already holds, handy as examples of tone
and terminology for the translator.

## Install

```sh
pip install git+https://github.com/ulliiiiiiiii/docx-inplace-translate
```

No dependencies besides Python 3.10+.

## License

MIT
