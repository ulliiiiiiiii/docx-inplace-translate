import zipfile

import pytest

import docx_translate as dt

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def cell(text):
    return f"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>" if text else "<w:tc><w:p/></w:tc>"


def row(*texts):
    return "<w:tr>" + "".join(cell(t) for t in texts) + "</w:tr>"


DOCUMENT = (
    f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>'
    '<w:p><w:r><w:t>Hel</w:t></w:r><w:r><w:t xml:space="preserve">lo world</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Name</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Value &amp; more</w:t></w:r></w:p>'
    "<w:p><w:r><w:t>2026</w:t></w:r></w:p>"
    "<w:tbl>" + row("Key", "English", "German (current)") + row("greeting", "Good morning", "")
    + row("farewell", "Goodbye", "Tschüss") + "</w:tbl>"
    "</w:body></w:document>"
)
HEADER = f'<?xml version="1.0" encoding="UTF-8"?><w:hdr {W}><w:p><w:r><w:t>Draft</w:t></w:r></w:p></w:hdr>'
IMAGE = bytes(range(256))


def make_docx(path):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/header1.xml", HEADER)
        z.writestr("word/media/image1.png", IMAGE)
    return path


def texts_of(path, part="word/document.xml"):
    tokens, found = dt.pieces(zipfile.ZipFile(path).read(part).decode())
    return [dt._text(tokens, p) for p in found]


def bracket(texts):
    return [f"[{t}]" for t in texts]


def test_every_text_is_translated_in_place(tmp_path):
    src, dst = make_docx(tmp_path / "in.docx"), tmp_path / "out.docx"
    seen = []
    dt.translate_docx(src, dst, lambda texts: seen.extend(texts) or bracket(texts))
    assert seen == ["Hello world", "Name", "Value & more", "Key", "English", "German (current)", "greeting",
                    "Good morning", "farewell", "Goodbye", "Tschüss", "Draft"]
    assert texts_of(dst)[:4] == ["[Hello world]", "[Name]", "[Value & more]", "2026"]
    assert texts_of(dst, "word/header1.xml") == ["[Draft]"]
    assert zipfile.ZipFile(dst).read("word/media/image1.png") == IMAGE
    assert "<w:tab/>" in zipfile.ZipFile(dst).read("word/document.xml").decode()


def test_two_language_table_only_fills_the_target_column(tmp_path):
    src, dst = make_docx(tmp_path / "in.docx"), tmp_path / "out.docx"
    columns = dt.language_columns("english", "german")
    assert dt.target_texts(src, columns) == ["Tschüss"]
    seen = []
    dt.translate_docx(src, dst, lambda texts: seen.extend(texts) or bracket(texts), columns)
    assert seen == ["Hello world", "Name", "Value & more", "Draft", "Good morning", "Goodbye"]
    assert texts_of(dst)[4:] == ["Key", "English", "German (current)", "greeting", "Good morning",
                                 "[Good morning]", "farewell", "Goodbye", "[Goodbye]"]


def test_wrong_number_of_translations_is_refused(tmp_path):
    src = make_docx(tmp_path / "in.docx")
    with pytest.raises(ValueError, match="translations for"):
        dt.translate_docx(src, tmp_path / "out.docx", lambda texts: texts[:-1])


def test_command_line_with_an_external_translator(tmp_path):
    src, dst = make_docx(tmp_path / "in.docx"), tmp_path / "out.docx"
    assert dt.main([str(src), str(dst), "--command", "tr a-z A-Z"]) == 0
    assert texts_of(dst)[0] == "HELLO WORLD"
