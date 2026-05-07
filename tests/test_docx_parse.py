import json
import re
from pathlib import Path

import pytest
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw

from docx_parse import (
    convert_docx_file_to_jsonl,
    convert_docx_file_to_markdown,
    convert_docx_file_to_middle_json,
    convert_docx_file_to_model_output,
)


DOCX_01_PATH = Path(r"D:\zxy\code\MinerU-master\MinerU-master\demo\office_docs\docx_01.docx")
DOCX_01_TABLE_BASELINE = Path(__file__).parent / "fixtures" / "docx_01_tables_baseline.json"


def _table_records_from_jsonl(jsonl: str) -> list[dict]:
    return [
        record
        for record in (json.loads(line) for line in jsonl.splitlines())
        if record.get("type") == "table"
    ]


def _normalize_table_html(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for image in soup.find_all("img"):
        image["src"] = "<image>"
    return re.sub(r"\s+", " ", str(soup)).strip()


def _make_sample_docx(path: Path, image_path: Path) -> None:
    image = Image.new("RGB", (96, 48), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((8, 8, 88, 40), outline="black", width=2)
    draw.text((20, 17), "DOCX", fill="black")
    image.save(image_path)

    doc = Document()
    doc.add_heading("Standalone DOCX Parser", level=1)

    paragraph = doc.add_paragraph("This parser keeps ")
    bold = paragraph.add_run("bold")
    bold.bold = True
    paragraph.add_run(" and ")
    italic = paragraph.add_run("italic")
    italic.italic = True
    paragraph.add_run(" text.")

    doc.add_heading("Data", level=2)
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Answer"
    table.cell(1, 1).text = "42"

    doc.add_picture(str(image_path))
    doc.add_paragraph("Figure 1: generated test image", style="Caption")
    doc.save(path)


def _set_outline_level(paragraph, level: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    outline = p_pr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        p_pr.append(outline)
    outline.set(qn("w:val"), str(level))


def test_docx_to_model_output_middle_json_markdown_and_jsonl(tmp_path):
    docx_path = tmp_path / "sample.docx"
    source_image = tmp_path / "source.png"
    image_output_dir = tmp_path / "images"
    _make_sample_docx(docx_path, source_image)

    model_output = convert_docx_file_to_model_output(docx_path)
    assert isinstance(model_output, list)
    assert any(block.get("type") == "title" for page in model_output for block in page)
    assert any(block.get("type") == "table" for page in model_output for block in page)
    assert any(block.get("type") == "image" for page in model_output for block in page)

    middle_json = convert_docx_file_to_middle_json(docx_path, image_output_dir=image_output_dir)
    assert middle_json["_backend"] == "office"
    assert middle_json["pdf_info"]
    assert any(image_output_dir.iterdir())

    markdown = convert_docx_file_to_markdown(
        docx_path,
        image_output_dir=image_output_dir,
        image_dir_name="images",
    )
    assert "# " in markdown
    assert "Standalone DOCX Parser" in markdown
    assert "## " in markdown
    assert "Data" in markdown
    assert "<table>" in markdown
    assert "Answer" in markdown
    assert "42" in markdown
    assert "![](images/" in markdown
    assert "Figure 1: generated test image" in markdown
    assert "**bold**" in markdown
    assert "*italic*" in markdown

    jsonl = convert_docx_file_to_jsonl(
        docx_path,
        image_output_dir=image_output_dir,
        image_dir_name="images",
    )
    records = [json.loads(line) for line in jsonl.splitlines()]
    assert records
    assert all(isinstance(record, dict) for record in records)
    assert any(record.get("text") and "Standalone DOCX Parser" in record["text"] for record in records)
    assert any(record.get("type") == "table" and "Answer" in record.get("table_body", "") for record in records)
    assert any(record.get("type") == "image" and record.get("img_path", "").startswith("images/") for record in records)
    assert all("page_idx" in record for record in records)


def test_outline_level_9_is_treated_as_body_text(tmp_path):
    docx_path = tmp_path / "outline_body.docx"
    doc = Document()
    heading = doc.add_heading("Real Heading", level=2)
    assert heading.text == "Real Heading"
    paragraph = doc.add_paragraph("Body paragraph with DOCX outline level 9.")
    _set_outline_level(paragraph, 9)
    doc.save(docx_path)

    markdown = convert_docx_file_to_markdown(docx_path)
    jsonl = convert_docx_file_to_jsonl(docx_path)
    records = [json.loads(line) for line in jsonl.splitlines()]

    assert "## " in markdown
    assert "Real Heading" in markdown
    assert "##########" not in markdown
    assert "Body paragraph with DOCX outline level 9." in markdown
    assert not any(
        record.get("text") == "Body paragraph with DOCX outline level 9."
        and record.get("text_level") == 10
        for record in records
    )


def test_table_cell_line_breaks_and_empty_paragraphs_are_preserved(tmp_path):
    docx_path = tmp_path / "table_breaks.docx"
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = ""
    first = cell.paragraphs[0]
    first.add_run("first line")
    first.add_run().add_break()
    first.add_run("second line")
    cell.add_paragraph("")
    cell.add_paragraph("after empty paragraph")
    doc.save(docx_path)

    records = [json.loads(line) for line in convert_docx_file_to_jsonl(docx_path).splitlines()]
    table_record = next(record for record in records if record.get("type") == "table")
    table_body = table_record["table_body"]

    assert "<p>first line<br/>second line</p>" in table_body
    assert "<p><br/></p>" in table_body
    assert "<p>after empty paragraph</p>" in table_body


def test_docx_01_table_output_matches_legacy_baseline():
    if not DOCX_01_PATH.exists():
        pytest.skip(f"local regression DOCX is missing: {DOCX_01_PATH}")

    baseline_tables = json.loads(DOCX_01_TABLE_BASELINE.read_text(encoding="utf-8"))
    current_tables = _table_records_from_jsonl(convert_docx_file_to_jsonl(DOCX_01_PATH))

    assert len(current_tables) == len(baseline_tables) == 8
    for current, baseline in zip(current_tables, baseline_tables):
        assert current.get("table_caption", []) == baseline.get("table_caption", [])
        assert current.get("page_idx") == baseline.get("page_idx")
        assert _normalize_table_html(current["table_body"]) == _normalize_table_html(
            baseline["table_body"]
        )
