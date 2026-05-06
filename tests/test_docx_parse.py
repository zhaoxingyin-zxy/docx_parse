import json
from pathlib import Path

from docx import Document
from PIL import Image, ImageDraw

from docx_parse import (
    convert_docx_file_to_jsonl,
    convert_docx_file_to_markdown,
    convert_docx_file_to_middle_json,
    convert_docx_file_to_model_output,
)


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
