from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

from docx_parse.backend.office.docx_analyze import office_docx_analyze
from docx_parse.backend.office.office_middle_json_mkcontent import union_make
from docx_parse.model.docx.main import convert_binary
from docx_parse.utils.enum_class import MakeMode


class FileImageWriter:
    """Small writer compatible with the Office image persistence helpers."""

    def __init__(self, output_dir: str | Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def write(self, name: str, data: bytes) -> None:
        path = self.output_dir / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)


def _ensure_bytes(file_bytes: bytes | bytearray | BinaryIO) -> bytes:
    if isinstance(file_bytes, bytes):
        return file_bytes
    if isinstance(file_bytes, bytearray):
        return bytes(file_bytes)
    return file_bytes.read()


def _records_to_jsonl(records: list[dict]) -> str:
    if not records:
        return ""
    return "\n".join(json.dumps(record, ensure_ascii=False) for record in records) + "\n"


def convert_docx_to_model_output(
    file_bytes: bytes | bytearray | BinaryIO,
    tolerant: bool = True,
) -> list:
    """Convert DOCX bytes to the lightweight block model output."""

    return convert_binary(BytesIO(_ensure_bytes(file_bytes)), tolerant=tolerant)


def convert_docx_to_middle_json(
    file_bytes: bytes | bytearray | BinaryIO,
    image_output_dir: str | Path | None = None,
    tolerant: bool = True,
) -> dict:
    """Convert DOCX bytes to MinerU-compatible middle JSON.

    When image_output_dir is provided, embedded images are written there and the
    middle JSON references local image paths instead of inline base64 payloads.
    """

    writer = FileImageWriter(image_output_dir) if image_output_dir else None
    middle_json, _ = office_docx_analyze(
        _ensure_bytes(file_bytes),
        image_writer=writer,
        tolerant=tolerant,
    )
    return middle_json


def convert_docx_to_markdown(
    file_bytes: bytes | bytearray | BinaryIO,
    image_output_dir: str | Path | None = None,
    image_dir_name: str = "images",
    mode: str = MakeMode.MM_MD,
    tolerant: bool = True,
) -> str:
    """Convert DOCX bytes to Markdown."""

    middle_json = convert_docx_to_middle_json(
        file_bytes,
        image_output_dir,
        tolerant=tolerant,
    )
    return union_make(middle_json["pdf_info"], mode, image_dir_name, tolerant=tolerant)


def convert_docx_to_jsonl(
    file_bytes: bytes | bytearray | BinaryIO,
    image_output_dir: str | Path | None = None,
    image_dir_name: str = "images",
    tolerant: bool = True,
) -> str:
    """Convert DOCX bytes to JSONL.

    The JSONL output is based on MinerU's flat content_list structure: each line
    is one JSON object, and each object includes fields such as type, text,
    table_body, img_path, and page_idx depending on the parsed block type.
    """

    middle_json = convert_docx_to_middle_json(
        file_bytes,
        image_output_dir,
        tolerant=tolerant,
    )
    records = union_make(
        middle_json["pdf_info"],
        MakeMode.CONTENT_LIST,
        image_dir_name,
        tolerant=tolerant,
    )
    return _records_to_jsonl(records)


def convert_docx_file_to_model_output(path: str | Path, tolerant: bool = True) -> list:
    with open(path, "rb") as f:
        return convert_docx_to_model_output(f, tolerant=tolerant)


def convert_docx_file_to_middle_json(
    path: str | Path,
    image_output_dir: str | Path | None = None,
    tolerant: bool = True,
) -> dict:
    with open(path, "rb") as f:
        return convert_docx_to_middle_json(
            f,
            image_output_dir=image_output_dir,
            tolerant=tolerant,
        )


def convert_docx_file_to_markdown(
    path: str | Path,
    image_output_dir: str | Path | None = None,
    image_dir_name: str = "images",
    mode: str = MakeMode.MM_MD,
    tolerant: bool = True,
) -> str:
    with open(path, "rb") as f:
        return convert_docx_to_markdown(
            f,
            image_output_dir=image_output_dir,
            image_dir_name=image_dir_name,
            mode=mode,
            tolerant=tolerant,
        )


def convert_docx_file_to_jsonl(
    path: str | Path,
    image_output_dir: str | Path | None = None,
    image_dir_name: str = "images",
    tolerant: bool = True,
) -> str:
    with open(path, "rb") as f:
        return convert_docx_to_jsonl(
            f,
            image_output_dir=image_output_dir,
            image_dir_name=image_dir_name,
            tolerant=tolerant,
        )
