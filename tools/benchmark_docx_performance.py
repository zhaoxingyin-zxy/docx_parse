from __future__ import annotations

import argparse
import json
import statistics
import sys
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any

from lxml import etree

from docx_parse.backend.office.model_output_to_middle_json import result_to_middle_json
from docx_parse.backend.office.office_middle_json_mkcontent import union_make
from docx_parse.model.docx.main import convert_binary
from docx_parse.utils.enum_class import MakeMode


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _count_docx_structure(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as docx_zip:
        document_xml = docx_zip.read("word/document.xml")
    root = etree.fromstring(document_xml)
    ns = {"w": W_NS, "m": M_NS}
    return {
        "document_xml_bytes": len(document_xml),
        "paragraphs": len(root.findall(".//w:p", namespaces=ns)),
        "runs": len(root.findall(".//w:r", namespaces=ns)),
        "tables": len(root.findall(".//w:tbl", namespaces=ns)),
        "table_rows": len(root.findall(".//w:tr", namespaces=ns)),
        "table_cells": len(root.findall(".//w:tc", namespaces=ns)),
        "omml_math": len(root.findall(".//m:oMath", namespaces=ns)),
        "drawings": len(root.findall(".//w:drawing", namespaces=ns)),
        "text_nodes": len(root.findall(".//w:t", namespaces=ns)),
    }


def _run_once(path: Path) -> dict[str, Any]:
    read_start = time.perf_counter()
    file_bytes = path.read_bytes()
    read_seconds = time.perf_counter() - read_start

    convert_start = time.perf_counter()
    model_output = convert_binary(BytesIO(file_bytes))
    convert_seconds = time.perf_counter() - convert_start

    middle_start = time.perf_counter()
    middle_json = result_to_middle_json(model_output, image_writer=None)
    middle_seconds = time.perf_counter() - middle_start

    content_start = time.perf_counter()
    records = union_make(middle_json["pdf_info"], MakeMode.CONTENT_LIST, "images")
    content_seconds = time.perf_counter() - content_start

    return {
        "read_seconds": read_seconds,
        "convert_binary_seconds": convert_seconds,
        "result_to_middle_json_seconds": middle_seconds,
        "content_list_seconds": content_seconds,
        "total_seconds": read_seconds + convert_seconds + middle_seconds + content_seconds,
        "pages": len(model_output),
        "blocks": sum(len(page) for page in model_output),
        "records": len(records),
    }


def _summarize_runs(runs: list[dict[str, Any]]) -> dict[str, Any]:
    timing_keys = [
        "read_seconds",
        "convert_binary_seconds",
        "result_to_middle_json_seconds",
        "content_list_seconds",
        "total_seconds",
    ]
    summary: dict[str, Any] = {
        "runs": len(runs),
        "pages": runs[-1]["pages"],
        "blocks": runs[-1]["blocks"],
        "records": runs[-1]["records"],
    }
    for key in timing_keys:
        values = [float(run[key]) for run in runs]
        summary[key] = {
            "min": min(values),
            "mean": statistics.fmean(values),
            "max": max(values),
        }
    return summary


def _print_human(result: dict[str, Any]) -> None:
    print(f"file: {result['file']}")
    print(f"size_mb: {result['size_mb']:.2f}")
    print("structure:")
    for key, value in result["structure"].items():
        print(f"  {key}: {value}")
    print("timings_seconds:")
    for key, value in result["summary"].items():
        if isinstance(value, dict):
            print(
                f"  {key}: min={value['min']:.3f} "
                f"mean={value['mean']:.3f} max={value['max']:.3f}"
            )
    print(f"pages: {result['summary']['pages']}")
    print(f"blocks: {result['summary']['blocks']}")
    print(f"records: {result['summary']['records']}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Benchmark DOCX parsing stages for performance verification."
    )
    parser.add_argument("docx_path", type=Path)
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--json", action="store_true", dest="as_json")
    parser.add_argument(
        "--max-convert-seconds",
        type=float,
        default=None,
        help="Exit with status 1 if the fastest convert_binary run exceeds this value.",
    )
    args = parser.parse_args()

    path = args.docx_path
    if not path.exists():
        print(f"DOCX file does not exist: {path}", file=sys.stderr)
        return 2
    if args.repeat < 1:
        print("--repeat must be >= 1", file=sys.stderr)
        return 2

    runs = [_run_once(path) for _ in range(args.repeat)]
    summary = _summarize_runs(runs)
    result = {
        "file": str(path),
        "size_mb": path.stat().st_size / 1024 / 1024,
        "structure": _count_docx_structure(path),
        "summary": summary,
        "runs": runs,
    }

    if args.as_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        _print_human(result)

    if args.max_convert_seconds is not None:
        fastest = summary["convert_binary_seconds"]["min"]
        if fastest > args.max_convert_seconds:
            print(
                f"convert_binary_seconds min {fastest:.3f}s exceeds "
                f"{args.max_convert_seconds:.3f}s",
                file=sys.stderr,
            )
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
