# AGENTS.md

Guidance for coding agents working in this repository.

## Project Overview

`docx_parse` is a standalone Python package extracted from MinerU for parsing
`.docx` files. It converts DOCX input into:

- lightweight model output blocks
- MinerU-compatible middle JSON
- Markdown
- JSONL, one parsed content block per line

The public API lives in `docx_parse/api.py`.

## Main Pipeline

The primary conversion flow is:

```text
api.py
-> backend/office/docx_analyze.py
-> model/docx/main.py
-> model/docx/docx_converter.py
-> backend/office/model_output_to_middle_json.py
-> backend/office/office_middle_json_mkcontent.py
```

Key modules:

- `docx_parse/model/docx/docx_converter.py`: core DOCX parser.
- `docx_parse/model/docx/table_xml.py`: OOXML table renderer.
- `docx_parse/backend/office/model_output_to_middle_json.py`: model output to middle JSON.
- `docx_parse/backend/office/office_middle_json_mkcontent.py`: Markdown and content-list generation.
- `docx_parse/backend/utils/office_image.py`: Office image serialization.
- `docx_parse/backend/utils/office_chart.py`: Office chart extraction to HTML.
- `docx_parse/model/docx/tools/math/omml.py`: OMML math handling.

## Public APIs

Keep these APIs stable unless the requested task explicitly changes the public
surface:

- `convert_docx_to_model_output(file_bytes, tolerant=True)`
- `convert_docx_to_middle_json(file_bytes, image_output_dir=None, tolerant=True)`
- `convert_docx_to_markdown(file_bytes, image_output_dir=None, image_dir_name="images", tolerant=True)`
- `convert_docx_to_jsonl(file_bytes, image_output_dir=None, image_dir_name="images", tolerant=True)`
- `convert_docx_file_to_model_output(path, tolerant=True)`
- `convert_docx_file_to_middle_json(path, image_output_dir=None, tolerant=True)`
- `convert_docx_file_to_markdown(path, image_output_dir=None, image_dir_name="images", tolerant=True)`
- `convert_docx_file_to_jsonl(path, image_output_dir=None, image_dir_name="images", tolerant=True)`

## Behavior Rules

- Default parsing mode is tolerant: `tolerant=True`.
- Local parsing failures should be skipped when possible and recorded in
  middle JSON `_parse_errors`.
- Strict mode, `tolerant=False`, must continue to raise parsing errors.
- `page_idx` is a DOCX structural page/section index, not a rendered Word page.
- Embedded images are written only when `image_output_dir` is provided.
- Markdown and JSONL image references use `image_dir_name`.
- Table output is HTML embedded in Markdown/JSONL. Preserve table structure over
  visual fidelity.

## Development Commands

Use these checks before finalizing parser changes:

```powershell
python -m compileall -q docx_parse tests
python -m pytest -q
```

If `pytest` is not installed in the active environment, install test
dependencies first:

```powershell
python -m pip install -e ".[test]"
```

For performance checks on a real DOCX:

```powershell
python tools/benchmark_docx_performance.py path\to\input.docx --repeat 3
```

## Testing Notes

Tests are in `tests/test_docx_parse.py`. They dynamically create DOCX fixtures
with `python-docx` and cover:

- public API conversion to model output, middle JSON, Markdown, and JSONL
- headings and outline levels
- table line breaks, empty cells, header rows, and formula text
- images and captions
- headers and footers
- hyperlinks and visible inline styles
- SDT text and formula preservation
- tolerant mode and strict mode behavior

The regression test for `DOCX_01_PATH` depends on a local MinerU fixture path
and should skip when that file is unavailable.

## Repository Hygiene

- Do not commit generated outputs.
- `output/` and `outputs/` are ignored local output directories.
- Do not commit `.pytest_cache/`, `docx_parse.egg-info/`, virtualenvs, or build
  artifacts.
- Preserve user changes in a dirty worktree. Do not reset or revert unrelated
  files.

## Coding Guidance

- Prefer the existing parser and renderer patterns over new abstractions.
- Use structured XML/OOXML parsing instead of ad hoc string parsing when
  possible.
- Keep edits focused; the parser has many DOCX edge cases.
- Add tests for behavior changes, especially in tolerant mode and table output.
- Be careful with source comments that appear garbled from historical encoding;
  avoid broad re-encoding or comment churn unless explicitly requested.

