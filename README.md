# docx_parse

Standalone DOCX parsing package extracted from MinerU. It converts `.docx` files to:

- lightweight model output blocks
- MinerU-compatible middle JSON
- Markdown
- JSONL, one parsed content block per line

## Usage

```python
from docx_parse import convert_docx_file_to_jsonl, convert_docx_file_to_markdown

markdown = convert_docx_file_to_markdown(
    "input.docx",
    image_output_dir="output/images",
    image_dir_name="images",
)

jsonl = convert_docx_file_to_jsonl(
    "input.docx",
    image_output_dir="output/images",
    image_dir_name="images",
)
print(jsonl)
```

## Public API

- `convert_docx_to_model_output(file_bytes)`
- `convert_docx_to_middle_json(file_bytes, image_output_dir=None)`
- `convert_docx_to_markdown(file_bytes, image_output_dir=None, image_dir_name="images")`
- `convert_docx_to_jsonl(file_bytes, image_output_dir=None, image_dir_name="images")`
- `convert_docx_file_to_model_output(path)`
- `convert_docx_file_to_middle_json(path, image_output_dir=None)`
- `convert_docx_file_to_markdown(path, image_output_dir=None, image_dir_name="images")`
- `convert_docx_file_to_jsonl(path, image_output_dir=None, image_dir_name="images")`

`image_output_dir` is optional. When provided, embedded images are written to disk and Markdown/JSONL references them by `image_dir_name/<filename>`.
