# DOCX 容错解析改造说明

本文档描述当前分支 `improve-tolerant-docx-parse` 相较于前一个基线分支 `verify-docx-performance` 的改动点。其他 agent 如果从 `verify-docx-performance` 开始，可以按照本文档复现同样的功能。

## 改造目标

默认启用非严格解析模式：DOCX 解析过程中，如果某个局部元素解析失败，不让整个文档解析中断，而是记录错误并继续解析后续内容。

需要保留严格模式能力：

```python
convert_docx_file_to_jsonl(path, tolerant=False)
```

当 `tolerant=False` 时，遇到异常仍然直接抛出，便于开发排查。

## API 行为

所有公开 DOCX API 增加或保留 `tolerant` 参数，并默认 `True`：

```python
convert_docx_to_model_output(file_bytes, tolerant=True)
convert_docx_to_middle_json(file_bytes, image_output_dir=None, tolerant=True)
convert_docx_to_markdown(file_bytes, image_output_dir=None, image_dir_name="images", tolerant=True)
convert_docx_to_jsonl(file_bytes, image_output_dir=None, image_dir_name="images", tolerant=True)

convert_docx_file_to_model_output(path, tolerant=True)
convert_docx_file_to_middle_json(path, image_output_dir=None, tolerant=True)
convert_docx_file_to_markdown(path, image_output_dir=None, image_dir_name="images", tolerant=True)
convert_docx_file_to_jsonl(path, image_output_dir=None, image_dir_name="images", tolerant=True)
```

默认调用时：

```python
convert_docx_file_to_jsonl(path)
```

等价于：

```python
convert_docx_file_to_jsonl(path, tolerant=True)
```

局部失败时会跳过失败元素，继续解析后续内容。

## 文件级改动

### 1. `docx_parse/model/docx/docx_converter.py`

在 `DocxConverter.__init__` 中增加容错状态：

```python
self.tolerant: bool = False
self.parse_errors: list[dict[str, Any]] = []
self._max_parse_errors: int = 100
```

新增两个辅助方法：

```python
def _record_parse_error(self, stage, exc, element=None, tag_name=None) -> None:
    ...

def _run_tolerant(self, stage, func, *args, element=None, tag_name=None):
    ...
```

`_record_parse_error` 负责记录局部错误，字段包括：

- `stage`
- `tag`
- `error_type`
- `message`

错误最多记录 100 条，避免异常过多时导致内存和日志膨胀。

`_run_tolerant` 负责执行局部解析函数：

- `self.tolerant=False`：异常继续抛出。
- `self.tolerant=True`：记录错误，写 warning，返回 `None`，继续后续解析。

修改 `DocxConverter.convert` 签名：

```python
def convert(self, file_stream: BinaryIO, tolerant: bool = True):
```

每次解析时重置：

```python
self.tolerant = tolerant
self.parse_errors = []
```

在 `_walk_linear` 主循环中，把以下局部处理点改为 `_run_tolerant` 包裹：

- DrawingML：`_handle_drawingml`
- 文本框：`_handle_textbox_content`
- 图片：`_handle_pictures`
- 普通段落：`_handle_text_elements`
- SDT 目录：`_handle_sdt_as_index`
- SDT 普通段落：`_handle_text_elements`

表格原本已有 `try/except`，保留其跳过行为，但在 `self.tolerant=True` 时补充 `_record_parse_error("table", ...)`。表格失败时不改变原有行为，仍然跳过坏表格。

### 2. `docx_parse/model/docx/main.py`

增加一个可挂载错误信息的 list 子类：

```python
class DocxPages(list):
    """List of parsed DOCX pages with optional tolerant-mode errors."""
```

修改入口默认值：

```python
def convert_path(file_path: str, tolerant: bool = True):
def convert_binary(file_binary: BinaryIO, tolerant: bool = True):
```

调用转换器时传递 `tolerant`：

```python
converter.convert(file_binary, tolerant=tolerant)
```

返回 `DocxPages(converter.pages)`。当容错模式下有错误时，把错误挂到结果对象上：

```python
if tolerant and converter.parse_errors:
    setattr(pages, "parse_errors", converter.parse_errors)
```

### 3. `docx_parse/backend/office/docx_analyze.py`

修改 `office_docx_analyze`：

```python
def office_docx_analyze(file_bytes, image_writer=None, tolerant: bool = True):
```

调用：

```python
results = convert_binary(file_stream, tolerant=tolerant)
middle_json = result_to_middle_json(results, image_writer, tolerant=tolerant)
```

如果 `results` 上存在 `parse_errors`，写入中间 JSON：

```python
parse_errors = getattr(results, "parse_errors", None)
if tolerant and parse_errors:
    middle_json["_parse_errors"] = list(parse_errors)
```

### 4. `docx_parse/backend/office/model_output_to_middle_json.py`

修改 `result_to_middle_json`：

```python
def result_to_middle_json(model_output_blocks_list, image_writer, tolerant: bool = True):
```

在 page 级别包裹：

```python
try:
    page_info = blocks_to_page_info(page_blocks, image_writer, index)
except Exception as exc:
    if not tolerant:
        raise
    middle_json["_parse_errors"].append(...)
    continue
```

这样某一页转换 middle JSON 失败时，不影响其他页。

如果没有错误，移除空的 `_parse_errors`：

```python
if tolerant and not middle_json.get("_parse_errors"):
    middle_json.pop("_parse_errors", None)
```

### 5. `docx_parse/backend/office/office_middle_json_mkcontent.py`

修改 `union_make`：

```python
def union_make(pdf_info_dict, make_mode, img_buket_path='', tolerant: bool = True):
```

在输出阶段增加 block 级容错：

- Markdown 模式：逐个 `para_block` 调用 `mk_blocks_to_markdown([para_block], ...)`，单个 block 失败时 warning 后跳过。
- `CONTENT_LIST`：逐个 block 调用 `make_blocks_to_content_list`，失败时 warning 后跳过。
- `CONTENT_LIST_V2`：逐个 block 调用 `make_blocks_to_content_list_v2`，失败时 warning 后跳过。

严格模式下保持原行为：

```python
if not tolerant:
    raise
```

### 6. `docx_parse/api.py`

公开 API 全部增加或透传 `tolerant=True`。

关键点：

- `convert_docx_to_model_output` 调用 `convert_binary(..., tolerant=tolerant)`
- `convert_docx_to_middle_json` 调用 `office_docx_analyze(..., tolerant=tolerant)`
- `convert_docx_to_markdown` 调用 `union_make(..., tolerant=tolerant)`
- `convert_docx_to_jsonl` 调用 `union_make(..., tolerant=tolerant)`
- file 版本 API 全部透传 `tolerant`

### 7. `README.md`

更新 Public API 说明，标明 `tolerant=True` 是默认值。

补充说明：

```text
`tolerant=True` is the default. Local parsing failures are skipped and recorded in middle JSON `_parse_errors` when available, so the rest of the document can continue parsing. Pass `tolerant=False` to use strict mode and raise parsing errors immediately.
```

## 测试要求

在 `tests/test_docx_parse.py` 中增加两类测试。

### 1. 默认容错解析测试

构造一个 DOCX：

- 段落 1：`before`
- 段落 2：`bad paragraph`
- 段落 3：`after`

用 `monkeypatch` 替换 `DocxConverter._handle_text_elements`，当段落文本是 `bad paragraph` 时抛出：

```python
raise RuntimeError("synthetic paragraph failure")
```

断言：

- `convert_docx_file_to_jsonl(docx_path, tolerant=False)` 会抛异常。
- `convert_docx_file_to_middle_json(docx_path)` 默认容错，不抛异常，并包含 `_parse_errors`。
- `_parse_errors[0]["stage"] == "paragraph"`。
- `convert_docx_file_to_jsonl(docx_path)` 默认容错，输出包含 `before` 和 `after`，不包含 `bad paragraph`。

### 2. 输出阶段 block 级容错测试

构造一个 `pdf_info`：

- 第一个 block 是合法 text。
- 第二个 block 是结构损坏的 image，例如 `"blocks": "not-a-list"`。

断言：

- `union_make(pdf_info, MakeMode.CONTENT_LIST, tolerant=False)` 抛 `TypeError`。
- `union_make(pdf_info, MakeMode.CONTENT_LIST)` 默认容错，只返回合法 text record。

## 验证命令

完成改造后执行：

```bash
python -m compileall -q docx_parse tests
python -m pytest -q tests/test_docx_parse.py
```

当前分支验证结果：

```text
12 passed
```

## 相关提交

当前分支相对 `verify-docx-performance` 多两个提交：

```text
8fb2588 Add tolerant DOCX parsing mode
3684314 Enable tolerant DOCX parsing by default
```

## 注意事项

- 不要提交本地 `outputs/` 目录。
- 容错模式只跳过局部失败元素，不保证失败元素内容完整保留。
- `_parse_errors` 是排查信息，不参与 Markdown / JSONL 输出。
- 默认容错是为了提高生产可用性；需要定位解析问题时使用 `tolerant=False`。
