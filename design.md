# DOCX 解析模块设计文档

## 1. 目标

本项目从 MinerU 中抽取 DOCX 解析能力，提供独立 Python 包，用于将 `.docx` 文件解析为：

- 轻量模型输出结构 `model_output`
- MinerU 兼容的 `middle_json`
- Markdown
- JSONL

当前设计重点是：

- 不依赖 `mammoth`
- 基于 `python-docx` 和底层 OOXML 解析 DOCX
- 表格、图片、图表、公式等复杂元素尽量保留
- 默认容错解析，局部失败不影响整篇文档输出

## 2. 对外 API

主要入口位于 `docx_parse/api.py`。

```python
convert_docx_file_to_model_output(path, tolerant=True)
convert_docx_file_to_middle_json(path, image_output_dir=None, tolerant=True)
convert_docx_file_to_markdown(path, image_output_dir=None, image_dir_name="images", tolerant=True)
convert_docx_file_to_jsonl(path, image_output_dir=None, image_dir_name="images", tolerant=True)
```

也支持 bytes / file-like 输入：

```python
convert_docx_to_model_output(file_bytes, tolerant=True)
convert_docx_to_middle_json(file_bytes, image_output_dir=None, tolerant=True)
convert_docx_to_markdown(file_bytes, image_output_dir=None, image_dir_name="images", tolerant=True)
convert_docx_to_jsonl(file_bytes, image_output_dir=None, image_dir_name="images", tolerant=True)
```

### 参数说明

- `image_output_dir`
  - `None`：图片不落盘保存。
  - 指定目录：图片保存到该目录。
- `image_dir_name`
  - 控制 Markdown / JSONL 中图片引用前缀。
- `tolerant`
  - 默认 `True`。
  - 局部元素解析失败时跳过当前元素，继续解析后续内容。
  - 设置为 `False` 时进入严格模式，异常直接抛出。

## 3. 总体处理流程

整体链路如下：

```text
DOCX 文件
  ↓
api.py
  ↓
office_docx_analyze()
  ↓
convert_binary()
  ↓
DocxConverter.convert()
  ↓
model_output blocks
  ↓
result_to_middle_json()
  ↓
middle_json
  ↓
union_make()
  ↓
Markdown / JSONL
```

核心分三层：

1. DOCX 原始结构解析层：`docx_parse/model/docx/docx_converter.py`
2. 中间结构转换层：`docx_parse/backend/office/model_output_to_middle_json.py`
3. 输出生成层：`docx_parse/backend/office/office_middle_json_mkcontent.py`

## 4. DOCX 解析层设计

核心类是：

```python
DocxConverter
```

位置：

```text
docx_parse/model/docx/docx_converter.py
```

职责：

- 读取 DOCX 二进制
- 使用 `python-docx` 加载文档
- 遍历 WordprocessingML body
- 识别段落、标题、列表、表格、图片、图表、文本框、目录等元素
- 输出轻量 `model_output` block 结构

### 主流程

```python
converter = DocxConverter()
converter.convert(file_binary, tolerant=True)
pages = converter.pages
```

内部会执行：

1. 清理失效 DOCX 内部关系
2. `Document(BytesIO(file_bytes))` 加载 DOCX
3. 预扫描目录 anchor、编号标题等信息
4. `_walk_linear(self.docx_obj.element.body)` 顺序遍历文档 body
5. 解析结果写入 `self.cur_page`

## 5. 元素处理策略

### 5.1 段落

普通段落通过：

```python
_handle_text_elements()
```

处理内容包括：

- 普通文本
- 标题识别
- 列表识别
- 超链接
- 加粗、斜体、下划线、删除线等样式
- 公式文本
- 目录 anchor

### 5.2 表格

表格通过 OOXML 直接解析，不使用 `mammoth`。

入口：

```python
_handle_tables()
```

核心渲染函数：

```python
render_table_html()
```

位置：

```text
docx_parse/model/docx/table_xml.py
```

输出为 HTML 表格：

```html
<table>
  <thead>
    <tr><th>...</th></tr>
  </thead>
  <tr><td>...</td></tr>
</table>
```

当前表格处理特性：

- 支持 `rowspan`
- 支持 `colspan`
- 支持表头行识别并输出 `thead`
- 表头单元格输出 `th`
- 单元格内多个 `w:p` 用 `<br/>` 保留段落边界
- 空单元格段落保留 `<br/>`
- 去除表格单元格中的 `strong`、`em` 等字体样式标签
- 表格公式不做 LaTeX 转换，尽量提取原始文本

### 5.3 图片

图片通过 DOCX relationship 解析：

```python
_handle_pictures()
```

图片序列化逻辑位于：

```text
docx_parse/backend/utils/office_image.py
```

当 `image_output_dir` 存在时，后续 middle_json 阶段会将 base64 图片保存到本地，并替换为本地路径。

当 `image_output_dir=None` 时，不落盘保存。

### 5.4 图表

图表处理逻辑位于：

```text
docx_parse/backend/utils/office_chart.py
```

优先尝试从图表绑定的 workbook 或 OOXML cache 中提取表格化数据，并输出 HTML 表格形式。

如果图表解析失败，容错模式下跳过图表，不影响其他内容。

### 5.5 公式

正文公式和表格公式都不依赖 `pylatexenc`。

当前策略：

- 不做复杂 LaTeX 转换
- 尽量从 OMML 内部提取文本
- 表格公式直接保留原始文本表达
- 公式解析失败时不影响整篇文档

## 6. 中间 JSON 转换设计

位置：

```text
docx_parse/backend/office/model_output_to_middle_json.py
```

入口：

```python
result_to_middle_json(model_output_blocks_list, image_writer, tolerant=True)
```

职责：

- 将 `model_output` 转为 MinerU 风格 `middle_json`
- 按 page 组织 `para_blocks`
- 处理图片落盘
- 处理表格 HTML 内联图片
- 处理图表图片
- 处理标题编号
- 处理目录 anchor 关联

输出结构示例：

```python
{
    "pdf_info": [
        {
            "para_blocks": [...],
            "discarded_blocks": [...],
            "page_idx": 0
        }
    ],
    "_backend": "office",
    "_version_name": "...",
    "_parse_errors": [...]
}
```

当容错模式下发生局部错误，会写入：

```python
"_parse_errors": [
    {
        "stage": "paragraph",
        "tag": "p",
        "error_type": "RuntimeError",
        "message": "..."
    }
]
```

## 7. Markdown / JSONL 输出设计

位置：

```text
docx_parse/backend/office/office_middle_json_mkcontent.py
```

统一入口：

```python
union_make(pdf_info_dict, make_mode, img_buket_path="", tolerant=True)
```

支持输出：

- `MakeMode.MM_MD`
- `MakeMode.NLP_MD`
- `MakeMode.CONTENT_LIST`
- `MakeMode.CONTENT_LIST_V2`

### Markdown

Markdown 输出会处理：

- 标题 `#`
- 正文段落
- 列表
- 表格 HTML
- 图片引用
- 图表内容
- 目录

### JSONL

JSONL 基于 `CONTENT_LIST`，每一行为一个 JSON object。

常见类型：

```json
{"type": "text", "text": "...", "page_idx": 0}
{"type": "table", "table_body": "<table>...</table>", "page_idx": 0}
{"type": "image", "img_path": "images/xxx.png", "page_idx": 0}
```

## 8. 容错机制设计

当前默认：

```python
tolerant=True
```

核心原则：

- 单个段落失败，不影响后续段落
- 单个表格失败，不影响后续内容
- 图片、图表、文本框、SDT 失败时跳过当前元素
- 输出阶段单个 block 损坏时跳过该 block
- 错误信息记录到 `_parse_errors`

严格模式：

```python
tolerant=False
```

用于开发排查，遇到异常直接抛出。

容错边界主要包括：

- `DocxConverter._walk_linear`
- `result_to_middle_json`
- `union_make`

## 9. 测试设计

测试文件：

```text
tests/test_docx_parse.py
```

覆盖内容：

- DOCX 到 model output / middle json / Markdown / JSONL
- 表格换行保留
- 空单元格 `<br/>` 保留
- 表头 `thead` / `th`
- 表格公式文本保留
- 页眉页脚
- 超链接和可见样式
- SDT 和正文公式
- `docx_01.docx` 表格回归
- 容错模式：
  - 默认容错继续解析
  - `tolerant=False` 严格抛错
  - 输出阶段坏 block 跳过

验证命令：

```bash
python -m compileall -q docx_parse tests
python -m pytest -q tests/test_docx_parse.py
```

## 10. 当前设计取舍

### 优点

- 不依赖 `mammoth`
- 表格控制更细
- 默认容错，稳定性更好
- 支持 Markdown / JSONL / middle_json 多种输出
- 可以通过 `_parse_errors` 追踪局部失败原因

### 限制

- DOCX 复杂排版不会完整还原
- 公式不做完整语义转换
- 图表解析依赖 OOXML / workbook 信息，复杂图表可能降级
- 表格视觉结构和 Word 实际渲染可能存在差异
- 容错模式下失败元素会丢失，而不是自动修复

## 11. 后续可优化方向

- 增加更细的错误类型和错误位置记录
- 增加图片处理模式：`inline / file / drop`
- 表格复杂结构继续增强
- 对大型 DOCX 做性能 profiling
- 增加更多真实 DOCX 回归样本
- 支持输出解析质量报告
- 对 `_parse_errors` 增加统计摘要，例如失败元素数量、类型分布等
