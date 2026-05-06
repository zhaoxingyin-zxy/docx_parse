# docx_parse 项目分析

## 项目定位

`docx_parse` 是从 MinerU 中抽取出来的独立 DOCX 解析包，目标是只保留 Word 文档解析能力，不依赖 MinerU 原项目中的 PDF、OCR、VLM、FastAPI、Gradio 等模块。

它可以把 `.docx` 文件转换为以下格式：

- 轻量级 model output：解析器直接生成的块结构。
- middle JSON：兼容 MinerU Office 后端的中间格式。
- Markdown：适合直接阅读、检索和入库前预览。
- JSONL：一行一个内容块，适合流式处理、RAG 切分和下游数据管道。

## 主要目录

```text
docx_parse/
  api.py                              # 对外公开 API
  model/docx/docx_converter.py        # DOCX 解析核心
  model/docx/tools/math/omml.py       # OMML 公式转 LaTeX
  backend/office/docx_analyze.py      # DOCX 解析入口
  backend/office/model_output_to_middle_json.py
                                      # model output 转 middle JSON
  backend/office/office_magic_model.py
                                      # Office 块结构归一化
  backend/office/office_middle_json_mkcontent.py
                                      # middle JSON 转 Markdown/content list
  backend/utils/office_image.py       # Office 图片序列化和占位图处理
  backend/utils/office_chart.py       # Office 图表转 HTML
  utils/                              # 枚举、配置、格式化、哈希等轻量工具
tests/
  test_docx_parse.py                  # 用 python-docx 动态生成样例文档并验证输出
```

## 对外 API

`docx_parse.api` 暴露了两组 API：一组接收 bytes 或二进制流，一组接收文件路径。

```python
from docx_parse import (
    convert_docx_file_to_markdown,
    convert_docx_file_to_jsonl,
    convert_docx_file_to_middle_json,
    convert_docx_file_to_model_output,
)

markdown = convert_docx_file_to_markdown("input.docx", image_output_dir="output/images")
jsonl = convert_docx_file_to_jsonl("input.docx", image_output_dir="output/images")
```

核心参数：

- `image_output_dir`：如果传入，解析出的图片会写入该目录，输出内容中使用相对路径引用图片。
- `image_dir_name`：Markdown/JSONL 中引用图片时使用的目录名前缀，默认是 `images`。

## DOCX 处理总流程

整体链路如下：

```text
DOCX 文件
  -> office_docx_analyze(file_bytes, image_writer)
  -> convert_binary(file_stream)
  -> DocxConverter.convert()
  -> model_output pages
  -> result_to_middle_json(model_output, image_writer)
  -> middle_json
  -> union_make(..., MakeMode.MM_MD)           # Markdown
  -> union_make(..., MakeMode.CONTENT_LIST)    # JSONL 的记录来源
```

其中 `DocxConverter` 是最核心的解析器，它负责把 WordprocessingML、python-docx 对象和 mammoth 输出整合成统一块结构。

## model output 是什么

model output 是 `DocxConverter.pages`，形态是分页数组：

```python
[
    [
        {"type": "title", "level": 1, "content": "..."},
        {"type": "text", "content": "..."},
        {"type": "table", "content": "<table>...</table>"},
        {"type": "image", "content": "data:image/...;base64,..."},
        {"type": "equation", "content": "...latex..."},
        {"type": "list", "content": [...]},
        {"type": "index", "content": [...]},
    ]
]
```

这里的 page 不一定等同于 Word 的视觉页码。DOCX 本身不是固定分页格式，解析器主要依据 section break 等结构切分。

## middle JSON 是什么

middle JSON 是 MinerU Office 后端的统一结构，顶层类似：

```json
{
  "pdf_info": [
    {
      "page_idx": 0,
      "para_blocks": [],
      "discarded_blocks": []
    }
  ],
  "_backend": "office",
  "_version_name": "0.1.0"
}
```

它的作用是把 DOCX、PPTX、XLSX 这类 Office 原生解析结果归一成类似 MinerU PDF 后端的块格式，后续 Markdown、content list、JSONL 都从它生成。

## 文本处理逻辑

文本段落由 `_handle_text_elements()` 处理，主要步骤：

1. 用 `python-docx` 把底层 XML 段落包装成 `Paragraph`。
2. 遍历段落内的 run、hyperlink、inline content control 等内联节点。
3. 提取文本、字体样式、超链接和字段代码形式的超链接。
4. 根据段落样式判断类型：标题、普通文本、列表、目录、caption、公式块等。
5. 生成对应的 block。

支持的文本特性包括：

- 标题层级：`Title`、`Heading 1/2/...`、outline level。
- 字体样式：bold、italic、underline、strikethrough、subscript、superscript。
- 超链接：标准 hyperlink 节点和 Word 字段代码形式的 hyperlink。
- inline content control：避免 `python-docx` 默认跳过 `w:sdt` 内联文本。
- 隐藏文本过滤：避免 TOC 页码字段等隐藏内容污染正文。

## 公式处理逻辑

Word 公式通常使用 OMML，即 Office Math Markup Language。项目中的：

```text
docx_parse/model/docx/tools/math/omml.py
```

负责把 OMML 转成 LaTeX。

处理方式：

- 段落中如果只有公式，生成 `equation` block。
- 段落中如果文本和公式混合，公式会被包成 `<eq>...</eq>`，之后在 Markdown 阶段转成 `$...$`。
- 表格中的 OMML 会额外处理，因为 mammoth 默认会静默丢弃表格单元格里的公式。

## 表格处理逻辑

表格由 `_handle_tables()` 处理。优先策略是：

1. 先用 mammoth 对完整 DOCX 进行 HTML 预解析，得到所有顶层表格。
2. 解析到具体 `w:tbl` 时，按顺序取 mammoth 预解析的表格 HTML。
3. 对表格 HTML 做清洗，只保留对结构有意义的标签和属性，例如 `table`、`tr`、`td`、`colspan`、`rowspan`。
4. 对表格内图片和公式做补充处理。

为什么使用 mammoth：

- mammoth 对 Word 表格、列表、样式和图片上下文处理比较成熟。
- 只解析孤立 `w:tbl` XML 容易丢编号、样式、图片等上下文。

## 图片处理逻辑

图片处理分两段：

1. `DocxConverter._handle_pictures()` 从 DOCX relationship 中找到 image part，把图片转成 base64 data URI。
2. `result_to_middle_json()` 中如果提供了 `image_writer`，会把 base64 图片写入 `image_output_dir`，并在输出中替换成相对路径。

对于 WMF/EMF 这类矢量图片：

- Windows 下会尝试用 Pillow 渲染。
- 非 Windows 或渲染失败时，会生成占位图，避免整个解析失败。

## 列表和目录处理逻辑

列表处理依赖 Word numbering 定义：

- 读取段落的 `numId` 和 `ilvl`。
- 判断有序或无序列表。
- 维护列表栈，构建嵌套列表结构。
- 对“看起来是章节标题的编号列表”做预扫描，把它们转换成 title block，而不是普通 list。

目录处理包括两类：

- 标准 TOC content control：识别 `sdt` 目录块并转成 `index`。
- 普通段落形式目录：根据样式和文本特征尝试转成 `index`。

如果目录条目带 `_Toc...` anchor，Markdown 中会保留链接目标。

## Markdown 输出逻辑

Markdown 由：

```text
backend/office/office_middle_json_mkcontent.py
```

中的 `union_make(pdf_info, MakeMode.MM_MD, image_dir_name)` 生成。

主要规则：

- title -> `#`、`##` 等标题。
- text -> 普通 Markdown 段落。
- image -> `![](images/xxx.jpg)`。
- table -> HTML table 原样嵌入 Markdown。
- inline equation -> `$...$`。
- interline equation -> LaTeX 公式文本。
- list/index -> 嵌套 Markdown 列表。
- hyperlink -> `[text](url)`。

## JSONL 输出逻辑

JSONL 是基于 `MakeMode.CONTENT_LIST` 生成的 flat content list：

```python
records = union_make(middle_json["pdf_info"], MakeMode.CONTENT_LIST, "images")
jsonl = "".join(json.dumps(record, ensure_ascii=False) + "\n" for record in records)
```

每一行是一个 JSON 对象，例如：

```json
{"type":"text","text":"...","page_idx":0}
{"type":"image","img_path":"images/xxx.jpg","image_caption":[],"page_idx":1}
{"type":"table","table_body":"<table>...</table>","table_caption":[],"page_idx":2}
```

这种格式适合：

- 大文档流式读取。
- RAG 分块前预处理。
- 数据管道逐行消费。
- 避免一次性加载完整 JSON 数组。

## 大文档处理经验

解析 `D:\zxy\毛泽东文集.docx` 时，项目输出了：

- 3012 个 page/section 级结果。
- 18348 行 JSONL。
- 约 151 万 Markdown 字符。
- 2 张图片。

这个文件暴露了一个 DOCX 边界问题：`python-docx` 读取页眉/页脚时可能因为 section 继承链过深或递归引用触发 `RecursionError`。

当前处理方式：

```python
try:
    self._add_header_footer(self.docx_obj)
except RecursionError:
    logger.warning("Skipping DOCX header/footer parsing ...")
```

也就是说：

- 正文、标题、表格、图片、目录继续解析。
- 页眉页脚异常时跳过，不让整个文档失败。

对于书籍类 DOCX，这通常是合理的，因为页眉页脚多是页码、章节名或重复装饰信息，不是主体内容。

## 测试策略

测试文件：

```text
tests/test_docx_parse.py
```

测试中使用 `python-docx` 动态创建一个 Word 文档，包含：

- 一级/二级标题。
- 加粗和斜体正文。
- 表格。
- 图片。
- caption。

然后验证：

- model output 中存在 title/table/image。
- middle JSON 标记 `_backend == "office"`。
- Markdown 中存在标题、表格、图片引用和样式文本。
- JSONL 每行都是 JSON 对象，并包含 text/table/image 等记录。

运行方式：

```powershell
python -m pytest -q
```

## 当前限制

- DOCX 不是真正分页格式，输出中的 `page_idx` 更接近 section/结构页，不等同于 Word 渲染后的物理页。
- 页眉页脚目前依赖 `python-docx`，遇到递归 section 引用时会跳过。
- 表格主要以 HTML 形式保留，复杂样式不会完整复刻 Word 视觉效果。
- 公式转换依赖 OMML 结构，极端复杂公式可能存在 LaTeX 简化或丢样式。
- 当前项目没有 CLI，主要作为 Python 包使用。

## 后续可改进点

- 增加命令行入口，例如 `docx-parse input.docx -o output/ --format md,jsonl`。
- 页眉页脚改为直接读取底层 OOXML relationship，绕开 `python-docx` 的递归接口。
- 为大文档增加单次解析的高级 API，避免用户连续调用 Markdown/JSONL 时重复解析。
- 增加 JSONL 分块策略，例如按标题层级、字数、token 数切分。
- 增加更多真实 DOCX 样例测试，覆盖目录、公式、嵌套表格、文本框、图表等复杂结构。
