# Copyright (c) Opendatalab. All rights reserved.
import re
from typing import Literal

from loguru import logger

from docx_parse.utils.enum_class import ContentType, BlockType
from docx_parse.utils.magic_model_utils import tie_up_category_by_index


class MagicModel:
    def __init__(self, page_blocks: list):
        self.page_blocks = page_blocks
        blocks = _build_magic_blocks(classify_caption_blocks(page_blocks))
        self._init_block_groups()
        self._split_blocks_by_type(blocks)
        self._fix_media_block_groups()

    def _init_block_groups(self):
        self.image_blocks = []
        self.table_blocks = []
        self.chart_blocks = []
        self.interline_equation_blocks = []
        self.text_blocks = []
        self.title_blocks = []
        self.discarded_blocks = []
        self.list_blocks = []
        self.index_blocks = []
        self.ref_text_blocks = []
        self.phonetic_blocks = []
        self.all_spans = []

    def _split_blocks_by_type(self, blocks: list):
        block_groups = [
            ([BlockType.IMAGE_BODY, BlockType.IMAGE_CAPTION, BlockType.IMAGE_FOOTNOTE], self.image_blocks),
            ([BlockType.TABLE_BODY, BlockType.TABLE_CAPTION, BlockType.TABLE_FOOTNOTE], self.table_blocks),
            ([BlockType.CHART_BODY, BlockType.CHART_CAPTION], self.chart_blocks),
            ([BlockType.INTERLINE_EQUATION], self.interline_equation_blocks),
            ([BlockType.TEXT], self.text_blocks),
            ([BlockType.TITLE], self.title_blocks),
            ([BlockType.REF_TEXT], self.ref_text_blocks),
            ([BlockType.PHONETIC], self.phonetic_blocks),
            ([BlockType.HEADER, BlockType.FOOTER, BlockType.PAGE_NUMBER, BlockType.ASIDE_TEXT, BlockType.PAGE_FOOTNOTE], self.discarded_blocks),
            ([BlockType.LIST], self.list_blocks),
            ([BlockType.INDEX], self.index_blocks),
        ]
        for block in blocks:
            block_type = block["type"]
            for type_names, target in block_groups:
                if block_type in type_names:
                    target.append(block)
                    break

    def _fix_media_block_groups(self):
        self.image_blocks, not_include_image_blocks = fix_two_layer_blocks(self.image_blocks, BlockType.IMAGE)
        self.table_blocks, not_include_table_blocks = fix_two_layer_blocks(self.table_blocks, BlockType.TABLE)
        self.chart_blocks, not_include_chart_blocks = fix_two_layer_blocks(self.chart_blocks, BlockType.CHART)
        for block in not_include_image_blocks + not_include_table_blocks + not_include_chart_blocks:
            block["type"] = BlockType.TEXT
            self.text_blocks.append(block)

    def get_list_blocks(self):
        return self.list_blocks

    def get_index_blocks(self):
        return self.index_blocks

    def get_image_blocks(self):
        return self.image_blocks

    def get_table_blocks(self):
        return self.table_blocks

    def get_chart_blocks(self):
        return self.chart_blocks

    def get_title_blocks(self):
        return self.title_blocks

    def get_text_blocks(self):
        return self.text_blocks

    def get_interline_equation_blocks(self):
        return self.interline_equation_blocks

    def get_discarded_blocks(self):
        return self.discarded_blocks


def _span_from_block(block_type: str, block_content: str, block_info: dict):
    if block_type in ["text", "title", "image_caption", "table_caption", "chart_caption", "header", "footer", "page_footnote"]:
        return block_type, parse_text_block_spans(block_content)
    if block_type == "image":
        return BlockType.IMAGE_BODY, {"type": ContentType.IMAGE, "image_base64": block_content}
    if block_type == "table":
        return BlockType.TABLE_BODY, {"type": ContentType.TABLE, "html": clean_table_html(block_content)}
    if block_type == "equation":
        return BlockType.INTERLINE_EQUATION, {"type": ContentType.INTERLINE_EQUATION, "content": block_content}
    if block_type == "chart":
        span = {"type": ContentType.CHART, "content": block_content}
        if block_info.get("image_base64"):
            span["image_base64"] = block_info["image_base64"]
        return BlockType.CHART_BODY, span
    return block_type, None


def _line_from_span(span):
    if isinstance(span, dict):
        return {"spans": [span]}
    if isinstance(span, list):
        return {"spans": span}
    raise ValueError(f"Unsupported span type: {type(span)}")


def _block_from_span(block_type: str, span, block_info: dict, index: int) -> dict:
    block = {"type": block_type, "lines": [_line_from_span(span)], "index": index}
    anchor = block_info.get("anchor")
    if isinstance(anchor, str) and anchor.strip() and block_type in [BlockType.TITLE, BlockType.TEXT, BlockType.INTERLINE_EQUATION]:
        block["anchor"] = anchor.strip()
    if block_type == BlockType.TITLE:
        block["is_numbered_style"] = block_info.get("is_numbered_style", False)
        block["level"] = block_info.get("level", 1)
    return block


def _append_structured_block(blocks: list, block_info: dict, index: int, block_type: str) -> bool:
    if block_type == "list":
        parsed = parse_list_block(block_info)
    elif block_type == "index":
        parsed = parse_index_block(block_info)
    else:
        return False
    if parsed:
        parsed["index"] = index
        blocks.append(parsed)
    return True


def _build_magic_blocks(page_blocks: list) -> list:
    blocks = []
    for index, block_info in enumerate(page_blocks):
        block_type = block_info["type"]
        block_content = block_info.get("content", "")
        if not block_content and block_type != BlockType.CHART:
            continue
        if _append_structured_block(blocks, block_info, index, block_type):
            continue
        block_type, span = _span_from_block(block_type, block_content, block_info)
        if span is None:
            continue
        blocks.append(_block_from_span(block_type, span, block_info, index))
    return blocks

_TEXT_TAG_RE = re.compile(r'<text(?:\s+style="([^"]*)")?>')


def _style_list(style_str: str | None) -> list:
    if not style_str:
        return []
    return [s.strip() for s in style_str.split(',') if s.strip()]


def _append_plain_span(spans: list, content: str) -> None:
    if content:
        spans.append({"type": ContentType.TEXT, "content": content})


def _next_markup_tag(content: str, pos: int):
    candidates = []
    eq_start = content.find('<eq>', pos)
    hyperlink_start = content.find('<hyperlink>', pos)
    text_tag_match = _TEXT_TAG_RE.search(content, pos)
    if eq_start != -1:
        candidates.append((eq_start, 'eq', None))
    if hyperlink_start != -1:
        candidates.append((hyperlink_start, 'hyperlink', None))
    if text_tag_match:
        candidates.append((text_tag_match.start(), 'text', text_tag_match))
    return min(candidates, key=lambda x: x[0]) if candidates else None


def _parse_eq_span(content: str, tag_pos: int, spans: list):
    eq_end = content.find('</eq>', tag_pos)
    if eq_end == -1:
        _append_plain_span(spans, content[tag_pos:])
        return None
    spans.append({
        "type": ContentType.INLINE_EQUATION,
        "content": content[tag_pos + 4:eq_end],
    })
    return eq_end + 5


def _parse_text_span(content: str, tag_pos: int, tag_match, spans: list):
    text_end = content.find('</text>', tag_pos)
    if text_end == -1:
        _append_plain_span(spans, content[tag_pos:])
        return None
    tag_open_end = content.find('>', tag_pos) + 1
    span = {"type": ContentType.TEXT, "content": content[tag_open_end:text_end]}
    style_str = tag_match.group(1) if tag_match and tag_match.start() == tag_pos else None
    if style_str:
        span["style"] = _style_list(style_str)
    spans.append(span)
    return text_end + 7


def _parse_hyperlink_span(content: str, tag_pos: int, spans: list):
    hyperlink_end = content.find('</hyperlink>', tag_pos)
    if hyperlink_end == -1:
        _append_plain_span(spans, content[tag_pos:])
        return None
    hyperlink_content = content[tag_pos + 11:hyperlink_end]
    span = _hyperlink_span_from_markup(hyperlink_content)
    if span is None:
        _append_plain_span(spans, content[tag_pos:])
        return None
    spans.append(span)
    return hyperlink_end + 12


def _hyperlink_span_from_markup(hyperlink_content: str):
    inner_text_match = _TEXT_TAG_RE.search(hyperlink_content)
    text_end = hyperlink_content.find('</text>')
    url_start = hyperlink_content.find('<url>')
    url_end = hyperlink_content.find('</url>')
    if not inner_text_match or text_end == -1 or url_start == -1 or url_end == -1:
        return None
    span = {
        "type": ContentType.HYPERLINK,
        "content": hyperlink_content[inner_text_match.end():text_end],
        "url": hyperlink_content[url_start + 5:url_end],
    }
    style = _style_list(inner_text_match.group(1))
    if style:
        span["style"] = style
    return span


def _parse_next_markup_span(content: str, pos: int, last_end: int, spans: list):
    next_tag = _next_markup_tag(content, pos)
    if not next_tag:
        _append_plain_span(spans, content[last_end:])
        return None
    tag_pos, tag_type, tag_match = next_tag
    _append_plain_span(spans, content[last_end:tag_pos])
    if tag_type == 'eq':
        return _parse_eq_span(content, tag_pos, spans)
    if tag_type == 'text':
        return _parse_text_span(content, tag_pos, tag_match, spans)
    return _parse_hyperlink_span(content, tag_pos, spans)


def parse_text_block_spans(content: str) -> list:
    if not content:
        return []
    spans = []
    last_end = 0
    pos = 0
    while pos < len(content):
        next_pos = _parse_next_markup_span(content, pos, last_end, spans)
        if next_pos is None:
            break
        pos = next_pos
        last_end = pos
    return spans

def parse_list_block(list_block: dict):
    """
    递归解析嵌套列表结构，生成与VLM一致的blocks结构。

    Args:
        list_block: 列表块字典

    Returns:
        tuple: (解析后的列表block, 下一个可用索引)
    """
    content = list_block.get("content", [])
    if not content:
        return None

    blocks = []

    for item in content:
        item_type = item.get("type", "")

        if item_type == "text":
            # 解析文本项（可能包含行内公式和超链接）
            text_content = item.get("content", "")
            spans = parse_text_block_spans(text_content)
            text_block = {
                "type": BlockType.TEXT,
                "lines": [{"spans": spans}]
            }
            blocks.append(text_block)

        elif item_type == "list":
            # 递归解析嵌套列表
            nested_list = parse_list_block(item)
            if nested_list:
                blocks.append(nested_list)

    # 构建当前列表block
    result = {
        "type": BlockType.LIST,
        "attribute": list_block.get("attribute", "unordered"),
        "ilevel": list_block.get("ilevel", 0),
        "blocks": blocks
    }

    return result


def parse_index_block(index_block: dict):
    """
    递归解析嵌套索引结构（目录），生成与list一致的blocks结构。

    Args:
        index_block: 索引块字典

    Returns:
        解析后的索引block字典，若内容为空则返回 None
    """
    content = index_block.get("content", [])
    if not content:
        return None

    blocks = []

    for item in content:
        item_type = item.get("type", "")

        if item_type == "text":
            text_content = item.get("content", "")
            spans = parse_text_block_spans(text_content)
            text_block = {
                "type": BlockType.TEXT,
                "lines": [{"spans": spans}]
            }
            anchor = item.get("anchor")
            if isinstance(anchor, str) and anchor.strip():
                text_block["anchor"] = anchor.strip()
            blocks.append(text_block)

        elif item_type == "index":
            nested_index = parse_index_block(item)
            if nested_index:
                blocks.append(nested_index)

    result = {
        "type": BlockType.INDEX,
        "ilevel": index_block.get("ilevel", 0),
        "blocks": blocks
    }

    return result


def _clean_table_tag(match, preserved_attrs: set, img_preserved_attrs: set):
    full_tag = match.group(0)
    tag_name = match.group(1).lower()
    is_self_closing = full_tag.rstrip().endswith('/>')
    current_preserved = preserved_attrs | (
        img_preserved_attrs if tag_name == 'img' else set()
    )
    kept_attrs = _kept_table_tag_attrs(full_tag, current_preserved)
    attrs_str = ' ' + ' '.join(kept_attrs) if kept_attrs else ''
    return f'<{tag_name}{attrs_str}/>' if is_self_closing else f'<{tag_name}{attrs_str}>'


def _kept_table_tag_attrs(full_tag: str, current_preserved: set) -> list[str]:
    kept_attrs = []
    attr_pattern = r'(\w+)\s*=\s*(?:"([^"]*)"|\'([^\']*)\'|(\S+))|(\w+)(?=\s|>|/>)'
    for attr_match in re.finditer(attr_pattern, full_tag):
        if attr_match.group(5):
            continue
        attr_name = attr_match.group(1)
        if attr_name is None:
            continue
        attr_name = attr_name.lower()
        attr_value = attr_match.group(2) or attr_match.group(3) or attr_match.group(4) or ""
        if attr_name in current_preserved:
            kept_attrs.append(f'{attr_name}="{attr_value}"')
    return kept_attrs


def clean_table_html(html: str) -> str:
    if not html:
        return ""
    preserved_attrs = {'colspan', 'rowspan'}
    img_preserved_attrs = {'src', 'alt', 'width', 'height'}
    tag_pattern = r'<(\w+)(?:\s+[^>]*)?\s*/?>'
    return re.sub(
        tag_pattern,
        lambda match: _clean_table_tag(match, preserved_attrs, img_preserved_attrs),
        html,
    )

def isolated_formula_clean(txt):
    latex = txt[:]
    if latex.startswith("\\["): latex = latex[2:]
    if latex.endswith("\\]"): latex = latex[:-2]
    latex = latex.strip()
    return latex


def code_content_clean(content):
    """清理代码内容，移除Markdown代码块的开始和结束标记"""
    if not content:
        return ""

    lines = content.splitlines()
    start_idx = 0
    end_idx = len(lines)

    # 处理开头的三个反引号
    if lines and lines[0].startswith("```"):
        start_idx = 1

    # 处理结尾的三个反引号
    if lines and end_idx > start_idx and lines[end_idx - 1].strip() == "```":
        end_idx -= 1

    # 只有在有内容时才进行join操作
    if start_idx < end_idx:
        return "\n".join(lines[start_idx:end_idx]).strip()
    return ""


def __tie_up_category_by_index(blocks, subject_block_type, object_block_type):
    """基于index的主客体关联包装函数"""
    # 定义获取主体和客体对象的函数
    def get_subjects():
        return list(
            map(
                lambda x: {"lines": x["lines"], "index": x["index"]},
                filter(
                    lambda x: x["type"] == subject_block_type,
                    blocks,
                ),
            )
        )

    def get_objects():
        return list(
            map(
                lambda x: {"lines": x["lines"], "index": x["index"]},
                filter(
                    lambda x: x["type"] == object_block_type,
                    blocks,
                ),
            )
        )

    # 调用通用方法
    return tie_up_category_by_index(
        get_subjects,
        get_objects,
        include_bbox=False,
    )


def get_type_blocks(blocks, block_type: Literal["image", "table", "chart"]):
    with_captions = __tie_up_category_by_index(blocks, f"{block_type}_body", f"{block_type}_caption")
    ret = []
    for v in with_captions:
        record = {
            f"{block_type}_body": v["sub_bbox"],
            f"{block_type}_caption_list": v["obj_bboxes"],
        }
        ret.append(record)
    return ret


def _filter_contiguous_captions(block: dict, fix_type: str) -> list:
    caption_list = block[f"{fix_type}_caption_list"]
    body_index = block[f"{fix_type}_body"]["index"]
    if not caption_list:
        return []
    caption_list.sort(key=lambda x: x["index"], reverse=True)
    filtered_captions = [caption_list[0]]
    for i in range(1, len(caption_list)):
        prev_index = caption_list[i - 1]["index"]
        curr_index = caption_list[i]["index"]
        gap_indices = set(range(curr_index + 1, prev_index))
        if curr_index == prev_index - 1 or gap_indices == {body_index}:
            filtered_captions.append(caption_list[i])
        else:
            block.setdefault("_not_include", []).extend(caption_list[i:])
            break
    filtered_captions.reverse()
    return filtered_captions


def _filter_caption_lists(need_fix_blocks: list, fix_type: str) -> list:
    not_include_blocks = []
    for block in need_fix_blocks:
        block[f"{fix_type}_caption_list"] = _filter_contiguous_captions(block, fix_type)
        not_include_blocks.extend(block.pop("_not_include", []))
    return not_include_blocks


def _build_two_layer_block(block: dict, fix_type: str, processed_indices: set) -> dict:
    body = block[f"{fix_type}_body"]
    caption_list = block[f"{fix_type}_caption_list"]
    body["type"] = f"{fix_type}_body"
    for caption in caption_list:
        caption["type"] = f"{fix_type}_caption"
        processed_indices.add(caption["index"])
    processed_indices.add(body["index"])
    two_layer_block = {"type": fix_type, "blocks": [body], "index": body["index"]}
    two_layer_block["blocks"].extend([*caption_list])
    two_layer_block["blocks"].sort(key=lambda x: x["index"])
    return two_layer_block


def _collect_unprocessed_blocks(blocks: list, processed_indices: set, not_include_blocks: list) -> None:
    for block in blocks:
        block.pop("type", None)
        if block["index"] not in processed_indices and block not in not_include_blocks:
            not_include_blocks.append(block)


def fix_two_layer_blocks(blocks, fix_type: Literal["image", "table", "chart"]):
    need_fix_blocks = get_type_blocks(blocks, fix_type)
    not_include_blocks = _filter_caption_lists(need_fix_blocks, fix_type)
    fixed_blocks = []
    processed_indices = set()
    for block in need_fix_blocks:
        fixed_blocks.append(_build_two_layer_block(block, fix_type, processed_indices))
    _collect_unprocessed_blocks(blocks, processed_indices, not_include_blocks)
    return fixed_blocks, not_include_blocks

CAPTION_PARENT_TYPES = ["table", "image", "chart"]
CAPTION_PREFIXES = {
    "table": ["?", "table"],
    "image": ["?", "fig"],
    "chart": ["?", "fig", "chart"],
}


def _mark_following_text_caption(page_blocks: list, index: int, block_type: str) -> None:
    if index + 1 >= len(page_blocks):
        return
    next_block = page_blocks[index + 1]
    if next_block.get("type") != "text":
        return
    content = next_block.get("content", "").strip().lower()
    prefixes = CAPTION_PREFIXES.get(block_type, [])
    if any(content.startswith(prefix.lower()) for prefix in prefixes):
        next_block = next_block.copy()
        next_block["type"] = "caption"
        page_blocks[index + 1] = next_block


def _mark_text_captions_after_media(page_blocks: list) -> None:
    for i, block in enumerate(page_blocks):
        block_type = block.get("type")
        if block_type in CAPTION_PARENT_TYPES:
            _mark_following_text_caption(page_blocks, i, block_type)


def _find_adjacent_caption_parent(page_blocks: list, start: int, step: int) -> str | None:
    i = start
    while 0 <= i < len(page_blocks):
        block_type = page_blocks[i].get("type")
        if block_type in CAPTION_PARENT_TYPES:
            return block_type
        if block_type != "caption":
            return None
        i += step
    return None


def _classify_caption_block(page_blocks: list, index: int, block: dict) -> dict:
    prev_parent_type = _find_adjacent_caption_parent(page_blocks, index - 1, -1)
    next_parent_type = _find_adjacent_caption_parent(page_blocks, index + 1, 1)
    new_block = block.copy()
    if prev_parent_type:
        new_block["type"] = f"{prev_parent_type}_caption"
    elif next_parent_type:
        new_block["type"] = f"{next_parent_type}_caption"
    else:
        new_block["type"] = "text"
    return new_block


def classify_caption_blocks(page_blocks: list) -> list:
    if not page_blocks:
        return page_blocks
    _mark_text_captions_after_media(page_blocks)
    result_blocks = []
    for i, block in enumerate(page_blocks):
        if block.get("type") == "caption":
            result_blocks.append(_classify_caption_block(page_blocks, i, block))
        else:
            result_blocks.append(block)
    return result_blocks

