# Copyright (c) Opendatalab. All rights reserved.
import posixpath
import re
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import BinaryIO, Optional, Union, Any, Final, Iterator
from zipfile import ZIP_DEFLATED, ZipFile

from loguru import logger
from docx import Document
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.text.paragraph import Paragraph
from docx.text.hyperlink import Hyperlink
from docx.text.run import Run
from lxml import etree
from pydantic import AnyUrl
from docx_parse.model.docx.tools.math.omml import oMath2Latex
from docx_parse.model.docx.table_xml import render_table_html
from docx_parse.utils.docx_formatting import Formatting, Script
from docx_parse.utils.enum_class import BlockType
from docx_parse.backend.utils.office_image import (
    serialize_office_image,
)
from docx_parse.backend.utils.office_chart import extract_chart_html_from_ooxml
from docx_parse.model.docx.docx_converter_structure_mixin import (
    DocxConverterStructureMixin,
)


class DocxConverter(DocxConverterStructureMixin):
    _BLIP_NAMESPACES: Final = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "v": "urn:schemas-microsoft-com:vml",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "w10": "urn:schemas-microsoft-com:office:word",
        "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
        "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    }
    _PARAGRAPH_TRANSPARENT_INLINE_CONTAINERS: Final = {
        "bdo",
        "customXml",
        "dir",
        "fldSimple",
        "ins",
        "moveTo",
        "smartTag",
    }
    """
    Word 文档中使用的 XML 命名空间映射。

    这些命名空间用于解析 DOCX 文件中的各种元素，包括：
    - a: DrawingML 主命名空间
    - r: Office 文档关系命名空间
    - w: WordprocessingML 主命名空间
    - wp: Wordprocessing Drawing 命名空间
    - mc: 标记兼容性命名空间
    - v: VML (Vector Markup Language) 命名空间
    - wps: Wordprocessing Shape 命名空间
    - w10: Office Word 命名空间
    - a14: Office 2010 Drawing 命名空间
    """

    def __init__(self):
        self.XML_KEY = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        )
        self.xml_namespaces = {
            "w": "http://schemas.microsoft.com/office/word/2003/wordml"
        }
        self.picture_xpath_expr = etree.XPath(
            ".//a:blip | .//v:imagedata", namespaces=DocxConverter._BLIP_NAMESPACES
        )

        self.docx_obj = None
        self.pages = []
        self.cur_page = []
        self.pre_num_id: int = -1  # 上一个处理元素的 numId
        self.pre_ilevel: int = -1  # 上一个处理元素的缩进等级, 用于判断列表层级
        self.list_block_stack: list = []  # 列表块堆栈
        self.list_counters: dict[tuple[int, int], int] = (
            {}
        )  # 列表计数器 (numId, ilvl) -> count
        self.index_block_stack: list = []  # 目录索引块堆栈
        self.pre_index_ilevel: int = -1  # 上一个目录项的缩进等级
        self.plain_toc_base_level: Optional[int] = None  # 普通目录段落的起始层级
        self.heading_list_numids: set = set()  # 用作章节标题的列表numId集合
        self.equation_bookends: str = "<eq>{EQ}</eq>"  # 公式标记格式
        self.processed_textbox_elements: list = []
        self.toc_anchor_set: set[str] = set()  # TOC 超链接目标锚点集合
        self._numbering_root: Optional[BaseOxmlElement] = None
        self._numbering_root_loaded: bool = False
        self._numbering_level_cache: dict[
            tuple[int, int], Optional[BaseOxmlElement]
        ] = {}
        self._style_chain_cache: dict[Any, tuple[Any, ...]] = {}
        self._paragraph_property_child_cache: dict[
            tuple[Any, str], Optional[BaseOxmlElement]
        ] = {}
        self._paragraph_label_level_cache: dict[
            Any, tuple[str, Optional[int]]
        ] = {}
        self._paragraph_num_cache: dict[
            Any, tuple[Optional[int], Optional[int]]
        ] = {}
        self._paragraph_toc_level_cache: dict[Any, Optional[int]] = {}
        self._paragraph_text_cache: dict[Any, str] = {}
        self._document_numbering_possible_cache: Optional[bool] = None
        self._header_footer_parts_cache: Optional[bool] = None
        self.tolerant: bool = False
        self.parse_errors: list[dict[str, Any]] = []
        self._max_parse_errors: int = 100

    def _record_parse_error(
        self,
        stage: str,
        exc: Exception,
        element: Optional[BaseOxmlElement] = None,
        tag_name: Optional[str] = None,
    ) -> None:
        if len(self.parse_errors) >= self._max_parse_errors:
            return
        if tag_name is None and element is not None:
            try:
                tag_name = etree.QName(element).localname
            except Exception:
                tag_name = None
        self.parse_errors.append(
            {
                "stage": stage,
                "tag": tag_name,
                "error_type": type(exc).__name__,
                "message": str(exc),
            }
        )

    def _run_tolerant(
        self,
        stage: str,
        func,
        *args,
        element: Optional[BaseOxmlElement] = None,
        tag_name: Optional[str] = None,
    ):
        try:
            return func(*args)
        except Exception as exc:
            if not self.tolerant:
                raise
            self._record_parse_error(stage, exc, element=element, tag_name=tag_name)
            logger.warning(
                f"Skipping DOCX {stage} due to {type(exc).__name__}: {exc}"
            )
            return None

    @staticmethod
    def _escape_hyperlink_text(text: str) -> str:
        """
        转义超链接文本中的方括号。

        Args:
            text: 要转义的文本

        Returns:
            str: 转义后的文本
        """
        if not text:
            return text
        # 转义方括号
        text = text.replace("[", "\\[").replace("]", "\\]")
        return text

    @staticmethod
    def _escape_hyperlink_url(url: str) -> str:
        """
        转义超链接 URL 中的括号。

        Args:
            url: 要转义的 URL

        Returns:
            str: 转义后的 URL
        """
        if not url:
            return url
        # 对括号进行 URL 编码
        url = url.replace("(", "%28").replace(")", "%29")
        return url

    @staticmethod
    def _get_style_str_from_format(format_obj) -> Optional[str]:
        """
        从 Formatting 对象提取样式字符串。

        Args:
            format_obj: Formatting 对象

        Returns:
            Optional[str]: 样式字符串（如 "bold,italic"），无样式时返回 None
        """
        if format_obj is None:
            return None
        styles = []
        if format_obj.bold:
            styles.append('bold')
        if format_obj.italic:
            styles.append('italic')
        if format_obj.underline:
            styles.append('underline')
        if format_obj.strikethrough:
            styles.append('strikethrough')
        return ','.join(styles) if styles else None

    @staticmethod
    def _has_visible_style(format_obj) -> bool:
        """
        检查格式是否包含可见样式（下划线或删除线）。

        空白文本在有这些样式时仍然是可见的，应当保留。

        Args:
            format_obj: Formatting 对象

        Returns:
            bool: 是否包含可见样式
        """
        if format_obj is None:
            return False
        return bool(format_obj.underline or format_obj.strikethrough)

    @staticmethod
    def _is_hidden_run(run: Run) -> bool:
        """Check whether a run is marked as hidden text in Word."""
        _W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        rpr = run._element.find(f"{{{_W}}}rPr")
        if rpr is None:
            return False
        # webHidden: commonly used by TOC page-number field runs
        if rpr.find(f"{{{_W}}}webHidden") is not None:
            return True
        # vanish: generic hidden text
        if rpr.find(f"{{{_W}}}vanish") is not None:
            return True
        return False

    @staticmethod
    def _text_tag(text: str, style_str: Optional[str] = None) -> str:
        if style_str:
            return f'<text style="{style_str}">{text}</text>'
        return f"<text>{text}</text>"

    @staticmethod
    def _plain_or_styled_text(text: str, style_str: Optional[str] = None) -> str:
        if style_str:
            return DocxConverter._text_tag(text, style_str)
        return text

    @staticmethod
    def _valid_hyperlink_str(hyperlink: Optional[Union[AnyUrl, Path, str]]) -> Optional[str]:
        if hyperlink is None:
            return None
        hyperlink_str = str(hyperlink)
        if not hyperlink_str or hyperlink_str.strip() == "" or hyperlink_str == ".":
            return None
        return hyperlink_str

    @classmethod
    def _format_text_with_hyperlink(
        cls,
        text: str,
        hyperlink: Optional[Union[AnyUrl, Path, str]],
        style_str: Optional[str] = None,
    ) -> str:
        """
        将文本和超链接格式化，支持字体样式标记。

        无超链接时：有样式包裹为 <text style="...">文本</text>，无样式直接返回文本。
        有超链接时：格式化为 <hyperlink><text [style="..."]>文本</text><url>链接</url></hyperlink>。

        Args:
            text: 文本内容
            hyperlink: 超链接地址
            style_str: 样式字符串（如 "bold,italic"），无样式时为 None

        Returns:
            str: 格式化后的文本
        """
        if not text:
            return text

        # 检查超链接是否有效（非空）
        if hyperlink is None:
            # 无超链接：只有有样式时才包裹 <text> 标签
            return cls._plain_or_styled_text(text, style_str)

        hyperlink_str = str(hyperlink)
        if not hyperlink_str or hyperlink_str.strip() == "" or hyperlink_str == ".":
            return cls._plain_or_styled_text(text, style_str)

        # 有超链接：构建 <text> 标签（含可选样式）
        text_tag = cls._text_tag(text, style_str)

        return f"<hyperlink>{text_tag}<url>{hyperlink_str}</url></hyperlink>"

    def _build_text_from_elements(
        self,
        paragraph_elements: list[
            tuple[str, Optional[Formatting], Optional[Union[AnyUrl, Path, str]]]
        ],
    ) -> str:
        """
        从 paragraph_elements 重组文本，应用超链接格式和字体样式。

        Args:
            paragraph_elements: 段落元素列表

        Returns:
            str: 重组后的文本
        """
        result_parts = []
        for text, format_obj, hyperlink in paragraph_elements:
            if text:
                style_str = self._get_style_str_from_format(format_obj)
                formatted_text = self._format_text_with_hyperlink(text, hyperlink, style_str)
                result_parts.append(formatted_text)
        return "".join(result_parts) if result_parts else ""

    @staticmethod
    def _normalize_text_block_content(content: str) -> str:
        """
        规范化普通文本块导出内容。

        DOCX 常用段首/段尾空格模拟版式对齐，导出普通文本块前去除这些前后空白。
        """
        if not content:
            return content
        return content.strip()

    @staticmethod
    def _paragraph_split_boundaries(non_eq_segments: list) -> set[int]:
        boundaries: set[int] = set()
        pos = 0
        for seg in non_eq_segments[:-1]:
            pos += len(seg)
            boundaries.add(pos)
        return boundaries

    @staticmethod
    def _paragraph_elements_match_segments(paragraph_elements: list, non_eq_segments: list) -> bool:
        concat_elem_text = "".join(text for text, _, _ in paragraph_elements)
        concat_seg_text = "".join(non_eq_segments)
        return concat_elem_text == concat_seg_text

    @staticmethod
    def _split_paragraph_element_at_boundaries(text: str, fmt, hyperlink, elem_start: int, boundaries: set[int]) -> list:
        elem_end = elem_start + len(text)
        splits = sorted(b - elem_start for b in boundaries if elem_start < b < elem_end)
        if not splits:
            return [(text, fmt, hyperlink)]
        result = []
        prev = 0
        for split_pos in splits:
            fragment = text[prev:split_pos]
            if fragment:
                result.append((fragment, fmt, hyperlink))
            prev = split_pos
        fragment = text[prev:]
        if fragment:
            result.append((fragment, fmt, hyperlink))
        return result

    @staticmethod
    def _split_paragraph_elements_at_eq_boundaries(paragraph_elements: list, non_eq_segments: list) -> list:
        if len(non_eq_segments) <= 1:
            return paragraph_elements
        boundaries = DocxConverter._paragraph_split_boundaries(non_eq_segments)
        if not boundaries:
            return paragraph_elements
        if not DocxConverter._paragraph_elements_match_segments(paragraph_elements, non_eq_segments):
            return paragraph_elements
        result = []
        text_pos = 0
        for text, fmt, hyperlink in paragraph_elements:
            if not text:
                result.append((text, fmt, hyperlink))
                text_pos += len(text)
                continue
            result.extend(
                DocxConverter._split_paragraph_element_at_boundaries(
                    text, fmt, hyperlink, text_pos, boundaries
                )
            )
            text_pos += len(text)
        return result

    @staticmethod
    def _paragraph_has_hyperlink(paragraph_elements: list) -> bool:
        return any(
            hyperlink is not None and str(hyperlink).strip() not in ("", ".")
            for _, _, hyperlink in paragraph_elements
        )

    @staticmethod
    def _paragraph_has_style(paragraph_elements: list) -> bool:
        return any(
            fmt is not None and (fmt.bold or fmt.italic or fmt.underline or fmt.strikethrough)
            for _, fmt, _ in paragraph_elements
        )

    def _build_element_replacements(self, paragraph_elements: list) -> list:
        element_mappings = []
        for text, format_obj, hyperlink in paragraph_elements:
            if text:
                style_str = self._get_style_str_from_format(format_obj)
                formatted_text = self._format_text_with_hyperlink(text, hyperlink, style_str)
                element_mappings.append((text, formatted_text))
        return element_mappings

    def _apply_element_replacements(self, result_text: str, element_mappings: list) -> str:
        for original_text, formatted_text in element_mappings:
            if original_text != formatted_text:
                result_text = self._replace_text_outside_equations(
                    result_text, original_text, formatted_text
                )
        return result_text

    def _build_text_with_equations_and_hyperlinks(
        self,
        paragraph_elements: list[
            tuple[str, Optional[Formatting], Optional[Union[AnyUrl, Path, str]]]
        ],
        text_with_equations: str,
        equations: list,
    ) -> str:
        if not equations:
            return self._build_text_from_elements(paragraph_elements)
        if not self._paragraph_has_hyperlink(paragraph_elements) and not self._paragraph_has_style(paragraph_elements):
            return text_with_equations
        eq_split_pattern = re.compile(r'<eq>.*?</eq>', re.DOTALL)
        non_eq_segments = eq_split_pattern.split(text_with_equations)
        paragraph_elements = self._split_paragraph_elements_at_eq_boundaries(
            paragraph_elements, non_eq_segments
        )
        return self._apply_element_replacements(
            text_with_equations,
            self._build_element_replacements(paragraph_elements),
        )

    def _replace_text_outside_equations(
        self, text: str, old_text: str, new_text: str
    ) -> str:
        """
        在公式标记外替换文本。

        Args:
            text: 原始文本
            old_text: 要替换的文本
            new_text: 替换后的文本

        Returns:
            str: 替换后的文本
        """
        # 分割文本为公式和非公式部分
        eq_pattern = re.compile(r"(<eq>.*?</eq>)")
        parts = eq_pattern.split(text)

        result_parts = []
        for part in parts:
            if part.startswith("<eq>") and part.endswith("</eq>"):
                # 公式部分，保持不变
                result_parts.append(part)
            else:
                # 非公式部分，进行替换
                result_parts.append(part.replace(old_text, new_text, 1))

        return "".join(result_parts)

    @staticmethod
    def _resolve_internal_relationship_target(
        rels_path: str, target: Optional[str]
    ) -> Optional[str]:
        """Resolve an OOXML relationship target to a package member path."""
        if not target:
            return None

        rels_posix = PurePosixPath(rels_path)
        if rels_posix.parent.name != "_rels":
            return None

        base_dir = rels_posix.parent.parent.as_posix()
        if target.startswith("/"):
            resolved = posixpath.normpath(target.lstrip("/"))
        else:
            resolved = posixpath.normpath(posixpath.join(base_dir, target))

        if resolved in {"", "."} or resolved.startswith("../"):
            return None
        return resolved

    def _remove_broken_relationships(self, source: ZipFile, info, package_members: set):
        try:
            root = etree.fromstring(source.read(info.filename))
        except Exception:
            return None, 0
        removed_count = 0
        for relationship in list(root):
            if etree.QName(relationship).localname != "Relationship":
                continue
            if relationship.get("TargetMode") == "External":
                continue
            resolved = self._resolve_internal_relationship_target(
                info.filename, relationship.get("Target")
            )
            if resolved is not None and resolved in package_members:
                continue
            root.remove(relationship)
            removed_count += 1
        return root, removed_count

    def _collect_rewritten_relationships(self, file_bytes: bytes) -> dict[str, bytes]:
        rewritten_rels: dict[str, bytes] = {}
        with ZipFile(BytesIO(file_bytes)) as source:
            package_members = set(source.namelist())
            for info in source.infolist():
                if not info.filename.endswith(".rels"):
                    continue
                root, removed_count = self._remove_broken_relationships(
                    source, info, package_members
                )
                if root is None or removed_count == 0:
                    continue
                logger.debug(
                    "Removed {} broken internal DOCX relationships from {}",
                    removed_count,
                    info.filename,
                )
                rewritten_rels[info.filename] = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone="yes"
                )
        return rewritten_rels

    @staticmethod
    def _write_sanitized_package(file_bytes: bytes, rewritten_rels: dict[str, bytes]) -> bytes:
        output = BytesIO()
        with ZipFile(BytesIO(file_bytes)) as source, ZipFile(output, "w", ZIP_DEFLATED) as target:
            for info in source.infolist():
                data = rewritten_rels.get(info.filename, source.read(info.filename))
                target.writestr(info, data)
        return output.getvalue()

    def _sanitize_missing_internal_relationships(self, file_bytes: bytes) -> bytes:
        try:
            rewritten_rels = self._collect_rewritten_relationships(file_bytes)
            if not rewritten_rels:
                return file_bytes
            return self._write_sanitized_package(file_bytes, rewritten_rels)
        except Exception:
            return file_bytes

    def _start_new_page(self) -> None:
        self.cur_page = []
        self.pages.append(self.cur_page)

    def _empty_paragraph_with_section(self, element: BaseOxmlElement, sect_pr) -> bool:
        if sect_pr is None:
            return False
        paragraph = Paragraph(element, self.docx_obj)
        if self._get_paragraph_text(paragraph).strip():
            return False
        return not self.picture_xpath_expr(element)

    @staticmethod
    def _is_continuous_section(sect_pr, w_ns: str) -> bool:
        sect_type = sect_pr.find(f"{{{w_ns}}}type")
        sect_val = (
            sect_type.get(f"{{{w_ns}}}val", "continuous")
            if sect_type is not None else "continuous"
        )
        return sect_val == "continuous"

    @staticmethod
    def _has_zero_section_margins(sect_pr, w_ns: str) -> bool:
        pg_mar = sect_pr.find(f"{{{w_ns}}}pgMar")
        if pg_mar is None:
            return False
        for attr in ("header", "footer", "top", "bottom", "left", "right"):
            if pg_mar.get(f"{{{w_ns}}}{attr}", "0") != "0":
                return False
        return True

    def _is_layout_only_section_break(self, element: BaseOxmlElement) -> bool:
        w_ns = DocxConverter._BLIP_NAMESPACES["w"]
        p_pr = element.find(f"{{{w_ns}}}pPr")
        sect_pr = p_pr.find(f"{{{w_ns}}}sectPr") if p_pr is not None else None
        if not self._empty_paragraph_with_section(element, sect_pr):
            return False
        if not self._is_continuous_section(sect_pr, w_ns):
            return False
        return self._has_zero_section_margins(sect_pr, w_ns)

    def convert(
        self,
        file_stream: BinaryIO,
        tolerant: bool = True,
    ):
        # 重置所有实例状态，确保同一实例多次调用 convert() 时不会残留上次的数据
        self.pages = []
        self.cur_page = []
        self.pre_num_id = -1
        self.pre_ilevel = -1
        self.list_block_stack = []
        self.list_counters = {}
        self.index_block_stack = []
        self.pre_index_ilevel = -1
        self.plain_toc_base_level = None
        self.heading_list_numids = set()
        self.processed_textbox_elements = []
        self.toc_anchor_set = set()
        self._numbering_root = None
        self._numbering_root_loaded = False
        self._numbering_level_cache = {}
        self._style_chain_cache = {}
        self._paragraph_property_child_cache = {}
        self._paragraph_label_level_cache = {}
        self._paragraph_num_cache = {}
        self._paragraph_toc_level_cache = {}
        self._paragraph_text_cache = {}
        self._document_numbering_possible_cache = None
        self._header_footer_parts_cache = None
        self.tolerant = tolerant
        self.parse_errors = []
        # 读取文件字节，先清理失效内部关系再交给 python-docx 解析。
        file_bytes = self._sanitize_missing_internal_relationships(file_stream.read())
        self.docx_obj = Document(BytesIO(file_bytes))
        self.toc_anchor_set = self._collect_toc_anchor_set()
        # 预扫描文档，识别用作章节标题的列表numId
        self.heading_list_numids = self._detect_heading_list_numids()
        self.pages.append(self.cur_page)
        self._walk_linear(self.docx_obj.element.body)
        if self._has_header_footer_parts():
            try:
                self._add_header_footer(self.docx_obj)
            except RecursionError as e:
                logger.warning(f"Skipping DOCX header/footer parsing due to recursive section references: {e}")
            except Exception as e:
                logger.warning(f"Skipping DOCX header/footer parsing: {e}")

    def _reset_index_state(self) -> None:
        """重置目录索引栈，避免相隔的多个目录块被错误合并。"""
        self.index_block_stack = []
        self.pre_index_ilevel = -1
        self.plain_toc_base_level = None

    def _collect_toc_anchor_set(self) -> set[str]:
        """Collect TOC hyperlink anchors from the entire document body."""
        anchor_attr = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor"
        )
        anchors: set[str] = set()
        for hl in self.docx_obj.element.body.findall(
            ".//w:hyperlink", namespaces=DocxConverter._BLIP_NAMESPACES
        ):
            anchor = hl.get(anchor_attr, "").strip()
            if anchor and anchor.startswith("_Toc"):
                anchors.add(anchor)
        return anchors

    def _has_header_footer_parts(self) -> bool:
        """Return True only when the DOCX package contains header/footer parts."""
        if self._header_footer_parts_cache is not None:
            return self._header_footer_parts_cache

        package = getattr(getattr(self.docx_obj, "part", None), "package", None)
        parts = getattr(package, "parts", []) if package is not None else []
        for part in parts:
            part_name = str(getattr(part, "partname", "")).lower()
            if "/header" in part_name or "/footer" in part_name:
                self._header_footer_parts_cache = True
                return True

        self._header_footer_parts_cache = False
        return False

    def _used_paragraph_style_ids(self) -> set[str]:
        body = self.docx_obj.element.body
        return {
            p_style.get(self.XML_KEY)
            for p_style in body.findall(".//w:pPr/w:pStyle", namespaces=DocxConverter._BLIP_NAMESPACES)
            if p_style.get(self.XML_KEY)
        }

    def _style_numbering_maps(self, styles_root) -> tuple[dict, dict, dict]:
        namespaces = DocxConverter._BLIP_NAMESPACES
        style_id_attr = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"
        style_elements = {}
        style_based_on = {}
        style_has_num_pr = {}
        for style in styles_root.findall("w:style", namespaces=namespaces):
            style_id = style.get(style_id_attr)
            if not style_id:
                continue
            style_elements[style_id] = style
            style_has_num_pr[style_id] = style.find(".//w:pPr/w:numPr", namespaces=namespaces) is not None
            based_on = style.find("w:basedOn", namespaces=namespaces)
            if based_on is not None and based_on.get(self.XML_KEY):
                style_based_on[style_id] = based_on.get(self.XML_KEY)
        return style_elements, style_based_on, style_has_num_pr

    @staticmethod
    def _style_chain_has_num_pr(style_id: str, style_based_on: dict, style_has_num_pr: dict) -> bool:
        seen: set[str] = set()
        current = style_id
        while current and current not in seen:
            seen.add(current)
            if style_has_num_pr.get(current):
                return True
            current = style_based_on.get(current)
        return False

    def _document_may_have_numbered_paragraphs(self) -> bool:
        if self._document_numbering_possible_cache is not None:
            return self._document_numbering_possible_cache
        body = self.docx_obj.element.body
        namespaces = DocxConverter._BLIP_NAMESPACES
        if body.find(".//w:pPr/w:numPr", namespaces=namespaces) is not None:
            self._document_numbering_possible_cache = True
            return True
        used_style_ids = self._used_paragraph_style_ids()
        if not used_style_ids:
            self._document_numbering_possible_cache = False
            return False
        styles_part = getattr(getattr(self.docx_obj, "part", None), "_styles_part", None)
        styles_root = getattr(styles_part, "element", None)
        if styles_root is None:
            self._document_numbering_possible_cache = False
            return False
        style_elements, style_based_on, style_has_num_pr = self._style_numbering_maps(styles_root)
        self._document_numbering_possible_cache = any(
            style_id in style_elements and self._style_chain_has_num_pr(style_id, style_based_on, style_has_num_pr)
            for style_id in used_style_ids
        )
        return self._document_numbering_possible_cache

    @staticmethod
    def _chart_relationship_types() -> tuple[set, set]:
        chart_rel_types = {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
        }
        package_rel_types = {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/package",
        }
        return chart_rel_types, package_rel_types

    def _chart_workbook_bytes(self, chart_part, package_rel_types: set):
        try:
            for rel in chart_part.rels.values():
                if rel.reltype in package_rel_types:
                    return rel.target_part.blob
        except Exception as e:
            logger.warning(f"Warning: chart workbook cannot be loaded: {e}")
        return None

    def _chart_xml_from_relationship(self, chart_rel):
        try:
            chart_part = chart_rel.target_part
            return chart_part, chart_part.blob
        except Exception as e:
            logger.warning(f"Warning: chart XML cannot be loaded: {e}")
            return None, None

    def _handle_chart_element(self, chart, chart_rel_types: set, package_rel_types: set) -> None:
        chart_block = {"type": BlockType.CHART, "content": ""}
        self.cur_page.append(chart_block)
        rel_id_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        rel_id = chart.get(rel_id_attr)
        if not rel_id:
            return
        try:
            chart_rel = self.docx_obj.part.rels[rel_id]
        except KeyError:
            return
        if chart_rel.reltype not in chart_rel_types:
            return
        chart_part, chart_xml = self._chart_xml_from_relationship(chart_rel)
        if chart_xml is None:
            return
        workbook_bytes = self._chart_workbook_bytes(chart_part, package_rel_types)
        try:
            chart_html = extract_chart_html_from_ooxml(chart_xml, workbook_bytes)
        except Exception as e:
            logger.warning(f"Warning: chart HTML cannot be extracted: {e}")
            return
        if chart_html:
            chart_block["content"] = chart_html

    def _handle_drawingml(self, elements: list[BaseOxmlElement]):
        chart_rel_types, package_rel_types = self._chart_relationship_types()
        for element in elements:
            chart = element.find(".//c:chart", namespaces=DocxConverter._BLIP_NAMESPACES)
            if chart is not None:
                self._handle_chart_element(chart, chart_rel_types, package_rel_types)

    def _handle_drawingml_for_element(self, element, tag_name: str) -> None:
        drawingml_els = element.findall(
            ".//w:drawing", namespaces=DocxConverter._BLIP_NAMESPACES
        )
        if drawingml_els:
            self._run_tolerant(
                "drawingml", self._handle_drawingml, drawingml_els,
                element=element, tag_name=tag_name,
            )

    def _textbox_elements_for_element(self, element, tag_name: str):
        txbx_xpath = etree.XPath(
            ".//w:txbxContent|.//v:textbox//w:p",
            namespaces=DocxConverter._BLIP_NAMESPACES,
        )
        textbox_elements = txbx_xpath(element)
        if textbox_elements or tag_name not in ["drawing", "pict"]:
            return textbox_elements
        alt_txbx_xpath = etree.XPath(
            ".//wps:txbx//w:p|.//w10:wrap//w:p|.//a:p//a:t",
            namespaces=DocxConverter._BLIP_NAMESPACES,
        )
        return alt_txbx_xpath(element)

    def _append_shape_text_if_present(self, element, textbox_elements: list) -> None:
        if textbox_elements:
            return
        shape_text_xpath = etree.XPath(
            ".//a:bodyPr/ancestor::*//a:t|.//a:txBody//a:t",
            namespaces=DocxConverter._BLIP_NAMESPACES,
        )
        shape_text_elements = shape_text_xpath(element)
        text_content = " ".join([t.text for t in shape_text_elements if t.text])
        text_content = self._normalize_text_block_content(text_content)
        if text_content.strip():
            logger.debug(f"Found shape text: {text_content[:50]}...")
            self.cur_page.append({"type": BlockType.TEXT, "content": text_content})

    def _handle_textbox_for_element(self, element, tag_name: str) -> None:
        if element in self.processed_textbox_elements:
            return
        textbox_elements = self._textbox_elements_for_element(element, tag_name)
        if not textbox_elements and tag_name in ["drawing", "pict"]:
            self._append_shape_text_if_present(element, textbox_elements)
        if not textbox_elements:
            return
        self.processed_textbox_elements.append(element)
        self.processed_textbox_elements.extend(textbox_elements)
        logger.debug(f"Found textbox content with {len(textbox_elements)} elements")
        self._run_tolerant(
            "textbox", self._handle_textbox_content, textbox_elements,
            element=element, tag_name=tag_name,
        )

    def _handle_table_element(self, element, tag_name: str) -> None:
        if self.pre_num_id != -1:
            self.pre_num_id = -1
            self.pre_ilevel = -1
            self.list_block_stack = []
            self.list_counters = {}
        try:
            self._handle_tables(element)
        except Exception as exc:
            if self.tolerant:
                self._record_parse_error("table", exc, element=element, tag_name=tag_name)
            logger.debug("could not parse a table, broken docx table")

    def _handle_picture_element(self, element, tag_name: str, picture_refs: Any) -> None:
        is_anchored = bool(element.findall(".//wp:anchor", namespaces=DocxConverter._BLIP_NAMESPACES))
        if is_anchored and tag_name == "p":
            self._run_tolerant("paragraph", self._handle_text_elements, element, element=element, tag_name=tag_name)
        self._run_tolerant("picture", self._handle_pictures, picture_refs, element=element, tag_name=tag_name)
        if not is_anchored and tag_name == "p":
            self._run_tolerant("paragraph", self._handle_text_elements, element, element=element, tag_name=tag_name)

    def _handle_sdt_element(self, element, tag_name: str) -> None:
        sdt_content = element.find(".//w:sdtContent", namespaces=DocxConverter._BLIP_NAMESPACES)
        if sdt_content is None:
            return
        if self._is_toc_sdt(element):
            self._run_tolerant("sdt_index", self._handle_sdt_as_index, sdt_content, element=element, tag_name=tag_name)
            return
        paragraphs = sdt_content.findall(".//w:p", namespaces=DocxConverter._BLIP_NAMESPACES)
        for p in paragraphs:
            self._run_tolerant("sdt_paragraph", self._handle_text_elements, p, element=p, tag_name="p")

    def _walk_linear_element(self, element) -> None:
        tag_name = etree.QName(element).localname
        picture_refs = self.picture_xpath_expr(element)
        self._handle_drawingml_for_element(element, tag_name)
        self._handle_textbox_for_element(element, tag_name)
        if tag_name == "tbl":
            self._handle_table_element(element, tag_name)
        elif picture_refs:
            self._handle_picture_element(element, tag_name, picture_refs)
        elif tag_name == "sdt":
            self._handle_sdt_element(element, tag_name)
        elif tag_name == "p":
            self._run_tolerant("paragraph", self._handle_text_elements, element, element=element, tag_name=tag_name)
        else:
            logger.debug(f"Ignoring element in DOCX with tag: {tag_name}")

    def _walk_linear(self, body: BaseOxmlElement):
        for element in body:
            self._walk_linear_element(element)

    def _handle_tables(self, element: BaseOxmlElement):
        """Render a DOCX table directly from OOXML into HTML."""
        html = self._normalize_table_colspans(render_table_html(self, element))
        table_block = {
            "type": BlockType.TABLE,
            "content": html,
        }
        self.cur_page.append(table_block)

    @staticmethod
    def _table_has_rowspan(table) -> bool:
        all_cells = table.find_all(['td', 'th'])
        return any(int(c.get('rowspan', 1)) > 1 for c in all_cells)

    @staticmethod
    def _table_row_col_counts(rows: list) -> list[int]:
        row_col_counts = []
        for row in rows:
            cells = row.find_all(['td', 'th'])
            row_col_counts.append(sum(int(c.get('colspan', 1)) for c in cells))
        return row_col_counts

    @staticmethod
    def _reduce_row_colspan(row, excess: int) -> bool:
        modified = False
        for cell in row.find_all(['td', 'th']):
            if excess <= 0:
                break
            span = int(cell.get('colspan', 1))
            if span <= 1:
                continue
            reduce_by = min(span - 1, excess)
            new_span = span - reduce_by
            if new_span == 1:
                cell.attrs.pop('colspan', None)
            else:
                cell['colspan'] = str(new_span)
            excess -= reduce_by
            modified = True
        return modified

    def _normalize_single_table_colspans(self, table, counter_cls) -> bool:
        rows = table.find_all('tr')
        if not rows or self._table_has_rowspan(table):
            return False
        row_col_counts = self._table_row_col_counts(rows)
        if not row_col_counts:
            return False
        count_freq = counter_cls(row_col_counts)
        if len(count_freq) == 1:
            return False
        target = count_freq.most_common(1)[0][0]
        modified = False
        for row, col_count in zip(rows, row_col_counts):
            if col_count > target:
                modified = self._reduce_row_colspan(row, col_count - target) or modified
        return modified

    def _normalize_table_colspans(self, html: str) -> str:
        try:
            from bs4 import BeautifulSoup
            from collections import Counter

            soup = BeautifulSoup(html, 'html.parser')
            modified = False
            for table in soup.find_all('table'):
                modified = self._normalize_single_table_colspans(table, Counter) or modified
            return str(soup) if modified else html
        except Exception as e:
            logger.debug(f"Failed to normalize table colspans: {e}")
            return html

    def _paragraph_section_end(self, element: BaseOxmlElement) -> bool:
        has_section_break = element.find(".//w:sectPr", namespaces=DocxConverter._BLIP_NAMESPACES) is not None
        if not has_section_break or self._is_layout_only_section_break(element):
            return False
        if element.text == "":
            self._start_new_page()
            return False
        return True

    def _reset_list_state(self) -> None:
        self.pre_num_id = -1
        self.pre_ilevel = -1
        self.list_block_stack = []
        self.list_counters = {}

    def _append_title_block(self, level: int, is_numbered_style: bool, content_text: str, anchor: Optional[str]) -> None:
        if content_text == "":
            return
        title_block = {
            "type": BlockType.TITLE,
            "level": level,
            "is_numbered_style": is_numbered_style,
            "content": content_text,
        }
        if anchor:
            title_block["anchor"] = anchor
        self.cur_page.append(title_block)

    def _append_text_block(self, content_text: str, anchor: Optional[str]) -> None:
        content_text = self._normalize_text_block_content(content_text)
        if content_text == "":
            return
        text_block = {"type": BlockType.TEXT, "content": content_text}
        if anchor:
            text_block["anchor"] = anchor
        self.cur_page.append(text_block)

    def _handle_plain_toc_result(self, is_section_end: bool) -> None:
        if self.pre_num_id != -1:
            self._reset_list_state()
        if is_section_end:
            self._start_new_page()

    def _handle_numbered_paragraph(self, ctx: dict) -> bool:
        numid = ctx["numid"]
        ilevel = ctx["ilevel"]
        if numid is None or ilevel is None or ctx["p_style_id"] in ["Title", "Heading"]:
            return False
        is_numbered = self._is_numbered_list(numid, ilevel)
        if numid in self.heading_list_numids:
            if self.pre_num_id != -1:
                self._reset_list_state()
            content_text = self._build_text_with_equations_and_hyperlinks(
                ctx["paragraph_elements"], ctx["text"], ctx["equations"]
            )
            self._append_title_block(ilevel + 1, is_numbered, content_text, ctx["anchor"])
        else:
            self._add_list_item(
                numid=numid, ilevel=ilevel, elements=ctx["paragraph_elements"],
                is_numbered=is_numbered, text=ctx["text"], equations=ctx["equations"],
            )
        return True

    def _paragraph_context(self, element: BaseOxmlElement) -> dict | None:
        paragraph = Paragraph(element, self.docx_obj)
        paragraph_elements = self._get_paragraph_elements(paragraph)
        paragraph_text = self._get_paragraph_text(paragraph)
        text, equations = self._handle_equations_in_text(element=element, text=paragraph_text)
        if text is None:
            return None
        p_style_id, p_level = self._get_label_and_level(paragraph)
        numid, ilevel = self._get_numId_and_ilvl(paragraph)
        return {
            "paragraph": paragraph,
            "paragraph_elements": paragraph_elements,
            "paragraph_text": paragraph_text,
            "anchor": self._extract_paragraph_bookmark(element),
            "text": text.strip(),
            "equations": equations,
            "p_style_id": p_style_id or "Normal",
            "p_level": p_level,
            "numid": None if numid == 0 else numid,
            "ilevel": ilevel,
        }

    def _is_numbered_heading_style(self, element: BaseOxmlElement, ctx: dict) -> bool:
        style_element = getattr(ctx["paragraph"].style, "element", None)
        if style_element is None:
            return "<w:numPr>" in element.xml
        return "<w:numPr>" in style_element.xml or "<w:numPr>" in element.xml

    def _handle_heading_paragraph_body(
        self, element: BaseOxmlElement, ctx: dict, content_text: str
    ) -> None:
        is_numbered_style = self._is_numbered_heading_style(element, ctx)
        level = ctx["p_level"] if ctx["p_level"] is not None else 2
        self._append_title_block(level, is_numbered_style, content_text, ctx["anchor"])

    def _is_plain_text_style(self, p_style_id: str) -> bool:
        plain_styles = [
            "Paragraph", "Normal", "Subtitle", "Author", "DefaultText",
            "ListParagraph", "ListBullet", "Quote",
        ]
        return p_style_id in plain_styles

    def _handle_styled_paragraph_body(self, element: BaseOxmlElement, ctx: dict) -> None:
        p_style_id = ctx["p_style_id"]
        content_text = self._build_text_with_equations_and_hyperlinks(
            ctx["paragraph_elements"], ctx["text"], ctx["equations"]
        )
        if p_style_id in ["Title"]:
            self._append_title_block(1, False, content_text, ctx["anchor"])
            return
        if "Heading" in p_style_id:
            self._handle_heading_paragraph_body(element, ctx, content_text)
            return
        if len(ctx["equations"]) > 0:
            self._handle_equation_paragraph(ctx, content_text)
            return
        if self._is_plain_text_style(p_style_id):
            self._append_text_block(content_text, ctx["anchor"])
            return
        if self._is_caption(element):
            self._append_caption_block(content_text)
            return
        self._append_text_block(content_text, ctx["anchor"])

    def _append_caption_block(self, content_text: str) -> None:
        if content_text != "":
            self.cur_page.append({"type": BlockType.CAPTION, "content": content_text})

    def _handle_equation_paragraph(self, ctx: dict, content_text: str) -> None:
        paragraph_text = ctx["paragraph_text"]
        text = ctx["text"]
        if (paragraph_text is None or len(paragraph_text.strip()) == 0) and len(text) > 0:
            self.cur_page.append({"type": BlockType.EQUATION, "content": text.replace("<eq>", "").replace("</eq>", "")})
        else:
            self._append_text_block(content_text, ctx["anchor"])

    def _handle_text_elements(self, element: BaseOxmlElement):
        is_section_end = self._paragraph_section_end(element)
        ctx = self._paragraph_context(element)
        if ctx is None:
            return None
        if self._handle_plain_toc_paragraph_as_index(
            paragraph=ctx["paragraph"], paragraph_element=element,
            paragraph_elements=ctx["paragraph_elements"], text=ctx["text"],
            equations=ctx["equations"],
        ):
            self._handle_plain_toc_result(is_section_end)
            return None
        self._reset_index_state()
        if self._handle_numbered_paragraph(ctx):
            return None
        if ctx["numid"] is None and self.pre_num_id != -1 and ctx["p_style_id"] not in ["Title", "Heading"]:
            self._reset_list_state()
        self._handle_styled_paragraph_body(element, ctx)
        if is_section_end:
            self._start_new_page()
        return None

    @staticmethod
    def _get_docx_image_rel_id(image: Any) -> Optional[str]:
        rel_id = image.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        )
        if not rel_id:
            rel_id = image.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
        return rel_id

    def _get_docx_image_part(self, image: Any) -> Optional[Any]:
        rel_id = self._get_docx_image_rel_id(image)
        if rel_id in self.docx_obj.part.rels:
            return self.docx_obj.part.rels[rel_id].target_part
        return None

    def _append_image_part_block(self, image_part: Any) -> None:
        img_base64 = serialize_office_image(
            image_part.blob,
            part_name=getattr(image_part, "partname", None),
            content_type=getattr(image_part, "content_type", None),
        )
        if img_base64 is not None:
            self.cur_page.append({"type": BlockType.IMAGE, "content": img_base64})

    def _handle_pictures(self, picture_refs: Any):
        seen_rel_ids: set[str] = set()
        for image in picture_refs:
            rel_id = self._get_docx_image_rel_id(image)
            if rel_id and rel_id in seen_rel_ids:
                continue
            if rel_id:
                seen_rel_ids.add(rel_id)
            image_part = self._get_docx_image_part(image)
            if image_part is None:
                logger.warning("Warning: image cannot be found")
                continue
            self._append_image_part_block(image_part)

    @staticmethod
    def _plain_paragraph_properties_allowed(child, w_ns: str) -> bool:
        disallowed = ("pStyle", "numPr", "outlineLvl")
        for tag_name in disallowed:
            if child.find(f"{{{w_ns}}}{tag_name}") is not None:
                return False
        return True

    @staticmethod
    def _append_plain_run_child(run_child, text_parts: list, allowed_run_property_tags: set) -> bool:
        run_child_name = etree.QName(run_child).localname
        if run_child_name == "rPr":
            return DocxConverter._plain_run_properties_allowed(
                run_child, allowed_run_property_tags
            )
        text_by_tag = {"t": run_child.text or "", "tab": "\t", "br": "\n", "cr": "\n"}
        if run_child_name not in text_by_tag:
            return False
        text_parts.append(text_by_tag[run_child_name])
        return True

    @staticmethod
    def _plain_run_properties_allowed(run_properties, allowed_run_property_tags: set) -> bool:
        for prop in run_properties:
            if etree.QName(prop).localname not in allowed_run_property_tags:
                return False
        return True

    def _get_plain_paragraph_text_fast(self, paragraph: Paragraph) -> Optional[str]:
        """Return plain text for simple paragraphs that cannot affect output styling."""
        paragraph_element = paragraph._element
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        allowed_run_property_tags = {
            "color",
            "lang",
            "noProof",
            "rFonts",
            "sz",
            "szCs",
        }
        text_parts = []
        has_run = False

        for child in paragraph_element:
            child_name = etree.QName(child).localname
            if child_name == "pPr":
                if not self._plain_paragraph_properties_allowed(child, w_ns):
                    return None
                continue
            if child_name != "r":
                return None

            has_run = True
            for run_child in child:
                if not self._append_plain_run_child(
                    run_child, text_parts, allowed_run_property_tags
                ):
                    return None

        if not has_run:
            return ""
        return "".join(text_parts)

    def _paragraph_empty_elements(self, paragraph: Paragraph, inner_contents: list, paragraph_text: str):
        if paragraph_text.strip() != "":
            return None
        has_visible_style_run = any(
            isinstance(c, Run) and c.text and self._has_visible_style(self._get_format_from_run(c))
            for c in inner_contents
        )
        return None if has_visible_style_run else [("", None, None)]

    def _flush_group(self, paragraph_elements: list, group_text: str, previous_format) -> None:
        has_visible = len(group_text.strip()) > 0 or (
            group_text and self._has_visible_style(previous_format)
        )
        if has_visible:
            paragraph_elements.append((group_text, previous_format, None))

    def _append_hyperlink_runs(self, paragraph_elements: list, hyperlink: Any, runs: list) -> None:
        for h_run in runs:
            if self._is_hidden_run(h_run):
                continue
            h_text = h_run.text or ""
            h_format = self._get_format_from_run(h_run)
            if h_text != "" or self._has_visible_style(h_format):
                paragraph_elements.append((h_text, h_format, hyperlink))

    @staticmethod
    def _hyperlink_target(content: Hyperlink):
        address = content.address
        if address and "://" in address:
            return address
        return Path(address) if address else Path(".")

    def _handle_hyperlink_content(self, content: Hyperlink, state: dict) -> bool:
        hyperlink = self._hyperlink_target(content)
        if content.runs and len(content.runs) > 0:
            self._flush_group(state["paragraph_elements"], state["group_text"], state["previous_format"])
            state["group_text"] = ""
            self._append_hyperlink_runs(state["paragraph_elements"], hyperlink, content.runs)
            return True
        state.update({"text": content.text, "format": None, "hyperlink": hyperlink})
        return False

    def _reset_field_state(self, state: dict) -> None:
        state["field_in"] = False
        state["field_url"] = None
        state["field_phase"] = None
        state["field_acc_text"] = ""
        state["field_acc_format"] = None

    def _handle_field_char(self, run: Run, state: dict, w_ns: str) -> bool:
        fld_char = run._element.find(f"{{{w_ns}}}fldChar")
        if fld_char is None:
            return False
        fld_type = fld_char.get(f"{{{w_ns}}}fldCharType")
        if fld_type == "begin":
            state.update({"field_in": True, "field_url": None, "field_phase": "instr", "field_acc_text": "", "field_acc_format": None})
            return True
        if fld_type == "separate":
            state["field_phase"] = "result"
            return True
        if fld_type != "end":
            return True
        self._finish_field_char(state)
        return state.get("text") is None

    def _finish_field_char(self, state: dict) -> None:
        acc_text = state["field_acc_text"]
        if state["field_url"] and acc_text.strip():
            state.update({"text": acc_text, "hyperlink": state["field_url"], "format": state["field_acc_format"]})
        elif acc_text.strip():
            state.update({"text": acc_text, "hyperlink": None, "format": state["field_acc_format"]})
        else:
            state["text"] = None
        self._reset_field_state(state)

    def _handle_run_field_content(self, run: Run, state: dict, w_ns: str) -> bool:
        instr_elem = run._element.find(f"{{{w_ns}}}instrText")
        if instr_elem is not None and state["field_phase"] == "instr":
            if instr_elem.text:
                match = re.search(r'HYPERLINK\s+"([^"]+)"', instr_elem.text)
                if match:
                    state["field_url"] = match.group(1)
            return True
        if state["field_in"] and state["field_phase"] == "result":
            t_elem = run._element.find(f"{{{w_ns}}}t")
            if t_elem is not None:
                state["field_acc_text"] += run.text
                if state["field_acc_format"] is None:
                    state["field_acc_format"] = self._get_format_from_run(run)
            return True
        return False

    def _handle_run_content(self, run: Run, state: dict, w_ns: str) -> bool:
        state.update({"text": None, "hyperlink": None, "format": None})
        if run._element.find(f"{{{w_ns}}}fldChar") is not None:
            self._handle_field_char(run, state, w_ns)
            return state.get("text") is None
        if self._handle_run_field_content(run, state, w_ns):
            return True
        state.update({"text": run.text, "hyperlink": None, "format": self._get_format_from_run(run)})
        return False

    def _append_grouped_content(self, state: dict) -> None:
        text = state["text"]
        fmt = state["format"]
        hyperlink = state["hyperlink"]
        has_visible = len(text.strip()) > 0 or self._has_visible_style(fmt)
        if (has_visible and fmt != state["previous_format"]) or hyperlink is not None:
            self._flush_group(state["paragraph_elements"], state["group_text"], state["previous_format"])
            state["group_text"] = ""
            if hyperlink is not None:
                state["paragraph_elements"].append((text.strip(), fmt, hyperlink))
                text = ""
            else:
                state["previous_format"] = fmt
        state["group_text"] += text

    def _get_paragraph_elements(self, paragraph: Paragraph):
        plain_text = self._get_plain_paragraph_text_fast(paragraph)
        if plain_text is not None:
            self._paragraph_text_cache[paragraph._element] = plain_text
            return [(plain_text, None, None)]
        inner_contents = list(self._iter_paragraph_inner_content(paragraph))
        empty_result = self._paragraph_empty_elements(
            paragraph, inner_contents, self._get_paragraph_text_from_contents(inner_contents)
        )
        if empty_result is not None:
            return empty_result
        state = {"paragraph_elements": [], "group_text": "", "previous_format": None}
        self._reset_field_state(state)
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for content in inner_contents:
            if self._paragraph_content_consumed(content, state, w_ns):
                continue
            self._append_grouped_content(state)
        self._flush_group(state["paragraph_elements"], state["group_text"], state["previous_format"])
        return state["paragraph_elements"]

    def _paragraph_content_consumed(self, content, state: dict, w_ns: str) -> bool:
        if isinstance(content, Hyperlink):
            return self._handle_hyperlink_content(content, state)
        if isinstance(content, Run):
            return self._handle_run_content(content, state, w_ns)
        return True

    def _iter_paragraph_inner_content(
        self,
        paragraph: Paragraph,
        container: Optional[BaseOxmlElement] = None,
    ) -> Iterator[Union[Run, Hyperlink]]:
        """Yield visible paragraph inline containers in document order.

        python-docx only walks direct ``w:r`` and ``w:hyperlink`` children of ``w:p``.
        Inline ``w:sdt`` content controls are skipped entirely, which drops their text
        from both ``paragraph.text`` and ``paragraph.iter_inner_content()``. This walker
        treats ``w:sdt`` and a few transparent wrapper nodes as pass-through containers
        and reuses the existing Run/Hyperlink wrappers for the actual visible content.
        """
        if container is None:
            container = paragraph._element

        _W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        for child in container:
            yield from self._iter_paragraph_child_content(paragraph, child, _W_NS)

    def _iter_paragraph_child_content(
        self,
        paragraph: Paragraph,
        child: BaseOxmlElement,
        w_ns: str,
    ) -> Iterator[Union[Run, Hyperlink]]:
        tag_name = etree.QName(child).localname
        if tag_name == "r":
            yield Run(child, paragraph)
            return
        if tag_name == "hyperlink":
            yield Hyperlink(child, paragraph)
            return
        if tag_name == "sdt":
            yield from self._iter_sdt_inner_content(paragraph, child, w_ns)
            return
        if tag_name in self._PARAGRAPH_TRANSPARENT_INLINE_CONTAINERS:
            yield from self._iter_paragraph_inner_content(paragraph, child)

    def _iter_sdt_inner_content(
        self,
        paragraph: Paragraph,
        child: BaseOxmlElement,
        w_ns: str,
    ) -> Iterator[Union[Run, Hyperlink]]:
        sdt_content = child.find(f"{{{w_ns}}}sdtContent")
        if sdt_content is not None:
            yield from self._iter_paragraph_inner_content(paragraph, sdt_content)

    @staticmethod
    def _get_paragraph_text_from_contents(
        inner_contents: list[Union[Run, Hyperlink]],
    ) -> str:
        """Rebuild paragraph plain text from visible inline containers."""
        return "".join(content.text or "" for content in inner_contents)

    def _get_paragraph_text(self, paragraph: Paragraph) -> str:
        """Return paragraph plain text, including inline ``w:sdt`` content."""
        cache_key = paragraph._element
        cached = self._paragraph_text_cache.get(cache_key)
        if cached is not None:
            return cached

        text = self._get_paragraph_text_from_contents(
            list(self._iter_paragraph_inner_content(paragraph))
        )
        self._paragraph_text_cache[cache_key] = text
        return text

    @classmethod
    def _resolve_style_chain_bool(
        cls,
        style_obj,
        attr_name: str,
    ) -> Optional[bool]:
        """从样式继承链中解析布尔字体属性。"""
        style = style_obj
        while style is not None:
            font = getattr(style, "font", None)
            if font is not None:
                if attr_name == "underline":
                    value = font.underline
                elif attr_name == "strikethrough":
                    value = font.strike
                else:
                    value = getattr(font, attr_name, None)
                if value is not None:
                    return bool(value)
            style = getattr(style, "base_style", None)
        return None

    @classmethod
    def _resolve_run_bool_with_inheritance(
        cls,
        run: Run,
        attr_name: str,
    ) -> bool:
        """解析 run 的字体属性，支持 run/字符样式/段落样式继承。"""
        if attr_name == "underline":
            direct_value = run.underline
        elif attr_name == "strikethrough":
            direct_value = run.font.strike
        else:
            direct_value = getattr(run, attr_name, None)

        if direct_value is not None:
            return bool(direct_value)

        # 先看 run 级字符样式链（跳过 Hyperlink 默认字符样式，避免把默认下划线
        # 误当作正文强调样式注入到解析结果中）
        run_style = getattr(run, "style", None)
        run_style_id = str(getattr(run_style, "style_id", "") or "").lower()
        run_style_name = str(getattr(run_style, "name", "") or "").lower()
        is_hyperlink_style = (
            run_style_id == "hyperlink" or "hyperlink" in run_style_name
        )
        if not is_hyperlink_style:
            inherited = cls._resolve_style_chain_bool(run_style, attr_name)
            if inherited is not None:
                return inherited

        # 再看所在段落样式链
        parent = getattr(run, "_parent", None)
        inherited = cls._resolve_style_chain_bool(getattr(parent, "style", None), attr_name)
        if inherited is not None:
            return inherited

        return False

    @classmethod
    def _get_format_from_run(cls, run: Run) -> Optional[Formatting]:
        """
        从 Run 对象获取格式信息。

        Args:
            run: Run 对象

        Returns:
            Optional[Formatting]: 格式对象
        """
        is_bold = cls._resolve_run_bool_with_inheritance(run, "bold")
        is_italic = cls._resolve_run_bool_with_inheritance(run, "italic")
        is_strikethrough = cls._resolve_run_bool_with_inheritance(run, "strikethrough")
        is_underline = cls._resolve_run_bool_with_inheritance(run, "underline")

        # 检测着重符号 (w:em)：若存在非 none 的 em 值，则视为下划线样式
        _W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        rPr = run._element.find(f'{{{_W}}}rPr')
        if rPr is not None:
            em = rPr.find(f'{{{_W}}}em')
            if em is not None:
                em_val = em.get(f'{{{_W}}}val', '')
                if em_val and em_val != 'none':
                    is_underline = True

        is_sub = run.font.subscript or False
        is_sup = run.font.superscript or False
        script = Script.SUB if is_sub else Script.SUPER if is_sup else Script.BASELINE

        return Formatting(
            bold=is_bold,
            italic=is_italic,
            underline=is_underline,
            strikethrough=is_strikethrough,
            script=script,
        )

    def _collect_text_and_equation_parts(self, element) -> tuple[list, list, list]:
        only_texts = []
        only_equations = []
        texts_and_equations = []
        for subt in element.iter():
            tag_name = etree.QName(subt).localname
            if tag_name == "t" and "math" not in subt.tag:
                if isinstance(subt.text, str):
                    only_texts.append(subt.text)
                    texts_and_equations.append(subt.text)
            elif "oMath" in subt.tag and "oMathPara" not in subt.tag:
                equation = self._latex_equation_from_omml(subt)
                if equation:
                    only_equations.append(equation)
                    texts_and_equations.append(equation)
        return only_texts, only_equations, texts_and_equations

    def _latex_equation_from_omml(self, element) -> str:
        try:
            latex_equation = str(oMath2Latex(element)).strip()
        except Exception as e:
            logger.debug(f"Failed to convert OMML equation to LaTeX: {e}")
            return ""
        if not latex_equation:
            return ""
        return self.equation_bookends.format(EQ=latex_equation)

    @staticmethod
    def _reconstructed_text_matches(only_texts: list, text: str) -> bool:
        return re.sub(r"\s+", "", "".join(only_texts)).strip() == re.sub(r"\s+", "", text).strip()

    @staticmethod
    def _insert_equation_parts(output_text: str, texts_and_equations: list) -> str:
        init_i = 0
        for i_substr, substr in enumerate(texts_and_equations):
            if len(substr) == 0:
                continue
            if substr in output_text[init_i:]:
                init_i += output_text[init_i:].find(substr) + len(substr)
            elif i_substr > 0:
                output_text = output_text[:init_i] + substr + output_text[init_i:]
                init_i += len(substr)
            else:
                output_text = substr + output_text
        return output_text

    def _handle_equations_in_text(self, element, text):
        only_texts, only_equations, texts_and_equations = self._collect_text_and_equation_parts(element)
        if len(only_equations) < 1:
            return text, []
        if not self._reconstructed_text_matches(only_texts, text):
            return text, []
        output_text = self._insert_equation_parts(text[:], texts_and_equations)
        return output_text, only_equations

    def _get_label_and_level(self, paragraph: Paragraph) -> tuple[str, Optional[int]]:
        """Return compact metadata for this paragraph or SDT."""
        cache_key = paragraph._element
        cached = self._paragraph_label_level_cache.get(cache_key)
        if cached is not None:
            return cached

        result = self._resolve_label_and_level(paragraph)
        self._paragraph_label_level_cache[cache_key] = result
        return result

    def _resolve_label_and_level(self, paragraph: Paragraph) -> tuple[str, Optional[int]]:
        """Resolve paragraph label and level without cache handling."""
        if paragraph.style is None:
            return ("Normal", None)

        label = paragraph.style.style_id
        name = paragraph.style.name

        if label is None:
            return ("Normal", None)

        for style in self._iter_style_chain(paragraph.style):
            result = self._label_and_level_from_style(style)
            if result is not None:
                return result

        outline_level = self._get_effective_outline_level(paragraph)
        if outline_level is not None:
            return ("Heading", outline_level + 1)

        return (name or label or "Normal", None)

    def _label_and_level_from_style(self, style: Any) -> Optional[tuple[str, Optional[int]]]:
        style_label = getattr(style, "style_id", None)
        style_name = getattr(style, "name", None)
        if style_label and ":" in style_label:
            parts = style_label.split(":")
            if len(parts) == 2:
                return (parts[0], self._str_to_int(parts[1], None))
        for candidate in (style_label, style_name):
            if candidate and "heading" in candidate.lower():
                return self._get_heading_and_level(candidate)
        return None

    def _iter_style_chain(self, style: Any) -> Iterator[Any]:
        """Yield a style and its base-style chain once each."""
        if style is None:
            return

        cache_key = getattr(style, "element", None)
        if cache_key is None:
            cache_key = (
                getattr(style, "style_id", None),
                getattr(style, "name", None),
            )
        cached = self._style_chain_cache.get(cache_key)
        if cached is not None:
            yield from cached
            return

        seen: set[int] = set()
        current = style
        chain = []
        while current is not None:
            current_id = id(current)
            if current_id in seen:
                break
            seen.add(current_id)
            chain.append(current)
            current = getattr(current, "base_style", None)
        self._style_chain_cache[cache_key] = tuple(chain)
        yield from chain

    def _get_paragraph_property_child(
        self, xml_element: Optional[BaseOxmlElement], child_tag: str
    ) -> Optional[BaseOxmlElement]:
        """Read a direct child from w:pPr without matching nested descendants."""
        if xml_element is None:
            return None

        cache_key = (xml_element, child_tag)
        if cache_key in self._paragraph_property_child_cache:
            return self._paragraph_property_child_cache[cache_key]

        namespaces = getattr(xml_element, "nsmap", None) or DocxConverter._BLIP_NAMESPACES
        pPr = xml_element.find("w:pPr", namespaces=namespaces)
        if pPr is None:
            self._paragraph_property_child_cache[cache_key] = None
            return None
        child = pPr.find(child_tag, namespaces=namespaces)
        self._paragraph_property_child_cache[cache_key] = child
        return child

    def _get_effective_numPr(
        self, paragraph: Paragraph
    ) -> Optional[BaseOxmlElement]:
        """Resolve paragraph numbering from direct properties, then style inheritance."""
        numPr = self._get_paragraph_property_child(paragraph._element, "w:numPr")
        if numPr is not None:
            return numPr

        for style in self._iter_style_chain(getattr(paragraph, "style", None)):
            style_element = getattr(style, "element", None)
            numPr = self._get_paragraph_property_child(style_element, "w:numPr")
            if numPr is not None:
                return numPr

        return None

    def _get_effective_outline_level(self, paragraph: Paragraph) -> Optional[int]:
        """Resolve outline level from paragraph properties or inherited styles."""
        outline_lvl = self._get_paragraph_property_child(
            paragraph._element, "w:outlineLvl"
        )
        if outline_lvl is None:
            for style in self._iter_style_chain(getattr(paragraph, "style", None)):
                style_element = getattr(style, "element", None)
                outline_lvl = self._get_paragraph_property_child(
                    style_element, "w:outlineLvl"
                )
                if outline_lvl is not None:
                    break

        if outline_lvl is None:
            return None

        level = self._str_to_int(outline_lvl.get(self.XML_KEY), None)
        if level is None or level >= 9:
            return None
        return level

    def _get_numId_and_ilvl(
        self, paragraph: Paragraph
    ) -> tuple[Optional[int], Optional[int]]:
        """
        获取段落的列表编号ID和层级。

        Args:
            paragraph: 段落对象

        Returns:
            tuple[Optional[int], Optional[int]]: (numId, ilvl) 元组
        """
        cache_key = paragraph._element
        cached = self._paragraph_num_cache.get(cache_key)
        if cached is not None:
            return cached

        numPr = self._get_effective_numPr(paragraph)

        if numPr is not None:
            # 获取 numId 元素并提取值
            namespaces = getattr(numPr, "nsmap", None) or DocxConverter._BLIP_NAMESPACES
            numId_elem = numPr.find("w:numId", namespaces=namespaces)
            ilvl_elem = numPr.find("w:ilvl", namespaces=namespaces)
            numId = numId_elem.get(self.XML_KEY) if numId_elem is not None else None
            ilvl = ilvl_elem.get(self.XML_KEY) if ilvl_elem is not None else None

            result = (self._str_to_int(numId, None), self._str_to_int(ilvl, None))
            self._paragraph_num_cache[cache_key] = result
            return result

        return None, None  # 如果段落不是列表的一部分

    def _get_numbering_root(self) -> Optional[BaseOxmlElement]:
        """Load and cache word/numbering.xml once per conversion."""
        if self._numbering_root_loaded:
            return self._numbering_root

        self._numbering_root_loaded = True

        if not hasattr(self.docx_obj, "part") or not hasattr(self.docx_obj.part, "package"):
            return None

        for part in self.docx_obj.part.package.parts:
            if "numbering" in part.partname:
                self._numbering_root = part.element
                break

        return self._numbering_root

    def _get_numbering_level_definition(
        self, numId: int, ilvl: int
    ) -> Optional[BaseOxmlElement]:
        """Resolve and cache the numbering level definition for a numId/ilvl pair."""
        cache_key = (numId, ilvl)
        if cache_key in self._numbering_level_cache:
            return self._numbering_level_cache[cache_key]

        numbering_root = self._get_numbering_root()
        namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }
        lvl_element: Optional[BaseOxmlElement] = None

        if numbering_root is not None:
            num_xpath = f".//w:num[@w:numId='{numId}']"
            num_element = numbering_root.find(num_xpath, namespaces=namespaces)

            if num_element is not None:
                abstract_num_id_elem = num_element.find(
                    ".//w:abstractNumId", namespaces=namespaces
                )
                if abstract_num_id_elem is not None:
                    abstract_num_id = abstract_num_id_elem.get(self.XML_KEY)
                    if abstract_num_id is not None:
                        abstract_num_xpath = (
                            f".//w:abstractNum[@w:abstractNumId='{abstract_num_id}']"
                        )
                        abstract_num_element = numbering_root.find(
                            abstract_num_xpath, namespaces=namespaces
                        )
                        if abstract_num_element is not None:
                            lvl_xpath = f".//w:lvl[@w:ilvl='{ilvl}']"
                            lvl_element = abstract_num_element.find(
                                lvl_xpath, namespaces=namespaces
                            )

        self._numbering_level_cache[cache_key] = lvl_element
        return lvl_element

    def _is_numbered_list(self, numId: int, ilvl: int) -> bool:
        """
        根据 numFmt 值检查列表是否为编号列表。

        Args:
            numId: 列表编号ID
            ilvl: 列表层级

        Returns:
            bool: 如果是编号列表返回 True，否则返回 False
        """
        try:
            lvl_element = self._get_numbering_level_definition(numId, ilvl)
            if lvl_element is None:
                return False
            namespaces = {
                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            }

            # 获取 numFmt 元素
            num_fmt_element = lvl_element.find(".//w:numFmt", namespaces=namespaces)
            if num_fmt_element is None:
                return False

            num_fmt = num_fmt_element.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
            )

            # 编号格式包括: decimal, lowerRoman, upperRoman, lowerLetter, upperLetter
            # 项目符号格式包括: bullet
            numbered_formats = {
                "decimal",
                "lowerRoman",
                "upperRoman",
                "lowerLetter",
                "upperLetter",
                "decimalZero",
            }

            return num_fmt in numbered_formats

        except Exception as e:
            logger.debug(f"Error determining if list is numbered: {e}")
            return False

    @staticmethod
    def _new_list_block(list_attribute: str, ilevel: int) -> dict:
        return {"type": BlockType.LIST, "attribute": list_attribute, "content": [], "ilevel": ilevel}

    @staticmethod
    def _new_list_text_item(content_text: str) -> dict:
        return {"type": BlockType.TEXT, "content": content_text}

    def _start_top_level_list(self, numid: int, ilevel: int, list_attribute: str, content_text: str) -> None:
        if self.pre_num_id != -1:
            self._reset_list_state()
        self._reset_list_counters_for_new_sequence(numid)
        list_block = self._new_list_block(list_attribute, ilevel)
        self.cur_page.append(list_block)
        self.list_block_stack.append(list_block)
        list_block["content"].append(self._new_list_text_item(content_text))
        self.pre_num_id = numid
        self.pre_ilevel = ilevel

    def _add_increased_indent_list_item(self, numid: int, ilevel: int, list_attribute: str, content_text: str) -> None:
        child_list_block = self._new_list_block(list_attribute, ilevel)
        if not self.list_block_stack:
            logger.warning(
                "Missing DOCX list parent for increased indent; "
                f"numid={numid}, ilevel={ilevel}. Starting a new list block."
            )
            self.cur_page.append(child_list_block)
        else:
            self.list_block_stack[-1]["content"].append(child_list_block)
        self.list_block_stack.append(child_list_block)
        child_list_block["content"].append(self._new_list_text_item(content_text))
        self.pre_ilevel = ilevel

    def _add_decreased_indent_list_item(self, numid: int, ilevel: int, list_attribute: str, content_text: str) -> None:
        while self.list_block_stack and self.list_block_stack[-1]["ilevel"] != ilevel:
            self.list_block_stack.pop()
        if not self.list_block_stack:
            logger.warning(
                "Malformed DOCX list nesting; "
                f"numid={numid}, ilevel={ilevel}. Starting a new list block."
            )
            list_block = self._new_list_block(list_attribute, ilevel)
            self.cur_page.append(list_block)
            self.list_block_stack.append(list_block)
        else:
            list_block = self.list_block_stack[-1]
        list_block["content"].append(self._new_list_text_item(content_text))
        self.pre_ilevel = ilevel

    def _add_same_indent_list_item(self, numid: int, ilevel: int, list_attribute: str, content_text: str) -> None:
        if not self.list_block_stack:
            logger.warning(
                "Missing DOCX list block for same indent; "
                f"numid={numid}, ilevel={ilevel}. Starting a new list block."
            )
            list_block = self._new_list_block(list_attribute, ilevel)
            self.cur_page.append(list_block)
            self.list_block_stack.append(list_block)
        else:
            list_block = self.list_block_stack[-1]
        list_block["content"].append(self._new_list_text_item(content_text))

    def _add_list_item_to_current_state(
        self, numid: int, ilevel: int, list_attribute: str, content_text: str
    ) -> None:
        if self.pre_num_id == -1 or self.pre_num_id != numid:
            self._start_top_level_list(numid, ilevel, list_attribute, content_text)
            return
        if self.pre_ilevel != -1 and self.pre_ilevel < ilevel:
            self._add_increased_indent_list_item(numid, ilevel, list_attribute, content_text)
            return
        if self.pre_ilevel != -1 and ilevel < self.pre_ilevel:
            self._add_decreased_indent_list_item(numid, ilevel, list_attribute, content_text)
            return
        if self.pre_ilevel == ilevel:
            self._add_same_indent_list_item(numid, ilevel, list_attribute, content_text)
            return
        self._warn_unexpected_list_state(numid, ilevel)

    def _warn_unexpected_list_state(self, numid: int, ilevel: int) -> None:
        logger.warning(
            "Unexpected DOCX list state in _add_list_item: "
            f"pre_num_id={self.pre_num_id}, numid={numid}, "
            f"pre_ilevel={self.pre_ilevel}, ilevel={ilevel}, "
            f"stack_depth={len(self.list_block_stack)}. "
        )

    def _add_list_item(self, *, numid: int, ilevel: int, elements: list, is_numbered: bool = False, text: str = "", equations: list = None) -> list:
        equations = equations or []
        if not elements:
            return None
        content_text = self._build_text_with_equations_and_hyperlinks(elements, text, equations)
        content_text = self._normalize_text_block_content(content_text)
        if content_text == "":
            return None
        list_attribute = "ordered" if is_numbered else "unordered"
        self._add_list_item_to_current_state(numid, ilevel, list_attribute, content_text)
        return None
