# Copyright (c) Opendatalab. All rights reserved.
import re
from typing import Optional

from loguru import logger
from docx.document import Document as DocxDocument
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.text.paragraph import Paragraph
from lxml import etree

from docx_parse.utils.enum_class import BlockType


class DocxConverterStructureMixin:
    def _read_heading_scan_paragraph(self, element):
        try:
            paragraph = Paragraph(element, self.docx_obj)
            p_style_id, _ = self._get_label_and_level(paragraph)
            numid, ilevel = self._get_numId_and_ilvl(paragraph)
            text = self._get_paragraph_text(paragraph).strip()
        except Exception:
            return None
        return {
            "p_style_id": p_style_id,
            "numid": None if numid == 0 else numid,
            "ilevel": ilevel,
            "text": text,
        }

    @staticmethod
    def _heading_scan_item_from_context(ctx: dict):
        has_body_text = ctx["p_style_id"] not in ["Title", "Heading"] and ctx["text"]
        if not has_body_text:
            return None
        if ctx["numid"] is not None and ctx["ilevel"] is not None:
            return ("list", ctx["numid"], ctx["ilevel"])
        return ("content", None, None)

    def _heading_scan_item_from_element(self, element):
        tag_name = etree.QName(element).localname
        if tag_name == "tbl":
            return ("content", None, None)
        if tag_name != "p":
            return None
        ctx = self._read_heading_scan_paragraph(element)
        if ctx is None:
            return None
        return self._heading_scan_item_from_context(ctx)

    def _collect_heading_scan_items(self) -> tuple[list, dict[int, set]]:
        items = []
        numid_ilvels: dict[int, set] = {}
        for element in self.docx_obj.element.body:
            item = self._heading_scan_item_from_element(element)
            if item is None:
                continue
            items.append(item)
            if item[0] == "list":
                numid_ilvels.setdefault(item[1], set()).add(item[2])
        return items, numid_ilvels

    @staticmethod
    def _mark_heading_candidates_seen(seen_numids: dict[int, bool]) -> None:
        for nid in seen_numids:
            seen_numids[nid] = True

    @staticmethod
    def _record_heading_candidate(
        heading_numids: set,
        seen_numids: dict[int, bool],
        numid: int,
    ) -> None:
        if seen_numids.get(numid):
            heading_numids.add(numid)
        seen_numids[numid] = False

    @staticmethod
    def _heading_numids_from_scan_items(items: list, numid_ilvels: dict[int, set]) -> set:
        heading_numids = set()
        seen_numids: dict[int, bool] = {}
        for item_type, numid, _ilevel in items:
            if item_type == "list":
                DocxConverterStructureMixin._record_heading_candidate(heading_numids, seen_numids, numid)
            elif item_type == "content":
                DocxConverterStructureMixin._mark_heading_candidates_seen(seen_numids)
        return {nid for nid in heading_numids if len(numid_ilvels.get(nid, set())) > 1}

    def _detect_heading_list_numids(self) -> set:
        if not self._document_may_have_numbered_paragraphs():
            return set()
        items, numid_ilvels = self._collect_heading_scan_items()
        heading_numids = self._heading_numids_from_scan_items(items, numid_ilvels)
        if heading_numids:
            logger.debug(
                f"Detected heading-style list numIds (will convert to title blocks): {heading_numids}"
            )
        return heading_numids

    def _reset_list_counters_for_new_sequence(self, numid: int):
        """
        开始新的编号序列时重置计数器。

        Args:
            numid: 列表编号ID
        """
        keys_to_reset = [key for key in self.list_counters.keys() if key[0] == numid]
        for key in keys_to_reset:
            self.list_counters[key] = 0

    def _sdt_pr_indicates_toc(self, sdt_pr) -> bool:
        if sdt_pr is None:
            return False
        doc_part_gallery = sdt_pr.find(
            ".//w:docPartGallery", namespaces=self._BLIP_NAMESPACES
        )
        if self._doc_part_gallery_indicates_toc(doc_part_gallery):
            return True
        tag_elem = sdt_pr.find("w:tag", namespaces=self._BLIP_NAMESPACES)
        return self._sdt_tag_indicates_toc(tag_elem)

    def _doc_part_gallery_indicates_toc(self, doc_part_gallery) -> bool:
        if doc_part_gallery is None:
            return False
        val = doc_part_gallery.get(self.XML_KEY, "")
        return "Table of Contents" in val or "toc" in val.lower()

    def _sdt_tag_indicates_toc(self, tag_elem) -> bool:
        if tag_elem is None:
            return False
        val = tag_elem.get(self.XML_KEY, "").lower().replace(" ", "")
        return "toc" in val or "contents" in val or "tableofcontents" in val

    @staticmethod
    def _is_toc_style_name(style_name: str) -> bool:
        localized_toc_pattern = '^' + '\u76ee\u5f55' + r'\s*\d+$'
        return bool(
            re.match(r'^TOC\s*\d+$', style_name, re.IGNORECASE)
            or re.match(localized_toc_pattern, style_name)
        )

    def _sdt_content_has_toc_style(self, element: BaseOxmlElement) -> bool:
        sdt_content = element.find(
            "w:sdtContent", namespaces=self._BLIP_NAMESPACES
        )
        if sdt_content is None:
            return False
        paragraphs = sdt_content.findall(
            "w:p", namespaces=self._BLIP_NAMESPACES
        )
        for paragraph_element in paragraphs[:5]:
            if self._paragraph_element_has_toc_style(paragraph_element):
                return True
        return False

    def _paragraph_element_has_toc_style(self, paragraph_element) -> bool:
        try:
            paragraph = Paragraph(paragraph_element, self.docx_obj)
            style_name = paragraph.style.name if paragraph.style else ""
        except Exception:
            return False
        return bool(style_name and self._is_toc_style_name(style_name))

    def _is_toc_sdt(self, element: BaseOxmlElement) -> bool:
        """Return compact metadata for this paragraph or SDT."""
        sdt_pr = element.find("w:sdtPr", namespaces=self._BLIP_NAMESPACES)
        if self._sdt_pr_indicates_toc(sdt_pr):
            return True
        return self._sdt_content_has_toc_style(element)

    def _get_toc_item_level(self, paragraph: Paragraph) -> Optional[int]:
        """
        从段落样式中获取目录项的层级（0-based）。

        "TOC 1" -> 0
        "TOC 2" -> 1
        "目录 1" -> 0

        Args:
            paragraph: 段落对象

        Returns:
            Optional[int]: 层级（0-based），如果不是目录样式则返回 None
        """
        cache_key = paragraph._element
        if cache_key in self._paragraph_toc_level_cache:
            return self._paragraph_toc_level_cache[cache_key]

        if paragraph.style is None:
            self._paragraph_toc_level_cache[cache_key] = None
            return None
        style_name = paragraph.style.name
        if style_name:
            match = re.match(r'^(?:TOC|目录)\s*(\d+)$', style_name, re.IGNORECASE)
            if match:
                level = int(match.group(1))
                return level - 1  # 转换为 0-based
        return None

    def _is_flat_list_toc(
        self, items: list[tuple[int, str, list, list, Optional[str]]]
    ) -> bool:
        """
        检测目录是否为扁平列表（插图清单、列表清单等），
        这类目录的所有条目应在同一层级，不应嵌套。

        策略：检查是否超过 50% 的条目以"图"或"表"开头。
        """
        match_count = 0
        total_count = 0
        for _level, text, _elements, _equations, _anchor in items:
            stripped = text.strip()
            if not stripped:
                continue
            total_count += 1
            if re.match(r'^[图表][\d\s.]', stripped) or re.match(
                r'^(Figure|Table)\s+\d', stripped, re.IGNORECASE
            ):
                match_count += 1
        if total_count == 0:
            return False
        return match_count / total_count > 0.5

    def _correct_toc_level_by_text(self, toc_level: int, text: str) -> int:
        """
        通过文本中的编号深度修正目录项的层级。

        仅对 toc_level > 0 的条目进行修正，避免影响顶层章节标题。
        例如：
        - "1.1 LYSO..." (toc 3 → ilevel=2) → text depth 2 → 返回 1
        - "1.1.1 LYSO..." (toc 3 → ilevel=2) → text depth 3 → 返回 2
        - "本章小结" (toc 1 → ilevel=0) → 返回 0（不修正）
        """
        if toc_level == 0:
            return 0
        stripped = text.strip()
        match = re.match(r'^(\d+(?:\.\d+)*)', stripped)
        if match:
            parts = match.group(1).split('.')
            # "1.1" -> 2 parts -> level 1; "1.1.1" -> 3 parts -> level 2
            return len(parts) - 1
        return toc_level

    @staticmethod
    def _new_index_block(ilevel: int) -> dict:
        return {"type": BlockType.INDEX, "content": [], "ilevel": ilevel}

    @staticmethod
    def _new_index_item(content_text: str, anchor: Optional[str]) -> dict:
        index_item = {"type": BlockType.TEXT, "content": content_text}
        if anchor:
            index_item["anchor"] = anchor
        return index_item

    def _start_index_block(self, ilevel: int, content_text: str, anchor: Optional[str]) -> None:
        index_block = self._new_index_block(ilevel)
        self.cur_page.append(index_block)
        self.index_block_stack.append(index_block)
        index_block["content"].append(self._new_index_item(content_text, anchor))
        self.pre_index_ilevel = ilevel

    def _add_child_index_item(self, ilevel: int, content_text: str, anchor: Optional[str]) -> None:
        child_index_block = self._new_index_block(ilevel)
        self.index_block_stack[-1]["content"].append(child_index_block)
        self.index_block_stack.append(child_index_block)
        child_index_block["content"].append(self._new_index_item(content_text, anchor))
        self.pre_index_ilevel = ilevel

    def _add_decreased_index_item(self, ilevel: int, content_text: str, anchor: Optional[str]) -> None:
        while self.index_block_stack and self.index_block_stack[-1]["ilevel"] != ilevel:
            self.index_block_stack.pop()
        if self.index_block_stack:
            self.index_block_stack[-1]["content"].append(
                self._new_index_item(content_text, anchor)
            )
        self.pre_index_ilevel = ilevel

    def _add_same_index_item(self, content_text: str, anchor: Optional[str]) -> None:
        if self.index_block_stack:
            self.index_block_stack[-1]["content"].append(
                self._new_index_item(content_text, anchor)
            )

    def _add_index_item(self, *, ilevel: int, elements: list, text: str = "", equations: list = None, anchor: Optional[str] = None) -> None:
        equations = equations or []
        if not elements:
            return
        content_text = self._build_text_with_equations_and_hyperlinks(elements, text, equations)
        content_text = self._normalize_text_block_content(content_text)
        if content_text == "":
            return
        if self.pre_index_ilevel == -1:
            self._start_index_block(ilevel, content_text, anchor)
        elif self.pre_index_ilevel < ilevel:
            self._add_child_index_item(ilevel, content_text, anchor)
        elif ilevel < self.pre_index_ilevel:
            self._add_decreased_index_item(ilevel, content_text, anchor)
        else:
            self._add_same_index_item(content_text, anchor)

    def _paragraph_bookmark_names(self, paragraph_element: BaseOxmlElement) -> list[str]:
        bookmark_name_attr = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name"
        )
        names = []
        for bm in paragraph_element.findall(
            ".//w:bookmarkStart", namespaces=self._BLIP_NAMESPACES
        ):
            name = bm.get(bookmark_name_attr, "").strip()
            if not name:
                continue
            # skip Word navigation artifacts
            if name.startswith("_GoBack"):
                continue
            names.append(name)
        return names

    def _select_paragraph_bookmark(self, names: list[str]) -> Optional[str]:
        if not names:
            return None
        toc_names = [name for name in names if name.startswith("_Toc")]
        if toc_names:
            # Prefer anchors that are actually referenced by TOC hyperlinks.
            for name in toc_names:
                if name in self.toc_anchor_set:
                    return name
            return toc_names[0]
        return names[0]

    def _extract_paragraph_bookmark(self, paragraph_element: BaseOxmlElement) -> Optional[str]:
        """Extract a bookmark name from a paragraph, prioritizing TOC bookmarks."""
        names = self._paragraph_bookmark_names(paragraph_element)
        return self._select_paragraph_bookmark(names)

    def _extract_toc_target_anchor(self, paragraph_element: BaseOxmlElement) -> Optional[str]:
        """Extract internal bookmark target from a TOC paragraph hyperlink."""
        anchor_attr = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor"
        )
        anchors = []
        for hl in paragraph_element.findall(
            ".//w:hyperlink", namespaces=self._BLIP_NAMESPACES
        ):
            anchor = hl.get(anchor_attr, "").strip()
            if anchor:
                anchors.append(anchor)
        if not anchors:
            return None
        for anchor in anchors:
            if anchor.startswith("_Toc"):
                return anchor
        return anchors[0]

    def _handle_plain_toc_paragraph_as_index(
        self,
        *,
        paragraph: Paragraph,
        paragraph_element: BaseOxmlElement,
        paragraph_elements: list,
        text: str,
        equations: list,
    ) -> bool:
        """将未包裹在 SDT 中的普通目录段落转换为 INDEX 项。"""
        toc_level = self._get_toc_item_level(paragraph)
        if toc_level is None:
            return False
        if not text:
            return True

        target_anchor = self._extract_toc_target_anchor(paragraph_element)
        # 只有已经进入目录序列后才允许无锚点条目，避免误收复用 TOC 样式的封面文本。
        if not target_anchor and self.pre_index_ilevel == -1:
            return False
        if target_anchor and target_anchor.startswith("_Toc"):
            self.toc_anchor_set.add(target_anchor)

        if self.plain_toc_base_level is None:
            self.plain_toc_base_level = toc_level
        normalized_level = max(0, toc_level - self.plain_toc_base_level)
        corrected_level = self._correct_toc_level_by_text(normalized_level, text)
        self._add_index_item(
            ilevel=corrected_level,
            elements=paragraph_elements,
            text=text,
            equations=equations,
            anchor=target_anchor,
        )
        return True

    def _toc_item_from_paragraph_element(self, paragraph_element) -> Optional[tuple[int, str, list, list, Optional[str]]]:
        try:
            paragraph = Paragraph(paragraph_element, self.docx_obj)
            paragraph_elements = self._get_paragraph_elements(paragraph)
            text, equations = self._handle_equations_in_text(
                element=paragraph_element, text=paragraph.text
            )
            target_anchor = self._extract_toc_target_anchor(paragraph_element)
            if target_anchor and target_anchor.startswith("_Toc"):
                self.toc_anchor_set.add(target_anchor)
            if text is None or not text.strip():
                return None
            toc_level = self._get_toc_item_level(paragraph)
            return (toc_level or 0, text.strip(), paragraph_elements, equations, target_anchor)
        except Exception as e:
            logger.debug(f"Error collecting TOC paragraph: {e}")
            return None

    def _collect_sdt_toc_items(self, sdt_content: BaseOxmlElement) -> list[tuple[int, str, list, list, Optional[str]]]:
        paragraphs = sdt_content.findall(".//w:p", namespaces=self._BLIP_NAMESPACES)
        toc_items = []
        for paragraph_element in paragraphs:
            item = self._toc_item_from_paragraph_element(paragraph_element)
            if item is not None:
                toc_items.append(item)
        return toc_items

    def _write_sdt_toc_items(self, toc_items: list[tuple[int, str, list, list, Optional[str]]]) -> None:
        is_flat = self._is_flat_list_toc(toc_items)
        self._reset_index_state()
        for toc_level, text, elements, equations, target_anchor in toc_items:
            corrected_level = 0 if is_flat else self._correct_toc_level_by_text(toc_level, text)
            self._add_index_item(
                ilevel=corrected_level,
                elements=elements,
                text=text,
                equations=equations,
                anchor=target_anchor,
            )
        self._reset_index_state()

    def _handle_sdt_as_index(self, sdt_content: BaseOxmlElement) -> None:
        self._write_sdt_toc_items(self._collect_sdt_toc_items(sdt_content))

    def _get_heading_and_level(self, style_label: str) -> tuple[str, Optional[int]]:
        """
        从样式标签获取标题和层级。

        Args:
            style_label: 样式标签

        Returns:
            tuple[str, Optional[int]]: (标签字符串, 层级) 元组
        """
        parts = self._split_text_and_number(style_label)

        if len(parts) == 2:
            parts.sort()
            label_str: str = ""
            label_level: Optional[int] = 0
            if parts[0].strip().lower() == "heading":
                label_str = "Heading"
                label_level = self._str_to_int(parts[1], None)
            if parts[1].strip().lower() == "heading":
                label_str = "Heading"
                label_level = self._str_to_int(parts[0], None)
            return label_str, label_level

        return style_label, None

    def _split_text_and_number(self, input_string: str) -> list[str]:
        """
        分割字符串中的文本和数字部分。

        Args:
            input_string: 输入字符串

        Returns:
            list[str]: 分割后的部分列表
        """
        match = re.match(r"(\D+)(\d+)$|^(\d+)(\D+)", input_string)
        if match:
            parts = list(filter(None, match.groups()))
            return parts
        else:
            return [input_string]

    def _str_to_int(
        self, s: Optional[str], default: Optional[int] = 0
    ) -> Optional[int]:
        """
        将字符串转换为整数。

        Args:
            s: 要转换的字符串
            default: 默认值，转换失败时返回

        Returns:
            Optional[int]: 转换后的整数，转换失败时返回默认值
        """
        if s is None:
            return None
        try:
            return int(s)
        except ValueError:
            return default

    def _process_header_footer_paragraph(self, paragraph: Paragraph) -> str:
        """
        处理页眉/页脚中的单个段落，支持行内公式和超链接。

        Args:
            paragraph: 段落对象

        Returns:
            str: 处理后的文本内容（包含公式标记和超链接格式）
        """
        paragraph_elements = self._get_paragraph_elements(paragraph)
        paragraph_text = self._get_paragraph_text(paragraph)
        text, equations = self._handle_equations_in_text(
            element=paragraph._element, text=paragraph_text
        )

        if text is None:
            return ""

        text = text.strip()
        if not text:
            return ""

        # 构建包含公式和超链接的文本
        content_text = self._build_text_with_equations_and_hyperlinks(
            paragraph_elements, text, equations
        )

        return content_text

    def _section_headers(self, section, is_odd_even_different: bool) -> list:
        headers = [section.header]
        if is_odd_even_different:
            headers.append(section.even_page_header)
        if section.different_first_page_header_footer:
            headers.append(section.first_page_header)
        return headers

    def _section_footers(self, section, is_odd_even_different: bool) -> list:
        footers = [section.footer]
        if is_odd_even_different:
            footers.append(section.even_page_footer)
        if section.different_first_page_header_footer:
            footers.append(section.first_page_footer)
        return footers

    def _header_footer_text(self, header_or_footer) -> str:
        processed_parts = []
        for paragraph in header_or_footer.paragraphs:
            content = self._process_header_footer_paragraph(paragraph)
            if content:
                processed_parts.append(content)
        return " ".join(processed_parts)

    def _append_header_footer_block(self, sec_idx: int, block_type: str, text: str, added: set) -> None:
        if text == "" or text.isdigit() or text in added:
            return
        added.add(text)
        try:
            self.pages[sec_idx].append({"type": block_type, "content": text})
        except IndexError:
            logger.error(f"Section index out of range when adding {block_type}.")

    def _add_header_footer(self, docx_obj: DocxDocument) -> None:
        is_odd_even_different = docx_obj.settings.odd_and_even_pages_header_footer
        for sec_idx, section in enumerate(docx_obj.sections):
            added_headers = set()
            added_footers = set()
            for header in self._section_headers(section, is_odd_even_different):
                self._append_header_footer_block(
                    sec_idx, BlockType.HEADER, self._header_footer_text(header), added_headers
                )
            for footer in self._section_footers(section, is_odd_even_different):
                self._append_header_footer_block(
                    sec_idx, BlockType.FOOTER, self._header_footer_text(footer), added_footers
                )

    def _is_caption(self, element: BaseOxmlElement) -> bool:
        """
        根据 insertText 中是否有 SEQ 字段来判断是否为 caption

        Args:
            element: 段落元素对象

        Returns:
            bool: 如果是标题返回 True，否则返回 False
        """
        instr_texts = element.findall(
            ".//w:instrText", namespaces=self._BLIP_NAMESPACES
        )

        for instr in instr_texts:
            if instr.text and "SEQ" in instr.text:
                return True

        return False
    
    @staticmethod
    def _sorted_textbox_paragraphs(container_paragraphs: dict) -> list:
        all_paragraphs = []
        for paragraphs in container_paragraphs.values():
            all_paragraphs.extend(
                sorted(
                    paragraphs,
                    key=lambda x: (x[1] is None, x[1] if x[1] is not None else float("inf")),
                )
            )
        return all_paragraphs

    def _handle_textbox_content(self, textbox_elements: list):
        all_paragraphs = self._sorted_textbox_paragraphs(
            self._collect_textbox_paragraphs(textbox_elements)
        )
        processed_paragraphs = set()
        for paragraph_element, position in all_paragraphs:
            paragraph = Paragraph(paragraph_element, self.docx_obj)
            text_content = self._get_paragraph_text(paragraph)
            paragraph_id = (text_content, position)
            if paragraph_id in processed_paragraphs:
                logger.debug(
                    f"Skipping duplicate paragraph: content='{text_content[:50]}...', position={position}"
                )
                continue
            processed_paragraphs.add(paragraph_id)
            self._handle_text_elements(paragraph_element)
        return

    def _textbox_container_id_for_paragraph(self, element) -> int | None:
        for ancestor in element.iterancestors():
            if any(ns in ancestor.tag for ns in ["textbox", "shape", "txbx"]):
                return id(ancestor)
        return None

    def _append_textbox_paragraph(self, container_paragraphs: dict, container_id, paragraph) -> None:
        container_paragraphs.setdefault(container_id, [])
        container_paragraphs[container_id].append(
            (paragraph, self._get_paragraph_position(paragraph))
        )

    def _append_nested_textbox_paragraphs(self, container_paragraphs: dict, element, processed_paragraphs: list) -> None:
        paragraphs = element.findall(".//w:p", namespaces=element.nsmap)
        container_id = id(element)
        for paragraph in paragraphs:
            paragraph_id = id(paragraph)
            if paragraph_id in processed_paragraphs:
                continue
            processed_paragraphs.append(paragraph_id)
            self._append_textbox_paragraph(container_paragraphs, container_id, paragraph)

    def _collect_textbox_paragraphs(self, textbox_elements):
        processed_paragraphs = []
        container_paragraphs = {}
        for element in textbox_elements:
            element_id = id(element)
            if element_id in processed_paragraphs:
                continue
            processed_paragraphs.append(element_id)
            tag_name = etree.QName(element).localname
            if tag_name == "p":
                container_id = self._textbox_container_id_for_paragraph(element)
                self._append_textbox_paragraph(container_paragraphs, container_id, element)
            else:
                self._append_nested_textbox_paragraphs(
                    container_paragraphs, element, processed_paragraphs
                )
        return container_paragraphs

    @staticmethod
    def _sibling_paragraph_index(paragraph_element):
        if not hasattr(paragraph_element, "getparent") or paragraph_element.getparent() is None:
            return None
        parent = paragraph_element.getparent()
        paragraphs = [p for p in parent.getchildren() if etree.QName(p).localname == "p"]
        try:
            return paragraphs.index(paragraph_element)
        except ValueError:
            return None

    @staticmethod
    def _numeric_position_attr(elem):
        for attr_name in ["y", "top", "positionY", "y-position", "position"]:
            value = elem.get(attr_name)
            if not value:
                continue
            try:
                clean_value = re.sub(r"[^0-9.]", "", value)
                if clean_value:
                    return float(clean_value)
            except (ValueError, TypeError):
                pass
        return None

    @staticmethod
    def _transform_position(elem):
        transform = elem.get("transform")
        if not transform:
            return None
        match = re.search(r"translate\([^,]+,\s*([0-9.]+)", transform)
        if not match:
            return None
        try:
            return float(match.group(1))
        except ValueError:
            return None

    @staticmethod
    def _vml_style_position(paragraph_element):
        for ns_uri in paragraph_element.nsmap.values():
            if "vml" not in ns_uri:
                continue
            style = paragraph_element.get("style")
            if not style:
                continue
            match = re.search(r"top:([0-9.]+)pt", style)
            if match:
                try:
                    return float(match.group(1))
                except ValueError:
                    pass
        return None

    def _get_paragraph_position(self, paragraph_element):
        sibling_index = self._sibling_paragraph_index(paragraph_element)
        if sibling_index is not None:
            return sibling_index
        for elem in (*[paragraph_element], *paragraph_element.iterancestors()):
            position = self._numeric_position_attr(elem)
            if position is not None:
                return position
            position = self._transform_position(elem)
            if position is not None:
                return position
            for attr_name in ["distT", "distB", "anchor", "relativeFrom"]:
                if elem.get(attr_name) is not None:
                    return elem.sourceline
        vml_position = self._vml_style_position(paragraph_element)
        if vml_position is not None:
            return vml_position
        return paragraph_element.sourceline if hasattr(paragraph_element, "sourceline") else None

