from __future__ import annotations

import base64
import re
from html import escape
from typing import Any

from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree
from loguru import logger

from docx_parse.backend.utils.office_image import (
    is_vector_image_part,
    serialize_office_image,
)
from docx_parse.model.docx.tools.math.omml import oMath2Latex
from docx_parse.utils.docx_formatting import Script

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def render_table_html(converter: Any, table_element) -> str:
    renderer = DocxTableXmlRenderer(converter)
    return renderer.render_table(table_element)


class DocxTableXmlRenderer:
    def __init__(self, converter: Any):
        self.converter = converter

    def render_table(self, table_element) -> str:
        rows = [child for child in table_element if self._local_name(child) == "tr"]
        row_infos = self._build_row_infos(rows)
        rendered_rows = []
        for row in row_infos:
            cells = []
            for cell in row:
                if cell["skip"]:
                    continue
                attrs = []
                if cell["grid_span"] > 1:
                    attrs.append(f'colspan="{cell["grid_span"]}"')
                if cell["rowspan"] > 1:
                    attrs.append(f'rowspan="{cell["rowspan"]}"')
                attr_text = " " + " ".join(attrs) if attrs else ""
                cells.append(
                    f"<td{attr_text}>{self.render_cell(cell['element'])}</td>"
                )
            rendered_rows.append(f"<tr>{''.join(cells)}</tr>")
        return self._merge_adjacent_inline_tags(f"<table>{''.join(rendered_rows)}</table>")

    def render_cell(self, cell_element) -> str:
        parts = []
        open_list_type: str | None = None

        def close_list() -> None:
            nonlocal open_list_type
            if open_list_type is not None:
                parts.append(f"</{open_list_type}>")
                open_list_type = None

        for child in cell_element:
            child_name = self._local_name(child)
            if child_name == "p":
                paragraph = Paragraph(child, self.converter.docx_obj)
                numid, ilevel = self.converter._get_numId_and_ilvl(paragraph)
                if numid is not None and ilevel is not None:
                    list_type = (
                        "ol"
                        if self.converter._is_numbered_list(numid, ilevel)
                        else "ul"
                    )
                    if open_list_type != list_type:
                        close_list()
                        parts.append(f"<{list_type}>")
                        open_list_type = list_type
                    parts.append(f"<li>{self.render_paragraph_inline(child)}</li>")
                    continue

                close_list()
                paragraph_html = self.render_paragraph_inline(child)
                if paragraph_html:
                    parts.append(f"<p>{paragraph_html}</p>")
            elif child_name == "tbl":
                close_list()
                parts.append(self.render_table(child))
        close_list()
        return "".join(parts)

    def render_paragraph_inline(self, paragraph_element) -> str:
        paragraph = Paragraph(paragraph_element, self.converter.docx_obj)
        return self._render_inline_children(paragraph_element, paragraph)

    def _render_inline_children(self, parent, paragraph: Paragraph) -> str:
        parts = []
        for child in parent:
            child_name = self._local_name(child)
            if child_name == "r":
                parts.append(self._render_run(child, paragraph))
            elif child_name == "hyperlink":
                parts.append(self._render_hyperlink(child, paragraph))
            elif child_name == "oMath":
                parts.append(self._render_equation(child))
            elif child_name == "oMathPara":
                for math_child in child:
                    if self._local_name(math_child) == "oMath":
                        parts.append(self._render_equation(math_child))
            elif (
                child_name in self.converter._PARAGRAPH_TRANSPARENT_INLINE_CONTAINERS
                or child_name in {"sdt", "ins", "moveTo", "fldSimple"}
            ):
                if child_name == "sdt":
                    sdt_content = child.find(f"{{{W_NS}}}sdtContent")
                    if sdt_content is not None:
                        parts.append(self._render_inline_children(sdt_content, paragraph))
                else:
                    parts.append(self._render_inline_children(child, paragraph))
        return "".join(parts)

    def _render_run(self, run_element, paragraph: Paragraph) -> str:
        run_parts = []
        for child in run_element:
            child_name = self._local_name(child)
            if child_name == "t":
                run_parts.append(escape(child.text or "", quote=False))
            elif child_name == "tab":
                run_parts.append("\t")
            elif child_name == "br":
                run_parts.append("<br/>")
            elif child_name in {"drawing", "pict"}:
                run_parts.append(self._render_images_in_element(child))
            elif child_name == "oMath":
                run_parts.append(self._render_equation(child))
        content = "".join(run_parts)
        if not content:
            return ""
        try:
            formatting = self.converter._get_format_from_run(Run(run_element, paragraph))
        except Exception as exc:
            logger.debug(f"Could not resolve table run formatting: {exc}")
            return content
        return self._apply_run_formatting(content, formatting)

    def _render_hyperlink(self, hyperlink_element, paragraph: Paragraph) -> str:
        content = self._render_inline_children(hyperlink_element, paragraph)
        if not content:
            return ""
        rel_id = hyperlink_element.get(f"{{{R_NS}}}id")
        anchor = hyperlink_element.get(f"{{{W_NS}}}anchor")
        href = ""
        if rel_id:
            rel = self.converter.docx_obj.part.rels.get(rel_id)
            href = str(
                getattr(rel, "target_ref", "")
                or getattr(rel, "target_part", "")
                or ""
            )
        elif anchor:
            href = f"#{anchor}"
        if not href:
            return content
        return f'<a href="{escape(href, quote=True)}">{content}</a>'

    def _render_equation(self, math_element) -> str:
        try:
            latex = str(oMath2Latex(math_element)).strip()
        except Exception as exc:
            logger.debug(f"Failed to convert table OMML equation to LaTeX: {exc}")
            return ""
        if not latex:
            return ""
        return f"<eq>{escape(latex, quote=False)}</eq>"

    def _render_images_in_element(self, element) -> str:
        html_parts = []
        seen_rel_ids = set()
        for image in self.converter.picture_xpath_expr(element):
            rel_id = image.get(f"{{{R_NS}}}embed") or image.get(f"{{{R_NS}}}id")
            if not rel_id or rel_id in seen_rel_ids:
                continue
            seen_rel_ids.add(rel_id)
            rel = self.converter.docx_obj.part.rels.get(rel_id)
            image_part = getattr(rel, "target_part", None) if rel is not None else None
            if image_part is None:
                continue
            data_uri = self._image_part_to_data_uri(image_part)
            if data_uri:
                html_parts.append(f'<img src="{data_uri}"/>')
        return "".join(html_parts)

    def _image_part_to_data_uri(self, image_part) -> str | None:
        part_name = getattr(image_part, "partname", None)
        content_type = getattr(image_part, "content_type", None)
        blob = getattr(image_part, "blob", None)
        if not blob:
            return None
        normalized_content_type = (content_type or "").split(";", 1)[0].strip().lower()
        if normalized_content_type.startswith("image/") and not is_vector_image_part(part_name, content_type):
            return (
                f"data:{normalized_content_type};base64,"
                f"{base64.b64encode(blob).decode('ascii')}"
            )
        return serialize_office_image(blob, part_name=part_name, content_type=content_type)

    def _apply_run_formatting(self, content: str, formatting) -> str:
        if formatting.script == Script.SUB:
            content = f"<sub>{content}</sub>"
        elif formatting.script == Script.SUPER:
            content = f"<sup>{content}</sup>"
        if formatting.italic:
            content = f"<em>{content}</em>"
        if formatting.bold:
            content = f"<strong>{content}</strong>"
        if formatting.strikethrough:
            content = f"<s>{content}</s>"
        return content

    @staticmethod
    def _merge_adjacent_inline_tags(html: str) -> str:
        """Match the legacy table output by folding adjacent equal style runs."""
        previous = None
        while previous != html:
            previous = html
            html = re.sub(r"</(strong|em|s)><\1>", "", html)
        return html

    def _build_row_infos(self, rows) -> list[list[dict]]:
        row_infos = []
        for row_index, row in enumerate(rows):
            col_index = 0
            cells = []
            for cell_element in row:
                if self._local_name(cell_element) != "tc":
                    continue
                grid_span = self._get_grid_span(cell_element)
                vmerge = self._get_vmerge(cell_element)
                cells.append(
                    {
                        "element": cell_element,
                        "row": row_index,
                        "col": col_index,
                        "grid_span": grid_span,
                        "vmerge": vmerge,
                        "rowspan": 1,
                        "skip": vmerge == "continue",
                    }
                )
                col_index += grid_span
            row_infos.append(cells)

        by_position = {
            (cell["row"], cell["col"]): cell for row in row_infos for cell in row
        }
        for row in row_infos:
            for cell in row:
                if cell["vmerge"] != "restart":
                    continue
                span = 1
                next_row = cell["row"] + 1
                while True:
                    next_cell = by_position.get((next_row, cell["col"]))
                    if next_cell is None or next_cell["vmerge"] != "continue":
                        break
                    span += 1
                    next_row += 1
                cell["rowspan"] = span
        return row_infos

    def _get_grid_span(self, cell_element) -> int:
        grid_span = cell_element.find(f"./{{{W_NS}}}tcPr/{{{W_NS}}}gridSpan")
        if grid_span is None:
            return 1
        value = grid_span.get(f"{{{W_NS}}}val")
        try:
            return max(1, int(value))
        except (TypeError, ValueError):
            return 1

    def _get_vmerge(self, cell_element) -> str | None:
        vmerge = cell_element.find(f"./{{{W_NS}}}tcPr/{{{W_NS}}}vMerge")
        if vmerge is None:
            return None
        value = vmerge.get(f"{{{W_NS}}}val")
        if value == "restart":
            return "restart"
        return "continue"

    @staticmethod
    def _local_name(element) -> str:
        return etree.QName(element).localname
