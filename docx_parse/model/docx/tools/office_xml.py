# Copyright (c) Opendatalab. All rights reserved.
from lxml import etree


def read_str(xml_str):
    """Parse an XML string into an lxml element.

    This helper is kept for compatibility with older internal imports. The DOCX
    converter no longer depends on the previous private XML model.
    """
    if isinstance(xml_str, str):
        xml_str = xml_str.encode("utf-8")
    return etree.fromstring(xml_str)

