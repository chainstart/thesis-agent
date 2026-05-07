from __future__ import annotations

from xml.etree import ElementTree as ET


COMMON_NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
}

for _prefix, _namespace in COMMON_NAMESPACES.items():
    ET.register_namespace(_prefix, _namespace)


def serialize_xml(root: ET.Element) -> bytes:
    body = ET.tostring(root, encoding="utf-8", xml_declaration=False)
    body = _ensure_ignorable_namespace_declarations(body)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


def serialize_package_xml(root: ET.Element, default_namespace: str) -> bytes:
    ET.register_namespace("", default_namespace)
    body = ET.tostring(root, encoding="utf-8", xml_declaration=False)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


def _ensure_ignorable_namespace_declarations(xml: bytes) -> bytes:
    start = xml.find(b"<w:document ")
    if start < 0:
        return xml
    end = xml.find(b">", start)
    if end < 0:
        return xml
    tag = xml[start:end]
    insertions: list[bytes] = []
    for prefix in _ignorable_prefixes(tag):
        namespace = COMMON_NAMESPACES.get(prefix)
        if not namespace:
            continue
        needle = f"xmlns:{prefix}=".encode()
        if needle not in tag:
            insertions.append(f' xmlns:{prefix}="{namespace}"'.encode())
    if not insertions:
        return xml
    return xml[:end] + b"".join(insertions) + xml[end:]


def _ignorable_prefixes(root_tag: bytes) -> list[str]:
    prefixes: list[str] = []
    marker = b":Ignorable=\""
    start = root_tag.find(marker)
    if start < 0:
        return prefixes
    start += len(marker)
    end = root_tag.find(b"\"", start)
    if end < 0:
        return prefixes
    for token in root_tag[start:end].decode("utf-8", errors="ignore").split():
        if token and token not in prefixes:
            prefixes.append(token)
    return prefixes
