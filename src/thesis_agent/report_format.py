from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_inspect import NS, W_NS, _paragraph_text
from .ooxml import serialize_package_xml, serialize_xml


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


@dataclass(frozen=True)
class ReportTemplateFormatReport:
    input: Path
    output: Path
    template: Path
    sections_rebased: int = 0
    relationship_parts_imported: int = 0
    template_tables_restyled: int = 0
    boilerplate_paragraphs_restyled: int = 0
    task_note_page_breaks_inserted: int = 0


def fix_report_template_format(input_path: Path, output_path: Path, template_path: Path) -> ReportTemplateFormatReport:
    """Make form/report documents inherit layout from their official template.

    This is intentionally narrower than ``fix_docx_format``.  Report materials
    such as task books and proposal forms are fixed-format school forms, so
    applying thesis-body rules like "chapter headings start new pages" can move
    them away from the template.  Here the template owns sections, boilerplate
    paragraph styles, and the official form tables; the target owns the filled
    text and student-added body content.
    """

    if input_path.suffix.lower() != ".docx" or template_path.suffix.lower() != ".docx":
        raise ValueError("report template formatting requires .docx input and template")
    if not zipfile.is_zipfile(input_path):
        raise ValueError(f"Not a valid docx file: {input_path}")
    if not zipfile.is_zipfile(template_path):
        raise ValueError(f"Not a valid template docx file: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(input_path) as target_zip, zipfile.ZipFile(template_path) as template_zip:
        target_root = ET.fromstring(target_zip.read("word/document.xml"))
        template_root = ET.fromstring(template_zip.read("word/document.xml"))
        target_rels = _read_relationships(target_zip)
        content_types = _read_content_types(target_zip)
        existing_names = set(target_zip.namelist())

        section_count, imported_parts, extra_parts = _rebase_sections_from_template(
            target_root,
            template_root,
            template_zip,
            target_rels,
            content_types,
            existing_names,
        )
        boilerplate_count = _restyle_boilerplate_paragraphs(target_root, template_root)
        table_count = _restyle_template_tables(target_root, template_root)
        task_breaks = _ensure_task_note_page(target_root)

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in target_zip.infolist():
                if item.filename == "word/document.xml":
                    data = serialize_xml(target_root)
                elif item.filename == "word/_rels/document.xml.rels":
                    data = serialize_package_xml(target_rels, REL_NS)
                elif item.filename == "[Content_Types].xml":
                    data = serialize_package_xml(content_types, CT_NS)
                else:
                    data = target_zip.read(item.filename)
                output_zip.writestr(item, data)
            if "word/_rels/document.xml.rels" not in existing_names:
                output_zip.writestr("word/_rels/document.xml.rels", serialize_package_xml(target_rels, REL_NS))
            if "[Content_Types].xml" not in existing_names:
                output_zip.writestr("[Content_Types].xml", serialize_package_xml(content_types, CT_NS))
            for name, data in extra_parts.items():
                if name not in existing_names:
                    output_zip.writestr(name, data)

    return ReportTemplateFormatReport(
        input=input_path,
        output=output_path,
        template=template_path,
        sections_rebased=section_count,
        relationship_parts_imported=imported_parts,
        template_tables_restyled=table_count,
        boilerplate_paragraphs_restyled=boilerplate_count,
        task_note_page_breaks_inserted=task_breaks,
    )


def _rebase_sections_from_template(
    target_root: ET.Element,
    template_root: ET.Element,
    template_zip: zipfile.ZipFile,
    target_rels: ET.Element,
    content_types: ET.Element,
    existing_names: set[str],
) -> tuple[int, int, dict[str, bytes]]:
    template_sections = list(template_root.findall(".//w:sectPr", NS))
    if not template_sections:
        return 0, 0, {}
    target_locations = _section_locations(target_root)
    if not target_locations:
        return 0, 0, {}

    rel_importer = _RelationshipImporter(template_zip, target_rels, content_types, existing_names)
    count = 0
    for idx, (parent, child_idx, _section) in enumerate(target_locations):
        source = template_sections[min(idx, len(template_sections) - 1)]
        replacement = rel_importer.copy_section(source)
        parent[child_idx] = replacement
        count += 1
    return count, rel_importer.imported_parts, rel_importer.extra_parts


def _section_locations(root: ET.Element) -> list[tuple[ET.Element, int, ET.Element]]:
    locations: list[tuple[ET.Element, int, ET.Element]] = []

    def visit(parent: ET.Element) -> None:
        for idx, child in enumerate(list(parent)):
            if child.tag == _w("sectPr"):
                locations.append((parent, idx, child))
            visit(child)

    visit(root)
    return locations


class _RelationshipImporter:
    def __init__(
        self,
        template_zip: zipfile.ZipFile,
        target_rels: ET.Element,
        content_types: ET.Element,
        existing_names: set[str],
    ) -> None:
        self.template_zip = template_zip
        self.target_rels = target_rels
        self.content_types = content_types
        self.existing_names = set(existing_names)
        self.extra_parts: dict[str, bytes] = {}
        self.imported_parts = 0
        self.template_rels = _read_relationships(template_zip)
        self.template_rels_by_id = {rel.attrib.get("Id"): rel for rel in self.template_rels}
        self._cache: dict[str, str] = {}

    def copy_section(self, section: ET.Element) -> ET.Element:
        copied = copy.deepcopy(section)
        for ref in copied.findall("w:headerReference", NS) + copied.findall("w:footerReference", NS):
            old_rid = ref.attrib.get(_r("id"))
            if not old_rid:
                continue
            new_rid = self._import_relationship(old_rid)
            if new_rid:
                ref.set(_r("id"), new_rid)
        return copied

    def _import_relationship(self, template_rid: str) -> str | None:
        if template_rid in self._cache:
            return self._cache[template_rid]
        rel = self.template_rels_by_id.get(template_rid)
        if rel is None:
            return None
        rel_type = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if not (rel_type.endswith("/header") or rel_type.endswith("/footer")):
            return None
        source_name = _word_part_name(target)
        if source_name not in self.template_zip.namelist():
            return None
        kind = "header" if rel_type.endswith("/header") else "footer"
        new_target = self._next_part_target(kind)
        new_part_name = "word/" + new_target
        self.extra_parts[new_part_name] = self.template_zip.read(source_name)
        self.existing_names.add(new_part_name)

        new_rid = self._next_rid()
        relationship = ET.Element(_rel("Relationship"))
        relationship.set("Id", new_rid)
        relationship.set("Type", rel_type)
        relationship.set("Target", new_target)
        self.target_rels.append(relationship)
        _ensure_content_type(self.content_types, "/" + new_part_name, _content_type_for_rel(rel_type))
        self._cache[template_rid] = new_rid
        self.imported_parts += 1
        return new_rid

    def _next_rid(self) -> str:
        used = {rel.attrib.get("Id") for rel in self.target_rels}
        idx = 1
        while f"rId{idx}" in used:
            idx += 1
        return f"rId{idx}"

    def _next_part_target(self, kind: str) -> str:
        idx = 1
        while f"word/{kind}Template{idx}.xml" in self.existing_names:
            idx += 1
        return f"{kind}Template{idx}.xml"


def _restyle_boilerplate_paragraphs(target_root: ET.Element, template_root: ET.Element) -> int:
    samples = _template_paragraph_samples(template_root)
    table_paragraphs = {id(p) for table in target_root.findall(".//w:tbl", NS) for p in table.findall(".//w:p", NS)}
    count = 0
    for paragraph in target_root.findall(".//w:p", NS):
        if id(paragraph) in table_paragraphs:
            continue
        key = _paragraph_match_key(_paragraph_text(paragraph))
        if not key:
            continue
        sample = samples.get(key)
        if sample is None:
            sample = samples.get(_field_prefix_key(key))
        if sample is None:
            continue
        _copy_paragraph_format(paragraph, sample)
        count += 1
    return count


def _template_paragraph_samples(template_root: ET.Element) -> dict[str, ET.Element]:
    samples: dict[str, ET.Element] = {}
    table_paragraphs = {id(p) for table in template_root.findall(".//w:tbl", NS) for p in table.findall(".//w:p", NS)}
    for paragraph in template_root.findall(".//w:p", NS):
        if id(paragraph) in table_paragraphs:
            continue
        text = _paragraph_text(paragraph)
        key = _paragraph_match_key(text)
        if not key:
            continue
        samples.setdefault(key, paragraph)
        prefix = _field_prefix_key(key)
        if prefix:
            samples.setdefault(prefix, paragraph)
    return samples


def _paragraph_match_key(text: str) -> str:
    compact = re.sub(r"\s+", "", text)
    compact = compact.replace("_", "")
    return compact


def _field_prefix_key(key: str) -> str:
    for prefix in (
        "课题名称",
        "课题",
        "专业",
        "班级",
        "姓名",
        "学号",
        "学院",
        "指导教师",
        "定稿日期",
        "年月日",
    ):
        if key.startswith(prefix):
            return prefix
    return ""


def _copy_paragraph_format(target: ET.Element, sample: ET.Element) -> None:
    existing_section = target.find("w:pPr/w:sectPr", NS)
    sample_ppr = sample.find("w:pPr", NS)
    copied_ppr = copy.deepcopy(sample_ppr) if sample_ppr is not None else None
    if copied_ppr is not None:
        for section in copied_ppr.findall("w:sectPr", NS):
            copied_ppr.remove(section)
        if existing_section is not None:
            copied_ppr.append(copy.deepcopy(existing_section))
    elif existing_section is not None:
        copied_ppr = ET.Element(_w("pPr"))
        copied_ppr.append(copy.deepcopy(existing_section))
    _replace_optional_child(target, "pPr", copied_ppr)
    sample_rpr = _first_run_properties(sample)
    if sample_rpr is not None:
        for run in target.findall("w:r", NS):
            _replace_optional_child(run, "rPr", sample_rpr)
    sample_breaks = len(sample.findall(".//w:br[@w:type='page']", NS))
    if sample_breaks:
        _ensure_page_break_after_text(target)


def _restyle_template_tables(target_root: ET.Element, template_root: ET.Element) -> int:
    target_tables = target_root.findall(".//w:tbl", NS)
    template_tables = template_root.findall(".//w:tbl", NS)
    if not target_tables or not template_tables:
        return 0
    count = 0
    for target_table, template_table in zip(target_tables[-len(template_tables) :], template_tables):
        _copy_table_format(target_table, template_table)
        count += 1
    return count


def _copy_table_format(target: ET.Element, sample: ET.Element) -> None:
    _replace_optional_child(target, "tblPr", sample.find("w:tblPr", NS), insert_at=0)
    sample_grid = sample.find("w:tblGrid", NS)
    target_grid = target.find("w:tblGrid", NS)
    if sample_grid is not None:
        if target_grid is not None:
            target.remove(target_grid)
        insert_at = 1 if target.find("w:tblPr", NS) is not None else 0
        target.insert(insert_at, copy.deepcopy(sample_grid))

    sample_rows = sample.findall("w:tr", NS)
    for row_idx, target_row in enumerate(target.findall("w:tr", NS)):
        sample_row = sample_rows[min(row_idx, len(sample_rows) - 1)] if sample_rows else None
        if sample_row is None:
            continue
        _replace_optional_child(target_row, "trPr", sample_row.find("w:trPr", NS), insert_at=0)
        sample_cells = sample_row.findall("w:tc", NS)
        for cell_idx, target_cell in enumerate(target_row.findall("w:tc", NS)):
            sample_cell = sample_cells[min(cell_idx, len(sample_cells) - 1)] if sample_cells else None
            if sample_cell is None:
                continue
            _copy_cell_format(target_cell, sample_cell)


def _copy_cell_format(target: ET.Element, sample: ET.Element) -> None:
    _replace_optional_child(target, "tcPr", sample.find("w:tcPr", NS), insert_at=0)
    sample_paragraphs = sample.findall("w:p", NS)
    if not sample_paragraphs:
        return
    for idx, target_paragraph in enumerate(target.findall("w:p", NS)):
        sample_paragraph = sample_paragraphs[min(idx, len(sample_paragraphs) - 1)]
        _copy_paragraph_format(target_paragraph, sample_paragraph)


def _ensure_task_note_page(root: ET.Element) -> int:
    body = root.find("w:body", NS)
    if body is None:
        return 0
    changes = 0
    for child in list(body):
        if child.tag != _w("p"):
            continue
        compact = _paragraph_match_key(_paragraph_text(child))
        if compact != "注：本任务书由上海电机学院教务处印制。":
            continue
        changes += int(_ensure_page_break_before_text(child))
        changes += int(_ensure_page_break_after_text(child))
        break
    return changes


def _ensure_page_break_before_text(paragraph: ET.Element) -> bool:
    first_run = _first_or_new_run(paragraph)
    children = list(first_run)
    if children and children[0].tag == _w("br") and children[0].attrib.get(_w("type")) == "page":
        return False
    page_break = ET.Element(_w("br"))
    page_break.set(_w("type"), "page")
    first_run.insert(1 if first_run.find("w:rPr", NS) is not None else 0, page_break)
    return True


def _ensure_page_break_after_text(paragraph: ET.Element) -> bool:
    if paragraph.findall(".//w:br[@w:type='page']", NS):
        return False
    run = ET.SubElement(paragraph, _w("r"))
    page_break = ET.SubElement(run, _w("br"))
    page_break.set(_w("type"), "page")
    return True


def _first_or_new_run(paragraph: ET.Element) -> ET.Element:
    run = paragraph.find("w:r", NS)
    if run is not None:
        return run
    return ET.SubElement(paragraph, _w("r"))


def _first_run_properties(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.findall("w:r", NS):
        if _run_text(run) or run.find("w:br", NS) is not None:
            return run.find("w:rPr", NS)
    return None


def _run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall("w:t", NS))


def _replace_optional_child(parent: ET.Element, local_name: str, source: ET.Element | None, insert_at: int = 0) -> None:
    existing = parent.find(f"w:{local_name}", NS)
    if existing is not None:
        parent.remove(existing)
    if source is not None:
        parent.insert(insert_at, copy.deepcopy(source))


def _read_relationships(package: zipfile.ZipFile) -> ET.Element:
    if "word/_rels/document.xml.rels" in package.namelist():
        return ET.fromstring(package.read("word/_rels/document.xml.rels"))
    root = ET.Element(_rel("Relationships"))
    return root


def _read_content_types(package: zipfile.ZipFile) -> ET.Element:
    if "[Content_Types].xml" in package.namelist():
        return ET.fromstring(package.read("[Content_Types].xml"))
    return ET.Element(_ct("Types"))


def _ensure_content_type(root: ET.Element, part_name: str, content_type: str) -> None:
    for child in root.findall(_ct("Override")):
        if child.attrib.get("PartName") == part_name:
            child.set("ContentType", content_type)
            return
    override = ET.Element(_ct("Override"))
    override.set("PartName", part_name)
    override.set("ContentType", content_type)
    root.append(override)


def _content_type_for_rel(rel_type: str) -> str:
    if rel_type.endswith("/header"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    if rel_type.endswith("/footer"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    return "application/octet-stream"


def _word_part_name(target: str) -> str:
    normalized = target.lstrip("/")
    if normalized.startswith("word/"):
        return normalized
    return "word/" + normalized


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _r(name: str) -> str:
    return f"{{{R_NS}}}{name}"


def _rel(name: str) -> str:
    return f"{{{REL_NS}}}{name}"


def _ct(name: str) -> str:
    return f"{{{CT_NS}}}{name}"
