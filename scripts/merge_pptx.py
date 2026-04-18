#!/usr/bin/env python3

"""Append slides from one PPTX deck into another while preserving source styles.

This utility is tailored to the fast-pptx workflow, where a regular deck and a
code deck are generated from different Pandoc reference documents. It appends
the source slides to the destination deck and carries forward the source slide
layouts, slide master, theme, and speaker notes so the appended slides retain
their original formatting.
"""

from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from typing import Dict, Iterable, Optional
import xml.etree.ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

NS = {"p": P_NS, "r": R_NS, "pr": PR_NS, "ct": CT_NS, "ep": EP_NS, "vt": VT_NS}

REL_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_TYPE_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
REL_TYPE_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
REL_TYPE_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
REL_TYPE_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
REL_TYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

CONTENT_TYPES = {
    "ppt/slides/slide": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    "ppt/notesSlides/notesSlide": "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    "ppt/slideLayouts/slideLayout": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "ppt/slideMasters/slideMaster": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "ppt/theme/theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "ppt/notesMasters/notesMaster": "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
}


class MergeError(RuntimeError):
    """Raised when the source deck uses relationships this merger cannot copy."""


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def rels_path(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def resolve_target(from_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(from_part), target))


def relative_target(from_part: str, to_part: str) -> str:
    return posixpath.relpath(to_part, posixpath.dirname(from_part))


def max_part_number(names: Iterable[str], directory: str, stem: str) -> int:
    pattern = re.compile(rf"^{re.escape(directory)}/{re.escape(stem)}(\d+)\.xml$")
    maximum = 0
    for name in names:
        match = pattern.match(name)
        if match:
            maximum = max(maximum, int(match.group(1)))
    return maximum


def next_numeric_id(existing: Iterable[str]) -> int:
    maximum = 0
    pattern = re.compile(r"rId(\d+)$")
    for value in existing:
        match = pattern.match(value)
        if match:
            maximum = max(maximum, int(match.group(1)))
    return maximum + 1


def part_content_type(part_name: str) -> str:
    for prefix, content_type in CONTENT_TYPES.items():
        if part_name.startswith(prefix) and part_name.endswith(".xml"):
            return content_type
    raise MergeError(f"Unsupported part type for content types: {part_name}")


def ensure_override(content_types_root: ET.Element, part_name: str) -> None:
    part_name = "/" + part_name.lstrip("/")
    for override in content_types_root.findall("ct:Override", NS):
        if override.get("PartName") == part_name:
            return
    ET.SubElement(
        content_types_root,
        qn(CT_NS, "Override"),
        {
            "PartName": part_name,
            "ContentType": part_content_type(part_name.lstrip("/")),
        },
    )


def relationship_by_id(rels_root: ET.Element) -> Dict[str, ET.Element]:
    return {rel.get("Id"): rel for rel in rels_root.findall("pr:Relationship", NS)}


def find_relationship(rels_root: ET.Element, rel_type: str) -> Optional[ET.Element]:
    for rel in rels_root.findall("pr:Relationship", NS):
        if rel.get("Type") == rel_type:
            return rel
    return None


def next_slide_id(presentation_root: ET.Element) -> int:
    maximum = 255
    slide_list = presentation_root.find("p:sldIdLst", NS)
    if slide_list is None:
        return 256
    for slide_id in slide_list.findall("p:sldId", NS):
        value = slide_id.get("id")
        if value is not None:
            maximum = max(maximum, int(value))
    return maximum + 1


def max_master_or_layout_id(entries: Dict[str, bytes], presentation_root: ET.Element) -> int:
    maximum = 2147483648
    master_list = presentation_root.find("p:sldMasterIdLst", NS)
    if master_list is not None:
        for master_id in master_list.findall("p:sldMasterId", NS):
            value = master_id.get("id")
            if value is not None:
                maximum = max(maximum, int(value))
    for name, data in entries.items():
        if not name.startswith("ppt/slideMasters/slideMaster") or not name.endswith(".xml"):
            continue
        root = parse_xml(data)
        layout_list = root.find("p:sldLayoutIdLst", NS)
        if layout_list is None:
            continue
        for layout_id in layout_list.findall("p:sldLayoutId", NS):
            value = layout_id.get("id")
            if value is not None:
                maximum = max(maximum, int(value))
    return maximum


def append_relationship(rels_root: ET.Element, rel_type: str, target: str, target_mode: Optional[str] = None) -> str:
    next_id = f"rId{next_numeric_id(rel.get('Id') for rel in rels_root.findall('pr:Relationship', NS))}"
    attrs = {"Id": next_id, "Type": rel_type, "Target": target}
    if target_mode:
        attrs["TargetMode"] = target_mode
    ET.SubElement(rels_root, qn(PR_NS, "Relationship"), attrs)
    return next_id


@dataclass
class PartAllocator:
    directory: str
    stem: str
    current: int

    def next(self) -> str:
        self.current += 1
        return f"{self.directory}/{self.stem}{self.current}.xml"


def copy_bytes(entries: Dict[str, bytes], src_entries: Dict[str, bytes], src_part: str, dst_part: str, content_types_root: ET.Element) -> None:
    entries[dst_part] = src_entries[src_part]
    ensure_override(content_types_root, dst_part)


def supported_slide_relationship(rel: ET.Element) -> bool:
    rel_type = rel.get("Type")
    if rel.get("TargetMode") == "External":
        return True
    return rel_type in {REL_TYPE_SLIDE_LAYOUT, REL_TYPE_NOTES_SLIDE}


def supported_notes_relationship(rel: ET.Element) -> bool:
    rel_type = rel.get("Type")
    if rel.get("TargetMode") == "External":
        return True
    return rel_type in {REL_TYPE_SLIDE, REL_TYPE_NOTES_MASTER}


def strip_missing_handout_master(entries: Dict[str, bytes]) -> None:
    presentation_rels_path = "ppt/_rels/presentation.xml.rels"
    presentation_path = "ppt/presentation.xml"

    if presentation_rels_path not in entries or presentation_path not in entries:
        return

    presentation_rels = parse_xml(entries[presentation_rels_path])
    removed = False

    for rel in list(presentation_rels.findall("pr:Relationship", NS)):
        if rel.get("Type") != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster":
            continue
        target = resolve_target("ppt/presentation.xml", rel.get("Target"))
        if target not in entries:
            presentation_rels.remove(rel)
            removed = True

    if not removed:
        return

    presentation = parse_xml(entries[presentation_path])
    handout_master_list = presentation.find("p:handoutMasterIdLst", NS)
    if handout_master_list is not None:
        presentation.remove(handout_master_list)

    entries[presentation_rels_path] = xml_bytes(presentation_rels)
    entries[presentation_path] = xml_bytes(presentation)


def slide_sort_key(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def slide_titles(entries: Dict[str, bytes]) -> list[str]:
    presentation = parse_xml(entries["ppt/presentation.xml"])
    presentation_rels = parse_xml(entries["ppt/_rels/presentation.xml.rels"])
    rel_by_id = relationship_by_id(presentation_rels)
    titles: list[str] = []

    slide_list = presentation.find("p:sldIdLst", NS)
    if slide_list is None:
        return titles

    for index, slide_id in enumerate(slide_list.findall("p:sldId", NS), start=1):
        rel_id = slide_id.get(qn(R_NS, "id"))
        slide_rel = rel_by_id.get(rel_id)
        if slide_rel is None or slide_rel.get("Type") != REL_TYPE_SLIDE:
            continue
        slide_part = resolve_target("ppt/presentation.xml", slide_rel.get("Target"))
        slide_root = parse_xml(entries[slide_part])
        title_text = ""
        for sp in slide_root.findall(".//p:sp", NS):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            if ph is None or ph.get("type") not in {None, "title", "ctrTitle"}:
                continue
            title_text = "".join(node.text or "" for node in sp.findall(".//a:t", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})).strip()
            if title_text:
                break
        titles.append(title_text or f"Slide {index}")

    return titles


def normalize_app_properties(entries: Dict[str, bytes]) -> None:
    app_path = "docProps/app.xml"
    if app_path not in entries:
        return

    app = parse_xml(entries[app_path])
    titles = slide_titles(entries)
    slide_count = len(titles)
    notes_count = len(
        [
            name
            for name in entries
            if name.startswith("ppt/notesSlides/notesSlide") and name.endswith(".xml")
        ]
    )

    slides_node = app.find("ep:Slides", NS)
    if slides_node is not None:
        slides_node.text = str(slide_count)

    notes_node = app.find("ep:Notes", NS)
    if notes_node is not None:
        notes_node.text = str(notes_count)

    heading_pairs = app.find("ep:HeadingPairs/vt:vector", NS)
    titles_of_parts = app.find("ep:TitlesOfParts/vt:vector", NS)
    if heading_pairs is not None and titles_of_parts is not None:
        variants = heading_pairs.findall("vt:variant", NS)
        slide_variant_index = None
        prefix_count = 0
        current_total = 0
        for idx in range(0, len(variants), 2):
            label_node = variants[idx].find("vt:lpstr", NS)
            count_node = variants[idx + 1].find("vt:i4", NS) if idx + 1 < len(variants) else None
            if label_node is None or count_node is None:
                continue
            label = label_node.text or ""
            count = int(count_node.text or "0")
            if label == "Slide Titles":
                slide_variant_index = idx + 1
                prefix_count = current_total
                break
            current_total += count

        if slide_variant_index is not None:
            variants[slide_variant_index].find("vt:i4", NS).text = str(slide_count)
            part_nodes = titles_of_parts.findall("vt:lpstr", NS)
            prefix_nodes = part_nodes[:prefix_count]
            for node in list(part_nodes):
                titles_of_parts.remove(node)
            for node in prefix_nodes:
                titles_of_parts.append(node)
            for title in titles:
                new_title = ET.SubElement(titles_of_parts, qn(VT_NS, "lpstr"))
                new_title.text = title
            titles_of_parts.set("size", str(prefix_count + slide_count))

    entries[app_path] = xml_bytes(app)


def fix_presentation(input_path: str, output_path: str) -> None:
    with zipfile.ZipFile(input_path) as pptx_zip:
        entries = OrderedDict((name, pptx_zip.read(name)) for name in pptx_zip.namelist())

    strip_missing_handout_master(entries)
    normalize_app_properties(entries)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for name, data in entries.items():
            out_zip.writestr(name, data)


def merge_presentations(destination_path: str, source_path: str, output_path: str) -> None:
    with zipfile.ZipFile(destination_path) as dst_zip:
        dest_entries = OrderedDict((name, dst_zip.read(name)) for name in dst_zip.namelist())
    with zipfile.ZipFile(source_path) as src_zip:
        src_entries = {name: src_zip.read(name) for name in src_zip.namelist()}

    strip_missing_handout_master(dest_entries)
    strip_missing_handout_master(src_entries)

    content_types = parse_xml(dest_entries["[Content_Types].xml"])
    dest_presentation = parse_xml(dest_entries["ppt/presentation.xml"])
    dest_presentation_rels = parse_xml(dest_entries["ppt/_rels/presentation.xml.rels"])
    src_presentation = parse_xml(src_entries["ppt/presentation.xml"])
    src_presentation_rels = parse_xml(src_entries["ppt/_rels/presentation.xml.rels"])

    src_presentation_rel_by_id = relationship_by_id(src_presentation_rels)

    slide_allocator = PartAllocator("ppt/slides", "slide", max_part_number(dest_entries.keys(), "ppt/slides", "slide"))
    notes_allocator = PartAllocator("ppt/notesSlides", "notesSlide", max_part_number(dest_entries.keys(), "ppt/notesSlides", "notesSlide"))
    layout_allocator = PartAllocator("ppt/slideLayouts", "slideLayout", max_part_number(dest_entries.keys(), "ppt/slideLayouts", "slideLayout"))
    master_allocator = PartAllocator("ppt/slideMasters", "slideMaster", max_part_number(dest_entries.keys(), "ppt/slideMasters", "slideMaster"))
    theme_allocator = PartAllocator("ppt/theme", "theme", max_part_number(dest_entries.keys(), "ppt/theme", "theme"))
    notes_master_allocator = PartAllocator("ppt/notesMasters", "notesMaster", max_part_number(dest_entries.keys(), "ppt/notesMasters", "notesMaster"))
    next_master_or_layout_numeric_id = max_master_or_layout_id(dest_entries, dest_presentation)

    layout_map: Dict[str, str] = {}
    master_map: Dict[str, str] = {}
    theme_map: Dict[str, str] = {}
    notes_slide_map: Dict[str, str] = {}
    appended_slides: list[tuple[str, str]] = []

    src_slide_ids = src_presentation.find("p:sldIdLst", NS)
    if src_slide_ids is None:
        raise MergeError("Source presentation does not contain any slides")

    dest_notes_master_part: Optional[str] = None
    dest_notes_master_rel = find_relationship(dest_presentation_rels, REL_TYPE_NOTES_MASTER)
    if dest_notes_master_rel is not None:
        dest_notes_master_part = resolve_target("ppt/presentation.xml", dest_notes_master_rel.get("Target"))

    source_uses_notes = False
    for src_slide_id in src_slide_ids.findall("p:sldId", NS):
        rel_id = src_slide_id.get(qn(R_NS, "id"))
        slide_rel = src_presentation_rel_by_id[rel_id]
        slide_part = resolve_target("ppt/presentation.xml", slide_rel.get("Target"))
        slide_rels_path = rels_path(slide_part)
        if slide_rels_path not in src_entries:
            continue
        slide_rels = parse_xml(src_entries[slide_rels_path])
        if find_relationship(slide_rels, REL_TYPE_NOTES_SLIDE) is not None:
            source_uses_notes = True
            break

    if source_uses_notes and dest_notes_master_part is None:
        src_notes_master_rel = find_relationship(src_presentation_rels, REL_TYPE_NOTES_MASTER)
        if src_notes_master_rel is None:
            raise MergeError("Source slides use notes, but the source deck has no notes master")
        src_notes_master_part = resolve_target("ppt/presentation.xml", src_notes_master_rel.get("Target"))
        new_notes_master_part = notes_master_allocator.next()
        src_notes_master_rels_part = rels_path(src_notes_master_part)
        if src_notes_master_rels_part not in src_entries:
            raise MergeError("Source notes master is missing its relationships part")
        notes_master_rels = parse_xml(src_entries[src_notes_master_rels_part])
        for rel in notes_master_rels.findall("pr:Relationship", NS):
            if rel.get("Type") != REL_TYPE_THEME:
                raise MergeError(f"Unsupported notes master relationship type: {rel.get('Type')}")
            old_theme_part = resolve_target(src_notes_master_part, rel.get("Target"))
            new_theme_part = theme_map.get(old_theme_part)
            if new_theme_part is None:
                new_theme_part = theme_allocator.next()
                theme_map[old_theme_part] = new_theme_part
                copy_bytes(dest_entries, src_entries, old_theme_part, new_theme_part, content_types)
            rel.set("Target", relative_target(new_notes_master_part, new_theme_part))
        copy_bytes(dest_entries, src_entries, src_notes_master_part, new_notes_master_part, content_types)
        dest_entries[rels_path(new_notes_master_part)] = xml_bytes(notes_master_rels)
        notes_master_rel_id = append_relationship(
            dest_presentation_rels,
            REL_TYPE_NOTES_MASTER,
            relative_target("ppt/presentation.xml", new_notes_master_part),
        )
        notes_master_id_list = dest_presentation.find("p:notesMasterIdLst", NS)
        if notes_master_id_list is None:
            slide_master_id_list = dest_presentation.find("p:sldMasterIdLst", NS)
            notes_master_id_list = ET.Element(qn(P_NS, "notesMasterIdLst"))
            notes_master_id = ET.SubElement(notes_master_id_list, qn(P_NS, "notesMasterId"))
            notes_master_id.set(qn(R_NS, "id"), notes_master_rel_id)
            insert_at = 1 if slide_master_id_list is not None else 0
            dest_presentation.insert(insert_at, notes_master_id_list)
        dest_notes_master_part = new_notes_master_part

    slide_master_id_list = dest_presentation.find("p:sldMasterIdLst", NS)
    if slide_master_id_list is None:
        slide_master_id_list = ET.SubElement(dest_presentation, qn(P_NS, "sldMasterIdLst"))

    for src_slide_id in src_slide_ids.findall("p:sldId", NS):
        rel_id = src_slide_id.get(qn(R_NS, "id"))
        src_slide_rel = src_presentation_rel_by_id[rel_id]
        src_slide_part = resolve_target("ppt/presentation.xml", src_slide_rel.get("Target"))
        src_slide_rels_path = rels_path(src_slide_part)
        if src_slide_rels_path not in src_entries:
            raise MergeError(f"Source slide is missing relationships: {src_slide_part}")
        src_slide_rels = parse_xml(src_entries[src_slide_rels_path])

        for rel in src_slide_rels.findall("pr:Relationship", NS):
            if not supported_slide_relationship(rel):
                raise MergeError(
                    f"Unsupported relationship on source slide {src_slide_part}: {rel.get('Type')}"
                )

        layout_rel = find_relationship(src_slide_rels, REL_TYPE_SLIDE_LAYOUT)
        if layout_rel is None:
            raise MergeError(f"Source slide does not reference a slide layout: {src_slide_part}")
        src_layout_part = resolve_target(src_slide_part, layout_rel.get("Target"))
        new_layout_part = layout_map.get(src_layout_part)

        if new_layout_part is None:
            src_layout_rels = parse_xml(src_entries[rels_path(src_layout_part)])
            layout_master_rel = find_relationship(src_layout_rels, REL_TYPE_SLIDE_MASTER)
            if layout_master_rel is None:
                raise MergeError(f"Source layout does not reference a slide master: {src_layout_part}")
            src_master_part = resolve_target(src_layout_part, layout_master_rel.get("Target"))
            new_master_part = master_map.get(src_master_part)

            if new_master_part is None:
                src_master_rels_part = rels_path(src_master_part)
                if src_master_rels_part not in src_entries:
                    raise MergeError(f"Source slide master is missing relationships: {src_master_part}")
                src_master_rels = parse_xml(src_entries[src_master_rels_part])
                src_master_root = parse_xml(src_entries[src_master_part])
                new_master_part = master_allocator.next()
                master_map[src_master_part] = new_master_part

                for rel in src_master_rels.findall("pr:Relationship", NS):
                    rel_type = rel.get("Type")
                    if rel_type == REL_TYPE_THEME:
                        old_theme_part = resolve_target(src_master_part, rel.get("Target"))
                        new_theme_part = theme_map.get(old_theme_part)
                        if new_theme_part is None:
                            new_theme_part = theme_allocator.next()
                            theme_map[old_theme_part] = new_theme_part
                            copy_bytes(dest_entries, src_entries, old_theme_part, new_theme_part, content_types)
                        rel.set("Target", relative_target(new_master_part, new_theme_part))
                    elif rel_type == REL_TYPE_SLIDE_LAYOUT:
                        old_layout_part = resolve_target(src_master_part, rel.get("Target"))
                        mapped_layout_part = layout_map.get(old_layout_part)
                        if mapped_layout_part is None:
                            mapped_layout_part = layout_allocator.next()
                            layout_map[old_layout_part] = mapped_layout_part
                            copy_bytes(dest_entries, src_entries, old_layout_part, mapped_layout_part, content_types)
                            old_layout_rels_part = rels_path(old_layout_part)
                            if old_layout_rels_part not in src_entries:
                                raise MergeError(f"Source layout is missing relationships: {old_layout_part}")
                            old_layout_rels = parse_xml(src_entries[old_layout_rels_part])
                            for layout_rel in old_layout_rels.findall("pr:Relationship", NS):
                                if layout_rel.get("Type") != REL_TYPE_SLIDE_MASTER:
                                    raise MergeError(
                                        f"Unsupported slide layout relationship type: {layout_rel.get('Type')}"
                                    )
                                layout_rel.set("Target", relative_target(mapped_layout_part, new_master_part))
                            dest_entries[rels_path(mapped_layout_part)] = xml_bytes(old_layout_rels)
                        rel.set("Target", relative_target(new_master_part, mapped_layout_part))
                    else:
                        raise MergeError(f"Unsupported slide master relationship type: {rel_type}")

                layout_id_list = src_master_root.find("p:sldLayoutIdLst", NS)
                new_master_numeric_id = None
                next_master_or_layout_numeric_id += 1
                new_master_numeric_id = next_master_or_layout_numeric_id
                if layout_id_list is not None:
                    for layout_id in layout_id_list.findall("p:sldLayoutId", NS):
                        next_master_or_layout_numeric_id += 1
                        layout_id.set("id", str(next_master_or_layout_numeric_id))

                dest_entries[new_master_part] = xml_bytes(src_master_root)
                ensure_override(content_types, new_master_part)
                dest_entries[rels_path(new_master_part)] = xml_bytes(src_master_rels)

                new_master_rel_id = append_relationship(
                    dest_presentation_rels,
                    REL_TYPE_SLIDE_MASTER,
                    relative_target("ppt/presentation.xml", new_master_part),
                )
                new_master_id = ET.SubElement(slide_master_id_list, qn(P_NS, "sldMasterId"))
                new_master_id.set("id", str(new_master_numeric_id))
                new_master_id.set(qn(R_NS, "id"), new_master_rel_id)
            new_layout_part = layout_map[src_layout_part]

        new_slide_part = slide_allocator.next()
        appended_slides.append((src_slide_part, new_slide_part))
        copy_bytes(dest_entries, src_entries, src_slide_part, new_slide_part, content_types)

        notes_rel = find_relationship(src_slide_rels, REL_TYPE_NOTES_SLIDE)
        if notes_rel is not None:
            src_notes_slide_part = resolve_target(src_slide_part, notes_rel.get("Target"))
            new_notes_slide_part = notes_slide_map.get(src_notes_slide_part)
            if new_notes_slide_part is None:
                new_notes_slide_part = notes_allocator.next()
                notes_slide_map[src_notes_slide_part] = new_notes_slide_part
                copy_bytes(dest_entries, src_entries, src_notes_slide_part, new_notes_slide_part, content_types)
                src_notes_slide_rels_part = rels_path(src_notes_slide_part)
                if src_notes_slide_rels_part not in src_entries:
                    raise MergeError(f"Source notes slide is missing relationships: {src_notes_slide_part}")
                notes_slide_rels = parse_xml(src_entries[src_notes_slide_rels_part])
                for rel in notes_slide_rels.findall("pr:Relationship", NS):
                    if not supported_notes_relationship(rel):
                        raise MergeError(
                            f"Unsupported relationship on source notes slide {src_notes_slide_part}: {rel.get('Type')}"
                        )
                    if rel.get("Type") == REL_TYPE_NOTES_MASTER:
                        if dest_notes_master_part is None:
                            raise MergeError("Notes slide requires a notes master, but none is available")
                        rel.set("Target", relative_target(new_notes_slide_part, dest_notes_master_part))
                dest_entries[rels_path(new_notes_slide_part)] = xml_bytes(notes_slide_rels)
            notes_rel.set("Target", relative_target(new_slide_part, new_notes_slide_part))

            notes_slide_rels = parse_xml(dest_entries[rels_path(new_notes_slide_part)])
            slide_rel = find_relationship(notes_slide_rels, REL_TYPE_SLIDE)
            if slide_rel is None:
                raise MergeError(f"Notes slide does not reference its owning slide: {new_notes_slide_part}")
            slide_rel.set("Target", relative_target(new_notes_slide_part, new_slide_part))
            dest_entries[rels_path(new_notes_slide_part)] = xml_bytes(notes_slide_rels)

        layout_rel.set("Target", relative_target(new_slide_part, new_layout_part))
        dest_entries[rels_path(new_slide_part)] = xml_bytes(src_slide_rels)

        new_slide_rel_id = append_relationship(
            dest_presentation_rels,
            REL_TYPE_SLIDE,
            relative_target("ppt/presentation.xml", new_slide_part),
        )
        slide_id_list = dest_presentation.find("p:sldIdLst", NS)
        if slide_id_list is None:
            slide_id_list = ET.SubElement(dest_presentation, qn(P_NS, "sldIdLst"))
        slide_id = ET.SubElement(slide_id_list, qn(P_NS, "sldId"))
        slide_id.set("id", str(next_slide_id(dest_presentation)))
        slide_id.set(qn(R_NS, "id"), new_slide_rel_id)

    for src_slide_part, new_slide_part in appended_slides:
        src_slide_rels = parse_xml(src_entries[rels_path(src_slide_part)])
        merged_slide_rels = parse_xml(dest_entries[rels_path(new_slide_part)])
        src_layout_rel = find_relationship(src_slide_rels, REL_TYPE_SLIDE_LAYOUT)
        merged_layout_rel = find_relationship(merged_slide_rels, REL_TYPE_SLIDE_LAYOUT)
        if src_layout_rel is not None and merged_layout_rel is not None:
            src_layout_part = resolve_target(src_slide_part, src_layout_rel.get("Target"))
            merged_layout_rel.set("Target", relative_target(new_slide_part, layout_map[src_layout_part]))
        src_notes_rel = find_relationship(src_slide_rels, REL_TYPE_NOTES_SLIDE)
        merged_notes_rel = find_relationship(merged_slide_rels, REL_TYPE_NOTES_SLIDE)
        if src_notes_rel is not None and merged_notes_rel is not None:
            src_notes_part = resolve_target(src_slide_part, src_notes_rel.get("Target"))
            merged_notes_rel.set("Target", relative_target(new_slide_part, notes_slide_map[src_notes_part]))
        dest_entries[rels_path(new_slide_part)] = xml_bytes(merged_slide_rels)

    dest_entries["ppt/presentation.xml"] = xml_bytes(dest_presentation)
    dest_entries["ppt/_rels/presentation.xml.rels"] = xml_bytes(dest_presentation_rels)
    dest_entries["[Content_Types].xml"] = xml_bytes(content_types)
    normalize_app_properties(dest_entries)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for name, data in dest_entries.items():
            out_zip.writestr(name, data)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Fix a PPTX package or append slides from one PPTX into another while preserving source styles."
    )
    parser.add_argument(
        "paths",
        nargs="+",
        help=(
            "Either INPUT OUTPUT to fix a PPTX in place/copy, or DESTINATION SOURCE OUTPUT "
            "to append SOURCE slides into DESTINATION."
        ),
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    try:
        if len(args.paths) == 2:
            fix_presentation(args.paths[0], args.paths[1])
        elif len(args.paths) == 3:
            merge_presentations(args.paths[0], args.paths[1], args.paths[2])
        else:
            raise MergeError("Expected either INPUT OUTPUT or DESTINATION SOURCE OUTPUT")
    except MergeError as exc:
        print(f"merge_pptx.py: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
