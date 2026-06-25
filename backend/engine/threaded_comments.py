"""
Threaded-comment post-processor for .xlsx files.

openpyxl only writes legacy "Notes" (the yellow hover-only sticky notes).
This module post-processes a saved .xlsx so the same comments show up as
modern *threaded Comments* in Excel and Google Sheets — meaning they:
    • appear together in the Comments side panel,
    • can be replied to,
    • can be @-mentioned to specific people.

How Excel signals "this comment is threaded":
    The legacy comments file keeps an entry per cell whose author is
    "tc={guid}". The same guid is the id of a <threadedComment> in
    xl/threadedComments/threadedCommentN.xml. The worksheet rels declares
    a relationship pointing at that threadedComment file.

Usage:
    inject_threaded_comments(
        xlsx_path,
        sheet_name="Track Check List",
        author_name="Metadata Decoder",
        comments=[("A2", "Misspelled artist…"), ("B5", "Duplicate ISRC…"), …],
    )
"""
from __future__ import annotations

import shutil
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

REL_COMMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_TC = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"

CT_TC = "application/vnd.ms-excel.threadedcomments+xml"
CT_PERSON = "application/vnd.ms-excel.person+xml"

DECODER_PERSON_ID = "{D6CD0001-0001-4DEC-0DEC-DECDECDECDEC}"


def _new_guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;"))


def _normalize_part_path(target: str, base_dir: str) -> str:
    """
    Resolve a Target attribute (which may be relative '../comments1.xml',
    sibling 'comments1.xml', or rooted '/xl/comments1.xml') into a canonical
    zip member path like 'xl/comments1.xml'.
    """
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base_dir.rstrip("/") + "/" + target).split("/")
    out: list[str] = []
    for p in parts:
        if p == "" or p == ".":
            continue
        if p == "..":
            if out:
                out.pop()
            continue
        out.append(p)
    return "/".join(out)


def _find_worksheet_member(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """Return the zip path of the worksheet xml whose tab is `sheet_name`."""
    wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {
        r.attrib["Id"]: r.attrib["Target"]
        for r in rels.findall(f"{{{NS_REL}}}Relationship")
    }
    for s in wb_xml.find(f"{{{NS_MAIN}}}sheets"):
        if s.attrib.get("name") != sheet_name:
            continue
        rid = s.attrib.get(f"{{{NS_DOC_REL}}}id")
        target = rid_to_target.get(rid, "")
        return _normalize_part_path(target, base_dir="xl")
    raise ValueError(f"Sheet {sheet_name!r} not found in workbook.")


def _ensure_relationship(rels_xml_bytes: bytes | None, rel_type: str, target: str) -> bytes:
    """Idempotently add a Relationship to a .rels XML body."""
    ET.register_namespace("", NS_REL)
    if rels_xml_bytes is None:
        tree = ET.fromstring(f'<Relationships xmlns="{NS_REL}"/>'.encode("utf-8"))
    else:
        tree = ET.fromstring(rels_xml_bytes)
    for rel in tree.findall(f"{{{NS_REL}}}Relationship"):
        if rel.attrib.get("Type") == rel_type and rel.attrib.get("Target") == target:
            return ET.tostring(tree, xml_declaration=True, encoding="UTF-8",
                               short_empty_elements=True)
    used = {r.attrib["Id"] for r in tree.findall(f"{{{NS_REL}}}Relationship")}
    n = 1
    while f"rId{n}" in used:
        n += 1
    new_rel = ET.SubElement(tree, f"{{{NS_REL}}}Relationship")
    new_rel.set("Id", f"rId{n}")
    new_rel.set("Type", rel_type)
    new_rel.set("Target", target)
    return ET.tostring(tree, xml_declaration=True, encoding="UTF-8",
                       short_empty_elements=True)


def _ensure_content_type_override(ct_bytes: bytes, part_name: str, content_type: str) -> bytes:
    ET.register_namespace("", NS_CT)
    tree = ET.fromstring(ct_bytes)
    for ovr in tree.findall(f"{{{NS_CT}}}Override"):
        if ovr.attrib.get("PartName") == part_name:
            ovr.set("ContentType", content_type)
            return ET.tostring(tree, xml_declaration=True, encoding="UTF-8",
                               short_empty_elements=True)
    ovr = ET.SubElement(tree, f"{{{NS_CT}}}Override")
    ovr.set("PartName", part_name)
    ovr.set("ContentType", content_type)
    return ET.tostring(tree, xml_declaration=True, encoding="UTF-8",
                       short_empty_elements=True)


def inject_threaded_comments(
    xlsx_path: Path,
    sheet_name: str,
    author_name: str,
    comments: list[tuple[str, str]],
) -> int:
    """
    Convert the legacy comments openpyxl wrote on `sheet_name` into modern
    threaded Comments. `comments` is the list of (cell_ref, text) pairs
    that should appear; this list is also what we'll write to the
    threaded-comments file.

    Returns the number of threaded comments written.
    """
    if not comments:
        return 0

    xlsx_path = Path(xlsx_path)
    tmp_path = xlsx_path.with_suffix(xlsx_path.suffix + ".tcwip")

    with zipfile.ZipFile(xlsx_path, "r") as src:
        all_members = {n: src.read(n) for n in src.namelist()}

        # Locate the worksheet xml + its rels file inside the zip.
        ws_member = _find_worksheet_member(src, sheet_name)        # xl/worksheets/sheet1.xml
        ws_dir = ws_member.rsplit("/", 1)[0]                       # xl/worksheets
        ws_basename = ws_member.rsplit("/", 1)[1]                  # sheet1.xml
        ws_rels_member = f"{ws_dir}/_rels/{ws_basename}.rels"      # xl/worksheets/_rels/sheet1.xml.rels

        # The legacy comments file path is whatever the worksheet rels
        # already points at (openpyxl picks its own naming).
        existing_ws_rels = all_members.get(ws_rels_member)
        comments_member: str | None = None
        if existing_ws_rels is not None:
            ws_rels_tree = ET.fromstring(existing_ws_rels)
            for r in ws_rels_tree.findall(f"{{{NS_REL}}}Relationship"):
                if r.attrib.get("Type") == REL_COMMENT:
                    comments_member = _normalize_part_path(
                        r.attrib.get("Target", ""), base_dir=ws_dir
                    )
                    break
        if comments_member is None:
            comments_member = "xl/comments1.xml"  # sane default

    # If a previous run already wrote threaded comments for this sheet, reuse
    # that path so we overwrite stale entries instead of accumulating files
    # across re-scans.
    threaded_member: str | None = None
    if existing_ws_rels is not None:
        ws_rels_tree = ET.fromstring(existing_ws_rels)
        for r in ws_rels_tree.findall(f"{{{NS_REL}}}Relationship"):
            if r.attrib.get("Type") == REL_TC:
                threaded_member = _normalize_part_path(
                    r.attrib.get("Target", ""), base_dir=ws_dir
                )
                break
    if threaded_member is None:
        n = 1
        while f"xl/threadedComments/threadedComment{n}.xml" in all_members:
            n += 1
        threaded_member = f"xl/threadedComments/threadedComment{n}.xml"
    persons_member = "xl/persons/person.xml"

    # ---- Build the new XML payloads --------------------------------------

    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.00Z")
    threaded_entries: list[dict] = []
    legacy_authors: list[str] = []
    for ref, text in comments:
        thread_id = _new_guid()
        legacy_authors.append(f"tc={thread_id}")
        threaded_entries.append({"ref": ref, "text": text, "thread_id": thread_id})

    # New legacy comments.xml — every entry has author tc={guid}, which is
    # how Excel knows it's a threaded comment.
    auth_xml = "".join(f"<author>{a}</author>" for a in legacy_authors)
    com_parts: list[str] = []
    for i, e in enumerate(threaded_entries):
        safe = _xml_escape(e["text"])
        com_parts.append(
            f'<comment ref="{e["ref"]}" authorId="{i}" shapeId="0" '
            f'xr:uid="{e["thread_id"]}">'
            f'<text><r><rPr><sz val="10"/><color theme="1"/>'
            f'<rFont val="Arial"/></rPr><t xml:space="preserve">{safe}</t></r></text>'
            '</comment>'
        )
    new_comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<comments xmlns="{NS_MAIN}" xmlns:xr="{NS_XR}">'
        f'<authors>{auth_xml}</authors>'
        f'<commentList>{"".join(com_parts)}</commentList>'
        '</comments>'
    ).encode("utf-8")

    tc_parts: list[str] = []
    for e in threaded_entries:
        safe = _xml_escape(e["text"])
        tc_parts.append(
            f'<threadedComment ref="{e["ref"]}" dT="{now_iso}" '
            f'personId="{DECODER_PERSON_ID}" id="{e["thread_id"]}">'
            f'<text>{safe}</text>'
            '</threadedComment>'
        )
    new_threaded_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<ThreadedComments xmlns="{NS_TC}">'
        f'{"".join(tc_parts)}'
        '</ThreadedComments>'
    ).encode("utf-8")

    new_persons_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<personList xmlns="{NS_TC}" xmlns:x="{NS_MAIN}">'
        f'<person displayName="{_xml_escape(author_name)}" '
        f'id="{DECODER_PERSON_ID}" providerId="None"/>'
        '</personList>'
    ).encode("utf-8")

    # Worksheet rels: add the threadedComment relationship (target is
    # relative to the worksheet's directory).
    tc_target_rel = f"../{threaded_member.split('/', 1)[1].rsplit('/', 1)[0]}/{threaded_member.rsplit('/', 1)[1]}"
    new_ws_rels = _ensure_relationship(existing_ws_rels, REL_TC, tc_target_rel)

    # Workbook rels: add the persons relationship (target is relative to xl/).
    existing_wb_rels = all_members.get("xl/_rels/workbook.xml.rels")
    persons_target_rel = "persons/person.xml"
    new_wb_rels = _ensure_relationship(existing_wb_rels, REL_PERSON, persons_target_rel)

    # Content_Types.xml: declare overrides for the new parts.
    new_ct = all_members["[Content_Types].xml"]
    new_ct = _ensure_content_type_override(new_ct, f"/{threaded_member}", CT_TC)
    new_ct = _ensure_content_type_override(new_ct, f"/{persons_member}", CT_PERSON)

    # ---- Write the modified zip ------------------------------------------

    replacements = {
        comments_member: new_comments_xml,
        threaded_member: new_threaded_xml,
        persons_member: new_persons_xml,
        ws_rels_member: new_ws_rels,
        "xl/_rels/workbook.xml.rels": new_wb_rels,
        "[Content_Types].xml": new_ct,
    }

    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out:
        written = set()
        for name, data in all_members.items():
            data_to_write = replacements[name] if name in replacements else data
            out.writestr(name, data_to_write)
            written.add(name)
        for name, data in replacements.items():
            if name not in written:
                out.writestr(name, data)

    shutil.move(tmp_path, xlsx_path)
    return len(threaded_entries)
