"""Reopen visual DOCX drafts and inspect Word-authored templates for editing."""

from hashlib import sha256
from io import BytesIO
import json
import re
from zipfile import ZipFile

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from lxml import etree
from report_context import REPORTS, loop_collections


METADATA = "customXml/virvig-template.xml"
METADATA_TAG = "{urn:virvig:visual-template}template"
VERSION = 1
TAG = re.compile(r"(\{\{.*?\}\}|\{%.*?%\})", re.S)
CONTROL = re.compile(r"(for|if|else|endif|endfor)(?:\s+(.+))?")
LOOP = re.compile(r"(\w+)\s+in\s+([\w.]+)")
LOOP_VARIABLES = {key: value for report in REPORTS.values() for key, value in loop_collections(report).items()}


def fingerprint(archive):
    digest = sha256()
    for name in sorted(set(archive.namelist()) - {METADATA}):
        digest.update(name.encode("utf-8") + b"\0")
        digest.update(sha256(archive.read(name)).digest())
    return digest.hexdigest()


def remove_visual_metadata(document):
    part_name = PackURI("/" + METADATA)
    for relationship in list(document.part.rels.values()):
        if not relationship.is_external and relationship.target_part.partname == part_name:
            document.part.drop_rel(relationship.rId)


def save_visual_template(document, payload):
    part_name = PackURI("/" + METADATA)
    metadata = etree.Element(METADATA_TAG, version=str(VERSION), fingerprint="")
    metadata.text = payload
    part = Part(part_name, "application/xml", etree.tostring(metadata), document.part.package)
    document.part.relate_to(part, RELATIONSHIP_TYPE.CUSTOM_XML)
    intermediate = BytesIO()
    document.save(intermediate)
    with ZipFile(intermediate) as archive:
        metadata.set("fingerprint", fingerprint(archive))
    part._blob = etree.tostring(metadata)
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def import_visual_template(data, kind):
    with ZipFile(BytesIO(data)) as archive:
        if METADATA in archive.namelist():
            metadata = etree.fromstring(archive.read(METADATA), parser=etree.XMLParser(resolve_entities=False))
            if metadata.tag != METADATA_TAG or metadata.get("version") != str(VERSION):
                raise ValueError("This visual template uses an unsupported editor version.")
            if metadata.get("fingerprint") != fingerprint(archive):
                raise ValueError("This DOCX changed after it was saved by the editor. Its earlier draft cannot be restored.")
            draft = json.loads(metadata.text or "{}")
            if draft.get("kind") != kind:
                raise ValueError("Select the report type this visual template was saved for.")
            if draft.get("document", {}).get("type") != "doc":
                raise ValueError("The saved visual draft is missing its document.")
            return draft
    return {"kind": kind, "document": import_word_body(Document(BytesIO(data)))}


def control_node(source, kind):
    match = CONTROL.fullmatch(source.strip())
    if not match:
        raise ValueError("This template uses Jinja control tags outside the editor’s supported subset.")
    action, value = match.groups()
    if action == "for":
        loop = LOOP.fullmatch(value or "")
        if not loop or LOOP_VARIABLES.get(loop[2]) != loop[1]:
            raise ValueError("This template repeats a collection the visual editor cannot represent.")
        value = loop[2]
        action = "repeat"
    elif action in ("endif", "endfor"):
        value = action
        action = "end"
    return {"type": kind, "attrs": {"action": action, "value": value or ""}}


def import_paragraph(paragraph):
    # Join adjacent Word runs when looking for a tag. Word can split a tag at a
    # proofing boundary, even when it has one visible character style.
    pieces = []
    for run in paragraph.runs:
        properties = run._element.rPr
        if properties is not None and any(child.tag not in {qn("w:b"), qn("w:i"), qn("w:u")} for child in properties):
            raise ValueError("This template contains character formatting the visual editor cannot preserve.")
        if any(getattr(run, name) is False for name in ("bold", "italic", "underline")) or run.underline not in (None, True):
            raise ValueError("This template uses character style overrides the visual editor cannot preserve.")
        if any(child.tag not in {qn("w:rPr"), qn("w:t"), qn("w:br"), qn("w:cr")} for child in run._element):
            raise ValueError("This template contains unsupported content inside a paragraph.")
        if any(br.get(qn("w:type")) not in (None, "textWrapping") for br in run._element.findall(qn("w:br"))):
            raise ValueError("This template contains a manual page or column break.")
        marks = [{"type": name} for name in ("bold", "italic", "underline") if getattr(run, name)]
        pieces.extend((character, marks) for character in run.text)
    text = "".join(character for character, _ in pieces)
    structural = re.fullmatch(r"\s*\{%p\s+(.*?)\s*%\}\s*", text, re.S)
    if structural:
        return control_node(structural[1], "control")

    content = []
    offset = 0
    for value in TAG.split(text):
        if value.startswith("{{") and value.endswith("}}"):
            rich = value.startswith("{{r ")
            content.append({"type": "field", "marks": pieces[offset][1],
                            "attrs": {"expression": value[3 if rich else 2:-2].strip(), "rich": rich}})
        elif value.startswith("{%") and value.endswith("%}"):
            content.append(control_node(value[2:-2], "inlineControl"))
        else:
            if re.search(r"\{[\{%#]", value):
                raise ValueError("This template contains a Jinja tag the editor cannot interpret.")
            for character, marks in pieces[offset:offset + len(value)]:
                if character == "\n":
                    content.append({"type": "hardBreak"})
                elif content and content[-1]["type"] == "text" and content[-1]["marks"] == marks:
                    content[-1]["text"] += character
                else:
                    content.append({"type": "text", "text": character, "marks": marks})
        offset += len(value)
    ppr = paragraph._element.pPr
    allowed = {qn("w:" + name) for name in ("pStyle", "keepNext", "keepLines", "spacing", "jc", "widowControl", "ind")}
    if ppr is not None and any(child.tag not in allowed for child in ppr):
        raise ValueError("This template includes paragraph formatting the visual editor cannot preserve.")
    style = paragraph.style.style_id
    heading = re.fullmatch(r"Heading([1-3])", style)
    attrs = {"wordProperties": etree.tostring(ppr, encoding="unicode") if ppr is not None else None,
             "textAlign": {0: "left", 1: "center", 2: "right", 3: "justify"}.get(paragraph.alignment)}
    if heading:
        attrs["level"] = int(heading[1])
    return {"type": "heading" if heading else "paragraph", "attrs": attrs, "content": content}


def import_word_body(document):
    body = document.element.body
    if any(node.tag not in (qn("w:p"), qn("w:sectPr")) for node in body):
        raise ValueError("This Word-authored template has tables or other document blocks the editor cannot preserve.")
    if body.sectPr is not None and any(node.tag not in {qn("w:" + name) for name in ("pgSz", "pgMar", "cols", "docGrid")} for node in body.sectPr):
        raise ValueError("This template contains section features, headers, or footers the editor cannot preserve.")
    allowed = {qn("w:" + name) for name in ("pPr", "r", "proofErr", "bookmarkStart", "bookmarkEnd")}
    if any(node.tag not in allowed for paragraph in document.paragraphs for node in paragraph._element):
        raise ValueError("This template has hyperlinks, revisions, or other Word content the editor cannot preserve.")
    nodes = [import_paragraph(paragraph) for paragraph in document.paragraphs]
    if not nodes or nodes[-1]["type"] == "control":
        nodes.append({"type": "paragraph"})
    return {"type": "doc", "content": nodes}
