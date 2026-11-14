"""Convert a visual template into a docxtpl-compatible Word document."""

import json
from io import BytesIO
from pathlib import Path
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu
from jinja2 import StrictUndefined, meta
from jinja2.sandbox import SandboxedEnvironment
from template_fields import REPORTS
from template_import import remove_visual_metadata, save_visual_template


ENV = SandboxedEnvironment(undefined=StrictUndefined)
ALIGNMENTS = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
              "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}


def expression(value, scope):
    value = value.strip()
    if not value or any(tag in value for tag in ("{{", "}}", "{%", "%}", "{#", "#}")):
        raise ValueError("Enter a Jinja expression without template delimiters.")
    tree = ENV.parse("{{ " + value + " }}")
    unknown = meta.find_undeclared_variables(tree) - scope
    if unknown:
        raise ValueError("Unknown template variable: " + ", ".join(sorted(unknown)))
    return value


def control_tag(attrs, scope, stack, collections):
    action, value = attrs["action"], attrs.get("value", "")
    if action == "repeat":
        iterable, variable = repeat_scope(value, collections, scope)
        stack.append(("repeat", scope, False))
        return f"for {variable} in {iterable}", scope | {variable, "loop"}
    if action == "if":
        stack.append(("if", scope, False))
        return "if " + expression(value, scope), scope
    if action == "else" and stack and not stack[-1][2]:
        block, previous, _ = stack[-1]
        stack[-1] = (block, previous, True)
        return "else", previous
    if action == "end" and stack:
        block, previous, _ = stack.pop()
        tag = "endfor" if block == "repeat" else "endif"
        if not value or value == tag:
            return tag, previous
    raise ValueError("A template control does not match its opening block.")


def add_inline(paragraph, nodes, scope):
    stack = []
    for index, node in enumerate(nodes):
        kind = node["type"]
        attrs = node.get("attrs", {})
        if kind == "hardBreak":
            paragraph.add_run().add_break()
            continue
        if kind == "inlineControl":
            tag, scope = control_tag(attrs, scope, stack, {
                "publications": "pub", "groups": "group", "group.publications": "pub",
                "activities": "activity", "theses": "thesis", "projects": "project",
            })
            paragraph.add_run("{% " + tag + " %}")
            continue
        if kind == "field":
            value = expression(attrs["expression"], scope)
            if attrs.get("condition"):
                condition = expression(attrs["condition"], scope)
                value = f"({value}) if ({condition}) else ''"
            opening = "{{r " if attrs.get("rich") else "{{ "
            text = opening + value + " }}"
        elif kind == "text":
            # Literal text typed in the editor must remain literal in Word.
            text = node.get("text", "")
            next_text = nodes[index + 1].get("text", "") if index + 1 < len(nodes) else ""
            if re.search(r"\{[\{%#]", text) or (text.endswith("{") and next_text[:1] in ("{", "%", "#")):
                raise ValueError("Insert Jinja expressions with the Field or Expression controls.")
        else:
            raise ValueError(f"Unsupported inline element: {kind}.")
        marks = {mark["type"] for mark in node.get("marks", ())}
        run = paragraph.add_run(text)
        run.bold = True if "bold" in marks else None
        run.italic = True if "italic" in marks else None
        run.underline = True if "underline" in marks else None
    if stack:
        raise ValueError("Close each inline loop or condition within its paragraph.")


def add_paragraph(target, node, scope, *, paragraph=None, style=None):
    attrs = node.get("attrs", {})
    if node["type"] == "heading":
        style = f"Heading {attrs.get('level', 1)}"
    if paragraph is None:
        paragraph = target.add_paragraph()
    properties = attrs.get("wordProperties")
    if properties:
        word_properties = parse_xml(properties.encode())
        if word_properties.tag != qn("w:pPr"):
            raise ValueError("Invalid paragraph properties in the visual template.")
        paragraph._element.insert(0, word_properties)
    if style or not properties or not attrs.get("useWordStyle", True):
        paragraph.style = style or "Normal"
    paragraph.alignment = ALIGNMENTS.get(attrs.get("textAlign"))
    add_inline(paragraph, node.get("content", ()), scope)


def add_list(target, node, scope, *, paragraph=None, level=0):
    style = "List Number" if node["type"] == "orderedList" else "List Bullet"
    for item in node.get("content", ()):
        for block in item.get("content", ()):
            if block["type"] in ("bulletList", "orderedList"):
                add_list(target, block, scope, level=level + 1)
            else:
                add_paragraph(target, block, scope, paragraph=paragraph, style=style + (f" {min(level + 1, 3)}" if level else ""))
                paragraph = None


def add_table(document, node, scope, collections):
    rows = node.get("content", ())
    # Table grids include the columns covered by vertical merges.
    occupancy = {}
    layout = []
    widths = {}
    for y, row in enumerate(rows):
        cells = []
        x = 0
        for cell in row.get("content", ()):
            while (y, x) in occupancy:
                x += 1
            attrs = cell.get("attrs", {})
            horizontal, vertical = attrs.get("colspan", 1), attrs.get("rowspan", 1)
            for yy in range(y, y + vertical):
                for xx in range(x, x + horizontal):
                    occupancy[yy, xx] = True
            for offset, width in enumerate(attrs.get("colwidth") or ()):
                if width:
                    widths[x + offset] = width
            cells.append((x, horizontal, vertical, cell))
            x += horizontal
        layout.append(cells)
    if not occupancy:
        return
    columns = max(x for _, x in occupancy) + 1
    repeat = [row.get("attrs", {}).get("repeat") for row in rows]
    for y, cells in enumerate(layout):
        for _, _, vertical, _ in cells:
            if vertical > 1 and any(repeat[y:y + vertical]):
                raise ValueError("A vertical merge cannot cross a repeating row.")
    table = document.add_table(rows=0, cols=columns)
    table.style = "Table Grid"
    table.autofit = False
    weights = [widths.get(x, 100) for x in range(columns)]
    page_width = document.sections[0].page_width - document.sections[0].left_margin - document.sections[0].right_margin
    for column, weight in zip(table.columns, weights):
        column.width = Emu(round(page_width * weight / sum(weights)))
    offset = {}
    for y, key in enumerate(repeat):
        row_scope = scope
        if key:
            iterable, variable = repeat_scope(key, collections, scope)
            table.add_row().cells[0].text = "{%tr for " + variable + " in " + iterable + " %}"
            row_scope = scope | {variable, "loop"}
        row = table.add_row()
        offset[y] = (len(table.rows) - 1, row_scope)
        if all(cell["type"] == "tableHeader" for _, _, _, cell in layout[y]) and layout[y]:
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        if key:
            table.add_row().cells[0].text = "{%tr endfor %}"
    for y, cells in enumerate(layout):
        index, row_scope = offset[y]
        for x, horizontal, vertical, node in cells:
            cell = table.cell(index, x)
            if horizontal > 1 or vertical > 1:
                cell = cell.merge(table.cell(offset[y + vertical - 1][0], x + horizontal - 1))
            for i, block in enumerate(node.get("content", ())):
                first = cell.paragraphs[0] if i == 0 else None
                if block["type"] in ("bulletList", "orderedList"):
                    add_list(cell, block, row_scope, paragraph=first)
                elif block["type"] in ("paragraph", "heading"):
                    add_paragraph(cell, block, row_scope, paragraph=first)
                else:
                    raise ValueError("Table cells support text, headings, and lists.")
            if node["type"] == "tableHeader":
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True


def repeat_scope(key, collections, scope):
    if key not in collections:
        raise ValueError("This collection is unavailable for the selected report.")
    variable = collections[key]
    expression(key, scope)
    return key, variable


def build_template(payload, base=None, *, embed=False):
    data = json.loads(payload)
    collection, item, _ = REPORTS[data["kind"]]
    collections = {collection: item}
    scope = {collection}
    if collection == "publications":
        collections.update({"groups": "group", "group.publications": "pub"})
        scope.add("groups")
    document = Document(base) if base else Document()
    if base:
        document.element.body.clear_content()
        remove_visual_metadata(document)
    stack = []
    for node in data["document"].get("content", ()):
        kind = node["type"]
        if kind == "control":
            tag, scope = control_tag(node["attrs"], scope, stack, collections)
            document.add_paragraph("{%p " + tag + " %}")
        elif kind == "table":
            add_table(document, node, scope, collections)
        elif kind in ("bulletList", "orderedList"):
            add_list(document, node, scope)
        elif kind in ("paragraph", "heading"):
            add_paragraph(document, node, scope)
        else:
            raise ValueError(f"Unsupported template element: {kind}.")
    if stack:
        raise ValueError("Close each repeat or condition with an End block.")
    if embed:
        Path("template.docx").write_bytes(save_visual_template(document, payload))
    else:
        output = BytesIO()
        document.save(output)
        Path("template.docx").write_bytes(output.getvalue())
