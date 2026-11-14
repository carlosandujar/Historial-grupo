"""Field and loop scopes for the visual editor and LLM prompt."""

import json

from report_context import REPORTS, field_alias, named_cells
from sample_workbook import create_sample_workbook


def editor_fields(kind, workbook=None):
    report = REPORTS[kind]
    fields = {}

    def add(name, label=None, rich=False):
        fields[f"{report.item}.{name}"] = {
            "label": label or name.replace("_", " ").capitalize(), "rich": rich,
        }

    source = workbook if workbook is not None and report.sheet in workbook.sheetnames else create_sample_workbook()
    headers = next(iter(source[report.sheet].values), ())
    for header in named_cells(headers, [None] * len(headers)):
        alias = field_alias(header)
        if alias.isidentifier() and alias not in {"raw_cell", "display_fields"} and (report.item != "pub" or alias != "excel"):
            fields.setdefault(f"{report.item}.{alias}", {"label": header, "rich": False})
        add("raw_cell[" + json.dumps(header, ensure_ascii=False) + "]", header + " (raw cell value)")
    for name in report.fields:
        if f"{report.item}.{name}" not in fields:
            add(name)
    for name in report.rich_fields:
        add(name, {"title_with_link": "Linked title", "doi_with_link": "Linked DOI", "link_text": "Project links"}[name], rich=True)
    fields.update({
        "field.label": {"label": "Display field label", "rich": False},
        "field.value": {"label": "Display field value", "rich": False},
        "field.rich": {"label": "Display field uses RichText", "rich": False},
    })
    fields["group.year"] = {"label": report.year_label.capitalize(), "rich": False}
    fields["loop.index"] = {"label": "Record number", "rich": False}
    fields[report.collection + " | length"] = {"label": "Number of records", "rich": False}
    return json.dumps({
        "collection": report.collection, "item": report.item, "fields": fields,
        "yearLabel": report.year_label,
    }, ensure_ascii=False)
