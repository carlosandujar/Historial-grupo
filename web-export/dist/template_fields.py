"""Field and loop scopes exposed by the existing report exporters."""

import json
import re

from sample_workbook import create_sample_workbook


REPORTS = {
    "journals": ("publications", "pub", "Journals"),
    "conferences": ("publications", "pub", "Conferences"),
    "activities": ("activities", "activity", "Activities"),
    "theses": ("theses", "thesis", "PhD Theses"),
    "progress": ("theses", "thesis", "PhD in progress"),
    "projects": ("projects", "project", "Projects-international"),
    "national": ("projects", "project", "Projects-national"),
}
PROJECT_FIELDS = (
    "reference acronym title funding_entity call ips ip_virvig_upc text_es team_ip start end "
    "active amount amount_display links has_links"
)


def editor_fields(kind, workbook=None):
    collection, item, sheet = REPORTS[kind]
    fields = {}

    def add(name, label=None, rich=False):
        fields[f"{item}.{name}"] = {"label": label or name.replace("_", " ").capitalize(), "rich": rich}

    if item == "pub":
        source = workbook if workbook is not None and sheet in workbook.sheetnames else create_sample_workbook()
        headers = next(iter(source[sheet].values), ())
        for header in headers:
            header = " ".join(str(header or "").split())
            if header:
                alias = re.sub(r"\W+", "_", header.casefold()).strip("_")
                if alias and alias.isidentifier():
                    add(alias, header)
                add("excel[" + json.dumps(header, ensure_ascii=False) + "]", header + " (Excel)")
        for field in ("title", "authors", "citation", "doi_url"):
            add(field)
        add("title_with_link", "Linked title", True)
        add("doi_with_link", "Linked DOI", True)
        fields["group.year"] = {"label": "Group year", "rich": False}
    else:
        names = {
            "activity": "title date year team description contacts activity_urls",
            "thesis": "author title supervisors team " + ("year" if kind == "theses" else "enrollment"),
            "project": PROJECT_FIELDS + (" eu" if kind == "projects" else " aei"),
        }[item]
        for name in names.split():
            add(name)
        if item == "project":
            add("link_text", "Project links", True)
    fields["loop.index"] = {"label": "Record number", "rich": False}
    fields[collection + " | length"] = {"label": "Number of records", "rich": False}
    return json.dumps({"collection": collection, "item": item, "fields": fields}, ensure_ascii=False)
