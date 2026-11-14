"""Shared worksheet fields and year groups for Word reports and the editor."""

from datetime import date, datetime
import re
from typing import NamedTuple


class Report(NamedTuple):
    collection: str
    item: str
    sheet: str
    year_field: str
    year_label: str
    fields: tuple[str, ...]
    rich_fields: tuple[str, ...] = ()


PUBLICATION_FIELDS = ("title", "authors", "citation", "doi_url")
PUBLICATION_LINKS = ("title_with_link", "doi_with_link")
THESIS_FIELDS = ("author", "title", "supervisors", "team")
PROJECT_FIELDS = (
    "reference", "acronym", "title", "funding_entity", "call", "ips", "ip_virvig_upc",
    "text_es", "team", "team_ip", "start", "end", "start_year", "active",
    "amount", "amount_display", "links", "has_links",
)
REPORTS = {
    "journals": Report("publications", "pub", "Journals", "year", "publication year",
                       PUBLICATION_FIELDS, PUBLICATION_LINKS),
    "conferences": Report("publications", "pub", "Conferences", "year", "publication year",
                          PUBLICATION_FIELDS, PUBLICATION_LINKS),
    "theses": Report("theses", "thesis", "PhD Theses", "year", "defense year", (*THESIS_FIELDS, "year")),
    "progress": Report("theses", "thesis", "PhD in progress", "enrollment_year", "enrollment year",
                       (*THESIS_FIELDS, "enrollment", "enrollment_year")),
    "projects": Report("projects", "project", "Projects-international", "start_year", "start year",
                       (*PROJECT_FIELDS, "eu"), ("link_text",)),
    "national": Report("projects", "project", "Projects-national", "start_year", "start year",
                       (*PROJECT_FIELDS, "aei"), ("link_text",)),
    "activities": Report("activities", "activity", "Activities", "year", "activity year",
                         ("title", "date", "year", "team", "description", "contacts", "activity_urls")),
}


def field_alias(header):
    return re.sub(r"\W+", "_", header.casefold()).strip("_")


def named_cells(headers, values):
    return {
        str(header): value for header, value in zip(headers, values)
        if header is not None and str(header).strip()
    }


def loop_collections(report):
    return {
        report.collection: report.item, "groups": "group", f"group.{report.collection}": report.item,
        f"{report.item}.display_fields": "field",
    }


def template_record(raw, **fields):
    record = {}
    for header, value in raw.items():
        alias = field_alias(header)
        if alias:
            record.setdefault(alias, value)
    return {**record, **fields, "raw_cell": raw}


def display_fields(report, entry):
    raw = entry.get("raw_cell", {})
    shown = {"title"}
    reserved = {"raw_cell", "display_fields"}
    computed = set(report.rich_fields)
    display_aliases = {}
    if report.item == "pub":
        reserved.update(("excel", "legacy_title", "doi_link"))
        computed.update(("citation", "doi_url"))
    if report.item == "thesis":
        computed.add("enrollment_year")
    if report.item == "project":
        reserved.update(("team_ip", "amount"))
        computed.update(("start_year", "amount_display", "links", "has_links"))
        display_aliases.update(amount="amount_display", ip_es_virvig_upc="ip_virvig_upc")
        if not any(field_alias(header) == "team" for header in raw):
            display_aliases["team_ip"] = "team"
    fields = []
    seen_headers = set()

    for header, value in raw.items():
        alias = field_alias(header)
        first = alias not in seen_headers
        seen_headers.add(alias)
        if alias == "title" and first:
            continue
        name = display_aliases.get(alias, alias)
        use_alias = alias not in computed | reserved or name != alias
        if first and use_alias and name in entry and name not in shown:
            value = entry[name]
            shown.add(name)
        fields.append({"label": header, "value": "" if value is None else value, "rich": False})

    for name in (*report.fields, *report.rich_fields):
        if name in shown or name in reserved or name not in entry:
            continue
        value = entry[name]
        if isinstance(value, list):
            value = ", ".join(map(str, value))
        fields.append({
            "label": name.replace("_", " ").capitalize(),
            "value": "" if value is None else value,
            "rich": name in report.rich_fields,
        })
    return fields


def date_year(value):
    if isinstance(value, date):
        return value.year
    if value is None:
        return None
    text = str(value).strip()
    if len(text) == 4 and text.isdecimal() and 1000 <= int(text) <= 9999:
        return int(text)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).year
        except ValueError:
            continue
    return None


def report_context(kind, entries):
    report = REPORTS[kind]
    entries = [{**entry, "display_fields": display_fields(report, entry)} for entry in entries]
    buckets = {}
    for entry in entries:
        year = date_year(entry.get(report.year_field))
        buckets.setdefault(year, []).append(entry)
    groups = [
        {"year": year, report.collection: buckets[year]}
        for year in sorted(buckets, key=lambda year: (year is None, -(year or 0)))
    ]
    return {report.collection: entries, "groups": groups}
