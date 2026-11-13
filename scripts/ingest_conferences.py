"""Import conference publications from ViRVIG into the Conferences worksheet."""

import argparse
from pathlib import Path
import re
import ssl
from urllib.request import urlopen
from zipfile import BadZipFile
from workbook_cli import parse_workbook_args

from lxml import etree, html
import truststore

from ingest_publications import (
    URL, append_worksheet_rows, bibtex_fields, clean, decode_html, doi_from, find_class,
)


SHEET = "Conferences"
HEADERS = (
    "Team", "Authors", "Title", "Conference proceedings", "Organization entity",
    "Year", "Month", "Start page", "End page", "International", "Impact factor",
    "DOI", "Internal comments",
)


def parse_conferences(content, year):
    listing = html.fromstring(decode_html(content)).get_element_by_id("publicationslist")
    if not listing.xpath("./div[@class='papertype'][normalize-space()='Conferences']"):
        raise ValueError("The listing does not contain a Conferences section.")
    items = listing.xpath(
        "./div[@class='publicationitem']"
        "[preceding-sibling::div[@class='papertype'][1][normalize-space()='Conferences']]"
    )
    records = []
    for item in items:
        fields = bibtex_fields(item)
        title = find_class(item, "publicationtitle")
        title_text = clean(title.text_content()) if title is not None else fields.get("title")
        if not title_text:
            continue
        authors = find_class(item, "publicationauthors")
        author_text = (
            clean(authors.text_content()) if authors is not None
            else re.sub(r"\s+and\s+", "; ", fields.get("author", ""))
        )

        description = find_class(item, "publicationdesc")
        citation = clean(description.text or "") if description is not None else ""
        match = re.fullmatch(
            r"(?P<venue>.+?)(?:,\s*pp?\.?\s*(?P<pages>[^,]+))?,\s*(?P<year>\d{4})\.?",
            citation, re.IGNORECASE,
        )
        visible = match.groupdict() if match else {}
        publication_year = fields.get("year") or visible.get("year")
        if publication_year != str(year):
            continue

        pages = fields.get("pages") or visible.get("pages") or ""
        span = re.fullmatch(r"(\d+(?::\d+)?)\s*(?:-+|[–—])\s*(\d+(?::\d+)?)", pages)
        start, end = (None, None)
        if span:
            start, end = (int(page) if page.isdecimal() else page for page in span.groups())

        doi = doi_from(fields.get("doi", ""))
        if not doi and description is not None:
            doi = next(
                (value for link in description.iter("a")
                 if (value := doi_from(link.get("href", "")))),
                None,
            )
        records.append({
            "Authors": author_text,
            "Title": title_text,
            "Conference proceedings": visible.get("venue") or fields.get("booktitle"),
            "Year": year,
            "Start page": start,
            "End page": end,
            "DOI": doi,
        })
    return records


def append_conferences(path, records, *, dry_run=False):
    return append_worksheet_rows(
        path, records, sheet_name=SHEET, headers=HEADERS, identity=("Title", "Year"),
        unique_field="DOI", reuse_empty_rows=True, dry_run=dry_run,
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("year", type=int, help="Publication year")
    parser.add_argument("--workbook", type=Path, help="Destination XLSX workbook")
    parser.add_argument("--dry-run", action="store_true", help="Count new rows without saving")
    args = parse_workbook_args(parser)
    if not 1000 <= args.year <= 9999:
        parser.error("Enter a four-digit year.")
    try:
        context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with urlopen(URL.format(year=args.year), timeout=30, context=context) as response:
            records = parse_conferences(response.read(), args.year)
        count = append_conferences(args.workbook, records, dry_run=args.dry_run)
    except (OSError, ValueError, BadZipFile, etree.LxmlError) as error:
        parser.exit(1, f"Error: {error}\n")
    action = "to add" if args.dry_run else "added"
    print(f"{len(records)} conference publications for {args.year}. Rows {action}: {count}.")


if __name__ == "__main__":
    main()
