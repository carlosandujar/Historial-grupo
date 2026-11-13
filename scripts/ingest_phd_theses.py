"""Añade las tesis doctorales de ViRVIG a la hoja PhD Theses."""

import argparse
from pathlib import Path
import re
import ssl
from urllib.parse import urljoin, urlsplit
from urllib.request import urlopen
from zipfile import BadZipFile
from workbook_cli import parse_workbook_args

from lxml import etree, html
import truststore

from ingest_publications import append_worksheet_rows, clean, decode_html, find_class


URL = "https://www.virvig.eu/publicationsphd.php"
SHEET = "PhD Theses"
HEADERS = ("Team", "Author", "Year", "Month", "Day", "Title", "Supervisors", "Link")


def parse_theses(content):
    listing = html.fromstring(decode_html(content)).get_element_by_id("phdlist")
    items = listing.xpath(
        "./div[@class='phditem']"
        "[preceding-sibling::div[@class='papertype'][1][normalize-space()='PhD Thesis']]"
    )
    if not items:
        raise ValueError("No se encontraron fichas en la sección PhD Thesis.")

    records = []
    for item in items:
        title = find_class(item, "phdtitle")
        author = find_class(item, "phdauthor")
        advisors = find_class(item, "phdadvidors")
        desc = find_class(item, "phddesc")
        bib = find_class(item, "publicationbibtex")
        title_text = clean(title.text_content()) if title is not None else ""
        if not title_text:
            continue
        date_text = clean(desc.text or "") if desc is not None else ""
        visible_date = re.fullmatch(
            r"PhD\.\s*Thesis\.\s*(?:(?P<month>[A-Za-z]+)\s+)?(?P<year>\d{4})\.?",
            date_text, flags=re.IGNORECASE,
        )
        date_fields = visible_date.groupdict() if visible_date else {}
        if bib is not None:
            for field, value in re.findall(
                r"(?m)^\s*(year|month)\s*=\s*\{([^{}]*)\}", "\n".join(bib.itertext()), re.I
            ):
                if value.strip():
                    date_fields[field.lower()] = value.strip()
        year = date_fields.get("year")
        supervisor_text = clean(advisors.text_content()) if advisors is not None else ""
        links = []
        for block in (title, desc):
            if block is None:
                continue
            for href in block.xpath(".//a/@href"):
                if not href.strip() or href.strip().startswith("#"):
                    continue
                link = urljoin(URL, href.strip())
                parsed = urlsplit(link)
                if parsed.scheme in ("http", "https") and parsed.netloc and link not in links:
                    links.append(link)
        records.append({
            "Author": clean(author.text_content()) if author is not None else None,
            "Year": int(year) if year and re.fullmatch(r"\d{4}", year) else None,
            "Month": date_fields.get("month"),
            "Title": title_text,
            "Supervisors": re.sub(r"^Advisors?\s*:\s*", "", supervisor_text, flags=re.I),
            "Link 1": links[0] if links else None,
            "Link 2": links[1] if len(links) > 1 else None,
        })
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("year", type=int, nargs="?", help="Año a importar. Por defecto, todos.")
    parser.add_argument("--workbook", type=Path, help="Libro XLSX de destino")
    parser.add_argument("--dry-run", action="store_true", help="Cuenta las filas nuevas sin guardar")
    args = parse_workbook_args(parser)
    if args.year is not None and not 1000 <= args.year <= 9999:
        parser.error("Indica un año de cuatro dígitos.")

    try:
        context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with urlopen(URL, timeout=30, context=context) as response:
            records = parse_theses(response.read())
        if args.year is not None:
            records = [record for record in records if record["Year"] == args.year]
        count = append_worksheet_rows(
            args.workbook, records, sheet_name=SHEET, headers=HEADERS,
            identity=("Title", "Author", "Year"), dry_run=args.dry_run,
        )
    except (OSError, ValueError, BadZipFile, etree.LxmlError) as error:
        parser.exit(1, f"Error: {error}\n")
    action = "por añadir" if args.dry_run else "añadidas"
    print(f"{len(records)} tesis en {SHEET}. Filas {action}: {count}.")


if __name__ == "__main__":
    main()
