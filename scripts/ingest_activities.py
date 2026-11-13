"""Import ViRVIG activities into the Activities worksheet."""

import argparse
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import re
import ssl
from urllib.parse import urljoin, urlsplit
from urllib.request import urlopen
from zipfile import BadZipFile, ZipFile

from lxml import etree, html
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.datetime import to_excel
from openpyxl.utils.exceptions import InvalidFileException
from workbook_cli import parse_workbook_args
import truststore

from ingest_publications import MAIN, clean, decode_html, find_class, replace_workbook_parts, worksheet_part


URL = "https://www.virvig.eu/activities.php"
SHEET = "Activities"
HEADERS = (
    "Team", "Title", "Date", "Time", "Year", "Description", "Contacts",
    "Activity URLs", "Internal comments", "Export",
)
WEB_URL = re.compile(r'\b(?:https?://|www\.)[^\s<>"\']+', re.I)


def description_text(node):
    if node is None:
        return ""
    node = deepcopy(node)
    for element in node.iter("br", "p", "li", "div"):
        element.tail = "\n" + (element.tail or "")
    return "\n".join(line for raw in node.text_content().splitlines() if (line := clean(raw)))


def activity_urls(*nodes):
    urls = {}
    for node in nodes:
        if node is None:
            continue
        for href in node.xpath(".//a/@href"):
            href = href.strip()
            if href and not href.startswith("#"):
                link = urljoin(URL, href)
                parsed = urlsplit(link)
                if parsed.scheme in ("http", "https") and parsed.netloc:
                    urls[link] = None
        for link in WEB_URL.findall(description_text(node)):
            link = link.rstrip(".,;!?")
            while link.endswith(")") and link.count(")") > link.count("("):
                link = link[:-1]
            urls["https://" + link if link.lower().startswith("www.") else link] = None
    return "\n".join(urls)


def parse_activities(content):
    page = html.fromstring(decode_html(content))
    items = page.find_class("activityconference")
    if not items:
        raise ValueError("The page does not contain the expected activity entries.")
    records = []
    for item in items:
        title_node = find_class(item, "activityconferencetitle")
        when_node = find_class(item, "activityconferencetime")
        description = find_class(item, "activityconferencedesc")
        contact = find_class(item, "activityconferencecontact")
        title = clean(title_node.text_content()) if title_node is not None else ""
        if not title or when_node is None:
            raise ValueError("An activity is missing its title or date and time.")
        when = datetime.strptime(clean(when_node.text_content()), "%d-%m-%Y %H:%M")
        contacts = clean(contact.text_content()) if contact is not None else ""
        records.append({
            "Title": title, "Date": when.date(), "Time": when.time(), "Year": when.year,
            "Description": description_text(description),
            "Contacts": re.sub(r"^Contacts?\s*:\s*", "", contacts, flags=re.I),
            "Activity URLs": activity_urls(title_node, description),
        })
    return records


def activity_key(title, date, time):
    if isinstance(date, datetime):
        date = date.date()
    if isinstance(time, datetime):
        time = time.time()
    return clean(str(title or "")).casefold(), str(date), str(time)


def write_cell(row, column, value, *, style=None):
    address = f"{get_column_letter(column)}{row.get('r')}"
    cell = row.find(f"{MAIN}c[@r='{address}']")
    if cell is None:
        cell = etree.SubElement(row, f"{MAIN}c", r=address)
    if style is not None:
        cell.set("s", str(style))
    if isinstance(value, str):
        cell.set("t", "inlineStr")
        inline = etree.SubElement(cell, f"{MAIN}is")
        etree.SubElement(inline, f"{MAIN}t").text = value
    else:
        etree.SubElement(cell, f"{MAIN}v").text = str(value)


def number_styles(styles):
    formats = styles.find(f"{MAIN}cellXfs")
    original_count = len(formats)
    result = {}
    for field, number_format in (("Date", "14"), ("Time", "20")):
        for index, style in enumerate(formats):
            if style.get("numFmtId") == number_format:
                result[field] = index
                break
        else:
            result[field] = len(formats)
            etree.SubElement(formats, f"{MAIN}xf", numFmtId=number_format, fontId="0",
                             fillId="0", borderId="0", xfId="0", applyNumberFormat="1")
    if len(formats) != original_count:
        formats.set("count", str(len(formats)))
    return result, len(formats) != original_count


def worksheet_rows(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if SHEET not in workbook.sheetnames:
            raise ValueError("The workbook does not contain an Activities worksheet.")
        rows = iter(workbook[SHEET].values)
        headers = tuple(next(rows, ()))
        columns = {name: index for index, name in enumerate(headers, start=1) if name}
        if not columns:
            if any(any(value is not None for value in row) for row in rows):
                raise ValueError("Activities has data below a blank header row.")
            return columns, set(), workbook.epoch
        if any(headers.count(name) != 1 for name in HEADERS[:-1]) or headers.count("Export") > 1:
            raise ValueError("Activities headers do not match the expected format.")
        indexes = [columns[name] - 1 for name in ("Title", "Date", "Time")]
        existing = {activity_key(*(row[index] for index in indexes)) for row in rows}
        return columns, existing, workbook.epoch
    finally:
        workbook.close()


def append_activities(path, records, *, dry_run=False):
    initial_stat = path.stat()
    columns, keys, epoch = worksheet_rows(path)
    pending = []
    for record in records:
        key = activity_key(record["Title"], record["Date"], record["Time"])
        if key not in keys:
            pending.append(record)
            keys.add(key)
    new_headers = [name for name in HEADERS if name not in columns]
    if dry_run or not (pending or new_headers):
        return len(pending)

    with ZipFile(path) as source:
        part = worksheet_part(source, SHEET)
        sheet = etree.fromstring(source.read(part))
        if sheet.find(f"{MAIN}tableParts") is not None or sheet.find(f"{MAIN}autoFilter") is not None:
            raise ValueError("Activities contains a table or filter whose range would need updating.")
        data = sheet.find(f"{MAIN}sheetData")
        dimension = sheet.find(f"{MAIN}dimension")
        _, _, last_column, last_row = range_boundaries(dimension.get("ref")) if dimension is not None else (1, 1, 1, 1)
        if new_headers:
            header = data.find(f"{MAIN}row[@r='1']")
            if header is None:
                header = etree.Element(f"{MAIN}row", r="1")
                data.insert(0, header)
            column = max(max(columns.values(), default=0), last_column) if columns else 0
            for name in new_headers:
                column += 1
                columns[name] = column
                write_cell(header, column, name)

        styles_part = "xl/styles.xml"
        styles = etree.fromstring(source.read(styles_part))
        date_styles, changed_styles = number_styles(styles)
        row_number = int(data[-1].get("r"))
        for record in pending:
            row_number += 1
            row = etree.SubElement(data, f"{MAIN}row", r=str(row_number))
            for name, column in sorted(columns.items(), key=lambda item: item[1]):
                value = record.get(name)
                if value is None or value == "":
                    continue
                style = date_styles.get(name)
                write_cell(row, column, to_excel(value, epoch) if style is not None else value, style=style)
        if dimension is not None:
            end = get_column_letter(max(last_column, max(columns.values())))
            dimension.set("ref", f"A1:{end}{max(last_row, row_number)}")
        parts = {part: etree.tostring(sheet, encoding="UTF-8", xml_declaration=True, standalone=True)}
        if changed_styles and pending:
            parts[styles_part] = etree.tostring(styles, encoding="UTF-8", xml_declaration=True, standalone=True)
        replace_workbook_parts(path, source, parts, initial_stat)
    return len(pending)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("year", nargs="?", type=int, help="Activity year. Defaults to all years.")
    parser.add_argument("--workbook", type=Path, help="Destination XLSX workbook")
    parser.add_argument("--dry-run", action="store_true", help="Count new rows without saving")
    args = parse_workbook_args(parser)
    if args.year is not None and not 1000 <= args.year <= 9999:
        parser.error("Enter a four-digit year.")

    try:
        context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with urlopen(URL, timeout=30, context=context) as response:
            records = parse_activities(response.read())
        if args.year is not None:
            records = [record for record in records if record["Year"] == args.year]
        count = append_activities(args.workbook, records, dry_run=args.dry_run)
    except (OSError, ValueError, BadZipFile, InvalidFileException, etree.LxmlError) as error:
        parser.exit(1, f"Error: {error}\n")
    action = "to add" if args.dry_run else "added"
    print(f"{len(records)} activities found. Rows {action}: {count}.")


if __name__ == "__main__":
    main()
