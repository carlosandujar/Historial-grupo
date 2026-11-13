"""Añade artículos de ViRVIG a la hoja Journals de un libro de historial."""

import argparse
from collections import Counter
from pathlib import Path
from posixpath import normpath
import re
import ssl
from tempfile import NamedTemporaryFile
from urllib.request import urlopen
from zipfile import BadZipFile, ZipFile
from workbook_cli import parse_workbook_args

from lxml import etree, html
import truststore


URL = "https://www.virvig.eu/publications.php?year={year}&author=all"
HEADERS = (
    "Team", "Authors", "Title", "Journal title", "Journal vol", "Start page",
    "End page", "Year", "Impact factor", "Quartile", "Position", "Area",
    "DOI", "URL DOI",
)
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def clean(value):
    return " ".join(value.split())


def doi_from(value):
    value = re.sub(r"Digital Object Identifier \(DOI\).*", "", value, flags=re.I)
    match = re.search(r"10\.\d{4,9}/[^\s<>\"']+", value)
    return match.group() if match else None


def find_class(node, name, tag="div"):
    return node.find(f".//{tag}[@class='{name}']")


def decode_html(content):
    # El sitio declara UTF-8, pero algunas entradas incluyen bytes Windows-1252.
    decoded = content.decode("utf-8", errors="surrogateescape")
    return re.sub(
        r"[\udc80-\udcff]",
        lambda m: bytes([ord(m.group()) - 0xDC00]).decode("cp1252", errors="replace"),
        decoded,
    )


def parse_publications(content, year):
    listing = html.fromstring(decode_html(content)).get_element_by_id("publicationslist")
    section = None
    found_journals = False
    records = []
    for item in listing.xpath(".//div[@class='papertype' or @class='publicationitem']"):
        if item.get("class") == "papertype":
            section = clean(item.text_content()).casefold()
            found_journals |= section == "journals"
        elif section == "journals":
            record = parse_journal(item)
            if record and record["Year"] == year:
                records.append(record)
    if not found_journals:
        raise ValueError("La página no contiene la sección Journals esperada.")
    return records


def bibtex_fields(item):
    bib = find_class(item, "publicationbibtex")
    fields = {}
    if bib is not None:
        # El sitio separa los campos del BibTeX con <br>.
        bib_text = "\n".join(bib.itertext())
        for match in re.finditer(r"(?m)^\s*(\w+)\s*=\s*\{(.*)\}\s*,?\s*$", bib_text):
            fields[match[1].lower()] = clean(match[2]).replace(r"\&", "&")
    return fields


def parse_journal(item):
    title_node = find_class(item, "publicationtitle")
    title = clean(title_node.text_content()) if title_node is not None else ""
    if not title:
        return None
    authors_node = find_class(item, "publicationauthors")
    desc = find_class(item, "publicationdesc")
    fields = bibtex_fields(item)
    citation = clean(desc.text or "") if desc is not None else ""
    citation_parts = re.match(
        r"^(?P<journal>.+?)(?:,\s*Vol\.?\s*(?P<volume>[^,]+))?"
        r"(?:,\s*pp?\.?\s*(?P<pages>[^,]+))?,\s*(?P<year>\d{4})\.?$",
        citation,
        re.I,
    )
    visible = citation_parts.groupdict() if citation_parts else {}
    year = fields.get("year") or visible.get("year")
    volume = fields.get("volume") or visible.get("volume")
    number = fields.get("number")
    if volume and number:
        volume = f"{volume}({number})"
    elif volume and volume.isdecimal():
        volume = int(volume)

    # Un solo número en pages puede ser un identificador de artículo.
    pages = fields.get("pages") or visible.get("pages") or ""
    page_range = re.fullmatch(r"(\d+)\s*(?:-+|[–—])\s*(\d+)", pages)
    start, end = map(int, page_range.groups()) if page_range else (None, None)
    doi = doi_from(fields.get("doi", ""))
    if not doi and desc is not None:
        doi = next(
            (value for link in desc.iter("a")
             if (value := doi_from(link.get("href", "")))),
            None,
        )
    return {
        "Authors": clean(authors_node.text_content()) if authors_node is not None else None,
        "Title": title,
        "Journal title": (fields.get("journal") or visible.get("journal") or "").strip() or None,
        "Journal vol": volume,
        "Start page": start,
        "End page": end,
        "Year": int(year) if year and year.isdecimal() else None,
        "DOI": doi,
        "URL DOI": f"https://doi.org/{doi}" if doi else None,
    }


def worksheet_part(archive, sheet_name):
    workbook = etree.fromstring(archive.read("xl/workbook.xml"))
    sheet = workbook.find(f"{MAIN}sheets/{MAIN}sheet[@name='{sheet_name}']")
    if sheet is None:
        raise ValueError(f"No se encontró la hoja {sheet_name}.")
    rels = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = next(rel.get("Target") for rel in rels if rel.get("Id") == sheet.get(REL_ID))
    return target.lstrip("/") if target.startswith("/") else normpath("xl/" + target)


def journal_part(archive):
    return worksheet_part(archive, "Journals")


def replace_workbook_parts(path, source, updated_parts, initial_stat):
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        with ZipFile(temp_path, "w") as destination:
            destination.comment = source.comment
            for entry in source.infolist():
                contents = updated_parts.get(entry.filename)
                destination.writestr(entry, source.read(entry) if contents is None else contents)
        source.close()
        current_stat = path.stat()
        if (current_stat.st_mtime_ns, current_stat.st_size, current_stat.st_ino) != (
            initial_stat.st_mtime_ns, initial_stat.st_size, initial_stat.st_ino
        ):
            raise ValueError("El libro cambió durante la importación. Vuelve a ejecutar el script.")
        temp_path.replace(path)
    finally:
        temp_path.unlink(missing_ok=True)


def append_journals(path, records, *, dry_run=False):
    return append_worksheet_rows(
        path, records, sheet_name="Journals", headers=HEADERS,
        identity=("Title", "Year"), unique_field="DOI", dry_run=dry_run,
    )


def append_worksheet_rows(
    path, records, *, sheet_name, headers, identity, unique_field=None,
    reuse_empty_rows=False, dry_run=False,
):
    initial_stat = path.stat()
    with ZipFile(path) as source:
        part = worksheet_part(source, sheet_name)
        sheet = etree.fromstring(source.read(part))
        if sheet.find(f"{MAIN}tableParts") is not None or sheet.find(f"{MAIN}autoFilter") is not None:
            raise ValueError("La hoja contiene una tabla o filtro con un rango que debe ampliarse.")
        data = sheet.find(f"{MAIN}sheetData")
        shared = []
        if "xl/sharedStrings.xml" in source.namelist():
            strings = etree.fromstring(source.read("xl/sharedStrings.xml"))
            shared = ["".join(node.itertext(f"{MAIN}t")) for node in strings]

        def read_row(row):
            values = {}
            for cell in row:
                col = cell.get("r").rstrip("0123456789")
                value = cell.findtext(f"{MAIN}v", "")
                if cell.get("t") == "s":
                    value = shared[int(value)]
                elif cell.get("t") == "inlineStr":
                    value = "".join(cell.itertext(f"{MAIN}t"))
                values[col] = value
            return values

        columns = dict(zip(headers, "ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
        if not len(data) or tuple(read_row(data[0]).get(col) for col in columns.values()) != headers:
            raise ValueError(f"Los encabezados de {sheet_name} no coinciden con el formato esperado.")

        def normalized(value):
            return clean(str(value if value is not None else "")).casefold()

        def key_for(record):
            return tuple(normalized(record.get(field)) for field in identity)

        def unique_key(record):
            value = normalized(record.get(unique_field))
            return (doi_from(value) or value) if unique_field == "DOI" else value

        existing = []
        last_populated = int(data[0].get("r"))
        for row in data[1:]:
            values = read_row(row)
            if any(values.values()) or row.find(f"{MAIN}c/{MAIN}f") is not None:
                last_populated = int(row.get("r"))
            existing.append({field: values.get(col) for field, col in columns.items()})
        keys = {key_for(record) for record in existing}
        unique_values = {unique_key(record) for record in existing}
        styles = {col: Counter() for col in columns.values()}
        for row in data[1:]:
            for cell in row:
                col = cell.get("r").rstrip("0123456789")
                if col in styles and cell.get("s") is not None:
                    styles[col][cell.get("s")] += 1
        row_number = last_populated if reuse_empty_rows else int(data[-1].get("r"))
        formatted_rows = {int(row.get("r")): row for row in data if int(row.get("r")) > row_number}
        count = 0
        for record in records:
            unique_value = unique_key(record)
            key = key_for(record)
            if (unique_value and unique_value in unique_values) or key in keys:
                continue
            unique_values.add(unique_value)
            keys.add(key)
            row_number += 1
            count += 1
            row = formatted_rows.get(row_number)
            if row is None:
                row = etree.Element(f"{MAIN}row", r=str(row_number))
                following = next((item for index, item in formatted_rows.items() if index > row_number), None)
                if following is None:
                    data.append(row)
                else:
                    following.addprevious(row)
            cells = {cell.get("r"): cell for cell in row}
            for field, col in columns.items():
                value = record.get(field)
                if value is None or value == "":
                    continue
                address = f"{col}{row_number}"
                cell = cells.get(address)
                if cell is None:
                    cell = etree.Element(f"{MAIN}c", r=address)
                    if styles[col]:
                        cell.set("s", styles[col].most_common(1)[0][0])
                    for other in row:
                        other_col = other.get("r").rstrip("0123456789")
                        if (len(other_col), other_col) > (len(col), col):
                            other.addprevious(cell)
                            break
                    else:
                        row.append(cell)
                else:
                    cell.attrib.pop("t", None)
                    for child in list(cell):
                        if child.tag in (f"{MAIN}v", f"{MAIN}is"):
                            cell.remove(child)
                if isinstance(value, int):
                    etree.SubElement(cell, f"{MAIN}v").text = str(value)
                else:
                    cell.set("t", "inlineStr")
                    inline = etree.SubElement(cell, f"{MAIN}is")
                    etree.SubElement(inline, f"{MAIN}t").text = value
        if dry_run or not count:
            return count

        dimension = sheet.find(f"{MAIN}dimension")
        if dimension is not None:
            start_ref, _, end_ref = dimension.get("ref").partition(":")
            last_col = end_ref.rstrip("0123456789") or next(reversed(columns.values()))
            last_row = max(row_number, int(re.search(r"\d+$", end_ref or start_ref).group()))
            dimension.set("ref", f"{start_ref}:{last_col}{last_row}")
        updated = etree.tostring(sheet, encoding="UTF-8", xml_declaration=True, standalone=True)
        replace_workbook_parts(path, source, {part: updated}, initial_stat)
    return count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("year", type=int, help="Año de publicación")
    parser.add_argument("--workbook", type=Path, help="Libro XLSX de destino")
    parser.add_argument("--dry-run", action="store_true", help="Cuenta las filas nuevas sin guardar")
    args = parse_workbook_args(parser)
    if not 1000 <= args.year <= 9999:
        parser.error("Indica un año de cuatro dígitos.")
    try:
        context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with urlopen(URL.format(year=args.year), timeout=30, context=context) as response:
            records = parse_publications(response.read(), args.year)
        count = append_journals(args.workbook, records, dry_run=args.dry_run)
    except (OSError, ValueError, BadZipFile, etree.LxmlError) as error:
        parser.exit(1, f"Error: {error}\n")
    action = "por añadir" if args.dry_run else "añadidas"
    print(f"{len(records)} publicaciones de {args.year} en Journals. Filas {action}: {count}.")


if __name__ == "__main__":
    main()
