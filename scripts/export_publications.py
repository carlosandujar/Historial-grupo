"""Genera un Word con las publicaciones de Journals por equipo y año."""

import argparse
from datetime import date
from pathlib import Path
import re
from urllib.parse import quote
from zipfile import BadZipFile

from docxtpl import DocxTemplate, RichText
from jinja2 import Environment, StrictUndefined, TemplateError
from openpyxl.utils.exceptions import InvalidFileException
from report_workbook import open_report_workbook
from report_context import named_cells, report_context, template_record
from workbook_cli import parse_workbook_args
from participant_filter import matches_participants, participant_names


SCRIPT_DIR = Path(__file__).resolve().parent
# The web build copies the exporters beside its templates directory.
TEMPLATE_DIR = (SCRIPT_DIR.parent if SCRIPT_DIR.name == "scripts" else SCRIPT_DIR) / "templates"
TEMPLATE = TEMPLATE_DIR / "publicaciones" / "detallada.docx"
REQUIRED = ("Team", "Authors", "Title", "Journal title", "Year", "DOI")


class PublicationTemplate(DocxTemplate):
    def patch_xml(self, src_xml):
        xml = super().patch_xml(src_xml)
        # docxtpl deja las expresiones {{r ...}} entre runs. El título
        # normal permanece dentro de w:t y conserva su valor de texto.
        return re.sub(
            r"(?<=</w:r>)\{\{\s*(\w+)\.title\s*\}\}(?=<w:r>)",
            r"{{ \1.legacy_title }}",
            xml,
        )


def cell_text(value):
    if value is None or isinstance(value, date):
        return ""
    return " ".join(str(value).split())


def team_names(value):
    return {name.strip().casefold() for name in re.split(r"[,;]", value) if name.strip()}


def doi_url(row):
    for field in ("DOI", "URL DOI"):
        value = cell_text(row.get(field))
        value = re.split(r"Digital Object Identifier \(DOI\)", value, maxsplit=1, flags=re.I)[0]
        match = re.search(r"10\.\d{4,9}/[^\s<>\"']+", value)
        if match:
            return "https://doi.org/" + quote(match.group(), safe="/:;().-_%")
    return ""


def read_publications(
    path, first_year, last_year, teams, participants="", *, sheet_name="Journals",
    required=REQUIRED,
):
    selected = participant_names(participants)
    with open_report_workbook(path) as workbook:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"No se encontró la hoja {sheet_name}.")
        rows = workbook[sheet_name].values
        source_headers = next(rows, ())
        headers = tuple(cell_text(cell) for cell in source_headers)
        if any(headers.count(field) != 1 for field in required):
            raise ValueError(f"{sheet_name} debe tener una columna de cada tipo: {', '.join(required)}.")
        named_headers = [header for header in headers if header]
        if len(named_headers) != len(set(named_headers)):
            raise ValueError(f"{sheet_name} tiene encabezados repetidos. Cada columna con nombre debe ser única.")
        publications = []
        for values in rows:
            row = {header: value for header, value in zip(headers, values) if header}
            year = cell_text(row["Year"])
            title = cell_text(row["Title"])
            if not (len(year) == 4 and year.isdecimal() and title):
                continue
            year = int(year)
            if not first_year <= year <= last_year:
                continue
            if teams and not teams & team_names(cell_text(row["Team"])):
                continue
            if not matches_participants(row["Authors"], selected):
                continue
            publications.append(PublicationRow(row, named_cells(source_headers, values)))
        return sorted(publications, key=publication_year, reverse=True)


class PublicationRow(dict):
    def __init__(self, row, raw_cell):
        super().__init__(row)
        self.raw_cell = raw_cell


def publication_year(row):
    return int(cell_text(row["Year"]))


def citation(row, *, venue_field="Journal title"):
    venue = cell_text(row[venue_field])
    volume = cell_text(row.get("Journal vol")) if venue_field == "Journal title" else ""
    if volume and volume[0].isdigit():
        volume = f"vol. {volume}"
    start = cell_text(row.get("Start page"))
    end = cell_text(row.get("End page"))
    pages = ""
    if start and end and start != end:
        pages = f"pp. {start}–{end}"
    elif start or end:
        pages = f"p. {start or end}"
    return ", ".join(part for part in (venue, volume, pages, cell_text(row["Year"])) if part) + "."


def template_publication(document, row, *, venue_field="Journal title"):
    title = cell_text(row["Title"])
    url = doi_url(row)
    legacy_title = RichText(title, bold=True)
    title_with_link = legacy_title
    doi_with_link = RichText()
    if url:
        url_id = document.build_url_id(url)
        title_with_link = RichText(title, bold=True, color="23578C", underline=True, url_id=url_id)
        doi_with_link = RichText(url, color="23578C", underline=True, url_id=url_id)
    return template_record(
        getattr(row, "raw_cell", row), excel=row, title=title, authors=cell_text(row["Authors"]),
        citation=citation(row, venue_field=venue_field),
        doi_url=url, title_with_link=title_with_link, doi_with_link=doi_with_link,
        legacy_title=legacy_title, doi_link=doi_with_link,
    )


def render(template, output, publications, *, jinja_env=None, venue_field="Journal title", kind="journals"):
    document = PublicationTemplate(template)
    entries = [template_publication(document, row, venue_field=venue_field) for row in publications]
    environment = jinja_env if jinja_env is not None else Environment(undefined=StrictUndefined)
    document.render(report_context(kind, entries), jinja_env=environment, autoescape=True)
    document.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help="Historial XLSX con la hoja Journals")
    parser.add_argument("--from-year", type=int, required=True, help="Primer año incluido")
    parser.add_argument("--to-year", type=int, required=True, help="Último año incluido")
    parser.add_argument("--teams", required=True, help="Equipos separados por comas, p. ej. UPC,UdG")
    parser.add_argument("--participants", default="", help="Apellidos de autores separados por comas")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="Plantilla DOCX")
    parser.add_argument("--output", type=Path, default=Path("publicaciones.docx"), help="Word de salida")
    args = parse_workbook_args(parser)
    teams = team_names(args.teams)
    if not teams or not 1000 <= args.from_year <= args.to_year <= 9999:
        parser.error("Indica al menos un equipo y un intervalo válido de años.")
    if args.output.suffix.lower() != ".docx":
        parser.error("La salida debe tener extensión .docx.")
    if args.template.resolve() == args.output.resolve():
        parser.error("La salida debe ser distinta de la plantilla.")
    try:
        publications = read_publications(args.workbook, args.from_year, args.to_year, teams, args.participants)
        render(args.template, args.output, publications)
    except (OSError, ValueError, BadZipFile, InvalidFileException, TemplateError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"{len(publications)} publicaciones exportadas a {args.output}")


if __name__ == "__main__":
    main()
