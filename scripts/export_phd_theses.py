"""Exporta las tesis del historial a Word con una plantilla DOCX."""

import argparse
from pathlib import Path
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined, TemplateError
from openpyxl.utils.exceptions import InvalidFileException
from report_workbook import open_report_workbook
from workbook_cli import parse_workbook_args

from export_publications import TEMPLATE_DIR, cell_text, team_names
from participant_filter import matches_participants, participant_names


SHEET = "PhD Theses"
TEMPLATE = TEMPLATE_DIR / "tesis-defendidas" / "numeradas.docx"
FIELDS = {"Author": "author", "Title": "title", "Supervisors": "supervisors", "Year": "year"}


def read_theses(path, first_year, last_year, teams, participants=""):
    selected = participant_names(participants)
    with open_report_workbook(path) as workbook:
        if SHEET not in workbook.sheetnames:
            raise ValueError(f"No se encontró la hoja {SHEET}.")
        rows = workbook[SHEET].values
        headers = tuple(cell_text(value) for value in next(rows, ()))
        required = (*FIELDS, "Team") if teams else FIELDS
        if any(headers.count(field) != 1 for field in required):
            raise ValueError(f"{SHEET} debe tener una columna de cada tipo: {', '.join(required)}.")

        theses = []
        for values in rows:
            row = dict(zip(headers, values))
            thesis = {field: cell_text(row.get(header)) for header, field in FIELDS.items()}
            if not thesis["title"]:
                continue
            raw_year = thesis["year"]
            year = int(raw_year) if len(raw_year) == 4 and raw_year.isdecimal() else None
            if first_year is not None and (year is None or year < first_year):
                continue
            if last_year is not None and (year is None or year > last_year):
                continue
            thesis["team"] = cell_text(row.get("Team"))
            if teams and not teams & team_names(thesis["team"]):
                continue
            if not matches_participants(row["Supervisors"], selected):
                continue
            thesis["year"] = year
            theses.append(thesis)
        return sorted(theses, key=lambda thesis: thesis["year"] or 0, reverse=True)


def render(template, output, theses, *, jinja_env=None):
    document = DocxTemplate(template)
    environment = jinja_env if jinja_env is not None else Environment(undefined=StrictUndefined, autoescape=True)
    document.render({"theses": theses}, jinja_env=environment)
    document.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help="Historial XLSX con la hoja PhD Theses")
    parser.add_argument("--from-year", type=int, help="Primer año incluido")
    parser.add_argument("--to-year", type=int, help="Último año incluido")
    parser.add_argument("--teams", help="Equipos separados por comas. Por defecto, incluye todas las tesis.")
    parser.add_argument("--participants", default="", help="Apellidos de directores separados por comas")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="Plantilla DOCX")
    parser.add_argument("--output", type=Path, default=Path("tesis_doctorales.docx"), help="Word de salida")
    args = parse_workbook_args(parser)

    first = args.from_year if args.from_year is not None else 1000
    last = args.to_year if args.to_year is not None else 9999
    if not 1000 <= first <= last <= 9999:
        parser.error("Indica un intervalo válido de años de cuatro dígitos.")
    teams = team_names(args.teams or "")
    if args.teams is not None and not teams:
        parser.error("Indica al menos un equipo en --teams.")
    if args.output.suffix.lower() != ".docx":
        parser.error("La salida debe tener extensión .docx.")
    if args.template.resolve() == args.output.resolve():
        parser.error("La salida debe ser distinta de la plantilla.")

    try:
        theses = read_theses(args.workbook, args.from_year, args.to_year, teams, args.participants)
        render(args.template, args.output, theses)
    except (OSError, ValueError, BadZipFile, InvalidFileException, PackageNotFoundError, TemplateError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"{len(theses)} tesis exportadas a {args.output}")


if __name__ == "__main__":
    main()
