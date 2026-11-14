"""Exporta las tesis en curso del historial a una plantilla Word."""

import argparse
from datetime import date
from pathlib import Path
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from jinja2 import TemplateError
from openpyxl.utils.exceptions import InvalidFileException
from report_workbook import open_report_workbook
from report_context import date_year, named_cells, template_record
from workbook_cli import parse_workbook_args

from export_phd_theses import render as render_theses
from export_publications import TEMPLATE_DIR, cell_text, team_names
from participant_filter import matches_participants, participant_names


SHEET = "PhD in progress"
TEMPLATE = TEMPLATE_DIR / "tesis-en-curso" / "numeradas.docx"
FIELDS = {"Team": "team", "Author": "author", "Title": "title", "Supervisors": "supervisors"}


def read_theses(path, teams=None, participants=""):
    selected = participant_names(participants)
    with open_report_workbook(path) as workbook:
        if SHEET not in workbook.sheetnames:
            raise ValueError(f"No se encontró la hoja {SHEET}.")
        rows = workbook[SHEET].values
        source_headers = next(rows, ())
        headers = tuple(cell_text(value) for value in source_headers)
        if any(headers.count(field) != 1 for field in ("Author", "Title", "Supervisors")):
            raise ValueError(f"{SHEET} debe tener las columnas Author, Title y Supervisors.")
        if teams and headers.count("Team") != 1:
            raise ValueError(f"{SHEET} debe tener la columna Team para filtrar por equipo.")
        theses = []
        for values in rows:
            row = dict(zip(headers, values))
            thesis = {name: cell_text(row.get(header)) for header, name in FIELDS.items()}
            if not thesis["title"]:
                continue
            if teams and not teams & team_names(thesis["team"]):
                continue
            if not matches_participants(row["Supervisors"], selected):
                continue
            enrollment = row.get("Enrollment")
            thesis["enrollment"] = (
                enrollment.strftime("%d/%m/%Y") if isinstance(enrollment, date) else cell_text(enrollment)
            )
            theses.append(template_record(named_cells(source_headers, values), **thesis, enrollment_year=date_year(enrollment)))
        return theses


def render(template, output, theses, *, jinja_env=None):
    render_theses(template, output, theses, jinja_env=jinja_env, kind="progress")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help=f"Historial XLSX con la hoja {SHEET}")
    parser.add_argument("--teams", help="Equipos separados por comas, p. ej. UPC,UdG")
    parser.add_argument("--participants", default="", help="Apellidos de directores separados por comas")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="Plantilla DOCX")
    parser.add_argument("--output", type=Path, default=Path("tesis_en_curso.docx"), help="Word de salida")
    args = parse_workbook_args(parser)
    teams = team_names(args.teams or "")
    if args.teams is not None and not teams:
        parser.error("Indica al menos un equipo en --teams.")
    if args.output.suffix.lower() != ".docx":
        parser.error("La salida debe tener extensión .docx.")
    if args.template.resolve() == args.output.resolve():
        parser.error("La salida debe ser distinta de la plantilla.")

    try:
        theses = read_theses(args.workbook, teams, args.participants)
        render(args.template, args.output, theses)
    except (OSError, ValueError, BadZipFile, InvalidFileException, PackageNotFoundError, TemplateError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"{len(theses)} tesis en curso exportadas a {args.output}")


if __name__ == "__main__":
    main()
