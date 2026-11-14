"""Exporta los proyectos internacionales del historial a una plantilla Word."""

import argparse
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlsplit
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate, RichText
from jinja2 import Environment, StrictUndefined, TemplateError
from openpyxl.utils.exceptions import InvalidFileException
from report_workbook import open_report_workbook
from workbook_cli import parse_workbook_args

from export_publications import TEMPLATE_DIR, cell_text, team_names
from participant_filter import matches_participants, participant_names


SHEET = "Projects-international"
TEMPLATE = TEMPLATE_DIR / "proyectos-internacionales" / "ficha.docx"
COMMON_TEXT_FIELDS = {
    "Reference": "reference", "Acronym": "acronym",
    "Title": "title", "Funding entity": "funding_entity", "Call": "call",
    "IPs": "ips", "Text - ES": "text_es",
}
TEXT_FIELDS = {**COMMON_TEXT_FIELDS, "EU?": "eu", "IP ViRVIG-UPC?": "ip_virvig_upc"}


def format_date(value):
    return value.strftime("%d/%m/%Y") if isinstance(value, date) else cell_text(value)


def date_year(value):
    if isinstance(value, date):
        return value.year
    text = cell_text(value)
    if len(text) == 4 and text.isdecimal() and 1000 <= int(text) <= 9999:
        return int(text)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).year
        except ValueError:
            continue
    return None


def read_projects(path, *, sheet=SHEET, text_fields=TEXT_FIELDS,
                  first_year=None, last_year=None, teams=None, participants=""):
    selected = participant_names(participants)
    with open_report_workbook(path) as workbook:
        if sheet not in workbook.sheetnames:
            raise ValueError(f"No se encontró la hoja {sheet}.")
        rows = workbook[sheet].iter_rows()
        headers = tuple(cell_text(cell.value).casefold() for cell in next(rows, ()))
        if any(headers.count(field) != 1 for field in ("reference", "title")):
            raise ValueError(f"{sheet} debe tener las columnas Reference y Title.")
        team_header = "team" if "team" in headers else "team ip"
        if teams and team_header not in headers:
            raise ValueError(f"{sheet} debe tener la columna Team o Team IP para filtrar por equipo.")
        if selected and "ips" not in headers:
            raise ValueError(f"{sheet} debe tener la columna IPs para filtrar por participantes.")
        if (first_year is not None and "end" not in headers) or (last_year is not None and "start" not in headers):
            raise ValueError(f"{sheet} debe tener las columnas de fechas necesarias para filtrar por años.")
        projects = []
        for cells in rows:
            row = dict(zip(headers, cells))
            values = {header: cell.value for header, cell in row.items()}
            project = {name: cell_text(values.get(header.casefold())) for header, name in text_fields.items()}
            if not (project["reference"] or project["title"]):
                continue
            team = cell_text(values.get(team_header))
            if teams and not teams & team_names(team):
                continue
            if not matches_participants(values.get("ips"), selected, ips=True):
                continue
            start, end = values.get("start"), values.get("end")
            if first_year is not None or last_year is not None:
                start_year, end_year = date_year(start), date_year(end)
                if start_year is not None and end_year is not None and start_year > end_year:
                    continue
                if first_year is not None and (end_year is None or end_year < first_year):
                    continue
                if last_year is not None and (start_year is None or start_year > last_year):
                    continue
            amount = values.get("amount")
            amount_display = cell_text(amount)
            if isinstance(amount, (int, float)) and not isinstance(amount, bool):
                amount_display = f"{amount:,.2f}".translate(str.maketrans(",.", ".,"))
                if "€" in (row["amount"].number_format or ""):
                    amount_display += " €"
            project.update(
                team_ip=team, start=format_date(start), end=format_date(end),
                active=values.get("active"), amount=amount, amount_display=amount_display,
                links=[url for header in ("link", "link 2", "link 3")
                       if (url := cell_text(values.get(header)))],
            )
            projects.append(project)
        return projects


def render(template, output, projects, *, jinja_env=None):
    document = DocxTemplate(template)
    entries = []
    for project in projects:
        links = RichText()
        count = 0
        for url in project["links"]:
            parsed = urlsplit(url)
            if parsed.scheme not in ("http", "https") or not parsed.netloc:
                continue
            if count:
                links.add(" · ")
            count += 1
            links.add(f"Enlace {count}", color="23578C", underline=True, url_id=document.build_url_id(url))
        entries.append({**project, "link_text": links, "has_links": count > 0})
    environment = jinja_env if jinja_env is not None else Environment(undefined=StrictUndefined, autoescape=True)
    document.render({"projects": entries}, jinja_env=environment)
    document.save(output)


def export_cli(description, sheet, template, output, text_fields):
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("workbook", nargs="?", type=Path, help=f"Historial XLSX con la hoja {sheet}")
    parser.add_argument("--teams", help="Equipos separados por comas, p. ej. UPC,UdG")
    parser.add_argument("--participants", default="", help="Apellidos de IPs separados por comas")
    parser.add_argument("--from-year", type=int, help="Incluye proyectos que terminan en este año o después")
    parser.add_argument("--to-year", type=int, help="Incluye proyectos que empiezan en este año o antes")
    parser.add_argument("--template", type=Path, default=template, help="Plantilla DOCX")
    parser.add_argument("--output", type=Path, default=Path(output), help="Word de salida")
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
        projects = read_projects(
            args.workbook, sheet=sheet, text_fields=text_fields,
            first_year=args.from_year, last_year=args.to_year, teams=teams, participants=args.participants,
        )
        render(args.template, args.output, projects)
    except (OSError, ValueError, BadZipFile, InvalidFileException, PackageNotFoundError, TemplateError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"{len(projects)} proyectos exportados a {args.output}")


def main():
    export_cli(__doc__, SHEET, TEMPLATE, "proyectos_internacionales.docx", TEXT_FIELDS)


if __name__ == "__main__":
    main()
