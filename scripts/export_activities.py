"""Export selected activities from the history workbook to Word."""

import argparse
from datetime import date, datetime
from pathlib import Path
import re
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined, TemplateError
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from workbook_cli import parse_workbook_args

from export_publications import TEMPLATE_DIR, cell_text, team_names
from participant_filter import matches_participants, participant_names


SHEET = "Activities"
TEMPLATE = TEMPLATE_DIR / "actividades" / "detallada.docx"
EXPORT_VALUES = {"y", "yes", "s", "verdadero"}


def read_activities(path, first_year=None, last_year=None, teams=(), participants=""):
    selected = participant_names(participants)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if SHEET not in workbook.sheetnames:
            raise ValueError(f"The workbook does not contain the {SHEET} worksheet.")
        rows = iter(workbook[SHEET].values)
        headers = tuple(cell_text(value) for value in next(rows, ()))
        required = ("Title", "Date", "Year", "Export")
        if teams:
            required += ("Team",)
        if selected:
            required += ("Contacts",)
        if any(headers.count(field) != 1 for field in required):
            raise ValueError(f"{SHEET} must contain one of each required column ({', '.join(required)}).")

        activities = []
        for values in rows:
            row = dict(zip(headers, values))
            flag = row["Export"]
            if flag is not True and cell_text(flag).casefold() not in EXPORT_VALUES:
                continue
            title = cell_text(row["Title"])
            if not title:
                continue
            raw_date = row["Date"]
            day = raw_date.date() if isinstance(raw_date, datetime) else raw_date
            raw_year = cell_text(row["Year"])
            year = day.year if isinstance(day, date) else (
                int(raw_year) if len(raw_year) == 4 and raw_year.isdecimal() else None
            )
            if first_year is not None and (year is None or year < first_year):
                continue
            if last_year is not None and (year is None or year > last_year):
                continue
            team = cell_text(row.get("Team"))
            if teams and not teams & team_names(team):
                continue
            contacts = cell_text(row.get("Contacts"))
            if not matches_participants(re.sub(r"\s+and\s+", "; ", contacts, flags=re.I), selected):
                continue

            activity = {
                "title": title,
                "date": day.strftime("%d/%m/%Y") if isinstance(day, date) else cell_text(day),
                "year": year,
                "team": team,
                "description": str(row.get("Description") or "").strip(),
                "contacts": contacts,
                "activity_urls": str(row.get("Activity URLs") or "").strip(),
            }
            activities.append((year or 0, day if isinstance(day, date) else date.min, activity))
        return [activity for _, _, activity in sorted(activities, key=lambda entry: entry[:2], reverse=True)]
    finally:
        workbook.close()


def render(template, output, activities, *, jinja_env=None):
    document = DocxTemplate(template)
    environment = jinja_env if jinja_env is not None else Environment(undefined=StrictUndefined)
    document.render({"activities": activities}, jinja_env=environment, autoescape=True)
    document.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help="History workbook with an Activities worksheet")
    parser.add_argument("--from-year", type=int, help="First year to include")
    parser.add_argument("--to-year", type=int, help="Last year to include")
    parser.add_argument("--teams", default="", help="Comma-separated teams")
    parser.add_argument("--participants", default="", help="Comma-separated contact surnames")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="DOCX template")
    parser.add_argument("--output", type=Path, default=Path("actividades.docx"), help="Output DOCX file")
    args = parse_workbook_args(parser)

    first = args.from_year if args.from_year is not None else 1000
    last = args.to_year if args.to_year is not None else 9999
    if not 1000 <= first <= last <= 9999:
        parser.error("Specify a valid range of four-digit years.")
    if args.output.suffix.lower() != ".docx":
        parser.error("The output file must have a .docx extension.")
    if args.template.resolve() == args.output.resolve():
        parser.error("The output file must differ from the template.")

    try:
        activities = read_activities(
            args.workbook, args.from_year, args.to_year, team_names(args.teams), args.participants,
        )
        render(args.template, args.output, activities)
    except (OSError, ValueError, BadZipFile, InvalidFileException, PackageNotFoundError, TemplateError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"Exported {len(activities)} activities to {args.output}")


if __name__ == "__main__":
    main()
