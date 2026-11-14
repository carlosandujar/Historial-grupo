"""Export conference publications to Word by team, year, and author surname."""

import argparse
from pathlib import Path
from zipfile import BadZipFile

from jinja2 import TemplateError
from openpyxl.utils.exceptions import InvalidFileException
from workbook_cli import parse_workbook_args

import export_publications


TEMPLATE = export_publications.TEMPLATE_DIR / "conferencias" / "detallada.docx"
VENUE = "Conference proceedings"
REQUIRED = ("Team", "Authors", "Title", VENUE, "Year", "DOI")


def read_conferences(path, first_year, last_year, teams, participants=""):
    return export_publications.read_publications(
        path, first_year, last_year, teams, participants,
        sheet_name="Conferences", required=REQUIRED,
    )


def render(template, output, publications, *, jinja_env=None):
    export_publications.render(
        template, output, publications, jinja_env=jinja_env, venue_field=VENUE, kind="conferences",
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help="History workbook with a Conferences worksheet")
    parser.add_argument("--from-year", type=int, required=True, help="First year to include")
    parser.add_argument("--to-year", type=int, required=True, help="Last year to include")
    parser.add_argument("--teams", required=True, help="Comma-separated teams, e.g. UPC,UdG")
    parser.add_argument("--participants", default="", help="Comma-separated author surnames")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="DOCX template")
    parser.add_argument("--output", type=Path, default=Path("conferencias.docx"), help="Output DOCX file")
    args = parse_workbook_args(parser)

    teams = export_publications.team_names(args.teams)
    if not teams or not 1000 <= args.from_year <= args.to_year <= 9999:
        parser.error("Specify at least one team and a valid year range.")
    if args.output.suffix.lower() != ".docx":
        parser.error("The output file must have a .docx extension.")
    if args.template.resolve() == args.output.resolve():
        parser.error("The output file must differ from the template.")

    try:
        publications = read_conferences(args.workbook, args.from_year, args.to_year, teams, args.participants)
        render(args.template, args.output, publications)
    except (OSError, ValueError, BadZipFile, InvalidFileException, TemplateError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"Exported {len(publications)} conference publications to {args.output}")


if __name__ == "__main__":
    main()
