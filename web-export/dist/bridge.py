"""Run the shared exporters in Pyodide."""

from pathlib import Path
from jinja2 import StrictUndefined
from jinja2.sandbox import SandboxedEnvironment

import export_publications as journals
import export_conferences as conferences
import export_activities as activities
import export_phd_theses as theses
import export_phd_in_progress as in_progress
import export_projects_international as projects
import export_projects_national as national


def generate(kind, first_year, last_year, teams, participants=""):
    workbook = Path("input.xlsx")
    template = Path("template.docx")
    output = Path("output.docx")
    env = SandboxedEnvironment(undefined=StrictUndefined, autoescape=True)
    selected = journals.team_names(teams)
    if kind == "journals":
        rows = journals.read_publications(workbook, first_year or 1000, last_year or 9999, selected, participants)
        journals.render(template, output, rows, jinja_env=env)
    elif kind == "conferences":
        rows = conferences.read_conferences(workbook, first_year or 1000, last_year or 9999, selected, participants)
        conferences.render(template, output, rows, jinja_env=env)
    elif kind == "activities":
        rows = activities.read_activities(workbook, first_year, last_year, selected, participants)
        activities.render(template, output, rows, jinja_env=env)
    elif kind == "theses":
        rows = theses.read_theses(workbook, first_year, last_year, selected, participants)
        theses.render(template, output, rows, jinja_env=env)
    elif kind == "progress":
        rows = in_progress.read_theses(workbook, selected, participants)
        in_progress.render(template, output, rows, jinja_env=env)
    elif kind == "projects":
        rows = projects.read_projects(
            workbook, first_year=first_year, last_year=last_year, teams=selected, participants=participants,
        )
        projects.render(template, output, rows, jinja_env=env)
    elif kind == "national":
        rows = projects.read_projects(
            workbook, sheet=national.SHEET, text_fields=national.TEXT_FIELDS,
            first_year=first_year, last_year=last_year, teams=selected, participants=participants,
        )
        projects.render(template, output, rows, jinja_env=env)
    else:
        raise ValueError("Unknown report type.")
    return len(rows)
