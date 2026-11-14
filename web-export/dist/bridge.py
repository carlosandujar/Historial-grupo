"""Load the workbook, select records, and render reports in Pyodide."""

import json
from io import BytesIO
from pathlib import Path

from jinja2 import StrictUndefined
from jinja2.sandbox import SandboxedEnvironment
from openpyxl import load_workbook

import export_publications as journals
import export_conferences as conferences
import export_activities as activities
import export_phd_theses as theses
import export_phd_in_progress as in_progress
import export_projects_international as projects
import export_projects_national as national
from sample_workbook import create_sample_workbook
from template_docx import build_template
from template_fields import editor_fields
from template_import import import_visual_template


workbook = None
sample_workbook = None
selection = None
selected_rows = None


def template_fields(kind):
    return editor_fields(kind, workbook)


def inspect_template(kind):
    source = Path("template-source.docx").read_bytes()
    draft = import_visual_template(source, kind)
    build_template(json.dumps(draft), "template-source.docx")
    # Check the final document syntax, including inline Word tags.
    from docxtpl import DocxTemplate
    DocxTemplate("template.docx").get_undeclared_template_variables()
    return json.dumps(draft)


def load_history():
    global workbook, selection, selected_rows
    if workbook is not None:
        workbook.close()
    workbook = None
    selection = selected_rows = None
    workbook = load_workbook(BytesIO(Path("input.xlsx").read_bytes()), read_only=True, data_only=True)
    return json.dumps(workbook.sheetnames)


def select_rows(source, kind, first_year, last_year, teams, participants):
    selected = journals.team_names(teams)
    if kind in ("journals", "conferences"):
        reader = journals.read_publications if kind == "journals" else conferences.read_conferences
        return reader(source, first_year or 1000, last_year or 9999, selected, participants)
    if kind == "activities":
        return activities.read_activities(source, first_year, last_year, selected, participants)
    if kind == "theses":
        return theses.read_theses(source, first_year, last_year, selected, participants)
    if kind == "progress":
        return in_progress.read_theses(source, selected, participants)
    options = dict(first_year=first_year, last_year=last_year, teams=selected, participants=participants)
    if kind == "national":
        options.update(sheet=national.SHEET, text_fields=national.TEXT_FIELDS)
    return projects.read_projects(source, **options)


renderers = {
    "journals": journals.render, "conferences": conferences.render, "activities": activities.render,
    "theses": theses.render, "progress": in_progress.render, "projects": projects.render,
    "national": projects.render,
}


def generate(kind, first_year, last_year, teams, participants="", *, preview=False, sample=False):
    global sample_workbook, selection, selected_rows
    if kind not in renderers:
        raise ValueError("Unknown report type.")
    if sample:
        if sample_workbook is None:
            sample_workbook = create_sample_workbook()
        source = sample_workbook
        first_year, last_year, teams, participants = None, None, "", ""
    else:
        source = workbook
        if source is None:
            raise ValueError("Choose a workbook before generating a report.")
    key = (sample, kind, first_year, last_year, teams, participants)
    if key != selection:
        selected_rows = select_rows(source, kind, first_year, last_year, teams, participants)
        selection = key
    rows = selected_rows[:5] if preview else selected_rows
    env = SandboxedEnvironment(undefined=StrictUndefined, autoescape=True)
    renderers[kind](Path("template.docx"), Path("output.docx"), rows, jinja_env=env)
    return json.dumps({"count": len(selected_rows), "shown": len(rows)})
