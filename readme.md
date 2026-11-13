# ViRVIG group history management

## Introduction

### Motivation

We often need to report on the group's research outputs and activities, including publications, theses, and projects. Each call (PID, SGR, FPU...) requires a specific set of information, for a specific time period, and in a specific format. Some calls ask for a group-wide report, while others ask us to highlight the contributions of specific members.

The database behind the [ViRVIG website](https://www.virvig.eu/) contains most of the information we need, but it has some limitations:
- The website focuses on the founding groups of ViRVIG (UPC and UdG), but the SGR group has members from other institutions (UB, UdL), and their contributions are not recorded.
- The website records the activity of the UPC members who have belonged to the group for years, but the new members might not have their past activities recorded.
- Some categories of information are missing, such as ongoing theses.
- For those categories that are present, some specific information might be missing or incomplete. For example, some calls require the quartile of the journal, but the website does not provide it. 
- The website's intranet is limited in terms of exporting the data in a configurable format. 

### Purpose

This repository contains a collection of scripts that:
- help maintain the group history workbook (as an Excel document)
- help preserve the quality of the data
- generate configurable (template-based) reports in Word format, with filters for year, team, and participants. 

Some key information is imported from the [ViRVIG website](https://www.virvig.eu/), while other information is entered manually. Currently, the scripts can import:
- Journal and conference publications
- Defended theses
- Activities (*events*)

Projects and theses in progress are entered manually, along with any information the importers leave blank, such as impact indicators in publications. 

### Workbook structure
You can find the latest version of the workbook in the [shared folder in Google Drive][drive-folder].

Workbooks are organized by date, e.g., **Historial 2026-10-21.xlsx**.

Each workbook has one sheet for each report type, with the following names:
- **Journals** and **Conferences** hold publications
- **Activities** holds events 
- **PhD Theses** and **PhD in progress** track defended and ongoing theses
- **Projects-international** and **Projects-national** track funded projects. 

**Important**: Do not modify the existing sheet names and column headers, as they are used by the scripts to identify and process the data. 

## Generating reports

### Generate Word Reports via the web app

The easiest way for ViRVIG members to generate reports is to use the [web-export app][web-export]. 

Here's a potential workflow for generating a report:

- Step 1: Download a local copy of the shared Excel file from the [shared folder in Google Drive][drive-folder]. 

- Step 2 (optional): The app includes some filters, but you can always remove manually, **in your local copy of the excel file**, the rows that you do not want to include in the report. 

- Step 3: Open the [web-export app][web-export] in a browser

- Step 4: Select your local copy of the workbook

- Step 5: Choose the report type (e.g. Journal publications, defended theses, etc.)

  For an Activities report, mark the rows you want to include in the workbook's `Export` column with `Y`, `Yes`, `S`, or `Verdadero`. Excel's Boolean `TRUE` also qualifies.

- Step 6: Set the filters (time period, participants...)

- Step 7: Choose a DOCX template. The tool provides a few built-in templates, but you can also use your own custom template.

- Step 8: Click the "Generate Word report" button. The report will be generated (in your browser) and you can download it as a DOCX file.



### Generate Word reports via Python scripts

The repository contains a set of Python scripts to generate reports from the workbook. Each script corresponds to a report type, and they share a common interface.

Each exporter reads the workbook and fills a DOCX template. Use `--template` to select a template and `--output` to name the report.

#### Python setup

Run the Python commands from the repository root. Install the dependencies with `python -m pip install lxml truststore openpyxl docxtpl "docxcompose>=2,<3"` and set `$workbook` to the path of your local copy.

```powershell
$workbook = "../../Historial del Grup/Historial 2026-11.xlsx"
```

You can also set `$env:VIRVIG_WORKBOOK = $workbook` in PowerShell. With that set, omit the workbook argument from export and check commands and omit `--workbook` from import commands. An explicit CLI path takes precedence. Every command requires one of these two sources.

#### Export examples

```powershell
python scripts/export_publications.py $workbook --from-year 2025 --to-year 2026 --teams "UPC,UdG" --output publicaciones.docx
python scripts/export_conferences.py $workbook --from-year 2024 --to-year 2025 --teams "UPC,UdG" --template templates/conferencias/compacta.docx
python scripts/export_phd_theses.py $workbook --from-year 2024 --to-year 2025
python scripts/export_phd_in_progress.py $workbook --teams "UPC,UdG"
python scripts/export_projects_international.py $workbook --teams "UPC,UdG" --from-year 2022 --to-year 2025
python scripts/export_projects_national.py $workbook --from-year 2025
python scripts/export_activities.py $workbook --from-year 2024 --to-year 2025 --output actividades.docx
```

The following filters are available for the exporters:

- **Year range.** Both bounds are inclusive. Journals and conferences require both bounds in the CLI. Defended theses, projects, and activities accept either bound independently. Projects qualify if their duration overlaps the selected years. Theses in progress have no year filter.
- **Teams.** Use comma-separated names, such as `UPC,UdG,UB`. Matching ignores case. The journal and conference CLIs require at least one team.
- **Participants.** Use `--participants "Andujar, Pelechano, Vazquez"` to match at least one publication author, thesis supervisor, project principal investigator (PI), or activity contact. Matching ignores case and accents and accepts either part of a hyphenated surname.
- **Activities.** Only rows marked `Y`, `Yes`, `S`, or `Verdadero` in `Export` are included. Matching ignores case and surrounding spaces. Excel's Boolean `TRUE` also qualifies. 

### Jinja2 in Word templates

The exporters use [Jinja2](https://jinja.palletsprojects.com/) to fill DOCX templates. The context for each report type is described in the table below.

Jinja2 provides a rich set of features for template rendering. Use `{{ item.field }}` to insert a value, `{% if item.field %}` to conditionally display text, and `{%p for item in items %}` to repeat paragraphs. In a DOCX template, put each `{%p ... %}` tag in its own Word paragraph and keep each tag together when applying formatting. Use `{{r ... }}` for rich-text fields such as hyperlinks.

For example, enter these as three separate Word paragraphs in a publication template.

```jinja
{%p for pub in publications %}
{{ pub.authors }}. {{ pub.title }}. {{ pub.citation }}
{%p endfor %}
```

### Template fields by report

| Report | Default template under `templates/` | Jinja2 context |
| --- | --- | --- |
| Journal publications | `publicaciones/detallada.docx` | `publications` and `groups`. Each publication exposes all named columns in `Journals`, including `journal_title`, `journal_vol`, `quartile`, and `intranet`. |
| Conference publications | `conferencias/detallada.docx` | `publications` and `groups`. Each publication exposes all named columns in `Conferences`, including `conference_proceedings`, `organization_entity`, `international`, and `internal_comments`. |
| Defended theses | `tesis-defendidas/numeradas.docx` | `theses` with `author`, `title`, `supervisors`, `year`, and `team`. |
| Theses in progress | `tesis-en-curso/numeradas.docx` | `theses` with `author`, `title`, `supervisors`, `enrollment`, and `team`. |
| International projects | `proyectos-internacionales/ficha.docx` | `projects` with the shared project fields below, plus `eu`. |
| National projects | `proyectos-nacionales/ficha.docx` | `projects` with the shared project fields below, plus `aei`. |
| Activities | `actividades/detallada.docx` | `activities` with `title`, `date`, `year`, `team`, `description`, `contacts`, and `activity_urls`. |

**Publication fields.** Both publication (journal and conference) exporters provide lowercase aliases with underscores for Excel column headers. For the original value, use `pub.excel['Column header']`. Each entry also provides `citation`, `doi_url`, `title_with_link`, and `doi_with_link`. For example, render a linked title with `{{r pub.title_with_link }}`. Each entry in `groups` contains `year` and `publications`. The publication lists are ordered newest year first, and both report types offer a `compacta.docx` template.

**Thesis fields.** Both thesis exporters provide a list named `theses`. Each entry has `author`, `title`, `supervisors`, and `team`. Defended theses also expose `year`, while theses in progress expose `enrollment`. Use `{{ theses | length }}` to show the number of theses in the report.

This numbered example works with either thesis exporter. Put each line in a separate Word paragraph.

```jinja
{%p for thesis in theses %}
[{{ loop.index }}] {{ thesis.author }} — {{ thesis.title }}{% if thesis.supervisors %} ({{ thesis.supervisors }}){% endif %}
{%p endfor %}
```

**Project fields.** `reference`, `acronym`, `title`, `funding_entity`, `call`, `ips`, `ip_virvig_upc`, `text_es`, `team_ip`, `start`, `end`, `active`, `amount`, `amount_display`, `links`, `has_links`, and `link_text`. Render clickable project links with `{{r project.link_text }}`.




## Updating the workbook with ViRVIG website data

The repository contains scripts to import publications, theses, and activities from the ViRVIG website. 

This task is usually performed by the head of the group, so the rest of the members can skip this section. 

### Running the importers

Follow the [Python setup](#python-setup) above and set `$workbook` to the file you want to update.

The importers are designed to **append new rows** to their destination sheets. They skip records they recognize as duplicates and preserve existing rows and manual edits. Missing source fields are left blank. 

Note: If you have the workbook open in Excel, close it before running the importers (otherwise the scripts will fail to write to it).

Add `--dry-run` to preview the number of rows to add. Use `--workbook` or set `VIRVIG_WORKBOOK` to select the destination.

| Importer | Destination sheet | Fields written (when available) |
| --- | --- | --- |
| `ingest_publications.py` | `Journals` | Authors, Title, Journal title, Journal vol (including the issue when available), Start page, End page, Year, DOI, URL DOI |
| `ingest_conferences.py` | `Conferences` | Authors, Title, Conference proceedings, Year, Start page, End page, DOI |
| `ingest_phd_theses.py` | `PhD Theses` | Author, Title, Supervisors, Year, Month. The parser collects up to two links, but the current importer does not write them to the sheet's Link column. |
| `ingest_activities.py` | `Activities` | Title, Date, Time, Year (from Date), Description, Contacts, Activity URLs. Fill in Export manually to select activities for reports. |

Journal and conference imports require a year. For defended theses and activities, omit the year to import all available years.

```powershell
python scripts/ingest_publications.py 2026 --workbook $workbook
python scripts/ingest_conferences.py 2024 --workbook $workbook --dry-run
python scripts/ingest_phd_theses.py 2025 --workbook $workbook --dry-run
python scripts/ingest_activities.py --workbook $workbook --dry-run
```

### Validating the workbook

The repository contains scripts to check for repeated rows. In particular, they check for repeated DOIs and titles in the `Journals` and `Conferences` sheets.

```powershell
python scripts/check_publications.py $workbook
python scripts/check_conferences.py $workbook
```

The checkers report the Excel row numbers for repeated DOIs and titles in `Journals` and `Conferences`, respectively. They compare titles across all years, ignore blank values and `?` DOI placeholders, and recognize DOI resolver URLs. An exit code of 1 means duplicates were found. The checkers open the workbook in read-only mode.

[web-export]: http://www.cs.upc.edu/~virtual/VIRVIG/historial/
[drive-folder]: https://drive.google.com/drive/folders/17PWZjTBQ6FDlCk0JEbg4rYrnN-Cv4PkK?usp=drive_link

## Updating the web-export app

Run `python build.py` from the `web-export` directory after changing an exporter in `scripts/` or a built-in template in `templates/`. The build copies the Python modules and templates into `web-export/dist/` and regenerates `dist/templates/catalogo.json`. The web interface files, including `index.html`, `style.css`, and `app.js`, are edited directly in `web-export/dist/`.

To preview the updated app locally, run the following from `web-export`.

```powershell
python build.py
python -m http.server 8000 --bind 127.0.0.1 --directory dist
```

Open [127.0.0.1:8000](http://127.0.0.1:8000), select a workbook, and generate a test report. The browser processes the workbook and template locally. It needs internet access on the first export to load Pyodide and its dependencies from jsDelivr and PyPI.

To publish the update, upload the **contents** of `web-export/dist/` to the static host so that `index.html` sits at the app's public URL. Preserve the folder structure and upload the app files and built-in templates only. Users select their own workbook in the browser.

