"""Fictional records for previewing templates before a workbook is selected."""

from datetime import date

from openpyxl import Workbook


def create_sample_workbook():
    workbook = Workbook()
    workbook.remove(workbook.active)
    for year, topic in ((2026, "Terrain visualization"), (2025, "Virtual heritage"), (2025, "Crowd simulation")):
        title = f"Example: {topic}"
        author = "Rivera, Alex"
        publication = dict(
            Team="UPC", Authors=author, Title=title, Year=year, DOI="10.0000/example" if year == 2026 else "",
        )
        pages = {"Start page": 12, "End page": 24, "Impact factor": ""}
        thesis = dict(Team="UPC", Author=author, Title=title, Supervisors="Martin, Casey")
        project = {
            "Team": "UPC", "Title": title, "Reference": f"DEMO-{year}",
            "Acronym": "DEMO", "Funding entity": "Example funding agency", "Call": "Example call",
            "IPs": author, "Start": date(year, 1, 1), "End": date(year + 2, 12, 31),
            "Active": True, "Amount": 10000, "Text - ES": "Proyecto de ejemplo.",
            "Link": "https://example.org", "Link 2": "", "Link 3": "",
        }
        for sheet, record in (
            ("Journals", {**publication, **pages, "Journal title": "Example Journal", "Journal vol": 4,
                          "Quartile": "", "Position": "", "Area": "", "URL DOI": "", "Intranet": ""}),
            ("Conferences", {**publication, **pages, "Conference proceedings": "Example Conference",
                             "Organization entity": "", "Month": "", "International": "",
                             "URL DOI": "", "Internal comments": ""}),
            ("Activities", {"Team": "UPC", "Title": title, "Year": year, "Date": date(year, 3, 1), "Time": None,
                            "Contacts": author, "Export": "Y", "Description": "An example activity.",
                            "Activity URLs": "https://example.org", "Internal comments": ""}),
            ("PhD Theses", {**thesis, "Year": year, "Month": 3, "Day": 1, "Link": "https://example.org/thesis"}),
            ("PhD in progress", {**thesis, "Enrollment": date(year, 9, 1)}),
            ("Projects-international", {**project, "EU?": True, "IP ViRVIG-UPC?": True}),
            ("Projects-national", {**project, "AEI?": True, "IP és ViRVIG-UPC?": True}),
        ):
            if sheet not in workbook.sheetnames:
                workbook.create_sheet(sheet).append(list(record))
            workbook[sheet].append(list(record.values()))
    return workbook
