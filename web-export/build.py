"""Prepare the exporters, templates, and template catalog for the web app."""

import json
from pathlib import Path
from shutil import copy2
from urllib.parse import quote


web = Path(__file__).resolve().parent
root = web.parent
dist = web / "dist"
files = (
    "workbook_cli.py",
    "participant_filter.py",
    "export_publications.py",
    "export_conferences.py",
    "export_activities.py",
    "export_phd_theses.py",
    "export_phd_in_progress.py",
    "export_projects_international.py",
    "export_projects_national.py",
)
groups = {
    "journals": ("publicaciones", "detallada.docx"),
    "conferences": ("conferencias", "detallada.docx"),
    "activities": ("actividades", "detallada.docx"),
    "theses": ("tesis-defendidas", "numeradas.docx"),
    "progress": ("tesis-en-curso", "numeradas.docx"),
    "projects": ("proyectos-internacionales", "ficha.docx"),
    "national": ("proyectos-nacionales", "ficha.docx"),
}
english_labels = {"detallada": "Detailed", "compacta": "Compact", "numeradas": "Numbered", "ficha": "Project profile"}

dist.mkdir(parents=True, exist_ok=True)
for name in files:
    copy2(root / "scripts" / name, dist / name)

catalog = {}
for kind, (folder, default) in groups.items():
    relative = Path("templates") / folder
    target = dist / relative
    target.mkdir(parents=True, exist_ok=True)
    templates = sorted(
        (file for file in (root / relative).glob("*")
         if file.is_file() and file.suffix.casefold() == ".docx" and not file.name.startswith(("~$", "."))),
        key=lambda file: (file.name.casefold() != default, file.name.casefold(), file.name),
    )
    catalog[kind] = []
    for file in templates:
        copy2(file, target / file.name)
        label = " ".join(file.stem.replace("-", " ").replace("_", " ").split())
        catalog[kind].append({
            "url": quote((relative / file.name).as_posix()),
            "label": english_labels.get(file.stem.casefold(), label[:1].upper() + label[1:] or file.name),
        })

(dist / "templates" / "catalogo.json").write_text(
    json.dumps(catalog, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
)

print(f"Web app ready in {dist}")
