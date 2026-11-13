"""Exporta los proyectos nacionales del historial a una plantilla Word."""

from export_projects_international import COMMON_TEXT_FIELDS, export_cli
from export_publications import TEMPLATE_DIR


SHEET = "Projects-national"
TEMPLATE = TEMPLATE_DIR / "proyectos-nacionales" / "ficha.docx"
TEXT_FIELDS = {**COMMON_TEXT_FIELDS, "AEI?": "aei", "IP és ViRVIG-UPC?": "ip_virvig_upc"}


def main():
    export_cli(__doc__, SHEET, TEMPLATE, "proyectos_nacionales.docx", TEXT_FIELDS)


if __name__ == "__main__":
    main()
