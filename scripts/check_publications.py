"""Checks for duplicate DOIs or titles in the Journals sheet."""

import argparse
from collections import defaultdict
from pathlib import Path
import re
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from workbook_cli import parse_workbook_args


COLUMNS = ("DOI", "Title")
DOI_PREFIX = re.compile(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)+", re.I)


def find_duplicates(rows):
    """Returns the duplicate values and their Excel rows by column."""
    headers = tuple(str(cell).strip() for cell in next(rows, ()))
    for name in COLUMNS:
        if headers.count(name) != 1:
            raise ValueError(f"Se esperaba exactamente una columna {name!r} en la primera fila.")
    columns = {name: headers.index(name) for name in COLUMNS}
    occurrences = {name: defaultdict(list) for name in COLUMNS}
    for row_number, row in enumerate(rows, start=2):
        for name, index in columns.items():
            value = row[index]
            if value is None:
                continue
            key = " ".join(str(value).split()).casefold()
            if name == "DOI":
                key = DOI_PREFIX.sub("", key).strip()
                if key == "?":
                    continue
            if key:
                occurrences[name][key].append((row_number, value))
    return {
        name: [entries for entries in groups.values() if len(entries) > 1]
        for name, groups in occurrences.items()
    }


def check_workbook(path, sheet_name):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return find_duplicates(workbook[sheet_name].values)
    finally:
        workbook.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help="Fichero XLSX del historial")
    args = parse_workbook_args(parser)
    try:
        duplicates = check_workbook(args.workbook, "Journals")
    except KeyError:
        parser.exit(2, "Error: No se encontró la hoja Journals.\n")
    except (OSError, ValueError, BadZipFile, InvalidFileException) as error:
        parser.exit(2, f"Error: {error}\n")

    for name, groups in duplicates.items():
        if not groups:
            print(f"{name}: sin duplicados.")
        for entries in groups:
            rows = ", ".join(str(row_number) for row_number, _ in entries)
            print(f"{name}: {entries[0][1]!r} (filas {rows})")
    return int(any(duplicates.values()))


if __name__ == "__main__":
    raise SystemExit(main())
