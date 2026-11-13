"""Check the Conferences worksheet for duplicate DOIs or titles."""

import argparse
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException
from workbook_cli import parse_workbook_args

from check_publications import check_workbook


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, help="History workbook (.xlsx)")
    args = parse_workbook_args(parser)
    try:
        duplicates = check_workbook(args.workbook, "Conferences")
    except KeyError:
        parser.exit(2, "Error: The Conferences worksheet was not found.\n")
    except (OSError, ValueError, BadZipFile, InvalidFileException) as error:
        parser.exit(2, f"Error: {error}\n")

    for name, groups in duplicates.items():
        if not groups:
            print(f"{name}: no duplicates.")
        for entries in groups:
            rows = ", ".join(str(row_number) for row_number, _ in entries)
            print(f"{name}: {entries[0][1]!r} (rows {rows})")
    return int(any(duplicates.values()))


if __name__ == "__main__":
    raise SystemExit(main())
