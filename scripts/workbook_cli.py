"""Resolve a workbook path for the command-line tools."""

from os import getenv
from pathlib import Path


WORKBOOK_ENV = "VIRVIG_WORKBOOK"


def parse_workbook_args(parser):
    args = parser.parse_args()
    if args.workbook is None:
        path = getenv(WORKBOOK_ENV)
        if not path or not path.strip():
            parser.error(f"Provide a workbook path on the command line or set {WORKBOOK_ENV}.")
        args.workbook = Path(path)
    return args
