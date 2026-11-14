"""Read reports from a workbook path or a workbook already open in the browser."""

from contextlib import contextmanager

from openpyxl import Workbook, load_workbook


@contextmanager
def open_report_workbook(source):
    if isinstance(source, Workbook):
        yield source
        return
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        yield workbook
    finally:
        workbook.close()
