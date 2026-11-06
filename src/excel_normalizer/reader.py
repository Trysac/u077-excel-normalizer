"""Bounded reading of ``.xlsx`` sheets into a :class:`~excel_normalizer.models.Grid`.

Before openpyxl runs, the file size is checked, the file is opened with :mod:`zipfile`, more
than 10_000 members are refused and the declared ``ZipInfo.file_size`` of every member is summed
against ``max_uncompressed_bytes``. Workbooks load with ``data_only=True`` and
``read_only=False`` (read-only sheets lack ``merged_cells``); formulas are never evaluated.
"""

from __future__ import annotations

import datetime
import importlib
import logging
import os
import zipfile
from typing import Any

from .errors import NormalizeError
from .models import CellValue, Grid

logger = logging.getLogger(__name__)

MAX_ZIP_MEMBERS = 10_000

# openpyxl ships no type information; importing it dynamically keeps ``mypy --strict`` clean.
_openpyxl: Any = importlib.import_module("openpyxl")


def _clip(text: str, limit: int) -> str:
    return text[:limit]


def _check_archive(path: str, max_uncompressed_bytes: int) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
    except (zipfile.BadZipFile, OSError, ValueError) as exc:
        raise NormalizeError("invalid_file", "file is not a readable zip archive") from exc
    if len(infos) > MAX_ZIP_MEMBERS:
        raise NormalizeError(
            "too_large", f"archive has {len(infos)} members; the limit is {MAX_ZIP_MEMBERS}"
        )
    declared = sum(info.file_size for info in infos)
    if declared > max_uncompressed_bytes:
        raise NormalizeError(
            "too_large",
            f"archive declares {declared} uncompressed bytes; "
            f"the limit is {max_uncompressed_bytes}",
        )


def _select_sheet(workbook: Any, sheet: str | int) -> Any:
    sheets = list(workbook.worksheets)
    titles = [str(item.title) for item in sheets]
    if isinstance(sheet, str):
        if sheet in titles:
            return sheets[titles.index(sheet)]
        wanted = repr(_clip(sheet, 50))
    else:
        if 0 <= sheet < len(sheets):
            return sheets[sheet]
        wanted = f"index {sheet}"
    available = ", ".join(_clip(title, 30) for title in titles[:10])
    raise NormalizeError("sheet_not_found", f"sheet {wanted} not found; available: {available}")


def _coerce(value: object) -> CellValue:
    if value is None:
        return None
    if isinstance(value, str | int | float | datetime.date):
        return value
    return str(value)


def _extract(worksheet: Any, max_cells: int) -> Grid:
    declared = int(worksheet.max_row) * int(worksheet.max_column)
    if declared > max_cells:
        raise NormalizeError("too_large", f"sheet has {declared} cells; the limit is {max_cells}")
    rows = tuple(
        tuple(_coerce(value) for value in row) for row in worksheet.iter_rows(values_only=True)
    )
    merged = tuple(
        sorted(
            (int(r.min_row) - 1, int(r.min_col) - 1, int(r.max_row) - 1, int(r.max_col) - 1)
            for r in worksheet.merged_cells.ranges
        )
    )
    return Grid(cells=rows, merged=merged)


def read_sheet(
    path: str | os.PathLike[str],
    sheet: str | int = 0,
    *,
    max_file_bytes: int = 50_000_000,
    max_uncompressed_bytes: int = 200_000_000,
    max_cells: int = 1_000_000,
) -> Grid:
    """Read one sheet of an ``.xlsx`` file into a :class:`Grid`.

    Cached formula results are read (``data_only=True``); a formula without a cached value
    reads as ``None`` and no formula is evaluated. Values that are not str, int, float, bool,
    date, datetime or ``None`` are converted with ``str()``.

    Args:
        path: Path of the workbook.
        sheet: Sheet title or zero-based index.
        max_file_bytes: Maximum size of the file on disk.
        max_uncompressed_bytes: Maximum sum of the declared sizes of the zip members.
        max_cells: Maximum ``max_row * max_column`` of the selected sheet.

    Raises:
        NormalizeError: ``too_large``, ``sheet_not_found`` or ``invalid_file``.
    """
    location = os.fspath(path)
    try:
        size = os.path.getsize(location)
    except OSError as exc:
        raise NormalizeError("invalid_file", "file cannot be read") from exc
    if size > max_file_bytes:
        raise NormalizeError("too_large", f"file has {size} bytes; the limit is {max_file_bytes}")
    _check_archive(location, max_uncompressed_bytes)
    try:
        workbook = _openpyxl.load_workbook(location, data_only=True, read_only=False)
    except Exception as exc:
        raise NormalizeError(
            "invalid_file", f"workbook cannot be parsed ({type(exc).__name__})"
        ) from exc
    try:
        grid = _extract(_select_sheet(workbook, sheet), max_cells)
    finally:
        workbook.close()
    logger.debug("read sheet with %d rows", len(grid.cells))
    return grid
