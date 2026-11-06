"""Immutable data model: grids read from sheets, header specs and normalized tables."""

from __future__ import annotations

import csv
import datetime
import io
from dataclasses import dataclass

CellValue = str | int | float | bool | datetime.date | datetime.datetime | None

_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


@dataclass(frozen=True)
class Grid:
    """Raw cell values of a sheet.

    Attributes:
        cells: Rows of cell values; rows may have different lengths.
        merged: Merged ranges as ``(min_row, min_col, max_row, max_col)``, zero-based inclusive.
    """

    cells: tuple[tuple[CellValue, ...], ...]
    merged: tuple[tuple[int, int, int, int], ...] = ()


@dataclass(frozen=True)
class HeaderSpec:
    """Location of the header rows inside a :class:`Grid` (zero-based ``start_row``)."""

    start_row: int
    row_count: int


def _render(value: CellValue, escape_formulas: bool) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, str):
        if escape_formulas and value.startswith(_FORMULA_PREFIXES):
            return "'" + value
        return value
    return str(value)


@dataclass(frozen=True)
class Table:
    """A clean table with unique column names."""

    columns: tuple[str, ...]
    rows: tuple[tuple[CellValue, ...], ...]

    def to_csv(self, *, escape_formulas: bool = True) -> str:
        """Render the table as CSV text using the ``csv`` module's default CRLF terminator.

        Cells render as: ``None`` -> empty, ``bool`` -> ``true``/``false``, dates and datetimes
        -> ISO 8601, everything else -> ``str()``. With ``escape_formulas`` (the default) a
        ``str`` cell, header names included, that starts with ``=``, ``+``, ``-``, ``@``, tab or
        carriage return gets a leading single quote (OWASP CSV Injection). Numbers are never
        escaped.
        """
        buffer = io.StringIO()
        writer = csv.writer(buffer)
        writer.writerow([_render(name, escape_formulas) for name in self.columns])
        for row in self.rows:
            writer.writerow([_render(value, escape_formulas) for value in row])
        return buffer.getvalue()
