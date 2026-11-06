"""Header detection and flattening of multi-row headers and merged cells."""

from __future__ import annotations

import logging

from .errors import NormalizeError
from .models import CellValue, Grid, HeaderSpec, Table

logger = logging.getLogger(__name__)

DEFAULT_MAX_CELLS = 1_000_000


def _is_blank(value: CellValue) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _filled(row: tuple[CellValue, ...]) -> int:
    return sum(1 for value in row if not _is_blank(value))


def _is_header_row(row: tuple[CellValue, ...]) -> bool:
    filled = 0
    strings = 0
    for value in row:
        if _is_blank(value):
            continue
        filled += 1
        if isinstance(value, str):
            strings += 1
    return filled > 0 and strings * 5 >= filled * 4


def detect_header(grid: Grid, *, max_header_rows: int = 5, max_title_rows: int = 5) -> HeaderSpec:
    """Locate the header rows of ``grid``.

    Up to ``max_title_rows`` leading rows with at most one non-empty cell are skipped as titles.
    Then consecutive rows in which at least 80 percent of the non-empty cells are strings are
    header rows, up to ``max_header_rows``.

    Raises:
        ValueError: if a limit argument is out of range.
        NormalizeError: with code ``header_not_found`` if no header row exists.
    """
    if max_header_rows < 1 or max_title_rows < 0:
        raise ValueError("max_header_rows must be >= 1 and max_title_rows >= 0")
    rows = grid.cells
    index = 0
    while index < len(rows) and index < max_title_rows and _filled(rows[index]) <= 1:
        index += 1
    anchors = {merge[0] for merge in grid.merged}
    count = 0
    while index + count < len(rows) and count < max_header_rows:
        if not _is_header_row(rows[index + count]):
            break
        if count > 0 and index + count in anchors:
            break
        count += 1
    if count == 0:
        raise NormalizeError("header_not_found", "no header row found in the grid")
    return HeaderSpec(start_row=index, row_count=count)


def _apply_merges(
    work: list[list[CellValue]],
    merges: tuple[tuple[int, int, int, int], ...],
    start: int,
    header_count: int,
    width: int,
    fill_merged: bool,
    budget: int,
) -> None:
    """Copy the top-left value of each merge into its empty cells, within a work budget."""
    for r0, c0, r1, c1 in merges:
        top = r0 - start
        if top < 0 or top >= len(work) or c0 < 0 or r1 < r0 or c1 < c0:
            continue
        in_header = top < header_count
        if not in_header and not fill_merged:
            continue
        last_row = header_count - 1 if in_header else len(work) - 1
        bottom = min(r1 - start, last_row)
        anchor = work[top]
        value = anchor[c0] if c0 < len(anchor) else None
        if _is_blank(value):
            continue
        right = min(c1, width - 1)
        for row in work[top : bottom + 1]:
            budget -= right - c0 + 1
            if budget < 0:
                raise NormalizeError("too_large", "merged ranges cover too many cells")
            if right >= len(row):
                row.extend([None] * (right + 1 - len(row)))
            for col in range(c0, right + 1):
                if row[col] is None:
                    row[col] = value


def _clean(value: CellValue) -> str:
    if value is None:
        return ""
    return " ".join(str(value).split())


def _column_names(header_rows: list[list[CellValue]], keep: list[int], joiner: str) -> list[str]:
    names: list[str] = []
    for position, col in enumerate(keep, start=1):
        parts: list[str] = []
        for row in header_rows:
            text = _clean(row[col]) if col < len(row) else ""
            if text and (not parts or parts[-1] != text):
                parts.append(text)
        names.append(joiner.join(parts) if parts else f"column_{position}")
    return names


def _unique(names: list[str]) -> list[str]:
    originals = set(names)
    taken: set[str] = set()
    next_suffix: dict[str, int] = {}
    result: list[str] = []
    for name in names:
        candidate = name
        if candidate in taken:
            suffix = next_suffix.get(name, 2)
            while f"{name}_{suffix}" in originals or f"{name}_{suffix}" in taken:
                suffix += 1
            candidate = f"{name}_{suffix}"
            next_suffix[name] = suffix + 1
        taken.add(candidate)
        result.append(candidate)
    return result


def normalize(
    grid: Grid,
    *,
    header: HeaderSpec | None = None,
    fill_merged: bool = True,
    joiner: str = " / ",
    drop_empty: bool = True,
    max_cells: int = DEFAULT_MAX_CELLS,
) -> Table:
    """Turn a raw :class:`Grid` into a clean :class:`Table`.

    Args:
        grid: The raw cells and merged ranges.
        header: Header location; detected with :func:`detect_header` when ``None``.
        fill_merged: Copy the top-left value of merges anchored in body rows into the other
            cells of the merge. Merges anchored in header rows always propagate inside the
            header rows; they never write into body rows.
        joiner: Separator between the parts of a multi-row header.
        drop_empty: Drop fully empty body rows and fully empty columns.
        max_cells: Limit on the cell count (sum of row lengths) and on the padded output size.

    Raises:
        NormalizeError: ``too_large`` when a limit is exceeded, ``header_not_found`` when no
            header can be detected or the explicit header is outside the grid.
    """
    total = sum(len(row) for row in grid.cells)
    if total > max_cells:
        raise NormalizeError("too_large", f"grid has {total} cells; the limit is {max_cells}")
    if header is None and all(_is_blank(v) for row in grid.cells for v in row):
        return Table(columns=(), rows=())
    spec = header if header is not None else detect_header(grid)
    start, count = spec.start_row, spec.row_count
    if start < 0 or count < 1 or start + count > len(grid.cells):
        raise NormalizeError("header_not_found", "header rows are outside the grid")

    work: list[list[CellValue]] = [list(row) for row in grid.cells[start:]]
    width = max(len(row) for row in work)
    _apply_merges(work, grid.merged, start, count, width, fill_merged, max_cells)

    header_rows = work[:count]
    body = work[count:]
    if drop_empty:
        occupied = {i for row in work for i, v in enumerate(row) if not _is_blank(v)}
        keep = sorted(occupied)
        body = [row for row in body if any(not _is_blank(v) for v in row)]
    else:
        keep = list(range(width))
    if (len(body) + count) * len(keep) > max_cells:
        raise NormalizeError("too_large", f"padded table exceeds the limit of {max_cells} cells")

    columns = tuple(_unique(_column_names(header_rows, keep, joiner)))
    rows = tuple(tuple(row[c] if c < len(row) else None for c in keep) for row in body)
    logger.debug("normalized table: %d columns, %d rows", len(columns), len(rows))
    return Table(columns=columns, rows=rows)
