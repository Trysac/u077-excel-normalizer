# excel-normalizer

[![PyPI version](https://img.shields.io/pypi/v/trysac-excel-normalizer)](https://pypi.org/project/trysac-excel-normalizer/)
[![License: MIT](https://img.shields.io/pypi/l/trysac-excel-normalizer)](LICENSE)
[![CI](https://github.com/Trysac/u077-excel-normalizer/actions/workflows/ci.yml/badge.svg)](https://github.com/Trysac/u077-excel-normalizer/actions/workflows/ci.yml)

Normalize messy Excel sheets with multi-row headers and merged cells into clean tables, with zip-bomb limits and CSV injection escaping

## Problem

Spreadsheets from finance, operations or clients arrive with title rows, two-row headers, merged cells and repeated column names. pandas.read_excel needs the header rows in advance and leaves merged cells empty, so every import script repeats the same cleanup. This library detects and flattens that structure with bounded resources and exports CSV that is safe to open.

## Install

```sh
pip install trysac-excel-normalizer
# or
uv add trysac-excel-normalizer
```

Requires Python 3.10 or newer. Runtime dependencies: openpyxl (>=3.1,<4) and defusedxml (>=0.7,<1).

## Usage

```python
from excel_normalizer import Grid, HeaderSpec, normalize

grid = Grid(
    cells=(("Sales", None), ("Q1", "Q2"), (1, 2)),
    merged=((0, 0, 0, 1),),
)
table = normalize(grid, header=HeaderSpec(start_row=0, row_count=2))
print(table.columns)
print(table.to_csv(), end="")
```

## Public API

- `read_sheet(path, sheet=0, *, max_file_bytes=50_000_000, max_uncompressed_bytes=200_000_000, max_cells=1_000_000) -> Grid`: reads one `.xlsx` sheet with data_only=True and read_only=False. Formulas are never evaluated.
- `detect_header(grid, *, max_header_rows=5, max_title_rows=5) -> HeaderSpec`: skips up to `max_title_rows` leading rows with at most one non-empty cell, then takes consecutive rows where at least 80 percent of non-empty cells are strings, up to `max_header_rows`.
- `normalize(grid, *, header=None, fill_merged=True, joiner=" / ", drop_empty=True, max_cells=1_000_000) -> Table`: flattens multi-row headers, fills merged body cells when `fill_merged` is true, drops empty rows and columns, and makes column names unique.
- `Table.to_csv(*, escape_formulas=True) -> str`: CSV with CRLF terminators; str cells starting with `=`, `+`, `-`, `@`, tab or carriage return get a leading single quote. Numbers are not escaped.
- Types: `Grid`, `HeaderSpec`, `Table`, `CellValue`, `__version__`.
- Errors: `NormalizeError` with `code` in `sheet_not_found`, `too_large`, `header_not_found`, `invalid_file`. Messages are truncated to 200 characters.

Limits: only `.xlsx` is read; a file over `max_file_bytes`, an archive with more than 10_000 members or declaring more than `max_uncompressed_bytes`, and a sheet or grid over `max_cells` are refused. `detect_header` raises ValueError for out-of-range limit arguments.

## Documentation

- [Architecture](docs/ARCHITECTURE.md) and [decisions](docs/decisions/)
- [Testing](docs/TESTING.md) and [security](docs/SECURITY.md)
- [Project record](docs/PROJECT.md) and [stack](docs/STACK.md)
- [Changelog](CHANGELOG.md) and [contributing guide](CONTRIBUTING.md)

## License

[MIT](LICENSE) (c) 2026 Miguel Gomez
