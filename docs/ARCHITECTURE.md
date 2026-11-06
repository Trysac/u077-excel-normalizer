# Architecture

## Purpose

`trysac-excel-normalizer` is a small Python library that reads `.xlsx` sheets, detects title and multi-row header rows, flattens merged cells and exports CSV that is safe to open in a spreadsheet. Its only runtime dependencies are openpyxl and defusedxml.

## Shape

```text
caller code --> excel_normalizer (public API in __init__.py)
                    |
                    +-- reader.py      read_sheet: file checks, openpyxl, Grid
                    +-- normalizer.py  detect_header, normalize
                    +-- models.py      Grid, HeaderSpec, Table, CSV rendering
                    +-- errors.py      NormalizeError with stable codes
```

- **Public API**: the names listed in `__all__` of `src/excel_normalizer/__init__.py`. Anything else is internal.
- **Errors**: failures raise `NormalizeError` with a stable `code`; `detect_header` raises `ValueError` for out-of-range limit arguments.
- **Side effects**: `read_sheet` reads the file system; the other functions are pure. The library logs counts at debug level through a `NullHandler`.

## Quality contract

`scripts/check.py` is the single definition of the gates; `fabrica.json` tells the factory and CI how to run it. See [TESTING.md](TESTING.md).

## Decisions

Architecture decisions are recorded in [decisions/](decisions/) as numbered ADRs.
