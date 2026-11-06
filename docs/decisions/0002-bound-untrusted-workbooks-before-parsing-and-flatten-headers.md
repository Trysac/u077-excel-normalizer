# Bound untrusted workbooks before parsing and flatten headers in a pure function

Status: accepted (generated, pending human review)

## Context

Spreadsheets from clients arrive with title rows, multi-row headers, merged cells and repeated names, and an .xlsx file is a zip archive that can expand enormously or carry hostile XML. pandas.read_excel needs header rows in advance and leaves merged cells empty.

## Decision

Split the work in two layers. read_sheet checks file size, member count (10_000) and the sum of declared member sizes with zipfile before openpyxl runs, then loads with data_only=True and read_only=False (merged_cells is needed) and relies on defusedxml through openpyxl. It returns an immutable Grid. normalize and detect_header are pure functions over the Grid, with a cell limit, a work budget for merges and a bound on padded output. Merges anchored in header rows propagate inside the header rows only; merges anchored in body rows fill the top-left value when fill_merged is true. Duplicate names get the smallest unused _N suffix. CSV export escapes str cells per OWASP. openpyxl is imported with importlib because it has no type stubs.

## Consequences

Hostile archives are rejected with a typed error before any XML is parsed, and normalization is testable without files. Memory use stays proportional to the declared size limit because read-only mode is not available. Single-column sheets are affected by the literal title-row rule. The size checks trust the central directory, backed by zipfile enforcing declared sizes while reading.
