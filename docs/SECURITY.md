# Security

This document is the threat model of `trysac-excel-normalizer`. It reads spreadsheets that usually come from third parties, so every input is treated as untrusted.

## Threat model

### Defended

| Threat | Mitigation |
|---|---|
| Zip bomb (small file, huge expansion) | `read_sheet` checks the file size (`max_file_bytes`, default 50_000_000), opens the file with `zipfile`, refuses more than 10_000 members and sums `ZipInfo.file_size` of every member against `max_uncompressed_bytes` (default 200_000_000) before openpyxl runs. `zipfile` stops decompressing at the declared size and then checks the CRC, so an under-declared member fails with `invalid_file` instead of expanding. |
| XML attacks (entity expansion, external entities) | openpyxl uses `defusedxml` automatically when it is installed. The test suite asserts `openpyxl.xml.DEFUSEDXML is True`. |
| Formula execution | Workbooks load with `data_only=True`; cached values are read and formulas are never evaluated. A formula without a cached value reads as `None`. |
| Huge sheets | `max_cells` (default 1_000_000) bounds `max_row * max_column` of the sheet in `read_sheet` and the sum of the row lengths of a `Grid` in `normalize`. The padded output size and the work done by merged ranges are bounded by the same limit. |
| CSV injection | `Table.to_csv` follows the OWASP CSV Injection guidance: a `str` cell (and header name) starting with `=`, `+`, `-`, `@`, tab or carriage return gets a leading single quote. Numbers are not escaped. `escape_formulas=False` disables this. |
| Information leak through errors | Every failure is a `NormalizeError` with a stable `code`. Messages are truncated to 200 characters and do not contain file contents. |
| Log leakage | The library uses `logging.getLogger(__name__)` with a `NullHandler`, logs counts at debug level and never logs cell values. |

### Not defended

- Memory use is not bounded by a streaming reader: `read_sheet` loads the workbook with `read_only=False` because read-only sheets lack `merged_cells`, so memory grows with the declared uncompressed size, capped by `max_uncompressed_bytes`.
- The zip limits apply to what the archive declares in its central directory; a forged declaration is only caught when `zipfile` enforces the declared size while reading.
- Files that are not `.xlsx` are not read; other formats are rejected by design rather than sanitized.
- CSV escaping covers only the listed leading characters; the consumer's spreadsheet application remains responsible for its own handling of the output.

## Reporting a vulnerability

See the repository `SECURITY.md` for the disclosure process.
