# Project record

## Objective

Read messy Excel sheets with title rows, multi-row headers and merged cells, and turn them into clean tables with unique column names and CSV output that is safe to open, while bounding the resources used on untrusted files.

## Scope

- In scope: the public API documented in [README.md](../README.md): reading one `.xlsx` sheet into a Grid, detecting header rows, flattening headers and merged cells, dropping empty rows and columns, and CSV export with formula escaping. Typed and tested.
- Out of scope: `.xls`, `.ods`, several tables per sheet, writing workbooks, type inference, and anything that needs runtime dependencies beyond openpyxl and defusedxml, or I/O beyond reading a sheet, without an accepted ADR.

## Timeline

| Date | Event |
|---|---|
| 2026 | Repository created from the factory utility template. |

## Effort metrics

Measured values only; a value that has not been measured stays "pending measurement".

| Metric | Value |
|---|---|
| Tokens (input / output) | 458110 / 69353 |
| Hours (human) | pending measurement |
| Cost (USD) | 0.6597 |

## Factory run u077 (measured from API usage)

| Phase | Model | Batch | Input | Output | Cache write | Cache read | Cost (USD) |
|---|---|---|---|---|---|---|---|
| generation | claude-sonnet-5-5 | yes | 2580 | 55922 | 0 | 4246 | 0.2826 |
| fix-1 | claude-sonnet-5-5 | no | 4 | 311 | 5707 | 24595 | 0.0198 |
| fix-2 | claude-sonnet-5-5 | no | 6 | 1498 | 11207 | 45334 | 0.0475 |
| fix-3 | claude-sonnet-5-5 | no | 8 | 1429 | 10108 | 61706 | 0.0457 |
| fix-4 | claude-sonnet-5-5 | no | 10 | 1224 | 9688 | 78945 | 0.0444 |
| audit | claude-sonnet-5-5 | no | 12 | 4075 | 35912 | 108834 | 0.1414 |
| docs | claude-haiku-5-5 | no | 131830 | 4894 | 0 | 0 | 0.0781 |

Total: $0.6597. Prices: Anthropic list prices.
