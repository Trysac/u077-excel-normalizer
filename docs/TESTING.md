# Testing

## Running

```sh
uv run pytest --cov=excel_normalizer   # tests with coverage report
uv run python scripts/check.py      # every gate, as CI and the factory run it
```

## Policy

- Branch coverage of `src/excel_normalizer` must stay at or above 80 %; the gate fails below it.
- Tests exercise the public API (`import excel_normalizer`) and assert observable effects: return values, raised exceptions and their codes.
- Each bug fix starts with a failing test that reproduces it.
- Tests are deterministic: no network, no wall clock, no dependence on test order.

## Layout

- `tests/test_excel_normalizer.py` covers the public API: header detection, two-row headers with merges, merge fill with and without `fill_merged`, empty row and column dropping, `drop_empty=False`, whitespace-only cells, duplicate and suffix naming, the `too_large` limits, the zip checks in `read_sheet` with openpyxl forbidden, the formula and sheet-not-found cases, the `DEFUSEDXML` assertion and CSV escaping.
- Fixtures build `.xlsx` files with openpyxl inside the tests.
- Shared fixtures would go in `tests/conftest.py`; none is present in the repository snapshot.
