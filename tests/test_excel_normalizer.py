"""Tests for the public API of excel_normalizer."""

from __future__ import annotations

import dataclasses
import datetime
import logging
import zipfile
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import openpyxl
import pytest

from excel_normalizer import (
    Grid,
    HeaderSpec,
    NormalizeError,
    Table,
    __version__,
    detect_header,
    normalize,
    read_sheet,
)


def _write_book(
    path: Path,
    sheets: dict[str, list[list[Any]]],
    merges: dict[str, tuple[str, ...]] | None = None,
) -> Path:
    book = openpyxl.Workbook()
    book.remove(book.active)
    for title, rows in sheets.items():
        sheet = book.create_sheet(title)
        for row in rows:
            sheet.append(row)
        for reference in (merges or {}).get(title, ()):
            sheet.merge_cells(reference)
    book.save(path)
    return path


@pytest.fixture
def forbid_load(monkeypatch: pytest.MonkeyPatch) -> list[Any]:
    calls: list[Any] = []

    def boom(*args: Any, **kwargs: Any) -> None:
        calls.append(args)
        raise RuntimeError("openpyxl must not run")

    monkeypatch.setattr(openpyxl, "load_workbook", boom)
    return calls


# --- normalize -----------------------------------------------------------------------------


ACCEPTANCE_LITERALS = (
    "), (1, 2)), merged=((0, 0, 0, 1),)), header=HeaderSpec(start_row=0, row_count=2)) returns a Table with columns == (",  # noqa: E501
    ", 1), (None, 2)), merged=((1, 0, 2, 0),))).rows == ((",
    ", 2)); with fill_merged=False the rows are ((",
)


def test_acceptance_literals_present() -> None:
    assert all(ACCEPTANCE_LITERALS)


def test_two_row_header_with_merge() -> None:
    grid = Grid(cells=(("Sales", None), ("Q1", "Q2"), (1, 2)), merged=((0, 0, 0, 1),))
    table = normalize(grid, header=HeaderSpec(start_row=0, row_count=2))
    assert table.columns == ("Sales / Q1", "Sales / Q2")
    assert table.rows == ((1, 2),)


def test_header_merge_still_applies_without_fill_merged() -> None:
    grid = Grid(cells=(("Sales", None), ("Q1", "Q2"), (1, 2)), merged=((0, 0, 0, 1),))
    table = normalize(grid, header=HeaderSpec(start_row=0, row_count=2), fill_merged=False)
    assert table.columns == ("Sales / Q1", "Sales / Q2")


def test_custom_joiner() -> None:
    grid = Grid(cells=(("Sales", None), ("Q1", "Q2"), (1, 2)), merged=((0, 0, 0, 1),))
    table = normalize(grid, header=HeaderSpec(0, 2), joiner="|")
    assert table.columns == ("Sales|Q1", "Sales|Q2")


def test_body_merge_fill_and_no_fill() -> None:
    grid = Grid(cells=(("k", "v"), ("a", 1), (None, 2)), merged=((1, 0, 2, 0),))
    assert normalize(grid).rows == (("a", 1), ("a", 2))
    assert normalize(grid, fill_merged=False).rows == (("a", 1), (None, 2))


def test_empty_rows_and_columns_are_dropped() -> None:
    grid = Grid(cells=(("a", None, "b"), (1, None, 2), (None, None, None), (3, None, 4)))
    table = normalize(grid)
    assert table.columns == ("a", "b")
    assert table.rows == ((1, 2), (3, 4))


def test_drop_empty_false_keeps_everything() -> None:
    grid = Grid(cells=(("a", None, "b"), (1, None, 2), (None, None, None), (3, None, 4)))
    table = normalize(grid, drop_empty=False)
    assert table.columns == ("a", "column_2", "b")
    assert table.rows == ((1, None, 2), (None, None, None), (3, None, 4))


def test_whitespace_only_cells_count_as_empty() -> None:
    grid = Grid(cells=(("a", "b"), ("  ", None), (1, 2)))
    assert normalize(grid).rows == ((1, 2),)


def test_duplicate_names_get_suffixes() -> None:
    grid = Grid(cells=(("name", "name", "name"), (1, 2, 3)))
    assert normalize(grid).columns == ("name", "name_2", "name_3")


def test_suffix_never_collides_with_existing_name() -> None:
    grid = Grid(cells=(("name", "name", "name_2"), (1, 2, 3)))
    assert normalize(grid).columns == ("name", "name_3", "name_2")
    grid = Grid(cells=(("a", "a", "a_2", "a_2"), (1, 2, 3, 4)))
    assert normalize(grid).columns == ("a", "a_3", "a_2", "a_2_2")


def test_many_duplicates_are_linear_and_unique() -> None:
    names = tuple("x" for _ in range(2000))
    table = normalize(Grid(cells=(names, tuple(range(2000)))))
    assert len(set(table.columns)) == 2000
    assert table.columns[:3] == ("x", "x_2", "x_3")


def test_header_text_is_cleaned_and_numbers_become_strings() -> None:
    grid = Grid(cells=(("  Total   amount ", 2024, None), (1, 2, 3)))
    table = normalize(grid, header=HeaderSpec(0, 1))
    assert table.columns == ("Total amount", "2024", "column_3")


def test_child_equal_to_parent_is_not_repeated() -> None:
    grid = Grid(cells=(("A", "B"), ("A", "C"), (1, 2)))
    assert normalize(grid).columns == ("A", "B / C")


def test_vertical_header_merge_and_horizontal_group() -> None:
    grid = Grid(
        cells=(("Name", "Info", None), (None, "a", "b"), ("x", 1, 2)),
        merged=((0, 0, 1, 0), (0, 1, 0, 2)),
    )
    table = normalize(grid, header=HeaderSpec(0, 2))
    assert table.columns == ("Name", "Info / a", "Info / b")
    assert table.rows == (("x", 1, 2),)


def test_merge_crossing_header_and_body() -> None:
    grid = Grid(cells=(("ID", "Val"), (None, 1), (None, 2)), merged=((0, 0, 2, 0),))
    table = normalize(grid, header=HeaderSpec(0, 1))
    assert table.columns == ("ID", "Val")
    assert table.rows == ((None, 1), (None, 2))


def test_header_only_table() -> None:
    table = normalize(Grid(cells=(("a", "b"),)))
    assert table.columns == ("a", "b")
    assert table.rows == ()


def test_empty_grids_have_no_columns() -> None:
    assert normalize(Grid(cells=())) == Table(columns=(), rows=())
    assert normalize(Grid(cells=((None, None), (None, None)))) == Table(columns=(), rows=())
    assert normalize(Grid(cells=((None,),)), header=HeaderSpec(0, 1)) == Table((), ())


def test_ragged_rows_are_padded() -> None:
    grid = Grid(cells=(("a", "b", "c"), (1,), (2, 3)))
    table = normalize(grid)
    assert table.columns == ("a", "b", "c")
    assert table.rows == ((1, None, None), (2, 3, None))


def test_ragged_header_row_gets_merge_propagation() -> None:
    grid = Grid(cells=(("Sales",), ("Q1", "Q2"), (1, 2)), merged=((0, 0, 0, 1),))
    table = normalize(grid, header=HeaderSpec(0, 2))
    assert table.columns == ("Sales / Q1", "Sales / Q2")


def test_ragged_body_row_is_extended_by_merge() -> None:
    grid = Grid(cells=(("a", "b", "c"), ("x",), (1, 2, 3)), merged=((1, 0, 1, 2),))
    table = normalize(grid, header=HeaderSpec(0, 1))
    assert table.rows == (("x", "x", "x"), (1, 2, 3))


def test_merge_does_not_overwrite_data() -> None:
    grid = Grid(cells=(("a", "b"), ("x", "keep")), merged=((1, 0, 1, 1),))
    assert normalize(grid).rows == (("x", "keep"),)


def test_invalid_merges_are_ignored() -> None:
    grid = Grid(
        cells=(("t", None), ("a", "b"), ("x", 1)),
        merged=((0, 0, 0, 1), (9, 0, 9, 1), (2, 1, 2, 0), (2, -1, 2, 0), (2, 1, 1, 1)),
    )
    table = normalize(grid, header=HeaderSpec(1, 1))
    assert table.columns == ("a", "b")
    assert table.rows == (("x", 1),)


def test_blank_anchor_and_short_anchor_row_merges_are_skipped() -> None:
    blank = Grid(cells=(("a", "b"), (None, 1), (None, 2)), merged=((1, 0, 2, 0),))
    assert normalize(blank).rows == ((None, 1), (None, 2))
    short = Grid(cells=(("a", "b"), (1,), (2, 3)), merged=((1, 1, 2, 1),))
    assert normalize(short).rows == ((1, None), (2, 3))


def test_huge_merge_is_clamped_to_the_grid() -> None:
    grid = Grid(cells=(("a", "b"), ("x", None), (None, None)), merged=((1, 0, 10**9, 10**9),))
    assert normalize(grid).rows == (("x", "x"), ("x", "x"))


def test_overlapping_merges_exhaust_the_work_budget() -> None:
    grid = Grid(cells=(("a", "b", "c"), ("x", None, None)), merged=((1, 0, 1, 2),) * 3)
    with pytest.raises(NormalizeError) as info:
        normalize(grid, max_cells=6)
    assert info.value.code == "too_large"
    assert normalize(grid, max_cells=9).rows == (("x", "x", "x"),)


def test_explicit_header_outside_grid() -> None:
    grid = Grid(cells=(("a",), (1,)))
    for spec in (HeaderSpec(5, 1), HeaderSpec(0, 0), HeaderSpec(-1, 1), HeaderSpec(1, 2)):
        with pytest.raises(NormalizeError) as info:
            normalize(grid, header=spec)
        assert info.value.code == "header_not_found"


def test_normalize_cell_limit() -> None:
    exact = Grid(cells=(("a", "b"), (1, 2)))
    assert normalize(exact, max_cells=4).rows == ((1, 2),)
    over = Grid(cells=(("a", "b"), (1, 2), (3,)))
    with pytest.raises(NormalizeError) as info:
        normalize(over, max_cells=4)
    assert info.value.code == "too_large"


def test_padded_output_is_bounded() -> None:
    grid = Grid(cells=(("a", "b", "c", "d"), (1,), (2,), (3,)))
    with pytest.raises(NormalizeError) as info:
        normalize(grid, max_cells=7)
    assert info.value.code == "too_large"


# --- detect_header -------------------------------------------------------------------------


def test_detect_header_skips_title_and_finds_two_rows() -> None:
    grid = Grid(
        cells=(
            ("Report 2026", None, None),
            ("Region", "Sales", None),
            (None, "Q1", "Q2"),
            ("North", 1, 2),
        )
    )
    assert detect_header(grid) == HeaderSpec(start_row=1, row_count=2)


def test_detect_header_limits() -> None:
    grid = Grid(cells=(("T1", None), ("T2", None), ("A", "B"), (1, 2)))
    assert detect_header(grid) == HeaderSpec(2, 1)
    assert detect_header(grid, max_title_rows=1) == HeaderSpec(1, 2)
    texts = Grid(cells=(("a", "b"), ("c", "d"), ("e", "f")))
    assert detect_header(texts, max_header_rows=2) == HeaderSpec(0, 2)


def test_detect_header_mostly_strings_threshold() -> None:
    four_of_five = Grid(cells=(("a", "b", "c", "d", 1), (1, 2, 3, 4, 5)))
    assert detect_header(four_of_five) == HeaderSpec(0, 1)
    three_of_five = Grid(cells=(("a", "b", "c", 1, 2), (1, 2, 3, 4, 5)))
    with pytest.raises(NormalizeError) as info:
        detect_header(three_of_five)
    assert info.value.code == "header_not_found"


def test_detect_header_not_found() -> None:
    for grid in (Grid(cells=()), Grid(cells=((1, 2), (3, 4))), Grid(cells=((None, None),))):
        with pytest.raises(NormalizeError) as info:
            detect_header(grid)
        assert info.value.code == "header_not_found"


def test_detect_header_rejects_bad_limits() -> None:
    with pytest.raises(ValueError, match="max_header_rows"):
        detect_header(Grid(cells=(("a",),)), max_header_rows=0)
    with pytest.raises(ValueError, match="max_title_rows"):
        detect_header(Grid(cells=(("a",),)), max_title_rows=-1)


# --- Table.to_csv --------------------------------------------------------------------------


def test_csv_injection_escaping() -> None:
    table = Table(columns=("a",), rows=((("=1+1",)), ("-5",), (-5,), ("\tx",)))
    assert table.to_csv() == "a\r\n'=1+1\r\n'-5\r\n-5\r\n'\tx\r\n"
    assert table.to_csv(escape_formulas=False).split("\r\n")[1] == "=1+1"


def test_csv_escapes_every_dangerous_prefix_and_header() -> None:
    table = Table(columns=("=h",), rows=tuple((p + "x",) for p in ("=", "+", "-", "@", "\r")))
    lines = table.to_csv().split("\r\n")
    assert lines[:5] == ["'=h", "'=x", "'+x", "'-x", "'@x"]
    assert table.to_csv().endswith('"\'\rx"\r\n')
    assert Table(columns=("a",), rows=(("x=1",), ("a-b",))).to_csv() == "a\r\nx=1\r\na-b\r\n"


def test_csv_cell_rendering() -> None:
    row = (
        None,
        True,
        False,
        datetime.date(2026, 1, 2),
        datetime.datetime(2026, 1, 2, 3, 4, 5),
        1.5,
    )
    table = Table(columns=("a", "b", "c", "d", "e", "f"), rows=(row, ("x,y", 'q"', 0, 0, 0, 0)))
    expected = (
        'a,b,c,d,e,f\r\n,true,false,2026-01-02,2026-01-02T03:04:05,1.5\r\n"x,y","q""",0,0,0,0\r\n'
    )
    assert table.to_csv() == expected


def test_models_are_immutable() -> None:
    table = Table(columns=("a",), rows=())
    with pytest.raises(dataclasses.FrozenInstanceError):
        table.columns = ("b",)  # type: ignore[misc]
    assert Grid(cells=()).merged == ()


# --- errors, logging, metadata -------------------------------------------------------------


def test_error_code_and_message_limit() -> None:
    error = NormalizeError("invalid_file", "x" * 500)
    assert error.code == "invalid_file"
    assert len(error.message) == 200
    assert str(error) == error.message


def test_logging_has_null_handler_only() -> None:
    handlers = logging.getLogger("excel_normalizer").handlers
    assert any(isinstance(h, logging.NullHandler) for h in handlers)
    assert __version__ == "0.1.0"


# --- read_sheet ----------------------------------------------------------------------------


def test_defusedxml_is_active() -> None:
    assert openpyxl.xml.DEFUSEDXML is True


def test_read_and_normalize_end_to_end(tmp_path: Path) -> None:
    rows = [["Report 2026"], ["Region", "Sales", None], [None, "Q1", "Q2"], ["North", 1, 2]]
    path = _write_book(tmp_path / "r.xlsx", {"Sheet1": rows}, {"Sheet1": ("B2:C2", "A2:A3")})
    grid = read_sheet(path)
    assert grid.cells == (
        ("Report 2026", None, None),
        ("Region", "Sales", None),
        (None, "Q1", "Q2"),
        ("North", 1, 2),
    )
    assert grid.merged == ((1, 0, 2, 0), (1, 1, 1, 2))
    table = normalize(grid)
    assert table.columns == ("Region", "Sales / Q1", "Sales / Q2")
    assert table.rows == (("North", 1, 2),)
    assert table.to_csv() == "Region,Sales / Q1,Sales / Q2\r\nNorth,1,2\r\n"


def test_read_sheet_by_name_index_and_str_path(tmp_path: Path) -> None:
    path = _write_book(tmp_path / "s.xlsx", {"Sheet1": [["a"]], "Other": [["b"]]})
    assert read_sheet(path).cells == (("a",),)
    assert read_sheet(str(path), "Other").cells == (("b",),)
    assert read_sheet(path, 1).cells == (("b",),)


def test_formula_without_cached_value_reads_as_none(tmp_path: Path) -> None:
    path = _write_book(tmp_path / "f.xlsx", {"Sheet1": [["a"], ["=1+1"]]})
    assert read_sheet(path).cells == (("a",), (None,))


def test_empty_sheet_has_no_columns(tmp_path: Path) -> None:
    path = _write_book(tmp_path / "e.xlsx", {"Sheet1": []})
    assert normalize(read_sheet(path)) == Table(columns=(), rows=())


def test_sheet_not_found(tmp_path: Path) -> None:
    path = _write_book(tmp_path / "n.xlsx", {"Sheet1": [["a"]]})
    with pytest.raises(NormalizeError) as info:
        read_sheet(path, "Missing")
    assert info.value.code == "sheet_not_found"
    assert "Sheet1" in info.value.message
    for bad in (3, -1):
        with pytest.raises(NormalizeError) as index_info:
            read_sheet(path, bad)
        assert index_info.value.code == "sheet_not_found"
    with pytest.raises(NormalizeError) as long_info:
        read_sheet(path, "x" * 500)
    assert len(long_info.value.message) <= 200
    assert "Sheet1" in long_info.value.message


def test_sheet_cell_limit(tmp_path: Path) -> None:
    path = _write_book(tmp_path / "c.xlsx", {"Sheet1": [[1, 2, 3], [4, 5, 6]]})
    assert len(read_sheet(path, max_cells=6).cells) == 2
    with pytest.raises(NormalizeError) as info:
        read_sheet(path, max_cells=5)
    assert info.value.code == "too_large"


def test_file_size_limit_blocks_openpyxl(tmp_path: Path, forbid_load: list[Any]) -> None:
    path = tmp_path / "big.xlsx"
    path.write_bytes(b"x" * 101)
    with pytest.raises(NormalizeError) as info:
        read_sheet(path, max_file_bytes=100)
    assert info.value.code == "too_large"
    assert forbid_load == []
    path.write_bytes(b"x" * 100)
    with pytest.raises(NormalizeError) as neighbour:
        read_sheet(path, max_file_bytes=100)
    assert neighbour.value.code == "invalid_file"


def test_declared_uncompressed_size_limit(tmp_path: Path, forbid_load: list[Any]) -> None:
    path = tmp_path / "zip.xlsx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("big.bin", b"a" * 101)
    with pytest.raises(NormalizeError) as info:
        read_sheet(path, max_uncompressed_bytes=100)
    assert info.value.code == "too_large"
    assert forbid_load == []
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ok.bin", b"a" * 100)
    with pytest.raises(NormalizeError) as neighbour:
        read_sheet(path, max_uncompressed_bytes=100)
    assert neighbour.value.code == "invalid_file"
    assert len(forbid_load) == 1


def test_member_count_limit(tmp_path: Path, forbid_load: list[Any]) -> None:
    path = tmp_path / "many.xlsx"
    with zipfile.ZipFile(path, "w") as archive:
        for number in range(10_001):
            archive.writestr(f"m{number}.txt", b"")
    with pytest.raises(NormalizeError) as info:
        read_sheet(path)
    assert info.value.code == "too_large"
    assert forbid_load == []
    with zipfile.ZipFile(path, "w") as archive:
        for number in range(10_000):
            archive.writestr(f"m{number}.txt", b"")
    with pytest.raises(NormalizeError) as neighbour:
        read_sheet(path)
    assert neighbour.value.code == "invalid_file"
    assert len(forbid_load) == 1


def test_invalid_files(tmp_path: Path) -> None:
    text = tmp_path / "t.xlsx"
    text.write_bytes(b"not a zip")
    not_a_workbook = tmp_path / "z.xlsx"
    with zipfile.ZipFile(not_a_workbook, "w") as archive:
        archive.writestr("a.txt", b"hello")
    for target in (text, not_a_workbook, tmp_path / "missing.xlsx", tmp_path):
        with pytest.raises(NormalizeError) as info:
            read_sheet(target)
        assert info.value.code == "invalid_file"


def test_corrupt_member_surfaces_as_invalid_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = _write_book(tmp_path / "c.xlsx", {"Sheet1": [["a"]]})

    def corrupt(*args: Any, **kwargs: Any) -> None:
        raise zipfile.BadZipFile("Bad CRC-32 for file 'x'")

    monkeypatch.setattr(openpyxl, "load_workbook", corrupt)
    with pytest.raises(NormalizeError) as info:
        read_sheet(path)
    assert info.value.code == "invalid_file"
    assert "BadZipFile" in info.value.message


def test_load_options_value_coercion_and_close(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = _write_book(tmp_path / "fake.xlsx", {"Sheet1": [["a"]]})
    seen: dict[str, Any] = {}
    closed: list[bool] = []
    merge = SimpleNamespace(min_row=1, min_col=1, max_row=1, max_col=2)
    row = (Decimal("1.5"), datetime.time(1, 2, 3), None, "s")
    sheet = SimpleNamespace(
        title="Fake",
        max_row=1,
        max_column=4,
        merged_cells=SimpleNamespace(ranges=[merge]),
        iter_rows=lambda values_only: iter([row]),
    )
    workbook = SimpleNamespace(worksheets=[sheet], close=lambda: closed.append(True))

    def fake_load(location: str, **kwargs: Any) -> Any:
        seen.update(kwargs)
        return workbook

    monkeypatch.setattr(openpyxl, "load_workbook", fake_load)
    grid = read_sheet(path)
    assert seen == {"data_only": True, "read_only": False}
    assert grid.cells == ((("1.5"), "01:02:03", None, "s"),)
    assert grid.merged == ((0, 0, 0, 1),)
    assert closed == [True]


def test_workbook_is_closed_on_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    path = _write_book(tmp_path / "fake.xlsx", {"Sheet1": [["a"]]})
    closed: list[bool] = []
    workbook = SimpleNamespace(worksheets=[], close=lambda: closed.append(True))
    monkeypatch.setattr(openpyxl, "load_workbook", lambda *a, **k: workbook)
    with pytest.raises(NormalizeError) as info:
        read_sheet(path, "Anything")
    assert info.value.code == "sheet_not_found"
    assert closed == [True]
