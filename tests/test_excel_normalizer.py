"""Behavior tests for the public API of excel_normalizer."""

from __future__ import annotations

import pytest

import excel_normalizer


def test_normalize_trims_and_lowercases() -> None:
    assert excel_normalizer.normalize("  Hello World ") == "hello world"


@pytest.mark.parametrize("value", ["", "   ", "\t\n"])
def test_normalize_rejects_blank_input(value: str) -> None:
    with pytest.raises(ValueError, match="must not be empty"):
        excel_normalizer.normalize(value)


def test_version_is_semver() -> None:
    parts = excel_normalizer.__version__.split(".")
    assert len(parts) == 3
    assert all(part.isdigit() for part in parts)
