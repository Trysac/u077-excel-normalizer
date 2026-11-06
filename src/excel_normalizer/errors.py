"""Typed error raised by the library."""

from __future__ import annotations

from typing import Literal

ErrorCode = Literal["sheet_not_found", "too_large", "header_not_found", "invalid_file"]

MAX_MESSAGE_CHARS = 200


class NormalizeError(Exception):
    """Base error of the library.

    Attributes:
        code: Stable machine-readable identifier of the failure.
        message: Human readable description, truncated to at most 200 characters.
    """

    def __init__(self, code: ErrorCode, message: str) -> None:
        clipped = message[:MAX_MESSAGE_CHARS]
        super().__init__(clipped)
        self.code: ErrorCode = code
        self.message: str = clipped
