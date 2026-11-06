"""Normalize messy Excel sheets with multi-row headers and merged cells into clean tables.

The library reads ``.xlsx`` sheets into an immutable :class:`Grid`, detects title and header
rows, flattens merged cells and exports CSV that is safe to open in a spreadsheet.

Security model (see ``docs/SECURITY.md``): before openpyxl runs, :func:`read_sheet` checks the
file size, opens the file with :mod:`zipfile`, refuses more than 10_000 members and sums the
declared ``ZipInfo.file_size`` of every member against ``max_uncompressed_bytes`` (zipfile stops
decompressing at the declared size and then checks the CRC, so an under-declared member fails
with ``invalid_file`` instead of expanding). openpyxl uses defusedxml automatically when it is
installed. Workbooks are loaded with ``data_only=True`` and ``read_only=False``; formulas are
never evaluated. CSV export follows the OWASP CSV Injection guidance.
"""

from __future__ import annotations

import logging

from .errors import NormalizeError
from .models import CellValue, Grid, HeaderSpec, Table
from .normalizer import detect_header, normalize
from .reader import read_sheet

__all__ = [
    "CellValue",
    "Grid",
    "HeaderSpec",
    "NormalizeError",
    "Table",
    "__version__",
    "detect_header",
    "normalize",
    "read_sheet",
]

__version__ = "0.1.0"

logging.getLogger(__name__).addHandler(logging.NullHandler())
