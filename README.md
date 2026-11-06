# excel-normalizer

[![PyPI version](https://img.shields.io/pypi/v/trysac-excel-normalizer)](https://pypi.org/project/trysac-excel-normalizer/)
[![License: MIT](https://img.shields.io/pypi/l/trysac-excel-normalizer)](LICENSE)
[![CI](https://github.com/Trysac/u077-excel-normalizer/actions/workflows/ci.yml/badge.svg)](https://github.com/Trysac/u077-excel-normalizer/actions/workflows/ci.yml)

Normalize messy Excel sheets with multi-row headers and merged cells into clean tables, with zip-bomb limits and CSV injection escaping

## Problem

Describe in two or three sentences the concrete problem this utility solves, who hits it, and what
they do today without it. Replace this paragraph when the first real feature lands.

## Install

```sh
pip install trysac-excel-normalizer
# or
uv add trysac-excel-normalizer
```

Requires Python 3.10 or newer. No runtime dependencies.

## Usage

```python
import excel_normalizer

# normalize() trims surrounding whitespace and lower-cases the text.
print(excel_normalizer.normalize("  Hello World "))  # -> "hello world"

# Blank input is rejected explicitly instead of returning an empty string.
try:
    excel_normalizer.normalize("   ")
except ValueError as exc:
    print(f"rejected: {exc}")
```

## Documentation

- [Architecture](docs/ARCHITECTURE.md) and [decisions](docs/decisions/)
- [Testing](docs/TESTING.md) and [security](docs/SECURITY.md)
- [Project record](docs/PROJECT.md) and [stack](docs/STACK.md)
- [Changelog](CHANGELOG.md) and [contributing guide](CONTRIBUTING.md)

## License

[MIT](LICENSE) (c) 2026 Miguel Gomez
