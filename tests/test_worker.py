from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import worker  # noqa: E402


def test_normalize_text_collapses_whitespace() -> None:
    assert worker.normalize_text("  SVO\xa0 -   LED  ") == "SVO - LED"


def test_normalize_airport_search_text_transliterates() -> None:
    assert worker.normalize_airport_search_text("Санкт-Петербург") == "sankt peterburg"


def test_parse_flight_numbers_deduplicates_and_strips_zeroes() -> None:
    assert worker.parse_flight_numbers("N4 00123 / 123 / 00456") == ["4", "123", "456"]


def test_parse_route_supports_dash_variants() -> None:
    assert worker.parse_route("SVO – LED — KZN") == ["SVO", "LED", "KZN"]


def test_resolve_airport_returns_match() -> None:
    search_strings, metadata = worker.build_airport_index()
    match = worker.resolve_airport("Sheremetyevo", search_strings, metadata)

    assert match is not None
    assert match.iata == "SVO"
    assert match.score > 70
