# -*- coding: utf-8 -*-
"""Sanitized worker for Flight Task Automation.

The production build connects to an authorized airline desktop environment,
collects flight metadata, controls Excel workbooks, validates route-specific
documents and assembles a final PDF package.

Company templates, proprietary menu mappings, internal paths and operational
records are intentionally excluded from this public portfolio snapshot.
"""
from __future__ import annotations

import json
import os
import re
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import airportsdata
import pythoncom
import win32com.client
from pypdf import PdfReader, PdfWriter
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from rapidfuzz import fuzz, process
from unidecode import unidecode

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "flight_task_config.json"
RESULT_PATH = APP_DIR / "_last_result.json"
ROOT_FOLDER_NAME = "flight task"
DATE_FOLDER_FORMAT = "%d.%m"

# Representative window markers retained to demonstrate the integration style.
TARGET_WINDOW_MARKERS = (
    "Планирование АК",
    "Полетные задания",
    "Авиабит",
    "Aviabit",
)


@dataclass
class FlightMeta:
    flight_date: datetime
    flight_raw: str
    flight_numbers: list[str]
    route: str
    route_points: list[str]
    international: bool
    turnaround: bool


@dataclass
class AirportMatch:
    original: str
    iata: str
    icao: str
    city: str
    name: str
    country: str
    score: float


def normalize_text(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").split()).strip()


def normalize_airport_search_text(value: str) -> str:
    text = unidecode(normalize_text(value)).lower()
    text = re.sub(r"[(){}\[\],.;:'\"«»/\\_-]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def build_airport_index() -> tuple[list[str], list[tuple[str, ...]]]:
    records = airportsdata.load("IATA")
    search_strings: list[str] = []
    metadata: list[tuple[str, ...]] = []

    for iata, record in records.items():
        if not iata:
            continue

        iata = normalize_text(iata).upper()
        icao = normalize_text(record.get("icao", "")).upper()
        city = normalize_text(record.get("city", ""))
        name = normalize_text(record.get("name", ""))
        country = normalize_text(record.get("country", "")).upper()

        variants = {iata, icao, city, name, f"{city} {name}".strip()}
        for variant in variants:
            query = normalize_airport_search_text(variant)
            if query:
                search_strings.append(query)
                metadata.append((iata, icao, city, name, country))

    return search_strings, metadata


def resolve_airport(
    point: str,
    search_strings: list[str],
    metadata: list[tuple[str, ...]],
) -> Optional[AirportMatch]:
    """Resolve an airport while preferring deterministic exact/contained matches.

    Fuzzy matching is used only as a fallback. This prevents short or distinctive
    airport names from being outranked by unrelated records with a coincidental
    fuzzy score.
    """
    query = normalize_airport_search_text(point)
    if not query:
        return None

    # Exact normalized code/name/city match first.
    for index, candidate in enumerate(search_strings):
        if candidate == query:
            iata, icao, city, name, country = metadata[index]
            return AirportMatch(point, iata, icao, city, name, country, 100.0)

    # Prefer a whole query contained in a longer airport name/city representation.
    contained_matches: list[tuple[int, int]] = []
    for index, candidate in enumerate(search_strings):
        if query in candidate:
            contained_matches.append((len(candidate), index))

    if contained_matches:
        _, index = min(contained_matches)
        iata, icao, city, name, country = metadata[index]
        return AirportMatch(point, iata, icao, city, name, country, 99.0)

    match = process.extractOne(query, search_strings, scorer=fuzz.WRatio)
    if not match:
        return None

    _, score, index = match
    iata, icao, city, name, country = metadata[index]
    return AirportMatch(
        original=point,
        iata=iata,
        icao=icao,
        city=city,
        name=name,
        country=country,
        score=float(score),
    )


def find_target_window():
    """Find the airline desktop application using both Win32 and UIA backends."""
    candidates = []

    for backend in ("win32", "uia"):
        try:
            windows = Desktop(backend=backend).windows(visible_only=True)
        except Exception:
            continue

        for window in windows:
            try:
                title = normalize_text(window.window_text())
            except Exception:
                continue

            score = sum(
                marker.lower() in title.lower()
                for marker in TARGET_WINDOW_MARKERS
            )
            if score:
                candidates.append((score, window))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def focus_window(window) -> None:
    if window is None:
        raise RuntimeError("Target application window was not found")

    try:
        if window.is_minimized():
            window.restore()
    except Exception:
        pass

    window.set_focus()
    time.sleep(0.15)


def invoke_menu_shortcut(keys: str) -> None:
    """Representative UI-automation entry point.

    Production keyboard/menu mappings are not included in the public snapshot.
    """
    app = find_target_window()
    focus_window(app)
    send_keys(keys, pause=0.05)


def connect_excel():
    """Attach to an already-running Excel instance via COM."""
    pythoncom.CoInitialize()
    return win32com.client.GetActiveObject("Excel.Application")


def export_sheet_to_pdf(sheet, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.ExportAsFixedFormat(0, str(output_path))


def close_generated_book(workbook) -> None:
    """Close a generated workbook without modifying the source document."""
    workbook.Close(SaveChanges=False)


def merge_pdfs(paths: list[Path], destination: Path) -> None:
    if not paths:
        raise ValueError("No PDF files supplied")

    writer = PdfWriter()
    for path in paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)

    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as output:
        writer.write(output)


def parse_flight_numbers(value: str) -> list[str]:
    values = re.findall(r"\d+", normalize_text(value))
    result: list[str] = []

    for item in values:
        normalized = item.lstrip("0") or "0"
        if normalized not in result:
            result.append(normalized)

    return result


def parse_route(value: str) -> list[str]:
    text = normalize_text(value)
    return [
        normalize_text(point)
        for point in re.split(r"\s*[-–—]\s*", text)
        if normalize_text(point)
    ]


def is_international(route_points: list[str]) -> bool:
    search_strings, metadata = build_airport_index()
    resolved = [
        resolve_airport(point, search_strings, metadata)
        for point in route_points
    ]
    countries = {
        airport.country
        for airport in resolved
        if airport is not None and airport.country
    }
    return bool(countries - {"RU"})


def read_config() -> dict:
    if not CONFIG_PATH.exists():
        return {}

    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def output_root() -> Path:
    config = read_config()
    if config.get("use_network_output"):
        configured = os.getenv("FLIGHT_TASK_NETWORK_OUTPUT")
        if configured:
            return Path(configured)
    return Path.home() / "Desktop" / ROOT_FOLDER_NAME


def write_last_result(meta: FlightMeta, final_pdf: Path, pages: int) -> None:
    payload = {
        "flight": meta.flight_raw,
        "route": meta.route,
        "date": meta.flight_date.strftime("%Y-%m-%d"),
        "pdf": str(final_pdf),
        "pages": pages,
    }
    RESULT_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def assemble_package(meta: FlightMeta, generated_parts: list[Path]) -> Path:
    """Build the final PDF through a temporary file, then replace atomically."""
    if not generated_parts:
        raise RuntimeError("No generated documents were provided")
    if not meta.flight_numbers:
        raise RuntimeError("Flight number is missing")

    first_flight = meta.flight_numbers[0]
    output_dir = output_root() / meta.flight_date.strftime(DATE_FOLDER_FORMAT)
    output_dir.mkdir(parents=True, exist_ok=True)
    final_pdf = output_dir / f"{first_flight}.pdf"

    with tempfile.TemporaryDirectory(
        prefix=f"_build_{first_flight}_",
        dir=str(output_dir),
    ) as temp_dir:
        ready_pdf = Path(temp_dir) / "ready.pdf"
        merge_pdfs(generated_parts, ready_pdf)
        ready_pdf.replace(final_pdf)

    pages = len(PdfReader(str(final_pdf)).pages)
    write_last_result(meta, final_pdf, pages)
    return final_pdf


def main() -> int:
    """Public integration boundary.

    The production build continues from here by reading the active flight from
    the authorized airline application, collecting generated Excel workbooks,
    validating the expected document set for the route, exporting PDFs and
    calling ``assemble_package``.
    """
    print("Flight Task Automation — sanitized portfolio snapshot")
    print(
        "Production-only airline integration is intentionally excluded from "
        "the public repository."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
