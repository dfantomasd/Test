# -*- coding: utf-8 -*-
"""Sanitized portfolio snapshot of the Flight Task worker.

Production-only templates, company paths, document identifiers and operational
records are intentionally excluded from the public repository.
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
from rapidfuzz import fuzz, process
from unidecode import unidecode
from pypdf import PdfReader, PdfWriter

import pythoncom
import win32com.client
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

ROOT_FOLDER_NAME = "flight task"
DATE_FOLDER_FORMAT = "%d.%m"
CONFIG_PATH = Path(__file__).resolve().parent / "flight_task_config.json"
RESULT_PATH = Path(__file__).resolve().parent / "_last_result.json"

# The production application/window names are retained only to show the
# integration approach. Internal document templates and company paths are not.
AVIABIT_TITLE_MARKERS = ("Планирование АК", "Полетные задания", "Авиабит", "Aviabit")


@dataclass
class MainMeta:
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
    value = unidecode(normalize_text(value)).lower()
    value = re.sub(r"[(){}\[\],.;:'\"«»/\\_-]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def build_airport_index():
    records = airportsdata.load("IATA")
    strings, meta = [], []
    for iata, rec in records.items():
        if not iata:
            continue
        iata = normalize_text(iata).upper()
        icao = normalize_text(rec.get("icao", "")).upper()
        city = normalize_text(rec.get("city", ""))
        name = normalize_text(rec.get("name", ""))
        country = normalize_text(rec.get("country", "")).upper()
        for variant in {iata, icao, city, name, f"{city} {name}".strip()}:
            q = normalize_airport_search_text(variant)
            if q:
                strings.append(q)
                meta.append((iata, icao, city, name, country))
    return strings, meta


def resolve_airport(point: str, strings, meta) -> Optional[AirportMatch]:
    query = normalize_airport_search_text(point)
    if not query:
        return None
    match = process.extractOne(query, strings, scorer=fuzz.WRatio)
    if not match:
        return None
    _, score, idx = match
    iata, icao, city, name, country = meta[idx]
    return AirportMatch(point, iata, icao, city, name, country, float(score))


def find_aviabit_window():
    candidates = []
    for backend in ("win32", "uia"):
        try:
            for w in Desktop(backend=backend).windows(visible_only=True):
                title = normalize_text(w.window_text())
                score = sum(m.lower() in title.lower() for m in AVIABIT_TITLE_MARKERS)
                if score:
                    candidates.append((score, w))
        except Exception:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def focus_window(window) -> None:
    if window is None:
        raise RuntimeError("Target application window was not found")
    try:
        window.restore()
    except Exception:
        pass
    window.set_focus()
    time.sleep(0.15)


def invoke_menu_shortcut(keys: str) -> None:
    """Representative UI-automation entry point.

    Production menu mappings are intentionally not published.
    """
    app = find_aviabit_window()
    focus_window(app)
    send_keys(keys, pause=0.05)


def connect_excel():
    pythoncom.CoInitialize()
    return win32com.client.GetActiveObject("Excel.Application")


def export_sheet_to_pdf(sheet, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.ExportAsFixedFormat(0, str(output_path))


def close_generated_book(workbook, excel_app=None) -> None:
    try:
        workbook.Close(SaveChanges=False)
    finally:
        pass


def merge_pdfs(paths: list[Path], destination: Path) -> None:
    writer = PdfWriter()
    for path in paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as fh:
        writer.write(fh)


def parse_flight_numbers(value: str) -> list[str]:
    values = re.findall(r"\d+", normalize_text(value))
    result = []
    for item in values:
        item = item.lstrip("0") or "0"
        if item not in result:
            result.append(item)
    return result


def parse_route(value: str) -> list[str]:
    text = normalize_text(value)
    return [normalize_text(p) for p in re.split(r"\s*[-–—]\s*", text) if normalize_text(p)]


def is_international(route_points: list[str]) -> bool:
    strings, meta = build_airport_index()
    resolved = [resolve_airport(p, strings, meta) for p in route_points]
    countries = {a.country for a in resolved if a is not None and a.country}
    return bool(countries - {"RU"})


def read_config() -> dict:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text("utf-8"))
    except Exception:
        return {}


def output_root() -> Path:
    cfg = read_config()
    if cfg.get("use_z_drive"):
        return Path(os.getenv("FLIGHT_TASK_NETWORK_OUTPUT", "Z:/flight task"))
    return Path.home() / "Desktop" / ROOT_FOLDER_NAME


def write_last_result(meta: MainMeta, final_pdf: Path, pages: int) -> None:
    payload = {
        "flight": meta.flight_raw,
        "route": meta.route,
        "date": meta.flight_date.strftime("%Y-%m-%d"),
        "pdf": str(final_pdf),
        "pages": pages,
    }
    RESULT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), "utf-8")


def assemble_package(meta: MainMeta, generated_parts: list[Path]) -> Path:
    """Merge a validated set of generated PDFs atomically."""
    if not generated_parts:
        raise RuntimeError("No generated documents were provided")

    first_flight = meta.flight_numbers[0]
    out_dir = output_root() / meta.flight_date.strftime(DATE_FOLDER_FORMAT)
    out_dir.mkdir(parents=True, exist_ok=True)
    final_pdf = out_dir / f"{first_flight}.pdf"

    with tempfile.TemporaryDirectory(prefix=f"_build_{first_flight}_", dir=str(out_dir)) as td:
        ready = Path(td) / "ready.pdf"
        merge_pdfs(generated_parts, ready)
        ready.replace(final_pdf)

    pages = len(PdfReader(str(final_pdf)).pages)
    write_last_result(meta, final_pdf, pages)
    return final_pdf


def main() -> int:
    """Portfolio entry point.

    The production version reads the active flight from the airline application,
    captures several Excel workbooks, validates the package for the route type,
    exports each workbook to PDF and merges the final set. Those company-specific
    menu mappings and document templates are not part of the public snapshot.
    """
    print("Flight Task portfolio snapshot")
    print("Connect this module to the authorized airline desktop environment to run it.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
