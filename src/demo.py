from __future__ import annotations

from pathlib import Path

from pypdf import PdfWriter

import worker


def make_blank_pdf(path: Path, width: float = 595, height: float = 842) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=width, height=height)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("wb") as handle:
        writer.write(handle)


def main() -> int:
    route_text = "SVO - LED"
    route_points = worker.parse_route(route_text)
    search_strings, metadata = worker.build_airport_index()
    resolved = [
        worker.resolve_airport(point, search_strings, metadata)
        for point in route_points
    ]

    print("Flight Task Automation — public demo")
    print(f"Route: {route_text}")
    for airport in resolved:
        if airport is None:
            print("  unresolved airport")
        else:
            print(
                f"  {airport.original} -> {airport.iata} / {airport.icao} "
                f"({airport.city}, {airport.country})"
            )

    output_dir = Path(__file__).resolve().parents[1] / "demo_output"
    part_a = output_dir / "01_flight_sheet.pdf"
    part_b = output_dir / "02_supporting_document.pdf"
    final_pdf = output_dir / "sample_flight_task.pdf"

    make_blank_pdf(part_a)
    make_blank_pdf(part_b)
    worker.merge_pdfs([part_a, part_b], final_pdf)

    print(f"Demo package created: {final_pdf}")
    print("This demo uses synthetic blank documents and no company data.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
