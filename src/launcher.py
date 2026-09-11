# -*- coding: utf-8 -*-
"""Flight Task Automation launcher.

Public portfolio version. Production-only menu mappings, internal paths and
company document identifiers are intentionally excluded.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog

APP_DIR = Path(__file__).resolve().parent
WORKER_PATH = APP_DIR / "worker.py"
CONFIG_PATH = APP_DIR / "flight_task_config.json"
RESULT_PATH = APP_DIR / "_last_result.json"


class FlightTaskApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Flight Task Automation")
        self.root.geometry("420x260")
        self.root.resizable(False, False)

        self.use_network_output = tk.BooleanVar(value=False)
        self._build_ui()

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=28, pady=24)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Flight Task Automation",
            font=("Segoe UI", 16, "bold"),
        ).pack(pady=(0, 8))

        tk.Label(
            frame,
            text="Portfolio launcher for airline document-package automation",
            font=("Segoe UI", 9),
        ).pack(pady=(0, 18))

        tk.Checkbutton(
            frame,
            text="Use configured network output",
            variable=self.use_network_output,
        ).pack(anchor="w", pady=(0, 14))

        tk.Button(
            frame,
            text="Build one flight task",
            command=self.run_single,
            width=30,
        ).pack(pady=5)

        tk.Button(
            frame,
            text="Build batch",
            command=self.run_batch,
            width=30,
        ).pack(pady=5)

    def _write_config(self) -> None:
        CONFIG_PATH.write_text(
            json.dumps(
                {"use_network_output": bool(self.use_network_output.get())},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    def _clear_result(self) -> None:
        try:
            RESULT_PATH.unlink()
        except FileNotFoundError:
            pass

    def _run_worker(self) -> dict:
        if not WORKER_PATH.exists():
            raise RuntimeError(f"Worker module not found: {WORKER_PATH}")

        self._write_config()
        self._clear_result()

        process = subprocess.run(
            [sys.executable, str(WORKER_PATH)],
            cwd=str(APP_DIR),
            check=False,
        )
        if process.returncode != 0:
            raise RuntimeError(f"Worker exited with code {process.returncode}")

        if RESULT_PATH.exists():
            return json.loads(RESULT_PATH.read_text(encoding="utf-8"))

        # The sanitized public worker stops at the proprietary integration
        # boundary, so a result file is not guaranteed outside production.
        return {"status": "portfolio_snapshot"}

    def run_single(self) -> None:
        try:
            result = self._run_worker()
            messagebox.showinfo(
                "Flight Task Automation",
                self._format_result(result),
            )
        except Exception as exc:
            messagebox.showerror("Flight Task Automation", str(exc))

    def run_batch(self) -> None:
        count = simpledialog.askinteger(
            "Batch processing",
            "How many flight tasks should be processed?",
            parent=self.root,
            minvalue=1,
            maxvalue=500,
        )
        if count is None:
            return

        completed = []
        try:
            for _ in range(count):
                completed.append(self._run_worker())
        except Exception as exc:
            messagebox.showerror("Flight Task Automation", str(exc))
            return

        messagebox.showinfo(
            "Flight Task Automation",
            f"Completed worker runs: {len(completed)}",
        )

    @staticmethod
    def _format_result(result: dict) -> str:
        if result.get("status") == "portfolio_snapshot":
            return (
                "Portfolio snapshot launched successfully.\n\n"
                "The proprietary airline integration layer is intentionally "
                "not included in the public repository."
            )

        flight = result.get("flight", "")
        route = result.get("route", "")
        pdf = result.get("pdf", "")
        return f"Flight: {flight}\nRoute: {route}\nOutput: {pdf}"

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    FlightTaskApp().run()


if __name__ == "__main__":
    main()
