from __future__ import annotations

import importlib
import os
import platform
import sys

REQUIRED_MODULES = (
    "airportsdata",
    "pypdf",
    "pythoncom",
    "win32com.client",
    "pywinauto",
    "rapidfuzz",
    "unidecode",
)


def check_modules() -> list[tuple[str, bool, str]]:
    results: list[tuple[str, bool, str]] = []
    for module_name in REQUIRED_MODULES:
        try:
            importlib.import_module(module_name)
            results.append((module_name, True, "installed"))
        except Exception as exc:
            results.append((module_name, False, str(exc)))
    return results


def check_excel_com() -> tuple[bool, str]:
    if platform.system() != "Windows":
        return False, "Excel COM is only available on Windows"

    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        version = str(excel.Version)
        excel.Quit()
        return True, f"Microsoft Excel COM available (version {version})"
    except Exception as exc:
        return False, f"Microsoft Excel COM unavailable: {exc}"


def main() -> int:
    print("Flight Task Automation diagnostics")
    print("=" * 40)
    print(f"OS: {platform.platform()}")
    print(f"Python: {sys.version.split()[0]}")
    print(f"Executable: {sys.executable}")
    print(
        "Network output: "
        + (os.getenv("FLIGHT_TASK_NETWORK_OUTPUT") or "not configured")
    )
    print()

    failed = False
    print("Python dependencies:")
    for module_name, ok, details in check_modules():
        marker = "OK" if ok else "FAIL"
        print(f"  [{marker}] {module_name}: {details}")
        failed = failed or not ok

    print()
    excel_ok, excel_details = check_excel_com()
    print(f"Excel: {'OK' if excel_ok else 'WARN'} - {excel_details}")

    if not sys.version.startswith("3.12"):
        print("WARN - Python 3.12.x is the tested/recommended portfolio runtime.")

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
