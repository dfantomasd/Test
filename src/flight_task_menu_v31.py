# -*- coding: utf-8 -*-
"""
Flight Task v31 — меню запуска.

При запуске показывает две кнопки:
1. Сделать одно задание
2. Сделать N заданий

Для пакетного режима после каждого успешно созданного задания:
- закрывает оставшиеся меню F11;
- возвращается в Aviabit;
- вызывает Статус (F5);
- в окне подтверждения выбирает "Да";
- если это не последний рейс — нажимает Down и переходит к следующему рейсу;
- запускает формирование следующего задания.

Финального сообщения "всё сделано" нет.
"""

from __future__ import annotations

import ctypes
import json
import os
import subprocess
import sys
import time
import tkinter as tk
from tkinter import simpledialog, messagebox
from pathlib import Path
from PIL import Image, ImageTk

import win32gui
import win32process
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


AVIABIT_TITLE_MARKERS = (
    "Планирование АК",
    "Полетные задания",
    "Авиабит",
    "Aviabit",
)

DOCUMENT_MENU_TITLE = "Выбор режима работы"
FLAG_WAIT_SECONDS = 5.0
CONFIG_PATH = Path(__file__).resolve().parent / "flight_task_config.json"
USE_Z_DRIVE = False


def normalize_text(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").split()).strip()


def worker_path() -> Path:
    return Path(__file__).resolve().parent / "flight_task_worker_v31.py"


def result_path() -> Path:
    return Path(__file__).resolve().parent / "_last_result.json"


def clear_last_result() -> None:
    try:
        result_path().unlink()
    except FileNotFoundError:
        pass
    except Exception:
        pass


def read_last_result() -> dict:
    path = result_path()
    if not path.exists():
        raise RuntimeError(
            "Рабочий скрипт завершился успешно, но не оставил данные "
            "о сформированном задании."
        )

    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as exc:
        raise RuntimeError(
            f"Не удалось прочитать результат последнего задания: {exc}"
        ) from exc

    if not data.get("flight") or not data.get("route"):
        raise RuntimeError(
            "В результате отсутствует номер рейса или направление."
        )

    return data


def find_aviabit_window():
    candidates = []

    for backend in ("win32", "uia"):
        try:
            for w in Desktop(backend=backend).windows(visible_only=True):
                try:
                    title = normalize_text(w.window_text())
                except Exception:
                    continue

                if not title:
                    continue

                score = sum(
                    marker.lower() in title.lower()
                    for marker in AVIABIT_TITLE_MARKERS
                )
                if score:
                    candidates.append((score, backend, w, title))
        except Exception:
            pass

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][2]


def focus_window(window) -> None:
    try:
        if window.is_minimized():
            window.restore()
    except Exception:
        pass

    try:
        window.set_focus()
        time.sleep(0.15)
        return
    except Exception:
        pass

    try:
        hwnd = int(window.handle)
        ctypes.windll.user32.ShowWindow(hwnd, 9)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.15)
    except Exception as exc:
        raise RuntimeError(f"Не удалось активировать Aviabit: {exc}") from exc


def close_document_menus() -> None:
    for _ in range(5):
        found = False

        for backend in ("win32", "uia"):
            try:
                windows = Desktop(backend=backend).windows(visible_only=True)
            except Exception:
                continue

            for w in windows:
                try:
                    title = normalize_text(w.window_text())
                except Exception:
                    continue

                if DOCUMENT_MENU_TITLE.lower() not in title.lower():
                    continue

                found = True
                try:
                    w.set_focus()
                except Exception:
                    pass

                send_keys("{ESC}", pause=0.03)
                time.sleep(0.10)

        if not found:
            break


def aviabit_process_id(aviabit) -> int | None:
    try:
        hwnd = int(aviabit.handle)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid)
    except Exception:
        return None


def find_yes_dialog(aviabit_pid: int | None):
    for backend in ("win32", "uia"):
        try:
            windows = Desktop(backend=backend).windows(visible_only=True)
        except Exception:
            continue

        for w in windows:
            try:
                hwnd = int(w.handle)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if aviabit_pid is not None and int(pid) != aviabit_pid:
                    continue
            except Exception:
                if aviabit_pid is not None:
                    continue

            try:
                descendants = w.descendants()
            except Exception:
                descendants = []

            for control in descendants:
                try:
                    ctext = normalize_text(control.window_text()).replace("&", "")
                except Exception:
                    continue

                if ctext.lower() in ("да", "yes"):
                    return w, control

    return None, None


def confirm_flag_yes() -> None:
    close_document_menus()

    aviabit = find_aviabit_window()
    if aviabit is None:
        raise RuntimeError("Не найдено окно Aviabit для установки флага.")

    focus_window(aviabit)
    pid = aviabit_process_id(aviabit)
    dialog, yes_button = find_yes_dialog(pid)

    if yes_button is None:
        print("      Ставлю флаг: F5 -> Да", flush=True)
        send_keys("{F5}", pause=0.05)

        deadline = time.time() + FLAG_WAIT_SECONDS
        while time.time() < deadline:
            dialog, yes_button = find_yes_dialog(pid)
            if yes_button is not None:
                break
            time.sleep(0.10)

    if yes_button is None:
        raise RuntimeError(
            "После F5 не найдено окно Aviabit с кнопкой «Да» для установки флага."
        )

    try:
        if dialog is not None:
            dialog.set_focus()
    except Exception:
        pass

    try:
        yes_button.click_input()
    except Exception:
        try:
            yes_button.click()
        except Exception as exc:
            raise RuntimeError(
                f"Не удалось нажать «Да» в окне установки флага: {exc}"
            ) from exc

    time.sleep(0.30)


def move_to_next_flight() -> None:
    aviabit = find_aviabit_window()
    if aviabit is None:
        raise RuntimeError("Не найден Aviabit для перехода к следующему рейсу.")

    focus_window(aviabit)
    send_keys("{DOWN}", pause=0.05)
    time.sleep(0.20)


def run_worker() -> int:
    worker = worker_path()
    if not worker.exists():
        raise RuntimeError(f"Не найден рабочий скрипт: {worker}")

    clear_last_result()

    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            json.dump({"use_z_drive": USE_Z_DRIVE}, f)
    except Exception as exc:
        raise RuntimeError(
            f"Не удалось сохранить настройку места сохранения: {exc}"
        ) from exc

    result = subprocess.run(
        [sys.executable, str(worker)],
        cwd=str(worker.parent),
    )
    return int(result.returncode)


def run_single() -> None:
    root.withdraw()
    code = run_worker()
    root.deiconify()
    root.lift()
    root.focus_force()
    if code != 0:
        raise SystemExit(code)

    try:
        result = read_last_result()
        flight = str(result.get("flight", "")).strip()
        route = str(result.get("route", "")).strip()
        messagebox.showinfo(
            "Flight Task — сформированное задание",
            f"Сформировано задание:\n{flight} — {route}",
        )
    except Exception as exc:
        messagebox.showinfo(
            "Flight Task — сформированное задание",
            f"Задание сформировано, но не удалось прочитать номер рейса:\n\n{exc}",
        )


def run_batch() -> None:
    count = simpledialog.askinteger(
        "Количество заданий",
        "Сколько заданий сформировать подряд?",
        parent=root,
        minvalue=1,
        maxvalue=500,
    )
    if count is None:
        return

    root.destroy()
    completed = []

    for index in range(1, count + 1):
        code = run_worker()
        if code != 0:
            raise SystemExit(code)

        result = read_last_result()
        completed.append(result)
        confirm_flag_yes()
        if index < count:
            move_to_next_flight()

    lines = []
    for item in completed:
        flight = str(item.get("flight", "")).strip()
        route = str(item.get("route", "")).strip()
        lines.append(f"{flight} — {route}")

    messagebox.showinfo(
        "Flight Task — сформированные задания",
        "Сформированные задания: " + f"{len(completed)}\n\n" + "\n".join(lines),
    )


WINDOW_W = 320
WINDOW_H = 290
BACKGROUND_FILE = "menu_background.png"
LOGO_FULL_FILE = "flight_task_logo_full.png"
ICON_FILE = "flight_task_icon.png"
ICON_ICO_FILE = "flight_task_icon.ico"

BUTTON_GREEN = "#0b4f3f"
BUTTON_GREEN_HOVER = "#17614f"
BUTTON_TEXT = "#ffffff"
BUTTON_BORDER = "#d7e1de"
LABEL_GREEN = "#0b4f3f"


def center_window(window, width: int, height: int) -> None:
    window.update_idletasks()
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = max(0, (screen_w - width) // 2)
    y = max(0, (screen_h - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def rounded_rect(canvas, x1, y1, x2, y2, radius=8, **kwargs):
    points = [
        x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
        x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
        x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


def make_canvas_button(canvas, center_x, center_y, width, height, text, command):
    x1 = center_x - width // 2
    y1 = center_y - height // 2
    x2 = center_x + width // 2
    y2 = center_y + height // 2
    shadow = rounded_rect(canvas, x1 + 1, y1 + 1, x2 + 1, y2 + 1, radius=8, fill="#81928d", outline="")
    body = rounded_rect(canvas, x1, y1, x2, y2, radius=8, fill=BUTTON_GREEN, outline=BUTTON_BORDER, width=1)
    label = canvas.create_text(center_x, center_y, text=text, fill=BUTTON_TEXT, font=("Segoe UI", 8, "bold"))
    tag = f"button_{body}"
    for item in (shadow, body, label):
        canvas.addtag_withtag(tag, item)
    canvas.tag_bind(tag, "<Enter>", lambda _e: (canvas.itemconfig(body, fill=BUTTON_GREEN_HOVER), canvas.config(cursor="hand2")))
    canvas.tag_bind(tag, "<Leave>", lambda _e: (canvas.itemconfig(body, fill=BUTTON_GREEN), canvas.config(cursor="")))
    canvas.tag_bind(tag, "<Button-1>", lambda _e: command())


def toggle_z_drive():
    global USE_Z_DRIVE
    USE_Z_DRIVE = not USE_Z_DRIVE
    draw_z_checkbox()


def draw_z_checkbox():
    canvas.itemconfig(z_box, fill=("#0b4f3f" if USE_Z_DRIVE else "white"))


root = tk.Tk()
root.title("Flight Task")
root.resizable(False, False)
center_window(root, WINDOW_W, WINDOW_H)

app_icon_image = None
icon_path = Path(__file__).resolve().parent / ICON_FILE
icon_ico_path = Path(__file__).resolve().parent / ICON_ICO_FILE
try:
    if icon_path.exists():
        app_icon_image = ImageTk.PhotoImage(Image.open(icon_path).convert("RGBA"))
        root.iconphoto(True, app_icon_image)
except Exception:
    pass

try:
    if icon_ico_path.exists():
        root.iconbitmap(default=str(icon_ico_path))
except Exception:
    pass

bg_path = Path(__file__).resolve().parent / BACKGROUND_FILE
if not bg_path.exists():
    messagebox.showerror("Flight Task", f"Не найден фон меню:\n{bg_path}")
    raise SystemExit(1)

original_bg = Image.open(bg_path).convert("RGB")
display_bg = original_bg.resize((400, 300), Image.Resampling.LANCZOS)
background_image = ImageTk.PhotoImage(display_bg)

canvas = tk.Canvas(root, width=WINDOW_W, height=WINDOW_H, highlightthickness=0, borderwidth=0)
canvas.pack(fill="both", expand=True)
canvas.create_image(-40, -55, image=background_image, anchor="nw")

z_box = canvas.create_rectangle(48, 182, 65, 199, outline="#0b4f3f", width=2, fill="white")
canvas.create_text(78, 190, text="Разместить задания на диске Z", anchor="w", fill="#0b4f3f", font=("Segoe UI", 8, "bold"), tags=("z_toggle",))
for item in (z_box, "z_toggle"):
    canvas.tag_bind(item, "<Button-1>", lambda _e: toggle_z_drive())

draw_z_checkbox()
make_canvas_button(canvas, WINDOW_W // 2, 228, 220, 24, "Сделать одно задание", run_single)
make_canvas_button(canvas, WINDOW_W // 2, 257, 220, 24, "Сделать N заданий", run_batch)
root.bind("<Escape>", lambda _e: root.destroy())
root.mainloop()
