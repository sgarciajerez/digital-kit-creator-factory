from __future__ import annotations

import shutil
import time
import ctypes
from pathlib import Path

import imageio.v2 as imageio
import mss
import numpy as np
import pyautogui
import pyperclip
import win32com.client
import win32gui
from PIL import Image


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "autopilot_3_Planificador_de_presupuesto_familiar_0.1.0.xlsx"
OUT_DIR = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "video"
DEMO_FILE = OUT_DIR / "planificador-demo.xlsx"
VIDEO_FILE = OUT_DIR / "planificador-demo-interactivo.mp4"
FPS = 24
TARGET_SIZE = (1920, 960)


def capture_frame(sct: mss.mss, monitor: dict[str, int]) -> np.ndarray:
    shot = Image.frombytes("RGB", (monitor["width"], monitor["height"]), sct.grab(monitor).rgb)
    shot = shot.resize(TARGET_SIZE, Image.Resampling.LANCZOS)
    return np.asarray(shot)


def record_for(writer, sct: mss.mss, monitor: dict[str, int], seconds: float) -> None:
    end = time.perf_counter() + seconds
    while time.perf_counter() < end:
        writer.append_data(capture_frame(sct, monitor))
        time.sleep(1 / FPS)


def activate(excel, workbook, sheet_name: str, cell: str = "A1") -> None:
    try:
        win32gui.ShowWindow(int(excel.Hwnd), 3)
        try:
            win32gui.SetForegroundWindow(int(excel.Hwnd))
        except Exception:
            pass
    except Exception:
        pass
    sheet = workbook.Worksheets(sheet_name)
    sheet.Activate()
    sheet.Range(cell).Select()
    excel.ActiveWindow.ScrollRow = 1
    excel.ActiveWindow.ScrollColumn = 1
    excel.ActiveWindow.Zoom = 90
    time.sleep(0.35)


def visible_type(excel, workbook, sheet_name: str, cell: str, value: str, writer, sct, monitor) -> None:
    """Type into the actual Excel cell while the screen recorder is running."""
    activate(excel, workbook, sheet_name, cell)
    pyautogui.press("delete")
    for character in str(value):
        if character.isascii():
            pyautogui.write(character, interval=0.035)
        else:
            pyperclip.copy(character)
            pyautogui.hotkey("ctrl", "v")
        record_for(writer, sct, monitor, 0.045)
    pyautogui.press("enter")
    excel.CalculateFull()
    record_for(writer, sct, monitor, 0.55)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, DEMO_FILE)
    excel = None
    workbook = None
    console = 0
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.WindowState = -4137
        workbook = excel.Workbooks.Open(str(DEMO_FILE), ReadOnly=False)
        workbook.Activate()
        time.sleep(2.2)
        try:
            win32gui.ShowWindow(int(excel.Hwnd), 3)
            win32gui.SetForegroundWindow(int(excel.Hwnd))
            excel.DisplayFullScreen = True
        except Exception:
            pass
        console = ctypes.windll.kernel32.GetConsoleWindow()
        if console:
            ctypes.windll.user32.ShowWindow(console, 0)
        pyautogui.hotkey("win", "d")
        time.sleep(0.35)
        try:
            win32gui.ShowWindow(int(excel.Hwnd), 3)
            win32gui.BringWindowToTop(int(excel.Hwnd))
            win32gui.SetForegroundWindow(int(excel.Hwnd))
        except Exception:
            pass
        try:
            win32gui.SetForegroundWindow(int(excel.Hwnd))
        except Exception:
            pass
        excel.CalculateFull()

        monitor = mss.mss().monitors[1]
        pyautogui.PAUSE = 0.05
        pyautogui.moveTo(monitor["width"] // 2, monitor["height"] // 2, duration=0.2)
        config_name = next(name.Name for name in workbook.Worksheets if str(name.Name).startswith("Configuraci"))

        with imageio.get_writer(VIDEO_FILE, fps=FPS, codec="libx264", quality=7, pixelformat="yuv420p", macro_block_size=1) as writer:
            with mss.mss() as sct:
                activate(excel, workbook, config_name, "A1")
                record_for(writer, sct, monitor, 1.0)
                visible_type(excel, workbook, config_name, "B6", "Ana", writer, sct, monitor)
                visible_type(excel, workbook, config_name, "B7", "Carlos", writer, sct, monitor)

                activate(excel, workbook, "Presupuesto", "B2")
                record_for(writer, sct, monitor, 0.8)
                for cell, value in (("B6", "950"), ("C6", "930"), ("B9", "90"), ("C9", "110"), ("B12", "260"), ("C12", "245")):
                    visible_type(excel, workbook, "Presupuesto", cell, value, writer, sct, monitor)
                activate(excel, workbook, "Vista previa", "A1")
                record_for(writer, sct, monitor, 1.7)

                for cell, value in (("A6", "05/01/2026"), ("B6", "Compra semanal"), ("C6", "Alimentación"), ("D6", "Ana"), ("E6", "86")):
                    visible_type(excel, workbook, "Gastos compartidos", cell, value, writer, sct, monitor)
                activate(excel, workbook, "Gastos compartidos", "E6")
                record_for(writer, sct, monitor, 1.1)
                activate(excel, workbook, "Vista previa", "A1")
                record_for(writer, sct, monitor, 1.5)
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.DisplayAlerts = True
            excel.Quit()
        if console:
            ctypes.windll.user32.ShowWindow(console, 5)
        if DEMO_FILE.exists():
            DEMO_FILE.unlink()
    print(VIDEO_FILE)


if __name__ == "__main__":
    main()
