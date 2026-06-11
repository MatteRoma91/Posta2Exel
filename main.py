# Posta2Excel - Interfaccia grafica
#
# Installazione dipendenze (eseguire nel terminale prima dell'avvio):
#   pip install customtkinter pandas openpyxl
#
# Solo per file Outlook (.pst / .ost):
#   pip install libpff-python
#   Su Windows, se libpff-python fallisce: pip install libpff-python-windows
#
# I file .mbox (Gmail Takeout, Thunderbird, Apple Mail) usano la libreria standard mailbox.
#
# Avvio applicazione:
#   python main.py
#
# Build PyInstaller (Windows):
#   pyinstaller --onefile --windowed --icon=app_icon.ico ^
#     --add-data "app_icon.ico;." --add-data "app_icon.png;." main.py

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

import backend

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

WINDOW_WIDTH = 650
WINDOW_HEIGHT = 500

BTN_SELECT_FG = "#3D4451"
BTN_SELECT_HOVER = "#4A5060"
BTN_START_DISABLED_FG = "#2B2B2B"
BTN_START_DISABLED_TEXT = "#666666"
BTN_START_ENABLED_FG = "#00C896"
BTN_START_ENABLED_HOVER = "#00E0A8"
PROGRESS_COLOR = "#00BFFF"
TITLE_COLOR = "#00BFFF"


def resource_path(relative_path: str) -> str:
    """Risolve percorsi asset in sviluppo e nel bundle PyInstaller (sys._MEIPASS)."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


class Posta2ExcelApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Posta2Excel - Estrattore Contatti Professionale")
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(False, False)

        self.source_path = ""
        self.dest_path = ""
        self.event_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.is_running = False
        self._icon_ref: tk.PhotoImage | None = None

        self._set_window_icon()
        self._center_window()
        self._build_ui()
        self._validate_ready()
        self.after(100, self._poll_queue)

    def _set_window_icon(self) -> None:
        try:
            if sys.platform == "win32":
                ico_path = resource_path("app_icon.ico")
                if os.path.isfile(ico_path):
                    self.iconbitmap(ico_path)
                elif os.path.isfile(resource_path("app_icon.png.jpg")):
                    pass
            elif sys.platform == "darwin":
                png_path = resource_path("app_icon.png")
                if not os.path.isfile(png_path) and os.path.isfile(resource_path("app_icon.png.jpg")):
                    png_path = resource_path("app_icon.png.jpg")
                if os.path.isfile(png_path):
                    self._icon_ref = tk.PhotoImage(file=png_path)
                    self.iconphoto(False, self._icon_ref)
        except Exception:
            pass

    def _center_window(self) -> None:
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - WINDOW_WIDTH) // 2
        y = (screen_h - WINDOW_HEIGHT) // 2
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(
            self,
            text="Posta2Excel",
            font=ctk.CTkFont(family="Arial", size=24, weight="bold"),
            text_color=TITLE_COLOR,
        )
        self.title_label.grid(row=0, column=0, padx=28, pady=(24, 18), sticky="n")

        self.source_button = ctk.CTkButton(
            self,
            text="Seleziona file .PST / .OST / .MBOX",
            command=self._select_source,
            fg_color=BTN_SELECT_FG,
            hover_color=BTN_SELECT_HOVER,
        )
        self.source_button.grid(row=1, column=0, padx=28, pady=(0, 8), sticky="ew")

        self.source_label = ctk.CTkLabel(
            self,
            text="Nessun file selezionato",
            anchor="w",
            wraplength=580,
        )
        self.source_label.grid(row=2, column=0, padx=28, pady=(0, 16), sticky="ew")

        self.dest_button = ctk.CTkButton(
            self,
            text="Seleziona Cartella di Destinazione",
            command=self._select_destination,
            fg_color=BTN_SELECT_FG,
            hover_color=BTN_SELECT_HOVER,
        )
        self.dest_button.grid(row=3, column=0, padx=28, pady=(0, 8), sticky="ew")

        self.dest_label = ctk.CTkLabel(
            self,
            text="Nessuna cartella selezionata",
            anchor="w",
            wraplength=580,
        )
        self.dest_label.grid(row=4, column=0, padx=28, pady=(0, 24), sticky="ew")

        self.extract_button = ctk.CTkButton(
            self,
            text="AVVIA ESTRAZIONE",
            height=48,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self._start_extraction,
            state="disabled",
            fg_color=BTN_START_DISABLED_FG,
            hover_color=BTN_START_DISABLED_FG,
            text_color=BTN_START_DISABLED_TEXT,
        )
        self.extract_button.grid(row=5, column=0, padx=28, pady=(0, 24), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self, progress_color=PROGRESS_COLOR)
        self.progress_bar.grid(row=6, column=0, padx=28, pady=(0, 10), sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            self,
            text="Stato: Pronto",
            anchor="w",
            wraplength=580,
        )
        self.status_label.grid(row=7, column=0, padx=28, pady=(0, 24), sticky="ew")

    def _truncate_path(self, path: str, max_len: int = 50) -> str:
        if len(path) <= max_len:
            return path
        keep = max_len - 3
        left = keep // 2
        right = keep - left
        return f"{path[:left]}...{path[-right:]}"

    def _select_source(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self,
            title="Seleziona file PST, OST o MBOX",
            filetypes=[
                ("Archivi email", "*.pst *.ost *.mbox"),
                ("Outlook", "*.pst *.ost"),
                ("MBOX (Gmail/Thunderbird/Apple)", "*.mbox"),
                ("Tutti i file", "*.*"),
            ],
        )
        if selected:
            self.source_path = selected
            self.source_label.configure(text=self._truncate_path(selected))
            self._validate_ready()

    def _select_destination(self) -> None:
        selected = filedialog.askdirectory(
            parent=self,
            title="Seleziona cartella di destinazione",
        )
        if selected:
            self.dest_path = selected
            self.dest_label.configure(text=self._truncate_path(selected))
            self._validate_ready()

    def _set_extract_button_ready(self, ready: bool) -> None:
        if ready:
            self.extract_button.configure(
                state="normal",
                fg_color=BTN_START_ENABLED_FG,
                hover_color=BTN_START_ENABLED_HOVER,
                text_color="#FFFFFF",
            )
        else:
            self.extract_button.configure(
                state="disabled",
                fg_color=BTN_START_DISABLED_FG,
                hover_color=BTN_START_DISABLED_FG,
                text_color=BTN_START_DISABLED_TEXT,
            )

    def _validate_ready(self) -> None:
        source_ok = bool(self.source_path) and os.path.isfile(self.source_path)
        dest_ok = bool(self.dest_path) and os.path.isdir(self.dest_path)
        if source_ok and dest_ok and not self.is_running:
            self._set_extract_button_ready(True)
        else:
            self._set_extract_button_ready(False)

    def _disable_controls(self) -> None:
        self.is_running = True
        self._set_extract_button_ready(False)
        self.source_button.configure(state="disabled")
        self.dest_button.configure(state="disabled")

    def _enable_controls(self) -> None:
        self.is_running = False
        self.source_button.configure(state="normal")
        self.dest_button.configure(state="normal")
        self._validate_ready()

    def _start_extraction(self) -> None:
        if self.is_running:
            return
        if not os.path.isfile(self.source_path):
            self.status_label.configure(text="Errore: file sorgente non valido.")
            return
        if not os.path.isdir(self.dest_path):
            self.status_label.configure(text="Errore: cartella di destinazione non valida.")
            return

        self.progress_bar.set(0)
        self.status_label.configure(text="Stato: Avvio estrazione...")
        self._disable_controls()

        self.worker = threading.Thread(
            target=backend.run_extraction,
            args=(self.source_path, self.dest_path, self.event_queue),
            daemon=True,
        )
        self.worker.start()

    def _poll_queue(self) -> None:
        while True:
            try:
                event = self.event_queue.get_nowait()
            except queue.Empty:
                break

            event_type = event.get("type")
            if event_type == "progress":
                self.progress_bar.set(float(event.get("value", 0)))
            elif event_type == "status":
                self.status_label.configure(text=event.get("text", ""))
            elif event_type == "done":
                self.progress_bar.set(1.0)
                self._enable_controls()
            elif event_type == "error":
                self.progress_bar.set(0)
                self.status_label.configure(text=event.get("text", "Errore sconosciuto."))
                self._enable_controls()

        self.after(100, self._poll_queue)


def main() -> None:
    app = Posta2ExcelApp()
    app.mainloop()


if __name__ == "__main__":
    main()
