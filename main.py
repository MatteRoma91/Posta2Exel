# Posta2Excel - Interfaccia grafica
#
# Installazione dipendenze (eseguire nel terminale prima dell'avvio):
#   pip install customtkinter pandas openpyxl libpff-python
#
# Su Windows, se libpff-python fallisce la compilazione:
#   pip install libpff-python-windows
#
# Avvio applicazione:
#   python main.py

from __future__ import annotations

import os
import queue
import threading
from tkinter import filedialog

import customtkinter as ctk

import backend

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class Posta2ExcelApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Posta2Excel - Estrattore Contatti Professionale")
        self.geometry("600x450")
        self.resizable(False, False)

        self.source_path = ""
        self.dest_path = ""
        self.event_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.is_running = False

        self._build_ui()
        self._validate_ready()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(
            self,
            text="Posta2Excel",
            font=ctk.CTkFont(family="Arial", size=24, weight="bold"),
            text_color="#00BFFF",
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 16), sticky="n")

        self.source_button = ctk.CTkButton(
            self,
            text="Seleziona file .PST / .OST",
            command=self._select_source,
        )
        self.source_button.grid(row=1, column=0, padx=20, pady=(0, 6), sticky="ew")

        self.source_label = ctk.CTkLabel(
            self,
            text="Nessun file selezionato",
            anchor="w",
            wraplength=540,
        )
        self.source_label.grid(row=2, column=0, padx=20, pady=(0, 12), sticky="ew")

        self.dest_button = ctk.CTkButton(
            self,
            text="Seleziona Cartella di Destinazione",
            command=self._select_destination,
        )
        self.dest_button.grid(row=3, column=0, padx=20, pady=(0, 6), sticky="ew")

        self.dest_label = ctk.CTkLabel(
            self,
            text="Nessuna cartella selezionata",
            anchor="w",
            wraplength=540,
        )
        self.dest_label.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="ew")

        self.extract_button = ctk.CTkButton(
            self,
            text="AVVIA ESTRAZIONE",
            height=48,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self._start_extraction,
            state="disabled",
        )
        self.extract_button.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=6, column=0, padx=20, pady=(0, 8), sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            self,
            text="Stato: Pronto",
            anchor="w",
            wraplength=540,
        )
        self.status_label.grid(row=7, column=0, padx=20, pady=(0, 20), sticky="ew")

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
            title="Seleziona file PST o OST",
            filetypes=[
                ("File Outlook", "*.pst *.ost"),
                ("PST", "*.pst"),
                ("OST", "*.ost"),
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

    def _validate_ready(self) -> None:
        source_ok = bool(self.source_path) and os.path.isfile(self.source_path)
        dest_ok = bool(self.dest_path) and os.path.isdir(self.dest_path)
        if source_ok and dest_ok and not self.is_running:
            self.extract_button.configure(state="normal")
        else:
            self.extract_button.configure(state="disabled")

    def _disable_controls(self) -> None:
        self.is_running = True
        self.extract_button.configure(state="disabled")
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
            self.status_label.configure(text="Errore: file PST/OST non valido.")
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
