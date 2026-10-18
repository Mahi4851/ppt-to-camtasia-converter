from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "PPT to Camtasia"

BG = "#F5F7FB"
CARD = "#FFFFFF"
TEXT = "#172033"
MUTED = "#6B7280"
ACCENT = "#4F46E5"
ACCENT_HOVER = "#4338CA"
BORDER = "#E4E7EC"
SUCCESS = "#15803D"
ERROR = "#B42318"


class ConverterUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("650x390")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.script_dir = Path(__file__).resolve().parent
        self.converter_script = self.script_dir / "proper_one_click.py"

        self.selected_path: Path | None = None
        self.last_output_folder: Path | None = None
        self.running = False
        self.last_error_lines: list[str] = []

        self.file_name = tk.StringVar(value="No PowerPoint selected")
        self.file_hint = tk.StringVar(value="Choose a .pptx file to begin")
        self.status = tk.StringVar(value="Ready")

        self._configure_style()
        self._build_ui()
        self._center_window()

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("FileName.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 11, "bold"))
        style.configure("FileHint.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            foreground="#FFFFFF",
            background=ACCENT,
            borderwidth=0,
            padding=(18, 10),
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_HOVER), ("disabled", "#A5B4FC")],
            foreground=[("disabled", "#F8FAFC")],
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            foreground=TEXT,
            background="#EEF2F7",
            borderwidth=0,
            padding=(16, 10),
        )
        style.map("Secondary.TButton", background=[("active", "#E5EAF1")])
        style.configure(
            "Clean.Horizontal.TProgressbar",
            troughcolor="#E8EBF2",
            background=ACCENT,
            bordercolor="#E8EBF2",
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            thickness=7,
        )

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=(30, 26, 30, 22))
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="PowerPoint to Camtasia", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Select a presentation and convert it with the existing editable Camtasia workflow.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 20))

        card_shell = tk.Frame(outer, bg=BORDER, bd=0)
        card_shell.pack(fill="x")
        card = ttk.Frame(card_shell, style="Card.TFrame", padding=(18, 16))
        card.pack(fill="x", padx=1, pady=1)

        file_row = ttk.Frame(card, style="Card.TFrame")
        file_row.pack(fill="x")

        file_text = ttk.Frame(file_row, style="Card.TFrame")
        file_text.pack(side="left", fill="x", expand=True)
        ttk.Label(file_text, textvariable=self.file_name, style="FileName.TLabel").pack(anchor="w")
        ttk.Label(file_text, textvariable=self.file_hint, style="FileHint.TLabel").pack(anchor="w", pady=(3, 0))

        self.select_button = ttk.Button(
            file_row,
            text="Choose PPT",
            style="Secondary.TButton",
            command=self.browse,
            cursor="hand2",
        )
        self.select_button.pack(side="right", padx=(16, 0))

        self.progress = ttk.Progressbar(
            outer,
            mode="indeterminate",
            style="Clean.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(22, 7))

        status_row = ttk.Frame(outer, style="App.TFrame")
        status_row.pack(fill="x")
        self.status_dot = tk.Label(status_row, text="●", fg="#98A2B3", bg=BG, font=("Segoe UI", 9))
        self.status_dot.pack(side="left")
        self.status_label = ttk.Label(status_row, textvariable=self.status, style="Status.TLabel")
        self.status_label.pack(side="left", padx=(6, 0))

        action_row = ttk.Frame(outer, style="App.TFrame")
        action_row.pack(fill="x", pady=(22, 0))

        self.convert_button = ttk.Button(
            action_row,
            text="Convert to Camtasia",
            style="Accent.TButton",
            command=self.start_conversion,
            state="disabled",
            cursor="hand2",
        )
        self.convert_button.pack(side="left")

        self.open_button = ttk.Button(
            action_row,
            text="Open Folder",
            style="Secondary.TButton",
            command=self.open_output,
            state="disabled",
            cursor="hand2",
        )
        self.open_button.pack(side="left", padx=(10, 0))

        ttk.Label(
            outer,
            text="Your Camtasia project keeps the same conversion behavior as the original tool.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(20, 0))

    def _center_window(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _set_status(self, text: str, kind: str = "neutral") -> None:
        self.status.set(text)
        color = {
            "neutral": "#98A2B3",
            "working": ACCENT,
            "success": SUCCESS,
            "error": ERROR,
        }.get(kind, "#98A2B3")
        self.status_dot.configure(fg=color)

    def browse(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Choose PowerPoint presentation",
            filetypes=[("PowerPoint Presentation", "*.pptx")],
        )
        if not chosen:
            return

        path = Path(chosen)
        self.selected_path = path
        self.last_output_folder = path.parent
        self.file_name.set(path.name)
        self.file_hint.set("PowerPoint presentation selected")
        self.convert_button.configure(state="normal")
        self.open_button.configure(state="disabled")
        self._set_status("Ready to convert", "neutral")

    def start_conversion(self) -> None:
        if self.running:
            return

        source = self.selected_path
        if source is None:
            messagebox.showwarning(APP_TITLE, "Please choose a PowerPoint file first.")
            return
        if not source.exists() or source.suffix.lower() != ".pptx":
            messagebox.showerror(APP_TITLE, "The selected PowerPoint file is no longer available.")
            return

        self.running = True
        self.last_error_lines = []
        self.select_button.configure(state="disabled")
        self.convert_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.progress.start(10)
        self.file_hint.set("Conversion in progress")
        self._set_status("Converting presentation...", "working")

        threading.Thread(target=self._run_conversion, args=(source,), daemon=True).start()

    def _run_conversion(self, source: Path) -> None:
        cmd = [
            sys.executable,
            "-u",
            str(self.converter_script),
            "--source",
            str(source),
            "--output-parent",
            str(source.parent),
        ]

        creationflags = 0
        if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags = subprocess.CREATE_NO_WINDOW

        try:
            process = subprocess.Popen(
                cmd,
                cwd=str(self.script_dir.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=creationflags,
            )
            assert process.stdout is not None
            for line in process.stdout:
                clean = line.strip()
                if clean:
                    self.last_error_lines.append(clean)
                    if len(self.last_error_lines) > 30:
                        self.last_error_lines.pop(0)
            code = process.wait()
        except Exception as exc:
            self.last_error_lines.append(str(exc))
            code = 1

        self.root.after(0, self._finish_conversion, code, source)

    def _finish_conversion(self, code: int, source: Path) -> None:
        self.running = False
        self.progress.stop()
        self.select_button.configure(state="normal")
        self.convert_button.configure(state="normal")

        stem = source.stem or "Untitled"
        project_folder = source.parent / f"{stem}_CAMTASIA_PROJECT"
        project_file = source.parent / f"{stem}.tscproj"
        packaged_project = project_folder / f"{stem}.tscproj"
        zip_file = source.parent / f"{stem}_CAMTASIA_PROJECT.zip"
        self.last_output_folder = source.parent

        success = code == 0 and project_file.exists() and packaged_project.exists() and zip_file.exists()
        if success:
            self.file_hint.set("Conversion completed successfully")
            self._set_status("Camtasia project is ready", "success")
            self.open_button.configure(state="normal")
            messagebox.showinfo(APP_TITLE, "Conversion complete. Your Camtasia project is ready.")
        else:
            self.file_hint.set("Conversion could not be completed")
            self._set_status("Conversion failed", "error")
            self.open_button.configure(state="normal" if source.parent.exists() else "disabled")
            messagebox.showerror(
                APP_TITLE,
                "Conversion could not be completed. Please make sure the PowerPoint file is available and try again.",
            )

    def open_output(self) -> None:
        folder = self.last_output_folder
        if not folder or not folder.exists():
            return
        try:
            if os.name == "nt":
                os.startfile(str(folder))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception:
            messagebox.showerror(APP_TITLE, "The folder could not be opened automatically.")


def main() -> None:
    root = tk.Tk()
    ConverterUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
