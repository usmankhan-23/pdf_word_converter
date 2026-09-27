"""
main.py
-------
Desktop GUI entry point. Built with Tkinter (ships with standard Python —
no extra GUI framework needed), so the app stays lightweight and dependency-
free for the UI layer itself.

Run with:
    python main.py
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import convert_pdf_to_docx, convert_docx_to_pdf, ConversionError

APP_TITLE = "PDF ⇄ Word Converter (Offline)"


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("560x360")
        self.resizable(False, False)

        self.input_path: str | None = None
        self.mode = tk.StringVar(value="pdf2docx")  # or "docx2pdf"

        self._build_ui()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        pad = {"padx": 16, "pady": 8}

        header = tk.Label(
            self, text=APP_TITLE, font=("Segoe UI", 14, "bold")
        )
        header.pack(anchor="w", **pad)

        subtitle = tk.Label(
            self,
            text="Fully local conversion — no internet connection required.",
            fg="#555",
        )
        subtitle.pack(anchor="w", padx=16)

        # Mode selector
        mode_frame = tk.Frame(self)
        mode_frame.pack(anchor="w", **pad)
        tk.Radiobutton(
            mode_frame, text="PDF → Word (.docx)", variable=self.mode,
            value="pdf2docx", command=self._reset_file,
        ).pack(side="left", padx=(0, 16))
        tk.Radiobutton(
            mode_frame, text="Word → PDF", variable=self.mode,
            value="docx2pdf", command=self._reset_file,
        ).pack(side="left")

        # File picker row
        file_frame = tk.Frame(self)
        file_frame.pack(fill="x", **pad)
        self.file_label = tk.Label(
            file_frame, text="No file selected.", anchor="w", fg="#333"
        )
        self.file_label.pack(side="left", fill="x", expand=True)
        tk.Button(
            file_frame, text="Choose File…", command=self._choose_file
        ).pack(side="right")

        # Convert button
        self.convert_btn = tk.Button(
            self, text="Convert", command=self._start_conversion,
            state="disabled", bg="#2563eb", fg="white",
            font=("Segoe UI", 11, "bold"), padx=20, pady=6,
        )
        self.convert_btn.pack(pady=12)

        # Progress bar
        self.progress = ttk.Progressbar(
            self, orient="horizontal", length=480, mode="determinate"
        )
        self.progress.pack(pady=4)

        # Status / log area
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(self, textvariable=self.status_var, fg="#333").pack(pady=4)

        self.log_box = tk.Text(self, height=6, width=64, state="disabled")
        self.log_box.pack(padx=16, pady=8)

    # ------------------------------------------------------------- helpers

    def _reset_file(self):
        self.input_path = None
        self.file_label.config(text="No file selected.")
        self.convert_btn.config(state="disabled")

    def _choose_file(self):
        if self.mode.get() == "pdf2docx":
            filetypes = [("PDF files", "*.pdf")]
        else:
            filetypes = [("Word documents", "*.docx")]

        path = filedialog.askopenfilename(title="Select a file", filetypes=filetypes)
        if not path:
            return

        self.input_path = path
        self.file_label.config(text=os.path.basename(path))
        self.convert_btn.config(state="normal")
        self._log(f"Selected: {path}")

    def _log(self, message: str):
        self.log_box.config(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def _set_progress(self, percent: int):
        # Called from a worker thread — marshal back onto the Tk main loop.
        self.after(0, lambda: self.progress.config(value=percent))

    # -------------------------------------------------------- conversion

    def _start_conversion(self):
        if not self.input_path:
            return

        default_ext = ".docx" if self.mode.get() == "pdf2docx" else ".pdf"
        save_path = filedialog.asksaveasfilename(
            title="Save converted file as…",
            defaultextension=default_ext,
            initialfile=os.path.splitext(os.path.basename(self.input_path))[0]
            + default_ext,
            filetypes=[("All files", "*.*")],
        )
        if not save_path:
            return

        self.convert_btn.config(state="disabled")
        self.progress.config(value=0)
        self.status_var.set("Converting…")
        self._log("Starting conversion…")

        thread = threading.Thread(
            target=self._run_conversion, args=(self.input_path, save_path),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, input_path: str, output_path: str):
        try:
            if self.mode.get() == "pdf2docx":
                result_path = convert_pdf_to_docx(
                    input_path, output_path, progress_callback=self._set_progress
                )
            else:
                result_path = convert_docx_to_pdf(
                    input_path, output_path, progress_callback=self._set_progress
                )
            self.after(0, self._on_success, result_path)
        except ConversionError as e:
            self.after(0, self._on_error, str(e))
        except Exception as e:  # safety net for anything unexpected
            self.after(0, self._on_error, f"Unexpected error: {e}")

    def _on_success(self, result_path: str):
        self.status_var.set("Done.")
        self._log(f"Saved: {result_path}")
        self.convert_btn.config(state="normal")
        messagebox.showinfo("Conversion complete", f"Saved to:\n{result_path}")

    def _on_error(self, message: str):
        self.status_var.set("Failed.")
        self._log(f"ERROR: {message}")
        self.convert_btn.config(state="normal")
        messagebox.showerror("Conversion failed", message)


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
