"""
Excel to PDF Form Filler — Desktop GUI App
-------------------------------------------
Fills an existing fillable PDF template with data from each row of an Excel file.
Generates one PDF per row.

How it works (STAR):
  - Situation: You have an Excel with rows of data and a PDF form template.
  - Task:      Fill the PDF for each row automatically.
  - Action:    Read Excel rows -> map columns to PDF fields -> stamp data -> save.
  - Result:    One filled PDF per row, saved to your chosen output folder.
"""

import os
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject


# -----------------------------
# Core PDF + Excel logic
# -----------------------------

def read_excel(path: str) -> pd.DataFrame:
    """Read an Excel or CSV file into a pandas DataFrame."""
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


def get_pdf_fields(pdf_path: str) -> list[str]:
    """Return the list of fillable field names in the PDF template."""
    reader = PdfReader(pdf_path)
    fields = reader.get_fields() or {}
    return list(fields.keys())


def fill_pdf(template_path: str, output_path: str, data: dict) -> None:
    """
    Fill a PDF template with the given data dictionary and save to output_path.
    Keys in `data` must match the field names in the PDF template.
    """
    reader = PdfReader(template_path)
    writer = PdfWriter(clone_from=reader)

    # Stringify all values (PDF fields expect text)
    str_data = {k: ("" if v is None else str(v)) for k, v in data.items()}

    # Update form fields on every page that has them
    for page in writer.pages:
        if "/Annots" in page:
            writer.update_page_form_field_values(page, str_data)

    # Make sure the values stay visible after saving (set NeedAppearances)
    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )

    with open(output_path, "wb") as f:
        writer.write(f)


def safe_filename(text: str) -> str:
    """Make a string safe to use as a filename."""
    keep = "-_.() "
    cleaned = "".join(c if c.isalnum() or c in keep else "_" for c in str(text))
    return cleaned.strip() or "record"


# -----------------------------
# GUI Application
# -----------------------------

class ExcelPdfFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → PDF Form Filler")
        self.geometry("780x640")
        self.minsize(720, 600)

        # State
        self.excel_path: str | None = None
        self.pdf_path: str | None = None
        self.output_dir: str | None = None
        self.df: pd.DataFrame | None = None
        self.pdf_fields: list[str] = []
        self.mappings: dict[str, tk.StringVar] = {}  # pdf_field -> excel_column var
        self.name_field_var = tk.StringVar()  # which PDF field to use for filename

        self._build_ui()

    # ---- UI construction ----
    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#666")
        style.configure("Step.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

        # Header
        header = ttk.Frame(self, padding=(16, 14, 16, 6))
        header.pack(fill="x")
        ttk.Label(header, text="Excel → PDF Form Filler", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Pick an Excel file and a fillable PDF template. Each row becomes one filled PDF.",
            style="Sub.TLabel",
        ).pack(anchor="w")

        # Step 1: File selection
        step1 = ttk.LabelFrame(self, text=" Step 1 — Choose files ", padding=12, style="Step.TLabelframe")
        step1.pack(fill="x", padx=16, pady=(10, 6))

        self._make_file_row(step1, "Excel / CSV file:", "excel", self._pick_excel)
        self._make_file_row(step1, "PDF template:", "pdf", self._pick_pdf)
        self._make_file_row(step1, "Output folder:", "out", self._pick_output_dir)

        # Step 2: Mapping
        step2 = ttk.LabelFrame(self, text=" Step 2 — Map PDF fields to Excel columns ", padding=12, style="Step.TLabelframe")
        step2.pack(fill="both", expand=True, padx=16, pady=6)

        # Scrollable mapping area
        canvas_container = ttk.Frame(step2)
        canvas_container.pack(fill="both", expand=True)

        self.map_canvas = tk.Canvas(canvas_container, highlightthickness=0, height=200)
        scrollbar = ttk.Scrollbar(canvas_container, orient="vertical", command=self.map_canvas.yview)
        self.map_canvas.configure(yscrollcommand=scrollbar.set)

        self.map_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.map_frame = ttk.Frame(self.map_canvas)
        self.map_canvas.create_window((0, 0), window=self.map_frame, anchor="nw")
        self.map_frame.bind(
            "<Configure>",
            lambda e: self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all")),
        )

        self.map_placeholder = ttk.Label(
            self.map_frame,
            text="↑ Choose an Excel file and a PDF template above to see field mappings here.",
            style="Sub.TLabel",
        )
        self.map_placeholder.pack(pady=20, padx=10, anchor="w")

        # Filename source selector
        name_frame = ttk.Frame(step2)
        name_frame.pack(fill="x", pady=(8, 0))
        ttk.Label(name_frame, text="Use this field for output filenames:").pack(side="left")
        self.name_combo = ttk.Combobox(name_frame, textvariable=self.name_field_var, state="disabled", width=30)
        self.name_combo.pack(side="left", padx=8)

        # Step 3: Generate
        step3 = ttk.LabelFrame(self, text=" Step 3 — Generate filled PDFs ", padding=12, style="Step.TLabelframe")
        step3.pack(fill="x", padx=16, pady=(6, 10))

        action_row = ttk.Frame(step3)
        action_row.pack(fill="x")
        self.generate_btn = ttk.Button(action_row, text="Generate PDFs", command=self._on_generate, state="disabled")
        self.generate_btn.pack(side="left")

        self.progress = ttk.Progressbar(step3, mode="determinate")
        self.progress.pack(fill="x", pady=(10, 4))

        self.status_lbl = ttk.Label(step3, text="Waiting...", style="Sub.TLabel")
        self.status_lbl.pack(anchor="w")

    def _make_file_row(self, parent, label_text, kind, command):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label_text, width=18).pack(side="left")
        var = tk.StringVar(value="(none selected)")
        setattr(self, f"{kind}_path_var", var)
        ttk.Label(row, textvariable=var, foreground="#444").pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse...", command=command).pack(side="right")

    # ---- File pickers ----
    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="Choose Excel or CSV file",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            self.df = read_excel(path)
        except Exception as e:
            messagebox.showerror("Cannot read file", f"Could not read the spreadsheet:\n{e}")
            return
        self.excel_path = path
        self.excel_path_var.set(f"{Path(path).name}  —  {len(self.df)} rows, {len(self.df.columns)} columns")
        self._refresh_mapping_area()
        self._refresh_generate_button()

    def _pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Choose fillable PDF template",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            fields = get_pdf_fields(path)
        except Exception as e:
            messagebox.showerror("Cannot read PDF", f"Could not read PDF fields:\n{e}")
            return
        if not fields:
            messagebox.showwarning(
                "No fillable fields found",
                "This PDF doesn't contain any fillable form fields.\n\n"
                "Tip: open it in a tool like Adobe Acrobat or LibreOffice and add form fields, "
                "or pick a different template.",
            )
            return
        self.pdf_path = path
        self.pdf_fields = fields
        self.pdf_path_var.set(f"{Path(path).name}  —  {len(fields)} fillable fields")
        self._refresh_mapping_area()
        self._refresh_generate_button()

    def _pick_output_dir(self):
        path = filedialog.askdirectory(title="Choose output folder")
        if not path:
            return
        self.output_dir = path
        self.out_path_var.set(path)
        self._refresh_generate_button()

    # ---- Mapping UI ----
    def _refresh_mapping_area(self):
        # Clear existing widgets
        for child in self.map_frame.winfo_children():
            child.destroy()
        self.mappings = {}

        if not (self.df is not None and self.pdf_fields):
            self.map_placeholder = ttk.Label(
                self.map_frame,
                text="↑ Choose an Excel file and a PDF template above to see field mappings here.",
                style="Sub.TLabel",
            )
            self.map_placeholder.pack(pady=20, padx=10, anchor="w")
            self.name_combo.configure(state="disabled", values=[])
            return

        # Header row
        header_row = ttk.Frame(self.map_frame)
        header_row.pack(fill="x", pady=(0, 4))
        ttk.Label(header_row, text="PDF field", width=30, font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        ttk.Label(header_row, text="Excel column", font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)

        excel_cols = list(self.df.columns)
        options = ["(skip)"] + [str(c) for c in excel_cols]

        for field in self.pdf_fields:
            row = ttk.Frame(self.map_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=field, width=30).pack(side="left", padx=4)

            var = tk.StringVar()
            # Auto-guess: exact match (case-insensitive) or substring
            guess = self._guess_match(field, excel_cols)
            var.set(guess if guess else "(skip)")

            combo = ttk.Combobox(row, textvariable=var, values=options, state="readonly", width=35)
            combo.pack(side="left", padx=4)
            self.mappings[field] = var

        # Filename source
        self.name_combo.configure(state="readonly", values=self.pdf_fields)
        # Default: pick a field that looks like a name
        default_name = next((f for f in self.pdf_fields if "name" in f.lower()), self.pdf_fields[0])
        self.name_field_var.set(default_name)

    def _guess_match(self, pdf_field: str, excel_cols: list) -> str | None:
        pf = pdf_field.lower().replace(" ", "").replace("_", "")
        for col in excel_cols:
            cl = str(col).lower().replace(" ", "").replace("_", "")
            if pf == cl or pf in cl or cl in pf:
                return str(col)
        return None

    def _refresh_generate_button(self):
        ready = bool(self.excel_path and self.pdf_path and self.output_dir)
        self.generate_btn.configure(state="normal" if ready else "disabled")

    # ---- Generation ----
    def _on_generate(self):
        if not (self.df is not None and self.pdf_path and self.output_dir):
            return

        # Build mapping that excludes "(skip)"
        active_map = {pdf_f: var.get() for pdf_f, var in self.mappings.items() if var.get() != "(skip)"}
        if not active_map:
            messagebox.showwarning("No mappings", "Please map at least one PDF field to an Excel column.")
            return

        name_field = self.name_field_var.get()

        # Run in background thread so the UI stays responsive
        self.generate_btn.configure(state="disabled")
        threading.Thread(
            target=self._run_generation,
            args=(active_map, name_field),
            daemon=True,
        ).start()

    def _run_generation(self, active_map: dict, name_field: str):
        total = len(self.df)
        self.progress.configure(maximum=total, value=0)
        successes = 0
        errors: list[str] = []

        for i, (_, row) in enumerate(self.df.iterrows(), start=1):
            try:
                # Build field data from mappings
                data = {pdf_f: row[excel_col] for pdf_f, excel_col in active_map.items() if excel_col in row.index}

                # Decide filename
                name_source = data.get(name_field) or f"row_{i}"
                fname = f"{i:03d}_{safe_filename(name_source)}.pdf"
                out_path = os.path.join(self.output_dir, fname)

                fill_pdf(self.pdf_path, out_path, data)
                successes += 1
            except Exception as e:
                errors.append(f"Row {i}: {e}")

            # Update UI from worker thread (Tk allows this in practice, but use after for safety)
            self.after(0, self._update_progress, i, total)

        self.after(0, self._finish_generation, successes, errors)

    def _update_progress(self, current: int, total: int):
        self.progress.configure(value=current)
        self.status_lbl.configure(text=f"Generating... {current} / {total}")

    def _finish_generation(self, successes: int, errors: list[str]):
        self.generate_btn.configure(state="normal")
        if errors:
            self.status_lbl.configure(text=f"Done with {len(errors)} error(s). {successes} PDFs created.")
            messagebox.showwarning(
                "Finished with errors",
                f"Generated {successes} PDFs.\n\nErrors:\n" + "\n".join(errors[:10]) +
                ("\n..." if len(errors) > 10 else ""),
            )
        else:
            self.status_lbl.configure(text=f"Done! {successes} PDFs saved to: {self.output_dir}")
            messagebox.showinfo("Success", f"Generated {successes} PDFs in:\n{self.output_dir}")


def main():
    app = ExcelPdfFillerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
