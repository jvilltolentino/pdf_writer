# Excel → PDF Form Filler

A desktop GUI app that fills a fillable PDF template with data from an Excel
file — one filled PDF per row.

## What it does (the STAR)

- **Situation:** You have an Excel file (e.g. employee list, invoice data) and
  a fillable PDF form template.
- **Task:** You need to produce one filled PDF per row, without doing it by hand.
- **Action:** This app reads the Excel, lets you map each PDF field to an
  Excel column, and stamps the data into the PDF.
- **Result:** A folder full of filled PDFs, one per row, named after a column
  you choose (e.g. employee name).

> Think of it like a **mail merge for PDFs** — same template, different data
> for each row.

## Setup (one-time)

You need Python 3.10 or newer.

```bash
# 1. Go into the project folder
cd excel_pdf_filler

# 2. (Recommended) Create a virtual environment
python -m venv venv

# Activate it
# On Windows:
venv\Scripts\activate
# On macOS / Linux:
source venv/bin/activate

# 3. Install the libraries
pip install -r requirements.txt
```

## Run the app

```bash
python app.py
```

A window will open with three steps:

1. **Choose files** — pick your Excel/CSV, your fillable PDF template,
   and an output folder.
2. **Map PDF fields to Excel columns** — the app auto-guesses matches; you
   can adjust any of them. Pick "(skip)" to leave a field blank.
3. **Generate filled PDFs** — click the button and watch the progress bar.

## Execution logs

Every run creates a timestamped log file in the **output folder**:

```
execution_2026-05-05_10-52-55.log
```

Each log records:
- The Excel file, PDF template, and output folder used
- The field mappings you picked
- A timestamped success line for **each PDF generated**, e.g.
  `[2026-05-05 10:52:55] INFO — Row 1/3 — generated successfully: 001_Ana Cruz.pdf`
- Any errors that happened (so you can investigate)
- A summary at the bottom (total successes, total failures)

> Think of it as a **receipt** for the run — proof of what was created and when.

Each run gets its own log file (named with the timestamp), so older logs are
never overwritten.

## Important: your PDF must be "fillable"

This app fills **existing form fields** in a PDF. Your template needs to have
real form fields (not just lines drawn on a page).

To check or add form fields, open your PDF in:
- **Adobe Acrobat** (Prepare Form tool), or
- **LibreOffice Draw** (free), or
- Any PDF form designer.

If the app says "No fillable fields found," that means the PDF is just an
image or flat document — you'll need to add form fields to it first.

## File structure

```
excel_pdf_filler/
├── app.py              # The main GUI app
├── requirements.txt    # Python libraries it needs
└── README.md           # This file
```

## Building a standalone .exe (optional)

If you want to share the app with someone who doesn't have Python:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "PDF Form Filler" app.py
```

The standalone executable will appear in the `dist/` folder.
