# PDF ⇄ Word Converter (Offline-First Desktop Utility)

A local, no-internet-required desktop tool for converting single files
between PDF and Microsoft Word (`.docx`), with a Tkinter GUI.

## Architecture

```text
pdf_word_converter/
├── main.py                          # Tkinter GUI — the entry point
├── converter/
│   ├── __init__.py                  # Public API surface
│   ├── pdf_to_docx_converter.py      # PDF -> DOCX (hybrid per-page engine)
│   ├── rtl_reconstruction.py        # Geometry-based RTL/Arabic text reconstruction
│   ├── docx_to_pdf_converter.py      # DOCX -> PDF (Word COM / LibreOffice)
│   └── utils.py                     # Validation, logging, temp dirs
├── requirements.txt
└── README.md
```

**Design principle:** `main.py` only knows about the GUI. It never touches
PDF/DOCX internals directly — it calls `converter.convert_pdf_to_docx(...)`
or `converter.convert_docx_to_pdf(...)` and handles the result/exception.
This means you can swap out either conversion backend later without
touching the UI at all, or reuse `converter/` as a CLI/headless tool.

### Why two different backends for the two directions?

| Direction   | Backend                          | Why                                                                 |
|-------------|-----------------------------------|----------------------------------------------------------------------|
| PDF → DOCX  | Custom PyMuPDF + python-docx pipeline | See "PDF → DOCX: why not just pdf2docx?" below. |
| DOCX → PDF  | LibreOffice headless, falling back to MS Word (COM) | High-fidelity OpenXML→PDF rendering essentially requires a full layout engine. Rather than reimplementing one, the tool drives whichever renderer is already installed locally. **Both options run entirely on-device** — no data ever leaves the machine, no network call is made — so this still satisfies "offline-first, zero cloud dependency." |

`docx_to_pdf_converter.py` tries backends in order and **actually attempts
each one**, falling back automatically on failure, rather than assuming a
backend works just because its Python package imported:
1. LibreOffice (`soffice --headless --convert-to pdf`) — tried first.
2. Microsoft Word via `docx2pdf` (Windows/macOS only) — tried if
   LibreOffice isn't installed, or as a second attempt if it is but fails.
3. If every available backend fails, raises a `ConversionError` explaining
   exactly what to install.

This matters because `docx2pdf` importing successfully does **not** mean
Word is actually installed/registered — an earlier version of this tool
assumed it did, which produced a `-2147221005 'Invalid class string'` COM
error with no fallback. That's fixed now.

### PDF → DOCX: why not just pdf2docx?

An earlier version used `pdf2docx`, which works well for ordinary text
PDFs but badly corrupts math-heavy academic PDFs (see
`formatting_comparison_report.md` for the original failure report against
a LaTeX-generated homework PDF): it stripped word spacing between
prose words, scattered matrix brackets into disjointed table cells, and
flattened subscripts/superscripts into plain digits.

Root cause: LaTeX-produced PDFs often position glyphs using pure
horizontal offsets instead of an actual space character, and draw tall
brackets as separate stacked Unicode glyph pieces. `pdf2docx` has no
semantic understanding of "this is a matrix" — it just forces whatever
characters it finds into paragraphs/tables.

The current pipeline instead:
1. **Reconstructs prose text from raw glyph coordinates**, inferring a
   space whenever the horizontal gap between two characters exceeds a
   threshold fraction of the local font size. This is the standard fix
   for "PDF text extraction with missing space characters" and needs no
   external NLP library.
2. **Detects math** (a block containing stacked bracket-extension glyphs,
   or a mix of font sizes on one line — how sub/superscripts render) and,
   instead of trying to re-typeset it, **crops that exact region straight
   out of the PDF page and embeds it as an image**. Adjacent math pieces
   that PyMuPDF splits into separate objects (a very common occurrence —
   one bracket glyph, one matrix column, and a trailing "∈ R⁴" can each
   land in their own internal "block") are spatially clustered back
   together first, so a whole matrix/equation becomes one clean image
   instead of several fragmented, misaligned ones.

This trades editability of math content for guaranteed visual correctness
— the matrix looks exactly like the source PDF because it genuinely is a
snapshot of it. Fully reconstructing math as native, editable OpenXML
Math (MathML) objects would need a dedicated math-OCR model; that's a
different, much larger project than "convert my PDF," and no mainstream
converter (commercial or open-source) does it reliably either.

**Known rough edge:** occasional small stray fragments (e.g. an isolated
punctuation mark near an equation boundary) can still slip through the
clustering as a lone paragraph. This is cosmetic — content isn't lost,
just occasionally an extra character floats in the wrong spot — and safe
to manually delete if noticed.

## Setup

```bash
# 1. Create a virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. (Linux only) Tkinter is a system package, not a pip package
sudo apt install python3-tk      # Debian/Ubuntu
# or: sudo dnf install python3-tkinter   # Fedora

# 4. (For DOCX -> PDF) Install a local renderer if you don't already have one:
#    - Windows/macOS: Microsoft Word, OR
#    - Any OS: LibreOffice — https://www.libreoffice.org/download

# 5. Run
python main.py
```

## Packaging as a standalone executable

To ship this as a single-file desktop app with no Python install required
on the target machine, use PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name PDFWordConverter main.py
```

The output binary will be in `dist/`. Note: PyInstaller bundles your
Python code and dependencies, but it does **not** bundle LibreOffice/Word —
those still need to be separately installed on the end-user's machine for
DOCX → PDF conversion (PDF → DOCX works with no external program at all).

## Extending this into a full "batch" tool

The prompt describes single-file conversion. If you later want batch
processing, the clean extension point is `converter/`: add a
`batch.py` that loops over a folder calling the same
`convert_pdf_to_docx` / `convert_docx_to_pdf` functions, and add a
"Convert Folder…" button in `main.py` that spins up a thread pool instead
of a single thread. No changes needed to the conversion logic itself.

## Known limitations

- DOCX → PDF fidelity depends entirely on the backend renderer (Word or
  LibreOffice) — this project doesn't control that rendering, only invokes it.
- Very complex PDFs (heavy vector art, unusual font embedding, unusual
  encodings) may not round-trip perfectly into editable DOCX — this is a
  `pdf2docx`/PyMuPDF limitation, not something this wrapper can fix.
- Currently handles one file at a time, per the stated requirement.
