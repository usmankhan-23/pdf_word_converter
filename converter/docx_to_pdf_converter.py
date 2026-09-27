"""
converter/docx_to_pdf_converter.py
-----------------------------------
DOCX -> PDF conversion.

v2 — the v1 code assumed that if the `docx2pdf` *package* imported
successfully, Microsoft Word itself must be installed and working. That's
false: `docx2pdf` imports fine even when Word isn't registered/installed,
and only fails at the moment it tries to create the COM object — which is
exactly the crash reported ("Invalid class string", -2147221005). v1 never
even attempted a fallback in that case.

v2 instead builds an ordered list of every backend that's plausibly
available, and actually TRIES each one in turn, falling back automatically
on any failure instead of assuming success:

  1. LibreOffice headless (`soffice --convert-to pdf`) — free, cross-
     platform, and (once installed) far more reliably scriptable than
     driving a GUI app over COM/AppleScript. Tried first.
  2. Microsoft Word, via `docx2pdf` (COM on Windows, AppleScript on macOS)
     — used if LibreOffice isn't installed, or as a second attempt if it
     is installed but fails for some other reason.

Both run entirely on the local machine — no network call is ever made.

Fix: the Word backend also now initializes its own COM apartment
(pythoncom.CoInitialize/CoUninitialize) around the call, since the GUI
invokes conversion from a background thread — see _convert_with_word.
"""

from __future__ import annotations
import os
import shutil
import subprocess
import platform
from .utils import (
    get_logger,
    validate_input_file,
    ensure_output_path,
    temp_workdir,
    which,
    ConversionError,
)

logger = get_logger(__name__)


def _get_candidate_backends() -> list[str]:
    """
    Return the backends worth trying, in the order to try them.
    Presence here means "plausibly usable" — actually invoking it may
    still fail (e.g. Word installed but not licensed/registered), which
    is exactly why the caller tries each one and falls back rather than
    trusting this list blindly.
    """
    backends = []

    if which("soffice") or which("libreoffice"):
        backends.append("libreoffice")

    if platform.system() in ("Windows", "Darwin"):
        try:
            import docx2pdf  # noqa: F401
            backends.append("word")
        except ImportError:
            pass

    return backends


def _convert_with_word(input_path: str, output_path: str) -> None:
    """
    Drive Microsoft Word via docx2pdf.

    On Windows this goes through win32com, which requires the calling
    thread to have its own initialized COM apartment before it can create
    any COM object. The GUI runs every conversion on a background
    threading.Thread, which never calls CoInitialize() — so Word's COM
    class registration looks invalid to that thread even though Word and
    docx2pdf are both installed correctly. That mismatch is exactly what
    produces "(-2147221005, 'Invalid class string', None, None)"
    (CO_E_CLASSSTRING). Initializing (and tearing down) the apartment
    around this call fixes it without touching the GUI's threading model.
    """
    pythoncom = None
    if platform.system() == "Windows":
        import pythoncom  # only importable/needed on Windows
        pythoncom.CoInitialize()

    try:
        from docx2pdf import convert as word_convert
        word_convert(input_path, output_path)
    finally:
        if pythoncom is not None:
            pythoncom.CoUninitialize()


def _convert_with_libreoffice(input_path: str, output_path: str) -> None:
    soffice = which("soffice") or which("libreoffice")
    if not soffice:
        raise ConversionError("LibreOffice binary not found on PATH.")

    with temp_workdir() as workdir:
        cmd = [
            soffice,
            "--headless",
            "--norestore",
            "--convert-to", "pdf",
            "--outdir", workdir,
            input_path,
        ]
        logger.info("Running local LibreOffice conversion: %s", " ".join(cmd))
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=180,
        )
        if result.returncode != 0:
            raise ConversionError(
                "LibreOffice conversion failed:\n"
                f"{result.stderr.decode(errors='ignore')}"
            )

        produced_name = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
        produced_path = os.path.join(workdir, produced_name)
        if not os.path.isfile(produced_path):
            raise ConversionError(
                "LibreOffice did not produce the expected output file."
            )

        shutil.move(produced_path, output_path)


_BACKEND_FUNCS = {
    "libreoffice": _convert_with_libreoffice,
    "word": _convert_with_word,
}


def convert_docx_to_pdf(
    input_path: str,
    output_path: str | None = None,
    progress_callback=None,
) -> str:
    """
    Convert a single Word document (.docx) into a standards-compliant PDF.

    Tries every plausibly-available local backend in order, falling back
    automatically if one fails, instead of trusting that a backend will
    work just because its Python package imported.

    Args:
        input_path: path to the source .docx
        output_path: path to write the .pdf to; derived automatically if None
        progress_callback: optional callable(percent: int) -> None

    Returns:
        The resolved output path on success.

    Raises:
        ConversionError if every available backend failed, or none exist,
        with a message describing what to install.
    """
    validate_input_file(input_path, ".docx")
    output_path = ensure_output_path(input_path, output_path, ".pdf")

    backends = _get_candidate_backends()
    if not backends:
        raise ConversionError(
            "No local rendering backend found.\n"
            "Install one of the following (both run fully offline):\n"
            "  - LibreOffice (any OS, free) — https://www.libreoffice.org/download\n"
            "  - Microsoft Word (Windows/macOS) + `pip install docx2pdf`"
        )

    if progress_callback:
        progress_callback(10)

    errors = []
    for backend in backends:
        try:
            logger.info(
                "Attempting DOCX->PDF via '%s' backend: %s -> %s",
                backend, input_path, output_path,
            )
            _BACKEND_FUNCS[backend](input_path, output_path)

            if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
                raise ConversionError(f"'{backend}' reported success but produced no output.")

            if progress_callback:
                progress_callback(100)
            logger.info("DOCX->PDF succeeded via '%s': %s", backend, output_path)
            return output_path

        except Exception as e:
            logger.warning("Backend '%s' failed: %s", backend, e)
            errors.append(f"  - {backend}: {e}")
            continue

    raise ConversionError(
        "All available conversion backends failed:\n"
        + "\n".join(errors)
        + "\n\nIf Word failed with a COM/'Invalid class string' error, Word is not "
        "properly installed or licensed on this machine. Install LibreOffice as a "
        "free, reliable, fully-offline fallback: https://www.libreoffice.org/download"
    )
