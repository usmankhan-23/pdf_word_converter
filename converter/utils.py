"""
converter/utils.py
------------------
Shared helpers used by both conversion modules:
  - input validation (extension / existence checks)
  - a simple rotating file logger (all activity stays local — no telemetry)
  - a safe temp-directory context manager for intermediate files
"""

import os
import logging
import platform
import tempfile
import shutil
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler

APP_NAME = "PDFWordConverter"


def get_log_dir() -> str:
    """Return (and create) a per-user local log directory. Fully offline."""
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    else:
        base = os.path.expanduser("~/.local/share")
    log_dir = os.path.join(base, APP_NAME, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


def get_logger(name: str = APP_NAME) -> logging.Logger:
    """Configured logger writing only to a local rotating file + console."""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger  # already configured

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
    )

    log_path = os.path.join(get_log_dir(), "converter.log")
    file_handler = RotatingFileHandler(
        log_path, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    file_handler.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)
    console_handler.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


class ConversionError(Exception):
    """Raised for any recoverable conversion failure, with a user-facing message."""


def validate_input_file(path: str, expected_ext: str) -> None:
    """
    Raise ConversionError with a clear message if the input file is missing,
    unreadable, or the wrong type. expected_ext e.g. '.pdf' or '.docx'
    """
    if not path:
        raise ConversionError("No input file selected.")
    if not os.path.isfile(path):
        raise ConversionError(f"File not found: {path}")
    if not os.access(path, os.R_OK):
        raise ConversionError(f"File is not readable (permissions?): {path}")

    ext = os.path.splitext(path)[1].lower()
    if ext != expected_ext.lower():
        raise ConversionError(
            f"Expected a {expected_ext} file but got '{ext}': {path}"
        )

    if os.path.getsize(path) == 0:
        raise ConversionError(f"File is empty: {path}")


def ensure_output_path(input_path: str, output_path: str | None, new_ext: str) -> str:
    """
    If output_path is None, derive one next to the input file with new_ext.
    Ensures the parent directory exists and is writable.
    """
    if output_path is None:
        base = os.path.splitext(input_path)[0]
        output_path = base + new_ext
    else:
        parent = os.path.dirname(os.path.abspath(output_path)) or "."
        os.makedirs(parent, exist_ok=True)
        if not os.access(parent, os.W_OK):
            raise ConversionError(f"Output directory is not writable: {parent}")
    return output_path


@contextmanager
def temp_workdir():
    """
    Context manager yielding a fresh local temp directory that is guaranteed
    to be cleaned up afterward, even on failure. No data leaves this machine.
    """
    d = tempfile.mkdtemp(prefix="pwc_")
    try:
        yield d
    finally:
        shutil.rmtree(d, ignore_errors=True)


def which(binary_name: str) -> str | None:
    """
    Locate a local binary (e.g. soffice). Tries PATH first via shutil.which,
    then falls back to LibreOffice's actual default install locations,
    because the official Windows/macOS installers do NOT add soffice to
    PATH — so shutil.which("soffice") alone reports "not found" even on a
    machine that has LibreOffice properly installed.
    """
    found = shutil.which(binary_name)
    if found:
        return found

    if binary_name not in ("soffice", "libreoffice"):
        return None

    import glob

    system = platform.system()
    candidates: list[str] = []

    if system == "Windows":
        for env_var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env_var)
            if base:
                candidates.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    elif system == "Darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    else:  # Linux and others
        candidates += [
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/snap/bin/libreoffice",
        ]
        candidates += glob.glob("/opt/libreoffice*/program/soffice")

    for path in candidates:
        if os.path.isfile(path):
            return path

    return None