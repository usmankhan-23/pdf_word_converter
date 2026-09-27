"""
converter package
------------------
Public API:
    convert_pdf_to_docx(input_path, output_path=None, ...)
    convert_docx_to_pdf(input_path, output_path=None, ...)
    ConversionError
"""

from .pdf_to_docx_converter import convert_pdf_to_docx
from .docx_to_pdf_converter import convert_docx_to_pdf
from .utils import ConversionError

__all__ = ["convert_pdf_to_docx", "convert_docx_to_pdf", "ConversionError"]
