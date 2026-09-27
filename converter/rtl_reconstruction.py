"""
converter/rtl_reconstruction.py
--------------------------------
Geometry-based text reconstruction for Arabic-script (Arabic/Urdu/Persian,
i.e. right-to-left, cursive-joining) PDF content.

--------------------------------------------------------------------------
THE BUG THIS WORKS AROUND
--------------------------------------------------------------------------
Some PDFs containing Arabic-script text (observed with Urdu text set in
Noto Nastaliq Urdu, produced via a Word/LibreOffice PDF export) have a
literal space character (U+0020) injected between letters that should be
part of the same word. This breaks the visual shaping (each letter renders
in its isolated, unjoined form) AND breaks naive text extraction (every
letter looks like its own one-character "word").

Critically, this is NOT a scrambled-character problem: the underlying
Unicode codepoints are correct Arabic-script letters, in the correct
sequence. It's purely an over-insertion of spaces at the wrong places.

--------------------------------------------------------------------------
THE FIX, AND HOW IT WAS VALIDATED
--------------------------------------------------------------------------
Measuring the real glyph-to-glyph gap (end of one letter's bbox to the
start of the next), IGNORING whether a literal space character sits
between them in the PDF's character stream, produces a clean bimodal
split: genuine within-word letter transitions cluster tightly near 0
(slightly negative to a few tenths of a point — cursive letters touch or
nearly touch), while genuine word boundaries are 5-20x larger. So real
word breaks can be recovered purely from geometry, independent of
whatever spaces the source PDF's content stream happens to contain.

This module's word/character reordering logic was validated against a
PDF with known, ground-truth source text (a cleanly-shaped Urdu paragraph
converted through this same project's docx->pdf path): running this
reconstruction against that PDF reproduced the original source string
byte-for-byte. See the project's development notes for the full
validation transcript.

Note on ordering: PDF content streams for RTL text list characters in
physical left-to-right page position, which is the REVERSE of logical
reading order at both the character level (within a word) and the word
level (within a line). Both need to be reversed to recover logical order
— reversing only one produces text where words are in the right order
but each word's letters are individually backwards (a real bug caught
during validation of this exact module).
"""

from __future__ import annotations
import re
import statistics
from collections import Counter

# Unicode ranges covering Arabic, Persian, and Urdu letters (Urdu-specific
# letters like ٹ ڈ ڑ ں ے live inside the main Arabic block), plus Arabic
# presentation forms.
_ARABIC_SCRIPT_RANGES = (
    (0x0600, 0x06FF),
    (0x0750, 0x077F),
    (0xFB50, 0xFDFF),
    (0xFE70, 0xFEFF),
)


def is_arabic_script_char(ch: str) -> bool:
    cp = ord(ch)
    return any(lo <= cp <= hi for lo, hi in _ARABIC_SCRIPT_RANGES)


def page_arabic_script_ratio(page) -> float:
    """Fraction of non-space characters on a PyMuPDF page that are Arabic-script."""
    text = page.get_text("text")
    non_space = [c for c in text if not c.isspace()]
    if not non_space:
        return 0.0
    arabic = sum(1 for c in non_space if is_arabic_script_char(c))
    return arabic / len(non_space)


def _gap_threshold_for_size(font_size: float) -> float:
    """
    Adaptive space-vs-no-space cutoff, scaled to font size. Validated at
    two different sizes (11pt and 18pt, both real documents) against
    ground truth: a genuine word gap is always at least ~5x a genuine
    intra-word gap, so a small fixed floor plus a modest per-point scale
    sits safely in the middle of that range in both cases.
    """
    return max(0.4, font_size * 0.04)


def clean_font_name(raw_name: str) -> str:
    """
    Turn a PDF-embedded PostScript font name into a real family name that
    Word/LibreOffice can actually look up on the system.

    PDFs store subsetted fonts with a 6-uppercase-letter subset tag
    ("BCDEEE+NotoNastaliqUrdu-Regular") rather than the plain family name
    ("Noto Nastaliq Urdu") a font-lookup by name expects. Also splits
    run-together CamelCase words, which is how these PostScript names
    encode multi-word family names.
    """
    name = re.sub(r"^[A-Z]{6}\+", "", raw_name)
    name = re.sub(r"-(Regular|Bold|Italic|BoldItalic|Medium|SemiBold|Light)$", "", name, flags=re.IGNORECASE)
    # Split CamelCase into words: insert a space before an uppercase letter
    # that follows a lowercase letter or digit ("NotoNastaliqUrdu" -> "Noto Nastaliq Urdu").
    name = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", name)
    return name.strip() or raw_name


def _reconstruct_block_text(block: dict) -> tuple[str, str, float]:
    """
    Rebuild one PDF text block's content in correct logical (reading-order)
    text, using the geometric word-gap heuristic instead of trusting
    whatever space characters happen to already be in the PDF.

    Returns (text, dominant_font_name, dominant_font_size).
    """
    font_counter = Counter()
    line_texts = []

    for line in block.get("lines", []):
        chars = []
        for span in line.get("spans", []):
            font_counter[(span.get("font", ""), span.get("size", 12.0))] += len(
                span.get("chars", [])
            )
            for ch in span.get("chars", []):
                chars.append({"c": ch["c"], "bbox": ch["bbox"], "size": span.get("size", 12.0)})

        chars.sort(key=lambda c: c["bbox"][0])
        non_space = [c for c in chars if c["c"] != " "]
        if not non_space:
            continue

        threshold = _gap_threshold_for_size(non_space[0]["size"])
        words = [[non_space[0]]]
        for a, b in zip(non_space, non_space[1:]):
            gap = b["bbox"][0] - a["bbox"][2]
            if gap > threshold:
                words.append([])
            words[-1].append(b)

        word_strings = ["".join(reversed([c["c"] for c in w])) for w in words]
        line_texts.append(" ".join(reversed(word_strings)))

    dominant_font = font_counter.most_common(1)[0][0] if font_counter else ("", 12.0)
    return " ".join(line_texts), dominant_font[0], dominant_font[1]


def reconstruct_page_paragraphs(page) -> list[dict]:
    """
    Reconstruct an Arabic-script PDF page as a list of paragraph records,
    each: {"text": str, "font": str, "size": float, "y0": float}, sorted
    top-to-bottom by actual page position (one paragraph per PDF text
    block — PDF blocks generally correspond to paragraphs).
    """
    raw = page.get_text("rawdict")
    paragraphs = []

    for block in raw.get("blocks", []):
        if block.get("type") != 0:  # skip images
            continue
        text, font, size = _reconstruct_block_text(block)
        if not text.strip():
            continue
        y0 = block["bbox"][1]
        paragraphs.append({"text": text, "font": clean_font_name(font), "size": size, "y0": y0})

    paragraphs.sort(key=lambda p: p["y0"])
    return paragraphs
