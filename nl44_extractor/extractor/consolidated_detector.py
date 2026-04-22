"""
consolidated_detector.py — finds the NL-44 page range in a consolidated PDF.

Detection strategy:
  START: first page where >= min_matches NL-44 keywords appear
  END:   page before the next form header, or last page of PDF

NL-44 is a single-page form.
"""

import re
import logging
import tempfile
import os
from typing import Optional, Tuple, List

logger = logging.getLogger(__name__)

DEFAULT_KEYWORDS = [
    "NL-44",
    "MOTOR TP OBLIGATIONS",
    "QUARTERLY RETURNS",   # always in the form title; old format also had
                           # "LIABILITY ONLY POLICIES" / "PACKAGE POLICIES" but
                           # the new IRDAI 2024 format (Chola etc.) does not
]
DEFAULT_MIN_MATCHES = 3

FORM_HEADER_PATTERN = re.compile(
    r"^\s*(?:FORM\s+)?NL[-\s]?(\d+)|\bFORM\s+NL[-\s]?(\d+)", 
    re.IGNORECASE | re.MULTILINE
)
def is_toc_page(text: str) -> bool:
    if re.search(r"TABLE\s+OF\s+CONTENTS|FORM\s+INDEX|INDEX\s+OF\s+FORMS", text, re.IGNORECASE):
        return True
    matches = re.findall(r"\bNL[-\s]?(\d+)\b", text, re.IGNORECASE)
    valid_forms = set(m for m in matches if 1 <= int(m) <= 45)
    return len(valid_forms) >= 4


def _page_keyword_count(text: str, keywords: List[str]) -> int:
    # Normalise spacing variants e.g. "NL - 44" → "NL-44" before matching
    text_upper = re.sub(r'NL\s*-\s*(\d+)', r'NL-\1', text.upper())
    return sum(1 for kw in keywords if kw.upper() in text_upper)


def find_nl44_pages(
    pdf_path: str,
    keywords: Optional[List[str]] = None,
    min_matches: int = DEFAULT_MIN_MATCHES,
) -> Optional[Tuple[int, int]]:
    try:
        import pdfplumber
    except ImportError:
        logger.error("pdfplumber not available")
        return None

    if keywords is None:
        keywords = DEFAULT_KEYWORDS

    try:
        with pdfplumber.open(pdf_path) as pdf:
            n_pages = len(pdf.pages)
            page_texts = [page.extract_text() or "" for page in pdf.pages]

        start_page = None
        for i, text in enumerate(page_texts):
            if is_toc_page(text):
                logger.debug(f"  page {i + 1}: TOC page, skipping")
                continue
            if _page_keyword_count(text, keywords) >= min_matches:
                start_page = i
                break

        if start_page is None:
            logger.warning(f"NL-44 section not found in: {pdf_path}")
            return None

        end_page = n_pages - 1
        for i in range(start_page + 1, n_pages):
            matches = FORM_HEADER_PATTERN.findall(page_texts[i])
            non_nl44 = [m for m in matches if m != "44"]
            if non_nl44:
                end_page = i - 1
                break

        logger.info(
            f"NL-44 found at pages {start_page}-{end_page} "
            f"(0-indexed) in {os.path.basename(pdf_path)}"
        )
        return (start_page, end_page)

    except Exception as e:
        logger.error(f"Error scanning consolidated PDF {pdf_path}: {e}")
        return None


def extract_nl44_to_temp(pdf_path: str, start_page: int, end_page: int) -> Optional[str]:
    try:
        import pypdf
    except ImportError:
        try:
            import PyPDF2 as pypdf
        except ImportError:
            logger.error("pypdf or PyPDF2 not available")
            return None

    try:
        reader = pypdf.PdfReader(pdf_path)
        writer = pypdf.PdfWriter()
        for page_num in range(start_page, end_page + 1):
            if page_num < len(reader.pages):
                writer.add_page(reader.pages[page_num])

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, prefix="nl44_extract_")
        with open(tmp.name, "wb") as f:
            writer.write(f)
        return tmp.name

    except Exception as e:
        logger.error(f"Error extracting pages from {pdf_path}: {e}")
        return None
