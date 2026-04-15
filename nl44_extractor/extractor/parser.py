"""
NL-44 Generic Parser (Motor TP Obligations).

NL-44 table layout:
  Col 0: Items (row label — sometimes spans two pdfplumber rows)
  Col 1: For the Quarter  (QTR value)
  Col 2: Upto the Quarter (YTD value)

Five extractable line items:
  liability_premium   — Gross Direct MTP Premium, Liability Only Policies (L)
  package_premium     — Gross Direct MTP Premium, Package Policies (P)
  total_mtp_premium   — Total Gross Direct MTP Business Premium (L+P)
  motor_od_premium    — Total Gross Direct Motor OD Insurance Business Premium
  total_premium       — Total Gross Direct Premium Income

Output model:
  data[metric_key] = {"qtr": float|None, "ytd": float|None}

Only extract.current_year is populated.
"""

import logging
from pathlib import Path
from typing import Optional

import pdfplumber

from config.company_registry import COMPANY_DISPLAY_NAMES, DEDICATED_PARSER
from config.row_registry import ROW_ALIASES
from extractor.models import CompanyExtract, PeriodData
from extractor.normaliser import clean_number, normalise_text

logger = logging.getLogger(__name__)


def _resolve_metric(label: str) -> Optional[str]:
    """Map a normalised label string to a metric key, or None if unrecognised."""
    norm = normalise_text(label)
    if not norm:
        return None
    # Direct lookup first
    if norm in ROW_ALIASES:
        return ROW_ALIASES[norm]
    # Substring fallback for robustness
    for alias, key in ROW_ALIASES.items():
        if alias in norm or norm in alias:
            return key
    return None


def _extract_table(table, period_data: PeriodData) -> int:
    """
    Parse one NL-44 table into period_data.

    Handles multi-line row labels: when a row has no numeric values, its
    label is carried forward and prepended to the next row's label before
    matching.

    Returns number of metrics extracted.
    """
    pending_label = ""
    extracted = 0

    for row in table:
        if not row or len(row) < 2:
            continue

        raw_label = str(row[0] or "")
        raw_qtr   = row[1] if len(row) > 1 else None
        raw_ytd   = row[2] if len(row) > 2 else None

        qtr_val = clean_number(raw_qtr)
        ytd_val = clean_number(raw_ytd)

        # Build the full label (pending continuation + current)
        full_label = (pending_label + " " + raw_label).strip() if pending_label else raw_label

        metric_key = _resolve_metric(full_label)

        if metric_key is None:
            # Also try the raw label alone (for rows like "Business Premium (L+P)")
            metric_key = _resolve_metric(raw_label)

        if metric_key and (qtr_val is not None or ytd_val is not None):
            period_data.data[metric_key] = {
                "qtr": qtr_val,
                "ytd": ytd_val,
            }
            extracted += 1
            pending_label = ""
            logger.debug(f"  {metric_key}: qtr={qtr_val} ytd={ytd_val}  (label='{full_label}')")
        elif qtr_val is None and ytd_val is None and raw_label.strip():
            # No values — carry label forward for the next row
            pending_label = full_label
        else:
            # Has values but no metric match — discard continuation
            pending_label = ""

    return extracted


def parse_pdf(pdf_path: str, company_key: str, quarter: str = "", year: str = "") -> CompanyExtract:
    """Main entry point — parses one NL-44 PDF."""
    logger.info(f"Parsing NL-44 PDF: {pdf_path} for company: {company_key}")

    company_name = COMPANY_DISPLAY_NAMES.get(company_key, str(company_key).title())

    dedicated_func_name = DEDICATED_PARSER.get(company_key)
    if dedicated_func_name:
        from extractor.companies import PARSER_REGISTRY
        dedicated_func = PARSER_REGISTRY.get(dedicated_func_name)
        if dedicated_func:
            logger.info(f"Routing to dedicated parser: {dedicated_func_name}")
            return dedicated_func(pdf_path, company_key, quarter, year)

    extract = CompanyExtract(
        source_file=Path(pdf_path).name,
        company_key=company_key,
        company_name=company_name,
        form_type="NL44",
        quarter=quarter,
        year=year,
    )

    period_data = PeriodData(period_label="current")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pg_idx, page in enumerate(pdf.pages):
                table = page.extract_table()
                if not table:
                    continue
                n = _extract_table(table, period_data)
                logger.debug(f"Page {pg_idx}: {n} metrics extracted")
                if n > 0:
                    break

    except Exception as e:
        logger.error(f"Failed to parse {pdf_path}: {e}", exc_info=True)
        extract.extraction_errors.append(str(e))
        return extract

    if not period_data.data:
        logger.warning(f"No metrics extracted from {pdf_path}")
        extract.extraction_warnings.append("No metrics extracted")
    else:
        logger.info(f"Extraction complete: {len(period_data.data)} metrics")

    extract.current_year = period_data
    return extract
