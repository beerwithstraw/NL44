"""
Dedicated NL-44 parser for Royal Sundaram General Insurance Co. Ltd.

Differences from the generic parser:
  1. Handles "shifted" layout where values are in the header rows.
  2. Converts units from Rs. to Lakh Rs. (divides by 100,000).
"""

import logging
from pathlib import Path

import pdfplumber

from config.company_registry import COMPANY_DISPLAY_NAMES
from extractor.models import CompanyExtract, PeriodData
from extractor.normaliser import clean_number

logger = logging.getLogger(__name__)


def parse_royal_sundaram(pdf_path: str, company_key: str,
                         quarter: str = "", year: str = "") -> CompanyExtract:
    company_name = COMPANY_DISPLAY_NAMES.get(company_key, "Royal Sundaram General Insurance")
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
            page = pdf.pages[0]
            table = page.extract_table()
            if table:
                _extract_royal_table(table, period_data)

    except Exception as e:
        logger.error(f"Failed to parse {pdf_path}: {e}", exc_info=True)
        extract.extraction_errors.append(str(e))
        return extract

    if not period_data.data:
        logger.warning(f"No metrics extracted from {pdf_path}")
        extract.extraction_warnings.append("No metrics extracted")
    else:
        logger.info(f"Extraction complete: {len(period_data.data)} metrics (converted from Rs. to Lacs)")

    extract.current_year = period_data
    return extract


def _extract_royal_table(table, period_data: PeriodData):
    """
    Royal Sundaram specific logic:
    - Handles shifted layout.
    - Dynamic unit conversion: if any value > 100,000,000 (10 Crore), 
      assumes Rs. and divides by 100,000 to get Lacs.
    """
    
    # First pass: collect raw values and determine if conversion is needed
    raw_data = []
    max_val = 0
    
    row_count = len(table)
    for i, row in enumerate(table):
        if not row or len(row) < 3:
            continue
            
        label = str(row[0] or "").strip().upper()
        qtr_val = clean_number(row[1])
        ytd_val = clean_number(row[2])
        
        if qtr_val is None and ytd_val is None:
            continue
            
        metric_key = None
        
        if "GROSS DIRECT MOTOR THIRD PARTY INSURANCE BUSINESS" in label:
            if i + 1 < row_count:
                next_label = str(table[i+1][0] or "").upper()
                if "LIABILITY ONLY POLICIES (L)" in next_label:
                    metric_key = "liability_premium"
                elif "PACKAGE POLICIES (P)" in next_label:
                    metric_key = "package_premium"
        elif "TOTAL GROSS DIRECT MOTOR THIRD PARTY INSURANCE" in label:
            metric_key = "total_mtp_premium"
        elif "TOTAL GROSS DIRECT MOTOR OWN DAMAGE" in label:
            metric_key = "motor_od_premium"
        elif "TOTAL GROSS DIRECT PREMIUM INCOME" in label:
            metric_key = "total_premium"
            
        if metric_key:
            raw_data.append((metric_key, qtr_val, ytd_val))
            if qtr_val: max_val = max(max_val, abs(qtr_val))
            if ytd_val: max_val = max(max_val, abs(ytd_val))

    # Second pass: apply conversion and store
    # Threshold: 100,000,000 (10 Crore). If a value is larger, it's likely in Rs.
    needs_conversion = max_val > 100000000
    divisor = 100000.0 if needs_conversion else 1.0
    
    if needs_conversion:
        logger.info(f"  Detected large values (max={max_val:,.0f}). Converting from Rs. to Lacs.")

    for metric_key, qtr_val, ytd_val in raw_data:
        qtr_final = qtr_val / divisor if qtr_val is not None else None
        ytd_final = ytd_val / divisor if ytd_val is not None else None
        
        period_data.data[metric_key] = {
            "qtr": qtr_final,
            "ytd": ytd_final,
        }
        logger.debug(f"  {metric_key}: qtr={qtr_final} ytd={ytd_final}")
