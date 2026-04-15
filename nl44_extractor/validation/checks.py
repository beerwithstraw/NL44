"""
Validation Checks for NL-44 (Motor TP Obligations).

Check families:
  1. COMPLETENESS  — all five metrics must be present
  2. TOTAL_SUM     — Liability (L) + Package (P) ≈ Total MTP (L+P)
                     checked for both QTR and YTD
"""

import csv
import logging
from typing import List, Optional
from dataclasses import dataclass, asdict

from extractor.models import CompanyExtract
from config.row_registry import ROW_ORDER

logger = logging.getLogger(__name__)

TOLERANCE = 2.0


@dataclass
class ValidationResult:
    company: str
    quarter: str
    year: str
    period: str        # "qtr" or "ytd"
    check_name: str
    status: str        # PASS / WARN / FAIL
    expected: Optional[float]
    actual: Optional[float]
    delta: Optional[float]
    note: str


def run_validations(extractions: List[CompanyExtract]) -> List[ValidationResult]:
    results: List[ValidationResult] = []
    for exc in extractions:
        if not exc.current_year:
            continue
        results.extend(_check_completeness(exc))
        for period in ("qtr", "ytd"):
            r = _check_lp_sum(exc, period)
            if r:
                results.append(r)
    return results


def _get(exc: CompanyExtract, metric: str, period: str) -> Optional[float]:
    v = exc.current_year.data.get(metric, {}).get(period)
    return float(v) if v is not None else None


def _check_completeness(exc: CompanyExtract) -> List[ValidationResult]:
    results = []
    for metric in ROW_ORDER:
        has_any = any(
            exc.current_year.data.get(metric, {}).get(p) is not None
            for p in ("qtr", "ytd")
        )
        if not has_any:
            results.append(ValidationResult(
                exc.company_name, exc.quarter, exc.year,
                "both", "COMPLETENESS", "WARN",
                expected=None, actual=None, delta=None,
                note=f"Metric '{metric}' is missing",
            ))
    return results


def _check_lp_sum(exc: CompanyExtract, period: str) -> Optional[ValidationResult]:
    """Liability (L) + Package (P) should equal Total MTP (L+P)."""
    l_val = _get(exc, "liability_premium", period)
    p_val = _get(exc, "package_premium", period)
    lp_val = _get(exc, "total_mtp_premium", period)

    if l_val is None or p_val is None or lp_val is None:
        return None

    expected = l_val + p_val
    delta = abs(lp_val - expected)
    status = "PASS" if delta <= TOLERANCE else "FAIL"
    period_label = "For the Quarter" if period == "qtr" else "Upto the Quarter"

    return ValidationResult(
        exc.company_name, exc.quarter, exc.year,
        period_label, f"TOTAL_SUM_MTP_{period.upper()}", status,
        expected=expected, actual=lp_val, delta=delta, note="",
    )


def write_validation_report(results: List[ValidationResult], output_path: str):
    with open(output_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "company", "quarter", "year", "period",
            "check_name", "status", "expected", "actual", "delta", "note",
        ])
        writer.writeheader()
        for r in results:
            writer.writerow(asdict(r))
    logger.info(f"Validation report saved to {output_path}")
