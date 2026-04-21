"""
pipeline.py — Config-driven NL-44 extraction pipeline.

Usage:
  python3 pipeline.py                          # incremental (default)
  python3 pipeline.py --force                  # re-extract everything
  python3 pipeline.py --force-company bajaj_allianz
  python3 pipeline.py --quarter Q1 Q3          # override config quarters
  python3 pipeline.py --dry-run
  python3 pipeline.py --skip-consolidated
"""

import argparse
import logging
import os
import sys
import yaml

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from extractor.path_scanner import scan
from extractor.consolidated_detector import find_nl44_pages, extract_nl44_to_temp
from extractor.processed_log import load as load_log, save as save_log
from extractor.processed_log import filter_unprocessed, mark_processed
from extractor.parser import parse_pdf
from validation.checks import run_validations, write_validation_report
from output.excel_writer import (
    save_workbook,
    write_validation_summary_sheet,
    write_validation_detail_sheet,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-5s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "extraction_config.yaml")


def load_config(config_path: str) -> dict:
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def validate_config(config: dict) -> None:
    for key in ("base_path", "master_sheet_path", "processed_log_path"):
        if not config.get(key, "").strip():
            raise ValueError(f"{key} is not set in extraction_config.yaml")


def main():
    parser = argparse.ArgumentParser(description="NL-44 Extraction Pipeline")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--force-company", type=str, default=None)
    parser.add_argument("--quarter", nargs="+", default=None)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--skip-consolidated", action="store_true")
    parser.add_argument("--config", type=str, default=CONFIG_PATH)
    args = parser.parse_args()

    try:
        config = load_config(args.config)
        validate_config(config)
    except (FileNotFoundError, ValueError) as e:
        logger.error(str(e))
        sys.exit(1)

    if args.quarter:
        config["quarters"] = args.quarter
    if args.skip_consolidated:
        config["consolidated_mode"] = "skip"

    logger.info("Scanning folder structure...")
    try:
        scan_results = scan(config)
    except (FileNotFoundError, ValueError) as e:
        logger.error(str(e))
        sys.exit(1)

    if not scan_results:
        logger.info("No PDFs found. Check base_path and fiscal_years.")
        sys.exit(0)

    log_path = config["processed_log_path"]
    log_data = load_log(log_path)
    to_process = filter_unprocessed(
        scan_results, log_data,
        force=args.force,
        force_company=args.force_company,
    )

    if not to_process:
        logger.info("All files up-to-date. Use --force to re-extract.")
        sys.exit(0)

    # --- Skip companies ---
    skip_set = set(config.get("skip_companies", []))
    if skip_set:
        skipped = [r for r in to_process if r.company_key in skip_set]
        to_process = [r for r in to_process if r.company_key not in skip_set]
        for r in skipped:
            logger.info(f"  Skipping [{r.company_key}] — in skip_companies list")

    if args.dry_run:
        logger.info(f"DRY RUN — would extract {len(to_process)} files:")
        for r in to_process:
            logger.info(f"  [{r.source_type:12s}] {r.company_key:30s} "
                        f"{r.fiscal_year} {r.quarter}  {r.pdf_path}")
        sys.exit(0)

    master_path = config["master_sheet_path"]
    consolidated_mode = config.get("consolidated_mode", "dynamic")
    succeeded = 0
    failed = 0
    failed_files = []
    all_extractions = []
    all_validation_results = []

    for result in to_process:
        logger.info(
            f"Extracting [{result.source_type}] {result.company_key} "
            f"{result.fiscal_year} {result.quarter}"
        )

        pdf_to_parse = result.pdf_path
        temp_file = None

        try:
            if result.source_type == "consolidated" and consolidated_mode != "skip":
                page_overrides = config.get("nl44_page_overrides", {})
                override = page_overrides.get(result.company_key) or {}
                keywords = config.get("nl44_keywords", None)

                if "start" in override:
                    start_0 = int(override["start"]) - 1
                    end_0   = int(override.get("end", override["start"])) - 1
                    pages   = (start_0, end_0)
                    logger.info(f"  Hard page override: pages {start_0}-{end_0}")
                else:
                    min_matches = int(override.get(
                        "min_matches",
                        config.get("nl44_keyword_min_matches", 3),
                    ))
                    pages = find_nl44_pages(result.pdf_path, keywords, min_matches)

                if pages is None:
                    logger.warning(f"  NL-44 section not found, skipping: {result.pdf_path}")
                    failed += 1
                    failed_files.append((result.pdf_path, "NL-44 section not found"))
                    continue

                temp_file = extract_nl44_to_temp(result.pdf_path, pages[0], pages[1])
                if temp_file is None:
                    logger.warning(f"  Page extraction failed, skipping: {result.pdf_path}")
                    failed += 1
                    failed_files.append((result.pdf_path, "Page extraction failed"))
                    continue

                pdf_to_parse = temp_file

            extract = parse_pdf(pdf_to_parse, result.company_key, result.quarter, result.year_code)
            validation_results = run_validations([extract])

            all_extractions.append(extract)
            all_validation_results.extend(validation_results)

            mark_processed(log_data, result, 0)
            save_log(log_path, log_data)

            succeeded += 1
            logger.info(f"  Done: {result.company_key}")

        except Exception as e:
            logger.error(f"  Failed: {result.company_key} — {e}", exc_info=True)
            failed += 1
            failed_files.append((result.pdf_path, str(e)))

        finally:
            if temp_file and os.path.exists(temp_file):
                os.unlink(temp_file)

    if all_extractions:
        stats = {
            "files_processed": succeeded + failed,
            "files_succeeded": succeeded,
            "files_failed": failed,
        }
        logger.info(f"Writing master workbook to {master_path}")
        save_workbook(all_extractions, master_path, stats=stats)

        report_path = os.path.join(
            os.path.dirname(master_path), "NL44_validation_report.csv"
        )
        write_validation_report(all_validation_results, report_path)
        write_validation_summary_sheet(report_path, master_path)
        write_validation_detail_sheet(report_path, master_path)
        logger.info(f"Validation report: {report_path}")

    logger.info("=" * 60)
    logger.info("Pipeline complete.")
    logger.info(f"  Succeeded: {succeeded}")
    logger.info(f"  Failed:    {failed}")
    if failed_files:
        for path, reason in failed_files:
            logger.info(f"    {path} — {reason}")


if __name__ == "__main__":
    main()
