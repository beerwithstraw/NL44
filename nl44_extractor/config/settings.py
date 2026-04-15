"""
Global constants for the NL-44 Extractor.

Master sheet layout follows the same Quarter_Info pattern as NL-39:
  - One row per Quarter_Info value ("For the Quarter" / "Upto the Quarter")
  - The five NL-44 items appear as metric columns
  - col_name.lower() must equal the metric key in row_registry.ROW_ORDER
"""

EXTRACTOR_VERSION = "1.0.0"
NUMBER_FORMAT = "#,##0.00"
LOW_CONFIDENCE_FILL_COLOR = "FFFF99"

MASTER_COLUMNS = [
    "Company_Name",             # A
    "Company",                  # B
    "NL",                       # C  — always "NL44"
    "Quarter",                  # D
    "Year",                     # E
    "Quarter_Info",             # F  — "For the Quarter" / "Upto the Quarter"
    "Sector",                   # G
    "Industry_Competitors",     # H
    "GI_Companies",             # I
    # --- Five NL-44 items (col .lower() == metric key) ---
    "Liability_Premium",        # J
    "Package_Premium",          # K
    "Total_MTP_Premium",        # L
    "Motor_OD_Premium",         # M
    "Total_Premium",            # N
    "Source_File",              # O
]


def company_key_to_pascal(key: str) -> str:
    return key.replace("_", " ").title().replace(" ", "")
