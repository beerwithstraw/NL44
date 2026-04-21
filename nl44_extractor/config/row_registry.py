"""
Row Registry for NL-44 (Motor TP Obligations).

The five line items on the form, in PDF order.
"""

ROW_ORDER = [
    "liability_premium",
    "package_premium",
    "total_mtp_premium",
    "motor_od_premium",
    "total_premium",
]

ROW_DISPLAY_NAMES = {
    "liability_premium":   "Gross Direct MTP Premium — Liability Only Policies (L)",
    "package_premium":     "Gross Direct MTP Premium — Package Policies (P)",
    "total_mtp_premium":   "Total Gross Direct Motor Third Party Insurance Business Premium (L+P)",
    "motor_od_premium":    "Total Gross Direct Motor Own Damage Insurance Business Premium",
    "total_premium":       "Total Gross Direct Premium Income",
}

# Normalised label fragments → metric key.
# normalise_text() lowercases and collapses whitespace before lookup.
ROW_ALIASES = {
    # Liability (L)
    "premium in respect of liability only policies (l)":      "liability_premium",
    "liability only policies (l)":                            "liability_premium",

    # Package (P)
    "premium in respect of package policies (p)":             "package_premium",
    "package policies (p)":                                   "package_premium",

    # Total MTP (L+P) — the explicit "(l+p)" row is the reliable anchor
    "business premium (l+p)":                                                "total_mtp_premium",
    "business (l+p)":                                                        "total_mtp_premium",
    "total gross direct motor third party insurance business premium (l+p)": "total_mtp_premium",
    "total gross direct motor third party insurance business (l+p)":         "total_mtp_premium",

    # Motor OD — multi-line cell comes through as single string after normalise_text
    "total gross direct motor own damage insurance business premium": "motor_od_premium",
    "total gross direct motor own damage insurance business\npremium": "motor_od_premium",

    # Total Premium Income
    "total gross direct premium income":                      "total_premium",
}
