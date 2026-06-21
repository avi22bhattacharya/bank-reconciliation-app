"""Extract the ending balance from a Wells Fargo WellsOne bank statement PDF.

Looks for the account summary table on page 1:
  Account number  Beginning balance  Total credits  Total debits  Ending balance
  4573468451      $0.00              $9,338,178.30  -$9,338,178.30  $0.00

Returns the ending balance as a float.
"""

from __future__ import annotations

import re
from pathlib import Path


def extract_ending_balance(pdf_path: str | Path) -> float:
    """Parse a Wells Fargo bank statement PDF and return the ending balance."""
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber is required: add pdfplumber to requirements.txt")

    with pdfplumber.open(str(pdf_path)) as pdf:
        text = pdf.pages[0].extract_text() or ""

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # Find the header row that contains "Ending balance", then read the data row below it
    for i, line in enumerate(lines):
        if "Ending balance" in line and i + 1 < len(lines):
            data = lines[i + 1]
            # Extract all numeric values that look like dollar amounts (with optional leading -)
            amounts = re.findall(r'-?\$?([\d,]+\.\d{2})', data)
            if amounts:
                return float(amounts[-1].replace(',', ''))

    # Fallback: scan the whole page for "Ending balance" followed by an amount on the same line
    match = re.search(r'Ending balance\s+[\S\s]*?(-?\$?[\d,]+\.\d{2})\s*$',
                      text, re.MULTILINE)
    if match:
        return float(match.group(1).lstrip('$').replace(',', ''))

    raise ValueError(
        "Could not find ending balance in the PDF. "
        "Confirm this is a Wells Fargo WellsOne account statement PDF."
    )
