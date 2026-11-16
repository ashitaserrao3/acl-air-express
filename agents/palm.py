"""
PALM bills (PDF, Palm Express Cargo Service, Hyderabad).

One table per invoice: SNo  Date  Cnsg No  Sec  Pcs  Wt  Rate  Total Price.
Rules kept from the original PALM processor: origin HYD, transport AIR, every
consignment Console. The bill has no flight column, so the airline is shown as
"Not on bill" instead of being treated as unknown.
"""

import re
from io import BytesIO

import pandas as pd
import pdfplumber

from utils.common import _one_num, parse_date, strip_ext

AGENT = "PALM"
KEY = "palm"
LABEL = "PALM"
DESCRIPTION = "Palm Express PDF invoices · origin HYD · no flight on the bill"
FILE_TYPES = ["pdf"]
OUTPUT_FILE = "PALM_Combined.xlsx"
EXTRA_COLS = []

ORIGIN = "HYD"
NO_AIRLINE = "Not on bill"
_INVOICE = re.compile(r"\b([A-Z]{1,3}/\d+/\d{2}-\d{2})\b")
_SUBTOTAL = re.compile(r"Subtotal\s*₹?\s*([\d,]+(?:\.\d+)?)", re.I)


def _num(x):
    return _one_num(str(x).replace(",", ""))


def _parse_line(parts):
    """'1 10.4.26 12391 MAA 1 15 50.00 750.00' -> dict, or None when it is not an AWB line."""
    if len(parts) != 8 or not parts[0].isdigit() or not parts[2].isdigit() or not re.fullmatch(r"[A-Z]{3}", parts[3]):
        return None
    date = parse_date(parts[1])
    nums = [_num(p) for p in parts[4:8]]
    if pd.isna(date) or any(n is None for n in nums):
        return None
    pcs, wt, rate, total = nums
    return {"AWB_NO": parts[2], "AWB_DATE": date, "DEST": parts[3], "PKGS": pcs, "CHR_WT": wt,
            "RATE": rate, "BASIC_FRT": total, "TOTAL_FRT": total}


def parse(data, name, log):
    with pdfplumber.open(BytesIO(data)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    lines = [ln.split() for ln in text.split("\n")]
    parsed = [_parse_line(parts) for parts in lines]
    df = pd.DataFrame([r for r in parsed if r])
    if df.empty:
        return df

    # the bill's own total line (Pcs, Wt) just below the table, and its Subtotal (₹)
    last_row = max(i for i, r in enumerate(parsed) if r)
    for parts in lines[last_row + 1:last_row + 3]:
        if len(parts) == 2 and all(_num(p) is not None for p in parts):
            pcs, wt = (_num(p) for p in parts)
            if (pcs, wt) != (df["PKGS"].sum(), df["CHR_WT"].sum()):
                log("warning", f"{name}: the bill's total line says {pcs:g} pcs / {wt:g} kg, but its rows add up to "
                               f"{df['PKGS'].sum():g} pcs / {df['CHR_WT'].sum():g} kg – please check the bill")
            break
    m = _SUBTOTAL.search(text)
    if m and abs(_num(m.group(1)) - df["TOTAL_FRT"].sum()) > 1:
        log("warning", f"{name}: Subtotal on the bill is {m.group(1)}, but the rows add up to "
                       f"{df['TOTAL_FRT'].sum():,.2f} – please check the bill")

    mismatch = (df["CHR_WT"] * df["RATE"] - df["TOTAL_FRT"]).abs() > 1
    df["DATA_ISSUE"] = pd.Series(pd.NA, index=df.index, dtype="string")
    df.loc[mismatch, "DATA_ISSUE"] = "Wt × Rate ≠ Total Price (check bill)"

    inv = _INVOICE.search(text)
    df["INVOICE_NO"] = inv.group(1) if inv else strip_ext(name)
    df["ORIGIN"] = ORIGIN
    df["TRNSPT_MODE"] = "AIR"
    df["AIRLINE"] = NO_AIRLINE
    df["LODGE_MODE"] = "Console"
    return df
