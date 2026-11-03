"""
Shared parser for PCF South and PCFL East bills.
The two formats are identical except for which column holds the total
(South = NET, East = GROSS), so the logic lives here once.
"""

import re

import pandas as pd

from utils.common import digits_only, find_col, find_header_row, frame_from_header, strip_ext, to_date, to_num
from utils.excel_utils import read_raw_any

EXTRA_COLS = ["AWBDO", "FSC", "OCDC", "TSP", "SSP", "SSP_AMT", "DIS%", "DISC", "CGST", "SGST", "IGST", "AMOUNT"]

# output column -> (header candidates, allow partial match)
COLUMN_MAP = {
    "AWB_NO": (["AWB NO", "AWB"], True),
    "AWB_DATE": (["DATE"], True),
    "DEST": (["DEST"], True),
    "FLIGHT_NO": (["FLIGHT"], True),
    "PKGS": (["PKT", "PKTS"], True),
    "CHR_WT": (["WGT", "WEIGHT"], True),
    "BASIC_FRT": (["BASIC"], True),
    "RATE": (["RATE"], True),
    "AWBDO": (["AWBDO", "AWB DO"], False),
    "FSC": (["FSC"], True),
    "OCDC": (["OCDC"], True),
    "TSP": (["TSP"], False),
    "SSP": (["SSP"], False),
    "SSP_AMT": (["SSP AMT", "SSPAMT"], True),
    "DIS%": (["DIS%", "DIS"], False),
    "DISC": (["DISC"], False),
    "CGST": (["CGST"], True),
    "SGST": (["SGST"], True),
    "IGST": (["IGST"], True),
    "AMOUNT": (["AMOUNT"], False),
}


_INVOICE = re.compile(r"INVOICE\s*NO\.?\s*[:\-]*\s*([A-Z]{3}\s*[/\-][^\s,]+)", re.I)


def invoice_from_header(raw, header_row):
    """'INVOICE NO.    : MAA/595/26-27' above the table -> 'MAA/595/26-27' (None if not there)."""
    for r in range(header_row):
        for v in raw.iloc[r].dropna():
            m = _INVOICE.search(str(v))
            if m:
                return re.sub(r"\s+", "", m.group(1)).upper()
    return None


def extract_origin(invoice):
    first = str(invoice).split("-")[0].strip().upper()
    return first if "-" in str(invoice) and first.isalpha() else None


def compute_mode(df):
    awbdo = df["AWBDO"].fillna(0)
    others = df[["FSC", "OCDC", "TSP"]].fillna(0).sum(axis=1)
    return ((awbdo - 50).abs() < 0.01) & (others == 0)


def parse_pcf(data, name, log, total_header, read_invoice_header=False):
    """
    read_invoice_header (PCF South): invoice no. and origin come from the bill's
    'INVOICE NO. : MAA/595/26-27' line, so the file name does not matter; the file
    name (MAA-PCF-0801.xlsx) is only the fallback. Off = file name only (PCFL East).
    """
    raw, _, _ = read_raw_any(data, name)
    header_row = find_header_row(
        raw, lambda cells: any("AWB" in c for c in cells) and any("DATE" in c for c in cells)
    )
    if header_row is None:
        raise ValueError("header row (AWB + DATE) not found")
    df = frame_from_header(raw, header_row)

    out = pd.DataFrame(index=df.index)
    for target, (cands, partial) in COLUMN_MAP.items():
        src = find_col(df, cands, allow_partial=partial)
        out[target] = df[src] if src is not None else pd.NA

    # AWB column must not be AWBDO
    if find_col(df, ["AWB NO", "AWB"], True) == find_col(df, ["AWBDO", "AWB DO"], False):
        raise ValueError("could not tell AWB No column apart from AWBDO")

    total_src = find_col(df, [total_header], True)
    if total_src is None:
        raise ValueError(f"total column '{total_header}' not found")
    out["TOTAL_FRT"] = df[total_src]

    out["AWB_NO"] = out["AWB_NO"].map(digits_only)
    # real AWB rows: an AWB number (PCF East uses short serials like 60203) AND an amount
    has_amount = to_num(out["TOTAL_FRT"]).notna() | to_num(out["BASIC_FRT"]).notna()
    out = out[out["AWB_NO"].notna() & (out["AWB_NO"].str.len() >= 4) & has_amount].copy()

    # "M" in weight = minimum charge. Kept as a note, weight left blank.
    wt_text = out["CHR_WT"].astype(str).str.strip().str.upper()
    is_min = wt_text == "M"
    out["CHR_WT"] = to_num(out["CHR_WT"].where(~is_min))
    out["DATA_ISSUE"] = pd.Series(pd.NA, index=out.index, dtype="string")
    out.loc[is_min, "DATA_ISSUE"] = "Weight shown as M (minimum)"
    if is_min.any():
        log("info", f"{name}: {int(is_min.sum())} rows with weight 'M'")

    for col in ["PKGS", "BASIC_FRT", "RATE", "TOTAL_FRT"] + EXTRA_COLS:
        out[col] = to_num(out[col])

    out["AWB_DATE"] = to_date(out["AWB_DATE"])
    out["LODGE_MODE"] = compute_mode(out).map({True: "Console", False: "Direct"})

    invoice = invoice_from_header(raw, header_row) if read_invoice_header else None
    if invoice:
        out["INVOICE_NO"] = invoice
        out["ORIGIN"] = invoice[:3]
    else:
        out["INVOICE_NO"] = strip_ext(name)
        out["ORIGIN"] = extract_origin(out["INVOICE_NO"].iloc[0]) if len(out) else None
        if read_invoice_header and len(out) and out["ORIGIN"].isna().all():
            log("warning", f"{name}: origin not found – no 'INVOICE NO. : MAA/…' line in the bill and the file "
                           "name does not start with a city code (e.g. MAA-…). Rows are left out until it is known.")
    out["TRNSPT_MODE"] = "AIR"
    return out
