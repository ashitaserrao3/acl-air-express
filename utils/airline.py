import pandas as pd
 
def get_airline(flight):
 
    if pd.isna(flight):
        return None
 
    code = str(flight).strip().replace(" ", "")[:2].upper()
 
    return {
        "AI": "AirIndia",
        "IX": "AirIndia",
        "6E": "Indigo",
        "SG": "SpiceJet",
        "S5": "StarAir",
        "QP": "Akasa"
    }.get(code, "Other")