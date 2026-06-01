import pandas as pd
 
def get_bill_period(date):
 
    if pd.isna(date):
        return None
 
    try:
        return "FFN" if date.day <= 15 else "SFN"
 
    except:
        return None