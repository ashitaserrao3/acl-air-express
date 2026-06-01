import pandas as pd
 
def get_slab(wt):
 
    if pd.isna(wt):
        return None
 
    try:
        wt = float(wt)
 
        if wt < 45:
            return "Less than 45Kg"
 
        elif wt < 100:
            return "45Kg +"
 
        elif wt < 250:
            return "100Kg +"
 
        elif wt < 300:
            return "250Kg +"
 
        elif wt < 500:
            return "300Kg +"
 
        elif wt < 1000:
            return "500Kg +"
 
        return "1000Kg +"
 
    except:
        return None
 