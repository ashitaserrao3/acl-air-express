import streamlit as st
import pandas as pd
 
from utils.slab import get_slab
from utils.airline import get_airline
from utils.bill_period import get_bill_period
from utils.exporter import export_excel
from utils.excel_utils import unmerge_excel
from utils.dashboard import show_dashboard
 
 
# ---------------------------------------------------
# HEADER DETECTION
# ---------------------------------------------------
def find_header_row(df_raw):
 
    for idx, row in df_raw.iterrows():
 
        if any(
            "awb" in str(cell).lower()
            and "no" in str(cell).lower()
            for cell in row
        ):
            return idx
 
    return None
 
 
# ---------------------------------------------------
# LODGE MODE
# ---------------------------------------------------
def get_mode(row):
 
    mode_cols = [
        "xray",
        "Tsp in",
        "Tsp out",
        "addntsp",
        "Addntsp Destination",
        "Unit",
        "Deunit",
        "Surch.",
        "Misc Chg.",
        "Other Dc",
        "Service",
        "Handling"
    ]
 
    def is_zero(v):
 
        try:
            return float(v) == 0
 
        except:
            return True
 
    if all(
        is_zero(row.get(c))
        for c in mode_cols
        if c in row.index
    ):
        return "Console"
 
    return "Direct"
 
 
# ---------------------------------------------------
# MAIN RUN FUNCTION
# ---------------------------------------------------
def run():
 
    st.title("📦 INDEX Processor")
 
    uploaded_files = st.file_uploader(
        "Upload INDEX Excel Files",
        type=["xlsx"],
        accept_multiple_files=True
    )
 
    if not uploaded_files:
        return
 
    # ---------------------------------------------------
    # CLEAR SESSION
    # ---------------------------------------------------
    #if st.button("Clear Processed Data"):
 
        if "index_df" in st.session_state:
            del st.session_state["index_df"]
 
        st.rerun()
 
    # ---------------------------------------------------
    # PROCESS ONLY ONCE
    # ---------------------------------------------------
    if "index_df" not in st.session_state:
 
        combined_df = pd.DataFrame()
 
        for file in uploaded_files:
 
            st.write(f"Processing: {file.name}")
 
            temp = unmerge_excel(file)
 
            df_raw = pd.read_excel(
                temp,
                header=None
            )
 
            header_row = find_header_row(df_raw)
 
            if header_row is None:
 
                st.warning(
                    f"Header not found in {file.name}"
                )
 
                continue
 
            # ---------------------------------------------------
            # LOAD CLEAN DF
            # ---------------------------------------------------
            temp.seek(0)
 
            df = pd.read_excel(
                temp,
                header=header_row
            )
 
            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.replace("\n", " ")
                .str.replace(r"\s+", " ", regex=True)
            )
 
            # ---------------------------------------------------
            # AWB CLEANING
            # ---------------------------------------------------
            awb_col = [
                c for c in df.columns
                if "AWB" in c.upper()
                and "NO" in c.upper()
            ]
 
            if not awb_col:
                continue
 
            df["Awb No"] = (
                df[awb_col[0]]
                .astype(str)
                .str.strip()
            )
 
            df = df[
                (df["Awb No"] != "")
                &
                (df["Awb No"].str.lower() != "nan")
                &
                (df["Awb No"].str.len() > 5)
            ]
 
            # ---------------------------------------------------
            # DATE
            # ---------------------------------------------------
            date_col = [
                c for c in df.columns
                if "DATE" in c.upper()
            ]
 
            if date_col:
 
                df["Awb Date"] = pd.to_datetime(
                    df[date_col[0]],
                    errors="coerce"
                )
 
            else:
                df["Awb Date"] = None
 
            # ---------------------------------------------------
            # AIRLINE
            # ---------------------------------------------------
            flight_col = [
                c for c in df.columns
                if "FLIGHT" in c.upper()
            ]
 
            if flight_col:
 
                df["Flight"] = df[flight_col[0]]
 
                df["Airline"] = df["Flight"].apply(
                    get_airline
                )
 
            else:
 
                df["Flight"] = None
                df["Airline"] = None
 
            # ---------------------------------------------------
            # NUMERIC CONVERSIONS
            # ---------------------------------------------------
            numeric_cols = [
                "Charge Wt.",
                "Total",
                "Basic Freight",
                "Rate",
                "Pkts"
            ]
 
            for col in numeric_cols:
 
                if col in df.columns:
 
                    df[col] = pd.to_numeric(
                        df[col],
                        errors="coerce"
                    )
 
            # ---------------------------------------------------
            # OD PAIR
            # ---------------------------------------------------
            df["Origin"] = (
                df.get("Origin", "")
                .astype(str)
            )
 
            df["Dest."] = (
                df.get("Dest.", "")
                .astype(str)
            )
 
            df["OD-Pair"] = (
                df["Origin"].str.strip()
                + "-"
                + df["Dest."].str.strip()
            )
 
            # ---------------------------------------------------
            # SLAB
            # ---------------------------------------------------
            df["Slab"] = df["Charge Wt."].apply(
                get_slab
            )
 
            # ---------------------------------------------------
            # CALCULATIONS
            # ---------------------------------------------------
            df["AddOns"] = (
                df["Total"]
                - df["Basic Freight"]
            )
 
            df["AddOn/Kg"] = (
                df["AddOns"]
                / df["Charge Wt."]
            ).round(2)
 
            df["CPKG"] = (
                df["Total"]
                / df["Charge Wt."]
            ).round(2)
 
            # ---------------------------------------------------
            # LODGE MODE
            # ---------------------------------------------------
            df["Mode"] = df.apply(
                get_mode,
                axis=1
            )
 
            # ---------------------------------------------------
            # STANDARD OUTPUT
            # ---------------------------------------------------
            df["INVOICE_NO"] = (
                file.name.replace(".xlsx", "")
            )
 
            df["BILL_PERIOD"] = (
                df["Awb Date"]
                .apply(get_bill_period)
            )
 
            df["AGENT"] = "INDEX"
 
            df["TRNSPT_MODE"] = "AIR"
 
            df["OD_PAIR"] = df["OD-Pair"]
 
            df["ORIGIN"] = df["Origin"]
 
            df["DEST"] = df["Dest."]
 
            df["AWB_NO"] = df["Awb No"]
 
            df["AWB_DATE"] = (
                df["Awb Date"]
                .dt.strftime("%d-%b-%Y")
            )
 
            df["FLIGHT_NO"] = df["Flight"]
 
            df["AIRLINE"] = df["Airline"]
 
            df["PKGS"] = df["Pkts"]
 
            df["CHR_WT"] = df["Charge Wt."]
 
            df["RATE"] = df["Rate"]
 
            df["BASIC_FRT"] = df["Basic Freight"]
 
            df["TOTAL_FRT"] = df["Total"]
 
            df["ADDON_CHR"] = df["AddOns"]
 
            df["ADDON_PER_KG"] = df["AddOn/Kg"]
 
            df["LODGE_MODE"] = df["Mode"]
 
            df["SLAB"] = df["Slab"]
 
            # ---------------------------------------------------
            # FINAL COLUMNS
            # ---------------------------------------------------
            final_columns = [
 
                "INVOICE_NO",
                "BILL_PERIOD",
                "AGENT",
                "TRNSPT_MODE",
            
                "OD_PAIR",
                "ORIGIN",
                "DEST",
            
                "AWB_NO",
                "AWB_DATE",
            
                "FLIGHT_NO",
                "AIRLINE",
            
                "PKGS",
                "CHR_WT",
            
                "RATE",
                "BASIC_FRT",
            
                "A.Do",
                "xray",
                "Tsp in",
                "Tsp out",
                "addntsp",
                "Addntsp Destination",
                "Unit",
                "Deunit",
                "Surch.",
                "Misc Chg.",
                "Other Dc",
                "Service",
                "Handling",
            
                "TOTAL_FRT",
            
                "CPKG",
            
                "ADDON_CHR",
                "ADDON_PER_KG",
            
                "SLAB",
                "LODGE_MODE"
            ]
 
            for col in final_columns:
 
                if col not in df.columns:
                    df[col] = None
 
            df = df[final_columns]
 
            # ---------------------------------------------------
            # COMBINE
            # ---------------------------------------------------
            combined_df = pd.concat(
                [combined_df, df],
                ignore_index=True
            )
 
        # ---------------------------------------------------
        # STORE SESSION
        # ---------------------------------------------------
        combined_df = combined_df.drop_duplicates(
            subset=["AWB_NO"]
        )
 
        st.session_state["index_df"] = combined_df
 
    # ---------------------------------------------------
    # LOAD SESSION DATA
    # ---------------------------------------------------
    combined_df = st.session_state["index_df"]
 
    if combined_df.empty:
 
        st.warning("No valid data found.")
        return
 
    # ---------------------------------------------------
    # SUCCESS
    # ---------------------------------------------------
    st.success(
        f"✅ Combined {len(combined_df)} rows"
    )
 
    # ---------------------------------------------------
    # DASHBOARD
    # ---------------------------------------------------
    filtered_df = show_dashboard(combined_df)
 
    # ---------------------------------------------------
    # DATAFRAME
    # ---------------------------------------------------
    st.dataframe(
        filtered_df,
        use_container_width=True
    )
 
    # ---------------------------------------------------
    # EXPORT
    # ---------------------------------------------------
    output = export_excel(
        combined_df,
        sheet_name="INDEX"
    )
 
    st.download_button(
        label="📥 Download Excel",
        data=output,
        file_name="INDEX_Combined.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
 