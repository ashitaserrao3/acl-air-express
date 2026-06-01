import streamlit as st
import pandas as pd
 
# =====================================================
# PAGE CONFIG
# =====================================================
 
st.set_page_config(
    page_title="Dashboard",
    layout="wide"
)
 
# =====================================================
# FONT SIZE CONTROLS
# =====================================================
 
METRIC_VALUE_SIZE = 20
METRIC_LABEL_SIZE = 11
 
# =====================================================
# CUSTOM CSS
# =====================================================
 
st.markdown(
    f"""
    <style>
 
    /* =========================
       METRIC VALUE FONT
    ========================= */
 
    [data-testid="stMetricValue"] {{
        font-size: {METRIC_VALUE_SIZE}px;
        font-weight: 600;
    }}
 
    /* =========================
       METRIC LABEL FONT
    ========================= */
 
    [data-testid="stMetricLabel"] {{
        font-size: {METRIC_LABEL_SIZE}px;
    }}
 
    /* =========================
       REMOVE EXTRA TOP PADDING
    ========================= */
 
    .block-container {{
        padding-top: 2rem;
    }}
 
    </style>
    """,
    unsafe_allow_html=True
)
 
# =====================================================
# DASHBOARD FUNCTION
# =====================================================
 
def show_dashboard(df):
 
    # =================================================
    # FILTERS
    # =================================================
 
    col1, col2 = st.columns(2)
 
    with col1:
 
        origin_options = ["All"] + sorted(
            df["ORIGIN"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
 
        selected_origin = st.selectbox(
            "Select Origin",
            origin_options
        )
 
    with col2:
 
        dest_options = ["All"] + sorted(
            df["DEST"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
 
        selected_dest = st.selectbox(
            "Select Destination",
            dest_options
        )
 
    # =================================================
    # FILTER DATA
    # =================================================
 
    filtered_df = df.copy()
 
    if selected_origin != "All":
 
        filtered_df = filtered_df[
            filtered_df["ORIGIN"]
            == selected_origin
        ]
 
    if selected_dest != "All":
 
        filtered_df = filtered_df[
            filtered_df["DEST"]
            == selected_dest
        ]
 
    # =================================================
    # KPI CALCULATIONS
    # =================================================
 
    total_awbs = len(filtered_df)
 
    total_frt = round(
        filtered_df["TOTAL_FRT"].sum(),
        2
    )
 
    total_wt = round(
        filtered_df["CHR_WT"].sum(),
        2
    )
 
    cpkg = round(
        total_frt / total_wt,
        2
    ) if total_wt != 0 else 0
 
    # =================================================
    # OVERVIEW SECTION
    # =================================================
 
    st.subheader("📌 Overview")
 
    with st.container(border=True):
 
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
 
        with kpi1:
 
            st.metric(
                "Total AWBs",
                f"{total_awbs:,}"
            )
 
        with kpi2:
 
            st.metric(
                "Total Freight",
                f"{total_frt:,.2f}"
            )
 
        with kpi3:
 
            st.metric(
                "Total CHR WT",
                f"{total_wt:,.2f}"
            )
 
        with kpi4:
 
            st.metric(
                "CPKG",
                f"{cpkg:,.2f}"
            )
 
    # =================================================
    # SLABWISE SUMMARY
    # =================================================
 
    st.subheader("📌 Slabwise Summary")
 
    if (
        "LODGE_MODE" in filtered_df.columns
        and
        "SLAB" in filtered_df.columns
    ):
 
        slab_mapping = {
 
            "Less than 45Kg": [
                "Less than 45Kg"
            ],
 
            "45Kg +": [
                "45Kg +"
            ],
 
            "Greater than 100Kg": [
                "100Kg +",
                "250Kg +",
                "300Kg +",
                "500Kg +",
                "1000Kg +"
            ]
        }
 
        cols = st.columns(3)
 
        for idx, (
            display_slab,
            actual_slabs
        ) in enumerate(slab_mapping.items()):
 
            slab_df = filtered_df[
                filtered_df["SLAB"]
                .isin(actual_slabs)
            ]
 
            # =============================================
            # CONSOLE
            # =============================================
 
            console_df = slab_df[
                slab_df["LODGE_MODE"]
                .astype(str)
                .str.upper()
                == "CONSOLE"
            ]
 
            console_awbs = len(console_df)
 
            console_total_frt = console_df[
                "TOTAL_FRT"
            ].sum()
 
            console_total_wt = console_df[
                "CHR_WT"
            ].sum()
 
            console_cpkg = round(
                console_total_frt
                / console_total_wt,
                2
            ) if console_total_wt != 0 else 0
 
            # =============================================
            # DIRECT
            # =============================================
 
            direct_df = slab_df[
                slab_df["LODGE_MODE"]
                .astype(str)
                .str.upper()
                == "DIRECT"
            ]
 
            direct_awbs = len(direct_df)
 
            direct_total_frt = direct_df[
                "TOTAL_FRT"
            ].sum()
 
            direct_total_wt = direct_df[
                "CHR_WT"
            ].sum()
 
            direct_cpkg = round(
                direct_total_frt
                / direct_total_wt,
                2
            ) if direct_total_wt != 0 else 0
 
            # =============================================
            # DISPLAY CARD
            # =============================================
 
            with cols[idx]:
 
                with st.container(border=True):
 
                    st.markdown(
                        f"#### 📦 {display_slab}"
                    )
 
                    left, right = st.columns(2)
 
                    # =====================================
                    # DIRECT BOX
                    # =====================================
 
                    with left:
 
                        with st.container(border=True):
 
                            st.markdown(
                                "##### Direct"
                            )
 
                            st.metric(
                                "AWBs",
                                f"{direct_awbs:,}"
                            )
 
                            st.metric(
                                "CPKG",
                                f"{direct_cpkg:,.2f}"
                            )
 
                    # =====================================
                    # CONSOLE BOX
                    # =====================================
 
                    with right:
 
                        with st.container(border=True):
 
                            st.markdown(
                                "##### Console"
                            )
 
                            st.metric(
                                "AWBs",
                                f"{console_awbs:,}"
                            )
 
                            st.metric(
                                "CPKG",
                                f"{console_cpkg:,.2f}"
                            )
 
    # =================================================
    # RETURN FILTERED DATA
    # =================================================
 
    return filtered_df
 