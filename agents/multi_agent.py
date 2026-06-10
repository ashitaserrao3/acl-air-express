import streamlit as st
 
def run():
 
    st.markdown("## 🚀 Multi Agent Processor")
 
    st.info(
        "Upload files from one or more agents and process them together."
    )
 
    st.markdown("### Upload Files")
 
    pobc_files = st.file_uploader(
        "📦 POBC Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_pobc"
    )
 
    pcf_south_files = st.file_uploader(
        "📦 PCF South Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_pcf_south"
    )
 
    pcfl_east_files = st.file_uploader(
        "📦 PCFL East Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_pcfl_east"
    )
 
    surya_files = st.file_uploader(
        "📦 Surya Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_surya"
    )
 
    bhagwati_files = st.file_uploader(
        "📦 Bhagwati Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_bhagwati"
    )
 
    palm_files = st.file_uploader(
        "📦 Palm Files",
        type=["pdf"],
        accept_multiple_files=True,
        key="multi_palm"
    )
 
    fdc_files = st.file_uploader(
        "📦 FDC Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="multi_fdc"
    )
 
    if st.button("🚀 Process All"):
 
        st.success("Ready for processing...")