import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
 
from agents import (
    index,
    pobc,
    pcf_south,
    pcfl_east,
    surya,
    bhagwati,
    palm,
    fdc
)
 
# =====================================================
# PAGE CONFIG
# =====================================================
 
st.set_page_config(
    page_title="ACL Air Express",
    layout="wide"
)

st.title("✈️ ACL Air Express")
 
# =====================================================
# LOAD CONFIG
# =====================================================
 
with open("config.yaml") as file:
    config = yaml.load(file, Loader=SafeLoader)
 
# =====================================================
# AUTHENTICATION
# =====================================================
 
authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"]
)
 
try:
 
    authenticator.login()
 
except Exception as e:
 
    st.error(e)
 
# =====================================================
# LOGIN SUCCESS
# =====================================================
 
if st.session_state.get("authentication_status"):
 
    authenticator.logout(
        location="sidebar"
    )
 
    st.sidebar.success(
        f"Welcome {st.session_state['name']}"
    )
 
    
 
    agent = st.sidebar.selectbox(
        "Select Agent",
        [
            "INDEX",
            "POBC",
            "PCF(South)",
            "PCFL(East)",
            "SURYA",
            "BHAGWATI",
            "PALM",
            "FDC"
        ]
    )
 
    if agent == "INDEX":
 
        index.run()
 
    elif agent == "POBC":
 
        pobc.run()
 
    elif agent == "PCF(South)":
 
        pcf_south.run()
 
    elif agent == "PCFL(East)":
 
        pcfl_east.run()
 
    elif agent == "SURYA":
 
        surya.run()
 
    elif agent == "BHAGWATI":
 
        bhagwati.run()
 
    elif agent == "PALM":
 
        palm.run()
 
    elif agent == "FDC":
 
        fdc.run()
 
# =====================================================
# LOGIN FAILED
# =====================================================
 
elif st.session_state.get("authentication_status") is False:
 
    st.error("❌ Username or password is incorrect")
 
# =====================================================
# NO LOGIN YET
# =====================================================
 
elif st.session_state.get("authentication_status") is None:
 
    st.info("🔐 Please login to continue")