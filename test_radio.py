import streamlit as st

st.set_page_config(layout="wide")

# Test CSS
st.markdown("""
<style>
    /* Try different selectors for radio buttons */
    div[data-testid="stRadio"] > div > div > div > div input[type="radio"]:checked + div {
        background-color: #000000 !important;
    }
    
    div[data-testid="stRadio"] label[data-baseweb="radio"] > div > div {
        border-color: #5b7ed8 !important;
    }
    
    div[data-testid="stRadio"] label[data-baseweb="radio"] input:checked ~ div > div {
        background-color: #000000 !important;
    }
</style>
""", unsafe_allow_html=True)

data_source = st.radio(
    "Load data from",
    ["Upload Excel file", "Connect Google Sheets"],
    index=0,
)

st.write(f"Selected: {data_source}")
