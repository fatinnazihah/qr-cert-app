import streamlit as st
from google.oauth2 import service_account

try:
    creds_dict = st.secrets["google_service_account"]
    creds = service_account.Credentials.from_service_account_info(creds_dict)
    st.success("✅ Google credentials loaded successfully!")
except Exception as e:
    st.error(f"❌ Failed to load credentials: {e}")
