from __future__ import annotations
import json
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from streamlit_autorefresh import st_autorefresh
import gspread
from google.oauth2.service_account import Credentials
from typing import List, Dict

# ==========================================
# Config & verbinding
# ==========================================

# Secrets ophalen — ondersteunt zowel losse key als [vegdel] blok
if "google_service_account" in st.secrets:
    creds_info = dict(st.secrets["google_service_account"])
else:
    st.stop()

SHEET_ID = (
    st.secrets.get("SHEET_ID")
    or st.secrets.get("vegdel", {}).get("SHEET_ID")
)

SHEET_ID_RAW = SHEET_ID

def _normalize_sheet_id(val: str | None) -> str | None:
    if not val:
        return None
    v = str(val).strip()
    if v.startswith("http://") or v.startswith("https://"):
        try:
            part = v.split("/d/")[1]
            return part.split("/")[0]
        except Exception:
            return None
    return v

SHEET_ID = _normalize_sheet_id(SHEET_ID_RAW)

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
client = gspread.authorize(creds)

# ==========================================
# Diagnose-tab setup
# ==========================================

def diag_ping() -> str:
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.sheet1
        ws.append_row(["diag", "ok"], value_input_option="RAW")
        return f"✅ Verbinding OK met {SHEET_ID}"
    except Exception as e:
        return f"❌ Fout tijdens verbindingstest: {e}"

# Streamlit UI
st.title("Vegdel Diagnosetool 🧰")

st.write("SHEETS_ID (raw):", SHEET_ID_RAW if SHEET_ID_RAW else "(niet gezet)")
st.write("SHEETS_ID (parsed):", SHEET_ID if SHEET_ID else "(kon niet parsen)")
sa_email = creds_info.get("client_email")
st.write("Service account:", sa_email)

if st.button("🔌 Test verbinding (ping)"):
    msg = diag_ping()
    if msg.startswith("✅"):
        st.success(msg)
    else:
        st.error(msg)
