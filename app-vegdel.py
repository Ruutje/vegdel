from __future__ import annotations
from io import BytesIO
from typing import Dict, List

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

import gspread
from google.oauth2.service_account import Credentials

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet


# ==========
# Config
# ==========
PRIMARY = "#b94216"   # bruin (frikandel)
ACCENT  = "#891418"   # curry
LIGHT   = "#ffffff"   # wit

st.set_page_config(page_title="Vegdel – Frikandel Speciaal", page_icon="🌭", layout="wide")
st.markdown(
    f"""
    <style>
      .stButton>button {{
        background: {PRIMARY};
        color: {LIGHT};
        border: 0;
        border-radius: 10px;
        padding: 0.6rem 1rem;
      }}
      .stButton>button:hover {{ filter: brightness(0.95); }}
      .stDownloadButton>button {{ background: {ACCENT}; color: {LIGHT}; border-radius: 10px; }}
      h1, h2, h3, h4 {{ color: '{PRIMARY}'; }}
    </style>
    """,
    unsafe_allow_html=True,
)

TEST_DATE = "18-10-2025"
VENUES = [
    "Den Hijzelaar",
    "Pieperz",
    "Cafetaria Marktzicht",
    "Snackbar Jopie",
    "De Bunders",
    "He Sushi en Snackbar",
    "De Boekt",
    "De Smulhoek",
]

# Criteria + weging (0–10 met stap 0,5)
CRITERIA = [
    ("Snelheid (pan → tafel)", 1),
    ("Prijs", 2),
    ("Overall smaak", 3),
    ("Curry", 2),
    ("Mayo", 2),
    ("Uitjes", 2),
    ("Service", 2),
]
WEIGHTS = {name: float(w) for name, w in CRITERIA}
TOTAL_WEIGHT = sum(WEIGHTS.values()) or 1.0

# Kolommen in Google Sheet
SHEET_COLUMNS = [
    "timestamp", "tester", "venue", "price",
    "Snelheid (pan → tafel)", "Prijs", "Overall smaak", "Curry", "Mayo", "Uitjes", "Service",
    "remark"
]

WS_SUBMISSIONS = "submissions"


# ==========
# Secrets → Google Sheets client
# ==========
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

creds_info = None
if "google_service_account" in st.secrets:
    creds_info = dict(st.secrets["google_service_account"])

SHEET_ID_RAW = st.secrets.get("SHEET_ID") or st.secrets.get("vegdel", {}).get("SHEET_ID")
SHEET_ID = _normalize_sheet_id(SHEET_ID_RAW)

client = None
if creds_info and SHEET_ID:
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
        client = gspread.authorize(creds)
    except Exception:
        client = None


def get_sheet():
    if client is None or not SHEET_ID:
        st.error("Geen Google Sheets-verbinding. Controleer Secrets (service account + SHEET_ID).")
        return None
    try:
        return client.open_by_key(SHEET_ID)
    except Exception:
        st.error("Kan Google Sheet niet openen. Deel de sheet als 'Bewerker' met het service account en controleer de ID.")
        return None


def ensure_submissions():
    sh = get_sheet()
    if sh is None:
        return None
    try:
        titles = {ws.title for ws in sh.worksheets()}
        if WS_SUBMISSIONS not in titles:
            sh.add_worksheet(title=WS_SUBMISSIONS, rows=1000, cols=26)
        ws = sh.worksheet(WS_SUBMISSIONS)
        vals = ws.get_all_values()
        if not vals:
            ws.update([SHEET_COLUMNS])  # header
        return ws
    except Exception:
        st.error("Kon werkblad 'submissions' niet initialiseren.")
        return None


def append_submission_row(row: Dict):
    ws = ensure_submissions()
    if ws is None:
        return False, "Geen toegang tot Google Sheet."
    try:
        ordered = [row.get(col, "") for col in SHEET_COLUMNS]
        ws.append_row(ordered, value_input_option="RAW")
        return True, None
    except Exception as e:
        return False, str(e)


def load_df() -> pd.DataFrame:
    ws = ensure_submissions()
    if ws is None:
        return pd.DataFrame(columns=SHEET_COLUMNS)
    try:
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame(columns=SHEET_COLUMNS)


# ==========
# Resultaten & PDF
# ==========
# ---- Resultaten
with results_tab:
    # Laad ruwe data en toon status
    df = load_df()
    st.caption(f"Rijen in 'submissions': {len(df)}")
    with st.expander("Ruwe data (submissions)"):
        st.dataframe(df.head(50), use_container_width=True)

    # Recompute ranking
    rank = compute_results(df)

    # Overzicht + winnaar
    if rank.empty:
        st.info("Nog geen scores ingevoerd of data niet herkend (controleer de header van het tabblad 'submissions').")
        winner_row = None
    else:
        winner_row = rank.iloc[0]
        with st.container(border=True):
            st.markdown("**Overall winnaar (huidig)**")
            st.markdown(f"### 🥇 {winner_row['Cafetaria']}")
            st.write(
                f"Gemiddelde score: **{winner_row['Gem. score (0-10)']:.2f}** "
                f"• op basis van {int(winner_row['# Testers'])} tester(s)."
            )

        st.markdown("#### Uitslag & Ranking")
        st.dataframe(rank, use_container_width=True)

    # PDF-actie altijd zichtbaar (disable als geen data)
    make_pdf = st.button("⛔ Einde Vegdel / Maak PDF", disabled=rank.empty)
    st.caption("Als deze knop uitstaat: voer eerst scores in. "
               "Blijft de ranking leeg? Check de header van het 'submissions'-werkblad (zie hieronder).")

    if make_pdf and not rank.empty:
        pdf_bytes = build_pdf(rank)
        st.session_state["final_pdf"] = pdf_bytes
        st.success("Eindrapport gegenereerd. Download hieronder.")

    if st.session_state.get("final_pdf"):
        st.download_button(
            "📄 Download eindrapport (PDF)",
            data=st.session_state["final_pdf"],
            file_name="vegdel_eindrapport.pdf",
            mime="application/pdf",
        )

    st.markdown("---")
    st.markdown("**Header die het 'submissions'-werkblad moet hebben (exact):**")
    st.code(
        "timestamp, tester, venue, price, "
        "Snelheid (pan → tafel), Prijs, Overall smaak, Curry, Mayo, Uitjes, Service, remark",
        language="text",
    )
    st.caption("Als de kopteksten afwijken, hernoem de eerste rij in het werkblad naar bovenstaande namen.")



# ==========
# UI
# ==========
st.title("🥇 Vegdel – Beste friettent (Frikandel Speciaal)")
st.caption(f"Testdatum: {TEST_DATE} • Regels: zelfde portie • Iedereen kan direct scoren")

st.markdown("### 🧪 Scoren")
colA, colB = st.columns(2)
with colA:
    tester = st.text_input("Jouw naam", placeholder="Voor- en achternaam")
with colB:
    venue = st.selectbox("Kies cafetaria", VENUES)

price = st.number_input("Prijs frikandel speciaal (€)", min_value=0.0, step=0.05, value=0.0)

st.markdown("#### Scores (0–10, stap 0,5) – verplicht")
inputs: Dict[str, float] = {}
for crit, _w in CRITERIA:
    inputs[crit] = st.number_input(crit, min_value=0.0, max_value=10.0, step=0.5, value=5.0)

remark = st.text_area("Opmerkingen (optioneel)")

save_clicked = st.button("Opslaan")
if save_clicked:
    if not tester.strip():
        st.error("Vul je naam in.")
    elif not venue:
        st.error("Kies een cafetaria.")
    else:
        row = {
            "timestamp": pd.Timestamp.utcnow().isoformat(),
            "tester": tester.strip(),
            "venue": venue,
            "price": float(price),
            "remark": remark.strip(),
        }
        for crit, _w in CRITERIA:
            row[crit] = float(inputs[crit])
        ok, err = append_submission_row(row)
        if ok:
            st.success(f"Bedankt {tester}! Je score voor {venue} is opgeslagen.")
        else:
            st.error(f"Opslaan is niet gelukt: {err}")

st.markdown("---")
st.markdown("### 🏆 Resultaten")
df = load_df()
rank = compute_results(df)
if rank.empty:
    st.info("Er zijn nog geen scores ingevoerd.")
else:
    winner = rank.iloc[0]["Cafetaria"]
    st.markdown(f"**Overall winnaar (huidig):** 🥇 {winner}")
    st.dataframe(rank, use_container_width=True)

    make_pdf = st.button("⛔ Einde Vegdel / Maak PDF")
    if make_pdf:
        pdf_bytes = build_pdf(rank)
        st.session_state["final_pdf"] = pdf_bytes
        st.success("Eindrapport gegenereerd. Download hieronder.")

    if st.session_state.get("final_pdf"):
        st.download_button(
            "📄 Download eindrapport (PDF)",
            data=st.session_state["final_pdf"],
            file_name="vegdel_eindrapport.pdf",
            mime="application/pdf",
        )
