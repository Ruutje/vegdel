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
def compute_results(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Rank", "Cafetaria", "Gem. score (0-10)", "# Testers", "Gem. prijs (€)"])

    # Zet kolommen goed
    for c in WEIGHTS.keys():
        if c not in df.columns:
            df[c] = np.nan

    # prijs naar float
    if "price" in df.columns:
        df["price"] = pd.to_numeric(df["price"], errors="coerce")
    else:
        df["price"] = np.nan

    # Weighted average per tester per venue
    scores_cols = list(WEIGHTS.keys())
    for col in scores_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    w = WEIGHTS
    TW = TOTAL_WEIGHT

    def row_weighted_mean(row):
        s = 0.0
        tot = 0.0
        for k, ww in w.items():
            val = row.get(k, np.nan)
            if pd.notna(val):
                s += float(val) * ww
                tot += ww
        return s / (tot or TW)

    df["_tester_avg"] = df[scores_cols].apply(row_weighted_mean, axis=1)

    # Aggregate per venue
    res = df.groupby("venue", as_index=False).agg(
        **{
            "Gem. score (0-10)": ("_tester_avg", "mean"),
            "# Testers": ("tester", "nunique"),
            "Gem. prijs (€)": ("price", "mean"),
        }
    )
    res["Gem. score (0-10)"] = res["Gem. score (0-10)"].round(2)
    res["Gem. prijs (€)"] = res["Gem. prijs (€)"].round(2)
    res = res.sort_values(["Gem. score (0-10)", "# Testers"], ascending=[False, False]).reset_index(drop=True)
    res.insert(0, "Rank", range(1, len(res) + 1))
    res.rename(columns={"venue": "Cafetaria"}, inplace=True)
    return res


def plot_ranking(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    plt.figure(figsize=(8, 4.5))
    plt.bar(df["Cafetaria"], df["Gem. score (0-10)"])
    plt.xticks(rotation=30, ha="right")
    plt.ylabel("Gemiddelde score (0-10)")
    plt.title("Vegdel – Ranking Frikandel Speciaal")
    plt.tight_layout()
    plt.savefig(buf, format="png", dpi=200)
    plt.close()
    buf.seek(0)
    return buf


def build_pdf(df_rank: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    body = []

    body += [Paragraph("<b>Vegdel – Frikandel Speciaal</b>", styles["Title"]), Spacer(1, 0.25*cm)]
    body += [Paragraph(f"Testdatum: {TEST_DATE}", styles["Normal"]), Spacer(1, 0.3*cm)]

    if not df_rank.empty:
        winner = df_rank.iloc[0]["Cafetaria"]
        score = df_rank.iloc[0]["Gem. score (0-10)"]
        body += [Paragraph(f"<b>Overall winnaar:</b> {winner} (gem. {score:.2f})", styles["Heading2"]), Spacer(1, 0.2*cm)]

        tbl = df_rank[["Rank", "Cafetaria", "Gem. score (0-10)", "# Testers", "Gem. prijs (€)"]]
        data = [list(tbl.columns)] + tbl.values.tolist()
        t = Table(data, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor(PRIMARY)),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.25, colors.gray),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor("#f7efe9")]),
        ]))
        body += [Paragraph("<b>Uitslag & Ranking</b>", styles["Heading2"]), t, Spacer(1, 0.4*cm)]

        chart_buf = plot_ranking(df_rank)
        body += [Paragraph("<b>Grafiek – Gemiddelde scores</b>", styles["Heading2"]), RLImage(chart_buf, width=16*cm, height=9*cm)]

    doc.build(body)
    buffer.seek(0)
    return buffer.read()


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
