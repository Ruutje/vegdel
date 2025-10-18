from __future__ import annotations
import uuid
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Any

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from streamlit_autorefresh import st_autorefresh

# Google Sheets
import gspread
from google.oauth2.service_account import Credentials

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet


# =========================
# Theme & Page config
# =========================
PRIMARY = "#b94216"     # bruin (frikandel)
ACCENT  = "#891418"     # curry
LIGHT   = "#ffffff"     # wit
ADMIN_PIN = "1000"

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

# =========================
# Models & Defaults
# =========================
def uid() -> str:
    return uuid.uuid4().hex[:8]

@dataclass
class Criterion:
    key: str
    name: str
    weight: float

@dataclass
class Venue:
    key: str
    name: str
    active: bool = True
    price: float = 0.0

DEFAULT_VENUE_NAMES = [
    "Den Hijzelaar",
    "Pieperz",
    "Cafetaria Marktzicht",
    "Snackbar Jopie",
    "De Bunders",
    "He Sushi en Snackbar",
    "De Boekt",
    "De Smulhoek",
]
DEFAULT_VENUES: List[Venue] = [Venue(uid(), n, True, 0.0) for n in DEFAULT_VENUE_NAMES]

DEFAULT_CRITERIA: List[Criterion] = [
    Criterion(uid(), "Snelheid (pan → tafel)", 1),
    Criterion(uid(), "Prijs", 2),
    Criterion(uid(), "Overall smaak", 3),
    Criterion(uid(), "Curry", 2),
    Criterion(uid(), "Mayo", 2),
    Criterion(uid(), "Uitjes", 2),
    Criterion(uid(), "Service", 2),
]

# Sheet tab names
WS_FLAGS      = "flags"
WS_VENUES     = "venues"
WS_CRITERIA   = "criteria"
WS_SUBMISSIONS= "submissions"


# =========================
# Secrets / Google Sheets client (defensief)
# =========================
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
    except Exception as e:
        client = None
else:
    # Laat UI gewoon laden: Diagnose-tab helpt verder
    pass

def get_sheet():
    """Open het Spreadsheet; return None met UI-melding als het niet lukt."""
    if client is None or not SHEET_ID:
        st.error("Geen Google Sheets-verbinding. Controleer Secrets (service account + SHEET_ID).")
        st.caption("Ga naar tab ‘🧰 Diagnose’ om dit te controleren en te testen.")
        return None
    try:
        return client.open_by_key(SHEET_ID)
    except Exception:
        st.error("Kan Google Sheet niet openen. Controleer of het service account Editor-rechten heeft, de ID klopt en Sheets/Drive API aanstaan.")
        st.caption("Tip: Deel de sheet met het service account e-mailadres (Editor). Zie Diagnose-tab.")
        return None

def ensure_worksheets():
    """Zorg dat alle tabbladen bestaan + basisheaders. Return None als sheet niet open kan."""
    sh = get_sheet()
    if sh is None:
        return None
    try:
        existing = {ws.title for ws in sh.worksheets()}
        wanted = {WS_FLAGS, WS_VENUES, WS_CRITERIA, WS_SUBMISSIONS}
        for name in wanted - existing:
            sh.add_worksheet(title=name, rows=1000, cols=26)
        # flags
        ws = sh.worksheet(WS_FLAGS)
        if not ws.get_all_values():
            ws.update([["key","value"], ["started","False"], ["finalized","False"]])
        # venues
        ws = sh.worksheet(WS_VENUES)
        if not ws.get_all_values():
            ws.update([["key","name","active","price"]] + [[v.key, v.name, True, 0.0] for v in DEFAULT_VENUES])
        # criteria
        ws = sh.worksheet(WS_CRITERIA)
        if not ws.get_all_values():
            ws.update([["key","name","weight"]] + [[c.key, c.name, c.weight] for c in DEFAULT_CRITERIA])
        # submissions
        ws = sh.worksheet(WS_SUBMISSIONS)
        if not ws.get_all_values():
            ws.update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
        return sh
    except Exception:
        st.error("Kon Google Sheet tabbladen niet initialiseren.")
        return None

def read_flags() -> Dict[str, Any]:
    sh = ensure_worksheets()
    if sh is None:
        return {"started": False, "finalized": False}
    try:
        ws = sh.worksheet(WS_FLAGS)
        data = ws.get_all_records()
        d = {r["key"]: r["value"] for r in data}
        return {"started": str(d.get("started","False")).lower()=="true",
                "finalized": str(d.get("finalized","False")).lower()=="true"}
    except Exception:
        return {"started": False, "finalized": False}

def write_flags(started: bool | None = None, finalized: bool | None = None):
    sh = get_sheet()
    if sh is None:
        st.error("Kon flags niet opslaan (geen toegang tot sheet).")
        return
    try:
        ws = sh.worksheet(WS_FLAGS)
        recs = ws.get_all_records()
        m = {r["key"]: r["value"] for r in recs}
        if started is not None:  m["started"] = str(bool(started))
        if finalized is not None: m["finalized"] = str(bool(finalized))
        ws.update([["key","value"]] + [[k, v] for k, v in m.items()])
    except Exception:
        st.error("Opslaan van flags mislukte.")

def read_venues() -> List[Venue]:
    sh = ensure_worksheets()
    if sh is None:
        st.warning("Geen verbinding met Google Sheets — toon standaard venues.")
        return [Venue(v.key, v.name, v.active, v.price) for v in DEFAULT_VENUES]
    try:
        ws = sh.worksheet(WS_VENUES)
        rows = ws.get_all_records()
        return [Venue(str(r["key"]), str(r["name"]).strip(), bool(r.get("active", True)), float(r.get("price",0.0))) for r in rows]
    except Exception:
        st.warning("Kon venues niet lezen — gebruik defaults.")
        return [Venue(v.key, v.name, v.active, v.price) for v in DEFAULT_VENUES]

def write_venues(vs: List[Venue]):
    sh = get_sheet()
    if sh is None:
        st.error("Kon venues niet opslaan (geen toegang tot sheet).")
        return
    try:
        ws = sh.worksheet(WS_VENUES)
        ws.update([["key","name","active","price"]] + [[v.key, v.name, v.active, v.price] for v in vs])
    except Exception:
        st.error("Opslaan van venues mislukte.")

def read_criteria() -> List[Criterion]:
    sh = ensure_worksheets()
    if sh is None:
        st.warning("Geen verbinding met Google Sheets — toon standaard criteria.")
        return [Criterion(c.key, c.name, c.weight) for c in DEFAULT_CRITERIA]
    try:
        ws = sh.worksheet(WS_CRITERIA)
        rows = ws.get_all_records()
        return [Criterion(str(r["key"]), str(r["name"]).strip(), float(r.get("weight",1))) for r in rows]
    except Exception:
        st.warning("Kon criteria niet lezen — gebruik defaults.")
        return [Criterion(c.key, c.name, c.weight) for c in DEFAULT_CRITERIA]

def write_criteria(cs: List[Criterion]):
    sh = get_sheet()
    if sh is None:
        st.error("Kon criteria niet opslaan (geen toegang tot sheet).")
        return
    try:
        ws = sh.worksheet(WS_CRITERIA)
        ws.update([["key","name","weight"]] + [[c.key, c.name, c.weight] for c in cs])
    except Exception:
        st.error("Opslaan van criteria mislukte.")

def append_submission(tester_name: str, v: Venue, scores_dict: Dict[str, float], remark: str):
    """Append submission; return (ok, error_msg|None)."""
    sh = ensure_worksheets()
    if sh is None:
        return False, "Geen toegang tot Google Sheet."
    try:
        ws = sh.worksheet(WS_SUBMISSIONS)
        # herlees criteria voor sleutel/naam-consistentie
        cs = read_criteria()
        rows = []
        for c in cs:
            val = float(scores_dict.get(c.key, 0.0))
            rows.append([tester_name, v.key, v.name, float(v.price), c.key, c.name, val, remark])
        ws.append_rows(rows, value_input_option="RAW")
        return True, None
    except Exception as e:
        return False, str(e)

def load_submissions_df() -> pd.DataFrame:
    sh = ensure_worksheets()
    if sh is None:
        return pd.DataFrame(columns=["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"])
    try:
        ws = sh.worksheet(WS_SUBMISSIONS)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame(columns=["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"])


# =========================
# State & shared flags
# =========================
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False
if "started" not in st.session_state:
    st.session_state.started = False
if "finalized" not in st.session_state:
    st.session_state.finalized = False
if "final_pdf" not in st.session_state:
    st.session_state.final_pdf: bytes | None = None

flags = read_flags()
st.session_state.started = flags.get("started", False)
st.session_state.finalized = flags.get("finalized", False)

# Auto-refresh zolang niet beëindigd (iedere 5s)
if not st.session_state.finalized:
    st_autorefresh(interval=5000, key="polling")

# Laad actuele venues/criteria
venues: List[Venue] = read_venues()
criteria: List[Criterion] = read_criteria()


# =========================
# Helpers: resultaten & PDF
# =========================
def total_weight() -> float:
    w = sum(max(0.0, float(c.weight)) for c in criteria)
    return w if w > 0 else 1.0

def per_tester_top1() -> pd.DataFrame:
    sub = load_submissions_df()
    if sub.empty:
        return pd.DataFrame(columns=["Tester","#1 Cafetaria","Score (0-10)"])
    w = {c.key: float(c.weight) for c in criteria}
    TW = sum(w.values()) or 1.0
    sub["w"] = sub["criterion_key"].map(w).fillna(0.0)
    agg = (
        sub.groupby(["tester","venue_key","venue_name"], as_index=False)
           .apply(lambda g: (g["score"]*g["w"]).sum()/TW)
           .reset_index(drop=True)
           .rename(columns={0:"tester_avg"})
    )
    idx = agg.groupby("tester")["tester_avg"].idxmax()
    best = agg.loc[idx, ["tester","venue_name","tester_avg"]].copy()
    best.rename(columns={"tester":"Tester","venue_name":"#1 Cafetaria","tester_avg":"Score (0-10)"}, inplace=True)
    best["Score (0-10)"] = best["Score (0-10)"].round(2)
    return best

def compute_results() -> pd.DataFrame:
    sub = load_submissions_df()
    if not sub.empty:
        w = {c.key: float(c.weight) for c in criteria}
        sub["w"] = sub["criterion_key"].map(w).fillna(0.0)
        TW = sum(w.values()) or 1.0
        agg = (
            sub.groupby(["tester","venue_key","venue_name","price"], as_index=False)
               .apply(lambda g: (g["score"]*g["w"]).sum()/TW)
               .reset_index(drop=True)
               .rename(columns={0:"tester_avg"})
        )
        res = agg.groupby(["venue_key","venue_name","price"], as_index=False).agg(
            **{"Gem. score (0-10)":("tester_avg","mean"), "# Testers":("tester","nunique")}
        )
        res["Gem. score (0-10)"] = res["Gem. score (0-10)"].round(2)
        res = res.sort_values(["Gem. score (0-10)","# Testers"], ascending=[False, False]).reset_index(drop=True)
        res.insert(0,"Rank", range(1, len(res)+1))
        res.rename(columns={"venue_name":"Cafetaria","price":"Prijs"}, inplace=True)
        return res
    return pd.DataFrame(columns=["Rank","Cafetaria","Prijs","Gem. score (0-10)","# Testers"])

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

def build_pdf(df: pd.DataFrame, top1_df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    story = []

    story += [Paragraph("<b>Vegdel – Frikandel Speciaal</b>", styles["Title"]), Spacer(1, 0.3*cm)]

    if not df.empty:
        tbl_df = df[["Rank", "Cafetaria", "Prijs", "Gem. score (0-10)", "# Testers"]]
        data = [list(tbl_df.columns)] + tbl_df.values.tolist()
        table = Table(data, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor(PRIMARY)),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.25, colors.gray),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor("#f7efe9")]),
        ]))
        story += [Paragraph("<b>Uitslag & Ranking</b>", styles["Heading2"]), table, Spacer(1, 0.4*cm)]

    if top1_df is not None and not top1_df.empty:
        tdata = [list(top1_df.columns)] + top1_df.values.tolist()
        ttable = Table(tdata, hAlign="LEFT")
        ttable.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor(ACCENT)),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.25, colors.gray),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor("#f7efe9")]),
        ]))
        story += [Paragraph("<b>Persoonlijke #1 per tester</b>", styles["Heading2"]), ttable, Spacer(1, 0.4*cm)]

    if not df.empty:
        chart_buf = plot_ranking(df)
        story += [Paragraph("<b>Grafiek – Gemiddelde scores</b>", styles["Heading2"]), RLImage(chart_buf, width=16*cm, height=9*cm)]

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


# =========================
# Header & Sidebar (admin)
# =========================
st.title("🥇 Vegdel – Beste friettent (Frikandel Speciaal)")
st.caption("Testdatum: 18-10-2025 • Regels: zelfde portie • Testers vullen in op eigen device")

with st.sidebar:
    st.header("🔐 Beheer")
    if not st.session_state.admin_mode:
        pin_try = st.text_input("Voer pincode in", type="password")
        if st.button("Inloggen beheer"):
            if pin_try == ADMIN_PIN:
                st.session_state.admin_mode = True
                st.success("Ingelogd als beheerder.")
            else:
                st.error("Onjuiste pincode.")
    else:
        st.success("Beheer actief")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("▶️ Start Vegdel", disabled=st.session_state.started or st.session_state.finalized):
                write_flags(started=True)
                st.session_state.started = True
                st.toast("Vegdel gestart — invoer geactiveerd voor iedereen.")
        with c2:
            if st.button("🔓 Uitloggen beheer"):
                st.session_state.admin_mode = False
                st.toast("Beheer uitgelogd.")

        st.divider()
        with st.expander("🗑️ Alles wissen (cloud + sessie)"):
            confirm = st.checkbox("Ik bevestig het wissen van alle data")
            if st.button("Alles wissen"):
                if confirm:
                    sh = ensure_worksheets()
                    if sh:
                        try:
                            ws = sh.worksheet(WS_SUBMISSIONS)
                            ws.clear()
                            ws.update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
                        except Exception:
                            pass
                        write_flags(started=False, finalized=False)
                        write_venues(DEFAULT_VENUES)
                        write_criteria(DEFAULT_CRITERIA)
                    st.session_state.finalized = False
                    st.session_state.final_pdf = None
                    st.session_state.started = False
                    st.toast("Alle data gewist (cloud + sessie).")
                else:
                    st.warning("Vink de bevestiging aan om te wissen.")

if st.session_state.finalized:
    st.warning("De test is beëindigd. Wijzigingen zijn vergrendeld.")
elif not st.session_state.started:
    st.info("Nog niet gestart. Beheerder kan starten via de sidebar.")


# =========================
# Tabs
# =========================
setup_tab, score_tab, results_tab, beheer_tab, diag_tab = st.tabs(
    ["⚙️ Instellen", "🧪 Scoren", "🏆 Resultaten", "🗂️ Beheer", "🧰 Diagnose"]
)

# ---- Instellen (alleen weergave; beheren doe je in Beheer-tab)
with setup_tab:
    st.subheader("Cafetaria's & prijzen (weergave – beheer in '🗂️ Beheer')")
    v_df = pd.DataFrame([{"key": v.key, "Actief": v.active, "Naam": v.name, "Prijs (EUR)": float(v.price)} for v in venues])
    st.dataframe(v_df.drop(columns=["key"]), use_container_width=True)

    st.subheader("Criteria & Weging")
    c_df = pd.DataFrame([{"key": c.key, "Criterium": c.name, "Weging": float(c.weight)} for c in criteria])
    st.dataframe(c_df.drop(columns=["key"]), use_container_width=True)

# ---- Scoren
with score_tab:
    if not st.session_state.started and not st.session_state.finalized:
        st.warning("De test is nog niet gestart. Beheerder kan starten via de sidebar.")
    if not any(v.active for v in venues) or not criteria:
        st.warning("Zorg dat er actieve cafetaria's en criteria zijn in '🗂️ Beheer'.")
    else:
        st.markdown("#### Wie ben je?")
        your_name = st.text_input("Vul je naam in", placeholder="Voor- en achternaam", disabled=st.session_state.finalized)
        current_name = your_name.strip() if your_name else ""
        if not current_name:
            st.info("Vul eerst je naam in om te kunnen scoren.")
        else:
            st.markdown("---")
            st.subheader(f"Beoordelen als: {current_name}")
            active_venues = [v for v in venues if v.active]

            for v in active_venues:
                with st.form(key=f"form_{current_name}_{v.key}", clear_on_submit=False):
                    st.markdown(f"### {v.name}")
                    cols = st.columns(2)
                    with cols[0]:
                        st.write("**Prijs (EUR)**")
                        price_val = st.number_input("Prijs", min_value=0.0, step=0.05, value=float(v.price), key=f"price_{v.key}")
                    with cols[1]:
                        st.write("**Opmerkingen (optioneel)**")
                        remark = st.text_area("Opmerkingen", value="", key=f"remark_{current_name}_{v.key}")

                    st.write("**Scores (0–10, stap 0,5):** Alle velden verplicht")
                    score_inputs: Dict[str, float] = {}
                    for c in criteria:
                        score_inputs[c.key] = st.number_input(
                            c.name, min_value=0.0, max_value=10.0, step=0.5, value=5.0,
                            key=f"score_{current_name}_{v.key}_{c.key}"
                        )

                    saved = st.form_submit_button("Opslaan cafetaria")
                    if saved and not st.session_state.finalized:
                        missing = [c.name for c in criteria if c.key not in score_inputs or score_inputs[c.key] is None]
                        if missing:
                            st.error("Niet alle criteria zijn ingevuld.")
                        else:
                            # schrijf prijsupdate naar venues-sheet
                            for i, vv in enumerate(venues):
                                if vv.key == v.key:
                                    venues[i].price = float(price_val)
                                    break
                            write_venues(venues)
                            ok, err = append_submission(current_name, v, score_inputs, remark)
                            if ok:
                                st.toast(f"Bedankt {current_name}! {v.name} is opgeslagen.")
                            else:
                                st.error(f"Opslaan in de cloud is niet gelukt: {err}")

# ---- Resultaten
with results_tab:
    df = compute_results()
    if df.empty:
        st.info("Nog geen scores ingevoerd.")
    else:
        winner_row = df.iloc[0]
        with st.container(border=True):
            st.markdown("**Overall winnaar**")
            st.markdown(f"### 🥇 {winner_row['Cafetaria']}")
            st.write(f"Gemiddelde score: **{winner_row['Gem. score (0-10)']:.2f}** • op basis van {int(winner_row['# Testers'])} tester(s).")

        st.markdown("#### Uitslag & Ranking")
        st.dataframe(df, use_container_width=True)

        top1 = per_tester_top1()
        if not top1.empty:
            st.markdown("#### Persoonlijke #1 per tester")
            st.dataframe(top1, use_container_width=True)

        c1, c2 = st.columns([1,2])
        with c1:
            end_clicked = st.button("⛔ Einde Vegdel", disabled=st.session_state.finalized or (not st.session_state.started) or (not st.session_state.admin_mode))
        with c2:
            st.caption("Bij beëindigen wordt alles vergrendeld en een PDF-rapport met grafieken gegenereerd. Actieve venues tellen mee; gesloten/inactieve niet.")

        if end_clicked:
            st.session_state.finalized = True
            write_flags(finalized=True)
            pdf_bytes = build_pdf(df, top1)
            st.session_state.final_pdf = pdf_bytes
            st.success("Eindrapport is gegenereerd en de invoer is vergrendeld.")

        if st.session_state.final_pdf:
            st.download_button("📄 Download eindrapport (PDF)", data=st.session_state.final_pdf, file_name="vegdel_eindrapport.pdf", mime="application/pdf")

# ---- Beheer (alleen actief na PIN)
with beheer_tab:
    st.subheader("🗂️ Beheer")
    if not st.session_state.admin_mode:
        st.warning("Alleen toegankelijk voor beheerder. Log in met PIN 1000 in de sidebar.")
    else:
        st.success("Beheer actief")

        # Venues
        st.markdown("### 🏪 Cafetaria's beheren")
        v_list = read_venues()
        v_df = pd.DataFrame([{"key": v.key, "Naam": v.name, "Actief": bool(v.active), "Prijs (EUR)": float(v.price)} for v in v_list])
        v_edit = st.data_editor(
            v_df, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "key": st.column_config.TextColumn("key", disabled=True),
                "Naam": st.column_config.TextColumn("Naam"),
                "Actief": st.column_config.CheckboxColumn("Actief"),
                "Prijs (EUR)": st.column_config.NumberColumn("Prijs (EUR)", min_value=0.0, step=0.05),
            },
        )
        cols_v = st.columns([1,2])
        with cols_v[0]:
            if st.button("💾 Venues opslaan"):
                new_vs = []
                for _, r in v_edit.iterrows():
                    name = str(r.get("Naam","")).strip()
                    if not name:
                        continue
                    new_vs.append(Venue(str(r["key"]), name, bool(r.get("Actief", True)), float(r.get("Prijs (EUR)") or 0.0)))
                write_venues(new_vs)
                st.toast("Venues opgeslagen in Google Sheets.")
        with cols_v[1]:
            st.caption("Tip: zet 'Actief' uit voor gesloten zaken. Nieuwe rijen kun je onderaan toevoegen.")

        st.divider()

        # Criteria
        st.markdown("### 📊 Criteria beheren")
        c_list = read_criteria()
        c_df = pd.DataFrame([{"key": c.key, "Criterium": c.name, "Weging": float(c.weight)} for c in c_list])
        c_edit = st.data_editor(
            c_df, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "key": st.column_config.TextColumn("key", disabled=True),
                "Criterium": st.column_config.TextColumn("Criterium"),
                "Weging": st.column_config.NumberColumn("Weging", min_value=0.0, step=1.0),
            },
        )
        cols_c = st.columns([1,2])
        with cols_c[0]:
            if st.button("💾 Criteria opslaan"):
                new_cs = []
                for _, r in c_edit.iterrows():
                    nm = str(r.get("Criterium","")).strip()
                    if not nm:
                        continue
                    new_cs.append(Criterion(str(r["key"]), nm, float(r.get("Weging") or 0.0)))
                write_criteria(new_cs)
                st.toast("Criteria opgeslagen in Google Sheets.")
        with cols_c[1]:
            st.caption("Let op: aanpassen van criteria/weging werkt door in de berekening van alle resultaten.")

        st.divider()

        # Submissions opschonen
        st.markdown("### ⚠️ Data opschonen")
        st.caption("Leeg alleen de testdata (submissions). Venues, criteria en flags blijven staan.")
        confirm_clear = st.checkbox("Ik bevestig dat ik alle testdata wil wissen", key="confirm_clear_submissions")
        if st.button("🧽 Wis alle testdata (submissions)"):
            if not confirm_clear:
                st.warning("Zet eerst de bevestigingsvink aan.")
            else:
                sh = ensure_worksheets()
                if sh:
                    try:
                        ws = sh.worksheet(WS_SUBMISSIONS)
                        ws.clear()
                        ws.update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
                        st.toast("Alle testdata gewist.")
                    except Exception:
                        st.error("Kon submissions niet wissen.")
                else:
                    st.error("Kon Google Sheet niet openen. Controleer de Diagnose-tab.")

# ---- Diagnose
with diag_tab:
    st.subheader("🧰 Diagnose & Verbinding")
    st.write("Service account geladen:", "✅" if creds_info else "❌ (niet gevonden)")
    if creds_info:
        st.write("Service account email:", creds_info.get("client_email"))
    st.write("SHEETS_ID (raw):", SHEET_ID_RAW if SHEET_ID_RAW else "(niet gezet)")
    st.write("SHEETS_ID (parsed):", SHEET_ID if SHEET_ID else "(kon niet parsen)")

    def diag_ping() -> str:
        sh = ensure_worksheets()
        if sh is None:
            return "❌ Geen toegang tot Google Sheet."
        try:
            ws = sh.worksheet(WS_SUBMISSIONS)
            ws.append_row(
                ["__DIAG__", "vkey", "vname", 0.0, "ckey", "cname", 5.0, "diagnose"],
                value_input_option="RAW"
            )
            return f"✅ Verbinding OK met {SHEET_ID}"
        except Exception as e:
            return f"❌ Fout tijdens verbindingstest: {e}"

    def diag_cleanup() -> str:
        sh = ensure_worksheets()
        if sh is None:
            return "❌ Geen toegang tot Google Sheet."
        try:
            ws = sh.worksheet(WS_SUBMISSIONS)
            data = ws.get_all_values()
            if not data:
                return "ℹ️ Geen data om op te schonen."
            header = data[0]
            rows = data[1:]
            try:
                idx = header.index("tester")
            except ValueError:
                return "❌ Kolom 'tester' niet gevonden."
            keep = [header] + [r for r in rows if (len(r) > idx and r[idx] != "__DIAG__")]
            ws.clear()
            ws.update(keep)
            return "🧽 Diagnose-rijen verwijderd."
        except Exception as e:
            return f"❌ Fout tijdens opschonen: {e}"

    colD1, colD2 = st.columns(2)
    with colD1:
        if st.button("🔌 Test verbinding (ping)"):
            st.info("Diagnose bezig…")
            msg = diag_ping()
            if msg.startswith("✅"):
                st.success(msg)
            elif msg.startswith("⚠️"):
                st.warning(msg)
            else:
                st.error(msg)
    with colD2:
        if st.button("🧽 Verwijder diagnose-rijen"):
            msg = diag_cleanup()
            if msg.startswith("🧽"):
                st.success(msg)
            else:
                st.info(msg)
