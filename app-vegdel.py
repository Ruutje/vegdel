from __future__ import annotations
import json
import math
import uuid
from dataclasses import dataclass, asdict
from io import BytesIO
from typing import Dict, List, Any

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from streamlit_autorefresh import st_autorefresh
import gspread
from google.oauth2.service_account import Credentials

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet

# ==========================================
# App Config & Theming
# ==========================================
PRIMARY = "#b94216"   # bruin (frikandel)
ACCENT = "#891418"     # curry
LIGHT  = "#ffffff"     # wit

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

# ==========================================
# Secrets & Google Sheets Client
# ==========================================
# Supports either top-level SHEET_ID, or [vegdel] block with SHEET_ID
if "google_service_account" not in st.secrets:
    st.error("Service account ontbreekt in secrets. Voeg [google_service_account] toe in Settings → Secrets.")
    st.stop()

creds_info = dict(st.secrets["google_service_account"])  # dict-like
SHEET_ID_RAW = (
    st.secrets.get("SHEET_ID")
    or st.secrets.get("vegdel", {}).get("SHEET_ID")
)


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

if not SHEET_ID:
    st.error("SHEET_ID ontbreekt. Zet 'SHEET_ID' (of [vegdel].SHEET_ID) in Secrets. Zie Diagnose-tab voor hulp.")
    st.stop()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
client = gspread.authorize(creds)

# ==========================================
# Data Models & Defaults
# ==========================================

ADMIN_PIN = "1000"

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

@dataclass
class Tester:
    key: str
    name: str

DEFAULT_CRITERIA: List[Criterion] = [
    Criterion(uid(), "Snelheid (pan → tafel)", 1),
    Criterion(uid(), "Prijs", 2),
    Criterion(uid(), "Overall smaak", 3),
    Criterion(uid(), "Curry", 2),
    Criterion(uid(), "Mayo", 2),
    Criterion(uid(), "Uitjes", 2),
    Criterion(uid(), "Service", 2),
]

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

# Worksheet names
WS_FLAGS = "flags"
WS_VENUES = "venues"
WS_CRITERIA = "criteria"
WS_SUBMISSIONS = "submissions"

# ==========================================
# Google Sheets Helpers
# ==========================================

def get_sheet():
    return client.open_by_key(SHEET_ID)


def ensure_worksheets():
    sh = get_sheet()
    existing = {ws.title for ws in sh.worksheets()}
    wanted = {WS_FLAGS, WS_VENUES, WS_CRITERIA, WS_SUBMISSIONS}
    for name in wanted - existing:
        sh.add_worksheet(title=name, rows=1000, cols=26)
    # headers & defaults
    ws = sh.worksheet(WS_FLAGS)
    if not ws.get_all_values():
        ws.update([["key","value"], ["started","False"], ["finalized","False"]])
    ws = sh.worksheet(WS_VENUES)
    if not ws.get_all_values():
        ws.update([["key","name","active","price"]] + [[v.key, v.name, True, 0.0] for v in DEFAULT_VENUES])
    ws = sh.worksheet(WS_CRITERIA)
    if not ws.get_all_values():
        ws.update([["key","name","weight"]] + [[c.key, c.name, c.weight] for c in DEFAULT_CRITERIA])
    ws = sh.worksheet(WS_SUBMISSIONS)
    if not ws.get_all_values():
        ws.update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
    return sh


def read_flags() -> Dict[str, Any]:
    sh = ensure_worksheets()
    ws = sh.worksheet(WS_FLAGS)
    data = ws.get_all_records()
    d = {r["key"]: r["value"] for r in data}
    return {
        "started": str(d.get("started","False")).lower()=="true",
        "finalized": str(d.get("finalized","False")).lower()=="true",
    }


def write_flags(started: bool | None = None, finalized: bool | None = None):
    sh = get_sheet()
    ws = sh.worksheet(WS_FLAGS)
    recs = ws.get_all_records()
    m = {r["key"]: r["value"] for r in recs}
    if started is not None: m["started"] = str(bool(started))
    if finalized is not None: m["finalized"] = str(bool(finalized))
    ws.update([["key","value"]] + [[k, v] for k, v in m.items()])


def read_venues() -> List[Venue]:
    sh = ensure_worksheets()
    ws = sh.worksheet(WS_VENUES)
    rows = ws.get_all_records()
    return [Venue(str(r["key"]), str(r["name"]).strip(), bool(r.get("active", True)), float(r.get("price",0.0))) for r in rows]


def write_venues(vs: List[Venue]):
    sh = get_sheet()
    ws = sh.worksheet(WS_VENUES)
    ws.update([["key","name","active","price"]] + [[v.key, v.name, v.active, v.price] for v in vs])


def read_criteria() -> List[Criterion]:
    sh = ensure_worksheets()
    ws = sh.worksheet(WS_CRITERIA)
    rows = ws.get_all_records()
    return [Criterion(str(r["key"]), str(r["name"]).strip(), float(r.get("weight",1))) for r in rows]


def write_criteria(cs: List[Criterion]):
    sh = get_sheet()
    ws = sh.worksheet(WS_CRITERIA)
    ws.update([["key","name","weight"]] + [[c.key, c.name, c.weight] for c in cs])


def append_submission(tester_name: str, v: Venue, scores_dict: Dict[str, float], remark: str):
    """Append submission rows. Returns (ok: bool, error: str|None)."""
    try:
        sh = ensure_worksheets()
        ws = sh.worksheet(WS_SUBMISSIONS)
        rows = []
        for c in read_criteria():
            val = float(scores_dict.get(c.key, 0.0))
            rows.append([tester_name, v.key, v.name, float(v.price), c.key, c.name, val, remark])
        ws.append_rows(rows, value_input_option="RAW")
        return True, None
    except Exception as e:
        return False, str(e)


def load_submissions_df() -> pd.DataFrame:
    sh = ensure_worksheets()
    ws = sh.worksheet(WS_SUBMISSIONS)
    data = ws.get_all_records()
    return pd.DataFrame(data)

# ==========================================
# Session State
# ==========================================
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False
if "started" not in st.session_state:
    st.session_state.started = False
if "finalized" not in st.session_state:
    st.session_state.finalized = False
if "final_pdf" not in st.session_state:
    st.session_state.final_pdf: bytes | None = None

# Pull shared flags
flags = read_flags()
st.session_state.started = flags.get("started", False)
st.session_state.finalized = flags.get("finalized", False)

# Auto-refresh while not finalized
if not st.session_state.finalized:
    st_autorefresh(interval=5000, key="polling")

# Load shared config
venues: List[Venue] = read_venues()
criteria: List[Criterion] = read_criteria()

# ==========================================
# Result Helpers
# ==========================================

def total_weight() -> float:
    s = sum(max(0.0, float(c.weight)) for c in criteria)
    return s if s > 0 else 1.0


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
           .reset_index(name="tester_avg")
    )
    # best per tester
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
               .reset_index(name="tester_avg")
        )
        res = agg.groupby(["venue_key","venue_name","price"], as_index=False).agg(
            **{"Gem. score (0-10)": ("tester_avg","mean"), "# Testers": ("tester","nunique")}
        )
        res["Gem. score (0-10)"] = res["Gem. score (0-10)"].round(2)
        res = res.sort_values(["Gem. score (0-10)", "# Testers"], ascending=[False, False]).reset_index(drop=True)
        res.insert(0, "Rank", range(1, len(res)+1))
        res.rename(columns={"venue_name":"Cafetaria","price":"Prijs"}, inplace=True)
        return res
    # empty fallback
    return pd.DataFrame(columns=["Rank","Cafetaria","Prijs","Gem. score (0-10)","# Testers"])    


def plot_ranking(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    plt.figure(figsize=(8, 4.5))
    plt.bar(df["Cafetaria"], df["Gem. score (0-10)"])
    plt.xticks(rotation=30, ha='right')
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

    story += [Paragraph("<b>Vegdel – Frikandel Speciaal</b>", styles['Title']), Spacer(1, 0.3*cm)]

    if not df.empty:
        tbl_df = df[["Rank", "Cafetaria", "Prijs", "Gem. score (0-10)", "# Testers"]]
        data = [list(tbl_df.columns)] + tbl_df.values.tolist()
        table = Table(data, hAlign='LEFT')
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(PRIMARY)),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.25, colors.gray),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#f7efe9')]),
        ]))
        story += [Paragraph("<b>Uitslag & Ranking</b>", styles['Heading2']), table, Spacer(1, 0.4*cm)]

    if top1_df is not None and not top1_df.empty:
        tdata = [list(top1_df.columns)] + top1_df.values.tolist()
        ttable = Table(tdata, hAlign='LEFT')
        ttable.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(ACCENT)),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.25, colors.gray),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#f7efe9')]),
        ]))
        story += [Paragraph("<b>Persoonlijke #1 per tester</b>", styles['Heading2']), ttable, Spacer(1, 0.4*cm)]

    if not df.empty:
        chart_buf = plot_ranking(df)
        story += [Paragraph("<b>Grafiek – Gemiddelde scores</b>", styles['Heading2']), RLImage(chart_buf, width=16*cm, height=9*cm)]

    doc.build(story)
    buffer.seek(0)
    return buffer.read()

# ==========================================
# Header & Sidebar (Admin)
# ==========================================
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
        colA, colB = st.columns(2)
        with colA:
            if st.button("▶️ Start Vegdel", disabled=st.session_state.started or st.session_state.finalized):
                write_flags(started=True)
                st.session_state.started = True
                st.toast("Vegdel gestart — invoer geactiveerd voor iedereen.")
        with colB:
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
                        sh.worksheet(WS_SUBMISSIONS).clear()
                        sh.worksheet(WS_SUBMISSIONS).update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
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

# ==========================================
# Tabs
# ==========================================
setup_tab, score_tab, results_tab, beheer_tab, diag_tab = st.tabs(["⚙️ Instellen", "🧪 Scoren", "🏆 Resultaten", "🗂️ Beheer", "🧰 Diagnose"])

# ------------------
# Tab: Instellen (basisbeheer zonder PIN, alleen weergave)
# ------------------
with setup_tab:
    st.subheader("Cafetaria's & prijzen (weergave – beheer in '🗂️ Beheer')")
    v_df = pd.DataFrame([{ "key": v.key, "Actief": v.active, "Naam": v.name, "Prijs (EUR)": float(v.price)} for v in venues])
    st.dataframe(v_df.drop(columns=["key"]), use_container_width=True)
    st.subheader("Criteria & Weging")
    c_df = pd.DataFrame([{ "key": c.key, "Criterium": c.name, "Weging": float(c.weight)} for c in criteria])
    st.dataframe(c_df.drop(columns=["key"]), use_container_width=True)

# ------------------
# Tab: Scoren
# ------------------
with score_tab:
    if not st.session_state.started and not st.session_state.finalized:
        st.warning("De test is nog niet gestart. Beheerder kan starten via de sidebar.")
    if not any(v.active for v in venues) or not criteria:
        st.warning("Zorg dat er actieve cafetaria's en criteria zijn in '🗂️ Beheer'.")
    else:
        st.markdown("#### Wie ben je?")
        your_name = st.text_input("Vul je naam in", placeholder="Voor- en achternaam", disabled=st.session_state.finalized)
        current_tester: Tester | None = None
        if your_name:
            your_name_clean = your_name.strip()
            if your_name_clean:
                current_tester = Tester(uid(), your_name_clean)
        else:
            st.info("Vul eerst je naam in om te kunnen scoren.")

        if current_tester is not None:
            st.markdown("---")
            st.subheader(f"Beoordelen als: {current_tester.name}")
            active_venues = [v for v in venues if v.active]

            for v in active_venues:
                with st.form(key=f"form_{current_tester.key}_{v.key}", clear_on_submit=False):
                    st.markdown(f"### {v.name}")
                    cols = st.columns(2)
                    with cols[0]:
                        st.write("**Prijs (EUR)**")
                        price_val = st.number_input("Prijs", min_value=0.0, step=0.05, value=float(v.price), key=f"price_{v.key}")
                    with cols[1]:
                        st.write("**Opmerkingen (optioneel)**")
                        remark = st.text_area("Opmerkingen", value="", key=f"remark_{current_tester.key}_{v.key}")

                    st.write("**Scores (0–10, stap 0,5):** Alle velden verplicht")
                    score_inputs: Dict[str, float] = {}
                    for c in criteria:
                        score_inputs[c.key] = st.number_input(c.name, min_value=0.0, max_value=10.0, step=0.5,
                                                              value=5.0, key=f"score_{current_tester.key}_{v.key}_{c.key}")

                    saved = st.form_submit_button("Opslaan cafetaria")
                    if saved and not st.session_state.finalized:
                        # Validate required: every criterion must be provided
                        missing = [c.name for c in criteria if c.key not in score_inputs or score_inputs[c.key] is None]
                        if missing:
                            st.error("Niet alle criteria zijn ingevuld.")
                        else:
                            # Update price in venues sheet
                            for i, vv in enumerate(venues):
                                if vv.key == v.key:
                                    venues[i].price = float(price_val)
                                    break
                            write_venues(venues)
                            ok, err = append_submission(current_tester.name, v, score_inputs, remark)
                            if ok:
                                st.toast(f"Bedankt {current_tester.name}! {v.name} is opgeslagen.")
                            else:
                                st.error(f"Opslaan in de cloud is niet gelukt: {err}")

# ------------------
# Tab: Resultaten
# ------------------
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

        col1, col2 = st.columns([1,2])
        with col1:
            end_clicked = st.button("⛔ Einde Vegdel", disabled=st.session_state.finalized or (not st.session_state.started) or (not st.session_state.admin_mode))
        with col2:
            st.caption("Bij beëindigen wordt alles vergrendeld en een PDF-rapport met grafieken gegenereerd. Actieve venues tellen mee; gesloten/inactieve niet.")

        if end_clicked:
            st.session_state.finalized = True
            write_flags(finalized=True)
            pdf_bytes = build_pdf(df, top1)
            st.session_state.final_pdf = pdf_bytes
            st.success("Eindrapport is gegenereerd en de invoer is vergrendeld.")

        if st.session_state.final_pdf:
            st.download_button("📄 Download eindrapport (PDF)", data=st.session_state.final_pdf, file_name="vegdel_eindrapport.pdf", mime="application/pdf")

# ------------------
# Tab: Beheer (alleen voor ingelogde beheerder)
# ------------------
with beheer_tab:
    st.subheader("🗂️ Beheer")
    if not st.session_state.admin_mode:
        st.warning("Alleen toegankelijk voor beheerder. Log in met PIN 1000 in de sidebar.")
    else:
        st.success("Beheer actief")
        # ---- Venues beheren ----
        st.markdown("### 🏪 Cafetaria's beheren")
        try:
            v_list = read_venues()
        except Exception:
            v_list = venues
        v_df = pd.DataFrame([{ "key": v.key, "Naam": v.name, "Actief": bool(v.active), "Prijs (EUR)": float(v.price)} for v in v_list])
        v_edit = st.data_editor(
            v_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "key": st.column_config.TextColumn("key", disabled=True),
                "Naam": st.column_config.TextColumn("Naam"),
                "Actief": st.column_config.CheckboxColumn("Actief"),
                "Prijs (EUR)": st.column_config.NumberColumn("Prijs (EUR)", min_value=0.0, step=0.05),
            },
        )
        col_save_v, col_hint_v = st.columns([1,2])
        with col_save_v:
            if st.button("💾 Venues opslaan"):
                new_vs = []
                for _, r in v_edit.iterrows():
                    name = str(r.get("Naam", "")).strip()
                    if not name:
                        continue
                    new_vs.append(Venue(str(r["key"]), name, bool(r.get("Actief", True)), float(r.get("Prijs (EUR)") or 0.0)))
                write_venues(new_vs)
                st.toast("Venues opgeslagen in Google Sheets.")
        with col_hint_v:
            st.caption("Tip: zet 'Actief' uit voor gesloten zaken. Nieuwe rijen kun je onderaan toevoegen.")

        st.divider()
        # ---- Criteria beheren ----
        st.markdown("### 📊 Criteria beheren")
        try:
            c_list = read_criteria()
        except Exception:
            c_list = criteria
        c_df = pd.DataFrame([{ "key": c.key, "Criterium": c.name, "Weging": float(c.weight)} for c in c_list])
        c_edit = st.data_editor(
            c_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "key": st.column_config.TextColumn("key", disabled=True),
                "Criterium": st.column_config.TextColumn("Criterium"),
                "Weging": st.column_config.NumberColumn("Weging", min_value=0.0, step=1.0),
            },
        )
        col_save_c, col_hint_c = st.columns([1,2])
        with col_save_c:
            if st.button("💾 Criteria opslaan"):
                new_cs = []
                for _, r in c_edit.iterrows():
                    nm = str(r.get("Criterium", "")).strip()
                    if not nm:
                        continue
                    new_cs.append(Criterion(str(r["key"]), nm, float(r.get("Weging") or 0.0)))
                write_criteria(new_cs)
                st.toast("Criteria opgeslagen in Google Sheets.")
        with col_hint_c:
            st.caption("Let op: aanpassen van criteria/weging werkt door in de berekening van alle resultaten.")

        st.divider()
        # ---- Data opschonen ----
        st.markdown("### ⚠️ Data opschonen")
        st.caption("Leeg alleen de testdata (submissions). Venues, criteria en flags blijven staan.")
        confirm_clear = st.checkbox("Ik bevestig dat ik alle testdata wil wissen", key="confirm_clear_submissions")
        if st.button("🧽 Wis alle testdata (submissions)"):
            if not confirm_clear:
                st.warning("Zet eerst de bevestigingsvink aan.")
            else:
                sh = ensure_worksheets()
                if sh:
                    ws = sh.worksheet(WS_SUBMISSIONS)
                    ws.clear()
                    ws.update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
                    st.toast("Alle testdata gewist.")
                else:
                    st.error("Kon Google Sheet niet openen. Controleer de Diagnose-tab.")

# ------------------
# Tab: Diagnose
# ------------------
with diag_tab:
    st.subheader("🧰 Diagnose & Verbinding")
    st.write("Service account geladen:", "✅" if st.secrets.get("google_service_account") else "❌ (niet gevonden)")
    sa_email = (st.secrets.get("google_service_account") or {}).get("client_email") if st.secrets.get("google_service_account") else None
    if sa_email:
        st.write("Service account email:", sa_email)
    st.write("SHEETS_ID (raw):", SHEET_ID_RAW if SHEET_ID_RAW else "(niet gezet)")
    st.write("SHEETS_ID (parsed):", SHEET_ID if SHEET_ID else "(kon niet parsen)")

    st.markdown("**Tijdelijk een Sheet ID/URL instellen (handig voor test):**")
    if "sheet_id_override" not in st.session_state:
        st.session_state.sheet_id_override = None
    tmp = st.text_input("Sheet ID of volledige URL", value=st.session_state.sheet_id_override or "")
    colx, coly = st.columns([1,1])
    with colx:
        if st.button("Gebruik tijdelijk deze SHEET_ID"):
            st.session_state.sheet_id_override = tmp.strip() or None
            st.experimental_rerun()
    with coly:
        if st.button("Wissen (override)"):
            st.session_state.sheet_id_override = None
            st.experimental_rerun()

    def diag_ping() -> str:
        try:
            sh = ensure_worksheets()
            ws = sh.worksheet(WS_SUBMISSIONS)
            ws.append_row(["__DIAG__", "vkey", "vname", 0.0, "ckey", "cname", 5.0, "diagnose"], value_input_option="RAW")
            return f"✅ Verbinding OK met {SHEET_ID}"
        except Exception as e:
            return f"❌ Fout tijdens verbindingstest: {e}"

    def diag_cleanup() -> str:
        try:
            sh = ensure_worksheets()
            ws = sh.worksheet(WS_SUBMISSIONS)
            data = ws.get_all_values()
            if not data:
                return "ℹ️ Geen data om op te schonen."
            header = data[0]
            rows = data[1:]
            try:
                idx = header.index('tester')
            except ValueError:
                return "❌ Kolom 'tester' niet gevonden."
            keep = [header] + [r for r in rows if (len(r) > idx and r[idx] != "__DIAG__")]
            ws.clear()
            ws.update(keep)
            return "🧽 Diagnose-rijen verwijderd."
        except Exception as e:
            return f"❌ Fout tijdens opschonen: {e}"

    if st.button("🔌 Test verbinding (ping)"):
        st.info("Diagnose bezig…")
        msg = diag_ping()
        if msg.startswith("✅"):
            st.success(msg)
        elif msg.startswith("⚠️"):
            st.warning(msg)
        else:
            st.error(msg)

    if st.button("🧽 Verwijder diagnose-rijen"):
        msg = diag_cleanup()
        if msg.startswith("🧽"):
            st.success(msg)
        else:
            st.info(msg)
