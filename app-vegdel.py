import json
import math
import uuid
from dataclasses import dataclass, asdict
from io import BytesIO
from typing import Dict, List, Any

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import gspread
from google.oauth2.service_account import Credentials
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# ==========================================
# Config & Theming
# ==========================================
PRIMARY = "#b94216"   # bruin (frikandel)
ACCENT = "#891418"     # curry
LIGHT  = "#ffffff"     # wit

st.set_page_config(page_title="Vegdel – Frikandel Speciaal", page_icon="🌭", layout="wide")

# Admin / beheer instellingen
ADMIN_PIN_SECRET = "1000"  # vaste pincode

# Google Sheets instellingen (vul in via Streamlit Secrets)
SHEET_ID = st.secrets.get("SHEET_ID", None)
GOOGLE_SA = st.secrets.get("google_service_account", None)

GC_SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_gs_client():
    if not GOOGLE_SA or not SHEET_ID:
        return None
    creds = Credentials.from_service_account_info(GOOGLE_SA, scopes=GC_SCOPE)
    client = gspread.authorize(creds)
    return client

def get_sheet():
    client = get_gs_client()
    if client is None:
        return None
    return client.open_by_key(SHEET_ID)

WS_FLAGS = "flags"
WS_VENUES = "venues"
WS_CRITERIA = "criteria"
WS_SUBMISSIONS = "submissions"

# Helper to ensure worksheets
def ensure_worksheets():
    sh = get_sheet()
    if sh is None:
        return None
    existing = {ws.title for ws in sh.worksheets()}
    wanted = {WS_FLAGS, WS_VENUES, WS_CRITERIA, WS_SUBMISSIONS}
    for w in wanted - existing:
        sh.add_worksheet(title=w, rows=1000, cols=26)
    # headers
    ws = sh.worksheet(WS_FLAGS)
    if ws.row_count == 0 or not ws.get_all_values():
        ws.update([ ["key", "value"], ["started", "False"], ["finalized", "False"] ])
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

# Load/save functions to Sheets

def read_flags() -> Dict[str, Any]:
    sh = ensure_worksheets()
    if sh is None:
        return {"started": st.session_state.started, "finalized": st.session_state.finalized}
    ws = sh.worksheet(WS_FLAGS)
    data = ws.get_all_records()
    out = {row["key"]: row["value"] for row in data}
    return {"started": str(out.get("started","False")).lower()=="true", "finalized": str(out.get("finalized","False")).lower()=="true"}

def write_flags(started: bool=None, finalized: bool=None):
    sh = get_sheet()
    if sh is None: return
    ws = sh.worksheet(WS_FLAGS)
    records = ws.get_all_records()
    m = {r["key"]: r["value"] for r in records}
    if started is not None: m["started"] = str(bool(started))
    if finalized is not None: m["finalized"] = str(bool(finalized))
    rows = [["key","value"]] + [[k,v] for k,v in m.items()]
    ws.update(rows)

def read_venues() -> List[Venue]:
    sh = ensure_worksheets()
    if sh is None:
        return venues
    ws = sh.worksheet(WS_VENUES)
    rows = ws.get_all_records()
    return [Venue(str(r["key"]), str(r["name"]), bool(r.get("active", True)), float(r.get("price",0.0))) for r in rows]

def write_venues(vs: List[Venue]):
    sh = get_sheet()
    if sh is None: return
    ws = sh.worksheet(WS_VENUES)
    rows = [["key","name","active","price"]] + [[v.key, v.name, v.active, v.price] for v in vs]
    ws.update(rows)

def read_criteria() -> List[Criterion]:
    sh = ensure_worksheets()
    if sh is None:
        return criteria
    ws = sh.worksheet(WS_CRITERIA)
    rows = ws.get_all_records()
    return [Criterion(str(r["key"]), str(r["name"]), float(r.get("weight",1))) for r in rows]

def write_criteria(cs: List[Criterion]):
    sh = get_sheet()
    if sh is None: return
    ws = sh.worksheet(WS_CRITERIA)
    rows = [["key","name","weight"]] + [[c.key, c.name, c.weight] for c in cs]
    ws.update(rows)

def append_submission(tester_name: str, v: Venue, scores_dict: Dict[str, float], remark: str):
    sh = ensure_worksheets()
    if sh is None:
        return False
    ws = sh.worksheet(WS_SUBMISSIONS)
    # Append one row per criterion for simple aggregation
    rows = []
    for c in criteria:
        val = float(scores_dict.get(c.key, 0.0))
        rows.append([tester_name, v.key, v.name, float(v.price), c.key, c.name, val, remark])
    ws.append_rows(rows, value_input_option="RAW")
    return True

def load_submissions_df() -> pd.DataFrame:
    sh = ensure_worksheets()
    if sh is None:
        return pd.DataFrame(columns=["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"])
    ws = sh.worksheet(WS_SUBMISSIONS)
    data = ws.get_all_records()
    return pd.DataFrame(data)

# ==========================================
# Data Models & Defaults
# ==========================================

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
DEFAULT_TESTERS: List[Tester] = []

if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False
if "started" not in st.session_state:
    st.session_state.started = False
if "finalized" not in st.session_state:
    st.session_state.finalized = False
if "venues" not in st.session_state:
    st.session_state.venues = DEFAULT_VENUES
if "criteria" not in st.session_state:
    st.session_state.criteria = DEFAULT_CRITERIA
if "testers" not in st.session_state:
    st.session_state.testers = DEFAULT_TESTERS
if "scores" not in st.session_state:
    st.session_state.scores = {}
if "remarks" not in st.session_state:
    st.session_state.remarks = {}
if "final_pdf" not in st.session_state:
    st.session_state.final_pdf = None

venues = st.session_state.venues
criteria = st.session_state.criteria
testers = st.session_state.testers
scores = st.session_state.scores
remarks = st.session_state.remarks

# Poll shared flags from Google Sheets for multi-user sync
flags = read_flags()
st.session_state.started = flags.get("started", st.session_state.started)
st.session_state.finalized = flags.get("finalized", st.session_state.finalized)

# Periodic auto-refresh while running
if not st.session_state.finalized:
    st.autorefresh(interval=5000, key="polling")

# ==========================================
# Helpers
# ==========================================

def diag_ping() -> str:
    """Try to write & read a diagnostic row in submissions."""
    try:
        sh = ensure_worksheets()
        if sh is None:
            return "❌ Geen verbinding met Google Sheets (controleer secrets SHEET_ID en google_service_account)."
        # ensure venues exist and pick first active
        vlist = read_venues()
        v = next((x for x in vlist if x.active), None)
        if v is None:
            return "⚠️ Er zijn geen actieve cafetaria's (Instellen-tab)."
        # build fake scores using current criteria (all 5's)
        fake_scores = {c.key: 5.0 for c in criteria}
        ok = append_submission("__DIAG__", v, fake_scores, "diagnose ping")
        if not ok:
            return "❌ Kon niet naar 'submissions' schrijven. Controleer schrijfrechten."
        # read back
        df = load_submissions_df()
        count = int((df['tester'] == "__DIAG__").sum()) if not df.empty else 0
        return f"✅ Verbinding OK. Submissions bevat nu {count} DIAG-rij(en)."
    except Exception as e:
        return f"❌ Fout tijdens diagnose: {e}"


def diag_cleanup() -> str:
    try:
        sh = ensure_worksheets()
        if sh is None:
            return "❌ Geen verbinding met Google Sheets."
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


def total_weight() -> float:
    s = sum(max(0.0, float(c.weight)) for c in criteria)
    return s if s > 0 else 1.0

def ensure_score_slot(tk: str, vk: str):
    if tk not in scores:
        scores[tk] = {}
    if vk not in scores[tk]:
        scores[tk][vk] = {}

def compute_results() -> pd.DataFrame:
    # Compute results from submissions if available; fall back to local scores
    sub_df = load_submissions_df()
    if not sub_df.empty:
        # Use current criteria weights
        w = {c.key: float(c.weight) for c in criteria}
        # Compute weighted avg per tester per venue
        sub_df["w"] = sub_df["criterion_key"].map(w).fillna(0.0)
        # Guard division
        TW = sum(w.values()) or 1.0
        agg = (sub_df.groupby(["tester","venue_key","venue_name","price"], as_index=False)
                      .apply(lambda g: (g["score"]*g["w"]).sum()/TW).reset_index(name="tester_avg"))
        # Average across testers
        res = agg.groupby(["venue_key","venue_name","price"], as_index=False).agg(
            **{"Gem. score (0-10)": ("tester_avg","mean"), "# Testers": ("tester","nunique")}
        )
        res["Gem. score (0-10)"] = res["Gem. score (0-10)"].round(2)
        res = res.sort_values(["Gem. score (0-10)", "# Testers"], ascending=[False, False]).reset_index(drop=True)
        res.insert(0, "Rank", range(1, len(res)+1))
        res.rename(columns={"venue_name":"Cafetaria","price":"Prijs"}, inplace=True)
        return res
    # Fallback to in-session (single-user) storage
    active_venues = [v for v in venues if v.active]
    TW = total_weight()
    rows = []
    for v in active_venues:
        tester_avgs = []
        for t in testers:
            t_scores = scores.get(t.key, {}).get(v.key, {})
            if not t_scores:
                continue
            s = sum(float(t_scores.get(c.key, 0.0)) * float(c.weight) for c in criteria)
            tester_avgs.append(s / TW)
        avg = float(np.mean(tester_avgs)) if tester_avgs else 0.0
        rows.append({
            "venue_key": v.key,
            "Cafetaria": v.name,
            "Prijs": v.price,
            "Gem. score (0-10)": round(avg, 2),
            "# Testers": len(tester_avgs),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(by=["Gem. score (0-10)", "# Testers"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))
    return df
    active_venues = [v for v in venues if v.active]
    TW = total_weight()
    rows = []
    for v in active_venues:
        tester_avgs = []
        for t in testers:
            t_scores = scores.get(t.key, {}).get(v.key, {})
            if not t_scores:
                continue
            s = sum(float(t_scores.get(c.key, 0.0)) * float(c.weight) for c in criteria)
            tester_avgs.append(s / TW)
        avg = float(np.mean(tester_avgs)) if tester_avgs else 0.0
        rows.append({
            "venue_key": v.key,
            "Cafetaria": v.name,
            "Prijs": v.price,
            "Gem. score (0-10)": round(avg, 2),
            "# Testers": len(tester_avgs),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(by=["Gem. score (0-10)", "# Testers"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))
    return df

def per_tester_top1() -> pd.DataFrame:
    active = [v for v in venues if v.active]
    rows = []
    TW = total_weight()
    for t in testers:
        best_name = None
        best_score = -1.0
        for v in active:
            t_scores = scores.get(t.key, {}).get(v.key, {})
            if not t_scores:
                continue
            s = sum(float(t_scores.get(c.key, 0.0)) * float(c.weight) for c in criteria) / TW
            if s > best_score:
                best_score = s
                best_name = v.name
        if best_name is not None:
            rows.append({"Tester": t.name, "#1 Cafetaria": best_name, "Score (0-10)": round(float(best_score), 2)})
    return pd.DataFrame(rows)

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
    story.append(Paragraph("<b>Vegdel – Frikandel Speciaal</b>", styles['Title']))
    story.append(Spacer(1, 0.3*cm))

    tbl_df = df[["Rank", "Cafetaria", "Prijs", "Gem. score (0-10)", "# Testers"]]
    data = [list(tbl_df.columns)] + tbl_df.values.tolist()
    table = Table(data, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(PRIMARY)),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.25, colors.gray),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#f7efe9')]),
    ]))
    story.append(Paragraph("<b>Uitslag & Ranking</b>", styles['Heading2']))
    story.append(table)
    story.append(Spacer(1, 0.4*cm))

    if top1_df is not None and not top1_df.empty:
        tdata = [list(top1_df.columns)] + top1_df.values.tolist()
        ttable = Table(tdata, hAlign='LEFT')
        ttable.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(ACCENT)),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.25, colors.gray),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#f7efe9')]),
        ]))
        story.append(Paragraph("<b>Persoonlijke #1 per tester</b>", styles['Heading2']))
        story.append(ttable)

    chart_buf = plot_ranking(df)
    from reportlab.platypus import Image as RLImage
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("<b>Grafiek – Gemiddelde scores</b>", styles['Heading2']))
    story.append(RLImage(chart_buf, width=16*cm, height=9*cm))

    doc.build(story)
    buffer.seek(0)
    return buffer.read()

# ==========================================
# Header & Sidebar (Beheer)
# ==========================================
st.title("🥇 Vegdel – Beste friettent (Frikandel Speciaal)")
st.caption("Testdatum: 18-10-2025 • Regels: zelfde portie • Testers vullen in op eigen device")

with st.sidebar:
    st.header("🔐 Beheer")
    if not st.session_state.admin_mode:
        pin_try = st.text_input("Voer pincode in", type="password")
        if st.button("Inloggen beheer"):
            if pin_try == ADMIN_PIN_SECRET:
                st.session_state.admin_mode = True
                st.success("Ingelogd als beheerder.")
            else:
                st.error("Onjuiste pincode.")
    else:
        st.success("Beheer actief")
        colA, colB = st.columns(2)
        with colA:
            if st.button("▶️ Start Vegdel", disabled=st.session_state.started or st.session_state.finalized):
                st.session_state.started = True
                write_flags(started=True)
                st.toast("Vegdel gestart — invoer geactiveerd voor iedereen.")
        with colB:
            if st.button("🗑️ Alles wissen", type="secondary"):
                confirm = st.checkbox("Ik bevestig het wissen van alle data")
                if confirm:
                    # Reset Google Sheet content
                    sh = ensure_worksheets()
                    if sh:
                        sh.worksheet(WS_SUBMISSIONS).clear()
                        sh.worksheet(WS_SUBMISSIONS).update([["tester","venue_key","venue_name","price","criterion_key","criterion_name","score","remark"]])
                        write_flags(started=False, finalized=False)
                        # reset venues & criteria to defaults
                        write_venues(DEFAULT_VENUES)
                        write_criteria(DEFAULT_CRITERIA)
                    # reset local session
                    st.session_state.scores = {}
                    st.session_state.remarks = {}
                    st.session_state.finalized = False
                    st.session_state.final_pdf = None
                    st.session_state.started = False
                    st.session_state.venues = DEFAULT_VENUES
                    st.session_state.criteria = DEFAULT_CRITERIA
                    st.session_state.testers = []
                    st.toast("Alle data gewist (cloud + sessie).")
                else:
                    st.warning("Vink de bevestiging aan om te wissen.")
        if st.button("🔓 Uitloggen beheer"):
            st.session_state.admin_mode = False
            st.toast("Beheer uitgelogd.")

if st.session_state.finalized:
    st.warning("De test is beëindigd. Wijzigingen zijn vergrendeld.")
elif not st.session_state.started:
    st.info("Nog niet gestart. Beheerder kan starten via de sidebar.")

# ==========================================
# Tabs
# ==========================================
setup_tab, score_tab, results_tab, diag_tab = st.tabs(["⚙️ Instellen", "🧪 Scoren", "🏆 Resultaten", "🧰 Diagnose"]) 

# ------------------
# Tab: Instellen
# ------------------
with setup_tab:
    st.subheader("Cafetaria's & prijzen")
    v_df = pd.DataFrame([
        {"key": v.key, "Actief": v.active, "Naam": v.name, "Prijs (EUR)": float(v.price)} for v in venues
    ])
    v_edit = st.data_editor(
        v_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        disabled=st.session_state.finalized,
        column_config={
            "key": st.column_config.TextColumn("key", disabled=True),
            "Actief": st.column_config.CheckboxColumn("Actief"),
            "Prijs (EUR)": st.column_config.NumberColumn("Prijs (EUR)", min_value=0.0, step=0.05),
        },
    )
    if not st.session_state.finalized:
        st.session_state.venues = [Venue(row["key"], str(row["Naam"]).strip(), bool(row["Actief"]), float(row["Prijs (EUR)"] or 0.0))
                                   for _, row in v_edit.iterrows() if str(row.get("Naam", "")).strip()]
        venues = st.session_state.venues

    st.divider()
    st.subheader("Criteria & Weging (0–10, stap 0,5)")
    c_df = pd.DataFrame([{ "key": c.key, "Criterium": c.name, "Weging": float(c.weight)} for c in criteria])
    c_edit = st.data_editor(
        c_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        disabled=st.session_state.finalized,
        column_config={
            "key": st.column_config.TextColumn("key", disabled=True),
            "Weging": st.column_config.NumberColumn("Weging", min_value=0.0, step=1.0),
        },
    )
    if not st.session_state.finalized:
        st.session_state.criteria = [Criterion(row["key"], str(row["Criterium"]).strip(), float(row["Weging"]))
                                     for _, row in c_edit.iterrows() if str(row.get("Criterium", "")).strip()]
        criteria = st.session_state.criteria

    st.caption(f"Totale weging: **{sum(max(0.0, float(c.weight)) for c in criteria):.0f}**")

    st.divider()
    st.subheader("Testers (optioneel vooraf toevoegen)")
    t_df = pd.DataFrame([{ "key": t.key, "Naam": t.name } for t in testers])
    t_edit = st.data_editor(
        t_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        disabled=st.session_state.finalized,
        column_config={"key": st.column_config.TextColumn("key", disabled=True)},
    )
    if not st.session_state.finalized:
        st.session_state.testers = [Tester(row["key"], str(row["Naam"]).strip()) for _, row in t_edit.iterrows() if str(row.get("Naam", "")).strip()]
        testers = st.session_state.testers

# ------------------
# Tab: Scoren
# ------------------
with score_tab:
    if not st.session_state.started and not st.session_state.finalized:
        st.warning("De test is nog niet gestart. Beheerder kan starten via de sidebar.")
    if not any(v.active for v in venues) or not criteria:
        st.warning("Zorg dat er actieve cafetaria's en criteria zijn in 'Instellen'.")
    else:
        st.markdown("#### Wie ben je?")
        your_name = st.text_input("Vul je naam in", placeholder="Voor- en achternaam", disabled=st.session_state.finalized)
        current_tester = None
        if your_name:
            existing = next((t for t in testers if t.name.strip().lower() == your_name.strip().lower()), None)
            if existing is None and not st.session_state.finalized:
                new_t = Tester(uid(), your_name.strip())
                st.session_state.testers.append(new_t)
                testers.append(new_t)
                current_tester = new_t
            else:
                current_tester = existing
        else:
            st.info("Vul eerst je naam in om te kunnen scoren.")

        if current_tester is not None:
            st.markdown("---")
            st.subheader(f"Beoordelen als: {current_tester.name}")
            active_venues = [v for v in venues if v.active]
            TW = sum(max(0.0, float(c.weight)) for c in criteria) or 1.0

            for v in active_venues:
                with st.form(key=f"form_{current_tester.key}_{v.key}", clear_on_submit=False):
                    st.markdown(f"### {v.name}  ")
                    cols = st.columns(2)
                    with cols[0]:
                        st.write("**Prijs (EUR)**")
                        price_val = st.number_input("Prijs", min_value=0.0, step=0.05, value=float(v.price), key=f"price_{v.key}")
                    with cols[1]:
                        st.write("**Opmerkingen (optioneel)**")
                        remark = st.text_area("Opmerkingen", value=st.session_state.remarks.get(current_tester.key, {}).get(v.key, ""), key=f"remark_{current_tester.key}_{v.key}")

                    st.write("**Scores (0–10, stap 0,5):** Alle velden verplicht")
                    score_inputs = {}
                    for c in criteria:
                        score_inputs[c.key] = st.number_input(c.name, min_value=0.0, max_value=10.0, step=0.5,
                                                              value=float(scores.get(current_tester.key, {}).get(v.key, {}).get(c.key, 0.0)),
                                                              key=f"score_{current_tester.key}_{v.key}_{c.key}")

                    saved = st.form_submit_button("Opslaan cafetaria")
                    if saved and not st.session_state.finalized:
                        # Validate required: every criterion must be provided
                        missing = [c.name for c in criteria if c.key not in score_inputs or score_inputs[c.key] is None]
                        if missing:
                            st.error("Niet alle criteria zijn ingevuld.")
                        else:
                            # Save price locally and to sheet venues
                            for i, vv in enumerate(st.session_state.venues):
                                if vv.key == v.key:
                                    st.session_state.venues[i].price = float(price_val)
                                    break
                            write_venues(st.session_state.venues)
                            # Persist submission to Google Sheets (append rows)
                            ok = append_submission(current_tester.name, v, score_inputs, remark)
                            if ok:
                                st.toast(f"Bedankt {current_tester.name}! {v.name} is opgeslagen.")
                            else:
                                st.error("Opslaan in de cloud is niet gelukt. Controleer de Google Sheets configuratie in secrets.")

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
# Tab: Diagnose
# ------------------
with diag_tab:
    st.subheader("🧰 Diagnose & Verbinding")
    st.write("SHEETS_ID:", SHEET_ID if SHEET_ID else "(niet gezet)")
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

