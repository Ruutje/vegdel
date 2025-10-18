import json
import math
import uuid
from dataclasses import dataclass, asdict
from io import BytesIO
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
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

# ==========================================
# Helpers
# ==========================================

def total_weight() -> float:
    s = sum(max(0.0, float(c.weight)) for c in criteria)
    return s if s > 0 else 1.0

def ensure_score_slot(tk: str, vk: str):
    if tk not in scores:
        scores[tk] = {}
    if vk not in scores[tk]:
        scores[tk][vk] = {}

def compute_results() -> pd.DataFrame:
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
                st.toast("Vegdel gestart — invoer geactiveerd.")
        with colB:
            if st.button("🗑️ Alles wissen", type="secondary"):
                confirm = st.checkbox("Ik bevestig het wissen van alle data")
                if confirm:
                    st.session_state.scores = {}
                    st.session_state.remarks = {}
                    st.session_state.finalized = False
                    st.session_state.final_pdf = None
                    st.session_state.started = False
                    st.session_state.venues = DEFAULT_VENUES
                    st.session_state.criteria = DEFAULT_CRITERIA
                    st.session_state.testers = []
                    st.toast("Alle data gewist en gereset.")
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
setup_tab, score_tab, results_tab = st.tabs(["⚙️ Instellen", "🧪 Scoren", "🏆 Resultaten"]) 

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
    if not testers or not any(v.active for v in venues) or not criteria:
        st.warning("Zorg dat er testers, actieve cafetaria's en criteria zijn in 'Instellen'.")
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
                            # Save price
                            for i, vv in enumerate(st.session_state.venues):
                                if vv.key == v.key:
                                    st.session_state.venues[i].price = float(price_val)
                                    break
                            # Save scores
                            if current_tester.key not in st.session_state.scores:
                                st.session_state.scores[current_tester.key] = {}
                            if v.key not in st.session_state.scores[current_tester.key]:
                                st.session_state.scores[current_tester.key][v.key] = {}
                            for ck, val in score_inputs.items():
                                st.session_state.scores[current_tester.key][v.key][ck] = float(val)
                            # Save remark (optional)
                            if current_tester.key not in st.session_state.remarks:
                                st.session_state.remarks[current_tester.key] = {}
                            st.session_state.remarks[current_tester.key][v.key] = remark
                            st.toast(f"Bedankt {current_tester.name}! {v.name} is opgeslagen.")

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
            pdf_bytes = build_pdf(df, top1)
            st.session_state.final_pdf = pdf_bytes
            st.success("Eindrapport is gegenereerd en de invoer is vergrendeld.")

        if st.session_state.final_pdf:
            st.download_button("📄 Download eindrapport (PDF)", data=st.session_state.final_pdf, file_name="vegdel_eindrapport.pdf", mime="application/pdf")
