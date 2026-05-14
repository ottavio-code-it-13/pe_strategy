# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════╗
║   PE STRATEGY RANKING ENGINE  ·  Fase 4                         ║
║   Integra output Persona 3, Persona 4, Persona 5                 ║
║   Run:  streamlit run pe_strategy_app.py                         ║
╚══════════════════════════════════════════════════════════════════╝
"""

import io
import os
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# ── Auto-detect local Excel files ────────────────────────────────
# Se i file sono nella stessa cartella dello script, vengono caricati
# automaticamente senza bisogno di upload manuale.

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

_LOCAL_FILES = {
    "p3": [
        "output_persona3_v4_no_leakage.xlsx",
        "output_persona3.xlsx",
    ],
    "p4": [
        "Output_persona4_cash_flow_resilience_output.xlsx",
        "cash_flow_resilience_output.xlsx",
        "output_persona4.xlsx",
    ],
    "p5": [
        "persona5_FMS_score_team__7_.xlsx",
        "persona5_FMS_score.xlsx",
        "output_persona5.xlsx",
    ],
}

def _find_local(key: str):
    """Restituisce il path del file locale se esiste, altrimenti None."""
    for fname in _LOCAL_FILES[key]:
        full = os.path.join(_SCRIPT_DIR, fname)
        if os.path.isfile(full):
            return full
    return None

_auto_p3 = _find_local("p3")
_auto_p4 = _find_local("p4")
_auto_p5 = _find_local("p5")

# ── Page config ──────────────────────────────────────────────────
st.set_page_config(
    page_title="PE Strategy Engine · Fase 4",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');

/* ── ROOT THEME ─────────────────────────── */
:root {
    --bg:        #0B0F1A;
    --surface:   #111827;
    --border:    #1E2D40;
    --accent:    #00C8FF;
    --accent2:   #FF6B35;
    --accent3:   #B8FF47;
    --text:      #E8EDF5;
    --muted:     #6B7B92;
    --op-color:  #00C8FF;
    --cf-color:  #B8FF47;
    --fms-color: #FF6B35;
}

html, body, [data-testid="stAppViewContainer"] {
    background-color: var(--bg) !important;
    font-family: 'DM Mono', monospace;
    color: var(--text);
}

[data-testid="stSidebar"] {
    background-color: var(--surface) !important;
    border-right: 1px solid var(--border);
}

/* ── HEADER BANNER ──────────────────────── */
.banner {
    background: linear-gradient(135deg, #0f1f35 0%, #0B0F1A 60%, #111827 100%);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 32px 40px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.banner::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 220px; height: 220px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(0,200,255,0.08) 0%, transparent 70%);
}
.banner-title {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 2.1rem;
    letter-spacing: -0.5px;
    color: var(--text);
    line-height: 1.1;
    margin-bottom: 6px;
}
.banner-sub {
    font-family: 'DM Mono', monospace;
    font-size: 0.75rem;
    color: var(--muted);
    letter-spacing: 2px;
    text-transform: uppercase;
}
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 4px;
    font-size: 0.65rem;
    font-weight: 600;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin-left: 10px;
    vertical-align: middle;
}
.badge-op  { background: rgba(0,200,255,0.15); color: var(--op-color);  border: 1px solid rgba(0,200,255,0.3); }
.badge-cf  { background: rgba(184,255,71,0.12); color: var(--cf-color); border: 1px solid rgba(184,255,71,0.3); }
.badge-fms { background: rgba(255,107,53,0.15); color: var(--fms-color);border: 1px solid rgba(255,107,53,0.3); }

/* ── METRIC CARDS ───────────────────────── */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 28px;
}
.metric-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 20px 22px;
    position: relative;
    overflow: hidden;
}
.metric-card::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 2px;
}
.metric-card.op::after  { background: var(--op-color); }
.metric-card.cf::after  { background: var(--cf-color); }
.metric-card.fms::after { background: var(--fms-color); }
.metric-card.total::after { background: var(--muted); }
.metric-label {
    font-size: 0.65rem;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 8px;
}
.metric-value {
    font-family: 'Syne', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    line-height: 1;
}
.metric-sub {
    font-size: 0.65rem;
    color: var(--muted);
    margin-top: 4px;
}

/* ── SECTION HEADERS ────────────────────── */
.section-header {
    font-family: 'Syne', sans-serif;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: var(--muted);
    border-bottom: 1px solid var(--border);
    padding-bottom: 10px;
    margin: 32px 0 20px 0;
}

/* ── STRATEGY SELECTOR CARDS ────────────── */
.strat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 20px;
    cursor: pointer;
    transition: border-color 0.2s, box-shadow 0.2s;
    margin-bottom: 8px;
}
.strat-card:hover { border-color: var(--accent); }
.strat-card.active { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }

/* ── RANK TABLE ─────────────────────────── */
.rank-table-wrap {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    overflow: hidden;
}

/* ── STREAMLIT OVERRIDES ────────────────── */
[data-testid="metric-container"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    padding: 16px !important;
}
div[data-testid="stDataFrame"] {
    background: var(--surface) !important;
}
.stSelectbox > div > div {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    color: var(--text) !important;
}
.stSlider > div { color: var(--text) !important; }
h1,h2,h3 {
    font-family: 'Syne', sans-serif !important;
    color: var(--text) !important;
}
.stButton > button {
    background: var(--accent) !important;
    color: var(--bg) !important;
    font-family: 'DM Mono', monospace !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 6px !important;
    letter-spacing: 1px;
}
.stDownloadButton > button {
    background: transparent !important;
    color: var(--accent) !important;
    border: 1px solid var(--accent) !important;
    font-family: 'DM Mono', monospace !important;
    border-radius: 6px !important;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'DM Mono', monospace !important;
    font-size: 0.75rem !important;
    letter-spacing: 1px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════

PLOTLY_TEMPLATE = dict(
    layout=dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Mono, monospace", color="#E8EDF5"),
        colorway=["#00C8FF", "#B8FF47", "#FF6B35", "#A78BFA", "#F59E0B"],
        xaxis=dict(gridcolor="#1E2D40", zerolinecolor="#1E2D40"),
        yaxis=dict(gridcolor="#1E2D40", zerolinecolor="#1E2D40"),
    )
)

COLOR_MAP = {
    "operational_upside_score":    "#00C8FF",
    "cash_flow_resilience_score":  "#B8FF47",
    "fms_score":                   "#FF6B35",
}

SCORE_LABELS = {
    "operational_upside_score":   "Operational Upside",
    "cash_flow_resilience_score": "Cash Flow Resilience",
    "fms_score":                  "FMS Score (Add-on)",
}


def normalize_mark(df: pd.DataFrame, mark_col: str) -> pd.DataFrame:
    df = df.copy()
    df = df.rename(columns={mark_col: "mark"})
    df["mark"] = df["mark"].astype(str).str.strip().str.replace(".0", "", regex=False)
    return df


@st.cache_data(show_spinner=False)
def load_persona3(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="Sheet1")
    mark_col = "Mark" if "Mark" in df.columns else "mark"
    df = normalize_mark(df, mark_col)
    keep = {
        "mark": "mark",
        "Ragione sociale": "company",
        "Provincia": "province",
        "operational_upside_score": "operational_upside_score",
        "recoverability_score_base": "recoverability_score_base",
        "recoverability_score_historical": "recoverability_score_historical",
        "distress_penalty": "distress_penalty",
        "ebitda_margin_2024": "ebitda_margin_2024",
        "peer_ebitda_median": "peer_ebitda_median",
        "gap_vs_peer": "gap_vs_peer",
        "operational_driver": "operational_driver",
        "upside_assoluto_pp": "upside_assoluto_pp",
        "cluster": "cluster",
    }
    df = df[[c for c in keep if c in df.columns]].rename(columns=keep)
    return df


@st.cache_data(show_spinner=False)
def load_persona4(file) -> pd.DataFrame:
    try:
        df = pd.read_excel(file, sheet_name="Cash Flow Score")
    except Exception:
        df = pd.read_excel(file)
    mark_col = "mark" if "mark" in df.columns else "Mark"
    df = normalize_mark(df, mark_col)
    keep = {
        "mark": "mark",
        "company": "company",
        "province": "province",
        "cash_flow_resilience_score": "cash_flow_resilience_score",
        "resilience_class": "resilience_class",
        "probability_of_default": "probability_of_default",
        "fcf_stability_score": "fcf_stability_score",
        "dscr_score": "dscr_score",
        "pd_score": "pd_score",
        "main_value_driver": "main_value_driver",
        "investment_thesis": "investment_thesis",
        "latest_revenue": "latest_revenue",
        "latest_ebitda": "latest_ebitda",
        "avg_fcf_margin": "avg_fcf_margin",
    }
    df = df[[c for c in keep if c in df.columns]].rename(columns=keep)
    return df


@st.cache_data(show_spinner=False)
def load_persona5(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="Sheet1")
    # Persona 5 does not have a numeric 'mark' – use company name as key
    df = df.copy()
    df["Azienda"] = df["Azienda"].astype(str).str.strip()
    keep = {
        "Azienda": "company_p5",
        "Provincia": "province_p5",
        "Macro Area": "macro_area",
        "Fatturato (€M)": "fatturato_m",
        "EBITDA Margin": "ebitda_margin_p5",
        "Net Debt/EBITDA": "net_debt_ebitda_p5",
        "Add-on Targets": "addon_targets",
        "FMS Score (0-100)": "fms_score",
        "Rank": "fms_rank",
        "Cluster": "cluster_p5",
        "Tier": "tier_p5",
    }
    df = df[[c for c in keep if c in df.columns]].rename(columns=keep)
    return df


def build_master(p3: pd.DataFrame, p4: pd.DataFrame,
                 p5: pd.DataFrame | None) -> pd.DataFrame:
    """Merge on mark (outer join P3 ∩ P4), P5 separate universe."""
    master = p3.merge(p4, on="mark", how="outer", suffixes=("_p3", "_p4"))

    # Consolidate company / province
    for field in ("company", "province"):
        c_p3 = f"{field}_p3" if f"{field}_p3" in master.columns else field
        c_p4 = f"{field}_p4" if f"{field}_p4" in master.columns else field
        if c_p3 in master.columns and c_p4 in master.columns:
            master[field] = master[c_p3].combine_first(master[c_p4])
            master = master.drop(columns=[c_p3, c_p4], errors="ignore")
        elif field not in master.columns:
            master[field] = np.nan

    # P5 score: fuzzy-join on company name where possible, else NaN
    master["fms_score"]     = np.nan
    master["tier_p5"]       = pd.Series(pd.NA, index=master.index, dtype=object)
    master["addon_targets"] = pd.Series(pd.NA, index=master.index, dtype=object)
    if p5 is not None and not p5.empty:
        p5_lookup = p5.set_index("company_p5")[["fms_score", "tier_p5", "addon_targets"]]
        for idx, row in master.iterrows():
            name = str(row.get("company", "")).strip()
            if name in p5_lookup.index:
                master.at[idx, "fms_score"]     = p5_lookup.at[name, "fms_score"]
                master.at[idx, "tier_p5"]        = p5_lookup.at[name, "tier_p5"]
                master.at[idx, "addon_targets"]  = p5_lookup.at[name, "addon_targets"]

    return master


def score_fmt(v):
    if pd.isna(v):
        return "—"
    return f"{v:.1f}"


def pct_fmt(v):
    if pd.isna(v):
        return "—"
    return f"{v*100:.1f}%"


def eur_fmt(v):
    if pd.isna(v):
        return "—"
    if v >= 1_000_000:
        return f"€{v/1_000_000:.1f}M"
    return f"€{v/1_000:.0f}K"


# ══════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("""
    <div style='padding:16px 0 24px 0'>
        <div style='font-family:Syne,sans-serif;font-size:1.1rem;font-weight:800;
                    color:#E8EDF5;letter-spacing:-0.3px;'>PE Strategy</div>
        <div style='font-size:0.6rem;letter-spacing:2.5px;color:#6B7B92;
                    text-transform:uppercase;margin-top:2px;'>Ranking Engine · Fase 4</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 📂 File di input")

    # ── Persona 3 ─────────────────────────────────────────────────
    if _auto_p3:
        st.success(f"✅ P3 rilevato automaticamente")
        st.caption(f"`{os.path.basename(_auto_p3)}`")
        p3_file = _auto_p3          # path stringa → gestito in load_persona3
        st.file_uploader(
            "Sostituisci Persona 3 (opzionale)",
            type=["xlsx"], key="p3",
            help="Trascina un file diverso per sovrascrivere quello locale"
        )
        if st.session_state.get("p3"):
            p3_file = st.session_state["p3"]
    else:
        st.warning("⚠️ P3 non trovato nella cartella")
        p3_file = st.file_uploader(
            "Persona 3 — Operational Improvement",
            type=["xlsx"], key="p3",
            help="output_persona3_v4_no_leakage.xlsx"
        )

    # ── Persona 4 ─────────────────────────────────────────────────
    if _auto_p4:
        st.success(f"✅ P4 rilevato automaticamente")
        st.caption(f"`{os.path.basename(_auto_p4)}`")
        p4_file = _auto_p4
        st.file_uploader(
            "Sostituisci Persona 4 (opzionale)",
            type=["xlsx"], key="p4",
            help="Trascina un file diverso per sovrascrivere quello locale"
        )
        if st.session_state.get("p4"):
            p4_file = st.session_state["p4"]
    else:
        st.warning("⚠️ P4 non trovato nella cartella")
        p4_file = st.file_uploader(
            "Persona 4 — Cash Flow Resilience",
            type=["xlsx"], key="p4",
            help="Output_persona4_cash_flow_resilience_output.xlsx"
        )

    # ── Persona 5 ─────────────────────────────────────────────────
    if _auto_p5:
        st.success(f"✅ P5 rilevato automaticamente")
        st.caption(f"`{os.path.basename(_auto_p5)}`")
        p5_file = _auto_p5
        st.file_uploader(
            "Sostituisci Persona 5 (opzionale)",
            type=["xlsx"], key="p5",
            help="Trascina un file diverso per sovrascrivere quello locale"
        )
        if st.session_state.get("p5"):
            p5_file = st.session_state["p5"]
    else:
        st.info("ℹ️ P5 non trovato (opzionale)")
        p5_file = st.file_uploader(
            "Persona 5 — FMS / Add-on  (opzionale)",
            type=["xlsx"], key="p5",
            help="persona5_FMS_score_team__7_.xlsx"
        )

    st.divider()

    st.markdown("#### ⚙️ Parametri ranking")

    strategy = st.radio(
        "Strategia PE",
        options=[
            "Operational Improvement",
            "Cash Flow Resilience / Deleverage",
            "FMS / Buy & Build (Add-on)",
            "Combined Score",
        ],
        index=0,
    )

    top_n = st.slider("Top N aziende", min_value=5, max_value=50, value=10, step=5)

    if strategy == "Combined Score":
        st.markdown("##### Pesi combined score")
        w_op = st.slider("Peso Operational", 0.0, 1.0, 0.40, 0.05)
        w_cf = st.slider("Peso Cash Flow",   0.0, 1.0, 0.40, 0.05)
        w_fms = st.slider("Peso FMS",         0.0, 1.0, 0.20, 0.05)
        total_w = w_op + w_cf + w_fms
        if abs(total_w - 1.0) > 0.01:
            st.warning(f"Pesi sommano a {total_w:.2f} — normalizzati automaticamente.")
    else:
        w_op, w_cf, w_fms = 0.4, 0.4, 0.2

    st.divider()
    st.markdown("""
    <div style='font-size:0.6rem;color:#3D4F63;line-height:1.6;'>
    Progetto Machine Learning<br>
    Team 7 · 2024/2025<br>
    Fase 4 — Strategy Ranking Engine
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# BANNER
# ══════════════════════════════════════════════════════════════════

st.markdown("""
<div class="banner">
  <div class="banner-title">
    PE Strategy Ranking Engine
    <span class="badge badge-op">P3</span>
    <span class="badge badge-cf">P4</span>
    <span class="badge badge-fms">P5</span>
  </div>
  <div class="banner-sub">Fase 4 · Integrazione output sottomodelli · Private Equity Target Selection</div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# DATA LOADING GATE
# ══════════════════════════════════════════════════════════════════

# ── Resolve file sources (path string or uploaded file object) ────
def _resolve(f):
    """Restituisce None se f è None o stringa vuota."""
    if f is None:
        return None
    if isinstance(f, str) and not os.path.isfile(f):
        return None
    return f

p3_src = _resolve(p3_file)
p4_src = _resolve(p4_file)
p5_src = _resolve(p5_file)

if not p3_src or not p4_src:
    st.markdown("""
    <div style='background:#111827;border:1px dashed #1E2D40;border-radius:10px;
                padding:48px;text-align:center;'>
        <div style='font-size:2.5rem;margin-bottom:16px;'>📁</div>
        <div style='font-family:Syne,sans-serif;font-size:1.1rem;color:#E8EDF5;
                    margin-bottom:8px;font-weight:700;'>
            Carica i file per iniziare
        </div>
        <div style='font-size:0.75rem;color:#6B7B92;max-width:420px;margin:0 auto;'>
            Carica almeno <strong style="color:#00C8FF">Persona 3</strong> e
            <strong style="color:#B8FF47">Persona 4</strong> dalla sidebar.
            Persona 5 è opzionale ma abilita la strategia FMS / Add-on.
        </div>
        <div style='margin-top:24px;display:flex;gap:12px;justify-content:center;
                    flex-wrap:wrap;font-size:0.7rem;color:#3D4F63;'>
            <span>output_persona3_v4_no_leakage.xlsx</span>
            <span>·</span>
            <span>Output_persona4_cash_flow_resilience_output.xlsx</span>
            <span>·</span>
            <span>persona5_FMS_score_team__7_.xlsx</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Load & merge ──────────────────────────────────────────────────
with st.spinner("Caricamento e integrazione dati..."):
    p3 = load_persona3(p3_src)
    p4 = load_persona4(p4_src)
    p5 = load_persona5(p5_src) if p5_src else None
    master = build_master(p3, p4, p5)

# ── Compute combined score ────────────────────────────────────────
total_w = w_op + w_cf + w_fms if (w_op + w_cf + w_fms) > 0 else 1.0
master["combined_score"] = (
    master.get("operational_upside_score", pd.Series(np.nan, index=master.index)).fillna(0) * (w_op / total_w)
    + master.get("cash_flow_resilience_score", pd.Series(np.nan, index=master.index)).fillna(0) * (w_cf / total_w)
    + master.get("fms_score", pd.Series(np.nan, index=master.index)).fillna(0) * (w_fms / total_w)
)

# ── Strategy → score column ───────────────────────────────────────
STRATEGY_MAP = {
    "Operational Improvement":           "operational_upside_score",
    "Cash Flow Resilience / Deleverage": "cash_flow_resilience_score",
    "FMS / Buy & Build (Add-on)":        "fms_score",
    "Combined Score":                    "combined_score",
}
selected_score_col = STRATEGY_MAP[strategy]

# FMS uses P5 universe directly
if strategy == "FMS / Buy & Build (Add-on)" and p5 is not None:
    ranking_df = p5.copy().rename(columns={"company_p5": "company", "province_p5": "province"})
    ranking_df["fms_score"] = ranking_df["fms_score"]
    ranking_df = ranking_df.sort_values("fms_score", ascending=False).head(top_n).reset_index(drop=True)
    ranking_df.insert(0, "rank", range(1, len(ranking_df) + 1))
    use_p5_only = True
else:
    use_p5_only = False
    if selected_score_col not in master.columns or master[selected_score_col].notna().sum() == 0:
        st.error(f"⚠️ Score `{selected_score_col}` non disponibile. Carica il file mancante o scegli un'altra strategia.")
        st.stop()
    ranking_df = (
        master.dropna(subset=[selected_score_col])
        .sort_values(selected_score_col, ascending=False)
        .head(top_n)
        .copy()
        .reset_index(drop=True)
    )
    ranking_df.insert(0, "rank", range(1, len(ranking_df) + 1))

# ══════════════════════════════════════════════════════════════════
# KPI CARDS
# ══════════════════════════════════════════════════════════════════

col1, col2, col3, col4 = st.columns(4)

n_merged   = master["mark"].nunique() if "mark" in master.columns else len(master)
n_op       = master["operational_upside_score"].notna().sum() if "operational_upside_score" in master.columns else 0
n_cf       = master["cash_flow_resilience_score"].notna().sum() if "cash_flow_resilience_score" in master.columns else 0
n_fms      = (p5.shape[0] if p5 is not None else 0)

with col1:
    st.metric("🏢 Aziende totali", n_merged)
with col2:
    avg_op = master["operational_upside_score"].mean() if "operational_upside_score" in master.columns else 0
    st.metric("⚙️ Avg Operational Score", f"{avg_op:.1f}" if not np.isnan(avg_op) else "—")
with col3:
    avg_cf = master["cash_flow_resilience_score"].mean() if "cash_flow_resilience_score" in master.columns else 0
    st.metric("💧 Avg CF Resilience", f"{avg_cf:.1f}" if not np.isnan(avg_cf) else "—")
with col4:
    avg_fms = p5["fms_score"].mean() if p5 is not None and "fms_score" in p5.columns else np.nan
    st.metric("🎯 Avg FMS Score", f"{avg_fms:.1f}" if not np.isnan(avg_fms) else "—")

# ══════════════════════════════════════════════════════════════════
# MAIN TABS
# ══════════════════════════════════════════════════════════════════

tab1, tab2, tab3, tab4 = st.tabs([
    "🏆  Top Ranking",
    "📊  Score Explorer",
    "🗺️  Analisi Geografica",
    "🔍  Dettaglio Azienda",
])

# ─────────────────────────────────────────────────────────────────
# TAB 1 — TOP RANKING
# ─────────────────────────────────────────────────────────────────
with tab1:
    strat_color = {
        "Operational Improvement":           "#00C8FF",
        "Cash Flow Resilience / Deleverage": "#B8FF47",
        "FMS / Buy & Build (Add-on)":        "#FF6B35",
        "Combined Score":                    "#A78BFA",
    }[strategy]

    st.markdown(
        f"<div class='section-header'>Top {top_n} · "
        f"<span style='color:{strat_color}'>{strategy}</span></div>",
        unsafe_allow_html=True,
    )

    # ── Horizontal bar chart ──────────────────────────────────────
    if use_p5_only:
        chart_df = ranking_df[["rank", "company", "fms_score"]].copy()
        chart_df["label"] = chart_df["rank"].astype(str) + ". " + chart_df["company"].astype(str)
        fig_bar = px.bar(
            chart_df.sort_values("fms_score"),
            x="fms_score", y="label",
            orientation="h",
            color_discrete_sequence=[strat_color],
            text=chart_df.sort_values("fms_score")["fms_score"].apply(lambda v: f"{v:.1f}"),
        )
    else:
        score_cols_avail = [
            c for c in ["operational_upside_score", "cash_flow_resilience_score", "fms_score", "combined_score"]
            if c in ranking_df.columns and ranking_df[c].notna().sum() > 0
        ]
        ranking_df["label"] = (
            ranking_df["rank"].astype(str) + ". " + ranking_df["company"].astype(str).str[:35]
        )
        melt_df = ranking_df[["rank", "label"] + score_cols_avail].melt(
            id_vars=["rank", "label"],
            value_vars=score_cols_avail,
            var_name="score_type",
            value_name="score",
        )
        melt_df["score_type"] = melt_df["score_type"].map(SCORE_LABELS).fillna(melt_df["score_type"])
        color_seq = [COLOR_MAP.get(c, "#A78BFA") for c in score_cols_avail]

        fig_bar = px.bar(
            melt_df.sort_values(["rank"], ascending=False),
            x="score", y="label",
            color="score_type",
            orientation="h",
            barmode="group",
            color_discrete_sequence=color_seq,
            text=melt_df.sort_values(["rank"], ascending=False)["score"].apply(
                lambda v: f"{v:.1f}" if not pd.isna(v) else ""
            ),
        )

    fig_bar.update_traces(textposition="outside", textfont_size=10)
    fig_bar.update_layout(
        **PLOTLY_TEMPLATE["layout"],
        height=420,
        margin=dict(l=0, r=40, t=10, b=10),
        xaxis_range=[0, 110],
        xaxis_title="Score",
        yaxis_title="",
        legend=dict(orientation="h", y=1.05, x=0),
        showlegend=(not use_p5_only),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # ── Ranking table ─────────────────────────────────────────────
    st.markdown("<div class='section-header'>Tabella Ranking</div>", unsafe_allow_html=True)

    if use_p5_only:
        display_cols = ["rank", "company", "province", "fms_score", "tier_p5",
                        "addon_targets", "ebitda_margin_p5", "fatturato_m"]
        display_cols = [c for c in display_cols if c in ranking_df.columns]
        show_df = ranking_df[display_cols].copy()
        show_df.columns = [c.replace("_p5", "").replace("_", " ").title() for c in show_df.columns]
    else:
        base_cols = ["rank", "company", "province"]
        score_extra = []
        if "operational_upside_score" in ranking_df.columns:
            score_extra.append("operational_upside_score")
        if "cash_flow_resilience_score" in ranking_df.columns:
            score_extra.append("cash_flow_resilience_score")
        if "fms_score" in ranking_df.columns:
            score_extra.append("fms_score")
        if strategy == "Combined Score":
            score_extra.append("combined_score")

        detail_cols = []
        if strategy == "Operational Improvement":
            detail_cols = [c for c in ["operational_driver", "upside_assoluto_pp",
                                       "recoverability_score_base", "recoverability_score_historical",
                                       "recoverability_score", "distress_penalty"]
                           if c in ranking_df.columns]
        elif strategy == "Cash Flow Resilience / Deleverage":
            detail_cols = [c for c in ["resilience_class", "probability_of_default",
                                       "main_value_driver", "latest_revenue", "latest_ebitda"]
                           if c in ranking_df.columns]

        display_cols = base_cols + score_extra + detail_cols
        show_df = ranking_df[display_cols].copy()

        fmt_map = {
            "operational_upside_score":   lambda v: score_fmt(v),
            "cash_flow_resilience_score": lambda v: score_fmt(v),
            "fms_score":                  lambda v: score_fmt(v),
            "combined_score":             lambda v: score_fmt(v),
            "probability_of_default":     lambda v: pct_fmt(v),
            "latest_revenue":             lambda v: eur_fmt(v),
            "latest_ebitda":              lambda v: eur_fmt(v),
            "ebitda_margin_2024":         lambda v: f"{v:.1f}%" if not pd.isna(v) else "—",
            "gap_vs_peer":                lambda v: f"{v:.1f} pp" if not pd.isna(v) else "—",
            "upside_assoluto_pp":         lambda v: f"{v:.1f} pp" if not pd.isna(v) else "—",
        }
        for col, fn in fmt_map.items():
            if col in show_df.columns:
                show_df[col] = show_df[col].apply(fn)

        rename_map = {
            "rank": "#", "company": "Azienda", "province": "Provincia",
            "operational_upside_score": "Op. Score",
            "cash_flow_resilience_score": "CF Score",
            "fms_score": "FMS Score",
            "combined_score": "Combined",
            "operational_driver": "Driver",
            "upside_assoluto_pp": "Upside pp",
            "recoverability_score_base": "Recover. B",
            "recoverability_score_historical": "Recover. H.",
            "recoverability_score": "Recover.",
            "distress_penalty": "Distress",
            "resilience_class": "Classe",
            "probability_of_default": "PD",
            "main_value_driver": "Driver",
            "latest_revenue": "Ricavi",
            "latest_ebitda": "EBITDA",
        }
        show_df.rename(columns={k: v for k, v in rename_map.items() if k in show_df.columns}, inplace=True)

    st.dataframe(show_df, use_container_width=True, hide_index=True)

    # ── Export ────────────────────────────────────────────────────
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        _safe_sheet = strategy[:20].replace("/","-").replace("\\","-").replace("*","").replace("?","").replace("[","").replace("]","").replace(":",""); ranking_df.to_excel(writer, sheet_name=f"Top{top_n}_{_safe_sheet}", index=False)
        master.to_excel(writer, sheet_name="Integrated Scores", index=False)
    buffer.seek(0)

    st.download_button(
        label="⬇️  Esporta Top Ranking (.xlsx)",
        data=buffer,
        file_name=f"fase4_top{top_n}_{strategy.lower().replace(' ','_')[:30]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─────────────────────────────────────────────────────────────────
# TAB 2 — SCORE EXPLORER
# ─────────────────────────────────────────────────────────────────
with tab2:
    st.markdown("<div class='section-header'>Distribuzione Score per Strategia</div>",
                unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    # Distribution histograms
    with col_a:
        if "operational_upside_score" in master.columns:
            fig_h1 = px.histogram(
                master.dropna(subset=["operational_upside_score"]),
                x="operational_upside_score", nbins=30,
                color_discrete_sequence=["#00C8FF"],
                title="Operational Upside Score",
            )
            fig_h1.update_layout(**PLOTLY_TEMPLATE["layout"], height=280,
                                  margin=dict(l=0,r=0,t=36,b=0), showlegend=False)
            st.plotly_chart(fig_h1, use_container_width=True)

    with col_b:
        if "cash_flow_resilience_score" in master.columns:
            fig_h2 = px.histogram(
                master.dropna(subset=["cash_flow_resilience_score"]),
                x="cash_flow_resilience_score", nbins=30,
                color_discrete_sequence=["#B8FF47"],
                title="Cash Flow Resilience Score",
            )
            fig_h2.update_layout(**PLOTLY_TEMPLATE["layout"], height=280,
                                  margin=dict(l=0,r=0,t=36,b=0), showlegend=False)
            st.plotly_chart(fig_h2, use_container_width=True)

    # Scatter OP vs CF
    st.markdown("<div class='section-header'>Scatter — Operational vs Cash Flow Resilience</div>",
                unsafe_allow_html=True)

    scatter_df = master.dropna(subset=["operational_upside_score", "cash_flow_resilience_score"]).copy()
    if not scatter_df.empty:
        scatter_df["company_short"] = scatter_df["company"].astype(str).str[:25]
        color_col = "province" if "province" in scatter_df.columns else None

        fig_sc = px.scatter(
            scatter_df,
            x="operational_upside_score",
            y="cash_flow_resilience_score",
            hover_name="company_short",
            color=color_col,
            size_max=12,
            opacity=0.75,
            labels={
                "operational_upside_score": "Operational Upside Score",
                "cash_flow_resilience_score": "Cash Flow Resilience Score",
            },
        )
        # Quadrant lines
        med_op = scatter_df["operational_upside_score"].median()
        med_cf = scatter_df["cash_flow_resilience_score"].median()
        for val, axis in [(med_op, "x"), (med_cf, "y")]:
            fig_sc.add_shape(
                type="line",
                **({f"{axis}0": val, f"{axis}1": val,
                    "y0" if axis == "x" else "x0": 0,
                    "y1" if axis == "x" else "x1": 110}),
                line=dict(color="#1E2D40", dash="dash", width=1),
            )
        # Quadrant labels
        annotations = [
            dict(x=5, y=105, text="Low Op · High CF", showarrow=False,
                 font=dict(size=9, color="#3D4F63")),
            dict(x=med_op + 2, y=105, text="⭐ High Op · High CF", showarrow=False,
                 font=dict(size=9, color="#B8FF47")),
            dict(x=5, y=2, text="Low Op · Low CF", showarrow=False,
                 font=dict(size=9, color="#3D4F63")),
            dict(x=med_op + 2, y=2, text="High Op · Low CF", showarrow=False,
                 font=dict(size=9, color="#3D4F63")),
        ]
        fig_sc.update_layout(
            **PLOTLY_TEMPLATE["layout"],
            height=480,
            margin=dict(l=0, r=0, t=10, b=0),
            annotations=annotations,
        )
        st.plotly_chart(fig_sc, use_container_width=True)

    # FMS distribution if available
    if p5 is not None and "fms_score" in p5.columns:
        st.markdown("<div class='section-header'>FMS Score — Distribuzione per Tier</div>",
                    unsafe_allow_html=True)
        fig_fms = px.box(
            p5, x="tier_p5", y="fms_score",
            color="tier_p5",
            color_discrete_sequence=["#FF6B35", "#F59E0B", "#6B7B92"],
            labels={"tier_p5": "Tier", "fms_score": "FMS Score"},
        )
        fig_fms.update_layout(**PLOTLY_TEMPLATE["layout"], height=320,
                               margin=dict(l=0,r=0,t=10,b=0), showlegend=False)
        st.plotly_chart(fig_fms, use_container_width=True)

    # Feature importance from P4
    st.markdown("<div class='section-header'>Feature Importance — Modello CF Resilience (P4)</div>",
                unsafe_allow_html=True)
    try:
        fi_df = pd.read_excel(p4_src, sheet_name="Feature Importance")
        fi_df = fi_df.sort_values("Importance")
        fig_fi = px.bar(
            fi_df, x="Importance", y="Feature",
            orientation="h",
            color="Importance",
            color_continuous_scale=[[0, "#1E2D40"], [1, "#B8FF47"]],
            text=fi_df["Importance"].apply(lambda v: f"{v:.3f}"),
        )
        fig_fi.update_traces(textposition="outside", textfont_size=10)
        fig_fi.update_layout(
            **PLOTLY_TEMPLATE["layout"],
            height=360,
            margin=dict(l=0, r=60, t=10, b=0),
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig_fi, use_container_width=True)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────
# TAB 3 — ANALISI GEOGRAFICA
# ─────────────────────────────────────────────────────────────────
with tab3:
    st.markdown("<div class='section-header'>Score Medio per Provincia</div>",
                unsafe_allow_html=True)

    geo_df = master.copy()
    if "province" not in geo_df.columns or geo_df["province"].isna().all():
        st.info("Dati geografici non disponibili.")
    else:
        agg_dict = {}
        for col in ["operational_upside_score", "cash_flow_resilience_score", "fms_score"]:
            if col in geo_df.columns:
                agg_dict[col] = "mean"
        agg_dict["mark"] = "count"
        rename_agg = {**{k: k for k in agg_dict if k != "mark"}, "mark": "n_aziende"}

        prov_df = (
            geo_df.groupby("province", as_index=False)
            .agg(agg_dict)
            .rename(columns=rename_agg)
            .sort_values("n_aziende", ascending=False)
        )

        col_geo1, col_geo2 = st.columns([2, 1])

        with col_geo1:
            score_to_plot = selected_score_col if selected_score_col in prov_df.columns else \
                next((c for c in ["operational_upside_score","cash_flow_resilience_score"] if c in prov_df.columns), None)
            if score_to_plot:
                top_prov = prov_df.nlargest(20, score_to_plot)
                fig_geo = px.bar(
                    top_prov.sort_values(score_to_plot),
                    x=score_to_plot, y="province",
                    orientation="h",
                    color=score_to_plot,
                    color_continuous_scale=[[0,"#1E2D40"],[1,strat_color]],
                    text=top_prov.sort_values(score_to_plot)[score_to_plot].apply(
                        lambda v: f"{v:.1f}" if not pd.isna(v) else ""
                    ),
                    labels={"province": "Provincia", score_to_plot: SCORE_LABELS.get(score_to_plot, score_to_plot)},
                    title=f"Top 20 Province — {SCORE_LABELS.get(score_to_plot, score_to_plot)}",
                )
                fig_geo.update_traces(textposition="outside", textfont_size=9)
                fig_geo.update_layout(
                    **PLOTLY_TEMPLATE["layout"],
                    height=520,
                    margin=dict(l=0, r=60, t=40, b=0),
                    coloraxis_showscale=False,
                )
                st.plotly_chart(fig_geo, use_container_width=True)

        with col_geo2:
            # Count by province
            fig_cnt = px.bar(
                prov_df.head(15).sort_values("n_aziende"),
                x="n_aziende", y="province",
                orientation="h",
                color_discrete_sequence=["#1E2D40"],
                labels={"n_aziende": "N. Aziende", "province": ""},
                title="N. Aziende per Provincia",
                text="n_aziende",
            )
            fig_cnt.update_traces(marker_color="#1E2D40",
                                  marker_line_color="#00C8FF", marker_line_width=1,
                                  textposition="outside", textfont_size=10)
            fig_cnt.update_layout(
                **PLOTLY_TEMPLATE["layout"],
                height=520,
                margin=dict(l=0, r=40, t=40, b=0),
            )
            st.plotly_chart(fig_cnt, use_container_width=True)

        # Radar chart by macro area (if P5 available)
        if p5 is not None and "macro_area" in p5.columns and "fms_score" in p5.columns:
            st.markdown("<div class='section-header'>FMS Score medio per Macro Area</div>",
                        unsafe_allow_html=True)
            ma_df = p5.groupby("macro_area")["fms_score"].mean().reset_index()
            fig_radar = go.Figure(go.Scatterpolar(
                r=ma_df["fms_score"].tolist() + [ma_df["fms_score"].iloc[0]],
                theta=ma_df["macro_area"].tolist() + [ma_df["macro_area"].iloc[0]],
                fill="toself",
                line_color="#FF6B35",
                fillcolor="rgba(255,107,53,0.15)",
            ))
            fig_radar.update_layout(
                **PLOTLY_TEMPLATE["layout"],
                height=380,
                polar=dict(
                    radialaxis=dict(visible=True, gridcolor="#1E2D40", range=[0,100]),
                    angularaxis=dict(gridcolor="#1E2D40"),
                    bgcolor="rgba(0,0,0,0)",
                ),
                margin=dict(l=40,r=40,t=40,b=40),
            )
            st.plotly_chart(fig_radar, use_container_width=True)

# ─────────────────────────────────────────────────────────────────
# TAB 4 — DETTAGLIO AZIENDA
# ─────────────────────────────────────────────────────────────────
with tab4:
    st.markdown("<div class='section-header'>Scorecard Azienda</div>",
                unsafe_allow_html=True)

    all_companies = sorted(
        master["company"].dropna().astype(str).unique().tolist()
    ) if "company" in master.columns else []

    if not all_companies:
        st.info("Nessun dato azienda disponibile.")
    else:
        selected_company = st.selectbox("Seleziona azienda", all_companies)

        row = master[master["company"].astype(str) == selected_company]
        if row.empty:
            st.warning("Azienda non trovata nel dataset integrato.")
        else:
            row = row.iloc[0]

            # ── Header scorecard ──────────────────────────────────
            op_score = row.get("operational_upside_score", np.nan)
            cf_score = row.get("cash_flow_resilience_score", np.nan)
            fms_sc   = row.get("fms_score", np.nan)

            st.markdown(f"""
            <div style='background:#111827;border:1px solid #1E2D40;border-radius:10px;
                        padding:24px 28px;margin-bottom:20px;'>
                <div style='font-family:Syne,sans-serif;font-size:1.4rem;font-weight:700;
                            margin-bottom:4px;'>{selected_company}</div>
                <div style='font-size:0.7rem;color:#6B7B92;'>{row.get("province","—")}</div>
                <div style='display:flex;gap:24px;margin-top:20px;flex-wrap:wrap;'>
                    <div>
                        <div style='font-size:0.6rem;letter-spacing:2px;text-transform:uppercase;
                                    color:#6B7B92;margin-bottom:4px;'>Operational Score</div>
                        <div style='font-family:Syne,sans-serif;font-size:1.8rem;font-weight:700;
                                    color:#00C8FF;'>{score_fmt(op_score)}</div>
                    </div>
                    <div>
                        <div style='font-size:0.6rem;letter-spacing:2px;text-transform:uppercase;
                                    color:#6B7B92;margin-bottom:4px;'>CF Resilience</div>
                        <div style='font-family:Syne,sans-serif;font-size:1.8rem;font-weight:700;
                                    color:#B8FF47;'>{score_fmt(cf_score)}</div>
                    </div>
                    <div>
                        <div style='font-size:0.6rem;letter-spacing:2px;text-transform:uppercase;
                                    color:#6B7B92;margin-bottom:4px;'>FMS Score</div>
                        <div style='font-family:Syne,sans-serif;font-size:1.8rem;font-weight:700;
                                    color:#FF6B35;'>{score_fmt(fms_sc)}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── Spider / Radar ────────────────────────────────────
            radar_vals = {}
            if not pd.isna(op_score):
                radar_vals["Op. Upside"] = min(op_score, 100)
            if not pd.isna(cf_score):
                radar_vals["CF Resilience"] = cf_score
            if not pd.isna(fms_sc):
                radar_vals["FMS Score"] = fms_sc

            # sub-scores for P3
            for col, label in [
                ("recoverability_score_base", "Recoverability"),
                ("ebitda_margin_2024", "EBITDA Margin"),
            ]:
                v = row.get(col, np.nan)
                if not pd.isna(v):
                    radar_vals[label] = min(abs(float(v)) * 3, 100)

            # sub-scores for P4
            for col, label in [
                ("fcf_stability_score", "FCF Stability"),
                ("dscr_score", "DSCR"),
            ]:
                v = row.get(col, np.nan)
                if not pd.isna(v):
                    radar_vals[label] = v

            if len(radar_vals) >= 3:
                cats = list(radar_vals.keys())
                vals = list(radar_vals.values())
                fig_spider = go.Figure(go.Scatterpolar(
                    r=vals + [vals[0]],
                    theta=cats + [cats[0]],
                    fill="toself",
                    line_color="#00C8FF",
                    fillcolor="rgba(0,200,255,0.10)",
                    marker=dict(color="#00C8FF", size=6),
                ))
                fig_spider.update_layout(
                    **PLOTLY_TEMPLATE["layout"],
                    height=380,
                    polar=dict(
                        radialaxis=dict(visible=True, gridcolor="#1E2D40", range=[0, 100]),
                        angularaxis=dict(gridcolor="#1E2D40"),
                        bgcolor="rgba(0,0,0,0)",
                    ),
                    margin=dict(l=40,r=40,t=20,b=20),
                )
                st.plotly_chart(fig_spider, use_container_width=True)

            # ── Investment Thesis ─────────────────────────────────
            thesis = row.get("investment_thesis", None)
            if thesis and str(thesis) not in ("nan", "—", ""):
                st.markdown("""
                <div style='background:#111827;border-left:3px solid #00C8FF;
                            padding:16px 20px;border-radius:0 8px 8px 0;margin-top:8px;'>
                    <div style='font-size:0.6rem;letter-spacing:2px;text-transform:uppercase;
                                color:#6B7B92;margin-bottom:8px;'>Investment Thesis (P4)</div>
                    <div style='font-size:0.8rem;line-height:1.7;color:#C8D5E8;'>
                """ + str(thesis) + """
                    </div>
                </div>
                """, unsafe_allow_html=True)

            # ── Raw data ──────────────────────────────────────────
            with st.expander("📋  Tutti i dati disponibili"):
                row_df = row.to_frame(name="Valore").reset_index()
                row_df.columns = ["Campo", "Valore"]
                row_df = row_df[row_df["Valore"].notna()]
                st.dataframe(row_df, use_container_width=True, hide_index=True)

