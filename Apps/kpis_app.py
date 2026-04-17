"""
KPIs Q2 2026 — PISO
Diseño: Dark theme inspirado en pisokpis.netlify.app
"""

import streamlit as st
import pandas as pd
import hashlib, json, base64, requests
from pathlib import Path

st.set_page_config(
    page_title="KPIs Q2 2026 — PISO",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR   = Path(__file__).parent.resolve()
EXCEL_PATH = BASE_DIR / "kpis_q2.xlsx"

# ─── GitHub storage ───────────────────────────────────────────────────────────
_GH_REPO = "svilla-sketch/reportes-bi-marketing"
_GH_PATH = "Apps/kpis_scores_q2_data.json"
_GH_API  = f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_PATH}"

def _gh_headers():
    return {"Authorization": f"token {st.secrets['github']['token']}",
            "Accept": "application/vnd.github.v3+json"}

def gh_load():
    try:
        r = requests.get(_GH_API, headers=_gh_headers(), timeout=6)
        if r.status_code == 200:
            return json.loads(base64.b64decode(r.json()["content"]).decode())
    except Exception: pass
    return {}

def gh_save(scores):
    try:
        content_b64 = base64.b64encode(
            json.dumps(scores, ensure_ascii=False, indent=2).encode()).decode()
        r = requests.get(_GH_API, headers=_gh_headers(), timeout=6)
        sha = r.json().get("sha") if r.status_code == 200 else None
        payload = {"message": "Update KPIs Q2 scores",
                   "content": content_b64,
                   "committer": {"name": "PISO KPIs", "email": "kpis@piso.app"}}
        if sha: payload["sha"] = sha
        requests.put(_GH_API, headers=_gh_headers(), json=payload, timeout=10)
    except Exception: pass

# ─── Constantes ───────────────────────────────────────────────────────────────
PROJECT_META = {
    "Kos":       "7 Escrituraciones",
    "P.C.":      "Construcción del 66% + 100% comercialización",
    "Santian":   "Inicio comercialización, 2 ventas",
    "Zen":       "Autorización CUS + inicio construcción",
    "Mar Negro": "Preparativo validado para inicio",
    "Colón":     "Compra terreno ante Notario",
    "Wacuz":     "Adquisición terreno e inicio preparativo N26",
    "RRO":       "Inicio Construcción Edificio Romero",
    "Admin":     "100% funciones Admin en ERP",
}

PROJECT_COLORS = {
    "Kos": "#3b82f6", "P.C.": "#ec4899", "Santian": "#10b981",
    "Zen": "#8b5cf6", "Mar Negro": "#06b6d4", "Colón": "#f97316",
    "Wacuz": "#14b8a6", "RRO": "#f59e0b", "Admin": "#64748b",
}

AVATAR_COLORS = [
    "#6366f1","#8b5cf6","#06b6d4","#10b981","#f59e0b",
    "#ef4444","#ec4899","#f97316","#14b8a6","#3b82f6",
    "#a855f7","#22d3ee","#84cc16","#fb923c","#e879f9",
    "#34d399","#fbbf24","#60a5fa",
]

INITIALS_MAP = {
    "Armando Pérez":            "ARM",
    "Rodrigo Rico":             "ROD",
    "Miguel Elizalde":          "MIG",
    "Joselin Vega":             "JSL",
    "Juan Ramón Moreno Flores": "JRM",
    "Mariana Irigoyen":         "MIR",
    "Mariana Méndez":           "MIA",
    "Saúl Villa":               "SVL",
    "Sergio Oliva":             "SRG",
    "Víctor Villa Walls":       "VVW",
    "Ángel Vázquez":            "ANG",
    "Óscar Valencia":           "OSC",
    "Areli Carrasco":           "ARC",
    "Adriana Carrasco":         "ADC",
    "Brenda Quintero":          "BRD",
    "Brianda Ramírez":          "BRI",
    "Ana Serrano":              "ANS",
    "Jorge Guerra":             "JRG",
}

SOCIOS = ["Juan Ramón Moreno Flores", "Víctor Villa Walls", "Saúl Villa"]

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* ── Reset Streamlit ── */
html, body, [data-testid="stAppViewContainer"], .stApp {
    background-color: #0f172a !important;
    color: #e2e8f0 !important;
    font-family: 'Inter', sans-serif !important;
}
[data-testid="stHeader"] { background: #0f172a !important; border-bottom: 1px solid #1e293b; }
[data-testid="stSidebar"] { background: #1e293b !important; }
section[data-testid="stMain"] > div { padding-top: 1rem !important; }
.block-container { padding: 1.5rem 2rem !important; max-width: 1400px !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: #1e293b !important;
    border-radius: 12px;
    padding: 4px;
    gap: 2px;
    border: 1px solid #334155;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #94a3b8 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    font-size: 0.85rem !important;
    padding: 6px 16px !important;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    background: #6366f1 !important;
    color: #fff !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.5rem !important; }
.stTabs [data-baseweb="tab-highlight"] { display: none !important; }

/* ── Cards ── */
.kpi-card {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 16px;
    padding: 1.25rem;
    margin-bottom: 1rem;
}
.stat-card {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 1.25rem 1.5rem;
    text-align: center;
}
.stat-card .stat-icon { font-size: 1.8rem; margin-bottom: 0.25rem; }
.stat-card .stat-value { font-size: 2rem; font-weight: 700; margin: 4px 0; }
.stat-card .stat-label { color: #94a3b8; font-size: 0.78rem; font-weight: 500; text-transform: uppercase; letter-spacing: .05em; }
.stat-card .stat-sub   { color: #64748b; font-size: 0.72rem; margin-top: 2px; }

/* ── Avatar ── */
.avatar {
    width: 52px; height: 52px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.9rem; color: #fff;
    flex-shrink: 0; letter-spacing: 0.05em;
    border: 2px solid rgba(255,255,255,0.15);
}
.avatar-sm {
    width: 28px; height: 28px; border-radius: 50%;
    display: inline-flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.6rem; color: #fff; flex-shrink: 0;
}

/* ── Person header ── */
.person-header {
    display: flex; align-items: flex-start; gap: 14px; margin-bottom: 1.2rem;
}
.person-info { flex: 1; }
.person-name { font-size: 1.1rem; font-weight: 700; color: #f1f5f9; }
.person-role { font-size: 0.78rem; color: #94a3b8; margin: 2px 0 6px; }
.person-score-big { font-size: 2.2rem; font-weight: 800; min-width: 90px; text-align: right; }

/* ── Badges ── */
.badge {
    display: inline-block; padding: 2px 8px; border-radius: 20px;
    font-size: 0.68rem; font-weight: 600; margin: 2px;
    border: 1px solid rgba(255,255,255,0.1);
}
.badge-equipo   { background: rgba(99,102,241,0.2); color: #818cf8; border-color: rgba(99,102,241,0.3); }
.badge-ind      { background: rgba(16,185,129,0.2); color: #34d399; border-color: rgba(16,185,129,0.3); }
.badge-socio    { background: rgba(245,158,11,0.2); color: #fbbf24; border-color: rgba(245,158,11,0.3); }

/* ── KPI rows ── */
.kpi-section-title {
    font-size: 0.68rem; font-weight: 700; letter-spacing: .1em;
    color: #64748b; text-transform: uppercase;
    border-bottom: 1px solid #273549;
    padding-bottom: 6px; margin-bottom: 10px; margin-top: 14px;
}
.kpi-row {
    padding: 8px 0;
    border-bottom: 1px solid #1a2540;
}
.kpi-row:last-child { border-bottom: none; }
.kpi-name { font-size: 0.85rem; font-weight: 500; color: #e2e8f0; margin-bottom: 3px; }
.kpi-meta { font-size: 0.7rem; color: #64748b; margin-bottom: 5px; }
.kpi-footer { display: flex; align-items: center; gap: 8px; }
.kpi-pct { font-size: 0.85rem; font-weight: 700; min-width: 42px; }
.kpi-bar-wrap {
    flex: 1; height: 5px; background: #273549; border-radius: 3px; overflow: hidden;
}
.kpi-bar-fill { height: 100%; border-radius: 3px; transition: width .4s; }
.kpi-status {
    font-size: 0.65rem; font-weight: 600; padding: 2px 7px;
    border-radius: 20px; white-space: nowrap;
}
.status-done   { background: rgba(16,185,129,0.15); color: #10b981; }
.status-pend   { background: rgba(100,116,139,0.15); color: #64748b; }
.status-prog   { background: rgba(245,158,11,0.15);  color: #f59e0b; }

/* ── Score boxes ── */
.score-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 10px; }
.score-box {
    background: #273549; border-radius: 10px;
    padding: 12px; border: 1px solid #334155;
}
.score-box-label { font-size: 0.7rem; color: #64748b; font-weight: 600; text-transform: uppercase; margin-bottom: 4px; }
.score-box-value { font-size: 1.5rem; font-weight: 800; margin-bottom: 4px; }
.score-box-bar { height: 4px; background: #334155; border-radius: 2px; overflow:hidden; margin-bottom: 4px; }
.score-box-sub { font-size: 0.65rem; color: #64748b; }

.total-box {
    background: linear-gradient(135deg, #1e293b 0%, #273549 100%);
    border: 1px solid #334155; border-radius: 12px;
    padding: 16px; text-align: center; margin-top: 10px;
}
.total-label { font-size: 0.7rem; color: #64748b; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 4px; }
.total-value { font-size: 2.8rem; font-weight: 800; line-height: 1; margin: 8px 0; }
.financials { display: flex; justify-content: space-around; margin-top: 12px; padding-top: 12px; border-top: 1px solid #334155; }
.fin-item { text-align: center; }
.fin-label { font-size: 0.65rem; color: #64748b; }
.fin-value { font-size: 0.9rem; font-weight: 700; color: #e2e8f0; }

/* ── Empresa table ── */
.obj-row {
    display: flex; align-items: center; gap: 12px;
    padding: 10px 0; border-bottom: 1px solid #1a2540;
}
.obj-row:last-child { border-bottom: none; }
.obj-proj-badge {
    padding: 3px 8px; border-radius: 6px; font-size: 0.7rem;
    font-weight: 700; min-width: 72px; text-align: center; flex-shrink: 0;
}
.obj-name { flex: 1; font-size: 0.85rem; color: #e2e8f0; }
.obj-pct { font-size: 0.9rem; font-weight: 700; min-width: 52px; text-align: right; }
.obj-bar-wrap { width: 120px; height: 6px; background: #273549; border-radius: 3px; overflow: hidden; flex-shrink: 0; }
.obj-bar-fill { height: 100%; border-radius: 3px; }

/* ── Rankings ── */
.rank-row {
    display: flex; align-items: center; gap: 10px;
    padding: 6px 0; border-bottom: 1px solid #1a2540;
}
.rank-num { color: #475569; font-size: 0.75rem; min-width: 24px; }
.rank-bar-wrap { flex: 1; height: 28px; background: #273549; border-radius: 6px; overflow: hidden; position: relative; }
.rank-bar-fill { height: 100%; border-radius: 6px; display: flex; align-items: center; padding-left: 10px; font-size: 0.78rem; font-weight: 600; color: #fff; }
.rank-val { min-width: 50px; text-align: right; font-size: 0.85rem; font-weight: 700; }

/* ── Inputs oscuros ── */
.stNumberInput input, .stTextInput input, .stSelectbox select {
    background: #273549 !important; color: #e2e8f0 !important;
    border: 1px solid #334155 !important; border-radius: 8px !important;
}
.stSlider [data-baseweb="slider"] { padding: 0 !important; }
div[data-testid="stMetricValue"] { color: #e2e8f0 !important; }

/* ── Selectbox ── */
div[data-baseweb="select"] > div {
    background: #273549 !important; border-color: #334155 !important; border-radius: 8px !important;
}
div[data-baseweb="select"] span { color: #e2e8f0 !important; }
[data-baseweb="popover"] { background: #1e293b !important; border: 1px solid #334155 !important; }
[role="option"] { background: #1e293b !important; color: #e2e8f0 !important; }
[role="option"]:hover { background: #273549 !important; }

/* ── Pill selector ── */
.pill-wrap { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 1.2rem; }
.pill {
    padding: 5px 12px; border-radius: 20px; font-size: 0.78rem; font-weight: 600;
    cursor: pointer; border: 1px solid #334155;
    background: #1e293b; color: #94a3b8;
    display: inline-flex; align-items: center; gap: 6px;
}
.pill.active { background: #6366f1; color: #fff; border-color: #6366f1; }

/* ── Misc ── */
h1, h2, h3 { color: #f1f5f9 !important; }
.stButton > button {
    background: #6366f1 !important; color: #fff !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important;
}
.stButton > button:hover { background: #4f46e5 !important; }
hr { border-color: #334155 !important; }
[data-testid="stExpander"] {
    background: #1e293b !important; border: 1px solid #334155 !important; border-radius: 10px !important;
}
.streamlit-expanderHeader { color: #e2e8f0 !important; font-weight: 600 !important; }
label { color: #94a3b8 !important; font-size: 0.82rem !important; }
</style>
""", unsafe_allow_html=True)


# ─── Data loading ─────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name="KPIs 2T", header=0)
    df.columns = df.columns.str.strip()
    for col in ["Proyecto","Objetivo Empresa","Objetivo Lider","KPI",
                "Responsable 1","Responsable 2","Responsable 3","Lider","Fecha"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan":"","NaN":""})
    def classify(r):
        return "equipo" if (r["Responsable 2"] not in ("","nan") or
                            r["Responsable 3"] not in ("","nan")) else "individual"
    df["tipo"] = df.apply(classify, axis=1)
    def make_uid(r):
        c = "|".join([r.get(x,"") for x in ["Proyecto","KPI","Responsable 1","Responsable 2","Responsable 3"]])
        return hashlib.md5(c.encode()).hexdigest()[:12]
    df["uid"] = df.apply(make_uid, axis=1)
    return df


# ─── Scores init/save ─────────────────────────────────────────────────────────
def init_scores(df):
    remote = gh_load()
    base = {"empresa": {p: 0 for p in PROJECT_META},
            "kpis":    {r["uid"]: 0 for _, r in df.iterrows()},
            "discrecional": {}}
    base["empresa"].update(remote.get("empresa", {}))
    for uid in base["kpis"]:
        base["kpis"][uid] = remote.get("kpis", {}).get(uid, 0)
    base["discrecional"] = remote.get("discrecional", {})
    st.session_state.scores = base

def save_scores():
    gh_save(st.session_state.scores)


# ─── Calculations ─────────────────────────────────────────────────────────────
def s_empresa():
    v = list(st.session_state.scores["empresa"].values())
    return sum(v)/len(v) if v else 0.0

def s_equipo(persona, df):
    m = (((df["Responsable 1"]==persona)|(df["Responsable 2"]==persona)|(df["Responsable 3"]==persona))
         & (df["tipo"]=="equipo"))
    k = df[m]
    if k.empty: return 0.0
    return sum(st.session_state.scores["kpis"].get(r["uid"],0) for _,r in k.iterrows()) / len(k)

def s_individual(persona, df):
    m = (df["Responsable 1"]==persona) & (df["tipo"]=="individual")
    k = df[m]
    if k.empty: return 0.0
    return sum(st.session_state.scores["kpis"].get(r["uid"],0) for _,r in k.iterrows()) / len(k)

def s_disc(persona):
    return float(st.session_state.scores.get("discrecional",{}).get(persona,0)) * 10.0

def s_total(persona, df):
    return s_empresa()*0.30 + s_equipo(persona,df)*0.30 + s_individual(persona,df)*0.30 + s_disc(persona)*0.10


# ─── Helpers UI ───────────────────────────────────────────────────────────────
def score_color(v):
    if v >= 80: return "#10b981"
    if v >= 60: return "#f59e0b"
    return "#ef4444"

def get_initials(name):
    return INITIALS_MAP.get(name, name[:3].upper())

def get_avatar_color(name):
    idx = abs(hash(name)) % len(AVATAR_COLORS)
    return AVATAR_COLORS[idx]

def all_people(df):
    return sorted([p for p in df["Responsable 1"].unique() if p])

def people_by_lider(df, lider):
    return sorted([p for p in df[df["Lider"]==lider]["Responsable 1"].unique() if p])

def project_badge_html(proyecto):
    c = PROJECT_COLORS.get(proyecto, "#64748b")
    return (f'<span class="badge" style="background:rgba({int(c[1:3],16)},{int(c[3:5],16)},'
            f'{int(c[5:7],16)},0.2);color:{c};border-color:rgba({int(c[1:3],16)},'
            f'{int(c[3:5],16)},{int(c[5:7],16)},0.4)">{proyecto}</span>')

def kpi_status_html(pct):
    if pct >= 100: return '<span class="kpi-status status-done">✓ Cumplido</span>'
    if pct >   0:  return '<span class="kpi-status status-prog">↗ En progreso</span>'
    return '<span class="kpi-status status-pend">○ Pendiente</span>'

def bar_html(pct, color=None, height=5):
    c = color or score_color(pct)
    return (f'<div class="kpi-bar-wrap" style="height:{height}px">'
            f'<div class="kpi-bar-fill" style="width:{min(pct,100)}%;background:{c}"></div></div>')


# ─── TOP HEADER ───────────────────────────────────────────────────────────────
def render_header():
    emp = s_empresa()
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown('<div style="display:flex;align-items:center;gap:12px">'
                    '<div style="background:#6366f1;border-radius:10px;padding:8px 12px;font-size:1.4rem">🏢</div>'
                    '<div><div style="font-size:1.15rem;font-weight:700;color:#f1f5f9">KPIs Dashboard</div>'
                    '<div style="font-size:0.75rem;color:#64748b">Medición de Objetivos — Q2 2026</div></div>'
                    '</div>', unsafe_allow_html=True)
    with col2:
        c = score_color(emp)
        st.markdown(f'<div style="text-align:right">'
                    f'<span style="color:#64748b;font-size:0.8rem">Score empresa: </span>'
                    f'<span style="color:{c};font-size:1.1rem;font-weight:800">{emp:.1f}%</span>&nbsp;&nbsp;'
                    f'<span style="background:#6366f1;color:#fff;padding:4px 14px;border-radius:20px;font-size:0.8rem;font-weight:600">Q2 2026</span>'
                    f'</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:1px;background:#334155;margin:12px 0 0"></div>', unsafe_allow_html=True)


# ─── VISTA: RESUMEN ───────────────────────────────────────────────────────────
def view_resumen(df):
    emp = s_empresa()
    people = all_people(df)
    scores_list = [s_total(p, df) for p in people]
    best_person = people[scores_list.index(max(scores_list))] if people else "—"
    avg_score = sum(scores_list)/len(scores_list) if scores_list else 0

    # Stat cards
    max_score = max(scores_list) if scores_list else 0
    n_eq = len(df[df.tipo == "equipo"]); n_ind = len(df[df.tipo == "individual"])
    cards_html = (
        f'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:1.5rem">'
        f'<div class="stat-card"><div class="stat-icon">🏢</div>'
        f'<div class="stat-value" style="color:{score_color(emp)}">{emp:.1f}%</div>'
        f'<div class="stat-label">Score Empresa</div><div class="stat-sub">Promedio ponderado</div></div>'
        f'<div class="stat-card"><div class="stat-icon">📋</div>'
        f'<div class="stat-value" style="color:#6366f1">{len(df)}</div>'
        f'<div class="stat-label">KPIs Total</div><div class="stat-sub">{n_eq} equipo · {n_ind} individual</div></div>'
        f'<div class="stat-card"><div class="stat-icon">🏆</div>'
        f'<div class="stat-value" style="color:{score_color(max_score)}">{max_score:.1f}%</div>'
        f'<div class="stat-label">Mejor Score</div><div class="stat-sub">{best_person.split()[0]}</div></div>'
        f'<div class="stat-card"><div class="stat-icon">📊</div>'
        f'<div class="stat-value" style="color:{score_color(avg_score)}">{avg_score:.1f}%</div>'
        f'<div class="stat-label">Score Promedio</div><div class="stat-sub">{len(people)} personas</div></div>'
        f'</div>'
    )
    st.markdown(cards_html, unsafe_allow_html=True)

    col_obj, col_rank = st.columns([1.1, 1])

    with col_obj:
        rows_html = ""
        for proyecto, meta in PROJECT_META.items():
            pct   = st.session_state.scores["empresa"].get(proyecto, 0)
            c     = PROJECT_COLORS.get(proyecto, "#64748b")
            c_rgb = f"{int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)}"
            rows_html += (
                f'<div class="obj-row">'
                f'<span class="obj-proj-badge" style="background:rgba({c_rgb},0.15);color:{c}">{proyecto}</span>'
                f'<span class="obj-name">{meta}</span>'
                f'<span class="obj-pct" style="color:{score_color(pct)}">{pct}%</span>'
                f'<div class="obj-bar-wrap"><div class="obj-bar-fill" style="width:{pct}%;background:{score_color(pct)}"></div></div>'
                f'</div>'
            )
        st.markdown(f'<div class="kpi-card"><div class="kpi-section-title">Objetivos de la empresa</div>{rows_html}</div>',
                    unsafe_allow_html=True)

    with col_rank:
        sorted_people = sorted(zip(people, scores_list), key=lambda x: x[1], reverse=True)
        rows_html = ""
        for i, (persona, score) in enumerate(sorted_people):
            ini = get_initials(persona)
            ac  = get_avatar_color(persona)
            c   = score_color(score)
            rows_html += (
                f'<div class="rank-row">'
                f'<span class="rank-num">#{i+1}</span>'
                f'<div class="avatar-sm" style="background:{ac}">{ini}</div>'
                f'<div class="rank-bar-wrap">'
                f'<div class="rank-bar-fill" style="width:{score:.1f}%;background:linear-gradient(90deg,{ac}88,{ac})">'
                f'{persona.split()[0]}'
                f'</div></div>'
                f'<span class="rank-val" style="color:{c}">{score:.1f}%</span>'
                f'</div>'
            )
        st.markdown(f'<div class="kpi-card"><div class="kpi-section-title">Ranking de personas</div>{rows_html}</div>',
                    unsafe_allow_html=True)


# ─── VISTA: PERSONAS ──────────────────────────────────────────────────────────
def view_personas(df):
    people = all_people(df)

    # Person selector
    if "sel_persona" not in st.session_state:
        st.session_state.sel_persona = people[0] if people else ""

    pills_html = '<div class="pill-wrap">'
    for p in people:
        ini = get_initials(p)
        ac  = get_avatar_color(p)
        active = "active" if p == st.session_state.sel_persona else ""
        pills_html += (f'<span class="pill {active}">'
                       f'<span style="background:{ac};width:16px;height:16px;border-radius:50%;display:inline-flex;'
                       f'align-items:center;justify-content:center;font-size:0.5rem;font-weight:700">{ini[0]}</span>'
                       f'{p.split()[0]}</span>')
    pills_html += '</div>'

    # Use selectbox (styled) as actual selector
    sel = st.selectbox("Persona", people,
                       index=people.index(st.session_state.sel_persona) if st.session_state.sel_persona in people else 0,
                       key="sel_persona_box", label_visibility="collapsed")
    st.session_state.sel_persona = sel
    persona = sel

    st.markdown(pills_html, unsafe_allow_html=True)

    # Scores
    emp  = s_empresa()
    equ  = s_equipo(persona, df)
    ind  = s_individual(persona, df)
    disc = s_disc(persona)
    tot  = emp*0.30 + equ*0.30 + ind*0.30 + disc*0.10

    # Person projects
    person_projects = list(df[df["Responsable 1"]==persona]["Proyecto"].unique())
    badges = "".join(project_badge_html(p) for p in person_projects)

    # Role
    lider_row = df[df["Responsable 1"]==persona]["Lider"]
    lider = lider_row.iloc[0] if not lider_row.empty else ""
    role = "Socio" if persona in SOCIOS else "Colaborador"

    # Avatar
    ini = get_initials(persona)
    ac  = get_avatar_color(persona)

    col_card, col_salarial = st.columns([1.4, 1])

    with col_card:
        # KPIs
        ind_kpis  = df[(df["Responsable 1"]==persona) & (df["tipo"]=="individual")]
        team_kpis = df[((df["Responsable 1"]==persona)|(df["Responsable 2"]==persona)|(df["Responsable 3"]==persona))
                       & (df["tipo"]=="equipo")]

        def kpi_rows_html(kpis, show_team=False):
            html = ""
            for _, row in kpis.iterrows():
                uid   = row["uid"]
                pct   = st.session_state.scores["kpis"].get(uid, 0)
                c     = score_color(pct)
                badge_cls = "equipo" if row["tipo"] == "equipo" else "ind"
                badge_lbl = "Equipo" if row["tipo"] == "equipo" else "Individual"
                tipo  = f'<span class="badge badge-{badge_cls}">{badge_lbl}</span>'
                rs    = [r for r in [row["Responsable 1"], row["Responsable 2"], row["Responsable 3"]] if r]
                r_str = " · ".join(rs) if show_team and len(rs) > 1 else ""
                rspan = f'<span style="color:#475569;font-size:0.68rem"> · {r_str}</span>' if r_str else ""
                html += (
                    f'<div class="kpi-row">'
                    f'<div style="display:flex;justify-content:space-between;align-items:flex-start">'
                    f'<div class="kpi-name">{row["KPI"]}</div>'
                    f'<div class="kpi-pct" style="color:{c}">{pct}%</div>'
                    f'</div>'
                    f'<div class="kpi-meta">{project_badge_html(row["Proyecto"])} {tipo}{rspan} · 📅 {row["Fecha"]}</div>'
                    f'<div class="kpi-footer">{bar_html(pct, c)}{kpi_status_html(pct)}</div>'
                    f'</div>'
                )
            return html or '<div style="color:#475569;font-size:0.82rem;padding:10px 0">Sin KPIs asignados</div>'

        lider_display = ""
        if lider and lider != persona and lider not in SOCIOS:
            parts = lider.split()
            lider_display = f" — {parts[0]} {parts[1]}" if len(parts) >= 2 else f" — {lider}"
        card_html = (
            f'<div class="kpi-card">'
            f'<div class="person-header">'
            f'<div class="avatar" style="background:{ac}">{ini}</div>'
            f'<div class="person-info">'
            f'<div class="person-name">{persona}</div>'
            f'<div class="person-role">{role}{lider_display}</div>'
            f'<div>{badges}</div>'
            f'</div>'
            f'<div class="person-score-big" style="color:{score_color(tot)}">{tot:.1f}%</div>'
            f'</div>'
            f'<div class="kpi-section-title">Objetivos Individuales · {len(ind_kpis)} KPIs</div>'
            f'{kpi_rows_html(ind_kpis)}'
            f'<div class="kpi-section-title">Objetivos Equipo · {len(team_kpis)} KPIs</div>'
            f'{kpi_rows_html(team_kpis, show_team=True)}'
            f'</div>'
        )
        st.markdown(card_html, unsafe_allow_html=True)

    with col_salarial:
        def score_box(label, pct, weight):
            c   = score_color(pct)
            pts = pct * weight / 100
            return (
                f'<div class="score-box">'
                f'<div class="score-box-label">{label}</div>'
                f'<div class="score-box-value" style="color:{c}">{pct:.1f}%</div>'
                f'<div class="score-box-bar"><div style="width:{min(pct,100)}%;height:100%;background:{c};border-radius:2px"></div></div>'
                f'<div class="score-box-sub">Aporta {pts:.1f}pts</div>'
                f'</div>'
            )

        tc = score_color(tot)
        salarial_html = (
            f'<div class="kpi-card">'
            f'<div class="kpi-section-title">Desglose Variable Salarial</div>'
            f'<div class="score-grid">'
            f'{score_box("Empresa (30%)", emp, 30)}'
            f'{score_box("Equipo (30%)", equ, 30)}'
            f'{score_box("Individual (30%)", ind, 30)}'
            f'{score_box("Socios (10%)", disc, 10)}'
            f'</div>'
            f'<div class="total-box">'
            f'<div class="total-label">Score Total Ponderado</div>'
            f'<div class="total-value" style="color:{tc}">{tot:.1f}%</div>'
            f'</div>'
            f'</div>'
        )
        st.markdown(salarial_html, unsafe_allow_html=True)


# ─── VISTA: EQUIPOS ───────────────────────────────────────────────────────────
def view_equipos(df):
    team_kpis = df[df["tipo"]=="equipo"].copy()
    proyectos = ["Todos"] + sorted(team_kpis["Proyecto"].unique().tolist())
    filtro = st.selectbox("Proyecto", proyectos, key="eq_f")
    if filtro != "Todos": team_kpis = team_kpis[team_kpis["Proyecto"]==filtro]

    for proyecto in sorted(team_kpis["Proyecto"].unique()):
        pc   = PROJECT_COLORS.get(proyecto,"#64748b")
        rows = team_kpis[team_kpis["Proyecto"]==proyecto]
        rows_html = ""
        for _, row in rows.iterrows():
            pct  = st.session_state.scores["kpis"].get(row["uid"], 0)
            c    = score_color(pct)
            rs   = [r for r in [row["Responsable 1"],row["Responsable 2"],row["Responsable 3"]] if r]
            resp_str = "  ·  ".join(f'<span style="color:#94a3b8">{r}</span>' for r in rs)
            rows_html += (
                f'<div class="kpi-row">'
                f'<div style="display:flex;justify-content:space-between;align-items:flex-start">'
                f'<div class="kpi-name">{row["KPI"]}</div>'
                f'<div class="kpi-pct" style="color:{c}">{pct}%</div>'
                f'</div>'
                f'<div class="kpi-meta">{resp_str} · 📅 {row["Fecha"]}</div>'
                f'<div class="kpi-footer">{bar_html(pct, c)}{kpi_status_html(pct)}</div>'
                f'</div>'
            )
        c_rgb = f"{int(pc[1:3],16)},{int(pc[3:5],16)},{int(pc[5:7],16)}"
        st.markdown(
            f'<div class="kpi-card">'
            f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px">'
            f'<span style="background:rgba({c_rgb},0.2);color:{pc};padding:3px 10px;border-radius:6px;font-weight:700;font-size:0.8rem">{proyecto}</span>'
            f'<span style="color:#64748b;font-size:0.75rem">{len(rows)} KPIs de equipo</span></div>'
            f'{rows_html}</div>', unsafe_allow_html=True)


# ─── VISTA: VARIABLE SALARIAL ─────────────────────────────────────────────────
def view_variable_salarial(df):
    emp_score = s_empresa()
    people    = all_people(df)

    rows_data = []
    for p in people:
        equ  = s_equipo(p, df)
        ind  = s_individual(p, df)
        disc = s_disc(p)
        tot  = emp_score*0.30 + equ*0.30 + ind*0.30 + disc*0.10
        lider_v = df[df["Responsable 1"]==p]["Lider"]
        rows_data.append((p, lider_v.iloc[0] if not lider_v.empty else "—",
                          emp_score, equ, ind, disc, tot))

    rows_data.sort(key=lambda x: x[6], reverse=True)

    # Ranking bars
    rank_html = ""
    for i, (p, lider, emp, equ, ind, disc, tot) in enumerate(rows_data):
        ini = get_initials(p); ac = get_avatar_color(p); c = score_color(tot)
        rank_html += (
            f'<div class="rank-row">'
            f'<span class="rank-num">#{i+1}</span>'
            f'<div class="avatar-sm" style="background:{ac}">{ini}</div>'
            f'<span style="min-width:140px;font-size:0.82rem;color:#e2e8f0">{p.split()[0]} {p.split()[-1]}</span>'
            f'<div class="rank-bar-wrap" style="height:22px">'
            f'<div class="rank-bar-fill" style="width:{tot:.1f}%;background:linear-gradient(90deg,{ac}88,{ac});font-size:0.72rem">'
            f'{tot:.1f}%</div></div>'
            f'<span class="rank-val" style="color:{c}">{tot:.1f}%</span>'
            f'</div>'
        )

    # Table
    header = ('<div style="display:grid;grid-template-columns:160px 100px repeat(4,80px) 80px;gap:8px;'
              'padding:6px 0;border-bottom:1px solid #334155;font-size:0.68rem;font-weight:600;color:#64748b;text-transform:uppercase">'
              '<span>Persona</span><span>Lider</span><span>Empresa</span><span>Equipo</span>'
              '<span>Indiv.</span><span>Disc.</span><span style="color:#6366f1">Total</span></div>')
    table_rows = ""
    for p, lider, emp, equ, ind, disc, tot in rows_data:
        c = score_color(tot)
        table_rows += (f'<div style="display:grid;grid-template-columns:160px 100px repeat(4,80px) 80px;gap:8px;'
                       f'padding:8px 0;border-bottom:1px solid #1a2540;font-size:0.82rem;align-items:center">'
                       f'<span style="color:#e2e8f0;font-weight:500">{p.split()[0]} {p.split()[-1]}</span>'
                       f'<span style="color:#64748b">{lider.split()[0] if lider != "—" else "—"}</span>'
                       f'<span style="color:{score_color(emp)}">{emp:.1f}%</span>'
                       f'<span style="color:{score_color(equ)}">{equ:.1f}%</span>'
                       f'<span style="color:{score_color(ind)}">{ind:.1f}%</span>'
                       f'<span style="color:{score_color(disc)}">{disc:.1f}%</span>'
                       f'<span style="color:{c};font-weight:700">{tot:.1f}%</span>'
                       f'</div>')

    col1, col2 = st.columns([1, 1.2])
    with col1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-section-title">Ranking</div>{rank_html}</div>',
                    unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-section-title">Detalle completo</div>{header}{table_rows}</div>',
                    unsafe_allow_html=True)


# ─── VISTA: EDITOR ────────────────────────────────────────────────────────────
def view_editor(df):
    st.markdown('<div style="color:#94a3b8;font-size:0.85rem;margin-bottom:1rem">Ingresa los porcentajes de cumplimiento. Presiona <b style="color:#6366f1">Guardar</b> cuando termines.</div>',
                unsafe_allow_html=True)

    tab_emp, tab_kpis, tab_disc = st.tabs(["📊 Empresa", "📋 KPIs", "⭐ Discrecional"])

    with tab_emp:
        st.markdown("#### Calificación por proyecto")
        for proyecto in PROJECT_META:
            current = int(st.session_state.scores["empresa"].get(proyecto, 0))
            v = st.slider(f"{proyecto}", 0, 100, current, format="%d%%", key=f"ed_emp__{proyecto}")
            st.session_state.scores["empresa"][proyecto] = v

    with tab_kpis:
        filtro_p = st.selectbox("Proyecto", ["Todos"] + sorted(df["Proyecto"].unique().tolist()), key="ed_proy")
        filtro_t = st.selectbox("Tipo", ["Todos", "equipo", "individual"], key="ed_tipo")
        show_df  = df.copy()
        if filtro_p != "Todos": show_df = show_df[show_df["Proyecto"]==filtro_p]
        if filtro_t != "Todos": show_df = show_df[show_df["tipo"]==filtro_t]

        for _, row in show_df.iterrows():
            uid     = row["uid"]
            current = int(st.session_state.scores["kpis"].get(uid, 0))
            rs      = [r for r in [row["Responsable 1"],row["Responsable 2"],row["Responsable 3"]] if r]
            label   = f"**{row['KPI']}** — {row['Proyecto']} · {', '.join(rs)}"
            v = st.number_input(label, 0, 100, current, key=f"ed_kpi__{uid}")
            st.session_state.scores["kpis"][uid] = v

    with tab_disc:
        st.markdown("#### Calificación discrecional (0 – 10)")
        cols = st.columns(3)
        for i, socio in enumerate(SOCIOS):
            with cols[i]:
                st.markdown(f"**{socio.split()[0]} {socio.split()[1]}**")
                equipo = [p for p in people_by_lider(df, socio) if p not in SOCIOS]
                for persona in equipo:
                    current = int(st.session_state.scores.get("discrecional",{}).get(persona, 0))
                    v = st.number_input(persona.split()[0], 0, 10, current, key=f"ed_disc__{persona}")
                    st.session_state.scores.setdefault("discrecional",{})[persona] = v

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 Guardar todos los cambios", type="primary", use_container_width=True):
        with st.spinner("Guardando en GitHub…"):
            save_scores()
        st.success("✅ Cambios guardados correctamente")


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    df = load_data()
    if "scores" not in st.session_state:
        init_scores(df)

    render_header()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📊 Resumen", "👤 Personas", "👥 Equipos", "💰 Variable Salarial", "✏️ Editor"])

    with tab1: view_resumen(df)
    with tab2: view_personas(df)
    with tab3: view_equipos(df)
    with tab4: view_variable_salarial(df)
    with tab5: view_editor(df)

if __name__ == "__main__":
    main()
