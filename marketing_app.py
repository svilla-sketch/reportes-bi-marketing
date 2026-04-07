import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests, re, json, os, calendar as _cal
import datetime as _dt

# Módulo Vambe separado — toda la lógica de conexión al CRM
from vambe_client import (
    load_vambe, load_vambe_pipeline, load_vambe_pipeline_with_tags,
    load_vambe_all_with_projects, vambe_funnel,
    STAGE_ORDER, STAGE_LABEL, STAGE_COLOR, FUNNEL_CUM,
    load_visits_month, VISIT_STATUS_LABELS,
)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Marketing BI", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

# ID real del spreadsheet (de la URL de edición: docs.google.com/spreadsheets/d/ESTE_ID/edit)
SHEET_ID       = "TU_SHEET_ID_AQUI"
SHEET_TAB      = 0          # índice de la pestaña (0 = primera)
SA_FILE        = "service_account.json"   # credentials de service account en esta carpeta

# Fallback: URL pública TSV (solo funciona si el sheet está publicado)
SHEET_TSV_URL  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTd5oP_xMw1I0J7z1A7STN_n3n6VyhzJXCsni3eWqY96wsINYYSSfcybVAWSFBWdUB44PrBLG-JCBPB/pub?output=tsv"

# ── Sheet columns ─────────────────────────────────────────────────────────────
PROJ_COLORS = {
    "KOS":"#3B82F6","Punto Calma":"#10B981","Zen":"#F59E0B",
    "DODEKA":"#EF4444","SANTIÁN":"#8B5CF6","SANTIAN":"#8B5CF6",
    "TOWNHOUSE":"#EC4899","KOS/ZEN":"#14B8A6","SERVICIO":"#F97316",
}
DEFAULT_COLOR = "#94A3B8"
MONTH_SHORT = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
               7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
MONTH_FULL  = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
               7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
               11:"Noviembre",12:"Diciembre"}
COLS = {
    0:"proyecto",1:"campaña",2:"fecha",3:"mes",4:"año",
    5:"leadsMeta",6:"leadsVambe",7:"leadsOrganicos",8:"recontactos",
    9:"leadsCalificados",10:"leadsArchivados",11:"rebotados",
    12:"prospectosCalificados",13:"pctProspectos",14:"cpp",
    15:"invertido",16:"visitaAgendada",17:"pctAgendada",
    18:"visitaConcretada",19:"pctConcretada",20:"costoVisita",
    21:"apartados",22:"pctApartadoLead",23:"pctApartadoVisita",
    24:"cac",25:"fuente",26:"medio",
}
NUMERIC = {"leadsMeta","leadsVambe","leadsOrganicos","recontactos","leadsCalificados",
           "prospectosCalificados","invertido","visitaAgendada","visitaConcretada",
           "apartados","cac","costoVisita","cpp"}

# ══════════════════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
[data-testid="stSidebarNav"]{display:none}
section[data-testid="stSidebar"]{background:#1E293B;min-width:230px;max-width:260px}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] small,
section[data-testid="stSidebar"] div{color:#E2E8F0!important}
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] .stSelectbox>div>div,
section[data-testid="stSidebar"] .stMultiSelect>div>div{
    background:#0F172A!important;color:#E2E8F0!important;border-color:#334155!important}
section[data-testid="stSidebar"] .stMultiSelect span[data-baseweb="tag"]{
    background:#334155!important}
section[data-testid="stSidebar"] hr{border-color:#334155}
section[data-testid="stSidebar"] .stRadio label{font-size:.88rem;padding:4px 2px}
.sec-lbl{font-size:.65rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
         color:#94A3B8!important;margin:14px 0 4px}
.block-container{padding-top:1.2rem;padding-bottom:.5rem}
div[data-testid="metric-container"]{
    background:white;border-radius:10px;padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.sec-title{font-size:.95rem;font-weight:600;color:#1E293B;
           border-left:4px solid #3B82F6;padding-left:10px;margin:6px 0 10px}
.conv-bar-bg{background:#F1F5F9;border-radius:6px;height:8px;flex:1}
.conv-row{margin-bottom:12px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def div(a,b):   return float(a)/float(b) if b and float(b)!=0 else 0.0
def money(v):   return f"${float(v):,.0f}"
def pct(v):     return f"{float(v):.1f}%"
def scol(df,c): return df[c].sum() if c in df.columns else 0.0

def clean_num(s):
    s=str(s).strip().replace("$","").replace(" ","")
    if s.lower() in ("","-","nan","none","n/a","#n/a","#ref!","#div/0!","#value!","#num!"):
        return 0.0
    s=s.replace("%","")
    if "," in s and "." in s:
        s=s.replace(",","") if s.rfind(".")>s.rfind(",") else s.replace(".","").replace(",",".")
    elif "," in s:
        parts=s.split(",")
        s=s.replace(",",".") if len(parts)==2 and len(parts[1])<=2 else s.replace(",","")
    try:    return float(s)
    except: return 0.0

def conv_bar(label, base, val, cumulative=False):
    rate = div(val,base)*100
    color = "#10B981" if rate>50 else "#F59E0B" if rate>15 else "#EF4444"
    tag = "↗ acum." if cumulative else "→ paso"
    return f"""
<div class="conv-row">
  <div style="display:flex;justify-content:space-between;font-size:.8rem;margin-bottom:3px">
    <span style="color:#475569">{label}</span>
    <span style="font-weight:700;color:{color}">{rate:.1f}%
      <span style="font-size:.65rem;color:#94A3B8;font-weight:400">{tag}</span></span>
  </div>
  <div style="display:flex;align-items:center;gap:6px">
    <div class="conv-bar-bg">
      <div style="width:{min(rate,100):.0f}%;background:{color};height:8px;border-radius:6px"></div>
    </div>
    <span style="font-size:.72rem;color:#94A3B8;min-width:70px">{val:,.0f} / {base:,.0f}</span>
  </div>
</div>"""

def _funnel_table_html(rows_data, show_invertido=True, show_recont=False):
    """
    Genera HTML de tabla de funnel por proyecto.
    rows_data: list of dicts con keys: nombre, leads, recont (opcional), cal, agen, conc, apar, inv (opcional), is_total
    show_recont: muestra columna '+Recont.' junto a Leads (solo datos Excel)
    """
    def _td_conv(val, base, bold):
        p = div(val, base) * 100
        c = "#16A34A" if p >= 50 else "#CA8A04" if p >= 20 else "#DC2626"
        s = "font-weight:700;" if bold else ""
        return (f"<td style='text-align:center;padding:8px 4px;{s}'>"
                f"{int(val):,} <span style='color:{c};font-size:.78rem'>({p:.1f}%)</span></td>")

    recont_th  = ("<th style='text-align:center;color:#F59E0B;font-size:.8rem'>+Recontactos</th>"
                  if show_recont else "")
    inv_th     = "<th style='text-align:right;color:#64748B'>Invertido</th>" if show_invertido else ""
    hdr = (
        "<table style='width:100%;border-collapse:collapse;font-size:.87rem'>"
        "<thead><tr style='border-bottom:2px solid #E2E8F0'>"
        "<th style='text-align:left;padding:8px 6px;color:#64748B'>Proyecto</th>"
        "<th style='text-align:center;color:#64748B'>Leads</th>"
        + recont_th +
        "<th style='text-align:center;color:#3B82F6'>→ Calificados</th>"
        "<th style='text-align:center;color:#8B5CF6'>→ Agendadas</th>"
        "<th style='text-align:center;color:#10B981'>→ Concretadas</th>"
        "<th style='text-align:center;color:#EC4899'>→ Apartados</th>"
        + inv_th +
        "</tr></thead><tbody>"
    )
    body = ""
    for r in rows_data:
        b   = r["is_total"]
        bg  = "#F8FAFC" if b else "white"
        brd = "border-top:2px solid #E2E8F0;" if b else "border-bottom:1px solid #F1F5F9;"
        ws  = "font-weight:700;" if b else ""
        lv  = r["leads"]          # leadsVambe + recontactos cuando show_recont=True
        rv  = r.get("recont", 0)  # recontactos por separado
        # Columna +Recontactos: muestra el número con fondo ámbar suave
        recont_td = ""
        if show_recont:
            recont_td = (
                f"<td style='text-align:center;{ws}padding:8px 4px;"
                f"background:#FFFBEB;color:#92400E;font-size:.85rem'>"
                f"+{int(rv):,}</td>"
            )
        body += (
            f"<tr style='background:{bg};{brd}'>"
            f"<td style='padding:8px 6px;{ws}'>{r['nombre']}</td>"
            f"<td style='text-align:center;{ws}'>{int(lv):,}</td>"
            + recont_td
            + _td_conv(r["cal"],  lv,        b)
            + _td_conv(r["agen"], r["cal"],   b)
            + _td_conv(r["conc"], r["agen"],  b)
            + _td_conv(r["apar"], r["conc"],  b)
            + (f"<td style='text-align:right;{ws}'>{money(r['inv'])}</td>" if show_invertido else "")
            + "</tr>"
        )
    return hdr + body + "</tbody></table>"

# ══════════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ══════════════════════════════════════════════════════════════════════════════
def _parse_rows(raw_rows):
    """Convierte lista de listas (valores crudos) en DataFrame limpio."""
    from datetime import datetime as _dt
    if len(raw_rows) < 2:
        return None, "Sheet vacío"
    rows = []
    for vals in raw_rows[1:]:
        row = {name: (clean_num(vals[i]) if name in NUMERIC else (vals[i].strip() if i < len(vals) else ""))
               for i, name in COLS.items() if i < len(vals)}
        rows.append(row)
    df = pd.DataFrame(rows)
    df = df[~df["proyecto"].str.lower().str.contains("total|proyecto", na=False)]
    df = df[df["proyecto"].str.strip() != ""]

    def parse_fecha(s):
        s = str(s).strip()
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                d = _dt.strptime(s, fmt)
                return d.month, d.year if d.year > 100 else d.year + 2000
            except: pass
        return None, None
    def _safe_year(v):
        try: f = float(v); return f if 2000 <= f <= 2100 else None
        except: return None
    def _safe_month(v):
        try: f = float(v); return f if 1 <= f <= 12 else None
        except: return None

    parsed = df["fecha"].apply(parse_fecha)
    df["mes_num"] = df.get("mes", pd.Series(dtype=str)).apply(_safe_month)
    df["año_num"] = df.get("año", pd.Series(dtype=str)).apply(_safe_year)
    df["mes_num"] = df["mes_num"].fillna(parsed.apply(lambda x: x[0]))
    df["año_num"] = df["año_num"].fillna(parsed.apply(lambda x: x[1]))
    df["mes_num"] = pd.to_numeric(df["mes_num"], errors="coerce")
    df["año_num"] = pd.to_numeric(df["año_num"], errors="coerce")
    return df, ""

@st.cache_data(ttl=300, show_spinner="Cargando sheet…")
def load_sheet():
    import os

    # ── Intento 1: Google Sheets API con service account ─────────────────────
    sa_path = os.path.join(os.path.dirname(__file__), SA_FILE)
    if os.path.exists(sa_path) and SHEET_ID != "TU_SHEET_ID_AQUI":
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            scopes = ["https://spreadsheets.google.com/feeds",
                      "https://www.googleapis.com/auth/drive"]
            creds  = Credentials.from_service_account_file(sa_path, scopes=scopes)
            client = gspread.authorize(creds)
            ws     = client.open_by_key(SHEET_ID).get_worksheet(SHEET_TAB)
            rows   = ws.get_all_values()
            df, err = _parse_rows(rows)
            if df is not None:
                return df, "api"   # "api" indica que vino de la API
            return None, err
        except Exception as e:
            pass   # cae al fallback TSV

    # ── Intento 2: URL TSV pública (fallback) ─────────────────────────────────
    try:
        resp = requests.get(SHEET_TSV_URL, timeout=20)
        resp.raise_for_status()
        if resp.text.strip().startswith("<"):
            return None, "sheet_not_published"
        lines = [l for l in resp.text.strip().split("\n") if l.strip()]
        raw   = [l.split("\t") for l in lines]
        df, err = _parse_rows(raw)
        if df is not None:
            return df, "tsv"
        return None, err
    except Exception as e:
        return None, str(e)

@st.cache_data(ttl=180, show_spinner="Cargando Vambe…")
def _load_vambe_cached(days): return load_vambe(days)

@st.cache_data(ttl=180, show_spinner="Cargando pipeline Vambe…")
def _load_vambe_pipeline_cached(): return load_vambe_pipeline()

@st.cache_data(ttl=1800, show_spinner="Cargando proyectos (Supabase)…")
def _load_vambe_tags_cached(): return load_vambe_pipeline_with_tags()

@st.cache_data(ttl=1800, show_spinner="Cargando leads por proyecto…")
def _load_vambe_all_projects_cached(): return load_vambe_all_with_projects(365)

# ── Metas helpers ─────────────────────────────────────────────────────────────
_METAS_FILE = os.path.join(os.getcwd(), "metas_q.json")
_METAS_PROJS = ["KOS", "Punto Calma", "Zen", "DODEKA", "SANTIÁN"]
_METAS_KEYS  = ["leads", "cal", "agen", "conc", "apar", "pres"]
_METAS_COLS  = ["Leads", "Cal.", "Agen.", "Conc.", "Apar.", "Presup. $"]
_Q1 = {1: "Enero", 2: "Febrero", 3: "Marzo"}
_Q2 = {4: "Abril", 5: "Mayo", 6: "Junio"}
_QUARTERS = {"Q1 2026": _Q1, "Q2 2026": _Q2}

def _metas_load():
    try:
        with open(_METAS_FILE) as f:
            return json.load(f)
    except Exception:
        return {}

def _metas_save(data):
    with open(_METAS_FILE, "w") as f:
        json.dump(data, f, indent=2)

def _meta_get(data, year, month, proj, key):
    return float(data.get(str(year), {}).get(str(month), {}).get(proj, {}).get(key, 0))

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📊 Marketing BI")
    st.caption("Inmobiliaria · Campañas")
    st.divider()

    # Navegación
    st.markdown('<p class="sec-lbl">Sección</p>', unsafe_allow_html=True)
    seccion = st.radio("nav",[
        "📈 Vista General",
        "📡 Vambe Live",
        "🔻 Embudo",
        "🎯 Metas",
        "🗓️ Visitas",
        "💰 Inversión",
        "🏗️ Proyectos",
        "📋 Datos",
    ], label_visibility="collapsed")

    st.divider()

    # Cargar sheet
    df_raw, sheet_src = load_sheet()
    if df_raw is None:
        if sheet_src == "sheet_not_published":
            st.warning("⚠️ Sheet no publicado — configura la API o publícalo en Drive.")
        else:
            st.error(f"❌ Sheet: {sheet_src}")
        df_raw = pd.DataFrame()

    # Vambe — cuatro fuentes:
    # vdf_raw:          todos los contactos (lead counts por fecha)
    # vdf_pipe_raw:     pipeline sin tags (rápido, embudo/etapas)
    # vdf_tags_raw:     pipeline + proyecto/canal (para conversiones por proyecto)
    # vdf_all_proj_raw: todos los contactos + proyecto (para leads totales por proyecto)
    vdf_raw, vambe_err              = _load_vambe_cached(365)
    vdf_pipe_raw, vambe_pipe_err    = _load_vambe_pipeline_cached()
    vdf_tags_raw, vambe_tags_err    = _load_vambe_tags_cached()
    vdf_all_proj_raw, _             = _load_vambe_all_projects_cached()

    # ── Filtros compactos ────────────────────────────────────────────────────
    st.markdown('<p class="sec-lbl">Filtros</p>', unsafe_allow_html=True)

    if not df_raw.empty:
        years = sorted(df_raw["año_num"].dropna().astype(int).unique().tolist(), reverse=True)
        projs = sorted(df_raw["proyecto"].dropna().unique().tolist())
        months = sorted(df_raw["mes_num"].dropna().astype(int).unique().tolist())
    else:
        years, projs, months = [], [], []

    sel_year  = st.selectbox("Año", ["Todos"]+[str(y) for y in years], label_visibility="visible")

    st.markdown('<p class="sec-lbl">Proyectos</p>', unsafe_allow_html=True)
    sel_projs = []
    for p in projs:
        if st.checkbox(p, value=True, key=f"proj_{p}"):
            sel_projs.append(p)
    if not sel_projs:
        sel_projs = projs  # si no hay ninguno marcado, mostrar todos

    month_opts = {MONTH_FULL.get(m,str(m)):m for m in months}
    st.markdown('<p class="sec-lbl">Meses</p>', unsafe_allow_html=True)
    sel_months = []
    for mname, mnum in month_opts.items():
        if st.checkbox(mname, value=True, key=f"mes_{mnum}"):
            sel_months.append(mnum)
    if not sel_months:
        sel_months = list(month_opts.values())

    st.divider()

    st.divider()

    # Recargar
    if st.button("🔄 Recargar datos", use_container_width=True):
        st.cache_data.clear(); st.rerun()

    # Status
    if not df_raw.empty:
        src_tag = "🔗 API" if sheet_src == "api" else "📄 TSV"
        st.success(f"✅ Sheet ({src_tag}): {len(df_raw):,} filas")
    if not vdf_raw.empty:
        st.success(f"✅ Vambe: {len(vdf_raw):,} contactos")
    elif vambe_err:
        st.error(f"❌ Vambe: {vambe_err[:60]}")

# ══════════════════════════════════════════════════════════════════════════════
# FILTRAR SHEET
# ══════════════════════════════════════════════════════════════════════════════
df = df_raw.copy() if not df_raw.empty else pd.DataFrame()
if not df.empty:
    if sel_year != "Todos":
        df = df[df["año_num"].fillna(0).astype(int)==int(sel_year)]
    if sel_projs:
        df = df[df["proyecto"].isin(sel_projs)]
    if sel_months:
        df = df[df["mes_num"].isin(sel_months)]

    def _periodo(row):
        try: return f"{MONTH_SHORT[int(row.mes_num)]} {int(row.año_num)}"
        except: return "?"
    def _periodo_sort(row):
        try: return int(row.año_num)*100+int(row.mes_num)
        except: return 0
    df["periodo"]      = df.apply(_periodo, axis=1)
    df["periodo_sort"] = df.apply(_periodo_sort, axis=1)

if not df.empty:
    inv         = scol(df,"invertido")
    leads_meta  = scol(df,"leadsMeta")
    leads_vambe = scol(df,"leadsVambe")
    leads_org   = scol(df,"leadsOrganicos")
    leads       = leads_vambe
    cal         = scol(df,"leadsCalificados")
    agen        = scol(df,"visitaAgendada")
    conc        = scol(df,"visitaConcretada")
    apar        = scol(df,"apartados")
    cac_val     = div(inv,apar)
    cpv_val     = div(inv,conc)
    cpl_val     = div(inv,leads)
    cpp_val     = div(inv,cal)
else:
    inv=leads_meta=leads_vambe=leads_org=leads=cal=agen=conc=apar=0
    cac_val=cpv_val=cpl_val=cpp_val=0.0

if df.empty and "Vista General" in seccion:
    st.warning("Sin datos. Verifica la conexión al sheet.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: VISTA GENERAL
# ══════════════════════════════════════════════════════════════════════════════
if "Vista General" in seccion:
    st.markdown("## 📈 Vista General")

    # ── KPIs fila 1: embudo ──────────────────────────────────────────────────
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("📥 Leads Vambe",  f"{leads:,.0f}")
    k2.metric("✅ Calificados",  f"{cal:,.0f}",
              delta=pct(div(cal,leads)*100)+" de leads")
    k3.metric("📅 Agendadas",    f"{agen:,.0f}",
              delta=pct(div(agen,cal)*100)+" de cal.")
    k4.metric("🏠 Concretadas",  f"{conc:,.0f}",
              delta=pct(div(conc,agen)*100)+" de agen.")
    k5.metric("🏆 Apartados",    f"{apar:,.0f}",
              delta=pct(div(apar,conc)*100)+" de vis.")

    # ── KPIs fila 2: costos ──────────────────────────────────────────────────
    c1,c2,c3,c4,_ = st.columns([1,1,1,1,1])
    c1.metric("💵 Invertido",      money(inv))
    c2.metric("💲 CPL",            money(cpl_val), delta="por lead",        delta_color="off")
    c3.metric("📌 CPP",            money(cpp_val), delta="por calificado",  delta_color="off")
    c4.metric("💳 CAC",            money(cac_val), delta="por apartado",    delta_color="off")

    st.divider()

    st.divider()

    # ── Gráficas cronológicas ─────────────────────────────────────────────────
    if not df.empty:
        c1,c2 = st.columns(2)
        with c1:
            st.markdown('<p class="sec-title">Leads Vambe por Mes</p>', unsafe_allow_html=True)
            mdf=(df.groupby(["periodo_sort","periodo"])
                   .agg(Leads=("leadsVambe","sum")).reset_index().sort_values("periodo_sort"))
            fig=px.bar(mdf,x="periodo",y="Leads",text="Leads",
                       color_discrete_sequence=["#3B82F6"])
            fig.update_traces(textposition="outside")
            fig.update_layout(height=250,margin=dict(t=10,b=0,l=0,r=0),
                              plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                              showlegend=False,
                              xaxis=dict(gridcolor="#F1F5F9",title="",
                                         categoryorder="array",categoryarray=mdf["periodo"].tolist()),
                              yaxis=dict(gridcolor="#F1F5F9",title=""))
            st.plotly_chart(fig,use_container_width=True)

        with c2:
            st.markdown('<p class="sec-title">Inversión por Mes</p>', unsafe_allow_html=True)
            idf=(df.groupby(["periodo_sort","periodo"])["invertido"]
                   .sum().reset_index().sort_values("periodo_sort"))
            idf["label"]=idf["invertido"].apply(money)
            fig2=px.bar(idf,x="periodo",y="invertido",text="label",
                        color_discrete_sequence=["#10B981"])
            fig2.update_traces(textposition="outside")
            fig2.update_layout(height=250,margin=dict(t=10,b=0,l=0,r=0),
                               plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                               showlegend=False,
                               xaxis=dict(gridcolor="#F1F5F9",title="",
                                          categoryorder="array",categoryarray=idf["periodo"].tolist()),
                               yaxis=dict(gridcolor="#F1F5F9",title="",showticklabels=False))
            st.plotly_chart(fig2,use_container_width=True)

        # Resumen tabla
        st.markdown('<p class="sec-title">Resumen por Mes</p>', unsafe_allow_html=True)
        res=(df.groupby(["periodo_sort","periodo"]).agg(
            Invertido=("invertido","sum"),Leads=("leadsVambe","sum"),
            Calificados=("leadsCalificados","sum"),Agendadas=("visitaAgendada","sum"),
            Concretadas=("visitaConcretada","sum"),Apartados=("apartados","sum"))
            .reset_index().sort_values("periodo_sort"))
        res["Cal%"]     = res.apply(lambda r: pct(div(r.Calificados,r.Leads)*100),axis=1)
        res["CAC"]      = res.apply(lambda r: money(div(r.Invertido,r.Apartados)),axis=1)
        res["Conv%"]    = res.apply(lambda r: pct(div(r.Apartados,r.Calificados)*100),axis=1)
        res["Invertido"]= res["Invertido"].apply(money)
        res=res.rename(columns={"periodo":"Mes"})
        show=["Mes","Invertido","Leads","Cal%","Calificados","Agendadas","Concretadas","Apartados","CAC","Conv%"]
        st.dataframe(res[[c for c in show if c in res.columns]].set_index("Mes"),
                     use_container_width=True, height=220)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: VAMBE LIVE
# ══════════════════════════════════════════════════════════════════════════════
elif "Vambe Live" in seccion:
    st.markdown("## 📡 Vambe Live — Pipeline en Tiempo Real")

    if vdf_raw.empty and vdf_pipe_raw.empty:
        st.error(f"No se pudo conectar a Vambe. {vambe_err}")
        st.stop()

    # ── Selector de modo de filtro ────────────────────────────────────────────
    import pandas as _pd
    now_utc = _pd.Timestamp.now(tz="UTC")

    fa, fb, fc = st.columns([2, 2, 4])
    with fa:
        modo_filtro = st.radio("Corte de tiempo", ["Relativo", "Por mes"],
                               horizontal=True, label_visibility="collapsed")

    if modo_filtro == "Relativo":
        with fb:
            rel_opts = {
                "Últimos 7 días": 7, "Últimos 30 días": 30,
                "Últimos 60 días": 60, "Últimos 90 días": 90,
                "Últimos 6 meses": 180, "Todo el año": 365,
            }
            rel_sel = st.selectbox("Período", list(rel_opts.keys()),
                                   index=1, label_visibility="collapsed")
        dias_rel  = rel_opts[rel_sel]
        fecha_min = now_utc - _pd.Timedelta(days=dias_rel)
        fecha_max = now_utc
        lbl_periodo = rel_sel

        # Para el modo relativo cargamos con la misma ventana
        @st.cache_data(ttl=180)
        def _load_v(d): return _load_vambe_cached(d)
        vdf_all, _ = _load_v(dias_rel)

    else:  # Por mes
        # Construir lista de meses disponibles en vdf_raw
        _base = vdf_raw if not vdf_raw.empty else vdf_pipe_raw
        if not _base.empty:
            meses_disp = (
                _base.dropna(subset=["mes_num","año_num"])
                .assign(sort=lambda d: d["año_num"].astype(int)*100 + d["mes_num"].astype(int))
                .drop_duplicates(subset=["sort"])
                .sort_values("sort", ascending=False)
            )
            mes_labels = [
                f"{MONTH_FULL.get(int(r.mes_num), str(int(r.mes_num)))} {int(r.año_num)}"
                for _, r in meses_disp.iterrows()
            ]
            mes_keys   = [(int(r.mes_num), int(r.año_num)) for _, r in meses_disp.iterrows()]
        else:
            mes_labels, mes_keys = [], []

        with fb:
            idx_sel = st.selectbox("Mes", range(len(mes_labels)),
                                   format_func=lambda i: mes_labels[i] if mes_labels else "—",
                                   label_visibility="collapsed")

        if mes_labels:
            sel_mes, sel_año = mes_keys[idx_sel]
            fecha_min = _pd.Timestamp(year=sel_año, month=sel_mes, day=1, tz="UTC")
            fecha_max = (fecha_min + _pd.offsets.MonthEnd(1)).replace(
                hour=23, minute=59, second=59)
            lbl_periodo = mes_labels[idx_sel]
        else:
            sel_mes, sel_año = now_utc.month, now_utc.year
            fecha_min, fecha_max = now_utc - _pd.Timedelta(days=30), now_utc
            lbl_periodo = "Sin datos"

        # Para mes específico filtramos vdf_raw (ya tiene 365 días)
        vdf_all = vdf_raw.copy() if not vdf_raw.empty else pd.DataFrame()

    # ── Aplicar filtro de fecha a ambas fuentes ───────────────────────────────
    def _filtrar_fecha(df):
        if df.empty or "fecha_creacion" not in df.columns:
            return df
        return df[(df["fecha_creacion"] >= fecha_min) &
                  (df["fecha_creacion"] <= fecha_max)].copy()

    vdf_periodo  = _filtrar_fecha(vdf_all)       # todos los leads del período
    vdf_pipe     = _filtrar_fecha(vdf_pipe_raw)  # leads del período con etapa en pipeline

    # ── Métricas ──────────────────────────────────────────────────────────────
    tv  = len(vdf_periodo)     # leads totales en el período
    fn  = vambe_funnel(vdf_pipe)
    tp  = fn["total"]
    cv  = fn["calificados"]
    av  = fn["vis_agendada"]
    vv  = fn["vis_concretada"]
    apv = fn["apartados"]

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("📥 Leads",           f"{tv:,}")
    k2.metric("✅ Calificados",     f"{cv:,}",  delta=pct(div(cv,tv)*100)+" de leads" if tv else "—")
    k3.metric("📅 Vis. Agendada",   f"{av:,}",  delta=pct(div(av,tv)*100)+" de leads" if tv else "—")
    k4.metric("🏠 Vis. Concretada", f"{vv:,}",  delta=pct(div(vv,tv)*100)+" de leads" if tv else "—")
    k5.metric("🏆 Apartados",       f"{apv:,}", delta=pct(div(apv,tv)*100)+" de leads" if tv else "—")
    st.caption(f"Período: **{lbl_periodo}** · {tv:,} leads totales · {tp:,} en etapas del pipeline")

    st.divider()

    # ── 1. Funnel por Proyecto (tabla principal) ───────────────────────────────
    st.markdown('<p class="sec-title">Funnel por Proyecto — Datos Vambe</p>',
                unsafe_allow_html=True)

    if vambe_tags_err:
        st.error(f"Error cargando proyectos: {vambe_tags_err}")
    elif vdf_tags_raw.empty and vdf_all_proj_raw.empty:
        st.info("Sin datos de proyecto disponibles.")
    else:
        # Leads totales por proyecto (todos los contactos del período)
        _va = _filtrar_fecha(vdf_all_proj_raw) if not vdf_all_proj_raw.empty else pd.DataFrame()
        # Conversiones (solo pipeline — tienen stage_key)
        _vt = _filtrar_fecha(vdf_tags_raw) if not vdf_tags_raw.empty else pd.DataFrame()

        if _va.empty and _vt.empty:
            st.info("Sin contactos con proyecto en este período.")
        else:
            projs_a = set(_va["proyecto"].unique()) if not _va.empty else set()
            projs_t = set(_vt["proyecto"].unique()) if not _vt.empty else set()
            proyectos_v = sorted(projs_a | projs_t)

            # ── CSS: números como links clickeables ───────────────────────
            st.markdown("""
            <style>
            .funnel-btn button {
                background:none!important;border:none!important;padding:2px 6px!important;
                font-size:.92rem!important;font-weight:700!important;color:#3B82F6!important;
                cursor:pointer!important;text-decoration:underline dotted!important;
                min-height:0!important;height:auto!important;line-height:1.4!important;
            }
            .funnel-btn button:hover{color:#1D4ED8!important;background:#EFF6FF!important;
                border-radius:4px!important;text-decoration:underline!important;}
            .funnel-btn-muted button {color:#94A3B8!important;}
            </style>""", unsafe_allow_html=True)

            _PROJ_C2 = {"KOS":"#3B82F6","Punto Calma":"#10B981","Zen":"#F59E0B",
                        "DODEKA":"#EF4444","SANTIÁN":"#8B5CF6","SANTIAN":"#8B5CF6"}

            # Columnas: Proyecto | Leads | →Cal | →Agen | →Conc | →Apar
            _HDR_COLS = [3, 1.2, 1.8, 1.8, 1.8, 1.8]
            _hc = st.columns(_HDR_COLS)
            _hc[0].markdown("<span style='font-size:.8rem;color:#64748B;font-weight:600'>Proyecto</span>",
                            unsafe_allow_html=True)
            for _i, _lbl in enumerate(["Leads","→ Calificados","→ Agendadas","→ Concretadas","→ Apartados"]):
                _hc[_i+1].markdown(
                    f"<span style='font-size:.8rem;color:#64748B;font-weight:600'>{_lbl}</span>",
                    unsafe_allow_html=True)

            st.markdown("<hr style='margin:4px 0 2px;border-color:#E2E8F0'>", unsafe_allow_html=True)

            # Datos por proyecto
            _funnel_data = []
            for proy in proyectos_v:
                leads_total = int((_va["proyecto"] == proy).sum()) if not _va.empty else 0
                _sub_pipe   = _vt[_vt["proyecto"] == proy] if not _vt.empty else pd.DataFrame()
                _fn         = vambe_funnel(_sub_pipe)
                _funnel_data.append({
                    "nombre": proy,
                    "leads":  leads_total,
                    "cal":    _fn["calificados"],
                    "agen":   _fn["vis_agendada"],
                    "conc":   _fn["vis_concretada"],
                    "apar":   _fn["apartados"],
                })
            # Fila TOTAL
            _fn_tot   = vambe_funnel(_vt) if not _vt.empty else vambe_funnel(pd.DataFrame())
            _tot_row  = {"nombre":"TOTAL","leads":len(_va) if not _va.empty else 0,
                         "cal":_fn_tot["calificados"],"agen":_fn_tot["vis_agendada"],
                         "conc":_fn_tot["vis_concretada"],"apar":_fn_tot["apartados"]}

            def _pct_lbl(val, base):
                p = div(val, base) * 100
                c = "#16A34A" if p >= 50 else "#CA8A04" if p >= 20 else "#DC2626"
                return f"<span style='color:{c};font-size:.75rem'>({p:.1f}%)</span>"

            def _render_funnel_row(row, is_total=False):
                """Renderiza una fila del funnel con botones clickeables."""
                _fw  = "700" if is_total else "500"
                _bg  = "background:#F8FAFC;" if is_total else ""
                _pc  = _PROJ_C2.get(row["nombre"], "#64748B")
                _rc  = st.columns(_HDR_COLS)
                # Nombre proyecto
                if is_total:
                    _rc[0].markdown(f"<div style='{_bg}padding:6px 4px;font-weight:700;"
                                    f"font-size:.88rem'>TOTAL</div>", unsafe_allow_html=True)
                else:
                    _rc[0].markdown(
                        f"<div style='padding:6px 4px;font-weight:600;color:{_pc}'>"
                        f"{row['nombre']}</div>", unsafe_allow_html=True)

                # Leads
                _key_l = f"fl_{row['nombre']}_leads"
                with _rc[1]:
                    st.markdown("<div class='funnel-btn'>", unsafe_allow_html=True)
                    if st.button(str(row["leads"]), key=_key_l, use_container_width=False):
                        st.session_state["funnel_drill"] = (row["nombre"], "leads")
                    st.markdown("</div>", unsafe_allow_html=True)

                # Cal / Agen / Conc / Apar
                for _ci, (_fk, _col_key) in enumerate([("cal","calificados"),("agen","vis_agendada"),
                                                        ("conc","vis_concretada"),("apar","apartados")]):
                    _val  = row[_fk]
                    _base = row["leads"] if _fk == "cal" else row[["cal","agen","conc"][_ci-1] if _ci > 0 else "cal"]
                    _pct  = _pct_lbl(_val, row["leads"])
                    _key  = f"fl_{row['nombre']}_{_fk}"
                    with _rc[_ci+2]:
                        st.markdown("<div class='funnel-btn" +
                                    (" funnel-btn-muted" if _val == 0 else "") + "'>",
                                    unsafe_allow_html=True)
                        if st.button(f"{_val}", key=_key, use_container_width=False):
                            st.session_state["funnel_drill"] = (row["nombre"], _col_key)
                        st.markdown("</div>", unsafe_allow_html=True)
                    # Porcentaje debajo del botón
                    _rc[_ci+2].markdown(_pct, unsafe_allow_html=True)

            for _fd in _funnel_data:
                _render_funnel_row(_fd)
            st.markdown("<hr style='margin:2px 0 4px;border-color:#E2E8F0;border-width:2px'>",
                        unsafe_allow_html=True)
            _render_funnel_row(_tot_row, is_total=True)

            _leads_total = len(_va) if not _va.empty else 0
            _pipe_total  = len(_vt) if not _vt.empty else 0
            st.caption(
                f"{_leads_total} leads totales · {_pipe_total} en pipeline · "
                f"{_leads_total - _pipe_total} sin etapa asignada"
            )

            # ── Panel de contactos (aparece al hacer clic en un número) ───
            if "funnel_drill" in st.session_state:
                _drill_proj, _drill_nivel = st.session_state["funnel_drill"]
                _nivel_map = {
                    "leads":         ("📥 Leads totales", None, True),
                    "calificados":   ("✅ Calificados+",  FUNNEL_CUM["calificados"],  False),
                    "vis_agendada":  ("📅 Agendadas+",    FUNNEL_CUM["vis_agendada"],  False),
                    "vis_concretada":("🏠 Concretadas+",  FUNNEL_CUM["vis_concretada"],False),
                    "apartados":     ("🏆 Apartados",      FUNNEL_CUM["apartados"],     False),
                }
                _nlbl, _stage_keys, _use_all = _nivel_map.get(
                    _drill_nivel, ("?", None, False))

                if _use_all:   # Leads: usar vdf_all_proj_raw
                    _drill_src = _va.copy() if not _va.empty else pd.DataFrame()
                elif _stage_keys:
                    _drill_src = (_vt[_vt["stage_key"].isin(_stage_keys)].copy()
                                  if not _vt.empty else pd.DataFrame())
                else:
                    _drill_src = pd.DataFrame()

                if _drill_proj != "TOTAL" and not _drill_src.empty:
                    _drill_src = _drill_src[_drill_src["proyecto"] == _drill_proj]

                _panel_title = (f"{'Todos los proyectos' if _drill_proj=='TOTAL' else _drill_proj}"
                                f" — {_nlbl}")
                st.markdown(f"<div style='margin-top:16px;padding:10px 0 4px;"
                            f"border-top:2px solid #E2E8F0;font-weight:700;color:#1E293B;"
                            f"font-size:.95rem'>{_panel_title}"
                            f"<span style='font-weight:400;color:#64748B;font-size:.8rem;"
                            f"margin-left:8px'>{len(_drill_src)} contactos</span></div>",
                            unsafe_allow_html=True)

                if _drill_src.empty:
                    st.info("Sin contactos en esta selección.")
                else:
                    _ctbl3 = (
                        "<table style='width:100%;border-collapse:collapse;font-size:.83rem'>"
                        "<thead><tr style='border-bottom:2px solid #E2E8F0'>"
                        "<th style='padding:6px 8px;color:#64748B;text-align:left'>Nombre</th>"
                        "<th style='padding:6px 6px;color:#64748B;text-align:left'>Proyecto</th>"
                        "<th style='padding:6px 6px;color:#64748B;text-align:left'>Canal</th>"
                        "<th style='padding:6px 6px;color:#64748B;text-align:left'>Etapa</th>"
                        "<th style='padding:6px 6px;color:#64748B;text-align:left'>Fecha</th>"
                        "<th style='padding:6px 6px;color:#64748B;text-align:left'>Teléfono</th>"
                        "</tr></thead><tbody>"
                    )
                    for _, _cr in _drill_src.sort_values("fecha_creacion",
                                                          ascending=False).iterrows():
                        _pc3  = _PROJ_C2.get(_cr.get("proyecto",""), "#94A3B8")
                        _sc3  = STAGE_COLOR.get(_cr.get("stage_key",""), "#94A3B8")
                        _fec3 = (pd.Timestamp(_cr["fecha_creacion"]).strftime("%d/%m/%Y")
                                 if pd.notna(_cr.get("fecha_creacion")) else "—")
                        _ctbl3 += (
                            f"<tr style='border-bottom:1px solid #F1F5F9'>"
                            f"<td style='padding:5px 8px;font-weight:600;color:#1E293B'>"
                            f"{_cr.get('nombre','—')}</td>"
                            f"<td style='padding:5px 6px'><span style='background:{_pc3}22;"
                            f"color:{_pc3};border-radius:6px;padding:1px 7px;font-size:.75rem;"
                            f"font-weight:600'>{_cr.get('proyecto') or 'Sin proyecto'}</span></td>"
                            f"<td style='padding:5px 6px;color:#64748B;font-size:.79rem'>"
                            f"{_cr.get('canal') or '—'}</td>"
                            f"<td style='padding:5px 6px;color:{_sc3};font-size:.79rem;"
                            f"font-weight:600'>{_cr.get('stage_label') or 'Sin etapa'}</td>"
                            f"<td style='padding:5px 6px;color:#64748B;font-size:.78rem'>{_fec3}</td>"
                            f"<td style='padding:5px 6px;color:#94A3B8;font-size:.76rem'>"
                            f"{_cr.get('telefono','—')}</td></tr>"
                        )
                    _ctbl3 += "</tbody></table>"
                    st.markdown(_ctbl3, unsafe_allow_html=True)

    st.divider()

    # ── 2. Embudo acumulado + conversiones ────────────────────────────────────
    c1, c2 = st.columns([2, 3])

    with c1:
        st.markdown('<p class="sec-title">Embudo acumulado</p>', unsafe_allow_html=True)
        stages_f = [
            ("Leads",           tv,  "#64748B"),
            ("Calificados",     cv,  "#3B82F6"),
            ("Vis. Agendada",   av,  "#8B5CF6"),
            ("Vis. Concretada", vv,  "#10B981"),
            ("Apartados",       apv, "#EC4899"),
        ]
        fig2 = go.Figure(go.Funnel(
            y=[s[0] for s in stages_f], x=[s[1] for s in stages_f],
            textinfo="value+percent initial",
            marker=dict(color=[s[2] for s in stages_f]),
        ))
        fig2.update_layout(height=320, margin=dict(t=10,b=10,l=10,r=10),
                           paper_bgcolor="rgba(0,0,0,0)", font=dict(size=13))
        st.plotly_chart(fig2, use_container_width=True)

    with c2:
        st.markdown('<p class="sec-title">Conversiones paso a paso</p>', unsafe_allow_html=True)
        html_s = ""
        for label, base, val in [
            ("Leads → Calificados",        tv, cv),
            ("Calificados → Vis. Agendada", cv, av),
            ("Vis. Agendada → Concretada",  av, vv),
            ("Vis. Concretada → Apartado",  vv, apv),
        ]:
            html_s += conv_bar(label, base, val, cumulative=False)
        st.markdown(html_s, unsafe_allow_html=True)

        st.markdown('<p class="sec-title" style="margin-top:14px">Acumuladas desde Lead</p>',
                    unsafe_allow_html=True)
        html_c = ""
        for label, val in [
            ("Lead → Calificado",      cv),
            ("Lead → Vis. Agendada",   av),
            ("Lead → Vis. Concretada", vv),
            ("Lead → Apartado",        apv),
        ]:
            html_c += conv_bar(label, tv, val, cumulative=True)
        st.markdown(html_c, unsafe_allow_html=True)

    st.divider()

    # ── 3. Histórico de leads por mes ─────────────────────────────────────────
    st.markdown('<p class="sec-title">Nuevos Leads por Mes (histórico)</p>',
                unsafe_allow_html=True)
    _hist = vdf_raw if not vdf_raw.empty else pd.DataFrame()
    if not _hist.empty and "año_num" in _hist.columns:
        _h2 = _hist.dropna(subset=["mes_num","año_num"]).copy()
        _h2["periodo_sort"] = _h2["año_num"].astype(int)*100 + _h2["mes_num"].astype(int)
        _h2["periodo"] = _h2.apply(
            lambda r: f"{MONTH_SHORT.get(int(r.mes_num),'?')} {int(r.año_num)}", axis=1)
        mc = (_h2.groupby(["periodo_sort","periodo"]).size()
                  .reset_index(name="Leads").sort_values("periodo_sort"))
        if modo_filtro == "Por mes" and mes_labels:
            lbl_resaltar = f"{MONTH_SHORT.get(sel_mes,'?')} {sel_año}"
            mc["color"] = mc["periodo"].apply(
                lambda p: "#3B82F6" if p == lbl_resaltar else "#CBD5E1")
        else:
            mc["color"] = "#3B82F6"
        fig3 = go.Figure(go.Bar(
            x=mc["periodo"], y=mc["Leads"], text=mc["Leads"],
            textposition="outside",
            marker_color=mc["color"],
        ))
        fig3.update_layout(height=240, margin=dict(t=10,b=0,l=0,r=0),
                           plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                           xaxis=dict(gridcolor="#F1F5F9", title="",
                                      categoryorder="array",
                                      categoryarray=mc["periodo"].tolist()),
                           yaxis=dict(gridcolor="#F1F5F9", title=""))
        st.plotly_chart(fig3, use_container_width=True)

    st.divider()

    # ── 4. Distribución por Etapa ─────────────────────────────────────────────
    st.markdown('<p class="sec-title">Distribución por Etapa</p>', unsafe_allow_html=True)

    stage_counts = vdf_pipe["stage_key"].value_counts().to_dict() if not vdf_pipe.empty else {}

    # Calcular "Lead No Calificado": leads totales del período que NO están en pipeline
    _va_dist = _filtrar_fecha(vdf_all_proj_raw) if not vdf_all_proj_raw.empty else pd.DataFrame()
    _vt_dist = _filtrar_fecha(vdf_tags_raw)     if not vdf_tags_raw.empty     else pd.DataFrame()
    if not _va_dist.empty and not _vt_dist.empty:
        _pipeline_ids = set(_vt_dist["id"].values)
        _no_cal_df = _va_dist[~_va_dist["id"].isin(_pipeline_ids)].copy()
    elif not _va_dist.empty:
        _no_cal_df = _va_dist.copy()
    else:
        _no_cal_df = pd.DataFrame()
    _no_cal_cnt = len(_no_cal_df)

    # Total para % incluye no-calificados
    _tp_total = tp + _no_cal_cnt

    grupos = {
        "⚪ Lead No Calificado": ([], "#94A3B8"),   # especial: viene de _no_cal_df
        "🔵 Entrada":            (["nuevo","en_com","recontacto","no_cal"],          "#60A5FA"),
        "🟢 Calificados":        (["perfilado","perfilado_com","cal_perdido"],        "#34D399"),
        "🟣 Visitas":            (["vis_agendada","reagendamiento","no_visito","vis_concretada","seg_post"], "#818CF8"),
        "🟡 Cierre":             (["evaluando","seg_largo"],                          "#FBBF24"),
        "🏆 Ganados":            (["apartado","contrato"],                            "#EC4899"),
        "🔴 Perdidos":           (["descartado","apt_cancel"],                        "#F87171"),
    }

    if _tp_total == 0:
        st.info("No hay contactos con etapa conocida en este período.")
    else:
        _max_all = max(list(stage_counts.values()) + [_no_cal_cnt]) if stage_counts or _no_cal_cnt else 1

        for grupo_label, (etapas, grupo_color) in grupos.items():
            # Grupo especial: Lead No Calificado
            if grupo_label == "⚪ Lead No Calificado":
                if _no_cal_cnt == 0:
                    continue
                _pct_nc = div(_no_cal_cnt, _tp_total) * 100
                _bar_nc = div(_no_cal_cnt, _max_all) * 100
                with st.expander(
                    f"{grupo_label} — {_no_cal_cnt} contactos · {_pct_nc:.0f}% del total",
                    expanded=False
                ):
                    # Fila única (sin sub-etapas)
                    st.markdown(
                        f"<table style='width:100%;border-collapse:collapse;font-size:.85rem;"
                        f"margin-bottom:12px'><tr style='border-bottom:1px solid #F1F5F9'>"
                        f"<td style='padding:6px 4px;color:#64748B;font-size:.83rem;width:35%'>"
                        f"Lead No Calificado</td>"
                        f"<td style='text-align:right;padding:6px 10px;font-weight:700;"
                        f"color:#94A3B8;font-size:.92rem;width:8%'>{_no_cal_cnt}</td>"
                        f"<td style='padding:6px 6px;width:45%'>"
                        f"<div style='background:#F1F5F9;border-radius:4px;height:7px'>"
                        f"<div style='width:{_bar_nc:.0f}%;background:#94A3B8;height:7px;"
                        f"border-radius:4px'></div></div></td>"
                        f"<td style='text-align:right;padding:6px 8px;color:#94A3B8;"
                        f"font-size:.76rem;width:12%'>{_pct_nc:.1f}%</td>"
                        f"</tr></table>",
                        unsafe_allow_html=True
                    )
                    # Contactos
                    if _no_cal_df.empty:
                        st.caption("Sin datos de contactos disponibles.")
                    else:
                        _PROJ_C_NC = {"KOS":"#3B82F6","Punto Calma":"#10B981","Zen":"#F59E0B",
                                      "DODEKA":"#EF4444","SANTIÁN":"#8B5CF6","SANTIAN":"#8B5CF6"}
                        _nc_tbl = (
                            f"<div style='font-size:.78rem;font-weight:700;color:#94A3B8;"
                            f"margin-bottom:4px'>👥 {_no_cal_cnt} contactos</div>"
                            "<table style='width:100%;border-collapse:collapse;font-size:.82rem'>"
                            "<thead><tr style='border-bottom:1px solid #E2E8F0'>"
                            "<th style='padding:5px 8px;color:#64748B;text-align:left'>Nombre</th>"
                            "<th style='padding:5px 6px;color:#64748B;text-align:left'>Proyecto</th>"
                            "<th style='padding:5px 6px;color:#64748B;text-align:left'>Canal</th>"
                            "<th style='padding:5px 6px;color:#64748B;text-align:left'>Fecha entrada</th>"
                            "<th style='padding:5px 6px;color:#64748B;text-align:left'>Teléfono</th>"
                            "</tr></thead><tbody>"
                        )
                        for _, _ncr in _no_cal_df.sort_values(
                                "fecha_creacion", ascending=False).iterrows():
                            _pc_nc = _PROJ_C_NC.get(_ncr.get("proyecto",""), "#94A3B8")
                            _fec_nc = (pd.Timestamp(_ncr["fecha_creacion"]).strftime("%d/%m/%Y")
                                       if pd.notna(_ncr.get("fecha_creacion")) else "—")
                            _nc_tbl += (
                                f"<tr style='border-bottom:1px solid #F8FAFC'>"
                                f"<td style='padding:5px 8px;font-weight:600;color:#1E293B'>"
                                f"{_ncr.get('nombre','—')}</td>"
                                f"<td style='padding:5px 6px'><span style='background:{_pc_nc}22;"
                                f"color:{_pc_nc};border-radius:6px;padding:1px 7px;"
                                f"font-size:.75rem;font-weight:600'>"
                                f"{_ncr.get('proyecto') or 'Sin proyecto'}</span></td>"
                                f"<td style='padding:5px 6px;color:#64748B;font-size:.79rem'>"
                                f"{_ncr.get('canal') or '—'}</td>"
                                f"<td style='padding:5px 6px;color:#64748B;font-size:.78rem'>"
                                f"{_fec_nc}</td>"
                                f"<td style='padding:5px 6px;color:#94A3B8;font-size:.76rem'>"
                                f"{_ncr.get('telefono','—')}</td></tr>"
                            )
                        _nc_tbl += "</tbody></table>"
                        st.markdown(_nc_tbl, unsafe_allow_html=True)
                continue  # saltar el loop normal para este grupo

            # ── Grupos normales del pipeline ───────────────────────────────
            total_grupo = sum(stage_counts.get(sk, 0) for sk in etapas)
            if total_grupo == 0:
                continue
            pct_grupo = div(total_grupo, tp) * 100

            # ── Expander por grupo: clic para ver contactos ────────────────
            _exp_title = (
                f"{grupo_label} — "
                f"{total_grupo} contactos · {pct_grupo:.0f}% del pipeline"
            )
            with st.expander(_exp_title, expanded=False):
                # Mini-tabla de etapas del grupo
                _rows_g = ""
                for sk in etapas:
                    cnt   = stage_counts.get(sk, 0)
                    color = STAGE_COLOR.get(sk, "#94A3B8")
                    label = STAGE_LABEL.get(sk, sk)
                    pct_p = div(cnt, tp) * 100
                    bar_w = div(cnt, _max_all) * 100
                    _rows_g += (
                        f"<tr style='border-bottom:1px solid #F1F5F9'>"
                        f"<td style='padding:6px 10px 6px 4px;color:#475569;"
                        f"font-size:.83rem;width:35%'>{label}</td>"
                        f"<td style='text-align:right;padding:6px 10px;font-weight:700;"
                        f"color:{color};font-size:.92rem;width:8%'>{cnt}</td>"
                        f"<td style='padding:6px 6px;width:45%'>"
                        f"  <div style='background:#F1F5F9;border-radius:4px;height:7px'>"
                        f"    <div style='width:{bar_w:.0f}%;background:{color};"
                        f"height:7px;border-radius:4px'></div></div></td>"
                        f"<td style='text-align:right;padding:6px 8px;"
                        f"color:#94A3B8;font-size:.76rem;width:12%'>{pct_p:.1f}%</td>"
                        f"</tr>"
                    )
                st.markdown(
                    "<table style='width:100%;border-collapse:collapse;font-size:.85rem;"
                    "margin-bottom:12px'>" + _rows_g + "</table>",
                    unsafe_allow_html=True
                )

                # ── Contactos del grupo ────────────────────────────────────
                _src = vdf_tags_raw if not vdf_tags_raw.empty else vdf_pipe
                if _src.empty:
                    st.caption("Sin datos de contactos enriquecidos. Carga el funnel por proyecto primero.")
                else:
                    _gc = _src[_src["stage_key"].isin(etapas)].copy()
                    if _gc.empty:
                        st.caption("Sin contactos en este grupo para el período seleccionado.")
                    else:
                        # Cabecera de tabla de contactos
                        _ctbl = (
                            f"<div style='font-size:.78rem;font-weight:700;color:{grupo_color};"
                            f"margin-bottom:4px'>👥 {len(_gc)} contactos</div>"
                            "<table style='width:100%;border-collapse:collapse;font-size:.82rem'>"
                            "<thead><tr style='border-bottom:1px solid #E2E8F0'>"
                            "<th style='text-align:left;padding:5px 8px;color:#64748B'>Nombre</th>"
                            "<th style='text-align:left;padding:5px 6px;color:#64748B'>Proyecto</th>"
                            "<th style='text-align:left;padding:5px 6px;color:#64748B'>Canal</th>"
                            "<th style='text-align:left;padding:5px 6px;color:#64748B'>Etapa</th>"
                            "<th style='text-align:left;padding:5px 6px;color:#64748B'>Fecha entrada</th>"
                            "<th style='text-align:left;padding:5px 6px;color:#64748B'>Teléfono</th>"
                            "</tr></thead><tbody>"
                        )
                        _PROJ_C = {"KOS":"#3B82F6","Punto Calma":"#10B981","Zen":"#F59E0B",
                                   "DODEKA":"#EF4444","SANTIÁN":"#8B5CF6","SANTIAN":"#8B5CF6"}
                        for _, _cr in _gc.sort_values("fecha_creacion", ascending=False).iterrows():
                            _pc  = _PROJ_C.get(_cr.get("proyecto",""), "#94A3B8")
                            _sc  = STAGE_COLOR.get(_cr.get("stage_key",""), "#94A3B8")
                            _fec = (pd.Timestamp(_cr["fecha_creacion"]).strftime("%d/%m/%Y")
                                    if pd.notna(_cr.get("fecha_creacion")) else "—")
                            _proy = _cr.get("proyecto") or "Sin proyecto"
                            _canal= _cr.get("canal") or "—"
                            _slbl = _cr.get("stage_label") or _cr.get("stage_name","—")
                            _ctbl += (
                                f"<tr style='border-bottom:1px solid #F8FAFC'>"
                                f"<td style='padding:5px 8px;font-weight:600;color:#1E293B'>"
                                f"{_cr.get('nombre','—')}</td>"
                                f"<td style='padding:5px 6px'>"
                                f"<span style='background:{_pc}22;color:{_pc};border-radius:6px;"
                                f"padding:1px 7px;font-size:.75rem;font-weight:600'>{_proy}</span></td>"
                                f"<td style='padding:5px 6px;color:#64748B;font-size:.8rem'>{_canal}</td>"
                                f"<td style='padding:5px 6px'>"
                                f"<span style='color:{_sc};font-size:.78rem;font-weight:600'>{_slbl}</span></td>"
                                f"<td style='padding:5px 6px;color:#64748B;font-size:.78rem'>{_fec}</td>"
                                f"<td style='padding:5px 6px;color:#94A3B8;font-size:.75rem'>"
                                f"{_cr.get('telefono','—')}</td>"
                                f"</tr>"
                            )
                        _ctbl += "</tbody></table>"
                        st.markdown(_ctbl, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: EMBUDO
# ══════════════════════════════════════════════════════════════════════════════
elif "Embudo" in seccion:
    st.markdown("## 🔻 Embudo de Conversión")

    # ── Filtro de tiempo (mismo que Vambe Live) ───────────────────────────────
    _enow = _dt.datetime.now()
    _efa, _efb, _efc = st.columns([2, 2, 4])
    with _efa:
        _emb_modo = st.radio("Corte", ["Relativo", "Por mes"], horizontal=True,
                             label_visibility="collapsed", key="emb_modo")

    if _emb_modo == "Relativo":
        with _efb:
            _erel_opts = {"Últimos 30 días": 1, "Últimos 60 días": 2,
                          "Últimos 90 días": 3, "Todo el año": 12}
            _erel_sel = st.selectbox("", list(_erel_opts.keys()), index=0,
                                     label_visibility="collapsed", key="emb_rel")
        _en_meses = _erel_opts[_erel_sel]
        _ecutoff = _enow.year * 100 + _enow.month - _en_meses + 1
        _df_emb = df_raw.copy() if not df_raw.empty else pd.DataFrame()
        if not _df_emb.empty:
            _df_emb["_ps"] = _df_emb["año_num"].fillna(0).astype(int)*100 + _df_emb["mes_num"].fillna(0).astype(int)
            _df_emb = _df_emb[_df_emb["_ps"] >= _ecutoff].drop(columns=["_ps"])
        _elbl = _erel_sel
    else:
        _ebase = df_raw if not df_raw.empty else pd.DataFrame()
        _emes_opts, _emes_keys = [], []
        if not _ebase.empty:
            _emd = (_ebase.dropna(subset=["mes_num","año_num"])
                    .assign(sort=lambda d: d["año_num"].astype(int)*100+d["mes_num"].astype(int))
                    .drop_duplicates("sort").sort_values("sort", ascending=False))
            for _, _er in _emd.iterrows():
                _emes_opts.append(f"{MONTH_FULL.get(int(_er.mes_num), str(int(_er.mes_num)))} {int(_er.año_num)}")
                _emes_keys.append((int(_er.mes_num), int(_er.año_num)))
        with _efb:
            _eidx = st.selectbox("Mes", range(len(_emes_opts)),
                                 format_func=lambda i: _emes_opts[i] if _emes_opts else "—",
                                 label_visibility="collapsed", key="emb_mes")
        if _emes_opts:
            _esm, _esy = _emes_keys[_eidx]
            _df_emb = df_raw[(df_raw["mes_num"].fillna(0).astype(int)==_esm) &
                              (df_raw["año_num"].fillna(0).astype(int)==_esy)].copy()
            _elbl = _emes_opts[_eidx]
        else:
            _df_emb = df_raw.copy() if not df_raw.empty else pd.DataFrame()
            _elbl = "Sin datos"

    # Aplicar filtro de proyectos del sidebar
    if not _df_emb.empty and sel_projs:
        _df_emb = _df_emb[_df_emb["proyecto"].isin(sel_projs)]

    # Período info y re-agregación local
    st.caption(f"Período: **{_elbl}**")

    # Sombrea las variables globales para que el resto de la sección use los datos filtrados
    df   = _df_emb
    # Leads = leadsVambe + recontactos; Cal = prospectosCalificados (col M)
    leads = (scol(df, "leadsVambe") + scol(df, "recontactos")) if not df.empty else 0
    cal   = (scol(df, "prospectosCalificados") if "prospectosCalificados" in df.columns
             else scol(df, "leadsCalificados")) if not df.empty else 0
    agen  = scol(df, "visitaAgendada")  if not df.empty else 0
    conc  = scol(df, "visitaConcretada") if not df.empty else 0
    apar  = scol(df, "apartados")       if not df.empty else 0
    inv   = scol(df, "invertido")       if not df.empty else 0

    # ── 1. Funnel de Conversión por Proyecto ──────────────────────────────────
    st.markdown('<p class="sec-title">Funnel de Conversión por Proyecto</p>',
                unsafe_allow_html=True)
    st.caption("Tasas de conversión entre etapas del funnel según los filtros activos")

    if df.empty:
        st.info("Sin datos del sheet para mostrar la tabla.")
    else:
        _cols_p = ["leadsVambe","recontactos","prospectosCalificados",
                   "visitaAgendada","visitaConcretada","apartados","invertido"]
        # Solo incluir columnas que existan en el df
        _cols_p = [c for c in _cols_p if c in df.columns]
        _grp = df.groupby("proyecto")[_cols_p].sum().reset_index()
        _grp = _grp[_grp["leadsVambe"] > 0].copy()
        _grp["_leads_total"] = _grp["leadsVambe"] + _grp.get("recontactos", 0)
        _tot = _grp[[c for c in _cols_p] + ["_leads_total"]].sum()
        _tot["proyecto"] = "TOTAL"
        _grp = pd.concat([_grp, _tot.to_frame().T], ignore_index=True)

        _sheet_rows = [
            {"nombre": r["proyecto"],
             "leads":  r.get("_leads_total", r["leadsVambe"]),
             "recont": r.get("recontactos", 0),
             "cal":    r.get("prospectosCalificados", r.get("leadsCalificados", 0)),
             "agen":   r.get("visitaAgendada", 0),
             "conc":   r["visitaConcretada"],
             "apar":   r["apartados"],
             "inv":    r["invertido"],
             "is_total": r["proyecto"] == "TOTAL"}
            for _, r in _grp.iterrows()
        ]
        st.markdown(_funnel_table_html(_sheet_rows, show_invertido=True, show_recont=True),
                    unsafe_allow_html=True)

    st.divider()

    # ── 2. Embudo acumulado ───────────────────────────────────────────────────
    c1, c2 = st.columns([3,2])
    with c1:
        stages_f = [
            ("Leads",          leads, "#64748B"),
            ("Calificados",    cal,   "#3B82F6"),
            ("Vis. Agendadas", agen,  "#8B5CF6"),
            ("Vis. Concretadas",conc, "#10B981"),
            ("Apartados",      apar,  "#EC4899"),
        ]
        fig = go.Figure(go.Funnel(
            y=[s[0] for s in stages_f], x=[s[1] for s in stages_f],
            textinfo="value+percent initial",
            marker=dict(color=[s[2] for s in stages_f]),
        ))
        fig.update_layout(height=400, margin=dict(t=20,b=20,l=20,r=20),
                          paper_bgcolor="rgba(0,0,0,0)", font=dict(size=14))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown('<p class="sec-title">Conversión por paso</p>', unsafe_allow_html=True)
        step_convs = [
            ("Leads → Calificados",  leads, cal),
            ("Cal. → Agendadas",     cal,   agen),
            ("Agendadas → Concretadas", agen, conc),
            ("Concretadas → Apartados", conc, apar),
        ]
        html = ""
        for label,base,val in step_convs:
            html += conv_bar(label, base, val, cumulative=False)
        st.markdown(html, unsafe_allow_html=True)

        st.markdown('<p class="sec-title" style="margin-top:16px">Conversión acumulada desde Lead</p>',
                    unsafe_allow_html=True)
        cum_convs = [
            ("Lead → Calificado",    leads, cal),
            ("Lead → Vis. Agendada", leads, agen),
            ("Lead → Vis. Concretada",leads, conc),
            ("Lead → Apartado",      leads, apar),
        ]
        html2 = ""
        for label,base,val in cum_convs:
            html2 += conv_bar(label, base, val, cumulative=True)
        st.markdown(html2, unsafe_allow_html=True)

    st.divider()

    # ── 3. Tabla: Paso vs Acumulado ───────────────────────────────────────────
    st.markdown('<p class="sec-title">Tabla: Paso vs Acumulado</p>', unsafe_allow_html=True)
    tabla = pd.DataFrame([
        {"Etapa":"Leads",           "Contactos":f"{leads:,.0f}", "Conv. Paso":"—",
         "Conv. Acum. (desde Lead)":"100%"},
        {"Etapa":"Calificados",     "Contactos":f"{cal:,.0f}",
         "Conv. Paso":pct(div(cal,leads)*100),
         "Conv. Acum. (desde Lead)":pct(div(cal,leads)*100)},
        {"Etapa":"Visita Agendada", "Contactos":f"{agen:,.0f}",
         "Conv. Paso":pct(div(agen,cal)*100),
         "Conv. Acum. (desde Lead)":pct(div(agen,leads)*100)},
        {"Etapa":"Visita Concretada","Contactos":f"{conc:,.0f}",
         "Conv. Paso":pct(div(conc,agen)*100),
         "Conv. Acum. (desde Lead)":pct(div(conc,leads)*100)},
        {"Etapa":"Apartados",       "Contactos":f"{apar:,.0f}",
         "Conv. Paso":pct(div(apar,conc)*100),
         "Conv. Acum. (desde Lead)":pct(div(apar,leads)*100)},
    ])
    st.dataframe(tabla.set_index("Etapa"), use_container_width=True, height=220)

    st.divider()

    # ── 4. Distribución por Proyecto (gráfica) ────────────────────────────────
    if not df.empty:
        st.markdown('<p class="sec-title">Embudo por Proyecto</p>', unsafe_allow_html=True)
        pf = df.groupby("proyecto").agg(
            Leads=("leadsVambe","sum"), Calificados=("leadsCalificados","sum"),
            Agendadas=("visitaAgendada","sum"), Concretadas=("visitaConcretada","sum"),
            Apartados=("apartados","sum")).reset_index()
        fig2 = go.Figure()
        for _, row in pf.iterrows():
            c = PROJ_COLORS.get(row.proyecto, DEFAULT_COLOR)
            fig2.add_trace(go.Bar(name=row.proyecto,
                x=["Leads","Calificados","Agendadas","Concretadas","Apartados"],
                y=[row.Leads,row.Calificados,row.Agendadas,row.Concretadas,row.Apartados],
                marker_color=c))
        fig2.update_layout(barmode="group",height=300,margin=dict(t=10,b=10,l=0,r=0),
                           legend=dict(orientation="h",y=1.1),
                           plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                           xaxis=dict(gridcolor="#F1F5F9"),yaxis=dict(gridcolor="#F1F5F9"))
        st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: METAS
# ══════════════════════════════════════════════════════════════════════════════
elif "Metas" in seccion:
    _mdata = _metas_load()
    _today = _dt.date.today()
    _META_YEAR = 2026

    # Selector de trimestre
    _q_sel = st.radio("Trimestre", list(_QUARTERS.keys()), horizontal=True,
                      key="metas_quarter")
    _Q_ACTIVE = _QUARTERS[_q_sel]
    _Q_LABEL  = _q_sel  # e.g. "Q1 2026"

    st.markdown(f"## 🎯 Metas Trimestrales {_Q_LABEL}")

    # ── Helper: celda de progreso ─────────────────────────────────────────────
    def _meta_cell(real, meta, factor, is_money=False):
        if meta == 0:
            return "<td style='text-align:center;padding:10px 8px;color:#94A3B8'>—</td>"
        esp = meta * factor
        pct_v = (real / esp * 100) if esp > 0 else 0
        if pct_v >= 90:
            bg, fg, arrow = "#D1FAE5", "#065F46", "▲"
        elif pct_v >= 70:
            bg, fg, arrow = "#FEF3C7", "#92400E", "→"
        else:
            bg, fg, arrow = "#FEE2E2", "#991B1B", "▼"
        fmt  = (lambda v: f"${v:,.0f}") if is_money else (lambda v: f"{v:,.0f}")
        return f"""<td style='text-align:center;padding:10px 8px'>
          <div style='font-size:1.15rem;font-weight:700;color:#1E293B'>{fmt(real)}</div>
          <div style='display:inline-block;background:{bg};color:{fg};border-radius:12px;
                      padding:2px 8px;font-size:.72rem;font-weight:700;margin-top:3px'>
            {arrow} {pct_v:.0f}% · esp {fmt(esp)}
          </div>
        </td>"""

    # ── 1. Como Vamos ─────────────────────────────────────────────────────────
    st.markdown('<p class="sec-title">Como Vamos</p>', unsafe_allow_html=True)

    # ── Helpers de reales (con pd.to_numeric para evitar concatenación de strings) ──
    def _to_int(series):
        return int(pd.to_numeric(series, errors="coerce").fillna(0).sum())

    def _rv_months(p, mnums):
        """Reales Vambe para uno o varios meses."""
        _leads = 0
        if not vdf_all_proj_raw.empty:
            _mask = ((vdf_all_proj_raw["mes_num"].isin(mnums)) &
                     (vdf_all_proj_raw["año_num"] == _META_YEAR) &
                     (vdf_all_proj_raw["proyecto"] == p))
            _leads = int(_mask.sum())
        _fn = vambe_funnel(pd.DataFrame())
        if not vdf_tags_raw.empty:
            _s = vdf_tags_raw[(vdf_tags_raw["mes_num"].isin(mnums)) &
                               (vdf_tags_raw["año_num"] == _META_YEAR) &
                               (vdf_tags_raw["proyecto"] == p)]
            _fn = vambe_funnel(_s)
        return {"leads": _leads, "cal": _fn["calificados"], "agen": _fn["vis_agendada"],
                "conc": _fn["vis_concretada"], "apar": _fn["apartados"], "pres": None}

    def _re_months(p, mnums):
        """Reales Excel para uno o varios meses.
        Leads = leadsVambe + recontactos | Cal = prospectosCalificados (col M)."""
        _empty = {"leads": 0, "cal": 0, "agen": 0, "conc": 0, "apar": 0, "pres": 0.0}
        if df_raw.empty: return _empty
        _s = df_raw[(df_raw["mes_num"].fillna(0).astype(int).isin(mnums)) &
                    (df_raw["año_num"].fillna(0).astype(int) == _META_YEAR) &
                    (df_raw["proyecto"] == p)]
        if _s.empty: return _empty
        _lv  = _to_int(_s["leadsVambe"])           if "leadsVambe"           in _s.columns else 0
        _rc  = _to_int(_s["recontactos"])           if "recontactos"          in _s.columns else 0
        _cal = (_to_int(_s["prospectosCalificados"]) if "prospectosCalificados" in _s.columns
                else _to_int(_s["leadsCalificados"]) if "leadsCalificados"      in _s.columns else 0)
        return {
            "leads": _lv + _rc, "cal": _cal,
            "agen":  _to_int(_s["visitaAgendada"])  if "visitaAgendada"  in _s.columns else 0,
            "conc":  _to_int(_s["visitaConcretada"]) if "visitaConcretada" in _s.columns else 0,
            "apar":  _to_int(_s["apartados"])        if "apartados"        in _s.columns else 0,
            "pres":  float(pd.to_numeric(_s["invertido"], errors="coerce").fillna(0).sum())
                     if "invertido" in _s.columns else 0.0,
        }

    def _meta_months(proj, key, mnums):
        return sum(_meta_get(_mdata, _META_YEAR, m, proj, key) for m in mnums)

    def _cv_render_table(mnums, factor, flbl):
        """Renderiza tabla Como Vamos para los meses dados con factor de proración."""
        st.caption(flbl)
        if factor == 0:
            st.info("Todavía no inicia este período.")
            return

        _th = "".join(
            f"<th style='text-align:center;padding:8px;color:#64748B;font-weight:600;"
            f"font-size:.82rem'>{c}</th>" for c in _METAS_COLS
        )
        _html = (
            "<table style='width:100%;border-collapse:collapse;background:white;"
            "border-radius:12px;box-shadow:0 1px 6px rgba(0,0,0,.07);overflow:hidden'>"
            "<thead><tr style='border-bottom:2px solid #E2E8F0'>"
            "<th style='text-align:left;padding:10px 14px;color:#64748B;font-weight:600;"
            "font-size:.82rem'>Proyecto / Fuente</th>"
            f"{_th}</tr></thead><tbody>"
        )

        # Acumuladores para fila TOTAL
        _acc_v = {k: 0   for k in _METAS_KEYS}
        _acc_e = {k: 0.0 for k in _METAS_KEYS}
        _acc_m = {k: 0   for k in _METAS_KEYS}

        for _proj in _METAS_PROJS:
            _has_meta = any(_meta_months(_proj, k, mnums) > 0 for k in _METAS_KEYS)
            _html += (
                f"<tr style='background:#F8FAFC;border-top:2px solid #E2E8F0'>"
                f"<td colspan='7' style='padding:6px 14px;font-weight:700;color:#1E293B;font-size:.9rem'>"
                f"{_proj}"
                f"{'<span style=\"font-weight:400;color:#94A3B8;font-size:.75rem;margin-left:8px\">sin meta</span>' if not _has_meta else ''}"
                f"</td></tr>"
            )
            _vr = _rv_months(_proj, mnums)
            _vcells = "".join(
                ("<td style='text-align:center;padding:7px 8px;color:#94A3B8;font-size:.82rem'>—</td>"
                 if _vr[k] is None else
                 _meta_cell(_vr[k], _meta_months(_proj, k, mnums), factor, is_money=(k == "pres")))
                for k in _METAS_KEYS
            )
            _html += (f"<tr style='border-bottom:1px solid #F1F5F9'>"
                      f"<td style='padding:6px 14px 6px 22px;color:#3B82F6;font-size:.8rem;font-weight:600'>"
                      f"🔵 Vambe</td>{_vcells}</tr>")

            _er = _re_months(_proj, mnums)
            _ecells = "".join(
                _meta_cell(_er[k], _meta_months(_proj, k, mnums), factor, is_money=(k == "pres"))
                for k in _METAS_KEYS
            )
            _html += (f"<tr style='border-bottom:1px solid #F1F5F9'>"
                      f"<td style='padding:6px 14px 6px 22px;color:#10B981;font-size:.8rem;font-weight:600'>"
                      f"🟢 Excel</td>{_ecells}</tr>")

            # Acumular totales
            for k in _METAS_KEYS:
                if _vr[k] is not None: _acc_v[k] += _vr[k]
                _acc_e[k] += _er[k]
                _acc_m[k] += _meta_months(_proj, k, mnums)

        # ── Fila TOTAL ────────────────────────────────────────────────────
        _html += (
            "<tr style='background:#1E293B;border-top:3px solid #334155'>"
            "<td colspan='7' style='padding:7px 14px;font-weight:700;color:white;"
            "font-size:.9rem;letter-spacing:.03em'>TOTAL</td></tr>"
        )
        _vcells_tot = "".join(
            ("<td style='text-align:center;padding:7px 8px;color:#94A3B8;font-size:.82rem'>—</td>"
             if _acc_v[k] == 0 else
             _meta_cell(_acc_v[k], _acc_m[k], factor, is_money=(k == "pres")))
            for k in _METAS_KEYS
        )
        _html += (f"<tr style='border-bottom:1px solid #334155;background:#EFF6FF'>"
                  f"<td style='padding:6px 14px 6px 22px;color:#3B82F6;font-size:.8rem;font-weight:700'>"
                  f"🔵 Vambe</td>{_vcells_tot}</tr>")
        _ecells_tot = "".join(
            _meta_cell(_acc_e[k], _acc_m[k], factor, is_money=(k == "pres"))
            for k in _METAS_KEYS
        )
        _html += (f"<tr style='border-bottom:1px solid #334155;background:#F0FDF4'>"
                  f"<td style='padding:6px 14px 6px 22px;color:#10B981;font-size:.8rem;font-weight:700'>"
                  f"🟢 Excel</td>{_ecells_tot}</tr>")

        _html += "</tbody></table>"
        st.markdown(_html, unsafe_allow_html=True)

    # ── Tabs: mes más reciente primero + Resumen trimestre ───────────────────
    _q_months = list(_Q_ACTIVE.items())  # e.g. [(1,"Enero"),(2,"Febrero"),(3,"Marzo")]
    _months_ordered = sorted(
        _q_months,
        key=lambda x: -(x[0] if (_META_YEAR < _today.year or x[0] <= _today.month) else 99)
    )  # Mes más reciente del trimestre primero
    _tab_labels = [mname for _, mname in _months_ordered] + [f"📊 Resumen {_Q_LABEL}"]
    _all_cv_tabs = st.tabs(_tab_labels)

    # Tabs mensuales
    for (_mnum, _mname), _tab in zip(_months_ordered, _all_cv_tabs[:-1]):
        with _tab:
            _days_in = _cal.monthrange(_META_YEAR, _mnum)[1]
            if (_META_YEAR, _mnum) < (_today.year, _today.month):
                _fac, _flbl = 1.0, "Mes completo"
            elif (_META_YEAR, _mnum) == (_today.year, _today.month):
                _fac = _today.day / _days_in
                _flbl = f"Día {_today.day} de {_days_in} — {_fac*100:.0f}% del mes transcurrido"
            else:
                _fac, _flbl = 0.0, "Mes futuro — sin comparación disponible"
            _cv_render_table([_mnum], _fac, _flbl)

    # Tab Resumen trimestre
    with _all_cv_tabs[-1]:
        _q_all_months = [m for m, _ in _q_months]
        _q_total_days = sum(_cal.monthrange(_META_YEAR, m)[1] for m in _q_all_months)
        if _today.year == _META_YEAR and _today.month <= max(_q_all_months):
            _q_elapsed = sum(
                _cal.monthrange(_META_YEAR, m)[1]
                for m in _q_all_months if m < _today.month
            ) + (_today.day if _today.month in _q_all_months else 0)
        else:
            _q_elapsed = _q_total_days
        _q_factor = min(_q_elapsed / _q_total_days, 1.0)
        _q_flbl = f"{_Q_LABEL}: {_q_elapsed} de {_q_total_days} días — {_q_factor*100:.0f}% del trimestre transcurrido"
        _cv_render_table(_q_all_months, _q_factor, _q_flbl)

    st.divider()

    # ── 2. Configurar Metas ───────────────────────────────────────────────────
    st.markdown('<p class="sec-title">Configurar Metas</p>', unsafe_allow_html=True)

    # Inicializar session_state desde el archivo si la key no existe todavía
    # (evita que Streamlit ignore el value= cuando ya hay un key en session_state)
    for _im in _Q_ACTIVE:
        for _ip in _METAS_PROJS:
            for _ik in _METAS_KEYS:
                _sk = f"m_{_im}_{_ip}_{_ik}"
                if _sk not in st.session_state:
                    st.session_state[_sk] = _meta_get(_mdata, _META_YEAR, _im, _ip, _ik)

    _cfg_tabs = st.tabs(list(_Q_ACTIVE.values()))
    # Preservar meses de todos los trimestres al guardar
    _new_mdata = {str(_META_YEAR): {str(m): _mdata.get(str(_META_YEAR), {}).get(str(m), {p: {} for p in _METAS_PROJS})
                  for m in range(1, 13)}}
    # Inicializar proyectos faltantes
    for _nm in range(1, 13):
        for _np in _METAS_PROJS:
            if _np not in _new_mdata[str(_META_YEAR)][str(_nm)]:
                _new_mdata[str(_META_YEAR)][str(_nm)][_np] = {}

    for _ctab, (_cmnum, _cmname) in zip(_cfg_tabs, _Q_ACTIVE.items()):
        with _ctab:
            _hcols = st.columns([2, 1, 1, 1, 1, 1, 1])
            for _ci, _clbl in enumerate(["Proyecto"] + _METAS_COLS):
                _hcols[_ci].markdown(
                    f"<span style='font-size:.78rem;color:#64748B;font-weight:600'>{_clbl}</span>",
                    unsafe_allow_html=True)

            for _proj in _METAS_PROJS:
                _pcols = st.columns([2, 1, 1, 1, 1, 1, 1])
                _pcols[0].markdown(f"**{_proj}**")
                for _ci, _key in enumerate(_METAS_KEYS):
                    _step = 1000 if _key == "pres" else (10 if _key == "leads" else 1)
                    _v = _pcols[_ci + 1].number_input(
                        "x", step=_step, min_value=0,
                        label_visibility="collapsed",
                        key=f"m_{_cmnum}_{_proj}_{_key}")
                    _new_mdata[str(_META_YEAR)][str(_cmnum)][_proj][_key] = _v

            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    if st.button("💾 Guardar Metas", type="primary", use_container_width=True):
        _metas_save(_new_mdata)
        # Limpiar session_state para que al recargar lea los valores del archivo
        for _im in _Q_ACTIVE:
            for _ip in _METAS_PROJS:
                for _ik in _METAS_KEYS:
                    st.session_state.pop(f"m_{_im}_{_ip}_{_ik}", None)
        st.success("✅ Metas guardadas.")
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: VISITAS
# ══════════════════════════════════════════════════════════════════════════════
elif "Visitas" in seccion:
    import calendar as _vis_cal
    st.markdown("## 🗓️ Visitas del Mes")

    # ── Selector de mes ───────────────────────────────────────────────────────
    _vis_now = _dt.datetime.now()
    if "vis_year"  not in st.session_state: st.session_state["vis_year"]  = _vis_now.year
    if "vis_month" not in st.session_state: st.session_state["vis_month"] = _vis_now.month

    _vc1, _vc2, _vc3 = st.columns([1, 2, 1])
    with _vc1:
        if st.button("◀", key="vis_prev"):
            if st.session_state["vis_month"] == 1:
                st.session_state["vis_month"] = 12
                st.session_state["vis_year"] -= 1
            else:
                st.session_state["vis_month"] -= 1
            st.rerun()
    with _vc2:
        _MONTH_NAMES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                           7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
        st.markdown(
            f"<div style='text-align:center;font-size:1.1rem;font-weight:700;padding:6px'>"
            f"{_MONTH_NAMES_ES[st.session_state['vis_month']]} {st.session_state['vis_year']}</div>",
            unsafe_allow_html=True)
    with _vc3:
        if st.button("▶", key="vis_next"):
            if st.session_state["vis_month"] == 12:
                st.session_state["vis_month"] = 1
                st.session_state["vis_year"] += 1
            else:
                st.session_state["vis_month"] += 1
            st.rerun()

    _vis_y = st.session_state["vis_year"]
    _vis_m = st.session_state["vis_month"]

    # ── Cargar datos ──────────────────────────────────────────────────────────
    @st.cache_data(ttl=300)
    def _load_visits(y, m): return load_visits_month(y, m)

    _vdf, _verr = _load_visits(_vis_y, _vis_m)

    if _verr:
        st.error(f"Error cargando visitas: {_verr}")
    elif _vdf.empty:
        st.info("No hay visitas registradas para este mes.")
    else:
        _now = _dt.datetime.now(_dt.timezone.utc)

        # ── KPI chips ────────────────────────────────────────────────────────
        _total      = len(_vdf)
        _cnt = {s: int((_vdf["status"] == s).sum()) for s in VISIT_STATUS_LABELS}
        _positivas  = _cnt["concretada"] + _cnt["activo"] + _cnt["ganado"]
        _pct_pos    = round(_positivas / _total * 100) if _total else 0

        # Fila de chips: Agendadas + 7 estados
        _chip_style = lambda bg, fg: (
            f"display:inline-block;background:{bg};color:{fg};"
            f"border-radius:12px;padding:8px 16px;text-align:center;"
            f"min-width:90px;margin:4px;"
        )
        _chips_html = (
            "<div style='display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px'>"
            f"<div style='{_chip_style('#EFF6FF','#1D4ED8')}'>"
            f"<div style='font-size:1.5rem;font-weight:800'>{_total}</div>"
            f"<div style='font-size:.75rem;font-weight:600'>📅 Agendadas</div></div>"
        )
        for _sk, (_slbl, _scol, _sico) in VISIT_STATUS_LABELS.items():
            _n = _cnt[_sk]
            _bg = _scol + "22"  # ~14% opacity background
            _chips_html += (
                f"<div style='{_chip_style(_bg, _scol)}'>"
                f"<div style='font-size:1.5rem;font-weight:800;color:{_scol}'>{_n}</div>"
                f"<div style='font-size:.75rem;font-weight:600'>{_sico} {_slbl}</div></div>"
            )
        _chips_html += "</div>"
        st.markdown(_chips_html, unsafe_allow_html=True)

        # ── Barra de progreso ─────────────────────────────────────────────
        _bar_html = (
            f"<div style='margin-bottom:20px'>"
            f"<div style='font-size:.82rem;font-weight:600;color:#374151;margin-bottom:4px'>"
            f"Progreso de {_MONTH_NAMES_ES[_vis_m]} {_vis_y} &nbsp;"
            f"<span style='color:#6B7280;font-weight:400'>"
            f"{_positivas} de {_total} visitas positivas ({_pct_pos}%)</span></div>"
            f"<div style='background:#E5E7EB;border-radius:8px;height:10px;overflow:hidden'>"
            f"<div style='width:{_pct_pos}%;background:linear-gradient(90deg,#3B82F6,#10B981);"
            f"height:100%;border-radius:8px;transition:width .4s'></div></div></div>"
        )
        st.markdown(_bar_html, unsafe_allow_html=True)

        # ── Tabla de visitas ──────────────────────────────────────────────
        st.markdown("**Todas las Visitas**")

        _PROJ_COLORS_VIS = {
            "KOS":"#3B82F6","Punto Calma":"#10B981","Zen":"#F59E0B",
            "DODEKA":"#EF4444","SANTIÁN":"#8B5CF6","SANTIAN":"#8B5CF6",
        }

        _tbl = (
            "<table style='width:100%;border-collapse:collapse;font-size:.87rem'>"
            "<thead><tr style='border-bottom:2px solid #E2E8F0'>"
            "<th style='text-align:left;padding:8px 10px;color:#64748B'>Contacto</th>"
            "<th style='text-align:left;padding:8px 6px;color:#64748B'>Proyecto</th>"
            "<th style='text-align:center;padding:8px 6px;color:#64748B'>Fecha Visita</th>"
            "<th style='text-align:center;padding:8px 6px;color:#64748B'>Estado</th>"
            "</tr></thead><tbody>"
        )

        for _, _row in _vdf.iterrows():
            _slbl, _scol, _sico = VISIT_STATUS_LABELS.get(_row["status"], ("—", "#94A3B8", ""))
            _pc = _PROJ_COLORS_VIS.get(_row["proyecto"], "#64748B")
            _vd = _row["visit_date"]
            if _vd is not None:
                try:
                    _vd_fmt = pd.Timestamp(_vd).strftime("%a %d de %b").lower()
                except Exception:
                    _vd_fmt = str(_vd)[:10]
            else:
                _vd_fmt = "—"

            # Fecha en rojo si es vencida
            _date_color = "#EF4444" if _row["status"] == "vencida" else "#374151"

            _tbl += (
                f"<tr style='border-bottom:1px solid #F1F5F9'>"
                f"<td style='padding:8px 10px'>"
                f"<div style='font-weight:600;color:#1E293B'>{_row['nombre']}</div>"
                f"<div style='font-size:.75rem;color:#94A3B8'>{_row['telefono']}</div></td>"
                f"<td style='padding:8px 6px'>"
                f"<span style='background:{_pc}22;color:{_pc};border-radius:8px;"
                f"padding:2px 8px;font-size:.78rem;font-weight:600'>{_row['proyecto']}</span></td>"
                f"<td style='text-align:center;padding:8px 6px;color:{_date_color};font-weight:600'>"
                f"{_vd_fmt}</td>"
                f"<td style='text-align:center;padding:8px 6px'>"
                f"<span style='background:{_scol}22;color:{_scol};border-radius:10px;"
                f"padding:3px 10px;font-size:.78rem;font-weight:700'>{_sico} {_slbl}</span></td>"
                f"</tr>"
            )
        _tbl += "</tbody></table>"
        st.markdown(_tbl, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: INVERSIÓN
# ══════════════════════════════════════════════════════════════════════════════
elif "Inversión" in seccion:
    st.markdown("## 💰 Análisis de Inversión")

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("💵 Total Invertido",  money(inv))
    k2.metric("💲 CPL",              money(cpl_val), delta="costo por lead",       delta_color="off")
    k3.metric("📌 CPP",              money(cpp_val), delta="costo por calificado", delta_color="off")
    k4.metric("💳 CAC",              money(cac_val), delta="costo por apartado",   delta_color="off")

    st.divider()

    if not df.empty:
        c1,c2 = st.columns(2)
        with c1:
            st.markdown('<p class="sec-title">Inversión por Proyecto</p>', unsafe_allow_html=True)
            ip=df.groupby("proyecto")["invertido"].sum().reset_index().sort_values("invertido")
            ip["color"]=ip["proyecto"].map(PROJ_COLORS).fillna(DEFAULT_COLOR)
            ip["label"]=ip["invertido"].apply(money)
            fig=go.Figure(go.Bar(y=ip["proyecto"],x=ip["invertido"],orientation="h",
                                 marker_color=ip["color"],text=ip["label"],textposition="outside"))
            fig.update_layout(height=300,margin=dict(t=10,b=10,l=0,r=80),
                              plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                              xaxis=dict(showticklabels=False),yaxis=dict(gridcolor="#F1F5F9"))
            st.plotly_chart(fig,use_container_width=True)
        with c2:
            st.markdown('<p class="sec-title">CAC por Proyecto</p>', unsafe_allow_html=True)
            eff=df.groupby("proyecto").agg(inv=("invertido","sum"),apt=("apartados","sum")).reset_index()
            eff["cac_v"]=eff.apply(lambda r: div(r.inv,r.apt),axis=1)
            eff=eff[eff["cac_v"]<2_000_000].sort_values("cac_v")
            eff["color"]=eff["proyecto"].map(PROJ_COLORS).fillna(DEFAULT_COLOR)
            eff["label"]=eff["cac_v"].apply(money)
            fig2=go.Figure(go.Bar(y=eff["proyecto"],x=eff["cac_v"],orientation="h",
                                  marker_color=eff["color"],text=eff["label"],textposition="outside"))
            fig2.update_layout(height=300,margin=dict(t=10,b=10,l=0,r=80),
                               plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                               xaxis=dict(showticklabels=False,title="menos = mejor"),
                               yaxis=dict(gridcolor="#F1F5F9"))
            st.plotly_chart(fig2,use_container_width=True)

        st.divider()
        st.markdown('<p class="sec-title">Inversión Mensual por Proyecto</p>', unsafe_allow_html=True)
        im=(df.groupby(["periodo_sort","periodo","proyecto"])["invertido"]
              .sum().reset_index().sort_values("periodo_sort"))
        orden=im.drop_duplicates("periodo_sort").sort_values("periodo_sort")["periodo"].tolist()
        fig3=px.bar(im,x="periodo",y="invertido",color="proyecto",
                    color_discrete_map=PROJ_COLORS,barmode="stack",
                    labels={"invertido":"Inversión","proyecto":"","periodo":""})
        fig3.update_layout(height=300,margin=dict(t=10,b=10,l=0,r=0),
                           legend=dict(orientation="h",y=1.08),
                           plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                           xaxis=dict(gridcolor="#F1F5F9",categoryorder="array",categoryarray=orden),
                           yaxis=dict(gridcolor="#F1F5F9"))
        st.plotly_chart(fig3,use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: PROYECTOS
# ══════════════════════════════════════════════════════════════════════════════
elif "Proyectos" in seccion:
    st.markdown("## 🏗️ Rendimiento por Proyecto")

    if df.empty:
        st.warning("Sin datos con los filtros actuales.")
    else:
        ps=df.groupby("proyecto").agg(
            Invertido=("invertido","sum"), Leads=("leadsVambe","sum"),
            Calificados=("leadsCalificados","sum"), Agendadas=("visitaAgendada","sum"),
            Concretadas=("visitaConcretada","sum"), Apartados=("apartados","sum")).reset_index()
        # conversiones paso a paso
        ps["pLead_Cal"] =ps.apply(lambda r: div(r.Calificados,r.Leads)*100,    axis=1)
        ps["pCal_Agen"] =ps.apply(lambda r: div(r.Agendadas,r.Calificados)*100,axis=1)
        ps["pAgen_Conc"]=ps.apply(lambda r: div(r.Concretadas,r.Agendadas)*100,axis=1)
        ps["pConc_Apt"] =ps.apply(lambda r: div(r.Apartados,r.Concretadas)*100,axis=1)
        # conversiones acumuladas desde lead
        ps["acLead_Agen"]=ps.apply(lambda r: div(r.Agendadas,r.Leads)*100,    axis=1)
        ps["acLead_Conc"]=ps.apply(lambda r: div(r.Concretadas,r.Leads)*100,  axis=1)
        ps["acLead_Apt"] =ps.apply(lambda r: div(r.Apartados,r.Leads)*100,    axis=1)
        ps["CAC"]        =ps.apply(lambda r: div(r.Invertido,r.Apartados),     axis=1)
        ps=ps.sort_values("Apartados",ascending=False)

        for i in range(0,len(ps),3):
            cols=st.columns(3)
            for j,(_,r) in enumerate(ps.iloc[i:i+3].iterrows()):
                c=PROJ_COLORS.get(r.proyecto,DEFAULT_COLOR)
                apar_i=int(r.Apartados)
                if apar_i==0:    badge_color,badge_txt="#FEE2E2","🔴 Sin apartados"
                elif r.CAC>200000: badge_color,badge_txt="#FEF9C3","🟡 CAC alto"
                else:            badge_color,badge_txt="#DCFCE7","🟢 En rango"
                with cols[j]:
                    st.markdown(f"""
                    <div style="background:white;border-radius:10px;padding:16px;
                                box-shadow:0 1px 4px rgba(0,0,0,.08);
                                border-left:4px solid {c};margin-bottom:14px">
                      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
                        <strong style="color:{c};font-size:.95rem">{r.proyecto}</strong>
                        <span style="background:{badge_color};font-size:.68rem;font-weight:600;
                                     padding:2px 7px;border-radius:20px">{badge_txt}</span>
                      </div>
                      <table style="width:100%;font-size:.81rem;border-collapse:collapse">
                        <tr><td style="color:#64748B;padding:3px 0">💵 Invertido</td>
                            <td style="text-align:right;font-weight:600">{money(r.Invertido)}</td></tr>
                        <tr><td colspan="2" style="padding:2px 0">
                          <div style="font-size:.67rem;font-weight:700;letter-spacing:.06em;
                                      text-transform:uppercase;color:#94A3B8;padding-top:6px">Embudo</div>
                        </td></tr>
                        <tr><td style="color:#64748B;padding:2px 0">📥 Leads</td>
                            <td style="text-align:right;font-weight:600">{int(r.Leads)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0">✅ Calificados</td>
                            <td style="text-align:right;font-weight:600">{int(r.Calificados)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0">📅 Agendadas</td>
                            <td style="text-align:right;font-weight:600">{int(r.Agendadas)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0">🏠 Concretadas</td>
                            <td style="text-align:right;font-weight:600">{int(r.Concretadas)}</td></tr>
                        <tr style="border-top:1px solid #F1F5F9">
                          <td style="color:{c};font-weight:700;padding:5px 0">🏆 Apartados</td>
                          <td style="text-align:right;font-weight:800;color:{c};font-size:1.05rem">{apar_i}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0">💳 CAC</td>
                            <td style="text-align:right;font-weight:600">{money(r.CAC)}</td></tr>
                        <tr><td colspan="2" style="padding:2px 0">
                          <div style="font-size:.67rem;font-weight:700;letter-spacing:.06em;
                                      text-transform:uppercase;color:#94A3B8;padding-top:6px">Conv. paso a paso</div>
                        </td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">Lead→Cal</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem">{pct(r.pLead_Cal)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">Cal→Agen</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem">{pct(r.pCal_Agen)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">Agen→Conc</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem">{pct(r.pAgen_Conc)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">Conc→Apt</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem">{pct(r.pConc_Apt)}</td></tr>
                        <tr><td colspan="2" style="padding:2px 0">
                          <div style="font-size:.67rem;font-weight:700;letter-spacing:.06em;
                                      text-transform:uppercase;color:#94A3B8;padding-top:6px">Conv. desde Lead</div>
                        </td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">→ Agendada</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem">{pct(r.acLead_Agen)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">→ Concretada</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem">{pct(r.acLead_Conc)}</td></tr>
                        <tr><td style="color:#64748B;padding:2px 0;font-size:.78rem">→ Apartado</td>
                            <td style="text-align:right;font-weight:600;font-size:.78rem;color:{c}">{pct(r.acLead_Apt)}</td></tr>
                      </table>
                    </div>""", unsafe_allow_html=True)

        st.divider()
        st.markdown('<p class="sec-title">Comparativa de Proyectos</p>', unsafe_allow_html=True)
        fig=go.Figure()
        for _,r in ps.iterrows():
            c=PROJ_COLORS.get(r.proyecto,DEFAULT_COLOR)
            fig.add_trace(go.Bar(name=r.proyecto,
                x=["Leads","Calificados","Agendadas","Concretadas","Apartados"],
                y=[r.Leads,r.Calificados,r.Agendadas,r.Concretadas,r.Apartados],
                marker_color=c))
        fig.update_layout(barmode="group",height=320,margin=dict(t=10,b=10,l=0,r=0),
                          legend=dict(orientation="h",y=1.08),
                          plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
                          xaxis=dict(gridcolor="#F1F5F9"),yaxis=dict(gridcolor="#F1F5F9"))
        st.plotly_chart(fig,use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: DATOS
# ══════════════════════════════════════════════════════════════════════════════
elif "Datos" in seccion:
    st.markdown("## 📋 Datos Crudos")

    tab1, tab2 = st.tabs(["📊 Sheet", "📡 Vambe"])

    with tab1:
        if df.empty:
            st.info("Sin datos de sheet con los filtros actuales.")
        else:
            st.caption(f"{len(df):,} filas · Año: {sel_year}")
            show_cols=["proyecto","mes_num","año_num","fuente","invertido",
                       "leadsMeta","leadsVambe","leadsOrganicos","leadsCalificados",
                       "visitaAgendada","visitaConcretada","apartados","cac"]
            disp=df[[c for c in show_cols if c in df.columns]].rename(columns={
                "mes_num":"Mes","año_num":"Año","proyecto":"Proyecto","fuente":"Fuente",
                "invertido":"Invertido","leadsMeta":"Leads Meta","leadsVambe":"Leads Vambe",
                "leadsOrganicos":"Leads Org","leadsCalificados":"Calificados",
                "visitaAgendada":"Agendadas","visitaConcretada":"Concretadas",
                "apartados":"Apartados","cac":"CAC"})
            if "Mes" in disp.columns:
                disp["Mes"]=pd.to_numeric(disp["Mes"],errors="coerce").map(MONTH_SHORT).fillna(disp["Mes"])
            st.dataframe(disp,use_container_width=True,height=480)
            st.download_button("⬇️ Descargar CSV",
                               df.to_csv(index=False).encode("utf-8"),
                               "marketing_sheet.csv","text/csv",use_container_width=True)

    with tab2:
        if vdf_raw.empty:
            st.info(f"Sin datos de Vambe. {vambe_err}")
        else:
            st.caption(f"{len(vdf_raw):,} contactos en el pipeline")
            disp_v=vdf_raw[["nombre","telefono","stage_label","fecha_creacion"]].rename(columns={
                "nombre":"Nombre","telefono":"Teléfono",
                "stage_label":"Etapa","fecha_creacion":"Fecha Entrada"})
            disp_v["Fecha Entrada"]=disp_v["Fecha Entrada"].dt.strftime("%d/%m/%Y")
            st.dataframe(disp_v,use_container_width=True,height=480)
            st.download_button("⬇️ Descargar CSV",
                               vdf_raw.to_csv(index=False).encode("utf-8"),
                               "vambe_contactos.csv","text/csv",use_container_width=True)
