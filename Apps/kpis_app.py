"""
KPIs Q2 2026 — PISO
Fuente: Apps/kpis_q2.xlsx  |  Persistencia: Supabase tabla kpis_scores_q2
"""

import streamlit as st
import pandas as pd
import hashlib
import requests
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KPIs Q2 2026 — PISO",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR   = Path(__file__).parent.resolve()
EXCEL_PATH = BASE_DIR / "kpis_q2.xlsx"

SOCIOS = [
    "Juan Ramón Moreno Flores",
    "Víctor Villa Walls",
    "Saúl Villa",
]

PROJECT_META = {
    "Kos":       "7 Escrituraciones",
    "P.C.":      "Construcción del 66% del proyecto + 100% de comercialización",
    "Santian":   "Inicio de comercialización, 2 ventas en el trimestre",
    "Zen":       "Autorización CUS e inicio construcción / 4 ventas + crédito puente",
    "Mar Negro": "Preparativo validado para inicio de proyecto",
    "Colón":     "Compra terreno ante Notario",
    "Wacuz":     "Adquisición terreno e inicio preparativo N26",
    "RRO":       "Inicio de Construcción Edificio Romero",
    "Admin":     "100% funciones Admin en ERP / Reportes implementados al 100%",
}

# ─────────────────────────────────────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.score-pill {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-weight: 600;
    font-size: 0.9rem;
}
.score-green  { background: #d1fae5; color: #065f46; }
.score-yellow { background: #fef9c3; color: #713f12; }
.score-red    { background: #fee2e2; color: #7f1d1d; }
.kpi-row { padding: 6px 0; border-bottom: 1px solid #f0f0f0; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SUPABASE
# ─────────────────────────────────────────────────────────────────────────────
def _sb_headers() -> dict:
    key = st.secrets["supabase"]["anon_key"]
    return {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates",
    }

def _sb_url() -> str:
    return st.secrets["supabase"]["url"] + "/rest/v1/kpis_scores_q2"


def sb_load() -> dict:
    """Devuelve dict {id: score} desde Supabase."""
    try:
        r = requests.get(_sb_url(), headers=_sb_headers(), timeout=5)
        if r.status_code == 200:
            return {row["id"]: row["score"] for row in r.json()}
    except Exception:
        pass
    return {}


def sb_save(records: list[dict]):
    """Upsert lista de {id, score} en Supabase."""
    try:
        requests.post(
            _sb_url(),
            headers=_sb_headers(),
            json=records,
            timeout=10,
        )
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, sheet_name="KPIs 2T", header=0)
    df.columns = df.columns.str.strip()

    text_cols = [
        "Proyecto", "Objetivo Empresa", "Objetivo Lider", "KPI",
        "Responsable 1", "Responsable 2", "Responsable 3", "Lider", "Fecha",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "NaN": ""})

    def classify(row):
        r2 = str(row.get("Responsable 2", "")).strip()
        r3 = str(row.get("Responsable 3", "")).strip()
        return "equipo" if (r2 not in ("", "nan") or r3 not in ("", "nan")) else "individual"

    df["tipo"] = df.apply(classify, axis=1)

    def make_uid(row):
        content = "|".join([
            row.get("Proyecto", ""),
            row.get("KPI", ""),
            row.get("Responsable 1", ""),
            row.get("Responsable 2", ""),
            row.get("Responsable 3", ""),
        ])
        return hashlib.md5(content.encode("utf-8")).hexdigest()[:12]

    df["uid"] = df.apply(make_uid, axis=1)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# SCORES EN SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
def init_scores(df: pd.DataFrame):
    """Carga scores desde Supabase y los pone en session_state."""
    remote = sb_load()

    scores = {
        "empresa":      {p: 0 for p in PROJECT_META},
        "kpis":         {row["uid"]: 0 for _, row in df.iterrows()},
        "discrecional": {},
    }

    for sid, val in remote.items():
        if sid.startswith("emp__"):
            proj = sid[5:]
            if proj in scores["empresa"]:
                scores["empresa"][proj] = int(val)
        elif sid.startswith("kpi__"):
            uid = sid[5:]
            if uid in scores["kpis"]:
                scores["kpis"][uid] = int(val)
        elif sid.startswith("disc__"):
            persona = sid[6:]
            scores["discrecional"][persona] = int(val)

    st.session_state.scores = scores


def save_scores(scores: dict):
    """Guarda todos los scores en Supabase."""
    records = []
    for proj, val in scores["empresa"].items():
        records.append({"id": f"emp__{proj}", "score": val})
    for uid, val in scores["kpis"].items():
        records.append({"id": f"kpi__{uid}", "score": val})
    for persona, val in scores.get("discrecional", {}).items():
        records.append({"id": f"disc__{persona}", "score": val})
    sb_save(records)


# ─────────────────────────────────────────────────────────────────────────────
# CÁLCULOS
# ─────────────────────────────────────────────────────────────────────────────
def calc_empresa(scores: dict) -> float:
    vals = [scores["empresa"].get(p, 0) for p in PROJECT_META]
    return sum(vals) / len(vals) if vals else 0.0


def calc_equipo(persona: str, df: pd.DataFrame, scores: dict) -> float:
    mask = (
        (
            (df["Responsable 1"] == persona)
            | (df["Responsable 2"] == persona)
            | (df["Responsable 3"] == persona)
        )
        & (df["tipo"] == "equipo")
    )
    kpis = df[mask]
    if kpis.empty:
        return 0.0
    vals = [scores["kpis"].get(row["uid"], 0) for _, row in kpis.iterrows()]
    return sum(vals) / len(vals)


def calc_individual(persona: str, df: pd.DataFrame, scores: dict) -> float:
    mask = (df["Responsable 1"] == persona) & (df["tipo"] == "individual")
    kpis = df[mask]
    if kpis.empty:
        return 0.0
    vals = [scores["kpis"].get(row["uid"], 0) for _, row in kpis.iterrows()]
    return sum(vals) / len(vals)


def calc_disc_pct(persona: str, scores: dict) -> float:
    return float(scores.get("discrecional", {}).get(persona, 0)) * 10.0


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS UI
# ─────────────────────────────────────────────────────────────────────────────
def pill(val: float) -> str:
    cls = "score-green" if val >= 80 else ("score-yellow" if val >= 60 else "score-red")
    return f'<span class="score-pill {cls}">{val:.1f}%</span>'


def semaforo(val: float) -> str:
    return "🟢" if val >= 80 else ("🟡" if val >= 60 else "🔴")


def all_people(df: pd.DataFrame) -> list[str]:
    return sorted([p for p in df["Responsable 1"].unique() if p])


def people_by_lider(df: pd.DataFrame, lider: str) -> list[str]:
    return sorted([p for p in df[df["Lider"] == lider]["Responsable 1"].unique() if p])


# ─────────────────────────────────────────────────────────────────────────────
# VISTA: EMPRESA
# ─────────────────────────────────────────────────────────────────────────────
def view_empresa(scores: dict):
    st.header("📊 Meta Empresa (30%)")
    st.markdown("Califica el cumplimiento de cada proyecto. El promedio = Score Empresa, igual para todos.")

    col_sliders, col_resumen = st.columns([3, 1])

    with col_sliders:
        for proyecto, meta in PROJECT_META.items():
            current = int(scores["empresa"].get(proyecto, 0))
            new_val = st.slider(
                f"**{proyecto}** — _{meta}_",
                min_value=0, max_value=100, value=current,
                key=f"emp__{proyecto}",
                format="%d%%",
            )
            scores["empresa"][proyecto] = new_val

    with col_resumen:
        avg = calc_empresa(scores)
        st.markdown("### Resumen")
        st.metric("Score Empresa", f"{avg:.1f}%")
        st.divider()
        for proyecto in PROJECT_META:
            v = scores["empresa"].get(proyecto, 0)
            st.markdown(f"{semaforo(v)} **{proyecto}**: {v}%")


# ─────────────────────────────────────────────────────────────────────────────
# VISTA: EQUIPO
# ─────────────────────────────────────────────────────────────────────────────
def view_equipo(df: pd.DataFrame, scores: dict):
    st.header("👥 Meta Equipo (30%)")
    st.markdown(
        "KPIs con Responsable 2 o 3 asignado. "
        "**Un solo % aplica igual para todos los responsables.**"
    )

    team_kpis = df[df["tipo"] == "equipo"].copy()

    proyectos = ["Todos"] + sorted(team_kpis["Proyecto"].unique().tolist())
    filtro = st.selectbox("Filtrar por proyecto", proyectos, key="eq_filtro")
    if filtro != "Todos":
        team_kpis = team_kpis[team_kpis["Proyecto"] == filtro]

    st.divider()

    for proyecto in sorted(team_kpis["Proyecto"].unique()):
        with st.expander(f"📁 {proyecto}", expanded=True):
            for _, row in team_kpis[team_kpis["Proyecto"] == proyecto].iterrows():
                uid = row["uid"]
                current = int(scores["kpis"].get(uid, 0))
                responsables = " · ".join([r for r in [row["Responsable 1"], row["Responsable 2"], row["Responsable 3"]] if r])

                c1, c2 = st.columns([4, 1])
                with c1:
                    st.markdown(
                        f"<div class='kpi-row'><b>{row['KPI']}</b><br>"
                        f"<small style='color:#6b7280'>Obj: {row['Objetivo Lider']} &nbsp;|&nbsp; "
                        f"👥 {responsables} &nbsp;|&nbsp; 📅 {row['Fecha']}</small></div>",
                        unsafe_allow_html=True,
                    )
                with c2:
                    new_val = st.number_input(
                        "%", min_value=0, max_value=100, value=current,
                        key=f"kpi__{uid}", label_visibility="collapsed",
                    )
                    scores["kpis"][uid] = new_val


# ─────────────────────────────────────────────────────────────────────────────
# VISTA: INDIVIDUAL
# ─────────────────────────────────────────────────────────────────────────────
def view_individual(df: pd.DataFrame, scores: dict):
    st.header("👤 Meta Individual (30%)")

    persona = st.selectbox("Seleccionar persona", all_people(df), key="ind_persona")
    if not persona:
        return

    ind_kpis  = df[(df["Responsable 1"] == persona) & (df["tipo"] == "individual")]
    team_kpis = df[
        ((df["Responsable 1"] == persona) | (df["Responsable 2"] == persona) | (df["Responsable 3"] == persona))
        & (df["tipo"] == "equipo")
    ]

    emp = calc_empresa(scores)
    equ = calc_equipo(persona, df, scores)
    ind = calc_individual(persona, df, scores)
    disc = calc_disc_pct(persona, scores)
    total = emp * 0.30 + equ * 0.30 + ind * 0.30 + disc * 0.10

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Empresa (30%)",      f"{emp:.1f}%")
    m2.metric("Equipo (30%)",       f"{equ:.1f}%")
    m3.metric("Individual (30%)",   f"{ind:.1f}%")
    m4.metric("Discrecional (10%)", f"{disc:.1f}%")
    m5.metric("TOTAL",              f"{total:.1f}%")

    st.divider()
    col_ind, col_eq = st.columns(2)

    with col_ind:
        st.markdown(f"### KPIs Individuales ({len(ind_kpis)})")
        if ind_kpis.empty:
            st.info("Sin KPIs individuales asignados.")
        else:
            for _, row in ind_kpis.iterrows():
                uid = row["uid"]
                current = int(scores["kpis"].get(uid, 0))
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.markdown(
                        f"<div class='kpi-row'><b>{row['KPI']}</b><br>"
                        f"<small style='color:#6b7280'>{row['Proyecto']} — {row['Objetivo Lider']} | 📅 {row['Fecha']}</small></div>",
                        unsafe_allow_html=True,
                    )
                with c2:
                    new_val = st.number_input(
                        "%", min_value=0, max_value=100, value=current,
                        key=f"ind__{uid}", label_visibility="collapsed",
                    )
                    scores["kpis"][uid] = new_val
        st.metric("Score Individual", f"{ind:.1f}%")

    with col_eq:
        st.markdown(f"### KPIs de Equipo ({len(team_kpis)})")
        if team_kpis.empty:
            st.info("Sin KPIs de equipo.")
        else:
            for _, row in team_kpis.iterrows():
                uid = row["uid"]
                kpi_score = scores["kpis"].get(uid, 0)
                responsables = " · ".join([r for r in [row["Responsable 1"], row["Responsable 2"], row["Responsable 3"]] if r])
                st.markdown(
                    f"<div class='kpi-row'><b>{row['KPI']}</b> &nbsp; {pill(kpi_score)}<br>"
                    f"<small style='color:#6b7280'>{row['Proyecto']} | 👥 {responsables} | 📅 {row['Fecha']}</small></div>",
                    unsafe_allow_html=True,
                )
        st.metric("Score Equipo", f"{equ:.1f}%")


# ─────────────────────────────────────────────────────────────────────────────
# VISTA: VARIABLE SALARIAL
# ─────────────────────────────────────────────────────────────────────────────
def view_variable_salarial(df: pd.DataFrame, scores: dict):
    st.header("💰 Variable Salarial")
    st.markdown("**Score Total = (Empresa × 30%) + (Equipo × 30%) + (Individual × 30%) + (Discrecional × 10%)**")

    with st.expander("⚙️ Calificación Discrecional (escala 0 – 10)", expanded=False):
        st.markdown("Cada líder asigna un valor de **0 a 10** a cada persona de su equipo.")
        cols = st.columns(3)
        for i, socio in enumerate(SOCIOS):
            with cols[i]:
                st.markdown(f"**{socio.split()[0]}**")
                equipo = [p for p in people_by_lider(df, socio) if p not in SOCIOS]
                for persona in equipo:
                    current = int(scores.get("discrecional", {}).get(persona, 0))
                    new_val = st.number_input(
                        persona, min_value=0, max_value=10, value=current,
                        key=f"disc__{persona}",
                    )
                    scores.setdefault("discrecional", {})[persona] = new_val

    st.divider()

    personas  = all_people(df)
    emp_score = calc_empresa(scores)

    rows = []
    for persona in personas:
        equ   = calc_equipo(persona, df, scores)
        ind   = calc_individual(persona, df, scores)
        disc  = calc_disc_pct(persona, scores)
        total = emp_score * 0.30 + equ * 0.30 + ind * 0.30 + disc * 0.10
        lider_vals = df[df["Responsable 1"] == persona]["Lider"]
        lider = lider_vals.iloc[0] if not lider_vals.empty else "—"
        rows.append({
            "Persona":            persona,
            "Lider":              lider,
            "Empresa (30%)":      round(emp_score, 1),
            "Equipo (30%)":       round(equ, 1),
            "Individual (30%)":   round(ind, 1),
            "Discrecional (10%)": round(disc, 1),
            "TOTAL":              round(total, 1),
        })

    result_df = pd.DataFrame(rows).sort_values("TOTAL", ascending=False)

    def highlight_total(val):
        if val >= 80:   return "background-color:#d1fae5;color:#065f46;font-weight:bold"
        if val >= 60:   return "background-color:#fef9c3;color:#713f12;font-weight:bold"
        return "background-color:#fee2e2;color:#7f1d1d;font-weight:bold"

    styled = result_df.style.applymap(highlight_total, subset=["TOTAL"]).format({
        "Empresa (30%)": "{:.1f}%", "Equipo (30%)": "{:.1f}%",
        "Individual (30%)": "{:.1f}%", "Discrecional (10%)": "{:.1f}%", "TOTAL": "{:.1f}%",
    })
    st.dataframe(styled, use_container_width=True, hide_index=True)
    st.markdown(f"**Score Empresa compartido:** {emp_score:.1f}%")


# ─────────────────────────────────────────────────────────────────────────────
# VISTA: ADMIN
# ─────────────────────────────────────────────────────────────────────────────
def view_admin(df: pd.DataFrame, scores: dict):
    st.header("⚙️ Admin — Vista de datos")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total KPIs", len(df))
    c2.metric("KPIs Equipo", len(df[df["tipo"] == "equipo"]))
    c3.metric("KPIs Individual", len(df[df["tipo"] == "individual"]))

    st.divider()
    tipo_f    = st.selectbox("Tipo",    ["Todos", "equipo", "individual"], key="adm_tipo")
    proy_f    = st.selectbox("Proyecto", ["Todos"] + sorted(df["Proyecto"].unique().tolist()), key="adm_proy")
    persona_f = st.selectbox("Persona (R1/R2/R3)", ["Todos"] + all_people(df), key="adm_persona")

    display = df.copy()
    if tipo_f    != "Todos": display = display[display["tipo"] == tipo_f]
    if proy_f    != "Todos": display = display[display["Proyecto"] == proy_f]
    if persona_f != "Todos":
        display = display[
            (display["Responsable 1"] == persona_f) |
            (display["Responsable 2"] == persona_f) |
            (display["Responsable 3"] == persona_f)
        ]

    display = display.copy()
    display["Score"] = display["uid"].map(lambda u: f"{scores['kpis'].get(u, 0)}%")

    show_cols = ["Proyecto","KPI","Objetivo Lider","Responsable 1","Responsable 2","Responsable 3","Lider","tipo","Fecha","Score"]
    st.dataframe(display[[c for c in show_cols if c in display.columns]], use_container_width=True, hide_index=True)

    st.divider()
    if st.button("🗑 Resetear TODOS los scores", type="secondary"):
        for uid in scores["kpis"]:    scores["kpis"][uid] = 0
        for p in scores["empresa"]:   scores["empresa"][p] = 0
        scores["discrecional"] = {}
        save_scores(scores)
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    df = load_data()

    if "scores" not in st.session_state:
        init_scores(df)

    scores = st.session_state.scores

    with st.sidebar:
        st.title("🏢 PISO KPIs Q2 2026")
        st.divider()

        vista = st.radio(
            "Navegación",
            ["📊 Empresa", "👥 Equipo", "👤 Individual", "💰 Variable Salarial", "⚙️ Admin"],
            label_visibility="collapsed",
        )

        st.divider()
        avg_emp = calc_empresa(scores)
        st.metric("Score Empresa", f"{avg_emp:.1f}%")
        st.divider()

        if st.button("💾 Guardar cambios", type="primary", use_container_width=True):
            save_scores(scores)
            st.success("✅ Guardado en la nube")

        st.caption("Cambios guardados en Supabase — accesibles desde cualquier dispositivo.")

    if vista == "📊 Empresa":
        view_empresa(scores)
    elif vista == "👥 Equipo":
        view_equipo(df, scores)
    elif vista == "👤 Individual":
        view_individual(df, scores)
    elif vista == "💰 Variable Salarial":
        view_variable_salarial(df, scores)
    elif vista == "⚙️ Admin":
        view_admin(df, scores)


if __name__ == "__main__":
    main()
