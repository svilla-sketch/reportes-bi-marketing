"""
vambe_client.py
Toda la lógica de conexión y procesamiento de datos de Vambe.
Importar en marketing_app.py con: from vambe_client import ...
"""
import requests
import pandas as pd

# ── Credenciales: lee de st.secrets si está disponible, sino usa defaults ──────
try:
    import streamlit as st
    _s = st.secrets
    VAMBE_BASE     = _s["vambe"]["base"]
    VAMBE_KEY      = _s["vambe"]["api_key"]
    VAMBE_PIPELINE = _s["vambe"]["pipeline"]
    SUPABASE_URL      = _s["supabase"]["url"]
    SUPABASE_ANON_KEY = _s["supabase"]["anon_key"]
except Exception:
    VAMBE_BASE     = "https://api.vambe.me"
    VAMBE_KEY      = "23dcedcbd2e421ec3da9a92d34c43db277e8ddb41fa95a2137cbbc6f1ac4e645"
    VAMBE_PIPELINE = "b6eb9000-a5c3-49aa-a413-9261d1224769"
    SUPABASE_URL      = "https://ybhvkhmuwljxpkrivyya.supabase.co"
    SUPABASE_ANON_KEY = (
        "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
        ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InliaHZraG11d2xqeHBrcml2eXlhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjgzMjQ5OTgsImV4cCI6MjA4MzkwMDk5OH0"
        ".lpSbKqlJCaVy5iypkA0Eo-GV-QYEBsJw0INkac9898k"
    )

# ── Mapeo de IDs de etapa → clave canónica ────────────────────────────────────
STAGE_MAP = {
    "f7fe6e71-834b-4ba3-b02a-fa720e8be8dd": "nuevo",
    "11688c74-accf-44e1-b069-8ecd60c3df9e": "en_com",
    "061f0cee-1835-40ca-9039-6f80d30a31c1": "recontacto",
    "479f7001-5bda-4ce6-9a97-4c9857456f09": "no_cal",
    "2470e46e-f268-4f8d-967d-59bf43fdd37d": "perfilado",
    "ddb7370c-1f04-415e-b415-19747f1c4cd3": "perfilado_com",
    "987f4fad-f5b3-4fb4-82cd-39b5a7e027bd": "cal_perdido",
    "b1a2ec53-1a6b-4f76-baab-3f2e6e5aee44": "vis_agendada",
    "10480786-2d0f-4677-916d-8bb779f96c61": "reagendamiento",
    "045c372b-9d83-44aa-aae0-217b552c3b71": "vis_concretada",
    "9d9ebf0d-9879-44ab-ba9c-4552721aa0a5": "seg_post",
    "91fa701c-fad9-479b-bb81-624dd331088b": "no_visito",
    "d112dfbf-3185-4662-86da-5b50cbd49b23": "evaluando",
    "eb3faccf-ab51-45ac-a18b-4ce7d1b4fd82": "apartado",
    "36ae7653-c531-4475-9391-3bb45e4acd88": "contrato",
    "c80bbf3f-3720-4bd7-982c-97c6a2ac8eaa": "descartado",
    "675926eb-f698-41ef-bd54-c1e9ef58ca48": "apt_cancel",
    "68ca2665-a6cd-445f-8a22-aad5f4362c27": "seg_largo",
}

# ── Etiquetas legibles ────────────────────────────────────────────────────────
STAGE_LABEL = {
    "nuevo":         "Nuevo",
    "en_com":        "En Comunicación",
    "recontacto":    "Recontacto",
    "no_cal":        "No Calificado",
    "perfilado":     "Perfilado",
    "perfilado_com": "Perfilado (Com.)",
    "cal_perdido":   "Calificado Perdido",
    "vis_agendada":  "Visita Agendada",
    "reagendamiento":"Re-Agendamiento",
    "vis_concretada":"Visita Concretada",
    "seg_post":      "Seg. Post-Visita",
    "no_visito":     "No Visitó",
    "evaluando":     "Evaluando",
    "apartado":      "Apartado",
    "contrato":      "Contrato",
    "descartado":    "Negocio Descartado",
    "apt_cancel":    "Apartado Cancelado",
    "seg_largo":     "Seg. Largo Plazo",
}

# ── Orden de visualización en gráficas ────────────────────────────────────────
STAGE_ORDER = [
    "nuevo", "en_com", "recontacto", "no_cal",
    "perfilado", "perfilado_com", "cal_perdido",
    "vis_agendada", "reagendamiento", "no_visito",
    "vis_concretada", "seg_post", "evaluando",
    "apartado", "contrato", "descartado", "apt_cancel", "seg_largo",
]

# ── Colores por etapa ─────────────────────────────────────────────────────────
STAGE_COLOR = {
    "nuevo":         "#94A3B8",
    "en_com":        "#60A5FA",
    "recontacto":    "#A78BFA",
    "no_cal":        "#F87171",
    "perfilado":     "#34D399",
    "perfilado_com": "#10B981",
    "cal_perdido":   "#FCA5A5",
    "vis_agendada":  "#818CF8",
    "reagendamiento":"#C4B5FD",
    "vis_concretada":"#6366F1",
    "seg_post":      "#A5B4FC",
    "no_visito":     "#FB923C",
    "evaluando":     "#FBBF24",
    "apartado":      "#EC4899",
    "contrato":      "#BE185D",
    "descartado":    "#EF4444",
    "apt_cancel":    "#F97316",
    "seg_largo":     "#CBD5E1",
}

# ── Grupos acumulados del embudo ──────────────────────────────────────────────
FUNNEL_CUM = {
    "calificados": {
        "perfilado", "perfilado_com", "cal_perdido",
        "vis_agendada", "reagendamiento", "vis_concretada",
        "seg_post", "no_visito", "evaluando",
        "apartado", "contrato", "descartado", "apt_cancel", "seg_largo",
    },
    "vis_agendada": {
        "vis_agendada", "reagendamiento", "vis_concretada",
        "seg_post", "no_visito", "evaluando",
        "apartado", "contrato", "descartado", "apt_cancel", "seg_largo",
    },
    "vis_concretada": {
        "vis_concretada", "seg_post", "evaluando",
        "apartado", "contrato", "descartado", "seg_largo",
    },
    "apartados": {"apartado", "contrato"},
}


def _parse_contacts(contacts: list) -> pd.DataFrame:
    """Convierte lista de contactos JSON a DataFrame con columnas estándar."""
    rows = []
    for c in contacts:
        ticket = c.get("active_ticket_v2") or {}
        # current_stage_id está disponible siempre; current_stage solo con pipeline filter
        sid  = ticket.get("current_stage_id", "") or ""
        skey = STAGE_MAP.get(sid, "?")
        rows.append({
            "id":          c.get("id", ""),
            "nombre":      c.get("name", ""),
            "telefono":    c.get("phone", ""),
            "created_at":  c.get("created_at", ""),
            "stage_id":    sid,
            "stage_key":   skey,
            "stage_label": STAGE_LABEL.get(skey, skey),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["fecha_creacion"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce")
        df["mes_num"]  = df["fecha_creacion"].dt.month
        df["año_num"]  = df["fecha_creacion"].dt.year
    return df


def load_vambe(days: int = 365) -> tuple[pd.DataFrame, str]:
    """
    Descarga todos los contactos de Vambe (sin filtro de pipeline).
    Usado para contar leads reales por fecha de creación (created_at).
    stage_key queda como '?' si el contacto no tiene un ticket activo en
    nuestro pipeline (normal: leads archivados, sin ticket, etc.).
    Retorna (DataFrame, error_string).
    """
    try:
        headers = {"x-api-key": VAMBE_KEY}
        r = requests.get(
            f"{VAMBE_BASE}/api/public/contacts",
            headers=headers,
            params={"days": days},   # sin pipeline_id → todos los contactos
            timeout=30,
        )
        r.raise_for_status()
        return _parse_contacts(r.json()), ""
    except Exception as e:
        return pd.DataFrame(), str(e)


def load_vambe_pipeline() -> tuple[pd.DataFrame, str]:
    """
    Descarga los contactos con ticket activo en nuestro pipeline.
    Incluye TODOS los períodos (days=730) para tener el estado completo del embudo.
    Usado para: tablero de etapas, embudo de conversión, conversiones.
    Retorna (DataFrame, error_string).
    """
    try:
        headers = {"x-api-key": VAMBE_KEY}
        r = requests.get(
            f"{VAMBE_BASE}/api/public/contacts",
            headers=headers,
            params={"days": 365, "pipeline_id": VAMBE_PIPELINE},
            timeout=30,
        )
        r.raise_for_status()
        return _parse_contacts(r.json()), ""
    except Exception as e:
        return pd.DataFrame(), str(e)


# ── Normalización de proyectos desde Supabase ─────────────────────────────────
_PROJECT_NORMALIZE = {
    "Punto Calma":  "Punto Calma",
    "Zen":          "Zen",
    "KOS":          "KOS",
    "DODEKA":       "DODEKA",
    "San Sebastián":"SANTIÁN",
    "Santian":      "SANTIÁN",
    "SANTIÁN":      "SANTIÁN",
}
_CHANNEL_NORMALIZE = {
    "Paid Media": "Paid Media",
    "Orgánico":   "Orgánico",
    "Página web": "Página web",
}


def _load_supabase_project_map() -> dict[str, tuple[str, str]]:
    """
    Consulta Supabase vambe_contacts y retorna {vambe_contact_id: (proyecto, canal)}.
    Una sola llamada REST paginada → muy rápido (~1-2 s en lugar de ~130 s).
    """
    sb_headers = {
        "apikey":        SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
    }
    contact_map: dict[str, tuple[str, str]] = {}
    page_size = 1000
    offset    = 0
    while True:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/vambe_contacts",
            headers=sb_headers,
            params={
                "select": "vambe_contact_id,project,channel",
                "limit":  page_size,
                "offset": offset,
            },
            timeout=15,
        )
        r.raise_for_status()
        rows = r.json()
        if not rows:
            break
        for row in rows:
            cid  = row.get("vambe_contact_id") or ""
            proj = _PROJECT_NORMALIZE.get(row.get("project") or "", "Sin proyecto")
            chan = _CHANNEL_NORMALIZE.get(row.get("channel") or "", "Sin canal")
            if cid:
                contact_map[cid] = (proj, chan)
        if len(rows) < page_size:
            break
        offset += page_size
    return contact_map


def load_vambe_pipeline_with_tags() -> tuple[pd.DataFrame, str]:
    """
    Descarga pipeline completo y enriquece con proyecto/canal desde Supabase.
    Rápido: una consulta Supabase en lugar de ~300 llamadas individuales a la API.
    Agrega columnas: 'proyecto' y 'canal'.
    """
    try:
        # 1. Pipeline contacts desde Vambe
        headers = {"x-api-key": VAMBE_KEY}
        r = requests.get(
            f"{VAMBE_BASE}/api/public/contacts",
            headers=headers,
            params={"days": 365, "pipeline_id": VAMBE_PIPELINE},
            timeout=30,
        )
        r.raise_for_status()
        contacts = r.json()

        # 2. Mapa proyecto/canal desde Supabase (rápido)
        contact_map = _load_supabase_project_map()

        # 3. Construir DataFrame
        rows = []
        for c in contacts:
            ticket = c.get("active_ticket_v2") or {}
            sid    = ticket.get("current_stage_id", "") or ""
            skey   = STAGE_MAP.get(sid, "?")
            cid    = c.get("id", "")
            proyecto, canal = contact_map.get(cid, ("Sin proyecto", "Sin canal"))

            rows.append({
                "id":          cid,
                "nombre":      c.get("name", ""),
                "telefono":    c.get("phone", ""),
                "created_at":  c.get("created_at", ""),
                "stage_id":    sid,
                "stage_key":   skey,
                "stage_label": STAGE_LABEL.get(skey, skey),
                "proyecto":    proyecto,
                "canal":       canal,
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df["fecha_creacion"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce")
            df["mes_num"]  = df["fecha_creacion"].dt.month
            df["año_num"]  = df["fecha_creacion"].dt.year
        return df, ""

    except Exception as e:
        return pd.DataFrame(), str(e)


def load_vambe_all_with_projects(days: int = 365) -> tuple[pd.DataFrame, str]:
    """
    Todos los contactos Vambe (sin filtro de pipeline) enriquecidos con
    proyecto y canal desde Supabase. Útil para contar leads TOTALES por proyecto.
    """
    try:
        headers = {"x-api-key": VAMBE_KEY}
        r = requests.get(
            f"{VAMBE_BASE}/api/public/contacts",
            headers=headers,
            params={"days": days},
            timeout=30,
        )
        r.raise_for_status()
        contacts = r.json()

        contact_map = _load_supabase_project_map()

        rows = []
        for c in contacts:
            cid = c.get("id", "")
            proyecto, canal = contact_map.get(cid, ("Sin proyecto", "Sin canal"))
            rows.append({
                "id":         cid,
                "nombre":     c.get("name", ""),
                "created_at": c.get("created_at", ""),
                "proyecto":   proyecto,
                "canal":      canal,
            })
        df = pd.DataFrame(rows)
        if not df.empty:
            df["fecha_creacion"] = pd.to_datetime(df["created_at"], utc=True, errors="coerce")
            df["mes_num"] = df["fecha_creacion"].dt.month
            df["año_num"] = df["fecha_creacion"].dt.year
        return df, ""
    except Exception as e:
        return pd.DataFrame(), str(e)


def vambe_funnel(df: pd.DataFrame) -> dict:
    """
    Calcula el embudo acumulado a partir de un DataFrame de contactos Vambe.
    Retorna dict con claves: total, calificados, vis_agendada, vis_concretada, apartados.
    """
    if df.empty:
        return {"total": 0, "calificados": 0, "vis_agendada": 0,
                "vis_concretada": 0, "apartados": 0}
    return {
        "total":          len(df),
        "calificados":    int(df["stage_key"].isin(FUNNEL_CUM["calificados"]).sum()),
        "vis_agendada":   int(df["stage_key"].isin(FUNNEL_CUM["vis_agendada"]).sum()),
        "vis_concretada": int(df["stage_key"].isin(FUNNEL_CUM["vis_concretada"]).sum()),
        "apartados":      int(df["stage_key"].isin(FUNNEL_CUM["apartados"]).sum()),
    }


# ─────────────────────────────────────────────────────────────────────────────
# VISITAS DEL MES
# Fuente: Supabase vambe_contacts → metadata.visit_date
# ─────────────────────────────────────────────────────────────────────────────

_VISIT_STAGE_STATUS = {
    "No Visitó":          "no_visito",
    "Visita Concretada":  "concretada",
    "Seg. Post-Visita":   "activo",
    "Evaluando":          "activo",
    "Seg. Largo Plazo":   "activo",
    "Re-Agendamiento":    "activo",
    "Negocio Descartado": "perdido",
    "Calificado Perdido": "perdido",
    "Apartado":           "ganado",
    "Contrato Firmado":   "ganado",
    "Apt. Cancelado":     "perdido",
}

VISIT_STATUS_LABELS = {
    "pendiente":  ("Pendiente",  "#F59E0B", "⏰"),
    "vencida":    ("Vencida",    "#EF4444", "⚠️"),
    "no_visito":  ("No Visitó",  "#6366F1", "🚫"),
    "concretada": ("Concretada", "#10B981", "✅"),
    "activo":     ("Activo",     "#8B5CF6", "👤"),
    "perdido":    ("Perdido",    "#DC2626", "❌"),
    "ganado":     ("Ganado",     "#D97706", "🏆"),
}


def load_visits_month(year: int, month: int) -> tuple[pd.DataFrame, str]:
    """
    Carga visitas del mes desde Supabase usando metadata.visit_date.
    Retorna DataFrame: nombre, telefono, proyecto, stage_name, visit_date, status
    status ∈ {pendiente, vencida, no_visito, concretada, activo, perdido, ganado}
    """
    from datetime import datetime, timezone

    try:
        first_day  = f"{year:04d}-{month:02d}-01"
        last_month = month + 1 if month < 12 else 1
        last_year  = year if month < 12 else year + 1
        first_next = f"{last_year:04d}-{last_month:02d}-01"

        sb_headers = {
            "apikey":        SUPABASE_ANON_KEY,
            "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        }

        rows_all: list[dict] = []
        page_size = 1000
        offset = 0
        while True:
            # Usamos lista de tuplas para poder repetir la misma key (PostgREST range filter)
            params = [
                ("select",                "name,phone,project,stage_name,metadata"),
                ("metadata->>visit_date", f"gte.{first_day}"),
                ("metadata->>visit_date", f"lt.{first_next}"),
                ("limit",                 page_size),
                ("offset",                offset),
            ]
            r = requests.get(
                f"{SUPABASE_URL}/rest/v1/vambe_contacts",
                headers=sb_headers,
                params=params,
                timeout=20,
            )
            r.raise_for_status()
            page = r.json()
            rows_all.extend(page)
            if len(page) < page_size:
                break
            offset += page_size

        if not rows_all:
            return pd.DataFrame(), ""

        now_utc = datetime.now(timezone.utc)
        records = []
        for c in rows_all:
            meta  = c.get("metadata") or {}
            vd_raw = meta.get("visit_date") or meta.get("fecha_reunion") or ""
            try:
                vd = datetime.fromisoformat(vd_raw.replace("Z", "+00:00"))
            except Exception:
                vd = None

            stage = c.get("stage_name") or ""
            if stage == "Visita Agendada":
                status = "pendiente" if (vd and vd > now_utc) else "vencida"
            else:
                status = _VISIT_STAGE_STATUS.get(stage, "activo")

            records.append({
                "nombre":     c.get("name", ""),
                "telefono":   c.get("phone", ""),
                "proyecto":   c.get("project") or "Sin proyecto",
                "stage_name": stage,
                "visit_date": vd,
                "status":     status,
            })

        df = pd.DataFrame(records)
        if not df.empty:
            df = df.sort_values("visit_date").reset_index(drop=True)
        return df, ""

    except Exception as e:
        return pd.DataFrame(), str(e)
