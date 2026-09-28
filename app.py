
import os
import io
import base64
import re
import uuid
import secrets
import hashlib
import hmac
import requests
import html
import json
import time
import zipfile
from urllib.parse import quote, urlparse
from datetime import datetime
from pathlib import Path
from lxml import etree

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.colors import HexColor
from reportlab.lib import colors
from reportlab.pdfgen import canvas

try:
    from openpyxl import Workbook as XLWorkbook, load_workbook
    from openpyxl.worksheet.datavalidation import DataValidation
except Exception:
    XLWorkbook = None
    load_workbook = None
    DataValidation = None

try:
    from supabase import create_client
except Exception:
    create_client = None

st.set_page_config(
    page_title="DIC | Informes mensuales",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

ITESO_BLUE = "#003B70"
ITESO_BLUE_2 = "#0B5A8F"
ITESO_CYAN = "#00A3E0"
ITESO_LIGHT = "#F3F6F8"
ITESO_BORDER = "#D7DEE5"
TEXT_GRAY = "#4B5563"

UNITS = {
    "CUE": "Centro Universidad Empresa",
    "COINCIDE": "Centro Universitario de Incidencia Social",
    "CUDJ": "Centro Universitario por la Dignidad y la Justicia Francisco Suárez, SJ",
    "CUI": "Centro Universitario Ignaciano",
    "CEJUVEN": "Centro de Acompañamiento y Estudios Juveniles",
    "CPC": "Centro de Promoción Cultural",
    "CEFSI": "Centro de Educación Física y Salud Integral",
}

MONTHS = [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
]

REPORT_RUBRICS = {
    "Vida universitaria": "Acciones para la comunidad universitaria: estudiantes, egresados, académicos y personal. Incluye procesos formativos no curriculares.",
    "Vinculación externa": "Acciones en alianza con instituciones o grupos externos, acuerdos o convenios y representación institucional.",
    "Desarrollo institucional": "Participación en comisiones institucionales, procesos de planeación, reestructura y mejora.",
    "Capacitación y formación del personal": "Talleres, cursos y seminarios dirigidos al propio equipo del centro; no incluye la formación que el centro ofrece hacia fuera.",
    "Participación en medios de difusión": "Apariciones en medios generalistas y redes sociales. Los productos científicos se reportan en Investigación.",
    "Oferta académica y docencia": "Materias curriculares de la DIC y participación docente. Se reporta únicamente en los tres cortes académicos del año.",
    "Investigación": "Proyectos, hitos, productos concluidos, participación en foros y difusión de investigación. Se reporta en enero y agosto.",
}
CATEGORIES = list(REPORT_RUBRICS.keys())
ACTION_PURPOSES = ["Formativa", "Incidencia", "Producción de conocimiento", "Gestión"]
ACTION_TYPES = ["Actividad", "Proceso", "Servicio"]
INCLUSION_CRITERIA = [
    "Atiende la misión del centro",
    "Trabajo colaborativo con otra instancia",
    "Relevante — Planeación institucional",
    "Relevante — Encargo de autoridades",
    "Otro",
]
TARGET_POPULATIONS = [
    "Estudiantes", "Egresados", "Académicos", "Personal administrativo",
    "Personal operativo", "Comunidad externa",
]
MEDIA_TYPES = ["Entrevista", "Nota", "Artículo", "Podcast", "Video", "Redes sociales", "Otro"]
RESEARCH_PARTICIPATION = ["Responsable / titular", "Colaborador", "Otro"]
RESEARCH_PROGRESS = ["Inicio", "En proceso", "Concluido"]
ACADEMIC_SEMESTERS = ["Primavera", "Verano", "Otoño"]
ACADEMIC_REPORTING_MONTHS = {"Enero", "Mayo", "Agosto"}  # Configurable si la DIC define otros cortes.
RESEARCH_REPORTING_MONTHS = {"Enero", "Agosto"}

# ---------- STYLE ----------
st.markdown(f"""
<style>
:root {{
    --iteso-blue: {ITESO_BLUE};
    --iteso-blue-2: {ITESO_BLUE_2};
    --iteso-cyan: {ITESO_CYAN};
    --iteso-light: {ITESO_LIGHT};
    --iteso-border: {ITESO_BORDER};
}}

html, body, [data-testid="stAppViewContainer"] {{
    background: #FFFFFF;
    color: #1F2937;
}}

[data-testid="stHeader"] {{
    background: rgba(255,255,255,0);
}}

.block-container {{
    padding-top: 1.0rem;
    max-width: 1360px;
}}

h1, h2, h3 {{
    color: var(--iteso-blue) !important;
    font-weight: 700 !important;
}}

p, label, span {{
    letter-spacing: 0;
}}

[data-testid="stSidebar"] {{
    background: #F4F6F8;
    border-right: 1px solid var(--iteso-border);
}}

[data-testid="stSidebar"] * {{
    color: #16324F !important;
}}

[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] input {{
    background: #FFFFFF !important;
    border-color: var(--iteso-border) !important;
    color: #16324F !important;
}}

[data-testid="stSidebar"] [role="radiogroup"] label {{
    padding: 0.15rem 0;
}}

.stButton > button {{
    border-radius: 8px;
    border: 1px solid var(--iteso-blue);
    background: #FFFFFF;
    color: var(--iteso-blue);
    font-weight: 600;
}}

.stButton > button:hover {{
    background: #EEF5FA;
    color: var(--iteso-blue);
    border-color: var(--iteso-blue-2);
}}

.stButton > button[kind="primary"] {{
    background: var(--iteso-blue);
    color: #FFFFFF;
    border-color: var(--iteso-blue);
}}

.stButton > button[kind="primary"]:hover {{
    background: var(--iteso-blue-2);
    color: #FFFFFF;
}}

[data-baseweb="select"] > div,
.stTextInput input,
.stNumberInput input,
.stTextArea textarea {{
    background: #FFFFFF !important;
    border: 1px solid var(--iteso-border) !important;
    color: #1F2937 !important;
}}

[data-testid="stExpander"] {{
    border: 1px solid var(--iteso-border);
    border-radius: 10px;
    background: #FFFFFF;
}}

[data-testid="stFileUploader"] section {{
    background: #FAFBFC !important;
    border: 1px dashed #AAB7C4 !important;
}}

[data-testid="stAlert"] {{
    border-radius: 10px;
}}

[data-testid="stMetric"] {{
    background: #FFFFFF;
}}

.dic-card {{
    background: #FFFFFF;
    border: 1px solid var(--iteso-border);
    border-radius: 12px;
    padding: 18px;
    margin-bottom: 12px;
}}

.small-muted {{
    color:#6B7280;
    font-size:0.9rem;
}}

.iteso-header {{
    display: flex;
    align-items: center;
    gap: 28px;
    padding: 6px 0 14px 0;
}}

.iteso-logo-wrap {{
    width: 270px;
    min-width: 270px;
    display: flex;
    align-items: center;
    justify-content: flex-start;
}}

.iteso-title-wrap {{
    display: flex;
    flex-direction: column;
    justify-content: center;
    min-height: 74px;
}}

.iteso-title {{
    color: var(--iteso-blue);
    font-size: 2.25rem;
    line-height: 1.1;
    font-weight: 700;
    margin: 0;
}}

.iteso-subtitle {{
    color: #53697A;
    font-size: 1rem;
    margin-top: 8px;
}}

.iteso-divider {{
    height: 1px;
    background: var(--iteso-border);
    margin: 6px 0 28px 0;
}}

[data-testid="stDownloadButton"] > button {{
    background: #FFFFFF;
    color: var(--iteso-blue);
    border: 1px solid var(--iteso-blue);
}}

[data-testid="stDownloadButton"] > button:hover {{
    background: #EEF5FA;
    color: var(--iteso-blue);
}}

/* Guardar acción: verde, sin afectar los demás botones de la aplicación */
div[class*="st-key-save_action_bar_"] .stButton > button {{
    background: #17823B !important;
    border-color: #17823B !important;
    color: #FFFFFF !important;
}}
div[class*="st-key-save_action_bar_"] .stButton > button:hover {{
    background: #116A30 !important;
    border-color: #116A30 !important;
    color: #FFFFFF !important;
}}
div[class*="st-key-save_action_bar_"] .stButton > button:disabled {{
    background: #DDEBE1 !important;
    border-color: #B7D5C0 !important;
    color: #66836E !important;
}}

</style>
""", unsafe_allow_html=True)

# ---------- DATABASE ----------
def get_supabase():
    if create_client is None:
        return None
    try:
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_KEY", "")
        if url and key:
            return create_client(url, key)
    except Exception:
        pass
    return None

supabase = get_supabase()

def db_mode():
    return "Supabase" if supabase else "Demo local (sesión)"

if "demo_reports" not in st.session_state:
    st.session_state.demo_reports = []
if "demo_activities" not in st.session_state:
    st.session_state.demo_activities = []
if "demo_photos" not in st.session_state:
    st.session_state.demo_photos = []
if "admin_authenticated" not in st.session_state:
    st.session_state.admin_authenticated = False
if "validated_center_email" not in st.session_state:
    st.session_state.validated_center_email = ""
if "show_center_preview" not in st.session_state:
    st.session_state.show_center_preview = False
if "show_consolidated_preview" not in st.session_state:
    st.session_state.show_consolidated_preview = False
if "final_selected_activities" not in st.session_state:
    st.session_state.final_selected_activities = {}
if "ranking_conflict_message" not in st.session_state:
    st.session_state.ranking_conflict_message = ""
if "daily_capsule_seen_key" not in st.session_state:
    st.session_state.daily_capsule_seen_key = ""
if "num_activities" not in st.session_state:
    st.session_state.num_activities = 5
if "director_page" not in st.session_state:
    st.session_state.director_page = "Nuevo reporte"
if "capture_month" not in st.session_state:
    st.session_state.capture_month = MONTHS[datetime.now().month - 1]
if "capture_year" not in st.session_state:
    st.session_state.capture_year = 2026
if "resuming_report_id" not in st.session_state:
    st.session_state.resuming_report_id = None
if "capture_method" not in st.session_state:
    st.session_state.capture_method = "Captura en línea"
if "docx_import_success" not in st.session_state:
    st.session_state.docx_import_success = ""
if "docx_import_warnings" not in st.session_state:
    st.session_state.docx_import_warnings = []
if "docx_import_filename" not in st.session_state:
    st.session_state.docx_import_filename = ""
if "center_authenticated" not in st.session_state:
    st.session_state.center_authenticated = False
if "center_user_email" not in st.session_state:
    st.session_state.center_user_email = ""
if "center_user_name" not in st.session_state:
    st.session_state.center_user_name = ""
if "center_user_unit" not in st.session_state:
    st.session_state.center_user_unit = ""
if "center_user_role" not in st.session_state:
    st.session_state.center_user_role = ""
if "password_setup_access_token" not in st.session_state:
    st.session_state.password_setup_access_token = ""
if "password_setup_email" not in st.session_state:
    st.session_state.password_setup_email = ""
if "password_setup_type" not in st.session_state:
    st.session_state.password_setup_type = ""
if "generated_activation_codes" not in st.session_state:
    st.session_state.generated_activation_codes = []
if "last_activity_at" not in st.session_state:
    st.session_state.last_activity_at = None
if "session_timeout_message" not in st.session_state:
    st.session_state.session_timeout_message = ""
if "stats_preview_enabled" not in st.session_state:
    st.session_state.stats_preview_enabled = False
if "database_backup_bytes" not in st.session_state:
    st.session_state.database_backup_bytes = None
if "database_backup_filename" not in st.session_state:
    st.session_state.database_backup_filename = ""
if "num_strategic_advances" not in st.session_state:
    st.session_state.num_strategic_advances = 1
if "num_strategic_issues" not in st.session_state:
    st.session_state.num_strategic_issues = 1
if "num_strategic_learnings" not in st.session_state:
    st.session_state.num_strategic_learnings = 1
if "no_strategic_issues" not in st.session_state:
    st.session_state.no_strategic_issues = False
if "monthly_strategic_summary" not in st.session_state:
    st.session_state.monthly_strategic_summary = ""

SESSION_TIMEOUT_SECONDS = 30 * 60

def clear_center_capture_state(reset_period=False):
    """Clear all temporary director-capture data from the current Streamlit session."""
    prefixes = (
        "title_", "desc_", "cat_", "other_cat_", "rank_", "part_", "rubro_",
        "purpose_", "action_type_", "criteria_", "criteria_other_", "planning_", "activity_date_",
        "location_", "population_", "external_population_", "dic_collab_", "dic_units_",
        "relevance_note_", "media_type_", "media_platform_", "media_topic_", "media_link_",
        "academic_semester_", "academic_credits_", "academic_opened_", "academic_groups_",
        "academic_fixed_prof_", "academic_variable_prof_", "faculty_name_", "faculty_contract_",
        "faculty_unit_", "faculty_category_", "faculty_promotion_", "research_role_",
        "research_members_", "research_field_", "research_actors_", "research_start_",
        "research_end_", "research_progress_", "research_products_", "research_pct_",
        "research_on_plan_", "research_on_plan_why_", "photos_", "social_", "chart_",
        "chart_title_", "existing_photos_", "existing_chart_",
        "adv_topic_", "adv_progress_", "adv_evidence_", "adv_level_", "adv_link_",
        "issue_type_", "issue_topic_", "issue_desc_", "issue_priority_", "issue_need_",
        "issue_dependency_", "issue_date_", "issue_link_", "issue_status_",
        "learn_type_", "learn_topic_", "learn_observation_", "learn_implication_",
        "learn_next_", "learn_link_",
    )
    exact_keys = {
        "show_center_preview",
        "ranking_conflict_message",
        "resuming_report_id",
        "docx_import_success",
        "docx_import_warnings",
        "docx_import_filename",
        "director_docx_upload",
        "fillable_docx_parsed",
        "fillable_docx_errors",
        "fillable_docx_warnings",
        "fillable_docx_upload",
        "highlight_selection", "learning_planning_advances", "learning_risks",
        "learning_opportunity", "media_appearances", "media_total_participations",
        "media_reach", "num_strategic_advances", "num_strategic_issues",
        "num_strategic_learnings", "no_strategic_issues", "monthly_strategic_summary",
    }

    for key in list(st.session_state.keys()):
        if key in exact_keys or key.startswith(prefixes):
            try:
                del st.session_state[key]
            except Exception:
                pass

    st.session_state.show_center_preview = False
    st.session_state.ranking_conflict_message = ""
    st.session_state.resuming_report_id = None
    st.session_state.docx_import_success = ""
    st.session_state.docx_import_warnings = []
    st.session_state.docx_import_filename = ""
    st.session_state.num_activities = 5
    st.session_state.num_strategic_advances = 1
    st.session_state.num_strategic_issues = 1
    st.session_state.num_strategic_learnings = 1
    st.session_state.no_strategic_issues = False
    st.session_state.monthly_strategic_summary = ""
    st.session_state.capture_method = "Captura en línea"

    if reset_period:
        st.session_state.capture_month = MONTHS[datetime.now().month - 1]
        st.session_state.capture_year = datetime.now().year


def _legacy_reflection_text(rows, kind):
    """Create a readable legacy text summary so older exports/searches remain useful."""
    parts = []
    for row in rows or []:
        if kind == "advance":
            text = f"{row.get('objective_topic','')}: {row.get('progress','')}"
            if row.get('evidence'):
                text += f" | Evidencia: {row.get('evidence')}"
            if row.get('progress_level'):
                text += f" | Nivel: {row.get('progress_level')}"
        elif kind == "issue":
            text = f"{row.get('issue_type','')}: {row.get('topic','')} — {row.get('description','')}"
            if row.get('priority'):
                text += f" | Prioridad: {row.get('priority')}"
            if row.get('need_type'):
                text += f" | Necesita: {row.get('need_type')}"
        else:
            text = f"{row.get('learning_type','')}: {row.get('topic','')} — {row.get('observation','')}"
            if row.get('implication'):
                text += f" | Implicación: {row.get('implication')}"
            if row.get('next_step'):
                text += f" | Próximo paso: {row.get('next_step')}"
        if text.strip(' :—|'):
            parts.append(text)
    return "\n".join(parts)


def get_structured_reflection(report_id):
    """Return the three structured strategic blocks for a monthly report."""
    empty = {"advances": [], "issues": [], "learnings": []}
    if not report_id:
        return empty
    if not supabase:
        rep = next((r for r in st.session_state.demo_reports if r.get("id") == report_id), None) or {}
        return {
            "advances": rep.get("strategic_advances") or [],
            "issues": rep.get("strategic_issues") or [],
            "learnings": rep.get("strategic_learnings") or [],
        }
    try:
        advances = supabase.table("report_strategic_advances").select("*").eq("report_id", report_id).order("order_index").execute().data or []
        issues = supabase.table("report_strategic_issues").select("*").eq("report_id", report_id).order("order_index").execute().data or []
        learnings = supabase.table("report_strategic_learnings").select("*").eq("report_id", report_id).order("order_index").execute().data or []
        return {"advances": advances, "issues": issues, "learnings": learnings}
    except Exception:
        return empty


def replace_structured_reflection(report_id, advances, issues, learnings, actor_email=""):
    """Replace report-level strategic rows. Only called on explicit draft/final saves."""
    if not report_id:
        return
    if not supabase:
        rep = next((r for r in st.session_state.demo_reports if r.get("id") == report_id), None)
        if rep is not None:
            rep["strategic_advances"] = advances or []
            rep["strategic_issues"] = issues or []
            rep["strategic_learnings"] = learnings or []
        return

    now = datetime.utcnow().isoformat()
    for table in ["report_strategic_advances", "report_strategic_issues", "report_strategic_learnings"]:
        supabase.table(table).delete().eq("report_id", report_id).execute()

    for idx, row in enumerate(advances or [], start=1):
        payload = {
            "report_id": report_id,
            "order_index": idx,
            "objective_topic": row.get("objective_topic"),
            "progress": row.get("progress"),
            "evidence": row.get("evidence") or None,
            "progress_level": row.get("progress_level"),
            "linked_activity_order": row.get("linked_activity_order"),
            "linked_activity_title": row.get("linked_activity_title") or None,
            "created_by": actor_email or None,
            "updated_by": actor_email or None,
            "updated_at": now,
        }
        supabase.table("report_strategic_advances").insert(payload).execute()

    for idx, row in enumerate(issues or [], start=1):
        target_date = row.get("target_date")
        if hasattr(target_date, "isoformat"):
            target_date = target_date.isoformat()
        payload = {
            "report_id": report_id,
            "order_index": idx,
            "issue_type": row.get("issue_type"),
            "topic": row.get("topic"),
            "description": row.get("description"),
            "priority": row.get("priority"),
            "need_type": row.get("need_type") or None,
            "dependency": row.get("dependency") or None,
            "target_date": target_date or None,
            "issue_status": row.get("issue_status") or "Abierto",
            "linked_activity_order": row.get("linked_activity_order"),
            "linked_activity_title": row.get("linked_activity_title") or None,
            "created_by": actor_email or None,
            "updated_by": actor_email or None,
            "updated_at": now,
        }
        supabase.table("report_strategic_issues").insert(payload).execute()

    for idx, row in enumerate(learnings or [], start=1):
        payload = {
            "report_id": report_id,
            "order_index": idx,
            "learning_type": row.get("learning_type"),
            "topic": row.get("topic"),
            "observation": row.get("observation"),
            "implication": row.get("implication"),
            "next_step": row.get("next_step") or None,
            "linked_activity_order": row.get("linked_activity_order"),
            "linked_activity_title": row.get("linked_activity_title") or None,
            "created_by": actor_email or None,
            "updated_by": actor_email or None,
            "updated_at": now,
        }
        supabase.table("report_strategic_learnings").insert(payload).execute()


def save_report(unit, month, year, status, sender_email="", report_extras=None):
    """Create/update the monthly report and its report-level strategic reflection."""
    report_extras = report_extras or {}
    advances = report_extras.get("strategic_advances") or []
    issues = report_extras.get("strategic_issues") or []
    learnings = report_extras.get("strategic_learnings") or []
    payload_extras = {
        k: v for k, v in {
            "monthly_highlights": report_extras.get("monthly_highlights"),
            "learning_planning_advances": report_extras.get("learning_planning_advances"),
            "learning_risks": report_extras.get("learning_risks"),
            "learning_opportunity": report_extras.get("learning_opportunity"),
            "media_monthly_summary": report_extras.get("media_monthly_summary"),
            "monthly_strategic_summary": report_extras.get("monthly_strategic_summary"),
            "no_strategic_issues": report_extras.get("no_strategic_issues"),
            "form_version": "2026-09-28-structured-reflection",
        }.items() if v is not None
    }
    if supabase:
        existing = (
            supabase.table("reports")
            .select("*")
            .eq("unit_code", unit)
            .eq("month", month)
            .eq("year", year)
            .execute()
        )
        if existing.data:
            rid = existing.data[0]["id"]
            payload = {
                "status": status,
                "sender_email": sender_email,
                "submitted_at": datetime.utcnow().isoformat() if status == "ENVIADO" else None,
                "updated_at": datetime.utcnow().isoformat(),
                **payload_extras,
            }
            supabase.table("reports").update(payload).eq("id", rid).execute()
            if any(k in report_extras for k in ("strategic_advances", "strategic_issues", "strategic_learnings")):
                replace_structured_reflection(rid, advances, issues, learnings, sender_email)
            return rid
        payload = {
            "unit_code": unit,
            "month": month,
            "year": year,
            "status": status,
            "sender_email": sender_email,
            **payload_extras,
        }
        res = supabase.table("reports").insert(payload).execute()
        rid = res.data[0]["id"]
        if any(k in report_extras for k in ("strategic_advances", "strategic_issues", "strategic_learnings")):
            replace_structured_reflection(rid, advances, issues, learnings, sender_email)
        return rid

    found = next((r for r in st.session_state.demo_reports
                  if r["unit_code"] == unit and r["month"] == month and r["year"] == year), None)
    if found:
        found["status"] = status
        found["sender_email"] = sender_email
        found["updated_at"] = datetime.now().isoformat()
        found.update(payload_extras)
        if status == "ENVIADO":
            found["submitted_at"] = datetime.now().isoformat()
        if any(k in report_extras for k in ("strategic_advances", "strategic_issues", "strategic_learnings")):
            replace_structured_reflection(found["id"], advances, issues, learnings, sender_email)
        return found["id"]
    rid = str(uuid.uuid4())
    st.session_state.demo_reports.append({
        "id": rid, "unit_code": unit, "month": month, "year": year,
        "status": status, "sender_email": sender_email,
        "submitted_at": datetime.now().isoformat() if status == "ENVIADO" else None,
        "created_at": datetime.now().isoformat(),
        **payload_extras,
    })
    if any(k in report_extras for k in ("strategic_advances", "strategic_issues", "strategic_learnings")):
        replace_structured_reflection(rid, advances, issues, learnings, sender_email)
    return rid


def save_single_activity(report_id, order_index, activity, actor_role="DIRECTOR"):
    """Persist one action without replacing the other actions in the monthly report."""
    persistence_errors = []
    actor_role = str(actor_role or "").upper()

    if supabase:
        existing_rows = (
            supabase.table("activities")
            .select("id,ranking,chart_storage_path")
            .eq("report_id", report_id)
            .eq("order_index", order_index)
            .limit(1)
            .execute()
            .data or []
        )
        existing = existing_rows[0] if existing_rows else None
        preserved_ranking = existing.get("ranking") if existing and actor_role != "DIRECTOR" else None

        payload = {
            "report_id": report_id,
            "title": activity["title"],
            "description_original": activity["description"],
            "description_edited": activity["description"],
            "category": activity["category"],
            "ranking": activity.get("ranking") if actor_role == "DIRECTOR" else preserved_ranking,
            "activity_date": serialize_date_value(activity.get("activity_date")),
            "participants": activity.get("participants"),
            "action_purpose": activity.get("action_purpose"),
            "action_type": activity.get("action_type"),
            "inclusion_criteria": activity.get("inclusion_criteria") or [],
            "location": activity.get("location"),
            "target_population": activity.get("target_population") or [],
            "target_external_name": activity.get("target_external_name"),
            "dic_collaboration": activity.get("dic_collaboration"),
            "dic_collaboration_units": activity.get("dic_collaboration_units"),
            "relevance_note": activity.get("relevance_note"),
            "detail_data": activity.get("detail_data") or {},
            "social_url": normalize_social_url(activity.get("social_url")),
            "chart_title": (activity.get("chart_title") or "").strip() or None,
            "chart_storage_path": None,
            "chart_original_filename": None,
            "order_index": order_index,
        }

        # Replace only the media for this one action. The other actions remain untouched.
        if existing:
            activity_id = existing["id"]
            old_photo_rows = (
                supabase.table("activity_photos")
                .select("storage_path")
                .eq("activity_id", activity_id)
                .execute()
                .data or []
            )
            old_paths = [r.get("storage_path") for r in old_photo_rows if r.get("storage_path")]
            if existing.get("chart_storage_path"):
                old_paths.append(existing["chart_storage_path"])
            if old_paths:
                try:
                    supabase.storage.from_("dic-activity-photos").remove(old_paths)
                except Exception:
                    pass
            try:
                supabase.table("activity_photos").delete().eq("activity_id", activity_id).execute()
            except Exception:
                pass
            res = supabase.table("activities").update(payload).eq("id", activity_id).execute()
            if not res.data:
                persistence_errors.append(f"No se pudo actualizar la acción `{activity.get('title','Acción')}`.")
                return persistence_errors
        else:
            res = supabase.table("activities").insert(payload).execute()
            if not res.data:
                persistence_errors.append(f"No se pudo crear la acción `{activity.get('title','Acción')}`.")
                return persistence_errors
            activity_id = res.data[0]["id"]

        all_photos = (activity.get("existing_photos", []) or []) + (activity.get("photos", []) or [])
        expected_photo_count = 0
        for photo_num, photo in enumerate(all_photos, start=1):
            photo_bytes = upload_bytes(photo)
            if not photo_bytes:
                persistence_errors.append(
                    f"`{activity['title']}`: la fotografía {photo_num} no contiene datos y no pudo guardarse."
                )
                continue
            expected_photo_count += 1
            photo_name = upload_name(photo, f"fotografia_{photo_num}.jpg")
            photo_mime = upload_mime(photo)
            ext = Path(photo_name).suffix.lower()
            if ext not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
                if "png" in (photo_mime or "").lower():
                    ext = ".png"
                elif "webp" in (photo_mime or "").lower():
                    ext = ".webp"
                else:
                    ext = ".jpg"
            storage_path = f"{report_id}/{activity_id}/photos/{uuid.uuid4().hex}{ext}"
            try:
                supabase.storage.from_("dic-activity-photos").upload(
                    path=storage_path,
                    file=photo_bytes,
                    file_options={"content-type": photo_mime or "image/jpeg", "upsert": "false"},
                )
                meta_res = supabase.table("activity_photos").insert({
                    "activity_id": activity_id,
                    "storage_path": storage_path,
                    "original_filename": photo_name,
                }).execute()
                if not meta_res.data:
                    raise RuntimeError("la fotografía se subió, pero no se registró en activity_photos")
                verify_bytes = supabase.storage.from_("dic-activity-photos").download(storage_path)
                if not verify_bytes:
                    raise RuntimeError("la verificación de la fotografía devolvió un archivo vacío")
            except Exception as exc:
                persistence_errors.append(
                    f"`{activity['title']}`: no se pudo guardar/verificar la fotografía {photo_num}. Detalle: {exc}"
                )

        if expected_photo_count:
            try:
                saved_photo_rows = (
                    supabase.table("activity_photos")
                    .select("id,storage_path")
                    .eq("activity_id", activity_id)
                    .execute()
                    .data or []
                )
                if len(saved_photo_rows) < expected_photo_count:
                    persistence_errors.append(
                        f"`{activity['title']}`: se esperaban {expected_photo_count} fotografía(s), "
                        f"pero sólo quedaron registradas {len(saved_photo_rows)}."
                    )
            except Exception as exc:
                persistence_errors.append(
                    f"`{activity['title']}`: no fue posible verificar las fotografías guardadas. Detalle: {exc}"
                )

        effective_chart = activity.get("chart") or activity.get("existing_chart")
        if effective_chart:
            chart_bytes = upload_bytes(effective_chart)
            if not chart_bytes:
                persistence_errors.append(f"`{activity['title']}`: la gráfica no contiene datos y no pudo guardarse.")
            else:
                chart_original_filename = upload_name(effective_chart, "grafica.jpg")
                chart_mime = upload_mime(effective_chart)
                ext = Path(chart_original_filename).suffix.lower()
                if ext not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
                    ext = ".png" if "png" in (chart_mime or "").lower() else ".jpg"
                chart_storage_path = f"{report_id}/{activity_id}/chart/{uuid.uuid4().hex}{ext}"
                try:
                    supabase.storage.from_("dic-activity-photos").upload(
                        path=chart_storage_path,
                        file=chart_bytes,
                        file_options={"content-type": chart_mime or "image/jpeg", "upsert": "false"},
                    )
                    verify_chart = supabase.storage.from_("dic-activity-photos").download(chart_storage_path)
                    if not verify_chart:
                        raise RuntimeError("la verificación de la gráfica devolvió un archivo vacío")
                    supabase.table("activities").update({
                        "chart_storage_path": chart_storage_path,
                        "chart_original_filename": chart_original_filename,
                    }).eq("id", activity_id).execute()
                except Exception as exc:
                    persistence_errors.append(
                        f"`{activity['title']}`: no se pudo guardar/verificar la gráfica. Detalle: {exc}"
                    )
        return persistence_errors

    # Demo/session mode: update or create only the requested order position.
    existing = next(
        (a for a in st.session_state.demo_activities
         if a.get("report_id") == report_id and int(a.get("order_index") or 0) == int(order_index)),
        None,
    )
    if existing:
        activity_id = existing["id"]
        preserved_ranking = existing.get("ranking") if actor_role != "DIRECTOR" else activity.get("ranking")
        existing.update({
            "title": activity["title"],
            "description_original": activity["description"],
            "description_edited": activity["description"],
            "category": activity["category"],
            "ranking": preserved_ranking,
            "activity_date": activity.get("activity_date"),
            "participants": activity.get("participants"),
            "action_purpose": activity.get("action_purpose"),
            "action_type": activity.get("action_type"),
            "inclusion_criteria": activity.get("inclusion_criteria") or [],
            "location": activity.get("location"),
            "target_population": activity.get("target_population") or [],
            "target_external_name": activity.get("target_external_name"),
            "dic_collaboration": activity.get("dic_collaboration"),
            "dic_collaboration_units": activity.get("dic_collaboration_units"),
            "relevance_note": activity.get("relevance_note"),
            "detail_data": activity.get("detail_data") or {},
            "social_url": normalize_social_url(activity.get("social_url")),
            "chart_title": activity.get("chart_title"),
            "chart_bytes": upload_bytes(activity.get("chart") or activity.get("existing_chart")),
            "chart_original_filename": upload_name(activity.get("chart") or activity.get("existing_chart"), "grafica.jpg") if (activity.get("chart") or activity.get("existing_chart")) else None,
            "order_index": order_index,
        })
        st.session_state.demo_photos = [p for p in st.session_state.demo_photos if p.get("activity_id") != activity_id]
    else:
        activity_id = str(uuid.uuid4())
        st.session_state.demo_activities.append({
            "id": activity_id,
            "report_id": report_id,
            "title": activity["title"],
            "description_original": activity["description"],
            "description_edited": activity["description"],
            "category": activity["category"],
            "ranking": activity.get("ranking") if actor_role == "DIRECTOR" else None,
            "activity_date": activity.get("activity_date"),
            "participants": activity.get("participants"),
            "action_purpose": activity.get("action_purpose"),
            "action_type": activity.get("action_type"),
            "inclusion_criteria": activity.get("inclusion_criteria") or [],
            "location": activity.get("location"),
            "target_population": activity.get("target_population") or [],
            "target_external_name": activity.get("target_external_name"),
            "dic_collaboration": activity.get("dic_collaboration"),
            "dic_collaboration_units": activity.get("dic_collaboration_units"),
            "relevance_note": activity.get("relevance_note"),
            "detail_data": activity.get("detail_data") or {},
            "social_url": normalize_social_url(activity.get("social_url")),
            "chart_title": activity.get("chart_title"),
            "chart_bytes": upload_bytes(activity.get("chart") or activity.get("existing_chart")),
            "chart_original_filename": upload_name(activity.get("chart") or activity.get("existing_chart"), "grafica.jpg") if (activity.get("chart") or activity.get("existing_chart")) else None,
            "order_index": order_index,
        })
    for photo in (activity.get("existing_photos", []) or []) + (activity.get("photos", []) or []):
        data = upload_bytes(photo)
        if data:
            st.session_state.demo_photos.append({
                "id": str(uuid.uuid4()),
                "activity_id": activity_id,
                "original_filename": upload_name(photo, "fotografia.jpg"),
                "mime_type": upload_mime(photo),
                "bytes": data,
            })
    return persistence_errors


def replace_activities(report_id, activities, actor_role="DIRECTOR"):
    """
    Replace activities and persist all media.
    Returns a list of persistence errors. An empty list means the save was verified.
    """
    persistence_errors = []

    if supabase:
        old_acts = (
            supabase.table("activities")
            .select("id,title,ranking,order_index,chart_storage_path")
            .eq("report_id", report_id)
            .execute()
            .data or []
        )
        # A collaborator can edit the report, but cannot assign or change rankings.
        # Preserve any ranking previously assigned by a director, matching title first
        # and order as a conservative fallback.
        prior_rank_by_title = {
            (oa.get("title") or "").strip().casefold(): oa.get("ranking")
            for oa in old_acts if oa.get("ranking")
        }
        prior_rank_by_order = {
            int(oa.get("order_index") or 0): oa.get("ranking")
            for oa in old_acts if oa.get("ranking")
        }

        # Clean physical files from the previous version.
        old_paths = []
        for oa in old_acts:
            photo_rows = (
                supabase.table("activity_photos")
                .select("storage_path")
                .eq("activity_id", oa["id"])
                .execute()
                .data or []
            )
            old_paths.extend([p["storage_path"] for p in photo_rows if p.get("storage_path")])
            if oa.get("chart_storage_path"):
                old_paths.append(oa["chart_storage_path"])

        if old_paths:
            try:
                supabase.storage.from_("dic-activity-photos").remove(old_paths)
            except Exception:
                pass

        for oa in old_acts:
            try:
                supabase.table("activity_photos").delete().eq("activity_id", oa["id"]).execute()
            except Exception:
                pass
        supabase.table("activities").delete().eq("report_id", report_id).execute()

        for i, a in enumerate(activities, start=1):
            res = supabase.table("activities").insert({
                "report_id": report_id,
                "title": a["title"],
                "description_original": a["description"],
                "description_edited": a["description"],
                "category": a["category"],
                "ranking": (
                    a["ranking"] if str(actor_role).upper() == "DIRECTOR"
                    else prior_rank_by_title.get((a.get("title") or "").strip().casefold(), prior_rank_by_order.get(i))
                ),
                "activity_date": a.get("activity_date").isoformat() if a.get("activity_date") else None,
                "participants": a.get("participants"),
                "action_purpose": a.get("action_purpose"),
                "action_type": a.get("action_type"),
                "inclusion_criteria": a.get("inclusion_criteria") or [],
                "location": a.get("location"),
                "target_population": a.get("target_population") or [],
                "target_external_name": a.get("target_external_name"),
                "dic_collaboration": a.get("dic_collaboration"),
                "dic_collaboration_units": a.get("dic_collaboration_units"),
                "relevance_note": a.get("relevance_note"),
                "detail_data": a.get("detail_data") or {},
                "social_url": normalize_social_url(a.get("social_url")),
                "chart_title": (a.get("chart_title") or "").strip() or None,
                "chart_storage_path": None,
                "chart_original_filename": None,
                "order_index": i,
            }).execute()

            if not res.data:
                persistence_errors.append(
                    f"No se pudo crear la actividad `{a.get('title','Actividad')}` en la base de datos."
                )
                continue

            activity_id = res.data[0]["id"]

            # -------- Photos: uploaded manually + imported from Word/resumed draft --------
            all_photos = (a.get("existing_photos", []) or []) + (a.get("photos", []) or [])
            expected_photo_count = 0

            for photo_num, photo in enumerate(all_photos, start=1):
                photo_bytes = upload_bytes(photo)
                if not photo_bytes:
                    persistence_errors.append(
                        f"`{a['title']}`: la fotografía {photo_num} no contiene datos y no pudo guardarse."
                    )
                    continue

                expected_photo_count += 1
                photo_name = upload_name(photo, f"fotografia_{photo_num}.jpg")
                photo_mime = upload_mime(photo)
                ext = Path(photo_name).suffix.lower()

                # Keep a safe image extension when Word/clipboard names do not provide one.
                if ext not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
                    if "png" in (photo_mime or "").lower():
                        ext = ".png"
                    elif "webp" in (photo_mime or "").lower():
                        ext = ".webp"
                    else:
                        ext = ".jpg"

                unique_name = f"{uuid.uuid4().hex}{ext}"
                storage_path = f"{report_id}/{activity_id}/photos/{unique_name}"

                try:
                    supabase.storage.from_("dic-activity-photos").upload(
                        path=storage_path,
                        file=photo_bytes,
                        file_options={
                            "content-type": photo_mime or "image/jpeg",
                            "upsert": "false",
                        },
                    )

                    meta_res = supabase.table("activity_photos").insert({
                        "activity_id": activity_id,
                        "storage_path": storage_path,
                        "original_filename": photo_name,
                    }).execute()

                    if not meta_res.data:
                        raise RuntimeError("la fotografía se subió, pero no se registró en activity_photos")

                    # Immediate read-back verification. This catches storage/permission/path problems
                    # before the report is reported as successfully saved.
                    verify_bytes = supabase.storage.from_("dic-activity-photos").download(storage_path)
                    if not verify_bytes:
                        raise RuntimeError("la verificación de la fotografía devolvió un archivo vacío")

                except Exception as exc:
                    persistence_errors.append(
                        f"`{a['title']}`: no se pudo guardar/verificar la fotografía {photo_num}. "
                        f"Detalle: {exc}"
                    )

            # Verify photo metadata count for this activity.
            if expected_photo_count:
                try:
                    saved_photo_rows = (
                        supabase.table("activity_photos")
                        .select("id,storage_path")
                        .eq("activity_id", activity_id)
                        .execute()
                        .data or []
                    )
                    if len(saved_photo_rows) < expected_photo_count:
                        persistence_errors.append(
                            f"`{a['title']}`: se esperaban {expected_photo_count} fotografía(s), "
                            f"pero sólo quedaron registradas {len(saved_photo_rows)}."
                        )
                except Exception as exc:
                    persistence_errors.append(
                        f"`{a['title']}`: no fue posible verificar las fotografías guardadas. Detalle: {exc}"
                    )

            # -------- Chart --------
            effective_chart = a.get("chart") or a.get("existing_chart")
            if effective_chart:
                chart_bytes = upload_bytes(effective_chart)
                if not chart_bytes:
                    persistence_errors.append(
                        f"`{a['title']}`: la gráfica no contiene datos y no pudo guardarse."
                    )
                else:
                    chart_original_filename = upload_name(effective_chart, "grafica.jpg")
                    chart_mime = upload_mime(effective_chart)
                    ext = Path(chart_original_filename).suffix.lower()
                    if ext not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
                        ext = ".png" if "png" in (chart_mime or "").lower() else ".jpg"

                    chart_storage_path = (
                        f"{report_id}/{activity_id}/chart/{uuid.uuid4().hex}{ext}"
                    )

                    try:
                        supabase.storage.from_("dic-activity-photos").upload(
                            path=chart_storage_path,
                            file=chart_bytes,
                            file_options={
                                "content-type": chart_mime or "image/jpeg",
                                "upsert": "false",
                            },
                        )
                        verify_chart = supabase.storage.from_("dic-activity-photos").download(
                            chart_storage_path
                        )
                        if not verify_chart:
                            raise RuntimeError("la verificación de la gráfica devolvió un archivo vacío")

                        supabase.table("activities").update({
                            "chart_storage_path": chart_storage_path,
                            "chart_original_filename": chart_original_filename,
                        }).eq("id", activity_id).execute()

                    except Exception as exc:
                        persistence_errors.append(
                            f"`{a['title']}`: no se pudo guardar/verificar la gráfica. Detalle: {exc}"
                        )

        return persistence_errors

    # Demo/session mode
    old_ids = [
        a["id"] for a in st.session_state.demo_activities
        if a["report_id"] == report_id
    ]
    st.session_state.demo_photos = [
        p for p in st.session_state.demo_photos
        if p["activity_id"] not in old_ids
    ]
    st.session_state.demo_activities = [
        a for a in st.session_state.demo_activities
        if a["report_id"] != report_id
    ]

    for i, a in enumerate(activities, start=1):
        aid = str(uuid.uuid4())
        chart = a.get("chart") or a.get("existing_chart")
        st.session_state.demo_activities.append({
            "id": aid,
            "report_id": report_id,
            "title": a["title"],
            "description_original": a["description"],
            "description_edited": a["description"],
            "category": a["category"],
            "ranking": a["ranking"] if str(actor_role).upper() == "DIRECTOR" else None,
            "activity_date": a.get("activity_date"),
            "participants": a.get("participants"),
            "action_purpose": a.get("action_purpose"),
            "action_type": a.get("action_type"),
            "inclusion_criteria": a.get("inclusion_criteria") or [],
            "location": a.get("location"),
            "target_population": a.get("target_population") or [],
            "target_external_name": a.get("target_external_name"),
            "dic_collaboration": a.get("dic_collaboration"),
            "dic_collaboration_units": a.get("dic_collaboration_units"),
            "relevance_note": a.get("relevance_note"),
            "detail_data": a.get("detail_data") or {},
            "social_url": normalize_social_url(a.get("social_url")),
            "chart_title": a.get("chart_title"),
            "chart_storage_path": None,
            "chart_original_filename": upload_name(chart, "grafica.jpg") if chart else None,
            "chart_bytes": upload_bytes(chart) if chart else None,
            "chart_mime_type": upload_mime(chart) if chart else None,
            "order_index": i,
        })

        all_photos = (a.get("existing_photos", []) or []) + (a.get("photos", []) or [])
        for photo in all_photos:
            data = upload_bytes(photo)
            if not data:
                persistence_errors.append(
                    f"`{a['title']}`: una fotografía no contenía datos."
                )
                continue
            st.session_state.demo_photos.append({
                "id": str(uuid.uuid4()),
                "activity_id": aid,
                "original_filename": upload_name(photo, "fotografia.jpg"),
                "mime_type": upload_mime(photo),
                "bytes": data,
            })

    return persistence_errors


def get_reports(month=None, year=None):
    if supabase:
        q = supabase.table("reports").select("*")
        if month:
            q = q.eq("month", month)
        if year:
            q = q.eq("year", year)
        return q.order("unit_code").execute().data or []
    rows = st.session_state.demo_reports
    if month:
        rows = [r for r in rows if r["month"] == month]
    if year:
        rows = [r for r in rows if r["year"] == year]
    return sorted(rows, key=lambda r: r["unit_code"])

def get_activities(report_id=None):
    if supabase:
        q = supabase.table("activities").select("*")
        if report_id:
            q = q.eq("report_id", report_id)
        return q.order("order_index").execute().data or []
    rows = st.session_state.demo_activities
    if report_id:
        rows = [a for a in rows if a["report_id"] == report_id]
    return sorted(rows, key=lambda a: a["order_index"])

def update_edited_description(activity_id, text):
    if supabase:
        supabase.table("activities").update({
            "description_edited": text,
            "updated_at": datetime.utcnow().isoformat(),
        }).eq("id", activity_id).execute()
    else:
        for a in st.session_state.demo_activities:
            if a["id"] == activity_id:
                a["description_edited"] = text

def get_activity_photos(activity_id):
    if supabase:
        rows = (
            supabase.table("activity_photos")
            .select("*")
            .eq("activity_id", activity_id)
            .execute()
            .data or []
        )
        out = []
        for row in rows:
            try:
                data = supabase.storage.from_("dic-activity-photos").download(row["storage_path"])
                out.append({
                    **row,
                    "bytes": data,
                    "mime_type": "image/png" if str(row.get("storage_path","")).lower().endswith(".png") else "image/jpeg",
                })
            except Exception:
                pass
        return out
    return [p for p in st.session_state.demo_photos if p["activity_id"] == activity_id]


def mark_session_activity():
    st.session_state.last_activity_at = time.time()


def clear_session_activity():
    st.session_state.last_activity_at = None


def enforce_session_timeout():
    """Expire authenticated sessions after 30 minutes without a Streamlit interaction."""
    authenticated = bool(
        st.session_state.get("admin_authenticated")
        or st.session_state.get("center_authenticated")
    )
    if not authenticated:
        return False

    now = time.time()
    last = st.session_state.get("last_activity_at")
    if last is not None and now - float(last) >= SESSION_TIMEOUT_SECONDS:
        if st.session_state.get("center_authenticated"):
            auth_logout_center()
        st.session_state.admin_authenticated = False
        clear_session_activity()
        st.session_state.session_timeout_message = (
            "La sesión se cerró automáticamente después de 30 minutos sin actividad."
        )
        return True

    mark_session_activity()
    return False


def check_admin_password(password):
    try:
        configured = st.secrets.get("ADMIN_PASSWORD", "")
    except Exception:
        configured = ""
    if not configured:
        return False, "ADMIN_PASSWORD no está configurada en los Secrets de Streamlit."
    if password == configured:
        st.session_state.admin_authenticated = True
        mark_session_activity()
        return True, ""
    return False, "Contraseña incorrecta."


def admin_gate():
    if st.session_state.admin_authenticated:
        if st.sidebar.button("Cerrar sesión de administración", use_container_width=True):
            st.session_state.admin_authenticated = False
            clear_session_activity()
            st.rerun()
        return True

    st.markdown("## Acceso restringido")
    st.info("Esta sección está reservada para la administración de la DIC.")
    pwd = st.text_input("Contraseña de administración", type="password", key="admin_password_input")
    if st.button("Desbloquear administración", type="primary"):
        ok, msg = check_admin_password(pwd)
        if ok:
            st.rerun()
        else:
            st.error(msg)
    return False


# ---------- HELPERS ----------

def add_docx_page_x_of_y(section):
    """Footer aligned right: PAGE/NUMPAGES."""
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def field(paragraph, instruction):
        run = paragraph.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instruction
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.extend([begin, instr, separate, end])

    field(p, "PAGE")
    p.add_run("/")
    field(p, "NUMPAGES")


class NumberedCanvas(canvas.Canvas):
    """Adds X/N at the bottom-right of every generated PDF."""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(total)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_page_number(self, total):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(HexColor("#6B7280"))
        self.drawRightString(letter[0] - 38, 22, f"{self._pageNumber}/{total}")
        self.restoreState()


def render_center_preview(unit, month, year, activities, report_extras=None):
    report_extras = report_extras or {}
    st.markdown(f"### Vista previa · {unit}")
    st.caption(f"{UNITS[unit]} · {month} {year}")
    for i, act in enumerate(activities, start=1):
        with st.container(border=True):
            st.markdown(f"#### {i}. {act['title']}")
            meta = [act.get("category", ""), rank_label(act.get("ranking"))]
            st.caption(" · ".join([x for x in meta if x]))
            for label, value in activity_detail_lines(act):
                st.markdown(f"**{label}:** {value}")
            if act.get("description"):
                st.write(act.get("description", ""))
            if act.get("social_url"):
                render_social_preview(act["social_url"], key_suffix=f"preview_{i}")
            chart = current_uploaded_chart(act)
            if chart:
                st.markdown(f"**{chart.get('title') or 'Gráfica'}**")
                st.image(chart["bytes"], use_container_width=True)
            photos = current_uploaded_photos(act)
            if photos:
                cols = st.columns(min(3, len(photos)))
                for idx, ph in enumerate(photos):
                    with cols[idx % len(cols)]:
                        st.image(ph["bytes"], use_container_width=True)

    highlights = report_extras.get("monthly_highlights") or []
    if highlights:
        st.markdown("### Lo más relevante del mes")
        for h in highlights:
            st.markdown(f"**Top {h.get('ranking')} · {h.get('title','')}**")
            st.write(h.get("why") or "")

    if any(report_extras.get(k) for k in ["learning_planning_advances", "learning_risks", "learning_opportunity"]):
        st.markdown("### Aprendizajes y seguimiento estratégico")
        advances = report_extras.get("strategic_advances") or []
        issues = report_extras.get("strategic_issues") or []
        learnings = report_extras.get("strategic_learnings") or []
        if advances:
            st.markdown("#### Avances sustantivos")
            for row in advances:
                st.markdown(f"**{row.get('objective_topic','')} · {row.get('progress_level','')}**")
                st.write(row.get("progress") or "")
                if row.get("evidence"): st.caption(f"Evidencia: {row.get('evidence')}")
        if report_extras.get("no_strategic_issues"):
            st.markdown("#### Riesgos, bloqueos y decisiones")
            st.write("Sin situaciones críticas reportadas este mes.")
        elif issues:
            st.markdown("#### Riesgos, bloqueos y decisiones")
            for row in issues:
                st.markdown(f"**{row.get('issue_type','')} · {row.get('topic','')} · Prioridad {row.get('priority','')}**")
                st.write(row.get("description") or "")
        if learnings:
            st.markdown("#### Aprendizajes y oportunidades")
            for row in learnings:
                st.markdown(f"**{row.get('learning_type','')} · {row.get('topic','')}**")
                st.write(row.get("observation") or "")
                st.caption(f"Implicación: {row.get('implication') or ''}")
        if report_extras.get("monthly_strategic_summary"):
            st.markdown("#### Síntesis del mes")
            st.write(report_extras.get("monthly_strategic_summary"))


def render_consolidated_preview(month, year, reports, activities_by_report):
    st.markdown("### Vista previa del informe consolidado")
    st.caption(f"Dirección de Integración Comunitaria · {month} {year}")
    st.markdown("#### Hitos por centro")
    for rep in reports:
        acts = [
            a for a in sorted(activities_by_report.get(rep["id"], []), key=activity_sort_key)
            if selected_for_final(a["id"])
        ]
        if not acts:
            continue

        with st.expander(
            f"{rep['unit_code']} · {UNITS.get(rep['unit_code'], '')} · {len(acts)} actividad(es)",
            expanded=False,
        ):
            for act in acts:
                with st.container(border=True):
                    st.markdown(
                        f"<div style='font-size:1.22rem;font-weight:800;color:{ITESO_BLUE};'>"
                        f"{act['title']}</div>",
                        unsafe_allow_html=True
                    )
                    st.caption(f"{rank_label(act.get('ranking'))} · {act.get('category','')}")
                    st.write(act.get("description_edited") or act.get("description_original") or "")
                    if act.get("social_url"):
                        render_social_preview(act["social_url"], key_suffix=f"consolidated_{act['id']}")
                    chart = get_activity_chart(act)
                    if chart:
                        st.markdown(f"**{chart.get('title') or 'Gráfica'}**")
                        st.image(chart["bytes"], use_container_width=True)
                    photos = get_activity_photos(act["id"])
                    if photos:
                        cols = st.columns(min(3, len(photos)))
                        for idx, ph in enumerate(photos):
                            with cols[idx % len(cols)]:
                                st.image(ph["bytes"], use_container_width=True)

def word_count(text):
    return len(re.findall(r"\b[\wÁÉÍÓÚÜÑáéíóúüñ'-]+\b", text or ""))


def activity_sort_key(activity):
    """Top 1, Top 2, Top 3, then unranked activities in their original order."""
    rank = activity.get("ranking")
    try:
        rank_int = int(rank) if rank is not None else None
    except Exception:
        rank_int = None

    if rank_int in (1, 2, 3):
        return (0, rank_int, activity.get("order_index", 999))
    return (1, 99, activity.get("order_index", 999))


def selected_for_final(activity_id):
    return bool(st.session_state.final_selected_activities.get(activity_id, False))


def rank_label(v):
    if v in [1, "1", "Top 1"]:
        return "Top 1"
    if v in [2, "2", "Top 2"]:
        return "Top 2"
    if v in [3, "3", "Top 3"]:
        return "Top 3"
    return "Sin ranking"



def available_rubrics_for_month(month, include_current=None):
    """Return the rubrics that should be offered for the selected reporting month."""
    rubrics = [
        "Vida universitaria",
        "Vinculación externa",
        "Desarrollo institucional",
        "Capacitación y formación del personal",
        "Participación en medios de difusión",
    ]
    if month in ACADEMIC_REPORTING_MONTHS:
        rubrics.append("Oferta académica y docencia")
    if month in RESEARCH_REPORTING_MONTHS:
        rubrics.append("Investigación")
    if include_current and include_current not in rubrics and include_current in REPORT_RUBRICS:
        rubrics.append(include_current)
    return rubrics



def coerce_date_value(value, fallback=None):
    """Convierte fechas de Supabase/Streamlit a datetime.date para st.date_input."""
    if value is None or value == "":
        return fallback if fallback is not None else datetime.now().date()

    # datetime -> date
    if hasattr(value, "date") and hasattr(value, "hour"):
        try:
            return value.date()
        except Exception:
            pass

    # date
    if (
        not isinstance(value, str)
        and hasattr(value, "year")
        and hasattr(value, "month")
        and hasattr(value, "day")
    ):
        return value

    # ISO/string
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return fallback if fallback is not None else datetime.now().date()

        for candidate in (raw, raw[:10]):
            try:
                return datetime.fromisoformat(candidate.replace("Z", "+00:00")).date()
            except Exception:
                pass

        for fmt in ("%d/%m/%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(raw, fmt).date()
            except Exception:
                pass

    return fallback if fallback is not None else datetime.now().date()


def serialize_date_value(value):
    """Serializa una fecha de forma segura aunque llegue como texto ISO."""
    if value in (None, ""):
        return None
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass
    if isinstance(value, str):
        return value[:10] if len(value) >= 10 else value
    return str(value)


def normalize_date_widget_state(key, fallback=None):
    """Normaliza session_state antes de renderizar un st.date_input."""
    if key in st.session_state:
        st.session_state[key] = coerce_date_value(
            st.session_state.get(key),
            fallback=fallback,
        )


def activity_from_session(index, actor_role=None, month=None):
    """Build one normalized activity record from Streamlit state and return validation messages."""
    actor_role = (actor_role or st.session_state.get("center_user_role") or "").upper()
    month = month or st.session_state.get("capture_month") or MONTHS[datetime.now().month - 1]
    rubro = st.session_state.get(f"rubro_{index}") or st.session_state.get(f"cat_{index}") or CATEGORIES[0]
    title = (st.session_state.get(f"title_{index}") or "").strip()
    desc = (st.session_state.get(f"desc_{index}") or "").strip()
    participants = int(st.session_state.get(f"part_{index}", 0) or 0)
    photos = st.session_state.get(f"photos_{index}", []) or []
    existing_photos = st.session_state.get(f"existing_photos_{index}", []) or []
    chart = st.session_state.get(f"chart_{index}")
    existing_chart = st.session_state.get(f"existing_chart_{index}")
    chart_title = (st.session_state.get(f"chart_title_{index}") or "").strip()
    social_url = (st.session_state.get(f"social_{index}") or "").strip()
    relevance_note = (st.session_state.get(f"relevance_note_{index}") or "").strip()
    detail = {}
    missing = []

    # An activity is considered started only when the user enters substantive content.
    started = bool(title or desc or participants or photos or existing_photos or chart or existing_chart or social_url)
    if rubro == "Participación en medios de difusión":
        started = started or bool(st.session_state.get(f"media_platform_{index}") or st.session_state.get(f"media_topic_{index}"))
    elif rubro == "Oferta académica y docencia":
        started = started or bool(st.session_state.get(f"faculty_name_{index}"))
    elif rubro == "Investigación":
        started = started or bool(st.session_state.get(f"research_members_{index}") or st.session_state.get(f"research_field_{index}"))

    if not started:
        return None, []

    record = {
        "title": title,
        "description": desc,
        "category": rubro,
        "ranking": None,
        "participants": participants or None,
        "activity_date": st.session_state.get(f"activity_date_{index}"),
        "action_purpose": None,
        "action_type": None,
        "inclusion_criteria": [],
        "location": None,
        "target_population": [],
        "target_external_name": None,
        "dic_collaboration": None,
        "dic_collaboration_units": None,
        "relevance_note": relevance_note or None,
        "detail_data": detail,
        "photos": photos,
        "social_url": normalize_social_url(social_url),
        "chart": chart,
        "chart_title": chart_title,
        "existing_photos": existing_photos,
        "existing_chart": existing_chart,
    }

    if rubro in {
        "Vida universitaria", "Vinculación externa", "Desarrollo institucional",
        "Capacitación y formación del personal"
    }:
        purpose = st.session_state.get(f"purpose_{index}") or ""
        action_type = st.session_state.get(f"action_type_{index}") or ""
        criteria = st.session_state.get(f"criteria_{index}", []) or []
        planning_type = st.session_state.get(f"planning_{index}") or ""
        criteria_other = (st.session_state.get(f"criteria_other_{index}") or "").strip()
        location = (st.session_state.get(f"location_{index}") or "").strip()
        population = st.session_state.get(f"population_{index}", []) or []
        external_name = (st.session_state.get(f"external_population_{index}") or "").strip()
        collab = st.session_state.get(f"dic_collab_{index}", "No")
        dic_units = (st.session_state.get(f"dic_units_{index}") or "").strip()
        if criteria_other and "Otro" in criteria:
            criteria = [x for x in criteria if x != "Otro"] + [f"Otro: {criteria_other}"]
        record.update({
            "action_purpose": purpose or None,
            "action_type": action_type or None,
            "inclusion_criteria": criteria,
            "location": location or None,
            "target_population": population,
            "target_external_name": external_name or None,
            "dic_collaboration": collab == "Sí",
            "dic_collaboration_units": dic_units or None,
        })
        detail["planning_type"] = planning_type or None
        if not title: missing.append("título de la acción")
        if not purpose: missing.append("fin de la acción")
        if not action_type: missing.append("tipo de acción")
        if not criteria: missing.append("criterio(s) de inclusión")
        if not location: missing.append("lugar")
        if participants <= 0: missing.append("número de participantes")
        if not population: missing.append("población a la que está dirigida")
        if "Comunidad externa" in population and not external_name:
            missing.append("nombre de la comunidad/institución externa")
        if collab == "Sí" and not dic_units:
            missing.append("instancias DIC que colaboran")
        if not desc: missing.append("descripción breve")
        if not relevance_note: missing.append("por qué importa / qué conviene que la dirección sepa")

    elif rubro == "Participación en medios de difusión":
        media_type = st.session_state.get(f"media_type_{index}") or ""
        platform = (st.session_state.get(f"media_platform_{index}") or "").strip()
        topic = (st.session_state.get(f"media_topic_{index}") or title).strip()
        link = (st.session_state.get(f"media_link_{index}") or social_url).strip()
        record["title"] = topic
        record["description"] = f"{media_type} en {platform}. Tema: {topic}".strip()
        record["social_url"] = normalize_social_url(link)
        record["participants"] = None
        record["relevance_note"] = None
        detail.update({"participation_type": media_type, "platform": platform, "topic": topic, "link": normalize_social_url(link)})
        if not media_type: missing.append("tipo de participación")
        if not platform: missing.append("medio o plataforma")
        if not topic: missing.append("tema de la participación")
        if link and not social_url_is_valid(normalize_social_url(link)): missing.append("enlace válido")

    elif rubro == "Oferta académica y docencia":
        semester = st.session_state.get(f"academic_semester_{index}") or ""
        credits = int(st.session_state.get(f"academic_credits_{index}", 0) or 0)
        opened = st.session_state.get(f"academic_opened_{index}") or ""
        groups = int(st.session_state.get(f"academic_groups_{index}", 0) or 0)
        fixed_prof = int(st.session_state.get(f"academic_fixed_prof_{index}", 0) or 0)
        variable_prof = int(st.session_state.get(f"academic_variable_prof_{index}", 0) or 0)
        faculty_name = (st.session_state.get(f"faculty_name_{index}") or "").strip()
        faculty_contract = st.session_state.get(f"faculty_contract_{index}") or ""
        faculty_unit = (st.session_state.get(f"faculty_unit_{index}") or "").strip()
        faculty_category = st.session_state.get(f"faculty_category_{index}") or ""
        faculty_promotion = st.session_state.get(f"faculty_promotion_{index}") or ""
        detail.update({
            "semester": semester, "credits": credits, "opened": opened, "groups": groups,
            "fixed_professors": fixed_prof, "variable_professors": variable_prof,
            "faculty_name": faculty_name, "faculty_contract": faculty_contract,
            "faculty_unit": faculty_unit, "faculty_category": faculty_category,
            "faculty_promotion": faculty_promotion,
        })
        record["description"] = (
            f"{semester}. Créditos: {credits}. Se abrió: {opened}. Grupos: {groups}. "
            f"Profesores tiempo fijo: {fixed_prof}; tiempo variable: {variable_prof}."
        )
        record["participants"] = None
        if not title: missing.append("nombre del curso ofertado")
        if not semester: missing.append("semestre")
        if credits <= 0: missing.append("créditos")
        if not opened: missing.append("si el curso se abrió")
        if opened == "Sí" and groups <= 0: missing.append("número de grupos")
        if faculty_name:
            if not faculty_contract: missing.append("tipo de contrato del integrante")
            if not faculty_unit: missing.append("dependencia de la materia")
            if not faculty_category: missing.append("categoría académica")
            if not faculty_promotion: missing.append("si está en proceso de promoción")

    elif rubro == "Investigación":
        role = st.session_state.get(f"research_role_{index}") or ""
        members = (st.session_state.get(f"research_members_{index}") or "").strip()
        field = (st.session_state.get(f"research_field_{index}") or "").strip()
        actors = (st.session_state.get(f"research_actors_{index}") or "").strip()
        progress = st.session_state.get(f"research_progress_{index}") or ""
        products = (st.session_state.get(f"research_products_{index}") or "").strip()
        pct = int(st.session_state.get(f"research_pct_{index}", 0) or 0)
        on_plan = st.session_state.get(f"research_on_plan_{index}") or ""
        why = (st.session_state.get(f"research_on_plan_why_{index}") or "").strip()
        start_date = st.session_state.get(f"research_start_{index}")
        end_date = st.session_state.get(f"research_end_{index}")
        detail.update({
            "participation_level": role, "members": members, "field": field, "linked_actors": actors,
            "start_date": start_date.isoformat() if hasattr(start_date, "isoformat") else start_date,
            "end_date": end_date.isoformat() if hasattr(end_date, "isoformat") else end_date,
            "progress_level": progress, "products": products, "products_progress_pct": pct,
            "on_plan": on_plan, "on_plan_why": why,
        })
        record["description"] = products or f"Proyecto de investigación en {field}."
        record["participants"] = None
        if not title: missing.append("proyecto")
        if not role: missing.append("nivel de participación")
        if not members: missing.append("integrantes en el proyecto")
        if not field: missing.append("temática o campo de conocimiento")
        if not actors: missing.append("actores internos o externos con quienes se vincula")
        if not progress: missing.append("nivel de avance")
        if not products: missing.append("productos o actividades de difusión/investigación")
        if not on_plan: missing.append("si va según lo planeado")
        if on_plan == "No" and not why: missing.append("explicación de por qué no va según lo planeado")

    if record.get("description") and word_count(record["description"]) > 250:
        missing.append(f"descripción: {word_count(record['description'])}/250 palabras")
    if (chart or existing_chart) and not chart_title:
        missing.append("título de la gráfica")
    return record, missing


def apply_highlights_to_activities(activities, highlights):
    """Apply Top 1/2/3 rankings from the report-level relevance selection."""
    for act in activities:
        act["ranking"] = None
    for h in highlights or []:
        idx = h.get("activity_index")
        rank = h.get("ranking")
        if isinstance(idx, int) and 0 <= idx < len(activities) and rank in (1, 2, 3):
            activities[idx]["ranking"] = rank
    return activities


def activity_detail_lines(act):
    """Human-readable field/value lines for previews and exports."""
    lines = []
    if act.get("action_purpose"):
        lines.append(("Fin de la acción", act.get("action_purpose")))
    if act.get("action_type"):
        lines.append(("Tipo de acción", act.get("action_type")))
    if act.get("inclusion_criteria"):
        lines.append(("Criterios de inclusión", ", ".join(act.get("inclusion_criteria") or [])))
    if act.get("activity_date"):
        lines.append(("Fecha", str(act.get("activity_date"))))
    if act.get("location"):
        lines.append(("Lugar", act.get("location")))
    if act.get("participants"):
        lines.append(("Participantes", str(act.get("participants"))))
    if act.get("target_population"):
        lines.append(("Población", ", ".join(act.get("target_population") or [])))
    if act.get("target_external_name"):
        lines.append(("Comunidad / institución externa", act.get("target_external_name")))
    if act.get("dic_collaboration") is not None:
        lines.append(("Colaboran otras instancias DIC", "Sí" if act.get("dic_collaboration") else "No"))
    if act.get("dic_collaboration_units"):
        lines.append(("Instancias DIC", act.get("dic_collaboration_units")))
    if act.get("relevance_note"):
        lines.append(("Por qué importa", act.get("relevance_note")))
    detail = act.get("detail_data") or {}
    rubric = act.get("category")
    if rubric == "Participación en medios de difusión":
        for label, key in [("Tipo de participación","participation_type"),("Medio o plataforma","platform"),("Tema","topic")]:
            if detail.get(key): lines.append((label, str(detail.get(key))))
    elif rubric == "Oferta académica y docencia":
        mapping = [
            ("Semestre","semester"),("Créditos","credits"),("Se abrió","opened"),("Número de grupos","groups"),
            ("Profesores Tiempo Fijo","fixed_professors"),("Profesores Tiempo Variable","variable_professors"),
            ("Docente","faculty_name"),("Tipo de contrato","faculty_contract"),("Dependencia","faculty_unit"),
            ("Categoría académica","faculty_category"),("En promoción","faculty_promotion")
        ]
        for label, key in mapping:
            if detail.get(key) not in (None, "", 0): lines.append((label, str(detail.get(key))))
    elif rubric == "Investigación":
        mapping = [
            ("Nivel de participación","participation_level"),("Integrantes","members"),("Campo de conocimiento","field"),
            ("Actores vinculados","linked_actors"),("Fecha de inicio","start_date"),("Fecha de término","end_date"),
            ("Nivel de avance","progress_level"),("Productos / difusión","products"),("% avance de productos","products_progress_pct"),
            ("Va según lo planeado","on_plan"),("Explicación","on_plan_why")
        ]
        for label, key in mapping:
            if detail.get(key) not in (None, "", 0): lines.append((label, str(detail.get(key))))
    return lines

def validate_current_activities(num_activities):
    activities = []
    messages = []
    month = st.session_state.get("capture_month") or MONTHS[datetime.now().month - 1]
    role = st.session_state.get("center_user_role") or ""
    for i in range(num_activities):
        record, missing = activity_from_session(i, actor_role=role, month=month)
        if record is None:
            continue
        activities.append(record)
        if missing:
            messages.append(f"Acción {i+1}: " + " · ".join(missing))
    return activities, messages

def show_validation_messages(messages):
    if messages:
        st.error("Completa los siguientes campos antes de continuar:")
        for message in messages:
            st.markdown(f"- {message}")


def handle_rank_change(index):
    """
    Detect duplicate rankings without forcing a rerun.
    This preserves every field already typed in the report.
    """
    current = st.session_state.get(f"rank_{index}", "Sin ranking")
    if current == "Sin ranking":
        st.session_state.ranking_conflict_message = ""
        return

    for j in range(st.session_state.num_activities):
        if j == index:
            continue
        if st.session_state.get(f"rank_{j}") == current:
            st.session_state.ranking_conflict_message = (
                f"{current} ya está asignado a la Actividad {j+1}. "
                f"Cambia el ranking de la Actividad {index+1} antes de continuar."
            )
            return

    st.session_state.ranking_conflict_message = ""


def delete_report_completely(report_id):
    if not supabase:
        act_ids = [a["id"] for a in st.session_state.demo_activities if a["report_id"] == report_id]
        st.session_state.demo_photos = [p for p in st.session_state.demo_photos if p["activity_id"] not in act_ids]
        st.session_state.demo_activities = [a for a in st.session_state.demo_activities if a["report_id"] != report_id]
        st.session_state.demo_reports = [r for r in st.session_state.demo_reports if r["id"] != report_id]
        return

    acts = (
        supabase.table("activities")
        .select("id,chart_storage_path")
        .eq("report_id", report_id)
        .execute()
        .data or []
    )
    act_ids = [a["id"] for a in acts]

    if act_ids:
        photo_rows = (
            supabase.table("activity_photos")
            .select("storage_path")
            .in_("activity_id", act_ids)
            .execute()
            .data or []
        )
        paths = [p["storage_path"] for p in photo_rows if p.get("storage_path")]
        paths.extend([a["chart_storage_path"] for a in acts if a.get("chart_storage_path")])
        if paths:
            supabase.storage.from_("dic-activity-photos").remove(paths)

    supabase.table("reports").delete().eq("id", report_id).execute()


@st.dialog("Eliminar envío", dismissible=False)
def delete_report_dialog(report_id, unit_code, month, year):
    st.warning(
        f"Se eliminará **{unit_code} · {month} {year}**, incluyendo actividades y fotografías. "
        "Esta acción no se puede deshacer."
    )
    confirmation = st.text_input("Escribe BORRAR para confirmar")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Cancelar", use_container_width=True):
            st.rerun()
    with c2:
        if st.button(
            "Eliminar definitivamente",
            type="primary",
            use_container_width=True,
            disabled=(confirmation.strip().upper() != "BORRAR"),
        ):
            delete_report_completely(report_id)
            clear_center_capture_state(reset_period=False)
            st.session_state["admin_delete_success"] = (
                f"Se eliminó el envío de {unit_code} · {month} {year}. "
                "También se limpió cualquier captura temporal de esta sesión."
            )
            st.rerun()



DAILY_QUOTES = [
    {
        "quote": "No perdamos nada del tiempo. Quizá los hubo más bellos, pero este es el nuestro.",
        "author": "Jean-Paul Sartre",
        "source": "Les Temps modernes (editorial inaugural, 1945)",
    },
    {
        "quote": "La vida no examinada no merece ser vivida.",
        "author": "Sócrates",
        "source": "Platón, Apología de Sócrates",
    },
    {
        "quote": "Somos lo que hacemos repetidamente.",
        "author": "Idea atribuida a Aristóteles",
        "source": "Síntesis moderna inspirada en la Ética nicomáquea",
    },
    {
        "quote": "La atención es la forma más rara y pura de generosidad.",
        "author": "Simone Weil",
        "source": "Cuadernos y correspondencia",
    },
    {
        "quote": "El hombre está condenado a ser libre.",
        "author": "Jean-Paul Sartre",
        "source": "El ser y la nada",
    },
    {
        "quote": "Lo que hacemos en la vida tiene su eco en la eternidad.",
        "author": "Marco Aurelio",
        "source": "Idea inspirada en Meditaciones",
    },
    {
        "quote": "No es que tengamos poco tiempo, sino que perdemos mucho.",
        "author": "Séneca",
        "source": "De la brevedad de la vida",
    },
    {
        "quote": "Yo soy yo y mi circunstancia.",
        "author": "José Ortega y Gasset",
        "source": "Meditaciones del Quijote",
    },
    {
        "quote": "La educación es un acto de amor, por tanto, un acto de valor.",
        "author": "Paulo Freire",
        "source": "Idea central de su obra pedagógica",
    },
    {
        "quote": "La esperanza necesita de la práctica para volverse historia concreta.",
        "author": "Paulo Freire",
        "source": "Pedagogía de la esperanza",
    },
    {
        "quote": "Conócete a ti mismo.",
        "author": "Máxima délfica",
        "source": "Templo de Apolo en Delfos",
    },
    {
        "quote": "La historia es un diálogo sin fin entre el presente y el pasado.",
        "author": "E. H. Carr",
        "source": "¿Qué es la historia?",
    },
]


def get_guadalajara_today():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("America/Mexico_City")).date()
    except Exception:
        return datetime.now().date()


def get_daily_quote(today):
    idx = today.toordinal() % len(DAILY_QUOTES)
    return DAILY_QUOTES[idx]


@st.cache_data(ttl=21600, show_spinner=False)
def fetch_today_ephemeris(month, day):
    """
    Retrieve a Spanish-language 'on this day' event from Wikipedia.
    Falls back gracefully if the external service is unavailable.
    """
    url = f"https://es.wikipedia.org/api/rest_v1/feed/onthisday/events/{month:02d}/{day:02d}"
    headers = {
        "User-Agent": "DIC-ITESO-Reportes/1.0 (educational internal app)"
    }
    try:
        response = requests.get(url, headers=headers, timeout=4)
        response.raise_for_status()
        data = response.json()
        events = data.get("events", [])

        # Prefer events with a year and a reasonably short description.
        candidates = [
            e for e in events
            if e.get("year") and e.get("text") and len(e.get("text", "")) <= 420
        ]
        if not candidates:
            candidates = [e for e in events if e.get("text")]

        if candidates:
            # Rotate deterministically so the same day is stable.
            chosen = candidates[(month * 31 + day) % len(candidates)]
            return {
                "year": chosen.get("year", ""),
                "text": chosen.get("text", "").strip(),
                "source": "Wikipedia · Efemérides del día",
            }
    except Exception:
        pass

    return {
        "year": "",
        "text": "Hoy también es una oportunidad para reconocer qué acontecimientos de nuestra comunidad merecen quedar registrados para el futuro.",
        "source": "Ecos e ideas del día",
    }


@st.dialog("Ecos e ideas del día", dismissible=False)
def daily_learning_capsule(email):
    today = get_guadalajara_today()
    quote = get_daily_quote(today)
    eph = fetch_today_ephemeris(today.month, today.day)

    st.caption(today.strftime("%d/%m/%Y"))

    st.markdown(
        f"""
        <div style="
            background:#F3F6F8;
            border-left:4px solid {ITESO_BLUE};
            padding:16px 18px;
            border-radius:8px;
            margin-bottom:16px;">
            <div style="font-size:1.05rem;font-weight:700;color:{ITESO_BLUE};margin-bottom:6px;">
                Para pensar
            </div>
            <div style="font-size:1.02rem;font-style:italic;line-height:1.45;">
                “{quote['quote']}”
            </div>
            <div style="margin-top:8px;font-size:.88rem;color:#64748B;">
                {quote['author']} · {quote['source']}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    year_text = f"{eph['year']} · " if eph.get("year") else ""
    st.markdown(
        f"""
        <div style="
            background:#FFFFFF;
            border:1px solid #D7DEE5;
            padding:16px 18px;
            border-radius:8px;">
            <div style="font-size:1.05rem;font-weight:700;color:{ITESO_BLUE};margin-bottom:6px;">
                Efeméride
            </div>
            <div style="line-height:1.45;">
                <b>{year_text}</b>{eph['text']}
            </div>
            <div style="margin-top:8px;font-size:.82rem;color:#7A8793;">
                {eph['source']}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption(
        "Una pausa breve antes de registrar el trabajo del mes: también construir memoria institucional es una forma de cuidar la comunidad."
    )

    if st.button("Comenzar reporte", type="primary", use_container_width=True):
        st.session_state.daily_capsule_seen_key = f"{email.lower()}|{today.isoformat()}"
        st.rerun()



def normalize_social_url(url):
    url = (url or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    return url


def detect_social_platform(url):
    url = normalize_social_url(url)
    if not url:
        return None
    host = (urlparse(url).netloc or "").lower()
    if "instagram.com" in host:
        return "Instagram"
    if "facebook.com" in host or "fb.watch" in host:
        return "Facebook"
    if "linkedin.com" in host:
        return "LinkedIn"
    return "Otro"


def social_url_is_valid(url):
    if not url:
        return True
    try:
        parsed = urlparse(normalize_social_url(url))
        return parsed.scheme in ("http", "https") and bool(parsed.netloc)
    except Exception:
        return False


def render_social_preview(url, key_suffix=""):
    """
    Best-effort preview for public social-media posts.
    Some platforms may block embedding depending on privacy/CSP settings;
    in that case the link remains available.
    """
    url = normalize_social_url(url)
    if not url:
        return

    platform = detect_social_platform(url)
    st.caption(f"Vista previa · {platform}")

    safe_url = html.escape(url, quote=True)

    if platform == "Instagram":
        embed_url = url.rstrip("/") + "/embed"
        st.components.v1.iframe(embed_url, height=520, scrolling=True)
    elif platform == "Facebook":
        embed_url = (
            "https://www.facebook.com/plugins/post.php?"
            f"href={quote(url, safe='')}&show_text=true&width=500"
        )
        st.components.v1.iframe(embed_url, height=540, scrolling=True)
    elif platform == "LinkedIn":
        # LinkedIn does not reliably allow arbitrary public post URLs in iframes.
        # Show a branded link card instead of failing the whole form.
        st.markdown(
            f"""
            <div style="border:1px solid #D7DEE5;border-radius:10px;padding:14px 16px;
                        background:#F8FAFC;margin-top:4px;margin-bottom:8px;">
              <div style="font-weight:700;color:#0A66C2;margin-bottom:5px;">LinkedIn</div>
              <div style="font-size:.9rem;color:#52606D;margin-bottom:10px;">
                La publicación se abrirá en LinkedIn para conservar su formato original.
              </div>
              <a href="{safe_url}" target="_blank" rel="noopener noreferrer"
                 style="color:#003B70;font-weight:600;text-decoration:none;">
                 Abrir publicación ↗
              </a>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div style="border:1px solid #D7DEE5;border-radius:10px;padding:14px 16px;
                        background:#F8FAFC;margin-top:4px;margin-bottom:8px;">
              <div style="font-weight:700;color:#003B70;margin-bottom:5px;">Enlace externo</div>
              <a href="{safe_url}" target="_blank" rel="noopener noreferrer"
                 style="color:#003B70;font-weight:600;text-decoration:none;">
                 Abrir publicación ↗
              </a>
            </div>
            """,
            unsafe_allow_html=True,
        )


def activity_fields_complete(index):
    record, missing = activity_from_session(
        index,
        actor_role=st.session_state.get("center_user_role"),
        month=st.session_state.get("capture_month"),
    )
    return record is not None and not missing



def upload_bytes(obj):
    if obj is None:
        return None
    if isinstance(obj, dict):
        return obj.get("bytes")
    try:
        return obj.getvalue()
    except Exception:
        return None


def upload_name(obj, fallback="archivo"):
    if isinstance(obj, dict):
        return obj.get("original_filename") or obj.get("name") or fallback
    return getattr(obj, "name", fallback)


def upload_mime(obj, fallback="image/jpeg"):
    if isinstance(obj, dict):
        return obj.get("mime_type") or fallback
    return getattr(obj, "type", fallback) or fallback



DOCX_TEMPLATE_FILENAME = "Plantilla_Importacion_Informe_DIC_ITESO.docx"
DOCX_GENERAL_FIELDS = ["CENTRO", "MES", "AÑO"]
DOCX_ACTIVITY_FIELDS = [
    "NOMBRE_ACTIVIDAD",
    "CATEGORIA",
    "CATEGORIA_OTRO",
    "IMPORTANCIA",
    "PARTICIPANTES_ALCANCE",
    "DESCRIPCION",
    "RED_SOCIAL_URL",
    "TITULO_GRAFICA",
    "GRAFICA",
    "FOTOGRAFIA",
]


DOCX_TEMPLATE_B64 = """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"""

def official_docx_template_bytes():
    """Return the approved DIC Word template embedded in the app itself."""
    return base64.b64decode(DOCX_TEMPLATE_B64)


def cell_all_text(cell):
    """Read visible text even when it lives inside Word content controls (SDTs)."""
    try:
        parts = cell._tc.xpath(".//w:t/text()")
        return "".join(parts).strip()
    except Exception:
        return (cell.text or "").strip()


def cell_images(cell, document):
    """Extract embedded images from a DOCX table cell."""
    images = []
    try:
        blips = cell._tc.xpath(".//a:blip")
    except Exception:
        blips = []
    for blip in blips:
        rel_id = blip.get(qn("r:embed"))
        if not rel_id:
            continue
        rel = document.part.rels.get(rel_id)
        if not rel:
            continue
        try:
            part = rel.target_part
            blob = part.blob
            content_type = getattr(part, "content_type", "image/jpeg") or "image/jpeg"
            partname = str(getattr(part, "partname", "imagen"))
            filename = Path(partname).name or "imagen"
            images.append({
                "bytes": blob,
                "mime_type": content_type,
                "original_filename": filename,
            })
        except Exception:
            continue
    return images


def normalize_docx_value(value):
    """Treat untouched Word dropdown placeholders as empty values."""
    text = (value or "").strip()
    if text.casefold() in {
        "selecciona una opción",
        "seleccionar una opción",
        "seleccione una opción",
    }:
        return ""
    return text


def compare_expected_labels(actual, expected, section_name):
    """Return human-readable structural differences for one table."""
    issues = []
    if actual == expected:
        return issues

    max_len = max(len(actual), len(expected))
    for pos in range(max_len):
        exp = expected[pos] if pos < len(expected) else None
        got = actual[pos] if pos < len(actual) else None
        if exp == got:
            continue
        if exp is None:
            issues.append(
                f"{section_name}: apareció un campo adicional `{got}` en la posición {pos + 1}."
            )
        elif got is None:
            issues.append(
                f"{section_name}: falta el campo `{exp}` en la posición {pos + 1}."
            )
        else:
            issues.append(
                f"{section_name}: se esperaba `{exp}` y se encontró `{got}` "
                f"en la posición {pos + 1}."
            )
    return issues


def parse_dic_docx(uploaded_bytes, expected_unit):
    """
    Validate the official DIC Word structure and extract report data.
    Structural changes block import; content issues are returned as warnings.
    """
    try:
        doc = Document(io.BytesIO(uploaded_bytes))
    except Exception as exc:
        return None, [f"No fue posible abrir el archivo como Word (.docx): {exc}"], []

    structural_errors = []
    content_warnings = []

    # Expected layout: instructions table + general table + 5 activity tables.
    if len(doc.tables) < 7:
        structural_errors.append(
            "El documento no contiene todas las tablas de la plantilla oficial. "
            "Vuelve a descargar la plantilla y captura la información sin eliminar secciones."
        )
        return None, structural_errors, content_warnings

    general_table = doc.tables[1]
    general_labels = [cell_all_text(r.cells[0]) for r in general_table.rows[1:]]
    structural_errors.extend(
        compare_expected_labels(general_labels, DOCX_GENERAL_FIELDS, "Datos del reporte")
    )

    activity_tables = doc.tables[2:7]
    if len(activity_tables) != 5:
        structural_errors.append(
            f"Se esperaban 5 bloques de actividad y se encontraron {len(activity_tables)}."
        )

    for i, table in enumerate(activity_tables[:5], start=1):
        labels = [cell_all_text(r.cells[0]) for r in table.rows[1:]]
        structural_errors.extend(
            compare_expected_labels(labels, DOCX_ACTIVITY_FIELDS, f"Actividad {i}")
        )

    if structural_errors:
        return None, structural_errors, content_warnings

    # General values.
    general = {}
    for row in general_table.rows[1:]:
        label = cell_all_text(row.cells[0])
        general[label] = cell_all_text(row.cells[1])

    center = normalize_docx_value(general.get("CENTRO", "")).upper()
    month = normalize_docx_value(general.get("MES", "")).title()
    year_raw = normalize_docx_value(general.get("AÑO", ""))

    if not center:
        content_warnings.append("El campo CENTRO está vacío.")
    elif center not in UNITS:
        content_warnings.append(f"El CENTRO `{center}` no es una sigla válida.")
    elif center != expected_unit:
        structural_errors.append(
            f"El archivo corresponde al centro `{center}`, pero estás ingresando como `{expected_unit}`. "
            "Carga la plantilla del centro que te corresponde."
        )

    if month not in MONTHS:
        content_warnings.append("El MES está vacío o no coincide con un mes válido.")

    try:
        year = int(year_raw)
    except Exception:
        year = datetime.now().year
        content_warnings.append("El AÑO está vacío o no es un número válido.")

    if structural_errors:
        return None, structural_errors, content_warnings

    activities = []
    for i, table in enumerate(activity_tables[:5], start=1):
        values = {}
        row_by_label = {}
        for row in table.rows[1:]:
            label = cell_all_text(row.cells[0])
            values[label] = cell_all_text(row.cells[1])
            row_by_label[label] = row

        title = normalize_docx_value(values.get("NOMBRE_ACTIVIDAD", ""))
        category = normalize_docx_value(values.get("CATEGORIA", ""))
        category_other = normalize_docx_value(values.get("CATEGORIA_OTRO", ""))
        ranking_text = normalize_docx_value(values.get("IMPORTANCIA", ""))
        participants_raw = normalize_docx_value(values.get("PARTICIPANTES_ALCANCE", ""))
        description = normalize_docx_value(values.get("DESCRIPCION", ""))
        social_url = normalize_docx_value(values.get("RED_SOCIAL_URL", ""))
        chart_title = normalize_docx_value(values.get("TITULO_GRAFICA", ""))

        chart_imgs = cell_images(row_by_label["GRAFICA"].cells[1], doc)
        photo_imgs = cell_images(row_by_label["FOTOGRAFIA"].cells[1], doc)

        # Completely unused activity: ignore it entirely.
        # Untouched Word dropdown placeholders have already been normalized to empty.
        started = any([
            title, category, category_other, ranking_text, participants_raw,
            description, social_url, chart_title, chart_imgs, photo_imgs,
        ])
        if not started:
            continue

        # If the activity has any real content, report only the fields that remain incomplete.
        missing_required = []
        if not title:
            missing_required.append("NOMBRE_ACTIVIDAD")
        if not category:
            missing_required.append("CATEGORIA")
        if not participants_raw:
            missing_required.append("PARTICIPANTES_ALCANCE")
        if not description:
            missing_required.append("DESCRIPCION")

        if missing_required:
            content_warnings.append(
                f"Actividad {i}: la actividad tiene contenido, pero faltan campos obligatorios: "
                + ", ".join(f"`{field}`" for field in missing_required)
                + ". Complétalos en el formulario antes de enviar."
            )

        # Category normalization.
        if category in CATEGORIES:
            category_selected = category
            other_category = category_other
        elif category:
            category_selected = "Otro"
            other_category = category_other or category
            content_warnings.append(
                f"Actividad {i}: la categoría `{category}` no coincide con la lista oficial; "
                "se importó como `Otro` para que puedas revisarla."
            )
        else:
            category_selected = CATEGORIES[0]
            other_category = ""

        # Ranking normalization.
        valid_ranks = ["Sin ranking", "Top 1", "Top 2", "Top 3"]
        if not ranking_text:
            ranking_text = "Sin ranking"
        elif ranking_text not in valid_ranks:
            content_warnings.append(
                f"Actividad {i}: IMPORTANCIA `{ranking_text}` no es válida; "
                "se dejó como `Sin ranking` para revisión."
            )
            ranking_text = "Sin ranking"

        # Participant normalization.
        try:
            participants = int(float(participants_raw.replace(",", ""))) if participants_raw else 0
        except Exception:
            participants = 0
            content_warnings.append(
                f"Actividad {i}: PARTICIPANTES_ALCANCE contiene `{participants_raw}` y debe ser un número entero."
            )

        if description and word_count(description) > 250:
            content_warnings.append(
                f"Actividad {i}: DESCRIPCION excede el máximo de 250 palabras."
            )

        if chart_imgs and not chart_title:
            content_warnings.append(
                f"Actividad {i}: se encontró una GRÁFICA, pero TITULO_GRAFICA está vacío."
            )

        if len(chart_imgs) > 1:
            content_warnings.append(
                f"Actividad {i}: se encontraron {len(chart_imgs)} imágenes en GRÁFICA; "
                "se utilizará la primera."
            )

        activities.append({
            "title": title,
            "category_selected": category_selected,
            "other_category": other_category,
            "ranking_text": ranking_text,
            "participants": participants,
            "description": description,
            "social_url": social_url,
            "chart_title": chart_title,
            "chart": chart_imgs[0] if chart_imgs else None,
            "photos": photo_imgs,
        })

    if not activities:
        content_warnings.append(
            "No se detectó ninguna actividad capturada en el documento."
        )

    return {
        "center": center or expected_unit,
        "month": month if month in MONTHS else st.session_state.get("capture_month", MONTHS[datetime.now().month - 1]),
        "year": year,
        "activities": activities,
    }, structural_errors, content_warnings


def hydrate_imported_docx(parsed):
    """Load imported Word contents into the same widgets used by manual capture."""
    # Clear any prior capture data.
    for key in list(st.session_state.keys()):
        if re.match(
            r"^(title|desc|cat|other_cat|rank|part|rubro|purpose|action_type|criteria|criteria_other|activity_date|"
            r"location|population|external_population|dic_collab|dic_units|relevance_note|media_type|media_platform|media_topic|media_link|"
            r"academic_semester|academic_credits|academic_opened|academic_groups|academic_fixed_prof|academic_variable_prof|faculty_name|faculty_contract|faculty_unit|faculty_category|faculty_promotion|"
            r"research_role|research_members|research_field|research_actors|research_start|research_end|research_progress|research_products|research_pct|research_on_plan|research_on_plan_why|"
            r"photos|social|chart|chart_title|existing_photos|existing_chart)_\d+$",
            key,
        ):
            st.session_state.pop(key, None)

    acts = parsed.get("activities", [])
    st.session_state.num_activities = max(5, len(acts) + 1)
    st.session_state.capture_month = parsed.get("month") or MONTHS[datetime.now().month - 1]
    st.session_state.capture_year = int(parsed.get("year") or datetime.now().year)
    st.session_state.resuming_report_id = None
    st.session_state.show_center_preview = False
    st.session_state.ranking_conflict_message = ""

    for i, act in enumerate(acts):
        st.session_state[f"title_{i}"] = act.get("title") or ""
        st.session_state[f"cat_{i}"] = act.get("category_selected") or CATEGORIES[0]
        st.session_state[f"other_cat_{i}"] = act.get("other_category") or ""
        st.session_state[f"rank_{i}"] = (
            (act.get("ranking_text") or "Sin ranking")
            if st.session_state.get("center_user_role") == "DIRECTOR"
            else "Sin ranking"
        )
        st.session_state[f"part_{i}"] = int(act.get("participants") or 0)
        st.session_state[f"desc_{i}"] = act.get("description") or ""
        st.session_state[f"social_{i}"] = act.get("social_url") or ""
        st.session_state[f"chart_title_{i}"] = act.get("chart_title") or ""
        st.session_state[f"existing_photos_{i}"] = act.get("photos") or []
        st.session_state[f"existing_chart_{i}"] = act.get("chart")

    st.session_state.docx_import_success = (
        f"Archivo `{st.session_state.get('docx_import_filename', 'Word')}` cargado correctamente. "
        "La información ya está disponible en el formulario para revisión."
    )




# ---------- IMPORTACIÓN WORD RELLENABLE V1.37 ----------
FILLABLE_DOCX_TEMPLATE_FILENAME = "Plantilla_Informe_Mensual_DIC_RELLENABLE.docx"
FILLABLE_DOCX_TEMPLATE_B64 = """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"""


def official_fillable_docx_template_bytes():
    return base64.b64decode(FILLABLE_DOCX_TEMPLATE_B64)


def _docx_control_map(docx_bytes):
    """Read tagged Word content controls (SDT) from the fillable template."""
    values = {}
    checked = {}
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    }
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            root = etree.fromstring(zf.read("word/document.xml"))
    except Exception as exc:
        raise ValueError(f"No fue posible leer la estructura del Word: {exc}")

    for sdt in root.xpath(".//w:sdt", namespaces=ns):
        pr = sdt.find("w:sdtPr", ns)
        if pr is None:
            continue
        tag_el = pr.find("w:tag", ns)
        if tag_el is None:
            continue
        tag = tag_el.get(f"{{{ns['w']}}}val", "").strip()
        if not tag:
            continue
        txt = "".join(sdt.xpath(".//w:t/text()", namespaces=ns)).strip()
        values[tag] = normalize_docx_value(txt)
        chk = pr.find("w14:checkbox", ns)
        is_checked = False
        if chk is not None:
            checked_el = chk.find("w14:checked", ns)
            if checked_el is not None:
                raw = (checked_el.get(f"{{{ns['w14']}}}val") or "").strip().lower()
                is_checked = raw in {"1", "true", "on"}
        if txt.strip() in {"☒", "☑", "✓", "✔"}:
            is_checked = True
        checked[tag] = is_checked
    return values, checked


def _docx_clean(value):
    text = normalize_docx_value(value)
    placeholders = {
        "nombre completo", "seleccione año", "dd/mm/aaaa", "dd/mm/aaaa o en blanco",
        "título breve de la acción", "lugar", "número", "haga clic y escriba",
        "describa qué ocurrió, resultados y actores", "una frase sobre por qué importa",
        "https://...", "sólo si el rubro es medios", "sólo si aplica", "curso", "nombre",
        "dependencia", "nombre del proyecto", "nombres", "temática", "actores internos/externos",
        "productos o difusión", "0-100", "título de la acción", "por qué es relevante",
        "objetivo / tema", "avance observado", "resultado / evidencia", "título o número de acción",
        "tema", "descripción", "qué se necesita", "responsable / instancia", "qué observamos",
        "qué implica", "próximo paso", "3 a 5 líneas",
    }
    return "" if text.strip().casefold() in placeholders else text.strip()


def _to_int(value, default=0):
    raw = _docx_clean(value).replace(",", "").replace("%", "").strip()
    if not raw:
        return default
    try:
        return int(float(raw))
    except Exception:
        return default


def _parse_word_date(value):
    raw = _docx_clean(value)
    if not raw:
        return datetime.now().date()
    return coerce_date_value(raw, fallback=datetime.now().date())


def _normalize_unit_from_word(value):
    raw = _docx_clean(value).strip()
    if not raw:
        return ""
    upper = raw.upper()
    if upper in UNITS:
        return upper
    for code, name in UNITS.items():
        if upper == name.upper() or upper.startswith(code + " ") or code in upper.split("·"):
            return code
    return raw


def _canonical_rubric(value):
    raw = _docx_clean(value)
    if raw in REPORT_RUBRICS:
        return raw
    # Accept numbered labels if Word preserves the explanatory prefix.
    for rubric in REPORT_RUBRICS:
        if rubric.casefold() in raw.casefold():
            return rubric
    return raw


def _extract_action_evidence(document, table_index, chart_title):
    """Extract images pasted into the evidence cell of a generic action table."""
    try:
        table = document.tables[table_index]
        evidence_cell = table.rows[17].cells[1]
        imgs = cell_images(evidence_cell, document)
    except Exception:
        imgs = []
    if not imgs:
        return [], None
    if chart_title:
        chart = imgs[0]
        photos = imgs[1:]
        return photos, chart
    return imgs, None


def parse_fillable_dic_docx(docx_bytes, expected_unit):
    """Parse the V1.35/V1.37 fillable DIC Word and return form-ready structured data."""
    try:
        doc = Document(io.BytesIO(docx_bytes))
    except Exception as exc:
        return None, [f"No fue posible abrir el archivo Word: {exc}"], []

    try:
        values, checked = _docx_control_map(docx_bytes)
    except Exception as exc:
        return None, [str(exc)], []

    errors = []
    warnings = []

    # Structural signature: key tags that uniquely identify this template.
    required_tags = ["GENERAL_CENTRO", "GENERAL_MES", "GENERAL_ANIO", "ACCION1_RUBRO", "SINTESIS_MES"]
    missing_tags = [tag for tag in required_tags if tag not in values]
    if missing_tags:
        errors.append(
            "El archivo no corresponde a la plantilla Word rellenable vigente. "
            "Faltan controles internos: " + ", ".join(missing_tags)
        )
        return None, errors, warnings

    center = _normalize_unit_from_word(values.get("GENERAL_CENTRO", ""))
    month = _docx_clean(values.get("GENERAL_MES", ""))
    year = _to_int(values.get("GENERAL_ANIO", ""), datetime.now().year)
    if center and expected_unit and center != expected_unit:
        errors.append(
            f"El Word corresponde al centro `{center}`, pero la sesión actual pertenece a `{expected_unit}`."
        )
    if month and month not in MONTHS:
        warnings.append(f"Mes `{month}` no reconocido; se conservará el mes actualmente seleccionado en la app.")
        month = st.session_state.get("capture_month", MONTHS[datetime.now().month - 1])
    if not month:
        month = st.session_state.get("capture_month", MONTHS[datetime.now().month - 1])

    activities = []

    criteria_tags = [
        ("CRIT_MISION", "Atiende la misión del centro"),
        ("CRIT_COLAB", "Trabajo colaborativo con otra instancia"),
        ("CRIT_PLANEACION", "Relevante — Planeación institucional"),
        ("CRIT_ENCARGO", "Relevante — Encargo de autoridades"),
        ("CRIT_OTRO", "Otro"),
    ]
    pop_tags = [
        ("POB_EST", "Estudiantes"), ("POB_EGR", "Egresados"),
        ("POB_ACA", "Académicos"), ("POB_ADM", "Personal administrativo"),
        ("POB_OPE", "Personal operativo"), ("POB_EXT", "Comunidad externa"),
    ]

    for n in range(1, 6):
        p = f"ACCION{n}_"
        rubro = _canonical_rubric(values.get(p + "RUBRO", ""))
        title = _docx_clean(values.get(p + "TITULO", ""))
        purpose = _docx_clean(values.get(p + "FIN", ""))
        action_type = _docx_clean(values.get(p + "TIPO", ""))
        planning = _docx_clean(values.get(p + "PLANEACION", ""))
        desc = _docx_clean(values.get(p + "DESCRIPCION", ""))
        relevance = _docx_clean(values.get(p + "IMPORTA", ""))
        link = _docx_clean(values.get(p + "ENLACE", ""))
        participants = _to_int(values.get(p + "PARTICIPANTES", ""), 0)
        criteria = [label for suffix, label in criteria_tags if checked.get(p + suffix, False)]
        population = [label for suffix, label in pop_tags if checked.get(p + suffix, False)]
        collab = _docx_clean(values.get(p + "DIC_COLAB", ""))
        dic_units = _docx_clean(values.get(p + "DIC_CUALES", ""))
        media_type = _docx_clean(values.get(p + "MEDIA_TIPO", ""))
        media_platform = _docx_clean(values.get(p + "MEDIA_PLATAFORMA", ""))
        media_topic = _docx_clean(values.get(p + "MEDIA_TEMA", ""))
        chart_title = _docx_clean(values.get(p + "GRAFICA_TITULO", ""))
        photos, chart = _extract_action_evidence(doc, 3 + n, chart_title)

        started = any([
            rubro, title, purpose, action_type, planning, desc, relevance, link,
            participants, criteria, population, media_type, media_platform, media_topic,
            photos, chart,
        ])
        if not started:
            continue
        if not rubro:
            warnings.append(f"Acción {n}: no se seleccionó rubro; revisa la acción después de importar.")
            rubro = CATEGORIES[0]
        if rubro not in REPORT_RUBRICS:
            warnings.append(f"Acción {n}: rubro `{rubro}` no reconocido; se importó como Vida universitaria para revisión.")
            rubro = CATEGORIES[0]

        activities.append({
            "rubro": rubro,
            "title": title,
            "purpose": purpose,
            "action_type": action_type,
            "planning_type": planning,
            "criteria": criteria,
            "activity_date": _parse_word_date(values.get(p + "FECHA", "")),
            "location": _docx_clean(values.get(p + "LUGAR", "")),
            "participants": participants,
            "population": population,
            "external_population": "",
            "dic_collab": "Sí" if collab.casefold() in {"sí", "si", "yes"} else "No",
            "dic_units": dic_units,
            "description": desc,
            "relevance_note": relevance,
            "social_url": link,
            "media_type": media_type,
            "media_platform": media_platform,
            "media_topic": media_topic or title,
            "media_link": link,
            "photos": photos,
            "chart": chart,
            "chart_title": chart_title,
        })

    # Academic section: one course becomes one app action.
    semester = _docx_clean(values.get("ACADEMICA_SEMESTRE", ""))
    faculty_rows = []
    for n in range(1, 5):
        faculty_rows.append({
            "name": _docx_clean(values.get(f"DOCENTE{n}_NOMBRE", "")),
            "contract": _docx_clean(values.get(f"DOCENTE{n}_CONTRATO", "")),
            "unit": _docx_clean(values.get(f"DOCENTE{n}_DEPENDENCIA", "")),
            "category": _docx_clean(values.get(f"DOCENTE{n}_CATEGORIA", "")),
            "promotion": _docx_clean(values.get(f"DOCENTE{n}_PROMOCION", "")),
        })
    course_count = 0
    for n in range(1, 6):
        title = _docx_clean(values.get(f"CURSO{n}_NOMBRE", ""))
        credits = _to_int(values.get(f"CURSO{n}_CREDITOS", ""), 0)
        opened = _docx_clean(values.get(f"CURSO{n}_ABRIO", ""))
        groups = _to_int(values.get(f"CURSO{n}_GRUPOS", ""), 0)
        fixed_prof = _to_int(values.get(f"CURSO{n}_PROF_TF", ""), 0)
        variable_prof = _to_int(values.get(f"CURSO{n}_PROF_TV", ""), 0)
        if not any([title, credits, opened, groups, fixed_prof, variable_prof]):
            continue
        course_count += 1
        fac = faculty_rows[n-1] if n <= len(faculty_rows) else {}
        activities.append({
            "rubro": "Oferta académica y docencia", "title": title,
            "academic_semester": semester, "academic_credits": credits,
            "academic_opened": opened, "academic_groups": groups,
            "academic_fixed_prof": fixed_prof, "academic_variable_prof": variable_prof,
            "faculty_name": fac.get("name", ""), "faculty_contract": fac.get("contract", ""),
            "faculty_unit": fac.get("unit", ""), "faculty_category": fac.get("category", ""),
            "faculty_promotion": fac.get("promotion", ""), "photos": [], "chart": None,
        })
    if course_count and month not in ACADEMIC_REPORTING_MONTHS:
        warnings.append(
            f"El Word contiene Oferta académica, pero `{month}` no está configurado como corte académico en la app. "
            "Los datos se importarán para revisión."
        )

    # Research section: one project becomes one app action.
    research_count = 0
    for n in range(1, 4):
        title = _docx_clean(values.get(f"INVEST{n}_PROYECTO", ""))
        members = _docx_clean(values.get(f"INVEST{n}_INTEGRANTES", ""))
        field = _docx_clean(values.get(f"INVEST{n}_TEMATICA", ""))
        products = _docx_clean(values.get(f"INVEST{n}_PRODUCTOS", ""))
        if not any([title, members, field, products]):
            continue
        research_count += 1
        activities.append({
            "rubro": "Investigación", "title": title,
            "research_role": _docx_clean(values.get(f"INVEST{n}_NIVEL", "")),
            "research_members": members,
            "research_field": field,
            "research_actors": _docx_clean(values.get(f"INVEST{n}_ACTORES", "")),
            "research_start": _parse_word_date(values.get(f"INVEST{n}_FECHA_INICIO", "")),
            "research_end": _parse_word_date(values.get(f"INVEST{n}_FECHA_FIN", "")),
            "research_progress": _docx_clean(values.get(f"INVEST{n}_AVANCE", "")),
            "research_products": products,
            "research_pct": _to_int(values.get(f"INVEST{n}_PORCENTAJE", ""), 0),
            "research_on_plan": _docx_clean(values.get(f"INVEST{n}_PLANEADO", "")),
            "research_on_plan_why": _docx_clean(values.get(f"INVEST{n}_PLANEADO_PORQUE", "")),
            "photos": [], "chart": None,
        })
    if research_count and month not in RESEARCH_REPORTING_MONTHS:
        warnings.append(
            f"El Word contiene Investigación, pero `{month}` no corresponde a enero/agosto. "
            "Los datos se importarán para revisión."
        )

    def resolve_link(raw):
        raw = _docx_clean(raw)
        if not raw:
            return None
        m = re.search(r"\b(\d{1,2})\b", raw)
        if m:
            idx = int(m.group(1))
            if 1 <= idx <= len(activities):
                return idx
        for idx, act in enumerate(activities, start=1):
            title = (act.get("title") or "").strip()
            if title and (title.casefold() == raw.casefold() or title.casefold() in raw.casefold()):
                return idx
        return None

    highlights = []
    for n in range(1, 4):
        title = _docx_clean(values.get(f"TOP{n}_TITULO", ""))
        why = _docx_clean(values.get(f"TOP{n}_PORQUE", ""))
        rubro = _canonical_rubric(values.get(f"TOP{n}_RUBRO", ""))
        if not any([title, why, rubro]):
            continue
        linked = resolve_link(title)
        if linked is None:
            warnings.append(f"Top {n}: no pude vincular `{title}` automáticamente con una acción; revísalo en la app.")
        highlights.append({"rank": n, "title": title, "why": why, "rubro": rubro, "linked_activity_order": linked})

    advances = []
    for n in range(1, 6):
        topic = _docx_clean(values.get(f"AVANCE{n}_TEMA", ""))
        progress = _docx_clean(values.get(f"AVANCE{n}_OBSERVADO", ""))
        evidence = _docx_clean(values.get(f"AVANCE{n}_EVIDENCIA", ""))
        if not any([topic, progress, evidence]):
            continue
        advances.append({
            "objective_topic": topic,
            "progress": progress,
            "evidence": evidence,
            "progress_level": _docx_clean(values.get(f"AVANCE{n}_NIVEL", "")) or "En proceso",
            "linked_activity_order": resolve_link(values.get(f"AVANCE{n}_ACCION", "")),
        })

    issues = []
    for n in range(1, 6):
        topic = _docx_clean(values.get(f"RIESGO{n}_TEMA", ""))
        desc = _docx_clean(values.get(f"RIESGO{n}_DESC", ""))
        if not any([topic, desc, _docx_clean(values.get(f"RIESGO{n}_NECESITA", ""))]):
            continue
        issues.append({
            "issue_type": _docx_clean(values.get(f"RIESGO{n}_TIPO", "")) or "Riesgo",
            "topic": topic,
            "description": desc,
            "priority": _docx_clean(values.get(f"RIESGO{n}_PRIORIDAD", "")) or "Medio",
            "need_type": _docx_clean(values.get(f"RIESGO{n}_NECESITA", "")) or "Seguimiento",
            "issue_status": _docx_clean(values.get(f"RIESGO{n}_ESTADO", "")) or "Abierto",
            "dependency": _docx_clean(values.get(f"RIESGO{n}_RESPONSABLE", "")),
            "linked_activity_order": resolve_link(values.get(f"RIESGO{n}_ACCION", "")),
        })

    learnings = []
    for n in range(1, 6):
        topic = _docx_clean(values.get(f"APRENDIZAJE{n}_TEMA", ""))
        observation = _docx_clean(values.get(f"APRENDIZAJE{n}_OBSERVAMOS", ""))
        if not any([topic, observation, _docx_clean(values.get(f"APRENDIZAJE{n}_IMPLICA", ""))]):
            continue
        learnings.append({
            "learning_type": _docx_clean(values.get(f"APRENDIZAJE{n}_TIPO", "")) or "Aprendizaje",
            "topic": topic,
            "observation": observation,
            "implication": _docx_clean(values.get(f"APRENDIZAJE{n}_IMPLICA", "")),
            "next_step": _docx_clean(values.get(f"APRENDIZAJE{n}_PASO", "")),
            "linked_activity_order": resolve_link(values.get(f"APRENDIZAJE{n}_ACCION", "")),
        })

    parsed = {
        "center": center or expected_unit,
        "month": month,
        "year": year,
        "director_name": _docx_clean(values.get("GENERAL_DIRECTOR", "")),
        "activities": activities,
        "media_summary": {
            "appearances": _to_int(values.get("MEDIOS_APARICIONES", ""), 0),
            "total_participations": _to_int(values.get("MEDIOS_PARTICIPACIONES", ""), 0),
            "reach": _to_int(values.get("MEDIOS_ALCANCE", ""), 0),
        },
        "highlights": highlights,
        "advances": advances,
        "issues": issues,
        "learnings": learnings,
        "summary": _docx_clean(values.get("SINTESIS_MES", "")),
    }
    if not activities:
        warnings.append("No se detectaron acciones, cursos ni proyectos de investigación capturados.")
    return parsed, errors, warnings


def hydrate_fillable_docx(parsed):
    """Populate the current V1.36 form from the parsed fillable Word without saving to Supabase."""
    # Clear current capture and structured reflection widget state.
    prefixes = (
        "title_", "desc_", "cat_", "rubro_", "purpose_", "action_type_", "planning_",
        "criteria_", "criteria_other_", "planning_", "activity_date_", "location_", "part_", "population_",
        "external_population_", "dic_collab_", "dic_units_", "relevance_note_", "media_type_",
        "media_platform_", "media_topic_", "media_link_", "academic_semester_", "academic_credits_",
        "academic_opened_", "academic_groups_", "academic_fixed_prof_", "academic_variable_prof_",
        "faculty_name_", "faculty_contract_", "faculty_unit_", "faculty_category_", "faculty_promotion_",
        "research_role_", "research_members_", "research_field_", "research_actors_", "research_start_",
        "research_end_", "research_progress_", "research_products_", "research_pct_", "research_on_plan_",
        "research_on_plan_why_", "photos_", "social_", "chart_", "chart_title_", "existing_photos_",
        "existing_chart_", "adv_topic_", "adv_progress_", "adv_evidence_", "adv_level_", "adv_link_",
        "issue_type_", "issue_topic_", "issue_desc_", "issue_priority_", "issue_need_", "issue_status_",
        "issue_dependency_", "issue_link_", "learn_type_", "learn_topic_", "learn_observation_",
        "learn_implication_", "learn_next_", "learn_link_", "highlight_reason_", "action_saved_ok_",
    )
    for key in list(st.session_state.keys()):
        if key.startswith(prefixes):
            st.session_state.pop(key, None)

    acts = parsed.get("activities") or []
    st.session_state.num_activities = max(5, len(acts) + 1)
    st.session_state.capture_month = parsed.get("month") or st.session_state.get("capture_month", MONTHS[datetime.now().month - 1])
    st.session_state.capture_year = int(parsed.get("year") or datetime.now().year)
    st.session_state.resuming_report_id = None
    st.session_state.show_center_preview = False
    st.session_state.ranking_conflict_message = ""

    for i, act in enumerate(acts):
        rubro = act.get("rubro") or CATEGORIES[0]
        st.session_state[f"rubro_{i}"] = rubro
        st.session_state[f"cat_{i}"] = rubro
        st.session_state[f"title_{i}"] = act.get("title") or ""
        st.session_state[f"purpose_{i}"] = act.get("purpose") or ""
        st.session_state[f"action_type_{i}"] = act.get("action_type") or ""
        st.session_state[f"planning_{i}"] = act.get("planning_type") or ""
        st.session_state[f"criteria_{i}"] = act.get("criteria") or []
        st.session_state[f"activity_date_{i}"] = act.get("activity_date") or datetime.now().date()
        st.session_state[f"location_{i}"] = act.get("location") or ""
        st.session_state[f"part_{i}"] = int(act.get("participants") or 0)
        st.session_state[f"population_{i}"] = act.get("population") or []
        st.session_state[f"external_population_{i}"] = act.get("external_population") or ""
        st.session_state[f"dic_collab_{i}"] = act.get("dic_collab") or "No"
        st.session_state[f"dic_units_{i}"] = act.get("dic_units") or ""
        st.session_state[f"desc_{i}"] = act.get("description") or ""
        st.session_state[f"relevance_note_{i}"] = act.get("relevance_note") or ""
        st.session_state[f"social_{i}"] = act.get("social_url") or ""
        st.session_state[f"media_type_{i}"] = act.get("media_type") or ""
        st.session_state[f"media_platform_{i}"] = act.get("media_platform") or ""
        st.session_state[f"media_topic_{i}"] = act.get("media_topic") or ""
        st.session_state[f"media_link_{i}"] = act.get("media_link") or ""
        st.session_state[f"academic_semester_{i}"] = act.get("academic_semester") or ""
        st.session_state[f"academic_credits_{i}"] = int(act.get("academic_credits") or 0)
        st.session_state[f"academic_opened_{i}"] = act.get("academic_opened") or ""
        st.session_state[f"academic_groups_{i}"] = int(act.get("academic_groups") or 0)
        st.session_state[f"academic_fixed_prof_{i}"] = int(act.get("academic_fixed_prof") or 0)
        st.session_state[f"academic_variable_prof_{i}"] = int(act.get("academic_variable_prof") or 0)
        st.session_state[f"faculty_name_{i}"] = act.get("faculty_name") or ""
        st.session_state[f"faculty_contract_{i}"] = act.get("faculty_contract") or "Tiempo Fijo"
        st.session_state[f"faculty_unit_{i}"] = act.get("faculty_unit") or ""
        st.session_state[f"faculty_category_{i}"] = act.get("faculty_category") or "Adjunto"
        st.session_state[f"faculty_promotion_{i}"] = act.get("faculty_promotion") or "No"
        st.session_state[f"research_role_{i}"] = act.get("research_role") or ""
        st.session_state[f"research_members_{i}"] = act.get("research_members") or ""
        st.session_state[f"research_field_{i}"] = act.get("research_field") or ""
        st.session_state[f"research_actors_{i}"] = act.get("research_actors") or ""
        st.session_state[f"research_start_{i}"] = act.get("research_start") or datetime.now().date()
        st.session_state[f"research_end_{i}"] = act.get("research_end") or datetime.now().date()
        st.session_state[f"research_progress_{i}"] = act.get("research_progress") or ""
        st.session_state[f"research_products_{i}"] = act.get("research_products") or ""
        st.session_state[f"research_pct_{i}"] = int(act.get("research_pct") or 0)
        st.session_state[f"research_on_plan_{i}"] = act.get("research_on_plan") or ""
        st.session_state[f"research_on_plan_why_{i}"] = act.get("research_on_plan_why") or ""
        st.session_state[f"existing_photos_{i}"] = act.get("photos") or []
        st.session_state[f"existing_chart_{i}"] = act.get("chart")
        st.session_state[f"chart_title_{i}"] = act.get("chart_title") or ""

    media = parsed.get("media_summary") or {}
    st.session_state.media_appearances = int(media.get("appearances") or 0)
    st.session_state.media_total_participations = int(media.get("total_participations") or 0)
    st.session_state.media_reach = int(media.get("reach") or 0)

    # Highlights are linked to the imported activity order.
    labels = [f"{idx+1}. {a.get('title') or a.get('rubro')}" for idx, a in enumerate(acts)]
    selected_labels = []
    for h in parsed.get("highlights") or []:
        order = h.get("linked_activity_order")
        if order and 1 <= int(order) <= len(labels):
            label = labels[int(order)-1]
            selected_labels.append(label)
            st.session_state[f"highlight_reason_{int(order)-1}"] = h.get("why") or ""
    st.session_state.highlight_selection = selected_labels[:3]

    advances = parsed.get("advances") or []
    issues = parsed.get("issues") or []
    learnings = parsed.get("learnings") or []
    st.session_state.num_strategic_advances = max(1, len(advances))
    st.session_state.num_strategic_issues = max(1, len(issues))
    st.session_state.num_strategic_learnings = max(1, len(learnings))
    st.session_state.no_strategic_issues = False
    st.session_state.monthly_strategic_summary = parsed.get("summary") or ""

    for idx, row in enumerate(advances):
        st.session_state[f"adv_topic_{idx}"] = row.get("objective_topic") or ""
        st.session_state[f"adv_progress_{idx}"] = row.get("progress") or ""
        st.session_state[f"adv_evidence_{idx}"] = row.get("evidence") or ""
        st.session_state[f"adv_level_{idx}"] = row.get("progress_level") or "En proceso"
        st.session_state[f"adv_link_{idx}"] = int(row.get("linked_activity_order") or 0)

    for idx, row in enumerate(issues):
        st.session_state[f"issue_type_{idx}"] = row.get("issue_type") or "Riesgo"
        st.session_state[f"issue_topic_{idx}"] = row.get("topic") or ""
        st.session_state[f"issue_desc_{idx}"] = row.get("description") or ""
        st.session_state[f"issue_priority_{idx}"] = row.get("priority") or "Medio"
        st.session_state[f"issue_need_{idx}"] = row.get("need_type") or "Seguimiento"
        st.session_state[f"issue_status_{idx}"] = row.get("issue_status") or "Abierto"
        st.session_state[f"issue_dependency_{idx}"] = row.get("dependency") or ""
        st.session_state[f"issue_link_{idx}"] = int(row.get("linked_activity_order") or 0)

    for idx, row in enumerate(learnings):
        st.session_state[f"learn_type_{idx}"] = row.get("learning_type") or "Aprendizaje"
        st.session_state[f"learn_topic_{idx}"] = row.get("topic") or ""
        st.session_state[f"learn_observation_{idx}"] = row.get("observation") or ""
        st.session_state[f"learn_implication_{idx}"] = row.get("implication") or ""
        st.session_state[f"learn_next_{idx}"] = row.get("next_step") or ""
        st.session_state[f"learn_link_{idx}"] = int(row.get("linked_activity_order") or 0)

    st.session_state.docx_import_success = (
        f"Word importado: {len(acts)} acción(es), {len(advances)} avance(s), "
        f"{len(issues)} situación(es) y {len(learnings)} aprendizaje(s). Revisa el formulario antes de guardar."
    )
    st.session_state.docx_import_warnings = st.session_state.get("fillable_docx_warnings", [])
    st.session_state.capture_method = "Importar Word rellenable"

def resume_draft(report):
    """Hydrate the capture form from a saved draft, including persisted media."""
    acts = get_activities(report["id"])

    # Clear previous capture widget state.
    keys_to_clear = []
    for key in list(st.session_state.keys()):
        if re.match(
            r"^(title|desc|cat|other_cat|rank|part|photos|social|chart|chart_title|"
            r"existing_photos|existing_chart|adv_topic|adv_progress|adv_evidence|adv_level|adv_link|"
            r"issue_type|issue_topic|issue_desc|issue_priority|issue_need|issue_dependency|issue_date|issue_link|issue_status|"
            r"learn_type|learn_topic|learn_observation|learn_implication|learn_next|learn_link)_\d+$",
            key,
        ):
            keys_to_clear.append(key)
    for key in keys_to_clear:
        st.session_state.pop(key, None)

    st.session_state.num_activities = max(5, len(acts) + 1)
    st.session_state.capture_month = report["month"]
    st.session_state.capture_year = int(report["year"])
    st.session_state.resuming_report_id = report["id"]
    st.session_state.show_center_preview = False
    st.session_state.ranking_conflict_message = ""

    for i, act in enumerate(acts):
        category = act.get("category") or CATEGORIES[0]
        st.session_state[f"rubro_{i}"] = category if category in REPORT_RUBRICS else CATEGORIES[0]
        st.session_state[f"cat_{i}"] = st.session_state[f"rubro_{i}"]
        st.session_state[f"title_{i}"] = act.get("title") or ""
        st.session_state[f"desc_{i}"] = act.get("description_original") or ""
        st.session_state[f"rank_{i}"] = rank_label(act.get("ranking"))
        st.session_state[f"part_{i}"] = int(act.get("participants") or 0)
        st.session_state[f"activity_date_{i}"] = coerce_date_value(act.get("activity_date"), fallback=datetime.now().date())
        st.session_state[f"purpose_{i}"] = act.get("action_purpose") or ""
        st.session_state[f"action_type_{i}"] = act.get("action_type") or ""
        st.session_state[f"planning_{i}"] = detail.get("planning_type") or ""
        st.session_state[f"criteria_{i}"] = act.get("inclusion_criteria") or []
        st.session_state[f"location_{i}"] = act.get("location") or ""
        st.session_state[f"population_{i}"] = act.get("target_population") or []
        st.session_state[f"external_population_{i}"] = act.get("target_external_name") or ""
        st.session_state[f"dic_collab_{i}"] = "Sí" if act.get("dic_collaboration") else "No"
        st.session_state[f"dic_units_{i}"] = act.get("dic_collaboration_units") or ""
        st.session_state[f"relevance_note_{i}"] = act.get("relevance_note") or ""
        st.session_state[f"social_{i}"] = act.get("social_url") or ""
        st.session_state[f"chart_title_{i}"] = act.get("chart_title") or ""
        detail = act.get("detail_data") or {}
        st.session_state[f"media_type_{i}"] = detail.get("participation_type") or MEDIA_TYPES[0]
        st.session_state[f"media_platform_{i}"] = detail.get("platform") or ""
        st.session_state[f"media_topic_{i}"] = detail.get("topic") or (act.get("title") or "")
        st.session_state[f"media_link_{i}"] = detail.get("link") or act.get("social_url") or ""
        st.session_state[f"academic_semester_{i}"] = detail.get("semester") or ACADEMIC_SEMESTERS[0]
        st.session_state[f"academic_credits_{i}"] = int(detail.get("credits") or 0)
        st.session_state[f"academic_opened_{i}"] = detail.get("opened") or "Sí"
        st.session_state[f"academic_groups_{i}"] = int(detail.get("groups") or 0)
        st.session_state[f"academic_fixed_prof_{i}"] = int(detail.get("fixed_professors") or 0)
        st.session_state[f"academic_variable_prof_{i}"] = int(detail.get("variable_professors") or 0)
        st.session_state[f"faculty_name_{i}"] = detail.get("faculty_name") or ""
        st.session_state[f"faculty_contract_{i}"] = detail.get("faculty_contract") or "Tiempo Fijo"
        st.session_state[f"faculty_unit_{i}"] = detail.get("faculty_unit") or ""
        st.session_state[f"faculty_category_{i}"] = detail.get("faculty_category") or "Adjunto"
        st.session_state[f"faculty_promotion_{i}"] = detail.get("faculty_promotion") or "No"
        st.session_state[f"research_role_{i}"] = detail.get("participation_level") or RESEARCH_PARTICIPATION[0]
        st.session_state[f"research_start_{i}"] = coerce_date_value(detail.get("start_date"), fallback=datetime.now().date())
        st.session_state[f"research_end_{i}"] = coerce_date_value(detail.get("end_date"), fallback=datetime.now().date())
        st.session_state[f"research_members_{i}"] = detail.get("members") or ""
        st.session_state[f"research_field_{i}"] = detail.get("field") or ""
        st.session_state[f"research_actors_{i}"] = detail.get("linked_actors") or ""
        st.session_state[f"research_progress_{i}"] = detail.get("progress_level") or RESEARCH_PROGRESS[0]
        st.session_state[f"research_products_{i}"] = detail.get("products") or ""
        st.session_state[f"research_pct_{i}"] = int(detail.get("products_progress_pct") or 0)
        st.session_state[f"research_on_plan_{i}"] = detail.get("on_plan") or "Sí"
        st.session_state[f"research_on_plan_why_{i}"] = detail.get("on_plan_why") or ""

        saved_photos = get_activity_photos(act["id"])
        st.session_state[f"existing_photos_{i}"] = saved_photos
        saved_chart = get_activity_chart(act)
        st.session_state[f"existing_chart_{i}"] = saved_chart

    highlights = report.get("monthly_highlights") or []
    st.session_state["highlight_selection"] = [h.get("label") for h in highlights if h.get("label")]
    for idx, h in enumerate(highlights):
        st.session_state[f"highlight_reason_{idx}"] = h.get("why") or ""
    st.session_state["learning_planning_advances"] = report.get("learning_planning_advances") or ""
    st.session_state["learning_risks"] = report.get("learning_risks") or ""
    st.session_state["learning_opportunity"] = report.get("learning_opportunity") or ""

    structured = get_structured_reflection(report.get("id"))
    advances = structured.get("advances") or []
    issues = structured.get("issues") or []
    learnings = structured.get("learnings") or []
    st.session_state.num_strategic_advances = max(1, len(advances))
    st.session_state.num_strategic_issues = max(1, len(issues))
    st.session_state.num_strategic_learnings = max(1, len(learnings))
    st.session_state.no_strategic_issues = bool(report.get("no_strategic_issues"))
    st.session_state.monthly_strategic_summary = report.get("monthly_strategic_summary") or ""

    for idx, row in enumerate(advances):
        st.session_state[f"adv_topic_{idx}"] = row.get("objective_topic") or ""
        st.session_state[f"adv_progress_{idx}"] = row.get("progress") or ""
        st.session_state[f"adv_evidence_{idx}"] = row.get("evidence") or ""
        st.session_state[f"adv_level_{idx}"] = row.get("progress_level") or "En proceso"
        linked_order = row.get("linked_activity_order")
        st.session_state[f"adv_link_{idx}"] = int(linked_order) if linked_order else 0

    for idx, row in enumerate(issues):
        st.session_state[f"issue_type_{idx}"] = row.get("issue_type") or "Riesgo"
        st.session_state[f"issue_topic_{idx}"] = row.get("topic") or ""
        st.session_state[f"issue_desc_{idx}"] = row.get("description") or ""
        st.session_state[f"issue_priority_{idx}"] = row.get("priority") or "Medio"
        st.session_state[f"issue_need_{idx}"] = row.get("need_type") or "Seguimiento"
        st.session_state[f"issue_dependency_{idx}"] = row.get("dependency") or ""
        st.session_state[f"issue_status_{idx}"] = row.get("issue_status") or "Abierto"
        linked_order = row.get("linked_activity_order")
        st.session_state[f"issue_link_{idx}"] = int(linked_order) if linked_order else 0

    for idx, row in enumerate(learnings):
        st.session_state[f"learn_type_{idx}"] = row.get("learning_type") or "Aprendizaje"
        st.session_state[f"learn_topic_{idx}"] = row.get("topic") or ""
        st.session_state[f"learn_observation_{idx}"] = row.get("observation") or ""
        st.session_state[f"learn_implication_{idx}"] = row.get("implication") or ""
        st.session_state[f"learn_next_{idx}"] = row.get("next_step") or ""
        linked_order = row.get("linked_activity_order")
        st.session_state[f"learn_link_{idx}"] = int(linked_order) if linked_order else 0

    media_summary = report.get("media_monthly_summary") or {}
    st.session_state["media_appearances"] = int(media_summary.get("appearances") or 0)
    st.session_state["media_total_participations"] = int(media_summary.get("total_participations") or 0)
    st.session_state["media_reach"] = int(media_summary.get("reach") or 0)

    st.session_state.director_page = "Nuevo reporte"
    st.session_state.capture_method = "Captura en línea"
    st.session_state.docx_import_success = ""
    st.session_state.docx_import_warnings = []


def current_uploaded_photos(activity):
    normalized = []
    for ph in (activity.get("existing_photos", []) or []) + (activity.get("photos", []) or []):
        data = upload_bytes(ph)
        if not data:
            continue
        normalized.append({
            "bytes": data,
            "mime_type": upload_mime(ph),
            "original_filename": upload_name(ph, "fotografia"),
        })
    return normalized



def current_uploaded_chart(activity):
    chart = activity.get("chart") or activity.get("existing_chart")
    if not chart:
        return None
    data = upload_bytes(chart)
    if not data:
        return None
    return {
        "bytes": data,
        "mime_type": upload_mime(chart),
        "original_filename": upload_name(chart, "grafica.jpg"),
        "title": activity.get("chart_title") or (
            chart.get("title", "") if isinstance(chart, dict) else ""
        ),
    }



def get_activity_chart(activity):
    """Return the saved chart image for an activity, if any."""
    path = activity.get("chart_storage_path")
    if not path:
        if activity.get("chart_bytes"):
            return {
                "bytes": activity["chart_bytes"],
                "mime_type": activity.get("chart_mime_type", "image/jpeg"),
                "original_filename": activity.get("chart_original_filename", "grafica.jpg"),
                "title": activity.get("chart_title") or "",
            }
        return None

    if supabase:
        try:
            data = supabase.storage.from_("dic-activity-photos").download(path)
            return {
                "bytes": data,
                "mime_type": "image/png" if str(path).lower().endswith(".png") else "image/jpeg",
                "original_filename": activity.get("chart_original_filename") or Path(path).name,
                "title": activity.get("chart_title") or "",
            }
        except Exception:
            return None
    return None


def add_reportlab_image(story, image_bytes, max_w=430, max_h=300, title=None, title_style=None):
    """Append an image to a ReportLab story, preserving aspect ratio."""
    try:
        if title:
            story.append(Spacer(1, 8))
            story.append(Paragraph(f"<b>{html.escape(title)}</b>", title_style))
        img = Image(io.BytesIO(image_bytes))
        ratio = min(max_w / img.imageWidth, max_h / img.imageHeight, 1)
        img.drawWidth = img.imageWidth * ratio
        img.drawHeight = img.imageHeight * ratio
        story += [Spacer(1, 7), img, Spacer(1, 7)]
    except Exception:
        pass


def format_report_datetime(value):
    """Format report timestamps in Guadalajara time."""
    if not value:
        return ""
    try:
        from datetime import timezone
        from zoneinfo import ZoneInfo

        if isinstance(value, str):
            dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        else:
            dt = value

        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)

        dt = dt.astimezone(ZoneInfo("America/Mexico_City"))
        months_short = {
            1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
            7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
        }
        return f"{dt.day} {months_short[dt.month]} {dt.year}, {dt.strftime('%H:%M')}"
    except Exception:
        return str(value)


def email_is_iteso(email):
    return bool(re.match(r"^[A-Za-z0-9._%+-]+@iteso\.mx$", (email or "").strip(), re.I))

def _supabase_config():
    try:
        return (
            (st.secrets.get("SUPABASE_URL", "") or "").rstrip("/"),
            st.secrets.get("SUPABASE_KEY", "") or "",
        )
    except Exception:
        return "", ""


def _auth_headers(access_token=None, admin=False):
    url, key = _supabase_config()
    headers = {"apikey": key, "Content-Type": "application/json"}
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"
    elif admin:
        headers["Authorization"] = f"Bearer {key}"
    return url, headers


def get_authorized_user(email):
    email = (email or "").strip().lower()
    if not email or not supabase:
        return None
    try:
        rows = (
            supabase.table("authorized_users")
            .select("*")
            .eq("email", email)
            .eq("active", True)
            .limit(1)
            .execute()
            .data or []
        )
        return rows[0] if rows else None
    except Exception:
        return None


def update_authorized_login(email, auth_user_id=None):
    if not supabase:
        return
    payload = {
        "last_login_at": datetime.utcnow().isoformat(),
        "updated_at": datetime.utcnow().isoformat(),
    }
    if auth_user_id:
        payload["auth_user_id"] = auth_user_id
    try:
        supabase.table("authorized_users").update(payload).eq("email", email.lower()).execute()
    except Exception:
        pass


def auth_sign_in(email, password):
    """Authenticate against Supabase Auth, then enforce the DIC authorization table."""
    email = (email or "").strip().lower()
    authorized = get_authorized_user(email)
    if not authorized:
        return False, "Este correo no está autorizado para acceder a la plataforma.", None

    url, headers = _auth_headers()
    if not url or not headers.get("apikey"):
        return False, "Supabase no está configurado correctamente.", None
    try:
        response = requests.post(
            f"{url}/auth/v1/token?grant_type=password",
            headers=headers,
            json={"email": email, "password": password},
            timeout=12,
        )
        if response.status_code >= 400:
            return False, "Correo o contraseña incorrectos.", None
        data = response.json()
        user = data.get("user") or {}
        auth_user_id = user.get("id")
        update_authorized_login(email, auth_user_id)
        authorized = get_authorized_user(email) or authorized
        return True, "", authorized
    except Exception as exc:
        return False, f"No fue posible iniciar sesión. Detalle: {exc}", None


def auth_logout_center():
    clear_center_capture_state(reset_period=False)
    for key in [
        "center_authenticated", "center_user_email", "center_user_name",
        "center_user_unit", "center_user_role", "validated_center_email",
    ]:
        st.session_state[key] = False if key == "center_authenticated" else ""
    st.session_state.director_page = "Nuevo reporte"
    clear_session_activity()


ACTIVATION_CODE_TTL_HOURS = 168  # 7 días
ACTIVATION_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"


def _activation_hash(code):
    normalized = re.sub(r"[^A-Z0-9]", "", (code or "").upper())
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def generate_activation_code():
    """Generate a one-time activation/reset code without ambiguous characters."""
    raw = "".join(secrets.choice(ACTIVATION_ALPHABET) for _ in range(10))
    return f"{raw[:5]}-{raw[5:]}"


def set_activation_code(email):
    """Create a new one-time activation/reset code. Only its hash is stored."""
    if not supabase:
        return False, "Supabase no está disponible.", None, None
    email = (email or "").strip().lower()
    user = get_authorized_user(email)
    if not user:
        return False, "El usuario no está activo o no existe.", None, None

    from datetime import timedelta, timezone
    code = generate_activation_code()
    now = datetime.now(timezone.utc)
    expires = now + timedelta(hours=ACTIVATION_CODE_TTL_HOURS)
    try:
        supabase.table("authorized_users").update({
            "activation_code_hash": _activation_hash(code),
            "activation_code_created_at": now.isoformat(),
            "activation_code_expires_at": expires.isoformat(),
            "updated_at": now.isoformat(),
        }).eq("email", email).execute()
        return True, "", code, expires.isoformat()
    except Exception as exc:
        return False, f"No fue posible generar el código: {exc}", None, None


def find_auth_user_id_by_email(email):
    """Find an existing Supabase Auth user without exposing the secret key to the browser."""
    url, headers = _auth_headers(admin=True)
    if not url or not headers.get("apikey"):
        return None
    try:
        response = requests.get(
            f"{url}/auth/v1/admin/users",
            headers=headers,
            params={"page": 1, "per_page": 1000},
            timeout=12,
        )
        if response.status_code >= 400:
            return None
        payload = response.json()
        users = payload.get("users", payload if isinstance(payload, list) else [])
        for user in users or []:
            if (user.get("email") or "").strip().lower() == (email or "").strip().lower():
                return user.get("id")
    except Exception:
        pass
    return None


def delete_authorized_user(email):
    """Delete access from authorized_users and, when present, remove the Supabase Auth identity.

    Historical reports/activities are intentionally preserved because they are not FK-linked to
    authorized_users/auth.users.
    """
    if not supabase:
        return False, "Supabase no está disponible."

    email = (email or "").strip().lower()
    if not email:
        return False, "Correo de usuario inválido."

    # Resolve Auth identity before deleting the authorization row.
    try:
        row_res = (
            supabase.table("authorized_users")
            .select("email,auth_user_id")
            .eq("email", email)
            .limit(1)
            .execute()
        )
        row = (row_res.data or [None])[0]
    except Exception as exc:
        return False, f"No fue posible consultar al usuario antes de eliminarlo: {exc}"

    if not row:
        return False, "El usuario ya no existe en la lista de accesos."

    auth_user_id = row.get("auth_user_id") or find_auth_user_id_by_email(email)

    # First revoke application authorization. Even if Auth cleanup later fails, the user can no
    # longer enter because every login is checked against authorized_users.
    try:
        supabase.table("authorized_users").delete().eq("email", email).execute()
    except Exception as exc:
        return False, f"No fue posible eliminar el acceso: {exc}"

    auth_warning = ""
    if auth_user_id:
        try:
            url, headers = _auth_headers(admin=True)
            if url and headers.get("apikey"):
                response = requests.delete(
                    f"{url}/auth/v1/admin/users/{auth_user_id}",
                    headers=headers,
                    timeout=12,
                )
                if response.status_code >= 400 and response.status_code != 404:
                    try:
                        detail = response.json()
                    except Exception:
                        detail = response.text
                    auth_warning = f" La autorización fue eliminada, pero Supabase Auth respondió: {detail}"
            else:
                auth_warning = " La autorización fue eliminada, pero no se pudo limpiar Supabase Auth."
        except Exception as exc:
            auth_warning = f" La autorización fue eliminada, pero no se pudo limpiar Supabase Auth: {exc}"

    return True, "Usuario eliminado. Sus reportes históricos se conservaron." + auth_warning


@st.dialog("Eliminar usuario", dismissible=False)
def delete_authorized_user_dialog(user):
    email = (user or {}).get("email", "")
    name = (user or {}).get("name", "")
    unit = (user or {}).get("unit_code", "")
    role = (user or {}).get("role", "")

    st.warning(
        f"Se eliminará el acceso de **{name or email}** · {unit} · {role}. "
        "La persona dejará de poder entrar a la aplicación. Sus reportes históricos NO se eliminarán."
    )
    confirmation = st.text_input("Escribe BORRAR para confirmar", key=f"delete_user_confirm_{email}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Cancelar", use_container_width=True, key=f"cancel_delete_user_{email}"):
            st.rerun()
    with c2:
        if st.button(
            "Eliminar definitivamente",
            type="primary",
            use_container_width=True,
            disabled=(confirmation.strip().upper() != "BORRAR"),
            key=f"confirm_delete_user_{email}",
        ):
            ok, msg = delete_authorized_user(email)
            if ok:
                st.session_state["admin_user_delete_success"] = msg
                st.rerun()
            else:
                st.error(msg)


def admin_set_auth_password(email, password, name="", unit_code="", role="", auth_user_id=None):
    """Create or update the Supabase Auth identity after a valid activation code."""
    url, headers = _auth_headers(admin=True)
    if not url or not headers.get("apikey"):
        return False, "Supabase no está configurado correctamente.", None

    email = (email or "").strip().lower()
    auth_user_id = auth_user_id or find_auth_user_id_by_email(email)
    payload = {
        "password": password,
        "email_confirm": True,
        "user_metadata": {"name": name, "unit_code": unit_code, "role": role},
    }
    try:
        if auth_user_id:
            response = requests.put(
                f"{url}/auth/v1/admin/users/{auth_user_id}",
                headers=headers,
                json=payload,
                timeout=12,
            )
        else:
            response = requests.post(
                f"{url}/auth/v1/admin/users",
                headers=headers,
                json={"email": email, **payload},
                timeout=12,
            )
            # If an Auth identity already exists from an earlier test, recover and update it.
            if response.status_code >= 400:
                recovered_id = find_auth_user_id_by_email(email)
                if recovered_id:
                    auth_user_id = recovered_id
                    response = requests.put(
                        f"{url}/auth/v1/admin/users/{auth_user_id}",
                        headers=headers,
                        json=payload,
                        timeout=12,
                    )

        if response.status_code >= 400:
            try:
                detail = response.json()
            except Exception:
                detail = response.text
            return False, f"No fue posible crear/actualizar la cuenta de acceso: {detail}", None

        data = response.json() if response.text else {}
        user_obj = data.get("user") if isinstance(data, dict) and isinstance(data.get("user"), dict) else data
        resolved_id = (user_obj or {}).get("id") or auth_user_id
        return True, "", resolved_id
    except Exception as exc:
        return False, f"No fue posible crear/actualizar la cuenta de acceso: {exc}", None


def activate_or_reset_account(email, activation_code, new_password):
    """Validate an admin-issued code and let the authorized user establish a password."""
    from datetime import timezone
    email = (email or "").strip().lower()
    user = get_authorized_user(email)
    if not user:
        return False, "Correo o código de activación incorrectos.", None

    stored_hash = user.get("activation_code_hash") or ""
    if not stored_hash:
        return False, "No hay un código de activación vigente para este usuario. Solicita uno al administrador.", None

    if not hmac.compare_digest(stored_hash, _activation_hash(activation_code)):
        return False, "Correo o código de activación incorrectos.", None

    try:
        expiry_raw = user.get("activation_code_expires_at")
        if not expiry_raw:
            return False, "El código de activación ya no es válido. Solicita uno nuevo.", None
        expiry = datetime.fromisoformat(str(expiry_raw).replace("Z", "+00:00"))
        if expiry.tzinfo is None:
            expiry = expiry.replace(tzinfo=timezone.utc)
        if datetime.now(timezone.utc) > expiry:
            return False, "El código de activación venció. Solicita uno nuevo al administrador.", None
    except Exception:
        return False, "No fue posible validar la vigencia del código. Solicita uno nuevo.", None

    if len(new_password or "") < 10:
        return False, "La contraseña debe tener al menos 10 caracteres.", None

    ok, msg, auth_user_id = admin_set_auth_password(
        email=email,
        password=new_password,
        name=user.get("name", ""),
        unit_code=user.get("unit_code", ""),
        role=user.get("role", ""),
        auth_user_id=user.get("auth_user_id"),
    )
    if not ok:
        return False, msg, None

    now = datetime.now(timezone.utc).isoformat()
    try:
        supabase.table("authorized_users").update({
            "auth_user_id": auth_user_id,
            "activated_at": user.get("activated_at") or now,
            "activation_code_hash": None,
            "activation_code_created_at": None,
            "activation_code_expires_at": None,
            "updated_at": now,
        }).eq("email", email).execute()
    except Exception as exc:
        return False, f"La contraseña se creó, pero no fue posible cerrar la activación: {exc}", None

    return True, "", get_authorized_user(email) or user


def authorized_users_template_bytes():
    if XLWorkbook is None or DataValidation is None:
        return b""
    wb = XLWorkbook()
    ws = wb.active
    ws.title = "Usuarios"
    ws.append(["nombre", "correo", "centro", "rol"])
    ws.append(["Nombre del director", "director.ejemplo@iteso.mx", "CUE", "DIRECTOR"])
    ws.append(["Nombre del colaborador", "colaborador.ejemplo@iteso.mx", "CUE", "COLABORADOR"])
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True, color="FFFFFF")
        cell.fill = cell.fill.copy(fill_type="solid", fgColor="003B70")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 20
    ws.freeze_panes = "A2"

    cat = wb.create_sheet("Catalogos")
    cat.append(["CENTROS", "ROLES"])
    for idx, code in enumerate(UNITS.keys(), start=2):
        cat.cell(idx, 1, code)
    cat.cell(2, 2, "COLABORADOR")
    cat.cell(3, 2, "DIRECTOR")
    dv_center = DataValidation(type="list", formula1="=Catalogos!$A$2:$A$8", allow_blank=False)
    dv_role = DataValidation(type="list", formula1="=Catalogos!$B$2:$B$3", allow_blank=False)
    ws.add_data_validation(dv_center); dv_center.add("C2:C500")
    ws.add_data_validation(dv_role); dv_role.add("D2:D500")

    ins = wb.create_sheet("Instrucciones")
    instructions = [
        ["Plantilla de usuarios autorizados · DIC", ""],
        ["Campo", "Cómo llenarlo"],
        ["nombre", "Nombre completo de la persona autorizada."],
        ["correo", "Correo institucional @iteso.mx. Un correo por persona."],
        ["centro", "Selecciona una de las 7 siglas disponibles."],
        ["rol", "COLABORADOR o DIRECTOR."],
        ["Importante", "Sólo el DIRECTOR puede asignar Top 1/2/3 y enviar el informe mensual."],
        ["Carga", "Puedes volver a cargar esta plantilla desde Administración para agregar o actualizar accesos."],
    ]
    for row in instructions:
        ins.append(row)
    ins.column_dimensions["A"].width = 22
    ins.column_dimensions["B"].width = 78
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def parse_authorized_users_excel(file_bytes):
    if load_workbook is None:
        return [], ["Falta la dependencia openpyxl en requirements.txt."]
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as exc:
        return [], [f"No se pudo abrir el Excel: {exc}"]
    if "Usuarios" not in wb.sheetnames:
        return [], ["El archivo debe contener una hoja llamada `Usuarios`."]
    ws = wb["Usuarios"]
    headers = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, 5)]
    if headers != ["nombre", "correo", "centro", "rol"]:
        return [], ["La hoja Usuarios debe tener exactamente: nombre, correo, centro, rol."]

    rows, errors, seen = [], [], set()
    director_by_unit = {}
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, 1).value or "").strip()
        email = str(ws.cell(r, 2).value or "").strip().lower()
        unit = str(ws.cell(r, 3).value or "").strip().upper()
        role = str(ws.cell(r, 4).value or "").strip().upper()
        if not any([name, email, unit, role]):
            continue
        row_errors = []
        if not name: row_errors.append("nombre vacío")
        if not email_is_iteso(email): row_errors.append("correo @iteso.mx inválido")
        if unit not in UNITS: row_errors.append("centro inválido")
        if role not in {"COLABORADOR", "DIRECTOR"}: row_errors.append("rol inválido")
        if email in seen: row_errors.append("correo duplicado")
        if role == "DIRECTOR" and unit in director_by_unit:
            row_errors.append(f"ya existe otro DIRECTOR para {unit} en este archivo")
        if row_errors:
            errors.append(f"Fila {r}: " + ", ".join(row_errors) + ".")
            continue
        seen.add(email)
        if role == "DIRECTOR": director_by_unit[unit] = email
        rows.append({"name": name, "email": email, "unit_code": unit, "role": role})
    if not rows and not errors:
        errors.append("El archivo no contiene usuarios para cargar.")
    return rows, errors


def list_authorized_users(include_inactive=True):
    if not supabase:
        return []
    try:
        q = supabase.table("authorized_users").select("*")
        if not include_inactive:
            q = q.eq("active", True)
        return q.order("unit_code").order("role").order("name").execute().data or []
    except Exception:
        return []


def active_director_conflict(unit_code, email):
    """Return the other active director for a center, if one exists."""
    unit_code = (unit_code or "").strip().upper()
    email = (email or "").strip().lower()
    if not supabase or not unit_code:
        return None
    try:
        rows = (
            supabase.table("authorized_users")
            .select("email,name,unit_code,role,active")
            .eq("unit_code", unit_code)
            .eq("role", "DIRECTOR")
            .eq("active", True)
            .execute()
            .data or []
        )
        for row in rows:
            if (row.get("email") or "").strip().lower() != email:
                return row
    except Exception:
        return None
    return None


def set_authorized_user_active(user, new_state):
    """Activate/deactivate a user with human-readable validation errors."""
    if not supabase:
        return False, "Supabase no está disponible."
    if not user:
        return False, "El usuario no existe."

    email = (user.get("email") or "").strip().lower()
    role = (user.get("role") or "").strip().upper()
    unit_code = (user.get("unit_code") or "").strip().upper()

    if new_state and role == "DIRECTOR":
        conflict = active_director_conflict(unit_code, email)
        if conflict:
            other = conflict.get("name") or conflict.get("email")
            return False, (
                f"No se puede reactivar este Director porque {other} ya figura como Director activo de {unit_code}. "
                "Desactiva, elimina o cambia de rol al Director actual antes de continuar."
            )

    try:
        res = (
            supabase.table("authorized_users")
            .update({
                "active": bool(new_state),
                "updated_at": datetime.utcnow().isoformat(),
            })
            .eq("email", email)
            .execute()
        )
        if not (res.data or []):
            return False, "No se encontró el usuario para actualizar su acceso."
        return True, "Acceso reactivado." if new_state else "Acceso desactivado."
    except Exception as exc:
        text = str(exc)
        if "authorized_users_one_active_director_per_unit" in text or "duplicate key" in text.lower():
            return False, (
                f"Ya existe otro Director activo para {unit_code}. "
                "Desactiva, elimina o cambia de rol al Director actual antes de reactivar este acceso."
            )
        return False, f"No fue posible actualizar el acceso: {exc}"


def upsert_authorized_users(rows, generate_codes=True):
    """Add/update authorized users and optionally issue one-time activation codes."""
    existing_rows = list_authorized_users(include_inactive=True)
    existing = {(u.get("email") or "").strip().lower(): u for u in existing_rows}
    results = []
    generated_codes = []
    now = datetime.utcnow().isoformat()

    for row in rows:
        email = (row.get("email") or "").strip().lower()
        unit_code = (row.get("unit_code") or "").strip().upper()
        role = (row.get("role") or "").strip().upper()
        previous = existing.get(email)

        # The database also enforces this rule, but validating here gives the administrator
        # a useful message instead of a Postgres exception.
        if role == "DIRECTOR":
            conflict = active_director_conflict(unit_code, email)
            if conflict:
                other = conflict.get("name") or conflict.get("email")
                results.append((
                    email, False,
                    f"{unit_code} ya tiene un Director activo: {other}. "
                    "Desactiva, elimina o cambia de rol al Director actual antes de cargar uno nuevo."
                ))
                continue

        payload = {
            "name": (row.get("name") or "").strip(),
            "email": email,
            "unit_code": unit_code,
            "role": role,
            "active": True,
            "updated_at": now,
        }
        if not previous:
            payload["created_at"] = now

        try:
            res = supabase.table("authorized_users").upsert(payload, on_conflict="email").execute()
            if not (res.data or []):
                results.append((email, False, "Supabase no confirmó el alta/actualización del usuario."))
                continue

            status = "Usuario actualizado" if previous else "Usuario agregado"
            should_generate = generate_codes and (not previous or not previous.get("auth_user_id"))
            if should_generate:
                ok, msg, code, expires = set_activation_code(email)
                if ok:
                    status = "Usuario activo y código temporal generado"
                    generated_codes.append({
                        "Nombre": payload.get("name", ""),
                        "Correo": email,
                        "Centro": unit_code,
                        "Rol": role,
                        "Código": code,
                        "Vence": format_access_datetime(expires),
                    })
                else:
                    status = f"Usuario guardado, pero {msg}"
            results.append((email, True, status))
            existing[email] = {**(previous or {}), **payload}
        except Exception as exc:
            text = str(exc)
            if "authorized_users_one_active_director_per_unit" in text or "duplicate key" in text.lower():
                message = (
                    f"{unit_code} ya tiene otro Director activo. "
                    "Desactiva, elimina o cambia de rol al Director actual antes de continuar."
                )
            else:
                message = f"No fue posible guardar el usuario: {exc}"
            results.append((email, False, message))

    return results, generated_codes


def format_access_datetime(value):
    if not value:
        return "Nunca"
    try:
        from datetime import timezone
        from zoneinfo import ZoneInfo
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        dt = dt.astimezone(ZoneInfo("America/Mexico_City"))
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(value)

def send_confirmation_email(to_email, unit, month, year):
    # V1: hook preparado. Si existe RESEND_API_KEY se puede activar.
    # Para evitar dependencias externas en el prototipo, por ahora registra éxito lógico.
    return {
        "ok": True,
        "message": f"Confirmación preparada para {to_email}: {unit}, {month} {year}."
    }

def generate_word(month, year, reports, activities_by_report):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)
    add_docx_page_x_of_y(section)

    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(logo), width=Inches(2.25))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Informe mensual Equipo de Consulta")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0, 76, 127)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p2.add_run(f"{month} {year}\nDirección de Integración Comunitaria")
    r.italic = True
    r.font.size = Pt(13)
    doc.add_paragraph("Hitos por centro")
    for rep in reports:
        doc.add_heading(f"{rep['unit_code']} — {UNITS.get(rep['unit_code'], rep['unit_code'])}", level=2)
        acts = [
            a for a in sorted(activities_by_report.get(rep["id"], []), key=activity_sort_key)
            if selected_for_final(a["id"])
        ]
        if not acts:
            continue
        for act in acts:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(act["title"])
            r.bold = True
            if act.get("ranking"):
                r.add_text(f" ({rank_label(act['ranking'])})")
            p.add_run(" — " + (act.get("description_edited") or act.get("description_original") or ""))
            if act.get("social_url"):
                doc.add_paragraph(f"Redes sociales: {act['social_url']}")

            chart = get_activity_chart(act)
            if chart:
                pchart = doc.add_paragraph()
                rchart = pchart.add_run(chart.get("title") or "Gráfica")
                rchart.bold = True
                try:
                    doc.add_picture(io.BytesIO(chart["bytes"]), width=Inches(5.7))
                except Exception:
                    pass

            photos = get_activity_photos(act["id"])
            for ph in photos:
                try:
                    doc.add_picture(io.BytesIO(ph["bytes"]), width=Inches(5.7))
                except Exception:
                    pass

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def generate_pdf(month, year, reports, activities_by_report):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter, rightMargin=45, leftMargin=45, topMargin=45, bottomMargin=45)
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "dicTitle", parent=styles["Title"], textColor=HexColor(ITESO_BLUE),
        fontSize=18, leading=22, alignment=TA_CENTER, spaceAfter=6
    )
    h2 = ParagraphStyle(
        "dicH2", parent=styles["Heading2"], textColor=HexColor(ITESO_BLUE),
        fontSize=13, leading=16, spaceBefore=10, spaceAfter=5
    )
    body = styles["BodyText"]
    body.fontSize = 9.5
    body.leading = 13

    story = []
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        img = Image(str(logo), width=190, height=42.5)
        img.hAlign = "RIGHT"
        story += [img, Spacer(1, 8)]
    story += [
        Paragraph("Informe mensual Equipo de Consulta", title),
        Paragraph(f"<i>{month} {year}</i><br/>Dirección de Integración Comunitaria", styles["Heading3"]),
        Spacer(1, 12),
        Paragraph("Hitos por centro", h2),
    ]
    for rep in reports:
        story.append(Paragraph(f"{rep['unit_code']} — {UNITS.get(rep['unit_code'], '')}", h2))
        acts = [
            a for a in sorted(activities_by_report.get(rep["id"], []), key=activity_sort_key)
            if selected_for_final(a["id"])
        ]
        if not acts:
            continue
        for act in acts:
            tag = f" <b>({rank_label(act['ranking'])})</b>" if act.get("ranking") else ""
            story.append(Paragraph(
                f"• <b>{act['title']}</b>{tag} — {act.get('description_edited') or act.get('description_original') or ''}",
                body
            ))
            story.append(Spacer(1, 5))
            if act.get("social_url"):
                story.append(Paragraph(f"<b>Redes sociales:</b> {html.escape(act['social_url'])}", body))
            chart = get_activity_chart(act)
            if chart:
                add_reportlab_image(
                    story, chart["bytes"], max_w=430, max_h=300,
                    title=chart.get("title") or "Gráfica", title_style=body
                )

            photos = get_activity_photos(act["id"])
            for ph in photos:
                try:
                    img = Image(io.BytesIO(ph["bytes"]))
                    max_w, max_h = 430, 280
                    ratio = min(max_w / img.imageWidth, max_h / img.imageHeight, 1)
                    img.drawWidth = img.imageWidth * ratio
                    img.drawHeight = img.imageHeight * ratio
                    story += [Spacer(1, 7), img, Spacer(1, 7)]
                except Exception:
                    pass

    doc.build(story, canvasmaker=NumberedCanvas)
    bio.seek(0)
    return bio.getvalue()


def generate_segment_word(rep, act, photos):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)
    add_docx_page_x_of_y(sec)

    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(logo), width=Inches(2.2))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = p.add_run("Extracto de actividad · DIC")
    rr.bold = True
    rr.font.size = Pt(17)
    rr.font.color.rgb = RGBColor(0, 76, 127)

    meta = doc.add_paragraph()
    meta.add_run(f"{rep['unit_code']} — {UNITS.get(rep['unit_code'], '')}\n").bold = True
    meta.add_run(f"{rep['month']} {rep['year']} · {rank_label(act.get('ranking'))}\n")
    meta.add_run(f"Categoría: {act.get('category','')}")

    h = doc.add_paragraph()
    r = h.add_run(act.get("title",""))
    r.bold = True
    r.font.size = Pt(14)

    doc.add_paragraph(act.get("description_edited") or act.get("description_original") or "")

    if act.get("participants"):
        doc.add_paragraph(f"Participantes / alcance: {act['participants']}")
    if act.get("social_url"):
        doc.add_paragraph(f"Redes sociales: {act['social_url']}")
    chart = get_activity_chart(act)
    if chart:
        pchart = doc.add_paragraph()
        pchart.add_run(chart.get("title") or "Gráfica").bold = True
        try:
            doc.add_picture(io.BytesIO(chart["bytes"]), width=Inches(5.8))
        except Exception:
            pass

    for ph in photos:
        try:
            bio = io.BytesIO(ph["bytes"])
            doc.add_picture(bio, width=Inches(5.8))
        except Exception:
            pass

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def generate_segment_pdf(rep, act, photos):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter, rightMargin=45, leftMargin=45, topMargin=45, bottomMargin=45)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "segTitle", parent=styles["Title"], textColor=HexColor(ITESO_BLUE),
        fontSize=17, leading=21, alignment=TA_CENTER, spaceAfter=12
    )
    h_style = ParagraphStyle(
        "segH", parent=styles["Heading2"], textColor=HexColor(ITESO_BLUE),
        fontSize=13, leading=16, spaceAfter=8
    )
    body = styles["BodyText"]
    body.fontSize = 10
    body.leading = 14

    story = []
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        img = Image(str(logo), width=190, height=45)
        img.hAlign = "RIGHT"
        story += [img, Spacer(1, 8)]

    story += [
        Paragraph("Extracto de actividad · DIC", title_style),
        Paragraph(
            f"<b>{rep['unit_code']} — {UNITS.get(rep['unit_code'], '')}</b><br/>"
            f"{rep['month']} {rep['year']} · {rank_label(act.get('ranking'))}<br/>"
            f"Categoría: {act.get('category','')}",
            body
        ),
        Spacer(1, 10),
        Paragraph(act.get("title",""), h_style),
        Paragraph(act.get("description_edited") or act.get("description_original") or "", body),
    ]
    if act.get("participants"):
        story += [Spacer(1, 6), Paragraph(f"Participantes / alcance: {act['participants']}", body)]
    if act.get("social_url"):
        story += [Spacer(1, 6), Paragraph(f"<b>Redes sociales:</b> {html.escape(act['social_url'])}", body)]
    chart = get_activity_chart(act)
    if chart:
        add_reportlab_image(
            story, chart["bytes"], max_w=430, max_h=300,
            title=chart.get("title") or "Gráfica", title_style=body
        )

    for ph in photos:
        try:
            img = Image(io.BytesIO(ph["bytes"]))
            max_w, max_h = 430, 300
            ratio = min(max_w / img.imageWidth, max_h / img.imageHeight, 1)
            img.drawWidth = img.imageWidth * ratio
            img.drawHeight = img.imageHeight * ratio
            story += [Spacer(1, 10), img]
        except Exception:
            pass

    doc.build(story, canvasmaker=NumberedCanvas)
    return bio.getvalue()



def generate_center_word(unit, month, year, activities, report_extras=None):
    report_extras = report_extras or {}
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55); sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.65); sec.right_margin = Inches(0.65)
    add_docx_page_x_of_y(sec)
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(logo), width=Inches(2.2))
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Informe mensual de centro a la DIC"); r.bold = True; r.font.size = Pt(18); r.font.color.rgb = RGBColor(0,76,127)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"Dirección de Integración Comunitaria\n{month} {year}")
    doc.add_paragraph(f"{unit} — {UNITS[unit]}").runs[0].bold = True
    doc.add_heading("Memoria del mes", level=1)
    for i, act in enumerate(activities, start=1):
        h = doc.add_paragraph(); rr = h.add_run(f"{i}. {act['title']}"); rr.bold = True; rr.font.color.rgb = RGBColor(0,76,127)
        meta = [act.get('category',''), rank_label(act.get('ranking'))]
        doc.add_paragraph(" · ".join([x for x in meta if x]))
        for label, value in activity_detail_lines(act):
            pdet = doc.add_paragraph(); rb = pdet.add_run(f"{label}: "); rb.bold = True; pdet.add_run(str(value))
        if act.get("description"):
            doc.add_paragraph(act.get("description") or "")
        if act.get("social_url"):
            doc.add_paragraph(f"Enlace: {act['social_url']}")
        chart = current_uploaded_chart(act)
        if chart:
            pchart = doc.add_paragraph(); rchart = pchart.add_run(chart.get("title") or "Gráfica"); rchart.bold = True
            try: doc.add_picture(io.BytesIO(chart["bytes"]), width=Inches(5.7))
            except Exception: pass
        for ph in current_uploaded_photos(act):
            try: doc.add_picture(io.BytesIO(ph["bytes"]), width=Inches(5.7))
            except Exception: pass

    highlights = report_extras.get("monthly_highlights") or []
    if highlights:
        doc.add_heading("Lo más relevante del mes", level=1)
        for h in highlights:
            p = doc.add_paragraph(); rb = p.add_run(f"Top {h.get('ranking')} · {h.get('title','')}"); rb.bold = True
            doc.add_paragraph(h.get("why") or "")

    media = report_extras.get("media_monthly_summary") or {}
    if media:
        doc.add_heading("Datos concentrados de medios del mes", level=1)
        doc.add_paragraph(f"Apariciones en medios: {media.get('appearances',0)}")
        doc.add_paragraph(f"Total de participaciones: {media.get('total_participations',0)}")
        doc.add_paragraph(f"Alcance o audiencia: {media.get('reach',0)}")

    doc.add_heading("Aprendizajes", level=1)
    for label, key in [
        ("Avances sustantivos en los objetivos de planeación", "learning_planning_advances"),
        ("Acciones paradas, en riesgo o que requieren decisiones", "learning_risks"),
        ("Un aprendizaje u oportunidad derivado del trabajo mensual", "learning_opportunity"),
    ]:
        p = doc.add_paragraph(); rb = p.add_run(label); rb.bold = True
        doc.add_paragraph(report_extras.get(key) or "")
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()


def generate_center_pdf(unit, month, year, activities, report_extras=None):
    report_extras = report_extras or {}
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter, rightMargin=45, leftMargin=45, topMargin=45, bottomMargin=45)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("centerTitleNew", parent=styles["Title"], textColor=HexColor(ITESO_BLUE), fontSize=18, leading=22, alignment=TA_CENTER, spaceAfter=8)
    h_style = ParagraphStyle("centerHNew", parent=styles["Heading2"], textColor=HexColor(ITESO_BLUE), fontSize=13, leading=16, spaceBefore=10, spaceAfter=6)
    body = styles["BodyText"]; body.fontSize = 9.5; body.leading = 13
    story = []
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        img = Image(str(logo), width=190, height=45); img.hAlign = "RIGHT"; story += [img, Spacer(1,8)]
    story += [Paragraph("Informe mensual de centro a la DIC", title_style), Paragraph(f"Dirección de Integración Comunitaria<br/>{month} {year}", styles["Heading3"]), Spacer(1,10), Paragraph(f"{unit} — {UNITS[unit]}", h_style), Paragraph("Memoria del mes", h_style)]
    for i, act in enumerate(activities, start=1):
        story.append(Paragraph(f"{i}. {html.escape(act['title'])}", h_style))
        story.append(Paragraph(html.escape(act.get('category','')) + " · " + rank_label(act.get('ranking')), body))
        for label, value in activity_detail_lines(act):
            story.append(Paragraph(f"<b>{html.escape(str(label))}:</b> {html.escape(str(value))}", body))
        if act.get("description"):
            story.append(Paragraph(html.escape(act.get("description") or ""), body))
        if act.get("social_url"):
            story.append(Paragraph(f"<b>Enlace:</b> {html.escape(act['social_url'])}", body))
        chart = current_uploaded_chart(act)
        if chart:
            add_reportlab_image(story, chart["bytes"], max_w=430, max_h=300, title=chart.get("title") or "Gráfica", title_style=body)
        for ph in current_uploaded_photos(act):
            try:
                img = Image(io.BytesIO(ph["bytes"])); max_w,max_h = 430,280
                ratio = min(max_w/img.imageWidth, max_h/img.imageHeight, 1)
                img.drawWidth = img.imageWidth*ratio; img.drawHeight = img.imageHeight*ratio
                story += [Spacer(1,8), img]
            except Exception: pass
        story.append(Spacer(1,8))
    highlights = report_extras.get("monthly_highlights") or []
    if highlights:
        story.append(Paragraph("Lo más relevante del mes", h_style))
        for h in highlights:
            story.append(Paragraph(f"<b>Top {h.get('ranking')} · {html.escape(h.get('title',''))}</b>", body))
            story.append(Paragraph(html.escape(h.get("why") or ""), body))
    media = report_extras.get("media_monthly_summary") or {}
    if media:
        story.append(Paragraph("Datos concentrados de medios del mes", h_style))
        story.append(Paragraph(f"Apariciones en medios: {media.get('appearances',0)} · Total de participaciones: {media.get('total_participations',0)} · Alcance/audiencia: {media.get('reach',0)}", body))
    story.append(Paragraph("Aprendizajes", h_style))
    for label, key in [
        ("Avances sustantivos en los objetivos de planeación", "learning_planning_advances"),
        ("Acciones paradas, en riesgo o que requieren decisiones", "learning_risks"),
        ("Un aprendizaje u oportunidad derivado del trabajo mensual", "learning_opportunity"),
    ]:
        story.append(Paragraph(f"<b>{label}</b>", body))
        story.append(Paragraph(html.escape(report_extras.get(key) or ""), body))
    doc.build(story, canvasmaker=NumberedCanvas); return bio.getvalue()


@st.dialog("Confirmar envío", dismissible=False)
def confirm_submission_dialog(unit, month, year, sender_email, activities, actor_role=None, report_extras=None):
    actor_role = (actor_role or st.session_state.get("center_user_role") or "").upper()
    if actor_role != "DIRECTOR":
        st.error("Sólo el Director del centro puede enviar el informe mensual.")
        if st.button("Cerrar", use_container_width=True):
            st.rerun()
        return
    st.write(
        f"Vas a enviar el reporte de **{UNITS[unit]}** correspondiente a **{month} {year}**."
    )
    st.warning(
        "Confirma que ya registraste todas las actividades que deseas reportar. "
        "Después del envío el reporte quedará marcado como enviado."
    )
    confirm_all = st.checkbox("Confirmo que ya no tengo más actividades que agregar.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Volver", use_container_width=True):
            st.rerun()
    with c2:
        if st.button(
            "Sí, enviar reporte",
            type="primary",
            use_container_width=True,
            disabled=not confirm_all,
        ):
            # First persist as BORRADOR. Only mark ENVIADO after all photos/charts
            # have been written and read back successfully.
            rid = save_report(unit, month, year, "BORRADOR", sender_email, report_extras=report_extras)
            persistence_errors = replace_activities(rid, activities, actor_role=actor_role)

            if persistence_errors:
                st.error(
                    "El reporte NO se marcó como enviado porque uno o más archivos "
                    "no pudieron guardarse correctamente."
                )
                for err in persistence_errors:
                    st.markdown(f"- {err}")
                st.info(
                    "La información textual quedó como borrador. Corrige el problema con "
                    "las imágenes y vuelve a intentar el envío."
                )
                return

            save_report(unit, month, year, "ENVIADO", sender_email, report_extras=report_extras)
            send_confirmation_email(sender_email, unit, month, year)

            # Date/time shown in Guadalajara local time.
            try:
                from zoneinfo import ZoneInfo
                sent_dt = datetime.now(ZoneInfo("America/Mexico_City"))
                sent_text = f"{sent_dt.strftime('%d/%m/%Y')} · {sent_dt.strftime('%H:%M')}"
            except Exception:
                sent_dt = datetime.now()
                sent_text = f"{sent_dt.strftime('%d/%m/%Y')} · {sent_dt.strftime('%H:%M')}"

            st.session_state["submission_receipt"] = {
                "period": f"{month} {year}",
                "sent_at": sent_text,
            }
            st.rerun()


@st.dialog("Muchas gracias por tu envío", dismissible=False)
def submission_success_dialog():
    receipt = st.session_state.get("submission_receipt", {})
    st.markdown("### Reporte enviado")
    st.markdown(
        f"""
        **Periodo:** {receipt.get('period', '')}  
        **Fecha · Hora de envío:** {receipt.get('sent_at', '')}
        """
    )
    st.success("La información quedó registrada correctamente.")

    if st.button("Cerrar sesión", type="primary", use_container_width=True):
        st.session_state.validated_center_email = ""
        st.session_state.daily_capsule_seen_key = ""
        st.session_state.director_page = "Nuevo reporte"
        clear_center_capture_state(reset_period=True)

        if "submission_receipt" in st.session_state:
            del st.session_state["submission_receipt"]

        st.rerun()



def generate_saved_center_word(rep, activities):
    """Generate a center report from activities already saved in DB/session."""
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)
    add_docx_page_x_of_y(sec)

    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(logo), width=Inches(2.2))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Informe de actividades")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0, 59, 112)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(
        f"Dirección de Integración Comunitaria\n{rep['month']} {rep['year']}"
    )
    r2.font.size = Pt(12)

    unit = rep["unit_code"]
    p3 = doc.add_paragraph()
    rr = p3.add_run(f"{unit} — {UNITS.get(unit, unit)}")
    rr.bold = True

    for i, act in enumerate(sorted(activities, key=activity_sort_key), start=1):
        h = doc.add_paragraph()
        title_run = h.add_run(f"{i}. {act.get('title','')}")
        title_run.bold = True
        title_run.font.size = Pt(14)
        title_run.font.color.rgb = RGBColor(0, 59, 112)

        meta = []
        if act.get("category"):
            meta.append(f"Categoría: {act['category']}")
        meta.append(rank_label(act.get("ranking")))
        if act.get("participants"):
            meta.append(f"Participantes / alcance: {act['participants']}")
        doc.add_paragraph(" · ".join(meta))
        for label, value in activity_detail_lines(act):
            pdet = doc.add_paragraph()
            rb = pdet.add_run(f"{label}: ")
            rb.bold = True
            pdet.add_run(str(value))

        doc.add_paragraph(
            act.get("description_original")
            or act.get("description_edited")
            or ""
        )
        if act.get("social_url"):
            doc.add_paragraph(f"Redes sociales: {act['social_url']}")

        chart = get_activity_chart(act)
        if chart:
            pchart = doc.add_paragraph()
            rchart = pchart.add_run(chart.get("title") or "Gráfica")
            rchart.bold = True
            try:
                doc.add_picture(io.BytesIO(chart["bytes"]), width=Inches(5.7))
            except Exception:
                pass

        photos = get_activity_photos(act["id"])
        for ph in photos:
            try:
                doc.add_picture(io.BytesIO(ph["bytes"]), width=Inches(5.7))
            except Exception:
                pass

    highlights = rep.get("monthly_highlights") or []
    if highlights:
        doc.add_heading("Lo más relevante del mes", level=1)
        for h in highlights:
            p = doc.add_paragraph()
            rb = p.add_run(f"Top {h.get('ranking')} · {h.get('title','')}")
            rb.bold = True
            doc.add_paragraph(h.get("why") or "")
    doc.add_heading("Aprendizajes", level=1)
    for label, key in [
        ("Avances sustantivos en los objetivos de planeación", "learning_planning_advances"),
        ("Acciones paradas, en riesgo o que requieren decisiones", "learning_risks"),
        ("Un aprendizaje u oportunidad derivado del trabajo mensual", "learning_opportunity"),
    ]:
        p = doc.add_paragraph(); rb = p.add_run(label); rb.bold = True
        doc.add_paragraph(rep.get(key) or "")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


def generate_saved_center_pdf(rep, activities):
    """Generate a PDF center report from activities already saved in DB/session."""
    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio,
        pagesize=letter,
        rightMargin=45,
        leftMargin=45,
        topMargin=45,
        bottomMargin=45,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "savedCenterTitle",
        parent=styles["Title"],
        textColor=HexColor(ITESO_BLUE),
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        spaceAfter=8,
    )
    h_style = ParagraphStyle(
        "savedCenterH",
        parent=styles["Heading2"],
        textColor=HexColor(ITESO_BLUE),
        fontSize=13,
        leading=16,
        spaceBefore=10,
        spaceAfter=6,
    )
    body = styles["BodyText"]
    body.fontSize = 9.5
    body.leading = 13

    story = []
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        img = Image(str(logo), width=190, height=45)
        img.hAlign = "RIGHT"
        story += [img, Spacer(1, 8)]

    unit = rep["unit_code"]
    story += [
        Paragraph("Informe de actividades", title_style),
        Paragraph(
            f"Dirección de Integración Comunitaria<br/>{rep['month']} {rep['year']}",
            styles["Heading3"],
        ),
        Spacer(1, 10),
        Paragraph(f"{unit} — {UNITS.get(unit, unit)}", h_style),
    ]

    for i, act in enumerate(sorted(activities, key=activity_sort_key), start=1):
        story.append(Paragraph(f"{i}. {act.get('title','')}", h_style))

        meta = []
        if act.get("category"):
            meta.append(f"Categoría: {act['category']}")
        meta.append(rank_label(act.get("ranking")))
        if act.get("participants"):
            meta.append(f"Participantes / alcance: {act['participants']}")
        story.append(Paragraph(" · ".join(meta), body))

        story.append(
            Paragraph(
                act.get("description_original")
                or act.get("description_edited")
                or "",
                body,
            )
        )
        if act.get("social_url"):
            story.append(Paragraph(f"<b>Redes sociales:</b> {html.escape(act['social_url'])}", body))

        chart = get_activity_chart(act)
        if chart:
            add_reportlab_image(
                story, chart["bytes"], max_w=430, max_h=300,
                title=chart.get("title") or "Gráfica", title_style=body
            )

        photos = get_activity_photos(act["id"])
        for ph in photos:
            try:
                img = Image(io.BytesIO(ph["bytes"]))
                max_w, max_h = 430, 280
                ratio = min(max_w / img.imageWidth, max_h / img.imageHeight, 1)
                img.drawWidth = img.imageWidth * ratio
                img.drawHeight = img.imageHeight * ratio
                story += [Spacer(1, 8), img]
            except Exception:
                pass

        story.append(Spacer(1, 8))

    highlights = rep.get("monthly_highlights") or []
    if highlights:
        story.append(Paragraph("Lo más relevante del mes", h_style))
        for h in highlights:
            story.append(Paragraph(f"<b>Top {h.get('ranking')} · {html.escape(h.get('title',''))}</b>", body))
            story.append(Paragraph(html.escape(h.get("why") or ""), body))
    story.append(Paragraph("Aprendizajes", h_style))
    for label, key in [
        ("Avances sustantivos en los objetivos de planeación", "learning_planning_advances"),
        ("Acciones paradas, en riesgo o que requieren decisiones", "learning_risks"),
        ("Un aprendizaje u oportunidad derivado del trabajo mensual", "learning_opportunity"),
    ]:
        story.append(Paragraph(f"<b>{label}</b>", body))
        story.append(Paragraph(html.escape(rep.get(key) or ""), body))
    doc.build(story, canvasmaker=NumberedCanvas)
    bio.seek(0)
    return bio.getvalue()



def _ppt_add_footer(slide, page_num=None):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    footer = slide.shapes.add_textbox(Inches(0.35), Inches(7.05), Inches(12.6), Inches(0.25))
    tf = footer.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "Dirección de Integración Comunitaria · ITESO"
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(100, 116, 139)

    if page_num:
        pn = slide.shapes.add_textbox(Inches(12.15), Inches(7.05), Inches(0.8), Inches(0.25))
        tfp = pn.text_frame
        tfp.clear()
        pp = tfp.paragraphs[0]
        rr = pp.add_run()
        rr.text = str(page_num)
        rr.font.size = Pt(8)
        rr.font.color.rgb = RGBColor(100, 116, 139)


def _ppt_add_logo(slide, prs):
    from pptx.util import Inches
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        slide.shapes.add_picture(str(logo), Inches(0.45), Inches(0.28), width=Inches(1.95))


def _ppt_title(slide, title, subtitle=None, y=0.75):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    box = slide.shapes.add_textbox(Inches(0.7), Inches(y), Inches(11.8), Inches(0.75))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.bold = True
    r.font.size = Pt(30)
    r.font.color.rgb = RGBColor(0, 59, 112)

    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.72), Inches(y + 0.62), Inches(11.4), Inches(0.45))
        stf = sub.text_frame
        stf.clear()
        sp = stf.paragraphs[0]
        sr = sp.add_run()
        sr.text = subtitle
        sr.font.size = Pt(13)
        sr.font.color.rgb = RGBColor(82, 96, 109)


def _ppt_text(slide, text, x, y, w, h, size=15, bold=False, color=(31, 41, 55)):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text or ""
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor(*color)
    return box


def _ppt_add_photo_grid(slide, photos, x, y, w, h):
    from pptx.util import Inches
    import tempfile

    if not photos:
        return

    max_photos = min(4, len(photos))
    if max_photos == 1:
        positions = [(x, y, w, h)]
    elif max_photos == 2:
        positions = [(x, y, w, h / 2 - 0.08), (x, y + h / 2 + 0.08, w, h / 2 - 0.08)]
    else:
        positions = [
            (x, y, w / 2 - 0.08, h / 2 - 0.08),
            (x + w / 2 + 0.08, y, w / 2 - 0.08, h / 2 - 0.08),
            (x, y + h / 2 + 0.08, w / 2 - 0.08, h / 2 - 0.08),
            (x + w / 2 + 0.08, y + h / 2 + 0.08, w / 2 - 0.08, h / 2 - 0.08),
        ]

    for ph, pos in zip(photos[:max_photos], positions):
        suffix = ".png" if ph.get("mime_type") == "image/png" else ".jpg"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(ph["bytes"])
            tmp_path = tmp.name
        try:
            slide.shapes.add_picture(
                tmp_path,
                Inches(pos[0]), Inches(pos[1]),
                width=Inches(pos[2]), height=Inches(pos[3])
            )
        except Exception:
            pass


def _ppt_activity_slide(prs, unit_code, activity, photos, page_num):
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(255, 255, 255)

    _ppt_add_logo(slide, prs)

    title = activity.get("title", "")
    _ppt_text(slide, title, 0.65, 0.9, 7.25, 0.85, size=24, bold=True, color=(0, 59, 112))

    meta = []
    meta.append(unit_code)
    meta.append(rank_label(activity.get("ranking")))
    if activity.get("category"):
        meta.append(activity.get("category"))
    if activity.get("participants"):
        meta.append(f"{activity.get('participants')} participantes / alcance")
    _ppt_text(slide, " · ".join(meta), 0.7, 1.68, 7.1, 0.35, size=11, color=(82, 96, 109))

    body = (
        activity.get("description_edited")
        or activity.get("description_original")
        or activity.get("description")
        or ""
    )
    _ppt_text(slide, body, 0.72, 2.18, 6.15, 3.35, size=14, color=(31, 41, 55))
    if activity.get("social_url"):
        platform = detect_social_platform(activity["social_url"]) or "Redes sociales"
        _ppt_text(
            slide,
            f"{platform}: {activity['social_url']}",
            0.72, 5.62, 6.15, 0.48,
            size=10, color=(0, 59, 112)
        )

    # Visual area
    if photos:
        _ppt_add_photo_grid(slide, photos, 7.25, 1.35, 5.4, 4.85)
    else:
        # Light placeholder
        shape = slide.shapes.add_shape(1, Inches(7.25), Inches(1.35), Inches(5.4), Inches(4.85))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(243, 246, 248)
        shape.line.color.rgb = RGBColor(215, 222, 229)
        _ppt_text(slide, "Sin fotografías cargadas", 8.35, 3.45, 3.3, 0.4, size=13, color=(100, 116, 139))

    _ppt_add_footer(slide, page_num)
    return slide



def _ppt_chart_slide(prs, unit_code, activity, chart, page_num):
    from pptx.util import Inches
    from pptx.dml.color import RGBColor
    import tempfile

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(255, 255, 255)
    _ppt_add_logo(slide, prs)

    _ppt_text(
        slide,
        chart.get("title") or "Gráfica",
        0.7, 0.9, 11.8, 0.65,
        size=25, bold=True, color=(0, 59, 112)
    )
    _ppt_text(
        slide,
        f"{unit_code} · {activity.get('title','')}",
        0.72, 1.55, 11.2, 0.35,
        size=11, color=(82, 96, 109)
    )

    suffix = ".png" if chart.get("mime_type") == "image/png" else ".jpg"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(chart["bytes"])
        tmp_path = tmp.name
    try:
        slide.shapes.add_picture(
            tmp_path,
            Inches(1.25), Inches(2.0),
            width=Inches(10.8), height=Inches(4.55)
        )
    except Exception:
        pass

    _ppt_add_footer(slide, page_num)
    return slide


def generate_saved_center_ppt(rep, activities):
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor
    import io

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(243, 246, 248)
    _ppt_add_logo(slide, prs)
    _ppt_title(
        slide,
        "Informe de actividades",
        f"{rep['unit_code']} · {UNITS.get(rep['unit_code'], '')} · {rep['month']} {rep['year']}",
        y=1.8,
    )
    _ppt_text(
        slide,
        "Dirección de Integración Comunitaria",
        0.75, 3.35, 8.0, 0.45,
        size=17, bold=True, color=(0, 59, 112),
    )
    _ppt_add_footer(slide, 1)

    page = 2
    for act in sorted(activities, key=activity_sort_key):
        photos = get_activity_photos(act["id"])
        _ppt_activity_slide(prs, rep["unit_code"], act, photos, page)
        page += 1
        chart = get_activity_chart(act)
        if chart:
            _ppt_chart_slide(prs, rep["unit_code"], act, chart, page)
            page += 1

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.getvalue()


def generate_consolidated_ppt(month, year, reports, activities_by_report):
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor
    import io

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(243, 246, 248)
    _ppt_add_logo(slide, prs)
    _ppt_title(slide, "Informe consolidado", f"Dirección de Integración Comunitaria · {month} {year}", y=1.8)
    selected_count = sum(
        1
        for rep in reports
        for a in activities_by_report.get(rep["id"], [])
        if selected_for_final(a["id"])
    )
    _ppt_text(slide, f"{selected_count} actividad(es) seleccionada(s) para versión final", 0.75, 3.35, 8.5, 0.45, size=17, bold=True, color=(0, 59, 112))
    _ppt_add_footer(slide, 1)

    page = 2
    # By center
    for rep in reports:
        acts = [
            a for a in sorted(activities_by_report.get(rep["id"], []), key=activity_sort_key)
            if selected_for_final(a["id"])
        ]
        if not acts:
            continue

        section = prs.slides.add_slide(prs.slide_layouts[6])
        section.background.fill.solid()
        section.background.fill.fore_color.rgb = RGBColor(243, 246, 248)
        _ppt_add_logo(section, prs)
        _ppt_text(section, rep["unit_code"], 0.75, 2.15, 5.0, 0.9, size=42, bold=True, color=(0, 59, 112))
        _ppt_text(section, UNITS.get(rep["unit_code"], ""), 0.78, 3.05, 10.8, 0.6, size=19, color=(82, 96, 109))
        _ppt_text(section, f"{len(acts)} actividad(es) seleccionada(s)", 0.8, 3.85, 6.0, 0.35, size=14, color=(100, 116, 139))
        _ppt_add_footer(section, page)
        page += 1

        for act in acts:
            photos = get_activity_photos(act["id"])
            _ppt_activity_slide(prs, rep["unit_code"], act, photos, page)
            page += 1
            chart = get_activity_chart(act)
            if chart:
                _ppt_chart_slide(prs, rep["unit_code"], act, chart, page)
                page += 1

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.getvalue()


def handle_center_change():
    """Prevent temporary data from one center appearing in another center."""
    st.session_state.validated_center_email = ""
    clear_center_capture_state(reset_period=False)
    st.session_state.director_page = "Nuevo reporte"



# ---------- DIRECTOR SEARCH / CENTER STATISTICS HELPERS ----------
def generate_selected_extract_word(selected_hits, unit_code):
    """Create one Word file with selected historical activity extracts from one center."""
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)
    add_docx_page_x_of_y(sec)

    p = doc.add_paragraph()
    r = p.add_run("Extractos seleccionados")
    r.bold = True
    r.font.size = Pt(19)
    r.font.color.rgb = RGBColor(0, 59, 112)
    doc.add_paragraph(f"{unit_code} · {UNITS.get(unit_code, '')}")
    doc.add_paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    for idx, (rep, act) in enumerate(selected_hits, start=1):
        if idx > 1:
            doc.add_page_break()
        doc.add_heading(act.get("title") or "Actividad", level=1)
        meta = f"{rep.get('month')} {rep.get('year')} · {act.get('category','')} · {rank_label(act.get('ranking'))}"
        doc.add_paragraph(meta)
        if act.get("participants"):
            doc.add_paragraph(f"Participantes / alcance: {act.get('participants')}")
        doc.add_paragraph(act.get("description_edited") or act.get("description_original") or "")
        if act.get("social_url"):
            doc.add_paragraph(f"Red social: {act.get('social_url')}")
        chart = get_activity_chart(act)
        if chart and chart.get("bytes"):
            doc.add_paragraph(chart.get("title") or "Gráfica")
            try:
                doc.add_picture(io.BytesIO(chart["bytes"]), width=Inches(6.2))
            except Exception:
                pass
        photos = get_activity_photos(act.get("id"))
        for ph in photos[:4]:
            try:
                doc.add_picture(io.BytesIO(ph["bytes"]), width=Inches(5.6))
            except Exception:
                pass

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def generate_selected_extract_pdf(selected_hits, unit_code):
    bio = io.BytesIO()
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DirExtractTitle", parent=styles["Title"], textColor=HexColor(ITESO_BLUE),
        fontSize=18, leading=21, spaceAfter=8,
    )
    h_style = ParagraphStyle(
        "DirExtractH", parent=styles["Heading2"], textColor=HexColor(ITESO_BLUE),
        fontSize=13, leading=16, spaceAfter=5,
    )
    story = [
        Paragraph("Extractos seleccionados", title_style),
        Paragraph(f"{html.escape(unit_code)} · {html.escape(UNITS.get(unit_code, ''))}", styles["BodyText"]),
        Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["BodyText"]),
        Spacer(1, 10),
    ]
    for idx, (rep, act) in enumerate(selected_hits, start=1):
        if idx > 1:
            story.append(PageBreak())
        story.append(Paragraph(html.escape(act.get("title") or "Actividad"), h_style))
        meta = f"{rep.get('month')} {rep.get('year')} · {act.get('category','')} · {rank_label(act.get('ranking'))}"
        story.append(Paragraph(html.escape(meta), styles["BodyText"]))
        if act.get("participants"):
            story.append(Paragraph(f"<b>Participantes / alcance:</b> {int(act.get('participants') or 0)}", styles["BodyText"]))
        desc = act.get("description_edited") or act.get("description_original") or ""
        story.append(Spacer(1, 5))
        story.append(Paragraph(html.escape(desc).replace("\n", "<br/>"), styles["BodyText"]))
        if act.get("social_url"):
            story.append(Paragraph(f"<b>Red social:</b> {html.escape(act.get('social_url'))}", styles["BodyText"]))
        chart = get_activity_chart(act)
        if chart and chart.get("bytes"):
            try:
                story.append(Spacer(1, 6))
                story.append(Paragraph(html.escape(chart.get("title") or "Gráfica"), styles["BodyText"]))
                story.append(Image(io.BytesIO(chart["bytes"]), width=430, height=240))
            except Exception:
                pass
        photos = get_activity_photos(act.get("id"))
        for ph in photos[:3]:
            try:
                story.append(Spacer(1, 6))
                story.append(Image(io.BytesIO(ph["bytes"]), width=360, height=220))
            except Exception:
                pass
    pdf = SimpleDocTemplate(
        bio, pagesize=letter, rightMargin=38, leftMargin=38, topMargin=42, bottomMargin=34,
        title=f"Extractos {unit_code}",
    )
    pdf.build(story, canvasmaker=NumberedCanvas)
    return bio.getvalue()


def generate_selected_extract_ppt(selected_hits, unit_code):
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _ppt_add_logo(slide, prs)
    _ppt_title(slide, "Extractos seleccionados", f"{unit_code} · {UNITS.get(unit_code, '')}", y=1.65)
    _ppt_text(slide, f"{len(selected_hits)} actividad(es) seleccionada(s)", 0.78, 3.25, 6.4, 0.45, size=17, bold=True, color=(0, 59, 112))
    _ppt_add_footer(slide, 1)
    page_num = 2
    for rep, act in selected_hits:
        photos = get_activity_photos(act.get("id"))
        _ppt_activity_slide(prs, unit_code, act, photos, page_num)
        page_num += 1
        chart = get_activity_chart(act)
        if chart:
            _ppt_chart_slide(prs, unit_code, act, chart, page_num)
            page_num += 1
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()


def center_statistics_data(unit_code):
    reports = [r for r in get_reports() if r.get("unit_code") == unit_code]
    report_ids = {r.get("id") for r in reports}
    activities = [a for a in get_activities() if a.get("report_id") in report_ids]
    return reports, activities


# ---------- HEADER ----------
header_left, header_right = st.columns([1.35, 4.65], vertical_alignment="center")
with header_left:
    st.image("assets/iteso_logo.png", width=255)

with header_right:
    st.markdown(
        """
        <div class="iteso-title-wrap">
            <div class="iteso-title">Sistema de informes mensuales</div>
            <div class="iteso-subtitle">Dirección de Integración Comunitaria · ITESO</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="iteso-divider"></div>', unsafe_allow_html=True)

# Access is handled with email + password or an administrator-issued activation code.


# ---------- ADMIN STATISTICS / BACKUP HELPERS ----------
def _stats_chart_png(df, label_col, value_col, title, kind="bar"):
    """Build a compact PNG chart for Word/PDF statistical exports."""
    if df is None or df.empty:
        return None
    fig, ax = plt.subplots(figsize=(8.5, 4.5))
    labels = [str(v) for v in df[label_col].tolist()]
    values = [float(v or 0) for v in df[value_col].tolist()]
    if kind == "line":
        ax.plot(labels, values, marker="o", linewidth=2)
        ax.tick_params(axis="x", rotation=45)
    elif len(labels) >= 6:
        ax.barh(labels, values)
        ax.invert_yaxis()
    else:
        ax.bar(labels, values)
        ax.tick_params(axis="x", rotation=25)
    ax.set_title(title)
    ax.set_ylabel(value_col if kind != "barh" else "")
    ax.grid(axis="y" if kind != "barh" else "x", alpha=0.2)
    fig.tight_layout()
    bio = io.BytesIO()
    fig.savefig(bio, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    bio.seek(0)
    return bio.getvalue()


def _docx_add_df_table(doc, df, max_rows=80):
    if df is None or df.empty:
        doc.add_paragraph("Sin datos para los filtros seleccionados.")
        return
    shown = df.head(max_rows)
    table = doc.add_table(rows=1, cols=len(shown.columns))
    table.style = "Table Grid"
    for idx, col in enumerate(shown.columns):
        table.rows[0].cells[idx].text = str(col)
    for _, row in shown.iterrows():
        cells = table.add_row().cells
        for idx, col in enumerate(shown.columns):
            value = row[col]
            if pd.isna(value):
                value = ""
            cells[idx].text = str(value)
    if len(df) > max_rows:
        doc.add_paragraph(f"Se muestran las primeras {max_rows} filas de {len(df)}.")


def _pdf_table_from_df(df, max_rows=60):
    if df is None or df.empty:
        return Paragraph("Sin datos para los filtros seleccionados.", getSampleStyleSheet()["BodyText"])
    shown = df.head(max_rows).copy()
    data = [[str(c) for c in shown.columns]]
    for _, row in shown.iterrows():
        values = []
        for col in shown.columns:
            v = row[col]
            values.append("" if pd.isna(v) else str(v))
        data.append(values)
    page_width = letter[0] - 72
    col_width = page_width / max(1, len(shown.columns))
    table = Table(data, colWidths=[col_width] * len(shown.columns), repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), HexColor(ITESO_BLUE)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("GRID", (0,0), (-1,-1), 0.35, HexColor("#D7DEE5")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, HexColor("#F8FAFC")]),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    return table


def generate_statistics_word(selected, filter_text, summary, datasets):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)
    add_docx_page_x_of_y(sec)

    p = doc.add_paragraph()
    r = p.add_run("Estadísticas de informes mensuales")
    r.bold = True
    r.font.size = Pt(19)
    r.font.color.rgb = RGBColor(0, 59, 112)
    doc.add_paragraph("Dirección de Integración Comunitaria · ITESO")
    doc.add_paragraph(filter_text)
    doc.add_paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    if selected.get("summary"):
        doc.add_heading("Resumen general", level=1)
        for label, value in summary.items():
            doc.add_paragraph(f"{label}: {value}")
        doc.add_paragraph(
            "* Personas impactadas corresponde a la suma de Participantes / alcance y puede incluir personas repetidas entre actividades."
        )

    if selected.get("reports"):
        doc.add_heading("Informes acumulados por centro", level=1)
        df = datasets.get("reports")
        _docx_add_df_table(doc, df)
        img = _stats_chart_png(df, "Centro", "Informes", "Informes acumulados por centro")
        if img:
            doc.add_picture(io.BytesIO(img), width=Inches(6.6))

    if selected.get("activities"):
        doc.add_heading("Actividades por reporte y mes", level=1)
        _docx_add_df_table(doc, datasets.get("activities"))

    if selected.get("categories"):
        doc.add_heading("Actividades por categoría", level=1)
        df = datasets.get("categories")
        img = _stats_chart_png(df, "Categoría", "Actividades", "Actividades por categoría")
        if img:
            doc.add_picture(io.BytesIO(img), width=Inches(6.6))
        _docx_add_df_table(doc, df)

    if selected.get("people"):
        doc.add_heading("Personas impactadas por mes", level=1)
        df = datasets.get("people")
        img = _stats_chart_png(df, "Periodo", "Personas impactadas*", "Personas impactadas por mes", kind="line")
        if img:
            doc.add_picture(io.BytesIO(img), width=Inches(6.6))
        _docx_add_df_table(doc, df)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def generate_statistics_pdf(selected, filter_text, summary, datasets):
    bio = io.BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="StatsTitle", parent=styles["Title"], textColor=HexColor(ITESO_BLUE),
        fontSize=18, leading=21, spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name="StatsH2", parent=styles["Heading2"], textColor=HexColor(ITESO_BLUE),
        fontSize=13, leading=16, spaceBefore=10, spaceAfter=6,
    ))
    story = [
        Paragraph("Estadísticas de informes mensuales", styles["StatsTitle"]),
        Paragraph("Dirección de Integración Comunitaria · ITESO", styles["BodyText"]),
        Paragraph(filter_text, styles["BodyText"]),
        Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["BodyText"]),
        Spacer(1, 10),
    ]

    def add_chart(img_bytes):
        if img_bytes:
            story.append(Image(io.BytesIO(img_bytes), width=468, height=248))
            story.append(Spacer(1, 7))

    if selected.get("summary"):
        story.append(Paragraph("Resumen general", styles["StatsH2"]))
        for label, value in summary.items():
            story.append(Paragraph(f"<b>{html.escape(str(label))}:</b> {html.escape(str(value))}", styles["BodyText"]))
        story.append(Paragraph(
            "* Personas impactadas corresponde a la suma de Participantes / alcance y puede incluir personas repetidas entre actividades.",
            styles["BodyText"],
        ))

    if selected.get("reports"):
        story.append(Paragraph("Informes acumulados por centro", styles["StatsH2"]))
        df = datasets.get("reports")
        story.append(_pdf_table_from_df(df))
        story.append(Spacer(1, 6))
        add_chart(_stats_chart_png(df, "Centro", "Informes", "Informes acumulados por centro"))

    if selected.get("activities"):
        story.append(Paragraph("Actividades por reporte y mes", styles["StatsH2"]))
        story.append(_pdf_table_from_df(datasets.get("activities")))

    if selected.get("categories"):
        story.append(Paragraph("Actividades por categoría", styles["StatsH2"]))
        df = datasets.get("categories")
        add_chart(_stats_chart_png(df, "Categoría", "Actividades", "Actividades por categoría"))
        story.append(_pdf_table_from_df(df))

    if selected.get("people"):
        story.append(Paragraph("Personas impactadas por mes", styles["StatsH2"]))
        df = datasets.get("people")
        add_chart(_stats_chart_png(df, "Periodo", "Personas impactadas*", "Personas impactadas por mes", kind="line"))
        story.append(_pdf_table_from_df(df))

    pdf = SimpleDocTemplate(
        bio, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=42, bottomMargin=34,
        title="Estadísticas DIC ITESO",
    )
    pdf.build(story, canvasmaker=NumberedCanvas)
    return bio.getvalue()


def fetch_all_table_rows(table_name, page_size=1000):
    """Retrieve all rows from a Supabase table in pages for administrator backups."""
    if not supabase:
        return []
    rows = []
    start = 0
    while True:
        chunk = (
            supabase.table(table_name)
            .select("*")
            .range(start, start + page_size - 1)
            .execute()
            .data or []
        )
        rows.extend(chunk)
        if len(chunk) < page_size:
            break
        start += page_size
    return rows


def generate_database_backup_zip():
    """Create a portable ZIP backup of the application database tables (not Storage binaries)."""
    tables = [
        "units", "reports", "activities", "activity_photos", "authorized_users", "audit_log",
        "report_strategic_advances", "report_strategic_issues", "report_strategic_learnings",
    ]
    bio = io.BytesIO()
    generated_at = datetime.utcnow().isoformat() + "Z"
    metadata = {
        "generated_at_utc": generated_at,
        "application": "DIC ITESO - Informes mensuales",
        "scope": "Postgres application tables",
        "includes_storage_files": False,
        "tables": {},
    }
    with zipfile.ZipFile(bio, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for table in tables:
            try:
                rows = fetch_all_table_rows(table)
                metadata["tables"][table] = len(rows)
                df = pd.DataFrame(rows)
                zf.writestr(f"{table}.csv", df.to_csv(index=False).encode("utf-8-sig"))
                zf.writestr(
                    f"{table}.json",
                    json.dumps(rows, ensure_ascii=False, indent=2, default=str).encode("utf-8"),
                )
            except Exception as exc:
                metadata["tables"][table] = f"ERROR: {exc}"
        readme = (
            "RESPALDO DIC ITESO\n\n"
            "Este archivo contiene una exportación de las tablas de la base de datos de la aplicación "
            "en formatos CSV y JSON.\n\n"
            "No incluye los archivos binarios almacenados en Supabase Storage (fotografías y gráficas); "
            "sí incluye las rutas de Storage registradas en las tablas.\n"
            "Las contraseñas de usuarios no forman parte de estas tablas y no se incluyen.\n"
        )
        zf.writestr("README.txt", readme.encode("utf-8"))
        zf.writestr("metadata.json", json.dumps(metadata, ensure_ascii=False, indent=2).encode("utf-8"))
    bio.seek(0)
    return bio.getvalue(), metadata


# ---------- ACCESS ----------
with st.sidebar:
    st.markdown("### Acceso")
    profile = st.radio("Perfil", ["Centro / Dirección", "Administración DIC"], label_visibility="collapsed")
    st.caption(f"Base de datos: **{db_mode()}**")
    st.divider()

if st.session_state.get("session_timeout_message"):
    st.warning(st.session_state.pop("session_timeout_message"))

if enforce_session_timeout():
    st.rerun()

@st.fragment(run_every="60s")
def session_timeout_watchdog():
    """Check idle authenticated sessions once per minute without refreshing activity."""
    authenticated = bool(
        st.session_state.get("admin_authenticated")
        or st.session_state.get("center_authenticated")
    )
    last = st.session_state.get("last_activity_at")
    if authenticated and last is not None and time.time() - float(last) >= SESSION_TIMEOUT_SECONDS:
        if st.session_state.get("center_authenticated"):
            auth_logout_center()
        st.session_state.admin_authenticated = False
        clear_session_activity()
        st.session_state.session_timeout_message = (
            "La sesión se cerró automáticamente después de 30 minutos sin actividad."
        )
        st.rerun()

session_timeout_watchdog()

if profile == "Centro / Dirección":
    with st.sidebar:
        if st.session_state.center_authenticated:
            unit = st.session_state.center_user_unit
            sender_email = st.session_state.center_user_email
            user_role = st.session_state.center_user_role
            user_name = st.session_state.center_user_name
            st.success(f"{user_name or sender_email}")
            st.caption(f"{unit} · {UNITS.get(unit, '')}")
            st.caption(f"Rol: **{user_role.title()}**")
            if st.button("Cerrar sesión", use_container_width=True):
                auth_logout_center()
                st.rerun()
            email_validated = True
        else:
            unit = ""
            sender_email = st.text_input("Correo institucional", placeholder="nombre@iteso.mx", key="login_email")
            password = st.text_input("Contraseña", type="password", key="login_password")
            if st.button("Iniciar sesión", type="primary", use_container_width=True):
                ok, msg, user = auth_sign_in(sender_email, password)
                if ok and user:
                    clear_center_capture_state(reset_period=False)
                    st.session_state.center_authenticated = True
                    st.session_state.center_user_email = user.get("email", sender_email).lower()
                    st.session_state.center_user_name = user.get("name", "")
                    st.session_state.center_user_unit = user.get("unit_code", "")
                    st.session_state.center_user_role = user.get("role", "")
                    st.session_state.validated_center_email = user.get("email", sender_email).lower()
                    st.session_state.director_page = "Nuevo reporte"
                    mark_session_activity()
                    st.rerun()
                else:
                    st.error(msg)

            with st.expander("Activar cuenta / Restablecer contraseña"):
                st.caption(
                    "Usa el código temporal que te entregue la Administración DIC. "
                    "El mismo proceso sirve para activar tu cuenta por primera vez o crear una nueva contraseña."
                )
                activation_email = st.text_input("Correo institucional", key="activation_email")
                activation_code = st.text_input("Código temporal", key="activation_code", placeholder="ABCDE-23456")
                activation_password = st.text_input("Nueva contraseña", type="password", key="activation_password")
                activation_password_2 = st.text_input("Confirmar contraseña", type="password", key="activation_password_2")
                st.caption("La contraseña debe tener al menos 10 caracteres.")
                if st.button("Activar / guardar nueva contraseña", use_container_width=True):
                    if activation_password != activation_password_2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        ok, msg, activated_user = activate_or_reset_account(
                            activation_email, activation_code, activation_password
                        )
                        if ok:
                            st.success("Contraseña guardada. Ya puedes iniciar sesión con tu correo y contraseña.")
                        else:
                            st.error(msg)
            email_validated = False
            user_role = ""
            user_name = ""

        st.divider()
        director_menu_options = ["Nuevo reporte", "Mis reportes"]
        if user_role == "DIRECTOR":
            director_menu_options += ["Buscador", "Estadísticas del centro"]
        page = st.radio(
            "Menú",
            director_menu_options,
            disabled=not email_validated,
            key="director_page",
        )

    if st.session_state.get("submission_receipt"):
        submission_success_dialog()
        st.stop()

    if email_validated:
        today_for_capsule = get_guadalajara_today()
        expected_capsule_key = f"{sender_email.strip().lower()}|{today_for_capsule.isoformat()}"
        if st.session_state.daily_capsule_seen_key != expected_capsule_key:
            daily_learning_capsule(sender_email.strip())
            st.stop()

    if not email_validated:
        st.header("Acceso a informes mensuales")
        st.info(
            "Ingresa con el correo institucional autorizado por Administración y tu contraseña. Si es tu primera vez, activa tu cuenta con el código temporal que te entregaron."
        )
        st.stop()

    if page == "Nuevo reporte":
        st.header("Nuevo reporte mensual")
        if user_role == "DIRECTOR":
            st.success("Perfil Director · puedes asignar Top 1/2/3 y enviar el informe final.")
        else:
            st.info("Perfil Colaborador · puedes capturar, editar y guardar borradores. El ranking y el envío corresponden al Director.")

        st.info(
            "La captura se organiza ahora por los siete rubros del Informe Mensual DIC. "
            "Para cada acción elige primero el rubro; el formulario se adapta automáticamente."
        )

        st.subheader("¿Cómo quieres cargar el reporte?")
        capture_method = st.radio(
            "Método de captura",
            ["Captura en línea", "Importar Word rellenable"],
            horizontal=True,
            key="capture_method",
            label_visibility="collapsed",
        )

        if capture_method == "Importar Word rellenable":
            st.markdown(
                "Carga la **plantilla Word rellenable oficial**. La app leerá los controles, "
                "las casillas, los bloques estratégicos y las imágenes pegadas en las fichas de acciones. "
                "Nada se guarda automáticamente: primero se poblará el formulario para que lo revises."
            )
            st.download_button(
                "⬇️ Descargar plantilla Word rellenable",
                data=official_fillable_docx_template_bytes(),
                file_name=FILLABLE_DOCX_TEMPLATE_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                on_click="ignore",
            )
            uploaded_fillable = st.file_uploader(
                "Subir Word ya llenado (.docx)",
                type=["docx"],
                key="fillable_docx_upload",
                help="Usa la plantilla descargada desde esta misma pantalla y conserva sus controles y estructura.",
            )
            if uploaded_fillable is not None:
                if st.button("🔎 Analizar Word", use_container_width=True):
                    parsed, parse_errors, parse_warnings = parse_fillable_dic_docx(
                        uploaded_fillable.getvalue(), unit
                    )
                    st.session_state.fillable_docx_parsed = parsed
                    st.session_state.fillable_docx_errors = parse_errors
                    st.session_state.fillable_docx_warnings = parse_warnings
                    st.session_state.docx_import_filename = uploaded_fillable.name

                parse_errors = st.session_state.get("fillable_docx_errors", []) or []
                parse_warnings = st.session_state.get("fillable_docx_warnings", []) or []
                parsed = st.session_state.get("fillable_docx_parsed")
                for err in parse_errors:
                    st.error(err)
                for warn in parse_warnings:
                    st.warning(warn)
                if parsed and not parse_errors:
                    st.success(
                        f"Word leído: **{len(parsed.get('activities') or [])} acciones**, "
                        f"**{len(parsed.get('advances') or [])} avances**, "
                        f"**{len(parsed.get('issues') or [])} situaciones** y "
                        f"**{len(parsed.get('learnings') or [])} aprendizajes**."
                    )
                    cimp1, cimp2 = st.columns([1, 1])
                    with cimp1:
                        st.caption(
                            f"Periodo detectado: {parsed.get('month')} {parsed.get('year')} · "
                            f"Centro: {parsed.get('center') or unit}"
                        )
                    with cimp2:
                        if st.button("📥 Importar al formulario", type="primary", use_container_width=True):
                            hydrate_fillable_docx(parsed)
                            st.rerun()

            if st.session_state.get("docx_import_success"):
                st.success(st.session_state.docx_import_success)
                for warn in st.session_state.get("docx_import_warnings", []) or []:
                    st.warning(warn)

        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.info(f"**{unit}** · {UNITS[unit]}")
        with col2:
            month = st.selectbox("Mes", MONTHS, key="capture_month")
        with col3:
            year = st.selectbox("Año", list(range(2025, 2031)), key="capture_year")

        if st.session_state.get("resuming_report_id"):
            st.success(
                f"Continuando borrador de **{month} {year}**. "
                "Puedes editar las acciones existentes o agregar nuevas."
            )

        periodic_notes = []
        if month in ACADEMIC_REPORTING_MONTHS:
            periodic_notes.append("Oferta académica y docencia está habilitada para este corte académico.")
        if month in RESEARCH_REPORTING_MONTHS:
            periodic_notes.append("Investigación está habilitada para este mes (enero/agosto).")
        if periodic_notes:
            st.caption(" · ".join(periodic_notes))
        st.caption(
            "Registra sólo las acciones relevantes del mes. Los rubros 1–5 están disponibles cada mes. "
            "Oferta académica y docencia se solicita tres veces al año; Investigación, dos veces al año. "
            "Al terminar, el Director seleccionará de 1 a 3 acciones como **Lo más relevante del mes** y completará **Aprendizajes**."
        )

        for i in range(st.session_state.num_activities):
            previous_complete = True if i == 0 else activity_fields_complete(i - 1)
            with st.expander(f"Acción {i+1}", expanded=(i < 2 and previous_complete)):
                if not previous_complete:
                    st.info(f"🔒 Completa los campos obligatorios de la Acción {i} antes de capturar la Acción {i+1}.")
                    continue

                current_rubric = st.session_state.get(f"rubro_{i}") or CATEGORIES[0]
                rubric_options = available_rubrics_for_month(month, include_current=current_rubric)
                if current_rubric not in rubric_options:
                    st.session_state[f"rubro_{i}"] = rubric_options[0]
                rubro = st.selectbox(
                    "Rubro de la acción",
                    rubric_options,
                    key=f"rubro_{i}",
                    help="Selecciona el apartado del Informe Mensual DIC al que corresponde esta acción.",
                )
                st.caption(REPORT_RUBRICS.get(rubro, ""))

                if rubro in {
                    "Vida universitaria", "Vinculación externa", "Desarrollo institucional",
                    "Capacitación y formación del personal"
                }:
                    st.text_input("Título de la acción", key=f"title_{i}")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.selectbox("Fin de la acción", [""] + ACTION_PURPOSES, key=f"purpose_{i}")
                    with c2:
                        st.selectbox("Tipo de acción", [""] + ACTION_TYPES, key=f"action_type_{i}")
                    st.selectbox(
                        "Planeación",
                        ["", "Planeada", "Emergente"],
                        key=f"planning_{i}",
                        help="Planeada: forma parte de la planeación aprobada. Emergente: responde a una situación, oportunidad o petición no prevista.",
                    )
                    criteria = st.multiselect(
                        "Criterio(s) de inclusión que cumple",
                        INCLUSION_CRITERIA,
                        key=f"criteria_{i}",
                    )
                    if "Otro" in criteria:
                        st.text_input("Especifica el otro criterio", key=f"criteria_other_{i}")
                    c3, c4, c5 = st.columns([1, 2, 1])
                    with c3:
                        normalize_date_widget_state(f"activity_date_{i}", fallback=datetime.now().date())
                        normalize_date_widget_state(f"activity_date_{i}", fallback=datetime.now().date())
                    st.date_input("Fecha", key=f"activity_date_{i}")
                    with c4:
                        st.text_input("Lugar", key=f"location_{i}")
                    with c5:
                        st.number_input("Número de participantes", min_value=0, step=1, key=f"part_{i}")
                    population = st.multiselect(
                        "Población a la que está dirigida",
                        TARGET_POPULATIONS,
                        key=f"population_{i}",
                    )
                    if "Comunidad externa" in population:
                        st.text_input(
                            "Nombre de la comunidad / institución externa",
                            key=f"external_population_{i}",
                        )
                    collab = st.radio(
                        "¿Colaboran otras instancias de la DIC?",
                        ["No", "Sí"], horizontal=True, key=f"dic_collab_{i}",
                    )
                    if collab == "Sí":
                        st.text_input("¿Cuáles instancias de la DIC?", key=f"dic_units_{i}")
                    desc = st.text_area(
                        "Descripción breve",
                        height=130,
                        key=f"desc_{i}",
                        placeholder="Qué ocurrió, principales resultados y actores participantes.",
                    )
                    wc = word_count(desc)
                    st.caption(f"{wc}/250 palabras")
                    if wc > 250:
                        st.error(f"Reduce la descripción en {wc-250} palabras.")
                    st.text_input(
                        "Por qué importa / qué conviene que la Dirección sepa (1 línea)",
                        key=f"relevance_note_{i}",
                    )
                    st.text_input(
                        "Enlace relacionado (opcional)",
                        key=f"social_{i}",
                        placeholder="https://...",
                    )

                elif rubro == "Participación en medios de difusión":
                    st.info("Registra una ficha por participación en medios o redes sociales.")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.selectbox("Tipo de participación", [""] + MEDIA_TYPES, key=f"media_type_{i}")
                    with c2:
                        st.text_input("Medio o plataforma", key=f"media_platform_{i}")
                    st.date_input("Fecha", key=f"activity_date_{i}")
                    st.text_input("Tema de la participación", key=f"media_topic_{i}")
                    st.text_input("Enlace (en caso de que haya)", key=f"media_link_{i}", placeholder="https://...")

                elif rubro == "Oferta académica y docencia":
                    st.info(
                        "Este rubro se solicita sólo en los tres cortes académicos del año. "
                        "Registra una ficha por curso ofertado."
                    )
                    st.selectbox("Semestre", [""] + ACADEMIC_SEMESTERS, key=f"academic_semester_{i}")
                    st.text_input("Nombre del curso ofertado", key=f"title_{i}")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.number_input("Créditos", min_value=0, step=1, key=f"academic_credits_{i}")
                    with c2:
                        st.selectbox("¿Se abrió?", ["", "Sí", "No"], key=f"academic_opened_{i}")
                    with c3:
                        st.number_input("Número de grupos", min_value=0, step=1, key=f"academic_groups_{i}")
                    c4, c5 = st.columns(2)
                    with c4:
                        st.number_input("# Profesores de Tiempo Fijo", min_value=0, step=1, key=f"academic_fixed_prof_{i}")
                    with c5:
                        st.number_input("# Profesores de Tiempo Variable", min_value=0, step=1, key=f"academic_variable_prof_{i}")
                    with st.expander("Integrante DIC con materias / proceso de promoción (opcional)"):
                        st.text_input("Nombre del docente", key=f"faculty_name_{i}")
                        c6, c7 = st.columns(2)
                        with c6:
                            st.selectbox("Tipo de contrato", ["Tiempo Fijo", "Tiempo Variable"], key=f"faculty_contract_{i}")
                            st.text_input("Dependencia a la que está adscrita la materia", key=f"faculty_unit_{i}")
                        with c7:
                            st.selectbox("Categoría", ["Adjunto", "Asociado", "Titular"], key=f"faculty_category_{i}")
                            st.selectbox("¿Está en proceso de promoción?", ["No", "Sí"], key=f"faculty_promotion_{i}")

                elif rubro == "Investigación":
                    st.info("Investigación se solicita dos veces al año: enero y agosto, en articulación con la DIP.")
                    st.text_input("Proyecto", key=f"title_{i}")
                    st.selectbox("Nivel de participación", [""] + RESEARCH_PARTICIPATION, key=f"research_role_{i}")
                    st.text_area("Integrantes en el proyecto", key=f"research_members_{i}", height=80)
                    st.text_input("Temática o campo de conocimiento", key=f"research_field_{i}")
                    st.text_area("Actores internos o externos con quienes se vincula", key=f"research_actors_{i}", height=80)
                    c1, c2 = st.columns(2)
                    with c1:
                        normalize_date_widget_state(f"research_start_{i}", fallback=datetime.now().date())
                        st.date_input("Fecha de inicio", key=f"research_start_{i}")
                    with c2:
                        normalize_date_widget_state(f"research_end_{i}", fallback=datetime.now().date())
                        st.date_input("Fecha de término", key=f"research_end_{i}")
                    st.selectbox("Nivel de avance", [""] + RESEARCH_PROGRESS, key=f"research_progress_{i}")
                    st.text_area("Productos o actividades de difusión / investigación", key=f"research_products_{i}", height=110)
                    st.number_input("% de avance de productos", min_value=0, max_value=100, step=5, key=f"research_pct_{i}")
                    on_plan = st.selectbox("¿Va según lo planeado?", ["", "Sí", "No"], key=f"research_on_plan_{i}")
                    if on_plan == "No":
                        st.text_area("¿Por qué no va según lo planeado?", key=f"research_on_plan_why_{i}", height=80)

                st.markdown("##### Evidencias opcionales")
                photos = st.file_uploader(
                    "Fotografías (opcional)", type=["jpg", "jpeg", "png"],
                    accept_multiple_files=True, key=f"photos_{i}"
                )
                existing_photos = st.session_state.get(f"existing_photos_{i}", []) or []
                if existing_photos:
                    ep_cols = st.columns(min(3, len(existing_photos)))
                    for ep_idx, ep in enumerate(existing_photos):
                        with ep_cols[ep_idx % len(ep_cols)]:
                            st.image(ep["bytes"], use_container_width=True)
                chart = st.file_uploader("Gráfica (opcional)", type=["jpg", "jpeg", "png"], key=f"chart_{i}")
                existing_chart = st.session_state.get(f"existing_chart_{i}")
                if chart or existing_chart:
                    st.text_input("Título de la gráfica", key=f"chart_title_{i}")
                    preview_chart = chart or existing_chart
                    try:
                        st.image(upload_bytes(preview_chart), use_container_width=True)
                    except Exception:
                        pass

                record, missing = activity_from_session(i, actor_role=user_role, month=month)
                if record is not None and missing:
                    st.warning("Completa: " + ", ".join(missing) + ".")

                action_ready = record is not None and not missing
                if st.session_state.get(f"action_saved_ok_{i}"):
                    st.success("✓ Acción guardada. Puedes seguir editándola y volver a guardar los cambios.")

                with st.container(key=f"save_action_bar_{i}"):
                    if st.button(
                        "💾 Guardar acción",
                        key=f"save_action_{i}",
                        type="primary",
                        use_container_width=True,
                        disabled=not action_ready,
                        help=(
                            "Completa los campos obligatorios de esta acción para habilitar el guardado."
                            if not action_ready else
                            "Guarda únicamente esta acción como parte del borrador mensual."
                        ),
                    ):
                        rid = save_report(unit, month, year, "BORRADOR", sender_email)
                        action_errors = save_single_activity(
                            rid, i + 1, record, actor_role=user_role
                        )
                        if action_errors:
                            st.session_state[f"action_saved_ok_{i}"] = False
                            st.error("La acción no se guardó completamente.")
                            for err in action_errors:
                                st.markdown(f"- {err}")
                        else:
                            st.session_state[f"action_saved_ok_{i}"] = True
                            st.session_state.resuming_report_id = rid
                            st.success("✓ Acción guardada en el borrador mensual.")

        activities, errors = validate_current_activities(st.session_state.num_activities)

        cadd, crem = st.columns([1, 4])
        with cadd:
            if st.button("➕ Agregar acción"):
                if errors:
                    st.error("Completa primero todas las acciones iniciadas antes de agregar otra.")
                else:
                    st.session_state.num_activities += 1
                    st.rerun()
        with crem:
            if st.session_state.num_activities > 5 and st.button("➖ Quitar última"):
                st.session_state.num_activities -= 1
                st.rerun()

        st.divider()
        st.subheader("Lo más relevante del mes")
        st.caption(
            "Selecciona de 1 a 3 acciones sustantivas. El orden de selección corresponde a Top 1, Top 2 y Top 3."
        )
        activity_labels = [f"{idx+1}. {a.get('title') or a.get('category')}" for idx, a in enumerate(activities)]
        monthly_highlights = None
        highlight_messages = []
        if user_role == "DIRECTOR":
            existing_sel = st.session_state.get("highlight_selection", []) or []
            st.session_state["highlight_selection"] = [x for x in existing_sel if x in activity_labels]
            selected_highlights = st.multiselect(
                "Acciones más relevantes",
                activity_labels,
                key="highlight_selection",
                max_selections=3,
                placeholder="Selecciona entre 1 y 3 acciones",
            )
            monthly_highlights = []
            for rank_pos, label in enumerate(selected_highlights, start=1):
                act_idx = activity_labels.index(label)
                reason_key = f"highlight_reason_{act_idx}"
                why = st.text_area(
                    f"¿Por qué es relevante? · Top {rank_pos} · {activities[act_idx].get('title','')}",
                    key=reason_key,
                    height=85,
                ).strip()
                activities[act_idx]["ranking"] = rank_pos
                monthly_highlights.append({
                    "label": label,
                    "activity_index": act_idx,
                    "title": activities[act_idx].get("title"),
                    "category": activities[act_idx].get("category"),
                    "why": why,
                    "ranking": rank_pos,
                })
                if not why:
                    highlight_messages.append(f"Explica por qué es relevante el Top {rank_pos}.")
            if activities and not selected_highlights:
                highlight_messages.append("Selecciona al menos una acción como lo más relevante del mes.")
        else:
            st.info("El Director seleccionará y justificará aquí las 1–3 acciones más relevantes antes del envío final.")

        media_actions = [a for a in activities if a.get("category") == "Participación en medios de difusión"]
        media_summary = None
        if media_actions:
            st.divider()
            st.subheader("Datos concentrados de medios del mes")
            st.caption("Se llenan una sola vez por mes, no por cada participación.")
            m1, m2, m3 = st.columns(3)
            with m1:
                media_appearances = st.number_input("Número de apariciones en medios", min_value=0, step=1, key="media_appearances")
            with m2:
                media_total = st.number_input("Número total de participaciones", min_value=0, step=1, key="media_total_participations")
            with m3:
                media_reach = st.number_input("Alcance o audiencia (si se cuenta con el dato)", min_value=0, step=1, key="media_reach")
            media_summary = {"appearances": int(media_appearances), "total_participations": int(media_total), "reach": int(media_reach)}

        st.divider()
        st.subheader("Aprendizajes y seguimiento estratégico")
        st.caption(
            "Registra cada avance, situación y aprendizaje como un renglón independiente. "
            "Así podrán buscarse, filtrarse y analizarse por mes, centro, tema y vínculo con las acciones reportadas."
        )

        activity_link_options = [0] + list(range(1, len(activities) + 1))
        def _activity_link_label(value):
            if not value:
                return "Sin vínculo a una acción específica"
            idx = int(value) - 1
            if 0 <= idx < len(activities):
                return f"Acción {value} · {activities[idx].get('title') or activities[idx].get('category','')}"
            return f"Acción {value}"

        strategic_advances = []
        strategic_issues = []
        strategic_learnings = []
        learning_messages = []

        if user_role == "DIRECTOR":
            st.markdown("#### A. Avances sustantivos")
            st.caption("Un renglón por avance. Registra qué objetivo o tema avanzó, qué cambió y qué evidencia lo muestra.")
            for idx in range(st.session_state.num_strategic_advances):
                with st.container(border=True):
                    a1, a2 = st.columns([2, 1])
                    with a1:
                        topic = st.text_input("Objetivo / tema", key=f"adv_topic_{idx}", placeholder="Ej. Fortalecer vinculación con organizaciones")
                    with a2:
                        level = st.selectbox(
                            "Nivel de avance",
                            ["Inicial", "En proceso", "Avance importante", "Concluido"],
                            key=f"adv_level_{idx}",
                        )
                    progress = st.text_area("Avance observado", key=f"adv_progress_{idx}", height=80, placeholder="¿Qué cambió o se consiguió este mes?")
                    evidence = st.text_input("Resultado / evidencia (opcional)", key=f"adv_evidence_{idx}", placeholder="Ej. convenio, producto, acuerdo, indicador")
                    linked = st.selectbox("Acción relacionada (opcional)", activity_link_options, format_func=_activity_link_label, key=f"adv_link_{idx}")
                    if topic.strip() or progress.strip() or evidence.strip():
                        if not topic.strip() or not progress.strip():
                            learning_messages.append(f"Avance {idx+1}: completa Objetivo / tema y Avance observado.")
                        strategic_advances.append({
                            "objective_topic": topic.strip(),
                            "progress": progress.strip(),
                            "evidence": evidence.strip(),
                            "progress_level": level,
                            "linked_activity_order": int(linked) if linked else None,
                            "linked_activity_title": (activities[int(linked)-1].get("title") if linked and int(linked)-1 < len(activities) else None),
                        })
            ca1, ca2 = st.columns([1, 4])
            with ca1:
                if st.button("➕ Agregar avance", key="add_strategic_advance"):
                    st.session_state.num_strategic_advances += 1
                    st.rerun()
            with ca2:
                if st.session_state.num_strategic_advances > 1 and st.button("➖ Quitar último avance", key="remove_strategic_advance"):
                    st.session_state.num_strategic_advances -= 1
                    st.rerun()

            st.markdown("#### B. Riesgos, bloqueos y decisiones")
            st.caption("Registra sólo situaciones que requieran atención. Si no hubo ninguna, márcalo explícitamente.")
            no_issues = st.checkbox("No hubo acciones paradas, riesgos ni decisiones pendientes este mes", key="no_strategic_issues")
            if not no_issues:
                for idx in range(st.session_state.num_strategic_issues):
                    with st.container(border=True):
                        i1, i2, i3 = st.columns([1.2, 2, 1])
                        with i1:
                            issue_type = st.selectbox("Tipo", ["Acción parada", "Riesgo", "Requiere decisión"], key=f"issue_type_{idx}")
                        with i2:
                            issue_topic = st.text_input("Tema", key=f"issue_topic_{idx}", placeholder="Tema o asunto")
                        with i3:
                            priority = st.selectbox("Prioridad", ["Bajo", "Medio", "Alto"], key=f"issue_priority_{idx}")
                        desc = st.text_area("Descripción breve", key=f"issue_desc_{idx}", height=75)
                        i4, i5, i6 = st.columns(3)
                        with i4:
                            need = st.selectbox("Qué se necesita", ["Seguimiento", "Decisión", "Recurso", "Coordinación", "Información", "Otro"], key=f"issue_need_{idx}")
                        with i5:
                            status = st.selectbox("Estado", ["Abierto", "En seguimiento", "Resuelto"], key=f"issue_status_{idx}")
                        with i6:
                            linked = st.selectbox("Acción relacionada", activity_link_options, format_func=_activity_link_label, key=f"issue_link_{idx}")
                        dependency = st.text_input("Responsable / instancia de quien depende (opcional)", key=f"issue_dependency_{idx}")
                        if issue_topic.strip() or desc.strip() or dependency.strip():
                            if not issue_topic.strip() or not desc.strip():
                                learning_messages.append(f"Situación {idx+1}: completa Tema y Descripción breve.")
                            strategic_issues.append({
                                "issue_type": issue_type,
                                "topic": issue_topic.strip(),
                                "description": desc.strip(),
                                "priority": priority,
                                "need_type": need,
                                "dependency": dependency.strip(),
                                "target_date": None,
                                "issue_status": status,
                                "linked_activity_order": int(linked) if linked else None,
                                "linked_activity_title": (activities[int(linked)-1].get("title") if linked and int(linked)-1 < len(activities) else None),
                            })
                ci1, ci2 = st.columns([1, 4])
                with ci1:
                    if st.button("➕ Agregar situación", key="add_strategic_issue"):
                        st.session_state.num_strategic_issues += 1
                        st.rerun()
                with ci2:
                    if st.session_state.num_strategic_issues > 1 and st.button("➖ Quitar última situación", key="remove_strategic_issue"):
                        st.session_state.num_strategic_issues -= 1
                        st.rerun()

            st.markdown("#### C. Aprendizajes y oportunidades")
            st.caption("Un renglón por hallazgo. Separa lo observado de lo que implica y del siguiente paso.")
            for idx in range(st.session_state.num_strategic_learnings):
                with st.container(border=True):
                    l1, l2 = st.columns([1, 2])
                    with l1:
                        learning_type = st.selectbox("Tipo", ["Aprendizaje", "Oportunidad"], key=f"learn_type_{idx}")
                    with l2:
                        learning_topic = st.text_input("Tema", key=f"learn_topic_{idx}", placeholder="Tema corto para poder buscarlo después")
                    observation = st.text_area("¿Qué observamos?", key=f"learn_observation_{idx}", height=75)
                    implication = st.text_area("¿Qué implica?", key=f"learn_implication_{idx}", height=75)
                    next_step = st.text_input("Próximo paso sugerido (opcional)", key=f"learn_next_{idx}")
                    linked = st.selectbox("Acción relacionada (opcional)", activity_link_options, format_func=_activity_link_label, key=f"learn_link_{idx}")
                    if learning_topic.strip() or observation.strip() or implication.strip() or next_step.strip():
                        if not learning_topic.strip() or not observation.strip() or not implication.strip():
                            learning_messages.append(f"Aprendizaje {idx+1}: completa Tema, ¿Qué observamos? y ¿Qué implica?.")
                        strategic_learnings.append({
                            "learning_type": learning_type,
                            "topic": learning_topic.strip(),
                            "observation": observation.strip(),
                            "implication": implication.strip(),
                            "next_step": next_step.strip(),
                            "linked_activity_order": int(linked) if linked else None,
                            "linked_activity_title": (activities[int(linked)-1].get("title") if linked and int(linked)-1 < len(activities) else None),
                        })
            cl1, cl2 = st.columns([1, 4])
            with cl1:
                if st.button("➕ Agregar aprendizaje", key="add_strategic_learning"):
                    st.session_state.num_strategic_learnings += 1
                    st.rerun()
            with cl2:
                if st.session_state.num_strategic_learnings > 1 and st.button("➖ Quitar último aprendizaje", key="remove_strategic_learning"):
                    st.session_state.num_strategic_learnings -= 1
                    st.rerun()

            st.markdown("#### D. Síntesis del mes · opcional")
            monthly_summary = st.text_area(
                "¿Qué debería recordar la Dirección de este mes?",
                key="monthly_strategic_summary",
                height=90,
                placeholder="Síntesis breve de 3 a 5 líneas.",
            )

            if not strategic_advances:
                learning_messages.append("Registra al menos un avance sustantivo del mes.")
            if not no_issues and not strategic_issues:
                learning_messages.append("Agrega al menos una situación o marca que no hubo riesgos/decisiones pendientes.")
            if not strategic_learnings:
                learning_messages.append("Registra al menos un aprendizaje u oportunidad del mes.")
        else:
            st.info("El Director completará aquí los avances, riesgos/decisiones y aprendizajes antes del envío final.")
            no_issues = bool(st.session_state.get("no_strategic_issues"))
            monthly_summary = st.session_state.get("monthly_strategic_summary", "")

        # Keep legacy report fields populated so historical exports/searches remain readable.
        learning_planning = _legacy_reflection_text(strategic_advances, "advance")
        learning_risks = "Sin situaciones críticas reportadas este mes." if no_issues else _legacy_reflection_text(strategic_issues, "issue")
        learning_opportunity = _legacy_reflection_text(strategic_learnings, "learning")

        report_extras = {
            "monthly_highlights": monthly_highlights,
            "learning_planning_advances": learning_planning,
            "learning_risks": learning_risks,
            "learning_opportunity": learning_opportunity,
            "strategic_advances": strategic_advances,
            "strategic_issues": strategic_issues,
            "strategic_learnings": strategic_learnings,
            "no_strategic_issues": bool(no_issues),
            "monthly_strategic_summary": (monthly_summary or "").strip(),
            "media_monthly_summary": media_summary,
        }
        final_report_messages = errors + (highlight_messages if user_role == "DIRECTOR" else []) + (learning_messages if user_role == "DIRECTOR" else [])

        st.subheader("Vista previa del reporte")
        st.caption(
            "Puedes previsualizar, descargar y revisar el reporte antes de enviarlo. "
            "Las fotografías cargadas se incluirán en estos archivos."
        )

        validated_activities, validation_messages = validate_current_activities(
            st.session_state.num_activities
        )
        if user_role == "DIRECTOR":
            validated_activities = apply_highlights_to_activities(validated_activities, monthly_highlights)
        preview_ready = bool(validated_activities) and not validation_messages

        pv_col, _ = st.columns([1, 3])
        with pv_col:
            if st.button(
                "👁️ Previsualizar reporte" if not st.session_state.show_center_preview else "Ocultar vista previa",
                use_container_width=True,
            ):
                if validation_messages:
                    show_validation_messages(validation_messages)
                elif not validated_activities:
                    st.warning("Captura al menos una actividad.")
                else:
                    st.session_state.show_center_preview = not st.session_state.show_center_preview
                    st.rerun()

        if preview_ready:
            preview_word = generate_center_word(unit, month, year, validated_activities, report_extras)
            preview_pdf = generate_center_pdf(unit, month, year, validated_activities, report_extras)
            if st.session_state.show_center_preview:
                render_center_preview(unit, month, year, validated_activities, report_extras)
        else:
            preview_word = b""
            preview_pdf = b""

        p1,p2 = st.columns(2)
        if preview_ready:
            with p1:
                st.download_button(
                    "⬇️ Descargar borrador en Word",
                    data=preview_word,
                    file_name=f"Reporte_{unit}_{month}_{year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    on_click="ignore",
                )
            with p2:
                st.download_button(
                    "⬇️ Descargar borrador en PDF",
                    data=preview_pdf,
                    file_name=f"Reporte_{unit}_{month}_{year}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    on_click="ignore",
                )
        else:
            with p1:
                if st.button("⬇️ Descargar borrador en Word", use_container_width=True):
                    if validation_messages:
                        show_validation_messages(validation_messages)
                    else:
                        st.warning("Captura al menos una actividad.")
            with p2:
                if st.button("⬇️ Descargar borrador en PDF", use_container_width=True):
                    if validation_messages:
                        show_validation_messages(validation_messages)
                    else:
                        st.warning("Captura al menos una actividad.")

        st.divider()
        b1, b2 = st.columns(2)
        with b1:
            if st.button("💾 Guardar borrador", use_container_width=True):
                validated_activities, validation_messages = validate_current_activities(
                    st.session_state.num_activities
                )
                if user_role == "DIRECTOR":
                    validated_activities = apply_highlights_to_activities(validated_activities, monthly_highlights)
                if validation_messages:
                    show_validation_messages(validation_messages)
                elif not validated_activities:
                    st.warning("Captura al menos una acción.")
                else:
                    rid = save_report(unit, month, year, "BORRADOR", sender_email, report_extras=report_extras)
                    persistence_errors = replace_activities(rid, validated_activities, actor_role=user_role)
                    if persistence_errors:
                        st.error(
                            "El borrador no se guardó completamente porque hubo un problema con "
                            "uno o más archivos. No cierres esta pantalla hasta corregirlo."
                        )
                        for err in persistence_errors:
                            st.markdown(f"- {err}")
                    else:
                        st.success(
                            f"Borrador guardado y archivos verificados para {month} {year}."
                        )
        with b2:
            if user_role != "DIRECTOR":
                st.button(
                    "📨 Enviar reporte · sólo Director",
                    type="primary",
                    use_container_width=True,
                    disabled=True,
                )
            elif st.button("📨 Enviar reporte", type="primary", use_container_width=True):
                validated_activities, validation_messages = validate_current_activities(
                    st.session_state.num_activities
                )
                validated_activities = apply_highlights_to_activities(validated_activities, monthly_highlights)
                if not email_is_iteso(sender_email):
                    st.error("El correo de la sesión no es válido.")
                elif validation_messages:
                    show_validation_messages(validation_messages)
                elif highlight_messages:
                    for msg in highlight_messages: st.error(msg)
                elif learning_messages:
                    for msg in learning_messages: st.error(msg)
                elif not validated_activities:
                    st.warning("Captura al menos una acción.")
                else:
                    confirm_submission_dialog(
                        unit, month, year, sender_email, validated_activities, actor_role=user_role, report_extras=report_extras
                    )

    elif page == "Mis reportes":
        st.header("Mis reportes")
        rows = [r for r in get_reports() if r["unit_code"] == unit]
        if not rows:
            st.info("Todavía no hay reportes guardados para este centro.")
        for r in rows:
            status = "✅ Enviado" if r["status"] == "ENVIADO" else "🟡 Borrador"
            sent_label = ""
            if r["status"] == "ENVIADO":
                sent_at = format_report_datetime(r.get("submitted_at"))
                if sent_at:
                    sent_label = f" · {sent_at}"

            with st.expander(f"{r['month']} {r['year']} · {status}{sent_label}"):
                acts = get_activities(r["id"])

                if r["status"] == "ENVIADO":
                    sent_at = format_report_datetime(r.get("submitted_at"))
                    if sent_at:
                        st.caption(f"Fecha y hora de envío: {sent_at}")
                else:
                    updated_at = format_report_datetime(r.get("updated_at"))
                    if updated_at:
                        st.caption(f"Última actualización: {updated_at}")
                    st.button(
                        "✏️ Continuar con el informe",
                        key=f"resume_{r['id']}",
                        type="primary",
                        use_container_width=True,
                        on_click=resume_draft,
                        args=(r,),
                    )

                if acts:
                    center_word = generate_saved_center_word(r, acts)
                    center_pdf = generate_saved_center_pdf(r, acts)
                    center_ppt = generate_saved_center_ppt(r, acts)

                    d1, d2, d3 = st.columns(3)
                    with d1:
                        st.download_button(
                            "⬇️ Word",
                            data=center_word,
                            file_name=f"Reporte_{r['unit_code']}_{r['month']}_{r['year']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"my_word_{r['id']}",
                            use_container_width=True,
                            on_click="ignore",
                        )
                    with d2:
                        st.download_button(
                            "⬇️ PDF",
                            data=center_pdf,
                            file_name=f"Reporte_{r['unit_code']}_{r['month']}_{r['year']}.pdf",
                            mime="application/pdf",
                            key=f"my_pdf_{r['id']}",
                            use_container_width=True,
                            on_click="ignore",
                        )
                    with d3:
                        st.download_button(
                            "⬇️ PPT",
                            data=center_ppt,
                            file_name=f"Reporte_{r['unit_code']}_{r['month']}_{r['year']}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"my_ppt_{r['id']}",
                            use_container_width=True,
                            on_click="ignore",
                        )

                st.divider()
                for a in acts:
                    st.markdown(f"**{a['title']}** · {rank_label(a.get('ranking'))}")
                    st.write(a.get("description_original", ""))
                    if a.get("social_url"):
                        render_social_preview(a["social_url"], key_suffix=f"saved_{a['id']}")
                    chart = get_activity_chart(a)
                    if chart:
                        st.markdown(f"**{chart.get('title') or 'Gráfica'}**")
                        st.image(chart["bytes"], use_container_width=True)
                    photos = get_activity_photos(a["id"])
                    if photos:
                        st.caption(
                            f"Fotografía{'s' if len(photos) != 1 else ''}: {len(photos)}"
                        )
                        pc = st.columns(min(3, len(photos)))
                        for idx, ph in enumerate(photos):
                            with pc[idx % len(pc)]:
                                st.image(ph["bytes"], use_container_width=True)
                    else:
                        st.caption("Sin fotografías guardadas para esta actividad.")


    elif page == "Buscador":
        if user_role != "DIRECTOR":
            st.error("Esta sección está disponible únicamente para el Director del centro.")
            st.stop()
        st.header("Buscador del centro")
        st.caption(
            f"Busca únicamente dentro del histórico de {unit}. Selecciona los resultados que quieras "
            "previsualizar y descargar juntos en Word, PDF o PowerPoint."
        )
        query = st.text_input(
            "Buscar palabras o temas",
            placeholder="Ej. inclusión, talleres, jóvenes, acompañamiento...",
            key="director_search_query",
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            search_years = sorted({int(r.get("year")) for r in get_reports() if r.get("unit_code") == unit and r.get("year") is not None})
            search_year = st.selectbox("Año", ["Todos"] + search_years, key="director_search_year")
        with c2:
            search_rank = st.selectbox("Ranking", ["Todos", "Top 1", "Top 2", "Top 3", "Sin ranking"], key="director_search_rank")
        with c3:
            search_category = st.selectbox("Categoría", ["Todas"] + CATEGORIES, key="director_search_category")

        if query.strip():
            reports = [r for r in get_reports() if r.get("unit_code") == unit]
            if search_year != "Todos":
                reports = [r for r in reports if int(r.get("year")) == int(search_year)]
            report_map = {r.get("id"): r for r in reports}
            q = query.strip().lower()
            hits = []
            for a in get_activities():
                rep = report_map.get(a.get("report_id"))
                if not rep:
                    continue
                if search_rank != "Todos" and rank_label(a.get("ranking")) != search_rank:
                    continue
                if search_category != "Todas" and a.get("category") != search_category:
                    continue
                haystack = " ".join([
                    a.get("title", ""), a.get("description_original", ""),
                    a.get("description_edited", ""), a.get("category", ""),
                ]).lower()
                if q in haystack:
                    hits.append((rep, a))

            st.write(f"**{len(hits)} resultado(s)**")
            selected_hits = []
            for rep, a in hits:
                with st.container(border=True):
                    sc, body = st.columns([0.08, 0.92])
                    with sc:
                        selected = st.checkbox("Seleccionar", key=f"dir_search_sel_{a['id']}", label_visibility="collapsed")
                    with body:
                        st.markdown(f"### {a.get('title','Actividad')}")
                        st.caption(
                            f"{rep.get('month')} {rep.get('year')} · {a.get('category','')} · "
                            f"{rank_label(a.get('ranking'))} · Participantes/alcance: {int(a.get('participants') or 0)}"
                        )
                        st.write(a.get("description_edited") or a.get("description_original") or "")
                    if selected:
                        selected_hits.append((rep, a))

            st.divider()
            st.subheader("Selección")
            st.write(f"**{len(selected_hits)} actividad(es) seleccionada(s)**")
            preview_selection = st.checkbox("Previsualizar selección", value=False, key="dir_search_preview")
            if preview_selection and selected_hits:
                for rep, a in selected_hits:
                    with st.expander(f"{a.get('title')} · {rep.get('month')} {rep.get('year')}", expanded=False):
                        st.caption(f"{a.get('category','')} · {rank_label(a.get('ranking'))}")
                        st.write(a.get("description_edited") or a.get("description_original") or "")
                        chart = get_activity_chart(a)
                        if chart:
                            st.markdown(f"**{chart.get('title') or 'Gráfica'}**")
                            st.image(chart["bytes"], use_container_width=True)
                        photos = get_activity_photos(a.get("id"))
                        if photos:
                            cols = st.columns(min(3, len(photos)))
                            for idx, ph in enumerate(photos):
                                with cols[idx % len(cols)]:
                                    st.image(ph["bytes"], use_container_width=True)
            if selected_hits:
                word_bytes = generate_selected_extract_word(selected_hits, unit)
                pdf_bytes = generate_selected_extract_pdf(selected_hits, unit)
                ppt_bytes = generate_selected_extract_ppt(selected_hits, unit)
                stamp = datetime.now().strftime("%Y%m%d_%H%M")
                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button("⬇️ Word", data=word_bytes, file_name=f"Extractos_{unit}_{stamp}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, on_click="ignore")
                with d2:
                    st.download_button("⬇️ PDF", data=pdf_bytes, file_name=f"Extractos_{unit}_{stamp}.pdf", mime="application/pdf", use_container_width=True, on_click="ignore")
                with d3:
                    st.download_button("⬇️ PowerPoint", data=ppt_bytes, file_name=f"Extractos_{unit}_{stamp}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True, on_click="ignore")
            else:
                st.info("Marca al menos un resultado para previsualizarlo o descargarlo.")
        else:
            st.info("Escribe una palabra o tema para buscar en los reportes de tu centro.")

    elif page == "Estadísticas del centro":
        if user_role != "DIRECTOR":
            st.error("Esta sección está disponible únicamente para el Director del centro.")
            st.stop()
        st.header("Estadísticas del centro")
        st.caption(f"Indicadores acumulados exclusivamente de {unit} · {UNITS.get(unit, '')}.")
        all_reports, all_activities = center_statistics_data(unit)
        if not all_reports:
            st.info("Todavía no hay informes guardados para construir estadísticas de este centro.")
        else:
            years_available = sorted({int(r.get("year")) for r in all_reports if r.get("year") is not None})
            f1, f2 = st.columns(2)
            with f1:
                stat_year = st.selectbox("Año", ["Todos"] + years_available, key="center_stats_year")
            with f2:
                stat_status = st.selectbox("Estatus", ["Todos", "Enviados", "Borradores"], key="center_stats_status")

            reports = list(all_reports)
            if stat_year != "Todos":
                reports = [r for r in reports if int(r.get("year")) == int(stat_year)]
            if stat_status == "Enviados":
                reports = [r for r in reports if r.get("status") == "ENVIADO"]
            elif stat_status == "Borradores":
                reports = [r for r in reports if r.get("status") == "BORRADOR"]
            report_ids = {r.get("id") for r in reports}
            activities = [a for a in all_activities if a.get("report_id") in report_ids]
            total_people = sum(int(a.get("participants") or 0) for a in activities)

            m1, m2, m3 = st.columns(3)
            m1.metric("Informes", len(reports))
            m2.metric("Actividades", len(activities))
            m3.metric("Personas impactadas*", f"{total_people:,}")

            month_order = {m: i + 1 for i, m in enumerate(MONTHS)}
            acts_by_report = {}
            for a in activities:
                acts_by_report.setdefault(a.get("report_id"), []).append(a)
            report_rows = []
            for r in reports:
                acts = acts_by_report.get(r.get("id"), [])
                report_rows.append({
                    "Año": int(r.get("year")), "Mes": r.get("month"),
                    "Estatus": "Enviado" if r.get("status") == "ENVIADO" else "Borrador",
                    "Actividades": len(acts),
                    "Personas impactadas*": sum(int(a.get("participants") or 0) for a in acts),
                    "_orden": int(r.get("year"))*100 + month_order.get(r.get("month"), 99),
                })
            report_df = pd.DataFrame(report_rows)
            if not report_df.empty:
                report_df = report_df.sort_values("_orden").drop(columns=["_orden"])

            categories = {}
            for a in activities:
                cat = (a.get("category") or "Sin categoría").strip() or "Sin categoría"
                categories[cat] = categories.get(cat, 0) + 1
            category_df = pd.DataFrame([{"Categoría": k, "Actividades": v} for k, v in categories.items()])
            if not category_df.empty:
                category_df = category_df.sort_values(["Actividades", "Categoría"], ascending=[False, True])

            report_lookup = {r.get("id"): r for r in reports}
            monthly = {}
            for a in activities:
                r = report_lookup.get(a.get("report_id"))
                if not r:
                    continue
                key = (int(r.get("year")), month_order.get(r.get("month"),99), r.get("month"))
                monthly[key] = monthly.get(key, 0) + int(a.get("participants") or 0)
            monthly_df = pd.DataFrame([
                {"Periodo": f"{month[:3]} {year}", "Personas impactadas*": people, "_orden": year*100+mn}
                for (year,mn,month), people in monthly.items()
            ])
            if not monthly_df.empty:
                monthly_df = monthly_df.sort_values("_orden").drop(columns=["_orden"])

            st.markdown("### Actividades por reporte y mes")
            st.dataframe(report_df, use_container_width=True, hide_index=True)
            st.markdown("### Actividades por categoría")
            if not category_df.empty:
                st.bar_chart(category_df.set_index("Categoría"), horizontal=True, use_container_width=True)
                st.dataframe(category_df, use_container_width=True, hide_index=True)
            st.markdown("### Personas impactadas por mes")
            if not monthly_df.empty:
                st.line_chart(monthly_df.set_index("Periodo"), use_container_width=True)
                st.dataframe(monthly_df, use_container_width=True, hide_index=True)
            st.caption("* Corresponde a la suma del campo Participantes / alcance. Una misma persona puede aparecer en más de una actividad.")

            # Allow the Director to export the center statistical module as Word/PDF.
            selected = {"summary": True, "reports": False, "activities": True, "categories": True, "people": True}
            filter_text = f"Centro: {unit} · Año: {stat_year} · Estatus: {stat_status}"
            summary = {"Informes": len(reports), "Actividades": len(activities), "Personas impactadas*": f"{total_people:,}", "Centro": unit}
            datasets = {
                "reports": pd.DataFrame([{"Centro": unit, "Nombre": UNITS.get(unit,''), "Informes": len(reports)}]),
                "activities": report_df,
                "categories": category_df,
                "people": monthly_df,
            }
            word_stats = generate_statistics_word(selected, filter_text, summary, datasets)
            pdf_stats = generate_statistics_pdf(selected, filter_text, summary, datasets)
            d1, d2 = st.columns(2)
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            with d1:
                st.download_button("⬇️ Descargar estadísticas en Word", data=word_stats, file_name=f"Estadisticas_{unit}_{stamp}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, on_click="ignore")
            with d2:
                st.download_button("⬇️ Descargar estadísticas en PDF", data=pdf_stats, file_name=f"Estadisticas_{unit}_{stamp}.pdf", mime="application/pdf", use_container_width=True, on_click="ignore")


else:
    if not admin_gate():
        st.stop()

    with st.sidebar:
        page = st.radio(
            "Menú",
            ["Seguimiento mensual", "Estadísticas", "Usuarios y accesos", "Respaldos", "Informe consolidado", "Buscador histórico"]
        )
        st.divider()
        st.markdown("### Herramientas")
        template_bytes = authorized_users_template_bytes()
        st.download_button(
            "⬇️ Plantilla de usuarios",
            data=template_bytes,
            file_name="Plantilla_Usuarios_Autorizados_DIC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=not bool(template_bytes),
        )
        st.caption("Nombre · Correo · Centro · Rol")

    if page == "Seguimiento mensual":
        st.header("Seguimiento mensual")
        if st.session_state.get("admin_delete_success"):
            st.success(st.session_state.pop("admin_delete_success"))
        f1, f2 = st.columns(2)
        with f1:
            month = st.selectbox("Mes", MONTHS, index=datetime.now().month - 1, key="dash_month")
        with f2:
            year = st.selectbox("Año", list(range(2025, 2031)), index=1, key="dash_year")

        rows = get_reports(month, year)
        by_unit = {r["unit_code"]: r for r in rows}

        submitted = sum(1 for r in rows if r["status"] == "ENVIADO")
        draft = sum(1 for r in rows if r["status"] == "BORRADOR")
        pending = len(UNITS) - len(rows)

        m1, m2, m3 = st.columns(3)
        m1.metric("Enviados", f"{submitted}/{len(UNITS)}")
        m2.metric("Borradores", draft)
        m3.metric("Pendientes", pending)

        st.divider()
        for code, full in UNITS.items():
            rep = by_unit.get(code)
            c1, c2, c3, c4, c5 = st.columns([1.3, 3.0, 2.0, 1.1, 2.4])
            c1.markdown(f"### {code}")
            c2.write(full)

            if not rep:
                c3.markdown(
                    """
                    <div style="background:#FDECEC;color:#A33A3A;border-radius:10px;
                    padding:14px 16px;border:1px solid #F4C9C9;">Pendiente</div>
                    """,
                    unsafe_allow_html=True
                )
            elif rep["status"] == "ENVIADO":
                sent_at = format_report_datetime(rep.get("submitted_at"))
                extra = f"<br><span style='font-size:.86rem;color:#52606D'>{sent_at}</span>" if sent_at else ""
                c3.markdown(
                    f"""
                    <div style="background:#EAF5EF;color:#1F6F4A;border-radius:10px;
                    padding:14px 16px;border:1px solid #CDE7D8;">
                    <b>Enviado</b>{extra}
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            else:
                updated_at = format_report_datetime(rep.get("updated_at"))
                extra = f"<br><span style='font-size:.86rem;color:#52606D'>Última actualización: {updated_at}</span>" if updated_at else ""
                c3.markdown(
                    f"""
                    <div style="background:#FFF6DF;color:#8A651A;border-radius:10px;
                    padding:14px 16px;border:1px solid #F0DFA9;">
                    <b>Borrador</b>{extra}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            if rep:
                with c4:
                    if st.button(
                        "🗑 Borrar",
                        key=f"delete_report_{rep['id']}",
                        use_container_width=True,
                    ):
                        delete_report_dialog(
                            rep["id"], rep["unit_code"], rep["month"], rep["year"]
                        )

            if rep:
                saved_activities = get_activities(rep["id"])
                if saved_activities:
                    center_word = generate_saved_center_word(rep, saved_activities)
                    center_pdf = generate_saved_center_pdf(rep, saved_activities)

                    center_ppt = generate_saved_center_ppt(rep, saved_activities)

                    with c5:
                        dw1, dw2, dw3 = st.columns(3)
                        with dw1:
                            st.download_button(
                                "Word",
                                data=center_word,
                                file_name=f"Reporte_{rep['unit_code']}_{rep['month']}_{rep['year']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"track_word_{rep['id']}",
                                use_container_width=True,
                                on_click="ignore",
                            )
                        with dw2:
                            st.download_button(
                                "PDF",
                                data=center_pdf,
                                file_name=f"Reporte_{rep['unit_code']}_{rep['month']}_{rep['year']}.pdf",
                                mime="application/pdf",
                                key=f"track_pdf_{rep['id']}",
                                use_container_width=True,
                                on_click="ignore",
                            )
                        with dw3:
                            st.download_button(
                                "PPT",
                                data=center_ppt,
                                file_name=f"Reporte_{rep['unit_code']}_{rep['month']}_{rep['year']}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                key=f"track_ppt_{rep['id']}",
                                use_container_width=True,
                                on_click="ignore",
                            )


    elif page == "Estadísticas":
        st.header("Estadísticas")
        st.caption(
            "Selecciona los resultados que quieras integrar en la vista previa y en la descarga Word/PDF. "
            "Los filtros se aplican a todos los resultados."
        )

        all_reports = get_reports()
        all_activities = get_activities()

        if not all_reports:
            st.info("Todavía no hay informes guardados para construir estadísticas.")
        else:
            years_available = sorted(
                {int(r.get("year")) for r in all_reports if r.get("year") is not None}
            )
            f1, f2, f3 = st.columns(3)
            with f1:
                stat_center = st.selectbox(
                    "Centro", ["Todos"] + list(UNITS.keys()), key="stats_center"
                )
            with f2:
                stat_year = st.selectbox(
                    "Año", ["Todos"] + years_available, key="stats_year"
                )
            with f3:
                stat_status_label = st.selectbox(
                    "Estatus", ["Todos", "Enviados", "Borradores"], key="stats_status"
                )

            filtered_reports = list(all_reports)
            if stat_center != "Todos":
                filtered_reports = [r for r in filtered_reports if r.get("unit_code") == stat_center]
            if stat_year != "Todos":
                filtered_reports = [r for r in filtered_reports if int(r.get("year")) == int(stat_year)]
            if stat_status_label == "Enviados":
                filtered_reports = [r for r in filtered_reports if r.get("status") == "ENVIADO"]
            elif stat_status_label == "Borradores":
                filtered_reports = [r for r in filtered_reports if r.get("status") == "BORRADOR"]

            filtered_report_ids = {r.get("id") for r in filtered_reports}
            filtered_activities = [
                a for a in all_activities if a.get("report_id") in filtered_report_ids
            ]

            total_reports = len(filtered_reports)
            total_activities = len(filtered_activities)
            total_people = sum(int(a.get("participants") or 0) for a in filtered_activities)
            centers_with_reports = len({r.get("unit_code") for r in filtered_reports})

            month_order = {m: i + 1 for i, m in enumerate(MONTHS)}
            report_counts = []
            for code, full_name in UNITS.items():
                count = sum(1 for r in filtered_reports if r.get("unit_code") == code)
                report_counts.append({"Centro": code, "Nombre": full_name, "Informes": count})
            report_counts_df = pd.DataFrame(report_counts)

            report_activity_rows = []
            activities_by_report = {}
            for a in filtered_activities:
                activities_by_report.setdefault(a.get("report_id"), []).append(a)
            for r in filtered_reports:
                acts = activities_by_report.get(r.get("id"), [])
                report_activity_rows.append({
                    "Año": int(r.get("year")),
                    "Mes": r.get("month"),
                    "Centro": r.get("unit_code"),
                    "Estatus": "Enviado" if r.get("status") == "ENVIADO" else "Borrador",
                    "Actividades": len(acts),
                    "Personas impactadas*": sum(int(a.get("participants") or 0) for a in acts),
                    "_mes_orden": month_order.get(r.get("month"), 99),
                })
            if report_activity_rows:
                report_activity_df = pd.DataFrame(report_activity_rows).sort_values(
                    ["Año", "_mes_orden", "Centro"], ascending=[False, False, True]
                ).drop(columns=["_mes_orden"])
            else:
                report_activity_df = pd.DataFrame(
                    columns=["Año", "Mes", "Centro", "Estatus", "Actividades", "Personas impactadas*"]
                )

            category_counts = {}
            for a in filtered_activities:
                category = (a.get("category") or "Sin categoría").strip() or "Sin categoría"
                category_counts[category] = category_counts.get(category, 0) + 1
            category_df = pd.DataFrame([
                {"Categoría": category, "Actividades": count}
                for category, count in category_counts.items()
            ])
            if not category_df.empty:
                category_df = category_df.sort_values(
                    ["Actividades", "Categoría"], ascending=[False, True]
                )
            else:
                category_df = pd.DataFrame(columns=["Categoría", "Actividades"])

            report_lookup = {r.get("id"): r for r in filtered_reports}
            monthly_people = {}
            for a in filtered_activities:
                r = report_lookup.get(a.get("report_id"))
                if not r:
                    continue
                year = int(r.get("year"))
                month = r.get("month")
                month_num = month_order.get(month, 99)
                key = (year, month_num, month)
                monthly_people[key] = monthly_people.get(key, 0) + int(a.get("participants") or 0)
            monthly_rows = []
            for (year, month_num, month), people in sorted(monthly_people.items()):
                monthly_rows.append({
                    "Periodo": f"{month[:3]} {year}",
                    "Personas impactadas*": people,
                    "_orden": year * 100 + month_num,
                })
            if monthly_rows:
                monthly_df = pd.DataFrame(monthly_rows).sort_values("_orden").drop(columns=["_orden"])
            else:
                monthly_df = pd.DataFrame(columns=["Periodo", "Personas impactadas*"])

            st.markdown("### Selección para vista previa y descarga")
            s1, s2, s3 = st.columns(3)
            with s1:
                include_summary = st.checkbox("Resumen general", value=True, key="stats_sel_summary")
                include_reports = st.checkbox("Informes acumulados por centro", value=True, key="stats_sel_reports")
            with s2:
                include_activities = st.checkbox("Actividades por reporte y mes", value=True, key="stats_sel_activities")
                include_categories = st.checkbox("Actividades por categoría", value=True, key="stats_sel_categories")
            with s3:
                include_people = st.checkbox("Personas impactadas por mes", value=True, key="stats_sel_people")

            selected_stats = {
                "summary": include_summary,
                "reports": include_reports,
                "activities": include_activities,
                "categories": include_categories,
                "people": include_people,
            }
            any_selected = any(selected_stats.values())

            center_text = "Todos los centros" if stat_center == "Todos" else stat_center
            year_text = "Todos los años" if stat_year == "Todos" else str(stat_year)
            filter_text = f"Filtros: {center_text} · {year_text} · {stat_status_label}"
            summary = {
                "Informes": total_reports,
                "Actividades": total_activities,
                "Personas impactadas*": f"{total_people:,}",
                "Centros con informes": centers_with_reports,
            }
            datasets = {
                "reports": report_counts_df,
                "activities": report_activity_df,
                "categories": category_df,
                "people": monthly_df,
            }

            b1, b2 = st.columns([1, 3])
            with b1:
                if st.button(
                    "Previsualizar selección", type="primary", use_container_width=True,
                    disabled=not any_selected,
                ):
                    st.session_state.stats_preview_enabled = True
            with b2:
                st.caption(filter_text)

            if st.session_state.stats_preview_enabled and any_selected:
                st.divider()
                st.subheader("Vista previa de estadísticas seleccionadas")

                if include_summary:
                    with st.container(border=True):
                        st.markdown("#### Resumen general")
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Informes", total_reports)
                        m2.metric("Actividades", total_activities)
                        m3.metric("Personas impactadas*", f"{total_people:,}")
                        m4.metric("Centros con informes", centers_with_reports)

                if include_reports:
                    with st.container(border=True):
                        st.markdown("#### Informes acumulados por centro")
                        c1, c2 = st.columns([1.15, 1.85])
                        with c1:
                            st.dataframe(report_counts_df, use_container_width=True, hide_index=True)
                        with c2:
                            st.bar_chart(report_counts_df[["Centro", "Informes"]].set_index("Centro"), use_container_width=True)

                if include_activities:
                    with st.container(border=True):
                        st.markdown("#### Actividades por reporte y mes")
                        if report_activity_df.empty:
                            st.info("No hay reportes que coincidan con los filtros seleccionados.")
                        else:
                            st.dataframe(report_activity_df, use_container_width=True, hide_index=True)

                if include_categories:
                    with st.container(border=True):
                        st.markdown("#### Actividades por categoría")
                        if category_df.empty:
                            st.info("No hay actividades que coincidan con los filtros seleccionados.")
                        else:
                            st.bar_chart(category_df.set_index("Categoría"), use_container_width=True, horizontal=True)
                            st.dataframe(category_df, use_container_width=True, hide_index=True)

                if include_people:
                    with st.container(border=True):
                        st.markdown("#### Personas impactadas por mes")
                        if monthly_df.empty:
                            st.info("No hay información de participantes para los filtros seleccionados.")
                        else:
                            st.line_chart(monthly_df.set_index("Periodo"), use_container_width=True)
                            st.dataframe(monthly_df, use_container_width=True, hide_index=True)

                st.caption(
                    "* Personas impactadas corresponde a la suma del campo Participantes / alcance. "
                    "Una misma persona puede estar contabilizada en más de una actividad."
                )

                word_stats = generate_statistics_word(selected_stats, filter_text, summary, datasets)
                pdf_stats = generate_statistics_pdf(selected_stats, filter_text, summary, datasets)
                d1, d2 = st.columns(2)
                stamp = datetime.now().strftime("%Y%m%d_%H%M")
                with d1:
                    st.download_button(
                        "⬇️ Descargar selección en Word",
                        data=word_stats,
                        file_name=f"Estadisticas_DIC_{stamp}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        on_click="ignore",
                    )
                with d2:
                    st.download_button(
                        "⬇️ Descargar selección en PDF",
                        data=pdf_stats,
                        file_name=f"Estadisticas_DIC_{stamp}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        on_click="ignore",
                    )
            elif not any_selected:
                st.info("Selecciona al menos un resultado para previsualizarlo o descargarlo.")

            st.caption(
                "Esta es la primera versión del módulo estadístico. Después podemos incorporar indicadores, "
                "comparaciones entre centros, rankings, tendencias y visualizaciones adicionales."
            )


    elif page == "Usuarios y accesos":
        st.header("Usuarios y accesos")
        if st.session_state.get("admin_user_delete_success"):
            st.success(st.session_state.pop("admin_user_delete_success"))
        if st.session_state.get("admin_user_access_success"):
            st.success(st.session_state.pop("admin_user_access_success"))
        st.caption(
            "Administra quién puede ingresar, su centro, su rol y los códigos temporales para activar o restablecer contraseñas. "
            "Esta versión no envía invitaciones por correo: el código se entrega directamente al usuario."
        )

        users = list_authorized_users(include_inactive=True)
        active_users = [u for u in users if u.get("active")]
        directors = [u for u in active_users if u.get("role") == "DIRECTOR"]
        collaborators = [u for u in active_users if u.get("role") == "COLABORADOR"]
        m1, m2, m3 = st.columns(3)
        m1.metric("Usuarios activos", len(active_users))
        m2.metric("Directores", len(directors))
        m3.metric("Colaboradores", len(collaborators))

        st.subheader("Panel de accesos")
        if users:
            display_rows = []
            for u in users:
                display_rows.append({
                    "Nombre": u.get("name") or "",
                    "Correo": u.get("email") or "",
                    "Centro": u.get("unit_code") or "",
                    "Rol": (u.get("role") or "").title(),
                    "Estado": "Activo" if u.get("active") else "Inactivo",
                    "Código temporal": (
                        "Vigente hasta " + format_access_datetime(u.get("activation_code_expires_at"))
                        if u.get("activation_code_hash") else "—"
                    ),
                    "Activado": format_access_datetime(u.get("activated_at")),
                    "Último acceso": format_access_datetime(u.get("last_login_at")),
                })
            st.dataframe(display_rows, use_container_width=True, hide_index=True)
        else:
            st.info("Todavía no hay usuarios autorizados cargados.")

        if st.session_state.get("generated_activation_codes"):
            st.warning(
                "Guarda estos códigos ahora. Por seguridad la aplicación sólo almacena una huella del código "
                "y no podrá volver a mostrar el mismo código después."
            )
            st.dataframe(st.session_state.generated_activation_codes, use_container_width=True, hide_index=True)
            if XLWorkbook is not None:
                code_wb = XLWorkbook()
                code_ws = code_wb.active
                code_ws.title = "Codigos"
                headers_codes = ["Nombre", "Correo", "Centro", "Rol", "Código", "Vence"]
                code_ws.append(headers_codes)
                for item in st.session_state.generated_activation_codes:
                    code_ws.append([item.get(h, "") for h in headers_codes])
                code_bio = io.BytesIO()
                code_wb.save(code_bio)
                st.download_button(
                    "⬇️ Descargar códigos de activación",
                    data=code_bio.getvalue(),
                    file_name="Codigos_Activacion_DIC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            if st.button("Ocultar códigos mostrados", use_container_width=True):
                st.session_state.generated_activation_codes = []
                st.rerun()

        st.divider()
        st.subheader("Agregar o actualizar usuarios desde Excel")
        st.download_button(
            "⬇️ Descargar plantilla",
            data=authorized_users_template_bytes(),
            file_name="Plantilla_Usuarios_Autorizados_DIC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        uploaded_users = st.file_uploader(
            "Cargar plantilla de usuarios (.xlsx)",
            type=["xlsx"],
            key="authorized_users_upload",
        )
        generate_codes = st.checkbox(
            "Generar código temporal para usuarios nuevos",
            value=True,
            help="El código se muestra una sola vez en Administración. Entrégalo al usuario por un medio separado.",
        )
        if uploaded_users is not None:
            parsed_users, excel_errors = parse_authorized_users_excel(uploaded_users.getvalue())
            if excel_errors:
                st.error("Corrige la plantilla antes de cargarla:")
                for err in excel_errors:
                    st.markdown(f"- {err}")
            else:
                st.success(f"Archivo válido: {len(parsed_users)} usuario(s).")
                st.dataframe([
                    {"Nombre": u["name"], "Correo": u["email"], "Centro": u["unit_code"], "Rol": u["role"]}
                    for u in parsed_users
                ], use_container_width=True, hide_index=True)
                if st.button("Agregar / actualizar usuarios", type="primary", use_container_width=True):
                    results, generated_codes = upsert_authorized_users(parsed_users, generate_codes=generate_codes)
                    st.session_state.generated_activation_codes = generated_codes
                    for email, ok, message in results:
                        if ok:
                            st.success(f"{email} · {message}")
                        else:
                            st.error(f"{email} · {message}")
                    st.rerun()

        if users:
            st.divider()
            st.subheader("Activar o desactivar un acceso")
            selected_email = st.selectbox(
                "Usuario",
                [u.get("email") for u in users],
                format_func=lambda e: next(
                    (f"{u.get('name','')} · {e} · {u.get('unit_code','')} · {u.get('role','')}" for u in users if u.get("email") == e),
                    e,
                ),
            )
            selected_user = next((u for u in users if u.get("email") == selected_email), None)
            if selected_user:
                new_state = not bool(selected_user.get("active"))
                label = "Reactivar acceso" if new_state else "Desactivar acceso"
                if st.button(label, use_container_width=True):
                    ok, msg = set_authorized_user_active(selected_user, new_state)
                    if ok:
                        st.session_state["admin_user_access_success"] = msg
                        st.rerun()
                    else:
                        st.error(msg)

                st.caption(
                    "Si la persona no ha activado su cuenta o olvidó su contraseña, genera un código nuevo. "
                    "El código anterior quedará invalidado."
                )
                if st.button("Generar código de activación / restablecimiento", use_container_width=True):
                    ok, msg, code, expires = set_activation_code(selected_email)
                    if ok:
                        st.session_state.generated_activation_codes = [{
                            "Nombre": selected_user.get("name", ""),
                            "Correo": selected_user.get("email", ""),
                            "Centro": selected_user.get("unit_code", ""),
                            "Rol": selected_user.get("role", ""),
                            "Código": code,
                            "Vence": format_access_datetime(expires),
                        }]
                        st.rerun()
                    else:
                        st.error(msg)

                st.markdown("#### Eliminar usuario")
                st.caption(
                    "Elimina a la persona de la lista de accesos. Si ya había creado contraseña, "
                    "también se elimina su identidad de acceso. Los reportes históricos se conservan."
                )
                if st.button(
                    "🗑️ Eliminar usuario",
                    use_container_width=True,
                    key=f"delete_authorized_user_{selected_email}",
                ):
                    delete_authorized_user_dialog(selected_user)


    elif page == "Respaldos":
        st.header("Respaldos")
        st.caption(
            "Genera una copia descargable de las tablas de la base de datos. "
            "El respaldo se entrega en un archivo ZIP con cada tabla en CSV y JSON."
        )
        st.info(
            "Este respaldo incluye datos de usuarios autorizados, informes, actividades, referencias de fotografías, "
            "unidades y bitácora. No incluye las fotografías/gráficas binarias de Supabase Storage; sus rutas sí quedan respaldadas."
        )

        if not supabase:
            st.warning("El respaldo de base de datos sólo está disponible cuando la aplicación está conectada a Supabase.")
        else:
            if st.button("Preparar respaldo de base de datos", type="primary", use_container_width=True):
                with st.spinner("Preparando respaldo..."):
                    try:
                        backup_bytes, backup_meta = generate_database_backup_zip()
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        st.session_state.database_backup_bytes = backup_bytes
                        st.session_state.database_backup_filename = f"Respaldo_DIC_{timestamp}.zip"
                        st.session_state.database_backup_meta = backup_meta
                    except Exception as exc:
                        st.session_state.database_backup_bytes = None
                        st.error(f"No fue posible generar el respaldo. Detalle: {exc}")

            if st.session_state.get("database_backup_bytes"):
                meta = st.session_state.get("database_backup_meta") or {}
                table_counts = meta.get("tables") or {}
                st.success("Respaldo preparado correctamente.")
                if table_counts:
                    rows = [{"Tabla": k, "Registros": v} for k, v in table_counts.items()]
                    st.dataframe(rows, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Descargar respaldo ZIP",
                    data=st.session_state.database_backup_bytes,
                    file_name=st.session_state.database_backup_filename or "Respaldo_DIC.zip",
                    mime="application/zip",
                    use_container_width=True,
                    on_click="ignore",
                )
                st.caption(
                    "Recomendación: guarda el ZIP en un espacio institucional seguro y genera un respaldo periódico."
                )


    elif page == "Informe consolidado":
        st.header("Informe consolidado")
        f1, f2 = st.columns(2)
        with f1:
            month = st.selectbox("Mes", MONTHS, index=datetime.now().month - 1, key="rep_month")
        with f2:
            year = st.selectbox("Año", list(range(2025, 2031)), index=1, key="rep_year")

        reports = get_reports(month, year)
        reports = [r for r in reports if r["status"] == "ENVIADO"]

        if not reports:
            st.info("Todavía no hay reportes enviados para este periodo.")
        else:
            acts_by_report = {r["id"]: get_activities(r["id"]) for r in reports}

            st.subheader("Edición DIC")
            st.caption(
                "La edición no modifica el texto original del centro. Se guarda una versión editada para el consolidado."
            )

            st.caption(
                "Las actividades aparecen ordenadas por prioridad dentro de cada centro. "
                "Marca únicamente las que deben pasar a la versión final."
            )

            for rep in reports:
                center_activities = sorted(
                    acts_by_report[rep["id"]],
                    key=activity_sort_key
                )

                with st.expander(
                    f"{rep['unit_code']} · {UNITS[rep['unit_code']]} · "
                    f"{len(center_activities)} actividad(es)",
                    expanded=False,
                ):
                    for a in center_activities:
                        rank_text = rank_label(a.get("ranking"))

                        st.markdown(
                            f"""
                            <div style="
                                color:{ITESO_BLUE};
                                font-size:1.35rem;
                                line-height:1.25;
                                font-weight:800;
                                margin:8px 0 2px 0;">
                                {a['title']}
                            </div>
                            <div style="
                                color:#64748B;
                                font-size:.92rem;
                                margin-bottom:8px;">
                                {rank_text} · {a.get('category','')}
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                        edited = st.text_area(
                            "Texto para consolidado",
                            value=a.get("description_edited") or a.get("description_original") or "",
                            key=f"edit_{a['id']}",
                            height=125,
                            label_visibility="collapsed",
                        )

                        edit_col, select_col = st.columns([1, 2.5])
                        with edit_col:
                            if st.button(
                                "Guardar edición",
                                key=f"save_{a['id']}",
                                use_container_width=True
                            ):
                                update_edited_description(a["id"], edited)
                                st.success("Edición guardada.")

                        with select_col:
                            current_selection = selected_for_final(a["id"])
                            selected = st.checkbox(
                                "✓ Incluir en la versión final del consolidado",
                                value=current_selection,
                                key=f"final_select_{a['id']}",
                            )
                            st.session_state.final_selected_activities[a["id"]] = selected

                        photos = get_activity_photos(a["id"])
                        if photos:
                            st.caption(f"Fotografías cargadas: {len(photos)}")
                            photo_cols = st.columns(min(3, len(photos)))
                            for idx, ph in enumerate(photos):
                                with photo_cols[idx % len(photo_cols)]:
                                    st.image(
                                        ph["bytes"],
                                        caption=ph.get("original_filename", f"Foto {idx+1}"),
                                        use_container_width=True
                                    )
                        else:
                            st.caption("Sin fotografías cargadas.")

                        st.markdown(
                            "<div style='height:1px;background:#E2E8F0;margin:20px 0;'></div>",
                            unsafe_allow_html=True
                        )

            # Reload after edits
            acts_by_report = {r["id"]: get_activities(r["id"]) for r in reports}

            pv_col, _ = st.columns([1, 3])
            with pv_col:
                if st.button(
                    "👁️ Previsualizar consolidado"
                    if not st.session_state.show_consolidated_preview
                    else "Ocultar vista previa",
                    use_container_width=True,
                ):
                    st.session_state.show_consolidated_preview = not st.session_state.show_consolidated_preview
                    st.rerun()

            if st.session_state.show_consolidated_preview:
                render_consolidated_preview(month, year, reports, acts_by_report)
                st.divider()

            selected_count = sum(
                1
                for rep in reports
                for a in acts_by_report.get(rep["id"], [])
                if selected_for_final(a["id"])
            )

            if selected_count == 0:
                st.warning(
                    "Todavía no has seleccionado actividades para la versión final. "
                    "Marca al menos una actividad con el check correspondiente."
                )
                word_bytes = b""
                pdf_bytes = b""
            else:
                st.info(f"Actividades seleccionadas para la versión final: **{selected_count}**")
                word_bytes = generate_word(month, year, reports, acts_by_report)
                pdf_bytes = generate_pdf(month, year, reports, acts_by_report)
                ppt_bytes = generate_consolidated_ppt(month, year, reports, acts_by_report)

            d1, d2, d3 = st.columns(3)
            with d1:
                st.download_button(
                    "⬇️ Descargar Word",
                    data=word_bytes,
                    file_name=f"Informe_DIC_{month}_{year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    disabled=(selected_count == 0)
                )
            with d2:
                st.download_button(
                    "⬇️ Descargar PDF",
                    data=pdf_bytes,
                    file_name=f"Informe_DIC_{month}_{year}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    disabled=(selected_count == 0)
                )
            with d3:
                st.download_button(
                    "⬇️ Descargar PPT",
                    data=ppt_bytes if selected_count else b"",
                    file_name=f"Informe_DIC_{month}_{year}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    disabled=(selected_count == 0)
                )

    else:
        st.header("Buscador histórico")
        query = st.text_input(
            "Buscar palabras o temas",
            placeholder="Ej. desaparición, inclusión, comunidades indígenas, AUSJAL..."
        )
        c1, c2 = st.columns(2)
        with c1:
            filter_unit = st.selectbox("Centro", ["Todos"] + list(UNITS.keys()))
        with c2:
            filter_rank = st.selectbox("Ranking", ["Todos", "Top 1", "Top 2", "Top 3", "Sin ranking"])

        if query.strip():
            reports = get_reports()
            report_map = {r["id"]: r for r in reports}
            hits = []
            q = query.lower().strip()
            for a in get_activities():
                rep = report_map.get(a["report_id"])
                if not rep:
                    continue
                if filter_unit != "Todos" and rep["unit_code"] != filter_unit:
                    continue
                if filter_rank != "Todos" and rank_label(a.get("ranking")) != filter_rank:
                    continue
                haystack = " ".join([
                    a.get("title", ""),
                    a.get("description_original", ""),
                    a.get("description_edited", ""),
                    a.get("category", ""),
                ]).lower()
                if q in haystack:
                    hits.append((rep, a))

            st.write(f"**{len(hits)} resultado(s)**")
            for rep, a in hits:
                with st.container(border=True):
                    st.markdown(f"### {a['title']}")
                    st.caption(
                        f"{rep['unit_code']} · {rep['month']} {rep['year']} · "
                        f"{a.get('category','')} · {rank_label(a.get('ranking'))}"
                    )
                    st.write(a.get('description_edited') or a.get('description_original') or '')

                    photos = get_activity_photos(a["id"])
                    if photos:
                        pcols = st.columns(min(3, len(photos)))
                        for idx, ph in enumerate(photos):
                            with pcols[idx % len(pcols)]:
                                st.image(
                                    ph["bytes"],
                                    caption=ph.get("original_filename", f"Foto {idx+1}"),
                                    use_container_width=True
                                )

                    word_seg = generate_segment_word(rep, a, photos)
                    pdf_seg = generate_segment_pdf(rep, a, photos)
                    d1, d2 = st.columns(2)
                    safe_title = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ_-]+", "_", a["title"])[:60]
                    with d1:
                        st.download_button(
                            "⬇️ Descargar extracto Word",
                            data=word_seg,
                            file_name=f"{rep['unit_code']}_{rep['month']}_{rep['year']}_{safe_title}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"hist_word_{a['id']}",
                            use_container_width=True,
                            on_click="ignore",
                        )
                    with d2:
                        st.download_button(
                            "⬇️ Descargar extracto PDF",
                            data=pdf_seg,
                            file_name=f"{rep['unit_code']}_{rep['month']}_{rep['year']}_{safe_title}.pdf",
                            mime="application/pdf",
                            key=f"hist_pdf_{a['id']}",
                            use_container_width=True,
                            on_click="ignore",
                        )
        else:
            st.info("Escribe una palabra o tema para buscar en el histórico.")

st.divider()
st.caption("Prototipo V1.38 · Dirección de Integración Comunitaria · ITESO")
