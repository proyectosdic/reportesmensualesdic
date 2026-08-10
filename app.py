
import os
import io
import re
import uuid
from datetime import datetime
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

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

CATEGORIES = [
    "Formación",
    "Vinculación interna",
    "Vinculación externa",
    "Incidencia social",
    "Cultura",
    "Salud y bienestar",
    "Deporte",
    "Pastoral / identidad ignaciana",
    "Derechos humanos",
    "Inclusión",
    "Sustentabilidad",
    "Otro",
]

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
if "num_activities" not in st.session_state:
    st.session_state.num_activities = 5

def save_report(unit, month, year, status, sender_email=""):
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
            supabase.table("reports").update({
                "status": status,
                "sender_email": sender_email,
                "submitted_at": datetime.utcnow().isoformat() if status == "ENVIADO" else None,
                "updated_at": datetime.utcnow().isoformat(),
            }).eq("id", rid).execute()
            return rid
        res = supabase.table("reports").insert({
            "unit_code": unit,
            "month": month,
            "year": year,
            "status": status,
            "sender_email": sender_email,
        }).execute()
        return res.data[0]["id"]

    # Demo session storage
    found = next((r for r in st.session_state.demo_reports
                  if r["unit_code"] == unit and r["month"] == month and r["year"] == year), None)
    if found:
        found["status"] = status
        found["sender_email"] = sender_email
        found["updated_at"] = datetime.now().isoformat()
        if status == "ENVIADO":
            found["submitted_at"] = datetime.now().isoformat()
        return found["id"]
    rid = str(uuid.uuid4())
    st.session_state.demo_reports.append({
        "id": rid, "unit_code": unit, "month": month, "year": year,
        "status": status, "sender_email": sender_email,
        "submitted_at": datetime.now().isoformat() if status == "ENVIADO" else None,
        "created_at": datetime.now().isoformat(),
    })
    return rid

def replace_activities(report_id, activities):
    """Replace activities and persist uploaded photos.

    Demo mode stores image bytes in session state.
    Supabase mode stores files in Storage bucket `dic-activity-photos`
    and metadata in `activity_photos`.
    """
    if supabase:
        # Remove old metadata/activities. Storage objects are left with unique paths
        # to avoid CDN stale-file issues; they can be cleaned later by an admin job.
        old_acts = supabase.table("activities").select("id").eq("report_id", report_id).execute().data or []
        for oa in old_acts:
            supabase.table("activity_photos").delete().eq("activity_id", oa["id"]).execute()
        supabase.table("activities").delete().eq("report_id", report_id).execute()

        for i, a in enumerate(activities, start=1):
            res = supabase.table("activities").insert({
                "report_id": report_id,
                "title": a["title"],
                "description_original": a["description"],
                "description_edited": a["description"],
                "category": a["category"],
                "ranking": a["ranking"],
                "activity_date": a.get("activity_date").isoformat() if a.get("activity_date") else None,
                "participants": a.get("participants"),
                "order_index": i,
            }).execute()
            activity_id = res.data[0]["id"]

            for photo in a.get("photos", []):
                try:
                    photo_bytes = photo.getvalue()
                    ext = Path(photo.name).suffix.lower() or ".jpg"
                    unique_name = f"{uuid.uuid4().hex}{ext}"
                    storage_path = f"{report_id}/{activity_id}/{unique_name}"
                    supabase.storage.from_("dic-activity-photos").upload(
                        path=storage_path,
                        file=photo_bytes,
                        file_options={
                            "content-type": getattr(photo, "type", "image/jpeg") or "image/jpeg",
                            "upsert": "false",
                        },
                    )
                    supabase.table("activity_photos").insert({
                        "activity_id": activity_id,
                        "storage_path": storage_path,
                        "original_filename": photo.name,
                    }).execute()
                except Exception as e:
                    st.warning(f"No se pudo guardar una fotografía de '{a['title']}': {e}")
        return

    # Demo session storage
    old_ids = [a["id"] for a in st.session_state.demo_activities if a["report_id"] == report_id]
    st.session_state.demo_photos = [
        p for p in st.session_state.demo_photos if p["activity_id"] not in old_ids
    ]
    st.session_state.demo_activities = [
        a for a in st.session_state.demo_activities if a["report_id"] != report_id
    ]

    for i, a in enumerate(activities, start=1):
        aid = str(uuid.uuid4())
        st.session_state.demo_activities.append({
            "id": aid,
            "report_id": report_id,
            "title": a["title"],
            "description_original": a["description"],
            "description_edited": a["description"],
            "category": a["category"],
            "ranking": a["ranking"],
            "activity_date": str(a.get("activity_date") or ""),
            "participants": a.get("participants"),
            "order_index": i,
        })
        for photo in a.get("photos", []):
            st.session_state.demo_photos.append({
                "id": str(uuid.uuid4()),
                "activity_id": aid,
                "original_filename": photo.name,
                "mime_type": getattr(photo, "type", "image/jpeg") or "image/jpeg",
                "bytes": photo.getvalue(),
            })


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


def check_admin_password(password):
    try:
        configured = st.secrets.get("ADMIN_PASSWORD", "")
    except Exception:
        configured = ""
    if not configured:
        return False, "ADMIN_PASSWORD no está configurada en los Secrets de Streamlit."
    if password == configured:
        st.session_state.admin_authenticated = True
        return True, ""
    return False, "Contraseña incorrecta."


def admin_gate():
    if st.session_state.admin_authenticated:
        if st.sidebar.button("Cerrar sesión de administración", use_container_width=True):
            st.session_state.admin_authenticated = False
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


def render_center_preview(unit, month, year, activities):
    st.markdown(f"### Vista previa · {unit}")
    st.caption(f"{UNITS[unit]} · {month} {year}")
    for i, act in enumerate(activities, start=1):
        with st.container(border=True):
            st.markdown(f"#### {i}. {act['title']}")
            meta = [act.get("category", ""), rank_label(act.get("ranking"))]
            if act.get("participants"):
                meta.append(f"Participantes / alcance: {act['participants']}")
            st.caption(" · ".join([x for x in meta if x]))
            st.write(act.get("description", ""))
            photos = current_uploaded_photos(act)
            if photos:
                cols = st.columns(min(3, len(photos)))
                for idx, ph in enumerate(photos):
                    with cols[idx % len(cols)]:
                        st.image(ph["bytes"], use_container_width=True)


def render_consolidated_preview(month, year, reports, activities_by_report):
    st.markdown("### Vista previa del informe consolidado")
    st.caption(f"Dirección de Integración Comunitaria · {month} {year}")

    ranked = []
    for rep in reports:
        for act in activities_by_report.get(rep["id"], []):
            if act.get("ranking") in [1, 2, 3, "1", "2", "3"]:
                ranked.append((rep, act))
    ranked.sort(key=lambda x: int(x[1]["ranking"]))

    if ranked:
        st.markdown("#### Principales hitos del mes")
        for rep, act in ranked:
            with st.container(border=True):
                st.markdown(f"**{rank_label(act.get('ranking'))} · {rep['unit_code']} · {act['title']}**")
                st.write(act.get("description_edited") or act.get("description_original") or "")
                photos = get_activity_photos(act["id"])
                if photos:
                    cols = st.columns(min(3, len(photos)))
                    for idx, ph in enumerate(photos):
                        with cols[idx % len(cols)]:
                            st.image(ph["bytes"], use_container_width=True)

    st.markdown("#### Hitos por centro")
    for rep in reports:
        st.markdown(f"### {rep['unit_code']} · {UNITS.get(rep['unit_code'], '')}")
        acts = activities_by_report.get(rep["id"], [])
        if not acts:
            st.caption("Sin actividades registradas.")
            continue
        for act in acts:
            with st.container(border=True):
                st.markdown(f"**{act['title']}** · {rank_label(act.get('ranking'))}")
                st.write(act.get("description_edited") or act.get("description_original") or "")
                photos = get_activity_photos(act["id"])
                if photos:
                    cols = st.columns(min(3, len(photos)))
                    for idx, ph in enumerate(photos):
                        with cols[idx % len(cols)]:
                            st.image(ph["bytes"], use_container_width=True)


def word_count(text):
    return len(re.findall(r"\b[\wÁÉÍÓÚÜÑáéíóúüñ'-]+\b", text or ""))

def rank_label(v):
    if v in [1, "1", "Top 1"]:
        return "Top 1"
    if v in [2, "2", "Top 2"]:
        return "Top 2"
    if v in [3, "3", "Top 3"]:
        return "Top 3"
    return "Sin ranking"


def activity_fields_complete(index):
    title = (st.session_state.get(f"title_{index}") or "").strip()
    desc = (st.session_state.get(f"desc_{index}") or "").strip()
    category_selected = st.session_state.get(f"cat_{index}", CATEGORIES[0])
    other_category = (st.session_state.get(f"other_cat_{index}") or "").strip()
    participants = st.session_state.get(f"part_{index}", 0) or 0
    if not title or not desc or word_count(desc) > 250 or participants <= 0:
        return False
    if category_selected == "Otro" and not other_category:
        return False
    return True


def current_uploaded_photos(activity):
    normalized = []
    for ph in activity.get("photos", []) or []:
        try:
            normalized.append({
                "bytes": ph.getvalue(),
                "mime_type": getattr(ph, "type", "image/jpeg") or "image/jpeg",
                "original_filename": getattr(ph, "name", "fotografia"),
            })
        except Exception:
            pass
    return normalized



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

    doc.add_paragraph("Principales hitos del mes")

    ranked = []
    unranked = []
    for rep in reports:
        for act in activities_by_report.get(rep["id"], []):
            item = (rep, act)
            if act.get("ranking") in [1, 2, 3, "1", "2", "3"]:
                ranked.append(item)
            else:
                unranked.append(item)
    ranked.sort(key=lambda x: int(x[1]["ranking"]))

    for rep, act in ranked:
        h = doc.add_paragraph()
        rr = h.add_run(f"{rank_label(act.get('ranking'))} · {rep['unit_code']} · {act['title']}")
        rr.bold = True
        rr.font.color.rgb = RGBColor(0, 76, 127)
        doc.add_paragraph(act.get("description_edited") or act.get("description_original") or "")

    doc.add_paragraph("Hitos por centro")
    for rep in reports:
        doc.add_heading(f"{rep['unit_code']} — {UNITS.get(rep['unit_code'], rep['unit_code'])}", level=2)
        acts = activities_by_report.get(rep["id"], [])
        if not acts:
            doc.add_paragraph("Sin actividades registradas.")
            continue
        for act in acts:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(act["title"])
            r.bold = True
            if act.get("ranking"):
                r.add_text(f" ({rank_label(act['ranking'])})")
            p.add_run(" — " + (act.get("description_edited") or act.get("description_original") or ""))

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
        Paragraph("Principales hitos del mes", h2),
    ]

    ranked = []
    for rep in reports:
        for act in activities_by_report.get(rep["id"], []):
            if act.get("ranking") in [1, 2, 3, "1", "2", "3"]:
                ranked.append((rep, act))
    ranked.sort(key=lambda x: int(x[1]["ranking"]))

    for rep, act in ranked:
        story.append(Paragraph(
            f"<b>{rank_label(act.get('ranking'))} · {rep['unit_code']} · {act['title']}</b>", body
        ))
        story.append(Paragraph(act.get("description_edited") or act.get("description_original") or "", body))
        story.append(Spacer(1, 7))

    story.append(Paragraph("Hitos por centro", h2))
    for rep in reports:
        story.append(Paragraph(f"{rep['unit_code']} — {UNITS.get(rep['unit_code'], '')}", h2))
        acts = activities_by_report.get(rep["id"], [])
        if not acts:
            story.append(Paragraph("Sin actividades registradas.", body))
        for act in acts:
            tag = f" <b>({rank_label(act['ranking'])})</b>" if act.get("ranking") else ""
            story.append(Paragraph(
                f"• <b>{act['title']}</b>{tag} — {act.get('description_edited') or act.get('description_original') or ''}",
                body
            ))
            story.append(Spacer(1, 5))

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



def generate_center_word(unit, month, year, activities):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)
    add_docx_page_x_of_y(sec)
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(logo), width=Inches(2.2))
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Informe de actividades"); r.bold = True; r.font.size = Pt(18); r.font.color.rgb = RGBColor(0,76,127)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"Dirección de Integración Comunitaria\n{month} {year}")
    doc.add_paragraph(f"{unit} — {UNITS[unit]}").runs[0].bold = True
    for i, act in enumerate(activities, start=1):
        h = doc.add_paragraph(); rr = h.add_run(f"{i}. {act['title']}"); rr.bold = True; rr.font.color.rgb = RGBColor(0,76,127)
        meta = [f"Categoría: {act.get('category','')}", rank_label(act.get('ranking'))]
        if act.get('participants'): meta.append(f"Participantes / alcance: {act['participants']}")
        doc.add_paragraph(" · ".join(meta))
        doc.add_paragraph(act.get("description") or "")
        for ph in current_uploaded_photos(act):
            try: doc.add_picture(io.BytesIO(ph["bytes"]), width=Inches(5.7))
            except Exception: pass
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()


def generate_center_pdf(unit, month, year, activities):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter, rightMargin=45, leftMargin=45, topMargin=45, bottomMargin=45)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("centerTitle", parent=styles["Title"], textColor=HexColor(ITESO_BLUE), fontSize=18, leading=22, alignment=TA_CENTER, spaceAfter=8)
    h_style = ParagraphStyle("centerH", parent=styles["Heading2"], textColor=HexColor(ITESO_BLUE), fontSize=13, leading=16, spaceBefore=10, spaceAfter=6)
    body = styles["BodyText"]; body.fontSize = 9.5; body.leading = 13
    story = []
    logo = Path("assets/iteso_logo.png")
    if logo.exists():
        img = Image(str(logo), width=190, height=45); img.hAlign = "RIGHT"; story += [img, Spacer(1,8)]
    story += [Paragraph("Informe de actividades", title_style), Paragraph(f"Dirección de Integración Comunitaria<br/>{month} {year}", styles["Heading3"]), Spacer(1,10), Paragraph(f"{unit} — {UNITS[unit]}", h_style)]
    for i, act in enumerate(activities, start=1):
        story.append(Paragraph(f"{i}. {act['title']}", h_style))
        meta = [f"Categoría: {act.get('category','')}", rank_label(act.get('ranking'))]
        if act.get('participants'): meta.append(f"Participantes / alcance: {act['participants']}")
        story.append(Paragraph(" · ".join(meta), body)); story.append(Paragraph(act.get("description") or "", body))
        for ph in current_uploaded_photos(act):
            try:
                img = Image(io.BytesIO(ph["bytes"])); max_w,max_h = 430,280
                ratio = min(max_w/img.imageWidth, max_h/img.imageHeight, 1)
                img.drawWidth = img.imageWidth*ratio; img.drawHeight = img.imageHeight*ratio
                story += [Spacer(1,8), img]
            except Exception: pass
        story.append(Spacer(1,8))
    doc.build(story, canvasmaker=NumberedCanvas); return bio.getvalue()


@st.dialog("Confirmar envío", dismissible=False)
def confirm_submission_dialog(unit, month, year, sender_email, activities):
    st.write(f"Vas a enviar el reporte de **{UNITS[unit]}** correspondiente a **{month} {year}**.")
    st.warning("Confirma que ya registraste todas las actividades que deseas reportar. Después del envío el reporte quedará marcado como enviado.")
    confirm_all = st.checkbox("Confirmo que ya no tengo más actividades que agregar.")
    c1,c2 = st.columns(2)
    with c1:
        if st.button("Volver", use_container_width=True): st.rerun()
    with c2:
        if st.button("Sí, enviar reporte", type="primary", use_container_width=True, disabled=not confirm_all):
            rid = save_report(unit, month, year, "ENVIADO", sender_email)
            replace_activities(rid, activities)
            email_result = send_confirmation_email(sender_email, unit, month, year)
            st.session_state["last_submission_message"] = f"Reporte de {month} {year} enviado correctamente. " + email_result["message"]
            st.rerun()


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

# ---------- DEMO ACCESS ----------
with st.sidebar:
    st.markdown("### Acceso")
    profile = st.radio("Perfil", ["Centro / Dirección", "Administración DIC"], label_visibility="collapsed")
    st.caption(f"Base de datos: **{db_mode()}**")
    st.divider()

if profile == "Centro / Dirección":
    with st.sidebar:
        unit = st.selectbox("Centro", list(UNITS.keys()), format_func=lambda x: f"{x} · {UNITS[x]}")
        sender_email = st.text_input("Correo institucional", placeholder="nombre@iteso.mx")

        # Cambiar el correo invalida una validación anterior.
        if st.session_state.validated_center_email and (
            sender_email.strip().lower() != st.session_state.validated_center_email.lower()
        ):
            st.session_state.validated_center_email = ""
            st.session_state.show_center_preview = False

        email_validated = (
            bool(sender_email.strip())
            and st.session_state.validated_center_email.lower() == sender_email.strip().lower()
        )

        if email_validated:
            st.success("Correo institucional validado")
        else:
            if st.button("Validar correo institucional", use_container_width=True):
                if email_is_iteso(sender_email):
                    st.session_state.validated_center_email = sender_email.strip()
                    st.success("Correo institucional validado.")
                    st.rerun()
                else:
                    st.error("Ingresa un correo institucional válido con dominio @iteso.mx.")

        st.divider()
        page = st.radio("Menú", ["Nuevo reporte", "Mis reportes"], disabled=not email_validated)

    if not email_validated:
        st.header("Valida tu correo institucional")
        st.info(
            "Para acceder a la captura y a tus reportes, primero valida un correo institucional "
            "con dominio @iteso.mx en el menú lateral."
        )
        st.stop()

    if page == "Nuevo reporte":
        st.header("Nuevo reporte mensual")
        if st.session_state.get("last_submission_message"):
            st.success(st.session_state.pop("last_submission_message"))
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.info(f"**{unit}** · {UNITS[unit]}")
        with col2:
            month = st.selectbox("Mes", MONTHS, index=datetime.now().month - 1)
        with col3:
            year = st.selectbox("Año", list(range(2025, 2031)), index=1)

        st.caption(
            "Registra inicialmente hasta 5 hitos. Si necesitas más, utiliza **Agregar actividad**. "
            "Cada descripción admite un máximo de 250 palabras. Toda actividad iniciada debe tener completos "
            "sus campos obligatorios antes de continuar; la fotografía es opcional. Sólo puede existir un Top 1, Top 2 y Top 3."
        )

        activities = []
        errors = []

        for i in range(st.session_state.num_activities):
            previous_complete = True if i == 0 else activity_fields_complete(i - 1)
            with st.expander(f"Actividad {i+1}", expanded=(i < 2 and previous_complete)):
                if not previous_complete:
                    st.info(f"🔒 Completa todos los campos obligatorios de la Actividad {i} antes de capturar la Actividad {i+1}.")
                    continue
                title = st.text_input("Nombre del hito / actividad", key=f"title_{i}")
                c1,c2,c3 = st.columns([2,1,1])
                with c1:
                    category_selected = st.selectbox("Tema / categoría", CATEGORIES, key=f"cat_{i}")
                    other_category = ""
                    if category_selected == "Otro":
                        other_category = st.text_input("Especifica la categoría", key=f"other_cat_{i}", placeholder="Escribe la categoría")
                    category = other_category.strip() if category_selected == "Otro" and other_category.strip() else category_selected
                with c2:
                    ranking_text = st.selectbox("Importancia", ["Sin ranking","Top 1","Top 2","Top 3"], key=f"rank_{i}")
                with c3:
                    participants = st.number_input("Participantes / alcance", min_value=0, step=1, value=0, key=f"part_{i}", help="Campo obligatorio para una actividad capturada.")
                desc = st.text_area("Descripción del hito", height=150, placeholder="Qué ocurrió, por qué fue relevante, resultados y actores participantes.", key=f"desc_{i}")
                wc = word_count(desc)
                if wc > 250: st.error(f"{wc}/250 palabras. Reduce la descripción en {wc-250} palabras.")
                else: st.caption(f"{wc}/250 palabras")
                photos = st.file_uploader("Fotografías (opcional)", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"photos_{i}", help="Puedes cargar una o más fotografías. Este campo es opcional.")
                activity_started = bool(title.strip() or desc.strip() or participants > 0 or ranking_text != "Sin ranking" or category_selected != CATEGORIES[0] or photos)
                if activity_started:
                    ranking = None if ranking_text == "Sin ranking" else int(ranking_text[-1])
                    activities.append({"title":title.strip(),"description":desc.strip(),"category":category,"ranking":ranking,"participants":participants or None,"photos":photos or []})
                    missing=[]
                    if not title.strip(): missing.append("nombre del hito / actividad")
                    if not desc.strip(): missing.append("descripción")
                    if participants <= 0: missing.append("participantes / alcance")
                    if category_selected == "Otro" and not other_category.strip(): missing.append("categoría específica")
                    if missing: errors.append(f"Actividad {i+1}: completa " + ", ".join(missing) + ".")
                    if wc > 250: errors.append(f"Actividad {i+1}: excede 250 palabras.")
                    if missing: st.warning("Esta actividad está incompleta. Antes de continuar, completa: " + ", ".join(missing) + ". La fotografía es opcional.")

        cadd, crem = st.columns([1, 4])
        with cadd:
            if st.button("➕ Agregar actividad"):
                incomplete = [
                    e for e in errors
                    if "completa" in e.lower() or "excede 250 palabras" in e.lower()
                ]
                if incomplete:
                    st.error("Completa primero todas las actividades iniciadas antes de agregar otra.")
                else:
                    st.session_state.num_activities += 1
                    st.rerun()
        with crem:
            if st.session_state.num_activities > 5 and st.button("➖ Quitar última"):
                st.session_state.num_activities -= 1
                st.rerun()

        ranks = [a["ranking"] for a in activities if a["ranking"]]
        if len(ranks) != len(set(ranks)):
            errors.append("Top 1, Top 2 y Top 3 no pueden repetirse.")

        st.divider()

        st.subheader("Vista previa del reporte")
        st.caption(
            "Puedes previsualizar, descargar y revisar el reporte antes de enviarlo. "
            "Las fotografías cargadas se incluirán en estos archivos."
        )

        preview_ready = bool(activities) and not errors

        pv_col, _ = st.columns([1, 3])
        with pv_col:
            if st.button(
                "👁️ Previsualizar reporte" if not st.session_state.show_center_preview else "Ocultar vista previa",
                use_container_width=True,
            ):
                st.session_state.show_center_preview = not st.session_state.show_center_preview
                st.rerun()

        if preview_ready:
            preview_word = generate_center_word(unit, month, year, activities)
            preview_pdf = generate_center_pdf(unit, month, year, activities)
            if st.session_state.show_center_preview:
                render_center_preview(unit, month, year, activities)
        else:
            preview_word = b""
            preview_pdf = b""
            if st.session_state.show_center_preview:
                st.warning("Completa correctamente las actividades iniciadas para previsualizar el reporte.")

        p1,p2 = st.columns(2)
        with p1:
            st.download_button("⬇️ Descargar borrador en Word", data=preview_word, file_name=f"Reporte_{unit}_{month}_{year}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", disabled=not preview_ready, use_container_width=True, on_click="ignore")
        with p2:
            st.download_button("⬇️ Descargar borrador en PDF", data=preview_pdf, file_name=f"Reporte_{unit}_{month}_{year}.pdf", mime="application/pdf", disabled=not preview_ready, use_container_width=True, on_click="ignore")

        st.divider()
        b1, b2 = st.columns(2)
        with b1:
            if st.button("💾 Guardar borrador", use_container_width=True):
                if errors:
                    for e in errors:
                        st.error(e)
                elif not activities:
                    st.warning("Captura al menos una actividad.")
                else:
                    rid = save_report(unit, month, year, "BORRADOR", sender_email)
                    replace_activities(rid, activities)
                    st.success(f"Borrador guardado para {month} {year}.")
        with b2:
            if st.button("📨 Enviar reporte", type="primary", use_container_width=True):
                submission_errors = list(errors)
                if not email_is_iteso(sender_email):
                    submission_errors.append("Captura un correo institucional válido @iteso.mx.")
                if submission_errors:
                    for e in dict.fromkeys(submission_errors): st.error(e)
                elif not activities:
                    st.warning("Captura al menos una actividad.")
                else:
                    confirm_submission_dialog(unit, month, year, sender_email, activities)

    else:
        st.header("Mis reportes")
        rows = [r for r in get_reports() if r["unit_code"] == unit]
        if not rows:
            st.info("Todavía no hay reportes guardados para este centro.")
        for r in rows:
            status = "✅ Enviado" if r["status"] == "ENVIADO" else "🟡 Borrador"
            with st.expander(f"{r['month']} {r['year']} · {status}"):
                acts = get_activities(r["id"])
                for a in acts:
                    st.markdown(f"**{a['title']}** · {rank_label(a.get('ranking'))}")
                    st.write(a.get("description_original", ""))
                    photos = get_activity_photos(a["id"])
                    if photos:
                        pc = st.columns(min(3, len(photos)))
                        for idx, ph in enumerate(photos):
                            with pc[idx % len(pc)]:
                                st.image(ph["bytes"], use_container_width=True)

else:
    if not admin_gate():
        st.stop()

    with st.sidebar:
        page = st.radio("Menú", ["Seguimiento mensual", "Informe consolidado", "Buscador histórico"])

    if page == "Seguimiento mensual":
        st.header("Seguimiento mensual")
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
            c1, c2, c3 = st.columns([2, 4, 2])
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

            for rep in reports:
                st.markdown(f"### {rep['unit_code']} · {UNITS[rep['unit_code']]}")
                for a in acts_by_report[rep["id"]]:
                    label = f"{rank_label(a.get('ranking'))} · {a['title']}"
                    edited = st.text_area(
                        label,
                        value=a.get("description_edited") or a.get("description_original") or "",
                        key=f"edit_{a['id']}",
                        height=125
                    )
                    if st.button("Guardar edición", key=f"save_{a['id']}"):
                        update_edited_description(a["id"], edited)
                        st.success("Edición guardada.")

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
                st.divider()

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

            word_bytes = generate_word(month, year, reports, acts_by_report)
            pdf_bytes = generate_pdf(month, year, reports, acts_by_report)

            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇️ Descargar Word",
                    data=word_bytes,
                    file_name=f"Informe_DIC_{month}_{year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with d2:
                st.download_button(
                    "⬇️ Descargar PDF",
                    data=pdf_bytes,
                    file_name=f"Informe_DIC_{month}_{year}.pdf",
                    mime="application/pdf",
                    use_container_width=True
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
st.caption("Prototipo V1 · Dirección de Integración Comunitaria · ITESO")
