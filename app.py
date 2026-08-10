
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
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.colors import HexColor

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

ITESO_BLUE = "#004C7F"
LIGHT_BLUE = "#EAF2F7"
TEXT_GRAY = "#555555"

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
.block-container {{
    padding-top: 1.2rem;
    max-width: 1380px;
}}
h1, h2, h3 {{
    color: {ITESO_BLUE};
}}
[data-testid="stSidebar"] {{
    background: #F4F7FA;
    border-right: 1px solid #D9E2EC;
}}
[data-testid="stSidebar"] * {{
    color: #12344D !important;
}}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stRadio label,
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] p {{
    color: #12344D !important;
}}
.dic-card {{
    background: white;
    border: 1px solid #d8dee5;
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 12px;
}}
.small-muted {{
    color:#6b7280; font-size:0.9rem;
}}
.header-wrap {{
    display:flex;
    align-items:center;
    gap:1rem;
}}
.header-title {{
    padding-top: 0.25rem;
}}
.header-subtitle {{
    color:#5B6B7A;
    margin-top: -8px;
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

    doc.build(story)
    bio.seek(0)
    return bio.getvalue()


def generate_segment_word(rep, act, photos):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)

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

    doc.build(story)
    return bio.getvalue()


# ---------- HEADER ----------
c1, c2 = st.columns([1.25, 3.75], gap="large")
with c1:
    st.image("assets/iteso_logo.png", width=240)
with c2:
    st.markdown("""
    <div class="header-title">
        <h1 style="margin-bottom:0.2rem;">Sistema de informes mensuales</h1>
        <div class="header-subtitle">Dirección de Integración Comunitaria · ITESO</div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

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
        page = st.radio("Menú", ["Nuevo reporte", "Mis reportes"])

    if page == "Nuevo reporte":
        st.header("Nuevo reporte mensual")
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
            with st.expander(f"Actividad {i+1}", expanded=(i < 2)):
                title = st.text_input("Nombre del hito / actividad", key=f"title_{i}")
                c1, c2, c3 = st.columns([2, 1, 1])
                with c1:
                    category_selected = st.selectbox("Tema / categoría", CATEGORIES, key=f"cat_{i}")
                    other_category = ""
                    if category_selected == "Otro":
                        other_category = st.text_input(
                            "Especifica la categoría",
                            key=f"other_cat_{i}",
                            placeholder="Escribe la categoría"
                        )
                    category = other_category.strip() if category_selected == "Otro" and other_category.strip() else category_selected
                with c2:
                    ranking_text = st.selectbox(
                        "Importancia",
                        ["Sin ranking", "Top 1", "Top 2", "Top 3"],
                        key=f"rank_{i}"
                    )
                with c3:
                    participants = st.number_input(
                        "Participantes / alcance",
                        min_value=0,
                        step=1,
                        value=0,
                        key=f"part_{i}",
                        help="Campo obligatorio para una actividad capturada. Si no aplica, registra 1 y acláralo en la descripción."
                    )

                desc = st.text_area(
                    "Descripción del hito",
                    height=150,
                    placeholder="Qué ocurrió, por qué fue relevante, resultados y actores participantes.",
                    key=f"desc_{i}"
                )
                wc = word_count(desc)
                if wc > 250:
                    st.error(f"{wc}/250 palabras. Reduce la descripción en {wc-250} palabras.")
                else:
                    st.caption(f"{wc}/250 palabras")

                photos = st.file_uploader(
                    "Fotografías (opcional)",
                    type=["jpg", "jpeg", "png"],
                    accept_multiple_files=True,
                    key=f"photos_{i}",
                    help="En V1 se habilita la captura; el almacenamiento permanente se activa al conectar Supabase Storage."
                )

                activity_started = bool(
                    title.strip()
                    or desc.strip()
                    or participants > 0
                    or ranking_text != "Sin ranking"
                    or category_selected != CATEGORIES[0]
                    or photos
                )

                if activity_started:
                    ranking = None if ranking_text == "Sin ranking" else int(ranking_text[-1])
                    activities.append({
                        "title": title.strip(),
                        "description": desc.strip(),
                        "category": category,
                        "ranking": ranking,
                        "participants": participants or None,
                        "photos": photos or [],
                    })

                    missing = []
                    if not title.strip():
                        missing.append("nombre del hito / actividad")
                    if not desc.strip():
                        missing.append("descripción")
                    if participants <= 0:
                        missing.append("participantes / alcance")
                    if category_selected == "Otro" and not other_category.strip():
                        missing.append("categoría específica")

                    if missing:
                        errors.append(
                            f"Actividad {i+1}: completa " + ", ".join(missing) + "."
                        )
                    if wc > 250:
                        errors.append(f"Actividad {i+1}: excede 250 palabras.")

                    if missing:
                        st.warning(
                            "Esta actividad está incompleta. Antes de continuar, completa: "
                            + ", ".join(missing)
                            + ". La fotografía es opcional."
                        )

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
                if not email_is_iteso(sender_email):
                    errors.append("Captura un correo institucional válido @iteso.mx.")
                if errors:
                    for e in dict.fromkeys(errors):
                        st.error(e)
                elif not activities:
                    st.warning("Captura al menos una actividad.")
                else:
                    rid = save_report(unit, month, year, "ENVIADO", sender_email)
                    replace_activities(rid, activities)
                    email_result = send_confirmation_email(sender_email, unit, month, year)
                    st.success(f"Reporte de {month} {year} enviado correctamente.")
                    st.info(email_result["message"])

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
                c3.error("Pendiente")
            elif rep["status"] == "ENVIADO":
                c3.success("Enviado")
            else:
                c3.warning("Borrador")

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
