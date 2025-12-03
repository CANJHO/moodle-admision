# app_streamlit_admision.py
# Interfaz Streamlit para tu exportador de Admisión
import streamlit as st
from pathlib import Path
from datetime import datetime
from io import BytesIO
import tempfile
import time

# Importamos tu lógica existente desde el script CLI
# (debe estar en el mismo repo y con ese nombre)
import moodle_admision_export as core

st.set_page_config(page_title="Admisión Moodle - Exportador", page_icon="📤", layout="wide")

# --- Encabezado ---
st.title("📤 Exportador de Admisión (Moodle)")
st.caption("Genera el Excel (RESULTADOS + RESUMEN) en base a Fecha, Curso(s) y Mapa Quiz→Área.")

# --- Secrets (token/base_url) ---
# Los tomamos de Streamlit Cloud (Settings → Secrets)
try:
    TOKEN = st.secrets["TOKEN"]
    BASE_URL = st.secrets["BASE_URL"]
except Exception:
    st.error("No se encontraron los *Secrets*. Ve a Settings → Secrets y define TOKEN y BASE_URL.")
    st.stop()

with st.sidebar:
    st.subheader("⚙️ Parámetros generales")
    # Botón/link rojo (texto blanco) a otra app
    st.markdown(
        """
        <a href="https://asignadorzoom-gqujexxocuamxss77jq7wy.streamlit.app/"
           target="_blank"
           style="
             display:block;
             text-align:center;
             padding:0.60rem 0.8rem;
             background:#d32f2f;
             color:#ffffff;
             border-radius:8px;
             text-decoration:none;
             font-weight:700;
             margin-bottom:0.75rem;
           ">
           ASIGNADOR UAI
        </a>
        """,
        unsafe_allow_html=True,
    )
    base_url = st.text_input(
        "Base URL de Moodle",
        value=BASE_URL,
        help="Ej.: https://aulavirtual.autonomadeica.edu.pe",
    )
    tz_offset = st.text_input("TZ offset local", value="-05:00", help="Ej.: -05:00")
    workers = st.slider("Hilos paralelos", min_value=4, max_value=32, value=16, step=1)
    only_roles = st.text_input(
        "Roles a incluir",
        value="student",
        help="Ej.: student (múltiples separados por coma)",
    )

    st.markdown("---")
    st.subheader("🧮 Nivelación")

    # Umbral base de nivelación en porcentaje (verde)
    nivel_threshold_pct = st.number_input(
        "Umbral de nivelación (%)",
        min_value=0.0,
        max_value=100.0,
        value=30.0,
        step=1.0,
        help="Si el porcentaje obtenido en un curso es menor o igual a este valor, el postulante requiere nivelación en ese curso.",
    )

    st.markdown("---")
    st.subheader("📊 Umbrales por área y curso")

    # Todos los cursos empiezan con el mismo % de nivelación (30 por defecto)
    nivel_por_area_pct: Dict[str, Dict[str, float]] = {} # type: ignore

    for area_key, area_label in [
        ("A", "Área A – Ingenierías"),
        ("B", "Área B – Ciencias de la Salud"),
        ("C", "Área C – Ciencias Humanas"),
    ]:
        with st.expander(f"{area_label} ({area_key})", expanded=(area_key == "A")):
            # valor base sugerido: el del input general
            base_val = nivel_threshold_pct

            com = st.number_input(
                f"{area_key} - Umbral COMUNICACIÓN (%)",
                min_value=0.0,
                max_value=100.0,
                value=base_val,
                step=1.0,
            )
            hab = st.number_input(
                f"{area_key} - Umbral HABILIDADES COMUNICATIVAS (%)",
                min_value=0.0,
                max_value=100.0,
                value=base_val,
                step=1.0,
            )
            mat = st.number_input(
                f"{area_key} - Umbral MATEMÁTICA (%)",
                min_value=0.0,
                max_value=100.0,
                value=base_val,
                step=1.0,
            )

            # CTA o CCSS según el área, pero internamente se sigue llamando CTA/CCSS
            if area_key == "C":
                label_cta = "CIENCIAS SOCIALES"
            else:
                label_cta = "CTA (CIENCIA, TECNOLOGÍA Y AMBIENTE)"

            cta = st.number_input(
                f"{area_key} - Umbral {label_cta} (%)",
                min_value=0.0,
                max_value=100.0,
                value=base_val,
                step=1.0,
            )

        nivel_por_area_pct[area_key] = {
            "COMUNICACIÓN": com,
            "HABILIDADES COMUNICATIVAS": hab,
            "MATEMÁTICA": mat,
            "CTA/CCSS": cta,  # clave interna, aunque en C sea CCSS
        }

    

col1, col2 = st.columns([1,1])
with col1:
    exam_date = st.date_input("📅 Día del examen (hora local)", help="Se filtra 00:00–23:59:59 según el TZ")
with col2:
    course_ids_str = st.text_input("🎓 ID(s) de curso (coma)",
                                   placeholder="Ej.: 11989 o 100,101")

quiz_map_str = st.text_input(
    "🧭 Mapa quiz→Área (A/B/C)",
    key="quiz_map_str",
    placeholder="Ej.: 11907=A,11908=B,11909=C",
    help="Puedes obtener los IDs desde Moodle o autollenarlo con 'Descubrir quizzes'."
)


# --- Utilidad: descubrir quizzes por cursos ---
# --- Utilidad: descubrir quizzes por cursos ---
def _guess_area_from_name(name: str) -> str:
    """Intenta deducir el área A/B/C a partir del nombre del quiz."""
    n = name.lower()
    if "ingenier" in n:
        return "A"          # Examen de Admisión – Ingenierías
    if "salud" in n:
        return "B"          # Ciencias de la Salud
    if "humana" in n:
        return "C"          # Ciencias Humanas
    return ""               # sin área detectada, se edita a mano

def discover_quizzes_ui():
    if not course_ids_str.strip():
        st.warning("Primero ingresa los ID(s) de curso.")
        return
    try:
        course_ids = [int(x) for x in course_ids_str.split(",") if x.strip()]
        quizzes = core.discover_quizzes(base_url, TOKEN, course_ids)
        if not quizzes:
            st.info("No se encontraron quizzes en esos cursos.")
            return

        st.success(f"Quizzes encontrados ({len(quizzes)}):")

        sugerencias = []
        for q in quizzes:
            area_guess = _guess_area_from_name(q["quizname"])
            # Mostrar el quiz con la sugerencia de área (si existe)
            if area_guess:
                st.write(
                    f"- **{q['quizname']}** — ID: `{q['quizid']}`  (curso {q['courseid']}) → área sugerida: **{area_guess}**"
                )
                sugerencias.append(f"{q['quizid']}={area_guess}")
            else:
                st.write(
                    f"- **{q['quizname']}** — ID: `{q['quizid']}`  (curso {q['courseid']})"
                )

        st.caption("Puedes editar el área sugerida (A/B/C) desde el cuadro de texto Mapa quiz→Área.")

        # Autollenar el input si hay sugerencias
        if sugerencias:
            st.session_state["quiz_map_str"] = ",".join(sugerencias)
            st.info("Se autocompletó el mapa quiz→Área. Revísalo y ajusta si es necesario.")
        else:
            st.info("No se pudo inferir áreas automáticamente. Completa el mapa a mano (A/B/C).")

    except Exception as e:
        st.error(f"Error al descubrir quizzes: {e}")

st.button("🔎 Descubrir quizzes en los cursos", on_click=discover_quizzes_ui)

st.markdown("---")


# --- Botón principal ---
# --- Botón principal ---
run = st.button("🚀 Generar Excel (RESULTADOS + RESUMEN)", type="primary")

if run:
    # Validaciones básicas (deja aquí tus checks de fecha, cursos, mapa, etc.)
    if not exam_date:
        st.error("Debes elegir la **Fecha** del examen.")
        st.stop()
    if not course_ids_str.strip():
        st.error("Debes ingresar al menos un **ID de curso**.")
        st.stop()
    quiz_map = core.parse_quiz_map(quiz_map_str)
    if not quiz_map:
        st.error("Debes ingresar un **Mapa quiz→Área** válido (ej. 11907=A,11908=B).")
        st.stop()

    # Convertimos el umbral general y los de área/curso a decimales (0.30, etc.)
    nivel_threshold_base = nivel_threshold_pct / 100.0
    nivel_por_area = {
        area: {sub: val / 100.0 for sub, val in subdict.items()}
        for area, subdict in nivel_por_area_pct.items()
    }

    try:
        # Recopila course_ids y demás; reemplaza esta sección con la lógica real.
        course_ids = [int(x) for x in course_ids_str.split(",") if x.strip()]
        # Intentar obtener quizzes si la función existe; en caso de fallo seguir con lista vacía.
        try:
            quizzes = core.discover_quizzes(base_url, TOKEN, course_ids)
        except Exception:
            quizzes = []

        # TODO: Construye 'rows' con la estructura que espera write_excel_all_in_one.
        # Actualmente se define una lista vacía para evitar NameError; reemplaza con tu código
        # que genere las filas (por ejemplo recopilando usuarios, resultados de quizzes, etc.).
        rows = []

        fname = f"RESULTADOS_ADMISION_{exam_date}.xlsx"
        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / fname
            core.write_excel_all_in_one(
                out_path,
                rows,
                nivel_threshold_base=nivel_threshold_base,
                nivel_by_area=nivel_por_area,
            )
            data = out_path.read_bytes()

        st.download_button(
            label="⬇️ Descargar Excel (RESULTADOS + RESUMEN)",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"❌ Ocurrió un error: {e}")




