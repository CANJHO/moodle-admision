# app_streamlit_admision.py
# Interfaz Streamlit para tu exportador de Admisión
import streamlit as st
from pathlib import Path
from datetime import datetime
from io import BytesIO
import tempfile
import time

# Importamos tu lógica existente desde el script CLI
import moodle_admision_export as core

st.set_page_config(
    page_title="Admisión Moodle - Exportador",
    page_icon="📤",
    layout="wide"
)

# --- Encabezado ---
st.title("📤 Exportador de Admisión (Moodle)")
st.caption("Genera el Excel (RESULTADOS + RESUMEN) en base a Fecha, Curso(s) y Mapa Quiz→Área.")

# --- Secrets (token/base_url) ---
try:
    TOKEN = st.secrets["TOKEN"]
    BASE_URL = st.secrets["BASE_URL"]
except Exception:
    st.error("No se encontraron los *Secrets*. Ve a Settings → Secrets y define TOKEN y BASE_URL.")
    st.stop()

# ---------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------
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

    # Umbral de nivelación en porcentaje (global).
    # Se compara contra el % obtenido en cada curso (0–100).
    nivel_threshold_pct = st.number_input(
        "Umbral de nivelación (%)",
        min_value=0.0,
        max_value=100.0,
        value=30.0,
        step=1.0,
        help="Si el porcentaje obtenido en un curso es menor o igual a este valor, "
             "el postulante requiere nivelación en ese curso.",
    )

    st.markdown("---")
    st.subheader("📊 Umbrales de nivelación por área y curso")

    # Estos controles son para ver/ajustar el umbral por curso y por área.
    # De momento todos parten en 30% (igual al global) y se pueden ajustar si lo necesitas.
    nivel_por_area = {}
    for area_key, area_label in [
        ("A", "Área A – Ingenierías"),
        ("B", "Área B – Ciencias de la Salud"),
        ("C", "Área C – Ciencias Humanas"),
    ]:
        with st.expander(f"{area_label} ({area_key})", expanded=(area_key == "A")):
            com_niv = st.number_input(
                f"{area_key} - Umbral COMUNICACIÓN (%)",
                min_value=0.0,
                max_value=100.0,
                value=nivel_threshold_pct,
                step=1.0,
            )
            hab_niv = st.number_input(
                f"{area_key} - Umbral HABILIDADES COMUNICATIVAS (%)",
                min_value=0.0,
                max_value=100.0,
                value=nivel_threshold_pct,
                step=1.0,
            )
            mat_niv = st.number_input(
                f"{area_key} - Umbral MATEMÁTICA (%)",
                min_value=0.0,
                max_value=100.0,
                value=nivel_threshold_pct,
                step=1.0,
            )
            cta_niv = st.number_input(
                f"{area_key} - Umbral CTA / CCSS (%)",
                min_value=0.0,
                max_value=100.0,
                value=nivel_threshold_pct,
                step=1.0,
            )

        # Guardamos por si más adelante quieres usar estos umbrales diferenciados
        nivel_por_area[area_key] = {
            "COMUNICACIÓN": com_niv,
            "HABILIDADES COMUNICATIVAS": hab_niv,
            "MATEMÁTICA": mat_niv,
            "CTA/CCSS": cta_niv,
        }

# ---------------------------------------------------------------------
# CUERPO PRINCIPAL
# ---------------------------------------------------------------------
col1, col2 = st.columns([1, 1])
with col1:
    exam_date = st.date_input(
        "📅 Día del examen (hora local)",
        help="Se filtra 00:00–23:59:59 según el TZ",
    )
with col2:
    course_ids_str = st.text_input(
        "🎓 ID(s) de curso (coma)",
        placeholder="Ej.: 11989 o 100,101",
    )

quiz_map_str = st.text_input(
    "🧭 Mapa quiz→Área (A/B/C)",
    key="quiz_map_str",
    placeholder="Ej.: 11907=A,11908=B,11909=C",
    help="Puedes obtener los IDs desde Moodle o autollenarlo con 'Descubrir quizzes'.",
)

# ---------------------------------------------------------------------
# Descubrir quizzes
# ---------------------------------------------------------------------
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
            if area_guess:
                st.write(
                    f"- **{q['quizname']}** — ID: `{q['quizid']}`  "
                    f"(curso {q['courseid']}) → área sugerida: **{area_guess}**"
                )
                sugerencias.append(f"{q['quizid']}={area_guess}")
            else:
                st.write(
                    f"- **{q['quizname']}** — ID: `{q['quizid']}`  "
                    f"(curso {q['courseid']})"
                )

        st.caption(
            "Puedes editar el área sugerida (A/B/C) desde el cuadro de texto "
            "“Mapa quiz→Área”."
        )

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

# ---------------------------------------------------------------------
# BOTÓN PRINCIPAL
# ---------------------------------------------------------------------
run = st.button("🚀 Generar Excel (RESULTADOS + RESUMEN)", type="primary")

if run:
    # Validaciones básicas
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

    # Convertimos el umbral global de % a decimal (0.30)
    nivel_threshold = nivel_threshold_pct / 100.0

    try:
        # Parseo de entradas
        course_ids = [int(x) for x in course_ids_str.split(",") if x.strip()]
        t_from, t_to, tz = core.day_range_epoch(exam_date.isoformat(), tz_offset)

        # Info inicial
        st.info(f"Cursos: {course_ids} | Día: {exam_date} (tz {tz_offset})")
        st.info(f"Quiz→Área: {quiz_map}")

        # Descubrir quizzes y quedarnos solo con los del mapa
        with st.status("🔁 Descubriendo quizzes…", expanded=False) as status:
            quizzes = core.discover_quizzes(base_url, TOKEN, course_ids)
            qids_in_cursos = {q["quizid"] for q in quizzes}
            target_qids = [qid for qid in quiz_map.keys() if qid in qids_in_cursos]
            target_quizzes = [q for q in quizzes if q["quizid"] in target_qids]
            status.update(
                label=f"Quizzes a procesar: {len(target_quizzes)}",
                state="complete",
            )

        # Usuarios por curso
        course_users = {}
        total_users = 0
        prog_bar = st.progress(0, text="Cargando usuarios por curso…")
        for i, cid in enumerate(course_ids, start=1):
            us = core.get_course_users(
                base_url,
                TOKEN,
                cid,
                only_roles=[x.strip() for x in only_roles.split(",") if x.strip()],
            )
            course_users[cid] = us
            total_users += len(us)
            prog_bar.progress(i / len(course_ids), text=f"Curso {cid}: {len(us)} usuarios")
        prog_bar.empty()

        if total_users == 0 or not target_quizzes:
            st.warning("Nada para procesar (sin usuarios o sin quizzes objetivo).")
            st.stop()

        # Procesar intentos
        st.write("⚙️ Procesando intentos (esto puede tardar)…")
        t0 = time.time()
        rows = []
        from concurrent.futures import ThreadPoolExecutor, as_completed
        futs = []
        with ThreadPoolExecutor(max_workers=workers) as ex:
            for q in target_quizzes:
                area_letter = quiz_map.get(q["quizid"])
                users = course_users.get(q["courseid"], [])
                for u in users:
                    futs.append(
                        ex.submit(
                            core._process_user_quiz,
                            base_url,
                            TOKEN,
                            q,
                            area_letter,
                            u,
                            t_from,
                            t_to,
                            tz,
                        )
                    )
            done = 0
            step_bar = st.progress(0.0)
            for fut in as_completed(futs):
                res = fut.result()
                if res:
                    rows.extend(res)
                done += 1
                step_bar.progress(done / max(1, len(futs)))
        step_bar.empty()

        st.success(f"Intentos dentro del día: {len(rows)}")
        if not rows:
            st.warning("No se encontraron intentos ese día.")
            st.stop()

        # Generar Excel en memoria y ofrecer descarga
        fname = f"RESULTADOS_ADMISION_{exam_date}.xlsx"
        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / fname
            core.write_excel_all_in_one(
                out_path,
                rows,
                nivel_threshold_base=nivel_threshold,  # usa la lógica ≤ 30% para nivelación
            )
            data = out_path.read_bytes()

        st.download_button(
            label="⬇️ Descargar Excel (RESULTADOS + RESUMEN)",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga el archivo generado",
        )

        st.caption(f"Tiempo total: {time.time() - t0:.1f} s")

    except Exception as e:
        st.error(f"❌ Ocurrió un error: {e}")
