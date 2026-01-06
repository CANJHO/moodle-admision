# app_streamlit_admision.py
# Interfaz Streamlit para tu exportador de Admisión (SIN BD / SIN MySQL)

import streamlit as st
from pathlib import Path
from io import BytesIO
import tempfile
import time
import json
import pandas as pd

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

    # (Quedan listos por si más adelante quieres usarlos; hoy no se pasan al core)
    nivel_por_area = {}
    for area_key, area_label in [
        ("A", "Área A – Ingenierías"),
        ("B", "Área B – Ciencias de la Salud"),
        ("C", "Área C – Ciencias Humanas"),
    ]:
        with st.expander(f"{area_label} ({area_key})", expanded=(area_key == "A")):
            com_niv = st.number_input(
                f"{area_key} - Umbral COMUNICACIÓN (%)",
                min_value=0.0, max_value=100.0,
                value=nivel_threshold_pct, step=1.0,
            )
            hab_niv = st.number_input(
                f"{area_key} - Umbral HABILIDADES COMUNICATIVAS (%)",
                min_value=0.0, max_value=100.0,
                value=nivel_threshold_pct, step=1.0,
            )
            mat_niv = st.number_input(
                f"{area_key} - Umbral MATEMÁTICA (%)",
                min_value=0.0, max_value=100.0,
                value=nivel_threshold_pct, step=1.0,
            )
            cta_niv = st.number_input(
                f"{area_key} - Umbral CTA / CCSS (%)",
                min_value=0.0, max_value=100.0,
                value=nivel_threshold_pct, step=1.0,
            )

        nivel_por_area[area_key] = {
            "COMUNICACIÓN": com_niv,
            "HABILIDADES COMUNICATIVAS": hab_niv,
            "MATEMÁTICA": mat_niv,
            "CTA/CCSS": cta_niv,
        }

# ---------------------------------------------------------------------
# CUERPO PRINCIPAL (GENERADOR)
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
    n = name.lower()
    if "ingenier" in n:
        return "A"
    if "salud" in n:
        return "B"
    if "humana" in n:
        return "C"
    return ""

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

        st.caption("Puedes editar el área sugerida (A/B/C) desde “Mapa quiz→Área”.")

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

    # Umbral global % → decimal
    nivel_threshold = nivel_threshold_pct / 100.0

    try:
        course_ids = [int(x) for x in course_ids_str.split(",") if x.strip()]
        t_from, t_to, tz = core.day_range_epoch(exam_date.isoformat(), tz_offset)

        st.info(f"Cursos: {course_ids} | Día: {exam_date} (tz {tz_offset})")
        st.info(f"Quiz→Área: {quiz_map}")

        with st.status("🔁 Descubriendo quizzes…", expanded=False) as status:
            quizzes = core.discover_quizzes(base_url, TOKEN, course_ids)
            qids_in_cursos = {q["quizid"] for q in quizzes}
            target_qids = [qid for qid in quiz_map.keys() if qid in qids_in_cursos]
            target_quizzes = [q for q in quizzes if q["quizid"] in target_qids]
            status.update(label=f"Quizzes a procesar: {len(target_quizzes)}", state="complete")

        course_users = {}
        total_users = 0
        prog_bar = st.progress(0, text="Cargando usuarios por curso…")
        for i, cid in enumerate(course_ids, start=1):
            us = core.get_course_users(
                base_url, TOKEN, cid,
                only_roles=[x.strip() for x in only_roles.split(",") if x.strip()],
            )
            course_users[cid] = us
            total_users += len(us)
            prog_bar.progress(i / len(course_ids), text=f"Curso {cid}: {len(us)} usuarios")
        prog_bar.empty()

        if total_users == 0 or not target_quizzes:
            st.warning("Nada para procesar (sin usuarios o sin quizzes objetivo).")
            st.stop()

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
                    futs.append(ex.submit(core._process_user_quiz, base_url, TOKEN, q, area_letter, u, t_from, t_to, tz))

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

        fname = f"RESULTADOS_ADMISION_{exam_date}.xlsx"
        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / fname
            core.write_excel_all_in_one(
                out_path,
                rows,
                nivel_threshold_base=nivel_threshold,  # <= 30% nivelación
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

# =====================================================================
# 📂 CONVERSOR A FORMATO BD
# =====================================================================
st.markdown("---")
st.header("📂 Conversor a formato BD")

tab1, tab2 = st.tabs(["✅ Desde Excel Moodle (RESULTADOS/RESUMEN)", "📤 Archivo de la comisión"])

# ==========================================================
# TAB 1: Desde Excel Moodle (RESULTADOS/RESUMEN)
# ==========================================================
with tab1:
    st.write(
        "Sube el Excel generado (con hojas **RESULTADOS** y **RESUMEN**) "
        "y lo convierto a la plantilla final para BD."
    )

    uploaded_file = st.file_uploader(
        "Sube el Excel con las hojas RESULTADOS y RESUMEN",
        type=["xlsx"],
        key="conv_excel_moodle",
    )

    c1, c2 = st.columns(2)
    with c1:
        periodo_value = st.text_input("Periodo", value="2026-1", key="periodo_moodle")
    with c2:
        fecha_registro_value = st.text_input(
            "Fecha de registro (AAAA-MM-DD hh:mm:ss)",
            value="2025-11-29 00:00:00",
            key="fecha_moodle",
        )

    convertir = st.button("🔄 Convertir a plantilla BD", key="btn_convertir_moodle")

    if convertir:
        if uploaded_file is None:
            st.error("Primero sube el archivo Excel generado (RESULTADOS + RESUMEN).")
            st.stop()

        try:
            xlsx = pd.ExcelFile(uploaded_file)
            hojas = xlsx.sheet_names
            if "RESULTADOS" not in hojas or "RESUMEN" not in hojas:
                st.error(
                    "❌ El archivo no contiene las hojas necesarias: 'RESULTADOS' y 'RESUMEN'. "
                    f"Hojas encontradas: {hojas}"
                )
                st.stop()

            df_resultados = pd.read_excel(xlsx, sheet_name="RESULTADOS")
            df_resumen = pd.read_excel(xlsx, sheet_name="RESUMEN")

            # DNI como texto
            df_resultados["Numero de DNI"] = df_resultados["Numero de DNI"].astype(str).str.strip()
            df_resumen["DNI"] = df_resumen["DNI"].astype(str).str.strip()

            # Cruce REAL para codigo_estudiante desde RESULTADOS ("Código de Matrícula")
            base_cols = ["Apellido(s)", "Nombre", "Numero de DNI"]
            if "Código de Matrícula" in df_resultados.columns:
                base_cols.append("Código de Matrícula")

            df_small = df_resultados[base_cols].copy()

            merged = df_resumen.merge(
                df_small,
                left_on="DNI",
                right_on="Numero de DNI",
                how="left",
            )

            if "Código de Matrícula" in merged.columns:
                codigo_estudiante = merged["Código de Matrícula"].astype(str).fillna("").str.strip()
            else:
                codigo_estudiante = pd.Series([""] * len(merged))

            # JSON de cursos nivelación
            course_cols = {
                "COMUNICACIÓN.1": "COMUNICACIÓN",
                "HABILIDADES COMUNICATIVAS.1": "HABILIDADES COMUNICATIVAS",
                "MATEMATICA": "MATEMATICA",
                "CIENCIA, TECNOLOGÍA Y AMBIENTE.1": "CIENCIA, TECNOLOGÍA Y AMBIENTE",
                "CIENCIAS SOCIALES": "CIENCIAS SOCIALES",
            }

            def build_json_courses(row):
                cursos = []
                for col, nombre in course_cols.items():
                    val = row.get(col)
                    if isinstance(val, str) and val.strip() != "":
                        cursos.append({"curso": nombre})
                return json.dumps(cursos, ensure_ascii=False)

            areas_nivelacion = merged.apply(build_json_courses, axis=1)

            # Requiere nivelación: acepta "SI" o "REQUIERE NIVELACIÓN"
            req = merged["PROGRAMA DE NIVELACIÓN"].fillna("").astype(str)
            requiere_nivelacion = req.apply(
                lambda x: "SI" if x.strip().upper() in ("REQUIERE NIVELACIÓN", "REQUIERE NIVELACION", "SI") else "NO"
            )

            out_df = pd.DataFrame({
                "id": None,
                "periodo": periodo_value,
                "codigo_estudiante": codigo_estudiante,
                "apellidos": merged["Apellido(s)"],
                "nombres": merged["Nombre"],
                "dni": merged["DNI"].astype(str),
                "area": merged["Área"],
                "programa": merged["Programa Académico"],
                "local_examen": merged["Sede o Filial"],
                "puntaje": pd.to_numeric(merged["TOTAL"], errors="coerce").fillna(0).astype(int),
                "asistio": merged["Asistencia"],
                "condicion": merged["CONDICIÓN"],
                "requiere_nivelacion": requiere_nivelacion,
                "areas_nivelacion": areas_nivelacion,
                "fecha_registro": fecha_registro_value,
                "estado": 1,
            })

            buffer = BytesIO()
            out_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.success("🎉 Archivo convertido correctamente (Moodle → BD).")
            st.download_button(
                label="⬇️ Descargar archivo para BD (postulantes_convertidos.xlsx)",
                data=buffer,
                file_name="postulantes_convertidos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.dataframe(out_df.head())

        except Exception as e:
            st.error(f"❌ Ocurrió un error durante la conversión: {e}")
            st.stop()

# ==========================================================
# TAB 2: Archivo de la comisión (cualquier nombre/hoja)
# ==========================================================
with tab2:
    st.write(
        "📤 **Subir archivo de la comisión (Cuadro de ingresantes / resultados / nivelación)**.\n\n"
        "- El archivo puede tener cualquier nombre.\n"
        "- La hoja puede tener cualquier nombre.\n"
        "- Se transformará al mismo formato BD (sin agregar columnas)."
    )

    com_file = st.file_uploader(
        "📤 Subir archivo de la comisión (Excel)",
        type=["xlsx"],
        key="comision_excel",
    )

    c1, c2 = st.columns(2)
    with c1:
        periodo_value_com = st.text_input("Periodo", value="2026-1", key="periodo_comision")
    with c2:
        fecha_registro_value_com = st.text_input(
            "Fecha de registro (AAAA-MM-DD hh:mm:ss)",
            value="2025-11-29 00:00:00",
            key="fecha_comision",
        )

    convertir_com = st.button("🔄 Convertir archivo de comisión → Plantilla BD", key="btn_convertir_comision")

    def _norm(s: str) -> str:
        return "".join(ch for ch in str(s).strip().lower() if ch.isalnum())

    def _find_col(df: pd.DataFrame, keywords):
        cols = list(df.columns)
        ncols = {c: _norm(c) for c in cols}
        for c, nc in ncols.items():
            if all(k in nc for k in keywords):
                return c
        return None

    if convertir_com:
        if com_file is None:
            st.error("Primero sube el Excel de la comisión.")
            st.stop()

        try:
            xlsx = pd.ExcelFile(com_file)
            if not xlsx.sheet_names:
                st.error("El archivo no contiene hojas.")
                st.stop()

            sheet = xlsx.sheet_names[0]
            df = pd.read_excel(xlsx, sheet_name=sheet)

            if df.empty:
                st.error("La hoja está vacía.")
                st.stop()

            col_ap = _find_col(df, ["apell"]) or _find_col(df, ["apellido"])
            col_nom = _find_col(df, ["nomb"])
            col_dni = _find_col(df, ["dni"])
            col_area = _find_col(df, ["area"])
            col_prog = _find_col(df, ["carrera"]) or _find_col(df, ["programa"])
            col_total = _find_col(df, ["total"]) or _find_col(df, ["puntaje"])
            col_asist = _find_col(df, ["asist"])
            col_cond = _find_col(df, ["condic"])
            col_prog_niv = _find_col(df, ["programa", "nivel"]) or _find_col(df, ["nivelacion"])

            col_cod = (
                _find_col(df, ["cod", "matr"]) or
                _find_col(df, ["codigo", "mat"]) or
                _find_col(df, ["matric"])
            )

            faltantes = []
            if not col_ap: faltantes.append("APELLIDOS")
            if not col_nom: faltantes.append("NOMBRES")
            if not col_dni: faltantes.append("DNI")
            if not col_area: faltantes.append("AREA")
            if not col_prog: faltantes.append("CARRERA/PROGRAMA")
            if not col_total: faltantes.append("TOTAL/PUNTAJE")

            if faltantes:
                st.error(f"No pude detectar estas columnas necesarias: {', '.join(faltantes)}")
                st.info(f"Columnas encontradas en la hoja '{sheet}': {list(df.columns)}")
                st.stop()

            dni = df[col_dni].astype(str).str.strip()
            apellidos = df[col_ap].astype(str).str.strip()
            nombres = df[col_nom].astype(str).str.strip()
            area = df[col_area].astype(str).str.strip()
            programa = df[col_prog].astype(str).str.strip()
            puntaje = pd.to_numeric(df[col_total], errors="coerce").fillna(0).astype(int)

            asistio = df[col_asist].astype(str).str.strip() if col_asist else "ASISTIÓ"
            condicion = df[col_cond].astype(str).str.strip() if col_cond else ""
            codigo_estudiante = df[col_cod].astype(str).fillna("").str.strip() if col_cod else ""

            if col_prog_niv:
                raw = df[col_prog_niv].fillna("").astype(str)
                requiere_nivelacion = raw.apply(
                    lambda x: "SI" if x.strip().upper() in ("SI", "REQUIERE NIVELACIÓN", "REQUIERE NIVELACION") else "NO"
                )
            else:
                requiere_nivelacion = pd.Series(["NO"] * len(df))

            # Intentar armar JSON de cursos si existen columnas por curso
            course_candidates = {
                "COMUNICACIÓN": ["comunic"],
                "HABILIDADES COMUNICATIVAS": ["habil"],
                "MATEMATICA": ["matemat"],
                "CIENCIA, TECNOLOGÍA Y AMBIENTE": ["ciencia", "tecn"],
                "CIENCIAS SOCIALES": ["ciencias", "social"],
            }

            detected_course_cols = {}
            for curso, keys in course_candidates.items():
                best = None
                for c in df.columns:
                    nc = _norm(c)
                    if all(_norm(k) in nc for k in keys):
                        best = c
                        break
                if best:
                    detected_course_cols[curso] = best

            def build_json_from_comision(row):
                cursos = []
                for curso, col in detected_course_cols.items():
                    v = row.get(col)
                    if isinstance(v, str) and v.strip() != "":
                        cursos.append({"curso": curso})
                    elif isinstance(v, (int, float)) and v != 0:
                        cursos.append({"curso": curso})
                return json.dumps(cursos, ensure_ascii=False)

            if detected_course_cols:
                areas_nivelacion = df.apply(build_json_from_comision, axis=1)
            else:
                areas_nivelacion = pd.Series([json.dumps([], ensure_ascii=False)] * len(df))

            col_sede = _find_col(df, ["sede"]) or _find_col(df, ["filial"]) or _find_col(df, ["local"])

            out_df = pd.DataFrame({
                "id": None,
                "periodo": periodo_value_com,
                "codigo_estudiante": codigo_estudiante,
                "apellidos": apellidos,
                "nombres": nombres,
                "dni": dni,
                "area": area,
                "programa": programa,
                "local_examen": df[col_sede].astype(str).str.strip() if col_sede else "",
                "puntaje": puntaje,
                "asistio": asistio,
                "condicion": condicion,
                "requiere_nivelacion": requiere_nivelacion,
                "areas_nivelacion": areas_nivelacion,
                "fecha_registro": fecha_registro_value_com,
                "estado": 1,
            })

            buffer = BytesIO()
            out_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.success("🎉 Archivo de comisión convertido correctamente → Plantilla BD.")
            st.download_button(
                label="⬇️ Descargar archivo para BD (postulantes_convertidos.xlsx)",
                data=buffer,
                file_name="postulantes_convertidos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.dataframe(out_df.head())

        except Exception as e:
            st.error(f"❌ Error convirtiendo archivo de comisión: {e}")
            st.stop()



