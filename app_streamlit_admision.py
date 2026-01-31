# app_streamlit_admision.py
# Interfaz Streamlit para tu exportador de Admisión (SIN BD / SIN MySQL)

import streamlit as st
from pathlib import Path
from io import BytesIO
import tempfile
import time
import json
import pandas as pd
import unicodedata
from datetime import datetime
import zipfile  # ✅ NUEVO: validar .xlsx (zip interno)

# Importamos tu lógica existente desde el script CLI
import moodle_admision_export as core

# ✅ NUEVO: Actas Finales (plantilla)
from actas_presentacion import build_excel_final_con_actas


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

# =====================================================================
# ✅ HELPERS (NUEVOS) - NO ROMPEN NADA, SOLO AYUDAN A DETECTAR COLUMNAS Y DNIs
# =====================================================================

def _norm_text(s: str) -> str:
    """Normaliza texto: minus, sin tildes, solo alfanumérico."""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return "".join(ch for ch in s if ch.isalnum())

def _find_col_flexible(df: pd.DataFrame, keyword_groups):
    """
    Busca una columna por grupos de keywords.
    keyword_groups: lista de listas. Retorna la primera columna que matchee algún grupo.
    Ej:
      [["codigo","matricula"], ["codigo","estudiante"], ["cod","matr"]]
    """
    cols = list(df.columns)
    norm_cols = {c: _norm_text(c) for c in cols}

    for group in keyword_groups:
        g = [_norm_text(x) for x in group]
        for c, nc in norm_cols.items():
            if all(k in nc for k in g):
                return c
    return None

def _norm_dni_value(v) -> str:
    """
    Normaliza DNI:
    - convierte a string
    - elimina '.0'
    - deja solo dígitos
    - rellena con 0 a la izquierda a 8 dígitos (clave para DNIs como 07489547)
    """
    s = "" if pd.isna(v) else str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits == "":
        return ""
    if len(digits) < 8:
        digits = digits.zfill(8)
    return digits

def _norm_dni_series(ser: pd.Series) -> pd.Series:
    return ser.apply(_norm_dni_value)

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

    # ==========================================================
    # ✅ FIX: plantilla en la MISMA carpeta (no assets)
    # ==========================================================
    BASE_DIR = Path(__file__).resolve().parent

    # Usa el modelo que tienes en la raíz del proyecto (según tu captura)
    modelo_path = BASE_DIR / "MODELO_RESULTADOS_EXAMEN.xlsx"

    # Si quisieras usar la otra, cambia a:
    # modelo_path = BASE_DIR / "PLANTILLA_DESCARGA_MOODLE_ADMISION.xlsx"

    if not modelo_path.exists():
        st.error(
            "❌ No encuentro la plantilla para Actas.\n\n"
            "Coloca el archivo en la misma carpeta del app_streamlit_admision.py:\n"
            f"- {modelo_path.as_posix()}"
        )
        st.stop()

    # ✅ Validación: un .xlsx real es un ZIP interno
    try:
        if not zipfile.is_zipfile(modelo_path):
            st.error(
                "❌ La plantilla NO es un .xlsx válido (no es ZIP interno).\n\n"
                f"Archivo: {modelo_path.name}\n\n"
                "✅ Solución:\n"
                "1) Ábrelo en Excel\n"
                "2) Guardar como → .xlsx\n"
                "3) Reemplaza el archivo y vuelve a intentar"
            )
            st.stop()
    except Exception as e:
        st.error(f"❌ Error validando la plantilla: {e}")
        st.stop()

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

        # ==========================================================
        # ✅ GENERAR EXCEL BASE + EXCEL FINAL (CON ACTAS) EN EL MISMO BOTÓN
        # ==========================================================
        fname_base = f"RESULTADOS_ADMISION_{exam_date}.xlsx"
        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / fname_base

            # 1) Excel base (RESULTADOS + RESUMEN)
            core.write_excel_all_in_one(
                out_path,
                rows,
                nivel_threshold_base=nivel_threshold,  # se mantiene igual para no romper tu core
            )
            base_bytes = out_path.read_bytes()

            # ✅ Validación: el excel base generado también debe ser ZIP interno
            if not zipfile.is_zipfile(BytesIO(base_bytes)):
                st.error("❌ El Excel base generado NO es un .xlsx válido (ZIP interno). Revisa openpyxl/pandas.")
                st.stop()

            # 2) Excel FINAL con actas dentro
            final_bytes = build_excel_final_con_actas(
                modelo_path=str(modelo_path),
                generated_excel_bytes=base_bytes,
                exam_date=datetime.combine(exam_date, datetime.min.time()),
                exam_label="EXAMEN ORDINARIO",
                output_add_resultados_resumen=True,
            )

        fname_final = f"ACTA_FINAL_Y_RESUMEN_{exam_date}.xlsx"

        st.download_button(
            label="⬇️ Descargar Excel (RESULTADOS + RESUMEN + ACTAS FINALES)",
            data=final_bytes,
            file_name=fname_final,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga el archivo final (incluye Actas Finales y Consolidados)",
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

            # Detectar columna DNI en RESULTADOS
            col_dni_res = "Numero de DNI" if "Numero de DNI" in df_resultados.columns else _find_col_flexible(
                df_resultados, [
                    ["numero", "dni"],
                    ["dni"],
                    ["documento", "dni"],
                    ["nro", "dni"],
                ]
            )

            # Detectar columna DNI en RESUMEN
            col_dni_sum = "DNI" if "DNI" in df_resumen.columns else _find_col_flexible(
                df_resumen, [
                    ["dni"],
                    ["numero", "dni"],
                    ["nro", "dni"],
                ]
            )

            if not col_dni_res or not col_dni_sum:
                st.error("No pude detectar la columna DNI en RESULTADOS o RESUMEN.")
                st.info(f"Columnas RESULTADOS: {list(df_resultados.columns)}")
                st.info(f"Columnas RESUMEN: {list(df_resumen.columns)}")
                st.stop()

            # Detectar columna de código: "Código de Matrícula" o "Código de Estudiante"
            col_cod = None
            for exact in [
                "Código de Matrícula", "Codigo de Matricula", "CÓDIGO DE MATRÍCULA", "CODIGO DE MATRICULA",
                "Código de Estudiante", "Codigo de Estudiante", "CÓDIGO DE ESTUDIANTE", "CODIGO DE ESTUDIANTE"
            ]:
                if exact in df_resultados.columns:
                    col_cod = exact
                    break

            if not col_cod:
                col_cod = _find_col_flexible(df_resultados, [
                    ["codigo", "matricula"],
                    ["codigo", "estudiante"],
                    ["cod", "matr"],
                    ["cod", "estud"],
                    ["codigo"],
                ])

            if not col_cod:
                st.warning("No encontré columna de CÓDIGO (MATRÍCULA/ESTUDIANTE) en RESULTADOS. Saldrá vacío.")
                st.info(f"Columnas RESULTADOS: {list(df_resultados.columns)}")

            # Normalizar DNI en ambos para merge
            df_resultados["_dni_norm"] = _norm_dni_series(df_resultados[col_dni_res])
            df_resumen["_dni_norm"] = _norm_dni_series(df_resumen[col_dni_sum])

            cols_small = ["_dni_norm"]
            if "Apellido(s)" in df_resultados.columns: cols_small.append("Apellido(s)")
            if "Nombre" in df_resultados.columns: cols_small.append("Nombre")
            if col_cod: cols_small.append(col_cod)

            df_small = df_resultados[cols_small].copy()

            merged = df_resumen.merge(
                df_small,
                on="_dni_norm",
                how="left",
            )

            if col_cod and col_cod in merged.columns:
                codigo_estudiante = merged[col_cod].astype(str).fillna("").str.strip().replace("nan", "")
            else:
                codigo_estudiante = pd.Series([""] * len(merged))

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

            req = merged["PROGRAMA DE NIVELACIÓN"].fillna("").astype(str) if "PROGRAMA DE NIVELACIÓN" in merged.columns else pd.Series([""] * len(merged))
            requiere_nivelacion = req.apply(
                lambda x: "SI" if str(x).strip().upper() in ("REQUIERE NIVELACIÓN", "REQUIERE NIVELACION", "SI") else "NO"
            )

            out_df = pd.DataFrame({
                "id": None,
                "periodo": periodo_value,
                "codigo_estudiante": codigo_estudiante,
                "apellidos": merged["Apellido(s)"] if "Apellido(s)" in merged.columns else "",
                "nombres": merged["Nombre"] if "Nombre" in merged.columns else "",
                "dni": merged[col_dni_sum].apply(_norm_dni_value),
                "area": merged["Área"] if "Área" in merged.columns else "",
                "programa": merged["Programa Académico"] if "Programa Académico" in merged.columns else "",
                "local_examen": merged["Sede o Filial"] if "Sede o Filial" in merged.columns else "",
                "puntaje": pd.to_numeric(merged["TOTAL"], errors="coerce").fillna(0).astype(int) if "TOTAL" in merged.columns else 0,
                "asistio": merged["Asistencia"] if "Asistencia" in merged.columns else "",
                "condicion": merged["CONDICIÓN"] if "CONDICIÓN" in merged.columns else "",
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

            filled = (out_df["codigo_estudiante"].astype(str).str.strip() != "").sum()
            st.info(f"Códigos de estudiante/matrícula encontrados: {filled} / {len(out_df)}")
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
        s = str(s).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        return "".join(ch for ch in s if ch.isalnum())

    def _find_col(df: pd.DataFrame, keywords):
        cols = list(df.columns)
        ncols = {c: _norm(c) for c in cols}
        k_norm = [_norm(k) for k in keywords]
        for c, nc in ncols.items():
            if all(k in nc for k in k_norm):
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
                _find_col(df, ["codigo", "estudiante"]) or
                _find_col(df, ["cod", "estudiante"]) or
                _find_col(df, ["codigo", "matricula"]) or
                _find_col(df, ["cod", "matr"]) or
                _find_col(df, ["codigo", "mat"]) or
                _find_col(df, ["matric"]) or
                _find_col(df, ["matricula"]) or
                _find_col(df, ["codigo"])
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

            st.info(f"Columna detectada para codigo_estudiante: {col_cod}")

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
