# app_streamlit_admision.py
# Interfaz Streamlit para tu exportador de Admisión (SIN BD / SIN MySQL)

import streamlit as st
from pathlib import Path
from io import BytesIO
import tempfile
import time
import json
import importlib
import pandas as pd
import unicodedata
from datetime import datetime
import zipfile  # ✅ validar .xlsx (zip interno)

# Importamos tu lógica existente desde el script CLI
import moodle_admision_export as core
core = importlib.reload(core)

# ✅ Actas Finales (plantilla)
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
# HELPERS GENERALES
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
    - rellena con 0 a la izquierda a 8 dígitos
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


def _clean_text(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.lower() == "nan":
        return ""
    return s


def _clean_upper_text(v) -> str:
    return _clean_text(v).upper()


def _read_padron_flexible(uploaded_file) -> pd.DataFrame:
    xlsx = pd.ExcelFile(uploaded_file)
    sheet_name = next((sh for sh in xlsx.sheet_names if _norm_text(sh) == "data"), xlsx.sheet_names[0])
    raw = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)

    header_row = None
    for idx in range(min(15, len(raw))):
        values = [_norm_text(v) for v in raw.iloc[idx].tolist() if _clean_text(v)]
        has_document = any(v in ("dni", "documento", "numerodni", "nrodocumento") for v in values)
        has_identity = any(v in ("codigo", "codigomatricula", "email", "correo", "email") for v in values)
        if has_document and has_identity:
            header_row = idx
            break

    if header_row is None:
        return pd.read_excel(xlsx, sheet_name=sheet_name)

    headers = []
    seen = {}
    for idx, value in enumerate(raw.iloc[header_row].tolist()):
        name = _clean_text(value) or f"COL_{idx}"
        count = seen.get(name, 0)
        seen[name] = count + 1
        headers.append(name if count == 0 else f"{name}.{count}")

    df = raw.iloc[header_row + 1:].copy().reset_index(drop=True)
    df.columns = headers
    return df


def _guess_area_from_program(programa) -> str:
    p = _norm_text(programa)
    if any(k in p for k in ["ingenier", "arquitect"]):
        return "A"
    if any(k in p for k in ["medicina", "enfermer", "obstetric", "psicolog", "odontolog", "farmac", "tecnologiamedica"]):
        return "B"
    if any(k in p for k in ["derecho", "administr", "contabil", "educacion", "comunic", "turismo"]):
        return "C"
    return ""


def _safe_float(v) -> float:
    if pd.isna(v):
        return 0.0
    s = str(v).strip().replace("%", "").replace(",", ".")
    if s == "":
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _to_upper_object_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(lambda x: x.upper() if isinstance(x, str) else x)
    return df


# ---------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------
with st.sidebar:
    st.subheader("⚙️ Parámetros generales")

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
        help="Si el porcentaje obtenido en un curso es menor o igual a este valor, el postulante requiere nivelación en ese curso.",
    )

    st.markdown("---")
    st.subheader("📊 Umbrales de nivelación por área y curso")

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
padron_file = st.file_uploader(
    "📋 Padrón de postulantes que rendirán en la fecha seleccionada",
    type=["xlsx"],
    key="padron_examen_fecha",
    help="Se usa para incluir solo a los postulantes programados para esta fecha e identificar quiénes no rindieron.",
)

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

    nivel_threshold = nivel_threshold_pct / 100.0

    BASE_DIR = Path(__file__).resolve().parent

    modelo_path = None
    candidates = [
        "MODELO DE RESULTADOS DEL EXAMEN.xlsx",
        "MODELO_RESULTADOS_EXAMEN.xlsx",
        "PLANTILLA_DESCARGA_MOODLE_ADMISION.xlsx",
    ]
    for name in candidates:
        p = BASE_DIR / name
        if p.exists():
            modelo_path = p
            break

    if modelo_path:
        st.info(f"✅ Plantilla detectada (opcional): {modelo_path.name}")
    else:
        st.info("✅ No se usará plantilla. Se generará ACTAS automáticamente.")

    excels_en_carpeta = sorted([x.name for x in BASE_DIR.glob("*.xlsx")])

    if not modelo_path:
        st.error(
            "❌ No encuentro la plantilla para Actas.\n\n"
            "Coloca el archivo en la misma carpeta del app_streamlit_admision.py.\n\n"
            f"📁 Carpeta actual: {BASE_DIR.as_posix()}\n"
            f"📄 Excel detectados: {excels_en_carpeta}\n\n"
            "Nombres esperados (cualquiera de estos):\n"
            "- MODELO DE RESULTADOS DEL EXAMEN.xlsx\n"
            "- MODELO_RESULTADOS_EXAMEN.xlsx\n"
            "- PLANTILLA_DESCARGA_MOODLE_ADMISION.xlsx"
        )
        st.stop()

    st.info(f"✅ Plantilla usada: {modelo_path.name}")

    try:
        course_ids = [int(x) for x in course_ids_str.split(",") if x.strip()]
        t_from, t_to, tz = core.day_range_epoch(exam_date.isoformat(), tz_offset)

        usar_padron = padron_file is not None
        padron_dnis = set()
        padron_codigos = set()
        padron_correos = set()
        padron_area_por_dni = {}
        padron_area_por_codigo = {}
        padron_area_por_correo = {}
        padron_records = []

        def _area_letter(v) -> str:
            area_text = _clean_upper_text(v)
            for letter in ("A", "B", "C"):
                if area_text == letter or area_text.endswith(f" {letter}") or area_text.endswith(f"-{letter}"):
                    return letter
            return ""

        if usar_padron:
            df_padron = _read_padron_flexible(padron_file)
            col_padron_dni = _find_col_flexible(df_padron, [["dni"], ["documento"], ["numero", "dni"]])
            col_padron_codigo = _find_col_flexible(df_padron, [["codigo", "matricula"], ["cod", "matr"], ["codigo"]])
            col_padron_correo = _find_col_flexible(df_padron, [["correo"], ["email"], ["mail"]])
            col_padron_area = _find_col_flexible(df_padron, [["area"]])
            col_padron_fecha = _find_col_flexible(df_padron, [["fecha", "examen"], ["dia", "examen"]])
            col_padron_programa = _find_col_flexible(df_padron, [["programa", "academico"], ["programa"], ["carrera"]])
            col_padron_nombre = _find_col_flexible(df_padron, [["apellidos", "nombres"], ["nombre", "completo"], ["nombres"]])
            col_padron_sede = _find_col_flexible(df_padron, [["sede", "filial"], ["sede"], ["filial"]])

            if not any([col_padron_dni, col_padron_codigo, col_padron_correo]):
                st.error("El padrón debe tener al menos una columna de DNI, Código de Matrícula o correo.")
                st.info(f"Columnas detectadas: {list(df_padron.columns)}")
                st.stop()

            st.info(
                "Columnas del padrón detectadas: "
                f"DNI={col_padron_dni or '-'} | Código={col_padron_codigo or '-'} | "
                f"Correo={col_padron_correo or '-'} | Programa={col_padron_programa or '-'} | "
                f"Área={col_padron_area or '-'} | Fecha={col_padron_fecha or '-'}"
            )

            if col_padron_fecha:
                fechas_padron = pd.to_datetime(df_padron[col_padron_fecha], errors="coerce", dayfirst=True).dt.date
                df_padron = df_padron[fechas_padron == exam_date].copy()
                if df_padron.empty:
                    st.error(f"El padrón no contiene postulantes para la fecha {exam_date}.")
                    st.stop()
                st.info(f"Padrón filtrado por {col_padron_fecha}: {len(df_padron)} postulantes para {exam_date}.")
            else:
                st.warning(
                    "El padrón no tiene una columna 'Fecha de examen'. "
                    "Se asumirá que todos los postulantes del archivo corresponden a la fecha seleccionada."
                )

            for idx, row in df_padron.iterrows():
                dni = _norm_dni_value(row.get(col_padron_dni)) if col_padron_dni else ""
                codigo = _clean_upper_text(row.get(col_padron_codigo)) if col_padron_codigo else ""
                correo = _clean_text(row.get(col_padron_correo)).lower() if col_padron_correo else ""
                area = _area_letter(row.get(col_padron_area)) if col_padron_area else ""
                if not area and col_padron_programa:
                    area = _guess_area_from_program(row.get(col_padron_programa))
                padron_records.append({
                    "index": idx,
                    "dni": dni,
                    "codigo": codigo,
                    "correo": correo,
                    "area": area,
                    "nombre": _clean_upper_text(row.get(col_padron_nombre)) if col_padron_nombre else "",
                    "programa": _clean_upper_text(row.get(col_padron_programa)) if col_padron_programa else "",
                    "sede": _clean_upper_text(row.get(col_padron_sede)) if col_padron_sede else "",
                })
                if dni:
                    padron_dnis.add(dni)
                    if area:
                        padron_area_por_dni[dni] = area
                if codigo:
                    padron_codigos.add(codigo)
                    if area:
                        padron_area_por_codigo[codigo] = area
                if correo:
                    padron_correos.add(correo)
                    if area:
                        padron_area_por_correo[correo] = area

            if not any([padron_dnis, padron_codigos, padron_correos]):
                st.error("El padrón no contiene postulantes válidos para procesar.")
                st.stop()
        else:
            st.info("No se subió padrón. Se procesarán los usuarios matriculados en Moodle como antes.")

        st.info(f"Cursos: {course_ids} | Día: {exam_date} (tz {tz_offset})")
        st.info(f"Quiz→Área: {quiz_map}")

        with st.status("🔁 Descubriendo quizzes…", expanded=False) as status:
            quizzes = core.discover_quizzes(base_url, TOKEN, course_ids)
            qids_in_cursos = {q["quizid"] for q in quizzes}
            target_qids = [qid for qid in quiz_map.keys() if qid in qids_in_cursos]
            target_quizzes = [q for q in quizzes if q["quizid"] in target_qids]
            status.update(label=f"Quizzes a procesar: {len(target_quizzes)}", state="complete")

        course_users = {}
        user_area_by_id = {}
        matched_padron_indices = set()
        total_users = 0
        prog_bar = st.progress(0, text="Cargando usuarios por curso…")
        for i, cid in enumerate(course_ids, start=1):
            us = core.get_course_users(
                base_url, TOKEN, cid,
                only_roles=[x.strip() for x in only_roles.split(",") if x.strip()],
            )
            usuarios_fecha = []
            for u in us:
                custom = u.get("custom", {})
                dni = _norm_dni_value(custom.get(core.CF_DNI, ""))
                codigo = _clean_upper_text(custom.get(core.CF_COD_MAT, ""))
                correo = _clean_text(u.get("email", "")).lower()
                if not usar_padron:
                    usuarios_fecha.append(u)
                    user_area_by_id[u["id"]] = ""
                    continue
                if dni in padron_dnis or codigo in padron_codigos or correo in padron_correos:
                    usuarios_fecha.append(u)
                    user_area_by_id[u["id"]] = (
                        padron_area_por_dni.get(dni)
                        or padron_area_por_codigo.get(codigo)
                        or padron_area_por_correo.get(correo)
                        or ""
                    )
                    for record in padron_records:
                        if (
                            (dni and record["dni"] == dni)
                            or (codigo and record["codigo"] == codigo)
                            or (correo and record["correo"] == correo)
                        ):
                            matched_padron_indices.add(record["index"])
            course_users[cid] = usuarios_fecha
            total_users += len(usuarios_fecha)
            label = "postulantes del padrón" if usar_padron else "usuarios"
            prog_bar.progress(i / len(course_ids), text=f"Curso {cid}: {len(usuarios_fecha)} {label}")
        prog_bar.empty()

        if not target_quizzes:
            st.warning("Nada para procesar: no hay quizzes objetivo.")
            st.stop()
        if total_users == 0 and usar_padron:
            st.warning("No se encontraron usuarios del padrón en Moodle. Se intentará incluirlos como NO ASISTIÓ.")
        elif total_users == 0:
            st.warning("No se encontraron usuarios en Moodle para procesar.")
            st.stop()

        st.write("⚙️ Procesando intentos (esto puede tardar)…")
        t0 = time.time()
        rows = []
        open_attempts = []
        if not hasattr(core, "inspect_user_quiz"):
            st.error(
                "El servidor está usando una versión anterior de moodle_admision_export.py. "
                "Reinicia o vuelve a desplegar la aplicación para cargar el módulo actualizado."
            )
            st.stop()
        target_quizzes_by_course = {}
        for q in target_quizzes:
            target_quizzes_by_course.setdefault(q["courseid"], []).append(q)
        from concurrent.futures import ThreadPoolExecutor, as_completed

        futs = []
        with ThreadPoolExecutor(max_workers=workers) as ex:
            for q in target_quizzes:
                area_letter = quiz_map.get(q["quizid"])
                users = course_users.get(q["courseid"], [])
                for u in users:
                    area_padron = user_area_by_id.get(u["id"], "")
                    include_no_attempt = usar_padron and (
                        len(target_quizzes_by_course.get(q["courseid"], [])) == 1
                        or area_padron == area_letter
                    )
                    futs.append(ex.submit(
                        core.inspect_user_quiz,
                        base_url, TOKEN, q, area_letter, u, t_from, t_to, tz,
                        include_no_attempt,
                    ))

            done = 0
            step_bar = st.progress(0.0)
            for fut in as_completed(futs):
                res = fut.result()
                if res.get("rows"):
                    rows.extend(res["rows"])
                if res.get("open_attempts"):
                    open_attempts.extend(res["open_attempts"])
                done += 1
                step_bar.progress(done / max(1, len(futs)))
        step_bar.empty()

        quizzes_by_area = {}
        for q in target_quizzes:
            quizzes_by_area.setdefault(quiz_map.get(q["quizid"]), []).append(q)
        padron_no_encontrados = []
        if usar_padron:
            for record in padron_records:
                if record["index"] in matched_padron_indices:
                    continue
                area_quizzes = quizzes_by_area.get(record["area"], [])
                if len(area_quizzes) != 1:
                    padron_no_encontrados.append(record)
                    continue
                pseudo_user = {
                    "id": f"padron-{record['index']}",
                    "firstname": "",
                    "lastname": record["nombre"],
                    "email": record["correo"],
                    "custom": {
                        core.CF_DNI: record["dni"],
                        core.CF_COD_MAT: record["codigo"],
                        core.CF_PROG: record["programa"],
                        core.CF_SEDE: record["sede"],
                    },
                }
                rows.append(core.build_row_no_attempt(pseudo_user, area_quizzes[0], record["area"]))

        if open_attempts:
            st.error(
                "No se puede generar la descarga porque Moodle todavía reporta "
                f"{len(open_attempts)} intento(s) en curso o vencido(s)."
            )
            st.dataframe(pd.DataFrame(open_attempts), use_container_width=True)
            st.info("Finaliza o cierra esos intentos en Moodle y vuelve a generar el reporte.")
            st.stop()

        asistentes = sum(1 for row in rows if row.get("Asistencia") == "ASISTIÓ")
        no_asistentes = sum(1 for row in rows if row.get("Asistencia") == "NO ASISTIÓ")
        st.success(f"Registros: {len(rows)} | Asistieron: {asistentes} | No asistieron: {no_asistentes}")

        if padron_no_encontrados:
            st.warning(
                f"{len(padron_no_encontrados)} postulante(s) del padrón no se encontraron en Moodle "
                "y no pudieron asignarse a un único quiz por área."
            )
            st.dataframe(pd.DataFrame(padron_no_encontrados), use_container_width=True)

        usuarios_sin_area = [
            u for users in course_users.values() for u in users
            if not user_area_by_id.get(u["id"], "")
        ]
        if usar_padron and usuarios_sin_area and any(len(qzs) > 1 for qzs in target_quizzes_by_course.values()):
            st.warning(
                "Algunos postulantes del padrón no tienen Área. Si no rindieron y su curso tiene varios quizzes, "
                "no será posible asignarlos a un examen. Agrega la columna Área al padrón."
            )

        if not rows:
            st.warning("No se encontraron intentos ese día.")
            st.stop()

        fname_base = f"RESULTADOS_ADMISION_{exam_date}.xlsx"
        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / fname_base

            core.write_excel_all_in_one(
                out_path,
                rows,
                nivel_threshold_base=nivel_threshold,
            )
            base_bytes = out_path.read_bytes()

            if not zipfile.is_zipfile(BytesIO(base_bytes)):
                st.error("❌ El Excel base generado NO es un .xlsx válido (ZIP interno). Revisa openpyxl/pandas.")
                st.stop()

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
# CONVERSOR A FORMATO BD
# =====================================================================
st.markdown("---")
st.header("📂 Conversor a formato BD")

tab1, tab2, tab3 = st.tabs([
    "✅ Desde Excel Moodle (RESULTADOS/RESUMEN)",
    "📤 Archivo de la comisión",
    "🔁 Reagrupando plantilla / sistema de admisión"
])


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

            col_dni_res = "Numero de DNI" if "Numero de DNI" in df_resultados.columns else _find_col_flexible(
                df_resultados, [
                    ["numero", "dni"],
                    ["dni"],
                    ["documento", "dni"],
                    ["nro", "dni"],
                ]
            )

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

            col_cod = None
            for exact in [
                "Código de Matrícula", "Codigo de Matricula", "CÓDIGO DE MATRÍCULA", "CODIGO DE MATRICULA",
            ]:
                if exact in df_resultados.columns:
                    col_cod = exact
                    break

            if not col_cod:
                col_cod = _find_col_flexible(df_resultados, [
                    ["codigo", "matricula"],
                    ["cod", "matr"],
                ])

            col_cod_resumen = None
            for exact in [
                "Código de Matrícula", "Codigo de Matricula", "CÓDIGO DE MATRÍCULA", "CODIGO DE MATRICULA",
            ]:
                if exact in df_resumen.columns:
                    col_cod_resumen = exact
                    break

            if not col_cod_resumen:
                col_cod_resumen = _find_col_flexible(df_resumen, [
                    ["codigo", "matricula"],
                    ["cod", "matr"],
                ])

            if not col_cod and not col_cod_resumen and len(df_resumen.columns) <= 2:
                st.warning("No encontré columna de CÓDIGO (MATRÍCULA/ESTUDIANTE) en RESULTADOS ni RESUMEN. Saldrá vacío.")
                st.info(f"Columnas RESULTADOS: {list(df_resultados.columns)}")
                st.info(f"Columnas RESUMEN: {list(df_resumen.columns)}")

            df_resultados["_dni_norm"] = _norm_dni_series(df_resultados[col_dni_res])
            df_resumen["_dni_norm"] = _norm_dni_series(df_resumen[col_dni_sum])

            codigo_estudiante_resumen = (
                df_resumen.iloc[:, 2].apply(_clean_text)
                if df_resumen.shape[1] > 2
                else pd.Series([""] * len(df_resumen))
            )
            condicion_resumen = (
                df_resumen.iloc[:, 25].apply(_clean_upper_text)
                if df_resumen.shape[1] > 25
                else pd.Series([""] * len(df_resumen))
            )

            cols_small = ["_dni_norm"]
            if "Apellido(s)" in df_resultados.columns:
                cols_small.append("Apellido(s)")
            if "Nombre" in df_resultados.columns:
                cols_small.append("Nombre")
            if col_cod:
                cols_small.append(col_cod)

            df_small = df_resultados[cols_small].copy()
            if col_cod:
                df_small["_codigo_estudiante_src"] = df_small[col_cod].apply(_clean_text)
                df_small = df_small.drop(columns=[col_cod])

            merged = df_resumen.merge(
                df_small,
                on="_dni_norm",
                how="left",
            )
            merged["_codigo_estudiante_resumen_directo"] = codigo_estudiante_resumen.reset_index(drop=True)
            merged["_condicion_resumen_directa"] = condicion_resumen.reset_index(drop=True)
            codigo_estudiante = merged["_codigo_estudiante_resumen_directo"].apply(_clean_text)

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
                        cursos.append({"curso": nombre.upper()})
                return json.dumps(cursos, ensure_ascii=False)

            areas_nivelacion = merged.apply(build_json_courses, axis=1)

            req = merged["PROGRAMA DE NIVELACIÓN"].fillna("").astype(str) if "PROGRAMA DE NIVELACIÓN" in merged.columns else pd.Series([""] * len(merged))
            requiere_nivelacion = req.apply(
                lambda x: "SI" if str(x).strip().upper() in ("REQUIERE NIVELACIÓN", "REQUIERE NIVELACION", "SI") else "NO"
            )
            condicion = merged["_condicion_resumen_directa"].apply(_clean_upper_text)
            asistio = merged["Asistencia"].apply(_clean_upper_text) if "Asistencia" in merged.columns else pd.Series([""] * len(merged))
            no_asistio_mask = asistio.apply(_norm_text).eq("noasistio")
            condicion = condicion.mask(no_asistio_mask, "NO INGRESÓ")
            requiere_nivelacion = requiere_nivelacion.mask(no_asistio_mask, "NO")
            areas_nivelacion = areas_nivelacion.mask(no_asistio_mask, "[]")

            out_df = pd.DataFrame({
                "id": None,
                "periodo": periodo_value.upper(),
                "codigo_estudiante": codigo_estudiante.astype(str).str.upper(),
                "apellidos": merged["Apellido(s)"].apply(_clean_upper_text) if "Apellido(s)" in merged.columns else "",
                "nombres": merged["Nombre"].apply(_clean_upper_text) if "Nombre" in merged.columns else "",
                "dni": merged[col_dni_sum].apply(_norm_dni_value),
                "area": merged["Área"].apply(_clean_upper_text) if "Área" in merged.columns else "",
                "programa": merged["Programa Académico"].apply(_clean_upper_text) if "Programa Académico" in merged.columns else "",
                "local_examen": merged["Sede o Filial"].apply(_clean_upper_text) if "Sede o Filial" in merged.columns else "",
                "modalidad_examen": "VIRTUAL",
                "puntaje": pd.to_numeric(merged["TOTAL"], errors="coerce").fillna(0).astype(int) if "TOTAL" in merged.columns else 0,
                "asistio": asistio,
                "condicion": condicion,
                "requiere_nivelacion": requiere_nivelacion.astype(str).str.upper(),
                "areas_nivelacion": areas_nivelacion.astype(str),
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
# TAB 2: Archivo de la comisión (formato consolidado)
# ==========================================================
with tab2:
    st.write(
        "📤 **Subir archivo de la comisión (Cuadro consolidado oficial)**.\n\n"
        "- Soporta el formato consolidado con 2 filas de encabezado.\n"
        "- Convierte el archivo a la plantilla final para BD.\n"
        "- La nivelación se calculará según el umbral porcentual que definas aquí."
    )

    com_file = st.file_uploader(
        "📤 Subir archivo de la comisión (Excel)",
        type=["xlsx"],
        key="comision_excel",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        periodo_value_com = st.text_input("Periodo", value="2026-1", key="periodo_comision")
    with c2:
        fecha_registro_value_com = st.text_input(
            "Fecha de registro (AAAA-MM-DD hh:mm:ss)",
            value="2025-11-29 00:00:00",
            key="fecha_comision",
        )
    with c3:
        umbral_nivelacion_com_pct = st.number_input(
            "Umbral nivelación comisión (%)",
            min_value=0.0,
            max_value=100.0,
            value=30.0,
            step=1.0,
            key="umbral_nivelacion_com_pct",
            help="Si el porcentaje de un curso es menor o igual a este valor, irá a nivelación."
        )

    convertir_com = st.button("🔄 Convertir archivo de comisión → Plantilla BD", key="btn_convertir_comision")

    def _norm_comm(s: str) -> str:
        s = str(s).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        return "".join(ch for ch in s if ch.isalnum())

    def _norm_dni_comm(v) -> str:
        s = "" if pd.isna(v) else str(v).strip()
        if s.endswith(".0"):
            s = s[:-2]
        digits = "".join(ch for ch in s if ch.isdigit())
        if not digits:
            return ""
        return digits.zfill(8) if len(digits) < 8 else digits

    def _build_two_row_header(df_raw: pd.DataFrame):
        h1 = df_raw.iloc[3].fillna("")
        h2 = df_raw.iloc[4].fillna("")
        cols = []

        current_main = ""

        for a, b in zip(h1, h2):
            a = str(a).strip()
            b = str(b).strip()

            if a:
                current_main = a

            main_norm = _norm_comm(current_main)
            sub_norm = _norm_comm(b)

            if main_norm == "comunicacion" and sub_norm.startswith("punt"):
                cols.append("COMUNICACION_PUNT")
            elif main_norm == "comunicacion" and sub_norm == "":
                cols.append("COMUNICACION_%")

            elif main_norm == "habilidadescomunicativas" and sub_norm.startswith("punt"):
                cols.append("HABILIDADES_COMUNICATIVAS_PUNT")
            elif main_norm == "habilidadescomunicativas" and sub_norm == "":
                cols.append("HABILIDADES_COMUNICATIVAS_%")

            elif main_norm == "matematica" and sub_norm.startswith("punt"):
                cols.append("MATEMATICA_PUNT")
            elif main_norm == "matematica" and sub_norm == "":
                cols.append("MATEMATICA_%")

            elif main_norm == "cta" and sub_norm.startswith("punt"):
                cols.append("CTA_CIENCIAS_SOCIALES_PUNT")
            elif main_norm == "cta" and sub_norm == "":
                cols.append("CTA_CIENCIAS_SOCIALES_%")

            elif a:
                cols.append(a)
            else:
                cols.append(f"COL_{len(cols)}")

        return cols

    def _parse_ratio(v):
        x = _safe_float(v)
        return x / 100.0 if x > 1 else x

    if convertir_com:
        if com_file is None:
            st.error("Primero sube el Excel de la comisión.")
            st.stop()

        try:
            raw = pd.read_excel(com_file, header=None)

            if raw.empty or len(raw) < 6:
                st.error("El archivo no tiene la estructura esperada.")
                st.stop()

            cols = _build_two_row_header(raw)
            df = raw.iloc[5:].copy().reset_index(drop=True)
            df.columns = cols

            if "DNI" not in df.columns:
                st.error("No pude detectar la columna DNI en el archivo consolidado.")
                st.info(f"Columnas detectadas: {list(df.columns)}")
                st.stop()

            df["DNI"] = df["DNI"].apply(_norm_dni_comm)
            df = df[df["DNI"] != ""].copy()

            if df.empty:
                st.error("No encontré registros válidos con DNI.")
                st.stop()

            col_ap = "APELLIDOS" if "APELLIDOS" in df.columns else None
            col_nom = "NOMBRES" if "NOMBRES" in df.columns else None
            col_dni = "DNI"
            col_area = "AREA" if "AREA" in df.columns else None
            col_prog = "PROGRAMA" if "PROGRAMA" in df.columns else None
            col_total = "PUNTAJE FINAL" if "PUNTAJE FINAL" in df.columns else None
            col_asist = "ASISTENCIA" if "ASISTENCIA" in df.columns else None
            col_cond = "CONDICIÓN" if "CONDICIÓN" in df.columns else ("CONDICION" if "CONDICION" in df.columns else None)
            col_cod = "CODIGO" if "CODIGO" in df.columns else None
            col_sede = "DIRECCIÓN LOCAL" if "DIRECCIÓN LOCAL" in df.columns else None

            faltantes = []
            if not col_ap:
                faltantes.append("APELLIDOS")
            if not col_nom:
                faltantes.append("NOMBRES")
            if not col_area:
                faltantes.append("AREA")
            if not col_prog:
                faltantes.append("PROGRAMA")
            if not col_total:
                faltantes.append("PUNTAJE FINAL")

            if faltantes:
                st.error(f"No pude detectar estas columnas necesarias: {', '.join(faltantes)}")
                st.info(f"Columnas detectadas: {list(df.columns)}")
                st.stop()

            pct_cols = {
                "COMUNICACIÓN": "COMUNICACION_%",
                "HABILIDADES COMUNICATIVAS": "HABILIDADES_COMUNICATIVAS_%",
                "MATEMATICA": "MATEMATICA_%",
                "CTA/CIENCIAS SOCIALES": "CTA_CIENCIAS_SOCIALES_%",
            }

            threshold_decimal = umbral_nivelacion_com_pct / 100.0

            def build_json_from_comision(row):
                cursos = []

                condicion_actual = _clean_text(row.get(col_cond)).upper() if col_cond else ""

                # ✅ Solo los que INGRESARON pueden llevar nivelación
                if condicion_actual != "INGRESÓ" and condicion_actual != "INGRESO":
                    return json.dumps([], ensure_ascii=False)

                val_com = _parse_ratio(row.get("COMUNICACION_%"))
                if val_com <= threshold_decimal:
                    cursos.append({"curso": "COMUNICACIÓN"})

                val_hab = _parse_ratio(row.get("HABILIDADES_COMUNICATIVAS_%"))
                if val_hab <= threshold_decimal:
                    cursos.append({"curso": "HABILIDADES COMUNICATIVAS"})

                val_mat = _parse_ratio(row.get("MATEMATICA_%"))
                if val_mat <= threshold_decimal:
                    cursos.append({"curso": "MATEMATICA"})

                val_cta = _parse_ratio(row.get("CTA_CIENCIAS_SOCIALES_%"))
                if val_cta <= threshold_decimal:
                    area_actual = _clean_text(row.get(col_area)).upper()
                    if area_actual == "C":
                        cursos.append({"curso": "CIENCIAS SOCIALES"})
                    else:
                        cursos.append({"curso": "CIENCIA, TECNOLOGÍA Y AMBIENTE"})

                return json.dumps(cursos, ensure_ascii=False)

            areas_nivelacion = df.apply(build_json_from_comision, axis=1)

            requiere_nivelacion = areas_nivelacion.apply(
                lambda x: "SI" if x != "[]" else "NO"
            )

            asistio = df[col_asist].apply(_clean_upper_text) if col_asist else pd.Series(["ASISTIÓ"] * len(df))
            condicion = df[col_cond].apply(_clean_upper_text) if col_cond else pd.Series([""] * len(df))
            codigo_estudiante = df[col_cod].apply(_clean_upper_text) if col_cod else pd.Series([""] * len(df))
            puntaje = pd.to_numeric(df[col_total], errors="coerce").fillna(0).astype(int)

            out_df = pd.DataFrame({
                "id": None,
                "periodo": periodo_value_com.upper(),
                "codigo_estudiante": codigo_estudiante,
                "apellidos": df[col_ap].apply(_clean_upper_text),
                "nombres": df[col_nom].apply(_clean_upper_text),
                "dni": df[col_dni].apply(_norm_dni_comm),
                "area": df[col_area].apply(_clean_upper_text),
                "programa": df[col_prog].apply(_clean_upper_text),
                "local_examen": df[col_sede].apply(_clean_upper_text) if col_sede else "",
                "puntaje": puntaje,
                "asistio": asistio,
                "condicion": condicion,
                "requiere_nivelacion": requiere_nivelacion.astype(str).str.upper(),
                "areas_nivelacion": areas_nivelacion.astype(str).str.upper(),
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

            con_nivel = (out_df["requiere_nivelacion"] == "SI").sum()
            st.info(f"Registros procesados: {len(out_df)} | Requieren nivelación: {con_nivel}")
            st.dataframe(out_df.head())

        except Exception as e:
            st.error(f"❌ Error convirtiendo archivo de comisión: {e}")
            st.stop()
            
# ==========================================================
# TAB 3: Reagrupando plantilla / sistema de admisión
# ==========================================================
with tab3:
    st.write(
        "🔁 **Reagrupa las preguntas del archivo ACTA_FINAL_Y_RESUMEN al orden presencial**.\n\n"
        "- Toma la hoja **RESULTADOS** del Excel generado.\n"
        "- Usa el **Área** de cada postulante para reordenar P.1...P.100 a Q1...Q100.\n"
        "- Convierte 0.2 → 1 y 0 → 0.\n"
        "- Usa la plantilla CSV como estructura de salida."
    )

    acta_file = st.file_uploader(
        "📥 Sube el archivo ACTA_FINAL_Y_RESUMEN (.xlsx)",
        type=["xlsx"],
        key="reagrupar_acta_file",
    )

    plantilla_file = st.file_uploader(
        "📥 Sube la plantilla del sistema de admisión (.csv)",
        type=["csv"],
        key="reagrupar_plantilla_file",
    )

    convertir_reagrupado = st.button(
        "🔄 Reagrupar preguntas y generar CSV",
        key="btn_reagrupar_plantilla"
    )

    def _val_to_q(v):
        # Si viene vacío, también cuenta como 0
        if pd.isna(v):
            return "0"

        s = str(v).strip().replace(",", ".")
        if s == "" or s.lower() == "nan":
            return "0"

        try:
            x = float(s)
        except Exception:
            return "0"

        if abs(x - 0.2) < 1e-9:
            return "1"

        # Todo lo demás lo tratamos como incorrecto = 0
        return "0"

    def _build_mapping_by_area(area: str):
        area = _clean_text(area).upper()

        if area == "A":
            return (
                list(range(1, 26)) +      # Q1-Q25   <- P1-P25
                list(range(26, 76)) +     # Q26-Q75  <- P26-P75
                list(range(76, 86)) +     # Q76-Q85  <- P76-P85
                list(range(86, 101))      # Q86-Q100 <- P86-P100
            )

        if area == "B":
            return (
                list(range(1, 26)) +      # Q1-Q25   <- P1-P25
                list(range(26, 56)) +     # Q26-Q55  <- P26-P55
                list(range(56, 66)) +     # Q56-Q65  <- P56-P65
                list(range(66, 101))      # Q66-Q100 <- P66-P100
            )

        # Área C
        return (
            list(range(1, 36)) +          # Q1-Q35   <- P1-P35
            list(range(36, 66)) +         # Q36-Q65  <- P36-P65
            list(range(66, 76)) +         # Q66-Q75  <- P66-P75
            list(range(76, 101))          # Q76-Q100 <- P76-P100
        )

    def _find_resultados_sheet(xlsx: pd.ExcelFile):
        if "RESULTADOS" in xlsx.sheet_names:
            return "RESULTADOS"
        for sh in xlsx.sheet_names:
            if _norm_text(sh) == "resultados":
                return sh
        return None

    def _find_col_case_insensitive(df: pd.DataFrame, target: str):
        target_norm = _norm_text(target)
        for c in df.columns:
            if _norm_text(c) == target_norm:
                return c
        return None

    if convertir_reagrupado:
        if acta_file is None:
            st.error("Primero sube el archivo ACTA_FINAL_Y_RESUMEN.")
            st.stop()

        if plantilla_file is None:
            st.error("Primero sube la plantilla CSV.")
            st.stop()

        try:
            # =============================
            # 1) Leer ACTA_FINAL_Y_RESUMEN
            # =============================
            xlsx = pd.ExcelFile(acta_file)
            sheet_resultados = _find_resultados_sheet(xlsx)

            if not sheet_resultados:
                st.error(f"No encontré la hoja RESULTADOS. Hojas detectadas: {xlsx.sheet_names}")
                st.stop()

            df_res = pd.read_excel(xlsx, sheet_name=sheet_resultados)

            col_dni = (
                _find_col_case_insensitive(df_res, "Numero de DNI")
                or _find_col_case_insensitive(df_res, "DNI")
            )
            col_area = (
                _find_col_case_insensitive(df_res, "Área")
                or _find_col_case_insensitive(df_res, "Area")
            )

            if not col_dni or not col_area:
                st.error("No pude detectar las columnas de DNI y/o Área en la hoja RESULTADOS.")
                st.info(f"Columnas detectadas: {list(df_res.columns)}")
                st.stop()

            # detectar P. 1 /0.2 ... P. 100 /0.2
            p_cols = {}
            for i in range(1, 101):
                expected = f"P. {i} /0.2"
                found = _find_col_case_insensitive(df_res, expected)
                if not found:
                    st.error(f"No encontré la columna '{expected}' en RESULTADOS.")
                    st.stop()
                p_cols[i] = found

            df_res["_dni_norm"] = df_res[col_dni].apply(_norm_dni_value)
            df_res["_area_norm"] = df_res[col_area].apply(lambda x: _clean_text(x).upper())

            df_res = df_res[df_res["_dni_norm"] != ""].copy()

            if df_res.empty:
                st.error("No encontré registros válidos con DNI en RESULTADOS.")
                st.stop()

            # =============================
            # 2) Leer plantilla CSV solo como estructura
            # =============================
            try:
                df_csv = pd.read_csv(plantilla_file, dtype=str, encoding="utf-8-sig")
            except Exception:
                plantilla_file.seek(0)
                df_csv = pd.read_csv(plantilla_file, dtype=str, encoding="latin-1")

            if df_csv.empty and len(df_csv.columns) == 0:
                st.error("No pude leer la estructura de la plantilla CSV.")
                st.stop()

            plantilla_cols = list(df_csv.columns)

            # =============================
            # 3) Construir salida desde RESULTADOS
            # =============================
            output_rows = []

            for _, row in df_res.iterrows():
                area = row["_area_norm"]
                mapping = _build_mapping_by_area(area)

                out = {col: "" for col in plantilla_cols}

                # DNI
                col_out_dni = None
                for cand in ["DNI", "NUMERO DE DNI", "NRO DNI", "DOCUMENTO", "DOC"]:
                    for c in plantilla_cols:
                        if _norm_text(c) == _norm_text(cand):
                            col_out_dni = c
                            break
                    if col_out_dni:
                        break

                if col_out_dni:
                    out[col_out_dni] = row["_dni_norm"]

                # AREA
                col_out_area = None
                for c in plantilla_cols:
                    if _norm_text(c) == _norm_text("AREA"):
                        col_out_area = c
                        break

                if col_out_area:
                    out[col_out_area] = area

                # Q1...Q100
                for q_idx, p_src in enumerate(mapping, start=1):
                    q_col = f"Q{q_idx}"
                    if q_col in out:
                        out[q_col] = _val_to_q(row[p_cols[p_src]])

                output_rows.append(out)

            out_df = pd.DataFrame(output_rows)

            # Mayúsculas en texto
            for col in out_df.columns:
                if out_df[col].dtype == object:
                    out_df[col] = out_df[col].apply(lambda x: x.upper() if isinstance(x, str) else x)

            # =============================
            # 4) Descargar CSV
            # =============================
            csv_bytes = out_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

            st.success(f"✅ Archivo generado correctamente. Registros: {len(out_df)}")
            st.download_button(
                label="⬇️ Descargar CSV reagrupado para sistema de admisión",
                data=csv_bytes,
                file_name="plantilla_reagrupada_sistema_admision.csv",
                mime="text/csv",
            )

            preview_cols = [c for c in out_df.columns if c in ["DNI", "AREA", "Q1", "Q2", "Q3", "Q4", "Q5", "Q96", "Q97", "Q98", "Q99", "Q100"]]
            if preview_cols:
                st.dataframe(out_df[preview_cols].head())

        except Exception as e:
            st.error(f"❌ Error al reagrupar la plantilla: {e}")
            st.stop()
