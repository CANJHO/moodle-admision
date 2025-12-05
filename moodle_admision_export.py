#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Exportador de Admisión (Moodle REST) — TODO EN UNO (sin plantilla)
- Filtra por DÍA (00:00–23:59:59, tz por defecto -05:00)
- Procesa los quizzes indicados y mapea quiz→Área (A/B/C)
- Genera un Excel con:
  * Hoja 'RESULTADOS' (todas las áreas juntas, ordenado por Programa Académico y luego DNI)
  * Hoja 'RESUMEN' (mismos datos agregados y orden)
- Sub-áreas por Área (rango de preguntas) tal como definiste:
  AREA A – INGENIERÍAS:
    COM: 1–21,97–100 (25) | MAT: 22–71 (50) | HABIL: 72–81 (10) | CTA: 82–96 (15)
  AREA B – CIENCIAS SALUD:
    COM: 1–21,97–100 (25) | MAT: 22–51 (30) | HABIL: 52–61 (10) | CTA: 62–96 (35)
  AREA C – CIENCIAS HUMANAS:
    COM: 1–31,97–100 (35) | MAT: 32–61 (30) | HABIL: 62–71 (10) | CCSS: 72–96 (25)

Requisitos:
  pip install requests pandas openpyxl python-dateutil
"""

import argparse
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date, time as dtime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -------- Config WS / HTTP --------
WS_PATH = "webservice/rest/server.php"
TIMEOUT = 60

# Custom fields (shortnames en Moodle)
CF_DNI = "DNI_CE"
CF_PROG = "PROGRAMA_ACADEMICO"    # etiqueta salida: "Programa Académico"
CF_SEDE = "SEDE_FILIAL"           # etiqueta salida: "Sede o Filial"
CF_COD_MAT = "CODIGO_DE_MATRICULA"   # etiqueta salida: "Código de Matrícula"

# -------- Áreas y sub-áreas (rangos P#) --------
def r(a: int, b: int) -> List[int]:
    return list(range(a, b + 1))

AREA_DEFS: Dict[str, Dict[str, Any]] = {
    "A": {  # INGENIERÍAS
        "label": "AREA-A INGENIERÍAS",
        "cta_label": "CIENCIA TECNOLOGÍA Y AMBIENTE",
        "ranges": {
            "COMUNICACIÓN": r(1, 21) + r(97, 100),      # 25
            "MATEMÁTICA":   r(22, 71),                  # 50
            "HABILIDADES COMUNICATIVAS": r(72, 81),     # 10
            "CTA/CCSS":     r(82, 96),                  # 15
        },
    },
    "B": {  # CIENCIAS SALUD
        "label": "AREA B-CIENCIAS SALUD",
        "cta_label": "CIENCIA TECNOLOGÍA Y AMBIENTE",
        "ranges": {
            "COMUNICACIÓN": r(1, 21) + r(97, 100),      # 25
            "MATEMÁTICA":   r(22, 51),                  # 30
            "HABILIDADES COMUNICATIVAS": r(52, 61),     # 10
            "CTA/CCSS":     r(62, 96),                  # 35
        },
    },
    "C": {  # CIENCIAS HUMANAS
        "label": "AREA C-CIENCIAS HUMANAS",
        "cta_label": "CCSS",
        "ranges": {
            "COMUNICACIÓN": r(1, 31) + r(97, 100),      # 35
            "MATEMÁTICA":   r(32, 61),                  # 30
            "HABILIDADES COMUNICATIVAS": r(62, 71),     # 10
            "CTA/CCSS":     r(72, 96),                  # 25   # CCSS
        },
    },
}

# -------- UMBRAL DE NIVELACIÓN --------
# 0.30 = 30% (se compara contra las columnas % (COM), % (HAB), etc. que van de 0 a 1)
NIVEL_THRESHOLD = 0.30


# -------- Sesión HTTP con reintentos --------
session = requests.Session()
retries = Retry(
    total=5,
    backoff_factor=0.2,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=frozenset(["GET", "POST"]),
)
adapter = HTTPAdapter(max_retries=retries, pool_connections=32, pool_maxsize=32)
session.mount("http://", adapter)
session.mount("https://", adapter)

def ws(base_url: str, token: str, wsfunction: str, **params) -> Any:
    url = f"{base_url.rstrip('/')}/{WS_PATH}"
    payload = dict(wstoken=token, wsfunction=wsfunction, moodlewsrestformat="json")
    payload.update(params)
    r = session.post(url, data=payload, timeout=TIMEOUT)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, dict) and data.get("exception"):
        raise RuntimeError(f"WS error [{data.get('errorcode','unknown')}]: {data.get('message','Unknown WS error')}")
    return data

# -------- Fechas --------
def day_range_epoch(day_str: str, tz_offset: str = "-05:00") -> Tuple[int, int, timezone]:
    sign = 1 if tz_offset.startswith("+") else -1
    hh, mm = tz_offset[1:].split(":")
    tz = timezone(sign * timedelta(hours=int(hh), minutes=int(mm)))
    d = date.fromisoformat(day_str)
    start_local = datetime.combine(d, dtime(0, 0, 0), tzinfo=tz)
    end_local   = datetime.combine(d, dtime(23, 59, 59), tzinfo=tz)
    return int(start_local.timestamp()), int(end_local.timestamp()), tz

# -------- Datos Moodle --------
def discover_quizzes(base_url: str, token: str, course_ids: List[int]) -> List[Dict[str, Any]]:
    quizzes = []
    idx = 0
    for cid in course_ids:
        qzs = ws(base_url, token, "mod_quiz_get_quizzes_by_courses", **{f"courseids[{idx}]": cid})
        idx += 1
        for q in qzs.get("quizzes", []):
            quizzes.append({"courseid": cid, "quizid": q["id"], "quizname": q["name"]})
    return quizzes

def get_course_users(base_url: str, token: str, course_id: int, only_roles: List[str]) -> List[Dict[str, Any]]:
    users = ws(base_url, token, "core_enrol_get_enrolled_users", courseid=course_id)
    out = []
    wanted = set([r.lower() for r in only_roles]) if only_roles else None
    for u in users:
        roles = [r.get("shortname","").lower() for r in u.get("roles",[])]
        if wanted and not (set(roles) & wanted):
            continue
        cf_map = {}
        for cf in u.get("customfields",[]):
            sn = cf.get("shortname","")
            val = cf.get("value","")
            if sn:
                cf_map[sn.upper()] = val
        out.append({
            "id": u["id"],
            "firstname": u.get("firstname",""),
            "lastname": u.get("lastname",""),
            "email": u.get("email",""),
            "idnumber": u.get("idnumber","") or "",
            "custom": cf_map,
        })
    return out

def get_user_attempts_in_range(base_url: str, token: str, quizid: int, userid: int, t_from: int, t_to: int) -> List[Dict[str, Any]]:
    tries = ws(base_url, token, "mod_quiz_get_user_attempts", quizid=quizid, userid=userid)
    attempts = tries.get("attempts", []) if isinstance(tries, dict) else tries
    out = []
    for a in attempts:
        ts = a.get("timestart") or 0
        tf = a.get("timefinish") or 0
        if tf == 0:
            continue
        if (t_from <= tf <= t_to) or (t_from <= ts <= t_to):
            out.append(a)
    return out

def get_attempt_review(base_url: str, token: str, attemptid: int) -> Dict[str, Any]:
    return ws(base_url, token, "mod_quiz_get_attempt_review", attemptid=attemptid)

# -------- Util notas / preguntas --------
def to_02(value: Any) -> Any:
    """
    Normaliza a 0.2 (correcto) o 0.0 (incorrecto), None si vacío.
    Acepta valores con coma decimal "0,2".
    """
    if value is None:
        return None
    if isinstance(value, str):
        value = value.replace(",", ".")  # convertir coma a punto
    try:
        x = float(value)
    except:
        return None
    if x > 0:
        return 0.2
    if abs(x) < 1e-9:
        return 0.0
    return 0.0

def count_correct(vals: List[Any]) -> int:
    return sum(1 for v in vals if v is not None and abs(v - 0.2) < 1e-9)

def count_responded(vals: List[Any]) -> int:
    # RESP = =0.2 o =0.0
    return sum(1 for v in vals if v is not None and (abs(v - 0.2) < 1e-9 or abs(v) < 1e-9))

def pct(numer: int, denom: int) -> float:
    return (numer / denom) if denom else 0.0

def build_row_from_review(user: Dict[str,Any], quiz: Dict[str,Any], area_letter: str,
                          attempt: Dict[str,Any], review: Dict[str,Any],
                          tz: timezone, max_questions: int = 100) -> Dict[str,Any]:
    # base fields
    ts = attempt.get("timestart") or 0
    tf = attempt.get("timefinish") or 0
    grade = review.get("grade")

    # vector preguntas 1..100
    q_vals = [None] * max_questions
    for q in (review.get("questions", []) if isinstance(review, dict) else []):
        slot = q.get("slot")
        if slot is None:
            continue
        try:
            slot = int(slot)
        except:
            continue
        if not (1 <= slot <= max_questions):
            continue
        mark = q.get("mark")
        if mark is None:
            # intentar fraction*maxmark si existe
            frac = q.get("fraction"); mx = q.get("maxmark")
            if isinstance(frac,(int,float)) and isinstance(mx,(int,float)):
                mark = float(frac) * float(mx)
        q_vals[slot - 1] = to_02(mark)

    # datos alumno + custom fields
    cf = user.get("custom", {})
    programa = cf.get(CF_PROG, "") or cf.get(CF_PROG.upper(), "") or ""
    sede     = cf.get(CF_SEDE, "") or cf.get(CF_SEDE.upper(), "") or ""
    dni      = cf.get(CF_DNI, "") or ""
    cod_mat  = cf.get(CF_COD_MAT, "") or cf.get(CF_COD_MAT.upper(), "") or ""
    base = {
        "Apellido(s)": user.get("lastname",""),
        "Nombre": user.get("firstname",""),
        "Dirección de correo": user.get("email",""),
        "Numero de DNI": dni,
        "Código de Matrícula": cod_mat,
        "Programa Académico": programa,
        "Sede o Filial": sede,
        "Área": area_letter,
        "Estado": attempt.get("state",""),
        "Comenzado el": datetime.fromtimestamp(ts, tz).strftime("%Y-%m-%d %H:%M:%S"),
        "Finalizado":   datetime.fromtimestamp(tf, tz).strftime("%Y-%m-%d %H:%M:%S"),
        "Tiempo requerido": f"{max(0, tf - ts)//60} min",
        "Calificación/20": grade,
        "_quizid": quiz["quizid"],
        "_courseid": quiz["courseid"],
        "_attemptid": attempt["id"],
    }
    # preguntas
    for i in range(1, max_questions+1):
        base[f"P. {i} /0.2"] = q_vals[i-1]

    # cálculos por área
    defs = AREA_DEFS.get(area_letter, {})
    ranges = defs.get("ranges", {})
    # percentajes por sub-área
    perc = {}
    pts  = {}
    for sub, qs in ranges.items():
        vals = [q_vals[i-1] for i in qs if 1 <= i <= max_questions]
        correct = count_correct(vals)
        total_q = len(qs)
        perc[sub] = pct(correct, total_q)
        pts[sub]  = perc[sub] * total_q * 0.2

    # CTA/CCSS label ya normalizado en defs
    base["% COMUNICACIÓN"] = perc.get("COMUNICACIÓN", 0.0)
    base["% HABILIDADES COMUNICATIVAS"] = perc.get("HABILIDADES COMUNICATIVAS", 0.0)
    base["% MATEMÁTICA"] = perc.get("MATEMÁTICA", 0.0)
    base["% CTA/CCSS"] = perc.get("CTA/CCSS", 0.0)

    base["P. COMUNICACIÓN"] = pts.get("COMUNICACIÓN", 0.0)
    base["P. HABILIDADES COMUNICATIVAS"] = pts.get("HABILIDADES COMUNICATIVAS", 0.0)
    base["P. MATEMÁTICA"] = pts.get("MATEMÁTICA", 0.0)
    base["P. CTA/CCSS"] = pts.get("CTA/CCSS", 0.0)

    base["PUNTAJE"] = (
        base["P. COMUNICACIÓN"] +
        base["P. HABILIDADES COMUNICATIVAS"] +
        base["P. MATEMÁTICA"] +
        base["P. CTA/CCSS"]
    )

    responded = count_responded(q_vals)
    base["PREGUNTAS RESPONDIDAS"] = responded
    base["PREGUNTAS NO RESPONDIDAS"] = 100 - responded
    base["%DE PREGUNTAS RESPONDIDAS"] = responded / 100.0
    base["%DE PREGUNTAS NO RESPONDIDAS"] = (100 - responded) / 100.0

    return base

# -------- CRITERIOS por Área (para RESUMEN) --------
CRITERIA_BY_AREA = {
    "A": {  # INGENIERÍAS
        "COMUNICACIÓN": 25,
        "HABILIDADES COMUNICATIVAS": 10,
        "MATEMÁTICA": 50,
        "CTA/CCSS": 15,
    },
    "B": {  # CIENCIAS SALUD
        "COMUNICACIÓN": 25,
        "HABILIDADES COMUNICATIVAS": 10,
        "MATEMÁTICA": 30,
        "CTA/CCSS": 35,
    },
    "C": {  # CIENCIAS HUMANAS
        "COMUNICACIÓN": 35,
        "HABILIDADES COMUNICATIVAS": 10,
        "MATEMÁTICA": 30,
        "CTA/CCSS": 25,
    },
}

def write_excel_all_in_one(
    out_path: Path,
    rows: List[Dict[str, Any]],
    criteria_by_area: Dict[str, Dict[str, float]] = None,
    nivel_threshold_base: float = None,
    nivel_by_area: Dict[str, Dict[str, float]] = None,
):
    """
    Genera el Excel con hojas:
      - RESULTADOS
      - RESUMEN  (incluye CONDICIÓN + NIVELACIÓN por curso)

    criteria_by_area: permite sobreescribir los criterios por área (A/B/C)
    nivel_threshold_base: umbral decimal base (0.30 = 30%) cuando no hay umbral específico
    nivel_by_area: umbral de nivelación por área y curso, en decimal (0.30 = 30%)
    """
    if not rows:
        raise RuntimeError("No hay filas para exportar.")

    # Valores por defecto cuando no se pasan desde la UI/CLI
    if criteria_by_area is None:
        criteria_by_area = CRITERIA_BY_AREA
    if nivel_threshold_base is None:
        nivel_threshold_base = NIVEL_THRESHOLD  # 0.30 por defecto

    df = pd.DataFrame(rows)

    # Orden: Programa Académico asc, luego DNI asc
    sort_cols = []
    if "Programa Académico" in df.columns:
        sort_cols.append("Programa Académico")
    if "Numero de DNI" in df.columns:
        sort_cols.append("Numero de DNI")
    if sort_cols:
        df = df.sort_values(by=sort_cols, kind="mergesort").reset_index(drop=True)

    # ---------- RESUMEN (Puntaje | CRITERIO | % + NIVELACIÓN) ----------
    resumen_rows = []
    for _, r in df.iterrows():
        area = str(r.get("Área", "")).upper().strip()
        crit = criteria_by_area.get(area, {})

        # % ya calculados en RESULTADOS (0.0–1.0)
        pct_com = float(r.get("% COMUNICACIÓN", 0) or 0)
        pct_hab = float(r.get("% HABILIDADES COMUNICATIVAS", 0) or 0)
        pct_mat = float(r.get("% MATEMÁTICA", 0) or 0)
        pct_cta = float(r.get("% CTA/CCSS", 0) or 0)

        # CRITERIOS por área (por fila)
        c_com = crit.get("COMUNICACIÓN", 0)
        c_hab = crit.get("HABILIDADES COMUNICATIVAS", 0)
        c_mat = crit.get("MATEMÁTICA", 0)
        c_cta = crit.get("CTA/CCSS", 0)

        # Puntajes = CRITERIO × %
        p_com = c_com * pct_com
        p_hab = c_hab * pct_hab
        p_mat = c_mat * pct_mat
        p_cta = c_cta * pct_cta

        total = p_com + p_hab + p_mat + p_cta

        # ---------- CONDICIÓN ----------
        # INGRESÓ si la nota (Calificación/20) es > 0
        grade = float(r.get("Calificación/20", 0) or 0)
        condicion = "INGRESÓ" if grade > 0 else ""

        # ---------- UMBRALES DE NIVELACIÓN POR CURSO ----------
        # Si no se pasó nada desde la UI, usamos el umbral base (todo 0.30)
        area_niveles = (nivel_by_area or {}).get(area, {})
        thr_com = area_niveles.get("COMUNICACIÓN", nivel_threshold_base)
        thr_hab = area_niveles.get("HABILIDADES COMUNICATIVAS", nivel_threshold_base)
        thr_mat = area_niveles.get("MATEMÁTICA", nivel_threshold_base)
        thr_cta = area_niveles.get("CTA/CCSS", nivel_threshold_base)

        # ---------- NIVELACIÓN POR CURSO ----------
        # Regla: si % obtenido es <= umbral ⇒ va a nivelación
        com_nivel = "COMUNICACIÓN" if pct_com <= thr_com else ""
        hab_nivel = "HABILIDADES COMUNICATIVAS" if pct_hab <= thr_hab else ""
        mat_nivel = "MATEMATICA" if pct_mat <= thr_mat else ""  # nombre de columna en plantilla

        cta_nivel = ""
        ccss_nivel = ""
        if pct_cta <= thr_cta:
            # Para Área C es CCSS (Ciencias Sociales), para A y B es CTA
            if area == "C":
                ccss_nivel = "CIENCIAS SOCIALES"
            else:
                cta_nivel = "CIENCIA, TECNOLOGÍA Y AMBIENTE"

        # Si al menos un curso requiere nivelación, marcamos el programa
        requiere_nivel = any([com_nivel, hab_nivel, mat_nivel, cta_nivel, ccss_nivel])
        programa_nivel = "SI" if requiere_nivel else "NO"

        resumen_rows.append({
            "Apellidos y nombres": f"{r.get('Apellido(s)','')} {r.get('Nombre','')}".strip(),
            "DNI": r.get("Numero de DNI",""),
            "Código de Matrícula": r.get("Código de Matrícula",""),
            "Programa Académico": r.get("Programa Académico",""),
            "Sede o Filial": r.get("Sede o Filial",""),
            "Área": area,
            "Asistencia": "ASISTIÓ",

            # COM
            "COMUNICACIÓN": p_com,
            "CRITERIO (COM)": c_com,
            "% (COM)": pct_com,

            # HAB
            "HABILIDADES COMUNICATIVAS": p_hab,
            "CRITERIO (HAB)": c_hab,
            "% (HAB)": pct_hab,

            # MAT
            "MATEMÁTICA": p_mat,
            "CRITERIO (MAT)": c_mat,
            "% (MAT)": pct_mat,

            # CTA/CCSS
            "CTA/CCSS": p_cta,
            "CRITERIO (CTA/CCSS)": c_cta,
            "% (CTA/CCSS)": pct_cta,

            "TOTAL": total,
            "%_TOTAL": total / 100.0,

            "PREGUNTAS RESPONDIDAS": r.get("PREGUNTAS RESPONDIDAS", 0),
            "PREGUNTAS NO RESPONDIDAS": r.get("PREGUNTAS NO RESPONDIDAS", 0),
            "% RESPONDIDAS": r.get("%DE PREGUNTAS RESPONDIDAS", 0.0),
            "% NO RESPONDIDAS": r.get("%DE PREGUNTAS NO RESPONDIDAS", 0.0),

            # columnas extra de la plantilla
            "CONDICIÓN": condicion,
            "PROGRAMA DE NIVELACIÓN": programa_nivel,
            "COMUNICACIÓN.1": com_nivel,
            "HABILIDADES COMUNICATIVAS.1": hab_nivel,
            "MATEMATICA": mat_nivel,
            "CIENCIA, TECNOLOGÍA Y AMBIENTE.1": cta_nivel,
            "CIENCIAS SOCIALES": ccss_nivel,
        })

    df_res = pd.DataFrame(resumen_rows)

    # Orden RESUMEN: Programa, DNI
    if {"Programa Académico","DNI"}.issubset(df_res.columns):
        df_res = df_res.sort_values(by=["Programa Académico","DNI"], kind="mergesort").reset_index(drop=True)

    # Columnas ordenadas
    ordered_cols = [
        "Apellidos y nombres","DNI","Código de Matrícula","Programa Académico","Sede o Filial","Área","Asistencia",

        "COMUNICACIÓN","CRITERIO (COM)","% (COM)",
        "HABILIDADES COMUNICATIVAS","CRITERIO (HAB)","% (HAB)",
        "MATEMÁTICA","CRITERIO (MAT)","% (MAT)",
        "CTA/CCSS","CRITERIO (CTA/CCSS)","% (CTA/CCSS)",

        "TOTAL","%_TOTAL",
        "PREGUNTAS RESPONDIDAS","PREGUNTAS NO RESPONDIDAS","% RESPONDIDAS","% NO RESPONDIDAS",

        "CONDICIÓN",
        "PROGRAMA DE NIVELACIÓN",
        "COMUNICACIÓN.1",
        "HABILIDADES COMUNICATIVAS.1",
        "MATEMATICA",
        "CIENCIA, TECNOLOGÍA Y AMBIENTE.1",
        "CIENCIAS SOCIALES",
    ]
    ordered_cols = [c for c in ordered_cols if c in df_res.columns] + [c for c in df_res.columns if c not in ordered_cols]

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="RESULTADOS", index=False)
        df_res[ordered_cols].to_excel(writer, sheet_name="RESUMEN", index=False)

    return out_path





# -------- CLI / Main --------
def parse_quiz_map(s: str) -> Dict[int, str]:
    """
    Ej: "11907=A,11908=B,11909=C"  -> {11907:'A', 11908:'B', 11909:'C'}
    """
    out = {}
    if not s:
        return out
    parts = [p for p in s.split(",") if p.strip()]
    for p in parts:
        if "=" not in p:
            continue
        left, right = p.split("=", 1)
        left = left.strip(); right = right.strip().upper()
        if left.isdigit() and right in ("A","B","C"):
            out[int(left)] = right
    return out

def default_downloads_path(filename: str) -> Path:
    # Windows
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    dl = home / "Downloads" / filename
    return dl

def main():
    ap = argparse.ArgumentParser(description="Exportador Admisión — TODO EN UNO (RESULTADOS+RESUMEN)")
    ap.add_argument("--base-url", required=True, help="https://tu-moodle")
    ap.add_argument("--token", required=True, help="Token WS")
    ap.add_argument("--course-ids", required=True, help="IDs de curso separados por coma (ej. 11989 o 100,101)")
    ap.add_argument("--quiz-map", required=True, help='Mapeo quiz→Área, ej.: 11907=A,11908=B,11909=C')
    ap.add_argument("--date", required=True, help="Día del examen (YYYY-MM-DD) hora local --tz-offset")
    ap.add_argument("--tz-offset", default="-05:00", help="Offset tz local, ej. -05:00")
    ap.add_argument("--workers", type=int, default=16, help="Hilos paralelos (16 recomendado)")
    ap.add_argument("--only-roles", default="student", help="Filtrar roles (ej. 'student'; multiple con coma)")
    ap.add_argument("--salida", default="", help="Ruta del Excel de salida (por defecto en Descargas)")
    args = ap.parse_args()

    base_url = args.base_url
    token = args.token
    course_ids = [int(x) for x in args.course_ids.split(",") if x.strip()]
    quiz_map = parse_quiz_map(args.quiz_map)
    only_roles = [x.strip() for x in args.only_roles.split(",") if x.strip()]
    t_from, t_to, tz = day_range_epoch(args.date, args.tz_offset)

    if not quiz_map:
        raise SystemExit("Debes proveer --quiz-map, ej.: 11907=A,11908=B,11909=C")

    # salida por defecto a Descargas
    if not args.salida:
        args.salida = f"RESULTADOS_ADMISION_{args.date}.xlsx"
    out_path = Path(args.salida)
    if not out_path.is_absolute():
        out_path = default_downloads_path(args.salida)

    print(f"[INFO] Cursos: {course_ids} | Día: {args.date} (tz {args.tz_offset})")
    print(f"[INFO] Ventana epoch: {t_from} → {t_to}")
    print(f"[INFO] Quiz→Área: {quiz_map}")

    # Descubrir quizzes y filtrar a los de --quiz-map
    quizzes = discover_quizzes(base_url, token, course_ids)
    qids_in_cursos = {q["quizid"] for q in quizzes}
    target_qids = [qid for qid in quiz_map.keys() if qid in qids_in_cursos]
    if not target_qids:
        print("[WARN] Ninguno de los quiz en --quiz-map aparece en los cursos dados.")
    target_quizzes = [q for q in quizzes if q["quizid"] in target_qids]
    print(f"[INFO] Quizzes a procesar ({len(target_quizzes)}): " + ", ".join(f"{q['quizname']}[{q['quizid']}]" for q in target_quizzes))

    # Usuarios por curso (filtrando roles)
    course_users: Dict[int, List[Dict[str,Any]]] = {}
    total_users = 0
    for cid in course_ids:
        us = get_course_users(base_url, token, cid, only_roles=only_roles)
        course_users[cid] = us
        total_users += len(us)
        print(f"[INFO] Usuarios en curso {cid}: {len(us)}")

    if total_users == 0 or not target_quizzes:
        print("[INFO] Nada para procesar.")
        return

    # Procesamiento paralelo (por (quiz, user))
    rows: List[Dict[str,Any]] = []
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = []
        for q in target_quizzes:
            area_letter = quiz_map.get(q["quizid"])
            users = course_users.get(q["courseid"], [])
            for u in users:
                futs.append(ex.submit(_process_user_quiz, base_url, token, q, area_letter, u, t_from, t_to, tz))
        done = 0
        for fut in as_completed(futs):
            res = fut.result()
            if res:
                rows.extend(res)
            done += 1
            if done % 100 == 0:
                print(f"[INFO] pares (quiz,usuario) procesados: {done}/{len(futs)}")

    print(f"[INFO] Intentos dentro del día: {len(rows)}")
    if not rows:
        print("No se encontraron intentos ese día.")
        return

    out = write_excel_all_in_one(out_path, rows)
    print(f"Archivo generado: {out} (filas: {len(rows)})")
    print(f"[INFO] Tiempo total: {time.time()-t0:.1f} s")

def _process_user_quiz(base_url: str, token: str, quiz: Dict[str,Any], area_letter: str,
                       user: Dict[str,Any], t_from: int, t_to: int, tz: timezone) -> List[Dict[str,Any]]:
    out = []
    try:
        attempts = get_user_attempts_in_range(base_url, token, quiz["quizid"], user["id"], t_from, t_to)
        for a in attempts:
            review = get_attempt_review(base_url, token, a["id"])
            out.append(build_row_from_review(user, quiz, area_letter, a, review, tz))
    except Exception as e:
        # para robustez, no detenemos todo por un usuario
        print(f"[WARN] usuario {user.get('id')} quiz {quiz['quizid']}: {e}")
    return out

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Cancelado por el usuario (Ctrl+C). Saliendo…")
