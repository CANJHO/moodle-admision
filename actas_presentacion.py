# actas_presentacion.py
# Genera hojas "ACTA" y "CONSOLIDADO" dentro del MISMO Excel usando una PLANTILLA modelo.
# - No reemplaza tu RESULTADOS/RESUMEN: las agrega al final (opcional)
# - No divide por CHINCHA/ICA: usa TODOS los registros en ambas hojas (ACTA y CONSOLIDADO)
# - Mantiene el formato del modelo (toma como base Acta_Final_* y Consolidado_* del archivo modelo)
#
# Requisitos: pandas, openpyxl

from __future__ import annotations

from io import BytesIO
from datetime import datetime
import pandas as pd
import openpyxl


# --------------------------
# Helpers
# --------------------------

def _norm_dni(v) -> str:
    s = "" if pd.isna(v) else str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits.zfill(8) if digits else ""


def _norm_text(s: str) -> str:
    s = str(s or "").strip().lower()
    import unicodedata
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return "".join(ch for ch in s if ch.isalnum())


def _find_col_flexible(df: pd.DataFrame, keyword_groups):
    cols = list(df.columns)
    norm_cols = {c: _norm_text(c) for c in cols}

    for group in keyword_groups:
        g = [_norm_text(x) for x in group]
        for c, nc in norm_cols.items():
            if all(k in nc for k in g):
                return c
    return None


def _clear_sheet_from_row(ws, start_row: int):
    """Limpia valores desde start_row hasta el max_row actual, respetando estilos del modelo."""
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(start_row, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None


def _dump_df_values(ws, df: pd.DataFrame):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))


def _pick_template_sheet(wb: openpyxl.Workbook, kind: str) -> str:
    """
    kind: 'acta' | 'consolidado'
    Devuelve el nombre de una hoja plantilla del modelo.
    Prioriza *_Chincha si existe, si no toma la primera que matchee.
    """
    names = wb.sheetnames

    if kind == "acta":
        preferred = [n for n in names if _norm_text(n) == _norm_text("Acta_Final_Chincha")]
        if preferred:
            return preferred[0]
        candidates = [n for n in names if "acta" in _norm_text(n)]
        if candidates:
            candidates.sort(key=lambda x: (0 if "actafinal" in _norm_text(x) else 1, x))
            return candidates[0]

    if kind == "consolidado":
        preferred = [n for n in names if _norm_text(n) == _norm_text("Consolidado_Chincha")]
        if preferred:
            return preferred[0]
        candidates = [n for n in names if "consolid" in _norm_text(n)]
        if candidates:
            candidates.sort(key=lambda x: (0 if "consolidado" in _norm_text(x) else 1, x))
            return candidates[0]

    raise RuntimeError(f"No pude detectar una hoja plantilla '{kind}' dentro del modelo. Hojas: {names}")


# --------------------------
# API principal
# --------------------------

def build_excel_final_con_actas(
    modelo_path: str,
    generated_excel_bytes: bytes,
    exam_date: datetime,
    exam_label: str = "EXAMEN ORDINARIO",
    output_add_resultados_resumen: bool = True,
) -> bytes:
    """
    Recibe el Excel ya generado por tu core (RESULTADOS + RESUMEN) como bytes
    y devuelve un único Excel final basado en la plantilla del modelo, rellenando:

      - Hoja "ACTA"
      - Hoja "CONSOLIDADO"

    Usando TODOS los registros (sin split por sede).

    Además, opcionalmente agrega RESULTADOS y RESUMEN dentro del mismo archivo final.
    """

    # 1) leer Excel generado
    gen_xlsx = pd.ExcelFile(BytesIO(generated_excel_bytes))
    if "RESULTADOS" not in gen_xlsx.sheet_names or "RESUMEN" not in gen_xlsx.sheet_names:
        raise RuntimeError(f"El Excel generado no tiene RESULTADOS/RESUMEN. Hojas: {gen_xlsx.sheet_names}")

    df_res = pd.read_excel(gen_xlsx, sheet_name="RESULTADOS")
    df_sum = pd.read_excel(gen_xlsx, sheet_name="RESUMEN")

    # 2) detectar columnas mínimas (flexible)
    col_dni_res = "Numero de DNI" if "Numero de DNI" in df_res.columns else _find_col_flexible(df_res, [["dni"], ["numero", "dni"]])
    col_dni_sum = "DNI" if "DNI" in df_sum.columns else _find_col_flexible(df_sum, [["dni"], ["numero", "dni"]])

    if not col_dni_res:
        raise RuntimeError(f"RESULTADOS no tiene columna DNI reconocible. Columnas: {list(df_res.columns)}")
    if not col_dni_sum:
        raise RuntimeError(f"RESUMEN no tiene columna DNI reconocible. Columnas: {list(df_sum.columns)}")

    # 3) normalizar DNI para join
    df_res["_dni_norm"] = df_res[col_dni_res].apply(_norm_dni)
    df_sum["_dni_norm"] = df_sum[col_dni_sum].apply(_norm_dni)

    # 4) armar df_small desde RESULTADOS (apellidos/nombres/correo/codigo)
    col_ap = "Apellido(s)" if "Apellido(s)" in df_res.columns else _find_col_flexible(df_res, [["apell"], ["apellido"]])
    col_nom = "Nombre" if "Nombre" in df_res.columns else _find_col_flexible(df_res, [["nomb"], ["nombre"]])
    col_mail = "Dirección de correo" if "Dirección de correo" in df_res.columns else _find_col_flexible(df_res, [["correo"], ["mail"], ["email"]])

    # código: matrícula o estudiante (cualquiera)
    col_cod = None
    for exact in [
        "Código de Matrícula", "Codigo de Matricula", "CÓDIGO DE MATRÍCULA", "CODIGO DE MATRICULA",
        "Código de Estudiante", "Codigo de Estudiante", "CÓDIGO DE ESTUDIANTE", "CODIGO DE ESTUDIANTE",
    ]:
        if exact in df_res.columns:
            col_cod = exact
            break
    if not col_cod:
        col_cod = _find_col_flexible(df_res, [["codigo", "matricula"], ["codigo", "estudiante"], ["cod", "matr"], ["cod", "estud"], ["codigo"]])

    df_small = df_res[["_dni_norm"]].copy()
    df_small["APELLIDOS"] = df_res[col_ap] if col_ap and col_ap in df_res.columns else ""
    df_small["NOMBRES"] = df_res[col_nom] if col_nom and col_nom in df_res.columns else ""
    df_small["CORREO"] = df_res[col_mail] if col_mail and col_mail in df_res.columns else ""
    df_small["CODIGO"] = df_res[col_cod] if col_cod and col_cod in df_res.columns else ""

    base = df_sum.merge(df_small, on="_dni_norm", how="left")

    # 5) abrir plantilla
    wb = openpyxl.load_workbook(modelo_path)

    # 6) seleccionar hojas plantilla del modelo y clonarlas como ACTA / CONSOLIDADO
    acta_tpl_name = _pick_template_sheet(wb, "acta")
    cons_tpl_name = _pick_template_sheet(wb, "consolidado")

    ws_acta_tpl = wb[acta_tpl_name]
    ws_cons_tpl = wb[cons_tpl_name]

    # Si ya existen ACTA/CONSOLIDADO por ejecuciones previas, borrarlas
    for fixed in ["ACTA", "CONSOLIDADO"]:
        if fixed in wb.sheetnames:
            del wb[fixed]

    ws_acta = wb.copy_worksheet(ws_acta_tpl)
    ws_acta.title = "ACTA"

    ws_cons = wb.copy_worksheet(ws_cons_tpl)
    ws_cons.title = "CONSOLIDADO"

    # Eliminar hojas acta/consolidado originales para que no salgan duplicadas
    for n in list(wb.sheetnames):
        nn = _norm_text(n)
        if nn.startswith(_norm_text("acta_final")) or nn.startswith(_norm_text("consolidado")):
            if n not in ("ACTA", "CONSOLIDADO"):
                del wb[n]

    # 7) llenado (misma estructura para ACTA y CONSOLIDADO)
    def fill_ws(ws, df: pd.DataFrame):
        _clear_sheet_from_row(ws, start_row=2)

        sort_cols = []
        if "Programa Académico" in df.columns:
            sort_cols.append("Programa Académico")
        if "DNI" in df.columns:
            sort_cols.append("DNI")
        if sort_cols:
            df = df.sort_values(sort_cols, kind="mergesort")

        for i in range(len(df)):
            rr = df.iloc[i]
            row_idx = i + 2

            ws.cell(row_idx, 1).value = i + 1
            ws.cell(row_idx, 2).value = rr.get("APELLIDOS", "")
            ws.cell(row_idx, 3).value = rr.get("NOMBRES", "")
            ws.cell(row_idx, 4).value = rr.get("DNI", "")
            ws.cell(row_idx, 5).value = rr.get("CODIGO", "") or ""
            ws.cell(row_idx, 6).value = ""  # TELEFONO
            ws.cell(row_idx, 7).value = rr.get("CORREO", "") or ""
            ws.cell(row_idx, 8).value = rr.get("Área", "") if "Área" in rr.index else ""
            ws.cell(row_idx, 9).value = rr.get("Programa Académico", "") or ""
            ws.cell(row_idx, 10).value = rr.get("Sede o Filial", "") or ""
            ws.cell(row_idx, 11).value = rr.get("Asistencia", "") if "Asistencia" in rr.index else ""

            ws.cell(row_idx, 12).value = rr.get("COMUNICACIÓN", "") if "COMUNICACIÓN" in rr.index else ""
            ws.cell(row_idx, 13).value = rr.get("% (COM)", "") if "% (COM)" in rr.index else ""
            ws.cell(row_idx, 14).value = rr.get("HABILIDADES COMUNICATIVAS", "") if "HABILIDADES COMUNICATIVAS" in rr.index else ""
            ws.cell(row_idx, 15).value = rr.get("% (HAB)", "") if "% (HAB)" in rr.index else ""
            ws.cell(row_idx, 16).value = rr.get("MATEMÁTICA", "") if "MATEMÁTICA" in rr.index else ""
            ws.cell(row_idx, 17).value = rr.get("% (MAT)", "") if "% (MAT)" in rr.index else ""
            ws.cell(row_idx, 18).value = rr.get("CTA/CCSS", "") if "CTA/CCSS" in rr.index else ""
            ws.cell(row_idx, 19).value = rr.get("% (CTA/CCSS)", "") if "% (CTA/CCSS)" in rr.index else ""
            ws.cell(row_idx, 20).value = rr.get("TOTAL", "") if "TOTAL" in rr.index else ""
            ws.cell(row_idx, 21).value = rr.get("%_TOTAL", "") if "%_TOTAL" in rr.index else ""
            ws.cell(row_idx, 22).value = rr.get("CONDICIÓN", "") if "CONDICIÓN" in rr.index else ""

            ws.cell(row_idx, 23).value = ""  # MODALIDAD
            ws.cell(row_idx, 24).value = exam_date.date()
            ws.cell(row_idx, 25).value = exam_label
            if ws.max_column >= 26:
                ws.cell(row_idx, 26).value = ""  # MODALIDAD DE INGRESO

    fill_ws(ws_acta, base.copy())
    fill_ws(ws_cons, base.copy())

    # 8) (Opcional) agregar RESULTADOS y RESUMEN
    if output_add_resultados_resumen:
        if "RESULTADOS" in wb.sheetnames:
            del wb["RESULTADOS"]
        if "RESUMEN" in wb.sheetnames:
            del wb["RESUMEN"]

        ws_r = wb.create_sheet("RESULTADOS", 0)
        ws_s = wb.create_sheet("RESUMEN", 1)

        _dump_df_values(ws_r, df_res.drop(columns=["_dni_norm"], errors="ignore"))
        _dump_df_values(ws_s, df_sum.drop(columns=["_dni_norm"], errors="ignore"))

    # 9) guardar a bytes
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
