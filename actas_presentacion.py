# actas_presentacion.py
# Genera ACTAS FINALES y CONSOLIDADOS dentro del MISMO Excel usando una PLANTILLA modelo
# - No rompe tu core
# - No usa nivelación
# - Deja MODALIDAD y MODALIDAD DE INGRESO en blanco
# Requisitos: pandas, openpyxl

from io import BytesIO
from datetime import datetime
import re
import pandas as pd
import openpyxl


def _norm_dni(v) -> str:
    s = "" if pd.isna(v) else str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return ""
    return digits.zfill(8)


def _sede_key(s: str) -> str:
    s = (s or "").upper()
    if "CHINCHA" in s:
        return "CHINCHA"
    if "ICA" in s:
        return "ICA"
    # fallback: si tu data trae otras sedes/filiales, puedes cambiar esta regla
    return "CHINCHA"


def _clear_sheet_from_row(ws, start_row: int, max_col: int):
    max_row = ws.max_row
    for r in range(start_row, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None


def _dump_df_values(ws, df: pd.DataFrame):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))


def build_excel_final_con_actas(
    modelo_path: str,
    generated_excel_bytes: bytes,
    exam_date: datetime,
    exam_label: str = "EXAMEN ORDINARIO",
    output_add_resultados_resumen: bool = True,
) -> bytes:
    """
    Recibe el Excel ya generado por tu core (RESULTADOS + RESUMEN)
    y devuelve un único Excel final basado en la plantilla del modelo,
    rellenando:
      - Acta_Final_Chincha / Acta_Final_Ica
      - Consolidado_Chincha / Consolidado_Ica
    y opcionalmente agregando hojas RESULTADOS/RESUMEN al mismo archivo final.

    NO usa nivelación.
    MODALIDAD y MODALIDAD DE INGRESO quedan en blanco.
    """

    # 1) leer Excel generado
    gen_xlsx = pd.ExcelFile(BytesIO(generated_excel_bytes))
    if "RESULTADOS" not in gen_xlsx.sheet_names or "RESUMEN" not in gen_xlsx.sheet_names:
        raise RuntimeError(f"El Excel generado no tiene RESULTADOS/RESUMEN. Hojas: {gen_xlsx.sheet_names}")

    df_res = pd.read_excel(gen_xlsx, sheet_name="RESULTADOS")
    df_sum = pd.read_excel(gen_xlsx, sheet_name="RESUMEN")

    # 2) validar columnas mínimas
    if "Numero de DNI" not in df_res.columns:
        raise RuntimeError("RESULTADOS no tiene la columna 'Numero de DNI'.")
    if "DNI" not in df_sum.columns:
        raise RuntimeError("RESUMEN no tiene la columna 'DNI'.")

    # 3) normalizar DNI para join
    df_res["_dni_norm"] = df_res["Numero de DNI"].apply(_norm_dni)
    df_sum["_dni_norm"] = df_sum["DNI"].apply(_norm_dni)

    # 4) armar df_small desde RESULTADOS
    # Nota: Código lo jalamos de 'Código de Matrícula' si existe
    cod_col = "Código de Matrícula" if "Código de Matrícula" in df_res.columns else None

    df_small = df_res[["_dni_norm"]].copy()
    df_small["APELLIDOS"] = df_res["Apellido(s)"] if "Apellido(s)" in df_res.columns else ""
    df_small["NOMBRES"] = df_res["Nombre"] if "Nombre" in df_res.columns else ""
    df_small["CORREO"] = df_res["Dirección de correo"] if "Dirección de correo" in df_res.columns else ""
    df_small["CODIGO"] = df_res[cod_col] if cod_col else ""

    base = df_sum.merge(df_small, on="_dni_norm", how="left")

    # validar columnas típicas del resumen (si alguna falta, se llenará vacío)
    # 'Sede o Filial' y 'Programa Académico' son claves para el acta
    if "Sede o Filial" not in base.columns:
        raise RuntimeError("RESUMEN no tiene la columna 'Sede o Filial'.")
    if "Programa Académico" not in base.columns:
        raise RuntimeError("RESUMEN no tiene la columna 'Programa Académico'.")

    base["SEDE_KEY"] = base["Sede o Filial"].astype(str).apply(_sede_key)

    # 5) abrir plantilla
    wb = openpyxl.load_workbook(modelo_path)

    sheets_acta = {
        "CHINCHA": "Acta_Final_Chincha",
        "ICA": "Acta_Final_Ica",
    }
    sheets_cons = {
        "CHINCHA": "Consolidado_Chincha",
        "ICA": "Consolidado_Ica",
    }

    for sname in list(sheets_acta.values()) + list(sheets_cons.values()):
        if sname not in wb.sheetnames:
            raise RuntimeError(f"El modelo no contiene la hoja requerida: {sname}")

    # 6) llenado por posiciones (respeta layout del modelo actual)
    # Columnas por posición (1-based):
    #  1 N°, 2 APELLIDOS, 3 NOMBRES, 4 DNI, 5 CODIGO, 6 TELEFONO, 7 CORREO,
    #  8 AREA, 9 CARRERA, 10 SEDE DE ESTUDIO, 11 ASISTENCIA,
    #  12 COM, 13 %COM, 14 HAB, 15 %HAB, 16 MAT, 17 %MAT, 18 CTA, 19 %CTA,
    #  20 TOTAL, 21 %TOTAL, 22 CONDICIÓN, 23 MODALIDAD, 24 FECHA, 25 EXAMEN,
    #  26 MODALIDAD DE INGRESO (solo en Acta_Final)
    def fill_ws(ws, df):
        _clear_sheet_from_row(ws, start_row=2, max_col=ws.max_column)

        # orden sugerido
        if "Programa Académico" in df.columns and "DNI" in df.columns:
            df = df.sort_values(["Programa Académico", "DNI"], kind="mergesort")

        for i in range(len(df)):
            rr = df.iloc[i]
            row_idx = i + 2

            ws.cell(row_idx, 1).value = i + 1
            ws.cell(row_idx, 2).value = rr.get("APELLIDOS", "")
            ws.cell(row_idx, 3).value = rr.get("NOMBRES", "")
            ws.cell(row_idx, 4).value = rr.get("DNI", "")
            ws.cell(row_idx, 5).value = rr.get("CODIGO", "") or ""
            ws.cell(row_idx, 6).value = ""  # TELEFONO (no viene)
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

            # MODALIDAD / MODALIDAD DE INGRESO en blanco
            ws.cell(row_idx, 23).value = ""
            ws.cell(row_idx, 24).value = exam_date.date()
            ws.cell(row_idx, 25).value = exam_label
            if ws.max_column >= 26:
                ws.cell(row_idx, 26).value = ""

    # 7) rellenar actas y consolidados
    for sede, sh in sheets_acta.items():
        ws = wb[sh]
        df_sede = base[base["SEDE_KEY"] == sede].copy()
        fill_ws(ws, df_sede)

    for sede, sh in sheets_cons.items():
        ws = wb[sh]
        df_sede = base[base["SEDE_KEY"] == sede].copy()
        fill_ws(ws, df_sede)

    # 8) (Opcional) agregar RESULTADOS y RESUMEN al workbook final
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
