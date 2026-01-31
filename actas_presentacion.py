# actas_presentacion.py
from __future__ import annotations

from io import BytesIO
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


def _load_wb_from_bytes(xlsx_bytes: bytes) -> openpyxl.Workbook:
    bio = BytesIO(xlsx_bytes)
    bio.seek(0)
    return openpyxl.load_workbook(bio)


def _wb_to_bytes(wb: openpyxl.Workbook) -> bytes:
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


def _safe_sheet_name(wb: openpyxl.Workbook, desired: str) -> str:
    """Evita colisiones de nombres de hoja."""
    name = desired[:31]
    if name not in wb.sheetnames:
        return name
    i = 2
    while True:
        cand = f"{name[:28]}_{i}"  # respeta 31 chars
        if cand not in wb.sheetnames:
            return cand
        i += 1


def _copy_sheet_values_only(src_ws: openpyxl.worksheet.worksheet.Worksheet,
                            dst_ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Copia valores (no estilos) para que sea robusto."""
    max_row = src_ws.max_row
    max_col = src_ws.max_column
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            dst_ws.cell(row=r, column=c, value=src_ws.cell(row=r, column=c).value)


def build_excel_final_con_actas(
    modelo_path: str | None,
    generated_excel_bytes: bytes,
    exam_date: datetime,
    exam_label: str = "EXAMEN",
    output_add_resultados_resumen: bool = True,
) -> bytes:
    """
    ✅ NUEVO COMPORTAMIENTO (lo que tú quieres):
    - NO exige nombres fijos de hojas (NO Acta_Final_Chincha)
    - Siempre genera una hoja 'ACTAS' dentro del mismo Excel
    - Si modelo_path existe, COPIA todas sus hojas al final (sin importar nombres)
      (pero si no hay modelo, NO falla)
    """

    # 1) Abrimos el Excel generado (RESULTADOS + RESUMEN)
    wb_gen = _load_wb_from_bytes(generated_excel_bytes)

    # 2) Leemos RESULTADOS/RESUMEN como DataFrame (si existen)
    df_resumen = None
    df_resultados = None

    if "RESUMEN" in wb_gen.sheetnames:
        ws = wb_gen["RESUMEN"]
        data = ws.values
        cols = next(data)
        df_resumen = pd.DataFrame(data, columns=cols)

    if "RESULTADOS" in wb_gen.sheetnames:
        ws = wb_gen["RESULTADOS"]
        data = ws.values
        cols = next(data)
        df_resultados = pd.DataFrame(data, columns=cols)

    # 3) Crear SIEMPRE hoja ACTAS (auto-generada)
    actas_name = _safe_sheet_name(wb_gen, "ACTAS")
    ws_actas = wb_gen.create_sheet(actas_name)

    # Encabezado simple
    ws_actas["A1"] = "ACTAS - GENERADO AUTOMÁTICAMENTE"
    ws_actas["A2"] = "Fecha examen"
    ws_actas["B2"] = exam_date.strftime("%Y-%m-%d")
    ws_actas["A3"] = "Tipo"
    ws_actas["B3"] = exam_label

    # 4) Construcción del contenido de ACTAS
    #    ✅ Lo hacemos robusto: si hay RESUMEN, lo usamos como “base” de actas.
    #    Si no hay RESUMEN, intentamos con RESULTADOS.
    start_row = 5

    if df_resumen is not None and not df_resumen.empty:
        # Elegimos columnas típicas si existen; si no, ponemos lo que haya
        preferred = [
            "DNI",
            "Apellido(s)",
            "Nombre",
            "Área",
            "Programa Académico",
            "Sede o Filial",
            "TOTAL",
            "Asistencia",
            "CONDICIÓN",
            "PROGRAMA DE NIVELACIÓN",
        ]
        cols = [c for c in preferred if c in df_resumen.columns]
        if not cols:
            cols = list(df_resumen.columns)

        df_actas = df_resumen[cols].copy()

        # Título de tabla
        ws_actas[f"A{start_row}"] = "CONSOLIDADO (desde RESUMEN)"
        start_row += 1

        # Volcar DF
        for r_idx, row in enumerate(dataframe_to_rows(df_actas, index=False, header=True), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws_actas.cell(row=r_idx, column=c_idx, value=value)

    elif df_resultados is not None and not df_resultados.empty:
        # Si no hay RESUMEN, usamos RESULTADOS (puede ser más grande)
        ws_actas[f"A{start_row}"] = "CONSOLIDADO (desde RESULTADOS)"
        start_row += 1

        # Para no hacer un monstruo, ponemos columnas preferidas si existen
        preferred = [
            "Numero de DNI",
            "DNI",
            "Apellido(s)",
            "Nombre",
            "Área",
            "Quiz",
            "Puntaje",
            "Porcentaje",
            "Fecha intento",
        ]
        cols = [c for c in preferred if c in df_resultados.columns]
        if not cols:
            cols = list(df_resultados.columns)

        df_actas = df_resultados[cols].copy()
        for r_idx, row in enumerate(dataframe_to_rows(df_actas, index=False, header=True), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws_actas.cell(row=r_idx, column=c_idx, value=value)

    else:
        ws_actas[f"A{start_row}"] = "No se encontró RESUMEN ni RESULTADOS en el archivo generado."

    # 5) (Opcional) Copiar todas las hojas del modelo, sin exigir nombres fijos
    if modelo_path:
        try:
            wb_modelo = openpyxl.load_workbook(modelo_path)
            for sname in wb_modelo.sheetnames:
                src = wb_modelo[sname]
                dst_name = _safe_sheet_name(wb_gen, sname)
                dst = wb_gen.create_sheet(dst_name)
                _copy_sheet_values_only(src, dst)
        except Exception:
            # Si el modelo falla, NO rompemos la generación
            pass

    # 6) Retornar el Excel final
    return _wb_to_bytes(wb_gen)
