"""
Fase 7 -- Nucleo de la sincronizacion Supabase (tabla `gastos`) -> hoja
GASTOS de "CONTROL DE COSTOS v2 - PLANTILLA.xlsm" (o el archivo que se
use por obra) en OneDrive.

Igual que sync_core.py (precios): NO toca formulas existentes. La hoja
GASTOS ya trae, precargadas hasta la fila 500, las formulas de IVA,
retenciones, NETO A PAGAR, etc. (dependen de columnas E..Q). Este
modulo solo escribe valores en las columnas de "dato" de la primera
fila libre, fila por fila, para cada gasto nuevo. No hace wb.save().

Decision importante: la hoja calcula el IVA (columna L) con una formula
que depende de si el NIT del proveedor (columna E) esta marcado como
responsable de IVA en la hoja PROVEEDORES. Los gastos que vienen de la
App (celular) NO traen NIT, solo el valor de IVA que la persona escribio
o que leyo el OCR de la factura -- por eso, SOLO para las filas que
vienen de la App, este modulo escribe BASE (K) e IVA (L) como valores
fijos (el dato real de la factura manda sobre el calculo generico).
El resto de columnas con formula (M, N, O, P, Q, X, AB..AG) se dejan
intactas -- ya vienen precargadas en la plantilla y calculan solas.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass

HOJA_GASTOS = "GASTOS"
FILA_INICIO_DATOS = 4
COL_ID_GASTO_APP = 38  # AL -- marcador para no duplicar al re-sincronizar
FILA_HEADER = 3

COL = {
    "FECHA": 2,
    "BUCKET": 3,
    "CAPITULO": 4,
    "PROVEEDOR": 6,
    "TIPO_DOC": 7,
    "CONCEPTO_RET": 9,
    "DESCRIPCION": 10,
    "BASE": 11,
    "IVA": 12,
    "ORIGEN": 22,
    "ESTADO_CONT": 25,
    "NOTAS": 26,
}


@dataclass
class GastoAplicado:
    gasto_id: str
    fila: int


def _asegurar_columna_marcador(ws):
    header = ws.cell(row=FILA_HEADER, column=COL_ID_GASTO_APP).value
    if header != "GASTO_ID_APP":
        ws.cell(row=FILA_HEADER, column=COL_ID_GASTO_APP).value = "GASTO_ID_APP"


def _ids_ya_sincronizados(ws) -> set:
    ids = set()
    fila = FILA_INICIO_DATOS
    while fila <= ws.max_row:
        v = ws.cell(row=fila, column=COL_ID_GASTO_APP).value
        if v:
            ids.add(str(v))
        fila += 1
    return ids


def _primera_fila_libre(ws) -> int:
    fila = FILA_INICIO_DATOS
    while fila <= ws.max_row:
        if ws.cell(row=fila, column=COL["BASE"]).value in (None, ""):
            return fila
        fila += 1
    return ws.max_row + 1


def sincronizar_gastos(wb, gastos: list) -> list[GastoAplicado]:
    """gastos: lista de dicts con id, fecha, bucket_excel, capitulo,
    proveedor, descripcion, valor_base, valor_iva, origen_nota.
    Devuelve los gastos realmente escritos (para marcar
    sincronizado_excel=true en Supabase despues)."""
    if HOJA_GASTOS not in wb.sheetnames:
        raise ValueError(f"No se encontro la hoja '{HOJA_GASTOS}' en el libro.")

    ws = wb[HOJA_GASTOS]
    _asegurar_columna_marcador(ws)
    ya_sincronizados = _ids_ya_sincronizados(ws)

    aplicados = []
    fila = _primera_fila_libre(ws)

    for g in gastos:
        gid = str(g["id"])
        if gid in ya_sincronizados:
            continue

        ws.cell(row=fila, column=COL["FECHA"]).value = g["fecha"]
        ws.cell(row=fila, column=COL["BUCKET"]).value = g.get("bucket_excel", "EJECUCION")
        ws.cell(row=fila, column=COL["CAPITULO"]).value = g.get("capitulo") or "GASTOS GENERALES"
        ws.cell(row=fila, column=COL["PROVEEDOR"]).value = g.get("proveedor") or "(sin proveedor)"
        ws.cell(row=fila, column=COL["TIPO_DOC"]).value = "DOCUMENTO SOPORTE"
        ws.cell(row=fila, column=COL["CONCEPTO_RET"]).value = "COMPRAS"
        ws.cell(row=fila, column=COL["DESCRIPCION"]).value = g.get("descripcion") or g.get("tipo_gasto_nombre", "")
        ws.cell(row=fila, column=COL["BASE"]).value = float(g["valor_base"])
        ws.cell(row=fila, column=COL["IVA"]).value = float(g.get("valor_iva") or 0)
        ws.cell(row=fila, column=COL["ORIGEN"]).value = "FOTO APP"
        ws.cell(row=fila, column=COL["ESTADO_CONT"]).value = "PENDIENTE"
        ws.cell(row=fila, column=COL["NOTAS"]).value = (
            f"Tipo: {g.get('tipo_gasto_nombre','')} · Registrado por: {g.get('usuario_nombre','')} "
            f"· Foto en Supabase: {g.get('foto_path','')}"
        )
        ws.cell(row=fila, column=COL_ID_GASTO_APP).value = gid

        aplicados.append(GastoAplicado(gasto_id=gid, fila=fila))
        fila += 1

    return aplicados
