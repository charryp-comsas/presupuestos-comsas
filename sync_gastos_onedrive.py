"""
Fase 7 -- corre en GitHub Actions (ver .github/workflows/sync-gastos.yml).
Sincroniza los gastos nuevos (tabla `gastos` en Supabase, registrados
desde la App con foto) hacia la hoja GASTOS del archivo CONTROL DE
COSTOS de la obra activa, guardado en OneDrive. Sentido UNICO:
Supabase -> Excel (el Excel nunca escribe de vuelta a Supabase).

Variables de entorno requeridas (Secrets en GitHub):
  SUPABASE_URL
  SUPABASE_SERVICE_ROLE_KEY   -- OJO: service role, NO la anon key (ver
                                  supabase_read.py:leer_gastos_pendientes).
                                  Se saca en Supabase > Project Settings >
                                  API > "service_role" (secreta, nunca
                                  compartirla ni ponerla en la App).
  MS_CLIENT_ID / MS_REFRESH_TOKEN   -- los mismos que ya usa sync-precios.
  PRESUPUESTO_ID_ACTIVO       -- uuid del presupuesto/obra cuyo Excel se
                                  esta sincronizando (por ahora, UNA obra
                                  a la vez -- ver nota abajo).
  ONEDRIVE_RUTA_ARCHIVO_COSTOS -- ej: "Escritorio/PRESUPUESTOS DE OBRA/
                                  07_Control de Costos/CONTROL DE COSTOS
                                  v2 - PLANTILLA.xlsm"

Nota sobre varias obras: cuando haya mas de una obra con Control de
Costos activo al tiempo, este script hay que correrlo una vez por obra
(un job de GitHub Actions distinto por obra, cada uno con su propio
PRESUPUESTO_ID_ACTIVO y ONEDRIVE_RUTA_ARCHIVO_COSTOS), o extenderlo
para que reciba un mapa obra->archivo. Se deja simple mientras solo hay
una obra en control de costos.
"""

from __future__ import annotations

import io
import os
import sys

import openpyxl

from gastos_core import sincronizar_gastos
from graph_onedrive import descargar_archivo, renovar_access_token, subir_archivo_grande
from supabase_read import leer_gastos_pendientes, marcar_gastos_sincronizados


def main() -> int:
    supabase_url = os.environ["SUPABASE_URL"]
    service_key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    client_id = os.environ["MS_CLIENT_ID"]
    refresh_token = os.environ["MS_REFRESH_TOKEN"]
    presupuesto_id_activo = os.environ["PRESUPUESTO_ID_ACTIVO"]
    ruta_archivo = os.environ["ONEDRIVE_RUTA_ARCHIVO_COSTOS"]

    print("1) Renovando token de acceso a OneDrive...")
    tokens = renovar_access_token(client_id, refresh_token)
    access_token = tokens["access_token"]
    nuevo_refresh = tokens.get("refresh_token", refresh_token)

    print("2) Leyendo gastos pendientes de sincronizar desde Supabase...")
    todos = leer_gastos_pendientes(supabase_url, service_key)
    gastos = [g for g in todos if g["presupuesto_id"] == presupuesto_id_activo]
    print(f"   {len(gastos)} gastos pendientes para esta obra "
          f"({len(todos) - len(gastos)} pendientes de otra obra, se ignoran aqui).")

    if not gastos:
        print("Sin gastos nuevos que subir. Fin.")
        _reportar_refresh_token(nuevo_refresh)
        return 0

    print("3) Descargando el Excel de Control de Costos desde OneDrive...")
    contenido = descargar_archivo(access_token, ruta_archivo)
    print(f"   {len(contenido) / 1024:.0f} KB descargados.")

    wb = openpyxl.load_workbook(io.BytesIO(contenido), keep_vba=True)

    print("4) Escribiendo gastos nuevos en la hoja GASTOS...")
    aplicados = sincronizar_gastos(wb, gastos)
    print(f"   {len(aplicados)} filas nuevas escritas.")

    if not aplicados:
        print("Nada nuevo que escribir (todo ya estaba sincronizado). Fin.")
        _reportar_refresh_token(nuevo_refresh)
        return 0

    print("5) Guardando libro en memoria...")
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    print("6) Subiendo a OneDrive...")
    subir_archivo_grande(access_token, ruta_archivo, buffer.getvalue())
    print("   Listo. Archivo actualizado en OneDrive.")

    print("7) Marcando gastos como sincronizados en Supabase...")
    marcar_gastos_sincronizados(supabase_url, service_key, [a.gasto_id for a in aplicados])
    print("   Listo.")

    _reportar_refresh_token(nuevo_refresh)
    return 0


def _reportar_refresh_token(nuevo_refresh: str) -> None:
    print(f"::add-mask::{nuevo_refresh}")
    print(f"NUEVO_REFRESH_TOKEN={nuevo_refresh}")


if __name__ == "__main__":
    sys.exit(main())
