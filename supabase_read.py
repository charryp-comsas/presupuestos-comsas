"""Lectura minima de la tabla `insumos` en Supabase via REST (PostgREST),
sin depender de la libreria supabase-py (evita problemas de proxy en
algunos entornos y es mas facil de auditar)."""

from __future__ import annotations

import requests


def leer_insumos(supabase_url: str, anon_key: str) -> list:
    url = f"{supabase_url}/rest/v1/insumos"
    headers = {"apikey": anon_key, "Authorization": f"Bearer {anon_key}"}
    out = []
    offset = 0
    page = 1000
    while True:
        params = {
            "select": "codigo,precio,activo",
            "limit": str(page),
            "offset": str(offset),
        }
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        datos = r.json()
        out.extend(datos)
        if len(datos) < page:
            break
        offset += page
    # solo insumos activos (los inactivos, ej. MAT-1723, no se sincronizan)
    return [d for d in out if d.get("activo", True)]


def leer_gastos_pendientes(supabase_url: str, service_key: str) -> list:
    """Lee gastos con sincronizado_excel=false, con el nombre del tipo
    de gasto y del usuario ya resueltos (para no tener que consultar
    otras tablas desde el script de sincronizacion).

    OJO: usa la SERVICE ROLE key (no la anon key) porque este script
    corre sin sesion de usuario (GitHub Actions) y las tablas tienen
    RLS -- la anon key sin sesion no veria ninguna fila. La service
    role key se guarda SOLO como secret de GitHub, nunca en el
    repositorio ni en la app de Streamlit."""
    url = f"{supabase_url}/rest/v1/gastos"
    headers = {"apikey": service_key, "Authorization": f"Bearer {service_key}"}
    params = {
        "select": (
            "id,presupuesto_id,fecha,capitulo,valor_total,valor_iva,proveedor,"
            "descripcion,foto_path,"
            "tipos_gasto(nombre,bucket_excel),"
            "usuarios(nombre)"
        ),
        "sincronizado_excel": "eq.false",
        "order": "creado_en.asc",
        "limit": "500",
    }
    r = requests.get(url, headers=headers, params=params, timeout=30)
    r.raise_for_status()
    filas = r.json()

    out = []
    for f in filas:
        tg = f.get("tipos_gasto") or {}
        us = f.get("usuarios") or {}
        valor_total = float(f["valor_total"])
        valor_iva = float(f.get("valor_iva") or 0)
        out.append(
            {
                "id": f["id"],
                "presupuesto_id": f["presupuesto_id"],
                "fecha": f["fecha"],
                "capitulo": f.get("capitulo"),
                "valor_base": valor_total - valor_iva,
                "valor_iva": valor_iva,
                "proveedor": f.get("proveedor"),
                "descripcion": f.get("descripcion"),
                "foto_path": f.get("foto_path"),
                "tipo_gasto_nombre": tg.get("nombre", ""),
                "bucket_excel": tg.get("bucket_excel", "EJECUCION"),
                "usuario_nombre": us.get("nombre", ""),
            }
        )
    return out


def marcar_gastos_sincronizados(supabase_url: str, service_key: str, ids: list) -> None:
    if not ids:
        return
    url = f"{supabase_url}/rest/v1/gastos"
    headers = {
        "apikey": service_key,
        "Authorization": f"Bearer {service_key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }
    lista_ids = ",".join(ids)
    params = {"id": f"in.({lista_ids})"}
    r = requests.patch(url, headers=headers, params=params, json={"sincronizado_excel": True}, timeout=30)
    r.raise_for_status()
