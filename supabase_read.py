"""Lectura de datos de Supabase.

`leer_insumos` sigue usando la API REST (PostgREST) con la anon key --
esa parte no esta afectada por el problema descrito abajo.

`leer_gastos_pendientes` y `marcar_gastos_sincronizados` se conectan
DIRECTO a Postgres (via psycopg2), en vez de pasar por la API REST de
Supabase con la service_role/secret key. Esto es a proposito: hubo un
periodo largo (agosto-septiembre 2026) en que un bug de la plataforma
de Supabase ("stale time cache" en la verificacion de JWT del gateway
REST) devolvia 403 Forbidden en TODAS las peticiones autenticadas con
service_role/secret key a este proyecto, sin importar que la clave
fuera valida y estuviera bien configurada. La base de datos Postgres
en si nunca estuvo afectada -- solo la capa de verificacion de tokens
de la API REST. Conectando directo a Postgres se evita el problema
por completo."""

from __future__ import annotations

import requests

import psycopg2
import psycopg2.extras

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
_SQL_GASTOS_PENDIENTES = """
    select
        g.id,
        g.presupuesto_id,
        g.fecha,
        g.capitulo,
        g.valor_total,
        g.valor_iva,
        g.proveedor,
        g.descripcion,
        g.foto_path,
        tg.nombre as tipo_gasto_nombre,
        tg.bucket_excel as bucket_excel,
        u.nombre as usuario_nombre
    from gastos g
    left join tipos_gasto tg on tg.codigo = g.tipo_gasto
    left join usuarios u on u.id = g.usuario_id
    where g.sincronizado_excel = false
    order by g.creado_en asc
    limit 500
"""

def leer_gastos_pendientes(db_url: str) -> list:
    """Lee gastos con sincronizado_excel=false, con el nombre del tipo
    de gasto y del usuario ya resueltos (para no tener que consultar
    otras tablas desde el script de sincronizacion).
    OJO: se conecta directo a Postgres (db_url = connection string,
    NO la service_role key) -- ver el docstring del modulo arriba
    para el porque. db_url se guarda SOLO como secret de GitHub
    (SUPABASE_DB_URL), nunca en el repositorio ni en la app de
    Streamlit."""
    with psycopg2.connect(db_url) as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(_SQL_GASTOS_PENDIENTES)
            filas = cur.fetchall()

    out = []
    for f in filas:
        valor_total = float(f["valor_total"])
        valor_iva = float(f["valor_iva"] or 0)
        out.append(
            {
                "id": str(f["id"]),
                "presupuesto_id": str(f["presupuesto_id"]),
                "fecha": f["fecha"].isoformat() if f["fecha"] else None,
                "capitulo": f.get("capitulo"),
                "valor_base": valor_total - valor_iva,
                "valor_iva": valor_iva,
                "proveedor": f.get("proveedor"),
                "descripcion": f.get("descripcion"),
                "foto_path": f.get("foto_path"),
                "tipo_gasto_nombre": f.get("tipo_gasto_nombre") or "",
                "bucket_excel": f.get("bucket_excel") or "EJECUCION",
                "usuario_nombre": f.get("usuario_nombre") or "",
            }
        )
    return out
def marcar_gastos_sincronizados(db_url: str, ids: list) -> None:
    if not ids:
        return
    with psycopg2.connect(db_url) as conn:
        with conn.cursor() as cur:
            cur.execute(
                "update gastos set sincronizado_excel = true where id = any(%s::uuid[])",
                (ids,),
            )
        conn.commit()
