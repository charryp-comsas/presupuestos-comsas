"""
Helpers minimos para leer/escribir un archivo en OneDrive personal via
Microsoft Graph, usando un refresh token de larga duracion (obtenido
una sola vez con get_refresh_token.py).

No usa msal para mantenerlo simple: son 3 llamadas HTTP directas.
"""

from __future__ import annotations

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
SCOPE = "Files.ReadWrite offline_access"


def renovar_access_token(client_id: str, refresh_token: str) -> dict:
    """Cambia el refresh_token guardado por un access_token fresco
    (dura ~1h) y devuelve tambien el refresh_token nuevo (Microsoft
    rota el refresh token en cada uso -- hay que volver a guardarlo)."""
    resp = requests.post(
        TOKEN_URL,
        data={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": SCOPE,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()  # incluye access_token y refresh_token nuevos


def descargar_archivo(access_token: str, ruta_onedrive: str) -> bytes:
    """ruta_onedrive: ruta relativa a la raiz de OneDrive, ej.
    'Escritorio/PRESUPUESTOS DE OBRA/06_Plantillas/BASE APU COSTOS 2026.xlsm'
    """
    url = f"{GRAPH}/me/drive/root:/{ruta_onedrive}:/content"
    resp = requests.get(url, headers={"Authorization": f"Bearer {access_token}"}, timeout=120)
    resp.raise_for_status()
    return resp.content


def subir_archivo_grande(access_token: str, ruta_onedrive: str, contenido: bytes, chunk_mb: int = 4):
    """Sube un archivo >4MB usando una sesion de carga por partes
    (obligatorio en Graph para archivos grandes; el xlsm real pesa
    ~30MB asi que la carga simple PUT no sirve)."""
    url_sesion = f"{GRAPH}/me/drive/root:/{ruta_onedrive}:/createUploadSession"
    resp = requests.post(
        url_sesion,
        headers={"Authorization": f"Bearer {access_token}"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    resp.raise_for_status()
    upload_url = resp.json()["uploadUrl"]

    tam = len(contenido)
    chunk = chunk_mb * 1024 * 1024
    inicio = 0
    while inicio < tam:
        fin = min(inicio + chunk, tam) - 1
        parte = contenido[inicio : fin + 1]
        headers = {
            "Content-Length": str(len(parte)),
            "Content-Range": f"bytes {inicio}-{fin}/{tam}",
        }
        r = requests.put(upload_url, headers=headers, data=parte, timeout=120)
        r.raise_for_status()
        inicio = fin + 1
    return True
