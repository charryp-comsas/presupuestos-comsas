"""
Fase 7 -- Lectura automatica (OCR) de facturas/recibos de obra.

En vez de un OCR tradicional (que falla mucho con facturas manuscritas o
tirillas de mala calidad), esto le manda la foto a Claude (API de
Anthropic) y le pide que devuelva SOLO un JSON con los campos que la app
necesita. El usuario siempre ve el resultado antes de guardar y puede
corregir cualquier campo -- este modulo nunca guarda nada por si solo.

Requiere el secret ANTHROPIC_API_KEY (ver fase7_control_gastos/LEEME.txt).
"""

from __future__ import annotations

import base64
import json
import re

import streamlit as st
from anthropic import Anthropic

MODELO = "claude-sonnet-4-5"

PROMPT = """Estas viendo la foto de una factura o recibo de una compra de obra de construccion en Colombia.
Lee la imagen y devuelve UNICAMENTE un JSON (sin texto antes ni despues, sin bloque markdown) con esta forma exacta:

{
  "proveedor": "<nombre del proveedor/establecimiento tal como aparece, o null si no se lee>",
  "valor_total": <numero total pagado, en pesos colombianos, sin puntos ni comas, o null>,
  "valor_iva": <valor del IVA discriminado en la factura, en pesos, 0 si no aplica o no se ve, o null si no se puede determinar>,
  "fecha": "<fecha de la factura en formato AAAA-MM-DD si se lee, o null>",
  "descripcion": "<que se compro, resumido en pocas palabras>",
  "confianza": "<ALTA, MEDIA o BAJA, segun que tan seguro estas de los valores que leiste>"
}

Si la imagen esta borrosa, incompleta, o no es claramente una factura/recibo, igual devuelve el JSON con los
campos que SI puedas leer en null y "confianza": "BAJA". No inventes valores."""


def _extraer_json(texto: str) -> dict:
    texto = texto.strip()
    texto = re.sub(r"^```(json)?|```$", "", texto, flags=re.MULTILINE).strip()
    return json.loads(texto)


def leer_recibo(imagen_bytes: bytes, media_type: str = "image/jpeg") -> dict | None:
    """Devuelve un dict con proveedor/valor_total/valor_iva/fecha/descripcion/confianza,
    o None si la llamada fallo (sin excepcion visible al usuario -- se maneja arriba)."""
    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        st.warning("Falta el secret ANTHROPIC_API_KEY -- la lectura automatica esta desactivada. "
                   "Puedes llenar los campos a mano igual.")
        return None

    try:
        cliente = Anthropic(api_key=api_key)
        imagen_b64 = base64.standard_b64encode(imagen_bytes).decode("utf-8")
        resp = cliente.messages.create(
            model=MODELO,
            max_tokens=500,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {"type": "base64", "media_type": media_type, "data": imagen_b64},
                        },
                        {"type": "text", "text": PROMPT},
                    ],
                }
            ],
        )
        texto = resp.content[0].text
        datos = _extraer_json(texto)
        return datos
    except Exception as e:
        st.warning(f"No se pudo leer la foto automaticamente ({e}). Llena los campos a mano.")
        return None
