import hmac
import secrets

from flask import request, session

CLAVE_SESION = "csrf"
ENCABEZADO = "X-CSRF-Token"
METODOS_SEGUROS = {"GET", "HEAD", "OPTIONS"}


def obtener_token() -> str:
    token = session.get(CLAVE_SESION)
    if not token:
        token = secrets.token_urlsafe(32)
        session[CLAVE_SESION] = token
    return token


def renovar_token() -> str:
    session[CLAVE_SESION] = secrets.token_urlsafe(32)
    return session[CLAVE_SESION]


def token_valido() -> bool:
    if request.method in METODOS_SEGUROS:
        return True
    esperado = session.get(CLAVE_SESION)
    recibido = request.headers.get(ENCABEZADO, "")
    return bool(esperado and recibido) and hmac.compare_digest(esperado, recibido)
