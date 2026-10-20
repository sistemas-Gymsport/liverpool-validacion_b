import hmac
import secrets
import threading
import time
from functools import wraps

from flask import current_app, session
from werkzeug.security import check_password_hash, generate_password_hash

from api.errores import ErrorApi

_HASH_REFERENCIA = generate_password_hash(secrets.token_urlsafe(16), method="pbkdf2:sha256")


def credenciales_validas(usuario: str, password: str, usuario_esperado: str, hash_esperado: str) -> bool:
    usuario_ok = hmac.compare_digest(usuario.encode("utf-8"), usuario_esperado.encode("utf-8"))
    hash_a_usar = hash_esperado if usuario_ok else _HASH_REFERENCIA
    try:
        password_ok = check_password_hash(hash_a_usar, password)
    except ValueError:
        password_ok = False
    return usuario_ok and password_ok


class SesionesRevocadas:
    """Identificadores de sesión cerrados antes de expirar; impide reutilizar una cookie tras el logout."""

    def __init__(self, duracion_segundos: float, reloj=time.time):
        self.duracion = duracion_segundos
        self.reloj = reloj
        self._revocadas: dict[str, float] = {}
        self._candado = threading.Lock()

    def revocar(self, sid: str) -> None:
        with self._candado:
            self._limpiar()
            self._revocadas[sid] = self.reloj() + self.duracion

    def esta_revocada(self, sid: str) -> bool:
        with self._candado:
            self._limpiar()
            return sid in self._revocadas

    def _limpiar(self) -> None:
        ahora = self.reloj()
        for sid in [s for s, vence in self._revocadas.items() if vence <= ahora]:
            del self._revocadas[sid]


def iniciar_sesion(usuario: str) -> None:
    session.clear()
    session.permanent = True
    session["usuario"] = usuario
    session["sid"] = secrets.token_urlsafe(24)
    session["inicio"] = time.time()


def sesion_actual() -> dict | None:
    sid = session.get("sid")
    inicio = session.get("inicio")
    if not sid or not isinstance(inicio, (int, float)):
        return None
    estado = current_app.extensions["verificador"]
    if time.time() - inicio > estado.config.duracion_sesion.total_seconds():
        return None
    if estado.revocadas.esta_revocada(sid):
        return None
    return {"usuario": session.get("usuario"), "sid": sid, "expira": inicio + estado.config.duracion_sesion.total_seconds()}


def requiere_sesion(vista):
    @wraps(vista)
    def envoltura(*args, **kwargs):
        actual = sesion_actual()
        if actual is None:
            if session.get("sid"):
                session.clear()
            raise ErrorApi(401, "Tu sesión no es válida o expiró. Inicia sesión nuevamente.", "sesion_requerida")
        return vista(*args, sesion=actual, **kwargs)

    return envoltura
