from flask import Blueprint, current_app, jsonify, request, session

from api.errores import ErrorApi
from security import csrf
from security.authentication import credenciales_validas, iniciar_sesion, sesion_actual

bp = Blueprint("auth", __name__, url_prefix="/api/auth")

MENSAJE_CREDENCIALES = "Usuario o contraseña incorrectos"


def _estado():
    return current_app.extensions["verificador"]


@bp.get("/csrf")
def obtener_csrf():
    return jsonify({"ok": True, "csrf": csrf.obtener_token()})


@bp.post("/login")
def login():
    estado = _estado()
    ip = request.remote_addr or "desconocida"
    espera = estado.limitador.segundos_restantes(ip)
    if espera:
        minutos = max(1, (espera + 59) // 60)
        raise ErrorApi(429, f"Demasiados intentos fallidos. Inténtalo de nuevo en {minutos} minutos.", "bloqueado",
                       reintentar_en=espera)

    datos = request.get_json(silent=True) or {}
    usuario = datos.get("usuario")
    password = datos.get("password")
    if not isinstance(usuario, str) or not isinstance(password, str) or not usuario or not password \
            or len(usuario) > 100 or len(password) > 200:
        raise ErrorApi(400, "Escribe tu usuario y contraseña.", "datos_incompletos")

    if not credenciales_validas(usuario.strip(), password, estado.config.usuario, estado.config.password_hash):
        estado.limitador.registrar_fallo(ip)
        raise ErrorApi(401, MENSAJE_CREDENCIALES, "credenciales")

    estado.limitador.registrar_exito(ip)
    iniciar_sesion(estado.config.usuario)
    actual = sesion_actual()
    return jsonify({"ok": True, "usuario": actual["usuario"], "expira": actual["expira"], "csrf": csrf.renovar_token()})


@bp.get("/session")
def consultar_sesion():
    actual = sesion_actual()
    if actual is None:
        if session.get("sid"):
            session.clear()
        return jsonify({"ok": False, "autenticado": False, "csrf": csrf.obtener_token()}), 401
    return jsonify({"ok": True, "autenticado": True, "usuario": actual["usuario"], "expira": actual["expira"],
                    "csrf": csrf.obtener_token()})


@bp.post("/logout")
def logout():
    estado = _estado()
    sid = session.get("sid")
    if sid:
        estado.reportes.eliminar_sesion(sid)
        estado.revocadas.revocar(sid)
    session.clear()
    return jsonify({"ok": True, "csrf": csrf.renovar_token()})
