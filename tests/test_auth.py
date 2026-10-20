import time

from werkzeug.security import check_password_hash

import crear_usuario
from conftest import PASSWORD, USUARIO, Sesion, configuracion
from factory import create_app
from security.authentication import credenciales_validas
from security.rate_limit import LimitadorLogin


def test_login_correcto(cliente):
    s = Sesion(cliente)
    respuesta = s.login()
    assert respuesta.status_code == 200
    assert respuesta.get_json()["usuario"] == USUARIO
    assert cliente.get("/api/auth/session").get_json()["autenticado"] is True


def test_login_incorrecto_mensaje_generico(cliente):
    s = Sesion(cliente)
    malo_password = s.login(password="otra-clave")
    malo_usuario = s.login(usuario="nadie")
    for respuesta in (malo_password, malo_usuario):
        assert respuesta.status_code == 401
        assert respuesta.get_json()["error"] == "Usuario o contraseña incorrectos"
    assert malo_password.get_json() == malo_usuario.get_json()


def test_login_sin_csrf(cliente):
    respuesta = cliente.post("/api/auth/login", json={"usuario": USUARIO, "password": PASSWORD})
    assert respuesta.status_code == 403
    cliente.get("/api/auth/csrf")
    respuesta = cliente.post("/api/auth/login", json={"usuario": USUARIO, "password": PASSWORD},
                             headers={"X-CSRF-Token": "falso"})
    assert respuesta.status_code == 403


def test_bloqueo_tras_cinco_intentos_y_liberacion(app, cliente):
    ahora = [1000.0]
    app.extensions["verificador"].limitador = LimitadorLogin(5, 15 * 60, reloj=lambda: ahora[0])
    s = Sesion(cliente)
    for _ in range(5):
        assert s.login(password="mala").status_code == 401
    bloqueado = s.login()
    assert bloqueado.status_code == 429
    assert "Demasiados intentos" in bloqueado.get_json()["error"]
    ahora[0] += 14 * 60
    assert s.login().status_code == 429
    ahora[0] += 61
    assert s.login().status_code == 200


def test_limitador_limpia_registros_vencidos():
    ahora = [0.0]
    limitador = LimitadorLogin(5, 900, reloj=lambda: ahora[0])
    limitador.registrar_fallo("1.1.1.1")
    assert len(limitador) == 1
    ahora[0] = 901
    assert len(limitador) == 0


def test_logout_invalida_la_cookie(cliente, sesion):
    cookie = cliente.get_cookie("gs_sesion").value
    assert sesion.post("/api/auth/logout").status_code == 200
    assert cliente.get("/api/auth/session").status_code == 401
    cliente.set_cookie("gs_sesion", cookie)
    assert cliente.get("/api/auth/session").status_code == 401


def test_cookies_seguras(directorio_temporal):
    app = create_app(configuracion(cookie_secure=True), directorio_temporal=str(directorio_temporal),
                     limpieza_periodica=False)
    cliente = app.test_client()
    s = Sesion(cliente)
    respuesta = s.login()
    cabecera = next(v for v in respuesta.headers.getlist("Set-Cookie") if v.startswith("gs_sesion="))
    assert "HttpOnly" in cabecera and "Secure" in cabecera and "SameSite=Lax" in cabecera
    assert "Expires=" in cabecera


def test_sesion_dura_ocho_horas(app, cliente, sesion, monkeypatch):
    import security.authentication as autenticacion

    real = time.time()
    monkeypatch.setattr(autenticacion.time, "time", lambda: real + 7.9 * 3600)
    assert cliente.get("/api/auth/session").status_code == 200
    monkeypatch.setattr(autenticacion.time, "time", lambda: real + 8.1 * 3600)
    respuesta = cliente.get("/api/auth/session")
    assert respuesta.status_code == 401
    assert app.permanent_session_lifetime.total_seconds() == 8 * 3600


def test_rutas_protegidas_sin_sesion(cliente):
    s = Sesion(cliente)
    assert s.post("/api/verificar").status_code == 401
    assert s.post("/api/nueva").status_code == 401
    assert cliente.get("/api/descargar/" + "a" * 32).status_code == 401
    assert cliente.get("/api/auth/session").status_code == 401


def test_aislamiento_entre_sesiones(app):
    a, b = Sesion(app.test_client()), Sesion(app.test_client())
    assert a.login().status_code == 200
    assert b.cliente.get("/api/auth/session").status_code == 401


def test_credenciales_validas_con_hash_de_referencia():
    from conftest import HASH
    assert credenciales_validas(USUARIO, PASSWORD, USUARIO, HASH)
    assert not credenciales_validas("otro", PASSWORD, USUARIO, HASH)
    assert not credenciales_validas(USUARIO, "x", USUARIO, "hash-invalido")


def test_crear_usuario_genera_env(tmp_path):
    lineas, existentes = [], {}
    variables = crear_usuario.generar_variables("adm'in", "secreta1", existentes)
    assert variables["APP_PASSWORD_HASH"].startswith("pbkdf2:sha256")
    assert check_password_hash(variables["APP_PASSWORD_HASH"], "secreta1")
    ruta = tmp_path / ".env"
    crear_usuario.guardar_env(ruta, crear_usuario.actualizar_env(lineas, variables))
    _, leidos = crear_usuario.leer_env(ruta)
    assert leidos["APP_USUARIO"] == "adm'in"
    assert leidos["APP_PASSWORD_HASH"] == variables["APP_PASSWORD_HASH"]

    from dotenv import dotenv_values
    valores = dotenv_values(ruta)
    assert valores["APP_USUARIO"] == "adm'in"
    assert valores["APP_PASSWORD_HASH"] == variables["APP_PASSWORD_HASH"]

    nuevas = crear_usuario.generar_variables("admin", "otra-clave", leidos)
    assert nuevas["SECRET_KEY"] == leidos["SECRET_KEY"]
    assert "COOKIE_SECURE" not in nuevas


def test_crear_usuario_rechaza_password_corta():
    import pytest
    with pytest.raises(ValueError):
        crear_usuario.generar_variables("admin", "123", {})
