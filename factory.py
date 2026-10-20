import logging
import tempfile
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, Request, jsonify, request
from werkzeug.middleware.proxy_fix import ProxyFix

from api import auth, verification
from api.errores import ErrorApi, registrar_manejadores
from config import MB, Configuracion
from security import csrf
from security.authentication import SesionesRevocadas
from security.rate_limit import LimitadorLogin
from services.report_service import AlmacenReportes, iniciar_limpieza_periodica
from services.verification_service import ServicioVerificacion

DIRECTORIO = Path(__file__).resolve().parent


@dataclass
class EstadoAplicacion:
    config: Configuracion
    limitador: LimitadorLogin
    reportes: AlmacenReportes
    revocadas: SesionesRevocadas
    servicio: ServicioVerificacion
    directorio_temporal: str | None = None


class SolicitudConTemporales(Request):
    """Los archivos subidos grandes se guardan en temporales del sistema que se borran al cerrarse."""

    def _get_file_stream(self, total_content_length, content_type, filename=None, content_length=None):
        from flask import current_app

        directorio = current_app.extensions["verificador"].directorio_temporal
        return tempfile.SpooledTemporaryFile(max_size=1 * MB, dir=directorio)


def create_app(config: Configuracion | None = None, *, directorio_temporal: str | None = None,
               limpieza_periodica: bool = True) -> Flask:
    if config is None:
        load_dotenv(DIRECTORIO / ".env")
        config = Configuracion.desde_entorno()

    app = Flask(__name__, static_folder=None)
    app.request_class = SolicitudConTemporales
    app.config.update(
        SECRET_KEY=config.secret_key,
        MAX_CONTENT_LENGTH=config.max_content_length,
        SESSION_COOKIE_NAME="gs_sesion",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=config.cookie_secure,
        SESSION_COOKIE_SAMESITE=config.cookie_samesite,
        PERMANENT_SESSION_LIFETIME=config.duracion_sesion,
        SESSION_REFRESH_EACH_REQUEST=False,
        JSON_SORT_KEYS=False,
    )
    if config.proxies_confiables > 0:
        app.wsgi_app = ProxyFix(app.wsgi_app, x_for=config.proxies_confiables, x_proto=1)

    reportes = AlmacenReportes(config.retencion_reporte.total_seconds(), config.max_reportes,
                               config.max_memoria_reportes_mb * MB)
    app.extensions["verificador"] = EstadoAplicacion(
        config=config,
        limitador=LimitadorLogin(config.max_intentos_login, config.bloqueo_login.total_seconds()),
        reportes=reportes,
        revocadas=SesionesRevocadas(config.duracion_sesion.total_seconds()),
        servicio=ServicioVerificacion(config.max_verificaciones_simultaneas),
        directorio_temporal=directorio_temporal,
    )
    if limpieza_periodica:
        iniciar_limpieza_periodica(reportes)

    _configurar_registro()
    registrar_manejadores(app)
    _registrar_seguridad(app, config)
    app.register_blueprint(auth.bp)
    app.register_blueprint(verification.bp)

    @app.get("/api/health")
    def salud():
        return jsonify({"ok": True})

    return app


def _configurar_registro() -> None:
    if not logging.getLogger("verificador").handlers:
        logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


def _registrar_seguridad(app: Flask, config: Configuracion) -> None:
    @app.before_request
    def _antes():
        if request.method == "OPTIONS" and request.headers.get("Origin") in config.cors_origins:
            return app.make_default_options_response()
        if request.path.startswith("/api/") and not csrf.token_valido():
            raise ErrorApi(403, "La solicitud no es válida (token de seguridad ausente o vencido). Recarga la página.",
                           "csrf")
        return None

    @app.after_request
    def _cabeceras(respuesta):
        respuesta.headers["X-Content-Type-Options"] = "nosniff"
        respuesta.headers["X-Frame-Options"] = "DENY"
        respuesta.headers["Referrer-Policy"] = "no-referrer"
        respuesta.headers["Content-Security-Policy"] = "default-src 'none'; frame-ancestors 'none'"
        respuesta.headers["Cache-Control"] = "no-store"
        respuesta.headers["Pragma"] = "no-cache"
        if config.cookie_secure:
            respuesta.headers["Strict-Transport-Security"] = "max-age=31536000"
        origen = request.headers.get("Origin")
        if origen and origen in config.cors_origins:
            respuesta.headers["Access-Control-Allow-Origin"] = origen
            respuesta.headers["Access-Control-Allow-Credentials"] = "true"
            respuesta.headers["Access-Control-Allow-Headers"] = "Content-Type, X-CSRF-Token"
            respuesta.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
            respuesta.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
            respuesta.headers["Access-Control-Max-Age"] = "600"
            respuesta.vary.add("Origin")
        return respuesta
