import logging

from flask import Flask, jsonify
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge

registro = logging.getLogger("verificador")

MENSAJE_GENERICO = "No fue posible completar la verificación. Inténtalo nuevamente."
MENSAJE_TAMANO = "El archivo supera el tamaño máximo permitido."


class ErrorApi(Exception):
    def __init__(self, estado: int, mensaje: str, codigo: str = "error", **extra):
        super().__init__(mensaje)
        self.estado = estado
        self.mensaje = mensaje
        self.codigo = codigo
        self.extra = extra


def respuesta_error(estado: int, mensaje: str, codigo: str, **extra):
    cuerpo = {"ok": False, "error": mensaje, "codigo": codigo}
    cuerpo.update(extra)
    return jsonify(cuerpo), estado


def registrar_manejadores(app: Flask) -> None:
    @app.errorhandler(ErrorApi)
    def _error_api(error: ErrorApi):
        return respuesta_error(error.estado, error.mensaje, error.codigo, **error.extra)

    @app.errorhandler(RequestEntityTooLarge)
    def _demasiado_grande(_error):
        return respuesta_error(413, MENSAJE_TAMANO, "archivo_grande")

    @app.errorhandler(HTTPException)
    def _http(error: HTTPException):
        mensajes = {
            400: "La solicitud no es válida.",
            404: "El recurso solicitado no existe.",
            405: "Método no permitido.",
        }
        return respuesta_error(error.code or 500, mensajes.get(error.code, MENSAJE_GENERICO), "http")

    @app.errorhandler(Exception)
    def _inesperado(error: Exception):
        registro.error("Error inesperado: %s", type(error).__name__)
        return respuesta_error(500, MENSAJE_GENERICO, "error_interno")
