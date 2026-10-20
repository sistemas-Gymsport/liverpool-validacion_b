import io
import re

from flask import Blueprint, current_app, jsonify, request, send_file

from api.errores import ErrorApi
from security.authentication import requiere_sesion
from services import file_service

bp = Blueprint("verificacion", __name__, url_prefix="/api")

RE_ID_REPORTE = re.compile(r"^[A-Za-z0-9_-]{20,64}$")
TIPO_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CAMPOS_ARCHIVOS = ("documentos", "excel", "medidas", "pdfs")


def _estado():
    return current_app.extensions["verificador"]


@bp.post("/verificar")
@requiere_sesion
def verificar(sesion):
    estado = _estado()
    archivos = [a for campo in CAMPOS_ARCHIVOS for a in request.files.getlist(campo)]
    try:
        estado.reportes.eliminar_sesion(sesion["sid"])
        documentos = file_service.validar_documentos(archivos, estado.config)
        datos, contenido, nombre = estado.servicio.ejecutar(documentos)
    finally:
        file_service.cerrar_archivos(archivos)

    datos["reporte_id"] = estado.reportes.guardar(sesion["sid"], contenido, nombre)
    datos["nombre_reporte"] = nombre
    datos["ok"] = True
    return jsonify(datos)


@bp.get("/descargar/<reporte_id>")
@requiere_sesion
def descargar(reporte_id: str, sesion):
    reporte = None
    if RE_ID_REPORTE.match(reporte_id):
        reporte = _estado().reportes.obtener(reporte_id, sesion["sid"])
    if reporte is None:
        raise ErrorApi(404, "El reporte ya no está disponible (expiró o el servidor se reinició). Repite la verificación.",
                       "reporte_no_disponible")
    return send_file(io.BytesIO(reporte.contenido), mimetype=TIPO_XLSX, as_attachment=True,
                     download_name=reporte.nombre_archivo, max_age=0)


@bp.post("/nueva")
@requiere_sesion
def nueva(sesion):
    _estado().reportes.eliminar_sesion(sesion["sid"])
    return jsonify({"ok": True})
