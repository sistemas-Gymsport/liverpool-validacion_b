import io
import logging
import threading
from dataclasses import asdict
from datetime import datetime
from zoneinfo import ZoneInfo

import verificar_pedidos as vp
from api.errores import MENSAJE_GENERICO, ErrorApi
from services.file_service import ArchivoValidado

ZONA_MEXICO = ZoneInfo("America/Mexico_City")
registro = logging.getLogger("verificador")


def ahora_mexico() -> datetime:
    return datetime.now(ZONA_MEXICO)


def nombre_reporte(momento: datetime) -> str:
    return f"verificacion_pedidos_{momento:%Y%m%d_%H%M}.xlsx"


def resultado_a_json(resultado: vp.Resultado) -> dict:
    return {
        "resumen": {
            "pedidos_liverpool": resultado.pedidos_liverpool,
            "verificados": resultado.verificados,
            "correctos": resultado.correctos,
            "errores": resultado.errores,
            "revisar": resultado.revisar,
            "sin_picking": resultado.sin_picking,
            "guias_por_revisar": resultado.guias_por_revisar,
        },
        "documentos": [asdict(d) for d in resultado.documentos],
        "guias": [asdict(g) for g in resultado.guias],
        "etiquetas": [asdict(e) for e in resultado.etiquetas],
        "medidas_cargadas": resultado.medidas_cargadas,
        "picking_cargado": resultado.picking_cargado,
        "guias_envio_cargadas": resultado.guias_envio_cargadas,
        "hallazgos": [asdict(h) for h in resultado.hallazgos],
        "detalle_pedidos": [asdict(d) for d in resultado.detalle],
        "pedidos_sin_picking": [asdict(s) for s in resultado.pedidos_sin_picking],
        "advertencias": list(resultado.advertencias),
        "pickings": list(resultado.pickings),
    }


class ServicioVerificacion:
    def __init__(self, max_simultaneas: int):
        self._turnos = threading.BoundedSemaphore(max_simultaneas)

    def ejecutar(self, documentos: list[ArchivoValidado]) -> tuple[dict, bytes, str]:
        if not self._turnos.acquire(timeout=30):
            raise ErrorApi(503, "El servidor está procesando otras verificaciones. Inténtalo en unos minutos.", "ocupado")
        try:
            return self._ejecutar(documentos)
        finally:
            self._turnos.release()

    def _ejecutar(self, documentos: list[ArchivoValidado]) -> tuple[dict, bytes, str]:
        try:
            resultado = vp.analizar([(d.nombre, d) for d in documentos])
        except vp.DocumentoInvalido as exc:
            raise ErrorApi(422, str(exc), "documento_invalido") from exc
        except MemoryError as exc:
            registro.error("Memoria insuficiente durante la verificación")
            raise ErrorApi(503, "El servidor no tiene memoria suficiente para estos documentos. Prueba con menos PDF.", "memoria") from exc

        momento = ahora_mexico()
        try:
            destino = io.BytesIO()
            vp.escribir_reporte(resultado, destino, generado=momento.replace(tzinfo=None))
            contenido = destino.getvalue()
        except Exception as exc:
            registro.error("Fallo al generar el reporte: %s", type(exc).__name__)
            raise ErrorApi(500, MENSAJE_GENERICO, "reporte_fallido") from exc

        datos = resultado_a_json(resultado)
        datos["generado"] = momento.strftime("%d/%m/%Y %H:%M")
        registro.info(
            "Verificación completada: %d documentos, %d renglones de picking, %d guías, %d errores, %d para revisar",
            len(documentos), len(resultado.detalle), len(resultado.etiquetas), resultado.errores, resultado.revisar,
        )
        return datos, contenido, nombre_reporte(momento)
