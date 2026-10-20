import os
import re
import zipfile
from dataclasses import dataclass
from typing import BinaryIO

from werkzeug.datastructures import FileStorage

from api.errores import MENSAJE_TAMANO, ErrorApi
from config import MB, Configuracion

FIRMA_ZIP = b"PK\x03\x04"
FIRMA_PDF = b"%PDF-"
MAX_DESCOMPRIMIDO_XLSX = 300 * MB
MAX_ENTRADAS_XLSX = 5000
MAX_RELACION_COMPRESION = 1000


@dataclass
class ArchivoValidado:
    """Archivo recibido y validado. Envuelve el flujo temporal de Werkzeug sin copiarlo."""

    nombre: str
    tamano: int
    flujo: BinaryIO

    @property
    def name(self) -> str:
        return self.nombre

    def read(self, n: int = -1) -> bytes:
        return self.flujo.read(n)

    def seek(self, posicion: int, desde: int = 0) -> int:
        return self.flujo.seek(posicion, desde)

    def tell(self) -> int:
        return self.flujo.tell()

    def seekable(self) -> bool:
        return True

    def readable(self) -> bool:
        return True

    def close(self) -> None:
        self.flujo.close()


def nombre_seguro(nombre: str | None) -> str:
    base = os.path.basename((nombre or "").replace("\\", "/"))
    base = re.sub(r"[\x00-\x1f\x7f<>:\"|?*]", "", base).strip(" .")
    return base[:150] or "archivo"


def _tamano(flujo: BinaryIO) -> int:
    flujo.seek(0, os.SEEK_END)
    tamano = flujo.tell()
    flujo.seek(0)
    return tamano


def _validar_basico(archivo: FileStorage, extension: str, firma: bytes, max_mb: int, mensaje_formato: str) -> ArchivoValidado:
    nombre = nombre_seguro(archivo.filename)
    if not nombre.lower().endswith(extension):
        raise ErrorApi(400, f"«{nombre}»: {mensaje_formato}", "extension_invalida")
    tamano = _tamano(archivo.stream)
    if tamano == 0:
        raise ErrorApi(400, f"«{nombre}» está vacío.", "archivo_vacio")
    if tamano > max_mb * MB:
        raise ErrorApi(413, f"«{nombre}»: {MENSAJE_TAMANO} (máximo {max_mb} MB).", "archivo_grande")
    cabecera = archivo.stream.read(1024)
    archivo.stream.seek(0)
    if (firma == FIRMA_PDF and firma not in cabecera) or (firma == FIRMA_ZIP and not cabecera.startswith(firma)):
        raise ErrorApi(400, f"«{nombre}»: {mensaje_formato}", "firma_invalida")
    return ArchivoValidado(nombre, tamano, archivo.stream)


def _validar_zip_xlsx(archivo: ArchivoValidado) -> None:
    mensaje = f"«{archivo.nombre}»: El archivo Excel no tiene un formato válido."
    try:
        with zipfile.ZipFile(archivo.flujo) as contenedor:
            entradas = contenedor.infolist()
            if len(entradas) > MAX_ENTRADAS_XLSX or "xl/workbook.xml" not in contenedor.namelist():
                raise ErrorApi(400, mensaje, "excel_invalido")
            total = 0
            for entrada in entradas:
                total += entrada.file_size
                relacion = entrada.file_size / max(entrada.compress_size, 1)
                if total > MAX_DESCOMPRIMIDO_XLSX or relacion > MAX_RELACION_COMPRESION:
                    raise ErrorApi(400, f"«{archivo.nombre}»: el Excel está excesivamente comprimido y se rechazó por seguridad.", "excel_sospechoso")
    except zipfile.BadZipFile as exc:
        raise ErrorApi(400, mensaje, "excel_invalido") from exc
    finally:
        archivo.seek(0)


def validar_documentos(archivos: list[FileStorage], config: Configuracion) -> list[ArchivoValidado]:
    """Valida cualquier combinación de documentos (.xlsx y .pdf). El tipo de cada uno se identifica después."""
    archivos = [a for a in archivos if a and a.filename]
    if not archivos:
        raise ErrorApi(400, "Sube al menos el reporte de pedidos de Liverpool (.xlsx).", "documentos_requeridos")
    nombres = [nombre_seguro(a.filename).lower() for a in archivos]
    for archivo, nombre in zip(archivos, nombres):
        if not nombre.endswith((".xlsx", ".pdf")):
            raise ErrorApi(400, f"«{nombre_seguro(archivo.filename)}»: solo se aceptan archivos Excel (.xlsx) y PDF.",
                           "extension_invalida")
    if sum(n.endswith(".pdf") for n in nombres) > config.max_pdfs:
        raise ErrorApi(400, f"Puedes subir un máximo de {config.max_pdfs} documentos PDF.", "demasiados_pdf")
    if sum(n.endswith(".xlsx") for n in nombres) > config.max_excels:
        raise ErrorApi(400, f"Puedes subir un máximo de {config.max_excels} archivos Excel.", "demasiados_excel")
    if len(set(nombres)) != len(nombres):
        raise ErrorApi(400, "Subiste dos veces el mismo documento. Quita el duplicado.", "documento_repetido")
    validados = [
        _validar_xlsx(a, config) if n.endswith(".xlsx") else
        _validar_basico(a, ".pdf", FIRMA_PDF, config.max_pdf_mb, "El documento PDF está dañado o no puede leerse.")
        for a, n in zip(archivos, nombres)
    ]
    validar_total(validados, config)
    return validados


def _validar_xlsx(archivo: FileStorage, config: Configuracion) -> ArchivoValidado:
    validado = _validar_basico(archivo, ".xlsx", FIRMA_ZIP, config.max_excel_mb,
                               "El archivo Excel no tiene un formato válido.")
    _validar_zip_xlsx(validado)
    return validado


def validar_total(archivos: list[ArchivoValidado], config: Configuracion) -> None:
    total = sum(a.tamano for a in archivos)
    if total > config.max_total_mb * MB:
        raise ErrorApi(413, f"{MENSAJE_TAMANO} El total no puede superar {config.max_total_mb} MB.", "total_grande")


def cerrar_archivos(archivos: list[FileStorage]) -> None:
    for archivo in archivos:
        try:
            archivo.close()
        except Exception:
            pass
