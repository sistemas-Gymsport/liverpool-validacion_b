import secrets
import threading
import time
from dataclasses import dataclass
from typing import Callable


@dataclass
class ReporteTemporal:
    sid: str
    contenido: bytes
    nombre_archivo: str
    creado: float


class AlmacenReportes:
    """Reportes en memoria del proceso. Uno por sesión, con caducidad y límite de memoria."""

    def __init__(self, retencion_segundos: float, max_reportes: int, max_bytes: int,
                 reloj: Callable[[], float] = time.monotonic):
        self.retencion = retencion_segundos
        self.max_reportes = max_reportes
        self.max_bytes = max_bytes
        self.reloj = reloj
        self._reportes: dict[str, ReporteTemporal] = {}
        self._candado = threading.Lock()

    def guardar(self, sid: str, contenido: bytes, nombre_archivo: str) -> str:
        identificador = secrets.token_urlsafe(24)
        with self._candado:
            self._limpiar()
            self._eliminar_de_sesion(sid)
            self._reportes[identificador] = ReporteTemporal(sid, contenido, nombre_archivo, self.reloj())
            self._aplicar_limites()
        return identificador

    def obtener(self, identificador: str, sid: str) -> ReporteTemporal | None:
        with self._candado:
            self._limpiar()
            reporte = self._reportes.get(identificador)
            if reporte is None or not secrets.compare_digest(reporte.sid, sid):
                return None
            return reporte

    def eliminar_sesion(self, sid: str) -> int:
        with self._candado:
            return self._eliminar_de_sesion(sid)

    def limpiar(self) -> None:
        with self._candado:
            self._limpiar()

    def __len__(self) -> int:
        with self._candado:
            return len(self._reportes)

    @property
    def bytes_usados(self) -> int:
        with self._candado:
            return sum(len(r.contenido) for r in self._reportes.values())

    def _eliminar_de_sesion(self, sid: str) -> int:
        ids = [i for i, r in self._reportes.items() if r.sid == sid]
        for i in ids:
            del self._reportes[i]
        return len(ids)

    def _limpiar(self) -> None:
        limite = self.reloj() - self.retencion
        for i in [i for i, r in self._reportes.items() if r.creado <= limite]:
            del self._reportes[i]

    def _aplicar_limites(self) -> None:
        ordenados = sorted(self._reportes.items(), key=lambda par: par[1].creado)
        total = sum(len(r.contenido) for r in self._reportes.values())
        while ordenados and (len(self._reportes) > self.max_reportes or total > self.max_bytes):
            identificador, reporte = ordenados.pop(0)
            total -= len(reporte.contenido)
            del self._reportes[identificador]


def iniciar_limpieza_periodica(almacen: AlmacenReportes, intervalo: float = 300) -> threading.Thread:
    def ciclo():
        while True:
            time.sleep(intervalo)
            almacen.limpiar()

    hilo = threading.Thread(target=ciclo, name="limpieza-reportes", daemon=True)
    hilo.start()
    return hilo
