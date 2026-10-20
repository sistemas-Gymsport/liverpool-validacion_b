import threading
import time
from dataclasses import dataclass
from typing import Callable


@dataclass
class _Registro:
    fallos: int
    primero: float
    bloqueado_hasta: float = 0.0


class LimitadorLogin:
    """Bloquea una IP tras varios intentos fallidos. Vive en memoria: se reinicia con el proceso."""

    def __init__(self, max_intentos: int, bloqueo_segundos: float, reloj: Callable[[], float] = time.monotonic):
        self.max_intentos = max_intentos
        self.bloqueo = bloqueo_segundos
        self.reloj = reloj
        self._registros: dict[str, _Registro] = {}
        self._candado = threading.Lock()

    def segundos_restantes(self, ip: str) -> int:
        with self._candado:
            self._limpiar()
            registro = self._registros.get(ip)
            if registro and registro.bloqueado_hasta > self.reloj():
                return int(registro.bloqueado_hasta - self.reloj()) + 1
            return 0

    def registrar_fallo(self, ip: str) -> None:
        with self._candado:
            ahora = self.reloj()
            registro = self._registros.get(ip)
            if registro is None or ahora - registro.primero > self.bloqueo:
                registro = _Registro(0, ahora)
                self._registros[ip] = registro
            registro.fallos += 1
            if registro.fallos >= self.max_intentos:
                registro.bloqueado_hasta = ahora + self.bloqueo

    def registrar_exito(self, ip: str) -> None:
        with self._candado:
            self._registros.pop(ip, None)

    def _limpiar(self) -> None:
        ahora = self.reloj()
        vencidos = [
            ip for ip, r in self._registros.items()
            if r.bloqueado_hasta <= ahora and ahora - r.primero > self.bloqueo
        ]
        for ip in vencidos:
            del self._registros[ip]

    def __len__(self) -> int:
        with self._candado:
            self._limpiar()
            return len(self._registros)
