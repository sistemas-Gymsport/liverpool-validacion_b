import os
from dataclasses import dataclass, field
from datetime import timedelta

MB = 1024 * 1024


class ConfiguracionIncompleta(RuntimeError):
    pass


def _bool(valor: str | None, defecto: bool = False) -> bool:
    if valor is None or valor == "":
        return defecto
    return valor.strip().lower() in {"1", "true", "si", "sí", "yes", "on"}


def _int(nombre: str, defecto: int) -> int:
    valor = os.environ.get(nombre, "").strip()
    return int(valor) if valor else defecto


@dataclass
class Configuracion:
    usuario: str
    password_hash: str
    secret_key: str
    cookie_secure: bool = False
    cookie_samesite: str = "Lax"
    cors_origins: tuple[str, ...] = ()
    proxies_confiables: int = 0
    duracion_sesion: timedelta = timedelta(hours=8)
    max_excel_mb: int = 20
    max_pdf_mb: int = 20
    max_pdfs: int = 10
    max_excels: int = 3
    max_total_mb: int = 260
    retencion_reporte: timedelta = timedelta(hours=1)
    max_reportes: int = 50
    max_memoria_reportes_mb: int = 100
    max_intentos_login: int = 5
    bloqueo_login: timedelta = timedelta(minutes=15)
    max_verificaciones_simultaneas: int = 2
    extras: dict = field(default_factory=dict)

    @property
    def max_content_length(self) -> int:
        return self.max_total_mb * MB + 2 * MB

    @classmethod
    def desde_entorno(cls) -> "Configuracion":
        faltantes = [n for n in ("APP_USUARIO", "APP_PASSWORD_HASH", "SECRET_KEY") if not os.environ.get(n)]
        if faltantes:
            raise ConfiguracionIncompleta(
                "Faltan variables de entorno obligatorias: " + ", ".join(faltantes)
                + ". En local ejecuta 'python crear_usuario.py' dentro de backend/; en Render configúralas en Environment."
            )
        secret = os.environ["SECRET_KEY"]
        if len(secret) < 32:
            raise ConfiguracionIncompleta("SECRET_KEY debe tener al menos 32 caracteres. Ejecuta 'python crear_usuario.py'.")
        samesite = os.environ.get("COOKIE_SAMESITE", "Lax").strip().capitalize()
        if samesite not in {"Lax", "Strict", "None"}:
            raise ConfiguracionIncompleta("COOKIE_SAMESITE debe ser Lax, Strict o None.")
        secure = _bool(os.environ.get("COOKIE_SECURE"), defecto=False)
        if samesite == "None" and not secure:
            raise ConfiguracionIncompleta("COOKIE_SAMESITE=None requiere COOKIE_SECURE=1.")
        origenes = tuple(o.strip().rstrip("/") for o in os.environ.get("CORS_ORIGINS", "").split(",") if o.strip())
        if "*" in origenes:
            raise ConfiguracionIncompleta("CORS_ORIGINS no admite comodines; indica los orígenes exactos.")
        return cls(
            usuario=os.environ["APP_USUARIO"],
            password_hash=os.environ["APP_PASSWORD_HASH"],
            secret_key=secret,
            cookie_secure=secure,
            cookie_samesite=samesite,
            cors_origins=origenes,
            proxies_confiables=_int("TRUSTED_PROXIES", 0),
            max_total_mb=_int("MAX_TOTAL_MB", 260),
        )
