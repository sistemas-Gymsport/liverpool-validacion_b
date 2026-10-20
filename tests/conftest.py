import copy
import io
import sys
from pathlib import Path

import pytest
from werkzeug.security import generate_password_hash

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import Configuracion  # noqa: E402
from factory import create_app  # noqa: E402
import generar_ejemplos as ej  # noqa: E402

USUARIO = "operador"
PASSWORD = "clave-de-prueba"
HASH = generate_password_hash(PASSWORD, method="pbkdf2:sha256")


def configuracion(**cambios) -> Configuracion:
    base = Configuracion(usuario=USUARIO, password_hash=HASH, secret_key="x" * 48)
    for clave, valor in cambios.items():
        setattr(base, clave, valor)
    return base


@pytest.fixture
def directorio_temporal(tmp_path):
    ruta = tmp_path / "subidas"
    ruta.mkdir()
    return ruta


@pytest.fixture
def app(directorio_temporal):
    aplicacion = create_app(configuracion(), directorio_temporal=str(directorio_temporal), limpieza_periodica=False)
    aplicacion.testing = True
    return aplicacion


@pytest.fixture
def cliente(app):
    return app.test_client()


class Sesion:
    def __init__(self, cliente):
        self.cliente = cliente
        self.csrf = cliente.get("/api/auth/csrf").get_json()["csrf"]

    def login(self, usuario=USUARIO, password=PASSWORD):
        respuesta = self.cliente.post("/api/auth/login", json={"usuario": usuario, "password": password},
                                      headers={"X-CSRF-Token": self.csrf})
        if respuesta.status_code == 200:
            self.csrf = respuesta.get_json()["csrf"]
        return respuesta

    def post(self, ruta, **kwargs):
        headers = kwargs.pop("headers", {})
        headers.setdefault("X-CSRF-Token", self.csrf)
        return self.cliente.post(ruta, headers=headers, **kwargs)

    def verificar(self, excel: bytes, pdfs: list[tuple[str, bytes]], nombre_excel="reporte de pedidos.xlsx"):
        datos = {
            "excel": (io.BytesIO(excel), nombre_excel),
            "pdfs": [(io.BytesIO(contenido), nombre) for nombre, contenido in pdfs],
        }
        return self.post("/api/verificar", data=datos, content_type="multipart/form-data")


@pytest.fixture
def sesion(cliente):
    s = Sesion(cliente)
    assert s.login().status_code == 200
    return s


def excel_bytes(pedidos=None, **kwargs) -> bytes:
    buffer = io.BytesIO()
    ej.crear_excel(buffer, pedidos if pedidos is not None else ej.PEDIDOS_BASE, **kwargs)
    return buffer.getvalue()


def pdf_bytes(renglones=None, **kwargs) -> bytes:
    buffer = io.BytesIO()
    ej.crear_picking(buffer, renglones if renglones is not None else ej.RENGLONES_BASE, **kwargs)
    return buffer.getvalue()


def pedidos_base():
    return copy.deepcopy(ej.PEDIDOS_BASE)


def renglones_base():
    return copy.deepcopy(ej.RENGLONES_BASE)


@pytest.fixture
def documentos():
    return excel_bytes(), pdf_bytes()
