"""Crea o actualiza las credenciales de acceso en backend/.env.

Uso:
    python crear_usuario.py              # pide usuario y contraseña, escribe .env
    python crear_usuario.py --render     # además muestra las variables para copiarlas en Render
"""

import argparse
import getpass
import os
import re
import secrets
import stat
import sys
from pathlib import Path

from werkzeug.security import generate_password_hash

RUTA_ENV = Path(__file__).resolve().parent / ".env"
LONGITUD_MINIMA = 6
VALORES_LOCALES = {"COOKIE_SECURE": "0", "FLASK_ENV": "development", "PORT": "8000"}
RE_LINEA = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$")


def entrecomillar(valor: str) -> str:
    return "'" + valor.replace("\\", "\\\\").replace("'", "\\'") + "'"


def desentrecomillar(valor: str) -> str:
    valor = valor.strip()
    if len(valor) >= 2 and valor[0] == valor[-1] and valor[0] in "'\"":
        valor = valor[1:-1].replace("\\'", "'").replace('\\"', '"').replace("\\\\", "\\")
    return valor


def leer_env(ruta: Path) -> tuple[list[str], dict[str, str]]:
    if not ruta.exists():
        return [], {}
    lineas = ruta.read_text(encoding="utf-8").splitlines()
    valores = {}
    for linea in lineas:
        coincidencia = RE_LINEA.match(linea)
        if coincidencia:
            valores[coincidencia.group(1)] = desentrecomillar(coincidencia.group(2))
    return lineas, valores


def actualizar_env(lineas: list[str], nuevos: dict[str, str]) -> str:
    pendientes = dict(nuevos)
    salida = []
    for linea in lineas:
        coincidencia = RE_LINEA.match(linea)
        if coincidencia and coincidencia.group(1) in pendientes:
            nombre = coincidencia.group(1)
            salida.append(f"{nombre}={entrecomillar(pendientes.pop(nombre))}")
        else:
            salida.append(linea)
    salida.extend(f"{nombre}={entrecomillar(valor)}" for nombre, valor in pendientes.items())
    return "\n".join(salida) + "\n"


def generar_variables(usuario: str, password: str, existentes: dict[str, str]) -> dict[str, str]:
    if len(password) < LONGITUD_MINIMA:
        raise ValueError(f"La contraseña debe tener al menos {LONGITUD_MINIMA} caracteres.")
    secret = existentes.get("SECRET_KEY") or ""
    if len(secret) < 32:
        secret = secrets.token_urlsafe(48)
    variables = {
        "APP_USUARIO": usuario,
        "APP_PASSWORD_HASH": generate_password_hash(password, method="pbkdf2:sha256"),
        "SECRET_KEY": secret,
    }
    for nombre, valor in VALORES_LOCALES.items():
        if nombre not in existentes:
            variables[nombre] = valor
    return variables


def guardar_env(ruta: Path, contenido: str) -> None:
    descriptor = os.open(ruta, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as archivo:
        archivo.write(contenido)
    try:
        os.chmod(ruta, stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass


def pedir_datos() -> tuple[str, str]:
    usuario = input("Usuario: ").strip()
    while not usuario or len(usuario) > 100:
        usuario = input("Escribe un usuario válido (1 a 100 caracteres): ").strip()
    while True:
        password = getpass.getpass(f"Contraseña (mínimo {LONGITUD_MINIMA} caracteres): ")
        if len(password) < LONGITUD_MINIMA:
            print(f"La contraseña debe tener al menos {LONGITUD_MINIMA} caracteres.")
            continue
        if getpass.getpass("Confirma la contraseña: ") != password:
            print("Las contraseñas no coinciden. Inténtalo de nuevo.")
            continue
        return usuario, password


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea las credenciales del Verificador de pedidos.")
    parser.add_argument("--render", action="store_true",
                        help="muestra APP_USUARIO, APP_PASSWORD_HASH y SECRET_KEY para configurarlas en Render")
    args = parser.parse_args()

    lineas, existentes = leer_env(RUTA_ENV)
    usuario, password = pedir_datos()
    variables = generar_variables(usuario, password, existentes)
    guardar_env(RUTA_ENV, actualizar_env(lineas, variables))
    print(f"Credenciales guardadas en {RUTA_ENV}")
    print("SECRET_KEY " + ("conservada." if existentes.get("SECRET_KEY") == variables["SECRET_KEY"] else "generada."))

    if args.render:
        print("\nVariables para Render (Environment). No las compartas ni las subas a Git:")
        for nombre in ("APP_USUARIO", "APP_PASSWORD_HASH", "SECRET_KEY"):
            print(f"{nombre}={variables[nombre]}")
        print("COOKIE_SECURE=1")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado.")
        sys.exit(1)
