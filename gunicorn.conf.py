"""Configuración de gunicorn. Se carga automáticamente al ejecutar `gunicorn app:app` desde backend/.

Escucha en el puerto que indique la plataforma (Render define PORT). Un solo worker porque los reportes
temporales, el bloqueo de intentos y las sesiones revocadas viven en la memoria del proceso.
"""

import os

bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"
workers = 1
threads = 4
timeout = 120
accesslog = "-"
