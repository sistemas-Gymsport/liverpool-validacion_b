# Backend en VPS + frontend en Vercel

```text
Navegador ──► https://liverpool-validacion.vercel.app   (frontend)
                 │  rewrite /api/:path*
                 ▼
              https://api.TU-DOMINIO.com  ── Nginx (HTTPS) ──► gunicorn 127.0.0.1:8000 (Flask)
```

El navegador solo habla con Vercel (mismo origen): no hace falta CORS y la cookie de sesión funciona con
`SameSite=Lax`. Vercel reenvía `/api` al VPS por HTTPS.

Guía para Ubuntu 22.04 / 24.04. Sustituye `api.TU-DOMINIO.com` por tu subdominio.

## 0. Antes de empezar

- Un **subdominio** (por ejemplo `api.gymsport.com.mx`) con un registro **A** que apunte a la IP pública del VPS.
  Sin dominio no se puede emitir el certificado HTTPS.
- Acceso SSH al VPS con un usuario con `sudo`.
- El repositorio del backend accesible desde el VPS (si es privado, con una *deploy key* de solo lectura).

## 1. Paquetes del sistema

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip git nginx certbot python3-certbot-nginx
sudo ufw allow OpenSSH
sudo ufw allow 'Nginx Full'
sudo ufw enable
```

El puerto 8000 **no** se abre: gunicorn solo escucha en `127.0.0.1`.

## 2. Usuario y código

```bash
sudo useradd --system --create-home --home-dir /opt/verificador --shell /usr/sbin/nologin verificador
sudo git clone https://github.com/sistemas-Gymsport/liverpool-validacion_b.git /opt/verificador/backend
sudo chown -R verificador:verificador /opt/verificador
```

## 3. Entorno virtual y dependencias

```bash
cd /opt/verificador/backend
sudo -u verificador python3 -m venv .venv
sudo -u verificador .venv/bin/pip install --upgrade pip
sudo -u verificador .venv/bin/pip install -r requirements.txt
```

## 4. Credenciales (`.env` en el VPS)

```bash
sudo -u verificador .venv/bin/python crear_usuario.py
```

Pide usuario y contraseña (usa los **reales**, no los de prueba) y crea `/opt/verificador/backend/.env` con permisos 600.
Después edítalo para producción:

```bash
sudo -u verificador nano /opt/verificador/backend/.env
```

Deja estas líneas así:

```text
COOKIE_SECURE='1'
FLASK_ENV='production'
TRUSTED_PROXIES='1'
```

Borra la línea `PORT` si existe (el servicio la define).

## 5. Servicio (systemd)

```bash
sudo cp /opt/verificador/backend/deploy/verificador.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now verificador
sudo systemctl status verificador
curl http://127.0.0.1:8000/api/health
```

Debe responder `{"ok":true}`. Registros: `sudo journalctl -u verificador -f`.

## 6. Nginx y HTTPS

```bash
sudo cp /opt/verificador/backend/deploy/nginx-verificador.conf /etc/nginx/sites-available/verificador
sudo sed -i 's/api.TU-DOMINIO.com/api.gymsport.com.mx/' /etc/nginx/sites-available/verificador
sudo ln -s /etc/nginx/sites-available/verificador /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
sudo certbot --nginx -d api.gymsport.com.mx --redirect
curl https://api.gymsport.com.mx/api/health
```

(Cambia `api.gymsport.com.mx` por tu subdominio real.) Certbot renueva el certificado automáticamente.

## 7. Vercel

En `frontend/vercel.json` cambia el destino del rewrite:

```json
{ "source": "/api/:path*", "destination": "https://api.gymsport.com.mx/api/:path*" }
```

Haz commit y push del frontend; Vercel vuelve a desplegar. Comprueba
`https://liverpool-validacion.vercel.app/api/health` → `{"ok":true}` y luego inicia sesión en el sitio.

No configures `VITE_API_BASE` ni `CORS_ORIGINS`: con el rewrite no hacen falta.

## 8. Actualizar el backend

```bash
cd /opt/verificador/backend
sudo -u verificador git pull
sudo -u verificador .venv/bin/pip install -r requirements.txt
sudo systemctl restart verificador
```

Al reiniciar se pierden los reportes temporales y las sesiones revocadas (comportamiento esperado).

## Límites a tener en cuenta

- **Tamaño:** Nginx acepta hasta 270 MB por petición; el proxy de Vercel no publica su límite. Si una verificación con
  documentos muy grandes falla con 413 o 502 desde Vercel, ver «Subida directa» en la guía de despliegue.
- **Tiempo:** Vercel espera hasta 120 s la respuesta del VPS; gunicorn también corta a los 120 s.
- **IP del cliente:** detrás de Vercel, la IP que ve el backend es la de Vercel; el bloqueo tras 5 intentos fallidos
  puede afectar a todo el equipo durante 15 minutos.
- **Un solo proceso:** no subas `workers` en `gunicorn.conf.py` (los reportes temporales viven en memoria).
