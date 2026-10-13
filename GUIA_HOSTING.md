# GUÍA: Subir SMCCI al Hosting

## OPCIÓN 1: RENDER (Recomendada - Gratis para empezar)

### Paso 1: Preparar archivos
Subir a GitHub toda la carpeta `backend/` excepto:
- `__pycache__/`
- `*.pyc`
- `.env` (las variables se ponen en Render)

### Paso 2: Crear cuenta en Render
1. Ir a https://render.com
2. Crear cuenta con GitHub
3. New → Web Service
4. Conectar repositorio de GitHub

### Paso 3: Configurar
- **Name:** smcci
- **Runtime:** Python
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT`

### Paso 4: Variables de entorno (en Render)
```
DB_HOST=dpg-xxxxx-a.oregon-postgres.render.com
DB_PORT=5432
DB_NAME=iglesia_mcci
DB_USER=mcci_user
DB_PASSWORD=tu_password_seguro
SECRET_KEY=tu_secret_key_largo
FLASK_ENV=production
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=587
MAIL_USE_TLS=true
MAIL_USERNAME=tu_correo@gmail.com
MAIL_PASSWORD=tu_clave_app
MAIL_DEFAULT_SENDER=SMCCI <tu_correo@gmail.com>
```

### Paso 5: Crear BD en Render
1. New → PostgreSQL
2. Copiar el Internal Database URL
3. Usar los datos en las variables de entorno

### Paso 6: Migrar datos
1. En Render, abrir Web Service → Shell
2. Ejecutar los scripts SQL en orden:
```bash
psql $DATABASE_URL -f sql/01_schema.sql
psql $DATABASE_URL -f sql/02_seed_chile.sql
# ... etc
```

---

## OPCIÓN 2: RAILWAY (Fácil y rápido)

### Paso 1: Instalar Railway CLI
```bash
npm install -g @railway/cli
railway login
```

### Paso 2: Inicializar proyecto
```bash
cd C:\Datos\MCCI\SMCCI\backend
railway init
```

### Paso 3: Agregar PostgreSQL
```bash
railway add postgresql
```

### Paso 4: Variables de entorno
Crear archivo `railway.json`:
```json
{
  "build": {
    "builder": "NIXPACKS"
  },
  "deploy": {
    "startCommand": "gunicorn app:app --bind 0.0.0.0:$PORT",
    "healthcheckPath": "/"
  }
}
```

### Paso 5: Subir
```bash
railway up
railway open
```

---

## OPCIÓN 3: PYTHONANYWHERE (Más simple)

### Paso 1: Crear cuenta
1. Ir a https://www.pythonanywhere.com
2. Crear cuenta gratuita

### Paso 2: Subir archivos
1. Files → Upload
2. Subir todos los archivos de `backend/`

### Paso 3: Configurar Web App
1. Web → Add a new web app
2. Framework: Flask
3. Python version: 3.10

### Paso 4: Configurar BD
1. Databases → Add new database
2. Copiar datos de conexión
3. Actualizar `.env`

### Paso 5: Instalar dependencias
```bash
pip install -r requirements.txt
```

### Paso 6: Migrar datos
```bash
psql < sql/01_schema.sql
```

---

## OPCIÓN 4: VPS DigitalOcean ($4-6/mes)

### Paso 1: Crear Droplet
1. Ir a https://www.digitalocean.com
2. Crear Droplet
3. Plan: Basic $4/mes (Ubuntu 22.04)
4. Agregar SSH key

### Paso 2: Conectar
```bash
ssh root@tu_ip
```

### Paso 3: Instalar Python
```bash
apt update
apt install python3-pip python3-venv postgresql postgresql-contrib
```

### Paso 4: Configurar PostgreSQL
```bash
sudo -u postgres psql
CREATE USER mcci WITH PASSWORD 'tu_password';
CREATE DATABASE "Iglesia_MCCI" OWNER mcci;
\q
```

### Paso 5: Subir archivos
```bash
# Desde tu PC
scp -r C:\Datos\MCCI\SMCCI\backend\* root@tu_ip:/var/www/smcci/
```

### Paso 6: Configurar app
```bash
cd /var/www/smcci
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Paso 7: Crear servicio systemd
```bash
nano /etc/systemd/system/smcci.service
```
```ini
[Unit]
Description=SMCCI Flask App
After=network.target

[Service]
User=root
WorkingDirectory=/var/www/smcci
ExecStart=/var/www/smcci/venv/bin/gunicorn app:app --bind 0.0.0.0:5000

[Install]
WantedBy=multi-user.target
```

### Paso 8: Iniciar
```bash
systemctl enable smcci
systemctl start smcci
```

---

## ARCHIVOS A SUBIR

### Carpeta `backend/`:
```
backend/
├── app.py              ← App principal
├── requirements.txt    ← Dependencias
├── .env               ← Variables (NO subir a GitHub)
├── iniciar.bat        ← Para Windows local
├── sync_sheets.py     ← Sync Google Sheets
├── sincronizar.bat    ← Ejecutar sync
├── notificar.py       ← Notificaciones
├── templates/         ← 18 archivos HTML
│   ├── base.html
│   ├── login.html
│   ├── hermanos.html
│   ├── hermano_form.html
│   ├── lideres.html
│   ├── lider_form.html
│   ├── discipulado.html
│   ├── discipulado_form.html
│   ├── encuentros.html
│   ├── encuentro_form.html
│   ├── seguimiento.html
│   ├── visitas.html
│   ├── visita_form.html
│   ├── celulas_informe.html
│   ├── celulas_informe_form.html
│   ├── reportes.html
│   ├── usuarios.html
│   ├── usuario_form.html
│   ├── cambiar_clave.html
│   └── recuperar.html
└── static/            ← CSS/JS si tiene
```

### Carpeta `sql/`:
```
sql/
├── 01_schema.sql
├── 02_seed_chile.sql
├── 03_reportes.sql
├── 04_datos_prueba.sql
├── 05_liderazgo_pastoral.sql
├── 06_seguimiento.sql
├── 07_visitas.sql
├── 08_relacion_seg_vis.sql
├── 09_visita_rut.sql
├── 10_roles_permisos.sql
└── 11_primer_login.sql
```

---

## CONFIGURACIÓN `requirements.txt` (actualizar)

```
Flask>=3.0
gunicorn>=23.0
psycopg2-binary>=2.9
python-dotenv>=1.0
openpyxl>=3.1
fpdf2>=2.8
Werkzeug>=3.0
google-api-python-client>=2.0
google-auth>=2.0
```

---

## VARIABLES DE ENTORNOS (ejemplo production)

```bash
# Base de datos
DB_HOST=tu_host_postgres
DB_PORT=5432
DB_NAME=iglesia_mcci
DB_USER=tu_usuario
DB_PASSWORD=tu_password_seguro

# Flask
SECRET_KEY=tu_secret_key_muy_largo_123
FLASK_ENV=production

# Email
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=587
MAIL_USE_TLS=true
MAIL_USERNAME=tu_correo@gmail.com
MAIL_PASSWORD=tu_clave_de_aplicacion
MAIL_DEFAULT_SENDER=SMCCI <tu_correo@gmail.com>
```

---

## RECOMENDACIÓN

| Opción | Costo | Dificultad | Ideal para |
|--------|-------|------------|------------|
| **Render** | Gratis/$7 | Fácil | Empezar gratis |
| **Railway** | $5 | Fácil | Proyectos pequeños |
| **PythonAnywhere** | Gratis/$10 | Muy fácil | Aprender |
| **VPS** | $4-6 | Media | Control total |

**Para tu caso:** Empieza con **Render** (gratis) o **Railway** ($5). Son los más fáciles y no necesitas configurar servidor.
