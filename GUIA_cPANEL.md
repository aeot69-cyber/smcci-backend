# GUÍA: Subir SMCCI a cPanel (Hosting Compartido)

## REQUISITOS
- Hosting con soporte Python (cPanel con "Setup Python App")
- Acceso a cPanel
- PostgreSQL en el hosting (o externo)

---

## PASO 1: Crear subdominio (recomendado)

1. En cPanel → **Subdominios**
2. Crear: `smcci.tudominio.com`
3. Document Root: `public_html/smcci`

---

## PASO 2: Crear aplicación Python en cPanel

1. En cPanel → **Software** → **Setup Python App**
2. Click **Create Application**
3. Configurar:
   - **Python Version:** 3.11 (o la más alta disponible)
   - **Application Root:** `smcci` (o la ruta de tu app)
   - **Application URL:** `smcci.tudominio.com`
   - **Application Startup File:** `passenger_wsgi.py`
   - **Application Entry Point:** `application`
4. Click **Create**

---

## PASO 3: Subir archivos

### Método 1: File Manager
1. En cPanel → **File Manager**
2. Ir a `public_html/smcci/`
3. Subir TODOS los archivos de `backend/`:
   - `app.py`
   - `requirements.txt`
   - `templates/` (carpeta completa)
   - `notificar.py`
   - `sync_sheets.py`

### Método 2: FTP
Usar FileZilla o similar:
```
Host: tudominio.com
User: tu_usuario
Password: tu_contraseña
Port: 21
```

---

## PASO 4: Crear archivo WSGI

1. En File Manager, crear archivo `passenger_wsgi.py` en la raíz
2. Contenido:

```python
import sys
import os

# Agregar la ruta de la app
sys.path.insert(0, os.path.dirname(__file__))

# Cargar variables de entorno
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))

# Importar la app
from app import app as application
```

---

## PASO 5: Crear archivo .env

1. Crear archivo `.env` en la raíz
2. Contenido (ajustar con tus datos):

```bash
# Base de datos (ajustar con datos del hosting)
DB_HOST=localhost
DB_PORT=5432
DB_NAME=tu_base_datos
DB_USER=tu_usuario
DB_PASSWORD=tu_password

# Flask
SECRET_KEY=tu_secret_key_muy_largo_aqui
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

## PASO 6: Crear PostgreSQL en cPanel

1. En cPanel → **Databases** → **PostgreSQL Databases**
2. Crear base de datos: `tu_usuario_iglesia_mcci`
3. Crear usuario de BD
4. Agregar usuario a la base con permisos ALL
5. Anotar: host, nombre, usuario, contraseña

---

## PASO 7: Instalar dependencias

1. En cPanel → **Terminal** (o SSH)
2. Activar entorno virtual:
```bash
source /home/tu_usuario/virtualenv/smcci/3.11/bin/activate
```
3. Instalar paquetes:
```bash
cd /home/tu_usuario/public_html/smcci
pip install -r requirements.txt
```

---

## PASO 8: Migrar base de datos

1. En cPanel → **phpMyAdmin** o **PostgreSQL Databases**
2. Importar cada archivo SQL en orden:
   - `01_schema.sql`
   - `02_seed_chile.sql`
   - `03_reportes.sql`
   - `04_datos_prueba.sql`
   - `05_liderazgo_pastoral.sql`
   - `06_seguimiento.sql`
   - `07_visitas.sql`
   - `08_relacion_seg_vis.sql`
   - `09_visita_rut.sql`
   - `10_roles_permisos.sql`
   - `11_primer_login.sql`

---

## PASO 9: Variables de entorno en cPanel

1. En **Setup Python App** → tu app
2. Click **Environment Variables**
3. Agregar cada variable:
   - `DB_HOST` = `localhost`
   - `DB_PORT` = `5432`
   - `DB_NAME` = `tu_base_datos`
   - `DB_USER` = `tu_usuario`
   - `DB_PASSWORD` = `tu_password`
   - `SECRET_KEY` = `tu_secret_key`
   - `FLASK_ENV` = `production`
   - etc.

---

## PASO 10: Reiniciar aplicación

1. En **Setup Python App** → tu app
2. Click **Restart**
3. Visitar `https://smcci.tudominio.com`

---

## ESTRUCTURA FINAL DE ARCHIVOS

```
public_html/smcci/
├── passenger_wsgi.py      ← Archivo WSGI
├── app.py                 ← App principal
├── requirements.txt       ← Dependencias
├── .env                   ← Variables de entorno
├── sync_sheets.py         ← Sync Google Sheets
├── notificar.py           ← Notificaciones
├── templates/
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
```

---

## SOLUCIÓN DE PROBLEMAS

### Error 500 Internal Server:
- Verificar `passenger_wsgi.py`
- Revisar `error_log` en cPanel
- Verificar variables de entorno

### Error "Module Not Found":
- Ejecutar `pip install` en el entorno virtual
- Verificar `requirements.txt`

### Error de BD:
- Verificar datos de conexión en `.env`
- Verificar que PostgreSQL esté habilitado

### La app no carga:
- Verificar que `passenger_wsgi.py` esté en la raíz
- Reiniciar la app en cPanel

---

## NOTAS IMPORTANTES

1. **NO subir** la carpeta `__pycache__/` ni archivos `.pyc`
2. **NO subir** tu entorno virtual local (`venv/`)
3. El hosting crea su propio entorno virtual
4. Las variables de entorno se configuran en cPanel, NO en `.env` para producción
5. Si tu hosting no tiene PostgreSQL, usar MySQL (cambiar `psycopg2` por `pymysql`)
