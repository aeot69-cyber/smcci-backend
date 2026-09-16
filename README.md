# SMCCI - Sistema de Consolidación MCCI

Sistema web para gestión de hermanos, líderes, células, discipulado, encuentros, visitas y reportes.

## Tecnologías
- Python 3.11
- Flask
- PostgreSQL
- Bootstrap 5

## Despliegue en Render
1. Subir estos archivos a GitHub
2. Crear PostgreSQL en Render
3. Crear Web Service conectado a GitHub
4. Configurar variables de entorno
5. Migrar datos SQL

## Variables de entorno necesarias
```
DB_HOST=
DB_PORT=5432
DB_NAME=iglesia_mcci
DB_USER=mcci
DB_PASSWORD=
SECRET_KEY=
FLASK_ENV=production
```
