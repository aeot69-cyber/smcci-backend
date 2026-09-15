# GUÍA: Sincronizar Google Forms con la BD

## REQUISITOS
- Cuenta de Google (con acceso a Google Cloud)
- El formulario de Google debe estar conectado a Google Sheets

---

## PASO 1: Configurar Google Cloud (1 vez)

1. Ir a **https://console.cloud.google.com/**
2. Crear proyecto nuevo: "SMCCI Sync"
3. Ir a **APIs y servicios → Habilitar API**
4. Buscar y habilitar: **Google Sheets API**
5. Ir a **Credenciales → Crear credenciales → Cuenta de servicio**
6. Nombre: "smcci-sync"
7. Copiar el email que aparece (algo como: smcci-sync@proyecto.iam.gserviceaccount.com)
8. Hacer clic en **Crear clave → JSON → Descargar**
9. Guardar el archivo como `credentials_google.json` en la carpeta `backend/`

---

## PASO 2: Compartir la hoja de Google Sheets

1. Abrir la hoja de cálculo del formulario
2. Hacer clic en **Compartir**
3. Pegar el email del Service Account (el que copiaste en Paso 1.7)
4. Dar permiso de **Editor**
5. Hacer clic en **Enviar**

---

## PASO 3: Obtener el ID de la hoja

1. Abrir la hoja de cálculo
2. Copiar el ID de la URL (está entre /d/ y /edit):
   ```
   https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
   ```
3. Abrir el archivo `sync_sheets.py`
4. Reemplazar `TU_SPREADSHEET_ID_AQUI` con el ID copiado

---

## PASO 4: Verificar el orden de columnas

1. Abrir la hoja de Google Sheets
2. Ver el orden de las columnas (A, B, C, etc.)
3. Comparar con el diccionario `COLUMN_MAP` en `sync_sheets.py`
4. Ajustar si es necesario

---

## PASO 5: Probar la sincronización

1. Abrir terminal en la carpeta `backend/`
2. Ejecutar: `py sync_sheets.py`
3. Verificar que diga "Insertados: X"

---

## PASO 6: Automatizar (opcional)

### Windows Task Scheduler:
1. Abrir "Programador de tareas"
2. Crear tarea nueva
3. Nombre: "SMCCI Sync Sheets"
4. Acción: Ejecutar `sincronizar.bat`
5. Triggers: Cada lunes a las 23:00 (después de que envíen el formulario)

### O ejecutar manualmente:
- Hacer doble clic en `sincronizar.bat`

---

## ESTRUCTURA DE LA HOJA (Referencia)

| Columna | Campo |
|---------|-------|
| A | Timestamp |
| B | NOMBRE DEL LIDER |
| C | NOMBRE DE TIMOTEO |
| D | NOMBRE DEL ANFITRION |
| E | LIDER O PASTOR A CARGO |
| F | GENERO DE LA CELULA |
| G | DIA DE CELULA (fecha) |
| H | MES |
| I | RED EN LA QUE PARTICIPA |
| J | DIRECCION |
| K | COMUNA |
| L | HORARIO |
| M | REALIZO SU CELULA (SI/NO) |
| N | JUSTIFICACION |
| O | DISCIPULOS QUE ASISTIERON |
| P | DISCIPULOS QUE NO ASISTIERON |
| Q | TOTAL OFRENDA |
| R | OFRENDA (tipo) |

---

## SOLUCIÓN DE PROBLEMAS

### Error "No se encontro credentials_google.json"
- Verificar que el archivo esté en `backend/`
- Verificar que se descargó correctamente de Google Cloud

### Error "The caller does not have permission"
- La hoja no está compartida con el Service Account
- Volver al Paso 2

### Error "Unable to find the spreadsheet"
- El SPREADSHEET_ID está incorrecto
- Verificar el ID en la URL de la hoja

### No inserta datos
- Verificar el orden de columnas (Paso 4)
- Verificar que los nombres de líder coincidan con la BD
