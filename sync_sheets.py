"""
Sincronizar Google Sheets → PostgreSQL (Informe de Célula)
Ejecutar con: py sync_sheets.py
"""
import os
import sys
import json
from datetime import datetime

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    print("Instalando dependencias de Google...")
    os.system("pip install google-api-python-client google-auth")
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

import psycopg2
from dotenv import load_dotenv

load_dotenv()

# ============ CONFIGURACIÓN ============
# ID de la hoja de Google Sheets (está en la URL después de /d/)
SPREADSHEET_ID = "TU_SPREADSHEET_ID_AQUI"  # <-- Cambiar
# Nombre de la pestaña (hoja)
RANGE_NAME = "Respuestas de formulario!A1:Z1000"
# Archivo de credenciales JSON de Google
CREDENTIALS_FILE = "credentials_google.json"  # <-- Poner en backend/

# ============ MAPEO DE COLUMNAS ============
# Ajustar según el orden de columnas de tu Google Sheet
COLUMN_MAP = {
    0: "fecha_respuesta",      # Timestamp
    1: "nombre_lider",         # NOMBRE DEL LIDER
    2: "nombre_timoteo",       # NOMBRE DE TIMOTEO
    3: "nombre_anfitrion",     # NOMBRE DEL ANFITRION
    4: "pastor_cargo",         # LIDER O PASTOR A CARGO
    5: "genero_celula",        # GENERO DE LA CELULA
    6: "dia_celula",           # DIA DE CELULA (fecha)
    7: "mes",                  # MES
    8: "red",                  # RED EN LA QUE PARTICIPA
    9: "direccion",            # DIRECCION
    10: "comuna",              # COMUNA
    11: "horario",             # HORARIO
    12: "realizado",           # REALIZO SU CELULA (SI/NO)
    13: "justificacion_no",    # JUSTIFICACION
    14: "discipulos_asistieron",  # DISCIPULOS QUE ASISTIERON
    15: "discipulos_no_asistieron",  # DISCIPULOS QUE NO ASISTIERON
    16: "total_ofrenda",       # TOTAL OFRENDA
    17: "tipo_ofrenda",        # OFRENDA (tipo)
}

# ============ FUNCIONES ============
def get_google_sheets_service():
    """Obtiene el servicio de Google Sheets."""
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"ERROR: No se encontro {CREDENTIALS_FILE}")
        print("Pasos:")
        print("1. Ir a console.cloud.google.com")
        print("2. Crear Service Account")
        print("3. Descargar JSON como credentials_google.json")
        print("4. Compartir la hoja con el email del Service Account")
        sys.exit(1)
    
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=scopes)
    return build('sheets', 'v4', credentials=creds)

def get_db():
    """Conexion a PostgreSQL."""
    return psycopg2.connect(
        host=os.environ.get("DB_HOST", "localhost"),
        port=int(os.environ.get("DB_PORT", 5432)),
        dbname=os.environ.get("DB_NAME", "Iglesia_MCCI"),
        user=os.environ.get("DB_USER", "mcci"),
        password=os.environ.get("DB_PASSWORD", "mcci12")
    )

def limpiar_rut(rut):
    """Quita puntos del RUT."""
    if rut:
        return rut.replace(".", "").replace(" ", "")
    return rut

def parse_fecha(fecha_str):
    """Convierte string de fecha de Google Sheets a formato SQL."""
    if not fecha_str:
        return None
    try:
        # Google Sheets puede enviar fecha como "2026-09-15" o "15/09/2026"
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(fecha_str.strip(), fmt).date()
            except ValueError:
                continue
    except:
        pass
    return None

def parse_decimal(valor):
    """Convierte string a decimal."""
    if not valor:
        return 0
    try:
        return float(str(valor).replace(",", "").replace("$", "").strip())
    except:
        return 0

def normalize_genero(gen):
    """Normaliza género de célula."""
    gen = gen.upper().strip() if gen else ""
    if "HOMBRE" in gen or "MASCULINO" in gen:
        return "HOMBRES"
    if "MUJER" in gen or "FEMENINO" in gen:
        return "MUJERES"
    return gen if gen in ("HOMBRES", "MUJERES") else "HOMBRES"

def normalize_red(red):
    """Normaliza red."""
    red = red.upper().strip() if red else ""
    if "ADOLE" in red or "TEEN" in red:
        return "ADOLESCENTES"
    if "JOVEN" in red or "YOUTH" in red:
        return "JOVENES"
    if "MUJER" in red:
        return "MUJERES"
    if "HOMBRE" in red:
        return "HOMBRES"
    return red if red in ("ADOLESCENTES", "JOVENES", "MUJERES", "HOMBRES") else "HOMBRES"

def normalize_mes(mes):
    """Normaliza mes."""
    mes = mes.upper().strip() if mes else ""
    meses_validos = ['JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
    if mes in meses_validos:
        return mes
    # Intentar mapear
    meses_map = {
        'JUN': 'JUNIO', 'JUL': 'JULIO', 'AGO': 'AGOSTO', 'SEP': 'SEPTIEMBRE',
        'OCT': 'OCTUBRE', 'NOV': 'NOVIEMBRE', 'DIC': 'DICIEMBRE'
    }
    return meses_map.get(mes[:3], mes)

def normalize_ofrenda(tipo):
    """Normaliza tipo de ofrenda."""
    tipo = tipo.upper().strip() if tipo else ""
    if "EFECT" in tipo:
        return "EFECTIVO"
    if "TRANSF" in tipo:
        return "TRANSFERENCIA"
    if "AMBOS" in tipo or "TODOS" in tipo:
        return "AMBOS"
    if "NO" in tipo:
        return "NO SE HACE"
    return tipo if tipo in ("EFECTIVO","TRANSFERENCIA","AMBOS","NO SE HACE") else "EFECTIVO"

def buscar_lider_id(conn, nombre_lider):
    """Busca el ID del líder por nombre en la BD."""
    if not nombre_lider:
        return None
    cur = conn.cursor()
    cur.execute("""
        SELECT l.id FROM lideres l 
        JOIN hermanos h ON h.id=l.hermano_id 
        WHERE h.nombre_completo ILIKE %s AND l.activo
        LIMIT 1
    """, (f"%{nombre_lider.strip()}%",))
    row = cur.fetchone()
    return row[0] if row else None

def sincronizar():
    """Función principal de sincronización."""
    print("=" * 50)
    print("  SINCRONIZAR Google Sheets → PostgreSQL")
    print("=" * 50)
    print()
    
    # 1. Conectar a Google Sheets
    print("1. Conectando a Google Sheets...")
    try:
        service = get_google_sheets_service()
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])
        print(f"   OK: {len(values)} filas leidas")
    except Exception as e:
        print(f"   ERROR: {e}")
        return
    
    if len(values) < 2:
        print("   No hay datos para sincronizar")
        return
    
    # 2. Conectar a BD
    print("2. Conectando a PostgreSQL...")
    try:
        conn = get_db()
        cur = conn.cursor()
        print("   OK")
    except Exception as e:
        print(f"   ERROR: {e}")
        return
    
    # 3. Procesar filas (saltar cabecera)
    print("3. Procesando datos...")
    headers = values[0]
    rows = values[1:]
    insertados = 0
    duplicados = 0
    errores = 0
    
    for i, row in enumerate(rows):
        try:
            # Mapear columnas
            data = {}
            for idx, col_name in COLUMN_MAP.items():
                data[col_name] = row[idx] if idx < len(row) else ""
            
            # Verificar duplicado (misma fecha + mismo líder)
            fecha = parse_fecha(data.get("dia_celula", ""))
            nombre_lider = data.get("nombre_lider", "").strip()
            
            if fecha and nombre_lider:
                cur.execute("""
                    SELECT id FROM celulas_informe 
                    WHERE fecha_dia=%s AND nombre_lider ILIKE %s
                """, (fecha, nombre_lider))
                if cur.fetchone():
                    duplicados += 1
                    continue
            
            # Buscar líder en BD
            lider_id = buscar_lider_id(conn, nombre_lider)
            
            # Normalizar valores
            realizado = data.get("realizado", "SI").upper().strip()
            realizado = realizado in ("SI", "SÍ", "YES", "TRUE", "1")
            
            # Insertar
            cur.execute("""
                INSERT INTO celulas_informe(
                    lider_id, nombre_lider, nombre_timoteo, nombre_anfitrion,
                    pastor_cargo, genero_celula, fecha_dia, mes, red,
                    direccion, comuna, horario, realizado, justificacion_no,
                    discipulos_asistieron, discipulos_no_asistieron,
                    total_ofrenda, tipo_ofrenda
                ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                lider_id,
                nombre_lider,
                data.get("nombre_timoteo", "").strip(),
                data.get("nombre_anfitrion", "").strip(),
                data.get("pastor_cargo", "").strip(),
                normalize_genero(data.get("genero_celula", "")),
                fecha,
                normalize_mes(data.get("mes", "")),
                normalize_red(data.get("red", "")),
                data.get("direccion", "").strip(),
                data.get("comuna", "").strip(),
                data.get("horario", "").strip(),
                realizado,
                data.get("justificacion_no", "").strip(),
                data.get("discipulos_asistieron", "").strip(),
                data.get("discipulos_no_asistieron", "").strip(),
                parse_decimal(data.get("total_ofrenda", "0")),
                normalize_ofrenda(data.get("tipo_ofrenda", ""))
            ))
            insertados += 1
            
        except Exception as e:
            errores += 1
            print(f"   Error fila {i+2}: {e}")
    
    conn.commit()
    
    # 4. Resumen
    print()
    print("=" * 50)
    print("  RESUMEN")
    print("=" * 50)
    print(f"  Total filas Google Sheets: {len(rows)}")
    print(f"  Insertados: {insertados}")
    print(f"  Duplicados (omitidos): {duplicados}")
    print(f"  Errores: {errores}")
    print()
    
    cur.close()
    conn.close()
    print("Sincronizacion completada!")

if __name__ == "__main__":
    sincronizar()
