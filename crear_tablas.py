"""
SMCCI - Creador de tablas Postgres.
Pide usuario y contraseña por consola (no quedan en código).
Uso: python crear_tablas.py
"""
import getpass
import psycopg2
from pathlib import Path

SQL_DIR = Path(__file__).resolve().parent.parent / "sql"

def conectar():
    print("=== SMCCI / Conexión PostgreSQL ===")
    host = input("Host [localhost]: ").strip() or "localhost"
    port = input("Puerto [5432]: ").strip() or "5432"
    db = input("Base de datos [smcci]: ").strip() or "smcci"
    user = input("Usuario Postgres: ").strip()
    pwd = getpass.getpass("Contraseña Postgres: ")
    conn = psycopg2.connect(host=host, port=port, dbname=db, user=user, password=pwd)
    conn.autocommit = True
    return conn

def ejecutar(conn, archivo):
    sql = (SQL_DIR / archivo).read_text(encoding="utf-8")
    with conn.cursor() as cur:
        cur.execute(sql)
    print(f"OK: {archivo}")

if __name__ == "__main__":
    conn = conectar()
    try:
        ejecutar(conn, "01_schema.sql")
        ejecutar(conn, "02_seed_chile.sql")
        # 03_reportes.sql son consultas, no se ejecuta como schema
        print("Tablas + catálogos creados. Verifique con: SELECT * FROM v_lideres_conteo;")
    finally:
        conn.close()
