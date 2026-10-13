"""SMCCI - Notificación diaria de seguimiento. Ejecutar todos los días 08:00.
Uso: py notificar.py  (lee .env para BD y SMTP opcional)
Si hay SMTP configurado envía correo, si no solo imprime + exporta Excel.
"""
import os, smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv
from openpyxl import Workbook
from datetime import date

load_dotenv()
def db():
    return psycopg2.connect(host=os.environ.get("DB_HOST", "localhost"), port=os.environ.get("DB_PORT", "5432"),
        dbname=os.environ.get("DB_NAME", "Iglesia_MCCI"), user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"], cursor_factory=RealDictCursor)

with db() as c, c.cursor() as cur:
    cur.execute("SELECT * FROM v_seguimiento_total")
    total = cur.fetchall()
    cur.execute("SELECT h.nombre_completo,l.red FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.estado_celula='SIN CELULA' AND l.activo")
    lsin = cur.fetchall()

print(f"[{date.today()}] Seguimiento total: {len(total)} (hermanos+visitas) | Lideres sin celula: {len(lsin)}")
for x in total:
    print(f" - [{x['tipo']}] {x['nombre_completo']} tel={x['telefono']} invito={x['invitado_por']} ultimo={x['ultimo_seg']}")

# Excel diario
wb = Workbook(); ws = wb.active; ws.title = "Seguimiento"
ws.append(["Tipo", "Nombre", "Telefono", "Invitado por", "Ultimo seguimiento"])
for x in total:
    ws.append([x["tipo"], x["nombre_completo"], x["telefono"], x["invitado_por"], str(x["ultimo_seg"] or "")])
out = f"seguimiento_{date.today()}.xlsx"
wb.save(out)
print("Excel:", out)

# Correo (opcional: configurar SMTP_HOST, SMTP_USER, SMTP_PASS, AVISO_PARA en .env)
host, para = os.environ.get("SMTP_HOST"), os.environ.get("AVISO_PARA")
if host and para and total:
    msg = MIMEMultipart(); msg["Subject"] = f"SMCCI seguimiento {date.today()} ({len(total)} pendientes)"
    msg["From"] = os.environ.get("SMTP_USER", "smcci"); msg["To"] = para
    cuerpo = "\n".join([f"- [{x['tipo']}] {x['nombre_completo']} ({x['telefono']}) invito: {x['invitado_por']}" for x in total])
    msg.attach(MIMEText(f"Pendientes (hermanos+visitas):\n{cuerpo}", "plain", "utf-8"))
    with smtplib.SMTP(host, int(os.environ.get("SMTP_PORT", "587"))) as s:
        s.starttls(); s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"]); s.send_message(msg)
    print("Correo enviado a", para)
else:
    print("Sin SMTP: revise la bandeja /seguimiento en la web.")
