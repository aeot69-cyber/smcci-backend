"""SMCCI - Sistema Consolidación MCCI (Flask + Postgres). CRUD + Reportes con menú."""
import os
from functools import wraps
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv
from flask import Flask, request, jsonify, session, redirect, url_for, render_template, flash, abort, send_file
from werkzeug.security import check_password_hash
from io import BytesIO
from openpyxl import Workbook
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()
app = Flask(__name__)

def enviar_email(destinatario, asunto, cuerpo_html):
    """Envía un email HTML usando SMTP."""
    try:
        msg = MIMEMultipart()
        msg['From'] = os.environ.get("MAIL_DEFAULT_SENDER", "SMCCI <noreply@mcci.cl>")
        msg['To'] = destinatario
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_html, 'html'))
        server = smtplib.SMTP(os.environ.get("MAIL_SERVER","smtp.gmail.com"), int(os.environ.get("MAIL_PORT",587)))
        server.starttls()
        server.login(os.environ.get("MAIL_USERNAME",""), os.environ.get("MAIL_PASSWORD",""))
        server.sendmail(msg['From'], destinatario, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando email: {e}")
        return False
app.secret_key = os.environ.get("SECRET_KEY", "cambiar-en-produccion")
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("FLASK_ENV") == "production", PERMANENT_SESSION_LIFETIME=3600)

def db():
    return psycopg2.connect(host=os.environ.get("DB_HOST", "localhost"),
        port=os.environ.get("DB_PORT", "5432"), dbname=os.environ.get("DB_NAME", "Iglesia_MCCI"),
        user=os.environ["DB_USER"], password=os.environ["DB_PASSWORD"], cursor_factory=RealDictCursor)

def q(sql, p=(), one=False, commit=False):
    with db() as c, c.cursor() as cur:
        cur.execute(sql, p)
        if commit: c.commit(); return cur.rowcount
        return cur.fetchone() if one else cur.fetchall()

def login_req(fn=None, *, roles=None):
    def deco(f):
        @wraps(f)
        def inner(*a, **kw):
            if "uid" not in session: return redirect(url_for("login"))
            if roles and session.get("rol") not in roles:
                abort(403)
            return f(*a, **kw)
        return inner
    if fn is not None:
        return deco(fn)
    return deco

def permiso_req(modulo, accion="leer"):
    def deco(fn):
        @wraps(fn)
        def inner(*a, **kw):
            if not tiene_permiso(session.get("rol",""), modulo, accion):
                flash("No tiene permiso para esta acción", "error")
                return redirect(url_for("index"))
            return fn(*a, **kw)
        return inner
    return deco

def permisos(rol):
    rows = q("SELECT modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar FROM roles_permisos WHERE rol=%s", (rol,))
    return {r["modulo"]: r for r in rows}

def tiene_permiso(rol, modulo, accion="leer"):
    p = permisos(rol).get(modulo)
    if not p: return False
    return bool(p.get(f"puede_{accion}", False))

def limpiar_rut(rut):
    """Quita puntos del RUT y deja guión: 10.829.271-7 → 10829271-7"""
    return rut.replace(".", "").replace(" ", "")

import re
def validar_clave(clave):
    """Valida que la contraseña cumpla requisitos de seguridad."""
    errores = []
    if len(clave) < 8:
        errores.append("mínimo 8 caracteres")
    if not re.search(r'[A-Z]', clave):
        errores.append("al menos 1 letra mayúscula")
    if not re.search(r'[a-z]', clave):
        errores.append("al menos 1 letra minúscula")
    if not re.search(r'[0-9]', clave):
        errores.append("al menos 1 número")
    if not re.search(r'[!@#$%^&*(),.?\":{}|<>_\-+=\[\]\\;\'~`/]', clave):
        errores.append("al menos 1 signo especial (!@#$%^&*...)")
    return errores

def generar_clave_temporal():
    """Genera una contraseña temporal segura y fácil de copiar."""
    import secrets, string
    mayus = secrets.choice(string.ascii_uppercase)
    minus = ''.join(secrets.choice(string.ascii_lowercase) for _ in range(3))
    numero = ''.join(secrets.choice(string.digits) for _ in range(2))
    signo = secrets.choice("!@#")
    resto = secrets.choice(string.ascii_lowercase)
    return f"{mayus}{minus}{numero}{signo}{resto}"

@app.context_processor
def inject_permisos():
    rol = session.get("rol", "")
    p = permisos(rol)
    modulos_menu = [
        ("hermanos",    "Hermanos",    "🙋", "hermanos",    "/hermanos",    "/hermanos/nuevo"),
        ("lideres",     "Líderes",     "👥", "lideres",     "/lideres",     "/lideres/nuevo"),
        ("discipulado", "Discipulado", "🤝", "discipulado", "/discipulado", "/discipulado/nuevo"),
        ("encuentros",  "Encuentros",  "🔥", "encuentros",  "/encuentros",  "/encuentros/nuevo"),
        ("seguimiento", "Seguimiento", "🔔", "seguimiento", "/seguimiento", None),
        ("visitas",     "Visitas",     "👋", "visitas",     "/visitas",     "/visitas/nuevo"),
        ("celulas",     "Informe Célula", "📋", "celulas",   "/celulas-informe", "/celulas-informe/nuevo"),
        ("reportes",    "Reportes",    "📈", "reportes",    "/reportes",    None),
        ("usuarios",    "Usuarios",    "⚙️", "usuarios",    "/usuarios",    "/usuarios/nuevo"),
    ]
    menu = "<a href='/'>🏠 INICIO</a>"
    for mod, nombre, icono, key, url_list, url_nuevo in modulos_menu:
        pk = p.get(key, {})
        if not pk.get("puede_leer"): continue
        menu += f"<button class='mbtn' data-bs-toggle='collapse' data-bs-target='#m-{key}'>{icono} {nombre} ▾</button>"
        menu += f"<div id='m-{key}' class='collapse sub'>"
        menu += f"<a href='{url_list}'>🔍 Consultar</a>"
        if pk.get("puede_crear") and url_nuevo:
            menu += f"<a href='{url_nuevo}'>➕ Ingresar</a>"
        if pk.get("puede_exportar"):
            menu += f"<a href='{url_list}?export=excel' target='_blank'>📊 Excel</a>"
            menu += f"<a href='{url_list}?export=pdf' target='_blank'>📄 PDF</a>"
        menu += "</div>"
    return dict(perm=p, rol=rol, menu=menu, tiene_permiso=lambda m, a="leer": tiene_permiso(rol, m, a))

def to_excel(rows, headers, filename):
    wb = Workbook(); ws = wb.active; ws.title = "Reporte"
    ws.append(headers)
    for r in rows: ws.append([("" if r.get(h) is None else str(r.get(h))) for h in headers])
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def to_pdf(title, rows, headers, filename):
    pdf = FPDF(); pdf.add_page(); pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, title[:80], new_x="LMARGIN", new_y="NEXT"); pdf.set_font("Helvetica", "", 8)
    w = max(190 // max(len(headers), 1), 25)
    for h in headers: pdf.cell(w, 8, str(h)[:20], border=1)
    pdf.ln()
    for r in rows:
        for h in headers: pdf.cell(w, 7, str(r.get(h) or "")[:22], border=1)
        pdf.ln()
    buf = BytesIO(pdf.output())
    return send_file(buf, as_attachment=True, download_name=filename, mimetype="application/pdf")

# ---------- AUTH ----------
@app.get("/login")
def login(): return render_template("login.html")

@app.post("/login")
def login_post():
    u, p = request.form.get("username", ""), request.form.get("password", "")
    row = q("SELECT id,password_hash,rol,primer_login FROM usuarios WHERE username=%s AND activo", (u,), one=True)
    if not row or not check_password_hash(row["password_hash"], p):
        flash("Usuario o clave inválidos", "error"); return redirect(url_for("login"))
    session["uid"], session["rol"], session["user"] = row["id"], row["rol"], u
    q("UPDATE usuarios SET ultimo_login=now() WHERE id=%s", (row["id"],), commit=True)
    if row.get("primer_login"):
        flash("Es su primer ingreso. Debe cambiar su contraseña.", "ok")
        return redirect(url_for("cambiar_clave"))
    return redirect(url_for("index"))

@app.get("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

# ---------- CAMBIO DE CONTRASEÑA ----------
@app.get("/cambiar-clave")
def cambiar_clave():
    if "uid" not in session: return redirect(url_for("login"))
    return render_template("cambiar_clave.html")

@app.post("/cambiar-clave")
def cambiar_clave_post():
    if "uid" not in session: return redirect(url_for("login"))
    from werkzeug.security import generate_password_hash, check_password_hash
    f = request.form
    actual = f.get("actual","")
    nueva = f.get("nueva","")
    confirmar = f.get("confirmar","")
    if nueva != confirmar:
        flash("Las contraseñas no coinciden", "error")
        return redirect(url_for("cambiar_clave"))
    errores = validar_clave(nueva)
    if errores:
        flash(f"La contraseña no cumple: {', '.join(errores)}", "error")
        return redirect(url_for("cambiar_clave"))
    row = q("SELECT password_hash FROM usuarios WHERE id=%s", (session["uid"],), one=True)
    if not row or not check_password_hash(row["password_hash"], actual):
        flash("La contraseña actual es incorrecta", "error")
        return redirect(url_for("cambiar_clave"))
    q("UPDATE usuarios SET password_hash=%s, primer_login=FALSE WHERE id=%s",
      (generate_password_hash(nueva), session["uid"]), commit=True)
    flash("Contraseña actualizada correctamente", "ok")
    return redirect(url_for("index"))

# ---------- RECUPERAR CONTRASEÑA ----------
@app.get("/recuperar")
def recuperar(): return render_template("recuperar.html")

@app.post("/recuperar")
def recuperar_post():
    from werkzeug.security import generate_password_hash
    u = request.form.get("username","").strip()
    row = q("""SELECT u.id, u.username, u.lider_id, h.nombre_completo AS nombre_user
        FROM usuarios u LEFT JOIN lideres l ON l.id=u.lider_id
        LEFT JOIN hermanos h ON h.id=u.hermano_id
        WHERE u.username=%s AND u.activo""", (u,), one=True)
    if not row:
        flash("Usuario no encontrado o inactivo", "error")
        return redirect(url_for("recuperar"))
    if not row.get("lider_id"):
        flash("Este usuario no tiene un líder vinculado. Contacte al administrador para recuperar su contraseña.", "error")
        return redirect(url_for("recuperar"))
    lider = q("""SELECT h.nombre_completo, h.correo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.id=%s""", (row["lider_id"],), one=True)
    if not lider or not lider.get("correo"):
        flash("El líder vinculado no tiene correo registrado. Contacte al administrador.", "error")
        return redirect(url_for("recuperar"))
    temp = generar_clave_temporal()
    q("UPDATE usuarios SET password_hash=%s, primer_login=TRUE WHERE id=%s",
      (generate_password_hash(temp), row["id"]), commit=True)
    asunto = "SMCCI - Recuperación de Contraseña"
    cuerpo = f"""<html><body style="font-family:Arial,sans-serif">
    <div style="max-width:500px;margin:auto;padding:20px;border:1px solid #ddd;border-radius:10px">
    <h2 style="color:#1a237e;text-align:center">Recuperación de Contraseña</h2>
    <p>Hola <b>{lider['nombre_completo']}</b>,</p>
    <p>Se solicitó una recuperación de contraseña para el usuario <b>{row['username']}</b>.</p>
    <p>Su contraseña temporal es:</p>
    <div style="background:#1a237e;color:#fff;padding:20px;text-align:center;border-radius:8px;margin:20px 0">
    <span style="font-family:monospace;font-size:28px;letter-spacing:4px;font-weight:bold">{temp}</span>
    </div>
    <p style="background:#fff3cd;padding:10px;border-radius:5px;border-left:4px solid #ffc107">
    <b>IMPORTANTE:</b> Copie la contraseña exactamente como aparece arriba.<br>
    Al ingresar, será <b>obligatorio cambiarla</b> por una nueva.</p>
    <p style="color:#888;font-size:12px">Si usted no solicitó este cambio, ignore este mensaje.</p>
    <hr><p style="text-align:center;color:#888;font-size:11px">SMCCI - Sistema de Consolidación</p>
    </div></body></html>"""
    if enviar_email(lider["correo"], asunto, cuerpo):
        flash(f"Se envió la contraseña temporal al correo de {lider['nombre_completo']} ({lider['correo']})", "ok")
    else:
        flash(f"Error al enviar el correo. Contraseña temporal: {temp} — Compártala de forma segura.", "error")
    return redirect(url_for("login"))

# ---------- USUARIOS (SUPERADMIN) ----------
MODULOS_INFO = [
    {"key":"hermanos","nombre":"Hermanos"},{"key":"lideres","nombre":"Líderes"},
    {"key":"discipulado","nombre":"Discipulado"},{"key":"encuentros","nombre":"Encuentros"},
    {"key":"seguimiento","nombre":"Seguimiento"},{"key":"visitas","nombre":"Visitas"},
    {"key":"reportes","nombre":"Reportes"},{"key":"usuarios","nombre":"Usuarios"},
    {"key":"celulas","nombre":"Informe Célula"},
]
ROLES_LIST = ["SUPERADMIN","ADMIN","CONSULTA"]

@app.get("/usuarios")
@login_req
@permiso_req("usuarios", "leer")
def usuarios():
    users = q("""SELECT u.*, h.nombre_completo AS h_nombre, l2.nombre_completo AS l_nombre
        FROM usuarios u LEFT JOIN hermanos h ON h.id=u.hermano_id LEFT JOIN lideres l ON l.id=u.lider_id
        LEFT JOIN hermanos l2 ON l2.id=l.hermano_id ORDER BY u.rol, u.username""")
    for u in users:
        u["vinculo"] = u.get("h_nombre") or u.get("l_nombre") or None
    perm = permisos(session.get("rol",""))
    perm_all = {}
    for rl in ROLES_LIST:
        perm_all[rl] = permisos(rl)
    return render_template("usuarios.html", usuarios=users, perm=perm_all, modulos=MODULOS_INFO, roles=ROLES_LIST)

@app.get("/usuarios/nuevo")
@login_req
@permiso_req("usuarios", "crear")
def usuario_nuevo():
    lideres = q("""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY h.nombre_completo""")
    return render_template("usuario_form.html", u=None, rol=session.get("rol",""), lideres=lideres)

@app.post("/usuarios/nuevo")
@login_req
@permiso_req("usuarios", "crear")
def usuario_crear():
    from werkzeug.security import generate_password_hash
    f = request.form
    username = f["username"].strip().lower()
    pwd = f.get("password","").strip()
    nuevo_rol = f["rol"]
    rol_actual = session.get("rol","")
    lider_id = f.get("lider_id") or None
    if not pwd:
        flash("La contraseña es obligatoria", "error"); return redirect(url_for("usuario_nuevo"))
    errores = validar_clave(pwd)
    if errores:
        flash(f"La contraseña no cumple: {', '.join(errores)}", "error"); return redirect(url_for("usuario_nuevo"))
    if rol_actual != "SUPERADMIN" and nuevo_rol == "SUPERADMIN":
        flash("No tiene permiso para crear usuarios SUPERADMIN", "error"); return redirect(url_for("usuario_nuevo"))
    try:
        q("INSERT INTO usuarios(username,password_hash,rol,activo,primer_login,lider_id) VALUES(%s,%s,%s,%s,TRUE,%s)",
          (username, generate_password_hash(pwd), f["rol"], f.get("activo")=="on", lider_id), commit=True)
        flash("Usuario creado", "ok")
    except Exception as e: flash(f"Error: {e}", "error"); return redirect(url_for("usuario_nuevo"))
    return redirect(url_for("usuarios"))

@app.get("/usuarios/<int:uid>/editar")
@login_req
@permiso_req("usuarios", "editar")
def usuario_editar(uid):
    u = q("SELECT * FROM usuarios WHERE id=%s", (uid,), one=True)
    if not u: flash("Usuario no encontrado", "error"); return redirect(url_for("usuarios"))
    lideres = q("""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY h.nombre_completo""")
    return render_template("usuario_form.html", u=u, rol=session.get("rol",""), lideres=lideres)

@app.post("/usuarios/<int:uid>/editar")
@login_req
@permiso_req("usuarios", "editar")
def usuario_update(uid):
    from werkzeug.security import generate_password_hash
    f = request.form
    pwd = f.get("password","").strip()
    rol_actual = session.get("rol","")
    nuevo_rol = f["rol"]
    lider_id = f.get("lider_id") or None
    if rol_actual != "SUPERADMIN" and nuevo_rol == "SUPERADMIN":
        flash("No tiene permiso para asignar rol SUPERADMIN", "error")
        return redirect(url_for("usuarios"))
    if pwd:
        errores = validar_clave(pwd)
        if errores:
            flash(f"La contraseña no cumple: {', '.join(errores)}", "error")
            return redirect(url_for("usuario_editar", uid=uid))
    try:
        if pwd:
            q("UPDATE usuarios SET rol=%s, activo=%s, password_hash=%s, lider_id=%s WHERE id=%s",
              (f["rol"], f.get("activo")=="on", generate_password_hash(pwd), lider_id, uid), commit=True)
        else:
            q("UPDATE usuarios SET rol=%s, activo=%s, lider_id=%s WHERE id=%s",
              (f["rol"], f.get("activo")=="on", lider_id, uid), commit=True)
        flash("Usuario actualizado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:uid>/toggle")
@login_req
@permiso_req("usuarios", "editar")
def usuario_toggle(uid):
    q("UPDATE usuarios SET activo=NOT activo WHERE id=%s", (uid,), commit=True)
    flash("Estado actualizado", "ok"); return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:uid>/eliminar")
@login_req
@permiso_req("usuarios", "eliminar")
def usuario_eliminar(uid):
    u = q("SELECT rol FROM usuarios WHERE id=%s", (uid,), one=True)
    if u and u["rol"] == "SUPERADMIN":
        flash("No se puede eliminar un SUPERADMIN", "error")
    else:
        q("DELETE FROM usuarios WHERE id=%s", (uid,), commit=True)
        flash("Usuario eliminado", "ok")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/permisos")
@login_req
@permiso_req("usuarios", "editar")
def usuario_permisos():
    f = request.form
    for rol in ROLES_LIST:
        for mod in MODULOS_INFO:
            key = mod["prefix"] if "prefix" in mod else mod["key"]
            leer = f.get(f"{rol}_{key}_leer") is not None
            crear = f.get(f"{rol}_{key}_crear") is not None
            editar = f.get(f"{rol}_{key}_editar") is not None
            eliminar = f.get(f"{rol}_{key}_eliminar") is not None
            exportar = f.get(f"{rol}_{key}_exportar") is not None
            q("""INSERT INTO roles_permisos(rol,modulo,puede_leer,puede_crear,puede_editar,puede_eliminar,puede_exportar)
                VALUES(%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (rol,modulo) DO UPDATE SET
                puede_leer=%s,puede_crear=%s,puede_editar=%s,puede_eliminar=%s,puede_exportar=%s""",
              (rol, key, leer, crear, editar, eliminar, exportar, leer, crear, editar, eliminar, exportar), commit=True)
    flash("Permisos actualizados", "ok"); return redirect(url_for("usuarios"))

# ---------- DASHBOARD ----------
@app.get("/")
@login_req
def index():
    d = {}
    d["hermanos"] = q("SELECT COUNT(*) c FROM hermanos WHERE activo", one=True)["c"]
    d["lideres"] = q("SELECT COUNT(*) c FROM lideres WHERE activo", one=True)["c"]
    d["disc"] = q("SELECT COUNT(*) c FROM discipulado WHERE fecha_fin IS NULL", one=True)["c"]
    d["pend"] = q("SELECT COUNT(*) c FROM encuentro_participacion WHERE estado='PENDIENTE'", one=True)["c"]
    d["top"] = q("SELECT * FROM v_lideres_conteo ORDER BY discipulos_activos DESC LIMIT 10")
    d["cumple"] = q("""SELECT nombre_completo,fecha_nacimiento FROM hermanos WHERE activo
        AND EXTRACT(MONTH FROM fecha_nacimiento)=EXTRACT(MONTH FROM CURRENT_DATE) ORDER BY EXTRACT(DAY FROM fecha_nacimiento) LIMIT 20""")
    return render_template("index.html", d=d)

# ---------- HERMANOS ----------
@app.get("/hermanos")
@login_req
def hermanos():
    texto = request.args.get("q", "")
    exp = request.args.get("export", "")
    rows = q("SELECT id, rut AS rut, nombre_completo, comuna, region, correo, telefono FROM v_hermanos_completo WHERE nombre_completo ILIKE %s OR rut ILIKE %s ORDER BY nombre_completo LIMIT 500",
             (f"%{texto}%", f"%{texto}%"))
    if exp == "excel": return to_excel(rows, ["rut", "nombre_completo", "comuna", "region", "correo", "telefono"], "hermanos.xlsx")
    if exp == "pdf": return to_pdf("Hermanos MCCI", rows, ["rut", "nombre_completo", "comuna", "region"], "hermanos.pdf")
    return render_template("hermanos.html", rows=rows, texto=texto)

@app.get("/hermanos/nuevo")
@login_req
@permiso_req("hermanos", "crear")
def hermano_nuevo():
    com = q("SELECT c.id,c.nombre,r.nombre_corto AS region FROM comunas c JOIN regiones r ON r.id=c.region_id ORDER BY c.nombre")
    her = q("SELECT id,nombre_completo FROM hermanos WHERE activo ORDER BY 2")
    return render_template("hermano_form.html", h=None, comunas=com, hermanos=her)

@app.post("/hermanos/nuevo")
@login_req
@permiso_req("hermanos", "crear")
def hermano_crear():
    f = request.form
    try:
        rut = limpiar_rut(f["rut"])
        q("""INSERT INTO hermanos(rut,nombres,apellido_paterno,apellido_materno,fecha_nacimiento,fecha_aniversario,
            correo,telefono,direccion,comuna_id,sexo,estado_civil,invitado_por_id,invitado_por_texto) VALUES(%s,%s,%s,%s,NULLIF(%s,'')::date,NULLIF(%s,'')::date,
            NULLIF(%s,''),%s,%s,NULLIF(%s,'')::int,%s,%s,NULLIF(%s,'')::int,NULLIF(%s,''))""",
          (rut, f["nombres"].strip(), f["paterno"].strip(), f.get("materno") or None,
           f.get("fnac") or None, f.get("fani") or None, f.get("correo") or None,
           f.get("telefono"), f.get("direccion"), f.get("comuna") or None, f.get("sexo") or None,
           f.get("ecivil") or None, f.get("invitado") or None, f.get("invtexto") or None), commit=True)
        flash("Hermano creado", "ok")
    except Exception as e: flash(f"Error: {e}", "error"); return redirect(url_for("hermano_nuevo"))
    return redirect(url_for("hermanos"))

@app.get("/hermanos/<int:hid>/editar")
@login_req
@permiso_req("hermanos", "editar")
def hermano_editar(hid):
    h = q("SELECT * FROM hermanos WHERE id=%s", (hid,), one=True)
    com = q("SELECT c.id,c.nombre FROM comunas c ORDER BY c.nombre")
    her = q("SELECT id,nombre_completo FROM hermanos WHERE activo AND id!=%s ORDER BY 2", (hid,))
    return render_template("hermano_form.html", h=h, comunas=com, hermanos=her)

@app.post("/hermanos/<int:hid>/editar")
@login_req
@permiso_req("hermanos", "editar")
def hermano_update(hid):
    f = request.form
    try:
        q("""UPDATE hermanos SET nombres=%s,apellido_paterno=%s,apellido_materno=%s,correo=NULLIF(%s,''),
            telefono=%s,direccion=%s,comuna_id=NULLIF(%s,'')::int,sexo=NULLIF(%s,''),estado_civil=NULLIF(%s,''),
            invitado_por_id=NULLIF(%s,'')::int,invitado_por_texto=NULLIF(%s,'') WHERE id=%s""",
          (f["nombres"], f["paterno"], f.get("materno"), f.get("correo"), f.get("telefono"),
           f.get("direccion"), f.get("comuna") or None, f.get("sexo") or None, f.get("ecivil") or None,
           f.get("invitado") or None, f.get("invtexto") or None, hid), commit=True)
        flash("Actualizado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("hermanos"))

@app.post("/hermanos/<int:hid>/toggle")
@login_req
@permiso_req("hermanos", "editar")
def hermano_toggle(hid):
    q("UPDATE hermanos SET activo=NOT activo WHERE id=%s", (hid,), commit=True)
    return redirect(url_for("hermanos"))

# ---------- LIDERES ----------
@app.get("/lideres")
@login_req
def lideres():
    exp = request.args.get("export", "")
    f12 = request.args.get("f12", "")
    where = "WHERE (%s='' OR tipo_12=%s)"
    rows = q(f"SELECT lider, red, estado_celula, info_enviada, cantidad_celulas, discipulos_activos, es_pastor, tipo_12 FROM v_lideres_conteo {where} ORDER BY discipulos_activos DESC", (f12, f12 or None))
    if exp == "excel": return to_excel(rows, ["lider", "red", "estado_celula", "info_enviada", "cantidad_celulas", "discipulos_activos", "es_pastor", "tipo_12"], "lideres.xlsx")
    if exp == "pdf": return to_pdf("Lideres MCCI", rows, ["lider", "red", "tipo_12", "discipulos_activos"], "lideres.pdf")
    return render_template("lideres.html", rows=q(f"SELECT * FROM v_lideres_conteo {where} ORDER BY discipulos_activos DESC", (f12, f12 or None)), f12=f12)

@app.get("/lideres/nuevo")
@login_req
@permiso_req("lideres", "crear")
def lider_nuevo():
    libres = q("""SELECT h.id,h.nombre_completo FROM hermanos h LEFT JOIN lideres l ON l.hermano_id=h.id
                  WHERE l.id IS NULL AND h.activo ORDER BY h.nombre_completo""")
    return render_template("lider_form.html", libres=libres, l=None)

@app.post("/lideres/nuevo")
@login_req
@permiso_req("lideres", "crear")
def lider_crear():
    f = request.form
    try:
        q("INSERT INTO lideres(hermano_id,red,estado_celula,info_enviada,cantidad_celulas,observacion,es_pastor,tipo_12) VALUES(%s,%s,%s,%s,%s,%s,%s,%s)",
          (f["hermano"], f["red"], f["estado"], f["info"], int(f.get("celulas") or 1), f.get("obs"),
           f.get("es_pastor") == "on", f.get("tipo12", "NINGUNO")), commit=True)
        flash("Líder creado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("lideres"))

@app.post("/lideres/<int:lid>/estado")
@login_req
@permiso_req("lideres", "editar")
def lider_estado(lid):
    f = request.form
    q("UPDATE lideres SET estado_celula=%s, info_enviada=%s, cantidad_celulas=%s, red=%s, es_pastor=%s, tipo_12=%s WHERE id=%s",
      (f["estado"], f["info"], int(f.get("celulas") or 1), f["red"], f.get("pastor") == "on", f.get("tipo12", "NINGUNO"), lid), commit=True)
    flash("Líder actualizado", "ok"); return redirect(url_for("lideres"))

# ---------- DISCIPULADO ----------
@app.get("/discipulado")
@login_req
def discipulado():
    exp = request.args.get("export", "")
    texto = request.args.get("q", "")
    f12 = request.args.get("f12", "")
    base = """FROM discipulado d JOIN hermanos h ON h.id=d.hermano_id JOIN lideres l ON l.id=d.lider_id
        JOIN hermanos hl ON hl.id=l.hermano_id LEFT JOIN periodicidad_celula p ON p.id=d.periodicidad_id
        WHERE (%s='' OR l.tipo_12=%s) AND (h.nombre_completo ILIKE %s OR hl.nombre_completo ILIKE %s)"""
    params = (f12, f12 or None, f"%{texto}%", f"%{texto}%")
    rows = q(f"""SELECT h.nombre_completo AS hermano,hl.nombre_completo AS lider,p.nombre AS periodicidad,
        l.tipo_12,d.es_discipulo_activo,d.fecha_asignacion,d.fecha_fin {base}
        ORDER BY d.fecha_fin NULLS FIRST, d.fecha_asignacion DESC LIMIT 500""", params)
    if exp == "excel": return to_excel(rows, ["hermano", "lider", "periodicidad", "tipo_12", "es_discipulo_activo", "fecha_asignacion", "fecha_fin"], "discipulado.xlsx")
    if exp == "pdf": return to_pdf("Discipulado MCCI", rows, ["hermano", "lider", "tipo_12"], "discipulado.pdf")
    full = q(f"""SELECT d.id,h.nombre_completo AS hermano,hl.nombre_completo AS lider,p.nombre AS periodicidad,
        l.tipo_12,d.es_discipulo_activo,d.fecha_asignacion,d.fecha_fin {base}
        ORDER BY d.fecha_fin NULLS FIRST, d.fecha_asignacion DESC LIMIT 300""", params)
    return render_template("discipulado.html", rows=full, texto=texto, f12=f12)

@app.get("/discipulado/nuevo")
@login_req
@permiso_req("discipulado", "crear")
def disc_nuevo():
    return render_template("disc_form.html",
        hermanos=q("SELECT id,nombre_completo FROM hermanos WHERE activo ORDER BY 2"),
        lideres=q("SELECT l.id,h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY 2"),
        per=q("SELECT * FROM periodicidad_celula ORDER BY 1"))

@app.post("/discipulado/nuevo")
@login_req
@permiso_req("discipulado", "crear")
def disc_crear():
    f = request.form
    try:
        q("INSERT INTO discipulado(hermano_id,lider_id,es_discipulo_activo,periodicidad_id,observacion) VALUES(%s,%s,%s,%s,%s)",
          (f["hermano"], f["lider"], f.get("activo") == "on", f.get("per") or None, f.get("obs")), commit=True)
        flash("Asignación creada", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("discipulado"))

@app.post("/discipulado/<int:did>/cerrar")
@login_req
@permiso_req("discipulado", "editar")
def disc_cerrar(did):
    q("UPDATE discipulado SET fecha_fin=CURRENT_DATE WHERE id=%s", (did,), commit=True)
    return redirect(url_for("discipulado"))

# ---------- ENCUENTROS ----------
@app.get("/encuentros")
@login_req
def encuentros():
    exp = request.args.get("export", "")
    texto = request.args.get("q", "")
    fest = request.args.get("fest", "")
    base = "FROM encuentro_participacion e JOIN hermanos h ON h.id=e.hermano_id JOIN tipos_encuentro t ON t.id=e.tipo_encuentro_id WHERE (%s='' OR e.estado=%s) AND h.nombre_completo ILIKE %s"
    params = (fest, fest or None, f"%{texto}%")
    if exp in ("excel", "pdf"):
        rows = q(f"SELECT h.nombre_completo,t.nombre AS etapa,e.estado,e.fecha_evento {base} ORDER BY 1 LIMIT 500", params)
        if exp == "excel": return to_excel(rows, ["nombre_completo", "etapa", "estado", "fecha_evento"], "encuentros.xlsx")
        return to_pdf("Encuentros MCCI", rows, ["nombre_completo", "etapa", "estado"], "encuentros.pdf")
    rows = q(f"SELECT e.id,h.nombre_completo,t.nombre AS etapa,e.estado,e.fecha_evento {base} ORDER BY e.fecha_evento DESC LIMIT 300", params)
    ruta = q("SELECT * FROM v_ruta_espiritual WHERE nombre_completo ILIKE %s ORDER BY nombre_completo LIMIT 200", (f"%{texto}%",))
    return render_template("encuentros.html", rows=rows, ruta=ruta, texto=texto, fest=fest)

@app.get("/encuentros/nuevo")
@login_req
@permiso_req("encuentros", "crear")
def enc_nuevo():
    return render_template("enc_form.html",
        hermanos=q("SELECT id,nombre_completo FROM hermanos WHERE activo ORDER BY 2"),
        tipos=q("SELECT * FROM tipos_encuentro ORDER BY orden"))

@app.post("/encuentros/nuevo")
@login_req
@permiso_req("encuentros", "crear")
def enc_crear():
    f = request.form
    try:
        q("""INSERT INTO encuentro_participacion(hermano_id,tipo_encuentro_id,estado,fecha_evento) VALUES(%s,%s,%s,%s)
             ON CONFLICT (hermano_id,tipo_encuentro_id) DO UPDATE SET estado=EXCLUDED.estado,fecha_evento=EXCLUDED.fecha_evento""",
          (f["hermano"], f["tipo"], f["estado"], f.get("fecha") or None), commit=True)
        flash("Encuentro registrado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("encuentros"))

# ---------- SEGUIMIENTO DIARIO (sin célula) ----------
@app.get("/seguimiento")
@login_req
def seguimiento():
    exp = request.args.get("export", "")
    rows = q("SELECT * FROM v_sin_celula LIMIT 300")
    visitas = q("""SELECT v.*, (SELECT MAX(s.fecha) FROM seguimiento s WHERE s.visita_id=v.id) AS ultimo_seg
        FROM v_visitas_completo v WHERE v.estado IN ('PENDIENTE','EN SEGUIMIENTO') ORDER BY ultimo_seg NULLS FIRST LIMIT 100""")
    total = q("SELECT * FROM v_seguimiento_total LIMIT 500")
    lideres_sin = q("""SELECT h.nombre_completo,l.red FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.estado_celula='SIN CELULA' AND l.activo""")
    if exp == "excel": return to_excel(total, ["tipo", "nombre_completo", "telefono", "invitado_por", "ultimo_seg"], "seguimiento_total.xlsx")
    if exp == "pdf": return to_pdf("Seguimiento total", total, ["tipo", "nombre_completo", "telefono"], "seguimiento.pdf")
    hist = q("""SELECT s.id,COALESCE(h.nombre_completo,v.nombre_completo) AS nombre,s.fecha,s.tipo,s.comentario,s.responsable,s.contactado,
        CASE WHEN s.visita_id IS NOT NULL THEN 'VISITA' ELSE 'HERMANO' END AS origen FROM seguimiento s
        LEFT JOIN hermanos h ON h.id=s.hermano_id LEFT JOIN visitas v ON v.id=s.visita_id ORDER BY s.fecha DESC LIMIT 100""")
    lideres = q("""SELECT h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.activo ORDER BY h.nombre_completo""")
    return render_template("seguimiento.html", rows=rows, visitas=visitas, total=total, lideres_sin=lideres_sin, hist=hist, lideres=lideres)

@app.post("/seguimiento/<int:hid>/registrar")
@login_req
@permiso_req("seguimiento", "crear")
def seg_registrar(hid):
    f = request.form
    q("INSERT INTO seguimiento(hermano_id,tipo,comentario,responsable,contactado,proximo_contacto) VALUES(%s,%s,%s,%s,%s,NULLIF(%s,'')::date)",
      (hid, f.get("tipo", "LLAMADA"), f.get("comentario"), f.get("responsable"), f.get("contactado") == "on", f.get("proximo") or None), commit=True)
    flash("Seguimiento registrado", "ok"); return redirect(url_for("seguimiento"))

@app.post("/seguimiento/visita/<int:vid>/registrar")
@login_req
@permiso_req("seguimiento", "crear")
def seg_visita_registrar(vid):
    f = request.form
    q("INSERT INTO seguimiento(visita_id,tipo,comentario,responsable,contactado,proximo_contacto) VALUES(%s,%s,%s,%s,%s,NULLIF(%s,'')::date)",
      (vid, f.get("tipo", "LLAMADA"), f.get("comentario"), f.get("responsable"), f.get("contactado") == "on", f.get("proximo") or None), commit=True)
    q("UPDATE visitas SET estado='EN SEGUIMIENTO' WHERE id=%s AND estado='PENDIENTE'", (vid,), commit=True)
    flash("Seguimiento de visita registrado (pasó a EN SEGUIMIENTO)", "ok"); return redirect(url_for("seguimiento"))

# ---------- VISITAS (culto) ----------
@app.get("/visitas")
@login_req
def visitas():
    exp = request.args.get("export", "")
    texto = request.args.get("q", "")
    fest = request.args.get("fest", "")
    base = "FROM v_visitas_completo WHERE (%s='' OR estado=%s) AND nombre_completo ILIKE %s"
    params = (fest, fest or None, f"%{texto}%")
    rows = q(f"SELECT * {base} ORDER BY fecha_visita DESC LIMIT 300", params)
    if exp == "excel": return to_excel(rows, ["rut", "nombre_completo", "direccion", "telefono", "comuna", "invitado_por", "invitado_por_texto", "fecha_visita", "motivo_oracion", "estado", "responsable"], "visitas.xlsx")
    if exp == "pdf": return to_pdf("Visitas MCCI", rows, ["nombre_completo", "telefono", "motivo_oracion", "estado"], "visitas.pdf")
    lideres = q("""SELECT l.id, h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.activo ORDER BY h.nombre_completo""")
    return render_template("visitas.html", rows=rows, texto=texto, fest=fest, lideres=lideres)

@app.get("/visitas/nuevo")
@login_req
@permiso_req("visitas", "crear")
def visita_nueva():
    return render_template("visita_form.html", v=None,
        comunas=q("SELECT id,nombre FROM comunas ORDER BY nombre"),
        hermanos=q("SELECT id,nombre_completo FROM hermanos WHERE activo ORDER BY 2"),
        lideres=q("SELECT h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY 1"))

@app.post("/visitas/nuevo")
@login_req
@permiso_req("visitas", "crear")
def visita_crear():
    f = request.form
    q("""INSERT INTO visitas(rut,nombre_completo,direccion,telefono,correo,comuna_id,invitado_por_id,invitado_por_texto,
        fecha_visita,motivo_oracion,estado,responsable,observacion)
        VALUES(NULLIF(%s,''),%s,%s,%s,%s,NULLIF(%s,'')::int,NULLIF(%s,'')::int,NULLIF(%s,''),COALESCE(NULLIF(%s,'')::date,CURRENT_DATE),%s,%s,%s,%s)""",
      (f.get("rut") or None, f["nombre"].strip(), f.get("direccion"), f.get("telefono"), f.get("correo") or None, f.get("comuna") or None,
       f.get("invitado") or None, f.get("invtexto") or None, f.get("fecha") or None, f["motivo"].strip(),
       f.get("estado", "PENDIENTE"), f.get("responsable") or None, f.get("obs")), commit=True)
    flash("Visita registrada", "ok"); return redirect(url_for("visitas"))

@app.post("/visitas/<int:vid>/estado")
@login_req
@permiso_req("visitas", "editar")
def visita_estado(vid):
    f = request.form
    q("UPDATE visitas SET estado=%s, responsable=%s, observacion=%s WHERE id=%s",
      (f["estado"], f.get("responsable") or None, f.get("obs"), vid), commit=True)
    flash("Visita actualizada", "ok"); return redirect(url_for("visitas"))

@app.post("/visitas/<int:vid>/integrar")
@login_req
@permiso_req("visitas", "crear")
def visita_integrar(vid):
    """Integra visita: crea hermano + lo asigna a célula (discipulado)."""
    f = request.form
    rut = (f.get("rut") or "").strip()
    lider_id = f.get("lider")
    if not rut or not lider_id:
        flash("Para integrar indica RUT y líder/célula destino", "error"); return redirect(url_for("visitas"))
    v = q("SELECT * FROM visitas WHERE id=%s", (vid,), one=True)
    if not v: flash("Visita no existe", "error"); return redirect(url_for("visitas"))
    if v["estado"] == "INTEGRADO":
        flash("Ya estaba integrada", "error"); return redirect(url_for("visitas"))
    partes = (v["nombre_completo"] or "").split()
    nombres = " ".join(partes[:2]) if len(partes) > 2 else (partes[0] if partes else "S/N")
    paterno = partes[-1] if len(partes) > 1 else "S/A"
    try:
        with db() as c, c.cursor() as cur:
            cur.execute("""INSERT INTO hermanos(rut,nombres,apellido_paterno,telefono,direccion,correo,comuna_id,invitado_por_id,invitado_por_texto)
                VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (rut, nombres, paterno, v["telefono"], v["direccion"], v["correo"], v["comuna_id"], v["invitado_por_id"], v["invitado_por_texto"]))
            hid = cur.fetchone()["id"]
            cur.execute("INSERT INTO discipulado(hermano_id,lider_id,es_discipulo_activo,observacion) VALUES(%s,%s,TRUE,'Integrado desde visita')", (hid, lider_id))
            cur.execute("UPDATE visitas SET estado='INTEGRADO', hermano_id=%s WHERE id=%s", (hid, vid))
            c.commit()
        flash(f"Integrado: ahora es hermano + asignado a célula (id {hid})", "ok")
    except Exception as e:
        flash(f"Error al integrar (¿RUT duplicado?): {e}", "error")
    return redirect(url_for("visitas"))

# ---------- INFORME CÉLULA ----------
@app.get("/celulas-informe")
@login_req
@permiso_req("celulas", "leer")
def celulas_informe_lista():
    exp = request.args.get("export", "")
    rows = q("""SELECT ci.*, h1.nombre_completo AS lider_nombre, h2.nombre_completo AS pastor_nombre
        FROM celulas_informe ci 
        LEFT JOIN lideres l1 ON l1.id=ci.lider_id LEFT JOIN hermanos h1 ON h1.id=l1.hermano_id
        LEFT JOIN lideres l2 ON l2.id=ci.pastor_id LEFT JOIN hermanos h2 ON h2.id=l2.hermano_id
        ORDER BY ci.id DESC""")
    if exp == "excel":
        return to_excel(rows, ["lider_nombre","pastor_nombre","fecha","mes","red","comuna",
            "horario","realizado","discipulos_asistieron","ofrenda","tipo_ofrenda"], "informe_celulas.xlsx")
    if exp == "pdf":
        return to_pdf("Informe Células MCCI", rows, ["lider_nombre","pastor_nombre","fecha","mes",
            "red","realizado","ofrenda"], "informe_celulas.pdf")
    return render_template("celulas_informe.html", rows=rows)

@app.get("/celulas-informe/nuevo")
@login_req
@permiso_req("celulas", "crear")
def celulas_informe_nuevo():
    lideres = q("""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY h.nombre_completo""")
    comunas = q("SELECT id, nombre FROM comunas ORDER BY nombre")
    return render_template("celulas_informe_form.html", c=None, lideres=lideres, comunas=comunas)

@app.post("/celulas-informe/nuevo")
@login_req
@permiso_req("celulas", "crear")
def celulas_informe_crear():
    f = request.form
    try:
        q("""INSERT INTO celulas_informe(lider_id,pastor_id,fecha,mes,red,direccion,comuna,
            horario,realizado,justificacion,discipulos_asistieron,discipulos_no_asistieron,
            ofrenda,tipo_ofrenda,modalidad)
            VALUES(NULLIF(%s,'')::int,NULLIF(%s,'')::int,NULLIF(%s,'')::date,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,NULLIF(%s,'')::decimal,%s,%s)""",
          (f.get("lider_id"), f.get("pastor_id"), f.get("fecha"), f.get("mes"), f.get("red"),
           f.get("direccion"), f.get("comuna"), f.get("horario"),
           f.get("realizo") == "on", f.get("justificacion"), f.get("asistieron"),
           f.get("no_asistieron"), f.get("ofrenda") or "0", f.get("tipo_ofrenda"),
           f.get("modalidad")), commit=True)
        flash("Informe registrado", "ok")
    except Exception as e:
        flash(f"Error: {e}", "error")
    return redirect(url_for("celulas_informe_lista"))

@app.get("/celulas-informe/<int:cid>/editar")
@login_req
@permiso_req("celulas", "editar")
def celulas_informe_editar(cid):
    c = q("SELECT * FROM celulas_informe WHERE id=%s", (cid,), one=True)
    if not c: abort(404)
    lideres = q("""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY h.nombre_completo""")
    comunas = q("SELECT id, nombre FROM comunas ORDER BY nombre")
    return render_template("celulas_informe_form.html", c=c, lideres=lideres, comunas=comunas)

@app.post("/celulas-informe/<int:cid>/editar")
@login_req
@permiso_req("celulas", "editar")
def celulas_informe_update(cid):
    f = request.form
    try:
        q("""UPDATE celulas_informe SET lider_id=NULLIF(%s,'')::int,pastor_id=NULLIF(%s,'')::int,
            fecha=NULLIF(%s,'')::date,mes=%s,red=%s,direccion=%s,comuna=%s,horario=%s,
            realizo=%s,justificacion=%s,discipulos_asistieron=%s,discipulos_no_asistieron=%s,
            ofrenda=NULLIF(%s,'')::decimal,tipo_ofrenda=%s,modalidad=%s WHERE id=%s""",
          (f.get("lider_id"), f.get("pastor_id"), f.get("fecha"), f.get("mes"), f.get("red"),
           f.get("direccion"), f.get("comuna"), f.get("horario"),
           f.get("realizo") == "on", f.get("justificacion"), f.get("asistieron"),
           f.get("no_asistieron"), f.get("ofrenda") or "0", f.get("tipo_ofrenda"),
           f.get("modalidad"), cid), commit=True)
        flash("Informe actualizado", "ok")
    except Exception as e:
        flash(f"Error: {e}", "error")
    return redirect(url_for("celulas_informe_lista"))

@app.post("/celulas-informe/<int:cid>/eliminar")
@login_req
@permiso_req("celulas", "eliminar")
def celulas_informe_eliminar(cid):
    try:
        q("DELETE FROM celulas_informe WHERE id=%s", (cid,), commit=True)
        flash("Informe eliminado", "ok")
    except Exception as e:
        flash(f"Error: {e}", "error")
    return redirect(url_for("celulas_informe_lista"))

# ---------- REPORTES ----------
@app.get("/reportes")
@login_req
def reportes():
    exp = request.args.get("export", "")
    ver = request.args.get("ver", "todos")
    r = {}
    r["por_red"] = q("SELECT red AS red,estado_celula,COUNT(*) AS c FROM lideres WHERE activo GROUP BY 1,2 ORDER BY 1,2")
    r["por_comuna"] = q("""SELECT r.nombre_corto AS region,c.nombre AS comuna,COUNT(*) AS c FROM hermanos h
        LEFT JOIN comunas c ON c.id=h.comuna_id LEFT JOIN regiones r ON r.id=c.region_id WHERE h.activo GROUP BY 1,2 ORDER BY 3 DESC LIMIT 30""")
    r["sin_celula"] = q("""SELECT h.nombre_completo,l.red AS red FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.estado_celula='SIN CELULA' AND l.activo""")
    r["pendientes"] = q("""SELECT h.nombre_completo,t.nombre AS etapa,e.fecha_evento AS fecha FROM encuentro_participacion e
        JOIN hermanos h ON h.id=e.hermano_id JOIN tipos_encuentro t ON t.id=e.tipo_encuentro_id
        WHERE e.estado='PENDIENTE' ORDER BY 3 LIMIT 50""")
    r["ruta_incompleta"] = q("SELECT nombre_completo FROM v_ruta_espiritual WHERE escuela_lideres=0 ORDER BY 1 LIMIT 100")
    r["periodicidad"] = q("""SELECT p.nombre AS nombre,COUNT(d.id) AS c FROM discipulado d JOIN periodicidad_celula p ON p.id=d.periodicidad_id
        WHERE d.fecha_fin IS NULL GROUP BY 1 ORDER BY 2 DESC""")
    r["por_12"] = q("SELECT tipo_12, COUNT(*) AS c, SUM(CASE WHEN es_pastor THEN 1 ELSE 0 END) AS pastores FROM lideres WHERE activo GROUP BY 1 ORDER BY 2 DESC")
    # Resumen único por líder + fila TOTAL (con encuentros desglosados por etapa)
    r["resumen"] = q("""SELECT v.lider, v.red, v.tipo_12, v.es_pastor, v.estado_celula, v.cantidad_celulas,
        v.total_asignados_vigentes, v.discipulos_activos,
        COALESCE(e.aprob_lider,0) AS enc_lider,
        COALESCE(s.tremendo,0) AS d_tremendo, COALESCE(s.fruto,0) AS d_fruto,
        COALESCE(s.reenc,0) AS d_reenc, COALESCE(s.crec,0) AS d_crec, COALESCE(s.lid,0) AS d_lid,
        COALESCE(s.total_aprob,0) AS disc_enc_aprob
        FROM v_lideres_conteo v
        LEFT JOIN (SELECT l.id AS lider_id, COUNT(ep.id) AS aprob_lider FROM lideres l
            LEFT JOIN encuentro_participacion ep ON ep.hermano_id=l.hermano_id AND ep.estado='APROBADO'
            GROUP BY 1) e ON e.lider_id=v.lider_id
        LEFT JOIN (SELECT d.lider_id,
            COUNT(*) FILTER (WHERE t.nombre='ENCUENTRO TREMENDO') AS tremendo,
            COUNT(*) FILTER (WHERE t.nombre='FRUTO FIEL') AS fruto,
            COUNT(*) FILTER (WHERE t.nombre='REENCUENTRO') AS reenc,
            COUNT(*) FILTER (WHERE t.nombre='ESCUELA CRECIMIENTO') AS crec,
            COUNT(*) FILTER (WHERE t.nombre='ESCUELA LIDERES') AS lid,
            COUNT(*) AS total_aprob
            FROM discipulado d JOIN encuentro_participacion ep ON ep.hermano_id=d.hermano_id AND ep.estado='APROBADO'
            JOIN tipos_encuentro t ON t.id=ep.tipo_encuentro_id
            WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo GROUP BY 1) s ON s.lider_id=v.lider_id
        ORDER BY v.discipulos_activos DESC""")
    tot = q("SELECT COUNT(*) AS lideres, COALESCE(SUM(discipulos_activos),0) AS activos, COALESCE(SUM(total_asignados_vigentes),0) AS asignados FROM v_lideres_conteo", one=True)
    r["total"] = tot
    # Totales de encuentros para la fila TOTAL
    te = q("""SELECT COALESCE(SUM(CASE WHEN ep.estado='APROBADO' AND ep.hermano_id=l.hermano_id THEN 1 ELSE 0 END),0) AS enc_lideres,
        COALESCE(SUM(CASE WHEN ep.estado='APROBADO' THEN 1 ELSE 0 END),0) AS enc_disc
        FROM discipulado d JOIN lideres l ON l.id=d.lider_id
        LEFT JOIN encuentro_participacion ep ON ep.hermano_id=d.hermano_id
        WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo""", one=True)
    r["total_enc"] = te
    catalog = {"por_red": ("Reporte por red", ["red", "estado_celula", "c"]), "por_comuna": ("Reporte por comuna", ["region", "comuna", "c"]),
        "sin_celula": ("Sin celula", ["nombre_completo", "red"]), "pendientes": ("Pendientes", ["nombre_completo", "etapa", "fecha"]),
        "ruta_incompleta": ("Ruta incompleta", ["nombre_completo"]), "periodicidad": ("Periodicidad", ["nombre", "c"]),
        "por_12": ("Por Tipo 12", ["tipo_12", "c", "pastores"]),
        "resumen": ("Resumen por lider (unico + total)", ["lider", "red", "tipo_12", "es_pastor", "estado_celula", "discipulos_activos", "enc_lider", "d_tremendo", "d_fruto", "d_reenc", "d_crec", "d_lid", "disc_enc_aprob"])}
    if exp in ("excel", "pdf"):
        if ver == "todos":
            all_rows = []
            all_heads = ["seccion", "campo1", "campo2", "campo3"]
            for key, (title, heads) in catalog.items():
                for row in r.get(key, []):
                    line = {"seccion": title}
                    for i, h in enumerate(heads[:3]):
                        line[f"campo{i+1}"] = str(row.get(h, ""))
                    all_rows.append(line)
            if not all_rows:
                all_rows = [{"seccion": "Sin datos", "campo1": "", "campo2": "", "campo3": ""}]
            if exp == "excel": return to_excel(all_rows, all_heads, "reporte_todos.xlsx")
            return to_pdf("Reportes MCCI - Todos", all_rows, all_heads, "reporte_todos.pdf")
        if ver in catalog:
            title, heads = catalog[ver]
            rows = r[ver]
            if not rows:
                rows = [{h: "" for h in heads}]
            if exp == "excel": return to_excel(rows, heads, f"reporte_{ver}.xlsx")
            return to_pdf(title, rows, heads, f"reporte_{ver}.pdf")
    return render_template("reportes.html", r=r, ver=ver)

# ---------- APIs (compatibilidad) ----------
@app.post("/api/login")
def api_login():
    data = request.get_json(force=True)
    u = q("SELECT id,password_hash,rol FROM usuarios WHERE username=%s AND activo", (data.get("user"),), one=True)
    if not u or not check_password_hash(u["password_hash"], data.get("password", "")): abort(401)
    session["uid"], session["rol"] = u["id"], u["rol"]
    return {"ok": True, "rol": u["rol"]}

@app.get("/api/lideres")
def api_lideres():
    if "uid" not in session: abort(401)
    return jsonify(q("SELECT * FROM v_lideres_conteo ORDER BY discipulos_activos DESC"))

@app.get("/api/hermanos")
def api_hermanos():
    if "uid" not in session: abort(401)
    texto = f"%{request.args.get('q','')}%"
    return jsonify(q("SELECT * FROM v_hermanos_completo WHERE nombre_completo ILIKE %s LIMIT 100", (texto,)))

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
