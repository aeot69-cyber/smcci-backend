"""SMCCI - Sistema Consolidación MCCI (Flask + Postgres). CRUD + Reportes con menú."""
import os
import time
import random
import secrets
from functools import wraps
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv
from flask import Flask, request, jsonify, session, redirect, url_for, render_template, flash, abort, send_file, make_response
from werkzeug.security import check_password_hash
from io import BytesIO
from openpyxl import Workbook
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()
app = Flask(__name__)
# Debug mode en desarrollo local
if os.environ.get("FLASK_ENV") == "development":
    app.debug = True

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
# La sesión de Flask es un cookie firmado con esta clave: con un valor por
# defecto fijo y público (como tenía antes) cualquiera podría fabricar una
# sesión válida, p. ej. SUPERADMIN, y tomar el sistema. Si falta la variable se
# genera una clave aleatoria: la sesión se cierra en cada reinicio, pero nadie
# puede falsificarla. Configure SECRET_KEY en el entorno para evitar los
# reinicios de sesión en cada despliegue.
_SECRET_KEY = os.environ.get("SECRET_KEY", "").strip()
if not _SECRET_KEY:
    _SECRET_KEY = secrets.token_hex(32)
    print("AVISO DE SEGURIDAD: falta SECRET_KEY en el entorno; se usó una clave "
          "efímera (las sesiones se cierran al reiniciar). Configure SECRET_KEY.")
app.secret_key = _SECRET_KEY
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("FLASK_ENV") == "production", PERMANENT_SESSION_LIFETIME=3600)

# ---------- DOMINIO OFICIAL ----------
# Que todos entren por smcci.iglesiamcci.cl: cualquier otro host (p. ej. el
# dominio de Render) redirige aquí en GET/HEAD conservando ruta y consulta.
# /api/* no se toca (consumidores programáticos) y localhost/127.0.0.1 quedan
# exentos para desarrollo local. El host se puede cambiar con DOMINIO_OFICIAL.
DOMINIO_OFICIAL = os.environ.get("DOMINIO_OFICIAL", "smcci.iglesiamcci.cl").strip().lower()

@app.before_request
def _redirigir_dominio_oficial():
    if request.method not in ("GET", "HEAD") or request.path.startswith("/api/"):
        return None
    host = (request.host or "").split(":")[0].strip().lower()
    if host and host not in (DOMINIO_OFICIAL, "localhost", "127.0.0.1"):
        destino = "https://" + DOMINIO_OFICIAL + (request.full_path if request.query_string else request.path)
        return redirect(destino, code=308)
    return None

def db():
    # Reusa una sola conexión por request (evita 23 handshakes en /configuracion → evita timeout en Render)
    try:
        if hasattr(g, "_db_conn") and g._db_conn and not g._db_conn.closed:
            return g._db_conn
    except RuntimeError:
        pass  # fuera de contexto de app, crea conexión directa
    conn = None
    last_err = None
    # 1) DATABASE_URL
    db_url = os.environ.get("DATABASE_URL", "")
    if db_url:
        try:
            from urllib.parse import urlparse, parse_qs
            u = urlparse(db_url)
            # Supabase exige SSL: pasar ?sslmode=require en el DATABASE_URL.
            # Por defecto 'prefer' (comportamiento anterior, ej. Render/local).
            sslmode = parse_qs(u.query).get("sslmode", ["prefer"])[0]
            conn = psycopg2.connect(host=u.hostname, port=u.port or 5432, dbname=u.path.lstrip("/"), user=u.username, password=u.password, sslmode=sslmode, cursor_factory=RealDictCursor, connect_timeout=5)
            try: g._db_conn = conn
            except RuntimeError: pass
            return conn
        except Exception as e:
            last_err = e
            print(f"DATABASE_URL connect failed: {e}")
    # 2) DB_* vars
    host = os.environ.get("DB_HOST")
    pwd = os.environ.get("DB_PASSWORD")
    if host and pwd:
        try:
            conn = psycopg2.connect(host=host, port=os.environ.get("DB_PORT", "5432"), dbname=os.environ.get("DB_NAME", "Iglesia_MCCI"), user=os.environ.get("DB_USER"), password=pwd, cursor_factory=RealDictCursor, connect_timeout=5)
            try: g._db_conn = conn
            except RuntimeError: pass
            return conn
        except Exception as e:
            last_err = e
            print(f"DB_* connect failed: {e}")
    # 3) Sin entorno configurado: NO hay fallback con credenciales guardadas en
    #    el código. Antes había un tercer intento contra una base vieja con la
    #    contraseña escrita en el fuente: eso compromete los datos de la iglesia
    #    para cualquiera que vea el repositorio. Mejor fallar con claridad.
    raise RuntimeError(
        "No hay base de datos configurada. Defina DATABASE_URL (o DB_HOST/DB_USER/"
        "DB_PASSWORD/DB_NAME/DB_PORT) en el entorno del servidor."
        + (f" Último error de conexión: {last_err}" if last_err else "")
    )

@app.teardown_appcontext
def close_db(exc):
    conn = getattr(g, "_db_conn", None)
    if conn is not None:
        try:
            if not conn.closed:
                conn.close()
        except: pass
        g._db_conn = None

def q(sql, p=(), one=False, commit=False):
    c = db()
    with c.cursor() as cur:
        cur.execute(sql, p)
        if commit:
            c.commit()
            return cur.rowcount
        return cur.fetchone() if one else cur.fetchall()

def update_parcial(tabla, id_val, campos_valores, id_col="id", permitir_null=None):
    """Actualiza solo los campos no vacíos. Si un campo es None o '', se conserva el valor anterior.
    Usa permitir_null=[\"col1\",\"col2\"] para permitir poner NULL explícitamente."""
    permitir_null = permitir_null or []
    # Sentinel para forzar NULL
    NULL_SENTINEL = "__NULL__"
    sets, vals = [], []
    for col, val in campos_valores.items():
        if val == NULL_SENTINEL:
            sets.append(f"{col}=NULL")
        elif col in permitir_null and val is None:
            sets.append(f"{col}=NULL")
        elif val is not None and val != "":
            sets.append(f"{col}=%s")
            vals.append(val)
    if not sets:
        return 0
    vals.append(id_val)
    return q(f"UPDATE {tabla} SET {','.join(sets)} WHERE {id_col}=%s", tuple(vals), commit=True)


def get_catalogo(tabla, incluir_inactivos=False, iglesia="sesion"):
    """Carga valores de un catálogo dinámico. Retorna lista de dicts con nombre.
    Por iglesia: globales (iglesia_id NULL) + propios. SUPERADMIN en TODAS: solo globales.
    iglesia: 'sesion' = iglesia en sesión; None = solo globales; int = esa iglesia (para
    el informe express público, que no tiene sesión de admin)."""
    try:
        where = "" if incluir_inactivos else "WHERE activo=true"
        order = "orden, nombre" if tabla in ("catalogo_meses","tipos_encuentro") else "nombre"
        if tabla == "catalogo_roles_usuario":
            # Los nombres de rol son globales a proposito
            return q(f"SELECT nombre FROM {tabla} {where} ORDER BY nombre")
        if iglesia == "sesion":
            igl, _ = iglesia_actual()
        else:
            igl = iglesia
        cond, params = "", []
        if igl is not None:
            cond = (" AND " if where else "WHERE ") + "(iglesia_id IS NULL OR iglesia_id=%s)"
            params = [igl]
        # catalogo_sexo tiene codigo, los demás solo nombre
        if tabla == "catalogo_sexo":
            rows = q(f"SELECT id, codigo, nombre FROM {tabla} {where}{cond} ORDER BY {order}", tuple(params))
        elif tabla == "catalogo_roles_lider":
            rows = q(f"SELECT nombre, icono FROM {tabla} {where}{cond} ORDER BY {order}", tuple(params))
            return rows
        else:
            rows = q(f"SELECT nombre FROM {tabla} {where}{cond} ORDER BY {order}", tuple(params))
            # Normalizar a dict con .nombre para compatibilidad
        return rows
    except:
        return []


# Mapeo catálogo → (tabla_afectada, columna, valor_por_defecto) para fallback al eliminar
CATALOGO_FALLBACK = {
    "catalogo_redes":             [("lideres", "red", "MIXTO")],
    "catalogo_tipo12":            [("lideres", "tipo_12", "NINGUNO")],
    "catalogo_roles_lider":       [("lideres", "rol_lider", "NINGUNO")],
    "catalogo_estado_civil":      [("hermanos", "estado_civil", None)],
    "catalogo_sexo":              [("hermanos", "sexo", None)],
    "catalogo_estado_celula":     [("lideres", "estado_celula", "EN FORMACION")],
    "catalogo_info_enviada":      [("lideres", "info_enviada", "EN CONSULTA")],
    "catalogo_estado_visita":     [("visitas", "estado", "PENDIENTE")],
    "catalogo_tipo_seguimiento":  [("seguimiento", "tipo", "LLAMADA")],
    "catalogo_tipo_hermano":      [("hermanos", "tipo_hermano", "NINGUNO")],
    "catalogo_modalidad_celula":  [("celulas_informe", "modalidad", "PRESENCIAL")],
    "catalogo_roles_usuario":     [("usuarios", "rol", "CONSULTA")],
}

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

from flask import g

def permisos(rol, iglesia="sesion"):
    """Permisos de un rol. iglesia: None=solo global, int=iglesia (con fallback a global),
    'sesion'=la iglesia en sesion (SUPERADMIN siempre global)."""
    if "_perm_cache" not in g:
        g._perm_cache = {}
    if iglesia == "sesion":
        try:
            if rol == "SUPERADMIN":
                iglesia = None
            else:
                iglesia = session.get("iglesia_id") or 1
        except RuntimeError:
            iglesia = None
    key = (rol, iglesia)
    if key not in g._perm_cache:
        if iglesia is None:
            rows = q("SELECT modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar FROM roles_permisos WHERE rol=%s AND iglesia_id IS NULL", (rol,))
        else:
            rows = q("SELECT modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar, iglesia_id FROM roles_permisos WHERE rol=%s AND (iglesia_id IS NULL OR iglesia_id=%s)", (rol, iglesia))
        d = {}
        for r in rows:
            # el valor propio de la iglesia pisa al global (herencia con override)
            if r["modulo"] not in d or r.get("iglesia_id") is not None:
                d[r["modulo"]] = r
        g._perm_cache[key] = d
    return g._perm_cache[key]

def _permisos_usuario(usuario_id):
    if "_uperm_cache" not in g:
        g._uperm_cache = {}
    if usuario_id not in g._uperm_cache:
        rows = q("SELECT modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar FROM usuario_permisos WHERE usuario_id=%s", (usuario_id,))
        g._uperm_cache[usuario_id] = {r["modulo"]: r for r in rows}
    return g._uperm_cache[usuario_id]

def tiene_permiso(rol, modulo, accion="leer", usuario_id=None, iglesia_id="sesion"):
    if rol == "SUPERADMIN":
        return True
    if usuario_id:
        up = _permisos_usuario(usuario_id).get(modulo)
        if up:
            return bool(up.get(f"puede_{accion}", False))
    p = permisos(rol, iglesia_id).get(modulo)
    if not p: return False
    return bool(p.get(f"puede_{accion}", False))

# ---------- MULTI-IGLESIA ----------
def iglesia_actual():
    """(id, nombre) de la iglesia en sesion. SUPERADMIN con ver_todo -> (None, 'TODAS')."""
    if session.get("rol") == "SUPERADMIN" and session.get("ver_todo"):
        return None, "TODAS"
    return session.get("iglesia_id") or 1, session.get("iglesia_nombre") or "CASA MATRIZ"

def filtro_igl(alias=""):
    """Fragmento ' AND [alias.]iglesia_id=%s' + params. Vacio si se ve todo."""
    igl, _ = iglesia_actual()
    if igl is None: return "", []
    col = f"{alias}.iglesia_id" if alias else "iglesia_id"
    return f" AND {col}=%s", [igl]

# Tablas cuya iglesia se valida por join (no tienen iglesia_id propia)
_IGL_VIA = {
    "discipulado": "SELECT l.iglesia_id FROM discipulado d JOIN lideres l ON l.id=d.lider_id WHERE d.id=%s",
    "encuentro_participacion": "SELECT h.iglesia_id FROM encuentro_participacion e JOIN hermanos h ON h.id=e.hermano_id WHERE e.id=%s",
}
def exigir_igl(tabla, rid):
    """True si el registro es de la iglesia en sesion (o se ve todo). Tablas con iglesia_id propia o via join."""
    igl, _ = iglesia_actual()
    if igl is None: return True
    try: rid = int(rid)
    except: return False
    if tabla in _IGL_VIA:
        r = q(_IGL_VIA[tabla], (rid,), one=True)
    else:
        # solo tablas conocidas (llamadas internas fijas, no input de usuario)
        if tabla not in ("hermanos","lideres","usuarios","visitas","seguimiento","celulas_informe","auditoria","iglesias"):
            return False
        r = q(f"SELECT iglesia_id FROM {tabla} WHERE id=%s", (rid,), one=True)
    return bool(r) and r["iglesia_id"] == igl

def iglesia_de(tabla, rid, col_id="id"):
    """iglesia_id de un registro (hermanos/lideres/visitas) o None."""
    try: rid = int(rid)
    except: return None
    if tabla not in ("hermanos","lideres","visitas"):
        return None
    r = q(f"SELECT iglesia_id FROM {tabla} WHERE {col_id}=%s", (rid,), one=True)
    return r["iglesia_id"] if r else None

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

# =====================================================================
#  AYUDAS / HELP EDITABLES  (solo SUPERADMIN las cambia en Configuración)
#  · Los textos por defecto viven en las plantillas (se escanean 1 vez).
#  · La tabla `ayudas` guarda SOLO los cambios hechos desde Configuración;
#    si no hay cambio, se usa el texto original de la plantilla.
#  · Formato de edición: texto plano.  **negrita**  ·  \n = salto
#    Los bloques ❓Ayuda son una viñeta por línea  "Etiqueta: texto".
# =====================================================================
from markupsafe import Markup, escape as _esc
from ayudas_defecto import AYUDAS_DEFECTO, AYUDAS_DINAMICO

_ay_cat = None   # catálogo de ayudas (una sola vez por proceso)

def _ay_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")

def escanear_ayudas():
    """Catálogo de ayudas: clave -> {plantilla,tipo,defecto,viñetas,dinamico}.
    Los textos por defecto vienen de ayudas_defecto.py (snapshot generado desde
    las plantillas originales); los cambios hechos en ⚙️ Configuración viven en
    la tabla `ayudas` y tienen prioridad."""
    global _ay_cat
    if _ay_cat is not None:
        return _ay_cat
    cat = {}
    for clave, defecto in AYUDAS_DEFECTO.items():
        base, _, resto = clave.partition(".")
        es_bloque = resto.startswith("ayuda")
        cat[clave] = {
            "plantilla": base + ".html",
            "tipo": "bloque" if es_bloque else "campo",
            "defecto": defecto,
            "viñetas": len([l for l in defecto.split("\n") if l.strip()]) if es_bloque else 0,
            "dinamico": AYUDAS_DINAMICO.get(clave, ""),
        }
    _ay_cat = cat
    return cat

def _ay_ovrs():
    """Cambios guardados en BD (se leen por request para verlos al instante en todas las instancias)."""
    from flask import g
    if not hasattr(g, "_ayudas_ovr"):
        try:
            g._ayudas_ovr = {r["clave"]: r["texto"] for r in q("SELECT clave, texto FROM ayudas")}
        except Exception:
            g._ayudas_ovr = {}
    return g._ayudas_ovr

def inval_ayudas():
    """Fuerza releer los cambios (se llama al guardar)."""
    from flask import g
    g._ayudas_ovr = None

def _ay_texto(clave):
    o = _ay_ovrs()
    if clave in o:
        return o[clave]
    return escanear_ayudas().get(clave, {}).get("defecto", "")

def _render_ay(txt, **ctx):
    s = str(_esc(txt or ""))                       # escapa HTML (el texto es de SUPERADMIN)
    for k, v in ctx.items():                       # tokens {semana}, {timoteos}, ...
        s = s.replace("{%s}" % k, str(_esc(str(v))))
    s = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', s, flags=re.S)   # **negrita**
    return Markup(s.replace("\n", "<br>\n"))       # salto de línea

def ay(clave, **ctx):
    """Ayuda de un campo (form-text)."""
    return _render_ay(_ay_texto(clave), **ctx)

def ayuda_li(clave, **ctx):
    """Bloque ❓Ayuda: cada línea del texto es una viñeta 'Etiqueta: texto'."""
    out = []
    for ln in _ay_texto(clave).split("\n"):
        ln = ln.strip()
        if not ln:
            continue
        m = re.match(r'^([^:*(){}\[\]→]{1,40}?):\s+(.+)$', ln)
        if m:
            out.append('<li><b>%s:</b> %s</li>' % (str(_esc(m.group(1))), _render_ay(m.group(2), **ctx)))
        else:
            out.append("<li>%s</li>" % _render_ay(ln, **ctx))
    return Markup("".join(out))

def guardar_ayudas(f, accion):
    """Guarda o restaura las ayudas editadas en ⚙️ Configuración (solo SUPERADMIN).
    Texto igual al original -> se borra el cambio y queda el valor por defecto."""
    if session.get("rol") != "SUPERADMIN":
        flash("Solo SUPERADMIN puede modificar las ayudas del sistema.", "error")
        return redirect(url_for("configuracion"))
    pagina = (f.get("pagina") or "").strip()
    pref = pagina[:-5] if pagina.endswith(".html") else pagina
    cat = escanear_ayudas()
    claves = [k for k in cat if k.split(".")[0] == pref]
    if not claves:
        flash("Página de ayuda no reconocida.", "error")
        return redirect(url_for("configuracion") + "#ayudas")
    if accion == "restaurar":
        try:
            n = q("DELETE FROM ayudas WHERE clave LIKE %s", (pref + ".%",), commit=True)
            flash(f"Ayudas de «{pagina}» restauradas a sus valores originales ({n}).", "success")
        except Exception as e:
            flash(f"No se pudieron restaurar las ayudas: {e}", "error")
    else:
        guard, rest = 0, 0
        for k in claves:
            if ("t__" + k) not in f:
                continue
            txt = (f.get("t__" + k) or "").replace("\r\n", "\n").replace("\r", "\n").strip("\n")
            try:
                if txt.strip() == cat[k]["defecto"].strip():
                    q("DELETE FROM ayudas WHERE clave=%s", (k,), commit=True)
                    rest += 1
                else:
                    q("""INSERT INTO ayudas (clave, texto, actualizado_en, actualizado_por)
                         VALUES (%s, %s, now(), %s)
                         ON CONFLICT (clave) DO UPDATE SET texto = EXCLUDED.texto,
                             actualizado_en = now(), actualizado_por = EXCLUDED.actualizado_por""",
                      (k, txt, session.get("username") or session.get("user") or ""), commit=True)
                    guard += 1
            except Exception as e:
                flash(f"Error al guardar «{k}»: {e}", "error")
        flash(f"Ayudas guardadas: {guard} modificada(s) · {rest} en su valor original.", "success")
    try:
        inval_ayudas()
    except Exception:
        pass
    return redirect(url_for("configuracion") + "#ayudas")

app.jinja_env.globals["ay"] = ay
app.jinja_env.globals["ayuda_li"] = ayuda_li

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
        ("configuracion","Configuración","⚙️","configuracion","/configuracion",None),
        ("auditoria",   "Auditoría",   "📜", "auditoria",   "/auditoria",   None),
    ]
    menu = "<a href='/'>🏠 INICIO</a>"
    for mod, nombre, icono, key, url_list, url_nuevo in modulos_menu:
        pk = p.get(key, {})
        if rol != "SUPERADMIN" and not pk.get("puede_leer"): continue
        menu += f"<button class='mbtn' data-bs-toggle='collapse' data-bs-target='#m-{key}'>{icono} {nombre} ▾</button>"
        menu += f"<div id='m-{key}' class='collapse sub'>"
        menu += f"<a href='{url_list}'>🔍 Consultar</a>"
        if pk.get("puede_crear") and url_nuevo:
            menu += f"<a href='{url_nuevo}'>➕ Ingresar</a>"
        if pk.get("puede_exportar"):
            menu += f"<a href='{url_list}?export=excel' target='_blank'>📊 Excel</a>"
            menu += f"<a href='{url_list}?export=pdf' target='_blank'>📄 PDF</a>"
        menu += "</div>"
        if key == "reportes":
            menu += "<a href='/graficos'>📊 Gráficos</a>"
    # Iglesias: solo SUPERADMIN (configuracion global)
    iglesias_lst, vista_igl = [], "CASA MATRIZ"
    if session.get("uid"):
        if rol == "SUPERADMIN" and session.get("ver_todo"):
            vista_igl = "TODAS"
        else:
            vista_igl = session.get("iglesia_nombre", "CASA MATRIZ")
        if rol == "SUPERADMIN":
            menu += "<a href='/iglesias'>⛪ Iglesias</a>"
            try: iglesias_lst = q("SELECT id,nombre,tipo FROM iglesias WHERE activo ORDER BY id")
            except: iglesias_lst = []
    return dict(perm=p, rol=rol, menu=menu, tiene_permiso=lambda m, a="leer": tiene_permiso(rol, m, a, session.get("uid")),
        vista_igl=vista_igl, iglesias_lst=iglesias_lst)

def log_audit(accion, modulo, detalle=""):
    """Registra auditoría sin bloquear si falla."""
    try:
        uid = session.get("uid")
        user = session.get("user", "")
        rol = session.get("rol", "")
        # IP real: detrás de proxy (Render/Koyeb) remote_addr es la IP interna
        # del proxy; la del cliente viene en X-Forwarded-For (primera = origen).
        ip = (request.headers.get("X-Forwarded-For", "") or "").split(",")[0].strip() or (request.remote_addr or "")
        igl = session.get("iglesia_id") or 1
        q("INSERT INTO auditoria(usuario_id, username, rol, accion, modulo, detalle, ip, iglesia_id) VALUES(%s,%s,%s,%s,%s,%s,%s,%s)",
          (uid, user, rol, accion, modulo, detalle[:500] if detalle else "", ip, igl), commit=True)
    except Exception as e:
        print(f"audit log failed: {e}")

def mensaje_error_informe(e, admin=False):
    """Traduce un error técnico a un mensaje comprensible. El detalle técnico
    no se muestra al líder: queda en el log (log_error_informe) y en
    Configuración › pestaña de Errores. admin=True lo deja más seco, para el
    formulario interno."""
    t = str(e).lower()
    if "duplicate key" in t or "unique" in t:
        return ("Ya enviaste un informe con esa fecha y ese timoteo. "
                "Si quieres corregir uno, repite exactamente la misma fecha y el mismo timoteo.")
    if "could not connect" in t or "connection" in t or "timeout" in t or "server closed" in t:
        return ("No pudimos conectar para guardar. Revisa tu internet e intenta de nuevo"
                + ("." if admin else ". Lo que escribiste quedó en pantalla: inténtalo de nuevo en un momento."))
    if "foreign key" in t:
        return ("No se pudo guardar porque falta información de tu célula"
                + ("; revisa sus datos." if admin else ". Avísale a tu pastor."))
    if "invalid input" in t or "invalid text representation" in t or "not a valid" in t:
        return "Revisa los datos ingresados: la ofrenda debe ser solo números."
    return "No se pudo guardar el informe" + ("." if admin else
           ". Lo que escribiste quedó en pantalla: presiona «Enviar informe» de nuevo. "
           "Si sigue fallando, avisa a tu pastor.")

def log_error_informe(origen, lid, e, extra=""):
    """Registra en auditoría (accion=ERROR_GUARDADO, modulo=celulas) un error al
    guardar un informe de célula, para que SUPERADMIN/ADMIN lo revisen en Configuración."""
    try:
        nombre = ""
        lid_int = None
        if lid:
            try:
                lid_int = int(lid)
            except (TypeError, ValueError):
                lid_int = None
        if lid_int:
            r = q("""SELECT h.nombre_completo FROM lideres l
                      JOIN hermanos h ON h.id=l.hermano_id WHERE l.id=%s""",
                   (lid_int,), one=True)
            nombre = r["nombre_completo"] if r else ""
        detalle = f"[{origen}] líder {lid or '?'} {nombre} — {e}"
        if extra:
            detalle += f" | {extra}"
        log_audit("ERROR_GUARDADO", "celulas", detalle)
    except Exception as ex:
        print(f"log_error_informe failed: {ex}")

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

# ---------- PROTECCIÓN ANTI-BOTS EN LOGIN ----------
LOGIN_FALLOS_USER   = 5     # claves incorrectas por usuario (ventana) → bloqueo
LOGIN_FALLOS_IP     = 30    # intentos fallidos por IP    (ventana) → bloqueo
LOGIN_VENTANA_MIN   = 15    # ventana de conteo en minutos
BLOQUEO_DURACIONES  = (15, 60, 1440)   # minutos por nivel: 1ª vez, 2ª vez, 3ª vez en adelante
BLOQUEO_MEMORIA_DIAS = 7    # cuántos días se recuerda para escalar el bloqueo
RECUPERAR_USER_HORA = 3     # recuperaciones de clave por usuario/hora
RECUPERAR_IP_HORA   = 10    # recuperaciones de clave por IP/hora
LOGIN_ESPERA_SEG    = 1.5   # tiempo mínimo de llenado del formulario

# --- Cloudflare Turnstile (CAPTCHA invisible) ---
# Se activa SOLO si hay ambas claves en el entorno (TURNSTILE_SITE_KEY y
# TURNSTILE_SECRET_KEY). Sin claves, el escudo sigue siendo el anterior
# (captcha matemático + honeypot + bloqueo exponencial) — fail-open documentado.
TURNSTILE_URL = "https://challenges.cloudflare.com/turnstile/v0/siteverify"

def claves_turnstile():
    """(site_key, secret_key) desde el entorno. Vacías = Turnstile desactivado."""
    return (os.environ.get("TURNSTILE_SITE_KEY", "").strip(),
            os.environ.get("TURNSTILE_SECRET_KEY", "").strip())

def turnstile_sitekey_visible():
    """Site key solo si está completa la pareja (para inyectar el widget)."""
    site, secret = claves_turnstile()
    return site if (site and secret) else ""

def _ip_privada(ip):
    """¿IP local/privada? Solo se le informa a Cloudflare la IP pública real:
    127.0.0.1 (desarrollo local) o IPs de red interna no aportan nada."""
    try:
        p = ip.split(".")
        if len(p) == 4 and all(x.isdigit() for x in p):
            a, b = int(p[0]), int(p[1])
            return (a == 127 or a == 10 or (a == 172 and 16 <= b <= 31)
                    or (a == 192 and b == 168) or (a == 169 and b == 254))
    except Exception:
        pass
    return ip.startswith(("::1", "fc", "fd", "fe80"))

def verificar_turnstile(token):
    """Verifica el token de Turnstile contra el servidor de Cloudflare.
    - Sin claves configuradas → True (sigue vigente el escudo anterior).
    - Con claves → False si falta el token o Cloudflare lo rechaza.
    - Si Cloudflare no responde (caída externa) → True: no se deja fuera a los
      usuarios por un fallo de terceros; el bloqueo por intentos sigue activo.
    - EN LOCALHOST (desarrollo) → True siempre: evita bloqueos por Tracking Prevention."""
    site, secret = claves_turnstile()
    if not (site and secret):
        return True
    # LOCAL: saltar Turnstile para pruebas (claves TEST + localhost)
    host = (request.headers.get("Host", "") or "").split(":")[0].lower()
    if host in ("localhost", "127.0.0.1") and site.startswith("1x000"):
        return True
    if not token:
        return False
    try:
        import urllib.request, urllib.parse, json as _json
        campos = {"secret": secret, "response": token}
        ip = ip_cliente()
        if ip and not _ip_privada(ip):
            campos["remoteip"] = ip
        datos = urllib.parse.urlencode(campos).encode()
        req = urllib.request.Request(TURNSTILE_URL, data=datos)
        with urllib.request.urlopen(req, timeout=6) as r:
            res = _json.loads(r.read().decode())
        return bool(res.get("success"))
    except Exception as e:
        print(f"turnstile verify failed: {e}")
        return True

def ip_cliente():
    """IP real del cliente: detrás del proxy de Render, remote_addr es la IP
    interna del proxy; la del cliente viene en X-Forwarded-For (primera = origen)."""
    return (request.headers.get("X-Forwarded-For", "") or "").split(",")[0].strip() or (request.remote_addr or "")

def registrar_intento(username, ip, exito, detalle=""):
    """Guarda un intento de login y purga los mayores a 30 días. Nunca rompe el login."""
    try:
        q("""WITH viejos AS (DELETE FROM intentos_login WHERE creado < now() - make_interval(days => 30))
             INSERT INTO intentos_login(ip, username, exito, detalle) VALUES(%s,%s,%s,%s)""",
          (ip or "", (username or "")[:100], bool(exito), (detalle or "")[:50]), commit=True)
    except Exception as e:
        print(f"registrar_intento failed: {e}")

def _bloqueos_recientes(detalle, usuario, ip):
    """Minutos transcurridos desde cada activación de bloqueo de la ventana (la más nueva al final).
    'bloqueo_user' se cuenta por usuario; 'bloqueo_ip' por IP (así un atacante que cambia de
    usuario en la misma IP también escala)."""
    if detalle == "bloqueo_user":
        return q("""SELECT extract(epoch from (now() - creado))/60 AS transcurridas
            FROM intentos_login
            WHERE detalle=%s AND username=%s
              AND creado > now() - make_interval(days => %s)
            ORDER BY creado""", (detalle, usuario or "", BLOQUEO_MEMORIA_DIAS))
    return q("""SELECT extract(epoch from (now() - creado))/60 AS transcurridas
        FROM intentos_login
        WHERE detalle=%s AND ip=%s
          AND creado > now() - make_interval(days => %s)
        ORDER BY creado""", (detalle, ip or "", BLOQUEO_MEMORIA_DIAS))

def _bloqueo_activo(detalle, usuario, ip):
    """(minutos_restantes, nivel) si hay un bloqueo vigente; None si no lo hay."""
    try:
        filas = _bloqueos_recientes(detalle, usuario, ip)
    except Exception as e:
        print(f"_bloqueo_activo failed: {e}")
        return None
    if not filas:
        return None
    nivel = min(len(filas), len(BLOQUEO_DURACIONES))
    duracion = BLOQUEO_DURACIONES[nivel - 1]
    restantes = duracion - float(filas[-1]["transcurridas"] or 0)
    return (restantes, nivel) if restantes > 0 else None

def _activar_bloqueo(detalle, usuario, ip):
    """Registra la activación de un bloqueo nuevo y devuelve (minutos, nivel) que le corresponde."""
    registrar_intento(usuario if detalle == "bloqueo_user" else "", ip, False, detalle)
    try:
        nivel = min(len(_bloqueos_recientes(detalle, usuario, ip) or []), len(BLOQUEO_DURACIONES))
    except Exception:
        nivel = 0
    nivel = max(nivel, 1)
    return BLOQUEO_DURACIONES[nivel - 1], nivel

def texto_espera(minutos):
    """15 → '15 minutos'; 60 → '1 hora'; 1440 → '24 horas'."""
    m = max(int(round(minutos)), 1)
    if m >= 120:
        h = round(m / 60)
        return f"{h} horas"
    if m >= 60:
        return "1 hora"
    return f"{m} minutos"

def minutos_bloqueo(username, ip):
    """Minutos restantes de bloqueo (0 = puede intentar).
    - Un bloqueo vigente manda sobre el conteo (la ventana de fallos es más corta
      que los bloqueos de nivel alto, si no se liberaría demasiado pronto).
    - Al superar el umbral se ACTIVA un bloqueo nuevo y este escala mientras siga
      reincidiendo: 15 min → 1 hora → 24 horas (memoria de 7 días), por usuario
      y por IP por separado.
    - Solo cuentan fallos posteriores al último login correcto, así un usuario
      válido no queda castigado."""
    # 0) ¿Hay un bloqueo vigente?
    for detalle, usuario in (("bloqueo_user", username), ("bloqueo_ip", "")):
        b = _bloqueo_activo(detalle, usuario, ip)
        if b:
            return max(b[0], 1)
    # 1) ¿Se superó el umbral? → activar bloqueo nuevo (escala)
    try:
        row = q("""
            SELECT
              (SELECT count(*) FROM intentos_login x
                 WHERE NOT x.exito AND x.detalle='credenciales' AND x.username=%s
                   AND x.creado > now() - make_interval(mins => %s)
                   AND x.creado > COALESCE((SELECT max(y.creado) FROM intentos_login y
                                            WHERE y.exito AND y.username=%s), '-infinity'::timestamp)
              ) AS fallos_user,
              (SELECT count(*) FROM intentos_login x
                 WHERE NOT x.exito AND x.ip=%s
                   AND x.creado > now() - make_interval(mins => %s)
                   AND x.creado > COALESCE((SELECT max(y.creado) FROM intentos_login y
                                            WHERE y.exito AND y.ip=%s), '-infinity'::timestamp)
              ) AS fallos_ip""",
            (username or "", LOGIN_VENTANA_MIN, username or "",
             ip or "", LOGIN_VENTANA_MIN, ip or ""), one=True)
        if not row:
            return 0
        if (row["fallos_user"] or 0) >= LOGIN_FALLOS_USER:
            return _activar_bloqueo("bloqueo_user", username or "", ip)[0]
        if (row["fallos_ip"] or 0) >= LOGIN_FALLOS_IP:
            return _activar_bloqueo("bloqueo_ip", "", ip)[0]
    except Exception as e:
        print(f"minutos_bloqueo failed: {e}")
    return 0

def nuevo_captcha(clave):
    """Genera una suma sencilla, la guarda en la sesión y devuelve el texto a mostrar."""
    a, b = random.randint(1, 9), random.randint(1, 9)
    session[clave] = a + b
    session[clave + "_ts"] = time.time()
    return f"¿Cuánto es {a} + {b}?"

def chequear_form_seguro(clave):
    """Valida honeypot + captcha matemático + CAPTCHA Turnstile + tiempo mínimo.
    Devuelve (mensaje_de_error, tipo) — (None, None) si todo está bien.
    El tipo se guarda en intentos_login (estos fallos cuentan por IP, no por
    usuario, para que nadie pueda bloquear una cuenta ajena desde fuera)."""
    esperado = session.pop(clave, None)
    ts = session.pop(clave + "_ts", None)
    if (request.form.get("sitio") or "").strip():
        return "Solicitud rechazada.", "honeypot"      # campo trampa: lo llenó un robot
    if esperado is None:
        return "La sesión expiró. Recargue la página e intente de nuevo.", "sesion"
    resp = (request.form.get("captcha") or "").strip()
    if not resp.isdigit() or int(resp) != int(esperado):
        return "Verificación incorrecta. Revise la respuesta e intente de nuevo.", "captcha"
    if not verificar_turnstile(request.form.get("cf-turnstile-response", "")):
        return "Verificación anti-bot fallida. Espere un segundo y reintente, o recargue la página.", "turnstile"
    if ts and (time.time() - float(ts)) < LOGIN_ESPERA_SEG:
        return "Envío demasiado rápido. Espere un momento y reintente.", "rapido"
    return None, None

# ---------- AUTH ----------
@app.get("/login")
def login():
    pregunta = nuevo_captcha("captcha_login")
    return render_template("login.html", captcha_pregunta=pregunta,
                           turnstile_sitekey=turnstile_sitekey_visible())

@app.post("/login")
def login_post():
    u = request.form.get("username", "").strip()
    p = request.form.get("password", "")
    ip = ip_cliente()
    # 1) ¿Bloqueado por intentos fallidos? (escala 15 min → 1 h → 24 h)
    mins = minutos_bloqueo(u, ip)
    if mins:
        registrar_intento(u, ip, False, "bloqueado")
        log_audit("LOGIN_BLOQUEADO", "usuarios", f"bloqueado {u or '-'} ip {ip}")
        flash(f"Demasiados intentos fallidos. Espere {texto_espera(mins)} antes de volver a intentar.", "error")
        return redirect(url_for("login"))
    # 2) Honeypot + captcha matemático + Turnstile + tiempo mínimo
    err, tipo = chequear_form_seguro("captcha_login")
    if err:
        registrar_intento(u, ip, False, tipo)
        flash(err, "error")
        return redirect(url_for("login"))
    # 3) Credenciales (la iglesia activa se valida aquí: hija desactivada no entra)
    row = q("""SELECT u.id,u.password_hash,u.rol,u.primer_login,u.iglesia_id,i.nombre AS iglesia_nombre,
        COALESCE(i.activo,true) AS iglesia_activa
        FROM usuarios u LEFT JOIN iglesias i ON i.id=u.iglesia_id WHERE u.username=%s AND u.activo""", (u,), one=True)
    if not row or not check_password_hash(row["password_hash"], p):
        registrar_intento(u, ip, False, "credenciales")
        log_audit("LOGIN_FALLIDO", "usuarios", f"login fallido {u} ip {ip}")
        flash("Usuario o clave inválidos", "error"); return redirect(url_for("login"))
    registrar_intento(u, ip, True, "web")      # login correcto: resetea contadores
    if row["rol"] != "SUPERADMIN" and row.get("iglesia_activa") is False:
        flash("Tu iglesia esta desactivada (fuera de operacion). Contacta a casa matriz.", "error")
        return redirect(url_for("login"))
    session["uid"], session["rol"], session["user"] = row["id"], row["rol"], u
    session["iglesia_id"] = row["iglesia_id"] or 1
    session["iglesia_nombre"] = row["iglesia_nombre"] or "CASA MATRIZ"
    if row["rol"] == "SUPERADMIN":
        session["ver_todo"] = True
    else:
        session.pop("ver_todo", None)
    q("UPDATE usuarios SET ultimo_login=now() WHERE id=%s", (row["id"],), commit=True)
    log_audit("LOGIN", "usuarios", f"login {u}")
    if row.get("primer_login"):
        flash("Es su primer ingreso. Debe cambiar su contraseña.", "ok")
        return redirect(url_for("cambiar_clave"))
    return redirect(url_for("index"))

@app.get("/ver-iglesia")
@login_req
def ver_iglesia():
    """SUPERADMIN: cambiar iglesia en vista (o volver a TODAS)."""
    if session.get("rol") != "SUPERADMIN":
        flash("Solo SUPERADMIN cambia de iglesia en vista", "error")
        return redirect(url_for("index"))
    ig = request.args.get("iglesia", "")
    if ig in ("", "todas"):
        session["ver_todo"] = True
    else:
        try: iid = int(ig)
        except: return redirect(request.referrer or url_for("index"))
        r = q("SELECT id,nombre FROM iglesias WHERE id=%s AND activo", (iid,), one=True)
        if r:
            session["ver_todo"] = False
            session["iglesia_id"] = r["id"]
            session["iglesia_nombre"] = r["nombre"]
    return redirect(request.referrer or url_for("index"))

@app.get("/logout")
def logout():
    try: log_audit("LOGOUT", "usuarios", f"logout {session.get('user','')}")
    except: pass
    session.clear(); return redirect(url_for("login"))

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
def recuperar():
    pregunta = nuevo_captcha("captcha_rec")
    return render_template("recuperar.html", captcha_pregunta=pregunta,
                           turnstile_sitekey=turnstile_sitekey_visible())

@app.post("/recuperar")
def recuperar_post():
    from werkzeug.security import generate_password_hash
    u = request.form.get("username","").strip()
    ip = ip_cliente()
    # Límite: si ya se pidieron demasiadas recuperaciones, se corta (evita que
    # un bot deje claves de otros usuarios en "primer login" sin control).
    try:
        row = q("""SELECT
            (SELECT count(*) FROM intentos_login
              WHERE detalle='recuperar' AND username=%s
                AND creado > now() - make_interval(hours => 1)) AS n_user,
            (SELECT count(*) FROM intentos_login
              WHERE detalle='recuperar' AND ip=%s
                AND creado > now() - make_interval(hours => 1)) AS n_ip""", (u, ip), one=True)
        if row and ((row["n_user"] or 0) >= RECUPERAR_USER_HORA or (row["n_ip"] or 0) >= RECUPERAR_IP_HORA):
            registrar_intento(u, ip, False, "recuperar")
            log_audit("RECUPERAR_BLOQUEADO", "usuarios", f"recuperar bloqueado {u} ip {ip}")
            flash("Demasiadas solicitudes de recuperación. Intente más tarde o contacte al administrador.", "error")
            return redirect(url_for("recuperar"))
    except Exception as e:
        print(f"limite recuperar failed: {e}")
    # Honeypot + captcha matemático + Turnstile + tiempo mínimo (mismo escudo del login)
    err, tipo = chequear_form_seguro("captcha_rec")
    if err:
        registrar_intento(u, ip, False, tipo)
        flash(err, "error")
        return redirect(url_for("recuperar"))
    registrar_intento(u, ip, False, "recuperar")   # cuenta para el límite por hora
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
    log_audit("RECUPERAR_CLAVE", "usuarios", f"recuperación solicitada para {u} ip {ip}")
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
    {"key":"celulas","nombre":"Informe Célula"},{"key":"configuracion","nombre":"Configuración"},
    {"key":"auditoria","nombre":"Auditoría"},
]
ROLES_LIST = ["SUPERADMIN","ADMIN","CONSULTA"]

def obtener_roles():
    try:
        rows = q("SELECT nombre FROM catalogo_roles_usuario WHERE activo=true ORDER BY nombre")
        return [r["nombre"] for r in rows] if rows else ROLES_LIST
    except:
        return ROLES_LIST

# ---------- Jerarquía de roles: nadie gestiona cuentas iguales o superiores ----------
# Un ADMIN no puede tomar una cuenta SUPERADMIN (ni la de otro ADMIN) cambiándole
# la contraseña, el rol o desactivándola. Eso solo lo hace SUPERADMIN.
_JERARQUIA_ROLES = {"SUPERADMIN": 100, "ADMIN": 80, "ADMIN_HIJA": 70,
                    "ADMIN_GRUPO": 60, "CONSULTA": 20, "VISTA": 10}

def puede_asignar_rol(rol_nuevo):
    """Rol que se le puede OTORGAR a una cuenta: siempre por debajo del propio.
    Un rol fuera de la jerarquía (catálogo nuevo) no se otorga, por prudencia."""
    mio = _JERARQUIA_ROLES.get((session.get("rol", "") or "").upper(), 0)
    if mio >= _JERARQUIA_ROLES["SUPERADMIN"]:
        return True
    r = _JERARQUIA_ROLES.get((rol_nuevo or "").upper())
    return r is not None and r < mio

def puede_gestionar_rol(rol_objetivo, uid_objetivo=None):
    """¿Se puede actuar sobre esa cuenta (editar, toggle, eliminar, resetear clave)?
    Nunca sobre cuentas del mismo nivel o por encima. Editarse a uno mismo sí está
    permitido: su rol igual no se puede elevar, eso lo decide puede_asignar_rol."""
    if uid_objetivo is not None and uid_objetivo == session.get("uid"):
        return True
    return puede_asignar_rol(rol_objetivo)

@app.get("/usuarios")
@login_req
@permiso_req("usuarios", "leer")
def usuarios():
    fi, pi = filtro_igl("u")
    users = q(f"""SELECT u.*, h.nombre_completo AS h_nombre, l2.nombre_completo AS l_nombre, i.nombre AS iglesia
        FROM usuarios u LEFT JOIN hermanos h ON h.id=u.hermano_id LEFT JOIN lideres l ON l.id=u.lider_id
        LEFT JOIN hermanos l2 ON l2.id=l.hermano_id LEFT JOIN iglesias i ON i.id=u.iglesia_id
        WHERE 1=1{fi} ORDER BY u.rol, u.username""", tuple(pi))
    for u in users:
        u["vinculo"] = u.get("h_nombre") or u.get("l_nombre") or None
    perm = permisos(session.get("rol",""))
    roles_dyn = obtener_roles()
    perm_all = {}
    for rl in roles_dyn:
        perm_all[rl] = permisos(rl)
    return render_template("usuarios.html", usuarios=users, perm=perm_all, modulos=MODULOS_INFO, roles=roles_dyn)

@app.get("/usuarios/nuevo")
@login_req
@permiso_req("usuarios", "crear")
def usuario_nuevo():
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY h.nombre_completo""", tuple(pl))
    igls = q("SELECT id,nombre,tipo FROM iglesias WHERE activo ORDER BY id") if session.get("rol") == "SUPERADMIN" else []
    return render_template("usuario_form.html", u=None, rol=session.get("rol",""), lideres=lideres, roles=obtener_roles(), iglesias=igls)

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
    if lider_id and not exigir_igl("lideres", lider_id):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("usuario_nuevo"))
    if q("SELECT id FROM usuarios WHERE username=%s", (username,), one=True):
        flash(f"El nombre '{username}' ya esta en uso en el sistema (los nombres son unicos globales). Prueba con otro.", "error")
        return redirect(url_for("usuario_nuevo"))
    # Iglesia: SUPERADMIN elige; el resto queda fijo en la suya
    if rol_actual == "SUPERADMIN":
        igl_nueva = int(f.get("iglesia_id") or 1)
    else:
        igl_nueva = session.get("iglesia_id") or 1
    if not pwd:
        flash("La contraseña es obligatoria", "error"); return redirect(url_for("usuario_nuevo"))
    errores = validar_clave(pwd)
    if errores:
        flash(f"La contraseña no cumple: {', '.join(errores)}", "error"); return redirect(url_for("usuario_nuevo"))
    if not puede_asignar_rol(nuevo_rol):
        flash("No tiene permiso para crear usuarios con un rol igual o superior al suyo", "error"); return redirect(url_for("usuario_nuevo"))
    try:
        q("INSERT INTO usuarios(username,password_hash,rol,activo,primer_login,lider_id,iglesia_id) VALUES(%s,%s,%s,%s,TRUE,%s,%s)",
          (username, generate_password_hash(pwd), f["rol"], f.get("activo")=="on", lider_id, igl_nueva), commit=True)
        log_audit("CREAR","usuarios", f"{username} rol={f.get('rol')} iglesia={igl_nueva}")
        flash("Usuario creado", "ok")
    except Exception as e: flash(f"Error: {e}", "error"); return redirect(url_for("usuario_nuevo"))
    return redirect(url_for("usuarios"))

@app.get("/usuarios/<int:uid>/editar")
@login_req
@permiso_req("usuarios", "editar")
def usuario_editar(uid):
    if not exigir_igl("usuarios", uid):
        flash("Ese usuario no es de tu iglesia", "error"); return redirect(url_for("usuarios"))
    u = q("SELECT u.*, i.nombre AS iglesia FROM usuarios u LEFT JOIN iglesias i ON i.id=u.iglesia_id WHERE u.id=%s", (uid,), one=True)
    if not u: flash("Usuario no encontrado", "error"); return redirect(url_for("usuarios"))
    if not puede_gestionar_rol(u["rol"], uid):
        flash("No puede editar una cuenta con igual o mayor privilegio que la suya", "error")
        return redirect(url_for("usuarios"))
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY h.nombre_completo""", tuple(pl))
    igls = q("SELECT id,nombre,tipo FROM iglesias WHERE activo ORDER BY id") if session.get("rol") == "SUPERADMIN" else []
    return render_template("usuario_form.html", u=u, rol=session.get("rol",""), lideres=lideres, roles=obtener_roles(), iglesias=igls)

@app.post("/usuarios/<int:uid>/editar")
@login_req
@permiso_req("usuarios", "editar")
def usuario_update(uid):
    from werkzeug.security import generate_password_hash
    f = request.form
    if not exigir_igl("usuarios", uid):
        flash("Ese usuario no es de tu iglesia", "error"); return redirect(url_for("usuarios"))
    objetivo = q("SELECT rol FROM usuarios WHERE id=%s", (uid,), one=True)
    if not objetivo:
        flash("Usuario no encontrado", "error"); return redirect(url_for("usuarios"))
    if not puede_gestionar_rol(objetivo["rol"], uid):
        flash("No puede editar una cuenta con igual o mayor privilegio que la suya", "error")
        return redirect(url_for("usuarios"))
    pwd = f.get("password","").strip()
    rol_actual = session.get("rol","")
    nuevo_rol = f["rol"]
    lider_id = f.get("lider_id") or None
    if lider_id and not exigir_igl("lideres", lider_id):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("usuarios"))
    # Iglesia: solo SUPERADMIN puede cambiarla
    igl_sql, igl_params = "", []
    if rol_actual == "SUPERADMIN" and f.get("iglesia_id"):
        igl_sql, igl_params = ", iglesia_id=%s", [int(f.get("iglesia_id"))]
    if not puede_asignar_rol(nuevo_rol):
        flash("No tiene permiso para asignar un rol igual o superior al suyo", "error")
        return redirect(url_for("usuarios"))
    if pwd:
        errores = validar_clave(pwd)
        if errores:
            flash(f"La contraseña no cumple: {', '.join(errores)}", "error")
            return redirect(url_for("usuario_editar", uid=uid))
    try:
        if pwd:
            q(f"UPDATE usuarios SET rol=%s, activo=%s, password_hash=%s, lider_id=%s{igl_sql} WHERE id=%s",
              (f["rol"], f.get("activo")=="on", generate_password_hash(pwd), lider_id) + tuple(igl_params) + (uid,), commit=True)
        else:
            q(f"UPDATE usuarios SET rol=%s, activo=%s, lider_id=%s{igl_sql} WHERE id=%s",
              (f["rol"], f.get("activo")=="on", lider_id) + tuple(igl_params) + (uid,), commit=True)
        log_audit("EDITAR","usuarios", f"id={uid}")
        flash("Usuario actualizado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:uid>/toggle")
@login_req
@permiso_req("usuarios", "editar")
def usuario_toggle(uid):
    if not exigir_igl("usuarios", uid):
        flash("Ese usuario no es de tu iglesia", "error"); return redirect(url_for("usuarios"))
    objetivo = q("SELECT rol FROM usuarios WHERE id=%s", (uid,), one=True)
    if not objetivo:
        flash("Usuario no encontrado", "error"); return redirect(url_for("usuarios"))
    if not puede_gestionar_rol(objetivo["rol"], uid):
        flash("No puede desactivar una cuenta con igual o mayor privilegio que la suya", "error")
        return redirect(url_for("usuarios"))
    q("UPDATE usuarios SET activo=NOT activo WHERE id=%s", (uid,), commit=True)
    flash("Estado actualizado", "ok"); return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:uid>/eliminar")
@login_req
@permiso_req("usuarios", "eliminar")
def usuario_eliminar(uid):
    if not exigir_igl("usuarios", uid):
        flash("Ese usuario no es de tu iglesia", "error"); return redirect(url_for("usuarios"))
    u = q("SELECT rol FROM usuarios WHERE id=%s", (uid,), one=True)
    if not u:
        flash("Usuario no encontrado", "error")
    elif not puede_gestionar_rol(u["rol"], uid):
        flash("No puede eliminar una cuenta con igual o mayor privilegio que la suya", "error")
    elif u["rol"] == "SUPERADMIN":
        flash("No se puede eliminar un SUPERADMIN", "error")
    else:
        q("DELETE FROM usuarios WHERE id=%s", (uid,), commit=True)
        log_audit("ELIMINAR","usuarios", f"id={uid}")
        flash("Usuario eliminado", "ok")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:uid>/reset_clave")
@login_req
@permiso_req("usuarios", "editar")
def usuario_reset_clave(uid):
    if not exigir_igl("usuarios", uid):
        flash("Ese usuario no es de tu iglesia", "error")
        return redirect(url_for("usuarios"))
    objetivo = q("SELECT rol FROM usuarios WHERE id=%s", (uid,), one=True)
    if not objetivo:
        flash("Usuario no encontrado", "error")
        return redirect(url_for("usuarios"))
    if not puede_gestionar_rol(objetivo["rol"], uid):
        flash("No puede cambiar la contraseña de una cuenta con igual o mayor privilegio que la suya", "error")
        return redirect(url_for("usuarios"))
    nueva = request.form.get("nueva_clave","").strip()
    if not nueva:
        flash("Debe ingresar una nueva contraseña", "error")
        return redirect(url_for("usuarios"))
    err = validar_clave(nueva)
    if err:
        flash(f"La clave no cumple: {', '.join(err)}", "error")
        return redirect(url_for("usuarios"))
    from werkzeug.security import generate_password_hash
    q("UPDATE usuarios SET password_hash=%s, primer_login=TRUE WHERE id=%s", (generate_password_hash(nueva), uid), commit=True)
    log_audit("RESET_CLAVE","usuarios", f"id={uid}")
    flash(f"Contraseña de usuario {uid} actualizada — deberá cambiarla al ingresar", "ok")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/permisos")
@login_req
@permiso_req("usuarios", "editar")
def usuario_permisos():
    if session.get("rol") != "SUPERADMIN":
        flash("Solo SUPERADMIN edita la matriz de permisos global", "error")
        return redirect(url_for("usuarios"))
    f = request.form
    for rol in obtener_roles():
        for mod in MODULOS_INFO:
            key = mod["prefix"] if "prefix" in mod else mod["key"]
            leer = f.get(f"{rol}_{key}_leer") is not None
            crear = f.get(f"{rol}_{key}_crear") is not None
            editar = f.get(f"{rol}_{key}_editar") is not None
            eliminar = f.get(f"{rol}_{key}_eliminar") is not None
            exportar = f.get(f"{rol}_{key}_exportar") is not None
            q("""INSERT INTO roles_permisos(rol,modulo,puede_leer,puede_crear,puede_editar,puede_eliminar,puede_exportar,iglesia_id)
                VALUES(%s,%s,%s,%s,%s,%s,%s,NULL) ON CONFLICT (rol,modulo,iglesia_id) DO UPDATE SET
                puede_leer=%s,puede_crear=%s,puede_editar=%s,puede_eliminar=%s,puede_exportar=%s""",
              (rol, key, leer, crear, editar, eliminar, exportar, leer, crear, editar, eliminar, exportar), commit=True)
    flash("Permisos actualizados", "ok"); return redirect(url_for("usuarios"))

# ---------- IGLESIAS (configuracion global, solo SUPERADMIN) ----------
def _solo_superadmin():
    if session.get("rol") != "SUPERADMIN":
        flash("Solo SUPERADMIN gestiona iglesias", "error")
        return False
    return True

@app.get("/iglesias")
@login_req
def iglesias():
    if not _solo_superadmin(): return redirect(url_for("index"))
    rows = q("""SELECT i.*, p.nombre AS padre_nombre,
        (SELECT COUNT(*) FROM hermanos h WHERE h.iglesia_id=i.id AND h.activo) AS hermanos,
        (SELECT COUNT(*) FROM lideres l WHERE l.iglesia_id=i.id AND l.activo) AS lideres,
        (SELECT COUNT(*) FROM usuarios u WHERE u.iglesia_id=i.id AND u.activo) AS usuarios
        FROM iglesias i LEFT JOIN iglesias p ON p.id=i.iglesia_padre_id ORDER BY i.id""")
    return render_template("iglesias.html", rows=rows)

@app.get("/iglesias/nuevo")
@login_req
def iglesia_nueva():
    if not _solo_superadmin(): return redirect(url_for("index"))
    padres = q("SELECT id,nombre,tipo FROM iglesias WHERE activo ORDER BY id")
    return render_template("iglesia_form.html", ig=None, padres=padres)

@app.post("/iglesias/nuevo")
@login_req
def iglesia_crear():
    if not _solo_superadmin(): return redirect(url_for("index"))
    f = request.form
    nombre = (f.get("nombre") or "").strip().upper()
    tipo = f.get("tipo") or "HIJA"
    padre = f.get("padre") or None
    descripcion = (f.get("descripcion") or "").strip()
    pref = (f.get("prefijo") or "").strip().upper()
    if not nombre:
        flash("El nombre es obligatorio", "error"); return redirect(url_for("iglesias"))
    if pref and (not pref.isascii() or not pref.isalnum() or not (2 <= len(pref) <= 6)):
        flash("El prefijo del código debe tener de 2 a 6 letras o números (ej. NOR)", "error")
        return redirect(url_for("iglesia_nueva"))
    if not pref:
        pref = _prefijo_sugerido(nombre)
    if q("SELECT 1 FROM iglesias WHERE UPPER(prefijo_codigo)=%s LIMIT 1", (pref,), one=True):
        flash(f"El prefijo {pref} ya está usado por otra iglesia: elige otro", "error")
        return redirect(url_for("iglesia_nueva"))
    if q("SELECT 1 FROM hermanos WHERE codigo ~ %s LIMIT 1", ("^" + pref + "[0-9]{4}$",), one=True):
        flash(f"Ya existen códigos {pref}#### en el sistema: elige otro prefijo", "error")
        return redirect(url_for("iglesia_nueva"))
    try:
        c = db()
        with c.cursor() as cur:
            cur.execute("""INSERT INTO iglesias(nombre,tipo,iglesia_padre_id,descripcion,prefijo_codigo)
                VALUES(%s,%s,%s,%s,%s) RETURNING id""",
                (nombre, tipo, int(padre) if padre else None, descripcion, pref))
            nuevo_id = cur.fetchone()["id"]
            cur.execute("CREATE SEQUENCE IF NOT EXISTS seq_codigo_igl_%d" % nuevo_id)
            c.commit()
        log_audit("CREAR","iglesias", f"{nombre} tipo={tipo} prefijo={pref}")
        flash(f"Iglesia creada: sus hermanos se numerarán como {pref}0001", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("iglesias"))

@app.get("/iglesias/<int:iid>/editar")
@login_req
def iglesia_editar(iid):
    if not _solo_superadmin(): return redirect(url_for("index"))
    ig = q("SELECT * FROM iglesias WHERE id=%s", (iid,), one=True)
    if not ig: flash("No encontrada", "error"); return redirect(url_for("iglesias"))
    padres = q("SELECT id,nombre,tipo FROM iglesias WHERE activo AND id!=%s ORDER BY id", (iid,))
    return render_template("iglesia_form.html", ig=ig, padres=padres)

@app.post("/iglesias/<int:iid>/editar")
@login_req
def iglesia_update(iid):
    if not _solo_superadmin(): return redirect(url_for("index"))
    if iid == 1:
        flash("CASA MATRIZ no se edita: es la raiz del sistema", "error"); return redirect(url_for("iglesias"))
    f = request.form
    ig = q("SELECT nombre, COALESCE(prefijo_codigo,'') AS prefijo FROM iglesias WHERE id=%s", (iid,), one=True)
    pref = (f.get("prefijo") or "").strip().upper() or (ig["prefijo"] if ig else "")
    if pref and (not pref.isascii() or not pref.isalnum() or not (2 <= len(pref) <= 6)):
        flash("El prefijo del código debe tener de 2 a 6 letras o números (ej. NOR)", "error")
        return redirect(url_for("iglesia_editar", iid=iid))
    if pref != (ig["prefijo"] if ig else "").upper():
        if q("SELECT 1 FROM iglesias WHERE UPPER(prefijo_codigo)=%s AND id!=%s LIMIT 1", (pref, iid), one=True):
            flash(f"El prefijo {pref} ya está usado por otra iglesia: elige otro", "error")
            return redirect(url_for("iglesia_editar", iid=iid))
        if q("SELECT 1 FROM hermanos WHERE codigo ~ %s AND iglesia_id!=%s LIMIT 1", ("^" + pref + "[0-9]{4}$", iid), one=True):
            flash(f"Ya existen códigos {pref}#### de otra iglesia: elige otro prefijo", "error")
            return redirect(url_for("iglesia_editar", iid=iid))
        # Con el prefijo nuevo la numeración continúa desde el máximo existente
        seq = "seq_codigo_igl_%d" % iid
        q("CREATE SEQUENCE IF NOT EXISTS " + seq, commit=True)
        mx = q("SELECT COALESCE(MAX((substring(codigo FROM '([0-9]+)$'))::INT),0) AS m FROM hermanos WHERE codigo ~ %s",
               ("^" + pref + "[0-9]{4}$",), one=True)["m"]
        q("SELECT setval(%s, %s, false)", (seq, mx + 1), commit=True)
    try:
        padre = f.get("padre") or None
        q("""UPDATE iglesias SET nombre=%s, tipo=%s, iglesia_padre_id=%s, activo=%s,
            descripcion=%s, prefijo_codigo=%s WHERE id=%s""",
          ((f.get("nombre") or "").strip().upper(), f.get("tipo") or "HIJA",
           int(padre) if padre else None, f.get("activo") == "on", (f.get("descripcion") or "").strip(),
           pref, iid), commit=True)
        log_audit("EDITAR","iglesias", f"id={iid} prefijo={pref}")
        flash("Iglesia actualizada", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("iglesias"))

@app.post("/iglesias/<int:iid>/convertir")
@login_req
def iglesia_convertir(iid):
    """GRUPO de crecimiento -> IGLESIA HIJA. Solo SUPERADMIN.
    Conserva id, datos, líderes, informes y configuración (solo cambia el tipo)."""
    if not _solo_superadmin(): return redirect(url_for("index"))
    if iid == 1:
        flash("CASA MATRIZ no se convierte", "error"); return redirect(url_for("iglesias"))
    ig = q("SELECT * FROM iglesias WHERE id=%s", (iid,), one=True)
    if not ig:
        flash("No encontrada", "error"); return redirect(url_for("iglesias"))
    if ig["tipo"] != "GRUPO":
        flash(f"{ig['nombre']} ya es {ig['tipo']}, no es un grupo", "error"); return redirect(url_for("iglesias"))
    try:
        q("UPDATE iglesias SET tipo='HIJA' WHERE id=%s", (iid,), commit=True)
        log_audit("CONVERTIR","iglesias", f"{ig['nombre']} GRUPO->HIJA id={iid}")
        flash(f"{ig['nombre']} ahora es IGLESIA HIJA: conserva sus datos, líderes, informes y configuración. "
              f"Asígnale su administrador (rol ADMIN_HIJA) en Usuarios.", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("iglesias"))

@app.post("/iglesias/<int:iid>/toggle")
@login_req
def iglesia_toggle(iid):
    if not _solo_superadmin(): return redirect(url_for("index"))
    if iid == 1:
        flash("CASA MATRIZ no se puede desactivar", "error"); return redirect(url_for("iglesias"))
    ig = q("SELECT nombre, activo FROM iglesias WHERE id=%s", (iid,), one=True)
    if not ig: flash("No encontrada", "error"); return redirect(url_for("iglesias"))
    q("UPDATE iglesias SET activo=NOT activo WHERE id=%s", (iid,), commit=True)
    log_audit("EDITAR","iglesias", f"toggle id={iid}")
    if ig["activo"]:
        flash(f"{ig['nombre']} DESACTIVADA: sus usuarios ya no pueden entrar (se conserva todo)", "ok")
    else:
        flash(f"{ig['nombre']} reactivada", "ok")
    return redirect(url_for("iglesias"))

@app.post("/iglesias/<int:iid>/eliminar")
@login_req
def iglesia_eliminar(iid):
    if not _solo_superadmin(): return redirect(url_for("index"))
    if iid == 1:
        flash("CASA MATRIZ no se puede eliminar", "error"); return redirect(url_for("iglesias"))
    ig = q("SELECT nombre FROM iglesias WHERE id=%s", (iid,), one=True)
    if not ig: flash("No encontrada", "error"); return redirect(url_for("iglesias"))
    # Solo eliminar si no tiene nada asociado (primero desactivar y vaciar/traspasar)
    deps = q("""SELECT
        (SELECT COUNT(*) FROM hermanos WHERE iglesia_id=%s) AS hermanos,
        (SELECT COUNT(*) FROM lideres WHERE iglesia_id=%s) AS lideres,
        (SELECT COUNT(*) FROM usuarios WHERE iglesia_id=%s) AS usuarios,
        (SELECT COUNT(*) FROM visitas WHERE iglesia_id=%s) AS visitas,
        (SELECT COUNT(*) FROM seguimiento WHERE iglesia_id=%s) AS seguimiento,
        (SELECT COUNT(*) FROM celulas_informe WHERE iglesia_id=%s) AS informes,
        (SELECT COUNT(*) FROM iglesias WHERE iglesia_padre_id=%s) AS hijas""", tuple([iid]*7), one=True)
    con_datos = [k for k in ("hermanos","lideres","usuarios","visitas","seguimiento","informes","hijas") if deps[k] > 0]
    if con_datos:
        # Hay datos: ir a purga completa (confirmacion escribiendo el nombre)
        return redirect(url_for("iglesia_purga", iid=iid))
    # Borrar config propia (catálogos y permisos de la iglesia) y luego la iglesia
    for t in ("catalogo_redes","catalogo_tipo12","catalogo_meses","tipos_encuentro","periodicidad_celula",
              "catalogo_roles_lider","catalogo_estado_civil","catalogo_sexo","catalogo_estado_celula",
              "catalogo_info_enviada","catalogo_estado_visita","catalogo_tipo_seguimiento",
              "catalogo_tipo_hermano","catalogo_modalidad_celula"):
        q(f"DELETE FROM {t} WHERE iglesia_id=%s", (iid,), commit=True)
    q("DELETE FROM roles_permisos WHERE iglesia_id=%s", (iid,), commit=True)
    q("DELETE FROM auditoria WHERE iglesia_id=%s", (iid,), commit=True)
    q("DELETE FROM iglesias WHERE id=%s", (iid,), commit=True)
    try:
        q("SET LOCAL lock_timeout='5s'; DROP SEQUENCE IF EXISTS seq_codigo_igl_%d" % iid, commit=True)
    except Exception:
        try: db().rollback()
        except: pass  # secuencia huérfana inofensiva; no romper la eliminación
    log_audit("ELIMINAR","iglesias", f"id={iid} {ig['nombre']}")
    flash(f"Iglesia {ig['nombre']} eliminada por completo", "ok")
    return redirect(url_for("iglesias"))

@app.get("/iglesias/<int:iid>/purga")
@login_req
def iglesia_purga(iid):
    """Confirmacion de purga completa: muestra conteos y pide escribir el nombre."""
    if not _solo_superadmin(): return redirect(url_for("index"))
    if iid == 1:
        flash("CASA MATRIZ no se puede eliminar", "error"); return redirect(url_for("iglesias"))
    ig = q("SELECT * FROM iglesias WHERE id=%s", (iid,), one=True)
    if not ig: flash("No encontrada", "error"); return redirect(url_for("iglesias"))
    deps = q("""SELECT
        (SELECT COUNT(*) FROM hermanos WHERE iglesia_id=%s) AS hermanos,
        (SELECT COUNT(*) FROM lideres WHERE iglesia_id=%s) AS lideres,
        (SELECT COUNT(*) FROM usuarios WHERE iglesia_id=%s) AS usuarios,
        (SELECT COUNT(*) FROM visitas WHERE iglesia_id=%s) AS visitas,
        (SELECT COUNT(*) FROM seguimiento WHERE iglesia_id=%s) AS seguimiento,
        (SELECT COUNT(*) FROM celulas_informe WHERE iglesia_id=%s) AS informes,
        (SELECT COUNT(*) FROM encuentro_participacion ep JOIN hermanos h ON h.id=ep.hermano_id WHERE h.iglesia_id=%s) AS encuentros,
        (SELECT COUNT(*) FROM discipulado d JOIN lideres l ON l.id=d.lider_id WHERE l.iglesia_id=%s) AS discipulados,
        (SELECT COUNT(*) FROM iglesias WHERE iglesia_padre_id=%s) AS hijas""", tuple([iid]*9), one=True)
    return render_template("iglesia_purga.html", ig=ig, deps=deps)

@app.post("/iglesias/<int:iid>/purga")
@login_req
def iglesia_purga_ejecutar(iid):
    if not _solo_superadmin(): return redirect(url_for("index"))
    if iid == 1:
        flash("CASA MATRIZ no se puede eliminar", "error"); return redirect(url_for("iglesias"))
    ig = q("SELECT nombre FROM iglesias WHERE id=%s", (iid,), one=True)
    if not ig: flash("No encontrada", "error"); return redirect(url_for("iglesias"))
    escrito = (request.form.get("confirmar") or "").strip().upper()
    if escrito != ig["nombre"].upper():
        flash("El nombre no coincide: purga cancelada", "error")
        return redirect(url_for("iglesia_purga", iid=iid))
    hijas = q("SELECT COUNT(*) AS c FROM iglesias WHERE iglesia_padre_id=%s", (iid,), one=True)["c"]
    if hijas > 0:
        flash("Primero elimina o reasigna sus iglesias hijas/grupos", "error")
        return redirect(url_for("iglesia_purga", iid=iid))
    try:
        c = db()
        with c.cursor() as cur:
            cur.execute("DELETE FROM encuentro_participacion WHERE hermano_id IN (SELECT id FROM hermanos WHERE iglesia_id=%s)", (iid,))
            cur.execute("DELETE FROM discipulado WHERE lider_id IN (SELECT id FROM lideres WHERE iglesia_id=%s) OR hermano_id IN (SELECT id FROM hermanos WHERE iglesia_id=%s)", (iid, iid))
            cur.execute("DELETE FROM seguimiento WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM celulas_informe WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM visitas WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM usuarios WHERE iglesia_id=%s", (iid,))
            cur.execute("UPDATE lideres SET lider_padre_id=NULL WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM lideres WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM hermanos WHERE iglesia_id=%s", (iid,))
            for t in ("catalogo_redes","catalogo_tipo12","catalogo_meses","tipos_encuentro","periodicidad_celula",
                      "catalogo_roles_lider","catalogo_estado_civil","catalogo_sexo","catalogo_estado_celula",
                      "catalogo_info_enviada","catalogo_estado_visita","catalogo_tipo_seguimiento",
                      "catalogo_tipo_hermano","catalogo_modalidad_celula"):
                cur.execute(f"DELETE FROM {t} WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM roles_permisos WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM auditoria WHERE iglesia_id=%s", (iid,))
            cur.execute("DELETE FROM iglesias WHERE id=%s", (iid,))
            c.commit()
        # Secuencia aparte y "aunque falle" (con timeout): si algún cliente
        # retiene la secuencia (p. ej. una consulta a pg_sequences en una
        # transacción abierta), la purga NO debe quedar colgada.
        try:
            q("SET LOCAL lock_timeout='5s'; DROP SEQUENCE IF EXISTS seq_codigo_igl_%d" % iid, commit=True)
        except Exception:
            try: db().rollback()
            except: pass  # queda una secuencia huérfana (inofensiva)
        log_audit("ELIMINAR","iglesias", f"purga id={iid} {ig['nombre']}")
        flash(f"Iglesia {ig['nombre']} y TODOS sus datos eliminados", "ok")
    except Exception as e:
        try: c.rollback()
        except: pass
        flash(f"No se pudo purgar: {e}", "error")
    return redirect(url_for("iglesias"))

# ---------- DASHBOARD ----------
@app.get("/")
@login_req
def index():
    d = {}
    fi, pi = filtro_igl()
    fh, ph = filtro_igl("h")
    d["hermanos"] = q(f"SELECT COUNT(*) c FROM hermanos WHERE activo{fi}", tuple(pi), one=True)["c"]
    d["lideres"] = q(f"SELECT COUNT(*) c FROM hermanos WHERE tipo_hermano='LIDER' AND activo{fi}", tuple(pi), one=True)["c"]
    d["disc"] = q(f"SELECT COUNT(*) c FROM hermanos WHERE tipo_hermano='DISCIPULO' AND activo{fi}", tuple(pi), one=True)["c"]
    d["ninguno"] = q(f"SELECT COUNT(*) c FROM hermanos WHERE tipo_hermano='NINGUNO' AND activo{fi}", tuple(pi), one=True)["c"]
    d["pend"] = q(f"SELECT COUNT(*) c FROM encuentro_participacion e JOIN hermanos h ON h.id=e.hermano_id WHERE e.estado='PENDIENTE'{fh}", tuple(ph), one=True)["c"]
    d["top"] = q(f"SELECT * FROM v_lideres_conteo{(' WHERE iglesia_id=%s' if iglesia_actual()[0] is not None else '')} ORDER BY discipulos_activos DESC LIMIT 10", ([iglesia_actual()[0]] if iglesia_actual()[0] is not None else []))
    d["cumple"] = q(f"""SELECT nombre_completo,fecha_nacimiento FROM hermanos WHERE activo{fi}
        AND EXTRACT(MONTH FROM fecha_nacimiento)=EXTRACT(MONTH FROM CURRENT_DATE) ORDER BY EXTRACT(DAY FROM fecha_nacimiento) LIMIT 20""", tuple(pi))
    return render_template("index.html", d=d)

# ---------- HERMANOS ----------
@app.get("/hermanos")
@login_req
def hermanos():
    texto = request.args.get("q", "")
    exp = request.args.get("export", "")
    ftipo = request.args.get("ftipo","")
    festado = request.args.get("festado","activos")
    fi, pi = filtro_igl()
    where_extra = ""
    params_extra = []
    if ftipo:
        where_extra += " AND tipo_hermano=%s"
        params_extra.append(ftipo)
    if festado == "activos":
        where_extra += " AND activo=true"
    elif festado == "inactivos":
        where_extra += " AND activo=false"
    # todos => no filtro activo
    base_where = "(nombre_completo ILIKE %s OR codigo ILIKE %s)"
    # incluir rut en busqueda solo si no hay ftipo para compatibilidad
    if not ftipo:
        base_where = "(nombre_completo ILIKE %s OR codigo ILIKE %s OR rut ILIKE %s)"
        base_params = (f"%{texto}%", f"%{texto}%", f"%{texto}%")
    else:
        base_params = (f"%{texto}%", f"%{texto}%")
    rows = q(f"SELECT id, codigo, nombre_completo, comuna, region, correo, telefono, tipo_hermano, activo, motivo_salida, fecha_baja FROM v_hermanos_completo WHERE {base_where} {where_extra} {fi} ORDER BY nombre_completo LIMIT 1000",
             base_params + tuple(params_extra) + tuple(pi))
    if exp == "excel": return to_excel(rows, ["codigo", "nombre_completo", "tipo_hermano", "comuna", "region", "correo", "telefono", "activo", "motivo_salida", "fecha_baja"], "hermanos.xlsx")
    if exp == "pdf": return to_pdf("Hermanos MCCI", rows, ["codigo", "nombre_completo", "tipo_hermano", "comuna", "region"], "hermanos.pdf")
    tipos_hermano = get_catalogo("catalogo_tipo_hermano")
    return render_template("hermanos.html", rows=rows, texto=texto, ftipo=ftipo, festado=festado, tipos_hermano=tipos_hermano)

# ---------- CÓDIGO DE HERMANOS POR IGLESIA (prefijo + secuencia propia) ----------
def _prefijo_sugerido(nombre, iglesia_id=None, excluir_id=None):
    """Prefijo por defecto a partir del nombre: 'IGLESIA NORTE' -> 'NOR'.
    Si colisiona con otra iglesia se alarga hasta 6 letras; en último
    caso usa el id (único por definición)."""
    base = (nombre or "IGLESIA").upper().strip()
    for p in ("IGLESIA", "GRUPO", "CASA", "COMUNIDAD", "RED"):
        if base.startswith(p + " "):
            base = base[len(p) + 1:]
            break
    base = "".join(ch for ch in base if ch.isalpha()) or "IGL"
    exi = lambda cand: q("SELECT 1 FROM iglesias WHERE UPPER(prefijo_codigo)=%s AND id!=%s LIMIT 1",
                         (cand, excluir_id or 0), one=True)
    cand, n = base[:3], 3
    while n < 6 and exi(cand):
        n += 1
        cand = base[:n]
    if exi(cand):
        cand = ("Z%d" % (iglesia_id or 0))[:6]
    return cand

def _codigo_hermano(iglesia_id):
    """Código de hermano único en todo el sistema: PREFIJO DE LA IGLESIA + su
    secuencia propia (MATRIZ->MCCI0886, hija->NOR0001, grupo->GRA0001).
    Cada sede numera desde 0001 y no se pisa con las demás."""
    iglesia_id = int(iglesia_id or 1)
    row = q("SELECT nombre, COALESCE(NULLIF(UPPER(prefijo_codigo),''), '') AS p FROM iglesias WHERE id=%s",
            (iglesia_id,), one=True)
    pref = (row["p"] if row else "")
    if not pref:
        pref = _prefijo_sugerido(row["nombre"] if row else "IGLESIA", iglesia_id)
        q("UPDATE iglesias SET prefijo_codigo=%s WHERE id=%s AND prefijo_codigo IS NULL",
          (pref, iglesia_id), commit=True)
    seq = "seq_codigo_igl_%d" % iglesia_id
    q("CREATE SEQUENCE IF NOT EXISTS " + seq, commit=True)
    return q("SELECT %s || LPAD(nextval(%s)::text,4,'0') AS c", (pref, seq), one=True)["c"]

@app.get("/hermanos/nuevo")
@login_req
@permiso_req("hermanos", "crear")
def hermano_nuevo():
    com = q("SELECT c.id,c.nombre,r.nombre_corto AS region FROM comunas c JOIN regiones r ON r.id=c.region_id ORDER BY c.nombre")
    fi, pi = filtro_igl()
    her = q(f"SELECT id,nombre_completo FROM hermanos WHERE activo{fi} ORDER BY 2", tuple(pi))
    sexos = get_catalogo("catalogo_sexo")
    estados_civil = get_catalogo("catalogo_estado_civil")
    tipos_hermano = get_catalogo("catalogo_tipo_hermano")
    return render_template("hermano_form.html", h=None, comunas=com, hermanos=her, sexos=sexos, estados_civil=estados_civil, tipos_hermano=tipos_hermano)

@app.post("/hermanos/nuevo")
@login_req
@permiso_req("hermanos", "crear")
def hermano_crear():
    f = request.form
    igl, _ = iglesia_actual()
    igl = igl if igl is not None else 1
    inv = f.get("invitado") or None
    if inv and not exigir_igl("hermanos", inv):
        flash("El hermano que invita no es de tu iglesia", "error"); return redirect(url_for("hermano_nuevo"))
    try:
        codigo = _codigo_hermano(igl)
        tipo_h = f.get("tipo_hermano") or 'NINGUNO'
        q("""INSERT INTO hermanos(codigo,tipo_hermano,rut,nombres,apellido_paterno,apellido_materno,fecha_nacimiento,fecha_aniversario,
            correo,telefono,direccion,comuna_id,sexo,estado_civil,invitado_por_id,invitado_por_texto,iglesia_id) VALUES(%s,%s,NULL,%s,%s,%s,NULLIF(%s,'')::date,NULLIF(%s,'')::date,
            NULLIF(%s,''),%s,%s,NULLIF(%s,'')::int,%s,%s,NULLIF(%s,'')::int,NULLIF(%s,''),%s)""",
          (codigo, tipo_h, f["nombres"].strip(), f["paterno"].strip(), f.get("materno") or None,
           f.get("fnac") or None, f.get("fani") or None, f.get("correo") or None,
           f.get("telefono"), f.get("direccion"), f.get("comuna") or None, f.get("sexo") or None,
           f.get("ecivil") or None, inv, f.get("invtexto") or None, igl), commit=True)
        # Si es LIDER, crear registro en lideres si no existe (de ahí salen líderes de célula)
        if tipo_h == 'LIDER':
            hid = q("SELECT id FROM hermanos WHERE codigo=%s", (codigo,), one=True)["id"]
            if not q("SELECT id FROM lideres WHERE hermano_id=%s", (hid,), one=True):
                q("INSERT INTO lideres(hermano_id, red, estado_celula, info_enviada, rol_lider, tipo_12, iglesia_id) VALUES(%s,'MIXTO','EN FORMACION','EN CONSULTA','NINGUNO','NINGUNO',%s)", (hid, igl), commit=True)
        log_audit("CREAR","hermanos", f"codigo={codigo} {f.get('nombres','')} tipo={tipo_h}")
        flash(f"Hermano creado con código {codigo}", "ok")
    except Exception as e: flash(f"Error: {e}", "error"); return redirect(url_for("hermano_nuevo"))
    return redirect(url_for("hermanos"))

@app.get("/hermanos/<int:hid>/editar")
@login_req
@permiso_req("hermanos", "editar")
def hermano_editar(hid):
    if not exigir_igl("hermanos", hid):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("hermanos"))
    h = q("SELECT * FROM hermanos WHERE id=%s", (hid,), one=True)
    com = q("SELECT c.id,c.nombre FROM comunas c ORDER BY c.nombre")
    fi, pi = filtro_igl()
    her = q(f"SELECT id,nombre_completo FROM hermanos WHERE activo AND id!=%s{fi} ORDER BY 2", tuple([hid]) + tuple(pi))
    sexos = get_catalogo("catalogo_sexo")
    estados_civil = get_catalogo("catalogo_estado_civil")
    tipos_hermano = get_catalogo("catalogo_tipo_hermano")
    return render_template("hermano_form.html", h=h, comunas=com, hermanos=her, sexos=sexos, estados_civil=estados_civil, tipos_hermano=tipos_hermano)

@app.post("/hermanos/<int:hid>/editar")
@login_req
@permiso_req("hermanos", "editar")
def hermano_update(hid):
    f = request.form
    if not exigir_igl("hermanos", hid):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("hermanos"))
    try:
        act = q("SELECT * FROM hermanos WHERE id=%s", (hid,), one=True)
        if not act: flash("No encontrado", "error"); return redirect(url_for("hermanos"))
        # Campos de texto: si vienen vacíos y antes tenían valor, permitir vaciar algunos
        # Para FK nullable (comuna, invitado, sexo, ecivil): "" → NULL
        sexo_val = f.get("sexo")
        if sexo_val == "": sexo_val = None
        ecivil_val = f.get("ecivil")
        if ecivil_val == "": ecivil_val = None
        comuna_val = f.get("comuna")
        comuna_val = int(comuna_val) if comuna_val and comuna_val.strip() else None
        invitado_val = f.get("invitado")
        invitado_val = int(invitado_val) if invitado_val and invitado_val.strip() else None
        if invitado_val and not exigir_igl("hermanos", invitado_val):
            flash("El hermano que invita no es de tu iglesia", "error"); return redirect(url_for("hermanos"))
        # Materno y correo pueden vaciarse ("None" string viene de template viejo)
        def clean(v):
            if v is None or v == "" or v == "None" or v == "null": return None
            return v.strip() if isinstance(v, str) else v
        materno_val = clean(f.get("materno"))
        if materno_val is None and f.get("materno") is None: materno_val = act["apellido_materno"]
        correo_val = clean(f.get("correo"))
        if correo_val is None and not f.get("correo"): correo_val = act["correo"] if act["correo"] not in ("None",) else None
        # Telefono y direccion: normalizar "None" -> None para update_parcial
        tel_raw = clean(f.get("telefono"))
        dir_raw = clean(f.get("direccion"))
        invtexto_val = clean(f.get("invtexto"))
        tipo_hermano_val = f.get("tipo_hermano") or act.get("tipo_hermano") or 'NINGUNO'
        if tipo_hermano_val in ("", "None", "null"): tipo_hermano_val = 'NINGUNO'
        # Resto: si viene vacío conservar anterior, excepto los que permiten NULL
        update_parcial("hermanos", hid, {
            "nombres": f.get("nombres") or act["nombres"],
            "apellido_paterno": f.get("paterno") or act["apellido_paterno"],
        })
        # Actualizar campos que pueden ser NULL por separado
        q("UPDATE hermanos SET apellido_materno=%s, correo=%s, telefono=%s, direccion=%s, comuna_id=%s, sexo=%s, estado_civil=%s, invitado_por_id=%s, invitado_por_texto=%s, tipo_hermano=%s WHERE id=%s",
          (materno_val, correo_val, tel_raw, dir_raw, comuna_val, sexo_val, ecivil_val, invitado_val, invtexto_val, tipo_hermano_val, hid), commit=True)
        # Sincronización LÍDER ↔ DISCÍPULO: de aquí salen los líderes de célula
        if tipo_hermano_val == 'LIDER':
            if not q("SELECT id FROM lideres WHERE hermano_id=%s", (hid,), one=True):
                q("INSERT INTO lideres(hermano_id, red, estado_celula, info_enviada, rol_lider, tipo_12, iglesia_id) VALUES(%s,'MIXTO','EN FORMACION','EN CONSULTA','NINGUNO','NINGUNO',%s)", (hid, act["iglesia_id"]), commit=True)
            else:
                q("UPDATE lideres SET activo=true WHERE hermano_id=%s", (hid,), commit=True)
        elif act.get("tipo_hermano") == 'LIDER' and tipo_hermano_val != 'LIDER':
            # Bajó de LIDER a DISCIPULO/NINGUNO → desactivar liderazgo pero mantener historial
            q("UPDATE lideres SET activo=false WHERE hermano_id=%s", (hid,), commit=True)
        log_audit("EDITAR","hermanos", f"id={hid} tipo={tipo_hermano_val}")
        flash("Actualizado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("hermanos"))

@app.post("/hermanos/<int:hid>/toggle")
@login_req
@permiso_req("hermanos", "editar")
def hermano_toggle(hid):
    if not exigir_igl("hermanos", hid):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("hermanos"))
    h = q("SELECT activo FROM hermanos WHERE id=%s", (hid,), one=True)
    if not h:
        flash("No encontrado", "error"); return redirect(url_for("hermanos"))
    if h["activo"]:
        # Pasa a inactivo: pedir motivo
        motivo = (request.form.get("motivo") or "").strip() or "Salida registrada"
        q("UPDATE hermanos SET activo=false, fecha_baja=CURRENT_DATE, motivo_salida=%s WHERE id=%s", (motivo, hid), commit=True)
        # Cerrar discipulado y desactivar liderazgo pero conservar código e historial
        q("UPDATE discipulado SET fecha_fin=CURRENT_DATE WHERE hermano_id=%s AND fecha_fin IS NULL", (hid,), commit=True)
        q("UPDATE lideres SET activo=false WHERE hermano_id=%s", (hid,), commit=True)
        log_audit("EDITAR","hermanos", f"baja id={hid} motivo={motivo}")
        flash(f"Hermano dado de baja (inactivo) con código conservado. Motivo: {motivo}", "ok")
    else:
        # Reactivar
        q("UPDATE hermanos SET activo=true, fecha_baja=NULL, motivo_salida=NULL WHERE id=%s", (hid,), commit=True)
        log_audit("EDITAR","hermanos", f"reactiva id={hid}")
        flash("Hermano reactivado", "ok")
    return redirect(url_for("hermanos"))

@app.post("/hermanos/<int:hid>/eliminar")
@login_req
@permiso_req("hermanos", "eliminar")
def hermano_eliminar(hid):
    if not exigir_igl("hermanos", hid):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("hermanos"))
    q("DELETE FROM hermanos WHERE id=%s", (hid,), commit=True)
    log_audit("ELIMINAR","hermanos", f"id={hid}")
    flash("Hermano eliminado", "ok")
    return redirect(url_for("hermanos"))

# ---------- LIDERES ----------
@app.get("/lideres")
@login_req
def lideres():
    exp = request.args.get("export", "")
    f12 = request.args.get("f12", "")
    texto = request.args.get("q", "")
    fpadre = request.args.get("fpadre", "")
    where = "WHERE (%s='' OR tipo_12=%s)"
    params = [f12, f12 or None]
    fi, pi = filtro_igl("l")
    where += fi
    params += pi
    if fpadre:
        if not exigir_igl("lideres", fpadre):
            flash("Ese lider padre no es de tu iglesia", "error"); return redirect(url_for("lideres"))
        where += " AND l.lider_padre_id=%s"
        params.append(int(fpadre))
    if texto:
        where += " AND l.lider ILIKE %s"
        params.append(f"%{texto}%")
    rows = q(f"SELECT * FROM v_lideres_conteo l {where} ORDER BY discipulos_activos DESC", tuple(params))
    # Timoteo por líder: opciones a cargo (discípulos activos + sub-líderes) y el guardado actualmente.
    tim_opts, tim_sel, tim_nombre = {}, {}, {}
    if rows:
        lids = tuple(r["lider_id"] for r in rows)
        ph = ",".join(["%s"] * len(lids))
        for r in q(f"""SELECT lid, id, nombre FROM (
            SELECT d.lider_id AS lid, d.hermano_id AS id, h.nombre_completo AS nombre
              FROM discipulado d JOIN hermanos h ON h.id=d.hermano_id
             WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo AND d.lider_id IN ({ph})
            UNION
            SELECT l2.lider_padre_id AS lid, l2.hermano_id AS id, h.nombre_completo AS nombre
              FROM lideres l2 JOIN hermanos h ON h.id=l2.hermano_id
             WHERE l2.activo AND l2.lider_padre_id IN ({ph})
        ) t ORDER BY nombre""", lids + lids):
            tim_opts.setdefault(r["lid"], []).append((r["id"], r["nombre"]))
        for r in q(f"""SELECT l.id AS lid, l.timoteo_id AS tid, h.nombre_completo AS tnom
                       FROM lideres l LEFT JOIN hermanos h ON h.id=l.timoteo_id
                      WHERE l.id IN ({ph})""", lids):
            tim_sel[r["lid"]] = r["tid"]
            tim_nombre[r["lid"]] = r["tnom"] or ""
        for r in rows:
            r["timoteo"] = tim_nombre.get(r["lider_id"], "")
    if exp == "excel": return to_excel(rows, ["codigo", "lider", "rol_lider", "pastor_nombre", "timoteo", "red", "tipo_12", "estado_celula", "info_enviada", "cantidad_celulas", "cantidad_macrocelulas", "discipulos_activos"], "lideres.xlsx")
    if exp == "pdf": return to_pdf("Lideres MCCI", rows, ["codigo", "lider", "rol_lider", "pastor_nombre", "timoteo", "red", "tipo_12", "cantidad_celulas", "cantidad_macrocelulas", "discipulos_activos"], "lideres.pdf")
    if fi:
        todos = q("""SELECT l.id, h.nombre_completo AS nombre,
            (SELECT COUNT(*) FROM lideres l2 WHERE l2.lider_padre_id=l.id AND l2.activo AND l2.iglesia_id=%s) AS hijos
            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo AND l.iglesia_id=%s ORDER BY h.nombre_completo""", (pi[0], pi[0]))
    else:
        todos = q("""SELECT l.id, h.nombre_completo AS nombre,
            (SELECT COUNT(*) FROM lideres l2 WHERE l2.lider_padre_id=l.id AND l2.activo) AS hijos
            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY h.nombre_completo""")
    catalogos_l = {
        "redes": get_catalogo("catalogo_redes"),
        "tipo12": get_catalogo("catalogo_tipo12"),
        "roles_lider": get_catalogo("catalogo_roles_lider"),
        "estado_celula": get_catalogo("catalogo_estado_celula"),
        "info_enviada": get_catalogo("catalogo_info_enviada"),
    }
    # Dict rol_nombre → icono para display
    rol_iconos = {r["nombre"]: r.get("icono","—") for r in catalogos_l["roles_lider"]}
    return render_template("lideres.html", rows=rows, f12=f12, texto=texto,
        fpadre=int(fpadre) if fpadre else None, todos_lideres=todos, catalogos_l=catalogos_l, rol_iconos=rol_iconos,
        tim_opts=tim_opts, tim_sel=tim_sel)

@app.get("/lideres/nuevo")
@login_req
@permiso_req("lideres", "crear")
def lider_nuevo():
    fi, pi = filtro_igl("h")
    libres = q(f"""SELECT h.id,h.nombre_completo FROM hermanos h LEFT JOIN lideres l ON l.hermano_id=h.id
                  WHERE l.id IS NULL AND h.activo{fi} ORDER BY h.nombre_completo""", tuple(pi))
    fl, pl = filtro_igl("l")
    todos = q(f"""SELECT l.id, h.nombre_completo AS nombre, l.tipo_12 FROM lideres l
                 JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY h.nombre_completo""", tuple(pl))
    catalogos_l = {
        "redes": get_catalogo("catalogo_redes"),
        "tipo12": get_catalogo("catalogo_tipo12"),
        "roles_lider": get_catalogo("catalogo_roles_lider"),
        "estado_celula": get_catalogo("catalogo_estado_celula"),
        "info_enviada": get_catalogo("catalogo_info_enviada"),
    }
    return render_template("lider_form.html", libres=libres, todos_lideres=todos, l=None, catalogos_l=catalogos_l)

# Roles que el sistema trata como pastorales. IMPORTANTE: incluye el femenino
# PASTORA. Antes solo 'PASTOR' activaba es_pastor, así que las 10 pastoras
# tenían es_pastor=false y NO aparecían en el reporte "Informe Célula · Por Pastor".
_ROLES_PASTORALES = ("PASTOR", "PASTORA")

@app.post("/lideres/nuevo")
@login_req
@permiso_req("lideres", "crear")
def lider_crear():
    f = request.form
    igl, _ = iglesia_actual()
    igl = igl if igl is not None else 1
    hid = f.get("hermano")
    if not hid or not exigir_igl("hermanos", hid):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("lideres"))
    padre = f.get("padre") or None
    if padre and not exigir_igl("lideres", padre):
        flash("El lider padre no es de tu iglesia", "error"); return redirect(url_for("lideres"))
    try:
        rol = f.get("rol") or "NINGUNO"
        q("INSERT INTO lideres(hermano_id,red,estado_celula,info_enviada,cantidad_celulas,cantidad_macrocelulas,observacion,es_pastor,rol_lider,tipo_12,lider_padre_id,iglesia_id) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
          (hid, f["red"], f["estado"], f["info"], int(f.get("celulas") or 1), int(f.get("macro") or 0), f.get("obs"),
           rol in _ROLES_PASTORALES, rol, f.get("tipo12", "NINGUNO"), padre, igl), commit=True)
        log_audit("CREAR","lideres", f"hermano={f.get('hermano')} rol={rol}")
        flash("Líder creado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("lideres"))

@app.post("/lideres/<int:lid>/estado")
@login_req
@permiso_req("lideres", "editar")
def lider_estado(lid):
    f = request.form
    if not exigir_igl("lideres", lid):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("lideres"))
    act = q("SELECT * FROM lideres WHERE id=%s", (lid,), one=True)
    if not act: flash("No encontrado", "error"); return redirect(url_for("lideres"))
    rol = f.get("rol") or "NINGUNO"
    # Preparar valores: padre puede ser "" → NULL (quitar padre)
    padre_raw = f.get("padre", None)
    # Si el campo padre vino en el form, respetarlo ("" = quitar, valor = asignar)
    # Si no vino, conservar el anterior
    if "padre" in f:
        padre_val = int(padre_raw) if padre_raw and padre_raw.strip() else None
        if padre_val and not exigir_igl("lideres", padre_val):
            flash("El lider padre no es de tu iglesia", "error"); return redirect(url_for("lideres"))
        # Update_parcial no soporta NULL directo, usamos permitir_null
        update_parcial("lideres", lid, {
            "estado_celula": f.get("estado") or act["estado_celula"],
            "info_enviada": f.get("info") or act["info_enviada"],
            "cantidad_celulas": int(f.get("celulas") or act["cantidad_celulas"]),
            "cantidad_macrocelulas": int(f.get("macro") or act["cantidad_macrocelulas"] or 0),
            "red": f.get("red") or act["red"],
            "rol_lider": rol,
            "es_pastor": rol in _ROLES_PASTORALES,
            "tipo_12": f.get("tipo12") or act["tipo_12"],
        })
        # Actualizar padre por separado para permitir NULL
        q("UPDATE lideres SET lider_padre_id=%s WHERE id=%s", (padre_val, lid), commit=True)
    else:
        update_parcial("lideres", lid, {
            "estado_celula": f.get("estado") or act["estado_celula"],
            "info_enviada": f.get("info") or act["info_enviada"],
            "cantidad_celulas": int(f.get("celulas") or act["cantidad_celulas"]),
            "cantidad_macrocelulas": int(f.get("macro") or act["cantidad_macrocelulas"] or 0),
            "red": f.get("red") or act["red"],
            "rol_lider": rol,
            "es_pastor": rol in _ROLES_PASTORALES,
            "tipo_12": f.get("tipo12") or act["tipo_12"],
        })
    log_audit("EDITAR","lideres", f"id={lid} rol={rol}")
    flash("Líder actualizado", "ok"); return redirect(url_for("lideres"))

@app.post("/lideres/<int:lid>/timoteo")
@login_req
@permiso_req("lideres", "editar")
def lider_timoteo(lid):
    if not exigir_igl("lideres", lid):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("lideres"))
    t = (request.form.get("timoteo") or "").strip()
    tid = int(t) if t.isdigit() else None
    # Seguridad: el timoteo elegido debe estar realmente a cargo de este líder.
    if tid is not None and tid not in [r["id"] for r in q(_SQL_TIMOTEOS, (lid, lid))]:
        flash("Ese timoteo no está a cargo del líder", "error"); return redirect(url_for("lideres"))
    q("UPDATE lideres SET timoteo_id=%s, updated_at=now() WHERE id=%s", (tid, lid), commit=True)
    log_audit("EDITAR", "lideres", f"id={lid} timoteo_id={tid}")
    flash("Timoteo actualizado", "ok"); return redirect(url_for("lideres"))

@app.post("/lideres/<int:lid>/eliminar")
@login_req
@permiso_req("lideres", "eliminar")
def lider_eliminar(lid):
    if not exigir_igl("lideres", lid):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("lideres"))
    try:
        row = q("SELECT hermano_id FROM lideres WHERE id=%s", (lid,), one=True)
        if not row:
            flash("Líder no encontrado", "error")
            return redirect(url_for("lideres"))
        hermano_id = row["hermano_id"]
        # Quitar referencia como padre
        q("UPDATE lideres SET lider_padre_id=NULL WHERE lider_padre_id=%s", (lid,), commit=True)
        # Eliminar discipulados de este líder (libera a los discípulos para reasignar)
        q("DELETE FROM discipulado WHERE lider_id=%s", (lid,), commit=True)
        # Quitar referencia en informes de célula
        try: q("UPDATE celulas_informe SET lider_id=NULL WHERE lider_id=%s", (lid,), commit=True)
        except: pass
        try: q("UPDATE celulas_informe SET pastor_id=NULL WHERE pastor_id=%s", (lid,), commit=True)
        except: pass
        # Borrar lider
        q("DELETE FROM lideres WHERE id=%s", (lid,), commit=True)
        # Actualizar tipo del hermano si era LIDER
        q("UPDATE hermanos SET tipo_hermano='DISCIPULO' WHERE id=%s AND tipo_hermano='LIDER'", (hermano_id,), commit=True)
        log_audit("ELIMINAR","lideres", f"id={lid} hermano={hermano_id}")
        flash("Líder eliminado y discípulos liberados", "ok")
    except Exception as e:
        flash(f"No se pudo eliminar líder (¿tiene dependencias?): {e}", "error")
    return redirect(url_for("lideres"))

# ---------- DISCIPULADO ----------
@app.get("/discipulado")
@login_req
def discipulado():
    exp = request.args.get("export", "")
    texto = request.args.get("q", "")
    f12 = request.args.get("f12", "")
    fi, pi = filtro_igl("l")
    base = f"""FROM discipulado d JOIN hermanos h ON h.id=d.hermano_id JOIN lideres l ON l.id=d.lider_id
        JOIN hermanos hl ON hl.id=l.hermano_id LEFT JOIN periodicidad_celula p ON p.id=d.periodicidad_id
        WHERE (%s='' OR l.tipo_12=%s) AND (h.nombre_completo ILIKE %s OR hl.nombre_completo ILIKE %s){fi}"""
    params = (f12, f12 or None, f"%{texto}%", f"%{texto}%") + tuple(pi)
    rows = q(f"""SELECT h.nombre_completo AS hermano,hl.nombre_completo AS lider,p.nombre AS periodicidad,
        l.tipo_12,d.es_discipulo_activo,d.fecha_asignacion,d.fecha_fin {base}
        ORDER BY d.fecha_fin NULLS FIRST, d.fecha_asignacion DESC LIMIT 1000""", params)
    if exp == "excel": return to_excel(rows, ["hermano", "lider", "periodicidad", "tipo_12", "es_discipulo_activo", "fecha_asignacion", "fecha_fin"], "discipulado.xlsx")
    if exp == "pdf": return to_pdf("Discipulado MCCI", rows, ["hermano", "lider", "tipo_12"], "discipulado.pdf")
    full = q(f"""SELECT d.id,h.nombre_completo AS hermano,hl.nombre_completo AS lider,p.nombre AS periodicidad,
        l.tipo_12,d.es_discipulo_activo,d.fecha_asignacion,d.fecha_fin {base}
        ORDER BY d.fecha_fin NULLS FIRST, d.fecha_asignacion DESC LIMIT 1000""", params)
    return render_template("discipulado.html", rows=full, texto=texto, f12=f12,
        tipos12=get_catalogo("catalogo_tipo12"))

@app.get("/discipulado/nuevo")
@login_req
@permiso_req("discipulado", "crear")
def disc_nuevo():
    fi, pi = filtro_igl()
    fl, pl = filtro_igl("l")
    return render_template("disc_form.html",
        hermanos=q(f"SELECT id,nombre_completo FROM hermanos WHERE activo{fi} ORDER BY 2", tuple(pi)),
        lideres=q(f"SELECT l.id,h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY 2", tuple(pl)),
        per=q("SELECT * FROM periodicidad_celula ORDER BY 1"))

@app.post("/discipulado/nuevo")
@login_req
@permiso_req("discipulado", "crear")
def disc_crear():
    f = request.form
    if not exigir_igl("hermanos", f.get("hermano") or 0) or not exigir_igl("lideres", f.get("lider") or 0):
        flash("Hermano o lider no son de tu iglesia", "error"); return redirect(url_for("discipulado"))
    try:
        q("UPDATE discipulado SET fecha_fin=CURRENT_DATE WHERE hermano_id=%s AND fecha_fin IS NULL", (f["hermano"],), commit=True)
        q("INSERT INTO discipulado(hermano_id,lider_id,es_discipulo_activo,periodicidad_id,observacion) VALUES(%s,%s,%s,%s,%s)",
          (f["hermano"], f["lider"], f.get("activo") == "on", f.get("per") or None, f.get("obs")), commit=True)
        log_audit("CREAR","discipulado", f"h={f.get('hermano')} l={f.get('lider')}")
        flash("Asignación creada", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("discipulado"))

@app.post("/discipulado/<int:did>/cerrar")
@login_req
@permiso_req("discipulado", "editar")
def disc_cerrar(did):
    if not exigir_igl("discipulado", did):
        flash("Esa asignacion no es de tu iglesia", "error"); return redirect(url_for("discipulado"))
    q("UPDATE discipulado SET fecha_fin=CURRENT_DATE WHERE id=%s", (did,), commit=True)
    return redirect(url_for("discipulado"))

@app.get("/discipulado/<int:did>/editar")
@login_req
@permiso_req("discipulado", "editar")
def disc_editar(did):
    if not exigir_igl("discipulado", did):
        flash("Esa asignacion no es de tu iglesia", "error"); return redirect(url_for("discipulado"))
    d = q("SELECT * FROM discipulado WHERE id=%s", (did,))
    if not d: flash("No encontrado", "error"); return redirect(url_for("discipulado"))
    fi, pi = filtro_igl()
    fl, pl = filtro_igl("l")
    return render_template("disc_editar.html", d=d[0],
        hermanos=q(f"SELECT id,nombre_completo FROM hermanos WHERE activo{fi} ORDER BY 2", tuple(pi)),
        lideres=q(f"SELECT l.id,h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY 2", tuple(pl)),
        per=q("SELECT * FROM periodicidad_celula ORDER BY 1"))

@app.post("/discipulado/<int:did>/editar")
@login_req
@permiso_req("discipulado", "editar")
def disc_editar_guardar(did):
    f = request.form
    if not exigir_igl("discipulado", did):
        flash("Esa asignacion no es de tu iglesia", "error"); return redirect(url_for("discipulado"))
    try:
        act = q("SELECT * FROM discipulado WHERE id=%s", (did,), one=True)
        if not act: flash("No encontrado", "error"); return redirect(url_for("discipulado"))
        nh = f.get("hermano") or act["hermano_id"]
        nl = f.get("lider") or act["lider_id"]
        if not exigir_igl("hermanos", nh) or not exigir_igl("lideres", nl):
            flash("Hermano o lider no son de tu iglesia", "error"); return redirect(url_for("discipulado"))
        update_parcial("discipulado", did, {
            "hermano_id": f.get("hermano") or act["hermano_id"],
            "lider_id": f.get("lider") or act["lider_id"],
            "es_discipulo_activo": f.get("activo") == "on" if f.get("activo") is not None else act["es_discipulo_activo"],
            "periodicidad_id": f.get("per") or act["periodicidad_id"],
            "observacion": f.get("obs") if f.get("obs") is not None else act["observacion"],
        })
        log_audit("EDITAR","discipulado", f"id={did}")
        flash("Asignación modificada", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("discipulado"))

@app.post("/discipulado/<int:did>/eliminar")
@login_req
@permiso_req("discipulado", "eliminar")
def disc_eliminar(did):
    if not exigir_igl("discipulado", did):
        flash("Esa asignacion no es de tu iglesia", "error"); return redirect(url_for("discipulado"))
    q("DELETE FROM discipulado WHERE id=%s", (did,), commit=True)
    log_audit("ELIMINAR","discipulado", f"id={did}")
    flash("Asignación eliminada", "ok")
    return redirect(url_for("discipulado"))

# ---------- ENCUENTROS ----------
@app.get("/encuentros")
@login_req
def encuentros():
    exp = request.args.get("export", "")
    texto = request.args.get("q", "")
    fest = request.args.get("fest", "")
    fh, ph = filtro_igl("h")
    base = f"FROM encuentro_participacion e JOIN hermanos h ON h.id=e.hermano_id JOIN tipos_encuentro t ON t.id=e.tipo_encuentro_id WHERE (%s='' OR e.estado=%s) AND h.nombre_completo ILIKE %s{fh}"
    params = (fest, fest or None, f"%{texto}%") + tuple(ph)
    if exp in ("excel", "pdf"):
        rows = q(f"SELECT h.nombre_completo,t.nombre AS etapa,e.estado,e.fecha_evento {base} ORDER BY 1 LIMIT 500", params)
        if exp == "excel": return to_excel(rows, ["nombre_completo", "etapa", "estado", "fecha_evento"], "encuentros.xlsx")
        return to_pdf("Encuentros MCCI", rows, ["nombre_completo", "etapa", "estado"], "encuentros.pdf")
    rows = q(f"SELECT e.id,h.nombre_completo,t.nombre AS etapa,e.estado,e.fecha_evento {base} ORDER BY e.fecha_evento DESC LIMIT 300", params)
    if ph:
        ruta = q("SELECT * FROM v_ruta_espiritual WHERE nombre_completo ILIKE %s AND iglesia_id=%s ORDER BY nombre_completo LIMIT 200", (f"%{texto}%", ph[0]))
    else:
        ruta = q("SELECT * FROM v_ruta_espiritual WHERE nombre_completo ILIKE %s ORDER BY nombre_completo LIMIT 200", (f"%{texto}%",))
    return render_template("encuentros.html", rows=rows, ruta=ruta, texto=texto, fest=fest)

@app.get("/encuentros/nuevo")
@login_req
@permiso_req("encuentros", "crear")
def enc_nuevo():
    from datetime import date
    fi, pi = filtro_igl()
    return render_template("enc_form.html",
        hermanos=q(f"SELECT id,nombre_completo FROM hermanos WHERE activo{fi} ORDER BY 2", tuple(pi)),
        tipos=q("SELECT * FROM tipos_encuentro ORDER BY orden"),
        fecha_hoy=date.today().isoformat())

@app.post("/encuentros/nuevo")
@login_req
@permiso_req("encuentros", "crear")
def enc_crear():
    from datetime import date
    f = request.form
    if not exigir_igl("hermanos", f.get("hermano") or 0):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("encuentros"))
    try:
        fecha = f.get("fecha") or date.today().isoformat()
        q("""INSERT INTO encuentro_participacion(hermano_id,tipo_encuentro_id,estado,fecha_evento) VALUES(%s,%s,%s,%s)
             ON CONFLICT (hermano_id,tipo_encuentro_id) DO UPDATE SET estado=EXCLUDED.estado,fecha_evento=EXCLUDED.fecha_evento""",
          (f["hermano"], f["tipo"], f["estado"], fecha), commit=True)
        log_audit("CREAR","encuentros", f"h={f.get('hermano')} tipo={f.get('tipo')}")
        flash("Encuentro registrado", "ok")
    except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("encuentros"))

@app.get("/encuentros/masivo")
@login_req
@permiso_req("encuentros", "crear")
def enc_masivo():
    from datetime import date
    sel = request.args.get("lider", "")
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT l.id, h.nombre_completo AS nombre, COUNT(d.id) AS discipulos
        FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        JOIN discipulado d ON d.lider_id=l.id AND d.fecha_fin IS NULL AND d.es_discipulo_activo=true
        WHERE l.activo{fl} GROUP BY l.id, h.nombre_completo ORDER BY h.nombre_completo""", tuple(pl))
    discipulos = []
    if sel:
        if not exigir_igl("lideres", sel):
            flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("enc_masivo"))
        discipulos = q("""SELECT h.id, h.nombre_completo,
            (SELECT t.nombre FROM encuentro_participacion ep JOIN tipos_encuentro t ON t.id=ep.tipo_encuentro_id
             WHERE ep.hermano_id=h.id ORDER BY ep.fecha_evento DESC LIMIT 1) AS ya_tiene,
            (SELECT ep.estado FROM encuentro_participacion ep WHERE ep.hermano_id=h.id ORDER BY ep.fecha_evento DESC LIMIT 1) AS estado_actual
            FROM discipulado d JOIN hermanos h ON h.id=d.hermano_id
            WHERE d.lider_id=%s AND d.fecha_fin IS NULL AND d.es_discipulo_activo=true ORDER BY h.nombre_completo""", (sel,))
    return render_template("enc_masivo.html", lideres=lideres, discipulos=discipulos,
        sel_lider=int(sel) if sel else None, tipos=q("SELECT * FROM tipos_encuentro ORDER BY orden"),
        fecha_hoy=date.today().isoformat())

@app.post("/encuentros/masivo")
@login_req
@permiso_req("encuentros", "crear")
def enc_masivo_guardar():
    from datetime import date
    f = request.form
    discs = f.getlist("disc")
    tipo = f["tipo"]
    estado = f["estado"]
    fecha = f.get("fecha") or date.today().isoformat()
    if not discs:
        flash("Seleccioná al menos un discípulo", "error")
        return redirect(url_for("enc_masivo", lider=f.get("lider","")))
    if f.get("lider") and not exigir_igl("lideres", f.get("lider")):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("enc_masivo"))
    count = 0
    for hid in discs:
        if not exigir_igl("hermanos", hid): continue
        try:
            q("""INSERT INTO encuentro_participacion(hermano_id,tipo_encuentro_id,estado,fecha_evento) VALUES(%s,%s,%s,%s)
                 ON CONFLICT (hermano_id,tipo_encuentro_id) DO UPDATE SET estado=EXCLUDED.estado,fecha_evento=EXCLUDED.fecha_evento""",
              (hid, tipo, estado, fecha), commit=True)
            count += 1
        except: pass
    log_audit("CREAR","encuentros", f"masivo {count} tipo={tipo}")
    flash(f"Encuentro registrado para {count} discípulos", "ok")
    return redirect(url_for("encuentros"))

@app.get("/encuentros/guia")
@login_req
@permiso_req("encuentros", "leer")
def enc_guia():
    from datetime import date
    sel = request.args.get("lider", "")
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT l.id, h.nombre_completo AS nombre, COUNT(d.id) AS discipulos
        FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        JOIN discipulado d ON d.lider_id=l.id AND d.fecha_fin IS NULL AND d.es_discipulo_activo=true
        WHERE l.activo{fl} GROUP BY l.id, h.nombre_completo ORDER BY h.nombre_completo""", tuple(pl))
    tipos = q("SELECT * FROM tipos_encuentro ORDER BY orden")
    discipulos = []
    stats = {}
    if sel:
        if not exigir_igl("lideres", sel):
            flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("enc_guia"))
        discipulos = q("""SELECT h.id, h.nombre_completo FROM discipulado d
            JOIN hermanos h ON h.id=d.hermano_id
            WHERE d.lider_id=%s AND d.fecha_fin IS NULL AND d.es_discipulo_activo=true
            ORDER BY h.nombre_completo""", (sel,))
        for d in discipulos:
            for t in tipos:
                enc = q("SELECT estado FROM encuentro_participacion WHERE hermano_id=%s AND tipo_encuentro_id=%s", (d["id"], t["id"]), one=True)
                d[f'enc_{t["id"]}'] = enc["estado"] if enc else None
        for t in tipos:
            ap = q("SELECT COUNT(*) AS c FROM encuentro_participacion ep JOIN discipulado d ON d.hermano_id=ep.hermano_id WHERE d.lider_id=%s AND ep.tipo_encuentro_id=%s AND ep.estado='APROBADO' AND d.fecha_fin IS NULL AND d.es_discipulo_activo=true", (sel, t["id"]), one=True)
            stats[t["id"]] = {"aprobados": ap["c"] if ap else 0, "total": len(discipulos)}
    return render_template("enc_guia.html", lideres=lideres, discipulos=discipulos,
        sel_lider=int(sel) if sel else None, tipos=tipos, stats=stats)

@app.post("/encuentros/guia")
@login_req
@permiso_req("encuentros", "editar")
def enc_guia_guardar():
    from datetime import date
    f = request.form
    sel = f.get("lider", "")
    if sel and not exigir_igl("lideres", sel):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("enc_guia"))
    action = f.get("action", "")
    if action == "masivo":
        marcados = f.getlist("marcado")
        estado = f.get("masivo_estado", "APROBADO")
        fecha = date.today().isoformat()
        count = 0
        for m in marcados:
            parts = m.split("|")
            if len(parts) == 2:
                try:
                    _hid = int(parts[0])
                except: continue
                if not exigir_igl("hermanos", _hid): continue
                try:
                    q("""INSERT INTO encuentro_participacion(hermano_id,tipo_encuentro_id,estado,fecha_evento) VALUES(%s,%s,%s,%s)
                         ON CONFLICT (hermano_id,tipo_encuentro_id) DO UPDATE SET estado=EXCLUDED.estado,fecha_evento=EXCLUDED.fecha_evento""",
                      (int(parts[0]), int(parts[1]), estado, fecha), commit=True)
                    count += 1
                except: pass
        flash(f"Actualizados {count} encuentros", "ok")
    elif "toggle" in f:
        parts = f["toggle"].split("|")
        if len(parts) == 3:
            hid, tid, estado = int(parts[0]), int(parts[1]), parts[2]
            if not exigir_igl("hermanos", hid):
                flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("enc_guia", lider=sel))
            fecha = date.today().isoformat()
            try:
                q("""INSERT INTO encuentro_participacion(hermano_id,tipo_encuentro_id,estado,fecha_evento) VALUES(%s,%s,%s,%s)
                     ON CONFLICT (hermano_id,tipo_encuentro_id) DO UPDATE SET estado=EXCLUDED.estado,fecha_evento=EXCLUDED.fecha_evento""",
                  (hid, tid, estado, fecha), commit=True)
                flash("Estado actualizado", "ok")
            except Exception as e: flash(f"Error: {e}", "error")
    return redirect(url_for("enc_guia", lider=sel))

@app.post("/encuentros/<int:eid>/eliminar")
@login_req
@permiso_req("encuentros", "eliminar")
def enc_eliminar(eid):
    if not exigir_igl("encuentro_participacion", eid):
        flash("Ese encuentro no es de tu iglesia", "error"); return redirect(url_for("encuentros"))
    q("DELETE FROM encuentro_participacion WHERE id=%s", (eid,), commit=True)
    log_audit("ELIMINAR","encuentros", f"id={eid}")
    flash("Encuentro eliminado", "ok")
    return redirect(url_for("encuentros"))

# ---------- SEGUIMIENTO DIARIO (sin célula) ----------
@app.get("/seguimiento")
@login_req
def seguimiento():
    exp = request.args.get("export", "")
    fi, pi = filtro_igl()
    if fi:
        rows = q(f"SELECT * FROM v_sin_celula WHERE iglesia_id=%s LIMIT 300", tuple(pi))
    else:
        rows = q("SELECT * FROM v_sin_celula LIMIT 300")
    visitas = q(f"""SELECT v.*, (SELECT MAX(s.fecha) FROM seguimiento s WHERE s.visita_id=v.id) AS ultimo_seg
        FROM v_visitas_completo v WHERE v.estado IN ('PENDIENTE','EN SEGUIMIENTO'){fi.replace('iglesia_id','v.iglesia_id') if fi else ''} ORDER BY ultimo_seg NULLS FIRST LIMIT 100""", tuple(pi))
    total = q(f"SELECT * FROM v_seguimiento_total{(' WHERE iglesia_id=%s' if fi else '')} LIMIT 500", tuple(pi))
    lideres_sin = q(f"""SELECT h.nombre_completo,l.red FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.estado_celula='SIN CELULA' AND l.activo{fi.replace('iglesia_id','l.iglesia_id') if fi else ''}""", tuple(pi))
    if exp == "excel": return to_excel(total, ["tipo", "nombre_completo", "telefono", "invitado_por", "ultimo_seg"], "seguimiento_total.xlsx")
    if exp == "pdf": return to_pdf("Seguimiento total", total, ["tipo", "nombre_completo", "telefono"], "seguimiento.pdf")
    if fi:
        hist = q("""SELECT s.id,COALESCE(h.nombre_completo,v.nombre_completo) AS nombre,s.fecha,s.tipo,s.comentario,s.responsable,s.contactado,
            CASE WHEN s.visita_id IS NOT NULL THEN 'VISITA' ELSE 'HERMANO' END AS origen FROM seguimiento s
            LEFT JOIN hermanos h ON h.id=s.hermano_id LEFT JOIN visitas v ON v.id=s.visita_id WHERE s.iglesia_id=%s ORDER BY s.fecha DESC LIMIT 100""", tuple(pi))
    else:
        hist = q("""SELECT s.id,COALESCE(h.nombre_completo,v.nombre_completo) AS nombre,s.fecha,s.tipo,s.comentario,s.responsable,s.contactado,
            CASE WHEN s.visita_id IS NOT NULL THEN 'VISITA' ELSE 'HERMANO' END AS origen FROM seguimiento s
            LEFT JOIN hermanos h ON h.id=s.hermano_id LEFT JOIN visitas v ON v.id=s.visita_id ORDER BY s.fecha DESC LIMIT 100""")
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.activo{fl} ORDER BY h.nombre_completo""", tuple(pl))
    tipos_seg = get_catalogo("catalogo_tipo_seguimiento")
    return render_template("seguimiento.html", rows=rows, visitas=visitas, total=total, lideres_sin=lideres_sin, hist=hist, lideres=lideres, tipos_seg=tipos_seg)

@app.post("/seguimiento/<int:hid>/registrar")
@login_req
@permiso_req("seguimiento", "crear")
def seg_registrar(hid):
    f = request.form
    if not exigir_igl("hermanos", hid):
        flash("Ese hermano no es de tu iglesia", "error"); return redirect(url_for("seguimiento"))
    igl, _ = iglesia_actual()
    q("INSERT INTO seguimiento(hermano_id,tipo,comentario,responsable,contactado,proximo_contacto,iglesia_id) VALUES(%s,%s,%s,%s,%s,NULLIF(%s,'')::date,%s)",
      (hid, f.get("tipo", "LLAMADA"), f.get("comentario"), f.get("responsable"), f.get("contactado") == "on", f.get("proximo") or None, igl if igl is not None else 1), commit=True)
    log_audit("CREAR","seguimiento", f"hermano {hid} tipo={f.get('tipo')}")
    flash("Seguimiento registrado", "ok"); return redirect(url_for("seguimiento"))

@app.post("/seguimiento/visita/<int:vid>/registrar")
@login_req
@permiso_req("seguimiento", "crear")
def seg_visita_registrar(vid):
    f = request.form
    if not exigir_igl("visitas", vid):
        flash("Esa visita no es de tu iglesia", "error"); return redirect(url_for("seguimiento"))
    vigl = iglesia_de("visitas", vid) or 1
    q("INSERT INTO seguimiento(visita_id,tipo,comentario,responsable,contactado,proximo_contacto,iglesia_id) VALUES(%s,%s,%s,%s,%s,NULLIF(%s,'')::date,%s)",
      (vid, f.get("tipo", "LLAMADA"), f.get("comentario"), f.get("responsable"), f.get("contactado") == "on", f.get("proximo") or None, vigl), commit=True)
    q("UPDATE visitas SET estado='EN SEGUIMIENTO' WHERE id=%s AND estado='PENDIENTE'", (vid,), commit=True)
    log_audit("CREAR","seguimiento", f"visita {vid} tipo={f.get('tipo')}")
    flash("Seguimiento de visita registrado (pasó a EN SEGUIMIENTO)", "ok"); return redirect(url_for("seguimiento"))

@app.post("/seguimiento/<int:sid>/eliminar")
@login_req
@permiso_req("seguimiento", "eliminar")
def seg_eliminar(sid):
    if not exigir_igl("seguimiento", sid):
        flash("Ese seguimiento no es de tu iglesia", "error"); return redirect(url_for("seguimiento"))
    q("DELETE FROM seguimiento WHERE id=%s", (sid,), commit=True)
    log_audit("ELIMINAR","seguimiento", f"id={sid}")
    flash("Seguimiento eliminado", "ok")
    return redirect(url_for("seguimiento"))

# ---------- VISITAS (culto) ----------
@app.get("/visitas")
@login_req
def visitas():
    exp = request.args.get("export", "")
    texto = request.args.get("q", "")
    fest = request.args.get("fest", "")
    fv, pv = filtro_igl()
    base = f"FROM v_visitas_completo WHERE (%s='' OR estado=%s) AND nombre_completo ILIKE %s{fv}"
    params = (fest, fest or None, f"%{texto}%") + tuple(pv)
    rows = q(f"SELECT * {base} ORDER BY fecha_visita DESC LIMIT 300", params)
    if exp == "excel": return to_excel(rows, ["nombre_completo", "direccion", "telefono", "comuna", "invitado_por", "invitado_por_texto", "fecha_visita", "motivo_oracion", "estado", "responsable"], "visitas.xlsx")
    if exp == "pdf": return to_pdf("Visitas MCCI", rows, ["nombre_completo", "telefono", "motivo_oracion", "estado"], "visitas.pdf")
    lideres = q(f"""SELECT l.id, h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.activo{fv.replace('iglesia_id','l.iglesia_id') if fv else ''} ORDER BY h.nombre_completo""", tuple(pv))
    estados_visita = get_catalogo("catalogo_estado_visita")
    tipos_seg = get_catalogo("catalogo_tipo_seguimiento")
    return render_template("visitas.html", rows=rows, texto=texto, fest=fest, lideres=lideres, estados_visita=estados_visita, tipos_seg=tipos_seg)

@app.get("/visitas/nuevo")
@login_req
@permiso_req("visitas", "crear")
def visita_nueva():
    fi, pi = filtro_igl()
    fl, pl = filtro_igl("l")
    return render_template("visita_form.html", v=None,
        comunas=q("SELECT id,nombre FROM comunas ORDER BY nombre"),
        hermanos=q(f"SELECT id,nombre_completo FROM hermanos WHERE activo{fi} ORDER BY 2", tuple(pi)),
        lideres=q(f"SELECT h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY 1", tuple(pl)),
        estados_visita=get_catalogo("catalogo_estado_visita"))

@app.post("/visitas/nuevo")
@login_req
@permiso_req("visitas", "crear")
def visita_crear():
    f = request.form
    igl, _ = iglesia_actual()
    igl = igl if igl is not None else 1
    inv = f.get("invitado") or None
    if inv and not exigir_igl("hermanos", inv):
        flash("Quien invita no es de tu iglesia", "error"); return redirect(url_for("visitas"))
    q("""INSERT INTO visitas(rut,nombre_completo,direccion,telefono,correo,comuna_id,invitado_por_id,invitado_por_texto,
        fecha_visita,motivo_oracion,estado,responsable,observacion,iglesia_id)
        VALUES(NULLIF(%s,''),%s,%s,%s,%s,NULLIF(%s,'')::int,NULLIF(%s,'')::int,NULLIF(%s,''),COALESCE(NULLIF(%s,'')::date,CURRENT_DATE),%s,%s,%s,%s,%s)""",
      (f.get("rut") or None, f["nombre"].strip(), f.get("direccion"), f.get("telefono"), f.get("correo") or None, f.get("comuna") or None,
       inv, f.get("invtexto") or None, f.get("fecha") or None, f["motivo"].strip(),
       f.get("estado", "PENDIENTE"), f.get("responsable") or None, f.get("obs"), igl), commit=True)
    log_audit("CREAR","visitas", f"{f.get('nombre','')}")
    flash("Visita registrada", "ok"); return redirect(url_for("visitas"))

@app.post("/visitas/<int:vid>/estado")
@login_req
@permiso_req("visitas", "editar")
def visita_estado(vid):
    f = request.form
    if not exigir_igl("visitas", vid):
        flash("Esa visita no es de tu iglesia", "error"); return redirect(url_for("visitas"))
    act = q("SELECT * FROM visitas WHERE id=%s", (vid,), one=True)
    if not act: flash("No encontrada", "error"); return redirect(url_for("visitas"))
    # responsable puede vaciarse → NULL
    resp = f.get("responsable")
    if resp == "": resp = None
    elif resp is None: resp = act["responsable"]
    estado_val = f.get("estado") or act["estado"]
    obs_val = f.get("obs") if f.get("obs") is not None else act["observacion"]
    q("UPDATE visitas SET estado=%s, responsable=%s, observacion=%s WHERE id=%s",
      (estado_val, resp, obs_val, vid), commit=True)
    log_audit("EDITAR","visitas", f"id={vid} estado={f.get('estado')}")
    flash("Visita actualizada", "ok"); return redirect(url_for("visitas"))

@app.post("/visitas/<int:vid>/integrar")
@login_req
@permiso_req("visitas", "crear")
def visita_integrar(vid):
    """Integra visita: crea hermano + lo asigna a célula (discipulado)."""
    f = request.form
    lider_id = f.get("lider")
    if lider_id and not exigir_igl("lideres", lider_id):
        flash("Ese lider no es de tu iglesia", "error"); return redirect(url_for("visitas"))
    if not lider_id:
        flash("Para integrar indica líder/célula destino", "error"); return redirect(url_for("visitas"))
    v = q("SELECT * FROM visitas WHERE id=%s", (vid,), one=True)
    if not v: flash("Visita no existe", "error"); return redirect(url_for("visitas"))
    if not exigir_igl("visitas", vid):
        flash("Esa visita no es de tu iglesia", "error"); return redirect(url_for("visitas"))
    if v["estado"] == "INTEGRADO":
        flash("Ya estaba integrada", "error"); return redirect(url_for("visitas"))
    partes = (v["nombre_completo"] or "").split()
    nombres = " ".join(partes[:2]) if len(partes) > 2 else (partes[0] if partes else "S/N")
    paterno = partes[-1] if len(partes) > 1 else "S/A"
    nombres = nombres.upper()
    paterno = paterno.upper()
    try:
        codigo = _codigo_hermano(v["iglesia_id"])
        c = db()
        with c.cursor() as cur:
            cur.execute("""INSERT INTO hermanos(codigo,tipo_hermano,nombres,apellido_paterno,telefono,direccion,correo,comuna_id,invitado_por_id,invitado_por_texto,iglesia_id)
                VALUES(%s,'DISCIPULO',%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (codigo, nombres, paterno, v["telefono"], v["direccion"], v["correo"], v["comuna_id"], v["invitado_por_id"], v["invitado_por_texto"], v["iglesia_id"]))
            hid = cur.fetchone()["id"]
            cur.execute("INSERT INTO discipulado(hermano_id,lider_id,es_discipulo_activo,observacion) VALUES(%s,%s,TRUE,'Integrado desde visita')", (hid, lider_id))
            cur.execute("UPDATE visitas SET estado='INTEGRADO', hermano_id=%s WHERE id=%s", (hid, vid))
            c.commit()
        log_audit("CREAR","visitas", f"integrar visita {vid} -> hermano {codigo} id {hid}")
        flash(f"Integrado: hermano {codigo} (DISCÍPULO) creado y asignado a célula (id {hid})", "ok")
    except Exception as e:
        flash(f"Error al integrar: {e}", "error")
    return redirect(url_for("visitas"))

@app.post("/visitas/<int:vid>/eliminar")
@login_req
@permiso_req("visitas", "eliminar")
def visita_eliminar(vid):
    if not exigir_igl("visitas", vid):
        flash("Esa visita no es de tu iglesia", "error"); return redirect(url_for("visitas"))
    q("DELETE FROM visitas WHERE id=%s", (vid,), commit=True)
    log_audit("ELIMINAR","visitas", f"id={vid}")
    flash("Visita eliminada", "ok")
    return redirect(url_for("visitas"))

# ---------- INFORME CÉLULA ----------
@app.get("/celulas-informe")
@login_req
@permiso_req("celulas", "leer")
def celulas_informe_lista():
    exp = request.args.get("export", "")
    fi, pi = filtro_igl("ci")
    rows = q(f"""SELECT ci.*, h1.nombre_completo AS lider_nombre, h2.nombre_completo AS pastor_nombre,
        i.nombre AS iglesia_nombre,
        CASE WHEN ci.realizo THEN 'SI' ELSE 'NO' END AS realizado
        FROM celulas_informe ci 
        LEFT JOIN iglesias i ON i.id=ci.iglesia_id
        LEFT JOIN lideres l1 ON l1.id=ci.lider_id LEFT JOIN hermanos h1 ON h1.id=l1.hermano_id
        LEFT JOIN lideres l2 ON l2.id=ci.pastor_id LEFT JOIN hermanos h2 ON h2.id=l2.hermano_id
        WHERE 1=1{fi} ORDER BY ci.id DESC""", tuple(pi))
    fl, pl = filtro_igl("l")
    sin_informe = q(f"""SELECT l.id, h.nombre_completo, l.red
        FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.activo AND l.estado_celula != 'SIN CELULA'{fl}
        AND l.id NOT IN (
            SELECT ci.lider_id FROM celulas_informe ci
            WHERE EXTRACT(MONTH FROM ci.fecha) = EXTRACT(MONTH FROM CURRENT_DATE)
            AND EXTRACT(YEAR FROM ci.fecha) = EXTRACT(YEAR FROM CURRENT_DATE)
            AND ci.lider_id IS NOT NULL
        )
        ORDER BY h.nombre_completo""", tuple(pl))
    if exp == "excel":
        return to_excel(rows, ["lider_nombre","pastor_nombre","fecha","mes","red","comuna",
            "horario","realizado","discipulos_asistieron","ofrenda","tipo_ofrenda"], "informe_celulas.xlsx")
    if exp == "pdf":
        return to_pdf("Informe Células MCCI", rows, ["lider_nombre","pastor_nombre","fecha","mes",
            "red","realizado","ofrenda"], "informe_celulas.pdf")
    return render_template("celulas_informe.html", rows=rows, sin_informe=sin_informe)

@app.get("/celulas-informe/nuevo")
@login_req
@permiso_req("celulas", "crear")
def celulas_informe_nuevo():
    """Alta de informe interno: mismo formulario que el informe exprés que
    cargan los líderes (más el selector de líder, que ahí lo da el código)."""
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY h.nombre_completo""", tuple(pl))
    return render_template("celulas_informe_form.html", c=None, lideres=lideres,
                           modalidades=get_catalogo("catalogo_modalidad_celula"),
                           fecha_hoy=date.today().isoformat(),
                           ayuda_timoteo=ayuda_timoteo(None))

@app.post("/celulas-informe/nuevo")
@login_req
@permiso_req("celulas", "crear")
def celulas_informe_crear():
    f = request.form
    err, lid, fecha_sel, timoteo_id = _leer_informe_interno(f)
    if err:
        flash(err, "error"); return redirect(url_for("celulas_informe_nuevo"))
    # Pastor, género, dirección, comuna y horario salen del líder (igual que en
    # el informe exprés); el mes se calcula solo desde la fecha.
    d = datos_celula_lider(lid)
    if not d:
        flash("No se encontró el líder elegido.", "error")
        return redirect(url_for("celulas_informe_nuevo"))
    cigl = iglesia_de("lideres", lid) or session.get("iglesia_id") or 1
    try:
        q("""INSERT INTO celulas_informe
                (lider_id, pastor_id, fecha, mes, red, direccion, comuna, horario,
                 realizo, justificacion, discipulos_asistieron, discipulos_no_asistieron,
                 ofrenda, tipo_ofrenda, modalidad, timoteo_id, iglesia_id)
             VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
          (lid, d["pastor_id"], fecha_sel, MESES_ES[fecha_sel.month], d["red"],
           d["direccion"], d["comuna"], d["horario"],
           f.get("realizo") == "1", (f.get("justificacion") or "").strip(),
           f.get("asistieron"), f.get("no_asistieron"),
           f.get("ofrenda") or "0", f.get("tipo_ofrenda"), f.get("modalidad"),
           timoteo_id, cigl), commit=True)
        flash("Informe registrado", "ok")
    except Exception as e:
        flash(f"{mensaje_error_informe(e, admin=True)} El detalle quedó registrado "
              "en Configuración › Errores.", "error")
        log_error_informe("ALTA_ADMIN", lid, e,
                          f"fecha={f.get('fecha')} tipo_ofrenda={f.get('tipo_ofrenda')} ofrenda={f.get('ofrenda')}")
    return redirect(url_for("celulas_informe_lista"))

@app.get("/celulas-informe/<int:cid>/editar")
@login_req
@permiso_req("celulas", "editar")
def celulas_informe_editar(cid):
    if not exigir_igl("celulas_informe", cid):
        flash("Ese informe no es de tu iglesia", "error"); return redirect(url_for("celulas_informe_lista"))
    c = q("SELECT * FROM celulas_informe WHERE id=%s", (cid,), one=True)
    if not c: abort(404)
    fl, pl = filtro_igl("l")
    lideres = q(f"""SELECT l.id, h.nombre_completo FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY h.nombre_completo""", tuple(pl))
    modalidades = get_catalogo("catalogo_modalidad_celula")
    # Si el informe tiene una modalidad que ya no está en el catálogo, se agrega
    # para que al editar no cambie sola.
    if c.get("modalidad") and not any(m["nombre"] == c["modalidad"] for m in modalidades):
        modalidades = list(modalidades) + [{"nombre": c["modalidad"]}]
    return render_template("celulas_informe_form.html", c=c, lideres=lideres,
                           modalidades=modalidades,
                           fecha_hoy=date.today().isoformat(),
                           ayuda_timoteo=ayuda_timoteo(c.get("lider_id")))

@app.post("/celulas-informe/<int:cid>/editar")
@login_req
@permiso_req("celulas", "editar")
def celulas_informe_update(cid):
    if not exigir_igl("celulas_informe", cid):
        flash("Ese informe no es de tu iglesia", "error"); return redirect(url_for("celulas_informe_lista"))
    c = q("SELECT * FROM celulas_informe WHERE id=%s", (cid,), one=True)
    if not c: abort(404)
    f = request.form
    err, lid, fecha_sel, timoteo_id = _leer_informe_interno(f)
    if err:
        flash(err, "error"); return redirect(url_for("celulas_informe_editar", cid=cid))
    # Los datos de la célula (pastor, género, dirección, comuna y horario) se
    # conservan como están; solo se recalculan si cambia el líder del informe.
    if str(c.get("lider_id") or "") == str(lid):
        pastor_id, red = c.get("pastor_id"), c.get("red")
        direccion, comuna, horario = c.get("direccion"), c.get("comuna"), c.get("horario")
    else:
        d = datos_celula_lider(lid)
        if not d:
            flash("No se encontró el líder elegido.", "error")
            return redirect(url_for("celulas_informe_editar", cid=cid))
        pastor_id, red = d["pastor_id"], d["red"]
        direccion, comuna, horario = d["direccion"], d["comuna"], d["horario"]
    try:
        q("""UPDATE celulas_informe SET
             lider_id=%s, pastor_id=%s, fecha=%s, mes=%s, red=%s, direccion=%s, comuna=%s,
             horario=%s, realizo=%s, justificacion=%s, discipulos_asistieron=%s,
             discipulos_no_asistieron=%s, ofrenda=%s, tipo_ofrenda=%s, modalidad=%s,
             timoteo_id=%s
           WHERE id=%s""",
          (lid, pastor_id, fecha_sel, MESES_ES[fecha_sel.month], red, direccion,
           comuna, horario, f.get("realizo") == "1", (f.get("justificacion") or "").strip(),
           f.get("asistieron"), f.get("no_asistieron"), f.get("ofrenda") or "0",
           f.get("tipo_ofrenda"), f.get("modalidad"), timoteo_id, cid), commit=True)
        log_audit("EDITAR","celulas", f"id={cid}")
        flash("Informe actualizado", "ok")
    except Exception as e:
        flash(f"{mensaje_error_informe(e, admin=True)} El detalle quedó registrado "
              "en Configuración › Errores.", "error")
        log_error_informe("EDICION_ADMIN", lid, e,
                          f"id={cid} fecha={f.get('fecha')} tipo_ofrenda={f.get('tipo_ofrenda')}")
    return redirect(url_for("celulas_informe_lista"))

@app.post("/celulas-informe/<int:cid>/eliminar")
@login_req
@permiso_req("celulas", "eliminar")
def celulas_informe_eliminar(cid):
    if not exigir_igl("celulas_informe", cid):
        flash("Ese informe no es de tu iglesia", "error"); return redirect(url_for("celulas_informe_lista"))
    try:
        q("DELETE FROM celulas_informe WHERE id=%s", (cid,), commit=True)
        log_audit("ELIMINAR","celulas", f"id={cid}")
        flash("Informe eliminado", "ok")
    except Exception as e:
        flash(f"Error: {e}", "error")
    return redirect(url_for("celulas_informe_lista"))

@app.get("/celulas-informe/timoteos/<int:lid>")
@login_req
@permiso_req("celulas", "leer")
def celulas_informe_timoteos(lid):
    """Timoteos (discípulos/sub-líderes) de un líder, para el selector del
    formulario interno de informe. Devuelve JSON: la lista, la ayuda (mismo
    texto que ve el líder en el informe exprés) y los datos de la célula que
    se guardarán (que salen del líder, igual que en el exprés)."""
    if not exigir_igl("lideres", lid):
        return jsonify(ok=False, error="iglesia"), 403
    rows = q(_SQL_TIMOTEOS, (lid, lid))
    d = datos_celula_lider(lid)
    return jsonify(ok=True,
                   timoteos=[{"id": r["id"], "nombre": r["nombre_completo"]} for r in rows],
                   ayuda=str(ayuda_timoteo(lid)),
                   lider=({"nombre": d["nombre_completo"], "red": d["red"] or "",
                           "pastor": d["pastor"] or "", "direccion": d["direccion"] or "",
                           "comuna": d["comuna"] or ""} if d else None))

# ---------- INFORME EXPRESS (público, sin login) ----------
from datetime import date, datetime

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

def semana_actual():
    """Devuelve 'Semana X del MES' para la etiqueta del formulario."""
    hoy = date.today()
    # semana del mes (1-5)
    inicio_mes = date(hoy.year, hoy.month, 1)
    dias_desde_inicio = (hoy - inicio_mes).days
    semana = (dias_desde_inicio // 7) + 1
    return f"Semana {semana} de {MESES_ES[hoy.month]}"

# Caducidad de la sesión del informe exprés: 30 minutos (ventana deslizante:
# cada envío/formulario activo reinicia el contador; si pasan 30 min sin activity, caduca).
INFORME_SESION_SEG = 30 * 60

# Timoteos disponibles para un líder: discípulos activos (discipulado) + sub-líderes.
# UNION (no ALL) evita duplicados si alguien es ambas cosas. Devuelve id (hermanos.id) y nombre.
_SQL_TIMOTEOS = """SELECT * FROM (
    SELECT h.id, h.nombre_completo FROM discipulado d
      JOIN hermanos h ON h.id=d.hermano_id
      WHERE d.lider_id=%s AND d.fecha_fin IS NULL AND d.es_discipulo_activo
    UNION
    SELECT h.id, h.nombre_completo FROM lideres l2
      JOIN hermanos h ON h.id=l2.hermano_id
      WHERE l2.lider_padre_id=%s AND l2.activo
) t ORDER BY nombre_completo"""

# Pastor REAL de un líder: sube la cadena de líder_padre_id hasta encontrar una
# PASTORA / PASTOR / APOSTOL. Antes solo se miraba el padre directo, así que un
# líder que estuviera bajo otro líder (y no bajo un pastor) no aparecía en el
# reporte "Informe Célula · Por Pastor" (se perdían 26 de los 48 informes).
_SQL_PASTOR_DE = """
WITH RECURSIVE subo AS (
    SELECT l.id AS lider_id, l.lider_padre_id AS anc_id, 1 AS nivel, ARRAY[l.id] AS ruta
    FROM lideres l
  UNION ALL
    SELECT s.lider_id, p.lider_padre_id, s.nivel + 1, s.ruta || p.id
    FROM subo s JOIN lideres p ON p.id = s.anc_id
    WHERE s.nivel < 20 AND NOT p.id = ANY(s.ruta)
), candidatos AS (
    -- el pastor más cercano por la cadena de líder_padre_id
    (SELECT DISTINCT ON (s.lider_id) s.lider_id, s.anc_id AS pastor_id, 0 AS prioridad
     FROM subo s
     JOIN lideres pp ON pp.id = s.anc_id
     WHERE s.anc_id IS NOT NULL
       AND (pp.es_pastor OR pp.rol_lider IN ('APOSTOL','PASTOR','PASTORA'))
     ORDER BY s.lider_id, s.nivel)
  UNION ALL
    -- pastor/a sin nadie arriba (los apóstoles): su propia célula queda bajo
    -- su nombre, para que ningún informe quede fuera del reporte "Por Pastor"
    (SELECT l.id, l.id, 1
     FROM lideres l
     WHERE (l.es_pastor OR l.rol_lider IN ('APOSTOL','PASTOR','PASTORA'))
       AND NOT EXISTS (
         SELECT 1 FROM subo s2
         WHERE s2.lider_id = l.id AND s2.anc_id IS NOT NULL
           AND EXISTS (SELECT 1 FROM lideres pp WHERE pp.id = s2.anc_id
                       AND (pp.es_pastor OR pp.rol_lider IN ('APOSTOL','PASTOR','PASTORA')))
       ))
), pastor_de AS (
    SELECT DISTINCT ON (lider_id) lider_id, pastor_id
    FROM candidatos ORDER BY lider_id, prioridad
)"""

def pastor_de_lider(lid):
    """Id del pastor real de un líder (None si no tiene ninguno arriba)."""
    if not lid:
        return None
    r = q(_SQL_PASTOR_DE + "SELECT pastor_id FROM pastor_de WHERE lider_id=%s", (lid,), one=True)
    return r["pastor_id"] if r else None

def datos_celula_lider(lid):
    """Datos de la célula que el informe exprés toma SOLOS del líder (pastor,
    género/red, dirección, comuna y horario). El formulario interno de informe
    hace lo mismo, para que lo que registra un admin sea idéntico a lo que
    carga el líder. El pastor es el REAL (sube la cadena de líderes)."""
    d = q("""SELECT l.id, l.red, h.nombre_completo, h.direccion,
                    c.nombre AS comuna,
                    (SELECT ci.horario FROM celulas_informe ci
                     WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS horario
             FROM lideres l
             JOIN hermanos h ON h.id=l.hermano_id
             LEFT JOIN comunas c ON c.id=h.comuna_id
             WHERE l.id=%s AND l.activo""", (lid,), one=True)
    if not d:
        return None
    d["pastor_id"] = pastor_de_lider(lid)
    d["pastor"] = ""
    if d["pastor_id"]:
        pn = q("""SELECT h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
                  WHERE l.id=%s""", (d["pastor_id"],), one=True)
        d["pastor"] = pn["nombre_completo"] if pn else ""
    return d

def ayuda_timoteo(lid):
    """Mismo texto de ayuda que ve el líder en el informe exprés (ft.1),
    calculado para el líder elegido: cuántos timoteos tiene y cuántos informes
    lleva esta semana ('N de 3')."""
    if not lid:
        return ay("celulas_informe_form.ft.0")
    n = len(q(_SQL_TIMOTEOS, (lid, lid)))
    hoy = date.today()
    anio_hoy, sem_hoy, _ = hoy.isocalendar()
    n_sem = q("""SELECT COUNT(*) AS c FROM celulas_informe
                 WHERE lider_id=%s
                   AND EXTRACT(YEAR FROM fecha)::int=%s
                   AND EXTRACT(WEEK FROM fecha)::int=%s""",
              (lid, anio_hoy, sem_hoy), one=True)["c"] or 0
    t = ("Tienes %d a tu cargo." % n) if n else "Aún no tienes timoteos registrados a tu cargo."
    return ay("informe_express.ft.1", timoteos=t, semana=n_sem)

def _leer_informe_interno(f):
    """Lee y valida el formulario interno de informe (mismas reglas que el
    informe exprés). Devuelve (error, líder, fecha, timoteo_id)."""
    lid = (f.get("lider_id") or "").strip()
    if not lid or not f.get("fecha"):
        return ("No se guardó: faltan datos obligatorios (líder y fecha).", None, None, None)
    if not exigir_igl("lideres", lid):
        return ("Ese líder no es de tu iglesia.", None, None, None)
    try:
        fecha_sel = date.fromisoformat(f.get("fecha", ""))
    except Exception:
        return ("Fecha inválida.", None, None, None)
    if fecha_sel > date.today():
        return ("La fecha no puede ser futura.", None, None, None)
    # Si NO se realizó, el motivo es obligatorio (como en el exprés).
    if f.get("realizo") != "1" and not (f.get("justificacion") or "").strip():
        return ("Si la célula NO se realizó, debe indicar el motivo.", lid, None, None)
    tid = (f.get("timoteo_id") or "").strip()
    timoteo_id = int(tid) if tid.isdigit() else None
    if timoteo_id is not None and timoteo_id not in [t["id"] for t in q(_SQL_TIMOTEOS, (lid, lid))]:
        return ("Ese timoteo no está a cargo del líder elegido.", lid, None, None)
    return (None, lid, fecha_sel, timoteo_id)

def _limpiar_sesion_informe():
    """Borra todas las claves de sesión del informe exprés."""
    for k in ("informe_lider_id", "informe_lider_nombre",
              "informe_iglesia_id", "informe_iglesia_nombre", "informe_expira"):
        session.pop(k, None)

@app.get("/informe")
def informe_express():
    """Paso 1: pide código de líder; si ya hay sesión (y no caducó), muestra formulario."""
    if session.get("informe_lider_id"):
        exp = session.get("informe_expira")
        if exp is None or time.time() > float(exp):
            # Sesión caducada (30 min de inactividad): se bloquea y se avisa.
            _limpiar_sesion_informe()
            flash("Tu sesión expiró a los 30 minutos. Ingresa tu código de líder nuevamente para continuar.", "error")
        else:
            session["informe_expira"] = time.time() + INFORME_SESION_SEG  # ventana deslizante
            if request.args.get("enviado"):
                return render_template("informe_enviado.html",
                                       nombre=session.get("informe_lider_nombre"))
            return _mostrar_formulario_informe()
    # Sin sesión vigente: muestra la pantalla de código
    pregunta = nuevo_captcha("captcha_informe")
    return render_template("informe_express.html",
                           lider=None,
                           captcha_pregunta=pregunta,
                           turnstile_sitekey=turnstile_sitekey_visible(),
                           semana_actual=semana_actual())

@app.post("/informe/codigo")
def informe_codigo():
    """Valida el código del líder y crea la sesión."""
    ip = ip_cliente()
    codigo = (request.form.get("codigo") or "").strip()

    # 1) Bloqueo por código (reutiliza la misma lógica: 5 fallos = 15 min → 1 h → 24 h)
    mins = _bloqueo_por_codigo(codigo, ip)
    if mins:
        registrar_intento(codigo, ip, False, "bloqueo_codigo")
        flash(f"Demasiados intentos fallidos. Espere {texto_espera(mins)} antes de volver a intentar.", "error")
        return redirect(url_for("informe_express"))

    # 2) Honeypot + captcha + Turnstile + tiempo mínimo
    err, tipo = chequear_form_seguro("captcha_informe")
    if err:
        registrar_intento(codigo, ip, False, tipo)
        flash(err, "error")
        return redirect(url_for("informe_express"))

    # 3) Buscar líder por código (+ último informe para horario/modalidad)
    # El pastor que se muestra es el REAL: sube la cadena de líder_padre_id.
    lider = q(_SQL_PASTOR_DE + """SELECT l.id, h.nombre_completo, l.red, l.iglesia_id,
                       i.nombre AS iglesia_nombre, COALESCE(i.activo,true) AS iglesia_activa,
                       hp.nombre_completo AS pastor_nombre,
                       h.direccion,
                       c.nombre AS comuna,
                       (SELECT ci.horario FROM celulas_informe ci
                        WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS horario,
                       (SELECT ci.modalidad FROM celulas_informe ci
                        WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS modalidad
                FROM lideres l
                JOIN hermanos h ON h.id=l.hermano_id
                LEFT JOIN comunas c ON c.id=h.comuna_id
                LEFT JOIN iglesias i ON i.id=l.iglesia_id
                LEFT JOIN pastor_de pd ON pd.lider_id = l.id
                LEFT JOIN lideres lp ON lp.id = pd.pastor_id
                LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
                WHERE l.codigo_informe=%s AND l.activo""", (codigo,), one=True)

    if not lider:
        registrar_intento(codigo, ip, False, "codigo_invalido")
        flash("Código inválido o líder inactivo. Verifique e intente de nuevo.", "error")
        return redirect(url_for("informe_express"))

    # Éxito: crea sesión ligera (1 hora)
    if lider.get("iglesia_activa") is False:
        registrar_intento(codigo, ip, False, "iglesia_inactiva")
        flash("La iglesia de este lider esta desactivada. Contacte a casa matriz.", "error")
        return redirect(url_for("informe_express"))

    session["informe_lider_id"] = lider["id"]
    session["informe_lider_nombre"] = lider["nombre_completo"]
    session["informe_iglesia_id"] = lider["iglesia_id"] or 1
    session["informe_iglesia_nombre"] = lider["iglesia_nombre"] or "CASA MATRIZ"
    session["informe_expira"] = time.time() + INFORME_SESION_SEG   # caduca a los 30 min (deslizante)
    registrar_intento(codigo, ip, True, "codigo_ok")
    return redirect(url_for("informe_express"))

def _bloqueo_por_codigo(codigo, ip):
    """Bloqueo exponencial por código (como login pero sin usuario).
    5 fallos en 15 min → 15 min → 1 h → 24 h (memoria 7 días)."""
    # ¿Bloqueo vigente?
    for detalle, clave in (("bloqueo_codigo", codigo), ("bloqueo_ip", ip)):
        b = _bloqueo_activo(detalle, clave, ip)
        if b:
            return max(b[0], 1)
    # ¿Supera umbral?
    try:
        row = q("""
            SELECT
              (SELECT count(*) FROM intentos_login x
                 WHERE NOT x.exito AND x.detalle IN ('codigo_invalido','captcha','turnstile','rapido','honeypot','sesion')
                   AND x.username=%s
                   AND x.creado > now() - make_interval(mins => %s)
                   AND x.creado > COALESCE((SELECT max(y.creado) FROM intentos_login y
                                            WHERE y.exito AND y.detalle='codigo_ok' AND y.username=%s), '-infinity'::timestamp)
              ) AS fallos_codigo,
              (SELECT count(*) FROM intentos_login x
                 WHERE NOT x.exito AND x.ip=%s
                   AND x.creado > now() - make_interval(mins => %s)
              ) AS fallos_ip""",
            (codigo, LOGIN_VENTANA_MIN, codigo, ip, LOGIN_VENTANA_MIN), one=True)
        if not row:
            return 0
        if (row["fallos_codigo"] or 0) >= LOGIN_FALLOS_USER:
            return _activar_bloqueo("bloqueo_codigo", codigo, ip)[0]
        if (row["fallos_ip"] or 0) >= LOGIN_FALLOS_IP:
            return _activar_bloqueo("bloqueo_ip", "", ip)[0]
    except Exception as e:
        print(f"_bloqueo_por_codigo failed: {e}")
    return 0

def _mostrar_formulario_informe():
    """Muestra el formulario con datos pre-cargados del líder."""
    lid = session.get("informe_lider_id")
    if not lid:
        return redirect(url_for("informe_express"))

    # Datos fijos del líder (para precargar). El pastor que se muestra es el
    # REAL: sube la cadena de líder_padre_id hasta una PASTORA / PASTOR / APOSTOL.
    lider = q(_SQL_PASTOR_DE + """SELECT l.id, h.nombre_completo AS nombre, l.red, l.iglesia_id,
                       i.nombre AS iglesia_nombre,
                       hp.nombre_completo AS pastor,
                       h.direccion,
                       c.nombre AS comuna,
                       (SELECT ci.horario FROM celulas_informe ci
                        WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS horario,
                       (SELECT ci.modalidad FROM celulas_informe ci
                        WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS modalidad
                FROM lideres l
                JOIN hermanos h ON h.id=l.hermano_id
                LEFT JOIN comunas c ON c.id=h.comuna_id
                LEFT JOIN iglesias i ON i.id=l.iglesia_id
                LEFT JOIN pastor_de pd ON pd.lider_id = l.id
                LEFT JOIN lideres lp ON lp.id = pd.pastor_id
                LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
                WHERE l.id=%s AND l.activo""", (lid,), one=True)
    if not lider:
        session.pop("informe_lider_id", None)
        session.pop("informe_lider_nombre", None)
        flash("Sesión expirada. Ingrese su código nuevamente.", "error")
        return redirect(url_for("informe_express"))

    # Timoteos disponibles (discípulos activos + sub-líderes) y cuántos informes
    # se han enviado esta semana (para mostrar "N de 3").
    timoteos = q(_SQL_TIMOTEOS, (lid, lid))
    hoy = date.today()
    anio_hoy, sem_hoy, _ = hoy.isocalendar()
    n_sem = q("""SELECT COUNT(*) AS c FROM celulas_informe
                 WHERE lider_id=%s
                   AND EXTRACT(YEAR FROM fecha)::int=%s
                   AND EXTRACT(WEEK FROM fecha)::int=%s""",
              (lid, anio_hoy, sem_hoy), one=True)["c"]

    # Formulario SIEMPRE en blanco: no se pre-carga el último informe, así
    # después de enviar uno el siguiente sale limpio para rellenar. La única
    # excepción es un borrador por un guardado fallido, para no perder lo que
    # el líder ya había escrito.
    borrador = session.pop("informe_borrador", None)
    pregunta = nuevo_captcha("captcha_informe2")
    return render_template("informe_express.html",
                           lider=lider,
                           existente=None,
                           borrador=borrador,
                           timoteos=timoteos,
                           informes_semana=n_sem or 0,
                           captcha_pregunta2=pregunta,
                           turnstile_sitekey=turnstile_sitekey_visible(),
                           fecha_hoy=hoy.isoformat(),
                           semana_actual=semana_actual(),
                           modalidades=get_catalogo("catalogo_modalidad_celula", iglesia=lider["iglesia_id"] or 1),
                           iglesia_nombre=lider["iglesia_nombre"])

@app.post("/informe/guardar")
def informe_guardar():
    """Guarda el informe (insert o update si ya existe esa semana)."""
    lid = session.get("informe_lider_id")
    if not lid:
        flash("Sesión expirada. Ingrese su código nuevamente.", "error")
        return redirect(url_for("informe_express"))
    # Caducidad (30 min): si pasó, se bloquea el envío.
    exp = session.get("informe_expira")
    if exp is None or time.time() > float(exp):
        _limpiar_sesion_informe()
        flash("Tu sesión expiró a los 30 minutos. Ingresa tu código de líder nuevamente.", "error")
        return redirect(url_for("informe_express"))

    ip = ip_cliente()
    f = request.form

    # Seguridad: honeypot + captcha + Turnstile + tiempo mínimo
    err, tipo = chequear_form_seguro("captcha_informe2")
    if err:
        registrar_intento(str(lid), ip, False, tipo)
        flash(err, "error")
        return redirect(url_for("informe_express"))

    # Validación: si NO se realizó, exige justificación
    realizo = f.get("realizo") == "1"
    justificacion = (f.get("justificacion") or "").strip()
    if not realizo and not justificacion:
        flash("Si la célula NO se realizó, debe indicar el motivo.", "error")
        return redirect(url_for("informe_express"))

    # Fecha elegida por el líder (viene del formulario)
    try:
        fecha_sel = date.fromisoformat(f.get("fecha", ""))
    except Exception:
        flash("Fecha inválida.", "error")
        return redirect(url_for("informe_express"))

    # No permitir fechas futuras
    if fecha_sel > date.today():
        flash("La fecha no puede ser futura.", "error")
        return redirect(url_for("informe_express"))

    mes_sel = MESES_ES[fecha_sel.month]
    anio_iso, semana_iso, _ = fecha_sel.isocalendar()

    # --- Timoteo (sub-líder / discípulo a cargo) de este informe ---
    timoteo_id = None
    t_raw = (f.get("timoteo") or "").strip()
    if t_raw:
        try:
            timoteo_id = int(t_raw)
        except ValueError:
            timoteo_id = None
    # Seguridad: el timoteo elegido debe estar realmente a cargo de este líder.
    if timoteo_id is not None:
        permitidos = [t["id"] for t in q(_SQL_TIMOTEOS, (lid, lid))]
        if timoteo_id not in permitidos:
            flash("Ese timoteo no está a tu cargo. Elige uno de tu lista.", "error")
            return redirect(url_for("informe_express"))

    # ¿Ya existe informe de ESTE líder + ESTE timoteo en ESTA FECHA? → corrige (UPDATE)
    existente = q("""SELECT id FROM celulas_informe
                     WHERE lider_id=%s
                       AND fecha=%s
                       AND timoteo_id IS NOT DISTINCT FROM %s
                     ORDER BY created_at DESC LIMIT 1""",
                  (lid, fecha_sel, timoteo_id), one=True)

    # Si es un informe NUEVO (timoteo distinto), respetar el máximo de 3 por semana.
    if not existente:
        n_sem = q("""SELECT COUNT(*) AS c FROM celulas_informe
                     WHERE lider_id=%s
                       AND EXTRACT(YEAR FROM fecha)::int=%s
                       AND EXTRACT(WEEK FROM fecha)::int=%s""",
                  (lid, anio_iso, semana_iso), one=True)["c"]
        if (n_sem or 0) >= 3:
            flash("Ya enviaste el máximo de 3 informes para esta semana. "
                  "Si quieres corregir uno, repite la misma fecha y el mismo timoteo que usaste antes.",
                  "error")
            return redirect(url_for("informe_express"))

    # Datos del líder para pastor_id, red, etc. El pastor es el REAL (sube la
    # cadena de líder_padre_id hasta una PASTORA / PASTOR / APOSTOL).
    ldata = q("""SELECT l.lider_padre_id, l.red, l.iglesia_id,
                       h.direccion,
                       c.nombre AS comuna,
                       (SELECT ci.horario FROM celulas_informe ci
                        WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS horario,
                       (SELECT ci.modalidad FROM celulas_informe ci
                        WHERE ci.lider_id=l.id ORDER BY ci.created_at DESC LIMIT 1) AS modalidad
                FROM lideres l
                JOIN hermanos h ON h.id=l.hermano_id
                LEFT JOIN comunas c ON c.id=h.comuna_id
                WHERE l.id=%s""", (lid,), one=True)
    ldata["pastor_id"] = pastor_de_lider(lid)

    try:
        if existente:
            q("""UPDATE celulas_informe SET
                     pastor_id=%s, mes=%s, red=%s, direccion=%s, comuna=%s, horario=%s,
                     realizo=%s, justificacion=%s, discipulos_asistieron=%s,
                     discipulos_no_asistieron=%s, ofrenda=%s, tipo_ofrenda=%s, modalidad=%s,
                     timoteo_id=%s
                 WHERE id=%s""",
                (ldata["pastor_id"], mes_sel, ldata["red"], ldata["direccion"],
                 ldata["comuna"], ldata["horario"], realizo, justificacion,
                 f.get("asistieron"), f.get("no_asistieron"),
                 f.get("ofrenda") or "0", f.get("tipo_ofrenda"), f.get("modalidad"),
                 timoteo_id, existente["id"]), commit=True)
        else:
            q("""INSERT INTO celulas_informe
                    (lider_id, pastor_id, fecha, mes, red, direccion, comuna, horario,
                     realizo, justificacion, discipulos_asistieron, discipulos_no_asistieron,
                     ofrenda, tipo_ofrenda, modalidad, timoteo_id, iglesia_id)
                 VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (lid, ldata["pastor_id"], fecha_sel, mes_sel,
                 ldata["red"], ldata["direccion"], ldata["comuna"], ldata["horario"],
                 realizo, justificacion, f.get("asistieron"), f.get("no_asistieron"),
                 f.get("ofrenda") or "0", f.get("tipo_ofrenda"), f.get("modalidad"),
                 timoteo_id, ldata["iglesia_id"] or 1), commit=True)
        log_audit("INFORME_EXPRESS", "celulas", f"lider={lid} realizo={realizo} ip={ip}")
        session["informe_expira"] = time.time() + INFORME_SESION_SEG  # refresca ventana tras enviar
        session.pop("informe_borrador", None)  # ya quedó guardado
        # Éxito: pantalla de confirmación con botón "Enviar otro informe".
        return redirect(url_for("informe_express", enviado=1))
    except Exception as e:
        # Se conserva lo que el líder ya escribió, para que no tenga que
        # repetir el formulario (se usa al volver a /informe y se limpia al
        # enviar con éxito). El mensaje que ve el líder es amable; el detalle
        # técnico queda en el log y en Configuración › pestaña de errores.
        session["informe_borrador"] = {
            "fecha": f.get("fecha", ""),
            "timoteo": (f.get("timoteo") or "").strip(),
            "realizo": "1" if realizo else "0",
            "justificacion": (f.get("justificacion") or "").strip(),
            "asistieron": f.get("asistieron") or "",
            "no_asistieron": f.get("no_asistieron") or "",
            "ofrenda": f.get("ofrenda") or "",
            "tipo_ofrenda": f.get("tipo_ofrenda") or "",
            "modalidad": f.get("modalidad") or "",
        }
        flash(mensaje_error_informe(e), "error")
        print(f"informe_guardar error: {type(e).__name__}: {e}")
        log_error_informe("INFORME_EXPRESS", lid, e,
                          f"fecha={f.get('fecha')} realizo={realizo} tipo_ofrenda={f.get('tipo_ofrenda')} ofrenda={f.get('ofrenda')} ip={ip}")

    return redirect(url_for("informe_express"))

@app.get("/informe/salir")
def informe_salir():
    """Limpia la sesión del informe express."""
    _limpiar_sesion_informe()
    flash("Sesión cerrada.", "ok")
    return redirect(url_for("informe_express"))

# ---------- REPORTES ----------
@app.get("/reportes")
@login_req
def reportes():
    exp = request.args.get("export", "")
    ver = request.args.get("ver", "todos")
    lider_sel = request.args.get("lider", type=int)
    r = {}
    # Filtros por iglesia en sesion (SUPERADMIN en TODAS -> vacios)
    fl, pl = filtro_igl("l")
    fh, ph = filtro_igl("h")
    f0, p0 = filtro_igl()
    fvv, pvv = filtro_igl("v")
    r["por_red"] = q(f"SELECT red AS red,estado_celula,COUNT(*) AS c FROM lideres WHERE activo{f0} GROUP BY 1,2 ORDER BY 1,2", tuple(p0))
    r["por_comuna"] = q(f"""SELECT r.nombre_corto AS region,c.nombre AS comuna,COUNT(*) AS c FROM hermanos h
        LEFT JOIN comunas c ON c.id=h.comuna_id LEFT JOIN regiones r ON r.id=c.region_id WHERE h.activo{fh} GROUP BY 1,2 ORDER BY 3 DESC LIMIT 30""", tuple(ph))
    r["sin_celula"] = q(f"""SELECT h.nombre_completo,l.red AS red FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
        WHERE l.estado_celula='SIN CELULA' AND l.activo{fl}""", tuple(pl))
    r["pendientes"] = q(f"""SELECT h.nombre_completo,t.nombre AS etapa,e.fecha_evento AS fecha FROM encuentro_participacion e
        JOIN hermanos h ON h.id=e.hermano_id JOIN tipos_encuentro t ON t.id=e.tipo_encuentro_id
        WHERE e.estado='PENDIENTE'{fh} ORDER BY 3 LIMIT 50""", tuple(ph))
    r["ruta_incompleta"] = q(f"SELECT nombre_completo FROM v_ruta_espiritual WHERE escuela_lideres=0{f0} ORDER BY 1 LIMIT 100", tuple(p0))
    r["periodicidad"] = q(f"""SELECT p.nombre AS nombre,COUNT(d.id) AS c FROM discipulado d JOIN lideres l ON l.id=d.lider_id
        JOIN periodicidad_celula p ON p.id=d.periodicidad_id
        WHERE d.fecha_fin IS NULL{fl} GROUP BY 1 ORDER BY 2 DESC""", tuple(pl))
    r["por_12"] = q(f"SELECT tipo_12, COUNT(*) AS c, SUM(CASE WHEN es_pastor THEN 1 ELSE 0 END) AS pastores FROM lideres WHERE activo{f0} GROUP BY 1 ORDER BY 2 DESC", tuple(p0))
    # NUEVO: Resumen de Informes de Célula por líder, pastor y red
    # (el pastor es el REAL: sube la cadena de líder_padre_id hasta una
    #  PASTORA / PASTOR / APOSTOL, no solo el padre directo)
    r["celulas_resumen"] = q(f"""{_SQL_PASTOR_DE}
        SELECT 
            l.id AS lider_id,
            h.nombre_completo AS lider,
            l.red,
            COALESCE(hp.nombre_completo, 'SIN PASTOR') AS pastor,
            COUNT(ci.id) AS total_informes,
            SUM(CASE WHEN ci.realizo THEN 1 ELSE 0 END) AS realizadas,
            SUM(CASE WHEN NOT ci.realizo THEN 1 ELSE 0 END) AS no_realizadas,
            COALESCE(SUM(ci.ofrenda), 0) AS total_ofrenda,
            MAX(ci.fecha) AS ultima_fecha,
            MAX(ci.created_at) AS ultimo_registro
        FROM lideres l
        JOIN hermanos h ON h.id = l.hermano_id
        LEFT JOIN pastor_de pd ON pd.lider_id = l.id
        LEFT JOIN lideres lp ON lp.id = pd.pastor_id
        LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
        LEFT JOIN celulas_informe ci ON ci.lider_id = l.id
        WHERE l.activo{fl}
        GROUP BY l.id, h.nombre_completo, l.red, hp.nombre_completo
        ORDER BY l.red, h.nombre_completo
    """, tuple(pl))
    # Resumen agrupado por pastor (es_pastor=true O rol pastoral: APOSTOL / PASTOR / PASTORA),
    # subiendo la cadena de líder_padre_id si el líder está bajo otro líder.
    r["celulas_por_pastor"] = q(f"""{_SQL_PASTOR_DE}
        SELECT 
            hp.nombre_completo AS pastor,
            lp.red AS red_pastor,
            COUNT(DISTINCT l.id) AS total_lideres,
            COUNT(ci.id) AS total_informes,
            SUM(CASE WHEN ci.realizo THEN 1 ELSE 0 END) AS realizadas,
            SUM(CASE WHEN NOT ci.realizo THEN 1 ELSE 0 END) AS no_realizadas,
            COALESCE(SUM(ci.ofrenda), 0) AS total_ofrenda
        FROM lideres l
        JOIN hermanos h ON h.id = l.hermano_id
        JOIN pastor_de pd ON pd.lider_id = l.id
        JOIN lideres lp ON lp.id = pd.pastor_id
        JOIN hermanos hp ON hp.id = lp.hermano_id
        LEFT JOIN celulas_informe ci ON ci.lider_id = l.id
        WHERE l.activo{fl}
        GROUP BY hp.nombre_completo, lp.red
        ORDER BY hp.nombre_completo
    """, tuple(pl))
    # Resumen agrupado por red
    r["celulas_por_red"] = q(f"""
        SELECT 
            l.red,
            COUNT(DISTINCT l.id) AS total_lideres,
            COUNT(ci.id) AS total_informes,
            SUM(CASE WHEN ci.realizo THEN 1 ELSE 0 END) AS realizadas,
            SUM(CASE WHEN NOT ci.realizo THEN 1 ELSE 0 END) AS no_realizadas,
            COALESCE(SUM(ci.ofrenda), 0) AS total_ofrenda
        FROM lideres l
        JOIN hermanos h ON h.id = l.hermano_id
        LEFT JOIN celulas_informe ci ON ci.lider_id = l.id
        WHERE l.activo{fl}
        GROUP BY l.red
        ORDER BY l.red
    """, tuple(pl))
    # Resumen único por líder + fila TOTAL (con encuentros desglosados por etapa)
    r["resumen"] = q(f"""SELECT v.lider, v.red, v.tipo_12, v.rol_lider, v.estado_celula, v.cantidad_celulas,
        v.cantidad_macrocelulas,
        (SELECT COUNT(*) FROM lideres c WHERE c.lider_padre_id=v.lider_id AND c.activo) AS celulas_hijas,
        v.discipulos_activos,
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
        WHERE 1=1{fvv} ORDER BY v.discipulos_activos DESC""", tuple(pvv))
    tot = q(f"""SELECT COUNT(*) AS lideres, COALESCE(SUM(discipulos_activos),0) AS activos,
        COALESCE(SUM(total_asignados_vigentes),0) AS asignados,
        COALESCE(SUM(cantidad_macrocelulas),0) AS macro,
        (SELECT COUNT(*) FROM lideres l2 WHERE l2.activo AND l2.lider_padre_id IS NOT NULL{f0.replace('iglesia_id','l2.iglesia_id') if f0 else ''}) AS hijos
        FROM v_lideres_conteo{(' WHERE iglesia_id=%s' if f0 else '')}""", (tuple(p0) + tuple(p0)) if f0 else (), one=True)
    r["total"] = tot
    # Totales de encuentros para la fila TOTAL
    te = q(f"""SELECT COALESCE(SUM(CASE WHEN ep.estado='APROBADO' AND ep.hermano_id=l.hermano_id THEN 1 ELSE 0 END),0) AS enc_lideres,
        COALESCE(SUM(CASE WHEN ep.estado='APROBADO' THEN 1 ELSE 0 END),0) AS enc_disc
        FROM discipulado d JOIN lideres l ON l.id=d.lider_id
        LEFT JOIN encuentro_participacion ep ON ep.hermano_id=d.hermano_id
        WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo{fl}""", tuple(pl), one=True)
    r["total_enc"] = te
    r["padre-discipulos"] = q(f"SELECT lider, tipo_12, n_hijos, discipulos_activos FROM v_lideres_padre_con_discipulos WHERE n_hijos > 0 AND discipulos_activos > 0{f0} ORDER BY n_hijos DESC, discipulos_activos DESC LIMIT 50", tuple(p0))
    # Informes de UN líder específico (drill-down: ver todos los informes de un líder)
    r["_lideres_opts"] = q(f"""SELECT l.id, h.nombre_completo, l.red FROM lideres l
        JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl}
        ORDER BY h.nombre_completo""", tuple(pl))
    r["lider_info"] = None
    r["lider_detalle"] = []
    if ver == "lider_detalle" and lider_sel:
        r["lider_info"] = q(f"""{_SQL_PASTOR_DE}
            SELECT l.id, h.nombre_completo, l.red, l.estado_celula,
            COALESCE(hp.nombre_completo,'SIN PASTOR') AS pastor
            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
            LEFT JOIN pastor_de pd ON pd.lider_id = l.id
            LEFT JOIN lideres lp ON lp.id = pd.pastor_id
            LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
            WHERE l.id=%s{fl}""", tuple([lider_sel]) + tuple(pl), one=True)
        if r["lider_info"]:
            r["lider_detalle"] = q("""SELECT ci.fecha, ci.mes, ci.red, ci.realizo, ci.justificacion,
                ci.discipulos_asistieron, ci.discipulos_no_asistieron, ci.ofrenda,
                ci.tipo_ofrenda, ci.modalidad, ci.created_at,
                ht.nombre_completo AS timoteo
                FROM celulas_informe ci
                LEFT JOIN hermanos ht ON ht.id=ci.timoteo_id
                WHERE ci.lider_id=%s
                ORDER BY ci.fecha DESC NULLS LAST, ci.created_at DESC""", (lider_sel,))
    catalog = {"por_red": ("Reporte por red", ["red", "estado_celula", "c"]), "por_comuna": ("Reporte por comuna", ["region", "comuna", "c"]),
        "sin_celula": ("Sin celula", ["nombre_completo", "red"]), "pendientes": ("Pendientes", ["nombre_completo", "etapa", "fecha"]),
        "ruta_incompleta": ("Ruta incompleta", ["nombre_completo"]), "periodicidad": ("Periodicidad", ["nombre", "c"]),
        "por_12": ("Por Tipo 12", ["tipo_12", "c", "pastores"]),
        "resumen": ("Resumen por lider (unico + total)", ["lider", "red", "tipo_12", "rol_lider", "estado_celula", "celulas_hijas", "cantidad_macrocelulas", "discipulos_activos", "enc_lider", "d_tremendo", "d_fruto", "d_reenc", "d_crec", "d_lid"]),
        "padre-discipulos": ("Líderes con hijos + discipulos a cargo", ["lider", "tipo_12", "n_hijos", "discipulos_activos"]),
        # NUEVOS: Informes de Célula
        "celulas_resumen": ("Informe Célula · Por Líder", ["lider", "red", "pastor", "total_informes", "realizadas", "no_realizadas", "total_ofrenda", "ultima_fecha"]),
        "celulas_por_pastor": ("Informe Célula · Por Pastor", ["pastor", "red_pastor", "total_lideres", "total_informes", "realizadas", "no_realizadas", "total_ofrenda"]),
        "celulas_por_red": ("Informe Célula · Por Red", ["red", "total_lideres", "total_informes", "realizadas", "no_realizadas", "total_ofrenda"]),
        "lider_detalle": ("Informe Célula · Detalle de un líder", ["fecha", "mes", "timoteo", "realizo", "discipulos_asistieron", "discipulos_no_asistieron", "justificacion", "ofrenda", "tipo_ofrenda", "modalidad"])}
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
    return render_template("reportes.html", r=r, ver=ver, lider_sel=lider_sel)

# ---------- GRÁFICOS (opción de menú: solo gráficos) ----------
@app.get("/graficos")
@login_req
def graficos():
    """9 gráficos validados: 6 de barras, 1 de líneas, 1 circular y 1 de anillo."""
    from datetime import timedelta
    fh, ph = filtro_igl("h")
    fvv, pvv = filtro_igl("v")
    fci, pci = filtro_igl("ci")

    def _iso(s):
        try: return date.fromisoformat(s or "")
        except Exception: return None
    hoy = date.today()
    # Rango por defecto: mismo patrón "agosto–octubre" (inicios de mes, últimos 3 meses)
    mdef, adef = hoy.month - 2, hoy.year
    if mdef < 1: mdef += 12; adef -= 1
    mfin, afin = hoy.month + 1, hoy.year
    if mfin > 12: mfin = 1; afin += 1
    desde = _iso(request.args.get("desde")) or date(adef, mdef, 1)
    hasta = _iso(request.args.get("hasta")) or (date(afin, mfin, 1) - timedelta(days=1))
    if desde > hasta: desde, hasta = hasta, desde
    if (hasta - desde).days > 730: desde = hasta - timedelta(days=730)
    p_fec = (desde.isoformat(), hasta.isoformat())

    # 1) Discípulos que asistieron a célula — por red × mes (de los informes)
    g_asistencia = q(f"""SELECT ci.red AS red, to_char(ci.fecha,'YYYY-MM') AS mes,
        SUM(COALESCE(array_length(string_to_array(NULLIF(ci.discipulos_asistieron,''), ','), 1), 0)) AS c
        FROM celulas_informe ci
        WHERE ci.realizo AND ci.red IS NOT NULL AND ci.red <> ''
          AND ci.fecha BETWEEN %s AND %s{fci}
        GROUP BY 1,2 ORDER BY 1,2""", p_fec + tuple(pci))
    # 2) Informes por red — realizadas vs no (período)
    g_inf_red = q(f"""SELECT ci.red AS red,
        SUM(CASE WHEN ci.realizo THEN 1 ELSE 0 END) AS realizadas,
        SUM(CASE WHEN NOT ci.realizo THEN 1 ELSE 0 END) AS no_realizadas
        FROM celulas_informe ci
        WHERE ci.red IS NOT NULL AND ci.red <> '' AND ci.fecha BETWEEN %s AND %s{fci}
        GROUP BY 1 ORDER BY 1""", p_fec + tuple(pci))
    # 3) Informes por mes — realizadas vs no (período)
    g_inf_mes = q(f"""SELECT to_char(ci.fecha,'YYYY-MM') AS mes,
        SUM(CASE WHEN ci.realizo THEN 1 ELSE 0 END) AS realizadas,
        SUM(CASE WHEN NOT ci.realizo THEN 1 ELSE 0 END) AS no_realizadas
        FROM celulas_informe ci
        WHERE ci.fecha BETWEEN %s AND %s{fci}
        GROUP BY 1 ORDER BY 1""", p_fec + tuple(pci))
    # 4) Discípulos actuales — por red × tipo (foto, sin filtro de fechas)
    g_disc = q(f"""WITH hred AS (
        SELECT h.id, h.tipo_hermano,
          (SELECT l.red FROM discipulado d JOIN lideres l ON l.id=d.lider_id
             WHERE d.hermano_id=h.id AND d.fecha_fin IS NULL ORDER BY d.id DESC LIMIT 1) AS red
        FROM hermanos h WHERE h.activo{fh})
        SELECT COALESCE(NULLIF(red,''),'SIN RED') AS red,
               COALESCE(NULLIF(tipo_hermano,''),'SIN CLASIFICAR') AS tipo,
               COUNT(*) AS c FROM hred GROUP BY 1,2 ORDER BY 1,2""", tuple(ph))
    # Red de una visita (visitas no tiene columna red): líder responsable →
    # invitador (líder o su líder) → hermano integrado → SIN RED.
    _RED_VISITA = """COALESCE(
        NULLIF((SELECT l.red FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
                WHERE v.responsable IS NOT NULL AND h.nombre_completo=v.responsable AND l.activo
                ORDER BY l.id DESC LIMIT 1), ''),
        NULLIF((SELECT l.red FROM lideres l
                WHERE l.hermano_id=v.invitado_por_id AND l.activo ORDER BY l.id DESC LIMIT 1), ''),
        NULLIF((SELECT l.red FROM discipulado d JOIN lideres l ON l.id=d.lider_id
                WHERE d.hermano_id=v.invitado_por_id AND d.fecha_fin IS NULL ORDER BY d.id DESC LIMIT 1), ''),
        NULLIF((SELECT l.red FROM discipulado d JOIN lideres l ON l.id=d.lider_id
                WHERE d.hermano_id=v.hermano_id AND d.fecha_fin IS NULL ORDER BY d.id DESC LIMIT 1), ''),
        'SIN RED')"""

    # 5) Visitas recibidas — por red × estado (vacío hoy; se irá llenando)
    g_vis = q(f"""SELECT {_RED_VISITA} AS red,
        COALESCE(NULLIF(v.estado,''),'(SIN ESTADO)') AS estado, COUNT(*) AS c
        FROM visitas v
        WHERE COALESCE(v.fecha_visita, v.created_at::date) BETWEEN %s AND %s{fvv}
        GROUP BY 1,2 ORDER BY 1,2""", p_fec + tuple(pvv))
    # 6) Visitas NUEVAS en el rango — por red × mes (nuevos a la iglesia; vacío hoy)
    g_vis_nuevas = q(f"""SELECT {_RED_VISITA} AS red,
        to_char(COALESCE(v.fecha_visita, v.created_at::date),'YYYY-MM') AS mes,
        COUNT(*) AS c
        FROM visitas v
        WHERE COALESCE(v.fecha_visita, v.created_at::date) BETWEEN %s AND %s{fvv}
        GROUP BY 1,2 ORDER BY 1,2""", p_fec + tuple(pvv))

    def _grupo(rows, cat, ser, val):
        """Estructura Chart.js: categorías en X, una serie por valor de `ser`."""
        cats = sorted({str(x[cat]) for x in rows})
        sers = sorted({str(x[ser]) for x in rows})
        idx = {(str(x[cat]), str(x[ser])): int(x[val] or 0) for x in rows}
        return {"labels": cats,
                "datasets": [{"label": s, "data": [idx.get((c, s), 0) for c in cats]} for s in sers]}

    def _dos(rows, cat):
        """Estructura Chart.js: categorías en X + series Realizadas/No realizadas."""
        cats = sorted({str(x[cat]) for x in rows})
        idx = {str(x[cat]): x for x in rows}
        col = lambda k: [int(idx[c][k] or 0) if c in idx else 0 for c in cats]
        return {"labels": cats, "datasets": [
            {"label": "Realizadas", "data": col("realizadas")},
            {"label": "No realizadas", "data": col("no_realizadas")}]}

    def _circular_inf(rows):
        """Circular: total de informes (realizadas+no) por red en el período."""
        rows = sorted(rows, key=lambda x: str(x["red"]))
        return {"labels": [str(x["red"]) for x in rows],
                "datasets": [{"data": [int(x["realizadas"] or 0) + int(x["no_realizadas"] or 0)
                                       for x in rows]}]}

    def _anillo_tipos(rows):
        """Anillo: composición actual de discípulos por tipo (total por tipo)."""
        agg = {}
        for x in rows:
            agg[str(x["tipo"])] = agg.get(str(x["tipo"]), 0) + int(x["c"] or 0)
        labels = sorted(agg, key=lambda k: -agg[k])
        return {"labels": labels, "datasets": [{"data": [agg[k] for k in labels]}]}

    charts = {
        "asistencia":         _grupo(g_asistencia, "red", "mes", "c"),
        "informes_red":       _dos(g_inf_red, "red"),
        "informes_mes":       _dos(g_inf_mes, "mes"),
        "discipulos":         _grupo(g_disc, "red", "tipo", "c"),
        "visitas":            _grupo(g_vis, "red", "estado", "c"),
        "visitas_nuevas":     _grupo(g_vis_nuevas, "red", "mes", "c"),
        "evolucion":          _grupo(g_asistencia, "mes", "red", "c"),   # líneas
        "informes_circular":  _circular_inf(g_inf_red),                  # circular
        "discipulos_anillo":  _anillo_tipos(g_disc),                     # anillo
    }
    return render_template("graficos.html", graficos=charts,
                           desde=desde.isoformat(), hasta=hasta.isoformat())

# ---------- APIs (compatibilidad) ----------
@app.post("/api/login")
def api_login():
    data = request.get_json(force=True)
    u = (data.get("user") or "").strip()
    p = data.get("password") or ""
    ip = ip_cliente()
    # Mismo bloqueo que /login (si no, un bot usaría la API para saltarse el límite)
    mins = minutos_bloqueo(u, ip)
    if mins:
        registrar_intento(u, ip, False, "bloqueado")
        log_audit("LOGIN_BLOQUEADO", "usuarios", f"api bloqueado {u} ip {ip}")
        resp = make_response("Demasiados intentos fallidos", 429)
        resp.headers["Retry-After"] = str(int(mins * 60))   # segundos para reintentar
        return resp
    row = q("""SELECT u.id,u.password_hash,u.rol,u.iglesia_id,i.nombre AS iglesia_nombre,
        COALESCE(i.activo,true) AS iglesia_activa
        FROM usuarios u LEFT JOIN iglesias i ON i.id=u.iglesia_id
        WHERE u.username=%s AND u.activo""", (u,), one=True)
    if not row or not check_password_hash(row["password_hash"], p):
        registrar_intento(u, ip, False, "credenciales")
        log_audit("LOGIN_FALLIDO", "usuarios", f"api login fallido {u} ip {ip}")
        abort(401)
    if row["rol"] != "SUPERADMIN" and row.get("iglesia_activa") is False:
        abort(403)
    registrar_intento(u, ip, True, "api")
    session["uid"], session["rol"] = row["id"], row["rol"]
    session["iglesia_id"] = row["iglesia_id"] or 1
    session["iglesia_nombre"] = row["iglesia_nombre"] or "CASA MATRIZ"
    if row["rol"] == "SUPERADMIN":
        session["ver_todo"] = True
    else:
        session.pop("ver_todo", None)
    return {"ok": True, "rol": row["rol"]}

@app.get("/api/lideres")
def api_lideres():
    if "uid" not in session: abort(401)
    fi, pi = filtro_igl()
    return jsonify(q(f"SELECT * FROM v_lideres_conteo{(' WHERE iglesia_id=%s' if fi else '')} ORDER BY discipulos_activos DESC", tuple(pi)))

@app.get("/api/hermanos")
def api_hermanos():
    if "uid" not in session: abort(401)
    texto = f"%{request.args.get('q','')}%"
    fi, pi = filtro_igl()
    return jsonify(q(f"SELECT * FROM v_hermanos_completo WHERE nombre_completo ILIKE %s{fi} LIMIT 100", (texto,) + tuple(pi)))

# ---------- AUDITORIA ----------
@app.get("/auditoria")
@login_req
def auditoria():
    if session.get("rol") != "SUPERADMIN" and not tiene_permiso(session.get("rol",""), "auditoria", "leer", session.get("uid")):
        flash("No tiene permiso para ver auditoría", "error")
        return redirect(url_for("index"))
    texto = request.args.get("q","")
    fmod = request.args.get("fmod","")
    facc = request.args.get("facc","")
    fdesde = request.args.get("desde","")
    fhasta = request.args.get("hasta","")
    exp = request.args.get("export","")
    where = "WHERE (username ILIKE %s OR detalle ILIKE %s)"
    params = [f"%{texto}%", f"%{texto}%"]
    if fmod:
        where += " AND modulo=%s"
        params.append(fmod)
    if facc:
        where += " AND accion=%s"
        params.append(facc)
    if fdesde:
        where += " AND created_at::date >= %s"
        params.append(fdesde)
    if fhasta:
        where += " AND created_at::date <= %s"
        params.append(fhasta)
    # Aislamiento por iglesia (SUPERADMIN en TODAS ve todo)
    fa, pa = filtro_igl()
    where += fa
    params += pa
    # Convertir UTC a America/Santiago para mostrar hora local
    rows = q(f"""
        SELECT 
            id, usuario_id, username, rol, accion, modulo, detalle, ip,
            (created_at AT TIME ZONE 'America/Santiago') AS created_at_local
        FROM auditoria {where} ORDER BY created_at DESC LIMIT 500
    """, tuple(params))
    if exp == "excel":
        return to_excel(rows, ["created_at_local","username","rol","accion","modulo","detalle","ip"], "auditoria.xlsx")
    if exp == "pdf":
        return to_pdf("Auditoría MCCI", rows, ["created_at_local","username","accion","modulo"], "auditoria.pdf")
    mods = q("SELECT DISTINCT modulo FROM auditoria ORDER BY modulo")
    accs = q("SELECT DISTINCT accion FROM auditoria ORDER BY accion")
    return render_template("auditoria.html", rows=rows, texto=texto, fmod=fmod, facc=facc, fdesde=fdesde, fhasta=fhasta, mods=mods, accs=accs)

# ---------- CONFIGURACION ----------
CATALOGOS = [
    {"key":"redes","nombre":"Redes","icon":"🔴","tabla":"catalogo_redes","tiene_orden":False},
    {"key":"tipo12","nombre":"Tipo 12","icon":"🔢","tabla":"catalogo_tipo12","tiene_orden":False},
    {"key":"meses","nombre":"Meses","icon":"📅","tabla":"catalogo_meses","tiene_orden":True},
    {"key":"etapas","nombre":"Etapas Encuentro","icon":"📖","tabla":"tipos_encuentro","tiene_orden":True},
    {"key":"periodicidad","nombre":"Periodicidad","icon":"🔁","tabla":"periodicidad_celula","tiene_orden":False},
    {"key":"roles_lider","nombre":"Roles Líder","icon":"👤","tabla":"catalogo_roles_lider","tiene_orden":False},
    {"key":"estado_civil","nombre":"Estado Civil","icon":"💍","tabla":"catalogo_estado_civil","tiene_orden":False},
    {"key":"sexo","nombre":"Sexo","icon":"⚧","tabla":"catalogo_sexo","tiene_orden":False},
    {"key":"estado_celula","nombre":"Estado Célula","icon":"🏠","tabla":"catalogo_estado_celula","tiene_orden":False},
    {"key":"info_enviada","nombre":"Info Enviada","icon":"📬","tabla":"catalogo_info_enviada","tiene_orden":False},
    {"key":"estado_visita","nombre":"Estado Visita","icon":"🙋","tabla":"catalogo_estado_visita","tiene_orden":False},
    {"key":"tipo_seguimiento","nombre":"Tipo Seguimiento","icon":"📞","tabla":"catalogo_tipo_seguimiento","tiene_orden":False},
    {"key":"tipo_hermano","nombre":"Tipo Hermano","icon":"👤","tabla":"catalogo_tipo_hermano","tiene_orden":False},
    {"key":"modalidad_celula","nombre":"Modalidad Célula","icon":"🏠","tabla":"catalogo_modalidad_celula","tiene_orden":False},
    {"key":"roles_usuario","nombre":"Roles de Usuario","icon":"🛡️","tabla":"catalogo_roles_usuario","tiene_orden":False},
]

@app.get("/configuracion")
@login_req
@permiso_req("usuarios", "leer")
def configuracion():
    igl_scope, _ = iglesia_actual()  # None = SUPERADMIN viendo todo (edita global)
    cats = []
    for cat in CATALOGOS:
        cols = ["v.id","v.nombre","v.activo"]
        if cat["tiene_orden"]: cols.append("v.orden")
        if cat["tabla"] == "catalogo_sexo": cols = ["v.id","v.codigo","v.nombre","v.activo"]
        elif cat["tabla"] == "catalogo_roles_lider": cols = ["v.id","v.nombre","v.icono","v.activo"]
        order = "orden, nombre" if cat["tiene_orden"] else "nombre"
        order = order.replace("orden", "v.orden").replace("nombre", "v.nombre")
        if cat["tabla"] == "catalogo_roles_usuario":
            # Nombres de rol: globales a proposito
            items = q(f"SELECT {','.join(cols)} FROM {cat['tabla']} v ORDER BY {order}")
            for it in items: it["iglesia"] = "GLOBAL"
        else:
            cols.append("v.iglesia_id")
            cols.append("i.nombre AS iglesia")
            if igl_scope is None:
                items = q(f"SELECT {','.join(cols)} FROM {cat['tabla']} v LEFT JOIN iglesias i ON i.id=v.iglesia_id ORDER BY {order}")
            else:
                items = q(f"SELECT {','.join(cols)} FROM {cat['tabla']} v LEFT JOIN iglesias i ON i.id=v.iglesia_id WHERE (v.iglesia_id IS NULL OR v.iglesia_id=%s) ORDER BY {order}", (igl_scope,))
            for it in items:
                it["iglesia"] = it.get("iglesia") or "GLOBAL"
        cats.append({**cat, "valores": items})
    if igl_scope is None:
        users = q("""SELECT u.*, h.nombre_completo AS h_nombre, h2.nombre_completo AS l_nombre, i.nombre AS iglesia
            FROM usuarios u LEFT JOIN hermanos h ON h.id=u.hermano_id
            LEFT JOIN lideres l ON l.id=u.lider_id LEFT JOIN hermanos h2 ON h2.id=l.hermano_id
            LEFT JOIN iglesias i ON i.id=u.iglesia_id
            ORDER BY u.rol, u.username""")
    else:
        users = q("""SELECT u.*, h.nombre_completo AS h_nombre, h2.nombre_completo AS l_nombre, i.nombre AS iglesia
            FROM usuarios u LEFT JOIN hermanos h ON h.id=u.hermano_id
            LEFT JOIN lideres l ON l.id=u.lider_id LEFT JOIN hermanos h2 ON h2.id=l.hermano_id
            LEFT JOIN iglesias i ON i.id=u.iglesia_id
            WHERE u.iglesia_id=%s ORDER BY u.rol, u.username""", (igl_scope,))
    for u in users:
        u["vinculo"] = u.get("h_nombre") or u.get("l_nombre") or None
    perm_all = {}
    for rl in obtener_roles():
        perm_all[rl] = {}
        for mod in MODULOS_INFO:
            perm_all[rl][mod["key"]] = {}
            for acc in ("leer","crear","editar","eliminar","exportar"):
                if rl == "SUPERADMIN":
                    perm_all[rl][mod["key"]][acc] = True
                else:
                    perm_all[rl][mod["key"]][acc] = tiene_permiso(rl, mod["key"], acc, iglesia_id=igl_scope)
    if igl_scope is None:
        lideres = q("""SELECT l.id, h.nombre_completo, l.red, l.codigo_informe FROM lideres l
            JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo ORDER BY h.nombre_completo""")
    else:
        lideres = q("""SELECT l.id, h.nombre_completo, l.red, l.codigo_informe FROM lideres l
            JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo AND l.iglesia_id=%s ORDER BY h.nombre_completo""", (igl_scope,))
    user_perms = {}
    for u in users:
        uid = u["id"]
        up = q("SELECT modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar FROM usuario_permisos WHERE usuario_id=%s", (uid,))
        user_perms[uid] = {r["modulo"]: r for r in up}
    igls_cfg = q("SELECT id,nombre,tipo FROM iglesias WHERE activo ORDER BY id") if session.get("rol") == "SUPERADMIN" else []
    # Errores al guardar informes de célula (solo SUPERADMIN/ADMIN los revisan)
    errores_informes = []
    if session.get("rol") in ("SUPERADMIN", "ADMIN"):
        errores_informes = q("""SELECT a.id, a.created_at, a.username, a.detalle, a.ip, i.nombre AS iglesia
            FROM auditoria a LEFT JOIN iglesias i ON i.id=a.iglesia_id
            WHERE a.accion='ERROR_GUARDADO' AND a.modulo='celulas'
            ORDER BY a.created_at DESC LIMIT 100""")
    # ---- Ayudas / help editables (solo SUPERADMIN) ----
    ayudas_cfg = []
    if session.get("rol") == "SUPERADMIN":
        ovr = _ay_ovrs()
        paginas = {}
        for clave, d in escanear_ayudas().items():
            defecto = d["defecto"]
            if d["tipo"] == "bloque":
                titulo = "❓ Ayuda de la página"
            else:
                primera = re.split(r'[\n.]', defecto)[0][:55].strip()
                titulo = f"📝 Ayuda de campo — «{primera}…»"
            paginas.setdefault(d["plantilla"], []).append({
                "clave": clave, "tipo": d["tipo"], "titulo": titulo,
                "texto": ovr.get(clave, defecto), "defecto": defecto,
                "es_custom": clave in ovr, "vinetas": d["viñetas"],
                "dinamico": d["dinamico"],
            })
        for pag in sorted(paginas):
            items = sorted(paginas[pag], key=lambda x: 0 if x["tipo"] == "bloque" else 1)
            ayudas_cfg.append({
                "plantilla": pag, "nombre": pag[:-5].replace("_", " ").title(),
                "ayudas": items,
                "custom": sum(1 for i in items if i["es_custom"]),
                "vinetas": sum(i["vinetas"] for i in items),
            })
    return render_template("configuracion.html", catalogos=cats, usuarios=users, perm_all=perm_all, user_perms=user_perms, roles=obtener_roles(), modulos=MODULOS_INFO, lideres=lideres, iglesias_cfg=igls_cfg, errores_informes=errores_informes, ayudas_cfg=ayudas_cfg)

@app.post("/configuracion")
@login_req
@permiso_req("usuarios", "editar")
def configuracion_accion():
    f = request.form
    tabla = f.get("catalogo","")
    accion = f.get("accion","")
    item_id = f.get("item_id")
    nombre = f.get("nombre","").strip()
    codigo = f.get("codigo","").strip()
    orden = f.get("orden")

    # ---- Ayudas / help editables (solo SUPERADMIN) ----
    if tabla == "__ayudas__":
        return guardar_ayudas(f, accion)

    try:
        if tabla == "__usuarios__":
            if accion == "crear_usuario":
                username = f.get("username","").strip()
                password = f.get("password","").strip()
                rol = f.get("rol","CONSULTA")
                vinculo_id = f.get("vinculo_id")
                if not username or not password:
                    flash("Usuario y contraseña requeridos", "error")
                else:
                    from werkzeug.security import generate_password_hash
                    lider_id = int(vinculo_id) if vinculo_id else None
                    if lider_id and not exigir_igl("lideres", lider_id):
                        flash("Ese lider no es de tu iglesia", "error")
                    elif rol == "SUPERADMIN" and session.get("rol") != "SUPERADMIN":
                        flash("No puede crear SUPERADMIN", "error")
                    elif q("SELECT id FROM usuarios WHERE username=%s", (username,), one=True):
                        flash(f"El nombre '{username}' ya esta en uso (unico global). Prueba con otro.", "error")
                    else:
                        hermano_id = None
                        if lider_id:
                            ex = q("SELECT hermano_id FROM lideres WHERE id=%s", (lider_id,), one=True)
                            if ex: hermano_id = ex["hermano_id"]
                        if session.get("rol") == "SUPERADMIN" and f.get("iglesia_id"):
                            igl_n = int(f.get("iglesia_id"))
                        else:
                            igl_n = session.get("iglesia_id") or 1
                        q("INSERT INTO usuarios (username,password_hash,rol,hermano_id,lider_id,iglesia_id) VALUES (%s,%s,%s,%s,%s,%s)",
                          (username, generate_password_hash(password), rol, hermano_id, lider_id, igl_n), commit=True)
                        flash(f"Usuario '{username}' creado", "ok")

            elif accion == "toggle_usuario" and item_id:
                if not exigir_igl("usuarios", item_id):
                    flash("Ese usuario no es de tu iglesia", "error")
                else:
                    q("UPDATE usuarios SET activo = NOT activo WHERE id=%s", (item_id,), commit=True)
                    flash("Estado actualizado", "ok")

            elif accion == "eliminar_usuario" and item_id:
                if not exigir_igl("usuarios", item_id):
                    flash("Ese usuario no es de tu iglesia", "error")
                else:
                    q("DELETE FROM usuarios WHERE id=%s", (item_id,), commit=True)
                    log_audit("ELIMINAR","usuarios", f"id={item_id}")
                    flash("Usuario eliminado", "ok")

            elif accion == "editar_usuario" and item_id:
                if not exigir_igl("usuarios", item_id):
                    flash("Ese usuario no es de tu iglesia", "error")
                else:
                    nuevo_rol = f.get("nuevo_rol")
                    nuevo_vinculo = f.get("nuevo_vinculo")
                    if nuevo_rol == "SUPERADMIN" and session.get("rol") != "SUPERADMIN":
                        flash("Solo SUPERADMIN asigna rol SUPERADMIN", "error")
                        nuevo_rol = None
                    updates = []
                    params = []
                    if nuevo_rol:
                        updates.append("rol=%s")
                        params.append(nuevo_rol)
                    if nuevo_vinculo is not None:
                        lid = int(nuevo_vinculo) if nuevo_vinculo else None
                        hid = None
                        if lid and not exigir_igl("lideres", lid):
                            flash("Ese lider no es de tu iglesia", "error")
                            lid = None
                        if lid:
                            ex = q("SELECT hermano_id FROM lideres WHERE id=%s", (lid,), one=True)
                            if ex: hid = ex["hermano_id"]
                        updates.append("lider_id=%s")
                        params.append(lid)
                        updates.append("hermano_id=%s")
                        params.append(hid)
                    if updates:
                        params.append(item_id)
                        q(f"UPDATE usuarios SET {','.join(updates)} WHERE id=%s", tuple(params), commit=True)
                        log_audit("EDITAR","usuarios", f"id={item_id}")
                        flash("Usuario actualizado", "ok")

            elif accion == "guardar_permisos" and item_id:
                user_rol = f.get("user_rol","")
                if user_rol == "SUPERADMIN":
                    flash("No se pueden editar permisos de SUPERADMIN", "error")
                else:
                    # Alcance: SUPERADMIN en TODAS edita global; el resto, su iglesia
                    igl_s, _ = iglesia_actual()
                    scope = None if (session.get("rol") == "SUPERADMIN" and igl_s is None) else (session.get("iglesia_id") or 1)
                    q("DELETE FROM roles_permisos WHERE rol=%s AND iglesia_id IS NOT DISTINCT FROM %s", (user_rol, scope), commit=True)
                    for mod in MODULOS_INFO:
                        leer = f"perm_{mod['key']}_leer" in f
                        crear = f"perm_{mod['key']}_crear" in f
                        editar = f"perm_{mod['key']}_editar" in f
                        eliminar = f"perm_{mod['key']}_eliminar" in f
                        exportar = f"perm_{mod['key']}_exportar" in f
                        if leer or crear or editar or eliminar or exportar:
                            q("INSERT INTO roles_permisos (rol,modulo,puede_leer,puede_crear,puede_editar,puede_eliminar,puede_exportar,iglesia_id) VALUES(%s,%s,%s,%s,%s,%s,%s,%s)",
                              (user_rol, mod["key"], leer, crear, editar, eliminar, exportar, scope), commit=True)
                    flash(f"Permisos de {user_rol} actualizados ({'global' if scope is None else 'tu iglesia'})", "ok")

            elif accion == "guardar_permisos_usuario" and item_id:
                if not exigir_igl("usuarios", item_id):
                    flash("Ese usuario no es de tu iglesia", "error")
                else:
                    user_id = item_id
                    q("DELETE FROM usuario_permisos WHERE usuario_id=%s", (user_id,), commit=True)
                    for mod in MODULOS_INFO:
                        leer = f"uperm_{mod['key']}_leer" in f
                        crear = f"uperm_{mod['key']}_crear" in f
                        editar = f"uperm_{mod['key']}_editar" in f
                        eliminar = f"uperm_{mod['key']}_eliminar" in f
                        exportar = f"uperm_{mod['key']}_exportar" in f
                        if leer or crear or editar or eliminar or exportar:
                            q("INSERT INTO usuario_permisos (usuario_id,modulo,puede_leer,puede_crear,puede_editar,puede_eliminar,puede_exportar) VALUES(%s,%s,%s,%s,%s,%s,%s)",
                              (user_id, mod["key"], leer, crear, editar, eliminar, exportar), commit=True)
                    flash("Permisos individuales actualizados", "ok")

            elif accion == "reset_clave" and item_id:
                if not exigir_igl("usuarios", item_id):
                    flash("Ese usuario no es de tu iglesia", "error")
                else:
                    nueva = f.get("nueva_clave","").strip()
                    if not nueva:
                        flash("Debe ingresar una nueva contraseña", "error")
                    else:
                        err = validar_clave(nueva)
                        if err:
                            flash(f"La clave no cumple: {', '.join(err)}", "error")
                        else:
                            from werkzeug.security import generate_password_hash
                            q("UPDATE usuarios SET password_hash=%s, primer_login=TRUE WHERE id=%s", (generate_password_hash(nueva), item_id), commit=True)
                            log_audit("RESET_CLAVE","usuarios", f"id={item_id}")
                            flash(f"Contraseña actualizada (usuario id {item_id}) — deberá cambiarla al ingresar", "ok")

            # === GENERAR CÓDIGOS DE LÍDERES (SUPERADMIN y ADMIN) ===
            elif accion == "generar_codigos_lideres":
                if session.get("rol") not in ("SUPERADMIN", "ADMIN", "ADMIN_HIJA", "ADMIN_GRUPO"):
                    flash("Solo SUPERADMIN y ADMIN pueden generar códigos de líderes", "error")
                else:
                    modo = f.get("modo", "faltantes")  # "faltantes" | "todos"
                    fl, pl = filtro_igl("l")
                    import secrets, string
                    from openpyxl import Workbook
                    from io import BytesIO
                    from flask import send_file

                    if modo == "todos":
                        # Regenerar todos: limpiar y volver a crear
                        lideres = q(f"SELECT l.id, h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.activo{fl} ORDER BY h.nombre_completo", tuple(pl))
                        q(f"UPDATE lideres SET codigo_informe=NULL WHERE activo{fl}", tuple(pl), commit=True)
                    else:
                        # Solo los que no tienen código
                        lideres = q(f"""SELECT l.id, h.nombre_completo FROM lideres l 
                            JOIN hermanos h ON h.id=l.hermano_id 
                            WHERE l.activo AND (l.codigo_informe IS NULL OR l.codigo_informe=''){fl}
                            ORDER BY h.nombre_completo""", tuple(pl))

                    if not lideres:
                        flash("No hay líderes para procesar", "error")
                    else:
                        codigos_usados = set(r["codigo_informe"] for r in q("SELECT codigo_informe FROM lideres WHERE codigo_informe IS NOT NULL"))
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "Codigos_Lideres"
                        ws.append(["Líder ID", "Nombre", "Red", "Código (6 dígitos)"])

                        for l in lideres:
                            # Generar código único de 6 dígitos
                            while True:
                                codigo = ''.join(secrets.choice(string.digits) for _ in range(6))
                                if codigo not in codigos_usados:
                                    codigos_usados.add(codigo)
                                    break
                            q("UPDATE lideres SET codigo_informe=%s WHERE id=%s", (codigo, l["id"]), commit=True)
                            # Obtener red para el Excel
                            red = q("SELECT red FROM lideres WHERE id=%s", (l["id"],), one=True)
                            ws.append([l["id"], l["nombre_completo"], red["red"] if red else "", codigo])

                        buf = BytesIO()
                        wb.save(buf)
                        buf.seek(0)
                        flash(f"✅ {len(lideres)} códigos generados ({'TODOS' if modo=='todos' else 'solo faltantes'})", "ok")
                        return send_file(buf, as_attachment=True, download_name=f"codigos_lideres_{date.today().isoformat()}.xlsx",
                            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # === GENERAR CÓDIGO PARA UN SOLO LÍDER ===
            elif accion == "generar_codigo_un_lider":
                if session.get("rol") not in ("SUPERADMIN", "ADMIN", "ADMIN_HIJA", "ADMIN_GRUPO"):
                    flash("Solo SUPERADMIN y ADMIN pueden generar códigos de líderes", "error")
                else:
                    lider_id = f.get("lider_id", "").strip()
                    forzar = f.get("forzar") == "on"
                    if lider_id and not exigir_igl("lideres", lider_id):
                        flash("Ese lider no es de tu iglesia", "error")
                        lider_id = None
                    if not lider_id:
                        flash("Debe seleccionar un líder", "error")
                    else:
                        # Verificar si ya tiene código
                        actual = q("SELECT codigo_informe FROM lideres WHERE id=%s", (lider_id,), one=True)
                        if actual and actual["codigo_informe"] and not forzar:
                            flash(f"El líder ya tiene código ({actual['codigo_informe']}). Marque 'Forzar regenerar' para cambiarlo.", "error")
                        else:
                            import secrets, string
                            codigos_usados = set(r["codigo_informe"] for r in q("SELECT codigo_informe FROM lideres WHERE codigo_informe IS NOT NULL"))
                            while True:
                                codigo = ''.join(secrets.choice(string.digits) for _ in range(6))
                                if codigo not in codigos_usados:
                                    codigos_usados.add(codigo)
                                    break
                            q("UPDATE lideres SET codigo_informe=%s WHERE id=%s", (codigo, lider_id), commit=True)
                            nombre = q("SELECT h.nombre_completo FROM lideres l JOIN hermanos h ON h.id=l.hermano_id WHERE l.id=%s", (lider_id,), one=True)
                            flash(f"✅ Código generado para {nombre['nombre_completo']}: <strong>{codigo}</strong>", "ok")

            # === GENERAR CÓDIGOS POR RED ===
            elif accion == "generar_codigos_por_red":
                if session.get("rol") not in ("SUPERADMIN", "ADMIN", "ADMIN_HIJA", "ADMIN_GRUPO"):
                    flash("Solo SUPERADMIN y ADMIN pueden generar códigos de líderes", "error")
                else:
                    red = f.get("red", "").strip()
                    modo = f.get("modo", "faltantes")
                    if not red:
                        flash("Debe seleccionar una red", "error")
                    else:
                        import secrets, string
                        from openpyxl import Workbook
                        from io import BytesIO
                        from flask import send_file
                        from datetime import date

                        if modo == "todos":
                            flr, plr = filtro_igl("l")
                            lideres = q(f"""SELECT l.id, h.nombre_completo, l.red FROM lideres l
                                JOIN hermanos h ON h.id=l.hermano_id 
                                WHERE l.activo AND l.red=%s{flr} ORDER BY h.nombre_completo""", (red,) + tuple(plr))
                            q(f"UPDATE lideres SET codigo_informe=NULL WHERE activo AND red=%s{flr}", (red,) + tuple(plr), commit=True)
                        else:
                            flr, plr = filtro_igl("l")
                            lideres = q(f"""SELECT l.id, h.nombre_completo, l.red FROM lideres l 
                                JOIN hermanos h ON h.id=l.hermano_id 
                                WHERE l.activo AND l.red=%s AND (l.codigo_informe IS NULL OR l.codigo_informe=''){flr}
                                ORDER BY h.nombre_completo""", (red,) + tuple(plr))

                        if not lideres:
                            flash("No hay líderes en esa red para procesar", "error")
                        else:
                            codigos_usados = set(r["codigo_informe"] for r in q("SELECT codigo_informe FROM lideres WHERE codigo_informe IS NOT NULL"))
                            wb = Workbook()
                            ws = wb.active
                            ws.title = "Codigos_Lideres"
                            ws.append(["Líder ID", "Nombre", "Red", "Código (6 dígitos)"])

                            for l in lideres:
                                while True:
                                    codigo = ''.join(secrets.choice(string.digits) for _ in range(6))
                                    if codigo not in codigos_usados:
                                        codigos_usados.add(codigo)
                                        break
                                q("UPDATE lideres SET codigo_informe=%s WHERE id=%s", (codigo, l["id"]), commit=True)
                                ws.append([l["id"], l["nombre_completo"], l["red"], codigo])

                            buf = BytesIO()
                            wb.save(buf)
                            buf.seek(0)
                            flash(f"✅ {len(lideres)} códigos generados para red {red} ({'TODOS' if modo=='todos' else 'solo faltantes'})", "ok")
                            return send_file(buf, as_attachment=True, download_name=f"codigos_lideres_{red}_{date.today().isoformat()}.xlsx",
                                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # === EXPORTAR CÓDIGOS ACTUALES ===
            elif accion == "exportar_codigos_lideres":
                if session.get("rol") not in ("SUPERADMIN", "ADMIN", "ADMIN_HIJA", "ADMIN_GRUPO"):
                    flash("Solo SUPERADMIN y ADMIN pueden exportar códigos", "error")
                else:
                    filtro = f.get("filtro", "total")
                    flx, plx = filtro_igl("l")
                    import secrets, string
                    from openpyxl import Workbook
                    from io import BytesIO
                    from flask import send_file
                    from datetime import date

                    # Construir query según filtro
                    if filtro == "total":
                        query = f"""SELECT l.id, h.nombre_completo, l.red, l.codigo_informe,
                                   l.created_at, l.updated_at
                            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
                            WHERE l.activo{flx} ORDER BY l.red, h.nombre_completo"""
                        params = tuple(plx)
                    elif filtro == "con_codigo":
                        query = f"""SELECT l.id, h.nombre_completo, l.red, l.codigo_informe,
                                   l.created_at, l.updated_at
                            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
                            WHERE l.activo AND l.codigo_informe IS NOT NULL AND l.codigo_informe != ''{flx}
                            ORDER BY l.red, h.nombre_completo"""
                        params = tuple(plx)
                    elif filtro == "sin_codigo":
                        query = f"""SELECT l.id, h.nombre_completo, l.red, l.codigo_informe,
                                   l.created_at, l.updated_at
                            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
                            WHERE l.activo AND (l.codigo_informe IS NULL OR l.codigo_informe = ''){flx}
                            ORDER BY l.red, h.nombre_completo"""
                        params = tuple(plx)
                    elif filtro.startswith("red_"):
                        red = filtro.replace("red_", "")
                        query = f"""SELECT l.id, h.nombre_completo, l.red, l.codigo_informe,
                                   l.created_at, l.updated_at
                            FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
                            WHERE l.activo AND l.red=%s{flx} ORDER BY h.nombre_completo"""
                        params = (red,) + tuple(plx)
                    else:
                        flash("Filtro no válido", "error")
                        return redirect(url_for("configuracion"))

                    lideres = q(query, params)

                    if not lideres:
                        flash("No hay datos para exportar con ese filtro", "error")
                        return redirect(url_for("configuracion"))

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Codigos_Lideres"
                    ws.append(["Líder ID", "Nombre", "Red", "Código"])

                    for l in lideres:
                        ws.append([
                            l["id"],
                            l["nombre_completo"],
                            l["red"],
                            l["codigo_informe"] or ""
                        ])

                    buf = BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    flash(f"✅ {len(lideres)} registros exportados", "ok")
                    return send_file(buf, as_attachment=True,
                        download_name=f"codigos_lideres_{filtro}_{date.today().isoformat()}.xlsx",
                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        else:
            tablas_validas = [c["tabla"] for c in CATALOGOS]
            if tabla not in tablas_validas:
                flash("Catálogo no válido", "error")
                return redirect(url_for("configuracion"))

            # Alcance configuracion: SUPERADMIN en TODAS -> global (NULL); resto -> su iglesia
            igl_s, _ = iglesia_actual()
            if session.get("rol") == "SUPERADMIN" and igl_s is None:
                scope = None
            else:
                scope = session.get("iglesia_id") or 1
            if accion == "crear" and nombre:
                if tabla == "catalogo_roles_usuario" and session.get("rol") != "SUPERADMIN":
                    flash("Solo SUPERADMIN crea nombres de rol", "error")
                elif tabla == "catalogo_sexo":
                    q(f"INSERT INTO {tabla} (codigo, nombre, iglesia_id) VALUES (%s, %s, %s)", (codigo.upper(), nombre.upper(), scope), commit=True)
                elif tabla == "tipos_encuentro":
                    max_ord = q(f"SELECT COALESCE(MAX(orden),0)+1 AS m FROM {tabla} WHERE iglesia_id IS NOT DISTINCT FROM %s", (scope,), one=True)["m"]
                    q(f"INSERT INTO {tabla} (nombre, orden, iglesia_id) VALUES (%s, %s, %s)", (nombre.upper(), int(orden) if orden else max_ord, scope), commit=True)
                elif tabla == "periodicidad_celula":
                    q(f"INSERT INTO {tabla} (nombre, intervalo_dias, descripcion, iglesia_id) VALUES (%s, 30, %s, %s)", (nombre.upper(), nombre, scope), commit=True)
                elif tabla == "catalogo_roles_lider":
                    icono = f.get("icono","").strip() or "—"
                    q(f"INSERT INTO {tabla} (nombre, icono, iglesia_id) VALUES (%s, %s, %s)", (nombre.upper(), icono, scope), commit=True)
                elif tabla == "catalogo_roles_usuario":
                    q(f"INSERT INTO {tabla} (nombre) VALUES (%s)", (nombre.upper(),), commit=True)
                else:
                    q(f"INSERT INTO {tabla} (nombre, iglesia_id) VALUES (%s, %s)", (nombre.upper(), scope), commit=True)
                flash(f"Valor '{nombre}' agregado ({'global' if scope is None else 'tu iglesia'})", "ok")

            elif accion == "toggle" and item_id:
                _row = q(f"SELECT iglesia_id FROM {tabla} WHERE id=%s", (item_id,), one=True)
                if _row and _row.get("iglesia_id") is None and session.get("rol") != "SUPERADMIN":
                    flash("Valor global: solo SUPERADMIN lo modifica", "error")
                elif _row and _row.get("iglesia_id") is not None and _row.get("iglesia_id") != (session.get("iglesia_id") or 1) and session.get("rol") != "SUPERADMIN":
                    flash("Ese valor no es de tu iglesia", "error")
                else:
                    q(f"UPDATE {tabla} SET activo = NOT activo WHERE id=%s", (item_id,), commit=True)
                    flash("Estado actualizado", "ok")

            elif accion == "eliminar" and item_id:
                _row = q(f"SELECT iglesia_id FROM {tabla} WHERE id=%s", (item_id,), one=True)
                _rowigl = _row.get("iglesia_id") if _row else None
                if _row and _rowigl is None and session.get("rol") != "SUPERADMIN":
                    flash("Valor global: solo SUPERADMIN lo elimina", "error")
                    return redirect(url_for("configuracion"))
                if _row and _rowigl is not None and _rowigl != (session.get("iglesia_id") or 1) and session.get("rol") != "SUPERADMIN":
                    flash("Ese valor no es de tu iglesia", "error")
                    return redirect(url_for("configuracion"))
                # Alcance del fallback: fila global -> todas las iglesias; fila propia -> solo su iglesia
                _uf, _up = ("", []) if _rowigl is None else (" AND iglesia_id=%s", [_rowigl])
                if tabla in ("tipos_encuentro", "periodicidad_celula"):
                    q(f"UPDATE {tabla} SET activo=false WHERE id=%s", (item_id,), commit=True)
                    flash("Desactivado (no se puede eliminar por relaciones)", "ok")
                else:
                    # Si el valor está en uso, poner "NINGUNO" antes de eliminar (solo en su alcance)
                    try:
                        fallback_info = CATALOGO_FALLBACK.get(tabla)
                        if fallback_info:
                            # Obtener nombre del valor a eliminar
                            val_row = q(f"SELECT nombre FROM {tabla} WHERE id=%s", (item_id,), one=True)
                            if val_row:
                                val_nombre = val_row["nombre"]
                                # Para catalogo_sexo el nombre es descriptivo pero el valor real es codigo
                                if tabla == "catalogo_sexo":
                                    val_nombre = val_row.get("codigo", val_row["nombre"])
                                    # sexo usa codigo M/F
                                    for t, col, default in fallback_info:
                                        cnt = q(f"SELECT COUNT(*) AS c FROM {t} WHERE {col}=%s{_uf}", tuple([val_nombre] + _up), one=True)
                                        if cnt and cnt["c"] > 0:
                                            if default is None:
                                                q(f"UPDATE {t} SET {col}=NULL WHERE {col}=%s{_uf}", tuple([val_nombre] + _up), commit=True)
                                            else:
                                                q(f"UPDATE {t} SET {col}=%s WHERE {col}=%s{_uf}", tuple([default, val_nombre] + _up), commit=True)
                                else:
                                    for t, col, default in fallback_info:
                                        cnt = q(f"SELECT COUNT(*) AS c FROM {t} WHERE {col}=%s{_uf}", tuple([val_nombre] + _up), one=True)
                                        if cnt and cnt["c"] > 0:
                                            if default is None:
                                                q(f"UPDATE {t} SET {col}=NULL WHERE {col}=%s{_uf}", tuple([val_nombre] + _up), commit=True)
                                            else:
                                                q(f"UPDATE {t} SET {col}=%s WHERE {col}=%s{_uf}", tuple([default, val_nombre] + _up), commit=True)
                                            flash(f"{cnt['c']} registro(s) con '{val_nombre}' cambiados a '{default or 'Ninguno'}'", "ok")
                    except Exception as e:
                        print(f"Fallback eliminar catalogo: {e}")
                    q(f"DELETE FROM {tabla} WHERE id=%s", (item_id,), commit=True)
                    flash("Eliminado", "ok")

            elif accion == "editar_icono" and item_id:
                _row = q(f"SELECT iglesia_id FROM {tabla} WHERE id=%s", (item_id,), one=True)
                if _row and _row.get("iglesia_id") is None and session.get("rol") != "SUPERADMIN":
                    flash("Valor global: solo SUPERADMIN lo modifica", "error")
                elif _row and _row.get("iglesia_id") is not None and _row.get("iglesia_id") != (session.get("iglesia_id") or 1) and session.get("rol") != "SUPERADMIN":
                    flash("Ese valor no es de tu iglesia", "error")
                else:
                    nuevo_icono = f.get("icono","").strip() or "—"
                    q(f"UPDATE {tabla} SET icono=%s WHERE id=%s", (nuevo_icono, item_id), commit=True)
                    flash("Icono actualizado", "ok")

            elif accion == "renombrar" and item_id and nombre:
                _row = q(f"SELECT iglesia_id FROM {tabla} WHERE id=%s", (item_id,), one=True)
                if _row and _row.get("iglesia_id") is None and session.get("rol") != "SUPERADMIN":
                    flash("Valor global: solo SUPERADMIN lo renombra", "error")
                elif _row and _row.get("iglesia_id") is not None and _row.get("iglesia_id") != (session.get("iglesia_id") or 1) and session.get("rol") != "SUPERADMIN":
                    flash("Ese valor no es de tu iglesia", "error")
                else:
                    q(f"UPDATE {tabla} SET nombre=%s WHERE id=%s", (nombre.upper(), item_id), commit=True)
                    flash("Renombrado", "ok")

    except Exception as e:
        msg = str(e)
        if "duplicate key" in msg or "llave duplicada" in msg:
            flash(f"'{nombre}' ya existe en este catálogo", "error")
        else:
            flash(f"Error: {e}", "error")

    return redirect(url_for("configuracion"))

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5000")), debug=False)
