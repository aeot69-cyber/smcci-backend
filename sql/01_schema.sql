-- =============================================
-- SMCCI - Sistema Consolidación MCCI
-- Motor: PostgreSQL 14+
-- Diseño 3FN, todas las tablas relacionadas vía FK
-- Orden de creación respeta dependencias
-- =============================================

-- Extensiones
CREATE EXTENSION IF NOT EXISTS "pgcrypto";
CREATE EXTENSION IF NOT EXISTS "citext";

-- ---------- 1. REGIONES ----------
CREATE TABLE IF NOT EXISTS regiones (
  id SERIAL PRIMARY KEY,
  codigo VARCHAR(5) UNIQUE NOT NULL,          -- Ej: 'RM','V','VIII'
  nombre VARCHAR(100) UNIQUE NOT NULL,        -- Ej: 'Región Metropolitana de Santiago'
  nombre_corto VARCHAR(50) NOT NULL,          -- Ej: 'Metropolitana'
  numeral VARCHAR(10),                        -- Ej: 'XIII','V'
  created_at TIMESTAMPTZ DEFAULT now()
);

-- ---------- 2. COMUNAS ----------
CREATE TABLE IF NOT EXISTS comunas (
  id SERIAL PRIMARY KEY,
  nombre VARCHAR(100) NOT NULL,
  region_id INT NOT NULL REFERENCES regiones(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  codigo_postal VARCHAR(10),
  UNIQUE (nombre, region_id),
  created_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_comunas_region ON comunas(region_id);
CREATE INDEX IF NOT EXISTS idx_comunas_nombre ON comunas(nombre);

-- ---------- 3. PERIODICIDAD CELULA ----------
CREATE TABLE IF NOT EXISTS periodicidad_celula (
  id SERIAL PRIMARY KEY,
  nombre VARCHAR(50) UNIQUE NOT NULL,   -- SEMANAL, QUINCENAL, MENSUAL
  intervalo_dias INT NOT NULL CHECK (intervalo_dias > 0),
  descripcion TEXT
);

-- ---------- 4. TIPOS ENCUENTRO ----------
CREATE TABLE IF NOT EXISTS tipos_encuentro (
  id SERIAL PRIMARY KEY,
  nombre VARCHAR(80) UNIQUE NOT NULL,  -- ENCUENTRO TREMENDO, FRUTO FIEL, REENCUENTRO, ESCUELA CRECIMIENTO, ESCUELA LIDERES
  descripcion TEXT,
  orden INT NOT NULL DEFAULT 0,        -- orden correlativo del proceso
  UNIQUE (orden)
);

-- ---------- 5. HERMANOS (Tabla Maestra) ----------
CREATE TABLE IF NOT EXISTS hermanos (
  id SERIAL PRIMARY KEY,
  rut VARCHAR(12) UNIQUE NOT NULL,  -- formato 12345678-9 (con guion, K mayúscula)
  nombres VARCHAR(100) NOT NULL,
  apellido_paterno VARCHAR(80) NOT NULL,
  apellido_materno VARCHAR(80),
  nombre_completo TEXT GENERATED ALWAYS AS (
    nombres || ' ' || apellido_paterno || COALESCE(' ' || apellido_materno,'')
  ) STORED,
  fecha_nacimiento DATE CHECK (fecha_nacimiento <= CURRENT_DATE),
  fecha_aniversario DATE,             -- matrimonio / aniversario espiritual
  correo CITEXT UNIQUE,               -- citext = case-insensitive
  telefono VARCHAR(20),
  direccion VARCHAR(200),
  comuna_id INT REFERENCES comunas(id) ON UPDATE CASCADE ON DELETE RESTRICT,
   sexo CHAR(1),
   estado_civil VARCHAR(20),
  fecha_ingreso DATE DEFAULT CURRENT_DATE,
  activo BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now(),
  CONSTRAINT chk_rut_formato CHECK (rut ~ '^[0-9]{7,8}-[0-9kK]$'),
  CONSTRAINT chk_correo_formato CHECK (correo IS NULL OR correo ~ '^[^@\s]+@[^@\s]+\.[^@\s]+$')
);
CREATE INDEX IF NOT EXISTS idx_hermanos_rut ON hermanos(rut);
CREATE INDEX IF NOT EXISTS idx_hermanos_nombre ON hermanos(nombre_completo);
CREATE INDEX IF NOT EXISTS idx_hermanos_comuna ON hermanos(comuna_id);
CREATE INDEX IF NOT EXISTS idx_hermanos_activo ON hermanos(activo);

-- ---------- 6. LIDERES (Tabla Maestra, 1:1 con hermanos) ----------
CREATE TABLE IF NOT EXISTS lideres (
  id SERIAL PRIMARY KEY,
  hermano_id INT UNIQUE NOT NULL REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  -- Un líder ES un hermano. UNIQUE garantiza 1:1.
  red VARCHAR(20) NOT NULL DEFAULT 'MIXTO',
  estado_celula VARCHAR(20) NOT NULL DEFAULT 'EN FORMACION',
  info_enviada VARCHAR(20) NOT NULL DEFAULT 'EN CONSULTA',
  cantidad_celulas INT NOT NULL DEFAULT 1 CHECK (cantidad_celulas >= 0),
  -- NOTA: discípulos activos NO se almacenan (se calculan en vista v_lideres_conteo)
  -- para evitar inconsistencia. Si exige columna histórica, usar vista materializada.
  fecha_nombramiento DATE DEFAULT CURRENT_DATE,
  activo BOOLEAN DEFAULT TRUE,
  observacion TEXT,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_lideres_red ON lideres(red);
CREATE INDEX IF NOT EXISTS idx_lideres_estado ON lideres(estado_celula);

-- ---------- 7. DISCIPULADO / LIDER_ACTUALIZA ----------
-- Asigna cada hermano a UN líder vigente (histórico conservado con fecha_fin).
CREATE TABLE IF NOT EXISTS discipulado (
  id SERIAL PRIMARY KEY,
  hermano_id INT NOT NULL REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  lider_id INT NOT NULL REFERENCES lideres(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  es_discipulo_activo BOOLEAN NOT NULL DEFAULT TRUE,
  periodicidad_id INT REFERENCES periodicidad_celula(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  fecha_asignacion DATE NOT NULL DEFAULT CURRENT_DATE,
  fecha_fin DATE CHECK (fecha_fin IS NULL OR fecha_fin >= fecha_asignacion),
  observacion TEXT,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now(),
  CONSTRAINT chk_no_autolider CHECK (True) -- validación lógica se hace en trigger (hermano<>líder hermano)
);
-- Solo UNA asignación vigente por hermano (histórico permitido cerrando fecha_fin)
CREATE UNIQUE INDEX IF NOT EXISTS uq_discipulado_vigente
  ON discipulado(hermano_id) WHERE fecha_fin IS NULL;
CREATE INDEX IF NOT EXISTS idx_discipulado_lider ON discipulado(lider_id);
CREATE INDEX IF NOT EXISTS idx_discipulado_activo ON discipulado(es_discipulo_activo);

-- ---------- 8. ENCUENTRO (participación por hermano) ----------
CREATE TABLE IF NOT EXISTS encuentro_participacion (
  id SERIAL PRIMARY KEY,
  hermano_id INT NOT NULL REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE CASCADE,
  tipo_encuentro_id INT NOT NULL REFERENCES tipos_encuentro(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  fecha_evento DATE NOT NULL DEFAULT CURRENT_DATE,
  estado VARCHAR(20) NOT NULL DEFAULT 'PENDIENTE'
    CHECK (estado IN ('PENDIENTE','APROBADO','REPROBADO','NO ASISTE')),
  lider_validador_id INT REFERENCES lideres(id) ON UPDATE CASCADE ON DELETE SET NULL,
  observacion TEXT,
  created_at TIMESTAMPTZ DEFAULT now(),
  UNIQUE (hermano_id, tipo_encuentro_id)  -- un registro por etapa; reintento = UPDATE estado
);
CREATE INDEX IF NOT EXISTS idx_encuentro_hermano ON encuentro_participacion(hermano_id);
CREATE INDEX IF NOT EXISTS idx_encuentro_tipo ON encuentro_participacion(tipo_encuentro_id);

-- ---------- 9. USUARIOS WEB (seguridad, separado de hermanos) ----------
CREATE TABLE IF NOT EXISTS usuarios (
  id SERIAL PRIMARY KEY,
  username CITEXT UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,  -- bcrypt/argon2 generado desde la app, NUNCA texto plano
  rol VARCHAR(20) NOT NULL DEFAULT 'CONSULTA',
  hermano_id INT REFERENCES hermanos(id) ON DELETE SET NULL,
  lider_id INT REFERENCES lideres(id) ON DELETE SET NULL,
  activo BOOLEAN DEFAULT TRUE,
  ultimo_login TIMESTAMPTZ,
  created_at TIMESTAMPTZ DEFAULT now()
);

-- ---------- FUNCIONES / TRIGGERS ----------
-- updated_at automático
CREATE OR REPLACE FUNCTION fn_touch_updated_at() RETURNS TRIGGER AS $$
BEGIN NEW.updated_at = now(); RETURN NEW; END; $$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_hermanos_touch ON hermanos;
CREATE TRIGGER trg_hermanos_touch BEFORE UPDATE ON hermanos
  FOR EACH ROW EXECUTE FUNCTION fn_touch_updated_at();
DROP TRIGGER IF EXISTS trg_lideres_touch ON lideres;
CREATE TRIGGER trg_lideres_touch BEFORE UPDATE ON lideres
  FOR EACH ROW EXECUTE FUNCTION fn_touch_updated_at();
DROP TRIGGER IF EXISTS trg_discipulado_touch ON discipulado;
CREATE TRIGGER trg_discipulado_touch BEFORE UPDATE ON discipulado
  FOR EACH ROW EXECUTE FUNCTION fn_touch_updated_at();

-- Validar dígito verificador RUT chileno
CREATE OR REPLACE FUNCTION fn_validar_rut(rut TEXT) RETURNS BOOLEAN AS $$
DECLARE cuerpo TEXT; dv CHAR(1); suma INT:=0; mult INT:=2; i INT; resto INT; dv_esp CHAR(1);
BEGIN
  IF rut !~ '^[0-9]{7,8}-[0-9kK]$' THEN RETURN FALSE; END IF;
  cuerpo := split_part(rut,'-',1); dv := upper(split_part(rut,'-',2));
  FOR i IN REVERSE length(cuerpo)..1 LOOP
    suma := suma + substring(cuerpo,i,1)::INT * mult;
    mult := CASE WHEN mult=7 THEN 2 ELSE mult+1 END;
  END LOOP;
  resto := 11 - (suma % 11);
  dv_esp := CASE resto WHEN 11 THEN '0' WHEN 10 THEN 'K' ELSE resto::TEXT END;
  RETURN dv = dv_esp;
END; $$ LANGUAGE plpgsql IMMUTABLE;

-- Evitar que un líder sea discípulo de sí mismo
CREATE OR REPLACE FUNCTION fn_discipulado_no_autolider() RETURNS TRIGGER AS $$
DECLARE h_lider INT;
BEGIN
  SELECT hermano_id INTO h_lider FROM lideres WHERE id = NEW.lider_id;
  IF h_lider = NEW.hermano_id THEN
    RAISE EXCEPTION 'Un hermano no puede ser discípulo de sí mismo (hermano_id=%)', NEW.hermano_id;
  END IF;
  RETURN NEW;
END; $$ LANGUAGE plpgsql;
DROP TRIGGER IF EXISTS trg_discipulado_noauto ON discipulado;
CREATE TRIGGER trg_discipulado_noauto BEFORE INSERT OR UPDATE ON discipulado
  FOR EACH ROW EXECUTE FUNCTION fn_discipulado_no_autolider();

-- Si líder queda SIN CELULA, sus discípulos vigentes pasan a inactivos (opcional, consistente)
-- Se deja como función manual, no trigger destructivo:
-- UPDATE discipulado SET es_discipulo_activo=false WHERE lider_id=?;

-- ---------- VISTAS (conteos siempre consistentes) ----------
-- Detalle hermano + comuna + región
CREATE OR REPLACE VIEW v_hermanos_completo AS
SELECT h.id, h.rut, h.nombre_completo, h.nombres, h.apellido_paterno, h.apellido_materno,
       h.fecha_nacimiento, h.fecha_aniversario, h.correo, h.telefono, h.direccion,
       h.sexo, h.estado_civil, h.activo, h.fecha_ingreso,
       c.nombre AS comuna, r.nombre_corto AS region, r.nombre AS region_larga
FROM hermanos h
LEFT JOIN comunas c ON c.id = h.comuna_id
LEFT JOIN regiones r ON r.id = c.region_id;

-- Lideres con conteo REAL de discípulos y células (no se desincroniza)
CREATE OR REPLACE VIEW v_lideres_conteo AS
SELECT l.id AS lider_id, h.nombre_completo AS lider,
       l.red, l.estado_celula, l.info_enviada, l.cantidad_celulas,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL) AS total_asignados_vigentes,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo) AS discipulos_activos,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL AND NOT d.es_discipulo_activo) AS discipulos_inactivos
FROM lideres l
JOIN hermanos h ON h.id = l.hermano_id
LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.nombre_completo, l.red, l.estado_celula, l.info_enviada, l.cantidad_celulas;

-- Ruta espiritual por hermano (qué encuentros tiene aprobados)
CREATE OR REPLACE VIEW v_ruta_espiritual AS
SELECT h.id AS hermano_id, h.nombre_completo,
       MAX(CASE WHEN t.nombre='ENCUENTRO TREMENDO' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS encuentro_tremendo,
       MAX(CASE WHEN t.nombre='FRUTO FIEL' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS fruto_fiel,
       MAX(CASE WHEN t.nombre='REENCUENTRO' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS reencuentro,
       MAX(CASE WHEN t.nombre='ESCUELA CRECIMIENTO' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS escuela_crecimiento,
       MAX(CASE WHEN t.nombre='ESCUELA LIDERES' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS escuela_lideres
FROM hermanos h
LEFT JOIN encuentro_participacion e ON e.hermano_id = h.id
LEFT JOIN tipos_encuentro t ON t.id = e.tipo_encuentro_id
GROUP BY h.id, h.nombre_completo;
