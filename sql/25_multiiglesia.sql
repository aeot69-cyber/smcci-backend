-- 25_multiiglesia.sql — MULTI-IGLESIA (Opción A: BD única, columna iglesia_id)
-- Fusiona 23_multiiglesia.sql + 24_config_por_iglesia.sql del prototipo SMCCI_IGLESIAS,
-- endurecido para producción Supabase. Ejecutar DESPUÉS de 24_codigo_informe.sql.
-- Idempotent: se puede ejecutar más de una vez sin dañar datos.
--
-- Modelo:
--   * tabla iglesias: MATRIZ (id=1 CASA MATRIZ, protegida) / HIJA / GRUPO, jerarquía por iglesia_padre_id.
--   * iglesia_id NOT NULL DEFAULT 1 en tablas de datos → todo lo existente queda en CASA MATRIZ.
--   * 14 catálogos + roles_permisos con iglesia_id NULL = valor GLOBAL (herencia de la matriz);
--     cada iglesia puede crear sus propios valores (fork al editar: solo su iglesia).
--   * roles ADMIN_HIJA / ADMIN_GRUPO (mismo poder que ADMIN; el alcance lo da iglesia_id).
--   * discipulado y encuentro_participacion NO llevan iglesia_id: se validan por JOIN
--     (lider→iglesia y hermano→iglesia) como en el prototipo probado (_IGL_VIA).

-- ============================ 1) TABLA IGLESIAS ============================
CREATE TABLE IF NOT EXISTS iglesias (
    id SERIAL PRIMARY KEY,
    nombre TEXT NOT NULL UNIQUE,
    tipo TEXT NOT NULL CHECK (tipo IN ('MATRIZ','HIJA','GRUPO')),
    iglesia_padre_id INT REFERENCES iglesias(id) ON DELETE RESTRICT,
    activo BOOLEAN NOT NULL DEFAULT TRUE,
    descripcion TEXT NOT NULL DEFAULT '',
    created_at TIMESTAMP NOT NULL DEFAULT NOW(),
    CHECK (id <> iglesia_padre_id)
);
INSERT INTO iglesias(nombre, tipo, descripcion)
VALUES ('CASA MATRIZ', 'MATRIZ', 'Iglesia matriz: centraliza la configuración global y los SUPERADMIN.')
ON CONFLICT (nombre) DO NOTHING;

-- ================= 2) iglesia_id EN TABLAS DE DATOS (DEFAULT 1) ================
-- NOTA: al agregar la columna con DEFAULT 1 NOT NULL, PostgreSQL rellena las
-- filas existentes con 1 (CASA MATRIZ) automáticamente. Backfill = gratis.
ALTER TABLE hermanos          ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);
ALTER TABLE lideres           ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);
ALTER TABLE usuarios          ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);
ALTER TABLE visitas           ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);
ALTER TABLE seguimiento       ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);
ALTER TABLE celulas_informe   ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);
ALTER TABLE auditoria         ADD COLUMN IF NOT EXISTS iglesia_id INT NOT NULL DEFAULT 1 REFERENCES iglesias(id);

-- Índices para que los filtros por iglesia no hagan seq-scan
CREATE INDEX IF NOT EXISTS idx_hermanos_iglesia        ON hermanos(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_lideres_iglesia         ON lideres(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_usuarios_iglesia        ON usuarios(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_visitas_iglesia         ON visitas(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_seguimiento_iglesia     ON seguimiento(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_celulas_informe_iglesia ON celulas_informe(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_auditoria_iglesia       ON auditoria(iglesia_id);
CREATE INDEX IF NOT EXISTS idx_lideres_codigo          ON lideres(codigo_informe) WHERE codigo_informe IS NOT NULL;

-- Los catálogos dejan de ser hardcodeados: cada iglesia define sus redes/meses/etc.
ALTER TABLE celulas_informe DROP CONSTRAINT IF EXISTS celulas_informe_red_check;
ALTER TABLE celulas_informe DROP CONSTRAINT IF EXISTS celulas_informe_mes_check;

-- ================ 3) ROLES ADMIN_HIJA / ADMIN_GRUPO ================
-- roles_permisos necesita iglesia_id antes de copiar permisos de ADMIN
ALTER TABLE roles_permisos ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
INSERT INTO catalogo_roles_usuario(nombre) VALUES ('ADMIN_HIJA'), ('ADMIN_GRUPO')
ON CONFLICT DO NOTHING;
INSERT INTO roles_permisos(rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar)
SELECT 'ADMIN_HIJA', modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar
FROM roles_permisos WHERE rol='ADMIN' AND iglesia_id IS NULL
ON CONFLICT DO NOTHING;
INSERT INTO roles_permisos(rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar)
SELECT 'ADMIN_GRUPO', modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar
FROM roles_permisos WHERE rol='ADMIN' AND iglesia_id IS NULL
ON CONFLICT DO NOTHING;

-- ================== 4) VISTAS CON iglesia_id (DROP + CREATE) ==================
DROP VIEW IF EXISTS v_hermanos_inactivos;
DROP VIEW IF EXISTS v_hermanos_completo;
CREATE VIEW v_hermanos_completo AS
SELECT h.id, h.codigo, h.rut, h.nombre_completo, h.nombres, h.apellido_paterno,
       h.apellido_materno, h.fecha_nacimiento, h.fecha_aniversario, h.correo,
       h.telefono, h.direccion, h.sexo, h.estado_civil, h.tipo_hermano, h.activo,
       h.fecha_ingreso, h.motivo_salida, h.fecha_baja, h.iglesia_id,
       c.nombre AS comuna, r.nombre_corto AS region, r.nombre AS region_larga
FROM hermanos h
LEFT JOIN comunas c ON c.id = h.comuna_id
LEFT JOIN regiones r ON r.id = c.region_id;
CREATE VIEW v_hermanos_inactivos AS
SELECT id, codigo, rut, nombre_completo, nombres, apellido_paterno,
       apellido_materno, fecha_nacimiento, fecha_aniversario, correo,
       telefono, direccion, sexo, estado_civil, tipo_hermano, activo,
       fecha_ingreso, motivo_salida, fecha_baja, iglesia_id,
       comuna, region, region_larga
FROM v_hermanos_completo
WHERE activo = false
ORDER BY fecha_baja DESC NULLS LAST, nombre_completo;

DROP VIEW IF EXISTS v_lideres_conteo;
CREATE VIEW v_lideres_conteo AS
SELECT l.id AS lider_id, h.codigo, h.nombre_completo AS lider,
       l.red, l.estado_celula, l.info_enviada, l.cantidad_celulas,
       l.rol_lider, l.tipo_12, l.lider_padre_id,
       hp.nombre_completo AS pastor_nombre,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL) AS total_asignados_vigentes,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo) AS discipulos_activos,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL AND NOT d.es_discipulo_activo) AS discipulos_inactivos,
       l.cantidad_macrocelulas, l.iglesia_id
FROM lideres l
  JOIN hermanos h ON h.id = l.hermano_id
  LEFT JOIN lideres lp ON lp.id = l.lider_padre_id
  LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
  LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.codigo, h.nombre_completo, l.red, l.estado_celula, l.info_enviada,
         l.cantidad_celulas, l.rol_lider, l.tipo_12, l.lider_padre_id,
         hp.nombre_completo, l.cantidad_macrocelulas, l.iglesia_id;

DROP VIEW IF EXISTS v_ruta_espiritual;
CREATE VIEW v_ruta_espiritual AS
SELECT h.id AS hermano_id, h.nombre_completo, h.iglesia_id,
       MAX(CASE WHEN t.nombre='ENCUENTRO TREMENDO' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS encuentro_tremendo,
       MAX(CASE WHEN t.nombre='FRUTO FIEL' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS fruto_fiel,
       MAX(CASE WHEN t.nombre='REENCUENTRO' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS reencuentro,
       MAX(CASE WHEN t.nombre='ESCUELA CRECIMIENTO' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS escuela_crecimiento,
       MAX(CASE WHEN t.nombre='ESCUELA LIDERES' AND e.estado='APROBADO' THEN 1 ELSE 0 END) AS escuela_lideres
FROM hermanos h
LEFT JOIN encuentro_participacion e ON e.hermano_id = h.id
LEFT JOIN tipos_encuentro t ON t.id = e.tipo_encuentro_id
GROUP BY h.id, h.nombre_completo, h.iglesia_id;

DROP VIEW IF EXISTS v_seguimiento_total;
DROP VIEW IF EXISTS v_sin_celula;
CREATE VIEW v_sin_celula AS
SELECT h.id, h.rut, h.nombre_completo, h.telefono, h.comuna_nombre, h.region,
       inv.nombre_completo AS invitado_por, h.invitado_por_texto, h.iglesia_id,
       (SELECT max(s.fecha) FROM seguimiento s WHERE s.hermano_id = h.id) AS ultimo_seg,
       (SELECT count(*) FROM seguimiento s WHERE s.hermano_id = h.id AND s.contactado) AS veces_contactado
FROM (SELECT h_1.id, h_1.rut, h_1.nombres, h_1.apellido_paterno, h_1.apellido_materno,
             h_1.nombre_completo, h_1.fecha_nacimiento, h_1.fecha_aniversario, h_1.correo,
             h_1.telefono, h_1.direccion, h_1.comuna_id, h_1.sexo, h_1.estado_civil,
             h_1.fecha_ingreso, h_1.activo, h_1.created_at, h_1.updated_at,
             h_1.invitado_por_id, h_1.invitado_por_texto, h_1.iglesia_id,
             c.nombre AS comuna_nombre, r.nombre_corto AS region
      FROM hermanos h_1
      LEFT JOIN comunas c ON c.id = h_1.comuna_id
      LEFT JOIN regiones r ON r.id = c.region_id) h
LEFT JOIN hermanos inv ON inv.id = h.invitado_por_id
LEFT JOIN discipulado d ON d.hermano_id = h.id AND d.fecha_fin IS NULL
LEFT JOIN lideres l ON l.hermano_id = h.id AND l.activo
WHERE h.activo AND d.id IS NULL AND l.id IS NULL
ORDER BY (SELECT max(s.fecha) FROM seguimiento s WHERE s.hermano_id = h.id) NULLS FIRST, h.nombre_completo;
CREATE VIEW v_seguimiento_total AS
SELECT 'HERMANO'::text AS tipo, h.id AS ref_id, h.nombre_completo, h.telefono,
       COALESCE(h.invitado_por, h.invitado_por_texto::text) AS invitado_por, h.iglesia_id,
       (SELECT max(s.fecha) FROM seguimiento s WHERE s.hermano_id = h.id) AS ultimo_seg
FROM v_sin_celula h
UNION ALL
SELECT 'VISITA'::text AS tipo, v.id AS ref_id, v.nombre_completo, v.telefono,
       COALESCE(inv.nombre_completo, v.invitado_por_texto::text) AS invitado_por, v.iglesia_id,
       (SELECT max(s.fecha) FROM seguimiento s WHERE s.visita_id = v.id) AS ultimo_seg
FROM visitas v
LEFT JOIN hermanos inv ON inv.id = v.invitado_por_id
WHERE v.estado::text = ANY (ARRAY['PENDIENTE'::character varying::text, 'EN SEGUIMIENTO'::character varying::text])
ORDER BY 7 NULLS FIRST, 3;

DROP VIEW IF EXISTS v_visitas_completo;
CREATE VIEW v_visitas_completo AS
SELECT v.id, v.nombre_completo, v.direccion, v.telefono, v.correo, v.comuna_id,
       v.invitado_por_id, v.invitado_por_texto, v.fecha_visita, v.motivo_oracion,
       v.estado, v.responsable, v.observacion, v.hermano_id, v.iglesia_id,
       v.created_at, v.updated_at,
       c.nombre AS comuna, inv.nombre_completo AS invitado_por
FROM visitas v
LEFT JOIN comunas c ON c.id = v.comuna_id
LEFT JOIN hermanos inv ON inv.id = v.invitado_por_id
ORDER BY v.fecha_visita DESC;

DROP VIEW IF EXISTS v_lideres_padre_con_discipulos;
CREATE VIEW v_lideres_padre_con_discipulos AS
SELECT l.id AS lider_id, h.codigo, h.nombre_completo AS lider,
       l.tipo_12, l.cantidad_celulas, l.iglesia_id,
       (SELECT COUNT(*) FROM lideres c
        WHERE c.lider_padre_id = l.id AND c.activo AND c.iglesia_id = l.iglesia_id) AS n_hijos,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo) AS discipulos_activos
FROM lideres l
JOIN hermanos h ON h.id = l.hermano_id
LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.codigo, h.nombre_completo, l.tipo_12, l.cantidad_celulas, l.iglesia_id;

-- ============ 5) CONFIG POR IGLESIA: CATÁLOGOS + PERMISOS (NULL = GLOBAL) ============
ALTER TABLE catalogo_redes            ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_tipo12           ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_meses            ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE tipos_encuentro           ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE periodicidad_celula       ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_roles_lider      ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_estado_civil     ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_sexo             ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_estado_celula    ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_info_enviada     ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_estado_visita    ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_tipo_seguimiento ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_tipo_hermano     ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE catalogo_modalidad_celula ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);
ALTER TABLE roles_permisos            ADD COLUMN IF NOT EXISTS iglesia_id INT REFERENCES iglesias(id);

-- Unicidad por ámbito: mismo nombre puede existir en global y en cada iglesia
ALTER TABLE catalogo_redes DROP CONSTRAINT IF EXISTS catalogo_redes_nombre_key;
ALTER TABLE catalogo_redes DROP CONSTRAINT IF EXISTS catalogo_redes_nombre_iglesia_key;
ALTER TABLE catalogo_redes ADD CONSTRAINT catalogo_redes_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_tipo12 DROP CONSTRAINT IF EXISTS catalogo_tipo12_nombre_key;
ALTER TABLE catalogo_tipo12 DROP CONSTRAINT IF EXISTS catalogo_tipo12_nombre_iglesia_key;
ALTER TABLE catalogo_tipo12 ADD CONSTRAINT catalogo_tipo12_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_meses DROP CONSTRAINT IF EXISTS catalogo_meses_nombre_key;
ALTER TABLE catalogo_meses DROP CONSTRAINT IF EXISTS catalogo_meses_nombre_iglesia_key;
ALTER TABLE catalogo_meses ADD CONSTRAINT catalogo_meses_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE tipos_encuentro DROP CONSTRAINT IF EXISTS tipos_encuentro_nombre_key;
ALTER TABLE tipos_encuentro DROP CONSTRAINT IF EXISTS tipos_encuentro_nombre_iglesia_key;
ALTER TABLE tipos_encuentro ADD CONSTRAINT tipos_encuentro_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE tipos_encuentro DROP CONSTRAINT IF EXISTS tipos_encuentro_orden_key;
ALTER TABLE tipos_encuentro DROP CONSTRAINT IF EXISTS tipos_encuentro_orden_iglesia_key;
ALTER TABLE tipos_encuentro ADD CONSTRAINT tipos_encuentro_orden_iglesia_key UNIQUE NULLS NOT DISTINCT (orden, iglesia_id);
ALTER TABLE periodicidad_celula DROP CONSTRAINT IF EXISTS periodicidad_celula_nombre_key;
ALTER TABLE periodicidad_celula DROP CONSTRAINT IF EXISTS periodicidad_celula_nombre_iglesia_key;
ALTER TABLE periodicidad_celula ADD CONSTRAINT periodicidad_celula_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_roles_lider DROP CONSTRAINT IF EXISTS catalogo_roles_lider_nombre_key;
ALTER TABLE catalogo_roles_lider DROP CONSTRAINT IF EXISTS catalogo_roles_lider_nombre_iglesia_key;
ALTER TABLE catalogo_roles_lider ADD CONSTRAINT catalogo_roles_lider_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_estado_civil DROP CONSTRAINT IF EXISTS catalogo_estado_civil_nombre_key;
ALTER TABLE catalogo_estado_civil DROP CONSTRAINT IF EXISTS catalogo_estado_civil_nombre_iglesia_key;
ALTER TABLE catalogo_estado_civil ADD CONSTRAINT catalogo_estado_civil_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_sexo DROP CONSTRAINT IF EXISTS catalogo_sexo_codigo_key;
ALTER TABLE catalogo_sexo DROP CONSTRAINT IF EXISTS catalogo_sexo_codigo_iglesia_key;
ALTER TABLE catalogo_sexo ADD CONSTRAINT catalogo_sexo_codigo_iglesia_key UNIQUE NULLS NOT DISTINCT (codigo, iglesia_id);
ALTER TABLE catalogo_estado_celula DROP CONSTRAINT IF EXISTS catalogo_estado_celula_nombre_key;
ALTER TABLE catalogo_estado_celula DROP CONSTRAINT IF EXISTS catalogo_estado_celula_nombre_iglesia_key;
ALTER TABLE catalogo_estado_celula ADD CONSTRAINT catalogo_estado_celula_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_info_enviada DROP CONSTRAINT IF EXISTS catalogo_info_enviada_nombre_key;
ALTER TABLE catalogo_info_enviada DROP CONSTRAINT IF EXISTS catalogo_info_enviada_nombre_iglesia_key;
ALTER TABLE catalogo_info_enviada ADD CONSTRAINT catalogo_info_enviada_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_estado_visita DROP CONSTRAINT IF EXISTS catalogo_estado_visita_nombre_key;
ALTER TABLE catalogo_estado_visita DROP CONSTRAINT IF EXISTS catalogo_estado_visita_nombre_iglesia_key;
ALTER TABLE catalogo_estado_visita ADD CONSTRAINT catalogo_estado_visita_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_tipo_seguimiento DROP CONSTRAINT IF EXISTS catalogo_tipo_seguimiento_nombre_key;
ALTER TABLE catalogo_tipo_seguimiento DROP CONSTRAINT IF EXISTS catalogo_tipo_seguimiento_nombre_iglesia_key;
ALTER TABLE catalogo_tipo_seguimiento ADD CONSTRAINT catalogo_tipo_seguimiento_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_tipo_hermano DROP CONSTRAINT IF EXISTS catalogo_tipo_hermano_nombre_key;
ALTER TABLE catalogo_tipo_hermano DROP CONSTRAINT IF EXISTS catalogo_tipo_hermano_nombre_iglesia_key;
ALTER TABLE catalogo_tipo_hermano ADD CONSTRAINT catalogo_tipo_hermano_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE catalogo_modalidad_celula DROP CONSTRAINT IF EXISTS catalogo_modalidad_celula_nombre_key;
ALTER TABLE catalogo_modalidad_celula DROP CONSTRAINT IF EXISTS catalogo_modalidad_celula_nombre_iglesia_key;
ALTER TABLE catalogo_modalidad_celula ADD CONSTRAINT catalogo_modalidad_celula_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id);
ALTER TABLE roles_permisos DROP CONSTRAINT IF EXISTS roles_permisos_rol_modulo_key;
ALTER TABLE roles_permisos DROP CONSTRAINT IF EXISTS roles_permisos_rol_modulo_iglesia_key;
ALTER TABLE roles_permisos ADD CONSTRAINT roles_permisos_rol_modulo_iglesia_key UNIQUE NULLS NOT DISTINCT (rol, modulo, iglesia_id);

COMMENT ON TABLE iglesias IS 'Sedes: CASA MATRIZ (id=1, protegida), iglesias hijas y grupos de crecimiento (jerarquia por iglesia_padre_id)';
COMMENT ON COLUMN iglesias.tipo IS 'MATRIZ | HIJA | GRUPO';
COMMENT ON COLUMN catalogo_redes.iglesia_id IS 'NULL = valor global/herencia de la matriz; con valor = config propia de la iglesia';
COMMENT ON COLUMN roles_permisos.iglesia_id IS 'NULL = permisos globales del rol; con valor = override por iglesia';
