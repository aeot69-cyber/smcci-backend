-- Migración: RUT → Código automático alfanumérico
-- El sistema genera código MCCI0001, MCCI0002, etc. al crear hermano/líder/visita integrada
-- RUT deja de ser obligatorio (se mantiene columna por compatibilidad pero nullable y sin CHECK)

CREATE SEQUENCE IF NOT EXISTS seq_codigo START 1;

ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS codigo VARCHAR(10) UNIQUE;
ALTER TABLE hermanos ALTER COLUMN rut DROP NOT NULL;
ALTER TABLE hermanos DROP CONSTRAINT IF EXISTS chk_rut_formato;
-- Mantener UNIQUE en rut pero permitir múltiples NULL (Postgres lo permite)
-- Si existe constraint de unique con nombre distinto, no tocar

-- Backfill códigos para hermanos existentes sin código
DO $$
DECLARE
  r RECORD;
  seq_val INT;
BEGIN
  FOR r IN SELECT id FROM hermanos WHERE codigo IS NULL ORDER BY id LOOP
    SELECT nextval('seq_codigo') INTO seq_val;
    UPDATE hermanos SET codigo = 'MCCI' || LPAD(seq_val::text, 4, '0') WHERE id = r.id;
  END LOOP;
END $$;

-- Ajustar secuencia al máximo actual
SELECT setval('seq_codigo', COALESCE((SELECT MAX(CAST(SUBSTRING(codigo FROM 5) AS INT)) FROM hermanos), 0));

-- Hacer código NOT NULL después de backfill
ALTER TABLE hermanos ALTER COLUMN codigo SET NOT NULL;

-- Recrear vistas para incluir código
DROP VIEW IF EXISTS v_hermanos_completo;
CREATE VIEW v_hermanos_completo AS
 SELECT h.id,
    h.codigo,
    h.rut,
    h.nombre_completo,
    h.nombres,
    h.apellido_paterno,
    h.apellido_materno,
    h.fecha_nacimiento,
    h.fecha_aniversario,
    h.correo,
    h.telefono,
    h.direccion,
    h.sexo,
    h.estado_civil,
    h.activo,
    h.fecha_ingreso,
    c.nombre AS comuna,
    r.nombre_corto AS region,
    r.nombre AS region_larga
   FROM ((hermanos h
     LEFT JOIN comunas c ON ((c.id = h.comuna_id)))
     LEFT JOIN regiones r ON ((r.id = c.region_id)));

DROP VIEW IF EXISTS v_lideres_conteo;
CREATE VIEW v_lideres_conteo AS
 SELECT l.id AS lider_id,
    h.codigo AS codigo,
    h.nombre_completo AS lider,
    l.red,
    l.estado_celula,
    l.info_enviada,
    l.cantidad_celulas,
    l.rol_lider,
    l.tipo_12,
    l.lider_padre_id,
    hp.nombre_completo AS pastor_nombre,
    count(d.id) FILTER (WHERE (d.fecha_fin IS NULL)) AS total_asignados_vigentes,
    count(d.id) FILTER (WHERE ((d.fecha_fin IS NULL) AND d.es_discipulo_activo)) AS discipulos_activos,
    count(d.id) FILTER (WHERE ((d.fecha_fin IS NULL) AND (NOT d.es_discipulo_activo))) AS discipulos_inactivos
   FROM ((((lideres l
     JOIN hermanos h ON ((h.id = l.hermano_id)))
     LEFT JOIN lideres lp ON ((lp.id = l.lider_padre_id)))
     LEFT JOIN hermanos hp ON ((hp.id = lp.hermano_id)))
     LEFT JOIN discipulado d ON ((d.lider_id = l.id)))
  WHERE l.activo
  GROUP BY l.id, h.codigo, h.nombre_completo, l.red, l.estado_celula, l.info_enviada, l.cantidad_celulas, l.rol_lider, l.tipo_12, l.lider_padre_id, hp.nombre_completo;
