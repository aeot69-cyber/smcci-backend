-- Catálogo dinámico para tipo de hermano: Líder / Discípulo / Ninguno
CREATE TABLE IF NOT EXISTS catalogo_tipo_hermano (
  id SERIAL PRIMARY KEY,
  nombre VARCHAR(30) NOT NULL UNIQUE,
  activo BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_tipo_hermano (nombre) VALUES ('LIDER'),('DISCIPULO'),('NINGUNO') ON CONFLICT (nombre) DO NOTHING;

-- Columna en hermanos
ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS tipo_hermano VARCHAR(30) DEFAULT 'NINGUNO';
-- Backfill: si ya es líder -> LIDER, si está en discipulado activo -> DISCIPULO, si no -> NINGUNO
UPDATE hermanos h SET tipo_hermano='LIDER' WHERE EXISTS (SELECT 1 FROM lideres l WHERE l.hermano_id=h.id AND l.activo);
UPDATE hermanos h SET tipo_hermano='DISCIPULO' WHERE tipo_hermano='NINGUNO' AND EXISTS (SELECT 1 FROM discipulado d WHERE d.hermano_id=h.id AND d.fecha_fin IS NULL AND d.es_discipulo_activo);

-- Índice para filtros
CREATE INDEX IF NOT EXISTS idx_hermanos_tipo ON hermanos(tipo_hermano);

-- Recrear vista para incluir tipo_hermano
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
    h.tipo_hermano,
    h.activo,
    h.fecha_ingreso,
    c.nombre AS comuna,
    r.nombre_corto AS region,
    r.nombre AS region_larga
   FROM ((hermanos h
     LEFT JOIN comunas c ON ((c.id = h.comuna_id)))
     LEFT JOIN regiones r ON ((r.id = c.region_id)));
