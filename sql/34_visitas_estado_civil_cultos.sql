-- 34: Visitas — estado civil + catálogo de cultos + vista actualizada
-- Fecha: 29-oct-2026. Idempotente (segura de re-ejecutar).
-- Decisiones del usuario (29-oct-2026):
--   · Visitas guarda sexo, estado civil y en qué culto fue consolidado.
--   · Cultos iniciales: RADICAL YOUTH, RADICAL TEEN, DOMINGO FAMILIAR,
--     JUEVES FAMILIAR, RED DE MUJERES, RED DE HOMBRE (editables desde
--     Configuración; se pueden agregar más).
--   · El culto de las visitas ya cargadas que venga vacío se deja vacío:
--     el administrador lo completa después con las opciones del catálogo.

ALTER TABLE visitas ADD COLUMN IF NOT EXISTS estado_civil varchar(20);
COMMENT ON COLUMN visitas.estado_civil IS 'Estado civil (catálogo catalogo_estado_civil)';

CREATE TABLE IF NOT EXISTS catalogo_cultos (
  id         serial PRIMARY KEY,
  nombre     varchar(100) NOT NULL,
  activo     boolean DEFAULT true,
  iglesia_id integer REFERENCES iglesias(id),
  created_at timestamptz DEFAULT now(),
  CONSTRAINT catalogo_cultos_nombre_iglesia_key UNIQUE NULLS NOT DISTINCT (nombre, iglesia_id)
);

INSERT INTO catalogo_cultos (nombre, iglesia_id) VALUES
  ('RADICAL YOUTH', NULL),
  ('RADICAL TEEN', NULL),
  ('DOMINGO FAMILIAR', NULL),
  ('JUEVES FAMILIAR', NULL),
  ('RED DE MUJERES', NULL),
  ('RED DE HOMBRE', NULL)
ON CONFLICT (nombre, iglesia_id) DO NOTHING;

COMMENT ON TABLE catalogo_cultos IS 'Cultos donde se consolidó la visita (Configuración → Cultos)';

-- Vista de visitas con los campos nuevos + provincia y región de la comuna.
-- IMPORTANTE: CREATE OR REPLACE exige mantener el orden de las columnas
-- existentes; las nuevas van SIEMPRE al final.
CREATE OR REPLACE VIEW v_visitas_completo AS
SELECT v.id,
       v.nombre_completo,
       v.direccion,
       v.telefono,
       v.correo,
       v.comuna_id,
       v.invitado_por_id,
       v.invitado_por_texto,
       v.fecha_visita,
       v.motivo_oracion,
       v.estado,
       v.responsable,
       v.observacion,
       v.hermano_id,
       v.iglesia_id,
       v.created_at,
       v.updated_at,
       c.nombre AS comuna,
       inv.nombre_completo AS invitado_por,
       v.rut,
       v.edad,
       v.sexo,
       v.culto,
       v.como_llego,
       v.asiste_celula,
       v.encargado_contacto,
       v.estado_civil,
       p.nombre AS provincia,
       r.nombre_corto AS region
FROM visitas v
LEFT JOIN comunas c ON c.id = v.comuna_id
LEFT JOIN provincias p ON p.id = c.provincia_id
LEFT JOIN regiones r ON r.id = c.region_id
LEFT JOIN hermanos inv ON inv.id = v.invitado_por_id;

-- Permiso del módulo nuevo "comunas" (Configuración → Regiones y Comunas).
-- SUPERADMIN ya tiene acceso total por código; ADMIN queda habilitado y el
-- resto de los roles se puede habilitar desde Configuración → Roles/Usuarios.
INSERT INTO roles_permisos (rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar, iglesia_id)
VALUES ('ADMIN', 'comunas', true, true, true, true, true, NULL)
ON CONFLICT (rol, modulo, iglesia_id) DO UPDATE
  SET puede_leer=true, puede_crear=true, puede_editar=true, puede_eliminar=true, puede_exportar=true;
