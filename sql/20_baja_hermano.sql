-- Control de hermanos que se van: quedan inactivos con motivo y fecha, código se conserva
ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS motivo_salida TEXT;
ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS fecha_baja DATE;

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
    h.motivo_salida,
    h.fecha_baja,
    c.nombre AS comuna,
    r.nombre_corto AS region,
    r.nombre AS region_larga
   FROM ((hermanos h
     LEFT JOIN comunas c ON ((c.id = h.comuna_id)))
     LEFT JOIN regiones r ON ((r.id = c.region_id)));

-- Vista solo inactivos para control histórico (alguna vez activos)
CREATE OR REPLACE VIEW v_hermanos_inactivos AS
 SELECT * FROM v_hermanos_completo WHERE activo=false ORDER BY fecha_baja DESC NULLS LAST, nombre_completo;
