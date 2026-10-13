-- Migración: liderazgo pastoral G12 MCCI
ALTER TABLE lideres ADD COLUMN IF NOT EXISTS es_pastor BOOLEAN NOT NULL DEFAULT FALSE;
ALTER TABLE lideres ADD COLUMN IF NOT EXISTS tipo_12 VARCHAR(20) NOT NULL DEFAULT 'NINGUNO';
CREATE INDEX IF NOT EXISTS idx_lideres_pastor ON lideres(es_pastor);
CREATE INDEX IF NOT EXISTS idx_lideres_tipo12 ON lideres(tipo_12);

-- Actualizar vista para incluir nuevos campos
DROP VIEW IF EXISTS v_lideres_conteo;
CREATE VIEW v_lideres_conteo AS
SELECT l.id AS lider_id, h.nombre_completo AS lider,
       l.red, l.estado_celula, l.info_enviada, l.cantidad_celulas,
       l.rol_lider, l.tipo_12,
       l.lider_padre_id,
       hp.nombre_completo AS pastor_nombre,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL) AS total_asignados_vigentes,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo) AS discipulos_activos,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL AND NOT d.es_discipulo_activo) AS discipulos_inactivos
FROM lideres l
JOIN hermanos h ON h.id = l.hermano_id
LEFT JOIN lideres lp ON lp.id = l.lider_padre_id
LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.nombre_completo, l.red, l.estado_celula, l.info_enviada, l.cantidad_celulas, l.rol_lider, l.tipo_12, l.lider_padre_id, hp.nombre_completo;
