-- 22_lideres_padre_discipulos.sql
-- Vista que une: líderes con sub-lideres (hijos) + discípulos activos a cargo.
-- Ejecutar: psql -U mcci -d Iglesia_MCCI -f sql/22_lideres_padre_discipulos.sql

CREATE OR REPLACE VIEW v_lideres_padre_con_discipulos AS
SELECT l.id AS lider_id,
       h.codigo,
       h.nombre_completo AS lider,
       l.tipo_12,
       l.cantidad_celulas,
       (SELECT COUNT(*) FROM lideres c WHERE c.lider_padre_id=l.id AND c.activo) AS n_hijos,
       COUNT(d.id) FILTER (WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo) AS discipulos_activos
FROM lideres l
  JOIN hermanos h ON h.id = l.hermano_id
  LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.codigo, h.nombre_completo, l.tipo_12, l.cantidad_celulas;

COMMENT ON VIEW v_lideres_padre_con_discipulos IS 'Líderes con n_hijos y discipulos_activos; usar en /lideres/padre-discipulos';