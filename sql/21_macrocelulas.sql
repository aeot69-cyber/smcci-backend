-- 21_macrocelulas.sql
-- N° de macrocélulas por líder: campo editable + exposición en la vista.
-- Ejecutar:  psql -U mcci -d Iglesia_MCCI -f sql/21_macrocelulas.sql

ALTER TABLE lideres ADD COLUMN IF NOT EXISTS cantidad_macrocelulas INT NOT NULL DEFAULT 0;

-- La vista se re-crea completa (DROP + CREATE) porque CREATE OR REPLACE no permite
-- cambiar el orden/cantidad de columnas y la vista real tiene 14 columnas.
DROP VIEW IF EXISTS v_lideres_conteo;
CREATE VIEW v_lideres_conteo AS
SELECT l.id AS lider_id,
       h.codigo,
       h.nombre_completo AS lider,
       l.red,
       l.estado_celula,
       l.info_enviada,
       l.cantidad_celulas,
       l.rol_lider,
       l.tipo_12,
       l.lider_padre_id,
       hp.nombre_completo AS pastor_nombre,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL) AS total_asignados_vigentes,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo) AS discipulos_activos,
       count(d.id) FILTER (WHERE d.fecha_fin IS NULL AND NOT d.es_discipulo_activo) AS discipulos_inactivos,
       l.cantidad_macrocelulas
FROM lideres l
  JOIN hermanos h ON h.id = l.hermano_id
  LEFT JOIN lideres lp ON lp.id = l.lider_padre_id
  LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
  LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.codigo, h.nombre_completo, l.red, l.estado_celula, l.info_enviada,
         l.cantidad_celulas, l.rol_lider, l.tipo_12, l.lider_padre_id,
         hp.nombre_completo, l.cantidad_macrocelulas;

COMMENT ON COLUMN lideres.cantidad_macrocelulas IS 'N° de macrocélulas a cargo del líder (se carga manualmente en /lideres)';
