-- ===========================================================================
-- 27_pastor_real.sql
-- Reporte "Informe Célula · Por Pastor" incompleto. Dos causas:
--
--   1) es_pastor solo se activaba con el rol "PASTOR" (masculino). Las 10
--      PASTORA quedaban con es_pastor=false y no aparecían en el reporte.
--
--   2) El reporte (y el listado) usaban SOLO el padre directo
--      (lider_padre_id). Un líder que estuviera bajo otro líder —y no bajo un
--      pastor— no aparecía bajo ningún pastor: se "perdían" 26 de los 48
--      informes. Además los apóstoles no tienen padre, por lo que su propia
--      célula no quedaba bajo ningún pastor.
--
-- Solución: el pastor de un líder es el REAL, subiendo la cadena de
-- lider_padre_id hasta encontrar la PASTORA / PASTOR / APOSTOL más cercana.
-- Si no hay nadie arriba y el líder es pastoral (los apóstoles), queda como
-- su propio pastor para que ningún informe quede fuera.
--
-- La misma lógica vive en app.py (_SQL_PASTOR_DE y pastor_de_lider) para que
-- los informes nuevos guarden ese pastor.
-- ===========================================================================

-- 1) Corrige los datos existentes: las pastoras pasan a ser pastoras de verdad
UPDATE lideres SET es_pastor = TRUE WHERE rol_lider = 'PASTORA' AND NOT es_pastor;

-- 2) Vista del listado de líderes: muestra el pastor real
DROP VIEW IF EXISTS v_lideres_conteo;
CREATE VIEW v_lideres_conteo AS
WITH RECURSIVE subo AS (
    SELECT l.id AS lider_id, l.lider_padre_id AS anc_id, 1 AS nivel, ARRAY[l.id] AS ruta
    FROM lideres l
  UNION ALL
    SELECT s.lider_id, p.lider_padre_id, s.nivel + 1, s.ruta || p.id
    FROM subo s JOIN lideres p ON p.id = s.anc_id
    WHERE s.nivel < 20 AND NOT p.id = ANY(s.ruta)
), candidatos AS (
    (SELECT DISTINCT ON (s.lider_id) s.lider_id, s.anc_id AS pastor_id, 0 AS prioridad
     FROM subo s
     JOIN lideres pp ON pp.id = s.anc_id
     WHERE s.anc_id IS NOT NULL
       AND (pp.es_pastor OR pp.rol_lider IN ('APOSTOL','PASTOR','PASTORA'))
     ORDER BY s.lider_id, s.nivel)
  UNION ALL
    -- pastor/a sin nadie arriba (los apóstoles): queda como su propio pastor
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
)
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
    l.cantidad_macrocelulas,
    l.iglesia_id
FROM lideres l
JOIN hermanos h ON h.id = l.hermano_id
LEFT JOIN pastor_de pd ON pd.lider_id = l.id
LEFT JOIN lideres lp ON lp.id = pd.pastor_id
LEFT JOIN hermanos hp ON hp.id = lp.hermano_id
LEFT JOIN discipulado d ON d.lider_id = l.id
WHERE l.activo
GROUP BY l.id, h.codigo, h.nombre_completo, l.red, l.estado_celula, l.info_enviada,
         l.cantidad_celulas, l.rol_lider, l.tipo_12, l.lider_padre_id,
         hp.nombre_completo, l.cantidad_macrocelulas, l.iglesia_id;

-- 3) Ajusta el pastor guardado en los informes existentes al pastor real
WITH RECURSIVE subo AS (
    SELECT l.id AS lider_id, l.lider_padre_id AS anc_id, 1 AS nivel, ARRAY[l.id] AS ruta
    FROM lideres l
  UNION ALL
    SELECT s.lider_id, p.lider_padre_id, s.nivel + 1, s.ruta || p.id
    FROM subo s JOIN lideres p ON p.id = s.anc_id
    WHERE s.nivel < 20 AND NOT p.id = ANY(s.ruta)
), candidatos AS (
    (SELECT DISTINCT ON (s.lider_id) s.lider_id, s.anc_id AS pastor_id, 0 AS prioridad
     FROM subo s
     JOIN lideres pp ON pp.id = s.anc_id
     WHERE s.anc_id IS NOT NULL
       AND (pp.es_pastor OR pp.rol_lider IN ('APOSTOL','PASTOR','PASTORA'))
     ORDER BY s.lider_id, s.nivel)
  UNION ALL
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
)
UPDATE celulas_informe ci
SET pastor_id = pd.pastor_id
FROM pastor_de pd
WHERE ci.lider_id = pd.lider_id
  AND ci.pastor_id IS DISTINCT FROM pd.pastor_id;

-- Control esperado: ningún informe debe quedar sin pastor
-- SELECT count(*) FROM celulas_informe WHERE pastor_id IS NULL;  --> 0
