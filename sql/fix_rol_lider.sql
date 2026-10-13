-- Agregar columna rol_lider
ALTER TABLE lideres ADD COLUMN IF NOT EXISTS rol_lider VARCHAR(20) DEFAULT 'NINGUNO';

-- Migrar datos de es_pastor a rol_lider
UPDATE lideres SET rol_lider='PASTOR' WHERE es_pastor=true;
UPDATE lideres SET rol_lider='NINGUNO' WHERE es_pastor=false AND rol_lider IS NULL;

-- Los que tienen tipo_12='12 APOSTOL' pero no son pastor -> APOSTOL
UPDATE lideres SET rol_lider='APOSTOL' WHERE tipo_12='12 APOSTOL' AND es_pastor=false;

-- Los que tienen tipo_12='12 PASTORES' y es_pastor -> PASTOR
UPDATE lideres SET rol_lider='PASTOR' WHERE tipo_12='12 PASTORES' AND es_pastor=true;

-- Los que tienen discipulos activos -> LIDER
UPDATE lideres SET rol_lider='LIDER' WHERE id IN (
    SELECT d.lider_id FROM discipulado d WHERE d.fecha_fin IS NULL AND d.es_discipulo_activo=true
    GROUP BY d.lider_id HAVING COUNT(*)>0
) AND rol_lider='NINGUNO';
