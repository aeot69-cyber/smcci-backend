-- =====================================================================
-- 28_esquema_faltante.sql   (idempotente: admite ejecutarse más de una vez)
-- ---------------------------------------------------------------------
-- Motivo: en producción existen la tabla «ayudas» y las columnas
--   lideres.lider_padre_id / lideres.timoteo_id / celulas_informe.timoteo_id
-- que no estaban en el esquema base. Las columnas de «lideres» ya se crean
-- en 01_schema.sql (hacían falta pronto: las vistas de 05, 17, 21, 22 y 25
-- las usan); siguen aquí a modo de parche idempotente para las bases ya
-- instaladas, junto con la tabla «ayudas» y el índice de la cadena de líder.
-- Sin esto, /lideres, /lideres/<id>, /informe, /celulas-informe y los
-- reportes devolvían error 500 en una instalación nueva.
-- Las definiciones están copiadas tal cual del esquema real de producción.
-- =====================================================================

ALTER TABLE lideres ADD COLUMN IF NOT EXISTS lider_padre_id integer;
ALTER TABLE lideres ADD COLUMN IF NOT EXISTS timoteo_id integer;
ALTER TABLE celulas_informe ADD COLUMN IF NOT EXISTS timoteo_id integer;

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint
        WHERE conrelid = 'lideres'::regclass
          AND conname  = 'lideres_lider_padre_id_fkey'
    ) THEN
        ALTER TABLE lideres
            ADD CONSTRAINT lideres_lider_padre_id_fkey
            FOREIGN KEY (lider_padre_id) REFERENCES lideres(id);
    END IF;
END $$;

-- Ayudas editables por SUPERADMIN: guarda SOLO las diferencias respecto al
-- snapshot por defecto (ayudas_defecto.py); clave p. ej. "informe_express.ft.0".
CREATE TABLE IF NOT EXISTS ayudas (
    clave           text PRIMARY KEY,
    texto           text NOT NULL,
    actualizado_en  timestamptz DEFAULT now(),
    actualizado_por text
);

-- Subir la cadena de líder hasta el pastor recorre lider_padre_id (CTE
-- recursiva de v_lideres_conteo y de pastor_de_lider).
CREATE INDEX IF NOT EXISTS idx_lideres_padre ON lideres(lider_padre_id);
