-- Relacionar Seguimiento <-> Visitas
ALTER TABLE seguimiento ADD COLUMN IF NOT EXISTS visita_id INT REFERENCES visitas(id) ON UPDATE CASCADE ON DELETE CASCADE;
ALTER TABLE seguimiento ALTER COLUMN hermano_id DROP NOT NULL;
-- Al menos uno debe estar presente
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname='chk_seg_uno') THEN
    ALTER TABLE seguimiento ADD CONSTRAINT chk_seg_uno CHECK (hermano_id IS NOT NULL OR visita_id IS NOT NULL);
  END IF;
END $$;
CREATE INDEX IF NOT EXISTS idx_seg_visita ON seguimiento(visita_id);

-- Bandeja unificada: hermanos sin célula + visitas pendientes/en seguimiento
CREATE OR REPLACE VIEW v_seguimiento_total AS
SELECT 'HERMANO' AS tipo, h.id AS ref_id, h.nombre_completo, h.telefono,
       COALESCE(h.invitado_por, h.invitado_por_texto) AS invitado_por,
       (SELECT MAX(s.fecha) FROM seguimiento s WHERE s.hermano_id=h.id) AS ultimo_seg
FROM v_sin_celula h
UNION ALL
SELECT 'VISITA', v.id, v.nombre_completo, v.telefono,
       COALESCE(inv.nombre_completo, v.invitado_por_texto),
       (SELECT MAX(s.fecha) FROM seguimiento s WHERE s.visita_id=v.id)
FROM visitas v LEFT JOIN hermanos inv ON inv.id=v.invitado_por_id
WHERE v.estado IN ('PENDIENTE','EN SEGUIMIENTO')
ORDER BY ultimo_seg NULLS FIRST, nombre_completo;
