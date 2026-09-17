-- Migración seguimiento diario: quién invitó + seguimientos
ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS invitado_por_id INT REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE SET NULL;
ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS invitado_por_texto VARCHAR(150);
CREATE INDEX IF NOT EXISTS idx_hermanos_invitado ON hermanos(invitado_por_id);

CREATE TABLE IF NOT EXISTS seguimiento (
  id SERIAL PRIMARY KEY,
  hermano_id INT NOT NULL REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE CASCADE,
  fecha DATE NOT NULL DEFAULT CURRENT_DATE,
  tipo VARCHAR(30) NOT NULL DEFAULT 'LLAMADA',
  comentario TEXT,
  responsable VARCHAR(100),
  contactado BOOLEAN DEFAULT FALSE,
  proximo_contacto DATE,
  created_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_seg_hermano ON seguimiento(hermano_id);
CREATE INDEX IF NOT EXISTS idx_seg_fecha ON seguimiento(fecha);

-- Vista: hermanos activos SIN asignación vigente (sin célula/líder) + quién invitó + último seguimiento
CREATE OR REPLACE VIEW v_sin_celula AS
SELECT h.id, h.rut, h.nombre_completo, h.telefono, h.comuna_nombre, h.region,
       inv.nombre_completo AS invitado_por, h.invitado_por_texto,
       (SELECT MAX(s.fecha) FROM seguimiento s WHERE s.hermano_id=h.id) AS ultimo_seg,
       (SELECT COUNT(*) FROM seguimiento s WHERE s.hermano_id=h.id AND s.contactado) AS veces_contactado
FROM (SELECT h.*, c.nombre AS comuna_nombre, r.nombre_corto AS region FROM hermanos h
      LEFT JOIN comunas c ON c.id=h.comuna_id LEFT JOIN regiones r ON r.id=c.region_id) h
LEFT JOIN hermanos inv ON inv.id=h.invitado_por_id
LEFT JOIN discipulado d ON d.hermano_id=h.id AND d.fecha_fin IS NULL
LEFT JOIN lideres l ON l.hermano_id=h.id AND l.activo
WHERE h.activo AND d.id IS NULL AND l.id IS NULL
ORDER BY ultimo_seg NULLS FIRST, h.nombre_completo;
