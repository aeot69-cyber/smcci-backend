-- Módulo Visitas: personas que solo van al culto, invitadas, con motivo de oración
CREATE TABLE IF NOT EXISTS visitas (
  id SERIAL PRIMARY KEY,
  nombre_completo VARCHAR(200) NOT NULL,
  direccion VARCHAR(200),
  telefono VARCHAR(20),
  correo VARCHAR(120),
  comuna_id INT REFERENCES comunas(id) ON UPDATE CASCADE ON DELETE SET NULL,
  invitado_por_id INT REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE SET NULL,
  invitado_por_texto VARCHAR(150),
  fecha_visita DATE NOT NULL DEFAULT CURRENT_DATE,
  motivo_oracion TEXT NOT NULL,
  estado VARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
  responsable VARCHAR(150),
  observacion TEXT,
  hermano_id INT REFERENCES hermanos(id) ON UPDATE CASCADE ON DELETE SET NULL, -- si se integra, link al hermano creado
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_visitas_nombre ON visitas(nombre_completo);
CREATE INDEX IF NOT EXISTS idx_visitas_estado ON visitas(estado);
CREATE INDEX IF NOT EXISTS idx_visitas_fecha ON visitas(fecha_visita);

CREATE OR REPLACE FUNCTION fn_touch_updated_at() RETURNS TRIGGER AS $$
BEGIN NEW.updated_at = now(); RETURN NEW; END; $$ LANGUAGE plpgsql;
DROP TRIGGER IF EXISTS trg_visitas_touch ON visitas;
CREATE TRIGGER trg_visitas_touch BEFORE UPDATE ON visitas FOR EACH ROW EXECUTE FUNCTION fn_touch_updated_at();

CREATE OR REPLACE VIEW v_visitas_completo AS
SELECT v.*, c.nombre AS comuna, inv.nombre_completo AS invitado_por
FROM visitas v LEFT JOIN comunas c ON c.id=v.comuna_id LEFT JOIN hermanos inv ON inv.id=v.invitado_por_id
ORDER BY v.fecha_visita DESC;
