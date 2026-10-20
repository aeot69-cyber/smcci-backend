-- Catálogo modalidad célula: Presencial, Online, Ambas
CREATE TABLE IF NOT EXISTS catalogo_modalidad_celula (
  id SERIAL PRIMARY KEY,
  nombre VARCHAR(30) NOT NULL UNIQUE,
  activo BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_modalidad_celula (nombre) VALUES ('PRESENCIAL'),('ONLINE'),('AMBAS') ON CONFLICT (nombre) DO NOTHING;

-- Quitar CHECK hardcodeado de celulas_informe.modalidad si existe
ALTER TABLE celulas_informe DROP CONSTRAINT IF EXISTS celulas_informe_modalidad_check;
ALTER TABLE celulas_informe DROP CONSTRAINT IF EXISTS chk_modalidad;
-- Asegurar columna existe
ALTER TABLE celulas_informe ADD COLUMN IF NOT EXISTS modalidad VARCHAR(20) DEFAULT 'PRESENCIAL';
