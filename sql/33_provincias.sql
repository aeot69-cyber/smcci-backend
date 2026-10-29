-- 33: Provincias de Chile + enlace comuna → provincia → región
-- Fecha: 29-oct-2026. Idempotente (segura de re-ejecutar).
-- El archivo regiones.xlsx trae 16 regiones, 56 provincias y 346 comunas.
-- Las regiones ya existen (tabla regiones, 16 filas, codigo = XV, I, ... RM).
-- Aquí se crea la tabla de provincias y se enlaza cada comuna a su provincia.

CREATE TABLE IF NOT EXISTS provincias (
  id         serial PRIMARY KEY,
  nombre     varchar(100) NOT NULL,
  region_id  integer NOT NULL REFERENCES regiones(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  created_at timestamptz DEFAULT now(),
  CONSTRAINT provincias_nombre_region_key UNIQUE (nombre, region_id)
);

CREATE INDEX IF NOT EXISTS idx_provincias_region ON provincias(region_id);

ALTER TABLE comunas
  ADD COLUMN IF NOT EXISTS provincia_id integer
  REFERENCES provincias(id) ON UPDATE CASCADE ON DELETE SET NULL;

CREATE INDEX IF NOT EXISTS idx_comunas_provincia ON comunas(provincia_id);

COMMENT ON TABLE provincias IS 'Provincias de Chile (56), dependen de regiones';
COMMENT ON COLUMN comunas.provincia_id IS 'Provincia de la comuna (catálogo regiones.xlsx)';
