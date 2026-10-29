-- 35_informe_celula_ubicacion.sql
-- Ubicación estructurada de la célula en el Informe de Célula:
-- Región → Provincia → Comuna, con el mismo catálogo oficial (regiones /
-- provincias / comunas) que usan visitas y hermanos.
--
-- Se conserva la columna de texto `comuna` (histórico y exportaciones); la
-- aplicación la mantiene sincronizada con el nombre de la comuna elegida.
-- Los informes antiguos con comuna escrita a mano se migran cuando el nombre
-- coincide con el catálogo (sin distinguir mayúsculas ni tildes); el resto
-- queda sin ubicación para que el administrador la complete al editarlos.

BEGIN;

ALTER TABLE celulas_informe
    ADD COLUMN IF NOT EXISTS region_id    integer REFERENCES regiones(id),
    ADD COLUMN IF NOT EXISTS provincia_id integer REFERENCES provincias(id),
    ADD COLUMN IF NOT EXISTS comuna_id    integer REFERENCES comunas(id);

CREATE INDEX IF NOT EXISTS idx_celulas_informe_comuna    ON celulas_informe(comuna_id);
CREATE INDEX IF NOT EXISTS idx_celulas_informe_provincia ON celulas_informe(provincia_id);
CREATE INDEX IF NOT EXISTS idx_celulas_informe_region    ON celulas_informe(region_id);

-- Migración de los informes que ya traen el nombre de la comuna en texto.
UPDATE celulas_informe ci
   SET comuna_id    = c.id,
       provincia_id = c.provincia_id,
       region_id    = c.region_id,
       comuna       = c.nombre
  FROM comunas c
 WHERE ci.comuna_id IS NULL
   AND btrim(coalesce(ci.comuna, '')) <> ''
   AND lower(translate(btrim(ci.comuna),
            'ÁÉÍÓÚÜÑáéíóúüñÁÉÍÓÚÜÑ', 'AEIOUUNaeiouunAEIOUUN'))
     = lower(translate(btrim(c.nombre),
            'ÁÉÍÓÚÜÑáéíóúüñÁÉÍÓÚÜÑ', 'AEIOUUNaeiouunAEIOUUN'));

COMMIT;
