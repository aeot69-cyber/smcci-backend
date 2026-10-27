-- 32: Importación de Excel de visitas (consolidación de llamados)
-- Fecha: 27-oct-2026. Idempotente (segura de re-ejecutar).
-- 1) Columnas nuevas en visitas para los campos del consolidado que no
--    tenían destino (decisiones del usuario 27-oct-2026).
-- 2) Tabla temporal donde queda parsed el archivo subido hasta confirmar.

ALTER TABLE visitas ADD COLUMN IF NOT EXISTS edad integer;
ALTER TABLE visitas ADD COLUMN IF NOT EXISTS sexo varchar(10);
ALTER TABLE visitas ADD COLUMN IF NOT EXISTS culto varchar(60);
ALTER TABLE visitas ADD COLUMN IF NOT EXISTS como_llego varchar(150);
ALTER TABLE visitas ADD COLUMN IF NOT EXISTS asiste_celula varchar(30);
ALTER TABLE visitas ADD COLUMN IF NOT EXISTS encargado_contacto varchar(150);

COMMENT ON COLUMN visitas.edad IS 'Edad al momento de la visita (Excel consolidado)';
COMMENT ON COLUMN visitas.sexo IS 'M / F (crudo del archivo)';
COMMENT ON COLUMN visitas.culto IS 'Culto al que asistió (RADICAL YOUTH, DOMINGO FAMILIAR, ...)';
COMMENT ON COLUMN visitas.como_llego IS 'Cómo llegó (crudo: Invitación, Impacto radical, RRSS, ...)';
COMMENT ON COLUMN visitas.asiste_celula IS 'SI / NO / EN CONSOLIDACIÓN / DESISTE (crudo)';
COMMENT ON COLUMN visitas.encargado_contacto IS 'Encargado del contacto/seguimiento (ENCARGADO CONTACTO)';
COMMENT ON COLUMN visitas.responsable IS 'Quien ingresó el registro (INGRESADO POR) o responsable manual';

CREATE TABLE IF NOT EXISTS visitas_import_tmp (
  token uuid PRIMARY KEY,
  usuario_id integer,
  iglesia_id integer NOT NULL DEFAULT 1,
  nombre_archivo text,
  creado_en timestamptz NOT NULL DEFAULT now(),
  filas jsonb NOT NULL
);
COMMENT ON TABLE visitas_import_tmp IS
  'Parseo provisional del Excel de visitas subido: vive hasta confirmar la importación (se purga >2h)';
