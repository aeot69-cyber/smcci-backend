-- Visitas: RUT para poder integrar a hermanos (hermanos.rut es obligatorio)
ALTER TABLE visitas ADD COLUMN IF NOT EXISTS rut VARCHAR(12);
