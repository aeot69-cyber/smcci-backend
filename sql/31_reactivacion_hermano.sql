-- 31. Reactivación de hermanos: FECHA DE ACTIVACIÓN (vuelta a MCCI)
-- Cuando alguien que se fue vuelve y se le da de alta (desde /hermanos o
-- reactivando su asignación en /discipulado) se registra aquí la fecha.
-- NULL = nunca se dio de baja. El historial de bajas sigue en fecha_baja,
-- motivo_salida y en la tabla auditoría; el de asignaciones, en discipulado.
-- Idempotente: puede ejecutarse varias veces.
ALTER TABLE hermanos ADD COLUMN IF NOT EXISTS fecha_activacion DATE;
COMMENT ON COLUMN hermanos.fecha_activacion IS
  'Fecha en que el hermano fue reactivado tras una baja (vuelta a la iglesia). NULL si nunca se dio de baja.';
