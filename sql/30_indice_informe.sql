-- 30_indice_informe.sql — Índice faltante en celulas_informe (auditoría, menores).
-- Lo usan: la comprobación "¿ya existe informe del líder en esta fecha?" y los
-- contadores de informes por semana en el informe exprés.
-- Idempotente: se puede ejecutar en producción dos veces sin efecto.

CREATE INDEX IF NOT EXISTS idx_celulas_informe_lider_fecha ON celulas_informe(lider_id, fecha);
