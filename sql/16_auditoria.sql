-- Auditoría: quién se conectó y qué hizo
CREATE TABLE IF NOT EXISTS auditoria (
  id SERIAL PRIMARY KEY,
  usuario_id INT REFERENCES usuarios(id) ON DELETE SET NULL,
  username VARCHAR(100),
  rol VARCHAR(20),
  accion VARCHAR(30) NOT NULL,  -- LOGIN, LOGOUT, CREAR, EDITAR, ELIMINAR, EXPORTAR, RESET_CLAVE
  modulo VARCHAR(30) NOT NULL,  -- hermanos, lideres, visitas, etc.
  detalle TEXT,
  ip VARCHAR(45),
  created_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_auditoria_usuario ON auditoria(usuario_id);
CREATE INDEX IF NOT EXISTS idx_auditoria_fecha ON auditoria(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_auditoria_modulo ON auditoria(modulo);
