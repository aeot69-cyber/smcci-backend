-- Catálogo de roles de usuario
CREATE TABLE IF NOT EXISTS catalogo_roles_usuario (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(50) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_roles_usuario (nombre) VALUES ('SUPERADMIN'),('ADMIN'),('CONSULTA') ON CONFLICT (nombre) DO NOTHING;
