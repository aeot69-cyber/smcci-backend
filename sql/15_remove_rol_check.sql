-- Migración: Eliminar CHECK constraint de usuarios.rol
-- Los roles ahora son dinámicos desde catalogo_roles_usuario
ALTER TABLE usuarios DROP CONSTRAINT IF EXISTS usuarios_rol_check;
