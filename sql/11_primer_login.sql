-- 11_primer_login.sql: Forzar cambio de contraseña en primer ingreso
ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS primer_login BOOLEAN DEFAULT TRUE;
UPDATE usuarios SET primer_login = FALSE WHERE ultimo_login IS NOT NULL;
