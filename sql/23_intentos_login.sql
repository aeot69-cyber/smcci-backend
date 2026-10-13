-- 23_intentos_login.sql — Registro de intentos de login (anti-bots)
-- Bloqueo tras intentos fallidos + límite de recuperaciones de contraseña.
-- Ejecutar en orden después de 22_lideres_padre_discipulos.sql
-- Local : psql -U mcci -d iglesia_MCCI -f 23_intentos_login.sql
-- Supabase: psql -h aws-0-us-west-2.pooler.supabase.com -U postgres.dcmwyuezcmukiwcxbpxp -d postgres -f 23_intentos_login.sql

CREATE TABLE IF NOT EXISTS intentos_login (
    id       BIGSERIAL PRIMARY KEY,
    ip       TEXT NOT NULL DEFAULT '',
    username TEXT NOT NULL DEFAULT '',
    exito    BOOLEAN NOT NULL DEFAULT FALSE,
    -- 'credenciales' = clave incorrecta (cuenta para el bloqueo por usuario)
    -- 'captcha' | 'honeypot' | 'rapido' = bot detectado (cuenta solo por IP)
    -- 'bloqueado' = intento mientras estaba bloqueado
    -- 'recuperar' = solicitud de recuperación de contraseña
    -- 'api' | ''   = login correcto (resetea contadores)
    detalle  TEXT NOT NULL DEFAULT '',
    creado   TIMESTAMP NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_intentos_login_creado ON intentos_login (creado);
CREATE INDEX IF NOT EXISTS ix_intentos_login_user    ON intentos_login (username, creado);
CREATE INDEX IF NOT EXISTS ix_intentos_login_ip      ON intentos_login (ip, creado);

-- Vista para revisar intentos sospechosos (menú Auditoría / SQL directo):
--   SELECT * FROM v_intentos_login_sospechosos;
CREATE OR REPLACE VIEW v_intentos_login_sospechosos AS
SELECT username, ip,
       count(*)          AS fallos,
       count(DISTINCT detalle) AS tipos,
       max(creado)       AS ultimo
FROM intentos_login
WHERE NOT exito
  AND creado > now() - interval '24 hours'
  AND detalle <> 'recuperar'
GROUP BY username, ip
HAVING count(*) >= 3
ORDER BY fallos DESC;
