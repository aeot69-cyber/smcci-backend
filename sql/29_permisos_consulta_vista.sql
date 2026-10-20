-- =====================================================================
-- 29_permisos_consulta_vista.sql   (idempotente)
-- ---------------------------------------------------------------------
-- Motivo: los roles CONSULTA y VISTA solo tenían la fila «celulas» en la
-- matriz de permisos, así que solo podían usar /celulas-informe. Ahora que
-- las rutas de listado llevan @permiso_req, se les da LECTURA de los
-- módulos de consulta para que no pierdan nada de lo que veían (nada de
-- crear, editar ni eliminar).
-- Si en el futuro se quiere que también consulten usuarios o
-- configuración, basta con añadir la fila correspondiente aquí.
-- =====================================================================

INSERT INTO roles_permisos
    (rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, iglesia_id)
SELECT r.rol, m.modulo, true, false, false, false, NULL
FROM (VALUES ('CONSULTA'), ('VISTA')) AS r(rol)
CROSS JOIN (VALUES
       ('hermanos'), ('lideres'), ('discipulado'), ('encuentros'),
       ('seguimiento'), ('visitas'), ('reportes')
) AS m(modulo)
ON CONFLICT (rol, modulo, iglesia_id) DO NOTHING;
