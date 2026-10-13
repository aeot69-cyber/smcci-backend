-- 10_roles_permisos.sql: Tabla de configuración de permisos por rol y módulo
CREATE TABLE IF NOT EXISTS roles_permisos (
    id SERIAL PRIMARY KEY,
    rol VARCHAR(20) NOT NULL,
    modulo VARCHAR(30) NOT NULL,
    puede_leer BOOLEAN DEFAULT TRUE,
    puede_crear BOOLEAN DEFAULT FALSE,
    puede_editar BOOLEAN DEFAULT FALSE,
    puede_eliminar BOOLEAN DEFAULT FALSE,
    puede_exportar BOOLEAN DEFAULT TRUE,
    UNIQUE(rol, modulo)
);

-- SUPERADMIN: todo permitido
INSERT INTO roles_permisos (rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar) VALUES
('SUPERADMIN', 'hermanos',     TRUE, TRUE, TRUE, TRUE, TRUE),
('SUPERADMIN', 'lideres',      TRUE, TRUE, TRUE, TRUE, TRUE),
('SUPERADMIN', 'discipulado',  TRUE, TRUE, TRUE, TRUE, TRUE),
('SUPERADMIN', 'encuentros',   TRUE, TRUE, TRUE, TRUE, TRUE),
('SUPERADMIN', 'seguimiento',  TRUE, TRUE, TRUE, TRUE, TRUE),
('SUPERADMIN', 'visitas',      TRUE, TRUE, TRUE, TRUE, TRUE),
('SUPERADMIN', 'reportes',     TRUE, FALSE, FALSE, FALSE, TRUE),
('SUPERADMIN', 'usuarios',     TRUE, TRUE, TRUE, TRUE, FALSE)
ON CONFLICT (rol, modulo) DO UPDATE SET
    puede_leer=EXCLUDED.puede_leer, puede_crear=EXCLUDED.puede_crear,
    puede_editar=EXCLUDED.puede_editar, puede_eliminar=EXCLUDED.puede_eliminar,
    puede_exportar=EXCLUDED.puede_exportar;

-- ADMIN: operaciones completas excepto usuarios
INSERT INTO roles_permisos (rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar) VALUES
('ADMIN', 'hermanos',     TRUE, TRUE, TRUE, FALSE, TRUE),
('ADMIN', 'lideres',      TRUE, TRUE, TRUE, FALSE, TRUE),
('ADMIN', 'discipulado',  TRUE, TRUE, TRUE, FALSE, TRUE),
('ADMIN', 'encuentros',   TRUE, TRUE, TRUE, FALSE, TRUE),
('ADMIN', 'seguimiento',  TRUE, TRUE, TRUE, FALSE, TRUE),
('ADMIN', 'visitas',      TRUE, TRUE, TRUE, FALSE, TRUE),
('ADMIN', 'reportes',     TRUE, FALSE, FALSE, FALSE, TRUE),
('ADMIN', 'usuarios',     FALSE, FALSE, FALSE, FALSE, FALSE)
ON CONFLICT (rol, modulo) DO UPDATE SET
    puede_leer=EXCLUDED.puede_leer, puede_crear=EXCLUDED.puede_crear,
    puede_editar=EXCLUDED.puede_editar, puede_eliminar=EXCLUDED.puede_eliminar,
    puede_exportar=EXCLUDED.puede_exportar;

-- CONSULTA: solo lectura y exportación
INSERT INTO roles_permisos (rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar) VALUES
('CONSULTA', 'hermanos',     TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'lideres',      TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'discipulado',  TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'encuentros',   TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'seguimiento',  TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'visitas',      TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'reportes',     TRUE, FALSE, FALSE, FALSE, TRUE),
('CONSULTA', 'usuarios',     FALSE, FALSE, FALSE, FALSE, FALSE)
ON CONFLICT (rol, modulo) DO UPDATE SET
    puede_leer=EXCLUDED.puede_leer, puede_crear=EXCLUDED.puede_crear,
    puede_editar=EXCLUDED.puede_editar, puede_eliminar=EXCLUDED.puede_eliminar,
    puede_exportar=EXCLUDED.puede_exportar;

-- Vista útil para consultar permisos
CREATE OR REPLACE VIEW v_roles_permisos AS
SELECT rol, modulo, puede_leer, puede_crear, puede_editar, puede_eliminar, puede_exportar
FROM roles_permisos ORDER BY rol, modulo;
