-- Permisos individuales por usuario (sobreescriben los del rol)
CREATE TABLE IF NOT EXISTS usuario_permisos (
    usuario_id INT NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
    modulo VARCHAR(50) NOT NULL,
    puede_leer BOOLEAN DEFAULT false,
    puede_crear BOOLEAN DEFAULT false,
    puede_editar BOOLEAN DEFAULT false,
    puede_eliminar BOOLEAN DEFAULT false,
    puede_exportar BOOLEAN DEFAULT false,
    PRIMARY KEY (usuario_id, modulo)
);
