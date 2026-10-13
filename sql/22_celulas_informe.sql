-- 22_celulas_informe.sql — Tabla de informes de célula (faltaba en el repo).
-- En producción ya existe (Supabase); este script es idempotent y solo la crea si falta.
-- NOTA: los CHECK de red/mes se eliminan en 25_multiiglesia.sql porque los
-- catálogos pasan a ser dinámicos por iglesia.

CREATE TABLE IF NOT EXISTS celulas_informe (
    id SERIAL PRIMARY KEY,
    lider_id INT REFERENCES lideres(id),
    pastor_id INT REFERENCES lideres(id),
    fecha DATE,
    mes TEXT,
    red TEXT,
    direccion TEXT,
    comuna TEXT,
    horario TEXT,
    realizo BOOLEAN DEFAULT true,
    modalidad TEXT,
    justificacion TEXT,
    discipulos_asistieron TEXT,
    discipulos_no_asistieron TEXT,
    ofrenda NUMERIC(10,2) DEFAULT 0,
    tipo_ofrenda TEXT,
    boucher_url TEXT,
    created_at TIMESTAMP NOT NULL DEFAULT NOW()
);
