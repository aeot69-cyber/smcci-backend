-- Tablas de catálogos dinámicos para Configuración

-- 1. Redes
CREATE TABLE IF NOT EXISTS catalogo_redes (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(50) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_redes (nombre) VALUES ('HOMBRES'),('MUJERES'),('JOVENES'),('MIXTO') ON CONFLICT (nombre) DO NOTHING;

-- 2. Tipo 12
CREATE TABLE IF NOT EXISTS catalogo_tipo12 (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(50) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_tipo12 (nombre) VALUES ('12 PASTORES'),('12 APOSTOL'),('12 JOVENES'),('NINGUNO') ON CONFLICT (nombre) DO NOTHING;

-- 3. Meses
CREATE TABLE IF NOT EXISTS catalogo_meses (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(20) NOT NULL UNIQUE,
    orden INT NOT NULL,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_meses (nombre, orden) VALUES
('ENERO',1),('FEBRERO',2),('MARZO',3),('ABRIL',4),('MAYO',5),('JUNIO',6),
('JULIO',7),('AGOSTO',8),('SEPTIEMBRE',9),('OCTUBRE',10),('NOVIEMBRE',11),('DICIEMBRE',12)
ON CONFLICT (nombre) DO NOTHING;

-- 4. Etapas/Encuentros (ya existe tipos_encuentro, solo agrego activo si no tiene)
ALTER TABLE tipos_encuentro ADD COLUMN IF NOT EXISTS activo BOOLEAN DEFAULT true;

-- 5. Periodicidad (ya existe periodicidad_celula, solo agrego activo si no tiene)
ALTER TABLE periodicidad_celula ADD COLUMN IF NOT EXISTS activo BOOLEAN DEFAULT true;

-- 6. Roles de líder (catálogo) — con icono configurable
CREATE TABLE IF NOT EXISTS catalogo_roles_lider (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(50) NOT NULL UNIQUE,
    icono VARCHAR(10) DEFAULT '—',
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_roles_lider (nombre, icono) VALUES ('PASTOR','⭐'),('APOSTOL','🔻'),('LIDER','👤'),('NINGUNO','—') ON CONFLICT (nombre) DO NOTHING;
ALTER TABLE catalogo_roles_lider ADD COLUMN IF NOT EXISTS icono VARCHAR(10) DEFAULT '—';

-- 7. Estado civil
CREATE TABLE IF NOT EXISTS catalogo_estado_civil (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(30) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_estado_civil (nombre) VALUES ('SOLTERO'),('CASADO'),('VIUDO'),('DIVORCIADO'),('CONVIVIENTE') ON CONFLICT (nombre) DO NOTHING;

-- 8. Sexo
CREATE TABLE IF NOT EXISTS catalogo_sexo (
    id SERIAL PRIMARY KEY,
    codigo CHAR(1) NOT NULL UNIQUE,
    nombre VARCHAR(20) NOT NULL,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_sexo (codigo, nombre) VALUES ('M','MASCULINO'),('F','FEMENINO') ON CONFLICT (codigo) DO NOTHING;

-- 9. Estado célula
CREATE TABLE IF NOT EXISTS catalogo_estado_celula (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(30) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_estado_celula (nombre) VALUES ('CELULA ACTIVA'),('SIN CELULA'),('EN FORMACION') ON CONFLICT (nombre) DO NOTHING;

-- 10. Info enviada
CREATE TABLE IF NOT EXISTS catalogo_info_enviada (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(30) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_info_enviada (nombre) VALUES ('SI'),('EN CONSULTA'),('NO') ON CONFLICT (nombre) DO NOTHING;

-- 11. Estado visita
CREATE TABLE IF NOT EXISTS catalogo_estado_visita (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(30) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_estado_visita (nombre) VALUES ('PENDIENTE'),('EN SEGUIMIENTO'),('INTEGRADO'),('DESCARTADO') ON CONFLICT (nombre) DO NOTHING;

-- 12. Tipo seguimiento
CREATE TABLE IF NOT EXISTS catalogo_tipo_seguimiento (
    id SERIAL PRIMARY KEY,
    nombre VARCHAR(30) NOT NULL UNIQUE,
    activo BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now()
);
INSERT INTO catalogo_tipo_seguimiento (nombre) VALUES ('LLAMADA'),('VISITA'),('WHATSAPP'),('CORREO'),('REUNION') ON CONFLICT (nombre) DO NOTHING;
