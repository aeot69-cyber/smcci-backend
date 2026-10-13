-- =============================================
-- SMCCI - Seed: 16 Regiones Chile + Comunas RM + catálogos
-- Ejecutar DESPUÉS de 01_schema.sql
-- =============================================

INSERT INTO regiones(codigo,nombre,nombre_corto,numeral) VALUES
('XV','Región de Arica y Parinacota','Arica','XV'),
('I','Región de Tarapacá','Tarapacá','I'),
('II','Región de Antofagasta','Antofagasta','II'),
('III','Región de Atacama','Atacama','III'),
('IV','Región de Coquimbo','Coquimbo','IV'),
('V','Región de Valparaíso','Valparaíso','V'),
('RM','Región Metropolitana de Santiago','Metropolitana','XIII'),
('VI','Región del Libertador Gral. Bernardo O''Higgins','O''Higgins','VI'),
('VII','Región del Maule','Maule','VII'),
('XVI','Región de Ñuble','Ñuble','XVI'),
('VIII','Región del Biobío','Biobío','VIII'),
('IX','Región de La Araucanía','Araucanía','IX'),
('XIV','Región de Los Ríos','Los Ríos','XIV'),
('X','Región de Los Lagos','Los Lagos','X'),
('XI','Región de Aysén del Gral. Carlos Ibáñez del Campo','Aysén','XI'),
('XII','Región de Magallanes y de la Antártica Chilena','Magallanes','XII')
ON CONFLICT (codigo) DO NOTHING;

-- Comunas RM principales (extensible a las 346 del país)
INSERT INTO comunas(nombre, region_id) VALUES
('Puente Alto',(SELECT id FROM regiones WHERE codigo='RM')),
('Maipú',(SELECT id FROM regiones WHERE codigo='RM')),
('Pirque',(SELECT id FROM regiones WHERE codigo='RM')),
('La Florida',(SELECT id FROM regiones WHERE codigo='RM')),
('La Pintana',(SELECT id FROM regiones WHERE codigo='RM')),
('San Bernardo',(SELECT id FROM regiones WHERE codigo='RM')),
('Las Condes',(SELECT id FROM regiones WHERE codigo='RM')),
('Providencia',(SELECT id FROM regiones WHERE codigo='RM')),
('Ñuñoa',(SELECT id FROM regiones WHERE codigo='RM')),
('Santiago',(SELECT id FROM regiones WHERE codigo='RM')),
('Peñalolén',(SELECT id FROM regiones WHERE codigo='RM')),
('El Bosque',(SELECT id FROM regiones WHERE codigo='RM')),
('La Granja',(SELECT id FROM regiones WHERE codigo='RM')),
('San José de Maipo',(SELECT id FROM regiones WHERE codigo='RM'))
ON CONFLICT DO NOTHING;

INSERT INTO periodicidad_celula(nombre, intervalo_dias, descripcion) VALUES
('SEMANAL',7,'Célula una vez por semana'),
('QUINCENAL',15,'Célula cada 15 días'),
('MENSUAL',30,'Célula una vez al mes')
ON CONFLICT (nombre) DO NOTHING;

INSERT INTO tipos_encuentro(nombre, descripcion, orden) VALUES
('ENCUENTRO TREMENDO','Primer encuentro espiritual',1),
('FRUTO FIEL','Segundo nivel / afirmación',2),
('REENCUENTRO','Restauración',3),
('ESCUELA CRECIMIENTO','Formación básica',4),
('ESCUELA LIDERES','Formación de liderazgo',5)
ON CONFLICT (nombre) DO NOTHING;
