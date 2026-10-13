-- Datos de prueba SMCCI (re-ejecutable)
INSERT INTO hermanos(rut,nombres,apellido_paterno,apellido_materno,fecha_nacimiento,correo,telefono,direccion,comuna_id,sexo,estado_civil)
VALUES
('11111111-1','Abraham','Garate','Perez','1980-05-10','abraham.garate@mcci.cl','+56911111111','Av. Test 123',(SELECT id FROM comunas WHERE nombre='Puente Alto'),'M','CASADO'),
('22222222-2','Jaime','Duarte','Soto','1995-08-20','jaime.duarte@mcci.cl','+56922222222','Calle Test 456',(SELECT id FROM comunas WHERE nombre='Maipú'),'M','SOLTERO'),
('12345678-5','Maria','Perez','Gonzalez','1990-03-15','maria.perez@mcci.cl','+56933333333','Pasaje Test 789',(SELECT id FROM comunas WHERE nombre='Pirque'),'F','CASADO')
ON CONFLICT (rut) DO NOTHING;

-- Líder: Abraham Garate (red HOMBRES, célula activa)
INSERT INTO lideres(hermano_id, red, estado_celula, info_enviada, cantidad_celulas, observacion)
SELECT id,'HOMBRES','CELULA ACTIVA','SI',1,'Líder prueba'
FROM hermanos WHERE rut='11111111-1'
ON CONFLICT DO NOTHING;

-- Discipulado: Jaime + Maria asignados a Abraham, periodicidad SEMANAL
INSERT INTO discipulado(hermano_id, lider_id, es_discipulo_activo, periodicidad_id, observacion)
SELECT h.id, l.id, TRUE, (SELECT id FROM periodicidad_celula WHERE nombre='SEMANAL'), 'Asignación prueba'
FROM hermanos h, lideres l
JOIN hermanos hl ON hl.id = l.hermano_id
WHERE h.rut IN ('22222222-2','12345678-5') AND hl.rut='11111111-1'
ON CONFLICT DO NOTHING;

-- Encuentros: Jaime aprobó Tremendo, Maria pendiente
INSERT INTO encuentro_participacion(hermano_id, tipo_encuentro_id, estado)
SELECT h.id, t.id, CASE WHEN h.rut='22222222-2' THEN 'APROBADO' ELSE 'PENDIENTE' END
FROM hermanos h CROSS JOIN tipos_encuentro t
WHERE h.rut IN ('22222222-2','12345678-5') AND t.nombre IN ('ENCUENTRO TREMENDO','FRUTO FIEL')
ON CONFLICT (hermano_id, tipo_encuentro_id) DO NOTHING;
