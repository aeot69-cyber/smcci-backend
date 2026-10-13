-- =============================================
-- SMCCI - Reportes y consultas pedidas
-- =============================================

-- R1. Discípulos por líder (usa vista, siempre consistente)
SELECT * FROM v_lideres_conteo ORDER BY discipulos_activos DESC;

-- R2. Detalle discípulos de un líder (ej: id=1)
-- SELECT h.nombre_completo, h.telefono, p.nombre AS periodicidad, d.es_discipulo_activo
-- FROM discipulado d
-- JOIN hermanos h ON h.id = d.hermano_id
-- LEFT JOIN periodicidad_celula p ON p.id = d.periodicidad_id
-- WHERE d.lider_id = 1 AND d.fecha_fin IS NULL;

-- R3. Células activas vs sin célula por red
SELECT red, estado_celula, COUNT(*) FROM lideres WHERE activo GROUP BY 1,2 ORDER BY 1,2;

-- R4. Hermanos por comuna / región
SELECT r.nombre_corto AS region, c.nombre AS comuna, COUNT(*) AS total
FROM hermanos h
LEFT JOIN comunas c ON c.id = h.comuna_id
LEFT JOIN regiones r ON r.id = c.region_id
WHERE h.activo GROUP BY 1,2 ORDER BY 1,2;

-- R5. Cumpleaños del mes (ej: mes actual)
SELECT nombre_completo, fecha_nacimiento,
       EXTRACT(DAY FROM fecha_nacimiento) AS dia
FROM hermanos WHERE activo
  AND EXTRACT(MONTH FROM fecha_nacimiento)=EXTRACT(MONTH FROM CURRENT_DATE)
ORDER BY dia;

-- R6. Aniversarios del mes
SELECT nombre_completo, fecha_aniversario
FROM hermanos WHERE activo AND fecha_aniversario IS NOT NULL
  AND EXTRACT(MONTH FROM fecha_aniversario)=EXTRACT(MONTH FROM CURRENT_DATE);

-- R7. Ruta espiritual incompleta (les falta Escuela Líderes)
SELECT * FROM v_ruta_espiritual WHERE escuela_lideres=0;

-- R8. Encuentros pendientes de validar
SELECT h.nombre_completo, t.nombre AS etapa, e.fecha_evento, e.estado
FROM encuentro_participacion e
JOIN hermanos h ON h.id=e.hermano_id
JOIN tipos_encuentro t ON t.id=e.tipo_encuentro_id
WHERE e.estado='PENDIENTE' ORDER BY e.fecha_evento;

-- R9. Líderes SIN CÉLULA (de tu planilla: Cristian Barriga, Daniel Moraga)
SELECT l.id, h.nombre_completo, l.red
FROM lideres l JOIN hermanos h ON h.id=l.hermano_id
WHERE l.estado_celula='SIN CELULA' AND l.activo;

-- R10. Periodicidad más usada
SELECT p.nombre, COUNT(d.id) AS asignaciones
FROM discipulado d JOIN periodicidad_celula p ON p.id=d.periodicidad_id
WHERE d.fecha_fin IS NULL GROUP BY 1 ORDER BY 2 DESC;
