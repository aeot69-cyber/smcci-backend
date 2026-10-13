-- 24) INFORME EXPRESS: código de líder para el formulario público /informe
-- Aplicar en local (iglesia_MCCI) y en Supabase (producción), en ese orden.

-- 1) Código personal por líder (6 dígitos, único). Se genera por lote después.
ALTER TABLE lideres ADD COLUMN IF NOT EXISTS codigo_informe VARCHAR(10);
CREATE UNIQUE INDEX IF NOT EXISTS ux_lideres_codigo_informe
    ON lideres (codigo_informe) WHERE codigo_informe IS NOT NULL;

-- 2) El mes del informe SIEMPRE se deriva de la fecha → debe admitir los 12
--    meses (antes solo JUNIO..DICIEMBRE; el formulario admin también lo deriva).
ALTER TABLE celulas_informe DROP CONSTRAINT IF EXISTS celulas_informe_mes_check;
ALTER TABLE celulas_informe ADD CONSTRAINT celulas_informe_mes_check
    CHECK (mes IS NULL OR upper(mes) IN
        ('ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO',
         'SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'));

-- 3) La red viene de lideres.red (catálogo completo: MIXTO, TEEN, NIÑOS...)
--    Antes solo aceptaba ADOLESCENTES/JOVENES/MUJERES/HOMBRES y los informes
--    de líderes MIXTO/TEEN fallaban.
ALTER TABLE celulas_informe DROP CONSTRAINT IF EXISTS celulas_informe_red_check;
ALTER TABLE celulas_informe ADD CONSTRAINT celulas_informe_red_check
    CHECK (red IS NULL OR red IN
        ('ADOLESCENTES','JOVENES','MUJERES','HOMBRES','MIXTO','TEEN','NIÑOS'));

-- Listo. Los códigos de líder NO se generan aquí: primero prueba local con
-- 2 líderes y después se pide autorización para el lote completo.
