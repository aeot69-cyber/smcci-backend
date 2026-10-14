-- =====================================================================
-- 26_prefijo_codigo.sql
-- Código de hermanos con PREFIJO POR IGLESIA (opción elegida)
-- Cada sede numera desde 0001 con su prefijo: MCCI0001 (matriz),
-- NOR0001 (hija), GRA0001 (grupo). La numeración sigue siendo única
-- en todo el sistema (prefijo único + secuencia propia por iglesia).
-- Idempotente (verificado con 3 corridas consecutivas).
-- =====================================================================

-- 1) Columna de prefijo en iglesias
ALTER TABLE iglesias ADD COLUMN IF NOT EXISTS prefijo_codigo VARCHAR(6);

-- 2) CASA MATRIZ conserva MCCI
UPDATE iglesias SET prefijo_codigo='MCCI' WHERE id=1 AND prefijo_codigo IS NULL;

-- 3) Derivar prefijo para iglesias existentes sin prefijo
--    'IGLESIA NORTE' -> 'NOR' (se quitan palabras genéricas); si colisiona
--    se alarga hasta 6 letras y, en último caso, se usa el id (único).
DO $$
DECLARE r RECORD; base TEXT; cand TEXT; n INT;
BEGIN
  FOR r IN SELECT id, nombre FROM iglesias
             WHERE prefijo_codigo IS NULL OR prefijo_codigo='' ORDER BY id LOOP
    base := upper(r.nombre);
    IF base ~ '^(IGLESIA|GRUPO|CASA|COMUNIDAD|RED) ' THEN
      base := substring(base FROM '^[A-Z]+ (.*)$');
    END IF;
    base := regexp_replace(base, '[^A-Z]', '', 'g');
    IF base='' THEN base := 'IGL'; END IF;
    cand := left(base,3); n := 3;
    WHILE n < 6 AND EXISTS (SELECT 1 FROM iglesias i WHERE upper(i.prefijo_codigo)=cand) LOOP
      n := n + 1; cand := left(base,n);
    END LOOP;
    IF EXISTS (SELECT 1 FROM iglesias i WHERE upper(i.prefijo_codigo)=cand) THEN
      cand := lpad(r.id::text, 6, '0');
    END IF;
    UPDATE iglesias SET prefijo_codigo=cand WHERE id=r.id;
  END LOOP;
END $$;

-- 4) Prefijos únicos en el sistema (impide que dos iglesias generen
--    códigos idénticos). NULL repetido está permitido.
CREATE UNIQUE INDEX IF NOT EXISTS ux_iglesias_prefijo
  ON iglesias (upper(prefijo_codigo));

-- 5) Secuencia propia por iglesia (numeración atómica, sin repeticiones)
DO $$
DECLARE r RECORD; seq TEXT; mx INT; lv BIGINT; es_nueva BOOLEAN;
BEGIN
  FOR r IN SELECT id, prefijo_codigo FROM iglesias ORDER BY id LOOP
    seq := 'seq_codigo_igl_' || r.id;
    EXECUTE format('CREATE SEQUENCE IF NOT EXISTS %I', seq);
    -- Sembrar la secuencia solo si nunca se usó (idempotente):
    SELECT (s.last_value IS NULL) INTO es_nueva
      FROM pg_sequences s WHERE s.schemaname='public' AND s.sequencename=seq;
    IF coalesce(es_nueva, false) THEN
      IF r.id = 1 THEN
        -- MATRIZ: continuar desde el máximo entre la secuencia global
        -- (MCCI de siempre) y los códigos MCCI existentes.
        SELECT COALESCE(MAX((substring(h.codigo FROM '([0-9]+)$'))::INT),0) INTO mx
          FROM hermanos h WHERE h.codigo ~ '^MCCI[0-9]+$';
        lv := 0;
        BEGIN
          EXECUTE 'SELECT last_value FROM seq_codigo' INTO lv;
        EXCEPTION WHEN undefined_table THEN lv := 0;
        END;
        EXECUTE format('SELECT setval(%L, GREATEST(%s,%s)+1, false)', seq, mx, lv);
      ELSE
        -- Hijas/grupos: partir desde 0001 (o desde el máximo si ya hay
        -- códigos con ese prefijo, p. ej. creados antes de esta versión).
        SELECT COALESCE(MAX((substring(h.codigo FROM '([0-9]+)$'))::INT),0) INTO mx
          FROM hermanos h WHERE h.codigo ~ ('^'||r.prefijo_codigo||'[0-9]+$');
        EXECUTE format('SELECT setval(%L, %s+1, false)', seq, mx);
      END IF;
    END IF;
  END LOOP;
END $$;
