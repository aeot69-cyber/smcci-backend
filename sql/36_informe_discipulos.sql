-- 36_informe_discipulos.sql
-- Asistencia por disc�pulo en el Informe de C�lula: en vez de dos textos
-- libres, el formulario muestra una fila por disc�pulo con checkbox
-- "vino / no vino" y un motivo opcional para los que no vinieron.
--
-- Se conservan las columnas de texto `discipulos_asistieron` /
-- `discipulos_no_asistieron`: la aplicaci�n las sigue escribiendo (resumen
-- "Juan, Mar�a" / "Pedro - trabajo") para que reportes, Excel, PDF y el
-- conteo de asistencia sigan funcionando sin cambios. Aqu� queda el dato
-- estructurado por disc�pulo para futuras estad�sticas.
--
-- `discipulo_id` es NULL en las filas de "otros" (nombres escritos a mano
-- que no est�n en la lista del l�der). Los informes antiguos no tienen
-- filas: al editarlos la aplicaci�n reconstruye las casillas desde los textos.

BEGIN;

CREATE TABLE IF NOT EXISTS celulas_informe_discipulos (
    id            serial PRIMARY KEY,
    informe_id    integer NOT NULL REFERENCES celulas_informe(id) ON DELETE CASCADE,
    discipulo_id  integer REFERENCES hermanos(id) ON DELETE SET NULL,
    nombre        text NOT NULL,
    asistio       boolean NOT NULL DEFAULT true,
    justificacion text,
    created_at    timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_informe_discipulos_informe
    ON celulas_informe_discipulos(informe_id);

-- Un disc�pulo no puede repetirse dentro del mismo informe.
CREATE UNIQUE INDEX IF NOT EXISTS uq_informe_discipulo
    ON celulas_informe_discipulos(informe_id, discipulo_id)
    WHERE discipulo_id IS NOT NULL;

COMMIT;
