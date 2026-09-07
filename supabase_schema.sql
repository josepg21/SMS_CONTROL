-- Esquema para Supabase SQL Editor
-- Pegar esto en: Dashboard Supabase > SQL Editor > New query > Run

-- "tallas_pedido" y "tallas_prog" se guardan como texto JSON (igual que en SQLite).
CREATE TABLE IF NOT EXISTS public.prendas (
    id TEXT PRIMARY KEY,
    hoja TEXT NOT NULL,
    style TEXT NOT NULL,
    po TEXT,
    name TEXT,
    tela TEXT,
    color TEXT,
    tallas_pedido TEXT DEFAULT '[]',
    total_pedido INTEGER DEFAULT 0,
    tallas_prog TEXT DEFAULT '[]',
    total_prog INTEGER DEFAULT 0,
    status TEXT,
    ingreso TEXT,
    observaciones TEXT DEFAULT '',
    fecha_importacion TEXT,
    ultima_actualizacion TEXT
);

-- Índice para búsquedas por hoja proyecto
CREATE INDEX IF NOT EXISTS idx_prendas_hoja ON public.prendas (hoja);
CREATE INDEX IF NOT EXISTS idx_prendas_status ON public.prendas (status);

-- RLS: la app usa la anon key, así que permitimos todo con anon/authenticated.
-- NOTA: esto significa que cualquiera con la URL puede leer/escribir. 
-- La app no tiene sistema de login propio.
ALTER TABLE public.prendas ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Permitir todo anon" ON public.prendas
    FOR ALL
    USING (true)
    WITH CHECK (true);

-- Opcional: forzar timestamps para fechas de importación/actualización
-- (los deja la app, pero ayuda si alguien inserta directo por el dashboard)
CREATE OR REPLACE FUNCTION public.set_fechas_defaults()
RETURNS TRIGGER AS $$
BEGIN
    IF NEW.fecha_importacion IS NULL THEN
        NEW.fecha_importacion := to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS');
    END IF;
    IF NEW.ultima_actualizacion IS NULL THEN
        NEW.ultima_actualizacion := to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS');
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_prendas_fechas ON public.prendas;
CREATE TRIGGER trg_prendas_fechas
    BEFORE INSERT OR UPDATE ON public.prendas
    FOR EACH ROW EXECUTE FUNCTION public.set_fechas_defaults();