-- Ejecutar en Supabase SQL Editor después de confirmar que existen
-- public.p y public.configuracion_pqrs.

-- La fila se conserva como configuración provisional. El administrador
-- podrá cambiar siguiente_numero desde /admin antes de iniciar operaciones.
INSERT INTO public.configuracion_pqrs (id, siguiente_numero)
VALUES (TRUE, 4)
ON CONFLICT (id) DO NOTHING;

-- La función toma un número y lo incrementa en una sola operación atómica.
CREATE OR REPLACE FUNCTION public.tomar_siguiente_numero()
RETURNS INTEGER
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  numero_tomado INTEGER;
BEGIN
  UPDATE public.configuracion_pqrs
  SET siguiente_numero = siguiente_numero + 1
  WHERE id = TRUE
  RETURNING siguiente_numero - 1 INTO numero_tomado;

  IF numero_tomado IS NULL THEN
    RAISE EXCEPTION 'No existe la configuración del contador de PQRS.';
  END IF;

  RETURN numero_tomado;
END;
$$;

REVOKE ALL ON FUNCTION public.tomar_siguiente_numero() FROM PUBLIC;
REVOKE ALL ON FUNCTION public.tomar_siguiente_numero() FROM anon;
REVOKE ALL ON FUNCTION public.tomar_siguiente_numero() FROM authenticated;
GRANT EXECUTE ON FUNCTION public.tomar_siguiente_numero() TO service_role;

-- El backend accede con la clave secreta de Supabase (solo en servidor).
-- No colocar esa clave en archivos públicos ni en el navegador.
