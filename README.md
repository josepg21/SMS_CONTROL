# SMS_CONTROL

Sistema de seguimiento de produccion de prendas SMS.

La app usa **Supabase** (Postgres en la nube) para almacenar los datos de forma
permanente. No depende de archivos locales.

## 1. Configuracion de Supabase (una sola vez)

1. Crea un proyecto gratis en [supabase.com](https://supabase.com)
2. En el **SQL Editor**, pega y ejecuta el contenido de `supabase_schema.sql`
3. En **Settings > API** copia:
   - **Project URL** (ej: `https://abcdxyz.supabase.co`)
   - **anon public** key

## 2. Credenciales locales (para desarrollo)

Copia `.streamlit/secrets.toml` y rellena tus credenciales:

```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_KEY = "tu-anon-public-key"
```

## 3. Despliegue en Streamlit Community Cloud (gratis)

1. Este repositorio debe ser **publico** en GitHub
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Inicia sesion con tu cuenta de GitHub
4. Haz clic en **"New app"**
5. Selecciona:
   - **Repository**: `josepg21/SMS_CONTROL`
   - **Branch**: `main`
   - **Main file path**: `app.py`
6. Antes de desplegar, agrega las credenciales en **Settings > Secrets**:
   `SUPABASE_URL` y `SUPABASE_KEY` con los valores de tu proyecto
7. Haz clic en **"Deploy"**

La app estara disponible en una URL publica como:
`https://SMS_CONTROL-xxxxx.streamlit.app`

### Notas

- Los datos viven en Supabase: **no se pierden** cuando la app se duerme o se re-despliega.
- La app se duerme tras ~15 min de inactividad y se despierta automaticamente al visitarla.
- La anon key esta expuesta en el cliente (es normal en apps sin login); cualquiera con la URL puede leer/escribir datos. Para proteger la app, agrega autenticacion de Supabase (Auth) mas adelante.