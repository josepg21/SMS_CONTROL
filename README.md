# SMS_CONTROL

Sistema de seguimiento de produccion de prendas SMS.

## Despliegue en Streamlit Community Cloud (gratis)

1. Este repositorio debe ser **publico** en GitHub
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Inicia sesion con tu cuenta de GitHub
4. Haz clic en **"New app"**
5. Selecciona:
   - **Repository**: `josepg21/SMS_CONTROL`
   - **Branch**: `main`
   - **Main file path**: `app.py`
6. Haz clic en **"Deploy"**

La app estara disponible en una URL publica como:
`https://SMS_CONTROL-xxxxx.streamlit.app`

### Notas

- La base de datos SQLite se reinicia cuando la app esta inactiva. Si necesitas persistencia, usa el file uploader para cargar el Excel nuevamente.
- La app se duerme tras ~15 min de inactividad y se despierta automaticamente al visitarla.