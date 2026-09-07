"""Capa de persistencia para el seguimiento usando Supabase (Postgres).

Estructura: una tabla 'prendas' donde cada fila es una variante (Hoja + Style + Color),
con los datos importados del Excel y los campos editables de seguimiento.

Cada HOJA del Excel es un proyecto independiente. La clave primaria es:
    hoja @@ style @@ color

Credenciales: se leen de st.secrets (SUPABASE_URL y SUPABASE_KEY) o de
variables de entorno con el mismo nombre.
"""
import json
import os
import datetime

import streamlit as st
from supabase import create_client

TALLAS = ['one size', '4-8y', '0-6m', '6-12m', '3-6m', '6-9m', '9-12m', '2y', '3y', '4Y', '5y', '6Y', '8y', '10y']


def _credenciales():
    url = key = None
    try:
        url = st.secrets.get('SUPABASE_URL')
        key = st.secrets.get('SUPABASE_KEY')
    except Exception:
        pass
    if not url or not key:
        url = os.environ.get('SUPABASE_URL')
        key = os.environ.get('SUPABASE_KEY')
    if not url or not key:
        raise RuntimeError(
            'Faltan credenciales de Supabase. Define SUPABASE_URL y SUPABASE_KEY '
            'en .streamlit/secrets.toml o en variables de entorno.'
        )
    return url, key


def _client():
    url, key = _credenciales()
    client = create_client(url, key)
    # Restaurar sesion guardada para que el login persista entre reruns
    sesion = st.session_state.get('sb_session')
    if sesion:
        try:
            client.auth.set_session(sesion['access_token'], sesion['refresh_token'])
        except Exception:
            pass
    return client


def _id(hoja, style, color):
    return f'{hoja}@@{style}@@{color}'


def init_db():
    """Verifica la conexión y que la tabla prendas exista."""
    client = _client()
    try:
        client.table('prendas').select('id').limit(1).execute()
    except Exception as e:
        if 'does not exist' in str(e).lower() or 'relation' in str(e).lower():
            raise RuntimeError(
                'La tabla prendas no existe en Supabase. Abre el SQL Editor y ejecuta '
                'el contenido de supabase_schema.sql.'
            ) from e
        raise


def reemplazar_datos(hoja, filas):
    """Borrar e insertar las filas de UNA hoja (proyecto), sin tocar otras hojas."""
    client = _client()
    client.table('prendas').delete().eq('hoja', hoja).execute()
    ahora = datetime.datetime.now().isoformat(timespec='seconds')
    registros = []
    for f in filas:
        registros.append({
            'id': _id(hoja, f['style'], f['color']),
            'hoja': hoja,
            'style': f['style'],
            'po': f.get('po'),
            'name': f.get('name'),
            'tela': f.get('tela'),
            'color': f.get('color'),
            'tallas_pedido': json.dumps(f.get('pedido', [])),
            'total_pedido': f.get('pedido_total', 0),
            'tallas_prog': json.dumps(f.get('prog', [])),
            'total_prog': f.get('prog_total', 0),
            'status': f.get('status'),
            'ingreso': f.get('ingreso', ''),
            'observaciones': f.get('obs', ''),
            'fecha_importacion': ahora,
        })
    if registros:
        client.table('prendas').upsert(registros).execute()


def actualizar_seguimiento(hoja, style, color, status=None, observaciones=None,
                           total_prog=None, tallas_prog=None):
    """Actualiza campos editables de una variante y registra la fecha de última modificación."""
    updates = {}
    if status is not None:
        updates['status'] = status
    if observaciones is not None:
        updates['observaciones'] = observaciones
    if total_prog is not None:
        updates['total_prog'] = total_prog
    if tallas_prog is not None:
        updates['tallas_prog'] = json.dumps(tallas_prog)
    if not updates:
        return
    updates['ultima_actualizacion'] = datetime.datetime.now().isoformat(timespec='seconds')
    client = _client()
    client.table('prendas').update(updates).eq('id', _id(hoja, style, color)).execute()


def _parsear(fila):
    d = dict(fila)
    for clave in ('tallas_pedido', 'tallas_prog'):
        valor = d.get(clave)
        if isinstance(valor, str):
            try:
                d[clave.replace('tallas_', '')] = json.loads(valor or '[]')
            except (ValueError, TypeError):
                d[clave.replace('tallas_', '')] = []
        elif valor is None:
            d[clave.replace('tallas_', '')] = []
        else:
            d[clave.replace('tallas_', '')] = valor
    return d


def obtener_todos(hoja=None):
    """Devuelve todas las filas, o solo las de una hoja si se especifica."""
    client = _client()
    query = client.table('prendas').select('*')
    if hoja is not None:
        query = query.eq('hoja', hoja).order('style')
    else:
        query = query.order('hoja').order('style')
    res = query.execute()
    return [_parsear(r) for r in res.data]


def obtener_por_id(hoja, style, color):
    client = _client()
    res = client.table('prendas').select('*').eq('id', _id(hoja, style, color)).execute()
    if res.data:
        return _parsear(res.data[0])
    return None


def obtener_hojas_disponibles():
    """Lista de hojas distintas presentes en la DB."""
    client = _client()
    res = client.table('prendas').select('hoja').order('hoja').execute()
    hojas = []
    for r in res.data:
        if r['hoja'] not in hojas:
            hojas.append(r['hoja'])
    return hojas


def existe_datos(hoja=None):
    client = _client()
    query = client.table('prendas').select('id').limit(1)
    if hoja is not None:
        query = query.eq('hoja', hoja)
    res = query.execute()
    return len(res.data) > 0


def conteo_por_estado(hoja=None):
    client = _client()
    query = client.table('prendas').select('status')
    if hoja is not None:
        query = query.eq('hoja', hoja)
    res = query.execute()
    conteo = {}
    for r in res.data:
        s = r['status']
        conteo[s] = conteo.get(s, 0) + 1
    return conteo


# ---------- Autenticacion (Supabase Auth) ----------

def login(email, password):
    """Inicia sesion y guarda los tokens en session_state para persistir."""
    client = _client()
    res = client.auth.sign_in_with_password({'email': email, 'password': password})
    st.session_state['sb_session'] = {
        'access_token': res.session.access_token,
        'refresh_token': res.session.refresh_token,
    }
    st.session_state['sb_email'] = res.user.email
    return res.user.email


def logout():
    client = _client()
    st.session_state.pop('sb_session', None)
    st.session_state.pop('sb_email', None)
    try:
        client.auth.sign_out()
    except Exception:
        pass


def sesion_activa():
    """True si hay una sesion de Supabase activa (guardada en session_state)."""
    return bool(st.session_state.get('sb_session'))


def email_actual():
    return st.session_state.get('sb_email')