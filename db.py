"""Capa de persistencia SQLite para el seguimiento.

Estructura: una tabla 'prendas' donde cada fila es una variante (Hoja + Style + Color),
con los datos importados del Excel y los campos editables de seguimiento.

Cada HOJA del Excel es un proyecto independiente. La clave primaria es:
    hoja @@ style @@ color
"""
import sqlite3
import json
import os
import datetime

TALLAS = ['one size', '4-8y', '0-6m', '6-12m', '3-6m', '6-9m', '9-12m', '2y', '3y', '4Y', '5y', '6Y', '8y', '10y']

SCHEMA = """
CREATE TABLE IF NOT EXISTS prendas (
    hoja TEXT NOT NULL,
    id TEXT PRIMARY KEY,
    style TEXT NOT NULL,
    po TEXT,
    name TEXT,
    tela TEXT,
    color TEXT,
    tallas_pedido TEXT,
    total_pedido INTEGER DEFAULT 0,
    tallas_prog TEXT,
    total_prog INTEGER DEFAULT 0,
    status TEXT,
    ingreso TEXT,
    observaciones TEXT DEFAULT '',
    fecha_importacion TEXT
);
"""


def _db_path():
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, 'seguimiento.db')


def _conn():
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    return conn


def _columnas(conn):
    rows = conn.execute('PRAGMA table_info(prendas)').fetchall()
    return [r['name'] for r in rows]


def _migrar(conn):
    """Si la DB es vieja (sin columna 'hoja'), se dropea y recrea con el nuevo schema."""
    cols = _columnas(conn)
    if 'hoja' not in cols:
        conn.execute('DROP TABLE IF EXISTS prendas')
        conn.execute(SCHEMA)
        conn.commit()


def init_db():
    conn = _conn()
    _migrar(conn)
    conn.execute(SCHEMA)
    conn.commit()
    conn.close()


def _id(hoja, style, color):
    return f'{hoja}@@{style}@@{color}'


def reemplazar_datos(hoja, filas):
    """Borrar e insertar las filas de UNA hoja (proyecto), sin tocar otras hojas."""
    conn = _conn()
    conn.execute('DELETE FROM prendas WHERE hoja = ?', (hoja,))
    ahora = datetime.datetime.now().isoformat(timespec='seconds')
    for f in filas:
        conn.execute(
            """INSERT OR REPLACE INTO prendas
               (hoja, id, style, po, name, tela, color, tallas_pedido, total_pedido,
                tallas_prog, total_prog, status, ingreso, observaciones, fecha_importacion)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (hoja, _id(hoja, f['style'], f['color']), f['style'], f.get('po'), f.get('name'),
             f.get('tela'), f.get('color'), json.dumps(f.get('pedido', [])),
             f.get('pedido_total', 0), json.dumps(f.get('prog', [])),
             f.get('prog_total', 0), f.get('status'), f.get('ingreso', ''),
             f.get('obs', ''), ahora),
        )
    conn.commit()
    conn.close()


def actualizar_seguimiento(hoja, style, color, status=None, observaciones=None,
                           total_prog=None, tallas_prog=None):
    """Actualiza campos editables de una variante."""
    conn = _conn()
    sets = []
    params = []
    if status is not None:
        sets.append('status = ?')
        params.append(status)
    if observaciones is not None:
        sets.append('observaciones = ?')
        params.append(observaciones)
    if total_prog is not None:
        sets.append('total_prog = ?')
        params.append(total_prog)
    if tallas_prog is not None:
        sets.append('tallas_prog = ?')
        params.append(json.dumps(tallas_prog))
    if not sets:
        conn.close()
        return
    params.append(_id(hoja, style, color))
    conn.execute(f'UPDATE prendas SET {", ".join(sets)} WHERE id = ?', params)
    conn.commit()
    conn.close()


def obtener_todos(hoja=None):
    """Devuelve todas las filas, o solo las de una hoja si se especifica."""
    conn = _conn()
    if hoja is not None:
        rows = conn.execute('SELECT * FROM prendas WHERE hoja = ? ORDER BY style', (hoja,)).fetchall()
    else:
        rows = conn.execute('SELECT * FROM prendas ORDER BY hoja, style').fetchall()
    conn.close()
    out = []
    for r in rows:
        d = dict(r)
        d['pedido'] = json.loads(d['tallas_pedido'] or '[]')
        d['prog'] = json.loads(d['tallas_prog'] or '[]')
        out.append(d)
    return out


def obtener_por_id(hoja, style, color):
    conn = _conn()
    row = conn.execute('SELECT * FROM prendas WHERE id = ?', (_id(hoja, style, color),)).fetchone()
    conn.close()
    if row:
        d = dict(row)
        d['pedido'] = json.loads(d['tallas_pedido'] or '[]')
        d['prog'] = json.loads(d['tallas_prog'] or '[]')
        return d
    return None


def obtener_hojas_disponibles():
    """Lista de hojas distintas presentes en la DB."""
    conn = _conn()
    rows = conn.execute('SELECT DISTINCT hoja FROM prendas ORDER BY hoja').fetchall()
    conn.close()
    return [r['hoja'] for r in rows]


def existe_datos(hoja=None):
    conn = _conn()
    if hoja is not None:
        n = conn.execute('SELECT COUNT(*) FROM prendas WHERE hoja = ?', (hoja,)).fetchone()[0]
    else:
        n = conn.execute('SELECT COUNT(*) FROM prendas').fetchone()[0]
    conn.close()
    return n > 0


def conteo_por_estado(hoja=None):
    conn = _conn()
    if hoja is not None:
        rows = conn.execute('SELECT status, COUNT(*) as n FROM prendas WHERE hoja = ? GROUP BY status',
                            (hoja,)).fetchall()
    else:
        rows = conn.execute('SELECT status, COUNT(*) as n FROM prendas GROUP BY status').fetchall()
    conn.close()
    return {r['status']: r['n'] for r in rows}
