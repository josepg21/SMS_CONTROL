"""Capa de persistencia SQLite para el seguimiento.

Estructura: una tabla 'prendas' donde cada fila es una variante (Style + Color),
con los datos importados del Excel y los campos editables de seguimiento.
"""
import sqlite3
import json
import os
import datetime

TALLAS = ['one size', '4-8y', '0-6m', '6-12m', '3-6m', '6-9m', '9-12m', '2y', '3y', '4Y', '5y', '6Y', '8y', '10y']

SCHEMA = """
CREATE TABLE IF NOT EXISTS prendas (
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


def init_db():
    conn = _conn()
    conn.execute(SCHEMA)
    conn.commit()
    conn.close()


def _id(style, color):
    return f'{style}@@{color}'


def reemplazar_datos(filas):
    """Reemplaza todo el contenido importable (usado al importar un Excel)."""
    conn = _conn()
    conn.execute('DROP TABLE IF EXISTS prendas')
    conn.execute(SCHEMA)
    ahora = datetime.datetime.now().isoformat(timespec='seconds')
    for f in filas:
        conn.execute(
            """INSERT OR REPLACE INTO prendas
               (id, style, po, name, tela, color, tallas_pedido, total_pedido,
                tallas_prog, total_prog, status, ingreso, observaciones, fecha_importacion)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (_id(f['style'], f['color']), f['style'], f.get('po'), f.get('name'),
             f.get('tela'), f.get('color'), json.dumps(f.get('pedido', [])),
             f.get('pedido_total', 0), json.dumps(f.get('prog', [])),
             f.get('prog_total', 0), f.get('status'), f.get('ingreso', ''),
             f.get('obs', ''), ahora),
        )
    conn.commit()
    conn.close()


def actualizar_seguimiento(style, color, status=None, observaciones=None,
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
    params.append(_id(style, color))
    conn.execute(f'UPDATE prendas SET {", ".join(sets)} WHERE id = ?', params)
    conn.commit()
    conn.close()


def obtener_todos():
    conn = _conn()
    rows = conn.execute('SELECT * FROM prendas ORDER BY style').fetchall()
    conn.close()
    out = []
    for r in rows:
        d = dict(r)
        d['pedido'] = json.loads(d['tallas_pedido'] or '[]')
        d['prog'] = json.loads(d['tallas_prog'] or '[]')
        out.append(d)
    return out


def obtener_por_id(style, color):
    conn = _conn()
    row = conn.execute('SELECT * FROM prendas WHERE id = ?', (_id(style, color),)).fetchone()
    conn.close()
    if row:
        d = dict(row)
        d['pedido'] = json.loads(d['tallas_pedido'] or '[]')
        d['prog'] = json.loads(d['tallas_prog'] or '[]')
        return d
    return None


def existe_datos():
    conn = _conn()
    n = conn.execute('SELECT COUNT(*) FROM prendas').fetchone()[0]
    conn.close()
    return n > 0


def conteo_por_estado():
    conn = _conn()
    rows = conn.execute('SELECT status, COUNT(*) as n FROM prendas GROUP BY status').fetchall()
    conn.close()
    return {r['status']: r['n'] for r in rows}
