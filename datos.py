"""Extracción de datos desde el Excel de seguimiento SMS.

Lee cualquier hoja con formato "STATUS ..." que tenga los bloques de tallas
PEDIDO y PROGRAMADO, más columnas de estilo, tela, color, estado y observaciones.

La extracción localiza dinámicamente la fila de encabezados y las columnas,
de modo que tolera variaciones en el archivo.
"""
import openpyxl

# Estados posibles que se usan en el editor (normalizados)
ESTADOS_ORDEN = [
    'TELA X ING 3-SET',
    'EN LAVANDER' + chr(205) + 'A',
    'EN CORTE',
    'EN COSTURA',
    'EN ACABADOS',
    'ENCAJADO',
    '(sin estado)',
]

TALLAS = ['one size', '4-8y', '0-6m', '6-12m', '3-6m', '6-9m', '9-12m', '2y', '3y', '4Y', '5y', '6Y', '8y', '10y']


def _norm_text(val):
    if val is None:
        return ''
    s = str(val)
    # limpia acentos corruptos por codificación (ej. LAVANDERÍA mal escrito)
    s = s.replace('LAVANDERIA', 'LAVANDER' + chr(205) + 'A')
    s = s.replace('LAVANDER' + 'IA', 'LAVANDER' + chr(205) + 'A')
    return s.strip()


def _celda(ws, r, idx, key):
    """Devuelve el valor crudo de una celda si la columna `key` está detectada, si no None."""
    c = idx.get(key)
    if c is None:
        return None
    return ws.cell(row=r, column=c + 1).value


def _to_num(val):
    if val is None or isinstance(val, bool):
        return 0
    if isinstance(val, (int, float)):
        try:
            return int(val)
        except Exception:
            return 0
    s = str(val).replace(',', '').strip()
    try:
        return int(float(s))
    except Exception:
        return 0


def _row_values(ws, r):
    cols = ws.max_column
    vals = []
    for c in range(1, cols + 1):
        v = ws.cell(row=r, column=c).value
        vals.append(_norm_text(v) if not isinstance(v, (int, float)) else v)
    return vals


def _find_header_row(ws):
    """Localiza la fila que contiene 'Style' y 'PO' en las primeras columnas."""
    for r in range(1, min(ws.max_row, 30) + 1):
        vals = _row_values(ws, r)
        text = '|'.join(str(v) for v in vals if v is not None).upper()
        if 'STYLE' in text and ('PO' in text or 'OP' in text):
            return r
    return None


def _detect_layout(ws, header_row):
    """Detecta índices de columnas (0-based) a partir de la fila de encabezado."""
    hr = _row_values(ws, header_row)
    hr_lower = [str(v).strip().lower() if v else '' for v in hr]

    idx = {}
    # búsqueda de columna Style
    for i, h in enumerate(hr_lower):
        h_norm = h.replace('.', '').replace(' ', '').replace('-', '')
        if h == 'style':
            idx['style'] = i
        elif h_norm in ('po', 'op'):
            idx.setdefault('po', i)
        elif h == 'tendido':
            idx['tendido'] = i
        elif h in ('total',):
            # capturamos los dos 'total' (PEDIDO y PROGRAMADO) por orden
            idx.setdefault('totales', []).append(i)
        elif h in ('name', 'nombre'):
            idx['name'] = i
        elif h in ('tela principal', 'tela'):
            idx.setdefault('tela', i)
        elif h in ('color cuerpo', 'color'):
            idx.setdefault('color', i)
        elif h in ('status tela', 'status', 'estado'):
            idx.setdefault('status', i)
        elif h in ('ingreso cost', 'ingreso costura', 'ingreso'):
            idx.setdefault('ingreso', i)
        elif h in ('observaciones', 'obs', 'observacion'):
            idx['obs'] = i
        # ===== Formato WINTER (canales de venta, sin tallas) =====
        elif h in ('style name', 'nombre'):
            idx.setdefault('name', i)
        elif h in ('tot ped',):
            idx['tot_ped'] = i
        elif h in ('tot prg',):
            idx['tot_prog'] = i
    return idx


def extractar_hoja(file_path, sheet_name):
    """Extrae filas de la hoja dada, agrupando por Style + Color Cuerpo.

    Devuelve (filas, meta):
      filas: lista de dicts con los campos del seguimiento
      meta: dict con info (estilos, variantes, totales)
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    try:
        return _leer_hoja(wb, sheet_name)
    finally:
        wb.close()


def _leer_hoja(wb, sheet_name):
    ws = wb[sheet_name]

    header_row = _find_header_row(ws)
    if not header_row:
        raise ValueError(f'No se encontraron encabezados (Style/PO) en la hoja "{sheet_name}".')

    idx = _detect_layout(ws, header_row)

    # Detectar formato WINTER (canales de venta, sin bloques de tallas).
    # Se reconoce por la presencia de 'Tot Ped' / 'Tot Prg' y ausencia de
    # bloques PEDIDO/PROGRAMADO completos.
    es_winter = ('tot_ped' in idx and 'tot_prog' in idx
                 and 'style' in idx and 'color' in idx)
    if es_winter:
        return _extraer_winter(ws, header_row, idx)

    # localizar bloques de tallas: buscamos en la fila inmediatamente anterior
    # (o la misma) la palabra PEDIDO / PROGRAMADO
    block_pedido = None
    block_prog = None
    for r in range(max(1, header_row - 3), header_row + 1):
        rv = _row_values(ws, r)
        joined = [str(v).upper() if v else '' for v in rv]
        for i, t in enumerate(joined):
            if t == 'PEDIDO':
                block_pedido = i
            elif t == 'PROGRAMADO':
                block_prog = i
    if block_pedido is None or block_prog is None:
        # fallback: el primer bloque de tallas es pedido, el segundo programado
        if 'totales' in idx and len(idx['totales']) >= 2:
            block_pedido = idx['totales'][0]
            block_prog = idx['totales'][1]

    if block_pedido is None or block_prog is None:
        raise ValueError('No se pudieron localizar los bloques PEDIDO/PROGRAMADO de tallas.')

    filas = []
    for r in range(header_row + 1, ws.max_row + 1):
        style = _norm_text(ws.cell(row=r, column=idx['style'] + 1).value)
        po = _norm_text(_celda(ws, r, idx, 'po'))
        if not style:
            continue
        fila = {
            'style': style,
            'po': po,
            'name': _norm_text(ws.cell(row=r, column=idx['name'] + 1).value if idx.get('name') else None),
            'tela': _norm_text(ws.cell(row=r, column=idx['tela'] + 1).value if idx.get('tela') else None),
            'color': _norm_text(ws.cell(row=r, column=idx['color'] + 1).value if idx.get('color') else None),
        }
        # pedido y programado por talla
        fila['pedido'] = []
        fila['prog'] = []
        for j in range(len(TALLAS)):
            pc = block_pedido + j
            gc = block_prog + j
            fila['pedido'].append(_to_num(ws.cell(row=r, column=pc + 1).value) if pc < ws.max_column else 0)
            fila['prog'].append(_to_num(ws.cell(row=r, column=gc + 1).value) if gc < ws.max_column else 0)
        fila['pedido_total'] = _to_num(ws.cell(row=r, column=block_pedido + len(TALLAS) + 1).value)
        fila['prog_total'] = _to_num(ws.cell(row=r, column=block_prog + len(TALLAS) + 1).value)
        if fila['pedido_total'] == 0:
            fila['pedido_total'] = sum(fila['pedido'])
        if fila['prog_total'] == 0:
            fila['prog_total'] = sum(fila['prog'])
        fila['status'] = _norm_text(ws.cell(row=r, column=idx['status'] + 1).value if idx.get('status') else None) or '(sin estado)'
        ingreso_raw = ws.cell(row=r, column=idx['ingreso'] + 1).value if idx.get('ingreso') else None
        if hasattr(ingreso_raw, 'strftime'):
            ingreso = ingreso_raw.strftime('%d-%m-%Y')
        else:
            ingreso = _norm_text(ingreso_raw)
        fila['ingreso'] = ingreso
        fila['obs'] = _norm_text(ws.cell(row=r, column=idx['obs'] + 1).value if idx.get('obs') else None)
        filas.append(fila)

    # normalizar estado
    for f in filas:
        f['status'] = _normalizar_estado(f['status'])

    return filas


def _normalizar_estado(st):
    st = _norm_text(st)
    if not st:
        return '(sin estado)'
    for k in ['TELA X ING 3-SET', 'TELA X ING', 'X ING 3-SET', 'X INGRESAR']:
        if k in st:
            return 'TELA X ING 3-SET'
    if 'LAVANDER' in st:
        return 'EN LAVANDER' + chr(205) + 'A'
    if 'CORTE' in st:
        return 'EN CORTE'
    if 'COSTUR' in st:
        return 'EN COSTURA'
    if 'ACABAD' in st:
        return 'EN ACABADOS'
    if 'ENCAJAD' in st:
        return 'ENCAJADO'
    return st


def obtener_hojas(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def _extraer_winter(ws, header_row, idx):
    """Extrae filas de una hoja en formato WINTER (canales de venta, sin tallas).

    Los campos numéricos (canales de venta) no son tallas. Para mantener la
    compatibilidad con el resto de la app (db.py, reporte.py, app.py), las
    listas pedido[]/prog[] se llenan con ceros (14 posiciones) y los totales reales
    se guardan en pedido_total / prog_total.

    Devuelve una lista de dicts con la misma estructura que el resto de hojas.
    Cada variante corresponde a Style + Color + Destino.
    """
    # Localizar columnas de destino: entre la columna 'color' y 'tot_ped' (exclusive)
    col_idx = idx.get('color')
    col_totped = idx.get('tot_ped')
    destinos = []
    if col_idx is not None and col_totped is not None:
        for c in range(col_idx + 1, col_totped):
            cab = _norm_text(ws.cell(row=header_row, column=c + 1).value)
            if cab:
                destinos.append((c, cab))
    col_prg = idx.get('tot_prog')

    filas = []
    for r in range(header_row + 1, ws.max_row + 1):
        style = _norm_text(ws.cell(row=r, column=idx['style'] + 1).value)
        if not style:
            continue
        color_base = _norm_text(ws.cell(row=r, column=idx['color'] + 1).value if idx.get('color') else None)
        row_tot_ped = _to_num(ws.cell(row=r, column=col_totped + 1).value) if col_totped is not None else 0
        row_tot_prg = _to_num(ws.cell(row=r, column=col_prg + 1).value) if col_prg is not None else 0
        status = _normalizar_estado(
            _norm_text(ws.cell(row=r, column=idx['status'] + 1).value if idx.get('status') else None)
            or '(sin estado)')
        po_base = _norm_text(_celda(ws, r, idx, 'po'))
        name_base = _norm_text(_celda(ws, r, idx, 'name'))
        tela_base = _norm_text(_celda(ws, r, idx, 'tela'))

        # Una variante por cada destino con cantidad > 0
        for (ecol, dname) in destinos:
            cant = _to_num(ws.cell(row=r, column=ecol + 1).value)
            if cant <= 0:
                continue
            # Repartir el programado proporcional al pedido del destino
            prog = 0
            if row_tot_ped:
                prog = round(row_tot_prg * cant / row_tot_ped)
            fila = {
                'style': style,
                'po': po_base,
                'name': name_base,
                'tela': tela_base,
                'color': f'{color_base} — {dname}',
                'pedido': [0] * len(TALLAS),
                'prog': [0] * len(TALLAS),
                'pedido_total': cant,
                'prog_total': prog,
                'status': status,
                'ingreso': '',
                'obs': _norm_text(ws.cell(row=r, column=idx['obs'] + 1).value if idx.get('obs') else None),
            }
            filas.append(fila)
    return filas
