"""Exportación de reportes (Excel y PDF) desde los datos de seguimiento."""
import os
import io
import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image
from reportlab.lib.enums import TA_CENTER

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

TALLAS = ['one size', '4-8y', '0-6m', '6-12m', '3-6m', '6-9m', '9-12m', '2y', '3y', '4Y', '5y', '6Y', '8y', '10y']

ESTADOS_ORDEN = ['TELA X ING 3-SET', 'EN LAVANDER' + chr(205) + 'A', 'EN CORTE', 'EN COSTURA', 'EN ACABADOS', 'ENCAJADO']


def _proc(datos):
    """Convierte filas de DB a estructura ordenada por estilo -> colores."""
    from collections import OrderedDict
    estilos = OrderedDict()
    for d in datos:
        es = estilos.setdefault(d['style'], {'name': d['name'], 'tela': d['tela'], 'colores': OrderedDict()})
        es['colores'].setdefault(d['color'], d)
    return estilos


def _totales(datos):
    tp = sum(d['total_pedido'] or 0 for d in datos)
    tg = sum(d['total_prog'] or 0 for d in datos)
    n_estilos = len({d['style'] for d in datos})
    n_var = len(datos)
    multi = sum(1 for d in datos)
    # estilos con más de una variante
    from collections import Counter
    c = Counter(d['style'] for d in datos)
    multi = sum(1 for k, v in c.items() if v > 1)
    return tp, tg, n_estilos, n_var, multi


def exportar_excel(datos):
    """Devuelve bytes de un archivo .xlsx de reporte."""
    estilos = _proc(datos)
    tp, tg, n_est, n_var, multi = _totales(datos)

    out = Workbook()
    out.remove(out.active)

    TITULO_FONT = Font(bold=True, size=13, color='FFFFFF')
    TITULO_FILL = PatternFill('solid', fgColor='1F4E78')
    HEAD_FONT = Font(bold=True, size=10, color='FFFFFF')
    HEAD_FILL = PatternFill('solid', fgColor='2E75B6')
    GRUPO_FONT = Font(bold=True, size=10, color='FFFFFF')
    GRUPO_FILL = PatternFill('solid', fgColor='548235')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal='center', vertical='center')
    LEFT = Alignment(horizontal='left', vertical='center')

    # ---- Resumen ----
    ws1 = out.create_sheet('RESUMEN')
    ws1['A1'] = 'REPORTE SEGUIMIENTO SMS'
    ws1['A1'].font = TITULO_FONT; ws1['A1'].fill = TITULO_FILL
    ws1.merge_cells('A1:H1')
    headers = ['#', 'Style', 'Prenda', '# Variantes', 'Total PEDIDO', 'Total PROGRAMADO', 'Tela', 'Estados']
    for j, h in enumerate(headers, start=1):
        c = ws1.cell(3, j, h); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = BORDER
    r = 4
    for n, (styl, es) in enumerate(estilos.items(), start=1):
        estados = {}
        for v in es['colores'].values():
            st = v['status'] or '(sin estado)'
            estados[st] = estados.get(st, 0) + 1
        estados_str = ', '.join(f'{k} ({v})' for k, v in estados.items())
        vals = [n, styl, es['name'], len(es['colores']),
                sum(v['total_pedido'] or 0 for v in es['colores'].values()),
                sum(v['total_prog'] or 0 for v in es['colores'].values()),
                es['tela'], estados_str]
        for j, val in enumerate(vals, start=1):
            c = ws1.cell(r, j, val); c.border = BORDER
            c.alignment = LEFT if j in (3, 7, 8) else CENTER
        r += 1
    ws1.cell(r, 1, 'TOTAL').font = GRUPO_FONT
    ws1.cell(r, 4, str(n_est)).font = GRUPO_FONT
    ws1.cell(r, 5, tp).font = GRUPO_FONT
    ws1.cell(r, 6, tg).font = GRUPO_FONT
    for col in range(1, 9):
        ws1.cell(r, col).fill = GRUPO_FILL
        ws1.cell(r, col).border = BORDER
    for j, w in enumerate([5, 12, 18, 15, 14, 16, 22, 30], start=1):
        ws1.column_dimensions[get_column_letter(j)].width = w

    # ---- Detalle ----
    ws2 = out.create_sheet('DETALLE')
    ws2['A1'] = 'SEGUIMIENTO POR ESTILO Y COLOR'
    ws2['A1'].font = TITULO_FONT; ws2['A1'].fill = TITULO_FILL
    ws2.merge_cells('A1:M1')
    det_headers = ['Style', 'Color', 'Prenda', 'Tela', 'Total PEDIDO', 'Total PROGRAMADO', 'STATUS', 'Ingreso', 'Observaciones', 'Dif', 'PO', 'Tendido']
    for j, h in enumerate(det_headers, start=1):
        c = ws2.cell(3, j, h); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = BORDER
    r = 4
    for styl, es in estilos.items():
        ws2.cell(r, 1, styl).font = GRUPO_FONT
        for col in range(1, 13):
            ws2.cell(r, col).fill = GRUPO_FILL; ws2.cell(r, col).border = BORDER
        ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        r += 1
        for color, f in es['colores'].items():
            dif = int(f['total_prog'] or 0) - int(f['total_pedido'] or 0)
            vals = [styl, color, f['name'], f['tela'], f['total_pedido'], f['total_prog'],
                    f['status'], f['ingreso'], f['observaciones'], dif, f['po'], f.get('tendido', '')]
            for j, val in enumerate(vals, start=1):
                c = ws2.cell(r, j, val); c.border = BORDER
                c.alignment = LEFT if j in (2, 3, 4, 9) else CENTER
            if dif < 0:
                ws2.cell(r, 10).fill = PatternFill('solid', fgColor='FFC7CE')
                ws2.cell(r, 10).font = Font(color='9C0006')
            elif dif > 0:
                ws2.cell(r, 10).fill = PatternFill('solid', fgColor='C6EFCE')
                ws2.cell(r, 10).font = Font(color='006100')
            r += 1
        r += 1
    for j, w in enumerate([12, 22, 18, 20, 12, 15, 15, 12, 40, 8, 12, 12], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    buf = io.BytesIO()
    out.save(buf)
    buf.seek(0)
    return buf.getvalue()


def exportar_pdf(datos):
    """Devuelve bytes de un archivo .pdf de resumen con gráficos."""
    estilos = _proc(datos)
    tp, tg, n_est, n_var, multi = _totales(datos)

    # conteo por estado
    from collections import Counter
    c_est = Counter(d['status'] or '(sin estado)' for d in datos)
    order = ESTADOS_ORDEN + [k for k in c_est if k not in ESTADOS_ORDEN and k != '(sin estado)']
    order = [k for k in order if k in c_est]
    if '(sin estado)' in c_est:
        order.append('(sin estado)')
    labels = order
    vals = [c_est[k] for k in order]
    if not vals:
        vals = [0]
        labels = ['(sin datos)']

    # chart 1: estados
    tmp1 = os.path.join(os.environ.get('TEMP', '/tmp'), 'sms_chart1.png')
    fig, ax = plt.subplots(figsize=(7.5, 3.6), dpi=150)
    cs = ['#2E75B6', '#FFC000', '#ED7D31', '#A9D18E', '#BFBFBF', '#C00000']
    bars = ax.bar(labels, vals, color=cs[:len(labels)], edgecolor='white')
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width()/2, v + 0.6, str(v), ha='center', va='bottom', fontsize=9, fontweight='bold')
    ax.set_ylabel('Variantes')
    ax.set_title('Variantes por Estado (Seguimiento)', fontweight='bold', fontsize=10)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.set_ylim(0, max(vals) * 1.15)
    plt.xticks(rotation=15, fontsize=8, ha='right')
    plt.tight_layout(); plt.savefig(tmp1, bbox_inches='tight'); plt.close()

    # chart 2: top 10 por programado
    top = sorted(estilos.items(), key=lambda kv: -sum(v['total_prog'] or 0 for v in kv[1]['colores'].values()))[:10]
    tlabels = [k for k, _ in top]
    tprog = [sum(v['total_prog'] or 0 for v in kv['colores'].values()) for _, kv in top]
    tped = [sum(v['total_pedido'] or 0 for v in kv['colores'].values()) for _, kv in top]
    tmp2 = os.path.join(os.environ.get('TEMP', '/tmp'), 'sms_chart2.png')
    fig, ax = plt.subplots(figsize=(7.5, 3.6), dpi=150)
    x = range(len(tlabels))
    ax.bar([i - 0.15 for i in x], tprog, width=0.3, label='Programado', color='#2E75B6')
    ax.bar([i + 0.15 for i in x], tped, width=0.3, label='Pedido', color='#A9D18E')
    ax.set_xticks(list(x)); ax.set_xticklabels([t[:8] for t in tlabels], fontsize=8, rotation=30, ha='right')
    ax.set_ylabel('Unidades')
    ax.set_title('Top 10 Estilos: Programado vs Pedido', fontweight='bold', fontsize=10)
    ax.legend(fontsize=8)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    plt.tight_layout(); plt.savefig(tmp2, bbox_inches='tight'); plt.close()

    fecha = datetime.date.today().strftime('%d/%m/%Y')
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    titulo = ParagraphStyle('titulo', parent=styles['Title'], fontName='Helvetica-Bold',
                            fontSize=18, alignment=TA_CENTER, textColor=colors.HexColor('#1F4E78'))
    subtitulo = ParagraphStyle('sub', parent=styles['Normal'], alignment=TA_CENTER,
                               fontSize=10, textColor=colors.HexColor('#555555'))
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontName='Helvetica-Bold',
                        fontSize=13, textColor=colors.HexColor('#1F4E78'), spaceBefore=8, spaceAfter=4)
    body = styles['BodyText']

    story = []
    story.append(Paragraph('REPORTE DE SEGUIMIENTO SMS', titulo))
    story.append(Spacer(1, 2))
    story.append(Paragraph(f'Resumen de avance - Generado el {fecha}', subtitulo))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width='100%', thickness=1, color=colors.HexColor('#1F4E78')))
    story.append(Spacer(1, 10))
    story.append(Paragraph('Resumen General', h2))
    res_data = [
        ['Estilos distintos', str(n_est)],
        ['Variantes (por color)', str(n_var)],
        ['Estilos con múltiples variantes', str(multi)],
        ['Total unidades PEDIDO', str(tp)],
        ['Total unidades PROGRAMADO', str(tg)],
        ['Avance de programación', f'{int(tg / tp * 100)}%' if tp else '0%'],
    ]
    rt = Table(res_data, colWidths=[70*mm, 40*mm])
    rt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F2F7FB')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BFBFBF')),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica-Bold'),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.white, colors.HexColor('#EAF1F8')]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(rt)
    story.append(Spacer(1, 8))

    story.append(Paragraph('Distribución por estado', h2))
    mayor = labels[0] if labels else '-'
    story.append(Paragraph(
        f'Del total de <b>{n_var}</b> variantes, el estado predominante es <b>{mayor}</b> '
        f'con <b>{c_est.get(mayor, 0)}</b> variantes, seguido de los colores restantes. '
        f'Se registran <b>{tp}</b> unidades pedidas y <b>{tg}</b> programadas.', body))
    story.append(Spacer(1, 6))
    story.append(Image(tmp1, width=175*mm, height=84*mm))
    story.append(Spacer(1, 12))

    story.append(Paragraph('Top 10 estilos por volumen programado', h2))
    if top:
        story.append(Paragraph(
            f'Los estilos con mayor volumen programado son prendas de bebé y pantalones; '
            f'el líder es <b>{tlabels[0]}</b> ({top[0][1]["name"]}) con '
            f'<b>{tprog[0]}</b> unidades.', body))
        story.append(Spacer(1, 6))
        story.append(Image(tmp2, width=175*mm, height=84*mm))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()
