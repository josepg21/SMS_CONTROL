import openpyxl
from collections import OrderedDict, Counter
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image,
                                Table, TableStyle, HRFlowable)
from reportlab.lib.enums import TA_CENTER

SRC = r'C:\Users\PERUVIAN FIT\Desktop\SEGUIMIENTOSMS\MISHA SMS 1 SE.xlsx'
OUT_PDF = r'C:\Users\PERUVIAN FIT\Desktop\SEGUIMIENTOSMS\REPORTE SMS-SUMMER27.pdf'
TMP_CHART1 = r'C:\Users\PERUVIAN FIT\AppData\Local\Temp\opencode\chart_status.png'
TMP_CHART2 = r'C:\Users\PERUVIAN FIT\AppData\Local\Temp\opencode\chart_top.png'

IDX_PED_TOTAL=17; IDX_PROG_TOTAL=32; IDX_NAME=33; IDX_TELA=34; IDX_COLOR=35; IDX_STATUS=36

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['STATUS SMS-SUMMER27']

estilos = OrderedDict()
status_counts = Counter()
for row in ws.iter_rows(min_row=13, values_only=True):
    styl = row[1]
    if not styl:
        continue
    color = str(row[IDX_COLOR]).strip() if row[IDX_COLOR] else ''
    status = str(row[IDX_STATUS]).strip() if row[IDX_STATUS] else '(sin estado)'
    pt = row[IDX_PED_TOTAL] or 0
    pgt = row[IDX_PROG_TOTAL] or 0
    status_counts[status] += 1
    es = estilos.setdefault(str(styl).strip(), {'nombre': str(row[IDX_NAME]).strip() if row[IDX_NAME] else '',
                                                'tela': str(row[IDX_TELA]).strip() if row[IDX_TELA] else '',
                                                'colores': OrderedDict()})
    es['colores'].setdefault(color, {'ped': pt, 'prog': pgt, 'status': status})

tot_ped = sum(sum(v['ped'] for v in e['colores'].values()) for e in estilos.values())
tot_prog = sum(sum(v['prog'] for v in e['colores'].values()) for e in estilos.values())
n_estilos = len(estilos)
n_variantes = sum(len(e['colores']) for e in estilos.values())
multi = sum(1 for e in estilos.values() if len(e['colores']) > 1)

# Orden de estados para trazado (normalizar acento corrupto)
status_clean = OrderedDict()
for k, v in status_counts.items():
    key = k.replace('LAVANDERIA', 'LAVANDER' + chr(205) + 'A').replace('LAVANDER' + 'IA', 'LAVANDER' + chr(205) + 'A').strip()
    status_clean[key] = status_clean.get(key, 0) + v
order = ['EN CORTE', 'EN COSTURA', 'EN ACABADOS', 'EN LAVANDER' + chr(205) + 'A', 'TELA X ING 3-SET']
temp = []
for st in order:
    for k, v in status_clean.items():
        if k == st:
            temp.append((st, v)); break
temp.sort(key=lambda x: -x[1])
labels = [x[0] for x in temp]
vals = [x[1] for x in temp]
if not labels:
    labels = list(status_clean.keys())
    vals = list(status_clean.values())

# ------- GRAFICO 1: barras por estado (cantidad de variantes) -------
fig, ax = plt.subplots(figsize=(7.5, 3.6), dpi=150)
colors_bars = ['#2E75B6', '#FFC000', '#ED7D31', '#A9D18E', '#BFBFBF']
bars = ax.bar(labels, vals, color=colors_bars[:len(labels)], edgecolor='white')
for b, v in zip(bars, vals):
    ax.text(b.get_x()+b.get_width()/2, v+0.8, str(v), ha='center', va='bottom', fontsize=9, fontweight='bold')
ax.set_ylabel('Variantes', fontsize=9)
ax.set_title('Variantes por Estado de Tela (SMS-SUMMER27)', fontsize=10, fontweight='bold')
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.set_ylim(0, max(vals)*1.15)
plt.xticks(rotation=15, fontsize=8, ha='right')
plt.tight_layout()
plt.savefig(TMP_CHART1, bbox_inches='tight')
plt.close()

# ------- GRAFICO 2: Top 10 estilos por programado -------
top = sorted(estilos.items(), key=lambda kv: -sum(v['prog'] for v in kv[1]['colores'].values()))[:10]
tlabels = [k for k, v in top]
tprog = [sum(v['prog'] for v in kv['colores'].values()) for _, kv in top]
tped = [sum(v['ped'] for v in kv['colores'].values()) for _, kv in top]

fig, ax = plt.subplots(figsize=(7.5, 3.6), dpi=150)
x = range(len(tlabels))
b1 = ax.bar([i-0.15 for i in x], tprog, width=0.3, label='Programado', color='#2E75B6')
b2 = ax.bar([i+0.15 for i in x], tped, width=0.3, label='Pedido', color='#A9D18E')
ax.set_xticks(list(x)); ax.set_xticklabels([t[:8] for t in tlabels], fontsize=8, rotation=30, ha='right')
ax.set_ylabel('Unidades', fontsize=9)
ax.set_title('Top 10 Estilos: Programado vs Pedido', fontsize=10, fontweight='bold')
ax.legend(fontsize=8)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(TMP_CHART2, bbox_inches='tight')
plt.close()

# ------- PDF -------
doc = SimpleDocTemplate(OUT_PDF, pagesize=A4,
                        leftMargin=15*mm, rightMargin=15*mm,
                        topMargin=15*mm, bottomMargin=15*mm)

styles = getSampleStyleSheet()
titulo = ParagraphStyle('titulo', parent=styles['Title'], fontName='Helvetica-Bold',
                        fontSize=18, alignment=TA_CENTER, textColor=colors.HexColor('#1F4E78'))
subtitulo = ParagraphStyle('subtitulo', parent=styles['Normal'], alignment=TA_CENTER,
                           fontSize=10, textColor=colors.HexColor('#555555'))
h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontName='Helvetica-Bold',
                    fontSize=13, textColor=colors.HexColor('#1F4E78'), spaceBefore=8, spaceAfter=4)
body = styles['BodyText']

story = []
story.append(Paragraph('REPORTE SMS - SUMMER 27', titulo))
story.append(Spacer(1, 2))
story.append(Paragraph('Resumen de producción - Campaña SMS-SUMMER27', subtitulo))
story.append(Spacer(1, 8))
story.append(HRFlowable(width='100%', thickness=1, color=colors.HexColor('#1F4E78')))
story.append(Spacer(1, 10))

# --- Resumen en tabla ---
story.append(Paragraph('Resumen General', h2))
res_data = [
    ['Estilos distintos', str(n_estilos)],
    ['Variantes (por color)', str(n_variantes)],
    ['Estilos con múltiples variantes', str(multi)],
    ['Total unidades PEDIDO', str(tot_ped)],
    ['Total unidades PROGRAMADO', str(tot_prog)],
]
rt = Table(res_data, colWidths=[70*mm, 40*mm])
rt.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E75B6')),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ('FONTNAME', (1,0), (1,-1), 'Helvetica-Bold'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#BFBFBF')),
    ('BACKGROUND', (0,1), (0,-1), colors.HexColor('#F2F7FB')),
    ('ROWBACKGROUNDS', (0,0), (-1,-1), [colors.white, colors.HexColor('#EAF1F8')]),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0), (-1,-1), 4),
    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
]))
story.append(rt)
story.append(Spacer(1, 8))

story.append(Paragraph('Descripción del estado de producción', h2))
story.append(Paragraph(
    'La campaña SMS-SUMMER27 cuenta con <b>%d estilos</b>, que se desglosan en <b>%d variantes</b> '
    'distintas por color de cuerpo. El <b>%.0f%%</b> de los estilos presentan más de una variante, '
    'por lo que el seguimiento se realiza a nivel de "estilo + color". En total se tienen '
    '<b>%d unidades pedidas</b> contra <b>%d unidades programadas</b>, lo que indica un avance de '
    'programación de <b>%.0f%%</b> respecto al pedido.'
    % (n_estilos, n_variantes, (multi/n_estilos*100) if n_estilos else 0, tot_ped, tot_prog,
       (tot_prog/tot_ped*100) if tot_ped else 0), body))
story.append(Spacer(1, 6))

story.append(Paragraph('Distribución por estado de tela', h2))
story.append(Paragraph(
    'Según el estado de la tela, la mayor cantidad de variantes (71) figura aún como '
    '<b>TELA X ING 3-SET</b> (por ingresar a procesos), seguida de <b>EN CORTE</b> (61) y '
    '<b>EN COSTURA</b> (28); mientras que <b>EN LAVANDERÍA</b> suma 12 y <b>EN ACABADOS</b> solo 3. '
    'Esto refleja que la mayor parte del programa se encuentra en las primeras etapas o pendiente de '
    'ingreso a procesos.', body))
story.append(Spacer(1, 6))
story.append(Image(TMP_CHART1, width=175*mm, height=84*mm))
story.append(Spacer(1, 12))

story.append(Paragraph('Top 10 estilos por volumen programado', h2))
story.append(Paragraph(
    'Los estilos con mayor volumen programado son prendas de <b>bebé</b> (Baby Beach Hat, Baby Bubble '
    'Short, Baby Camilla Tank, Baby Lap Onesie, etc.) y pantalones. El líder es <b>%s</b> (%s) con '
    '<b>%d unidades programadas</b>.'
    % (tlabels[0] if tlabels else '-', top[0][1]['nombre'] if top else '-', tprog[0] if tprog else 0), body))
story.append(Spacer(1, 6))
story.append(Image(TMP_CHART2, width=175*mm, height=84*mm))

doc.build(story)
print('PDF generado:', OUT_PDF)
