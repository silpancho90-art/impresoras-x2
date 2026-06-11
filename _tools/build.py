# -*- coding: utf-8 -*-
"""
Genera DATOS_BASE_IMPRESORAS_ACTUALIZADO.xlsx:
- Actualiza la hoja IMPRESORAS TERMICAS del archivo base con los datos mas recientes
  de DATOS RECIBIDOS, comparando por CODIGO (con normalizacion).
- Conserva diseno, formatos, colores, formulas, dibujos y nombres de hoja originales.
- Crea hoja REVISION_MANUAL (no coincidencias) y hoja RESUMEN (indicadores).
- Usa inlineStr para no tocar sharedStrings; no toca styles, theme, drawings ni person.
"""
import zipfile, re, sys, os
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import (find_file, load_shared_strings, sheet_target_by_name,
                     read_rows, normalize)

BASE = find_file('datos base impresoras')
UPD = find_file('datos recibidos')
OUT = 'DATOS_BASE_IMPRESORAS_ACTUALIZADO.xlsx'

BASE_COLS = {'codigo_completo': 'A', 'modelo': 'B', 'codigo': 'C', 'ciudad': 'D',
             'estado': 'E', 'supervisor': 'F', 'usuario': 'G', 'proyecto': 'H',
             'observacion': 'I'}
UPD_COLS = {'codigo_completo': 'B', 'modelo': 'C', 'codigo': 'D', 'ciudad': 'E',
            'estado': 'F', 'supervisor': 'G', 'usuario': 'H', 'proyecto': 'I',
            'observacion': 'J'}
FIELDS = ['ciudad', 'estado', 'supervisor', 'usuario', 'proyecto', 'observacion']
OBS_MARKERS = {'COINCIDE', 'COINDICE'}


def norm_code(v):
    if v is None:
        return ''
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return repr(f)
    except ValueError:
        return normalize(s)


def esc(v):
    if v is None:
        v = ''
    return str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def istr(value, ref=None, style=None):
    a = []
    if ref:
        a.append('r="%s"' % ref)
    if style is not None:
        a.append('s="%s"' % style)
    a.append('t="inlineStr"')
    return '<c %s><is><t xml:space="preserve">%s</t></is></c>' % (' '.join(a), esc(value))


def ncell(value, ref=None, style=None):
    a = []
    if ref:
        a.append('r="%s"' % ref)
    if style is not None:
        a.append('s="%s"' % style)
    return '<c %s><v>%s</v></c>' % (' '.join(a), value)


def num_to_col(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ===== 1. Leer datos =====
zb = zipfile.ZipFile(BASE)
sb = load_shared_strings(zb)
base_rows = read_rows(zb, sheet_target_by_name(zb, 'IMPRESORAS TERMICAS'), sb)

zu = zipfile.ZipFile(UPD)
su = load_shared_strings(zu)
upd_rows = read_rows(zu, sheet_target_by_name(zu, 'Códigos recepcionados'), su)

base = {}
base_order = []
base_dups = []
for r in base_rows:
    if r['row'] <= 1:
        continue
    c = r['cells']
    code = c.get('C')
    if code is None or str(code).strip() == '':
        continue
    key = norm_code(code)
    rec = {f: c.get(BASE_COLS[f]) for f in BASE_COLS}
    rec['_row'] = r['row']
    rec['_codedisp'] = str(code).strip()
    if key in base:
        base_dups.append((key, base[key]['_row'], r['row']))
    else:
        base_order.append(key)
        base[key] = rec

upd = {}
upd_order = []
upd_dups = []
for r in upd_rows:
    if r['row'] <= 2:
        continue
    c = r['cells']
    code = c.get('D')
    if code is None or str(code).strip() == '':
        continue
    key = norm_code(code)
    rec = {f: c.get(UPD_COLS[f]) for f in UPD_COLS}
    rec['_codedisp'] = str(code).strip()
    if key in upd:
        upd_dups.append((key, r['row']))
    else:
        upd_order.append(key)
        upd[key] = rec

m_keys = set(base)
u_keys = set(upd)
coincidentes = [k for k in base_order if k in u_keys]
nuevos = [k for k in upd_order if k not in m_keys]                 # NUEVO REGISTRO
no_encontrados = [k for k in base_order if k not in u_keys]        # NO ENCONTRADO EN ACTUALIZACION


# ===== 2. Calcular actualizaciones =====
change_counts = {f: 0 for f in FIELDS}
row_updates = {}       # base rownum -> {field: nuevo_valor}
records_updated = 0
for key in coincidentes:
    b = base[key]
    u = upd[key]
    ups = {}
    for f in FIELDS:
        uv = u.get(f)
        if uv is None or str(uv).strip() == '':
            continue  # no sobrescribir con vacio
        if f == 'observacion' and normalize(uv) in OBS_MARKERS:
            continue  # marcador de validacion, no sustantivo
        if normalize(b.get(f)) != normalize(uv):
            ups[f] = uv
            change_counts[f] += 1
    if ups:
        row_updates[b['_row']] = ups
        records_updated += 1


# ===== 3. Editar hoja base (sheet1.xml) =====
sheet1 = zb.read('xl/worksheets/sheet1.xml').decode('utf-8')


def get_cell_style(row_xml, ref):
    m = re.search(r'<c r="%s"(?:\s+s="(\d+)")?' % re.escape(ref), row_xml)
    return m.group(1) if (m and m.group(1)) else None


def replace_cell(row_xml, ref, value, default_style):
    style = get_cell_style(row_xml, ref) or default_style
    newc = istr(value, ref, style)
    pat_full = re.compile(r'<c r="%s"[^>]*>.*?</c>' % re.escape(ref), re.S)
    if pat_full.search(row_xml):
        return pat_full.sub(lambda _: newc, row_xml, count=1)
    pat_self = re.compile(r'<c r="%s"[^>]*/>' % re.escape(ref))
    if pat_self.search(row_xml):
        return pat_self.sub(lambda _: newc, row_xml, count=1)
    return row_xml


def process_row(match):
    row_xml = match.group(0)
    rn = int(re.match(r'<row r="(\d+)"', row_xml).group(1))
    ups = row_updates.get(rn)
    if not ups:
        return row_xml
    for f, val in ups.items():
        row_xml = replace_cell(row_xml, '%s%d' % (BASE_COLS[f], rn), val, '4')
    return row_xml


sheet1 = re.sub(r'<row r="\d+"[^>]*>.*?</row>', process_row, sheet1, flags=re.S)


# ===== 4. Hojas nuevas =====
def build_sheet(headers, data_rows, colwidth=22, freeze=True):
    ncols = len(headers)
    last_col = num_to_col(ncols)
    nrows = len(data_rows) + 1
    cols_xml = ''.join('<col customWidth="1" min="%d" max="%d" width="%d"/>' % (i + 1, i + 1, colwidth)
                       for i in range(ncols))
    rows_xml = ''
    hcells = ''.join(istr(h, '%s1' % num_to_col(i + 1), None) for i, h in enumerate(headers))
    rows_xml += '<row r="1">%s</row>' % hcells
    for ri, rowvals in enumerate(data_rows, start=2):
        cells = ''
        for ci, val in enumerate(rowvals):
            ref = '%s%d' % (num_to_col(ci + 1), ri)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells += ncell(val, ref, None)
            else:
                cells += istr('' if val is None else val, ref, None)
        rows_xml += '<row r="%d">%s</row>' % (ri, cells)
    pane = ('<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
            'activePane="bottomLeft" state="frozen"/></sheetView>') if freeze else \
           '<sheetView workbookViewId="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="A1:%s%d"/>'
            '<sheetViews>%s</sheetViews>'
            '<sheetFormatPr defaultRowHeight="15.75"/>'
            '<cols>%s</cols>'
            '<sheetData>%s</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '</worksheet>') % (last_col, nrows, pane, cols_xml, rows_xml)


# REVISION_MANUAL
REV_HEADERS = ['Tipo', 'CODIGO', 'CODIGO COMPLETO', 'MODELO', 'CIUDAD', 'ESTADO',
               'SUPERVISOR', 'USUARIO', 'PROYECTO', 'OBSERVACION']
rev_data = []
for key in nuevos:
    u = upd[key]
    rev_data.append(['NUEVO REGISTRO', u['_codedisp'], u.get('codigo_completo'), u.get('modelo'),
                     u.get('ciudad'), u.get('estado'), u.get('supervisor'), u.get('usuario'),
                     u.get('proyecto'),
                     (u.get('observacion') if normalize(u.get('observacion')) not in OBS_MARKERS else '')])
for key in no_encontrados:
    b = base[key]
    rev_data.append(['NO ENCONTRADO EN ACTUALIZACION', b['_codedisp'], b.get('codigo_completo'),
                     b.get('modelo'), b.get('ciudad'), b.get('estado'), b.get('supervisor'),
                     b.get('usuario'), b.get('proyecto'), b.get('observacion')])
sheet2_xml = build_sheet(REV_HEADERS, rev_data)

# RESUMEN
res_rows = [
    ['Concepto', 'Valor'],
    ['Total de códigos en DATOS BASE IMPRESORAS', len(base_order)],
    ['Total de códigos en DATOS RECIBIDOS', len(upd_order)],
    ['Códigos coincidentes', len(coincidentes)],
    ['Códigos NO coincidentes (total)', len(nuevos) + len(no_encontrados)],
    ['  - Nuevos registros (en RECIBIDOS, no en BASE)', len(nuevos)],
    ['  - No encontrados en actualización (en BASE, no en RECIBIDOS)', len(no_encontrados)],
    ['Registros actualizados (con al menos un cambio)', records_updated],
    ['Campos actualizados - Ciudad', change_counts['ciudad']],
    ['Campos actualizados - Estado', change_counts['estado']],
    ['Campos actualizados - Supervisor', change_counts['supervisor']],
    ['Campos actualizados - Usuario', change_counts['usuario']],
    ['Campos actualizados - Proyecto', change_counts['proyecto']],
    ['Campos actualizados - Observación', change_counts['observacion']],
    ['Códigos duplicados en BASE', len(base_dups)],
    ['Códigos duplicados en RECIBIDOS', len(upd_dups)],
]
# build RESUMEN as simple 2-col sheet (header + rows)
sheet3_xml = build_sheet(res_rows[0], res_rows[1:], colwidth=48, freeze=False)


# ===== 5. workbook / rels / content types =====
workbook = zb.read('xl/workbook.xml').decode('utf-8')
workbook = workbook.replace(
    '<sheet state="visible" name="IMPRESORAS TERMICAS" sheetId="1" r:id="rId5"/></sheets>',
    '<sheet state="visible" name="IMPRESORAS TERMICAS" sheetId="1" r:id="rId5"/>'
    '<sheet state="visible" name="REVISION_MANUAL" sheetId="2" r:id="rId6"/>'
    '<sheet state="visible" name="RESUMEN" sheetId="3" r:id="rId7"/></sheets>')

rels = zb.read('xl/_rels/workbook.xml.rels').decode('utf-8')
rels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
    '</Relationships>')

ctypes = zb.read('[Content_Types].xml').decode('utf-8')
ctypes = ctypes.replace(
    '</Types>',
    '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet2.xml"/>'
    '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet3.xml"/>'
    '</Types>')

overrides = {
    'xl/worksheets/sheet1.xml': sheet1,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    '[Content_Types].xml': ctypes,
}
additions = {
    'xl/worksheets/sheet2.xml': sheet2_xml,
    'xl/worksheets/sheet3.xml': sheet3_xml,
}

with zipfile.ZipFile(BASE) as zin, zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        name = item.filename
        if name in overrides:
            zout.writestr(item, overrides[name].encode('utf-8'))
        else:
            zout.writestr(item, zin.read(name))
    for name, data in additions.items():
        zout.writestr(name, data.encode('utf-8'))

# ===== 6. Resumen en consola =====
print('=' * 60)
print('ARCHIVO GENERADO:', OUT)
print('=' * 60)
print('Total base:', len(base_order), '| Total recibidos:', len(upd_order))
print('Coincidentes:', len(coincidentes))
print('No coincidentes:', len(nuevos) + len(no_encontrados),
      '(nuevos:', len(nuevos), '| no encontrados:', len(no_encontrados), ')')
print('Registros actualizados:', records_updated)
print('Cambios por campo:', change_counts)
print('Duplicados base:', base_dups, '| recibidos:', upd_dups)
print('NUEVO REGISTRO:', [upd[k]['_codedisp'] for k in nuevos])
print('NO ENCONTRADO:', [base[k]['_codedisp'] for k in no_encontrados])
