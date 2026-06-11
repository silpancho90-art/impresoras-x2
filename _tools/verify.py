# -*- coding: utf-8 -*-
import zipfile, sys, os
import xml.etree.ElementTree as ET
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import (load_shared_strings, get_sheets, sheet_target_by_name,
                     read_rows, col_to_num, split_ref)

OUT = 'DATOS_BASE_IMPRESORAS_ACTUALIZADO.xlsx'
BASE = 'datos base impresoras.xlsx'
M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
zo = zipfile.ZipFile(OUT)
zb = zipfile.ZipFile(BASE)

print('=== 1. Well-formedness de TODAS las partes ===')
bad = 0
for n in zo.namelist():
    if n.endswith('.xml') or n.endswith('.rels'):
        try:
            ET.fromstring(zo.read(n))
        except Exception as e:
            bad += 1; print('  INVALIDO:', n, e)
print('  Partes invalidas:', bad)

print('\n=== 2. Conservacion de partes originales ===')
faltan = set(zb.namelist()) - set(zo.namelist())
print('  Faltantes:', faltan if faltan else 'NINGUNA')
print('  Nuevas:', sorted(set(zo.namelist()) - set(zb.namelist())))
for crit in ['xl/styles.xml', 'xl/sharedStrings.xml', 'xl/theme/theme1.xml',
             'xl/drawings/drawing1.xml', 'xl/persons/person.xml']:
    print('  %s: %s' % (crit, 'IDENTICO' if zb.read(crit) == zo.read(crit) else 'MODIFICADO'))

print('\n=== 3. Hojas ===')
for s in get_sheets(zo):
    print('  -', s['name'], '->', s['target'])

print('\n=== 4. Lint estructural (orden filas/columnas, refs) ===')
problems = 0
for s in get_sheets(zo):
    root = ET.fromstring(zo.read(s['target']))
    data = root.find(M + 'sheetData')
    if data is None:
        continue
    prev_row = 0
    for r in data.findall(M + 'row'):
        rn = int(r.get('r'))
        if rn <= prev_row:
            problems += 1; print('  fila fuera de orden', s['name'], rn)
        prev_row = rn
        prev_col = 0
        for c in r.findall(M + 'c'):
            col, rr = split_ref(c.get('r'))
            if rr != rn:
                problems += 1; print('  ref incoherente', s['name'], c.get('r'))
            cn = col_to_num(col)
            if cn <= prev_col:
                problems += 1; print('  columna fuera de orden', s['name'], c.get('r'))
            prev_col = cn
print('  Problemas:', problems)

print('\n=== 5. Muestras de actualizacion en IMPRESORAS TERMICAS ===')
shared = load_shared_strings(zo)
rows = read_rows(zo, sheet_target_by_name(zo, 'IMPRESORAS TERMICAS'), shared)
hdr = [r for r in rows if r['row'] == 1][0]['cells']
print('  Header J:', hdr.get('J'), '(historico, debe seguir intacto)')
for r in rows:
    c = r['cells']
    if c.get('C') in ('Y308021032', 'Y404229696'):
        print('  COD', c.get('C'), '| CIUDAD', c.get('D'), '| ESTADO', c.get('E'),
              '| SUP', c.get('F'), '| USR', c.get('G'), '| PROY', c.get('H'),
              '| OBS', c.get('I'), '| J(hist)', c.get('J'))
codes = [r['cells'].get('C') for r in rows if r['row'] > 1 and r['cells'].get('C')]
print('  Total CODIGO en base final:', len(codes))

print('\n=== 6. REVISION_MANUAL ===')
rows = read_rows(zo, sheet_target_by_name(zo, 'REVISION_MANUAL'), shared)
print('  Filas (con header):', len(rows))
from collections import Counter
tipos = Counter(r['cells'].get('A') for r in rows[1:])
print('  Tipos:', dict(tipos))
print('  Header:', [rows[0]['cells'].get(chr(65+i)) for i in range(10)])
for r in rows[1:3]:
    print('   ', [r['cells'].get(chr(65+i)) for i in range(10)])

print('\n=== 7. RESUMEN ===')
rows = read_rows(zo, sheet_target_by_name(zo, 'RESUMEN'), shared)
for r in rows:
    print('  ', r['cells'].get('A'), '=', r['cells'].get('B'))
