# -*- coding: utf-8 -*-
"""Utilidades para leer datos de hojas .xlsx con la libreria estandar."""
import zipfile, re, glob, os
import xml.etree.ElementTree as ET
import unicodedata

M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def find_file(hint):
    if os.path.exists(hint):
        return hint
    for f in glob.glob('*.xlsx'):
        if hint.lower() in f.lower():
            return f
    raise FileNotFoundError(hint)


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def num_to_col(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    return m.group(1), int(m.group(2))


def load_shared_strings(z):
    strings = []
    if 'xl/sharedStrings.xml' not in z.namelist():
        return strings
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    for si in root.findall(M + 'si'):
        texts = [t.text or '' for t in si.iter(M + 't')]
        strings.append(''.join(texts))
    return strings


def get_sheets(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_to_target = {rel.get('Id'): rel.get('Target') for rel in rels}
    sheets = []
    for s in wb.find(M + 'sheets'):
        name = s.get('name')
        rid = s.get(R + 'id')
        target = rid_to_target.get(rid)
        if target and not target.startswith('xl/'):
            target = 'xl/' + target
        sheets.append({'name': name, 'rid': rid, 'target': target,
                       'sheetId': s.get('sheetId')})
    return sheets


def sheet_target_by_name(z, name):
    for s in get_sheets(z):
        if s['name'] == name:
            return s['target']
    return None


def read_rows(z, target, shared):
    """Devuelve lista de dicts: {col_letter: value} por fila no vacia, con numero de fila."""
    root = ET.fromstring(z.read(target))
    data = root.find(M + 'sheetData')
    out = []
    if data is None:
        return out
    for r in data.findall(M + 'row'):
        rownum = int(r.get('r'))
        cells = {}
        for c in r.findall(M + 'c'):
            ref = c.get('r')
            col, _ = split_ref(ref)
            t = c.get('t')
            v = c.find(M + 'v')
            isnode = c.find(M + 'is')
            val = None
            if t == 's' and v is not None:
                val = shared[int(v.text)]
            elif t == 'inlineStr' and isnode is not None:
                val = ''.join(x.text or '' for x in isnode.iter(M + 't'))
            elif v is not None:
                val = v.text
            if val is not None and str(val).strip() != '':
                cells[col] = val
        if cells:
            out.append({'row': rownum, 'cells': cells})
    return out


def normalize(s):
    """Normaliza texto: quita tildes, espacios extra, mayusculas."""
    if s is None:
        return ''
    s = str(s)
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    s = re.sub(r'\s+', ' ', s).strip().upper()
    return s
