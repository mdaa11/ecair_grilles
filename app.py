from flask import Flask, request, send_file
from flask_cors import CORS
import io, os, shutil, zipfile
from lxml import etree
from datetime import date

app = Flask(__name__)
CORS(app)

TEMPLATES_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_FILES = {
    'A4':'GQualité_A4_-_Prénom_Nom.xlsx','A5':'GQualité_A5_-_Prénom_Nom.xlsx',
    'A6':'GQualité_A6_-_Prénom_Nom.xlsx','A7':'GQualité_A7_et_A8_-_Prénom_Nom.xlsx',
    'A8':'GQualité_A7_et_A8_-_Prénom_Nom.xlsx','M7':'GQualité_M7_et_M8_-_Prénom_Nom.xlsx',
    'M8':'GQualité_M7_et_M8_-_Prénom_Nom.xlsx','M3_ACE':'GQualité_M3_-_Prénom_Nom.xlsx',
    'M3_ECE':'GQualité_M3_-_Prénom_Nom.xlsx','M3_CVAD':'GQualité_M3_-_Prénom_Nom.xlsx',
}
SHEET_NAMES = {
    'A4':'A4','A5':'A5  SPEKTY','A6':'A6','A7':'A7','A8':'A8',
    'M7':'M7','M8':'M8','M3_ACE':'M3  PAC-ACE','M3_ECE':'M3  PAC-ECE','M3_CVAD':'M3  CVAD',
}
OKKO_COL   = {'A4':2,'A5':2,'A6':2,'A7':3,'A8':3,'M7':3,'M8':3,'M3_ACE':2,'M3_ECE':2,'M3_CVAD':2}
COMMENT_COL= {'A4':6,'A5':6,'A6':6,'A7':7,'A8':7,'M7':7,'M8':7,'M3_ACE':5,'M3_ECE':5,'M3_CVAD':5}
CRITERIA_ROWS = {
    'A4':[4,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43],
    'A5':[4,5,7,8,10,11,12,13,15,16,18,19,20,21,23,24,25,26,28,29,30,31,33,34,35,36,38,39,40,41,43,44,45,47,48,49,51,52,54,56,57,58,59,61,62,64,65,67,68,69,70,72,73,75,76,78,79],
    'A6':[4,5,6,7,8,9,10,11,12,13,14],
    'A7':[4,5,6,7,8,9,10,11,12,13,14,15,16,17,18],
    'A8':[5,6,7,8,9,10,11,12,14,15,16,17,18,19,20,21,22],
    'M7':[4,5,6,7,8,9,10,11,12,13,14,15,16,17,18],
    'M8':[5,6,7,8,9,10,11,12,14,15,16,17,18,19,20,21,22],
    'M3_ACE':[6,7,8,9,11,12,13,14,15,17,18,19,20,21,23,24,25,26,27,28,29,31],
    'M3_ECE':[5,6,7,8,9,10,11,12,13,14,15],
    'M3_CVAD':[5,6,7,8,9,10,11,12,13,14,15,16],
}
COMBINED = {'A7A8':['A7','A8'],'M7M8':['M7','M8'],'M3':['M3_ACE','M3_ECE','M3_CVAD']}

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def inject(src, sheet_name, values):
    with zipfile.ZipFile(src, 'r') as z:
        all_files = {name: z.read(name) for name in z.namelist()}

    # Find sheet path
    root = etree.fromstring(all_files['xl/workbook.xml'])
    rels  = etree.fromstring(all_files['xl/_rels/workbook.xml.rels'])
    rId = None
    for s in root.findall(f'.//{{{NS}}}sheet'):
        if s.get('name') == sheet_name:
            rId = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    sheet_path = None
    for rel in rels:
        if rel.get('Id') == rId:
            sheet_path = 'xl/' + rel.get('Target')

    sheet_root = etree.fromstring(all_files[sheet_path])
    ss_root    = etree.fromstring(all_files['xl/sharedStrings.xml'])

    shared = []
    for si in ss_root.findall(f'{{{NS}}}si'):
        t = si.find(f'{{{NS}}}t')
        shared.append(t.text or '' if t is not None else ''.join(p.text or '' for p in si.findall(f'.//{{{NS}}}t')))

    def add_string(s):
        if s in shared:
            return shared.index(s)
        idx = len(shared); shared.append(s)
        si = etree.SubElement(ss_root, f'{{{NS}}}si')
        t  = etree.SubElement(si, f'{{{NS}}}t')
        t.text = s
        return idx

    sd    = sheet_root.find(f'{{{NS}}}sheetData')
    rows  = {int(r.get('r')): r for r in sd.findall(f'{{{NS}}}row')}
    cells = {}
    for r_el in sd.findall(f'{{{NS}}}row'):
        for c_el in r_el.findall(f'{{{NS}}}c'):
            cells[c_el.get('r')] = c_el

    for (row, col), val in values.items():
        addr = col_letter(col) + str(row)
        idx  = add_string(str(val))

        if addr in cells:
            c = cells[addr]
        else:
            style = next((ce.get('s') for a, ce in cells.items() if a[1:] == str(row) and ce.get('s')), None)
            if row not in rows:
                r_el = etree.SubElement(sd, f'{{{NS}}}row')
                r_el.set('r', str(row)); rows[row] = r_el
            c = etree.SubElement(rows[row], f'{{{NS}}}c')
            c.set('r', addr)
            if style: c.set('s', style)
            cells[addr] = c

        for f in c.findall(f'{{{NS}}}f'): c.remove(f)
        c.set('t', 's')
        v = c.find(f'{{{NS}}}v')
        if v is None: v = etree.SubElement(c, f'{{{NS}}}v')
        v.text = str(idx)

    ss_root.set('count', str(len(shared)))
    ss_root.set('uniqueCount', str(len(shared)))
    all_files[sheet_path]             = etree.tostring(sheet_root, xml_declaration=True, encoding='UTF-8', standalone=True)
    all_files['xl/sharedStrings.xml'] = etree.tostring(ss_root,    xml_declaration=True, encoding='UTF-8', standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf

with open(os.path.join(TEMPLATES_DIR, 'index.html'), 'r', encoding='utf-8') as f:
    HTML_CONTENT = f.read()

@app.route('/')
def index():
    return HTML_CONTENT

@app.route('/generate', methods=['POST'])
def generate():
    data        = request.json
    grille_code = data.get('code')
    nom         = data.get('nom', 'Prénom Nom')
    date_str    = data.get('date', str(date.today()))
    boa         = data.get('boa', '')
    answers     = data.get('answers', {})
    comments    = data.get('comments', {})
    codes       = COMBINED.get(grille_code, [grille_code])
    src_path    = os.path.join(TEMPLATES_DIR, TEMPLATE_FILES[codes[0]])

    # For combined grilles (A7A8, M7M8, M3) inject all sheets into same file
    # Start from the source file and inject sheet by sheet
    with zipfile.ZipFile(src_path, 'r') as z:
        all_files = {name: z.read(name) for name in z.namelist()}

    for code in codes:
        sheet_name = SHEET_NAMES[code]
        values = {(1,1): 'Dossier : ' + nom, (2,1): 'BOA : ' + boa}
        ans = answers.get(code, {})
        cmt = comments.get(code, {})
        for i, row in enumerate(CRITERIA_ROWS[code]):
            values[(row, OKKO_COL[code])] = ans.get(str(i), 'OK')
            if cmt.get(str(i)):
                values[(row, COMMENT_COL[code])] = cmt[str(i)]

        # Find sheet path in all_files
        root = etree.fromstring(all_files['xl/workbook.xml'])
        rels = etree.fromstring(all_files['xl/_rels/workbook.xml.rels'])
        rId = None
        for s in root.findall(f'.//{{{NS}}}sheet'):
            if s.get('name') == sheet_name:
                rId = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        sheet_path = None
        for rel in rels:
            if rel.get('Id') == rId:
                sheet_path = 'xl/' + rel.get('Target')

        sheet_root = etree.fromstring(all_files[sheet_path])
        ss_root    = etree.fromstring(all_files['xl/sharedStrings.xml'])

        shared = []
        for si in ss_root.findall(f'{{{NS}}}si'):
            t = si.find(f'{{{NS}}}t')
            shared.append(t.text or '' if t is not None else ''.join(p.text or '' for p in si.findall(f'.//{{{NS}}}t')))

        def add_string(s):
            if s in shared: return shared.index(s)
            idx = len(shared); shared.append(s)
            si = etree.SubElement(ss_root, f'{{{NS}}}si')
            t  = etree.SubElement(si, f'{{{NS}}}t'); t.text = s
            return idx

        sd    = sheet_root.find(f'{{{NS}}}sheetData')
        rows  = {int(r.get('r')): r for r in sd.findall(f'{{{NS}}}row')}
        cells = {}
        for r_el in sd.findall(f'{{{NS}}}row'):
            for c_el in r_el.findall(f'{{{NS}}}c'):
                cells[c_el.get('r')] = c_el

        for (row, col), val in values.items():
            addr = col_letter(col) + str(row)
            idx  = add_string(str(val))
            if addr in cells:
                c = cells[addr]
            else:
                style = next((ce.get('s') for a, ce in cells.items() if a[1:] == str(row) and ce.get('s')), None)
                if row not in rows:
                    r_el = etree.SubElement(sd, f'{{{NS}}}row'); r_el.set('r', str(row)); rows[row] = r_el
                c = etree.SubElement(rows[row], f'{{{NS}}}c'); c.set('r', addr)
                if style: c.set('s', style)
                cells[addr] = c
            for f in c.findall(f'{{{NS}}}f'): c.remove(f)
            c.set('t', 's')
            v = c.find(f'{{{NS}}}v')
            if v is None: v = etree.SubElement(c, f'{{{NS}}}v')
            v.text = str(idx)

        ss_root.set('count', str(len(shared))); ss_root.set('uniqueCount', str(len(shared)))
        all_files[sheet_path]             = etree.tostring(sheet_root, xml_declaration=True, encoding='UTF-8', standalone=True)
        all_files['xl/sharedStrings.xml'] = etree.tostring(ss_root,    xml_declaration=True, encoding='UTF-8', standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
    buf.seek(0)

    display  = {'A7A8':'A7-A8','M7M8':'M7-M8','M3':'M3'}.get(grille_code, grille_code)
    filename = f'{date_str} - GQualité {display} - {nom}.xlsx'
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
