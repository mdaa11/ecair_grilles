from flask import Flask, request, send_file
from flask_cors import CORS
import io, os, shutil, zipfile, tempfile
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
OKKO_COL  = {'A4':2,'A5':2,'A6':2,'A7':3,'A8':3,'M7':3,'M8':3,'M3_ACE':2,'M3_ECE':2,'M3_CVAD':2}
COMMENT_COL={'A4':6,'A5':6,'A6':6,'A7':7,'A8':7,'M7':7,'M8':7,'M3_ACE':5,'M3_ECE':5,'M3_CVAD':5}
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
    s=''
    while n>0:
        n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def cell_addr(row,col): return col_letter(col)+str(row)

def get_sheet_path(zip_path, sheet_name):
    with zipfile.ZipFile(zip_path,'r') as z:
        wb_xml=z.read('xl/workbook.xml'); wb_rels=z.read('xl/_rels/workbook.xml.rels')
    root=etree.fromstring(wb_xml); rId=None
    for s in root.findall(f'.//{{{NS}}}sheet'):
        if s.get('name')==sheet_name:
            rId=s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'); break
    rels_root=etree.fromstring(wb_rels)
    for rel in rels_root:
        if rel.get('Id')==rId: return 'xl/'+rel.get('Target')

def inject_into_xlsx(src_path, sheet_values):
    tmp=tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False); tmp_path=tmp.name; tmp.close()
    shutil.copy2(src_path, tmp_path)

    for sheet_name, values in sheet_values.items():
        sheet_path=get_sheet_path(tmp_path, sheet_name)
        with zipfile.ZipFile(tmp_path,'r') as z:
            sheet_xml=z.read(sheet_path)
            has_ss='xl/sharedStrings.xml' in z.namelist()
            ss_xml=z.read('xl/sharedStrings.xml') if has_ss else None

        sheet_root=etree.fromstring(sheet_xml)
        shared_strings=[]
        if ss_xml:
            ss_root=etree.fromstring(ss_xml)
            for si in ss_root.findall(f'{{{NS}}}si'):
                t=si.find(f'{{{NS}}}t')
                if t is not None: shared_strings.append(t.text or '')
                else: shared_strings.append(''.join(p.text or '' for p in si.findall(f'.//{{{NS}}}t')))
        else:
            ss_root=etree.Element(f'{{{NS}}}sst')

        def get_or_add(s):
            if s in shared_strings: return shared_strings.index(s)
            idx=len(shared_strings); shared_strings.append(s)
            si=etree.SubElement(ss_root,f'{{{NS}}}si')
            t_el=etree.SubElement(si,f'{{{NS}}}t'); t_el.text=s
            t_el.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            return idx

        sd=sheet_root.find(f'{{{NS}}}sheetData')
        row_map={int(r.get('r')):r for r in sd.findall(f'{{{NS}}}row')}
        cell_map={}
        for r_el in sd.findall(f'{{{NS}}}row'):
            for c_el in r_el.findall(f'{{{NS}}}c'): cell_map[c_el.get('r')]=c_el

        for (row,col),val in values.items():
            addr=cell_addr(row,col); str_val=str(val); idx=get_or_add(str_val)
            if addr in cell_map:
                c_el=cell_map[addr]  # preserve existing s= style
            else:
                # new cell: inherit style from same row
                row_style=None
                for ca,ce in cell_map.items():
                    if ca.endswith(str(row)) and ce.get('s'): row_style=ce.get('s'); break
                if row not in row_map:
                    r_el=etree.SubElement(sd,f'{{{NS}}}row'); r_el.set('r',str(row)); row_map[row]=r_el
                c_el=etree.SubElement(row_map[row],f'{{{NS}}}c'); c_el.set('r',addr)
                if row_style: c_el.set('s',row_style)
                cell_map[addr]=c_el
            f_el=c_el.find(f'{{{NS}}}f')
            if f_el is not None: c_el.remove(f_el)
            c_el.set('t','s')
            v_el=c_el.find(f'{{{NS}}}v')
            if v_el is None: v_el=etree.SubElement(c_el,f'{{{NS}}}v')
            v_el.text=str(idx)

        ss_root.set('count',str(len(shared_strings))); ss_root.set('uniqueCount',str(len(shared_strings)))
        sheet_xml_out=etree.tostring(sheet_root,xml_declaration=True,encoding='UTF-8',standalone=True)
        ss_xml_out=etree.tostring(ss_root,xml_declaration=True,encoding='UTF-8',standalone=True)

        out_path=tmp_path+'.out'
        with zipfile.ZipFile(tmp_path,'r') as zin:
            with zipfile.ZipFile(out_path,'w',zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename==sheet_path: zout.writestr(item,sheet_xml_out)
                    elif item.filename=='xl/sharedStrings.xml': zout.writestr(item,ss_xml_out)
                    else: zout.writestr(item,zin.read(item.filename))
        os.replace(out_path,tmp_path)

    with open(tmp_path,'rb') as f: data=f.read()
    os.unlink(tmp_path)
    return data

with open(os.path.join(TEMPLATES_DIR,'index.html'),'r',encoding='utf-8') as f:
    HTML_CONTENT=f.read()

@app.route('/')
def index(): return HTML_CONTENT

@app.route('/generate',methods=['POST'])
def generate():
    data=request.json
    grille_code=data.get('code'); nom=data.get('nom','Prénom Nom')
    date_str=data.get('date',str(date.today())); boa=data.get('boa','')
    answers=data.get('answers',{}); comments=data.get('comments',{})
    codes=COMBINED.get(grille_code,[grille_code])
    template_path=os.path.join(TEMPLATES_DIR,TEMPLATE_FILES[codes[0]])
    sheet_values={}
    for code in codes:
        sname=SHEET_NAMES[code]; vals={}
        vals[(1,1)]='Dossier : '+nom; vals[(2,1)]='BOA : '+boa
        ans=answers.get(code,{}); cmt=comments.get(code,{})
        for i,row in enumerate(CRITERIA_ROWS[code]):
            vals[(row,OKKO_COL[code])]=ans.get(str(i),'OK')
            comment=cmt.get(str(i),'')
            if comment: vals[(row,COMMENT_COL[code])]=comment
        sheet_values[sname]=vals
    xlsx_bytes=inject_into_xlsx(template_path,sheet_values)
    display={'A7A8':'A7-A8','M7M8':'M7-M8','M3':'M3'}.get(grille_code,grille_code)
    filename=f'{date_str} - GQualité {display} - {nom}.xlsx'
    return send_file(io.BytesIO(xlsx_bytes),as_attachment=True,download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__':
    app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)))
