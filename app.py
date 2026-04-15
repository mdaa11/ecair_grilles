from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook
import io, os, copy, json
from datetime import date

app = Flask(__name__)
CORS(app)

TEMPLATES_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_FILES = {
    'A4': 'GQualité_A4_-_Prénom_Nom.xlsx',
    'A5': 'GQualité_A5_-_Prénom_Nom.xlsx',
    'A6': 'GQualité_A6_-_Prénom_Nom.xlsx',
    'A7': 'GQualité_A7_et_A8_-_Prénom_Nom.xlsx',
    'A8': 'GQualité_A7_et_A8_-_Prénom_Nom.xlsx',
}

SHEET_NAMES = {
    'A4': 'A4',
    'A5': 'A5  SPEKTY',
    'A6': 'A6',
    'A7': 'A7',
    'A8': 'A8',
}

OKKO_COL = {
    'A4': 2, 'A5': 2, 'A6': 2,
    'A7': 3, 'A8': 3,
}

COMMENT_COL = {
    'A4': 6, 'A5': 6, 'A6': 6,
    'A7': 7, 'A8': 7,
}

CRITERIA_ROWS = {
    'A4': [4,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43],
    'A5': [4,5,7,8,10,11,12,13,15,16,18,19,20,21,23,24,25,26,28,29,30,31,33,34,35,36,38,39,40,41,43,44,45,47,48,49,51,52,54,56,57,58,59,61,62,64,65,67,68,69,70,72,73,75,76,78,79],
    'A6': [4,5,6,7,8,9,10,11,12,13,14],
    'A7': [4,5,6,7,8,9,10,11,12,13,14,15,16,17,18],
    'A8': [5,6,7,8,9,10,11,12,14,15,16,17,18,19,20,21,22],
}

with open(os.path.join(TEMPLATES_DIR, 'index.html'), 'r', encoding='utf-8') as f:
    HTML_CONTENT = f.read()

@app.route('/')
def index():
    return HTML_CONTENT

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    grille_code = data.get('code')
    nom = data.get('nom', 'Prénom Nom')
    date_str = data.get('date', str(date.today()))
    boa = data.get('boa', '')
    answers = data.get('answers', {})
    comments = data.get('comments', {})

    # For A7A8 combined, generate both sheets in one file
    codes = ['A7', 'A8'] if grille_code == 'A7A8' else [grille_code]

    # Load template (use A7 template for both A7 and A8)
    template_file = TEMPLATE_FILES[codes[0]]
    wb = load_workbook(os.path.join(TEMPLATES_DIR, template_file))

    for code in codes:
        ws = wb[SHEET_NAMES[code]]
        okko_col = OKKO_COL[code]
        comment_col = COMMENT_COL[code]
        criteria_rows = CRITERIA_ROWS[code]

        # Fill dossier and BOA
        if code in ('A4', 'A6'):
            ws['A1'] = 'Dossier : ' + nom
            ws['A2'] = 'BOA : ' + boa
        elif code in ('A7', 'A8'):
            ws['A1'] = 'Dossier : ' + nom
            ws['A2'] = 'BOA : ' + boa

        # Fill OK/KO/NA values and add eval formulas where missing
        ans = answers.get(code, {})
        cmt = comments.get(code, {})
        eval_col = okko_col + 3  # E col for A4/A5/A6, F col for A7/A8

        for i, row in enumerate(criteria_rows):
            val = ans.get(str(i), 'OK')
            ws.cell(row=row, column=okko_col).value = val
            comment = cmt.get(str(i), '')
            if comment:
                ws.cell(row=row, column=comment_col).value = comment
            # Add evaluation formula if not already present
            eval_cell = ws.cell(row=row, column=eval_col)
            if eval_cell.value is None:
                okko_letter = chr(64 + okko_col)
                poids_letter = chr(64 + okko_col + 2)
                if code in ('A7', 'A8'):
                    eval_cell.value = f'=IF({okko_letter}{row}="OK",{poids_letter}{row},IF({okko_letter}{row}="KO",0,IF({okko_letter}{row}="NA","NA","")))'
                else:
                    eval_cell.value = f'=IF({okko_letter}{row}="OK",{poids_letter}{row},IF({okko_letter}{row}="KO",0,""))'

    # Save to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    display_code = 'A7-A8' if grille_code == 'A7A8' else grille_code
    filename = f'{date_str} - GQualité {display_code} - {nom}.xlsx'

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
