from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os, io, json, base64
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

app = Flask(__name__)
CORS(app)

DATABASE_URL = os.environ.get('DATABASE_URL', '')

def get_db():
    if DATABASE_URL:
        import psycopg2
        import psycopg2.extras
        conn = psycopg2.connect(DATABASE_URL)
        return conn, 'pg'
    else:
        import sqlite3
        conn = sqlite3.connect('inspecciones.db')
        conn.row_factory = sqlite3.Row
        return conn, 'sqlite'

def init_db():
    conn, dbtype = get_db()
    if dbtype == 'pg':
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS visitas (
                id SERIAL PRIMARY KEY,
                cliente TEXT NOT NULL,
                fecha TEXT NOT NULL,
                participantes TEXT,
                asesora TEXT,
                creado TEXT NOT NULL
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS hallazgos (
                id SERIAL PRIMARY KEY,
                visita_id INTEGER NOT NULL,
                lugar TEXT,
                situacion TEXT,
                recomendacion TEXT,
                foto_antes TEXT,
                foto_despues TEXT,
                estado TEXT DEFAULT 'pendiente'
            )
        ''')
    else:
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS visitas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente TEXT NOT NULL,
                fecha TEXT NOT NULL,
                participantes TEXT,
                asesora TEXT,
                creado TEXT NOT NULL
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS hallazgos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                visita_id INTEGER NOT NULL,
                lugar TEXT,
                situacion TEXT,
                recomendacion TEXT,
                foto_antes TEXT,
                foto_despues TEXT,
                estado TEXT DEFAULT 'pendiente'
            )
        ''')
    conn.commit()
    cur.close()
    conn.close()

init_db()

@app.route('/')
def index():
    with open('index.html', encoding='utf-8') as f:
        return f.read()

@app.route('/admin')
def admin():
    with open('admin.html', encoding='utf-8') as f:
        return f.read()

@app.route('/cliente')
def cliente():
    with open('cliente.html', encoding='utf-8') as f:
        return f.read()

@app.route('/api/visita', methods=['POST'])
def crear_visita():
    data = request.get_json()
    cliente = data.get('cliente','').strip()
    fecha = data.get('fecha','').strip()
    if not cliente or not fecha:
        return jsonify({'error': 'Cliente y fecha requeridos'}), 400
    ahora = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn, dbtype = get_db()
    cur = conn.cursor()
    if dbtype == 'pg':
        cur.execute(
            'INSERT INTO visitas (cliente, fecha, participantes, asesora, creado) VALUES (%s,%s,%s,%s,%s) RETURNING id',
            (cliente, fecha, data.get('participantes',''), data.get('asesora',''), ahora)
        )
        visita_id = cur.fetchone()[0]
        for h in data.get('hallazgos', []):
            cur.execute(
                'INSERT INTO hallazgos (visita_id, lugar, situacion, recomendacion, foto_antes, estado) VALUES (%s,%s,%s,%s,%s,%s)',
                (visita_id, h.get('lugar',''), h.get('situacion',''), h.get('recomendacion',''), h.get('fotoBefore',''), 'pendiente')
            )
    else:
        cur.execute(
            'INSERT INTO visitas (cliente, fecha, participantes, asesora, creado) VALUES (?,?,?,?,?)',
            (cliente, fecha, data.get('participantes',''), data.get('asesora',''), ahora)
        )
        visita_id = cur.lastrowid
        for h in data.get('hallazgos', []):
            cur.execute(
                'INSERT INTO hallazgos (visita_id, lugar, situacion, recomendacion, foto_antes, estado) VALUES (?,?,?,?,?,?)',
                (visita_id, h.get('lugar',''), h.get('situacion',''), h.get('recomendacion',''), h.get('fotoBefore',''), 'pendiente')
            )
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'ok': True, 'visita_id': visita_id})

@app.route('/api/visita/<int:visita_id>/hallazgo', methods=['POST'])
def agregar_hallazgo(visita_id):
    data = request.get_json()
    conn, dbtype = get_db()
    cur = conn.cursor()
    ph = '%s' if dbtype == 'pg' else '?'
    cur.execute(
        f'INSERT INTO hallazgos (visita_id, lugar, situacion, recomendacion, foto_antes, estado) VALUES ({ph},{ph},{ph},{ph},{ph},{ph})',
        (visita_id, data.get('lugar',''), data.get('situacion',''), data.get('recomendacion',''), data.get('fotoBefore',''), 'pendiente')
    )
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'ok': True})

@app.route('/api/visitas')
def listar_visitas():
    cliente_q = request.args.get('cliente', '')
    fecha_q = request.args.get('fecha', '')
    conn, dbtype = get_db()
    ph = '%s' if dbtype == 'pg' else '?'
    if dbtype == 'pg':
        import psycopg2.extras
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    else:
        conn.row_factory = __import__('sqlite3').Row
        cur = conn.cursor()
    query = f'''SELECT v.*,
        COUNT(h.id) as total_hallazgos,
        SUM(CASE WHEN h.estado='cerrado' THEN 1 ELSE 0 END) as cerrados
        FROM visitas v LEFT JOIN hallazgos h ON h.visita_id=v.id
        WHERE 1=1'''
    params = []
    if cliente_q:
        query += f' AND LOWER(v.cliente) LIKE {ph}'
        params.append(f'%{cliente_q.lower()}%')
    if fecha_q:
        query += f' AND v.fecha={ph}'
        params.append(fecha_q)
    query += ' GROUP BY v.id ORDER BY v.creado DESC LIMIT 100'
    cur.execute(query, params)
    rows = [dict(r) for r in cur.fetchall()]
    cur.close()
    conn.close()
    return jsonify(rows)

@app.route('/api/visita/<int:visita_id>')
def detalle_visita(visita_id):
    conn, dbtype = get_db()
    ph = '%s' if dbtype == 'pg' else '?'
    if dbtype == 'pg':
        import psycopg2.extras
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    else:
        conn.row_factory = __import__('sqlite3').Row
        cur = conn.cursor()
    cur.execute(f'SELECT * FROM visitas WHERE id={ph}', (visita_id,))
    v = cur.fetchone()
    if not v:
        cur.close(); conn.close()
        return jsonify({'error': 'No encontrada'}), 404
    cur.execute(f'SELECT * FROM hallazgos WHERE visita_id={ph} ORDER BY id', (visita_id,))
    hs = [dict(h) for h in cur.fetchall()]
    cur.close()
    conn.close()
    return jsonify({'visita': dict(v), 'hallazgos': hs})

@app.route('/api/hallazgo/<int:hallazgo_id>/despues', methods=['POST'])
def subir_despues(hallazgo_id):
    data = request.get_json()
    foto = data.get('foto_despues', '')
    conn, dbtype = get_db()
    ph = '%s' if dbtype == 'pg' else '?'
    cur = conn.cursor()
    cur.execute(f"UPDATE hallazgos SET foto_despues={ph}, estado='cerrado' WHERE id={ph}", (foto, hallazgo_id))
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'ok': True})

def b64_to_xl_image(b64_str, max_w=200, max_h=150):
    try:
        data = base64.b64decode(b64_str.split(',')[1])
        img = PILImage.open(io.BytesIO(data)).convert('RGB')
        img.thumbnail((max_w, max_h), PILImage.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        return XLImage(buf)
    except:
        return None

@app.route('/api/exportar/<int:visita_id>')
def exportar_excel(visita_id):
    conn, dbtype = get_db()
    ph = '%s' if dbtype == 'pg' else '?'
    if dbtype == 'pg':
        import psycopg2.extras
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    else:
        conn.row_factory = __import__('sqlite3').Row
        cur = conn.cursor()
    cur.execute(f'SELECT * FROM visitas WHERE id={ph}', (visita_id,))
    v = dict(cur.fetchone())
    cur.execute(f'SELECT * FROM hallazgos WHERE visita_id={ph} ORDER BY id', (visita_id,))
    hs = [dict(h) for h in cur.fetchall()]
    cur.close()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Informe de inspección'
    thin = Side(style='thin', color='000000')
    ALL_BORDERS = Border(top=thin, bottom=thin, left=thin, right=thin)

    def apply(cell, val, bold=False, sz=10, h='center', v_align='center', fill=None, border=False):
        cell.value = val
        cell.font = Font(name='Arial', bold=bold, size=sz)
        cell.alignment = Alignment(horizontal=h, vertical=v_align, wrap_text=True)
        if fill:
            cell.fill = PatternFill('solid', fgColor=fill)
        if border:
            cell.border = ALL_BORDERS

    for col, w in [('A',16),('B',10),('C',12),('D',21),('E',38),('F',38),('G',25),('H',12)]:
        ws.column_dimensions[col].width = w
    for i, h in enumerate([14,14,33,33,35], 1):
        ws.row_dimensions[i].height = h

    apply(ws['A1'], 'LOGO', bold=True)
    apply(ws['C1'], 'SEGURIDAD SALUD EN EL TRABAJO', bold=True)
    apply(ws['F1'], f'Fecha: {v["fecha"]}', bold=True, sz=8)
    apply(ws['C2'], 'INFORME Y SEGUIMIENTO A INSPECCIONES', bold=True)
    apply(ws['F2'], 'CODIGO: FT-SST-020', bold=True, sz=8)
    apply(ws['F3'], 'Pag 1 de 2', bold=True, sz=8)
    apply(ws['G3'], 'Versión 1', bold=True)
    apply(ws['A4'], f'Fecha inspección: {v["fecha"]}', bold=True, h='left')
    apply(ws['E4'], f'PARTICIPANTES: {v["participantes"]}', bold=True, h='left')

    gray = 'D3D1C7'
    for col, txt in [('A','LUGAR'),('B','SITUACIÓN ENCONTRADA'),('E','FOTO ANTES'),('F','FOTO DESPUÉS'),('G','RECOMENDACIÓN'),('H','ESTADO')]:
        apply(ws[f'{col}5'], txt, bold=True, fill=gray, border=True)

    for merge in ['A1:B3','C1:E1','C2:E3','F1:G1','F2:G2','F3:G3','A4:B4','E4:H4','B5:D5']:
        ws.merge_cells(merge)

    ROW_H = 140
    for i, h in enumerate(hs):
        r = 6 + i
        ws.row_dimensions[r].height = ROW_H
        apply(ws[f'A{r}'], h.get('lugar',''), border=True, h='left', v_align='top')
        apply(ws[f'B{r}'], h.get('situacion',''), border=True, h='left', v_align='top')
        apply(ws[f'G{r}'], h.get('recomendacion',''), border=True, h='left', v_align='top')
        estado = h.get('estado','pendiente')
        apply(ws[f'H{r}'], 'Cerrado' if estado=='cerrado' else 'Pendiente',
              border=True, fill='C0DD97' if estado=='cerrado' else 'FAC775')
        apply(ws[f'E{r}'], '', border=True)
        apply(ws[f'F{r}'], '', border=True)
        ws.merge_cells(f'B{r}:D{r}')
        if h.get('foto_antes'):
            img = b64_to_xl_image(h['foto_antes'])
            if img:
                img.anchor = f'E{r}'
                ws.add_image(img)
        if h.get('foto_despues'):
            img2 = b64_to_xl_image(h['foto_despues'])
            if img2:
                img2.anchor = f'F{r}'
                ws.add_image(img2)

    obs_row = 6 + len(hs)
    sig_row = obs_row + 1
    ws.row_dimensions[obs_row].height = 80
    ws.row_dimensions[sig_row].height = 50
    apply(ws[f'A{obs_row}'], 'OBSERVACIONES GENERALES', h='left', v_align='top')
    ws.merge_cells(f'A{obs_row}:H{obs_row}')
    apply(ws[f'A{sig_row}'], 'Firma de quien realizó la inspección' + ' '*35 + 'Firma de quien recibió la inspección', sz=11, h='center')
    ws.merge_cells(f'A{sig_row}:H{sig_row}')

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"FT-SST-020_{v['cliente'].replace(' ','_')}_{v['fecha']}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
