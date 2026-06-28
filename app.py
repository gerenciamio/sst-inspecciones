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
        import psycopg2, psycopg2.extras
        conn = psycopg2.connect(DATABASE_URL)
        return conn, 'pg'
    else:
        import sqlite3
        conn = sqlite3.connect('inspecciones.db')
        conn.row_factory = sqlite3.Row
        return conn, 'sqlite'

def init_db():
    conn, dbtype = get_db()
    cur = conn.cursor()
    if dbtype == 'pg':
        cur.execute('''CREATE TABLE IF NOT EXISTS visitas (
            id SERIAL PRIMARY KEY, cliente TEXT NOT NULL, fecha TEXT NOT NULL,
            participantes TEXT, asesora TEXT, creado TEXT NOT NULL)''')
        cur.execute('''CREATE TABLE IF NOT EXISTS hallazgos (
            id SERIAL PRIMARY KEY, visita_id INTEGER NOT NULL,
            lugar TEXT, situacion TEXT, recomendacion TEXT,
            foto_antes TEXT, foto_despues TEXT, estado TEXT DEFAULT 'pendiente',
            factor TEXT, prioridad TEXT, responsable TEXT,
            estado_acpm TEXT, fecha_ejecucion TEXT, fecha_seguimiento TEXT)''')
        for col in ['factor TEXT','prioridad TEXT','responsable TEXT','estado_acpm TEXT','fecha_ejecucion TEXT','fecha_seguimiento TEXT']:
            try:
                cur.execute(f"ALTER TABLE hallazgos ADD COLUMN {col}")
            except:
                pass
    else:
        cur.execute('''CREATE TABLE IF NOT EXISTS visitas (
            id INTEGER PRIMARY KEY AUTOINCREMENT, cliente TEXT NOT NULL, fecha TEXT NOT NULL,
            participantes TEXT, asesora TEXT, creado TEXT NOT NULL)''')
        cur.execute('''CREATE TABLE IF NOT EXISTS hallazgos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, visita_id INTEGER NOT NULL,
            lugar TEXT, situacion TEXT, recomendacion TEXT,
            foto_antes TEXT, foto_despues TEXT, estado TEXT DEFAULT 'pendiente',
            factor TEXT, prioridad TEXT, responsable TEXT,
            estado_acpm TEXT, fecha_ejecucion TEXT, fecha_seguimiento TEXT)''')
    conn.commit()
    cur.close()
    conn.close()

init_db()

@app.route('/')
def index():
    with open('index.html', encoding='utf-8') as f: return f.read()

@app.route('/admin')
def admin():
    with open('admin.html', encoding='utf-8') as f: return f.read()

@app.route('/cliente')
def cliente():
    with open('cliente.html', encoding='utf-8') as f: return f.read()

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
        cur.execute('INSERT INTO visitas (cliente,fecha,participantes,asesora,creado) VALUES (%s,%s,%s,%s,%s) RETURNING id',
            (cliente, fecha, data.get('participantes',''), data.get('asesora',''), ahora))
        visita_id = cur.fetchone()[0]
        for h in data.get('hallazgos', []):
            cur.execute('INSERT INTO hallazgos (visita_id,lugar,situacion,recomendacion,foto_antes,estado,factor,prioridad,responsable,estado_acpm,fecha_ejecucion,fecha_seguimiento) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                (visita_id,h.get('lugar',''),h.get('situacion',''),h.get('recomendacion',''),h.get('fotoBefore',''),'pendiente',h.get('factor',''),h.get('prioridad',''),h.get('responsable',''),h.get('estado_acpm',''),h.get('fecha_ejecucion',''),h.get('fecha_seguimiento','')))
    else:
        cur.execute('INSERT INTO visitas (cliente,fecha,participantes,asesora,creado) VALUES (?,?,?,?,?)',
            (cliente, fecha, data.get('participantes',''), data.get('asesora',''), ahora))
        visita_id = cur.lastrowid
        for h in data.get('hallazgos', []):
            cur.execute('INSERT INTO hallazgos (visita_id,lugar,situacion,recomendacion,foto_antes,estado,factor,prioridad,responsable,estado_acpm,fecha_ejecucion,fecha_seguimiento) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                (visita_id,h.get('lugar',''),h.get('situacion',''),h.get('recomendacion',''),h.get('fotoBefore',''),'pendiente',h.get('factor',''),h.get('prioridad',''),h.get('responsable',''),h.get('estado_acpm',''),h.get('fecha_ejecucion',''),h.get('fecha_seguimiento','')))
    conn.commit(); cur.close(); conn.close()
    return jsonify({'ok': True, 'visita_id': visita_id})

@app.route('/api/visita/<int:visita_id>/hallazgo', methods=['POST'])
def agregar_hallazgo(visita_id):
    data = request.get_json()
    conn, dbtype = get_db()
    cur = conn.cursor()
    ph = '%s' if dbtype == 'pg' else '?'
    cur.execute(f'INSERT INTO hallazgos (visita_id,lugar,situacion,recomendacion,foto_antes,estado,factor,prioridad,responsable,estado_acpm,fecha_ejecucion,fecha_seguimiento) VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})',
        (visita_id,data.get('lugar',''),data.get('situacion',''),data.get('recomendacion',''),data.get('fotoBefore',''),'pendiente',data.get('factor',''),data.get('prioridad',''),data.get('responsable',''),data.get('estado_acpm',''),data.get('fecha_ejecucion',''),data.get('fecha_seguimiento','')))
    conn.commit(); cur.close(); conn.close()
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
    query = 'SELECT v.*, COUNT(h.id) as total_hallazgos, SUM(CASE WHEN h.estado=\'cerrado\' THEN 1 ELSE 0 END) as cerrados FROM visitas v LEFT JOIN hallazgos h ON h.visita_id=v.id WHERE 1=1'
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
    cur.close(); conn.close()
    return jsonify(rows)
