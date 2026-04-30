import sqlite3, uuid, os
from datetime import datetime, timezone, timedelta
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

DATABASE = os.environ.get('DB_PATH', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data.db'))
TZ = timezone(timedelta(hours=8))

def get_db():
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA journal_mode=WAL")
    return db

def init_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS entries (
            id TEXT PRIMARY KEY,
            store TEXT NOT NULL,
            date TEXT NOT NULL,
            items TEXT NOT NULL DEFAULT '[]',
            photo_count INTEGER DEFAULT 0,
            created_at TEXT
        )''')

init_db()

@app.route('/api/entries', methods=['GET'])
def list_entries():
    with get_db() as db:
        rows = db.execute('SELECT * FROM entries ORDER BY created_at DESC').fetchall()
    return jsonify([{
        'id': r['id'], 'store': r['store'], 'date': r['date'],
        'items': __import__('json').loads(r['items']),
        'photo_count': r['photo_count'], 'created_at': r['created_at']
    } for r in rows])

@app.route('/api/entries', methods=['POST'])
def create_entry():
    data = request.get_json(force=True)
    if not data.get('store') or not data.get('date'):
        return jsonify({'error': 'store and date are required'}), 400
    eid = str(uuid.uuid4())
    items_json = __import__('json').dumps(data.get('items', []), ensure_ascii=False)
    now = datetime.now(TZ).isoformat()
    with get_db() as db:
        db.execute('INSERT INTO entries (id, store, date, items, photo_count, created_at) VALUES (?,?,?,?,?,?)',
                   [eid, data['store'], data['date'], items_json, data.get('photo_count', 0), now])
    return jsonify({'id': eid, 'created_at': now}), 201

@app.route('/api/entries/<eid>', methods=['DELETE'])
def delete_entry(eid):
    with get_db() as db:
        cur = db.execute('DELETE FROM entries WHERE id=?', [eid])
        if cur.rowcount == 0:
            return jsonify({'error': 'not found'}), 404
    return jsonify({'deleted': eid})

@app.route('/api/entries', methods=['DELETE'])
def delete_all_entries():
    with get_db() as db:
        count = db.execute('SELECT COUNT(*) FROM entries').fetchone()[0]
        db.execute('DELETE FROM entries')
    return jsonify({'deleted': count})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
