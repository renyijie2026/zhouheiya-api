"""周黑鸭数据同步服务 — Flask + SQLite，部署到 Sealos"""
import sqlite3, json, os
from datetime import datetime
from flask import Flask, request, jsonify

app = Flask(__name__)
DB = 'data.db'

def init_db():
    with sqlite3.connect(DB) as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            store TEXT NOT NULL,
            date TEXT NOT NULL,
            items TEXT NOT NULL,
            photo_count INTEGER DEFAULT 0,
            created_at TEXT NOT NULL
        )''')
        conn.commit()

def row_to_dict(r):
    return {
        'id': r[0],
        'store': r[1],
        'date': r[2],
        'items': json.loads(r[3]),
        'photo_count': r[4],
        'created_at': r[5]
    }

@app.route('/api/entries', methods=['GET'])
def get_entries():
    with sqlite3.connect(DB) as conn:
        cur = conn.execute('SELECT * FROM entries ORDER BY created_at DESC LIMIT 500')
        rows = [row_to_dict(r) for r in cur.fetchall()]
    # Deduplicate: latest per store+date
    seen = {}
    for r in rows:
        k = f"{r['store']}|{r['date']}"
        if k not in seen:
            seen[k] = r
    return jsonify(list(seen.values()))

@app.route('/api/entries', methods=['POST'])
def add_entry():
    data = request.json
    if not data or not data.get('store') or not data.get('date'):
        return jsonify({'error': '缺少门店或日期'}), 400
    with sqlite3.connect(DB) as conn:
        conn.execute(
            'INSERT INTO entries (store, date, items, photo_count, created_at) VALUES (?, ?, ?, ?, ?)',
            (data['store'], data['date'], json.dumps(data.get('items', []), ensure_ascii=False),
             data.get('photo_count', 0), datetime.now().isoformat())
        )
        conn.commit()
    return jsonify({'ok': True}), 201

@app.route('/api/entries', methods=['DELETE'])
def clear_entries():
    """清空全部数据（需要确认参数）"""
    if request.args.get('confirm') != 'yes':
        return jsonify({'error': '需要 confirm=yes 参数'}), 400
    with sqlite3.connect(DB) as conn:
        conn.execute('DELETE FROM entries')
        conn.commit()
    return jsonify({'ok': True})

@app.route('/')
def health():
    return 'OK'

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
