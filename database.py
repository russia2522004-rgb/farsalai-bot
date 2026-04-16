import os
import json
from datetime import datetime

# Используем PostgreSQL если есть DATABASE_URL, иначе SQLite
DATABASE_URL = os.getenv('DATABASE_URL')

if DATABASE_URL:
    import psycopg2
    import psycopg2.extras
    def get_conn():
        return psycopg2.connect(DATABASE_URL)
    PLACEHOLDER = '%s'
else:
    import sqlite3
    def get_conn():
        return sqlite3.connect('farsalai.db')
    PLACEHOLDER = '?'


def init_db():
    """Инициализация базы данных"""
    conn = get_conn()
    c = conn.cursor()

    if DATABASE_URL:
        # PostgreSQL
        c.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                model TEXT NOT NULL UNIQUE,
                description TEXT,
                specs TEXT,
                construction TEXT,
                extra_tables TEXT,
                production_time TEXT DEFAULT '25-30 дней',
                packaging TEXT DEFAULT 'экспортная деревянная тара (ящик)',
                base_price REAL,
                currency TEXT DEFAULT 'ЮАНЕЙ',
                photo_path TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS kp_journal (
                id SERIAL PRIMARY KEY,
                kp_number TEXT NOT NULL UNIQUE,
                kp_date TEXT NOT NULL,
                client TEXT,
                equipment_list TEXT,
                total_price REAL,
                currency TEXT,
                payment_terms TEXT,
                manager_id BIGINT,
                manager_name TEXT,
                yandex_word_url TEXT,
                yandex_pdf_url TEXT,
                yandex_word_id TEXT,
                yandex_pdf_id TEXT,
                sheets_row INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
    else:
        # SQLite
        c.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                model TEXT NOT NULL UNIQUE,
                description TEXT,
                specs TEXT,
                construction TEXT,
                extra_tables TEXT,
                production_time TEXT DEFAULT '25-30 дней',
                packaging TEXT DEFAULT 'экспортная деревянная тара (ящик)',
                base_price REAL,
                currency TEXT DEFAULT 'ЮАНЕЙ',
                photo_path TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS kp_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_number TEXT NOT NULL UNIQUE,
                kp_date TEXT NOT NULL,
                client TEXT,
                equipment_list TEXT,
                total_price REAL,
                currency TEXT,
                payment_terms TEXT,
                manager_id INTEGER,
                manager_name TEXT,
                yandex_word_url TEXT,
                yandex_pdf_url TEXT,
                yandex_word_id TEXT,
                yandex_pdf_id TEXT,
                sheets_row INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

    conn.commit()
    conn.close()
    print("База данных инициализирована")


def _row_to_dict(cursor, row):
    """Конвертирует строку в словарь"""
    if DATABASE_URL:
        return dict(row)
    else:
        cols = [d[0] for d in cursor.description]
        return dict(zip(cols, row))


# ─── Оборудование ────────────────────────────────────────────────────────────

def add_equipment(data: dict) -> int:
    conn = get_conn()
    c = conn.cursor()

    if DATABASE_URL:
        c.execute('''
            INSERT INTO equipment
            (name, model, description, specs, construction, extra_tables,
             production_time, packaging, base_price, currency, photo_path)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT (model) DO UPDATE SET
            name=EXCLUDED.name, description=EXCLUDED.description,
            specs=EXCLUDED.specs, construction=EXCLUDED.construction,
            base_price=EXCLUDED.base_price, currency=EXCLUDED.currency
            RETURNING id
        ''', (
            data.get('name'), data.get('model'), data.get('description'),
            data.get('specs'), data.get('construction'), data.get('extra_tables'),
            data.get('production_time', '25-30 дней'),
            data.get('packaging', 'экспортная деревянная тара (ящик)'),
            data.get('base_price'), data.get('currency', 'ЮАНЕЙ'),
            data.get('photo_path'),
        ))
        eq_id = c.fetchone()[0]
    else:
        c.execute('''
            INSERT OR REPLACE INTO equipment
            (name, model, description, specs, construction, extra_tables,
             production_time, packaging, base_price, currency, photo_path)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
        ''', (
            data.get('name'), data.get('model'), data.get('description'),
            data.get('specs'), data.get('construction'), data.get('extra_tables'),
            data.get('production_time', '25-30 дней'),
            data.get('packaging', 'экспортная деревянная тара (ящик)'),
            data.get('base_price'), data.get('currency', 'ЮАНЕЙ'),
            data.get('photo_path'),
        ))
        eq_id = c.lastrowid

    conn.commit()
    conn.close()
    return eq_id


def get_equipment_by_model(model: str) -> dict | None:
    conn = get_conn()
    c = conn.cursor()
    if DATABASE_URL:
        c.execute('SELECT * FROM equipment WHERE model ILIKE %s', (f'%{model}%',))
        row = c.fetchone()
        if row:
            cols = [d[0] for d in c.description]
            result = dict(zip(cols, row))
        else:
            result = None
    else:
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM equipment WHERE model LIKE ?', (f'%{model}%',))
        row = c.fetchone()
        result = dict(row) if row else None
    conn.close()
    return result


def search_equipment(query: str) -> list:
    conn = get_conn()
    c = conn.cursor()
    if DATABASE_URL:
        c.execute('''
            SELECT * FROM equipment
            WHERE name ILIKE %s OR model ILIKE %s
        ''', (f'%{query}%', f'%{query}%'))
        cols = [d[0] for d in c.description]
        rows = [dict(zip(cols, row)) for row in c.fetchall()]
    else:
        import sqlite3 as sq
        conn.row_factory = sq.Row
        c = conn.cursor()
        c.execute('''
            SELECT * FROM equipment
            WHERE name LIKE ? OR model LIKE ?
        ''', (f'%{query}%', f'%{query}%'))
        rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_all_equipment() -> list:
    conn = get_conn()
    c = conn.cursor()
    if DATABASE_URL:
        c.execute('SELECT id, name, model, base_price, currency FROM equipment ORDER BY name')
        cols = [d[0] for d in c.description]
        rows = [dict(zip(cols, row)) for row in c.fetchall()]
    else:
        import sqlite3 as sq
        conn.row_factory = sq.Row
        c = conn.cursor()
        c.execute('SELECT id, name, model, base_price, currency FROM equipment ORDER BY name')
        rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def update_equipment(model: str, data: dict):
    conn = get_conn()
    c = conn.cursor()
    ph = '%s' if DATABASE_URL else '?'
    fields = ', '.join(f'{k} = {ph}' for k in data.keys())
    values = list(data.values()) + [model]
    c.execute(f'UPDATE equipment SET {fields} WHERE model = {ph}', values)
    conn.commit()
    conn.close()


def delete_equipment(model: str):
    conn = get_conn()
    c = conn.cursor()
    ph = '%s' if DATABASE_URL else '?'
    c.execute(f'DELETE FROM equipment WHERE model = {ph}', (model,))
    conn.commit()
    conn.close()


# ─── Журнал КП ───────────────────────────────────────────────────────────────

def generate_kp_number(equipment_models: list) -> str:
    conn = get_conn()
    c = conn.cursor()
    c.execute('SELECT COUNT(*) FROM kp_journal')
    row = c.fetchone()
    count = (row[0] if row else 0) + 1
    conn.close()
    model_part = equipment_models[0].replace('-', '') if equipment_models else 'KP'
    return f"{model_part}-{count:03d}"


def save_kp(data: dict) -> int:
    conn = get_conn()
    c = conn.cursor()
    ph = '%s' if DATABASE_URL else '?'

    if DATABASE_URL:
        c.execute(f'''
            INSERT INTO kp_journal
            (kp_number, kp_date, client, equipment_list, total_price, currency,
             payment_terms, manager_id, manager_name,
             yandex_word_url, yandex_pdf_url, yandex_word_id, yandex_pdf_id, sheets_row)
            VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})
            RETURNING id
        ''', (
            data.get('kp_number'), data.get('kp_date', datetime.now().strftime('%d.%m.%Y')),
            data.get('client'), data.get('equipment_list'),
            data.get('total_price'), data.get('currency', 'ЮАНЕЙ'),
            data.get('payment_terms', '50% предоплата, 50% по факту поставки'),
            data.get('manager_id'), data.get('manager_name'),
            data.get('yandex_word_url'), data.get('yandex_pdf_url'),
            data.get('yandex_word_id'), data.get('yandex_pdf_id'),
            data.get('sheets_row'),
        ))
        kp_id = c.fetchone()[0]
    else:
        c.execute(f'''
            INSERT INTO kp_journal
            (kp_number, kp_date, client, equipment_list, total_price, currency,
             payment_terms, manager_id, manager_name,
             yandex_word_url, yandex_pdf_url, yandex_word_id, yandex_pdf_id, sheets_row)
            VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph})
        ''', (
            data.get('kp_number'), data.get('kp_date', datetime.now().strftime('%d.%m.%Y')),
            data.get('client'), data.get('equipment_list'),
            data.get('total_price'), data.get('currency', 'ЮАНЕЙ'),
            data.get('payment_terms', '50% предоплата, 50% по факту поставки'),
            data.get('manager_id'), data.get('manager_name'),
            data.get('yandex_word_url'), data.get('yandex_pdf_url'),
            data.get('yandex_word_id'), data.get('yandex_pdf_id'),
            data.get('sheets_row'),
        ))
        kp_id = c.lastrowid

    conn.commit()
    conn.close()
    return kp_id


def update_kp(kp_number: str, data: dict):
    data['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_conn()
    c = conn.cursor()
    ph = '%s' if DATABASE_URL else '?'
    fields = ', '.join(f'{k} = {ph}' for k in data.keys())
    values = list(data.values()) + [kp_number]
    c.execute(f'UPDATE kp_journal SET {fields} WHERE kp_number = {ph}', values)
    conn.commit()
    conn.close()


def get_kp_by_number(kp_number: str) -> dict | None:
    conn = get_conn()
    c = conn.cursor()
    ph = '%s' if DATABASE_URL else '?'
    c.execute(f'SELECT * FROM kp_journal WHERE kp_number = {ph}', (kp_number,))
    row = c.fetchone()
    if row:
        cols = [d[0] for d in c.description]
        result = dict(zip(cols, row))
    else:
        result = None
    conn.close()
    return result


def search_kp(query: str) -> list:
    conn = get_conn()
    c = conn.cursor()
    if DATABASE_URL:
        c.execute('''
            SELECT * FROM kp_journal
            WHERE client ILIKE %s OR equipment_list ILIKE %s OR kp_number ILIKE %s
            ORDER BY created_at DESC LIMIT 10
        ''', (f'%{query}%', f'%{query}%', f'%{query}%'))
    else:
        c.execute('''
            SELECT * FROM kp_journal
            WHERE client LIKE ? OR equipment_list LIKE ? OR kp_number LIKE ?
            ORDER BY created_at DESC LIMIT 10
        ''', (f'%{query}%', f'%{query}%', f'%{query}%'))
    cols = [d[0] for d in c.description]
    rows = [dict(zip(cols, row)) for row in c.fetchall()]
    conn.close()
    return rows


def get_recent_kp(manager_id: int = None, limit: int = 10) -> list:
    conn = get_conn()
    c = conn.cursor()
    ph = '%s' if DATABASE_URL else '?'
    if manager_id:
        c.execute(f'''
            SELECT * FROM kp_journal WHERE manager_id = {ph}
            ORDER BY created_at DESC LIMIT {ph}
        ''', (manager_id, limit))
    else:
        c.execute(f'SELECT * FROM kp_journal ORDER BY created_at DESC LIMIT {ph}', (limit,))
    cols = [d[0] for d in c.description]
    rows = [dict(zip(cols, row)) for row in c.fetchall()]
    conn.close()
    return rows


if __name__ == '__main__':
    init_db()
