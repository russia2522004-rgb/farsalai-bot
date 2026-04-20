import os
import json
from datetime import datetime

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
    conn = get_conn()
    c = conn.cursor()

    if DATABASE_URL:
        c.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                model TEXT NOT NULL UNIQUE,
                description TEXT,
                specs TEXT,
                warranty TEXT DEFAULT '1 год. Изнашиваемые детали гарантийному обслуживанию не подлежат.',
                production_time TEXT DEFAULT '25-30 дней',
                packaging TEXT DEFAULT 'экспортная деревянная тара (ящик)',
                delivery TEXT DEFAULT 'до завода покупателя',
                payment_terms TEXT DEFAULT '50% – предоплата, 50% – по факту поставки',
                base_price REAL,
                currency TEXT DEFAULT 'ЮАНЕЙ',
                photo_path TEXT,
                original_file_path TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS equipment_blocks (
                id SERIAL PRIMARY KEY,
                equipment_id INTEGER REFERENCES equipment(id) ON DELETE CASCADE,
                block_type TEXT NOT NULL,
                block_title TEXT,
                xml_content TEXT,
                images TEXT DEFAULT '[]',
                sort_order INTEGER DEFAULT 0,
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
                sheets_row INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Добавляем новые колонки если их нет (миграция)
        for col, definition in [
            ('warranty', "TEXT DEFAULT '1 год.'"),
            ('delivery', "TEXT DEFAULT 'до завода покупателя'"),
            ('payment_terms', "TEXT DEFAULT '50% – предоплата, 50% – по факту поставки'"),
            ('original_file_path', 'TEXT'),
        ]:
            try:
                c.execute(f'ALTER TABLE equipment ADD COLUMN {col} {definition}')
            except Exception:
                pass

    else:
        c.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                model TEXT NOT NULL UNIQUE,
                description TEXT,
                specs TEXT,
                warranty TEXT DEFAULT '1 год. Изнашиваемые детали гарантийному обслуживанию не подлежат.',
                production_time TEXT DEFAULT '25-30 дней',
                packaging TEXT DEFAULT 'экспортная деревянная тара (ящик)',
                delivery TEXT DEFAULT 'до завода покупателя',
                payment_terms TEXT DEFAULT '50% – предоплата, 50% – по факту поставки',
                base_price REAL,
                currency TEXT DEFAULT 'ЮАНЕЙ',
                photo_path TEXT,
                original_file_path TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS equipment_blocks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipment_id INTEGER REFERENCES equipment(id) ON DELETE CASCADE,
                block_type TEXT NOT NULL,
                block_title TEXT,
                xml_content TEXT,
                images TEXT DEFAULT '[]',
                sort_order INTEGER DEFAULT 0,
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
                sheets_row INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

    conn.commit()
    conn.close()
    print("База данных инициализирована")


def _rows_to_dicts(cursor):
    cols = [d[0] for d in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


def _row_to_dict(cursor, row):
    cols = [d[0] for d in cursor.description]
    return dict(zip(cols, row)) if row else None


# ─── Оборудование ────────────────────────────────────────────────────────────

def add_equipment(data: dict) -> int:
    conn = get_conn()
    c = conn.cursor()
    ph = PLACEHOLDER

    fields = ['name', 'model', 'description', 'specs', 'warranty',
              'production_time', 'packaging', 'delivery', 'payment_terms',
              'base_price', 'currency', 'photo_path', 'original_file_path']
    values = [data.get(f) for f in fields]

    if DATABASE_URL:
        placeholders = ', '.join([ph] * len(fields))
        c.execute(f'''
            INSERT INTO equipment ({', '.join(fields)})
            VALUES ({placeholders})
            ON CONFLICT (model) DO UPDATE SET
            {', '.join(f"{f}=EXCLUDED.{f}" for f in fields if f != 'model')}
            RETURNING id
        ''', values)
        eq_id = c.fetchone()[0]
    else:
        placeholders = ', '.join([ph] * len(fields))
        c.execute(f'''
            INSERT OR REPLACE INTO equipment ({', '.join(fields)})
            VALUES ({placeholders})
        ''', values)
        eq_id = c.lastrowid

    conn.commit()
    conn.close()
    return eq_id


def get_equipment_by_model(model: str) -> dict | None:
    conn = get_conn()
    c = conn.cursor()
    if DATABASE_URL:
        c.execute('SELECT * FROM equipment WHERE model ILIKE %s', (f'%{model}%',))
    else:
        c.execute('SELECT * FROM equipment WHERE model LIKE ?', (f'%{model}%',))
    row = c.fetchone()
    result = _row_to_dict(c, row)
    conn.close()
    return result


def search_equipment(query: str) -> list:
    conn = get_conn()
    c = conn.cursor()
    if DATABASE_URL:
        c.execute('SELECT * FROM equipment WHERE name ILIKE %s OR model ILIKE %s',
                  (f'%{query}%', f'%{query}%'))
    else:
        c.execute('SELECT * FROM equipment WHERE name LIKE ? OR model LIKE ?',
                  (f'%{query}%', f'%{query}%'))
    result = _rows_to_dicts(c)
    conn.close()
    return result


def get_all_equipment() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute('SELECT id, name, model, base_price, currency FROM equipment ORDER BY name')
    result = _rows_to_dicts(c)
    conn.close()
    return result


def update_equipment(model: str, data: dict):
    conn = get_conn()
    c = conn.cursor()
    ph = PLACEHOLDER
    fields = ', '.join(f'{k} = {ph}' for k in data.keys())
    values = list(data.values()) + [model]
    c.execute(f'UPDATE equipment SET {fields} WHERE model = {ph}', values)
    conn.commit()
    conn.close()


def delete_equipment(model: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(f'DELETE FROM equipment WHERE model = {PLACEHOLDER}', (model,))
    conn.commit()
    conn.close()


# ─── Блоки оборудования ──────────────────────────────────────────────────────

def save_equipment_blocks(equipment_id: int, blocks: list):
    """Сохраняет блоки оборудования. Удаляет старые и сохраняет новые."""
    conn = get_conn()
    c = conn.cursor()
    ph = PLACEHOLDER

    # Удаляем старые блоки
    c.execute(f'DELETE FROM equipment_blocks WHERE equipment_id = {ph}', (equipment_id,))

    # Сохраняем новые
    for i, block in enumerate(blocks):
        images = json.dumps(block.get('images', []), ensure_ascii=False)
        c.execute(f'''
            INSERT INTO equipment_blocks
            (equipment_id, block_type, block_title, xml_content, images, sort_order)
            VALUES ({ph}, {ph}, {ph}, {ph}, {ph}, {ph})
        ''', (
            equipment_id,
            block.get('type', 'unknown'),
            block.get('title', ''),
            block.get('xml', ''),
            images,
            i
        ))

    conn.commit()
    conn.close()


def get_equipment_blocks(equipment_id: int) -> list:
    """Получает блоки оборудования в правильном порядке."""
    conn = get_conn()
    c = conn.cursor()
    c.execute(f'''
        SELECT * FROM equipment_blocks
        WHERE equipment_id = {PLACEHOLDER}
        ORDER BY sort_order
    ''', (equipment_id,))
    result = _rows_to_dicts(c)
    conn.close()
    return result


def update_equipment_block(block_id: int, data: dict):
    conn = get_conn()
    c = conn.cursor()
    ph = PLACEHOLDER
    fields = ', '.join(f'{k} = {ph}' for k in data.keys())
    values = list(data.values()) + [block_id]
    c.execute(f'UPDATE equipment_blocks SET {fields} WHERE id = {ph}', values)
    conn.commit()
    conn.close()


def delete_equipment_block(block_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute(f'DELETE FROM equipment_blocks WHERE id = {PLACEHOLDER}', (block_id,))
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
    ph = PLACEHOLDER

    fields = ['kp_number', 'kp_date', 'client', 'equipment_list', 'total_price',
              'currency', 'payment_terms', 'manager_id', 'manager_name',
              'yandex_word_url', 'yandex_pdf_url', 'sheets_row']
    values = [data.get(f) for f in fields]
    placeholders = ', '.join([ph] * len(fields))

    if DATABASE_URL:
        c.execute(f'''
            INSERT INTO kp_journal ({', '.join(fields)})
            VALUES ({placeholders}) RETURNING id
        ''', values)
        kp_id = c.fetchone()[0]
    else:
        c.execute(f'''
            INSERT INTO kp_journal ({', '.join(fields)})
            VALUES ({placeholders})
        ''', values)
        kp_id = c.lastrowid

    conn.commit()
    conn.close()
    return kp_id


def update_kp(kp_number: str, data: dict):
    data['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_conn()
    c = conn.cursor()
    ph = PLACEHOLDER
    fields = ', '.join(f'{k} = {ph}' for k in data.keys())
    values = list(data.values()) + [kp_number]
    c.execute(f'UPDATE kp_journal SET {fields} WHERE kp_number = {ph}', values)
    conn.commit()
    conn.close()


def get_kp_by_number(kp_number: str) -> dict | None:
    conn = get_conn()
    c = conn.cursor()
    c.execute(f'SELECT * FROM kp_journal WHERE kp_number = {PLACEHOLDER}', (kp_number,))
    row = c.fetchone()
    result = _row_to_dict(c, row)
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
    result = _rows_to_dicts(c)
    conn.close()
    return result


def get_recent_kp(manager_id: int = None, limit: int = 10) -> list:
    conn = get_conn()
    c = conn.cursor()
    ph = PLACEHOLDER
    if manager_id:
        c.execute(f'''
            SELECT * FROM kp_journal WHERE manager_id = {ph}
            ORDER BY created_at DESC LIMIT {ph}
        ''', (manager_id, limit))
    else:
        c.execute(f'SELECT * FROM kp_journal ORDER BY created_at DESC LIMIT {ph}', (limit,))
    result = _rows_to_dicts(c)
    conn.close()
    return result


if __name__ == '__main__':
    init_db()
