import sqlite3
import os
from datetime import datetime

DB_PATH = "farsalai.db"


def init_db():
    """Инициализация базы данных"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Таблица оборудования (библиотека)
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

    # Таблица журнала КП
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


# ─── Оборудование ────────────────────────────────────────────────────────────

def add_equipment(data: dict) -> int:
    """Добавить оборудование в библиотеку"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO equipment
        (name, model, description, specs, construction, extra_tables,
         production_time, packaging, base_price, currency, photo_path)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data.get('name'),
        data.get('model'),
        data.get('description'),
        data.get('specs'),
        data.get('construction'),
        data.get('extra_tables'),
        data.get('production_time', '25-30 дней'),
        data.get('packaging', 'экспортная деревянная тара (ящик)'),
        data.get('base_price'),
        data.get('currency', 'ЮАНЕЙ'),
        data.get('photo_path'),
    ))
    eq_id = c.lastrowid
    conn.commit()
    conn.close()
    return eq_id


def get_equipment_by_model(model: str) -> dict | None:
    """Найти оборудование по модели"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM equipment WHERE model LIKE ?', (f'%{model}%',))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_equipment() -> list:
    """Получить весь список оборудования"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT id, name, model, base_price, currency FROM equipment ORDER BY name')
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def search_equipment(query: str) -> list:
    """Поиск оборудования по названию или модели"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('''
        SELECT * FROM equipment
        WHERE name LIKE ? OR model LIKE ?
    ''', (f'%{query}%', f'%{query}%'))
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_equipment(model: str, data: dict):
    """Обновить данные оборудования"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    fields = ', '.join(f'{k} = ?' for k in data.keys())
    values = list(data.values()) + [model]
    c.execute(f'UPDATE equipment SET {fields} WHERE model = ?', values)
    conn.commit()
    conn.close()


def delete_equipment(model: str):
    """Удалить оборудование"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM equipment WHERE model = ?', (model,))
    conn.commit()
    conn.close()


# ─── Журнал КП ───────────────────────────────────────────────────────────────

def generate_kp_number(equipment_models: list) -> str:
    """Генерация номера КП"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT COUNT(*) FROM kp_journal')
    count = c.fetchone()[0] + 1
    conn.close()
    model_part = equipment_models[0].replace('-', '') if equipment_models else 'KP'
    return f"{model_part}-{count:03d}"


def save_kp(data: dict) -> int:
    """Сохранить КП в журнал"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT INTO kp_journal
        (kp_number, kp_date, client, equipment_list, total_price, currency,
         payment_terms, manager_id, manager_name,
         yandex_word_url, yandex_pdf_url, yandex_word_id, yandex_pdf_id, sheets_row)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data.get('kp_number'),
        data.get('kp_date', datetime.now().strftime('%d.%m.%Y')),
        data.get('client'),
        data.get('equipment_list'),
        data.get('total_price'),
        data.get('currency', 'ЮАНЕЙ'),
        data.get('payment_terms', '50% предоплата, 50% по факту поставки'),
        data.get('manager_id'),
        data.get('manager_name'),
        data.get('yandex_word_url'),
        data.get('yandex_pdf_url'),
        data.get('yandex_word_id'),
        data.get('yandex_pdf_id'),
        data.get('sheets_row'),
    ))
    kp_id = c.lastrowid
    conn.commit()
    conn.close()
    return kp_id


def update_kp(kp_number: str, data: dict):
    """Обновить КП (при правках)"""
    data['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    fields = ', '.join(f'{k} = ?' for k in data.keys())
    values = list(data.values()) + [kp_number]
    c.execute(f'UPDATE kp_journal SET {fields} WHERE kp_number = ?', values)
    conn.commit()
    conn.close()


def get_kp_by_number(kp_number: str) -> dict | None:
    """Получить КП по номеру"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM kp_journal WHERE kp_number = ?', (kp_number,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def search_kp(query: str) -> list:
    """Поиск КП по клиенту или оборудованию"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('''
        SELECT * FROM kp_journal
        WHERE client LIKE ? OR equipment_list LIKE ? OR kp_number LIKE ?
        ORDER BY created_at DESC
        LIMIT 10
    ''', (f'%{query}%', f'%{query}%', f'%{query}%'))
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_recent_kp(manager_id: int = None, limit: int = 10) -> list:
    """Получить последние КП"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    if manager_id:
        c.execute('''
            SELECT * FROM kp_journal WHERE manager_id = ?
            ORDER BY created_at DESC LIMIT ?
        ''', (manager_id, limit))
    else:
        c.execute('SELECT * FROM kp_journal ORDER BY created_at DESC LIMIT ?', (limit,))
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]


if __name__ == '__main__':
    init_db()