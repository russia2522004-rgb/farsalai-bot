import os
import json
import time
import requests
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

YANDEX_TOKEN = os.getenv('YANDEX_DISK_TOKEN')
YANDEX_API = 'https://cloud-api.yandex.net/v1/disk'
HEADERS = {'Authorization': f'OAuth {YANDEX_TOKEN}'}

SHEETS_ID = os.getenv('GOOGLE_SHEETS_ID')
CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_FILE')

YANDEX_BASE_FOLDER = 'ФарсалИИ/КП'


# ─── Яндекс Диск ─────────────────────────────────────────────────────────────

def _ensure_folder(path: str):
    """Создаёт папку на Яндекс Диске если не существует. Обрабатывает вложенные папки."""
    parts = [p for p in path.strip('/').split('/') if p]
    current = ''
    for part in parts:
        current = f'{current}/{part}' if current else part
        # Проверяем существует ли папка
        r = requests.get(f'{YANDEX_API}/resources',
                         headers=HEADERS,
                         params={'path': current})
        if r.status_code == 200:
            continue  # папка уже есть — идём дальше

        # Папка не существует — создаём
        r = requests.put(f'{YANDEX_API}/resources',
                         headers=HEADERS,
                         params={'path': current})

        if r.status_code == 201:
            time.sleep(0.5)  # небольшая пауза после создания
            continue

        if r.status_code == 409:
            # Проверяем причину 409
            error_code = r.json().get('error', '')
            if 'ExistentDirectory' in error_code or 'AlreadyExists' in error_code:
                continue  # папка уже есть — ок
            # Иначе — реальная ошибка
            print(f"Ошибка создания папки {current}: {error_code}")
        else:
            print(f"Неожиданный ответ при создании папки {current}: {r.status_code}")


def _get_folder_for_kp() -> str:
    """Возвращает путь папки для текущего месяца"""
    now = datetime.now()
    month_names = {
        1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель',
        5: 'май', 6: 'июнь', 7: 'июль', 8: 'август',
        9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декабрь'
    }
    folder = f'{YANDEX_BASE_FOLDER}/{now.year}/{month_names[now.month]}'
    _ensure_folder(folder)
    return folder


def upload_file_to_yandex(local_path: str, remote_name: str, existing_resource_id: str = None) -> tuple[str, str]:
    """
    Загружает файл на Яндекс Диск.
    Возвращает (публичная ссылка, path на диске)
    """
    folder = _get_folder_for_kp()
    remote_path = f'{folder}/{remote_name}'

    # Удаляем файл если уже существует
    requests.delete(f'{YANDEX_API}/resources',
                    headers=HEADERS,
                    params={'path': remote_path, 'permanently': 'true'})
    time.sleep(1)

    # Получаем URL для загрузки
    r = requests.get(f'{YANDEX_API}/resources/upload',
                     headers=HEADERS,
                     params={'path': remote_path})
    r.raise_for_status()
    upload_url = r.json()['href']

    # Загружаем файл
    with open(local_path, 'rb') as f:
        requests.put(upload_url, data=f)

    time.sleep(1)

    # Публикуем файл
    requests.put(f'{YANDEX_API}/resources/publish',
                 headers=HEADERS,
                 params={'path': remote_path})

    time.sleep(1)

    # Получаем публичную ссылку
    r = requests.get(f'{YANDEX_API}/resources',
                     headers=HEADERS,
                     params={'path': remote_path})
    public_url = r.json().get('public_url', '')

    return public_url, remote_path


def upload_kp_files(word_path: str, pdf_path: str, kp_number: str) -> tuple[str, str]:
    """
    Загружает Word и PDF файлы КП на Яндекс Диск.
    Возвращает (ссылка на Word, ссылка на PDF)
    """
    word_name = f'КП_{kp_number}.docx'
    pdf_name = f'КП_{kp_number}.pdf'

    word_url, _ = upload_file_to_yandex(word_path, word_name)
    pdf_url, _ = upload_file_to_yandex(pdf_path, pdf_name)

    return word_url, pdf_url


def upload_equipment_photo(local_path: str, model_or_path: str) -> str:
    """
    Загружает фото оборудования на Яндекс Диск.
    model_or_path: модель (например 'IE-2') или путь внутри Библиотеки (например 'IE-2/blocks/block0_img.png')
    """
    if '/' in model_or_path:
        # Произвольный путь внутри Библиотеки
        parts = model_or_path.split('/')
        folder = f'{YANDEX_BASE_FOLDER}/Библиотека/{"/".join(parts[:-1])}'
        remote_name = parts[-1]
    else:
        # Просто модель — сохраняем как фото
        folder = f'{YANDEX_BASE_FOLDER}/Библиотека/{model_or_path}'
        ext = os.path.splitext(local_path)[1]
        remote_name = f'фото{ext}'

    _ensure_folder(folder)
    remote_path = f'{folder}/{remote_name}'

    requests.delete(f'{YANDEX_API}/resources',
                    headers=HEADERS,
                    params={'path': remote_path, 'permanently': 'true'})
    time.sleep(1)

    r = requests.get(f'{YANDEX_API}/resources/upload',
                     headers=HEADERS,
                     params={'path': remote_path})
    r.raise_for_status()
    upload_url = r.json()['href']

    with open(local_path, 'rb') as f:
        requests.put(upload_url, data=f)

    return remote_path


# ─── Google Sheets ────────────────────────────────────────────────────────────

def _get_sheet():
    """Подключение к Google Sheets"""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    # Пробуем сначала из переменной окружения
    creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    else:
        # Fallback на файл (для локальной разработки)
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)

    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEETS_ID).sheet1


def add_kp_to_sheets(kp_data: dict, word_url: str, pdf_url: str) -> int:
    """
    Добавляет строку КП в Google Sheets.
    Возвращает номер строки.
    """
    sheet = _get_sheet()

    row = [
        kp_data.get('kp_date', datetime.now().strftime('%d.%m.%Y')),
        kp_data.get('kp_number', ''),
        kp_data.get('client', ''),
        kp_data.get('equipment_list', ''),
        kp_data.get('total_price', ''),
        kp_data.get('currency', 'ЮАНЕЙ'),
        kp_data.get('manager_name', ''),
        word_url,
        pdf_url,
    ]

    sheet.append_row(row)
    return len(sheet.get_all_values())


def update_kp_in_sheets(row_number: int, word_url: str, pdf_url: str):
    """Обновляет ссылки на файлы в Google Sheets при перезаписи"""
    sheet = _get_sheet()
    sheet.update_cell(row_number, 8, word_url)
    sheet.update_cell(row_number, 9, pdf_url)


def ensure_headers():
    """Проверяет и создаёт заголовки таблицы если их нет"""
    sheet = _get_sheet()
    first_row = sheet.row_values(1)
    if not first_row or first_row[0] != 'Дата':
        headers = ['Дата', 'Номер КП', 'Клиент', 'Оборудование',
                   'Сумма', 'Валюта', 'Менеджер', 'Word', 'PDF']
        sheet.insert_row(headers, 1)
        sheet.format('A1:I1', {'textFormat': {'bold': True}})
