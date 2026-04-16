import os
import json
import subprocess
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from database import get_equipment_by_model

OUTPUT_DIR = 'output'
TEMPLATE_PATH = 'template/kp_template.docx'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs('template', exist_ok=True)


def _replace_text_in_runs(paragraph, placeholder, value):
    """Заменяет плейсхолдер в параграфе с сохранением форматирования"""
    # Сначала проверяем весь текст параграфа
    full_text = ''.join([run.text for run in paragraph.runs])
    if placeholder not in full_text:
        return False

    # Если плейсхолдер в одном run
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            return True

    # Если плейсхолдер разбит по нескольким runs — объединяем
    if placeholder in full_text:
        new_text = full_text.replace(placeholder, value)
        # Очищаем все runs кроме первого
        for i, run in enumerate(paragraph.runs):
            if i == 0:
                run.text = new_text
            else:
                run.text = ''
        return True

    return False


def _replace_in_document(doc, replacements: dict):
    """Заменяет все плейсхолдеры в документе"""
    # В параграфах
    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            _replace_text_in_runs(paragraph, placeholder, str(value))

    # В таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for placeholder, value in replacements.items():
                        _replace_text_in_runs(paragraph, placeholder, str(value))


def _replace_specs_table(doc, specs: list):
    """Заменяет таблицу характеристик данными из библиотеки"""
    if not specs:
        return

    # Находим таблицу с характеристиками (ищем по заголовку)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text = ''.join([p.text for p in cell.paragraphs])
                if 'Характеристика' in full_text or 'Материал' in full_text:
                    # Нашли таблицу характеристик
                    # Удаляем все строки кроме заголовка
                    while len(table.rows) > 1:
                        tr = table.rows[-1]._tr
                        tr.getparent().remove(tr)

                    # Добавляем строки из библиотеки
                    for spec in specs:
                        row = table.add_row()
                        row.cells[0].text = spec.get('name', '')
                        row.cells[1].text = spec.get('value', '')
                        # Делаем значение жирным
                        for para in row.cells[1].paragraphs:
                            for run in para.runs:
                                run.bold = True
                    return


def generate_kp_document(kp_data: dict, manager_name: str) -> tuple[str, str]:
    """
    Генерирует КП на основе шаблона.
    Возвращает (путь к docx, путь к pdf)
    """
    kp_number = kp_data.get('kp_number', 'KP-001')
    kp_date = kp_data.get('kp_date', datetime.now().strftime('%d.%m.%Y'))
    items = kp_data.get('items', [])

    # Берём первую позицию (основное оборудование)
    item = items[0] if items else {}
    model = item.get('model', '')
    eq = get_equipment_by_model(model)

    # Название оборудования
    equipment_name = eq['name'] if eq else item.get('name', model)

    # Цена
    quantity = item.get('quantity', 1)
    unit_price = item.get('unit_price', 0)
    currency = item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))

    if len(items) > 1:
        # Несколько позиций — показываем итог
        price_str = f"{kp_data.get('total_price', 0):,.0f} {kp_data.get('currency', 'ЮАНЕЙ')} (итого за все позиции)"
    else:
        price_str = f"{unit_price:,.0f} {currency}"

    # Загружаем шаблон
    if os.path.exists(TEMPLATE_PATH):
        doc = Document(TEMPLATE_PATH)
    else:
        raise FileNotFoundError(f"Шаблон не найден: {TEMPLATE_PATH}")

    # Заменяем плейсхолдеры
    replacements = {
        '{{DATE}}': kp_date,
        '{{KP_NUMBER}}': kp_number,
        '{{EQUIPMENT_NAME}}': equipment_name,
        '{{PRICE}}': price_str,
        '{{PRODUCTION_TIME}}': kp_data.get('production_time', '25–30 дней'),
        '{{PACKAGING}}': kp_data.get('packaging', 'экспортная деревянная тара (ящик)'),
        '{{PAYMENT_TERMS}}': kp_data.get('payment_terms', '50% – предоплата, 50% – по факту поставки'),
    }

    _replace_in_document(doc, replacements)

    # Заменяем таблицу характеристик если есть данные в библиотеке
    if eq and eq.get('specs'):
        try:
            specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
            if specs:
                _replace_specs_table(doc, specs)
        except Exception as e:
            print(f"Ошибка замены характеристик: {e}")

    # Сохраняем Word
    docx_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.docx')
    doc.save(docx_path)

    # Конвертируем в PDF
    pdf_path = _convert_to_pdf(docx_path, kp_number)

    return docx_path, pdf_path


def _convert_to_pdf(docx_path: str, kp_number: str) -> str:
    """Конвертирует Word в PDF"""
    pdf_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.pdf')

    try:
        # Пробуем LibreOffice
        result = subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf',
            '--outdir', OUTPUT_DIR, docx_path
        ], check=True, capture_output=True, timeout=60)

        # LibreOffice сохраняет с тем же именем но .pdf
        generated = docx_path.replace('.docx', '.pdf')
        if os.path.exists(generated) and generated != pdf_path:
            os.rename(generated, pdf_path)

    except Exception as e1:
        print(f"LibreOffice недоступен: {e1}")
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
        except Exception as e2:
            print(f"docx2pdf недоступен: {e2}")
            # Возвращаем docx как fallback
            return docx_path

    return pdf_path


def cleanup_temp_files(docx_path: str, pdf_path: str):
    """Удаляет временные файлы"""
    for path in [docx_path, pdf_path]:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass
