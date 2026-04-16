import os
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from database import get_equipment_by_model

OUTPUT_DIR = 'output'
TEMPLATE_PATH = 'template/kp_template.docx'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs('template', exist_ok=True)


def _replace_text_in_runs(paragraph, placeholder, value):
    """Заменяет плейсхолдер в параграфе с сохранением форматирования"""
    full_text = ''.join([run.text for run in paragraph.runs])
    if placeholder not in full_text:
        return False

    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            return True

    if placeholder in full_text:
        new_text = full_text.replace(placeholder, value)
        for i, run in enumerate(paragraph.runs):
            if i == 0:
                run.text = new_text
            else:
                run.text = ''
        return True

    return False


def _replace_in_document(doc, replacements: dict):
    """Заменяет все плейсхолдеры в документе"""
    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            _replace_text_in_runs(paragraph, placeholder, str(value))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for placeholder, value in replacements.items():
                        _replace_text_in_runs(paragraph, placeholder, str(value))


def _add_specs_after_heading(doc, specs: list):
    """Добавляет таблицу характеристик после заголовка 'Технические характеристики'"""
    if not specs:
        return

    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Cm

    # Находим параграф с заголовком
    target_idx = None
    for i, para in enumerate(doc.paragraphs):
        if 'Технические характеристики' in para.text:
            target_idx = i
            break

    if target_idx is None:
        return

    # Создаём таблицу характеристик
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'

    # Заголовок таблицы
    hdr = table.rows[0].cells
    hdr[0].text = 'Характеристика'
    hdr[1].text = 'Значение'
    for cell in hdr:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True

    # Строки характеристик
    for spec in specs:
        row = table.add_row().cells
        row[0].text = spec.get('name', '')
        para = row[1].paragraphs[0]
        run = para.add_run(str(spec.get('value', '')))
        run.bold = True

    # Перемещаем таблицу после заголовка
    target_para = doc.paragraphs[target_idx]._element
    target_para.addnext(table._tbl)


def generate_kp_document(kp_data: dict, manager_name: str) -> tuple[str, str]:
    """Генерирует КП на основе шаблона. Возвращает (путь к docx, путь к pdf)"""
    kp_number = kp_data.get('kp_number', 'KP-001')
    kp_date = kp_data.get('kp_date', datetime.now().strftime('%d.%m.%Y'))
    items = kp_data.get('items', [])

    item = items[0] if items else {}
    model = item.get('model', '')
    eq = get_equipment_by_model(model)

    equipment_name = eq['name'] if eq else item.get('name', model)
    unit_price = item.get('unit_price', 0)
    currency = item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))

    if len(items) > 1:
        price_str = f"{kp_data.get('total_price', 0):,.0f} {kp_data.get('currency', 'ЮАНЕЙ')} (итого за все позиции)"
    else:
        price_str = f"{unit_price:,.0f} {currency}"

    if os.path.exists(TEMPLATE_PATH):
        doc = Document(TEMPLATE_PATH)
    else:
        raise FileNotFoundError(f"Шаблон не найден: {TEMPLATE_PATH}")

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

    # Добавляем характеристики из библиотеки
    if eq and eq.get('specs'):
        try:
            specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
            if specs:
                _add_specs_after_heading(doc, specs)
        except Exception as e:
            print(f"Ошибка добавления характеристик: {e}")

    docx_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.docx')
    doc.save(docx_path)

    pdf_path = _convert_to_pdf_reportlab(kp_data, equipment_name, kp_number, kp_date, eq)

    return docx_path, pdf_path


def _convert_to_pdf_reportlab(kp_data: dict, equipment_name: str, kp_number: str, kp_date: str, eq: dict) -> str:
    """Генерирует PDF через reportlab"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import io

    pdf_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.pdf')

    # Регистрируем шрифт с поддержкой кириллицы
    try:
        pdfmetrics.registerFont(TTFont('Arial', '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-Bold', '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'))
        font_name = 'Arial'
        font_bold = 'Arial-Bold'
    except Exception:
        try:
            pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
            pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
            font_name = 'DejaVu'
            font_bold = 'DejaVu-Bold'
        except Exception:
            font_name = 'Helvetica'
            font_bold = 'Helvetica-Bold'

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=1.5*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    normal = ParagraphStyle('normal', fontName=font_name, fontSize=10, leading=14)
    bold = ParagraphStyle('bold', fontName=font_bold, fontSize=10, leading=14)
    title = ParagraphStyle('title', fontName=font_bold, fontSize=14, leading=18, alignment=1)
    subtitle = ParagraphStyle('subtitle', fontName=font_bold, fontSize=12, leading=16, alignment=1)
    right = ParagraphStyle('right', fontName=font_name, fontSize=10, leading=14, alignment=2)

    story = []

    # Шапка
    story.append(Paragraph(f'от {kp_date}    №    {kp_number}', right))
    story.append(Paragraph('г. Таганрог', right))
    story.append(Spacer(1, 0.5*cm))

    # Заголовок
    story.append(Paragraph('КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ', title))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph('ООО «Фарсал» предлагает к поставке', subtitle))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(equipment_name, subtitle))
    story.append(Spacer(1, 0.5*cm))

    # Технические характеристики
    if eq and eq.get('specs'):
        try:
            specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
            if specs:
                story.append(Paragraph('Технические характеристики', bold))
                story.append(Spacer(1, 0.3*cm))

                table_data = [['Характеристика', 'Значение']]
                for spec in specs:
                    table_data.append([spec.get('name', ''), spec.get('value', '')])

                t = Table(table_data, colWidths=[9*cm, 8*cm])
                t.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, 0), font_bold),
                    ('FONTNAME', (0, 1), (-1, -1), font_name),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                    ('FONTNAME', (1, 1), (1, -1), font_bold),
                ]))
                story.append(t)
                story.append(Spacer(1, 0.5*cm))
        except Exception as e:
            print(f"Ошибка PDF характеристик: {e}")

    # Условия
    items = kp_data.get('items', [])
    item = items[0] if items else {}
    unit_price = item.get('unit_price', 0)
    currency = item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))

    story.append(Paragraph('Гарантия', bold))
    story.append(Paragraph('Гарантийный срок: 1 год. Изнашиваемые детали гарантийному обслуживанию не подлежат.', normal))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(f'Сроки изготовления: {kp_data.get("production_time", "25–30 дней")}.', normal))
    story.append(Paragraph(f'Упаковка: {kp_data.get("packaging", "экспортная деревянная тара (ящик)")}.', normal))
    story.append(Paragraph(f'Условия оплаты: {kp_data.get("payment_terms", "50% – предоплата, 50% – по факту поставки")}.', normal))
    story.append(Spacer(1, 0.3*cm))

    if len(items) > 1:
        price_str = f"{kp_data.get('total_price', 0):,.0f} {kp_data.get('currency', 'ЮАНЕЙ')} (итого)"
    else:
        price_str = f"{unit_price:,.0f} {currency}"

    story.append(Paragraph(f'Цена с НДС с доставкой до завода покупателя за 1 штуку: {price_str}.', bold))
    story.append(Spacer(1, 1*cm))

    # Подпись
    story.append(Paragraph('С уважением,', normal))
    story.append(Paragraph('директор ООО «Фарсал»,', normal))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph('МП     _______________       А. Ю. Лавришко', normal))

    doc.build(story)
    return pdf_path


def cleanup_temp_files(docx_path: str, pdf_path: str):
    """Удаляет временные файлы"""
    for path in [docx_path, pdf_path]:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass
