import os
import json
from datetime import datetime
from docx import Document
from database import get_equipment_by_model

OUTPUT_DIR = 'output'
TEMPLATE_PATH = 'template/kp_template.docx'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs('template', exist_ok=True)


def _replace_text_in_runs(paragraph, placeholder, value):
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
    if not specs:
        return
    target_idx = None
    for i, para in enumerate(doc.paragraphs):
        if 'Технические характеристики' in para.text:
            target_idx = i
            break
    if target_idx is None:
        return
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = 'Характеристика'
    hdr[1].text = 'Значение'
    for cell in hdr:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
    for spec in specs:
        row = table.add_row().cells
        row[0].text = spec.get('name', '')
        para = row[1].paragraphs[0]
        run = para.add_run(str(spec.get('value', '')))
        run.bold = True
    target_para = doc.paragraphs[target_idx]._element
    target_para.addnext(table._tbl)


def _find_cyrillic_font():
    """Ищет шрифт с поддержкой кириллицы на сервере"""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    font_paths = [
        # DejaVu
        ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
        # Liberation
        ('/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf', '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'),
        # FreeSans
        ('/usr/share/fonts/truetype/freefont/FreeSans.ttf', '/usr/share/fonts/truetype/freefont/FreeSansBold.ttf'),
        # Nix paths
        ('/run/current-system/sw/share/X11/fonts/DejaVuSans.ttf', '/run/current-system/sw/share/X11/fonts/DejaVuSans-Bold.ttf'),
    ]

    # Ищем любой ttf с кириллицей
    import subprocess
    try:
        result = subprocess.run(['find', '/usr', '/nix', '/run', '-name', '*.ttf', '-type', 'f'],
                               capture_output=True, text=True, timeout=5)
        all_fonts = result.stdout.strip().split('\n')
        # Предпочитаем DejaVu или Liberation
        for font in all_fonts:
            if 'DejaVuSans.ttf' in font or 'LiberationSans-Regular.ttf' in font or 'FreeSans.ttf' in font:
                bold_font = font.replace('DejaVuSans.ttf', 'DejaVuSans-Bold.ttf')\
                               .replace('LiberationSans-Regular.ttf', 'LiberationSans-Bold.ttf')\
                               .replace('FreeSans.ttf', 'FreeSansBold.ttf')
                if os.path.exists(font):
                    try:
                        pdfmetrics.registerFont(TTFont('CyrFont', font))
                        if os.path.exists(bold_font):
                            pdfmetrics.registerFont(TTFont('CyrFont-Bold', bold_font))
                        else:
                            pdfmetrics.registerFont(TTFont('CyrFont-Bold', font))
                        print(f"Найден шрифт: {font}")
                        return 'CyrFont', 'CyrFont-Bold'
                    except Exception as e:
                        print(f"Шрифт {font} не подошёл: {e}")
                        continue
    except Exception as e:
        print(f"Поиск шрифтов не удался: {e}")

    # Пробуем заранее известные пути
    for regular, bold in font_paths:
        if os.path.exists(regular):
            try:
                pdfmetrics.registerFont(TTFont('CyrFont', regular))
                if os.path.exists(bold):
                    pdfmetrics.registerFont(TTFont('CyrFont-Bold', bold))
                else:
                    pdfmetrics.registerFont(TTFont('CyrFont-Bold', regular))
                print(f"Найден шрифт: {regular}")
                return 'CyrFont', 'CyrFont-Bold'
            except Exception as e:
                print(f"Шрифт {regular} не подошёл: {e}")

    print("Кириллический шрифт не найден, используем Helvetica")
    return 'Helvetica', 'Helvetica-Bold'


def generate_kp_document(kp_data: dict, manager_name: str) -> tuple[str, str]:
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
        price_str = f"{kp_data.get('total_price', 0):,.0f} {kp_data.get('currency', 'ЮАНЕЙ')} (итого)"
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
        '{{PRODUCTION_TIME}}': kp_data.get('production_time', '25-30 дней'),
        '{{PACKAGING}}': kp_data.get('packaging', 'экспортная деревянная тара (ящик)'),
        '{{PAYMENT_TERMS}}': kp_data.get('payment_terms', '50% предоплата, 50% по факту поставки'),
    }

    _replace_in_document(doc, replacements)

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
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    pdf_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.pdf')
    font_name, font_bold = _find_cyrillic_font()

    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                            leftMargin=2*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    normal = ParagraphStyle('normal', fontName=font_name, fontSize=10, leading=14)
    bold_style = ParagraphStyle('bold', fontName=font_bold, fontSize=10, leading=14)
    title_style = ParagraphStyle('title', fontName=font_bold, fontSize=14, leading=18, alignment=1)
    subtitle_style = ParagraphStyle('subtitle', fontName=font_bold, fontSize=12, leading=16, alignment=1)
    right_style = ParagraphStyle('right', fontName=font_name, fontSize=10, leading=14, alignment=2)

    story = []
    story.append(Paragraph(f'от {kp_date}    №    {kp_number}', right_style))
    story.append(Paragraph('г. Таганрог', right_style))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph('КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ', title_style))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph('ООО «Фарсал» предлагает к поставке', subtitle_style))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(equipment_name, subtitle_style))
    story.append(Spacer(1, 0.5*cm))

    if eq and eq.get('specs'):
        try:
            specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
            if specs:
                story.append(Paragraph('Технические характеристики', bold_style))
                story.append(Spacer(1, 0.3*cm))
                table_data = [['Характеристика', 'Значение']]
                for spec in specs:
                    table_data.append([spec.get('name', ''), str(spec.get('value', ''))])
                t = Table(table_data, colWidths=[9*cm, 8*cm])
                t.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, 0), font_bold),
                    ('FONTNAME', (0, 1), (-1, -1), font_name),
                    ('FONTNAME', (1, 1), (1, -1), font_bold),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ]))
                story.append(t)
                story.append(Spacer(1, 0.5*cm))
        except Exception as e:
            print(f"Ошибка PDF характеристик: {e}")

    items = kp_data.get('items', [])
    item = items[0] if items else {}
    unit_price = item.get('unit_price', 0)
    currency = item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))

    story.append(Paragraph('Гарантия', bold_style))
    story.append(Paragraph('Гарантийный срок: 1 год. Изнашиваемые детали гарантийному обслуживанию не подлежат.', normal))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(f'Сроки изготовления: {kp_data.get("production_time", "25-30 дней")}.', normal))
    story.append(Paragraph(f'Упаковка: {kp_data.get("packaging", "экспортная деревянная тара (ящик)")}.', normal))
    story.append(Paragraph(f'Условия оплаты: {kp_data.get("payment_terms", "50% предоплата, 50% по факту поставки")}.', normal))
    story.append(Spacer(1, 0.3*cm))

    if len(items) > 1:
        price_str = f"{kp_data.get('total_price', 0):,.0f} {kp_data.get('currency', 'ЮАНЕЙ')} (итого)"
    else:
        price_str = f"{unit_price:,.0f} {currency}"

    story.append(Paragraph(f'Цена с НДС с доставкой до завода покупателя за 1 штуку: {price_str}.', bold_style))
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph('С уважением,', normal))
    story.append(Paragraph('директор ООО «Фарсал»,', normal))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph('МП     _______________       А. Ю. Лавришко', normal))

    doc.build(story)
    return pdf_path


def cleanup_temp_files(docx_path: str, pdf_path: str):
    for path in [docx_path, pdf_path]:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass
