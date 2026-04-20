import os
import json
import zipfile
import shutil
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from database import get_equipment_by_model, get_equipment_blocks

OUTPUT_DIR = 'output'
TEMPLATE_PATH = 'template/kp_template.docx'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs('template', exist_ok=True)


def _replace_text_in_runs(paragraph, placeholder, value):
    full_text = ''.join([run.text for run in paragraph.runs])
    if placeholder not in full_text:
        return False
    new_full = full_text.replace(placeholder, value)
    if paragraph.runs:
        paragraph.runs[0].text = new_full
        for run in paragraph.runs[1:]:
            run.text = ''
    return True


def _replace_in_document(doc, replacements: dict):
    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            _replace_text_in_runs(paragraph, placeholder, str(value) if value else '')
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for placeholder, value in replacements.items():
                        _replace_text_in_runs(paragraph, placeholder, str(value) if value else '')


def _find_placeholder_paragraph(doc, placeholder: str):
    """Находит параграф с плейсхолдером"""
    for i, para in enumerate(doc.paragraphs):
        if placeholder in ''.join([r.text for r in para.runs]):
            return i, para
    return None, None


def _add_paragraph_after(doc, ref_para, text, bold=False, size=11):
    """Добавляет параграф после указанного"""
    new_para = OxmlElement('w:p')
    ref_para._element.addnext(new_para)
    from docx.oxml import OxmlElement
    # Проще создать через docx и потом переместить
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    # Перемещаем в нужное место
    ref_para._element.addnext(p._element)
    return p


def _add_specs_table(doc, ref_para, specs: list):
    """Добавляет таблицу характеристик после параграфа"""
    if not specs:
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
        p = row[1].paragraphs[0]
        run = p.add_run(str(spec.get('value', '')))
        run.bold = True

    ref_para._element.addnext(table._tbl)


def _add_summary_table(doc, ref_para, items: list):
    """Добавляет итоговую таблицу для нескольких позиций"""
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'

    headers = ['Оборудование', 'Кол-во', 'Цена за ед.', 'Сумма']
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = headers[i]
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True

    total = 0
    for item in items:
        row = table.add_row().cells
        row[0].text = item.get('name', item.get('model', ''))
        row[1].text = str(item.get('quantity', 1))
        price = item.get('unit_price', 0)
        currency = item.get('currency', '')
        qty = item.get('quantity', 1)
        subtotal = price * qty
        total += subtotal
        row[2].text = f"{price:,.0f} {currency}"
        row[3].text = f"{subtotal:,.0f} {currency}"

    # Итоговая строка
    total_row = table.add_row().cells
    total_row[0].text = 'ИТОГО'
    for cell in total_row:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
    currency = items[0].get('currency', '') if items else ''
    total_row[3].text = f"{total:,.0f} {currency}"
    for para in total_row[3].paragraphs:
        for run in para.runs:
            run.bold = True

    ref_para._element.addnext(table._tbl)


def generate_kp_document(kp_data: dict, manager_name: str) -> tuple[str, str]:
    """Генерирует КП из блоков оборудования"""
    kp_number = kp_data.get('kp_number', 'KP-001')
    kp_date = kp_data.get('kp_date', datetime.now().strftime('%d.%m.%Y'))
    items = kp_data.get('items', [])

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Шаблон не найден: {TEMPLATE_PATH}")

    doc = Document(TEMPLATE_PATH)

    # Заменяем шапку
    _replace_in_document(doc, {
        '{{DATE}}': kp_date,
        '{{KP_NUMBER}}': kp_number,
    })

    # Находим плейсхолдер {{CONTENT}}
    content_idx, content_para = _find_placeholder_paragraph(doc, '{{CONTENT}}')

    if content_para is None:
        # Если нет {{CONTENT}} — добавляем в конец перед подписью
        content_para = doc.paragraphs[-3] if len(doc.paragraphs) > 3 else doc.paragraphs[-1]

    # Очищаем плейсхолдер
    content_para.runs[0].text = '' if content_para.runs else ''

    # Вставляем блоки оборудования в обратном порядке (addnext вставляет после)
    insert_after = content_para

    # Если несколько позиций — добавляем итоговую таблицу последней
    if len(items) > 1:
        summary_title = doc.add_paragraph()
        run = summary_title.add_run('Итоговая стоимость')
        run.bold = True
        run.font.size = Pt(11)
        insert_after._element.addnext(summary_title._element)
        _add_summary_table(doc, summary_title, items)
        insert_after = summary_title

    # Добавляем блоки каждой позиции (в обратном порядке чтобы получить правильный порядок)
    for item in reversed(items):
        model = item.get('model', '')
        eq = get_equipment_by_model(model)
        blocks = get_equipment_blocks(eq['id']) if eq else []

        # Условия для этой позиции
        warranty = item.get('warranty') or (eq.get('warranty') if eq else None) or '1 год.'
        production_time = item.get('production_time') or (eq.get('production_time') if eq else None) or '25-30 дней'
        packaging = item.get('packaging') or (eq.get('packaging') if eq else None) or 'экспортная деревянная тара (ящик)'
        delivery = item.get('delivery') or (eq.get('delivery') if eq else None) or 'до завода покупателя'
        payment_terms = item.get('payment_terms') or kp_data.get('payment_terms') or (eq.get('payment_terms') if eq else None) or '50% предоплата, 50% по факту поставки'
        unit_price = item.get('unit_price', 0)
        currency = item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))

        # Блок условий позиции
        conditions_para = doc.add_paragraph()
        insert_after._element.addnext(conditions_para._element)

        def add_condition(text, bold_label):
            p = doc.add_paragraph()
            run_label = p.add_run(f'{bold_label} ')
            run_label.bold = True
            run_label.font.size = Pt(10)
            run_value = p.add_run(text)
            run_value.font.size = Pt(10)
            conditions_para._element.addnext(p._element)
            return p

        # Добавляем условия снизу вверх
        add_condition(f"{unit_price:,.0f} {currency}.",
                      'Цена с НДС с доставкой до завода покупателя за 1 штуку:')
        add_condition(payment_terms + '.', 'Условия оплаты:')
        add_condition(packaging + '.', 'Упаковка:')
        add_condition(production_time + '.', 'Сроки изготовления:')

        warranty_para = doc.add_paragraph()
        conditions_para._element.addnext(warranty_para._element)
        run_w = warranty_para.add_run('Гарантия')
        run_w.bold = True
        run_w.font.size = Pt(11)

        warranty_text = doc.add_paragraph()
        conditions_para._element.addnext(warranty_text._element)
        run_wt = warranty_text.add_run(warranty)
        run_wt.font.size = Pt(10)

        # Добавляем блоки из библиотеки (в обратном порядке)
        for block in reversed(blocks):
            block_title = block.get('block_title', '')
            xml_content = block.get('xml', block.get('xml_content', ''))

            if block_title:
                title_para = doc.add_paragraph()
                title_run = title_para.add_run(block_title)
                title_run.bold = True
                title_run.font.size = Pt(11)
                conditions_para._element.addnext(title_para._element)

            # Если есть XML — вставляем как есть (сохраняет форматирование)
            if xml_content:
                try:
                    from lxml import etree
                    xml_elem = etree.fromstring(xml_content)
                    conditions_para._element.addnext(xml_elem)
                except Exception:
                    # Fallback — добавляем как текст
                    if block.get('content'):
                        content_p = doc.add_paragraph(block.get('content', ''))
                        content_p.runs[0].font.size = Pt(10) if content_p.runs else None
                        conditions_para._element.addnext(content_p._element)

        # Если нет блоков из библиотеки — добавляем характеристики из specs
        if not blocks and eq and eq.get('specs'):
            try:
                specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
                if specs:
                    specs_title = doc.add_paragraph()
                    run_st = specs_title.add_run('Технические характеристики')
                    run_st.bold = True
                    run_st.font.size = Pt(11)
                    conditions_para._element.addnext(specs_title._element)
                    _add_specs_table(doc, specs_title, specs)
            except Exception as e:
                print(f"Ошибка добавления характеристик: {e}")

        # Фото оборудования
        if eq and eq.get('photo_path'):
            photo_local = f'temp_photo_{kp_number}_{model}.jpg'
            try:
                import requests
                token = os.getenv('YANDEX_DISK_TOKEN')
                headers = {'Authorization': f'OAuth {token}'}
                r = requests.get('https://cloud-api.yandex.net/v1/disk/resources/download',
                                 headers=headers,
                                 params={'path': eq['photo_path']})
                if r.status_code == 200:
                    download_url = r.json().get('href')
                    if download_url:
                        img_r = requests.get(download_url)
                        with open(photo_local, 'wb') as f:
                            f.write(img_r.content)

                        photo_para = doc.add_paragraph()
                        photo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run_photo = photo_para.add_run()
                        run_photo.add_picture(photo_local, width=Inches(4))
                        note_para = doc.add_paragraph(
                            '* Фото для справки, реальные фотографии будут предоставлены после завершения производства.')
                        if note_para.runs:
                            note_para.runs[0].font.size = Pt(9)
                            note_para.runs[0].italic = True
                        conditions_para._element.addnext(note_para._element)
                        conditions_para._element.addnext(photo_para._element)
            except Exception as e:
                print(f"Ошибка загрузки фото: {e}")
            finally:
                if os.path.exists(photo_local):
                    os.remove(photo_local)

        # Название оборудования
        name = eq['name'] if eq else item.get('name', model)
        name_para = doc.add_paragraph()
        run_name = name_para.add_run(name)
        run_name.bold = True
        run_name.font.size = Pt(13)
        name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        conditions_para._element.addnext(name_para._element)

        # Разделитель между позициями
        if len(items) > 1:
            sep = doc.add_paragraph()
            conditions_para._element.addnext(sep._element)

    # Сохраняем
    docx_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.docx')
    doc.save(docx_path)

    pdf_path = _convert_to_pdf_reportlab(kp_data, kp_number)

    return docx_path, pdf_path


def _convert_to_pdf_reportlab(kp_data: dict, kp_number: str) -> str:
    """Генерирует PDF через reportlab"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    pdf_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.pdf')

    # Шрифт
    font_name, font_bold = 'Helvetica', 'Helvetica-Bold'
    for regular, bold in [
        ('fonts/DejaVuSans.ttf', 'fonts/DejaVuSans-Bold.ttf'),
        ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
    ]:
        if os.path.exists(regular):
            try:
                pdfmetrics.registerFont(TTFont('CyrFont', regular))
                pdfmetrics.registerFont(TTFont('CyrFont-Bold', bold if os.path.exists(bold) else regular))
                font_name, font_bold = 'CyrFont', 'CyrFont-Bold'
                break
            except Exception:
                pass

    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                            leftMargin=2*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    normal = ParagraphStyle('n', fontName=font_name, fontSize=10, leading=14)
    bold_s = ParagraphStyle('b', fontName=font_bold, fontSize=10, leading=14)
    title_s = ParagraphStyle('t', fontName=font_bold, fontSize=14, leading=18, alignment=1)
    sub_s = ParagraphStyle('s', fontName=font_bold, fontSize=12, leading=16, alignment=1)
    right_s = ParagraphStyle('r', fontName=font_name, fontSize=10, leading=14, alignment=2)

    story = []
    kp_date = kp_data.get('kp_date', datetime.now().strftime('%d.%m.%Y'))
    kp_num = kp_data.get('kp_number', kp_number)
    items = kp_data.get('items', [])

    story.append(Paragraph(f'от {kp_date}    №    {kp_num}', right_s))
    story.append(Paragraph('г. Таганрог', right_s))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph('КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ', title_s))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph('ООО «Фарсал» предлагает к поставке', sub_s))
    story.append(Spacer(1, 0.5*cm))

    for item in items:
        model = item.get('model', '')
        eq = get_equipment_by_model(model)
        name = eq['name'] if eq else item.get('name', model)

        story.append(Paragraph(name, sub_s))
        story.append(Spacer(1, 0.3*cm))

        # Характеристики
        if eq and eq.get('specs'):
            try:
                specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
                if specs:
                    story.append(Paragraph('Технические характеристики', bold_s))
                    story.append(Spacer(1, 0.2*cm))
                    td = [['Характеристика', 'Значение']]
                    for s in specs:
                        td.append([s.get('name', ''), str(s.get('value', ''))])
                    t = Table(td, colWidths=[9*cm, 8*cm])
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
            except Exception:
                pass

        # Условия позиции
        warranty = item.get('warranty') or (eq.get('warranty') if eq else None) or '1 год.'
        production_time = item.get('production_time') or (eq.get('production_time') if eq else None) or '25-30 дней'
        packaging = item.get('packaging') or (eq.get('packaging') if eq else None) or 'экспортная деревянная тара (ящик)'
        payment_terms = item.get('payment_terms') or kp_data.get('payment_terms') or (eq.get('payment_terms') if eq else None) or '50% предоплата, 50% по факту'
        unit_price = item.get('unit_price', 0)
        currency = item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))

        story.append(Paragraph(f'Гарантия: {warranty}', normal))
        story.append(Paragraph(f'Сроки изготовления: {production_time}.', normal))
        story.append(Paragraph(f'Упаковка: {packaging}.', normal))
        story.append(Paragraph(f'Условия оплаты: {payment_terms}.', normal))
        story.append(Paragraph(f'Цена с НДС с доставкой {item.get("delivery", "до завода покупателя")} за 1 шт.: {unit_price:,.0f} {currency}.', bold_s))
        story.append(Spacer(1, 0.5*cm))

    # Итоговая таблица
    if len(items) > 1:
        story.append(Paragraph('Итоговая стоимость', bold_s))
        td = [['Оборудование', 'Кол-во', 'Цена/шт', 'Сумма']]
        total = 0
        for item in items:
            qty = item.get('quantity', 1)
            price = item.get('unit_price', 0)
            curr = item.get('currency', '')
            sub = price * qty
            total += sub
            td.append([item.get('name', item.get('model', '')), str(qty),
                       f"{price:,.0f} {curr}", f"{sub:,.0f} {curr}"])
        curr = items[0].get('currency', '') if items else ''
        td.append(['ИТОГО', '', '', f"{total:,.0f} {curr}"])
        t = Table(td, colWidths=[8*cm, 2*cm, 4*cm, 4*cm])
        t.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), font_bold),
            ('FONTNAME', (0, -1), (-1, -1), font_bold),
            ('FONTNAME', (0, 1), (-1, -2), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.5*cm))

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
