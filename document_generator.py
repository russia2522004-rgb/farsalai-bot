import os
import json
import subprocess
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from database import get_equipment_by_model

OUTPUT_DIR = 'output'
TEMPLATE_DIR = 'template'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)


def _set_cell_bg(cell, color: str):
    """Устанавливает цвет фона ячейки"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def _add_paragraph_with_style(doc, text, bold=False, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=6):
    """Добавляет параграф с настройками"""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return p


def generate_kp_document(kp_data: dict, manager_name: str) -> tuple[str, str]:
    """
    Генерирует КП в формате Word и PDF.
    Возвращает (путь к docx, путь к pdf)
    """
    kp_number = kp_data.get('kp_number', 'KP-001')
    kp_date = kp_data.get('kp_date', datetime.now().strftime('%d.%m.%Y'))

    # Проверяем есть ли шаблон
    template_path = os.path.join(TEMPLATE_DIR, 'kp_template.docx')
    if os.path.exists(template_path):
        doc = Document(template_path)
        # Очищаем содержимое после шапки (оставляем первые параграфы с реквизитами)
        # Это упрощённый вариант — в будущем доработать под конкретный шаблон
    else:
        doc = _create_document_from_scratch(kp_data, kp_number, kp_date, manager_name)

    # Сохраняем Word
    docx_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.docx')
    doc.save(docx_path)

    # Конвертируем в PDF
    pdf_path = _convert_to_pdf(docx_path, kp_number)

    return docx_path, pdf_path


def _create_document_from_scratch(kp_data: dict, kp_number: str, kp_date: str, manager_name: str) -> Document:
    """Создаёт документ КП с нуля"""
    doc = Document()

    # Настройки страницы (A4)
    from docx.shared import Cm
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # ── Шапка ──
    header_table = doc.add_table(rows=1, cols=2)
    header_table.style = 'Table Grid'
    left_cell = header_table.cell(0, 0)
    right_cell = header_table.cell(0, 1)

    # Реквизиты компании (левая колонка)
    left_cell.text = 'ООО «Фарсал»\nИНН/КПП: XXXXXXXXXX/XXXXXXXXX\nАдрес: г. Таганрог\nТел: +7 (XXX) XXX-XX-XX'
    for para in left_cell.paragraphs:
        para.runs[0].font.size = Pt(9) if para.runs else None

    # Дата и номер (правая колонка)
    right_para = right_cell.paragraphs[0]
    right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right_run = right_para.add_run(f'от {kp_date}  №  {kp_number}')
    right_run.font.size = Pt(10)
    right_run.bold = True

    doc.add_paragraph()

    # ── Заголовок ──
    title = doc.add_paragraph('КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True
    title.runs[0].font.size = Pt(14)

    subtitle = doc.add_paragraph('ООО «Фарсал» предлагает к поставке')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].bold = True
    subtitle.runs[0].font.size = Pt(12)

    doc.add_paragraph()

    # ── Позиции оборудования ──
    items = kp_data.get('items', [])
    for i, item in enumerate(items, 1):
        model = item.get('model', '')
        eq = get_equipment_by_model(model)

        # Название оборудования
        eq_name = eq['name'] if eq else item.get('name', model)
        name_para = doc.add_paragraph(f'{eq_name}')
        name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        name_para.runs[0].bold = True
        name_para.runs[0].font.size = Pt(13)

        doc.add_paragraph()

        # Фото (если есть)
        if eq and eq.get('photo_path') and os.path.exists(eq['photo_path']):
            try:
                photo_para = doc.add_paragraph()
                photo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = photo_para.add_run()
                run.add_picture(eq['photo_path'], width=Inches(4))
                note = doc.add_paragraph('* Фото для справки, реальные фотографии вашего оборудования будут предоставлены после завершения производства.')
                note.runs[0].font.size = Pt(9)
                note.runs[0].italic = True
            except Exception:
                pass

        doc.add_paragraph()

        # Описание (если есть)
        if eq and eq.get('description'):
            desc_para = doc.add_paragraph(eq['description'])
            desc_para.runs[0].font.size = Pt(10)
            doc.add_paragraph()

        # Технические характеристики
        specs_title = doc.add_paragraph('Технические характеристики')
        specs_title.runs[0].bold = True
        specs_title.runs[0].font.size = Pt(11)

        specs = []
        if eq and eq.get('specs'):
            try:
                specs = json.loads(eq['specs']) if isinstance(eq['specs'], str) else eq['specs']
            except Exception:
                specs = []

        if specs:
            specs_table = doc.add_table(rows=1, cols=2)
            specs_table.style = 'Table Grid'
            # Заголовок
            hdr = specs_table.rows[0].cells
            hdr[0].text = 'Характеристика'
            hdr[1].text = 'Значение'
            for cell in hdr:
                cell.paragraphs[0].runs[0].bold = True
                _set_cell_bg(cell, 'D9E1F2')

            for spec in specs:
                row = specs_table.add_row().cells
                row[0].text = spec.get('name', '')
                val_run = row[1].paragraphs[0].add_run(str(spec.get('value', '')))
                val_run.bold = True

        doc.add_paragraph()

        # Конструктивное исполнение (если есть)
        if eq and eq.get('construction'):
            const_title = doc.add_paragraph('Конструктивное исполнение')
            const_title.runs[0].bold = True
            doc.add_paragraph(eq['construction'])
            doc.add_paragraph()

    # ── Блок условий ──
    doc.add_paragraph()
    conditions = [
        ('Гарантийный срок:', '1 год. Изнашиваемые детали гарантийному обслуживанию не подлежат.'),
        ('Сроки изготовления:', kp_data.get('production_time', '25–30 дней')),
        ('Упаковка:', kp_data.get('packaging', 'экспортная деревянная тара (ящик)')),
        ('Условия оплаты:', kp_data.get('payment_terms', '50% – предоплата, 50% – по факту поставки')),
    ]

    # Цены по позициям
    if len(items) > 1:
        price_title = doc.add_paragraph('Стоимость:')
        price_title.runs[0].bold = True
        for item in items:
            total = item['quantity'] * item['unit_price']
            p = doc.add_paragraph(
                f"• {item.get('name', item.get('model'))}: "
                f"{item['quantity']} шт. × {item['unit_price']:,.0f} = "
                f"{total:,.0f} {item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))}"
            )
            p.runs[0].font.size = Pt(10)

        total_p = doc.add_paragraph(
            f"Итого: {kp_data.get('total_price', 0):,.0f} {kp_data.get('currency', 'ЮАНЕЙ')} с НДС с доставкой до завода покупателя"
        )
        total_p.runs[0].bold = True
    else:
        item = items[0] if items else {}
        price_p = doc.add_paragraph(
            f"Цена с НДС с доставкой до завода покупателя за 1 штуку: "
            f"{item.get('unit_price', 0):,.0f} {item.get('currency', kp_data.get('currency', 'ЮАНЕЙ'))}"
        )
        price_p.runs[0].bold = True

    doc.add_paragraph()

    for label, value in conditions:
        p = doc.add_paragraph()
        run_label = p.add_run(f'{label} ')
        run_label.bold = True
        run_label.font.size = Pt(10)
        run_value = p.add_run(value)
        run_value.font.size = Pt(10)

    doc.add_paragraph()
    doc.add_paragraph()

    # ── Подпись ──
    sign_para = doc.add_paragraph('С уважением,\nдиректор ООО «Фарсал»,')
    sign_para.runs[0].font.size = Pt(10)

    sign_line = doc.add_paragraph('МП     _______________       А. Ю. Лавришко')
    sign_line.runs[0].font.size = Pt(10)

    return doc


def _convert_to_pdf(docx_path: str, kp_number: str) -> str:
    """Конвертирует Word в PDF через LibreOffice"""
    pdf_path = os.path.join(OUTPUT_DIR, f'КП_{kp_number}.pdf')

    try:
        # Пробуем LibreOffice
        subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf',
            '--outdir', OUTPUT_DIR, docx_path
        ], check=True, capture_output=True)

        # LibreOffice сохраняет с тем же именем
        generated = docx_path.replace('.docx', '.pdf')
        if generated != pdf_path and os.path.exists(generated):
            os.rename(generated, pdf_path)

    except (subprocess.CalledProcessError, FileNotFoundError):
        try:
            # Пробуем docx2pdf
            from docx2pdf import convert
            convert(docx_path, pdf_path)
        except Exception as e:
            print(f"PDF конвертация недоступна: {e}")
            # Возвращаем путь к docx как fallback
            return docx_path

    return pdf_path


def cleanup_temp_files(docx_path: str, pdf_path: str):
    """Удаляет временные файлы после загрузки на Яндекс Диск"""
    for path in [docx_path, pdf_path]:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass