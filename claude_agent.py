import os
import json
import re
import copy
from anthropic import Anthropic
from database import search_equipment, get_all_equipment, get_equipment_by_model

client = Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

SECTION_HEADERS = {
    'технические характеристики': 'specs',
    'гарантия': 'warranty_block',
    'конструктивное исполнение': 'construction',
    'расход рабочей жидкости': 'flow',
    'габаритные размеры': 'dimensions',
    'габаритный чертеж': 'drawing',
    'назначение оборудования': 'purpose',
    'назначение': 'purpose',
    'конструкция насоса': 'design',
    'комплект поставки': 'supply',
    'дополнительные опции': 'options',
    'дополнительная опция': 'options',
    'график рабочих характеристик': 'chart',
    'наименование': 'naming',
}

# Тексты которые НЕ являются заголовками разделов даже если совпадают
NOT_SECTION_HEADERS = [
    'гарантия не распространяется',
    'гарантийный',
    'гарантия на',
]

CONDITIONS_KEYWORDS = [
    'сроки изготовления', 'сроки поставки', 'срок изготовления', 'срок поставки',
    'упаковка', 'условия оплаты', 'цена с ндс', 'цена за',
    'с уважением', 'директор', 'коммерческое предложение',
    'ооо «фарсал» предлагает',
]

SYSTEM_PROMPT = """Ты — ИИ-агент компании ООО «Фарсал», помогающий менеджерам создавать коммерческие предложения.

Компания продаёт промышленное оборудование — насосы и компрессоры производства Китай.

ОБЯЗАТЕЛЬНЫЕ данные для КП:
- Оборудование (одна или несколько позиций): название, модель
- Количество каждой позиции
- Цена за единицу (и валюта)
- Клиент (название организации)

НЕОБЯЗАТЕЛЬНЫЕ (если не указаны — берём из библиотеки):
- Условия оплаты
- Срок поставки/изготовления
- Способ доставки
- Гарантия

ПРАВИЛА:
1. Задавай уточняющие вопросы по одному
2. Если данных достаточно — верни JSON
3. Понимай разговорный стиль
4. "готово", "сохрани", "отправляй" — финализировать КП
5. Поддерживай несколько позиций

Когда все данные собраны, верни ТОЛЬКО JSON:
```json
{
  "ready": true,
  "client": "название клиента",
  "contact_person": null,
  "items": [
    {
      "model": "модель",
      "name": "полное название",
      "quantity": 1,
      "unit_price": 80000,
      "currency": "ЮАНЕЙ",
      "payment_terms": null,
      "production_time": null,
      "delivery": null,
      "warranty": null
    }
  ],
  "total_price": 80000,
  "currency": "ЮАНЕЙ",
  "notes": null
}
```
null = возьмём из библиотеки.
"""


def parse_json_from_text(text: str) -> dict | None:
    json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except:
            pass
    try:
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    return None


def parse_claude_response(text: str) -> tuple[str, dict | None]:
    data = parse_json_from_text(text)
    if data and data.get('ready'):
        return text, data
    return text, None


def chat_with_claude(conversation_history: list, user_message: str) -> tuple[str, dict | None]:
    conversation_history.append({"role": "user", "content": user_message})
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=conversation_history
    )
    assistant_message = response.content[0].text
    conversation_history.append({"role": "assistant", "content": assistant_message})
    return parse_claude_response(assistant_message)


def process_edit(conversation_history: list, edit_message: str) -> tuple[str, dict | None]:
    edit_prompt = f"Менеджер хочет внести правку: {edit_message}\nОбнови JSON КП."
    return chat_with_claude(conversation_history, edit_prompt)


# ─── Извлечение блоков из Word файла ─────────────────────────────────────────

def _get_elem_text(elem) -> str:
    return ''.join(t.text or '' for t in elem.iter(f'{{{NS}}}t')).strip()


def _is_section_header(elem) -> tuple:
    if elem.tag.split('}')[-1] != 'p':
        return None, None
    style_elems = list(elem.iter(f'{{{NS}}}pStyle'))
    style = style_elems[0].get(f'{{{NS}}}val', '') if style_elems else ''
    text = _get_elem_text(elem).lower()
    if not text:
        return None, None

    # Проверяем исключения — тексты которые НЕ являются заголовками
    for not_header in NOT_SECTION_HEADERS:
        if text.startswith(not_header):
            return None, None

    is_header_style = bool(re.search(r'[Hh]eading|ХХХ|^\d+$', style)) or style == '1'
    if is_header_style:
        for keyword, block_type in SECTION_HEADERS.items():
            if keyword in text:
                return block_type, _get_elem_text(elem)
    # Точное совпадение по тексту (без стиля)
    for keyword, block_type in SECTION_HEADERS.items():
        if text == keyword:
            return block_type, _get_elem_text(elem)
    return None, None


def _is_conditions_element(elem) -> bool:
    text = _get_elem_text(elem).lower()
    return any(kw in text for kw in CONDITIONS_KEYWORDS)


def extract_numbering_xml(doc_path: str) -> str:
    """Извлекает numbering.xml из Word файла"""
    try:
        import zipfile
        with zipfile.ZipFile(doc_path, 'r') as z:
            if 'word/numbering.xml' in z.namelist():
                return z.read('word/numbering.xml').decode('utf-8')
    except Exception as e:
        print(f"Ошибка извлечения numbering.xml: {e}")
    return ''


def extract_blocks_from_docx(doc_path: str) -> list:
    """
    Извлекает блоки из Word файла как XML фрагменты.
    Сохраняет форматирование — таблицы, шрифты, стили.
    Извлекает картинки внутри блоков и сохраняет локально.
    """
    try:
        from docx import Document
        from lxml import etree as et
        import zipfile
    except ImportError:
        return []

    doc = Document(doc_path)
    body = doc.element.body
    elements = list(body)

    # Читаем все медиафайлы из docx архива
    media_files = {}
    try:
        with zipfile.ZipFile(doc_path, 'r') as z:
            for name in z.namelist():
                if name.startswith('word/media/'):
                    media_files[name] = z.read(name)
    except Exception as e:
        print(f"Ошибка чтения медиафайлов: {e}")

    # Читаем relationships для сопоставления rId → файл
    rels = {}
    try:
        with zipfile.ZipFile(doc_path, 'r') as z:
            if 'word/_rels/document.xml.rels' in z.namelist():
                rel_xml = z.read('word/_rels/document.xml.rels')
                rel_root = et.fromstring(rel_xml)
                for rel in rel_root:
                    rid = rel.get('Id', '')
                    target = rel.get('Target', '')
                    rels[rid] = target
    except Exception as e:
        print(f"Ошибка чтения relationships: {e}")

    def get_block_images_base64(block_elements) -> list:
        """Извлекает картинки из элементов блока как base64"""
        import base64
        images_b64 = []
        DRAW_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        BLIP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

        for elem in block_elements:
            for inline in list(elem.iter(f'{{{DRAW_NS}}}inline')) + list(elem.iter(f'{{{DRAW_NS}}}anchor')):
                blips = list(inline.iter(f'{{{BLIP_NS}}}blip'))
                for blip in blips:
                    rid = blip.get(f'{{{REL_NS}}}embed', '')
                    if rid and rid in rels:
                        target = rels[rid]
                        media_key = f'word/{target}' if not target.startswith('word/') else target
                        if media_key in media_files:
                            b64 = base64.b64encode(media_files[media_key]).decode('utf-8')
                            # Определяем тип
                            ext = os.path.splitext(target)[1].lower().lstrip('.')
                            mime = {'jpg': 'jpeg', 'jpeg': 'jpeg', 'png': 'png', 'gif': 'gif'}.get(ext, 'png')
                            images_b64.append(f'data:image/{mime};base64,{b64}')
        return images_b64

    blocks = []
    current_block = None
    current_elements = []

    for elem in elements:
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'sectPr':
            continue

        block_type, block_title = _is_section_header(elem)

        if block_type:
            # Сохраняем предыдущий блок
            if current_block and current_elements:
                wrapper = et.Element('block')
                for e in current_elements:
                    wrapper.append(copy.deepcopy(e))
                images_b64 = get_block_images_base64(current_elements)
                blocks.append({
                    'type': current_block['type'],
                    'title': current_block['title'],
                    'xml': et.tostring(wrapper, encoding='unicode'),
                    'images': [], 'images_base64': images_b64  # локальные пути к картинкам
                })
            current_block = {'type': block_type, 'title': block_title}
            current_elements = []

        elif current_block:
            if _is_conditions_element(elem):
                if current_elements:
                    wrapper = et.Element('block')
                    for e in current_elements:
                        wrapper.append(copy.deepcopy(e))
                    images_b64 = get_block_images_base64(current_elements)
                    blocks.append({
                        'type': current_block['type'],
                        'title': current_block['title'],
                        'xml': et.tostring(wrapper, encoding='unicode'),
                        'images': [], 'images_base64': images_b64
                    })
                current_block = None
                current_elements = []
            else:
                text = _get_elem_text(elem)
                if text or tag == 'tbl':
                    current_elements.append(elem)
                elif tag == 'p':
                    # Пустой параграф с картинкой
                    DRAW_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                    if list(elem.iter(f'{{{DRAW_NS}}}inline')) or list(elem.iter(f'{{{DRAW_NS}}}anchor')):
                        current_elements.append(elem)

    # Последний блок
    if current_block and current_elements:
        wrapper = et.Element('block')
        for e in current_elements:
            wrapper.append(copy.deepcopy(e))
        images_b64 = get_block_images_base64(current_elements)
        blocks.append({
            'type': current_block['type'],
            'title': current_block['title'],
            'xml': et.tostring(wrapper, encoding='unicode'),
            'images': [], 'images_base64': images_b64
        })

    return blocks

    # Последний блок
    if current_block and current_elements:
        from lxml import etree as et
        wrapper = et.Element('block')
        for e in current_elements:
            wrapper.append(copy.deepcopy(e))
        blocks.append({
            'type': current_block['type'],
            'title': current_block['title'],
            'xml': et.tostring(wrapper, encoding='unicode'),
            'images': []
        })

    return blocks


def extract_equipment_info_from_text(doc_text: str) -> dict:
    """Извлекает данные об оборудовании через Claude"""
    prompt = f"""Из текста КП извлеки данные оборудования.

Текст:
{doc_text[:3000]}

Верни JSON:
```json
{{
  "name": "полное название",
  "model": "модель",
  "warranty": "гарантия",
  "production_time": "срок",
  "packaging": "упаковка",
  "delivery": "место доставки",
  "payment_terms": "условия оплаты",
  "base_price": числовое значение или null,
  "currency": "валюта"
}}
```"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )
    return parse_json_from_text(response.content[0].text) or {}


def extract_all_equipment_from_doc(doc_text: str, doc_path: str = None) -> list:
    """Извлекает оборудование из документа. doc_path нужен для извлечения XML блоков."""
    equipment = extract_equipment_info_from_text(doc_text)
    if not equipment:
        return []

    blocks = []
    numbering_xml = ''
    if doc_path and os.path.exists(doc_path):
        blocks = extract_blocks_from_docx(doc_path)
        numbering_xml = extract_numbering_xml(doc_path)

    equipment['blocks'] = blocks
    equipment['numbering_xml'] = numbering_xml
    return [equipment]


def extract_equipment_from_doc(doc_text: str) -> dict:
    items = extract_all_equipment_from_doc(doc_text)
    return items[0] if items else {}


def compare_equipment(existing: dict, new_data: dict) -> dict:
    """
    Сравнивает существующее оборудование с новыми данными.
    Логика: дополнять новыми данными, при конфликте — спрашивать.
    """
    differences = {
        'fields_to_add': {},       # поля которых нет в БД — добавить автоматически
        'fields_conflict': [],     # поля которые отличаются — спросить менеджера
        'specs_to_add': [],        # характеристики которых нет в БД
        'specs_conflict': [],      # характеристики которые отличаются
        'has_conflicts': False,    # есть ли конфликты требующие решения
        'has_additions': False,    # есть ли что добавить автоматически
    }

    # Сравниваем скалярные поля
    for field in ['base_price', 'currency', 'warranty', 'production_time', 'packaging', 'delivery', 'payment_terms']:
        old_val = existing.get(field)
        new_val = new_data.get(field)

        if not new_val:
            continue  # в новом файле нет — игнорируем

        if not old_val:
            # В БД пусто — добавляем автоматически
            differences['fields_to_add'][field] = new_val
            differences['has_additions'] = True
        elif field == 'base_price':
            if abs(float(old_val) - float(new_val)) > 0.01:
                differences['fields_conflict'].append({
                    'field': field,
                    'old': old_val, 'new': new_val,
                    'currency': new_data.get('currency', existing.get('currency', ''))
                })
                differences['has_conflicts'] = True
        else:
            if str(old_val).strip() != str(new_val).strip():
                differences['fields_conflict'].append({
                    'field': field,
                    'old': old_val, 'new': new_val
                })
                differences['has_conflicts'] = True

    # Сравниваем specs (список характеристик)
    old_specs_raw = existing.get('specs', '[]')
    new_specs_raw = new_data.get('specs', [])

    try:
        old_specs = json.loads(old_specs_raw) if isinstance(old_specs_raw, str) else (old_specs_raw or [])
    except Exception:
        old_specs = []

    new_specs = new_specs_raw if isinstance(new_specs_raw, list) else []

    old_map = {s['name'].strip().lower(): s['value'] for s in old_specs if isinstance(s, dict) and 'name' in s}

    for spec in new_specs:
        if not isinstance(spec, dict) or 'name' not in spec:
            continue
        key = spec['name'].strip().lower()
        new_val = spec.get('value', '')
        if key not in old_map:
            # Новая характеристика — добавить
            differences['specs_to_add'].append(spec)
            differences['has_additions'] = True
        elif str(old_map[key]).strip() != str(new_val).strip():
            # Значение отличается — конфликт
            differences['specs_conflict'].append({
                'name': spec['name'],
                'old': old_map[key],
                'new': new_val
            })
            differences['has_conflicts'] = True

    return differences


def resolve_equipment_conflict(existing: dict, new_data: dict, differences: dict, manager_instruction: str) -> dict:
    prompt = f"""Менеджер решает конфликт данных оборудования.
Существующие: {json.dumps(existing, ensure_ascii=False)}
Новые: {json.dumps(new_data, ensure_ascii=False)}
Различия: {json.dumps(differences, ensure_ascii=False)}
Инструкция: {manager_instruction}
Верни ТОЛЬКО JSON с итоговыми данными."""
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    return parse_json_from_text(response.content[0].text) or new_data
