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
    'назначение': 'purpose',
    'конструкция насоса': 'design',
    'комплект поставки': 'supply',
    'дополнительные опции': 'options',
    'график рабочих характеристик': 'chart',
    'наименование': 'naming',
    'назначение оборудования': 'purpose',
}

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
    is_header_style = bool(re.search(r'[Hh]eading|ХХХ|^\d+$', style))
    if is_header_style:
        for keyword, block_type in SECTION_HEADERS.items():
            if keyword in text:
                return block_type, _get_elem_text(elem)
    for keyword, block_type in SECTION_HEADERS.items():
        if text == keyword or text.startswith(keyword):
            return block_type, _get_elem_text(elem)
    return None, None


def _is_conditions_element(elem) -> bool:
    text = _get_elem_text(elem).lower()
    return any(kw in text for kw in CONDITIONS_KEYWORDS)


def extract_blocks_from_docx(doc_path: str) -> list:
    """
    Извлекает блоки из Word файла как XML фрагменты.
    Сохраняет всё форматирование — таблицы, шрифты, стили.
    """
    try:
        from docx import Document
        from lxml import etree
    except ImportError:
        return []

    doc = Document(doc_path)
    body = doc.element.body
    elements = list(body)

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
            current_block = {'type': block_type, 'title': block_title}
            current_elements = []

        elif current_block:
            if _is_conditions_element(elem):
                # Конец блоков — начались условия
                if current_elements:
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
                current_block = None
                current_elements = []
            else:
                text = _get_elem_text(elem)
                if text or tag == 'tbl':
                    current_elements.append(elem)

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
    # Данные оборудования через Claude
    equipment = extract_equipment_info_from_text(doc_text)
    if not equipment:
        return []

    # Блоки через XML если есть путь к файлу
    blocks = []
    if doc_path and os.path.exists(doc_path):
        blocks = extract_blocks_from_docx(doc_path)

    equipment['blocks'] = blocks
    return [equipment]


def extract_equipment_from_doc(doc_text: str) -> dict:
    items = extract_all_equipment_from_doc(doc_text)
    return items[0] if items else {}


def compare_equipment(existing: dict, new_data: dict) -> dict:
    differences = {
        'specs_changed': [],
        'price_changed': None,
        'conditions_changed': [],
        'has_changes': False
    }

    old_price = existing.get('base_price')
    new_price = new_data.get('base_price')
    if old_price and new_price and abs(float(old_price) - float(new_price)) > 0.01:
        differences['price_changed'] = {
            'old': old_price, 'new': new_price,
            'currency': new_data.get('currency', existing.get('currency', ''))
        }
        differences['has_changes'] = True

    for field in ['warranty', 'production_time', 'packaging', 'delivery', 'payment_terms']:
        old_val = existing.get(field, '')
        new_val = new_data.get(field, '')
        if old_val and new_val and str(old_val).strip() != str(new_val).strip():
            differences['conditions_changed'].append(
                {'field': field, 'old': old_val, 'new': new_val}
            )
            differences['has_changes'] = True

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
