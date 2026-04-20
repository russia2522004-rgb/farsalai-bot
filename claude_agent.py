import os
import json
import re
from anthropic import Anthropic
from database import search_equipment, get_all_equipment, get_equipment_by_model

client = Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))

SYSTEM_PROMPT = """Ты — ИИ-агент компании ООО «Фарсал», помогающий менеджерам создавать коммерческие предложения.

Компания продаёт промышленное оборудование — насосы и компрессоры производства Китай.

Твоя задача — вести диалог с менеджером, собрать все необходимые данные для КП и вернуть их в структурированном виде.

ОБЯЗАТЕЛЬНЫЕ данные для КП:
- Оборудование (одна или несколько позиций): название, модель
- Количество каждой позиции
- Цена за единицу (и валюта)
- Клиент (название организации)

НЕОБЯЗАТЕЛЬНЫЕ (если не указаны — берём из библиотеки или стандартные):
- Условия оплаты
- Срок поставки/изготовления
- Способ доставки
- Гарантия
- Контактное лицо клиента

ПРАВИЛА:
1. Задавай уточняющие вопросы по одному — не анкетой
2. Если менеджер назвал модель — используй её
3. Если данных достаточно — верни JSON
4. Веди деловой но дружелюбный тон
5. Понимай разговорный стиль ("сделай подешевле" = изменить цену)
6. "готово", "сохрани", "отправляй" — сигнал финализировать КП
7. Поддерживай несколько позиций в одном КП

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

Если для позиции не указаны условия — оставь null (возьмём из библиотеки).
Если данных не хватает — веди диалог без JSON.
"""


def parse_json_from_text(text: str) -> dict | None:
    json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except json.JSONDecodeError:
            pass
    try:
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
    except Exception:
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
    text, kp_data = parse_claude_response(assistant_message)
    return assistant_message, kp_data


def process_edit(conversation_history: list, edit_message: str) -> tuple[str, dict | None]:
    edit_prompt = f"""Менеджер хочет внести правку в КП: {edit_message}
Обнови данные КП и верни обновлённый JSON. Если правка понятна — сразу верни JSON."""
    return chat_with_claude(conversation_history, edit_prompt)


def extract_blocks_from_doc(doc_text: str, doc_path: str = None) -> tuple[dict, list]:
    """
    Извлекает данные оборудования и блоки из КП документа.
    Возвращает (карточка оборудования, список блоков)
    """
    prompt = f"""Из текста коммерческого предложения извлеки:
1. Данные об оборудовании
2. Все разделы документа как отдельные блоки

Текст:
{doc_text}

Верни JSON:
```json
{{
  "equipment": {{
    "name": "полное название",
    "model": "модель/артикул",
    "description": "краткое описание или null",
    "warranty": "гарантийные условия",
    "production_time": "срок изготовления/поставки",
    "packaging": "упаковка",
    "delivery": "место доставки",
    "payment_terms": "условия оплаты",
    "base_price": числовое значение или null,
    "currency": "валюта"
  }},
  "blocks": [
    {{
      "type": "specs",
      "title": "Технические характеристики",
      "content": "текстовое содержимое раздела"
    }},
    {{
      "type": "construction",
      "title": "Конструктивное исполнение",
      "content": "текстовое содержимое"
    }}
  ]
}}
```

Типы блоков: specs, construction, dimensions, flow, warranty_block, delivery_info, options, purpose, design, supply_set, chart, other
Если раздела нет — не включай его в список.
"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text
    data = parse_json_from_text(text)

    if not data:
        return {}, []

    equipment = data.get('equipment', {})
    blocks = data.get('blocks', [])
    return equipment, blocks


def extract_all_equipment_from_doc(doc_text: str) -> list:
    """Извлекает все позиции оборудования из КП (для обратной совместимости)"""
    equipment, blocks = extract_blocks_from_doc(doc_text)
    if equipment:
        equipment['blocks'] = blocks
        return [equipment]
    return []


def extract_equipment_from_doc(doc_text: str) -> dict:
    """Извлекает первую позицию (для обратной совместимости)"""
    items = extract_all_equipment_from_doc(doc_text)
    return items[0] if items else {}


def compare_equipment(existing: dict, new_data: dict) -> dict:
    differences = {
        'specs_changed': [],
        'price_changed': None,
        'conditions_changed': [],
        'has_changes': False
    }

    # Сравниваем цену
    old_price = existing.get('base_price')
    new_price = new_data.get('base_price')
    if old_price and new_price and abs(float(old_price) - float(new_price)) > 0.01:
        differences['price_changed'] = {
            'old': old_price, 'new': new_price,
            'currency': new_data.get('currency', existing.get('currency', ''))
        }
        differences['has_changes'] = True

    # Сравниваем условия
    condition_fields = ['warranty', 'production_time', 'packaging', 'delivery', 'payment_terms']
    for field in condition_fields:
        old_val = existing.get(field, '')
        new_val = new_data.get(field, '')
        if old_val and new_val and str(old_val).strip() != str(new_val).strip():
            differences['conditions_changed'].append({
                'field': field, 'old': old_val, 'new': new_val
            })
            differences['has_changes'] = True

    # Сравниваем характеристики
    try:
        old_specs = json.loads(existing.get('specs', '[]')) if isinstance(existing.get('specs'), str) else (existing.get('specs') or [])
        new_specs = new_data.get('specs', [])
        old_dict = {s['name']: s['value'] for s in old_specs}
        new_dict = {s['name']: s['value'] for s in new_specs}
        for name, new_val in new_dict.items():
            old_val = old_dict.get(name)
            if old_val is None:
                differences['specs_changed'].append({'name': name, 'old': None, 'new': new_val, 'type': 'added'})
                differences['has_changes'] = True
            elif str(old_val) != str(new_val):
                differences['specs_changed'].append({'name': name, 'old': old_val, 'new': new_val, 'type': 'changed'})
                differences['has_changes'] = True
        for name in old_dict:
            if name not in new_dict:
                differences['specs_changed'].append({'name': name, 'old': old_dict[name], 'new': None, 'type': 'removed'})
                differences['has_changes'] = True
    except Exception as e:
        print(f"Ошибка сравнения характеристик: {e}")

    return differences


def resolve_equipment_conflict(existing: dict, new_data: dict, differences: dict, manager_instruction: str) -> dict:
    prompt = f"""Менеджер решает конфликт данных оборудования.

Существующие данные: {json.dumps(existing, ensure_ascii=False, indent=2)}
Новые данные: {json.dumps(new_data, ensure_ascii=False, indent=2)}
Различия: {json.dumps(differences, ensure_ascii=False, indent=2)}
Инструкция менеджера: {manager_instruction}

Составь итоговые данные оборудования на основе инструкции.
Верни ТОЛЬКО JSON с итоговыми данными."""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    result = parse_json_from_text(response.content[0].text)
    return result or new_data
