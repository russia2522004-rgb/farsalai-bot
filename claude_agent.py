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
- Условия оплаты
- Срок поставки/изготовления

НЕОБЯЗАТЕЛЬНЫЕ данные:
- Контактное лицо клиента
- Способ доставки
- Особые условия

ПРАВИЛА:
1. Задавай уточняющие вопросы по одному — не анкетой
2. Если менеджер назвал модель оборудования — используй её
3. Если данных достаточно — верни JSON с данными КП
4. Веди деловой но дружелюбный тон
5. Понимай сокращения и разговорный стиль ("сделай подешевле" = изменить цену)
6. Если менеджер говорит "готово", "сохрани", "отправляй" — это сигнал финализировать КП
7. Поддерживай несколько позиций в одном КП

Когда все данные собраны, верни ТОЛЬКО JSON в таком формате (без лишнего текста):
```json
{
  "ready": true,
  "client": "название клиента",
  "contact_person": "контактное лицо или null",
  "items": [
    {
      "model": "модель оборудования",
      "name": "полное название",
      "quantity": 1,
      "unit_price": 80000,
      "currency": "ЮАНЕЙ"
    }
  ],
  "payment_terms": "условия оплаты",
  "delivery": "способ доставки или null",
  "production_time": "срок изготовления",
  "total_price": 80000,
  "currency": "ЮАНЕЙ",
  "notes": "дополнительные примечания или null"
}
```

Если данных ещё не хватает — просто веди диалог, без JSON.
"""


def parse_json_from_text(text: str) -> dict | None:
    """Извлекает JSON из текста"""
    # Ищем в блоке ```json
    json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except json.JSONDecodeError:
            pass

    # Ищем любой JSON объект
    try:
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
    except Exception:
        pass

    return None


def parse_claude_response(text: str) -> tuple[str, dict | None]:
    """Разбирает ответ Claude — текст и возможный JSON"""
    data = parse_json_from_text(text)
    if data and data.get('ready'):
        return text, data
    return text, None


def chat_with_claude(conversation_history: list, user_message: str) -> tuple[str, dict | None]:
    """Ведёт диалог с Claude. Возвращает (текст ответа, данные КП если готовы)"""
    conversation_history.append({
        "role": "user",
        "content": user_message
    })

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=conversation_history
    )

    assistant_message = response.content[0].text
    conversation_history.append({
        "role": "assistant",
        "content": assistant_message
    })

    text, kp_data = parse_claude_response(assistant_message)
    return assistant_message, kp_data


def process_edit(conversation_history: list, edit_message: str) -> tuple[str, dict | None]:
    """Обрабатывает правки к готовому КП"""
    edit_prompt = f"""Менеджер хочет внести правку в КП: {edit_message}

Обнови данные КП и верни обновлённый JSON. Если правка понятна — сразу верни JSON без лишних вопросов."""
    return chat_with_claude(conversation_history, edit_prompt)


def extract_all_equipment_from_doc(doc_text: str) -> list:
    """
    Извлекает ВСЕ позиции оборудования из текста КП.
    Возвращает список карточек оборудования.
    """
    prompt = f"""Из следующего текста коммерческого предложения извлеки данные обо ВСЕХ позициях оборудования.

Текст:
{doc_text}

Верни ТОЛЬКО JSON массив без лишнего текста. Каждый элемент — отдельная позиция оборудования:
```json
[
  {{
    "name": "полное название оборудования",
    "model": "модель/артикул",
    "description": "краткое описание назначения или null",
    "specs": [{{"name": "характеристика", "value": "значение"}}],
    "construction": "описание конструктивного исполнения или null",
    "production_time": "срок изготовления",
    "packaging": "упаковка",
    "base_price": числовое значение цены,
    "currency": "валюта"
  }}
]
```

Если оборудование не найдено — верни пустой массив [].
"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text

    # Ищем массив JSON
    array_match = re.search(r'```json\s*(\[.*?\])\s*```', text, re.DOTALL)
    if array_match:
        try:
            return json.loads(array_match.group(1))
        except Exception:
            pass

    # Пробуем найти массив без тегов
    try:
        array_match = re.search(r'\[.*\]', text, re.DOTALL)
        if array_match:
            return json.loads(array_match.group())
    except Exception:
        pass

    return []


def extract_equipment_from_doc(doc_text: str) -> dict:
    """Извлекает первую позицию оборудования (для обратной совместимости)"""
    items = extract_all_equipment_from_doc(doc_text)
    return items[0] if items else {}


def compare_equipment(existing: dict, new_data: dict) -> dict:
    """
    Сравнивает существующее оборудование с новыми данными.
    Возвращает словарь с различиями.
    """
    differences = {
        'specs_changed': [],
        'price_changed': None,
        'has_changes': False
    }

    # Сравниваем цену
    old_price = existing.get('base_price')
    new_price = new_data.get('base_price')
    if old_price and new_price and abs(float(old_price) - float(new_price)) > 0.01:
        differences['price_changed'] = {
            'old': old_price,
            'new': new_price,
            'currency': new_data.get('currency', existing.get('currency', ''))
        }
        differences['has_changes'] = True

    # Сравниваем характеристики
    try:
        old_specs = json.loads(existing.get('specs', '[]')) if isinstance(existing.get('specs'), str) else (existing.get('specs') or [])
        new_specs = new_data.get('specs', [])

        old_specs_dict = {s['name']: s['value'] for s in old_specs}
        new_specs_dict = {s['name']: s['value'] for s in new_specs}

        # Изменённые или новые характеристики
        for name, new_val in new_specs_dict.items():
            old_val = old_specs_dict.get(name)
            if old_val is None:
                differences['specs_changed'].append({
                    'name': name,
                    'old': None,
                    'new': new_val,
                    'type': 'added'
                })
                differences['has_changes'] = True
            elif str(old_val) != str(new_val):
                differences['specs_changed'].append({
                    'name': name,
                    'old': old_val,
                    'new': new_val,
                    'type': 'changed'
                })
                differences['has_changes'] = True

        # Удалённые характеристики
        for name in old_specs_dict:
            if name not in new_specs_dict:
                differences['specs_changed'].append({
                    'name': name,
                    'old': old_specs_dict[name],
                    'new': None,
                    'type': 'removed'
                })
                differences['has_changes'] = True

    except Exception as e:
        print(f"Ошибка сравнения характеристик: {e}")

    return differences


def resolve_equipment_conflict(existing: dict, new_data: dict, differences: dict, manager_instruction: str) -> dict:
    """
    Применяет инструкции менеджера для разрешения конфликта характеристик.
    manager_instruction — голосовое/текстовое указание что оставить.
    Возвращает итоговые данные для сохранения.
    """
    prompt = f"""Менеджер решает конфликт данных оборудования.

Существующие данные в библиотеке:
{json.dumps(existing, ensure_ascii=False, indent=2)}

Новые данные из документа:
{json.dumps(new_data, ensure_ascii=False, indent=2)}

Различия:
{json.dumps(differences, ensure_ascii=False, indent=2)}

Инструкция менеджера: {manager_instruction}

На основе инструкции менеджера составь итоговые данные оборудования.
Верни ТОЛЬКО JSON с итоговыми данными оборудования (тот же формат что и входные данные).
"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )

    result = parse_json_from_text(response.content[0].text)
    return result or new_data
