import os
import json
from anthropic import Anthropic
from database import search_equipment, get_all_equipment

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
- Особые условия

ПРАВИЛА:
1. Задавай уточняющие вопросы по одному — не анкетой
2. Если менеджер назвал модель оборудования — используй её
3. Если данных достаточно — верни JSON с данными КП
4. Веди деловой но дружелюбный тон
5. Понимай сокращения и разговорный стиль ("сделай подешевле" = изменить цену)
6. Если менеджер говорит "готово", "сохрани", "отправляй" — это сигнал финализировать КП

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
  "production_time": "срок изготовления",
  "total_price": 80000,
  "currency": "ЮАНЕЙ",
  "notes": "дополнительные примечания или null"
}
```

Если данных ещё не хватает — просто веди диалог, без JSON.
"""


def parse_claude_response(text: str) -> tuple[str, dict | None]:
    """Разбирает ответ Claude — текст и возможный JSON"""
    import re
    json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if json_match:
        try:
            data = json.loads(json_match.group(1))
            if data.get('ready'):
                return text, data
        except json.JSONDecodeError:
            pass
    return text, None


def chat_with_claude(conversation_history: list, user_message: str) -> tuple[str, dict | None]:
    """
    Ведёт диалог с Claude.
    Возвращает (текст ответа, данные КП если готовы)
    """
    # Добавляем сообщение пользователя
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

    # Добавляем ответ в историю
    conversation_history.append({
        "role": "assistant",
        "content": assistant_message
    })

    # Проверяем есть ли готовые данные КП
    text, kp_data = parse_claude_response(assistant_message)

    return assistant_message, kp_data


def process_edit(conversation_history: list, edit_message: str) -> tuple[str, dict | None]:
    """Обрабатывает правки к готовому КП"""
    edit_prompt = f"""Менеджер хочет внести правку в КП: {edit_message}

Обнови данные КП и верни обновлённый JSON. Если правка понятна — сразу верни JSON без лишних вопросов."""

    return chat_with_claude(conversation_history, edit_prompt)


def generate_kp_text(kp_data: dict, equipment_from_db: list) -> str:
    """
    Генерирует живой текст для КП на основе данных.
    Возвращает текст в деловом стиле.
    """
    items_desc = []
    for item in kp_data.get('items', []):
        eq = next((e for e in equipment_from_db if item['model'] in e.get('model', '')), None)
        if eq:
            items_desc.append(f"- {item['name']} (модель {item['model']}): {item['quantity']} шт. по {item['unit_price']:,.0f} {item['currency']}")
        else:
            items_desc.append(f"- {item['name']}: {item['quantity']} шт. по {item['unit_price']:,.0f} {item['currency']}")

    prompt = f"""Ты помогаешь составить коммерческое предложение для промышленного оборудования.

Данные КП:
- Клиент: {kp_data.get('client')}
- Оборудование: {chr(10).join(items_desc)}
- Итого: {kp_data.get('total_price'):,.0f} {kp_data.get('currency')}
- Условия оплаты: {kp_data.get('payment_terms')}
- Срок изготовления: {kp_data.get('production_time')}
- Примечания: {kp_data.get('notes') or 'нет'}

Напиши краткое вводное предложение для КП в деловом стиле (1-2 предложения).
Начинай сразу с текста, без приветствий и пояснений.
Пример стиля: "ООО «Фарсал» предлагает к рассмотрению данное коммерческое предложение на поставку промышленного оборудования."
"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=300,
        messages=[{"role": "user", "content": prompt}]
    )

    return response.content[0].text.strip()


def extract_equipment_from_doc(doc_text: str) -> dict:
    """
    Извлекает данные оборудования из текста КП документа.
    Используется при загрузке КП в библиотеку.
    """
    prompt = f"""Из следующего текста коммерческого предложения извлеки данные об оборудовании.

Текст:
{doc_text}

Верни ТОЛЬКО JSON без лишнего текста:
```json
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
```
"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text
    _, data = parse_claude_response(text)

    # Попробуем распарсить без ```json``` тегов
    if not data:
        import re
        try:
            json_match = re.search(r'\{.*\}', text, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group())
        except Exception:
            pass

    return data or {}