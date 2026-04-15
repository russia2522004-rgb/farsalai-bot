import os
import json
import logging
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)
from openai import OpenAI

from database import (
    init_db, add_equipment, get_all_equipment, search_equipment,
    get_equipment_by_model, delete_equipment, update_equipment,
    save_kp, update_kp, get_kp_by_number, search_kp, get_recent_kp,
    generate_kp_number
)
from claude_agent import chat_with_claude, process_edit, extract_equipment_from_doc
from document_generator import generate_kp_document, cleanup_temp_files
from storage import upload_kp_files, add_kp_to_sheets, update_kp_in_sheets, ensure_headers, upload_equipment_photo

# ── Логирование ──
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Константы состояний ──
MAIN_MENU, CREATING_KP, EDITING_KP, ADDING_EQUIPMENT, SEARCHING_KP = range(5)

# ── Авторизация ──
ALLOWED_IDS = set(map(int, os.getenv('ALLOWED_USER_IDS', '').split(',')))
ADMIN_ID = int(os.getenv('ALLOWED_USER_IDS', '0').split(',')[0])

# ── OpenAI для Whisper ──
openai_client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

# ── Хранилище сессий пользователей ──
# user_sessions[user_id] = {
#   'state': ...,
#   'conversation': [...],  # история для Claude
#   'kp_data': {...},       # текущие данные КП
#   'kp_number': '...',     # номер КП
#   'sheets_row': ...,      # строка в Sheets
# }
user_sessions = {}


def is_allowed(user_id: int) -> bool:
    return user_id in ALLOWED_IDS


def get_session(user_id: int) -> dict:
    if user_id not in user_sessions:
        user_sessions[user_id] = {
            'state': MAIN_MENU,
            'conversation': [],
            'kp_data': None,
            'kp_number': None,
            'sheets_row': None,
        }
    return user_sessions[user_id]


async def transcribe_voice(file_path: str) -> str:
    """Транскрибирует голосовое сообщение через Whisper"""
    with open(file_path, 'rb') as audio:
        transcript = openai_client.audio.transcriptions.create(
            model='whisper-1',
            file=audio,
            language='ru'
        )
    return transcript.text


async def handle_voice_or_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    """Получает текст из голосового или текстового сообщения"""
    if update.message.voice:
        voice = update.message.voice
        file = await context.bot.get_file(voice.file_id)
        voice_path = f'temp_voice_{update.effective_user.id}.ogg'
        await file.download_to_drive(voice_path)

        await update.message.reply_text('🎙 Распознаю голос...')
        text = await transcribe_voice(voice_path)
        os.remove(voice_path)

        await update.message.reply_text(f'📝 Распознано: _{text}_', parse_mode='Markdown')
        return text

    return update.message.text


# ─── Команды ─────────────────────────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text('⛔ У вас нет доступа к этому боту.')
        return

    session = get_session(user_id)
    session['state'] = MAIN_MENU

    keyboard = [
        ['📄 Создать КП'],
        ['📚 Библиотека оборудования', '🔍 Найти КП'],
        ['📋 Последние КП'],
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

    await update.message.reply_text(
        '👋 Добро пожаловать в *ФарсалИИ*!\n\n'
        'Я помогу вам быстро подготовить коммерческое предложение.\n\n'
        'Выберите действие:',
        reply_markup=reply_markup,
        parse_mode='Markdown'
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    await update.message.reply_text(
        '*Доступные команды:*\n\n'
        '/start — главное меню\n'
        '/newkp — создать новое КП\n'
        '/cancel — отменить текущее действие\n'
        '/equipment — библиотека оборудования\n'
        '/history — последние КП\n'
        '/find — найти КП\n\n'
        'Вы можете отправлять *голосовые* или *текстовые* сообщения на любом этапе.',
        parse_mode='Markdown'
    )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    session = get_session(user_id)
    session['state'] = MAIN_MENU
    session['conversation'] = []
    session['kp_data'] = None

    keyboard = [
        ['📄 Создать КП'],
        ['📚 Библиотека оборудования', '🔍 Найти КП'],
        ['📋 Последние КП'],
    ]
    await update.message.reply_text(
        '❌ Действие отменено. Возвращаемся в главное меню.',
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )


# ─── Создание КП ─────────────────────────────────────────────────────────────

async def start_kp_creation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = get_session(user_id)
    session['state'] = CREATING_KP
    session['conversation'] = []
    session['kp_data'] = None

    await update.message.reply_text(
        '📄 *Создание нового КП*\n\n'
        'Опишите что нужно включить в коммерческое предложение.\n'
        'Можно голосом или текстом.\n\n'
        '_Например: "КП для ООО Ромашка, насос IE-2, 1 штука, цена 80 тысяч юаней"_',
        parse_mode='Markdown',
        reply_markup=ReplyKeyboardRemove()
    )


async def process_kp_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает сообщения при создании КП"""
    user_id = update.effective_user.id
    session = get_session(user_id)

    text = await handle_voice_or_text(update, context)
    if not text:
        return

    # Показываем что обрабатываем
    await context.bot.send_chat_action(update.effective_chat.id, 'typing')

    # Отправляем Claude
    response, kp_data = chat_with_claude(session['conversation'], text)

    if kp_data:
        # Данные собраны — генерируем КП
        session['kp_data'] = kp_data
        await update.message.reply_text(
            '✅ *Данные собраны!* Генерирую документ...',
            parse_mode='Markdown'
        )
        await generate_and_send_kp(update, context, session, kp_data)
    else:
        # Продолжаем диалог
        # Убираем JSON из ответа если есть
        clean_response = response.split('```json')[0].strip() if '```json' in response else response
        await update.message.reply_text(clean_response)


async def generate_and_send_kp(update, context, session, kp_data):
    """Генерирует и отправляет КП"""
    user_id = update.effective_user.id
    user = update.effective_user
    manager_name = user.full_name

    # Генерируем номер КП
    models = [item.get('model', '') for item in kp_data.get('items', [])]
    kp_number = generate_kp_number(models)
    kp_data['kp_number'] = kp_number
    kp_data['kp_date'] = datetime.now().strftime('%d.%m.%Y')
    session['kp_number'] = kp_number

    # Генерируем документ
    await context.bot.send_chat_action(update.effective_chat.id, 'upload_document')

    try:
        docx_path, pdf_path = generate_kp_document(kp_data, manager_name)

        # Загружаем на Яндекс Диск
        await update.message.reply_text('☁️ Загружаю на Яндекс Диск...')
        word_url, pdf_url = upload_kp_files(docx_path, pdf_path, kp_number)

        # Сохраняем в БД
        equipment_list = ', '.join([item.get('name', item.get('model')) for item in kp_data.get('items', [])])
        db_data = {
            'kp_number': kp_number,
            'kp_date': kp_data['kp_date'],
            'client': kp_data.get('client'),
            'equipment_list': equipment_list,
            'total_price': kp_data.get('total_price'),
            'currency': kp_data.get('currency', 'ЮАНЕЙ'),
            'payment_terms': kp_data.get('payment_terms'),
            'manager_id': user_id,
            'manager_name': manager_name,
            'yandex_word_url': word_url,
            'yandex_pdf_url': pdf_url,
        }

        # Добавляем в Google Sheets
        sheets_row = add_kp_to_sheets({**db_data}, word_url, pdf_url)
        db_data['sheets_row'] = sheets_row
        session['sheets_row'] = sheets_row

        save_kp(db_data)

        # Отправляем файлы менеджеру
        with open(docx_path, 'rb') as f:
            await update.message.reply_document(f, filename=f'КП_{kp_number}.docx')
        with open(pdf_path, 'rb') as f:
            await update.message.reply_document(f, filename=f'КП_{kp_number}.pdf')

        cleanup_temp_files(docx_path, pdf_path)

        session['state'] = EDITING_KP

        keyboard = [['✅ Готово, сохранить', '✏️ Внести правки'], ['❌ Отменить']]
        await update.message.reply_text(
            f'📄 *КП №{kp_number} готово!*\n\n'
            f'🔗 [Word на Яндекс Диске]({word_url})\n'
            f'🔗 [PDF на Яндекс Диске]({pdf_url})\n\n'
            'Проверьте документ. Нужны правки?',
            parse_mode='Markdown',
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True),
            disable_web_page_preview=True
        )

    except Exception as e:
        logger.error(f'Ошибка генерации КП: {e}')
        await update.message.reply_text(
            f'❌ Ошибка при генерации КП: {str(e)}\n\nПопробуйте ещё раз или обратитесь к администратору.'
        )


# ─── Редактирование КП ───────────────────────────────────────────────────────

async def process_edit_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает правки к КП"""
    user_id = update.effective_user.id
    session = get_session(user_id)

    text = await handle_voice_or_text(update, context)
    if not text:
        return

    # Проверяем команды
    if text in ['✅ Готово, сохранить', 'готово', 'сохранить', 'отправляй']:
        await finalize_kp(update, context, session)
        return

    if text in ['❌ Отменить']:
        await cancel(update, context)
        return

    # Обрабатываем правку
    await context.bot.send_chat_action(update.effective_chat.id, 'typing')
    await update.message.reply_text('✏️ Вношу правки...')

    response, updated_data = process_edit(session['conversation'], text)

    if updated_data:
        session['kp_data'] = updated_data
        await update.message.reply_text('🔄 Перегенерирую документ...')
        await regenerate_kp(update, context, session, updated_data)
    else:
        await update.message.reply_text(response)


async def regenerate_kp(update, context, session, kp_data):
    """Перегенерирует КП с правками"""
    user = update.effective_user
    kp_number = session['kp_number']
    kp_data['kp_number'] = kp_number

    try:
        docx_path, pdf_path = generate_kp_document(kp_data, user.full_name)

        # Перезаписываем на Яндекс Диске
        word_url, pdf_url = upload_kp_files(docx_path, pdf_path, kp_number)

        # Обновляем Sheets
        if session.get('sheets_row'):
            update_kp_in_sheets(session['sheets_row'], word_url, pdf_url)

        # Обновляем БД
        update_kp(kp_number, {
            'yandex_word_url': word_url,
            'yandex_pdf_url': pdf_url,
        })

        # Отправляем новые файлы
        with open(docx_path, 'rb') as f:
            await update.message.reply_document(f, filename=f'КП_{kp_number}.docx')
        with open(pdf_path, 'rb') as f:
            await update.message.reply_document(f, filename=f'КП_{kp_number}.pdf')

        cleanup_temp_files(docx_path, pdf_path)

        keyboard = [['✅ Готово, сохранить', '✏️ Внести правки'], ['❌ Отменить']]
        await update.message.reply_text(
            f'✅ КП обновлено!\n\n'
            f'🔗 [Word]({word_url}) | [PDF]({pdf_url})\n\n'
            'Ещё правки нужны?',
            parse_mode='Markdown',
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True),
            disable_web_page_preview=True
        )

    except Exception as e:
        logger.error(f'Ошибка перегенерации: {e}')
        await update.message.reply_text(f'❌ Ошибка: {str(e)}')


async def finalize_kp(update, context, session):
    """Финализирует КП"""
    kp_number = session.get('kp_number', '—')
    session['state'] = MAIN_MENU
    session['conversation'] = []
    session['kp_data'] = None

    keyboard = [
        ['📄 Создать КП'],
        ['📚 Библиотека оборудования', '🔍 Найти КП'],
        ['📋 Последние КП'],
    ]
    await update.message.reply_text(
        f'✅ *КП №{kp_number} сохранено!*\n\n'
        'Документы сохранены на Яндекс Диске и в журнале Google Sheets.',
        parse_mode='Markdown',
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )


# ─── Библиотека оборудования ─────────────────────────────────────────────────

async def equipment_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    equipment = get_all_equipment()
    if not equipment:
        text = '📚 *Библиотека оборудования пуста.*\n\nДобавьте оборудование командой /add\_equipment'
    else:
        lines = ['📚 *Библиотека оборудования:*\n']
        for eq in equipment:
            price = f"{eq['base_price']:,.0f} {eq['currency']}" if eq['base_price'] else 'цена не указана'
            lines.append(f"• *{eq['model']}* — {eq['name']} ({price})")
        text = '\n'.join(lines)
        text += '\n\n_Для добавления нового оборудования отправьте файл КП или используйте /add\_equipment_'

    await update.message.reply_text(text, parse_mode='Markdown')


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает загруженный документ (КП для библиотеки)"""
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    session = get_session(user_id)
    doc = update.message.document

    if not doc.file_name.endswith(('.docx', '.pdf')):
        await update.message.reply_text('Пожалуйста, отправьте файл в формате .docx или .pdf')
        return

    await update.message.reply_text('📄 Читаю документ...')

    # Скачиваем файл
    file = await context.bot.get_file(doc.file_id)
    local_path = f'temp_{user_id}_{doc.file_name}'
    await file.download_to_drive(local_path)

    try:
        # Извлекаем текст
        if doc.file_name.endswith('.docx'):
            from docx import Document as DocxDocument
            d = DocxDocument(local_path)
            doc_text = '\n'.join([p.text for p in d.paragraphs if p.text.strip()])
        else:
            doc_text = 'PDF файл — текст извлечён частично'

        await update.message.reply_text('🤖 Анализирую содержимое...')

        # Claude извлекает данные
        eq_data = extract_equipment_from_doc(doc_text)

        if eq_data:
            # Показываем что распознали
            specs_preview = ''
            if eq_data.get('specs'):
                specs = eq_data['specs'][:3]
                specs_preview = '\n'.join([f"  • {s['name']}: {s['value']}" for s in specs])
                if len(eq_data['specs']) > 3:
                    specs_preview += f'\n  • ...ещё {len(eq_data["specs"]) - 3} характеристик'

            preview = (
                f'📋 *Распознано:*\n\n'
                f'Название: {eq_data.get("name", "—")}\n'
                f'Модель: {eq_data.get("model", "—")}\n'
                f'Цена: {eq_data.get("base_price", "—")} {eq_data.get("currency", "")}\n'
                f'Срок: {eq_data.get("production_time", "—")}\n'
            )
            if specs_preview:
                preview += f'\nХарактеристики (первые 3):\n{specs_preview}\n'

            preview += '\n✅ Сохранить в библиотеку?'

            # Сохраняем данные в сессию для подтверждения
            session['pending_equipment'] = eq_data
            session['pending_equipment']['doc_path'] = local_path

            keyboard = [['✅ Сохранить', '❌ Отменить']]
            await update.message.reply_text(
                preview,
                parse_mode='Markdown',
                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            )
            session['state'] = ADDING_EQUIPMENT
        else:
            await update.message.reply_text('❌ Не удалось распознать данные. Попробуйте другой файл.')
            os.remove(local_path)

    except Exception as e:
        logger.error(f'Ошибка обработки документа: {e}')
        await update.message.reply_text(f'❌ Ошибка: {str(e)}')
        if os.path.exists(local_path):
            os.remove(local_path)


async def confirm_add_equipment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Подтверждение добавления оборудования"""
    user_id = update.effective_user.id
    session = get_session(user_id)
    text = update.message.text

    if text == '✅ Сохранить' and session.get('pending_equipment'):
        eq_data = session['pending_equipment']
        doc_path = eq_data.pop('doc_path', None)

        # Конвертируем specs в JSON строку
        if isinstance(eq_data.get('specs'), list):
            eq_data['specs'] = json.dumps(eq_data['specs'], ensure_ascii=False)

        eq_id = add_equipment(eq_data)

        if doc_path and os.path.exists(doc_path):
            os.remove(doc_path)

        session['pending_equipment'] = None
        session['state'] = MAIN_MENU

        keyboard = [
            ['📄 Создать КП'],
            ['📚 Библиотека оборудования', '🔍 Найти КП'],
            ['📋 Последние КП'],
        ]
        await update.message.reply_text(
            f'✅ *{eq_data.get("name")}* добавлено в библиотеку!',
            parse_mode='Markdown',
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
    else:
        session['state'] = MAIN_MENU
        await cancel(update, context)


# ─── История и поиск КП ──────────────────────────────────────────────────────

async def show_history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    # Администратор видит все, менеджер — только свои
    if user_id == ADMIN_ID:
        kps = get_recent_kp(limit=10)
    else:
        kps = get_recent_kp(manager_id=user_id, limit=10)

    if not kps:
        await update.message.reply_text('📋 История КП пуста.')
        return

    lines = ['📋 *Последние КП:*\n']
    for kp in kps:
        lines.append(
            f"• *{kp['kp_number']}* от {kp['kp_date']}\n"
            f"  Клиент: {kp['client'] or '—'}\n"
            f"  Оборудование: {kp['equipment_list'] or '—'}\n"
            f"  Сумма: {kp['total_price']:,.0f} {kp['currency']}\n"
        )

    await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')


async def find_kp(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    session = get_session(user_id)
    session['state'] = SEARCHING_KP

    await update.message.reply_text(
        '🔍 Введите для поиска:\n'
        '• Название клиента\n'
        '• Номер КП\n'
        '• Название оборудования',
        reply_markup=ReplyKeyboardRemove()
    )


async def process_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = get_session(user_id)
    query = update.message.text

    results = search_kp(query)
    session['state'] = MAIN_MENU

    if not results:
        await update.message.reply_text('❌ Ничего не найдено.')
    else:
        lines = [f'🔍 *Результаты поиска по "{query}":*\n']
        for kp in results:
            word_link = f'[Word]({kp["yandex_word_url"]})' if kp.get('yandex_word_url') else '—'
            pdf_link = f'[PDF]({kp["yandex_pdf_url"]})' if kp.get('yandex_pdf_url') else '—'
            lines.append(
                f"• *{kp['kp_number']}* от {kp['kp_date']}\n"
                f"  {kp['client'] or '—'} | {kp['equipment_list'] or '—'}\n"
                f"  {word_link} | {pdf_link}\n"
            )
        await update.message.reply_text(
            '\n'.join(lines),
            parse_mode='Markdown',
            disable_web_page_preview=True
        )

    keyboard = [
        ['📄 Создать КП'],
        ['📚 Библиотека оборудования', '🔍 Найти КП'],
        ['📋 Последние КП'],
    ]
    await update.message.reply_text(
        'Что дальше?',
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )


# ─── Главный обработчик сообщений ────────────────────────────────────────────

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text('⛔ У вас нет доступа.')
        return

    session = get_session(user_id)
    text = update.message.text or ''
    state = session['state']

    # Главное меню — кнопки
    if text == '📄 Создать КП' or text == '/newkp':
        await start_kp_creation(update, context)
        session['state'] = CREATING_KP
        return

    if text == '📚 Библиотека оборудования' or text == '/equipment':
        await equipment_menu(update, context)
        return

    if text == '🔍 Найти КП' or text == '/find':
        await find_kp(update, context)
        return

    if text == '📋 Последние КП' or text == '/history':
        await show_history(update, context)
        return

    # Маршрутизация по состоянию
    if state == CREATING_KP:
        await process_kp_message(update, context)

    elif state == EDITING_KP:
        await process_edit_message(update, context)

    elif state == ADDING_EQUIPMENT:
        await confirm_add_equipment(update, context)

    elif state == SEARCHING_KP:
        await process_search(update, context)

    else:
        # Главное меню — неизвестная команда
        await update.message.reply_text(
            'Используйте кнопки меню или /start для начала.',
        )


# ─── Запуск ──────────────────────────────────────────────────────────────────

def main():
    init_db()

    # Инициализируем заголовки Google Sheets
    try:
        ensure_headers()
    except Exception as e:
        logger.warning(f'Google Sheets недоступен: {e}')

    token = os.getenv('TELEGRAM_BOT_TOKEN')
    app = Application.builder().token(token).build()

    # Команды
    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_command))
    app.add_handler(CommandHandler('cancel', cancel))
    app.add_handler(CommandHandler('newkp', start_kp_creation))
    app.add_handler(CommandHandler('equipment', equipment_menu))
    app.add_handler(CommandHandler('history', show_history))
    app.add_handler(CommandHandler('find', find_kp))

    # Документы
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    # Голосовые и текстовые сообщения
    app.add_handler(MessageHandler(
        filters.TEXT | filters.VOICE,
        handle_message
    ))

    logger.info('ФарсалИИ запущен!')
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()