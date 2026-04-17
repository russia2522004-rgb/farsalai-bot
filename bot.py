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

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

MAIN_MENU, CREATING_KP, EDITING_KP, ADDING_EQUIPMENT, SEARCHING_KP = range(5)

ALLOWED_IDS = set(map(int, os.getenv('ALLOWED_USER_IDS', '').split(',')))
ADMIN_ID = int(os.getenv('ALLOWED_USER_IDS', '0').split(',')[0])
LOG_CHANNEL_ID = os.getenv('LOG_CHANNEL_ID')

# Имена менеджеров
MANAGER_NAMES = {
    177592975: 'ЛавАлеСер',
    922595157: 'ЛавАлеЮр',
}

openai_client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
user_sessions = {}


def is_allowed(user_id: int) -> bool:
    return user_id in ALLOWED_IDS


def get_manager_name(user_id: int, fallback: str) -> str:
    return MANAGER_NAMES.get(user_id, fallback)


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


async def send_log(context, message: str):
    """Отправляет сообщение в канал логов"""
    if not LOG_CHANNEL_ID:
        return
    try:
        await context.bot.send_message(
            chat_id=LOG_CHANNEL_ID,
            text=message,
            parse_mode='HTML'
        )
    except Exception as e:
        logger.error(f"Ошибка отправки лога: {e}")


async def log_bot_response(context, manager_name: str, text: str):
    """Логирует ответ бота менеджеру"""
    if not LOG_CHANNEL_ID:
        return
    # Обрезаем длинные ответы
    short_text = text[:500] + '...' if len(text) > 500 else text
    # Убираем markdown разметку для лога
    clean = short_text.replace('*', '').replace('_', '').replace('`', '')
    await send_log(context,
        f"🤖 <b>Ответ бота</b>\n"
        f"👤 Менеджеру: {manager_name}\n"
        f"💬 {clean}"
    )


async def log_files(context, manager_name: str, kp_number: str, docx_path: str, pdf_path: str):
    """Пересылает файлы КП в канал логов"""
    if not LOG_CHANNEL_ID:
        return
    try:
        await context.bot.send_message(
            chat_id=LOG_CHANNEL_ID,
            text=f"📎 <b>Файлы КП №{kp_number}</b>\n👤 Менеджер: {manager_name}",
            parse_mode='HTML'
        )
        if os.path.exists(docx_path):
            with open(docx_path, 'rb') as f:
                await context.bot.send_document(
                    chat_id=LOG_CHANNEL_ID,
                    document=f,
                    filename=f'КП_{kp_number}.docx'
                )
        if pdf_path and pdf_path.endswith('.pdf') and os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                await context.bot.send_document(
                    chat_id=LOG_CHANNEL_ID,
                    document=f,
                    filename=f'КП_{kp_number}.pdf'
                )
    except Exception as e:
        logger.error(f"Ошибка отправки файлов в лог: {e}")


async def transcribe_voice(file_path: str) -> str:
    with open(file_path, 'rb') as audio:
        transcript = openai_client.audio.transcriptions.create(
            model='whisper-1',
            file=audio,
            language='ru'
        )
    return transcript.text


async def handle_voice_or_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    user = update.effective_user
    manager_name = get_manager_name(user.id, user.full_name)

    if update.message.voice:
        voice = update.message.voice
        file = await context.bot.get_file(voice.file_id)
        voice_path = f'temp_voice_{user.id}.ogg'
        await file.download_to_drive(voice_path)
        await update.message.reply_text('🎙 Распознаю голос...')
        text = await transcribe_voice(voice_path)
        os.remove(voice_path)
        await update.message.reply_text(f'📝 Распознано: _{text}_', parse_mode='Markdown')

        # Лог голосового
        await send_log(context,
            f"🎙 <b>Голосовое сообщение</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"📝 Транскрипция: {text}"
        )
        return text

    text = update.message.text
    # Лог текстового (только если не кнопка меню)
    menu_buttons = ['📄 Создать КП', '📚 Библиотека оборудования', '🔍 Найти КП',
                    '📋 Последние КП', '✅ Готово, сохранить', '✏️ Внести правки', '❌ Отменить']
    if text and text not in menu_buttons and not text.startswith('/'):
        await send_log(context,
            f"💬 <b>Текстовое сообщение</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"📝 Текст: {text}"
        )
    return text


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
        '👋 Добро пожаловать в *ФарсалИИ*!\n\nВыберите действие:',
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
        '/find — найти КП',
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
        '❌ Действие отменено.',
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )


# ─── Создание КП ─────────────────────────────────────────────────────────────

async def start_kp_creation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = get_session(user_id)
    session['state'] = CREATING_KP
    session['conversation'] = []
    session['kp_data'] = None

    manager_name = get_manager_name(user_id, update.effective_user.full_name)
    await send_log(context,
        f"📄 <b>Начало создания КП</b>\n"
        f"👤 Менеджер: {manager_name}"
    )

    await update.message.reply_text(
        '📄 *Создание нового КП*\n\n'
        'Опишите что нужно включить в коммерческое предложение.\n'
        'Можно голосом или текстом.',
        parse_mode='Markdown',
        reply_markup=ReplyKeyboardRemove()
    )


async def process_kp_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = get_session(user_id)

    text = await handle_voice_or_text(update, context)
    if not text:
        return

    await context.bot.send_chat_action(update.effective_chat.id, 'typing')
    response, kp_data = chat_with_claude(session['conversation'], text)

    if kp_data:
        session['kp_data'] = kp_data
        await update.message.reply_text('✅ *Данные собраны!* Генерирую документ...', parse_mode='Markdown')

        # Лог собранных данных
        manager_name = get_manager_name(user_id, update.effective_user.full_name)
        items_str = ', '.join([f"{i.get('name', i.get('model'))} x{i.get('quantity')} = {i.get('unit_price')} {i.get('currency', '')}" for i in kp_data.get('items', [])])
        await send_log(context,
            f"✅ <b>Данные КП собраны</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"🏢 Клиент: {kp_data.get('client', '—')}\n"
            f"📦 Позиции: {items_str}\n"
            f"💰 Итого: {kp_data.get('total_price')} {kp_data.get('currency', '')}\n"
            f"💳 Оплата: {kp_data.get('payment_terms', '—')}\n"
            f"⏱ Срок: {kp_data.get('production_time', '—')}"
        )

        await generate_and_send_kp(update, context, session, kp_data)
    else:
        clean_response = response.split('```json')[0].strip() if '```json' in response else response
        await update.message.reply_text(clean_response)
        manager_name = get_manager_name(user_id, update.effective_user.full_name)
        await log_bot_response(context, manager_name, clean_response)


async def generate_and_send_kp(update, context, session, kp_data):
    user_id = update.effective_user.id
    manager_name = get_manager_name(user_id, update.effective_user.full_name)

    models = [item.get('model', '') for item in kp_data.get('items', [])]
    kp_number = generate_kp_number(models)
    kp_data['kp_number'] = kp_number
    kp_data['kp_date'] = datetime.now().strftime('%d.%m.%Y')
    session['kp_number'] = kp_number

    await context.bot.send_chat_action(update.effective_chat.id, 'upload_document')

    try:
        docx_path, pdf_path = generate_kp_document(kp_data, manager_name)

        await update.message.reply_text('☁️ Загружаю на Яндекс Диск...')
        word_url, pdf_url = upload_kp_files(docx_path, pdf_path, kp_number)

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

        sheets_row = add_kp_to_sheets({**db_data}, word_url, pdf_url)
        db_data['sheets_row'] = sheets_row
        session['sheets_row'] = sheets_row
        save_kp(db_data)

        with open(docx_path, 'rb') as f:
            await update.message.reply_document(f, filename=f'КП_{kp_number}.docx')
        if pdf_path.endswith('.pdf') and os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                await update.message.reply_document(f, filename=f'КП_{kp_number}.pdf')

        # Логируем файлы
        await log_files(context, manager_name, kp_number, docx_path, pdf_path)

        cleanup_temp_files(docx_path, pdf_path)

        session['state'] = EDITING_KP

        keyboard = [['✅ Готово, сохранить', '✏️ Внести правки'], ['❌ Отменить']]
        await update.message.reply_text(
            f'📄 *КП №{kp_number} готово!*\n\n'
            f'🔗 [Word]({word_url}) | [PDF]({pdf_url})\n\n'
            'Проверьте документ. Нужны правки?',
            parse_mode='Markdown',
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True),
            disable_web_page_preview=True
        )

        # Лог успешного КП
        await send_log(context,
            f"📄 <b>КП сформировано</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"🔢 Номер: {kp_number}\n"
            f"🏢 Клиент: {kp_data.get('client', '—')}\n"
            f"📦 Оборудование: {equipment_list}\n"
            f"🔗 <a href='{word_url}'>Word</a> | <a href='{pdf_url}'>PDF</a>"
        )

    except Exception as e:
        logger.error(f'Ошибка генерации КП: {e}')
        await send_log(context,
            f"❌ <b>Ошибка генерации КП</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"⚠️ Ошибка: {str(e)}"
        )
        await update.message.reply_text(f'❌ Ошибка при генерации КП: {str(e)}')


# ─── Редактирование КП ───────────────────────────────────────────────────────

async def process_edit_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = get_session(user_id)

    text = await handle_voice_or_text(update, context)
    if not text:
        return

    if text in ['✅ Готово, сохранить', 'готово', 'сохранить', 'отправляй']:
        await finalize_kp(update, context, session)
        return

    if text in ['❌ Отменить']:
        await cancel(update, context)
        return

    await context.bot.send_chat_action(update.effective_chat.id, 'typing')
    await update.message.reply_text('✏️ Вношу правки...')

    manager_name = get_manager_name(user_id, update.effective_user.full_name)
    await send_log(context,
        f"✏️ <b>Правка КП</b>\n"
        f"👤 Менеджер: {manager_name}\n"
        f"🔢 КП: {session.get('kp_number', '—')}\n"
        f"📝 Правка: {text}"
    )

    response, updated_data = process_edit(session['conversation'], text)

    if updated_data:
        session['kp_data'] = updated_data
        await update.message.reply_text('🔄 Перегенерирую документ...')
        await regenerate_kp(update, context, session, updated_data)
    else:
        await update.message.reply_text(response)
        await log_bot_response(context, manager_name, response)


async def regenerate_kp(update, context, session, kp_data):
    user = update.effective_user
    manager_name = get_manager_name(user.id, user.full_name)
    kp_number = session['kp_number']
    kp_data['kp_number'] = kp_number

    try:
        docx_path, pdf_path = generate_kp_document(kp_data, manager_name)
        word_url, pdf_url = upload_kp_files(docx_path, pdf_path, kp_number)

        if session.get('sheets_row'):
            update_kp_in_sheets(session['sheets_row'], word_url, pdf_url)

        update_kp(kp_number, {
            'yandex_word_url': word_url,
            'yandex_pdf_url': pdf_url,
        })

        with open(docx_path, 'rb') as f:
            await update.message.reply_document(f, filename=f'КП_{kp_number}.docx')
        if pdf_path.endswith('.pdf') and os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                await update.message.reply_document(f, filename=f'КП_{kp_number}.pdf')

        # Логируем файлы
        await log_files(context, manager_name, kp_number, docx_path, pdf_path)

        cleanup_temp_files(docx_path, pdf_path)

        keyboard = [['✅ Готово, сохранить', '✏️ Внести правки'], ['❌ Отменить']]
        await update.message.reply_text(
            f'✅ КП обновлено!\n\n🔗 [Word]({word_url}) | [PDF]({pdf_url})\n\nЕщё правки нужны?',
            parse_mode='Markdown',
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True),
            disable_web_page_preview=True
        )

        await send_log(context,
            f"🔄 <b>КП обновлено</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"🔢 КП: {kp_number}\n"
            f"🔗 <a href='{word_url}'>Word</a> | <a href='{pdf_url}'>PDF</a>"
        )

    except Exception as e:
        logger.error(f'Ошибка перегенерации: {e}')
        await update.message.reply_text(f'❌ Ошибка: {str(e)}')


async def finalize_kp(update, context, session):
    kp_number = session.get('kp_number', '—')
    manager_name = get_manager_name(update.effective_user.id, update.effective_user.full_name)
    session['state'] = MAIN_MENU
    session['conversation'] = []
    session['kp_data'] = None

    keyboard = [
        ['📄 Создать КП'],
        ['📚 Библиотека оборудования', '🔍 Найти КП'],
        ['📋 Последние КП'],
    ]
    await update.message.reply_text(
        f'✅ *КП №{kp_number} сохранено!*',
        parse_mode='Markdown',
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )

    await send_log(context,
        f"✅ <b>КП финализировано</b>\n"
        f"👤 Менеджер: {manager_name}\n"
        f"🔢 КП: {kp_number}"
    )


# ─── Библиотека оборудования ─────────────────────────────────────────────────

async def equipment_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return
    equipment = get_all_equipment()
    if not equipment:
        text = '📚 *Библиотека оборудования пуста.*\n\nОтправьте файл КП чтобы добавить оборудование.'
    else:
        lines = ['📚 *Библиотека оборудования:*\n']
        for eq in equipment:
            price = f"{eq['base_price']:,.0f} {eq['currency']}" if eq['base_price'] else 'цена не указана'
            lines.append(f"• *{eq['model']}* — {eq['name']} ({price})")
        text = '\n'.join(lines)
        text += '\n\n_Отправьте файл КП чтобы добавить новое оборудование_'
    await update.message.reply_text(text, parse_mode='Markdown')


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    session = get_session(user_id)
    doc = update.message.document

    if not doc.file_name.endswith(('.docx', '.pdf')):
        await update.message.reply_text('Пожалуйста, отправьте файл в формате .docx или .pdf')
        return

    await update.message.reply_text('📄 Читаю документ...')

    file = await context.bot.get_file(doc.file_id)
    local_path = f'temp_{user_id}_{doc.file_name}'
    await file.download_to_drive(local_path)

    try:
        if doc.file_name.endswith('.docx'):
            from docx import Document as DocxDocument
            d = DocxDocument(local_path)
            doc_text = '\n'.join([p.text for p in d.paragraphs if p.text.strip()])
        else:
            doc_text = 'PDF файл'

        await update.message.reply_text('🤖 Анализирую содержимое...')
        eq_data = extract_equipment_from_doc(doc_text)

        if eq_data:
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
            await update.message.reply_text('❌ Не удалось распознать данные.')
            os.remove(local_path)

    except Exception as e:
        logger.error(f'Ошибка обработки документа: {e}')
        await update.message.reply_text(f'❌ Ошибка: {str(e)}')
        if os.path.exists(local_path):
            os.remove(local_path)


async def confirm_add_equipment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = get_session(user_id)
    text = update.message.text

    if text == '✅ Сохранить' and session.get('pending_equipment'):
        eq_data = session['pending_equipment']
        doc_path = eq_data.pop('doc_path', None)

        if isinstance(eq_data.get('specs'), list):
            eq_data['specs'] = json.dumps(eq_data['specs'], ensure_ascii=False)

        eq_id = add_equipment(eq_data)

        if doc_path and os.path.exists(doc_path):
            os.remove(doc_path)

        session['pending_equipment'] = None
        session['state'] = MAIN_MENU

        manager_name = get_manager_name(user_id, update.effective_user.full_name)
        await send_log(context,
            f"📚 <b>Добавлено оборудование</b>\n"
            f"👤 Менеджер: {manager_name}\n"
            f"📦 {eq_data.get('name')} (модель: {eq_data.get('model')})"
        )

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
            f"  {kp.get('client') or '—'} | {kp.get('equipment_list') or '—'}\n"
        )
    await update.message.reply_text('\n'.join(lines), parse_mode='Markdown')


async def find_kp(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return
    session = get_session(user_id)
    session['state'] = SEARCHING_KP
    await update.message.reply_text(
        '🔍 Введите для поиска:\n• Название клиента\n• Номер КП\n• Название оборудования',
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
        lines = [f'🔍 *Результаты по "{query}":*\n']
        for kp in results:
            word_link = f'[Word]({kp["yandex_word_url"]})' if kp.get('yandex_word_url') else '—'
            pdf_link = f'[PDF]({kp["yandex_pdf_url"]})' if kp.get('yandex_pdf_url') else '—'
            lines.append(
                f"• *{kp['kp_number']}* от {kp['kp_date']}\n"
                f"  {kp.get('client') or '—'} | {kp.get('equipment_list') or '—'}\n"
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
    await update.message.reply_text('Что дальше?', reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))


# ─── Главный обработчик ──────────────────────────────────────────────────────

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text('⛔ У вас нет доступа.')
        return

    session = get_session(user_id)
    text = update.message.text or ''
    state = session['state']

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

    if state == CREATING_KP:
        await process_kp_message(update, context)
    elif state == EDITING_KP:
        await process_edit_message(update, context)
    elif state == ADDING_EQUIPMENT:
        await confirm_add_equipment(update, context)
    elif state == SEARCHING_KP:
        await process_search(update, context)
    else:
        await update.message.reply_text('Используйте кнопки меню или /start для начала.')


# ─── Запуск ──────────────────────────────────────────────────────────────────

def main():
    init_db()
    try:
        ensure_headers()
    except Exception as e:
        logger.warning(f'Google Sheets недоступен: {e}')

    token = os.getenv('TELEGRAM_BOT_TOKEN')
    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_command))
    app.add_handler(CommandHandler('cancel', cancel))
    app.add_handler(CommandHandler('newkp', start_kp_creation))
    app.add_handler(CommandHandler('equipment', equipment_menu))
    app.add_handler(CommandHandler('history', show_history))
    app.add_handler(CommandHandler('find', find_kp))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT | filters.VOICE, handle_message))

    logger.info('ФарсалИИ запущен!')
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
