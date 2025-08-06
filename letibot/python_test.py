from imports import *
from data_sort import *

# Конфигурация
exe_path = r'C:\Users\Максим\Desktop\letibot\LETI_Parser_2.exe'
TOKEN = "7995657436:AAFIVQygLKQ9-Gppo1x72u1-wPwmBtAPTts"
DB_FILE = "abit_ids.json"
EXCEL_FILE = "результат_распределения.xlsx"
with open('directions.json', 'r', encoding='utf-8') as f:
    directions = json.load(f)
codes = list(directions.keys())

# Обновление КС
async def periodic_task(interval_minutes=30):
    while True:
        result = subprocess.run([exe_path], capture_output=True, text=True)
        data_sort()
        await asyncio.sleep(interval_minutes * 60)

# Загрузка базы данных
def load_db(file):
    try:
        with open(file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

# Сохранение в базу данных
def save_db(file, data):
    with open(file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# Команда /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    buttons = [["Добавить ID", "Помощь"], ["Обновить данные"]]
    reply_markup = ReplyKeyboardMarkup(buttons, resize_keyboard=True)
    await update.message.reply_text( # type: ignore
        "<b>Привет!</b> Я помогаю узнать место абитуриента в конкурсных списках ЛЭТИ\n\n"
        "Используй команду <b><code>/id</code></b> , чтобы указать свой <strong>Уникальный код поступающего</strong>\n"
        "Узнай свою текущую позицию в списке при помощи команды <b><code>/get</code></b>",
        reply_markup=reply_markup, parse_mode='HTML'
    )

# Команда /id
async def id_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Введите <b>Уникальный код поступающего</b> (7 цифр):", parse_mode='HTML') # type: ignore
    context.user_data['awaiting_id'] = True # type: ignore

# Проверка ID абитуриента на валидность
def is_valid_abiturient_id(abiturient_id):
    return re.fullmatch(r'\d{7}', abiturient_id) is not None

# Обработка введённого ID
async def handle_abiturient_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.user_data.get('awaiting_id'): # type: ignore
        return
    user_id = str(update.message.from_user.id) # type: ignore
    abiturient_id = update.message.text.strip() # type: ignore
    if not is_valid_abiturient_id(abiturient_id):
        await update.message.reply_text("❌ <b>Ошибка!</b> Код должен содержать ровно 7 цифр. Попробуйте ещё раз:", parse_mode='HTML') # type: ignore
        return
    db = load_db(DB_FILE)
    db[user_id] = int(abiturient_id)
    save_db(DB_FILE, db)
    await update.message.reply_text(f"✅ Код <code>{abiturient_id}</code> успешно сохранён!", parse_mode='HTML') # type: ignore
    context.user_data['awaiting_id'] = False # type: ignore

# Кнопки под клавиатурой
async def handle_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text # type: ignore
    if text == "Помощь":
        await update.message.reply_text( # type: ignore
            "Используйте команды:\n"
            "/start - приветствие\n"
            "/id - добавить/изменить ID абитуриента\n"
            "/get - узнать текущее место в конкурсном списке",
            parse_mode='HTML'
        )
    elif text == "Добавить ID":
        await id_command(update, context)
    elif text == "Обновить данные":
        await get_admission_info(update, context)

# Кнопки под сообщениями
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer() # type: ignore
    if query.data == "update_data": # type: ignore
        await get_admission_info(update, context, 1)

# Обработчик текстовых сообщений
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if context.user_data.get('awaiting_id'): # type: ignore
        await handle_abiturient_id(update, context)

# Команда /get
async def get_admission_info(update: Update, context: ContextTypes.DEFAULT_TYPE, type=None) -> None:
    try:
        telegram_id = str(update.effective_user.id) # type: ignore
        db = load_db(DB_FILE)
        if telegram_id not in db:
            await update.message.reply_text("❌ Вы не ввели <b>Уникальный код поступающего</b>! Используйте <code>/id</code> чтобы добавить его.", parse_mode='HTML') # type: ignore
            return
        abiturient_id = db[telegram_id]
        message = f"🔍 Результаты поиска ID <code>{abiturient_id}</code>:\n\n"
        sheets = pd.ExcelFile(EXCEL_FILE).sheet_names
        counter = 0
        flag = False
        for sheet in sheets:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
            applicant_row = df[df['Уникальный код поступающего'] == abiturient_id]
            if applicant_row.empty:
                counter += 1
                if counter == 3:
                    await update.message.reply_text("❌ Ваш <b>Уникальный код потупающего</b> не найден в списке. Проверьте правильность введённого кода, и используйте <code>/id</code> чтобы изменить его.", parse_mode='HTML') # type: ignore
                    return
                continue
            if not flag:
                balls = int(applicant_row['Конкурсный балл'].values[0])
                message += f'📊 Баллы: <b>{balls}</b>\n\n'
                flag = True
            if sheet == 'Целевое':
                result_value = applicant_row['Результат'].values[0]
                filtered_df = df[df['Результат'] == result_value].copy()
                filtered_df.reset_index(drop=True, inplace=True)
                position = filtered_df.index[filtered_df['Уникальный код поступающего'] == abiturient_id].tolist()[0] + 1
                results = [str(applicant_row['Согласие на зачисление'].values[0]), str(applicant_row['Результат'].values[0]), f'{position}/{len(filtered_df)}', int(applicant_row['Приоритет №'].values[0]), int(applicant_row['ID заказчика'].values[0]), str(applicant_row['Номер предложения'].values[0])]
                message += f"📌 Направление: {results[1]} {directions[results[1]]}\n"
                message += f"📒 ID заказчика: {results[4]}\n"
                message += f"📝 Номер предложения: {results[5]}\n"
                message += f"🏆 Позиция: {results[2]}\n"
                message += f"📍 Приоритет №: {results[3]}\n"
                message += f'🔄 Согласие: {results[0]}\n\n'
            elif sheet == 'Бюджет':
                result_value = applicant_row['Результат'].values[0]
                filtered_df = df[df['Результат'] == result_value].copy()
                filtered_df.reset_index(drop=True, inplace=True)
                position = filtered_df.index[filtered_df['Уникальный код поступающего'] == abiturient_id].tolist()[0] + 1
                results = [str(applicant_row['Согласие на зачисление'].values[0]), str(applicant_row['Результат'].values[0]), f'{position}/{len(filtered_df)}', int(applicant_row['Приоритет №'].values[0])]
                message += f"📌 Направление: {results[1]} {directions[results[1]]}\n"
                message += f"🏆 Позиция: {results[2]}\n"
                message += f"📍 Приоритет №: {results[3]}\n"
                message += f'🔄 Согласие: {results[0]}\n\n'
            elif sheet == 'Контракт':
                result_value = applicant_row['Результат'].values[0]
                filtered_df = df[df['Результат'] == result_value].copy()
                filtered_df.reset_index(drop=True, inplace=True)
                position = filtered_df.index[filtered_df['Уникальный код поступающего'] == abiturient_id].tolist()[0] + 1
                results = [str(applicant_row['Статус договора'].values[0]), str(applicant_row['Результат'].values[0]), f'{position}/{len(filtered_df)}', int(applicant_row['Приоритет №'].values[0])]
                message += f"📌 Направление: {results[1]} {directions[results[1]]}\n"
                message += f"🏆 Позиция: {results[2]}\n"
                message += f"📍 Приоритет №: {results[3]}\n"
                message += f'🔄 Статус договора: {results[0]}\n\n'
        moscow_time = datetime.now(ZoneInfo("Europe/Moscow"))
        formatted = moscow_time.strftime("%d.%m.%Y %H:%M:%S (%Z)")
        message += f"🕓 Обновлено: {formatted}"
        keyboard = [
            [InlineKeyboardButton("🔄 Обновить данные", callback_data="update_data")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        if type != None:
            await update.callback_query.edit_message_text(message, parse_mode='HTML', reply_markup=reply_markup) # type: ignore
            time.sleep(1)
        else:
            await update.message.reply_text(message, parse_mode='HTML', reply_markup=reply_markup) # type: ignore
            
    except FileNotFoundError:
        if type != None:
            await update.callback_query.edit_message_text("⚠ Файл с конкурсными списками временно недоступен.") # type: ignore
        else:
            await update.message.reply_text("⚠ Файл с конкурсными списками временно недоступен.") # type: ignore
        print('Ошибка получения конкурсного списка!')
    except Exception as e:
        if type != None:
            await update.callback_query.edit_message_text("⚠️ Произошла ошибка при обработке запроса.") # type: ignore
        else:
            await update.message.reply_text("⚠️ Произошла ошибка при обработке запроса.") # type: ignore
        print(e)

def main() -> None:
    application = Application.builder().token(TOKEN).build()
    application.add_handler(MessageHandler(filters.Text(["Добавить ID", "Помощь", "Обновить данные"]), handle_buttons))
    application.add_handler(CallbackQueryHandler(button_handler))
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("id", id_command))
    application.add_handler(CommandHandler("get", get_admission_info))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    print("Бот запущен...")
    application.run_polling()

def run_async_in_thread():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(periodic_task())

if __name__ == "__main__":
    thread = Thread(target=run_async_in_thread, daemon=True)
    thread.start()
    main()