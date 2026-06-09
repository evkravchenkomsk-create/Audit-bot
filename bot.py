"""
Telegram-бот: ТРЕКЕР — Дневник, Привычки, Задачи, Финансы
"""
import asyncio
import logging
import os
from datetime import datetime, date, timedelta

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton

from storage import TrackerStorage

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

TOKEN = os.getenv("BOT_TOKEN", "")
router = Router()
db = TrackerStorage("tracker_data")

class S(StatesGroup):
    main = State()
    diary_mood = State()
    diary_text = State()
    habit_name = State()
    habit_icon = State()
    habit_goal = State()
    habit_unit = State()
    task_text = State()
    task_type = State()
    task_due = State()
    fin_type = State()
    fin_amount = State()
    fin_currency = State()
    fin_category = State()
    fin_desc = State()

MOODS = {1: "😤", 2: "😑", 3: "🙂", 4: "😊", 5: "🔥"}
MOOD_LABELS = {1: "Плохо", 2: "Нейтрально", 3: "Нормально", 4: "Хорошо", 5: "Огонь"}
MONTHS_RU = ["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"]
MONTHS_SHORT = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]
CAT_ICONS = {"food":"🍔","transport":"🚗","business":"💼","health":"💪","personal":"👤","other":"📦"}
CAT_NAMES = {"food":"Еда","transport":"Транспорт","business":"Бизнес","health":"Здоровье","personal":"Личное","other":"Прочее"}
TASK_TYPES = {"today":"📅 Сегодня","shortterm":"📆 Краткосрочная","longterm":"🗓 Долгосрочная"}
CURRENCIES = ["KGS","KZT","RUB","USD"]

def today_str():
    return date.today().isoformat()

def fmt_date(iso_str):
    d = datetime.fromisoformat(iso_str)
    return f"{d.day} {MONTHS_SHORT[d.month-1]}"

def bar(val, mx, w=10):
    if mx == 0: return "░" * w
    filled = round(val / mx * w)
    return "█" * filled + "░" * (w - filled)

def kb_main():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📖 Дневник", callback_data="diary"),
         InlineKeyboardButton(text="⚡ Привычки", callback_data="habits")],
        [InlineKeyboardButton(text="✅ Задачи", callback_data="tasks"),
         InlineKeyboardButton(text="💰 Финансы", callback_data="finance")],
    ])

# ══ START ══
@router.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await state.set_state(S.main)
    name = msg.from_user.first_name or "друг"
    now = datetime.now()
    uid = msg.from_user.id
    today = today_str()
    habits = db.get_habits(uid)
    done_today = sum(1 for h in habits if db.is_habit_done(h["id"], today))
    tasks_pending = len([t for t in db.get_tasks(uid) if not t.get("done")])
    await msg.answer(
        f"👋 <b>Привет, {name}!</b>\n\n"
        f"📅 <b>{now.day} {MONTHS_RU[now.month-1]} {now.year}</b>\n\n"
        f"⚡ Привычек сегодня: <b>{done_today}/{len(habits)}</b>\n"
        f"✅ Задач в работе: <b>{tasks_pending}</b>\n\n"
        f"Выбери раздел:",
        reply_markup=kb_main(), parse_mode="HTML"
    )

@router.callback_query(F.data == "main")
async def cb_main(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.main)
    now = datetime.now()
    uid = cb.from_user.id
    today = today_str()
    habits = db.get_habits(uid)
    done_today = sum(1 for h in habits if db.is_habit_done(h["id"], today))
    tasks_pending = len([t for t in db.get_tasks(uid) if not t.get("done")])
    await cb.message.edit_text(
        f"🏠 <b>Главное меню</b>\n\n"
        f"📅 <b>{now.day} {MONTHS_RU[now.month-1]} {now.year}</b>\n\n"
        f"⚡ Привычек сегодня: <b>{done_today}/{len(habits)}</b>\n"
        f"✅ Задач в работе: <b>{tasks_pending}</b>\n\nВыбери раздел:",
        reply_markup=kb_main(), parse_mode="HTML"
    )

# ══ DIARY ══
def kb_diary(uid):
    entries = db.get_diary(uid)
    rows = []
    for e in sorted(entries, key=lambda x: x["created_at"], reverse=True)[:5]:
        mood = MOODS.get(e.get("mood", 3), "")
        preview = e["text"][:28] + ("…" if len(e["text"]) > 28 else "")
        rows.append([InlineKeyboardButton(text=f"{mood} {fmt_date(e['created_at'])} — {preview}", callback_data=f"de_{e['id']}")])
    rows.append([InlineKeyboardButton(text="✏️ Новая запись", callback_data="diary_new")])
    rows.append([InlineKeyboardButton(text="🏠 Меню", callback_data="main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

@router.callback_query(F.data == "diary")
async def cb_diary(cb: CallbackQuery):
    uid = cb.from_user.id
    entries = db.get_diary(uid)
    moods = db.get_moods(uid)
    today_mood = moods.get(today_str())
    mood_line = f"Настроение сегодня: {MOODS.get(today_mood,'—')} {MOOD_LABELS.get(today_mood,'')}" if today_mood else "Настроение сегодня: не задано"
    await cb.message.edit_text(
        f"📖 <b>Дневник</b>\n\n{mood_line}\nВсего записей: <b>{len(entries)}</b>",
        reply_markup=kb_diary(uid), parse_mode="HTML"
    )

@router.callback_query(F.data == "diary_new")
async def cb_diary_new(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.diary_mood)
    mood_btns = [InlineKeyboardButton(text=f"{MOODS[i]} {MOOD_LABELS[i]}", callback_data=f"dm_{i}") for i in range(1,6)]
    await cb.message.edit_text(
        "📖 <b>Новая запись</b>\n\nКак ты сегодня?",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[mood_btns[:3], mood_btns[3:],
            [InlineKeyboardButton(text="⬅️ Назад", callback_data="diary")]]),
        parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("dm_"))
async def cb_diary_mood(cb: CallbackQuery, state: FSMContext):
    mood = int(cb.data.split("_")[1])
    await state.update_data(diary_mood=mood)
    await state.set_state(S.diary_text)
    await cb.message.edit_text(
        f"📖 Настроение: {MOODS[mood]} <b>{MOOD_LABELS[mood]}</b>\n\nНапиши свою запись:",
        parse_mode="HTML"
    )

@router.message(S.diary_text)
async def handle_diary_text(msg: Message, state: FSMContext):
    data = await state.get_data()
    mood = data.get("diary_mood", 3)
    uid = msg.from_user.id
    db.add_diary(uid, msg.text.strip(), mood)
    db.set_mood(uid, today_str(), mood)
    await state.set_state(S.main)
    await msg.answer(
        f"✅ Запись сохранена! {MOODS[mood]}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📖 К дневнику", callback_data="diary")],
            [InlineKeyboardButton(text="🏠 Меню", callback_data="main")]
        ])
    )

@router.callback_query(F.data.startswith("de_"))
async def cb_diary_entry(cb: CallbackQuery):
    eid = cb.data[3:]
    entries = db.get_diary(cb.from_user.id)
    entry = next((e for e in entries if e["id"] == eid), None)
    if not entry:
        await cb.answer("Запись не найдена"); return
    mood = MOODS.get(entry.get("mood", 3), "")
    await cb.message.edit_text(
        f"📖 <b>{fmt_date(entry['created_at'])}</b> {mood}\n\n{entry['text']}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🗑 Удалить", callback_data=f"dd_{eid}")],
            [InlineKeyboardButton(text="⬅️ Назад", callback_data="diary")]
        ]), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("dd_"))
async def cb_diary_delete(cb: CallbackQuery):
    db.delete_diary(cb.from_user.id, cb.data[3:])
    await cb.answer("Удалено ✓")
    entries = db.get_diary(cb.from_user.id)
    await cb.message.edit_text(
        f"📖 <b>Дневник</b>\n\nЗаписей: {len(entries)}",
        reply_markup=kb_diary(cb.from_user.id), parse_mode="HTML"
    )

# ══ HABITS ══
def get_week_days():
    today = date.today()
    return [(today - timedelta(days=i)).isoformat() for i in range(6, -1, -1)]

def kb_habits(uid):
    habits = db.get_habits(uid)
    today = today_str()
    rows = []
    for h in habits:
        done = db.is_habit_done(h["id"], today)
        rows.append([InlineKeyboardButton(
            text=f"{'✅' if done else '⬜'} {h['icon']} {h['name']}",
            callback_data=f"ht_{h['id']}"
        )])
    rows.append([InlineKeyboardButton(text="➕ Новая привычка", callback_data="habit_new")])
    rows.append([InlineKeyboardButton(text="🏠 Меню", callback_data="main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

@router.callback_query(F.data == "habits")
async def cb_habits(cb: CallbackQuery):
    uid = cb.from_user.id
    habits = db.get_habits(uid)
    today = today_str()
    done_today = sum(1 for h in habits if db.is_habit_done(h["id"], today))
    week = get_week_days()
    streak = 0
    for day in reversed(week[:-1]):
        if habits and all(db.is_habit_done(h["id"], day) for h in habits):
            streak += 1
        else:
            break
    if habits and all(db.is_habit_done(h["id"], today) for h in habits):
        streak += 1
    await cb.message.edit_text(
        f"⚡ <b>Привычки</b>\n\nСегодня: <b>{done_today}/{len(habits)}</b> · Серия: <b>{streak} дн.</b>\n\nНажми чтобы отметить:",
        reply_markup=kb_habits(uid), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("ht_"))
async def cb_habit_toggle(cb: CallbackQuery):
    hid = cb.data[3:]
    uid = cb.from_user.id
    today = today_str()
    was_done = db.is_habit_done(hid, today)
    db.toggle_habit_done(hid, uid, today)
    await cb.answer("✅ Отмечено!" if not was_done else "↩️ Снято")
    habits = db.get_habits(uid)
    done_today = sum(1 for h in habits if db.is_habit_done(h["id"], today))
    week = get_week_days()
    streak = 0
    for day in reversed(week[:-1]):
        if habits and all(db.is_habit_done(h["id"], day) for h in habits):
            streak += 1
        else:
            break
    if habits and all(db.is_habit_done(h["id"], today) for h in habits):
        streak += 1
    await cb.message.edit_text(
        f"⚡ <b>Привычки</b>\n\nСегодня: <b>{done_today}/{len(habits)}</b> · Серия: <b>{streak} дн.</b>\n\nНажми чтобы отметить:",
        reply_markup=kb_habits(uid), parse_mode="HTML"
    )

@router.callback_query(F.data == "habit_new")
async def cb_habit_new(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.habit_name)
    await cb.message.edit_text("➕ <b>Новая привычка</b>\n\nВведи название:", parse_mode="HTML")

@router.message(S.habit_name)
async def handle_habit_name(msg: Message, state: FSMContext):
    await state.update_data(habit_name=msg.text.strip())
    await state.set_state(S.habit_icon)
    await msg.answer("Иконка (эмодзи), например 💪:")

@router.message(S.habit_icon)
async def handle_habit_icon(msg: Message, state: FSMContext):
    await state.update_data(habit_icon=msg.text.strip()[:2] or "⚡")
    await state.set_state(S.habit_goal)
    await msg.answer("Цель в день (число), например 1:")

@router.message(S.habit_goal)
async def handle_habit_goal(msg: Message, state: FSMContext):
    try:
        goal = int(msg.text.strip())
    except:
        goal = 1
    await state.update_data(habit_goal=goal)
    await state.set_state(S.habit_unit)
    await msg.answer("Единица измерения (раз / мин / км):")

@router.message(S.habit_unit)
async def handle_habit_unit(msg: Message, state: FSMContext):
    data = await state.get_data()
    db.add_habit(msg.from_user.id, data["habit_name"], data.get("habit_icon","⚡"), data.get("habit_goal",1), msg.text.strip() or "раз")
    await state.set_state(S.main)
    await msg.answer(
        f"✅ Привычка <b>{data['habit_name']}</b> создана!",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⚡ К привычкам", callback_data="habits")],
            [InlineKeyboardButton(text="🏠 Меню", callback_data="main")]
        ]), parse_mode="HTML"
    )

# ══ TASKS ══
def kb_tasks(uid, filter_type="all"):
    tasks = db.get_tasks(uid)
    if filter_type == "today":
        show = [t for t in tasks if t["type"] == "today" and not t.get("done")]
    elif filter_type == "longterm":
        show = [t for t in tasks if t["type"] == "longterm"]
    elif filter_type == "done":
        show = [t for t in tasks if t.get("done")]
    else:
        show = [t for t in tasks if not t.get("done")]
    rows = []
    for t in show[-8:]:
        icon = "✅" if t.get("done") else {"today":"📅","shortterm":"📆","longterm":"🗓"}.get(t["type"],"•")
        text = t["text"][:35] + ("…" if len(t["text"])>35 else "")
        rows.append([InlineKeyboardButton(text=f"{icon} {text}", callback_data=f"tc_{t['id']}")])
    rows.append([
        InlineKeyboardButton(text="Все", callback_data="tf_all"),
        InlineKeyboardButton(text="Сегодня", callback_data="tf_today"),
        InlineKeyboardButton(text="✓ Готово", callback_data="tf_done"),
    ])
    rows.append([InlineKeyboardButton(text="➕ Новая задача", callback_data="task_new")])
    rows.append([InlineKeyboardButton(text="🏠 Меню", callback_data="main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

@router.callback_query(F.data == "tasks")
async def cb_tasks(cb: CallbackQuery):
    uid = cb.from_user.id
    tasks = db.get_tasks(uid)
    pending = len([t for t in tasks if not t.get("done")])
    done = len([t for t in tasks if t.get("done")])
    await cb.message.edit_text(
        f"✅ <b>Задачи</b>\n\nВ работе: <b>{pending}</b> · Выполнено: <b>{done}</b>",
        reply_markup=kb_tasks(uid), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("tf_"))
async def cb_task_filter(cb: CallbackQuery):
    uid = cb.from_user.id
    tasks = db.get_tasks(uid)
    pending = len([t for t in tasks if not t.get("done")])
    await cb.message.edit_text(
        f"✅ <b>Задачи</b>\n\nВ работе: <b>{pending}</b>",
        reply_markup=kb_tasks(uid, cb.data[3:]), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("tc_"))
async def cb_task_card(cb: CallbackQuery):
    tid = cb.data[3:]
    tasks = db.get_tasks(cb.from_user.id)
    task = next((t for t in tasks if t["id"] == tid), None)
    if not task:
        await cb.answer("Не найдено"); return
    status = "✅ Выполнено" if task.get("done") else "🔄 В работе"
    type_name = TASK_TYPES.get(task["type"], task["type"])
    due_line = f"\n📅 Срок: {task['due']}" if task.get("due") else ""
    await cb.message.edit_text(
        f"<b>{task['text']}</b>\n\n{type_name} · {status}{due_line}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(
                text="↩️ Снять" if task.get("done") else "✅ Выполнено",
                callback_data=f"tdone_{tid}"
            )],
            [InlineKeyboardButton(text="🗑 Удалить", callback_data=f"tdel_{tid}")],
            [InlineKeyboardButton(text="⬅️ К задачам", callback_data="tasks")]
        ]), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("tdone_"))
async def cb_task_done(cb: CallbackQuery):
    db.toggle_task_done(cb.from_user.id, cb.data[6:])
    await cb.answer("✓ Обновлено")
    uid = cb.from_user.id
    tasks = db.get_tasks(uid)
    pending = len([t for t in tasks if not t.get("done")])
    await cb.message.edit_text(
        f"✅ <b>Задачи</b>\n\nВ работе: <b>{pending}</b>",
        reply_markup=kb_tasks(uid), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("tdel_"))
async def cb_task_delete(cb: CallbackQuery):
    db.delete_task(cb.from_user.id, cb.data[5:])
    await cb.answer("Удалено ✓")
    uid = cb.from_user.id
    tasks = db.get_tasks(uid)
    pending = len([t for t in tasks if not t.get("done")])
    await cb.message.edit_text(
        f"✅ <b>Задачи</b>\n\nВ работе: <b>{pending}</b>",
        reply_markup=kb_tasks(uid), parse_mode="HTML"
    )

@router.callback_query(F.data == "task_new")
async def cb_task_new(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.task_text)
    await cb.message.edit_text("➕ <b>Новая задача</b>\n\nЧто нужно сделать?", parse_mode="HTML")

@router.message(S.task_text)
async def handle_task_text(msg: Message, state: FSMContext):
    await state.update_data(task_text=msg.text.strip())
    await state.set_state(S.task_type)
    await msg.answer(
        "Тип задачи:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📅 Сегодня", callback_data="tt_today"),
             InlineKeyboardButton(text="📆 Краткосрочная", callback_data="tt_shortterm")],
            [InlineKeyboardButton(text="🗓 Долгосрочная", callback_data="tt_longterm")]
        ])
    )

@router.callback_query(F.data.startswith("tt_"))
async def handle_task_type(cb: CallbackQuery, state: FSMContext):
    await state.update_data(task_type=cb.data[3:])
    await state.set_state(S.task_due)
    await cb.message.edit_text("Срок (необязательно, или отправь '-'):")

@router.message(S.task_due)
async def handle_task_due(msg: Message, state: FSMContext):
    data = await state.get_data()
    due = msg.text.strip() if msg.text.strip() not in ("-","нет","") else ""
    db.add_task(msg.from_user.id, data["task_text"], data.get("task_type","today"), due)
    await state.set_state(S.main)
    await msg.answer(
        "✅ Задача добавлена!",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ К задачам", callback_data="tasks")],
            [InlineKeyboardButton(text="🏠 Меню", callback_data="main")]
        ])
    )

# ══ FINANCE ══
def get_finance_summary(uid):
    records = db.get_finance(uid)
    month = datetime.now().strftime("%Y-%m")
    month_rec = [r for r in records if r["created_at"][:7] == month]
    income = sum(r["amount"] for r in month_rec if r["type"] == "income")
    expense = sum(r["amount"] for r in month_rec if r["type"] == "expense")
    return income, expense

def kb_finance(uid, filter_type="all"):
    records = sorted(db.get_finance(uid), key=lambda r: r["created_at"], reverse=True)
    if filter_type == "income":
        display = [r for r in records if r["type"] == "income"]
    elif filter_type == "expense":
        display = [r for r in records if r["type"] == "expense"]
    else:
        display = records
    rows = []
    for r in display[:6]:
        icon = "💰" if r["type"] == "income" else CAT_ICONS.get(r.get("category","other"),"📦")
        sign = "+" if r["type"] == "income" else "-"
        amt = f"{sign}{int(r['amount'])}"
        desc = (r.get("description") or CAT_NAMES.get(r.get("category","other"),""))[:22]
        rows.append([InlineKeyboardButton(text=f"{icon} {desc} {amt} {r.get('currency','')}", callback_data=f"fc_{r['id']}")])
    rows.append([
        InlineKeyboardButton(text="Все", callback_data="ff_all"),
        InlineKeyboardButton(text="📈 Доходы", callback_data="ff_income"),
        InlineKeyboardButton(text="📉 Расходы", callback_data="ff_expense"),
    ])
    rows.append([InlineKeyboardButton(text="➕ Добавить", callback_data="fin_new")])
    rows.append([InlineKeyboardButton(text="🏠 Меню", callback_data="main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

@router.callback_query(F.data == "finance")
async def cb_finance(cb: CallbackQuery):
    uid = cb.from_user.id
    income, expense = get_finance_summary(uid)
    balance = income - expense
    now = datetime.now()
    await cb.message.edit_text(
        f"💰 <b>Финансы — {MONTHS_RU[now.month-1]} {now.year}</b>\n\n"
        f"📈 Доходы: <b>{int(income):,}</b>\n"
        f"📉 Расходы: <b>{int(expense):,}</b>\n"
        f"💼 Баланс: <b>{'+' if balance>=0 else ''}{int(balance):,}</b>",
        reply_markup=kb_finance(uid), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("ff_"))
async def cb_finance_filter(cb: CallbackQuery):
    uid = cb.from_user.id
    income, expense = get_finance_summary(uid)
    balance = income - expense
    await cb.message.edit_text(
        f"💰 <b>Финансы</b>\n\n📈 {int(income):,} / 📉 {int(expense):,}\n💼 {'+' if balance>=0 else ''}{int(balance):,}",
        reply_markup=kb_finance(uid, cb.data[3:]), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("fc_"))
async def cb_finance_card(cb: CallbackQuery):
    rid = cb.data[3:]
    records = db.get_finance(cb.from_user.id)
    r = next((x for x in records if x["id"] == rid), None)
    if not r:
        await cb.answer("Не найдено"); return
    icon = "💰" if r["type"] == "income" else CAT_ICONS.get(r.get("category","other"),"📦")
    sign = "+" if r["type"] == "income" else "-"
    await cb.message.edit_text(
        f"{icon} <b>{r.get('description') or CAT_NAMES.get(r.get('category','other'),'')}</b>\n\n"
        f"Сумма: <b>{sign}{int(r['amount'])} {r.get('currency','')}</b>\n"
        f"Категория: {CAT_NAMES.get(r.get('category','other'),'—')}\n"
        f"Дата: {fmt_date(r['created_at'])}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🗑 Удалить", callback_data=f"fdel_{rid}")],
            [InlineKeyboardButton(text="⬅️ К финансам", callback_data="finance")]
        ]), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("fdel_"))
async def cb_finance_delete(cb: CallbackQuery):
    db.delete_finance(cb.from_user.id, cb.data[5:])
    await cb.answer("Удалено ✓")
    uid = cb.from_user.id
    income, expense = get_finance_summary(uid)
    balance = income - expense
    await cb.message.edit_text(
        f"💰 <b>Финансы</b>\n\n📈 {int(income):,} / 📉 {int(expense):,}\n💼 {'+' if balance>=0 else ''}{int(balance):,}",
        reply_markup=kb_finance(uid), parse_mode="HTML"
    )

@router.callback_query(F.data == "fin_new")
async def cb_fin_new(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.fin_type)
    await cb.message.edit_text(
        "💰 <b>Новая запись</b>\n\nТип:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📉 Расход", callback_data="fnt_expense"),
             InlineKeyboardButton(text="📈 Доход", callback_data="fnt_income")]
        ]), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("fnt_"))
async def cb_fin_type(cb: CallbackQuery, state: FSMContext):
    await state.update_data(fin_type=cb.data[4:])
    await state.set_state(S.fin_amount)
    await cb.message.edit_text("Введи сумму:")

@router.message(S.fin_amount)
async def handle_fin_amount(msg: Message, state: FSMContext):
    try:
        amount = float(msg.text.strip().replace(",","."))
    except:
        await msg.answer("Введи число, например: 1500"); return
    await state.update_data(fin_amount=amount)
    await state.set_state(S.fin_currency)
    await msg.answer(
        "Валюта:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=c, callback_data=f"fc2_{c}") for c in CURRENCIES]
        ])
    )

@router.callback_query(F.data.startswith("fc2_"))
async def cb_fin_currency(cb: CallbackQuery, state: FSMContext):
    await state.update_data(fin_currency=cb.data[4:])
    data = await state.get_data()
    if data.get("fin_type") == "income":
        db.add_finance(cb.from_user.id, "income", data["fin_amount"], cb.data[4:], "income", "")
        await state.set_state(S.main)
        await cb.message.edit_text(
            f"✅ Доход: +{int(data['fin_amount'])} {cb.data[4:]}",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="💰 К финансам", callback_data="finance")],
                [InlineKeyboardButton(text="🏠 Меню", callback_data="main")]
            ])
        )
    else:
        await state.set_state(S.fin_category)
        cats = [("food","🍔 Еда"),("transport","🚗 Транспорт"),("business","💼 Бизнес"),
                ("health","💪 Здоровье"),("personal","👤 Личное"),("other","📦 Прочее")]
        rows = [[InlineKeyboardButton(text=name, callback_data=f"fcat_{key}")] for key,name in cats]
        await cb.message.edit_text("Категория:", reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))

@router.callback_query(F.data.startswith("fcat_"))
async def cb_fin_category(cb: CallbackQuery, state: FSMContext):
    await state.update_data(fin_category=cb.data[5:])
    await state.set_state(S.fin_desc)
    await cb.message.edit_text("Описание (или '-' пропустить):")

@router.message(S.fin_desc)
async def handle_fin_desc(msg: Message, state: FSMContext):
    data = await state.get_data()
    desc = msg.text.strip() if msg.text.strip() not in ("-","нет") else ""
    db.add_finance(msg.from_user.id, data["fin_type"], data["fin_amount"], data.get("fin_currency","KGS"), data.get("fin_category","other"), desc)
    await state.set_state(S.main)
    await msg.answer(
        f"✅ Расход: -{int(data['fin_amount'])} {data.get('fin_currency','KGS')}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="💰 К финансам", callback_data="finance")],
            [InlineKeyboardButton(text="🏠 Меню", callback_data="main")]
        ])
    )

# ══ MAIN ══
async def main():
    bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    log.info("✅ Трекер-бот запущен!")
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
