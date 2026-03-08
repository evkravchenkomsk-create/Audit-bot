"""
EVK Business Bot — система управления бизнесом
Авторизация + отчёты по проектам из AmoCRM
"""
import asyncio
import logging
import os
from datetime import datetime

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, CallbackQuery,
    InlineKeyboardMarkup, InlineKeyboardButton
)

import database as db
from config import BOT_TOKEN, OWNER_ID, PROJECTS
from amo import get_project_report

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

router = Router()


class S(StatesGroup):
    waiting_note = State()


# --- КЛАВИАТУРЫ ---

def kb_request_access():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📋 Запросить доступ", callback_data="request_access")]
    ])


def kb_report_period(project_key: str):
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="📅 Сегодня",  callback_data=f"rpt_{project_key}_1"),
            InlineKeyboardButton(text="📅 7 дней",   callback_data=f"rpt_{project_key}_7"),
        ],
        [
            InlineKeyboardButton(text="📅 30 дней",  callback_data=f"rpt_{project_key}_30"),
        ],
    ])


def kb_projects(user_id: int):
    """Клавиатура с проектами пользователя"""
    projects = db.get_user_projects(user_id)
    rows = []
    for key in projects:
        cfg = PROJECTS.get(key)
        if cfg:
            rows.append([InlineKeyboardButton(
                text=f"{cfg['emoji']} {cfg['name']}",
                callback_data=f"project_{key}"
            )])
    rows.append([InlineKeyboardButton(text="📊 Все проекты сразу", callback_data="report_all")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def kb_approve(user_id: int):
    """Клавиатура для владельца: принять/отклонить"""
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="✅ Принять",   callback_data=f"approve_{user_id}"),
            InlineKeyboardButton(text="❌ Отклонить", callback_data=f"reject_{user_id}"),
        ]
    ])


def kb_select_projects(user_id: int, selected: list = None):
    """Выбор проектов при одобрении"""
    selected = selected or []
    rows = []
    for key, cfg in PROJECTS.items():
        check = "✅" if key in selected else "⬜"
        rows.append([InlineKeyboardButton(
            text=f"{check} {cfg['emoji']} {cfg['name']}",
            callback_data=f"toggleproj_{user_id}_{key}_{'|'.join(selected)}"
        )])
    rows.append([InlineKeyboardButton(
        text="💾 Сохранить и одобрить",
        callback_data=f"saveapprove_{user_id}_{'|'.join(selected)}"
    )])
    return InlineKeyboardMarkup(inline_keyboard=rows)


# --- СТАРТОВЫЙ ЭКРАН ---

@router.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await state.clear()
    user_id = msg.from_user.id

    if db.is_owner(user_id):
        db.approve_user(user_id, list(PROJECTS.keys()))
        await show_main_menu(msg, user_id)
        return

    user = db.get_user(user_id)

    if user and user["status"] == "approved":
        await show_main_menu(msg, user_id)
    elif user and user["status"] == "pending":
        await msg.answer(
            "⏳ <b>Ваш запрос на рассмотрении</b>\n\n"
            "Ожидайте — руководитель получил уведомление и скоро даст доступ.",
        )
    elif user and user["status"] == "rejected":
        await msg.answer(
            "❌ <b>В доступе отказано</b>\n\n"
            "Если вы считаете это ошибкой — свяжитесь с руководителем.",
        )
    else:
        name = msg.from_user.full_name or msg.from_user.first_name or "Пользователь"
        await msg.answer(
            f"👋 <b>Добро пожаловать, {name}!</b>\n\n"
            "Это система управления бизнесом EVK.\n\n"
            "Для получения доступа нажмите кнопку ниже и кратко опишите свою должность.",
            reply_markup=kb_request_access()
        )


async def show_main_menu(msg_or_cb, user_id: int):
    projects = db.get_user_projects(user_id)
    is_own = db.is_owner(user_id)

    project_names = [f"{PROJECTS[p]['emoji']} {PROJECTS[p]['name']}" for p in projects if p in PROJECTS]

    text = (
        "🏠 <b>Главное меню</b>\n\n"
        f"Ваши проекты: {', '.join(project_names) if project_names else '—'}\n\n"
        "Выберите проект для отчёта:"
    )

    rows = []
    for key in projects:
        cfg = PROJECTS.get(key)
        if cfg:
            rows.append([InlineKeyboardButton(
                text=f"{cfg['emoji']} {cfg['name']}",
                callback_data=f"project_{key}"
            )])

    if len(projects) > 1:
        rows.append([InlineKeyboardButton(text="📊 Все проекты", callback_data="report_all")])

    if is_own:
        rows.append([InlineKeyboardButton(text="👥 Управление доступами", callback_data="manage_users")])

    kb = InlineKeyboardMarkup(inline_keyboard=rows)

    if isinstance(msg_or_cb, Message):
        await msg_or_cb.answer(text, reply_markup=kb)
    else:
        await msg_or_cb.message.edit_text(text, reply_markup=kb)


# --- ЗАПРОС ДОСТУПА ---

@router.callback_query(F.data == "request_access")
async def cb_request(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.waiting_note)
    await cb.message.edit_text(
        "📋 <b>Запрос доступа</b>\n\n"
        "Введите вашу должность и к какому проекту нужен доступ:\n\n"
        "<i>Пример: Руководитель Filter.kg</i>"
    )


@router.message(S.waiting_note)
async def handle_note(msg: Message, state: FSMContext, bot: Bot):
    await state.clear()
    user = msg.from_user
    note = msg.text.strip()

    db.create_user(user.id, user.username, user.full_name, note)

    # Уведомление владельцу
    uname = f"@{user.username}" if user.username else "нет username"
    await bot.send_message(
        OWNER_ID,
        f"🔔 <b>Новый запрос доступа</b>\n\n"
        f"👤 {user.full_name} ({uname})\n"
        f"🆔 ID: <code>{user.id}</code>\n"
        f"📝 {note}\n\n"
        f"Выдать доступ?",
        reply_markup=kb_approve(user.id)
    )

    await msg.answer(
        "✅ <b>Запрос отправлен!</b>\n\n"
        "Руководитель получил уведомление. "
        "Как только доступ будет выдан — вы получите сообщение."
    )


# --- ОДОБРЕНИЕ / ОТКЛОНЕНИЕ ---

@router.callback_query(F.data.startswith("approve_"))
async def cb_approve(cb: CallbackQuery):
    if not db.is_owner(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return

    user_id = int(cb.data.split("_")[1])
    user = db.get_user(user_id)
    if not user:
        await cb.answer("Пользователь не найден", show_alert=True)
        return

    await cb.message.edit_text(
        f"✅ <b>Выбор проектов для {user['full_name']}</b>\n\n"
        f"Отметьте проекты к которым дать доступ:",
        reply_markup=kb_select_projects(user_id, [])
    )


@router.callback_query(F.data.startswith("toggleproj_"))
async def cb_toggle_project(cb: CallbackQuery):
    parts = cb.data.split("_", 3)
    user_id = int(parts[1])
    project_key = parts[2]
    current = parts[3].split("|") if parts[3] else []

    if project_key in current:
        current.remove(project_key)
    else:
        current.append(project_key)

    current = [p for p in current if p]

    user = db.get_user(user_id)
    await cb.message.edit_text(
        f"✅ <b>Выбор проектов для {user['full_name'] if user else user_id}</b>\n\n"
        f"Выбрано: {len(current)} проект(ов)\n"
        f"Отметьте проекты к которым дать доступ:",
        reply_markup=kb_select_projects(user_id, current)
    )


@router.callback_query(F.data.startswith("saveapprove_"))
async def cb_save_approve(cb: CallbackQuery, bot: Bot):
    parts = cb.data.split("_", 2)
    user_id = int(parts[1])
    projects = [p for p in parts[2].split("|") if p] if parts[2] else []

    if not projects:
        await cb.answer("Выберите хотя бы один проект", show_alert=True)
        return

    db.approve_user(user_id, projects)

    user = db.get_user(user_id)
    project_names = [f"{PROJECTS[p]['emoji']} {PROJECTS[p]['name']}" for p in projects if p in PROJECTS]

    await cb.message.edit_text(
        f"✅ Пользователь <b>{user['full_name'] if user else user_id}</b> одобрен\n"
        f"Проекты: {', '.join(project_names)}"
    )

    try:
        await bot.send_message(
            user_id,
            f"🎉 <b>Доступ открыт!</b>\n\n"
            f"Вам выдан доступ к проектам:\n"
            f"{chr(10).join(project_names)}\n\n"
            f"Нажмите /start для начала работы."
        )
    except Exception:
        pass


@router.callback_query(F.data.startswith("reject_"))
async def cb_reject(cb: CallbackQuery, bot: Bot):
    if not db.is_owner(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return

    user_id = int(cb.data.split("_")[1])
    db.reject_user(user_id)

    user = db.get_user(user_id)
    await cb.message.edit_text(f"❌ Пользователь {user['full_name'] if user else user_id} — отказано")

    try:
        await bot.send_message(user_id, "❌ <b>В доступе отказано.</b>\nСвяжитесь с руководителем для уточнений.")
    except Exception:
        pass


# --- ОТЧЁТЫ ---

@router.callback_query(F.data.startswith("project_"))
async def cb_project(cb: CallbackQuery):
    project_key = cb.data.removeprefix("project_")
    user_projects = db.get_user_projects(cb.from_user.id)

    if project_key not in user_projects and not db.is_owner(cb.from_user.id):
        await cb.answer("Нет доступа к этому проекту", show_alert=True)
        return

    cfg = PROJECTS.get(project_key)
    await cb.message.edit_text(
        f"{cfg['emoji']} <b>{cfg['name']}</b>\n\nВыберите период:",
        reply_markup=kb_report_period(project_key)
    )


@router.callback_query(F.data.startswith("rpt_"))
async def cb_report(cb: CallbackQuery, bot: Bot):
    parts = cb.data.split("_")
    project_key = parts[1]
    days = int(parts[2])

    user_projects = db.get_user_projects(cb.from_user.id)
    if project_key not in user_projects and not db.is_owner(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return

    cfg = PROJECTS.get(project_key)
    await cb.message.edit_text(f"⏳ Загружаю отчёт {cfg['emoji']} {cfg['name']}...")

    report = await get_project_report(cfg, days=days)
    period = "сегодня" if days == 1 else f"за {days} дней"

    await cb.message.edit_text(
        report,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text="🔄 Обновить", callback_data=cb.data),
                InlineKeyboardButton(text="🏠 Меню",     callback_data="menu"),
            ]
        ])
    )


@router.callback_query(F.data == "report_all")
async def cb_report_all(cb: CallbackQuery):
    user_projects = db.get_user_projects(cb.from_user.id)
    if not user_projects:
        await cb.answer("Нет доступных проектов", show_alert=True)
        return

    await cb.message.edit_text("⏳ Загружаю отчёты по всем проектам...")

    reports = []
    for key in user_projects:
        cfg = PROJECTS.get(key)
        if cfg:
            r = await get_project_report(cfg, days=1)
            reports.append(r)

    text = "\n\n".join(reports) if reports else "Нет данных"
    if len(text) > 4000:
        text = text[:4000] + "\n\n... (обрезано)"

    await cb.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔄 Обновить", callback_data="report_all")],
            [InlineKeyboardButton(text="🏠 Меню",     callback_data="menu")],
        ])
    )


@router.callback_query(F.data == "menu")
async def cb_menu(cb: CallbackQuery, state: FSMContext):
    await state.clear()
    await show_main_menu(cb, cb.from_user.id)


# --- УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ (только owner) ---

@router.callback_query(F.data == "manage_users")
async def cb_manage_users(cb: CallbackQuery):
    if not db.is_owner(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return

    approved = db.get_approved_users()
    pending = db.get_pending_users()

    lines = ["👥 <b>Управление доступами</b>\n"]

    if pending:
        lines.append(f"⏳ <b>Ожидают ({len(pending)}):</b>")
        for u in pending:
            uname = f"@{u['username']}" if u['username'] else "нет username"
            lines.append(f"  • {u['full_name']} ({uname}) — {u.get('note', '—')}")
        lines.append("")

    if approved:
        lines.append(f"✅ <b>Одобрены ({len(approved)}):</b>")
        for u in approved:
            projs = db.get_user_projects(u["telegram_id"])
            proj_names = [PROJECTS[p]["name"] for p in projs if p in PROJECTS]
            lines.append(f"  • {u['full_name']}: {', '.join(proj_names) or '—'}")

    rows = []
    for u in pending:
        rows.append([InlineKeyboardButton(
            text=f"✅ Принять {u['full_name'][:20]}",
            callback_data=f"approve_{u['telegram_id']}"
        )])
    rows.append([InlineKeyboardButton(text="🏠 Меню", callback_data="menu")])

    await cb.message.edit_text(
        "\n".join(lines),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows)
    )


# --- КОМАНДЫ ---

@router.message(Command("report"))
async def cmd_report(msg: Message):
    if not db.is_approved(msg.from_user.id) and not db.is_owner(msg.from_user.id):
        await msg.answer("❌ Нет доступа. Нажмите /start для запроса.")
        return

    projects = db.get_user_projects(msg.from_user.id)
    if not projects:
        await msg.answer("У вас нет доступных проектов.")
        return

    await msg.answer("Выберите проект:", reply_markup=kb_projects(msg.from_user.id))


@router.message(Command("users"))
async def cmd_users(msg: Message):
    if not db.is_owner(msg.from_user.id):
        return
    await msg.answer("👥 Управление:", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="👥 Управление доступами", callback_data="manage_users")]
    ]))


# --- ЕЖЕДНЕВНЫЙ ОТЧЁТ ---

async def send_daily_reports(bot: Bot):
    """Отправляет ежедневные отчёты всем авторизованным пользователям"""
    log.info("Отправка ежедневных отчётов...")
    users = db.get_approved_users()

    for user in users:
        projects = db.get_user_projects(user["telegram_id"])
        if not projects:
            continue

        reports = []
        for key in projects:
            cfg = PROJECTS.get(key)
            if cfg:
                r = await get_project_report(cfg, days=1)
                reports.append(r)

        if not reports:
            continue

        text = f"🌅 <b>Ежедневный отчёт — {datetime.now().strftime('%d.%m.%Y')}</b>\n\n"
        text += "\n\n".join(reports)

        if len(text) > 4000:
            text = text[:4000] + "\n\n..."

        try:
            await bot.send_message(user["telegram_id"], text)
        except Exception as e:
            log.error(f"Ошибка отправки отчёта {user['telegram_id']}: {e}")

        await asyncio.sleep(0.5)


async def daily_scheduler(bot: Bot):
    from config import REPORT_HOUR_UTC
    while True:
        now = datetime.utcnow()
        if now.hour == REPORT_HOUR_UTC and now.minute == 0:
            await send_daily_reports(bot)
            await asyncio.sleep(60)
        await asyncio.sleep(30)


# --- ЗАПУСК ---

async def main():
    db.init_db()

    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    asyncio.create_task(daily_scheduler(bot))

    log.info("✅ EVK Business Bot запущен!")
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
