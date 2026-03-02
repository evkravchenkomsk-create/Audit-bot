"""
Telegram-бот: Аудит бизнеса перед партнёрским входом
"""
import asyncio
import logging
import os
import sys
from datetime import datetime

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, CallbackQuery,
    InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
)

from data import AUDIT_BLOCKS, STOP_FACTORS
from storage import AuditStorage
from report_generator import generate_excel_report, generate_text_report

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

TOKEN = os.getenv("BOT_TOKEN", "")
router = Router()
db = AuditStorage("audit_data")

class S(StatesGroup):
    menu           = State()
    entering_name  = State()
    in_audit       = State()
    adding_comment = State()

SCORE_EMOJI = ["❌", "1️⃣", "2️⃣", "3️⃣", "4️⃣", "⭐"]
SCORE_LABEL = ["Нет/критично", "Очень слабо", "Слабо", "Средне", "Хорошо", "Отлично"]

def bar(score, mx, w=12):
    if mx == 0: return "░" * w
    filled = round(score / mx * w)
    return "█" * filled + "░" * (w - filled)

def pct(score, mx):
    return round(score / mx * 100) if mx else 0

def get_decision(total, mx):
    p = pct(total, mx)
    if   p >= 80: return "🟢 GO — ВХОДИТЬ",   "Высокий потенциал. Переговоры немедленно."
    elif p >= 60: return "🟡 GO + УСЛОВИЕ",    "Потенциал есть. Пилот 90 дней с KPI."
    elif p >= 40: return "🟠 ОСТОРОЖНО",       "Серьёзные дыры. Жёсткие условия договора."
    else:         return "🔴 NO GO",           "Высокий риск. Отказ или аудит через 6 мес."

def kb_main():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📋 Новый аудит",  callback_data="new")],
        [InlineKeyboardButton(text="📂 Мои аудиты",   callback_data="list")],
        [InlineKeyboardButton(text="❓ Как работает", callback_data="help")],
    ])

def kb_blocks(aid, completed):
    rows = []
    for i, bl in enumerate(AUDIT_BLOCKS):
        ans   = db.get_block_answers(aid, i)
        done  = len(ans)
        total = len(bl["questions"])
        sc    = sum(a.get("score", 0) for a in ans.values())
        icon  = "✅" if i in completed else ("🔄" if done else "⬜")
        rows.append([InlineKeyboardButton(
            text=f"{icon} {bl['short']}  {sc}/{bl['max']}  ({done}/{total})",
            callback_data=f"blk_{aid}_{i}"
        )])
    rows.append([InlineKeyboardButton(text="📊 Итоги + отчёт", callback_data=f"res_{aid}")])
    rows.append([InlineKeyboardButton(text="🏠 Меню",           callback_data="menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

def kb_questions(aid, bidx):
    block = AUDIT_BLOCKS[bidx]
    ans   = db.get_block_answers(aid, bidx)
    rows  = []
    for i, q in enumerate(block["questions"]):
        a = ans.get(str(i))
        if a:
            sc   = a.get("score", 0)
            flag = "💬" if a.get("comment") else ""
            txt  = f"{SCORE_EMOJI[sc]}{flag} {q['title'][:40]}"
        else:
            txt = f"⬜ {q['title'][:42]}"
        rows.append([InlineKeyboardButton(text=txt[:55], callback_data=f"q_{aid}_{bidx}_{i}")])
    rows.append([InlineKeyboardButton(text="⬅️ К блокам", callback_data=f"blks_{aid}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

def kb_score(aid, bidx, qidx):
    btns = [InlineKeyboardButton(text=SCORE_EMOJI[s], callback_data=f"sc_{aid}_{bidx}_{qidx}_{s}") for s in range(6)]
    return InlineKeyboardMarkup(inline_keyboard=[
        btns[:3], btns[3:],
        [InlineKeyboardButton(text="💬 Добавить комментарий", callback_data=f"cmt_{aid}_{bidx}_{qidx}")],
        [InlineKeyboardButton(text="➡️ Следующий",            callback_data=f"nxt_{aid}_{bidx}_{qidx}")],
        [InlineKeyboardButton(text="⬅️ К вопросам блока",    callback_data=f"blk_{aid}_{bidx}")],
    ])

def kb_after_comment(aid, bidx, qidx):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➡️ Следующий вопрос",  callback_data=f"nxt_{aid}_{bidx}_{qidx}")],
        [InlineKeyboardButton(text="⬅️ К вопросам блока", callback_data=f"blk_{aid}_{bidx}")],
    ])

def kb_results(aid):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📥 Скачать Excel-отчёт", callback_data=f"xlsx_{aid}")],
        [InlineKeyboardButton(text="📋 К блокам",             callback_data=f"blks_{aid}")],
        [InlineKeyboardButton(text="🏠 Меню",                 callback_data="menu")],
    ])

def kb_audits(audits):
    rows = []
    for a in audits[-8:]:
        dt  = datetime.fromisoformat(a["created_at"]).strftime("%d.%m %H:%M")
        co  = a.get("company", "?")[:22]
        rows.append([InlineKeyboardButton(text=f"🏢 {co}  ({dt})", callback_data=f"open_{a['id']}")])
    rows.append([InlineKeyboardButton(text="🏠 Меню", callback_data="menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

@router.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await state.set_state(S.menu)
    name = msg.from_user.first_name or "партнёр"
    await msg.answer(
        f"👋 <b>Привет, {name}!</b>\n\n"
        "Бот для аудита бизнеса перед партнёрским входом.\n\n"
        "<b>6 блоков · 116 вопросов · 600 баллов</b>\n"
        "По итогам — решение GO / NO GO и Excel-отчёт.\n\n"
        "📊 Финансы — 120 б.\n"
        "💼 Продажи — 140 б.\n"
        "📣 Маркетинг — 100 б.\n"
        "⚙️ Операции и команда — 100 б.\n"
        "👤 Собственник — 80 б.\n"
        "🌍 Рынок — 60 б.",
        reply_markup=kb_main(), parse_mode="HTML"
    )

@router.callback_query(F.data == "menu")
async def cb_menu(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.menu)
    await cb.message.edit_text("🏠 <b>Главное меню</b>", reply_markup=kb_main(), parse_mode="HTML")

@router.callback_query(F.data == "help")
async def cb_help(cb: CallbackQuery):
    await cb.message.edit_text(
        "❓ <b>Как пользоваться</b>\n\n"
        "1️⃣ <b>Новый аудит</b> → введите название компании\n"
        "2️⃣ Выберите блок аудита\n"
        "3️⃣ Нажмите на вопрос → поставьте оценку <b>0–5</b>\n"
        "4️⃣ Добавьте комментарий с фактами (опционально)\n"
        "5️⃣ Пройдите все блоки → <b>«Итоги + отчёт»</b>\n\n"
        "Шкала оценок:\n"
        "❌ 0 — Нет / критично\n"
        "1️⃣ — Очень слабо\n"
        "2️⃣ — Слабо\n"
        "3️⃣ — Средне\n"
        "4️⃣ — Хорошо\n"
        "⭐ 5 — Отлично\n\n"
        "💡 Блоки проходите в любом порядке и возвращайтесь.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🏠 Меню", callback_data="menu")]
        ]), parse_mode="HTML"
    )

@router.callback_query(F.data == "new")
async def cb_new(cb: CallbackQuery, state: FSMContext):
    await state.set_state(S.entering_name)
    await cb.message.edit_text("🏢 <b>Новый аудит</b>\n\nВведите название компании:", parse_mode="HTML")

@router.message(S.entering_name)
async def handle_name(msg: Message, state: FSMContext):
    company = msg.text.strip()
    if len(company) < 2:
        await msg.answer("Введите корректное название (минимум 2 символа):")
        return
    aid = db.create_audit(msg.from_user.id, company)
    await state.set_state(S.in_audit)
    await msg.answer(
        f"✅ Аудит создан!\n🏢 <b>{company}</b>\n\nВыберите блок:",
        reply_markup=kb_blocks(aid, []), parse_mode="HTML"
    )

@router.callback_query(F.data == "list")
async def cb_list(cb: CallbackQuery):
    audits = db.get_user_audits(cb.from_user.id)
    if not audits:
        await cb.message.edit_text(
            "📂 У вас пока нет аудитов.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📋 Новый аудит", callback_data="new")],
                [InlineKeyboardButton(text="🏠 Меню",        callback_data="menu")],
            ])
        )
        return
    await cb.message.edit_text(
        f"📂 <b>Ваши аудиты</b> ({len(audits)} шт.)",
        reply_markup=kb_audits(audits), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("open_"))
async def cb_open(cb: CallbackQuery, state: FSMContext):
    aid   = cb.data.removeprefix("open_")
    audit = db.get_audit(aid)
    if not audit:
        await cb.answer("Аудит не найден", show_alert=True); return
    completed = db.get_completed_blocks(aid)
    total, mx = db.get_total_score(aid)
    co = audit.get("company", "?")
    await state.set_state(S.in_audit)
    await cb.message.edit_text(
        f"🏢 <b>{co}</b>\n"
        f"📊 {total}/{mx} баллов ({pct(total, mx)}%)\n"
        f"Блоков пройдено: {len(completed)}/6\n\nВыберите блок:",
        reply_markup=kb_blocks(aid, completed), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("blks_"))
async def cb_blks(cb: CallbackQuery):
    aid   = cb.data.removeprefix("blks_")
    audit = db.get_audit(aid)
    if not audit:
        await cb.answer("Аудит не найден", show_alert=True); return
    completed = db.get_completed_blocks(aid)
    total, mx = db.get_total_score(aid)
    co = audit.get("company", "?")
    await cb.message.edit_text(
        f"🏢 <b>{co}</b>\n"
        f"📊 {total}/{mx} ({pct(total, mx)}%) [{bar(total, mx, 14)}]\n\nВыберите блок:",
        reply_markup=kb_blocks(aid, completed), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("blk_"))
async def cb_blk(cb: CallbackQuery):
    parts = cb.data.split("_")
    aid   = parts[1]
    bidx  = int(parts[2])
    block = AUDIT_BLOCKS[bidx]
    ans   = db.get_block_answers(aid, bidx)
    sc    = sum(a.get("score", 0) for a in ans.values())
    done  = len(ans)
    total = len(block["questions"])
    await cb.message.edit_text(
        f"<b>{block['title']}</b>\n"
        f"Отвечено: {done}/{total}   Баллов: {sc}/{block['max']}\n"
        f"[{bar(sc, block['max'], 14)}]\n\n"
        f"<i>{block['description']}</i>\n\n"
        f"Нажмите на вопрос:",
        reply_markup=kb_questions(aid, bidx), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("q_"))
async def cb_q(cb: CallbackQuery):
    parts = cb.data.split("_")
    aid   = parts[1]
    bidx  = int(parts[2])
    qidx  = int(parts[3])
    block = AUDIT_BLOCKS[bidx]
    q     = block["questions"][qidx]
    ans   = db.get_block_answers(aid, bidx)
    cur   = ans.get(str(qidx), {})

    current = ""
    if cur:
        sc = cur.get("score", 0)
        current = f"\n\n✏️ <b>Текущая оценка:</b> {SCORE_EMOJI[sc]} {SCORE_LABEL[sc]}"
        if cur.get("comment"):
            current += f"\n💬 <i>{cur['comment'][:150]}</i>"

    await cb.message.edit_text(
        f"<b>{block['short']}</b> · Вопрос {qidx + 1}/{len(block['questions'])}\n"
        f"{'─' * 32}\n\n"
        f"<b>{q['title']}</b>\n\n"
        f"🔍 <i>{q['how']}</i>"
        f"{current}\n\n"
        f"Поставьте оценку:",
        reply_markup=kb_score(aid, bidx, qidx), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("sc_"))
async def cb_sc(cb: CallbackQuery):
    parts = cb.data.split("_")
    aid   = parts[1]
    bidx  = int(parts[2])
    qidx  = int(parts[3])
    score = int(parts[4])
    db.save_answer(aid, bidx, qidx, score)
    block = AUDIT_BLOCKS[bidx]
    q     = block["questions"][qidx]
    cur   = db.get_block_answers(aid, bidx).get(str(qidx), {})
    cline = f"\n💬 <i>{cur['comment'][:150]}</i>" if cur.get("comment") else ""
    await cb.message.edit_text(
        f"<b>{block['short']}</b> · Вопрос {qidx + 1}/{len(block['questions'])}\n"
        f"{'─' * 32}\n\n"
        f"<b>{q['title']}</b>\n\n"
        f"✅ Оценка: {SCORE_EMOJI[score]} <b>{SCORE_LABEL[score]}</b>{cline}",
        reply_markup=kb_score(aid, bidx, qidx), parse_mode="HTML"
    )
    await cb.answer(f"Сохранено: {score}/5 ✓")

@router.callback_query(F.data.startswith("cmt_"))
async def cb_cmt(cb: CallbackQuery, state: FSMContext):
    parts = cb.data.split("_")
    aid   = parts[1]
    bidx  = int(parts[2])
    qidx  = int(parts[3])
    q     = AUDIT_BLOCKS[bidx]["questions"][qidx]
    await state.set_state(S.adding_comment)
    await state.update_data(aid=aid, bidx=bidx, qidx=qidx)
    await cb.message.edit_text(
        f"💬 <b>Комментарий к вопросу</b>\n\n"
        f"<b>{q['title']}</b>\n\n"
        f"Введите факты, цифры, наблюдения:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⏭ Пропустить", callback_data=f"nxt_{aid}_{bidx}_{qidx}")]
        ]),
        parse_mode="HTML"
    )

@router.message(S.adding_comment)
async def handle_comment(msg: Message, state: FSMContext):
    data    = await state.get_data()
    aid, bidx, qidx = data["aid"], data["bidx"], data["qidx"]
    comment = msg.text.strip()
    db.save_comment(aid, bidx, qidx, comment)
    await state.set_state(S.in_audit)
    cur   = db.get_block_answers(aid, bidx).get(str(qidx), {})
    sc    = cur.get("score", 0)
    q     = AUDIT_BLOCKS[bidx]["questions"][qidx]
    await msg.answer(
        f"💬 Комментарий сохранён!\n\n"
        f"<b>{q['title']}</b>\n"
        f"Оценка: {SCORE_EMOJI[sc]} {SCORE_LABEL[sc]}\n"
        f"Факты: <i>{comment[:200]}</i>",
        reply_markup=kb_after_comment(aid, bidx, qidx), parse_mode="HTML"
    )

@router.callback_query(F.data.startswith("nxt_"))
async def cb_nxt(cb: CallbackQuery, state: FSMContext):
    parts = cb.data.split("_")
    aid   = parts[1]
    bidx  = int(parts[2])
    qidx  = int(parts[3])
    block = AUDIT_BLOCKS[bidx]
    nxt   = qidx + 1
    if nxt < len(block["questions"]):
        cb.data = f"q_{aid}_{bidx}_{nxt}"
        await state.set_state(S.in_audit)
        await cb_q(cb)
    else:
        ans = db.get_block_answers(aid, bidx)
        sc  = sum(a.get("score", 0) for a in ans.values())
        p   = pct(sc, block["max"])
        await cb.message.edit_text(
            f"🎉 <b>Блок завершён!</b>\n\n"
            f"{block['title']}\n"
            f"Результат: <b>{sc}/{block['max']} ({p}%)</b>\n"
            f"[{bar(sc, block['max'], 16)}]",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📋 К блокам",         callback_data=f"blks_{aid}")],
                [InlineKeyboardButton(text="📊 Посмотреть итоги", callback_data=f"res_{aid}")],
            ]),
            parse_mode="HTML"
        )

@router.callback_query(F.data.startswith("res_"))
async def cb_res(cb: CallbackQuery):
    aid   = cb.data.removeprefix("res_")
    audit = db.get_audit(aid)
    if not audit:
        await cb.answer("Аудит не найден", show_alert=True); return
    co          = audit.get("company", "?")
    total, mx   = db.get_total_score(aid)
    p           = pct(total, mx)
    dec, advice = get_decision(total, mx)
    lines = [
        f"📊 <b>РЕЗУЛЬТАТЫ АУДИТА</b>",
        f"🏢 <b>{co}</b>",
        f"{'─' * 32}", "",
    ]
    for i, block in enumerate(AUDIT_BLOCKS):
        ans  = db.get_block_answers(aid, i)
        sc   = sum(a.get("score", 0) for a in ans.values())
        bp   = pct(sc, block["max"])
        done = len(ans)
        tq   = len(block["questions"])
        icon = "🟢" if bp >= 70 else ("🟡" if bp >= 50 else ("🟠" if bp >= 30 and done else ("⬜" if not done else "🔴")))
        lines.append(f"{icon} <b>{block['short']}</b>: {sc}/{block['max']} ({bp}%)")
        lines.append(f"    [{bar(sc, block['max'], 10)}] {done}/{tq} отвечено")
    lines += [
        "", f"{'─' * 32}",
        f"🎯 <b>ИТОГО: {total}/{mx} ({p}%)</b>",
        f"[{bar(total, mx, 18)}]",
        "", f"<b>{dec}</b>",
        f"💡 {advice}",
    ]
    await cb.message.edit_text("\n".join(lines), reply_markup=kb_results(aid), parse_mode="HTML")

@router.callback_query(F.data.startswith("xlsx_"))
async def cb_xlsx(cb: CallbackQuery):
    aid   = cb.data.removeprefix("xlsx_")
    audit = db.get_audit(aid)
    if not audit:
        await cb.answer("Аудит не найден", show_alert=True); return
    await cb.answer("⏳ Генерирую отчёт...")
    co = audit.get("company", "Компания")
    try:
        filepath = await generate_excel_report(aid, audit, db)
        fname    = f"Аудит_{co[:20].replace(' ', '_')}.xlsx"
        await cb.message.answer_document(
            FSInputFile(filepath, filename=fname),
            caption=(
                f"📊 <b>Excel-отчёт готов!</b>\n"
                f"🏢 {co}\n"
                f"📅 {datetime.now().strftime('%d.%m.%Y')}\n\n"
                f"В файле: итоговый дашборд, детали по каждому блоку,\n"
                f"стоп-факторы и матрица быстрых побед."
            ),
            parse_mode="HTML"
        )
        os.remove(filepath)
    except Exception as e:
        log.error(f"Excel error: {e}")
        report = generate_text_report(aid, audit, db)
        await cb.message.answer(f"⚠️ Excel недоступен:\n\n{report}", parse_mode="HTML")

async def main():
    bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp  = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    log.info("✅ Бот запущен! Ожидаю сообщений...")
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
