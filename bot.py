# -*- coding: utf-8 -*-
"""
Hisobchi — Telegram bot (aiogram v3)
- Foydalanuvchi /start bosadi
- Ro'yxatdan o'tmagan bo'lsa: raqamini yuborishi so'raladi (tugma orqali, faqat OWN contact)
- Ro'yxatdan o'tgach: "Hisobchini ochish" Web App tugmasi beriladi
- Ro'yxatdan o'tmasdan boshqa hech narsa qila olmaydi
"""
import asyncio
import logging
import os

from aiogram import Bot, Dispatcher, Router, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.types import (
    Message,
    CallbackQuery,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    WebAppInfo,
    MenuButtonWebApp,
)

from app import get_user, upsert_user, init_db, get_monthly_summary

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("hisobchi-bot")

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
WEBAPP_URL = os.environ.get("WEBAPP_URL", "")  # masalan: https://sizning-app.onrender.com

# MUHIM: faqat Bot obyekti har safar start_bot() chaqirilganda yangidan
# yaratiladi (chunki aiohttp sessiya aynan Bot ichida, joriy event loop'ga
# bog'langan holda saqlanadi — eski/yopilgan loop bilan qayta ishlatilsa
# "RuntimeError: Event loop is closed" beradi).
# Dispatcher va Router esa faqat BIR MARTA yaratiladi — chunki Router faqat
# bitta Dispatcher'ga bog'lanishi mumkin, qayta include_router() qilinsa
# "Router is already attached" xatosini beradi.
router = Router()
dp = Dispatcher()
dp.include_router(router)


def contact_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="📱 Raqamni yuborish", request_contact=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )


def webapp_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(
                text="📊 Hisobchini ochish",
                web_app=WebAppInfo(url=WEBAPP_URL),
            )
        ],
        [
            InlineKeyboardButton(
                text="📅 Oylik hisobot",
                callback_data="monthly_report",
            )
        ],
    ])


UZ_MONTHS = [
    "yanvar", "fevral", "mart", "aprel", "may", "iyun",
    "iyul", "avgust", "sentyabr", "oktyabr", "noyabr", "dekabr",
]


def format_monthly_report(summary: dict) -> str:
    import datetime
    now = datetime.datetime.now()
    month_name = UZ_MONTHS[now.month - 1]

    def fmt(n):
        return f"{n:,.0f}".replace(",", " ")

    lines = [f"📅 <b>{month_name.capitalize()} oyi uchun hisobot</b>\n"]
    lines.append(f"➕ Kirim: <b>{fmt(summary['income'])} so'm</b>")
    lines.append(f"➖ Xarajat: <b>{fmt(summary['expense'])} so'm</b>")
    lines.append(f"💰 Balans: <b>{fmt(summary['balance'])} so'm</b>")

    if summary["categories"]:
        lines.append("\n<b>Xarajatlar taqsimoti:</b>")
        total_expense = summary["expense"] or 1
        for cat, val in summary["categories"][:8]:
            pct = round(val / total_expense * 100)
            lines.append(f"• {cat} — {fmt(val)} so'm ({pct}%)")
    else:
        lines.append("\nBu oyda hali xarajat kiritilmagan.")

    return "\n".join(lines)


@router.message(CommandStart())
async def cmd_start(message: Message):
    user = get_user(message.from_user.id)
    if user:
        await message.answer(
            f"Salom, <b>{message.from_user.first_name}</b>! 👋\n\n"
            "Hisobchi tayyor — kirim va xarajatlaringizni boshqarish uchun "
            "quyidagi tugmani bosing.",
            reply_markup=webapp_keyboard(),
        )
        return

    await message.answer(
        "Assalomu alaykum! 👋\n\n"
        "<b>Hisobchi</b> botiga xush kelibsiz — bu bot orqali kirim va "
        "xarajatlaringizni qulay tarzda hisoblab borishingiz mumkin.\n\n"
        "Davom etish uchun avval telefon raqamingizni tasdiqlashingiz kerak. "
        "Pastdagi tugmani bosing 👇",
        reply_markup=contact_keyboard(),
    )


@router.message(F.contact)
async def on_contact(message: Message):
    contact = message.contact
    # Xavfsizlik: faqat o'zining raqamini qabul qilamiz, boshqa userning
    # kontaktini forward qilib yuborishning oldini olamiz
    if contact.user_id != message.from_user.id:
        await message.answer(
            "❗️ Iltimos, faqat <b>o'zingizning</b> raqamingizni yuboring.",
            reply_markup=contact_keyboard(),
        )
        return

    upsert_user(
        telegram_id=message.from_user.id,
        phone=contact.phone_number,
        first_name=message.from_user.first_name,
        last_name=message.from_user.last_name,
        username=message.from_user.username,
    )

    await message.answer(
        "✅ Ro'yxatdan muvaffaqiyatli o'tdingiz!",
        reply_markup=ReplyKeyboardRemove(),
    )
    await message.answer(
        "Endi Hisobchidan foydalanishingiz mumkin 👇",
        reply_markup=webapp_keyboard(),
    )


@router.message()
async def block_unregistered(message: Message):
    """Ro'yxatdan o'tmagan foydalanuvchi boshqa hech narsa yoza olmaydi —
    doim raqam so'raladi. Ro'yxatdan o'tganlar uchun esa webapp tugmasi
    qayta yuboriladi."""
    user = get_user(message.from_user.id)
    if user:
        await message.answer(
            "Hisobchini ochish uchun tugmani bosing 👇",
            reply_markup=webapp_keyboard(),
        )
    else:
        await message.answer(
            "Davom etish uchun avval telefon raqamingizni yuboring 👇",
            reply_markup=contact_keyboard(),
        )


@router.callback_query(F.data == "monthly_report")
async def on_monthly_report(callback: CallbackQuery):
    user = get_user(callback.from_user.id)
    if not user:
        await callback.answer("Avval ro'yxatdan o'ting", show_alert=True)
        return

    summary = get_monthly_summary(callback.from_user.id)
    text = format_monthly_report(summary)
    await callback.message.answer(text)
    await callback.answer()


async def _run_polling():
    """Har chaqirilganda TOZA Bot obyekti yaratadi — joriy event loop'ga
    bog'langan aiohttp sessiya bilan. Dispatcher/Router global (bir marta
    yaratilgan) qoladi. Shu tufayli qayta urinishlarda eski (yopilgan)
    loop'ga bog'langan sessiya ishlatilmaydi."""
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    try:
        await bot.delete_webhook(drop_pending_updates=True)

        # Doimiy menyu tugmasi — xabar yozish maydonining oldida, botning
        # pastki chap burchagida doim ko'rinib turadi. Bosilganda /start
        # bosmasdan ham to'g'ridan-to'g'ri saytni ochadi. Sayt o'zi
        # initData va ro'yxatdan o'tganlikni tekshiradi, shuning uchun
        # ro'yxatdan o'tmagan foydalanuvchi baribir "avval ro'yxatdan
        # o'ting" ekranini ko'radi.
        await bot.set_chat_menu_button(
            menu_button=MenuButtonWebApp(
                text="📊 Hisobchi",
                web_app=WebAppInfo(url=WEBAPP_URL),
            )
        )

        # handle_signals=False: bot alohida threadda ishlaydi (asosiy thread
        # emas), shuning uchun aiogram SIGINT/SIGTERM ushlagichlarini
        # o'rnatishga urinmasligi kerak (bu faqat asosiy threadda ishlaydi).
        await dp.start_polling(bot, handle_signals=False)
    finally:
        await bot.session.close()


def start_bot():
    """Alohida thread ichidan chaqiriladi (main.py orqali)."""
    init_db()
    asyncio.run(_run_polling())


if __name__ == "__main__":
    start_bot()
