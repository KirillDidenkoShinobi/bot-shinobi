import asyncio
import logging
import json
import os
import time
import random
import string
import re
import difflib
import html
import aiohttp
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F, BaseMiddleware
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message,
    ChatMemberUpdated,
    ChatJoinRequest,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    CallbackQuery,
    FSInputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
    WebAppInfo,
    MessageReactionUpdated,
    ChatPermissions,
)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

# PostgreSQL — постоянное хранилище данных Railway.
# Восстановление выполняется ДО загрузки data.json/Excel модулей.
from db_storage import restore_or_seed, persistence_loop
restore_or_seed()

from stats_api import create_report_snapshot, start_stats_api
from chat_summary import record_message, snapshot as chat_stats_snapshot, random_fact as mlbb_random_fact

from schedule_system import (
    init_schedule_system,
    schedule_router,
    show_schedule_menu,
)


# =========================================================
# НАСТРОЙКИ
# =========================================================

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не задан. Создайте .env по примеру .env.example")

# ID группы Shinobi Team
CHAT_ID = -1004368701081

# ID администраторов, которым доступна админ-панель
OWNER_IDS = [1197818066]
ADMIN_IDS = sorted(set(OWNER_IDS + [int(x.strip()) for x in os.getenv("SHINOBI_ADMIN_IDS", "").split(",") if x.strip().isdigit()]))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.json")
EXCEL_FILE = os.path.join(BASE_DIR, "shinobi_links_log.xlsx")
ADVERTISING_LINK = "https://t.me/+xXGHYqokIiMzNjYy"  # legacy default, do not overwrite persisted link
ADMIN_STAFF_CHAT_ID = -1004462686385
ERROR_DELETE_AFTER = 15
WEBAPP_URL = os.getenv("SHINOBI_WEBAPP_URL", "").strip()

# =========================================================
# SERVER TAGS (Telegram Bot API 9.5+)
# =========================================================
SERVER_TAG_PENDING_SECONDS = 20 * 60
SERVER_TAG_CONFIRM_DELETE_SECONDS = 20 * 60
SERVER_TAG_FILE = os.path.join(BASE_DIR, "shinobi_server_tags.txt")
SERVER_TAGS = [
    # Старые теги — не удаляем.
    "Ру серв", "Крд серв", "Мск серв", "Еу серв", "Узб серв", "Екб серв",
    "Укр серв", "Сиб серв", "Дв серв", "Вн серв", "Еу/Ру серв", "Ру/Крд серв",
    "Екб/Серв Ру/Мск серв", "Сиб/Дв серв", "Ру/Еу серв", "Ру/Вн серв",
    "Ру/Екб серв", "Ру/Сиб/Мск серв", "Унив серв",
    # Расширенный список регионов/серверов сообщества MLBB.
    "Кз серв", "Бел серв", "Мд серв", "Кг серв", "Аз серв", "Ам серв", "Гео серв",
    "Тж серв", "Тм серв", "Фил серв", "Индо серв", "Сг серв", "Мал серв", "Тай серв",
    "Мья серв", "Инд серв", "Бд серв", "Пак серв", "Неп серв", "Шл серв", "Яп серв",
    "Кор серв", "Тайв серв", "MENA серв", "Афр серв", "NA серв", "LATAM серв",
    "Бра серв", "Авст серв",
]
SERVER_ALIASES = {
    "ру": "Ру", "ru": "Ру", "россия": "Ру",
    "крд": "Крд", "krd": "Крд", "краснодар": "Крд",
    "мск": "Мск", "msk": "Мск", "москва": "Мск",
    "еу": "Еу", "eu": "Еу", "европа": "Еу",
    "узб": "Узб", "uzb": "Узб", "узбекистан": "Узб",
    "екб": "Екб", "ekb": "Екб", "екатеринбург": "Екб",
    "укр": "Укр", "ua": "Укр", "украина": "Укр",
    "сиб": "Сиб", "sib": "Сиб", "сибирь": "Сиб",
    "дв": "Дв", "dv": "Дв", "дальнийвосток": "Дв", "дальний-восток": "Дв",
    "вн": "Вн", "vn": "Вн", "вьетнам": "Вн",
    "кз": "Кз", "kz": "Кз", "казахстан": "Кз",
    "бел": "Бел", "by": "Бел", "беларусь": "Бел", "белоруссия": "Бел",
    "мд": "Мд", "md": "Мд", "молдова": "Мд",
    "кг": "Кг", "kg": "Кг", "кыргызстан": "Кг", "киргизия": "Кг",
    "аз": "Аз", "az": "Аз", "азербайджан": "Аз",
    "ам": "Ам", "am": "Ам", "армения": "Ам",
    "гео": "Гео", "geo": "Гео", "грузия": "Гео",
    "тж": "Тж", "tj": "Тж", "таджикистан": "Тж",
    "тм": "Тм", "tm": "Тм", "туркменистан": "Тм",
    "фил": "Фил", "ph": "Фил", "филиппины": "Фил",
    "индо": "Индо", "id": "Индо", "индонезия": "Индо",
    "сг": "Сг", "sg": "Сг", "сингапур": "Сг",
    "мал": "Мал", "my": "Мал", "малайзия": "Мал",
    "тай": "Тай", "th": "Тай", "таиланд": "Тай",
    "мья": "Мья", "mm": "Мья", "мьянма": "Мья",
    "инд": "Инд", "in": "Инд", "индия": "Инд",
    "бд": "Бд", "bd": "Бд", "бангладеш": "Бд",
    "пак": "Пак", "pk": "Пак", "пакистан": "Пак",
    "неп": "Неп", "np": "Неп", "непал": "Неп",
    "шл": "Шл", "lk": "Шл", "шриланка": "Шл", "шри-ланка": "Шл",
    "яп": "Яп", "jp": "Яп", "япония": "Яп",
    "кор": "Кор", "kr": "Кор", "корея": "Кор", "южнаякорея": "Кор",
    "тайв": "Тайв", "tw": "Тайв", "тайвань": "Тайв",
    "mena": "MENA", "мена": "MENA", "ближнийвосток": "MENA",
    "афр": "Афр", "africa": "Афр", "африка": "Афр",
    "na": "NA", "севернаяамерика": "NA",
    "latam": "LATAM", "латам": "LATAM", "латинскаяамерика": "LATAM",
    "бра": "Бра", "br": "Бра", "бразилия": "Бра",
    "авст": "Авст", "au": "Авст", "австралия": "Авст",
    "унив": "Унив", "универсальный": "Унив",
}
SERVER_COMBOS_ORDERED = {
    ("Еу", "Ру"): "Еу/Ру серв",
    ("Ру", "Еу"): "Ру/Еу серв",
    ("Ру", "Крд"): "Ру/Крд серв",
    ("Сиб", "Дв"): "Сиб/Дв серв",
    ("Ру", "Вн"): "Ру/Вн серв",
    ("Ру", "Екб"): "Ру/Екб серв",
    ("Ру", "Сиб", "Мск"): "Ру/Сиб/Мск серв",
    ("Екб", "Ру", "Мск"): "Екб/Серв Ру/Мск серв",
}
SERVER_SINGLE = {x.split()[0]: x for x in SERVER_TAGS if "/" not in x and x != "Унив серв"}


def _server_pending_root():
    return data.setdefault("pending_server_tags", {})

def _server_assignments_root():
    return data.setdefault("server_tag_assignments", {})

def _server_tags_enabled():
    """Единый переключатель Web ↔ Bot. PostgreSQL — источник истины, data.json — fallback."""
    enabled = bool(data.get("server_tags_enabled", True))
    try:
        from db_storage import get_setting
        db_value = get_setting("server_tags_enabled", None)
        if isinstance(db_value, bool):
            enabled = db_value
            data["server_tags_enabled"] = enabled
            return enabled
    except Exception as exc:
        print(f"[SERVER TAG] Не удалось прочитать переключатель из БД: {exc}")
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as file:
            disk_data = json.load(file)
        enabled = bool(disk_data.get("server_tags_enabled", enabled))
        data["server_tags_enabled"] = enabled
    except Exception:
        pass
    return enabled

def _normalize_server_input(text: str):
    raw = " ".join((text or "").strip().casefold().replace("ё", "е").split())
    cleaned = re.sub(r"\b(серв|сервер|server)\b", " ", raw)
    parts = [x for x in re.split(r"[\s,;+|/]+", cleaned) if x]
    names = []
    corrected = False
    for part in parts:
        key = part.replace("-", "")
        name = SERVER_ALIASES.get(key)
        if not name:
            close = difflib.get_close_matches(key, SERVER_ALIASES.keys(), n=1, cutoff=0.72)
            if close:
                name = SERVER_ALIASES[close[0]]
                corrected = True
        if name and name not in names:
            names.append(name)
            # Любой ввод не в каноническом виде (включая малую букву/полное название) показываем как исправление.
            if part != name:
                corrected = True
        elif not name:
            return [], False
    return names, corrected


def _parse_server_tag(text: str):
    raw = " ".join((text or "").strip().casefold().replace("ё", "е").split())
    if not raw:
        return None
    for tag in SERVER_TAGS:
        if raw == tag.casefold():
            return tag
    names, _ = _normalize_server_input(text)
    if not names:
        return None
    if "Унив" in names or len(names) > 3:
        return "Унив серв"
    if len(names) == 1:
        return SERVER_SINGLE.get(names[0])
    direct = SERVER_COMBOS_ORDERED.get(tuple(names))
    if direct:
        return direct
    for combo, tag in SERVER_COMBOS_ORDERED.items():
        if len(combo) == len(names) and frozenset(combo) == frozenset(names):
            return tag
    # Для новых серверов 2–3 значения собираются автоматически.
    generated = "/".join(names) + " серв"
    # Telegram member tag ограничен 16 символами.
    if len(generated) <= 16:
        return generated
    return "Унив серв"


def _canonical_server_input(text: str):
    names, corrected = _normalize_server_input(text)
    if not names:
        return None, False
    return "/".join(names), corrected

async def _telegram_set_member_tag(chat_id: int, user_id: int, tag: str):
    # Последняя защита: даже уже открытый 20-минутный запрос не выдаст тег после OFF на сайте.
    if not _server_tags_enabled():
        return False

    set_url = f"https://api.telegram.org/bot{BOT_TOKEN}/setChatMemberTag"
    get_url = f"https://api.telegram.org/bot{BOT_TOKEN}/getChatMember"
    timeout = aiohttp.ClientTimeout(total=15)

    async def api_call(session, url, payload):
        async with session.post(url, json=payload) as response:
            result = await response.json(content_type=None)
            if not response.ok or not result.get("ok"):
                raise RuntimeError(result.get("description") or f"Telegram HTTP {response.status}")
            return result

    async with aiohttp.ClientSession(timeout=timeout) as session:
        # setChatMemberTag должен заменять существующий тег. Сначала пробуем обычную замену.
        await api_call(session, set_url, {"chat_id": chat_id, "user_id": user_id, "tag": tag})

        # Не доверяем только ответу ok=True: читаем фактический тег участника.
        member = await api_call(session, get_url, {"chat_id": chat_id, "user_id": user_id})
        actual_tag = str((member.get("result") or {}).get("tag") or "")
        if actual_tag == tag:
            return True

        # Если Telegram почему-то оставил старый тег (например, УКР), принудительно
        # снимаем его пустой строкой и сразу ставим выбранный пользователем новый тег.
        await api_call(session, set_url, {"chat_id": chat_id, "user_id": user_id, "tag": ""})
        await api_call(session, set_url, {"chat_id": chat_id, "user_id": user_id, "tag": tag})

        # Финальная проверка: сообщение об успехе будет только если реально стоит новый тег.
        member = await api_call(session, get_url, {"chat_id": chat_id, "user_id": user_id})
        actual_tag = str((member.get("result") or {}).get("tag") or "")
        if actual_tag != tag:
            raise RuntimeError(f"Telegram не применил новый тег: ожидался {tag}, сейчас стоит {actual_tag or 'без тега'}")

    return True

async def _delete_message_later(chat_id: int, message_id: int, delay: int):
    await asyncio.sleep(delay)
    try:
        await bot.delete_message(chat_id, message_id)
    except Exception:
        pass

async def _ask_server_tag_delayed(chat_id: int, user, delay: int = 30):
    # Если функция выключена на сайте — бот вообще не запускает сценарий тегов.
    if not _server_tags_enabled():
        return
    # Ждём после вступления, чтобы запрос тега не появлялся мгновенно.
    await asyncio.sleep(delay)
    # За время задержки функцию могли выключить через сайт.
    if not _server_tags_enabled():
        return
    key = f"{chat_id}:{user.id}"
    # Не создаём второй запрос, если для участника уже есть активный.
    pending = _server_pending_root().get(key)
    if pending and int(pending.get("expires_at", 0)) > int(time.time()):
        return
    # Если за время ожидания пользователь уже вышел из чата — ничего не отправляем.
    try:
        member = await bot.get_chat_member(chat_id, user.id)
        if member.status in {"left", "kicked"}:
            return
    except Exception:
        pass
    await _ask_server_tag(chat_id, user)

SERVER_DISPLAY_NAMES = {
    "Ру": "Россия", "Крд": "Краснодар", "Мск": "Москва", "Еу": "Европа",
    "Узб": "Узбекистан", "Екб": "Екатеринбург", "Укр": "Украина",
    "Сиб": "Сибирь", "Дв": "Дальний Восток", "Вн": "Вьетнам",
    "Кз": "Казахстан", "Бел": "Беларусь", "Мд": "Молдова",
    "Кг": "Кыргызстан", "Аз": "Азербайджан", "Ам": "Армения",
    "Гео": "Грузия", "Тж": "Таджикистан", "Тм": "Туркменистан",
    "Фил": "Филиппины", "Индо": "Индонезия", "Сг": "Сингапур",
    "Мал": "Малайзия", "Тай": "Таиланд", "Мья": "Мьянма",
    "Инд": "Индия", "Бд": "Бангладеш", "Пак": "Пакистан",
    "Неп": "Непал", "Шл": "Шри-Ланка", "Яп": "Япония",
    "Кор": "Южная Корея", "Тайв": "Тайвань", "MENA": "Ближний Восток",
    "Афр": "Африка", "NA": "Северная Америка",
    "LATAM": "Латинская Америка", "Бра": "Бразилия", "Авст": "Австралия",
}

def _server_tag_options_text():
    options = [tag[:-5] for tag in SERVER_TAGS if tag.endswith(" серв") and "/" not in tag and tag != "Унив серв"]
    return "\n".join(
        f"• <code>{code}</code> — {SERVER_DISPLAY_NAMES.get(code, code)}"
        for code in options
    )

async def _ask_server_tag(chat_id: int, user):
    if not _server_tags_enabled():
        return
    mention = f'<a href="tg://user?id={user.id}">{user.full_name}</a>'
    msg = await bot.send_message(
        chat_id,
        f"🏷 {mention}, напиши свой сервер <b>ответом на это сообщение</b>.\n\n"
        "🌍 <b>Доступные серверы и регионы:</b>\n"
        f"{_server_tag_options_text()}\n\n"
        "✍️ Напиши сокращение из списка, например: <code>Мск</code> или <code>Ру/Еу</code>.\n"
        "Можно указать 2–3 сервера через / или запятую.\n"
        "Если серверов больше 3 — будет выдан тег <b>Унив серв</b>.\n\n"
        "🚫 Если у тебя спам-блок и ты не можешь написать сообщение — поставь <b>любую реакцию</b> на это сообщение.\n"
        "Бот примет реакцию только от пользователя, к которому обратился.\n\n"
        "⏳ На ответ или реакцию есть 20 минут.",
        parse_mode="HTML",
    )
    _server_pending_root()[f"{chat_id}:{user.id}"] = {
        "chat_id": chat_id, "user_id": user.id, "prompt_message_id": msg.message_id,
        "expires_at": int(time.time()) + SERVER_TAG_PENDING_SECONDS,
        "name": user.full_name,
    }
    save_data()
    return msg

async def server_tag_cleanup_loop():
    await asyncio.sleep(5)
    while True:
        now = int(time.time())
        changed = False
        for key, item in list(_server_pending_root().items()):
            if int(item.get("expires_at", 0)) <= now:
                _server_pending_root().pop(key, None)
                changed = True
                try:
                    await bot.delete_message(int(item["chat_id"]), int(item["prompt_message_id"]))
                except Exception:
                    pass
        if changed:
            save_data()
        await asyncio.sleep(10)


# =========================================================
# RANDOM REPLIES IN ALL GROUP CHATS
# =========================================================
# Для каждого группового чата ведётся отдельная база фраз и целей для ответа.
# Когда база фраз достигает 150 сообщений, она полностью очищается и цикл
# накопления начинается заново. Telegram Bot API видит только новые сообщения.
RANDOM_MIN_MINUTES = int(os.getenv("STAFF_RANDOM_MIN_MINUTES", "30"))
RANDOM_MAX_MINUTES = int(os.getenv("STAFF_RANDOM_MAX_MINUTES", "120"))
RANDOM_POOL_RESET_AT = 150
RANDOM_TARGET_LIMIT = 150

# AI replies via OpenAI Responses API. If OPENAI_API_KEY is not configured or
# the API is temporarily unavailable, the old random-message reply is used.
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()  # legacy; no longer required for AI replies
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5-mini").strip() or "gpt-5-mini"
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip() or "gemini-3.5-flash-lite"
AI_REPLY_GROQ_MODEL = os.getenv("AI_REPLY_GROQ_MODEL", AI_MODERATION_MODEL if "AI_MODERATION_MODEL" in globals() else "openai/gpt-oss-20b").strip() or "openai/gpt-oss-20b"
AI_CONTEXT_MESSAGES = max(5, min(40, int(os.getenv("AI_CONTEXT_MESSAGES", "20"))))
AI_MAX_OUTPUT_TOKENS = max(40, min(300, int(os.getenv("AI_MAX_OUTPUT_TOKENS", "120"))))


# =========================================================
# AI MODERATION
# =========================================================
# Отдельна от ИИ-ответов. Модерация анализирует обычные сообщения в группах,
# хранит нарушения в data.json и применяет эскалацию наказаний.
AI_MODERATION_DEFAULT = os.getenv("AI_MODERATION_ENABLED", "1").strip().lower() not in {"0", "false", "off", "no"}
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
AI_MODERATION_MODEL = os.getenv("AI_MODERATION_MODEL", "openai/gpt-oss-20b").strip() or "openai/gpt-oss-20b"
GEMINI_MODERATION_MODEL = os.getenv("GEMINI_MODERATION_MODEL", GEMINI_MODEL).strip() or GEMINI_MODEL
AI_MODERATION_MIN_CONFIDENCE = float(os.getenv("AI_MODERATION_MIN_CONFIDENCE", "0.88"))
AI_MODERATION_TIMEOUT = int(os.getenv("AI_MODERATION_TIMEOUT", "25"))
AI_MODERATION_MAX_TEXT = int(os.getenv("AI_MODERATION_MAX_TEXT", "1000"))
AI_MODERATION_CONTEXT_MESSAGES = max(4, min(20, int(os.getenv("AI_MODERATION_CONTEXT_MESSAGES", "8"))))
AI_MODERATION_PREFILTER = os.getenv("AI_MODERATION_PREFILTER", "1").strip().lower() not in {"0", "false", "off", "no"}
_AI_MOD_RATE_LIMIT_UNTIL = 0.0
AI_MODERATION_CONTEXT_TTL = max(300, int(os.getenv("AI_MODERATION_CONTEXT_TTL", "1800")))
_AI_MOD_CONTEXT = {}
AI_MODERATION_RULES = """Shinobi Team moderation rules. Be lenient on harmless chat and strict on serious harm.
1 insults/harassment: light=request (PRED); continued=warn; if continued after WARN=mute 30-60m; serious=2-12h; systematic=1-3d then ban. Friendly banter is allowed.
2 politics: neutral mentions allowed; conflict=request stop; continued=warning/mute 30-120m; agitation/provocation=2-12h; systematic=long mute/ban.
3 NSFW: porn=delete+mute 6-24h; repeat=1-3d then ban; mass spam=long mute/ban; sexual content involving minors=ban.
4 gore: real killings/corpses/severe injury=delete+mute 6-24h; repeat=1-7d; mass intentional spam=ban.
5 spam: small=request stop; continued=mute 15-60m; mass=1-6h; bots/raids/systematic=long mute/ban. Normal active chat/memes/several messages are allowed.
6 provocation/conflict: start=request (PRED); continued=warn; repeated after WARN=30m-2h; intentional escalation=2-12h; regular=1-3d then ban.
7 admin requests: first ignoring=request (PRED), continued=warn then underlying rule. Disagreement itself allowed.
8 admin conflict: same standards as any user; reasonable criticism allowed.
9 privacy: private info/chats=delete+mute 6-24h; addresses/docs/intimate material=long mute/ban; doxxing/threat to leak/repeat=ban.
10 prefix: missing server prefix is NOT a violation.
11 links/ads: normal useful links allowed; suspicious=delete+request (PRED); repeated suspicious=warn; intrusive ads=1-12h; scam/malware/systematic ads=ban.
12 alts: alt itself allowed; mute evasion=sanction alt; ban evasion=ban alt+escalate main; deception/provocation=case-by-case.
13 appeals: allowed; insults/spam/provocation during appeal use normal rules.
Minor violations expire after 30 days of normal behavior. Permanent ban mainly for scam, doxxing, raids, repeated ban evasion and other severe/systematic violations.
"""

def _ai_mod_context_add(message: Message):
    """Keep a short in-memory social context for moderation. No punishment state lives here."""
    if not message.from_user or message.from_user.is_bot:
        return
    text = (message.text or message.caption or "").strip()
    if not text:
        return
    now = int(time.time())
    cid = int(message.chat.id)
    bucket = _AI_MOD_CONTEXT.setdefault(cid, [])
    reply_to = getattr(message, "reply_to_message", None)
    reply_uid = getattr(getattr(reply_to, "from_user", None), "id", None)
    bucket.append({
        "ts": now,
        "message_id": int(message.message_id),
        "user_id": int(message.from_user.id),
        "name": (message.from_user.full_name or message.from_user.username or str(message.from_user.id))[:80],
        "reply_to_user_id": int(reply_uid) if reply_uid else None,
        "text": text[:700],
    })
    cutoff = now - AI_MODERATION_CONTEXT_TTL
    bucket[:] = [x for x in bucket if int(x.get("ts", 0)) >= cutoff][-AI_MODERATION_CONTEXT_MESSAGES:]

def _ai_mod_context_text(chat_id: int, current_message_id: int) -> str:
    now = int(time.time())
    cutoff = now - AI_MODERATION_CONTEXT_TTL
    rows = [x for x in _AI_MOD_CONTEXT.get(int(chat_id), []) if int(x.get("ts", 0)) >= cutoff]
    lines = []
    for x in rows[-AI_MODERATION_CONTEXT_MESSAGES:]:
        marker = " [CURRENT]" if int(x.get("message_id", 0)) == int(current_message_id) else ""
        reply = f" reply_to_user={x['reply_to_user_id']}" if x.get("reply_to_user_id") else ""
        lines.append(f"user_id={x.get('user_id')} name={x.get('name')}{reply}{marker}: {x.get('text','')}")
    return "\n".join(lines) if lines else "[no prior chat context available]"

def _ai_mod_root():
    root = data.setdefault("ai_moderation", {})
    root.setdefault("enabled", AI_MODERATION_DEFAULT)
    root.setdefault("chats", {})
    root.setdefault("violations", {})
    root.setdefault("log", [])
    return root

def _ai_mod_chat(chat_id: int):
    root = _ai_mod_root()
    state = root["chats"].setdefault(str(chat_id), {})
    state.setdefault("enabled", True)
    return state

def _ai_mod_enabled(chat_id: int) -> bool:
    root = _ai_mod_root()
    return bool(root.get("enabled", True) and _ai_mod_chat(chat_id).get("enabled", True))

def _ai_mod_user_state(chat_id: int, user_id: int):
    key = f"{chat_id}:{user_id}"
    st = _ai_mod_root()["violations"].setdefault(key, {"events": [], "last_at": 0})
    st.pop("points", None); st.pop("count", None)
    now = int(time.time())
    st["events"] = [e for e in st.setdefault("events", []) if now-int(e.get("ts",0)) <= 30*86400][-50:]
    return st

def _ai_mod_history(st):
    counts = {}
    for e in st.get("events", []): counts[e.get("category","other")] = counts.get(e.get("category","other"),0)+1
    return counts

def _ai_mod_candidate(message: Message) -> bool:
    if message.chat.type not in ("group", "supergroup") or not message.from_user or message.from_user.is_bot:
        return False
    # Do not moderate human replies to messages sent by any Telegram bot.
    # This check happens before Groq/context handling, so such replies cost zero AI tokens.
    replied = message.reply_to_message
    if replied and replied.from_user and replied.from_user.is_bot:
        return False
    text = (message.text or message.caption or "").strip()
    if not text or text.startswith(("/", ".")):
        return False
    return True

def _ai_mod_needs_ai(text: str) -> bool:
    """Cheap fail-open prefilter: skip only clearly neutral chat to save Groq tokens."""
    if not AI_MODERATION_PREFILTER:
        return True
    t = (text or "").strip().lower()
    if not t:
        return False
    # Links, invitations/ads, threats, conflict language, politics, NSFW/gore/privacy,
    # profanity and spam-like text always go to the contextual AI classifier.
    suspicious = re.compile(
        r"(?:https?://|t\.me/|@[a-z0-9_]{4,}|вступ|переход|заход|подпис|мой чат|мой канал|реклам|"
        r"убью|убить|сдох|дохни|угрож|найду тебя|адрес|телефон|докс|слив|личн(?:ые|ую) данн|"
        r"полит|путин|зеленск|росси|украин|войн|нато|пропаганд|"
        r"порно|секс|нюд|гол(?:ая|ый)|18\+|труп|расчлен|кровищ|убийств|"
        r"лох|дебил|идиот|туп(?:ой|ая)|кончен|клоун|уеб|ху[йея]|пизд|еба|бля|сука|мраз|чмо|заткни|закрой\s+(?:уже\s+)?рот|иди\s+нах|пош[её]л\s+нах|отстань|хватит|"
        r"спам|флуд|скам|мошенн|бан|мут|админ)"
    )
    if suspicious.search(t):
        return True
    # Obvious flood/repetition should still be checked.
    if len(t) > 500 or re.search(r"(.)\1{7,}", t) or len(re.findall(r"[!?]{2,}", t)) >= 3:
        return True
    # Multiple repeated tokens are a cheap spam signal.
    words = re.findall(r"[a-zа-яё0-9_]+", t)
    if len(words) >= 8 and len(set(words)) <= max(2, len(words)//4):
        return True
    return False

async def _ai_mod_is_exempt(message: Message) -> bool:
    if not message.from_user:
        return True
    if message.from_user.id in ADMIN_IDS:
        return True
    try:
        member = await bot.get_chat_member(message.chat.id, message.from_user.id)
        status = getattr(getattr(member, "status", None), "value", getattr(member, "status", None))
        return status in {"creator", "administrator"}
    except Exception:
        return False

async def _ai_mod_classify(message: Message, history=None, text_override=None):
    global _AI_MOD_RATE_LIMIT_UNTIL
    if not GROQ_API_KEY and not GEMINI_API_KEY:
        return None
    # Never mutate aiogram Message: it is a frozen Pydantic model. Tests pass text separately.
    source_text = text_override if text_override is not None else (message.text or message.caption or "")
    text=str(source_text or "").strip()[:AI_MODERATION_MAX_TEXT]
    # Manual `.мод ии тест` must always reach Groq; automatic moderation may skip
    # only clearly neutral messages. None here means fail-open/no punishment.
    if text_override is None and not _ai_mod_needs_ai(text):
        return {"violation":False,"category":"none","rule":0,"subpoint":1,"stage":"first","action":"none","mute_minutes":0,"delete":False,"confidence":1.0,"reason":"Локальный фильтр: нейтральное сообщение"}
    now=time.time()
    groq_in_cooldown = now < _AI_MOD_RATE_LIMIT_UNTIL
    history=history or {}
    schema='{"violation":true,"category":"none|insult|politics|nsfw|gore|spam|provocation|admin|privacy|link|alt","rule":0,"subpoint":1,"stage":"first|continued|serious|systematic|mass|critical","action":"none|request|warn|mute|ban","mute_minutes":0,"delete":false,"confidence":0.0,"reason":"Russian short reason"}'
    base_prompt=f"""You are a careful Telegram moderator. {AI_MODERATION_RULES}
Prior violations in last 30 days by category: {json.dumps(history, ensure_ascii=False)}.

RECENT CHAT CONTEXT (chronological; [CURRENT] is the message being judged):
{_ai_mod_context_text(message.chat.id, message.message_id)}

SOCIAL-CONTEXT POLICY — IMPORTANT:
- Profanity, rude vocabulary and insults are NOT violations by themselves. Judge behavior, not individual words.
- Friendly banter, mutual teasing and reciprocal joking insults are allowed, including phrases such as «дурак», «дебил», «уеба», «лох», «ты хуже», «ахах уеба 😂», «лох ебаный 😂», when the surrounding exchange looks friendly or reciprocal.
- A single MILD everyday jab with insufficient context can be violation=false/action=none. This leniency does NOT apply to clearly hostile targeted degradation or commands to leave/shut up.
- Explicit targeted phrases such as «конченый идиот», «ты вообще тупой?», «закрой рот», «иди нахуй отсюда», especially when combined in one message, are evidence of a real hostile attack unless the surrounding context clearly proves reciprocal friendly banter.
- If the SAME author sends multiple hostile/derogatory messages toward the SAME person within the recent context, treat repetition as strong evidence of continued harassment/conflict. Do not reset the judgment just because each message is analyzed separately.
- Repeated one-sided attacks like «ты тупой/клоун» → «иди нахуй отсюда» → «конченый идиот» should normally be a Rule 1 violation and should escalate according to prior/contextual continuation.
- Do NOT punish merely because a message contains obscene language or is addressed with «ты».
- Look at BOTH participants in the recent dialogue. Mutual joking/aggressive slang without distress is normally allowed.
- Treat it as a real violation when context shows one-sided repeated attacks, a person asks the other to stop, targeted humiliation/harassment, credible threats, sustained hostility, or deliberate escalation.
- A clear request like «хватит», «не оскорбляй меня», «отстань», «мне неприятно» followed by continued attacks is strong evidence the banter is no longer consensual.
- Serious threats, systematic bullying, doxxing, scams, sexual content involving minors and other severe categories remain enforceable immediately under the rules.
- Advertising/poaching does NOT require a clickable link. A clear invitation to move/join another chat, channel, service or project (for example «пошлите все в мой чат») can fall under Rule 11 when context shows promotion or member poaching without admin approval. Do not punish if context shows they mean the current Shinobi chat, an approved project, or it is not actually promotional.
- Never infer friendship or conflict from one rude phrase alone. If the available context cannot distinguish joking from genuine hostility, choose violation=false, action=none.
- Confidence must reflect SOCIAL context, not confidence that a word is insulting.

Do not invent missing context. If unsure whether banter is friendly, politics is neutral, spam is actually mass, or an alt evades punishment, mark no violation.
Choose subpoint as the exact numbered bullet within that rule's punishment list (1 = first bullet, 2 = second bullet, etc.). Return exactly ONE JSON object and nothing else. No markdown, no code fences, no comments. Use this shape: {schema}
CURRENT MESSAGE TEXT: {text}"""
    headers={"Authorization":f"Bearer {GROQ_API_KEY}","Content-Type":"application/json"}

    def parse_json(raw):
        raw=(raw or "").strip()
        raw=re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I).strip()
        if not raw:
            return None
        try:
            return json.loads(raw)
        except Exception:
            pass
        # Extract a JSON object from accidental explanatory text.
        a=raw.find("{"); b=raw.rfind("}")
        if a>=0 and b>a:
            try:
                return json.loads(raw[a:b+1])
            except Exception:
                pass
        # Groq can occasionally stop after a valid JSON value but before the final
        # closing brace. Repair ONLY missing structural closers; never invent text,
        # values, quotes or commas. This safely handles e.g. ..."delete":false
        # becoming ..."delete":false}.
        if a >= 0:
            candidate=raw[a:]
            in_string=False; escaped=False; stack=[]; structurally_ok=True
            for ch in candidate:
                if in_string:
                    if escaped:
                        escaped=False
                    elif ch == "\\":
                        escaped=True
                    elif ch == '"':
                        in_string=False
                    continue
                if ch == '"':
                    in_string=True
                elif ch in "{[":
                    stack.append(ch)
                elif ch in "}]":
                    if not stack or (ch == "}" and stack[-1] != "{") or (ch == "]" and stack[-1] != "["):
                        structurally_ok=False; break
                    stack.pop()
            if structurally_ok and not in_string and stack:
                repaired=candidate + "".join("}" if ch == "{" else "]" for ch in reversed(stack))
                try:
                    obj=json.loads(repaired)
                    print("[AI MOD] repaired truncated model JSON")
                    return obj
                except Exception:
                    pass
        # Safe partial NO-VIOLATION recovery. If Groq clearly emitted both
        # violation=false and action=none before truncation, we may safely
        # accept ONLY the non-punitive verdict. Never recover a violation=true
        # or any punitive action from incomplete JSON.
        low = raw.lower()
        if ('"violation":false' in low.replace(" ", "") and
                '"action":"none"' in low.replace(" ", "")):
            print("[AI MOD] accepted partial no-violation verdict (fail-safe)")
            return {
                "violation": False, "category": "none", "rule": 0,
                "subpoint": 1, "stage": "first", "action": "none",
                "mute_minutes": 0, "delete": False, "confidence": 1.0,
                "reason": "Нарушение не подтверждено"
            }
        return None

    async def call_gemini(session, prompt):
        if not GEMINI_API_KEY:
            return None
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODERATION_MODEL}:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.0, "maxOutputTokens": 320, "responseMimeType": "application/json"},
        }
        try:
            async with session.post(url, json=payload) as response:
                body = await response.text()
                if response.status >= 400:
                    print(f"[AI MOD] Gemini HTTP {response.status}: {body[:500]}")
                    return None
                result = json.loads(body)
            parts = (((result.get("candidates") or [{}])[0].get("content") or {}).get("parts") or [])
            raw = "".join(str(x.get("text", "")) for x in parts if isinstance(x, dict)).strip()
            obj = parse_json(raw)
            if isinstance(obj, dict):
                print("[AI MOD] Provider: GEMINI")
                return obj
            print(f"[AI MOD] Gemini invalid JSON: {raw[:500]}")
        except Exception as exc:
            print(f"[AI MOD] Gemini error: {exc}")
        return None

    try:
        timeout=aiohttp.ClientTimeout(total=AI_MODERATION_TIMEOUT)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            obj=None
            # Groq is primary. During its rate-limit cooldown we skip it entirely
            # and send the request straight to Gemini.
            if GROQ_API_KEY and not groq_in_cooldown:
                for attempt in range(2):
                    prompt=base_prompt if attempt==0 else base_prompt+"\nIMPORTANT: Your previous response was invalid. Output ONLY a valid JSON object matching the requested shape."
                    payload={"model":AI_MODERATION_MODEL,"messages":[{"role":"user","content":prompt}],"temperature":0.0,"max_completion_tokens":320}
                    async with session.post("https://api.groq.com/openai/v1/chat/completions",headers=headers,json=payload) as response:
                        body=await response.text()
                        if response.status>=400:
                            print(f"[AI MOD] HTTP {response.status}: {body[:500]}")
                            if response.status == 429:
                                wait = 600.0
                                retry_header = response.headers.get("Retry-After", "").strip()
                                try:
                                    if retry_header: wait = max(1.0, float(retry_header))
                                except Exception: pass
                                m = re.search(r"try again in\s+(?:(\d+)m)?([0-9.]+)?s", body, flags=re.I)
                                if m: wait = (float(m.group(1) or 0) * 60.0) + float(m.group(2) or 0)
                                _AI_MOD_RATE_LIMIT_UNTIL = time.time() + max(5.0, wait) + 2.0
                                print(f"[AI MOD] Groq rate-limit cooldown: {int(max(5.0, wait))+2}s; switching to Gemini")
                                break
                            if attempt==0 and response.status==400: continue
                            break
                        try: result=json.loads(body)
                        except Exception:
                            print(f"[AI MOD] invalid API response: {body[:500]}")
                            continue
                    raw=(((result.get("choices") or [{}])[0].get("message") or {}).get("content") or "").strip()
                    if not raw:
                        print(f"[AI MOD] empty model response attempt {attempt+1}")
                        continue
                    obj=parse_json(raw)
                    if isinstance(obj,dict):
                        print("[AI MOD] Provider: GROQ")
                        break
                    print(f"[AI MOD] invalid model JSON attempt {attempt+1}: {raw[:500]}")
            elif GROQ_API_KEY and groq_in_cooldown:
                print("[AI MOD] Groq cooldown active; switching to Gemini")

            if not isinstance(obj, dict):
                obj = await call_gemini(session, base_prompt)
            if not isinstance(obj,dict):
                print("[AI MOD] GROQ/GEMINI unavailable; fail-open — no moderation action")
                return None
        cmap={"insult":1,"politics":2,"nsfw":3,"gore":4,"spam":5,"provocation":6,"admin":7,"privacy":9,"link":11,"alt":12}
        cat=str(obj.get("category","none")); action=str(obj.get("action","none"))
        max_subpoints={1:4,2:4,3:4,4:3,5:4,6:4,7:2,8:3,9:3,10:2,11:3,12:3,13:3}
        rule=cmap.get(cat,0)
        try: subpoint=int(obj.get("subpoint",1) or 1)
        except Exception: subpoint=1
        subpoint=max(1,min(max_subpoints.get(rule,1),subpoint))
        return {"violation":bool(obj.get("violation",False)),"category":cat,"rule":rule,"subpoint":subpoint,"stage":str(obj.get("stage","first")),"action":action if action in {"none","request","warn","mute","ban"} else "none","mute_minutes":max(0,min(10080,int(obj.get("mute_minutes",0) or 0))),"delete":bool(obj.get("delete",False)),"confidence":max(0,min(1,float(obj.get("confidence",0) or 0))),"reason":str(obj.get("reason","Нарушение правил"))[:220]}
    except Exception as exc:
        print(f"[AI MOD] classify: {exc}")
        return None

async def _ai_mod_apply(message: Message, verdict: dict):
    if not verdict.get("violation") or verdict.get("confidence",0)<AI_MODERATION_MIN_CONFIDENCE: return False
    cat=verdict.get("category","none"); rule=int(verdict.get("rule",0)); subpoint=int(verdict.get("subpoint",1) or 1); stage=verdict.get("stage","first")
    if not rule: return False
    st=_ai_mod_user_state(message.chat.id,message.from_user.id); prior=_ai_mod_history(st).get(cat,0)
    action=verdict.get("action","none"); mins=int(verdict.get("mute_minutes",0) or 0)
    ranges={"insult":(30,4320),"politics":(30,1440),"nsfw":(360,4320),"gore":(360,10080),"spam":(15,1440),"provocation":(30,4320),"privacy":(360,10080),"link":(60,720)}
    if action=="mute":
        lo,hi=ranges.get(cat,(30,1440)); mins=max(lo,min(hi,mins or lo))
    immediate=(cat=="nsfw" and stage=="critical") or (cat=="privacy" and stage=="critical") or (cat=="link" and stage in {"serious","systematic","mass","critical"}) or (cat=="spam" and stage in {"mass","critical"}) or (cat=="alt" and stage in {"systematic","critical"})
    if action=="ban" and not immediate and prior<2:
        action="mute"; lo,hi=ranges.get(cat,(120,1440)); mins=max(lo,min(hi,mins or hi))
    if action=="none": return False
    if verdict.get("delete"):
        try: await message.delete()
        except Exception as exc: print(f"[AI MOD] delete: {exc}")
    applied=action
    central_rec=None; warn_count=0; auto_banned=False
    try:
        if action=="request":
            central_rec=await _mod_issue_pred(message.chat.id,message.from_user,verdict.get("reason","Нарушение правил"),rule,subpoint,0,"AI","ai",message.message_id); applied="pred"
        elif action=="warn":
            central_rec,warn_count,auto_banned=await _mod_issue_warn(message.chat.id,message.from_user,verdict.get("reason","Нарушение правил"),rule,subpoint,0,"AI","ai",message.message_id); applied="warn"
        elif action=="mute":
            central_rec=await _mod_issue_mute(message.chat.id,message.from_user,mins*60,verdict.get("reason","Нарушение правил"),rule,subpoint,0,"AI","ai",message.message_id)
        elif action=="ban":
            central_rec=await _mod_issue_ban(message.chat.id,message.from_user,0,verdict.get("reason","Нарушение правил"),rule,subpoint,0,"AI","ai",message.message_id)
    except Exception as exc: print(f"[AI MOD] punishment: {exc}"); return False
    now=int(time.time()); st["events"].append({"ts":now,"category":cat,"rule":rule,"subpoint":subpoint,"action":applied}); st["events"]=st["events"][-50:]; st["last_at"]=now
    names={1:"Оскорбления и травля",2:"Политика",3:"18+ контент",4:"Жесть / шок-контент",5:"Спам и флуд",6:"Провокации и конфликты",7:"Требования администрации",9:"Личные данные",11:"Ссылки и реклама",12:"Твинки"}
    if applied=="pred": act="⚠️ ПРЕД на 7 дней"
    elif applied=="warn": act=f"🟨 WARN {min(warn_count,3)}/3" + (" → ⛔ бан 2 дня" if auto_banned else "")
    elif applied=="ban": act="⛔ бан"
    else:
        dur=f"{mins//1440} дн." if mins%1440==0 else (f"{mins//60} ч." if mins%60==0 else f"{mins} мин."); act=f"🔇 мут на {dur}"
    mention=f'<a href="tg://user?id={message.from_user.id}">{html.escape(message.from_user.full_name)}</a>'
    try: await bot.send_message(message.chat.id,f"🧪 <b>AI-модерация • BETA</b>\n\n{mention}: <b>{act}</b>\nПричина: {html.escape(verdict.get('reason','Нарушение правил'))}\nПравило: <b>Правило {rule}, пункт {subpoint}</b> — {names.get(rule,'правила чата')}.",parse_mode="HTML")
    except Exception as exc: print(f"[AI MOD] notice: {exc}")
    log=_ai_mod_root()["log"]
    entry={"ts":now,"chat_id":message.chat.id,"chat_title":getattr(message.chat,"title","") or "","user_id":message.from_user.id,"user_name":message.from_user.full_name,"username":message.from_user.username or "","message_id":message.message_id,"message_text":(message.text or message.caption or "")[:1200],"action":applied,"mute_minutes":mins,"category":cat,"rule":rule,"subpoint":subpoint,"stage":stage,"confidence":verdict.get("confidence"),"reason":verdict.get("reason")}
    log.append(entry); del log[:-500]; save_data()
    try:
        staff_user=f"@{message.from_user.username}" if message.from_user.username else f"ID <code>{message.from_user.id}</code>"
        original=html.escape((message.text or message.caption or "")[:700]) or "[медиа/без текста]"
        staff_text=(f"🧪 <b>AI MODERATION • BETA</b>\n\n"
                    f"👤 <b>{html.escape(message.from_user.full_name)}</b> ({staff_user})\n"
                    f"💬 Чат: <b>{html.escape(getattr(message.chat,'title','') or str(message.chat.id))}</b>\n"
                    f"📝 Сообщение: <blockquote>{original}</blockquote>\n"
                    f"⚖️ Наказание: <b>{act}</b>\n"
                    f"📜 Правило: <b>Правило {rule}, пункт {subpoint}</b> — {html.escape(names.get(rule,'правила чата'))}\n"
                    f"🔎 Причина: {html.escape(verdict.get('reason','Нарушение правил'))}\n"
                    f"🎯 Уверенность AI: <b>{float(verdict.get('confidence',0)):.0%}</b>")
        await bot.send_message(ADMIN_STAFF_CHAT_ID,staff_text,parse_mode="HTML")
    except Exception as exc: print(f"[AI MOD] staff log: {exc}")
    return True

class AIModerationMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, Message) and _ai_mod_candidate(event) and _ai_mod_enabled(event.chat.id):
            # Add first so the classifier sees the current line marked [CURRENT]
            # plus the preceding dialogue from both participants.
            _ai_mod_context_add(event)
            if not await _ai_mod_is_exempt(event):
                verdict = await _ai_mod_classify(event, _ai_mod_history(_ai_mod_user_state(event.chat.id, event.from_user.id)))
                if verdict and await _ai_mod_apply(event, verdict):
                    return
        return await handler(event, middleware_data)


def _random_root():
    root = data.setdefault("random_replies", {})
    root.setdefault("chats", {})
    return root

def _random_chat_state(chat_id: int):
    chats = _random_root()["chats"]
    key = str(chat_id)
    state = chats.setdefault(key, {})
    state.setdefault("enabled", True)
    state.setdefault("min_minutes", RANDOM_MIN_MINUTES)
    state.setdefault("max_minutes", RANDOM_MAX_MINUTES)
    state.setdefault("messages", [])
    state.setdefault("targets", [])
    state.setdefault("title", "")
    state.setdefault("next_send_at", 0)
    state.setdefault("ai_enabled", True)
    return state

def _random_message_ok(message: Message) -> bool:
    if message.chat.type not in ("group", "supergroup") or not message.from_user or message.from_user.is_bot:
        return False
    text = (message.text or "").strip()
    if not text or len(text) < 2 or len(text) > 500:
        return False
    low = text.casefold()
    if text.startswith(("/", ".")) or "http://" in low or "https://" in low or "t.me/" in low:
        return False
    return True

class StaffRandomCollectorMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, Message) and _random_message_ok(event):
            state = _random_chat_state(event.chat.id)
            state["title"] = event.chat.title or str(event.chat.id)
            display_name = (event.from_user.full_name or event.from_user.username or str(event.from_user.id)).strip()
            state["messages"].append({
                "text": event.text.strip(),
                "user_id": event.from_user.id,
                "user_name": display_name[:80],
                "message_id": event.message_id,
            })
            state["targets"].append({
                "message_id": event.message_id,
                "user_id": event.from_user.id,
            })
            state["targets"] = state["targets"][-RANDOM_TARGET_LIMIT:]
            # Ровно при достижении 150 фраз очищаем базу и начинаем новый цикл.
            if len(state["messages"]) >= RANDOM_POOL_RESET_AT:
                state["messages"] = []
                state["targets"] = []
                state["next_send_at"] = 0
            save_data()
        return await handler(event, middleware_data)

async def _generate_ai_reply(state, target):
    if not state.get("ai_enabled", True) or (not GEMINI_API_KEY and not GROQ_API_KEY):
        return None

    pool = state.get("messages", [])
    target_id = target.get("message_id")
    target_item = next((x for x in reversed(pool) if x.get("message_id") == target_id), None)
    if not target_item:
        return None
    context_items = pool[-AI_CONTEXT_MESSAGES:]
    context_lines=[]
    for item in context_items:
        name=item.get("user_name") or f"user_{item.get('user_id','?')}"
        marker=" [СООБЩЕНИЕ ДЛЯ ОТВЕТА]" if item.get("message_id")==target_id else ""
        context_lines.append(f"{name}: {item.get('text','')}{marker}")
    prompt=("Ты — встроенный чат-бот сообщества Shinobi Team по Mobile Legends. "
            "Ответь именно на отмеченное сообщение с учётом контекста. Пиши естественно по-русски, обычно 1-2 коротких предложения. "
            "Не говори, что ты ИИ, не придумывай факты, не выдавай себя за конкретного участника. Без кавычек и префикса Ответ:.\n\n"
            "Контекст чата:\n"+"\n".join(context_lines))
    timeout=aiohttp.ClientTimeout(total=35)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        # Gemini is primary for casual replies so OpenAI credits are no longer required.
        if GEMINI_API_KEY:
            try:
                url=f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
                payload={"contents":[{"parts":[{"text":prompt}]}],"generationConfig":{"temperature":0.8,"maxOutputTokens":AI_MAX_OUTPUT_TOKENS}}
                async with session.post(url,json=payload) as response:
                    body=await response.text()
                    if response.status < 400:
                        result=json.loads(body); parts=(((result.get("candidates") or [{}])[0].get("content") or {}).get("parts") or [])
                        text="".join(str(x.get("text","")) for x in parts if isinstance(x,dict)).strip()
                        if text:
                            print("[AI REPLIES] Provider: GEMINI")
                            return text[:1000]
                    else: print(f"[AI REPLIES] Gemini HTTP {response.status}: {body[:500]}")
            except Exception as exc: print(f"[AI REPLIES] Gemini error: {exc}")
        if GROQ_API_KEY:
            try:
                payload={"model":AI_REPLY_GROQ_MODEL,"messages":[{"role":"user","content":prompt}],"temperature":0.8,"max_completion_tokens":AI_MAX_OUTPUT_TOKENS}
                headers={"Authorization":f"Bearer {GROQ_API_KEY}","Content-Type":"application/json"}
                async with session.post("https://api.groq.com/openai/v1/chat/completions",headers=headers,json=payload) as response:
                    body=await response.text()
                    if response.status < 400:
                        result=json.loads(body); text=(((result.get("choices") or [{}])[0].get("message") or {}).get("content") or "").strip()
                        if text:
                            print("[AI REPLIES] Provider: GROQ")
                            return text[:1000]
                    else: print(f"[AI REPLIES] Groq HTTP {response.status}: {body[:500]}")
            except Exception as exc: print(f"[AI REPLIES] Groq error: {exc}")
    return None

async def _send_random_reply(chat_id: int):
    state = _random_chat_state(chat_id)
    pool = state.get("messages", [])
    targets = state.get("targets", [])
    if len(pool) < 5 or not targets:
        return False, "Недостаточно сообщений в базе (нужно минимум 5)."

    # Prefer recent targets so the AI answers something that is still relevant.
    target = random.choice(targets[-30:])
    ai_text = await _generate_ai_reply(state, target)
    if ai_text:
        picked_text = ai_text
        source = "AI"
    else:
        choices = [x for x in pool if x.get("user_id") != target.get("user_id") and x.get("message_id") != target.get("message_id")]
        if not choices:
            choices = pool
        picked_text = random.choice(choices)["text"]
        source = "fallback"

    try:
        await bot.send_message(chat_id, picked_text, reply_to_message_id=target["message_id"])
    except Exception:
        await bot.send_message(chat_id, picked_text)
    return True, f"[{source}] {picked_text}"

def _schedule_next_random(state):
    import time
    lo = max(1, int(state.get("min_minutes", RANDOM_MIN_MINUTES)))
    hi = max(lo, int(state.get("max_minutes", RANDOM_MAX_MINUTES)))
    state["next_send_at"] = int(time.time()) + random.randint(lo * 60, hi * 60)

async def staff_random_reply_loop():
    import time
    await asyncio.sleep(60)
    while True:
        try:
            now = int(time.time())
            chats = list(_random_root().get("chats", {}).items())
            changed = False
            for chat_key, state in chats:
                if not state.get("enabled", True):
                    continue
                if not state.get("next_send_at"):
                    _schedule_next_random(state)
                    changed = True
                    continue
                if now >= int(state.get("next_send_at", 0)):
                    await _send_random_reply(int(chat_key))
                    _schedule_next_random(state)
                    changed = True
            if changed:
                save_data()
        except Exception as error:
            print(f"[RANDOM REPLIES] {error}")
        await asyncio.sleep(30)


# =========================================================
# БОТ
# =========================================================

# PythonAnywhere Free: исходящие подключения должны идти через их HTTP-прокси.
# Сначала берём адрес из переменных окружения PythonAnywhere; fallback оставлен
# на стандартный адрес их прокси.
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# Управление доступом к командам в группах.
# Состояние хранится в data.json, поэтому не сбрасывается после перезапуска.
COMMAND_BUTTON_TEXTS = {
    "🔗 Создать ссылку", "📊 Моя статистика", "🔗 Мои ссылки",
    "👥 Приглашённые", "🏆 Топ", "📋 Мой график",
    "🏖 Запросить выходной", "🗓 График / заполнить",
    "⚙️ Админ-панель", "🗓 Управление графиком", "📥 Скачать Excel",
    "📊 Общая статистика", "📅 Excel с графиками", "🏖 Excel с выходными",
    "🛠 Тех. работы",
}
BLOCK_NOTICE_COOLDOWN = 4 * 60 * 60
_block_notice_last = {}

# Команды, которые обязательно подчиняются групповому запрету.
# В том числе графики: в запрещённой группе они не должны обходить middleware.
BLOCKABLE_GROUP_COMMANDS = {
    ".график", "график",
    ".мой график", "мой график",
}

# Игровые команды тоже подчиняются групповому запрету. Список нужен для явной классификации/будущих настроек.
# Жёсткий whitelist игрового хаба: игровые команды и игровые inline-кнопки
# работают ТОЛЬКО в этой группе. Telegram может отдавать ID супергруппы со знаком "-",
# поэтому сравниваем по абсолютному значению.
GAME_COMMAND_ALLOWED_CHAT_ID = -1004326906997
CALL_ALLOWED_CHAT_ID = -1004368701081

def _game_chat_allowed(chat_id: int) -> bool:
    try:
        return int(chat_id) == GAME_COMMAND_ALLOWED_CHAT_ID
    except (TypeError, ValueError):
        return False

GAME_COMMAND_PREFIXES = (
    ".рынок", ".акции", ".купить", ".портфель",
    ".досье", ".анкета", ".матч", ".mlbb", ".шиноби", ".фарм", ".рынок график",
    ".продать", ".передать", ".подарить", ".дивиденды", ".рынок топ", ".рынок события",
)

def _is_game_command(text: str) -> bool:
    normalized = " ".join((text or "").strip().casefold().split())
    return any(normalized == prefix or normalized.startswith(prefix + " ") for prefix in GAME_COMMAND_PREFIXES)

MODERATION_COMMANDS = (".пред", ".преды", ".снятьпред", ".варн", ".варны", ".снятьварн", ".мут", ".размут", ".бан", ".разбан", ".нарушения", ".модлог", ".наказание")

def _is_moderation_command(text: str) -> bool:
    n = " ".join((text or "").casefold().split())
    return any(n == c or n.startswith(c + " ") for c in MODERATION_COMMANDS)

AI_MOD_COMMAND_PREFIX = ".мод ии"

def _is_ai_moderation_command(text: str) -> bool:
    n = " ".join((text or "").strip().casefold().split())
    return n == AI_MOD_COMMAND_PREFIX or n.startswith(AI_MOD_COMMAND_PREFIX + " ")

def _looks_like_bot_command(text: str) -> bool:
    text = (text or "").strip()
    normalized = " ".join(text.casefold().split())
    return (
        normalized in BLOCKABLE_GROUP_COMMANDS
        or text.startswith("/")
        or text.startswith(".")
        or text in COMMAND_BUTTON_TEXTS
    )

def _group_is_installed(chat_id: int) -> bool:
    if "data" not in globals():
        return False
    return str(chat_id) in {str(x) for x in data.get("installed_groups", [])}

def _chat_commands_blocked(chat_id: int) -> bool:
    settings = data.get("group_command_access", {}) if "data" in globals() else {}
    return settings.get(str(chat_id), "allow") == "deny"

class GroupCommandAccessMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, Message) and event.chat.type in {"group", "supergroup"}:
            text = (event.text or "").strip()
            is_install = text.casefold() in {".установить", "/установить", "/install"}

            # Спец-доступы имеют абсолютный приоритет над установкой группы и
            # общим переключателем команд на сайте.
            # Игровые команды работают только в GAME_COMMAND_ALLOWED_CHAT_ID.
            if _is_game_command(text):
                if not _game_chat_allowed(event.chat.id):
                    # В основном чате не выполняем игровую команду, а раз в 4 часа
                    # на пользователя подсказываем, где она доступна.
                    if int(event.chat.id) == CHAT_ID and event.from_user:
                        key = ("game_redirect", int(event.from_user.id))
                        now = time.time()
                        if now - _block_notice_last.get(key, 0) >= BLOCK_NOTICE_COOLDOWN:
                            _block_notice_last[key] = now
                            try:
                                user = event.from_user
                                if user.username:
                                    mention = f"@{html.escape(user.username)}"
                                else:
                                    display_name = html.escape(user.full_name or "Игрок")
                                    mention = f'<a href="tg://user?id={user.id}">{display_name}</a>'
                                await event.reply(
                                    f"{mention}, ❌ <b>ошибка!</b>\n\n"
                                    "🎮 Игровые команды доступны только в <b>игровом чате</b>.\n"
                                    "Используй <code>.фарм</code>, <code>.рынок</code>, <code>.статистика</code> "
                                    "и остальные игровые команды там.",
                                    parse_mode="HTML",
                                )
                            except Exception as exc:
                                print(f"[GAME REDIRECT NOTICE] {exc}")
                    return
                return await handler(event, middleware_data)

            # .вызов должен доходить до своего обработчика в CALL_ALLOWED_CHAT_ID,
            # где отдельно проверяется членство пользователя в Staff.
            normalized_text = " ".join(text.casefold().split())
            if normalized_text == ".вызов" or normalized_text.startswith(".вызов "):
                if int(event.chat.id) != CALL_ALLOWED_CHAT_ID:
                    return
                return await handler(event, middleware_data)

            # Команды модерации проходят общий запрет команд. Доступ проверяет их обработчик по Admin Staff.
            if _is_moderation_command(text) or _is_ai_moderation_command(text):
                return await handler(event, middleware_data)

            # До ручной установки бот не обслуживает команды в группе.
            # Команда .установить — единственное исключение.
            if not _group_is_installed(event.chat.id):
                if not is_install and _looks_like_bot_command(text):
                    return
            elif (
                _chat_commands_blocked(event.chat.id)
                and _looks_like_bot_command(text)
                and not is_install
                and not (_is_game_command(text) and _game_chat_allowed(event.chat.id))
            ):
                key = (event.chat.id, event.from_user.id if event.from_user else 0)
                now = time.time()
                if now - _block_notice_last.get(key, 0) >= BLOCK_NOTICE_COOLDOWN:
                    _block_notice_last[key] = now
                    code = "".join(random.choices(string.ascii_letters + string.digits, k=25))
                    try:
                        notice = await event.reply(f"error 404\n{code}")
                        async def _delete_error_message(chat_id: int, message_id: int):
                            await asyncio.sleep(ERROR_DELETE_AFTER)
                            try:
                                await bot.delete_message(chat_id, message_id)
                            except Exception:
                                pass
                        asyncio.create_task(_delete_error_message(notice.chat.id, notice.message_id))
                    except Exception:
                        pass
                return
        return await handler(event, middleware_data)

dp.message.outer_middleware(GroupCommandAccessMiddleware())

class GameCallbackWhitelistMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, CallbackQuery) and (event.data or "").startswith("game:"):
            msg = event.message
            if msg and msg.chat.type in {"group", "supergroup"} and not _game_chat_allowed(msg.chat.id):
                try:
                    await event.answer("Игровые команды в этой группе отключены.", show_alert=True)
                except Exception:
                    pass
                return
        return await handler(event, middleware_data)

dp.callback_query.outer_middleware(GameCallbackWhitelistMiddleware())

class ChatSummaryCollectorMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, Message) and event.chat.type in {"group", "supergroup"}:
            try:
                record_message(event.chat.id, event.chat.title or str(event.chat.id), event.from_user.id if event.from_user else None)
            except Exception as exc:
                print(f"[CHAT SUMMARY] collector: {exc}")
            # Запоминаем участников, которых бот реально видел в каждой группе.
            # Telegram Bot API не предоставляет метод получения полного списка участников.
            if event.from_user and not event.from_user.is_bot:
                try:
                    cid = str(event.chat.id)
                    uid = str(event.from_user.id)
                    members = data.setdefault("group_seen_members", {}).setdefault(cid, {})
                    members[uid] = {
                        "id": event.from_user.id,
                        "username": event.from_user.username or "",
                        "name": event.from_user.full_name or event.from_user.first_name or str(event.from_user.id),
                        "last_seen": datetime.now().isoformat(),
                    }
                    # Keep the global cache in sync too, so moderation by @username
                    # keeps working after redeploys and from other allowed handlers.
                    cached = get_user(event.from_user.id)
                    cached["id"] = event.from_user.id
                    cached["username"] = event.from_user.username or ""
                    cached["name"] = event.from_user.full_name or event.from_user.first_name or str(event.from_user.id)
                    save_data()
                except Exception as exc:
                    print(f"[SEEN MEMBERS] {exc}")
        return await handler(event, middleware_data)

dp.message.outer_middleware(ChatSummaryCollectorMiddleware())

# Мгновенные случайные реакции на каждое сообщение основного чата.
# Отправляем асинхронно, чтобы запрос Telegram не задерживал обработчики.
from aiogram.types import ReactionTypeEmoji

AUTO_REACTION_EMOJIS = ("❤️", "🔥", "😂", "👍", "🥰", "🎉")
# Telegram accepts only reactions enabled in the group. Cache the current list.
_reaction_cache = {}  # chat_id -> (timestamp, tuple of emoji)
_reaction_cache_ttl = 300

async def _allowed_reactions(chat_id: int, force_refresh: bool = False):
    now = asyncio.get_running_loop().time()
    cached = _reaction_cache.get(chat_id)
    if cached and not force_refresh and now - cached[0] < _reaction_cache_ttl:
        return cached[1]
    try:
        chat = await bot.get_chat(chat_id)
        available = getattr(chat, "available_reactions", None)
        if available is None:
            # None means the chat uses Telegram's default set.
            emojis = AUTO_REACTION_EMOJIS
        else:
            emojis = tuple(r.emoji for r in available if isinstance(r, ReactionTypeEmoji))
        _reaction_cache[chat_id] = (now, emojis)
        return emojis
    except Exception as exc:
        print(f"[AUTO REACTION CONFIG] {chat_id}: {exc}")
        return cached[1] if cached else ("👍", "❤️")

# Shared site setting, refreshed frequently without restarting the bot.
_reaction_setting_cache = (0.0, True)
def _auto_reactions_enabled():
    global _reaction_setting_cache
    now = asyncio.get_running_loop().time()
    if now - _reaction_setting_cache[0] < 2:
        return _reaction_setting_cache[1]
    enabled = _reaction_setting_cache[1]
    try:
        from db_storage import get_setting
        value = get_setting('auto_reactions_enabled', True)
        enabled = value is not False and str(value).lower() not in ('false', '0', 'off')
    except Exception as exc:
        print(f'[AUTO REACTION SETTING] {exc}')
    _reaction_setting_cache = (now, enabled)
    return enabled

async def _send_instant_reaction(chat_id: int, message_id: int):
    if not _auto_reactions_enabled():
        return
    emojis = await _allowed_reactions(chat_id)
    if not emojis:
        return  # Reactions are disabled or only paid/custom reactions are allowed.
    try:
        await bot.set_message_reaction(
            chat_id=chat_id,
            message_id=message_id,
            reaction=[ReactionTypeEmoji(emoji=random.choice(emojis))],
        )
    except Exception as exc:
        if "REACTION_INVALID" in str(exc).upper():
            # Admins may have changed the chat's allowed reactions.
            # Refresh for the next message; do not spam invalid retries.
            _reaction_cache.pop(chat_id, None)
        else:
            print(f"[AUTO REACTION] {chat_id}/{message_id}: {exc}")

class InstantReactionMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if (isinstance(event, Message) and event.chat.id == CHAT_ID
                and event.from_user and not event.from_user.is_bot):
            asyncio.create_task(_send_instant_reaction(event.chat.id, event.message_id))
        return await handler(event, middleware_data)

# Внешний middleware срабатывает и для сообщений без подходящего обработчика.
dp.message.outer_middleware(InstantReactionMiddleware())




# Доступ к боту в ЛС разрешён только участникам группы Admin Staff.
# Проверка выполняется через Telegram при каждом действии в личном чате,
# поэтому после выхода/исключения из Admin Staff доступ прекращается.
async def _is_admin_staff_member(user_id: int) -> bool:
    try:
        member = await bot.get_chat_member(ADMIN_STAFF_CHAT_ID, user_id)
        status = getattr(member, "status", None)
        status_value = getattr(status, "value", status)
        return status_value in {"creator", "administrator", "member", "restricted"}
    except Exception:
        return False


class PrivateAdminStaffOnlyMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        user = getattr(event, "from_user", None)
        if not user or user.is_bot:
            return await handler(event, middleware_data)

        is_private = False
        if isinstance(event, Message):
            is_private = event.chat.type == "private"
        elif isinstance(event, CallbackQuery):
            is_private = bool(event.message and event.message.chat.type == "private")

        if not is_private:
            return await handler(event, middleware_data)

        if await _is_admin_staff_member(user.id):
            return await handler(event, middleware_data)

        if isinstance(event, CallbackQuery):
            try:
                await event.answer("⛔ Доступ только для участников Admin Staff.", show_alert=True)
            except Exception:
                pass
        else:
            try:
                await event.answer(
                    "⛔ <b>Доступ запрещён</b>\n\n"
                    "Использовать Shinobi Team Bot в личных сообщениях могут только участники <b>Admin Staff</b>.",
                    parse_mode="HTML"
                )
            except Exception:
                pass
        return


dp.message.outer_middleware(PrivateAdminStaffOnlyMiddleware())
dp.callback_query.outer_middleware(PrivateAdminStaffOnlyMiddleware())
dp.message.outer_middleware(AIModerationMiddleware())
dp.message.outer_middleware(StaffRandomCollectorMiddleware())

init_schedule_system(ADMIN_IDS)
dp.include_router(schedule_router)


# =========================================================
# JSON
# =========================================================

def default_data():
    return {
        "users": {},
        "links": {},
        "bot_admins": [],
        "known_groups": {},
        "installed_groups": [],
        "group_command_access": {},
        "pending_join_requests": {},
        "advertising_link": "",
        "maintenance_mode": False,
        "invite_adjustments": [],
        "pending_server_tags": {},
        "server_tag_assignments": {},
        "game_hub": {}
    }


def _load_data_from_excel_backup():
    """Восстанавливает полную data.json из скрытого листа Excel Backup JSON."""
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        workbook = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
        if "Backup JSON" not in workbook.sheetnames:
            workbook.close()
            return None
        sheet = workbook["Backup JSON"]
        chunks = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row and row[0] is not None:
                chunks.append(str(row[0]))
        workbook.close()
        if not chunks:
            return None
        restored = json.loads("".join(chunks))
        if not isinstance(restored, dict):
            return None
        restored.setdefault("users", {})
        restored.setdefault("links", {})
        restored.setdefault("bot_admins", [])
        restored.setdefault("known_groups", {})
        restored.setdefault("installed_groups", [])
        restored.setdefault("group_command_access", {})
        restored.setdefault("pending_join_requests", {})
        restored.setdefault("advertising_link", "")
        restored.setdefault("maintenance_mode", False)
        restored.setdefault("invite_adjustments", [])
        restored.setdefault("pending_server_tags", {})
        restored.setdefault("server_tag_assignments", {})
        restored.setdefault("game_hub", {})
        restored.setdefault("group_seen_members", {})
        print("База восстановлена из Excel-резерва.")
        return restored
    except Exception as error:
        print("Не удалось восстановить базу из Excel:", error)
        return None


def load_data():

    if not os.path.exists(DATA_FILE):
        restored = _load_data_from_excel_backup()
        return restored if restored is not None else default_data()

    try:

        with open(DATA_FILE, "r", encoding="utf-8") as file:
            loaded = json.load(file)

        loaded.setdefault("users", {})
        loaded.setdefault("links", {})
        loaded.setdefault("bot_admins", [])
        loaded.setdefault("known_groups", {})
        loaded.setdefault("installed_groups", [])
        loaded.setdefault("group_command_access", {})
        loaded.setdefault("pending_join_requests", {})
        loaded.setdefault("advertising_link", "")
        loaded.setdefault("maintenance_mode", False)
        loaded.setdefault("invite_adjustments", [])
        loaded.setdefault("pending_server_tags", {})
        loaded.setdefault("server_tag_assignments", {})
        loaded.setdefault("game_hub", {})
        loaded.setdefault("group_seen_members", {})

        return loaded

    except (json.JSONDecodeError, OSError):
        restored = _load_data_from_excel_backup()
        return restored if restored is not None else default_data()


data = load_data()

# Фиксированная рекламная ссылка Shinobi Team. Сохраняем её как отдельный источник.
# Preserve saved advertising link; never overwrite it on every restart.
ADVERTISING_LINK = data.get("advertising_link") or ADVERTISING_LINK
data["advertising_link"] = ADVERTISING_LINK
data.setdefault("links", {})
_ad_info = data["links"].setdefault(ADVERTISING_LINK, {})
_ad_info.setdefault("owner_id", None)
_ad_info["type"] = "advertising"
_ad_info["label"] = "Рекламная ссылка"
_ad_info.setdefault("created", datetime.now().isoformat())
_ad_info.setdefault("invited", 0)
_ad_info.setdefault("users", [])


# =========================================================
# SHINOBI GAME HUB: рынок / досье / MLBB Match
# =========================================================
MARKET_WALLETS_FILE = Path(__file__).resolve().parent / "market_wallets.json"
FARM_COOLDOWN = 60 * 60

MARKET_FACTS = [
    "💡 Цены рынка меняются случайно и не зависят от активности участников.",
    "💡 Портфель может дорожать и дешеветь даже без новых покупок.",
    "💡 Фарм выдаёт случайно от 1 до 100 💎 за попытку.",
    "💡 Дивиденды можно забирать раз в сутки — выплата зависит от стоимости портфеля.",
    "💡 Акции можно продавать, дарить и смотреть через reply на сообщение участника.",
    "💡 Редкие рыночные события дают более сильные скачки цены.",
    "💡 Чем больше разных активов в портфеле, тем выше коллекционные достижения.",
    "💡 Покупать собственные акции нельзя — рынок любит конкуренцию.",
]
MARKET_EVENT_TEXTS = [
    "залетел в рекомендации", "стал мемом дня", "нашёл секретный бафф", "поймал волну хайпа",
    "проиграл спор рынку", "попал под коррекцию", "внезапно стал любимчиком инвесторов", "словил рыночную турбулентность",
]
RARITIES = [("Обычный",55),("Редкий",25),("Эпический",15),("Легендарный",5)]
DOSSIER_NOTES = [
    "🕵️ Подозрительно часто появляется именно тогда, когда чат оживает.",
    "📡 Радар Shinobi продолжает наблюдение за активностью.",
    "🎮 Возможный мотив: ещё одна катка перед сном.",
    "🥷 Статус дела: наблюдение продолжается.",
    "📁 Архив пополняется только действиями внутри чата.",
]
FARM_EVENTS = [
    "нашёл тайник под башней", "забрал награду за ежедневку", "выиграл мини-контракт Shinobi",
    "продал редкий мем на чёрном рынке", "поймал золотой крип", "нашёл забытый кошелёк рынка",
    "получил премию за активность", "удачно закрыл секретное задание", "сорвал маленький джекпот",
]


def _read_market_wallets() -> dict:
    try:
        return json.loads(MARKET_WALLETS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _market_balance(chat_id: int, uid: str, default: int = 0) -> int:
    # PostgreSQL is the source of truth when configured, so web changes are
    # immediately visible to the bot even when web/bot run in separate processes.
    key = f"market_wallet:{int(chat_id)}:{str(uid)}"
    try:
        from db_storage import get_setting
        value = get_setting(key, None)
        if value is not None:
            return max(0, int(value))
    except Exception:
        pass
    wallets = _read_market_wallets()
    return max(0, int(wallets.get(str(chat_id), {}).get(str(uid), default)))


def _set_market_balance(chat_id: int, uid: str, amount: int):
    wallets = _read_market_wallets()
    wallets.setdefault(str(chat_id), {})[str(uid)] = max(0, int(amount))
    tmp = MARKET_WALLETS_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(wallets, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(MARKET_WALLETS_FILE)
    try:
        from db_storage import set_setting, sync_to_db
        set_setting(f"market_wallet:{int(chat_id)}:{str(uid)}", max(0, int(amount)))
        sync_to_db()
    except Exception:
        pass


def _game_chat(chat_id: int) -> dict:
    root = data.setdefault("game_hub", {})
    chat = root.setdefault(str(chat_id), {})
    chat.setdefault("players", {})
    return chat


def _game_player(chat_id: int, user) -> dict:
    chat = _game_chat(chat_id)
    uid = str(user.id)
    p = chat["players"].setdefault(uid, {
        "username": user.username or "", "name": user.full_name or str(user.id),
        "messages": 0, "replies": 0, "reactions": 0, "balance": 0,
        "portfolio": {}, "profile": {}, "created_at": datetime.now().isoformat(),
        "last_farm": 0, "farm_count": 0, "farm_total": 0,
        "dossier_started_at": datetime.now().isoformat(), "dossier_week_messages": 0,
        "dossier_night_games": 0, "dossier_words": {}, "dossier_commands": {},
    })
    p["username"] = user.username or p.get("username", "")
    p["name"] = user.full_name or p.get("name", str(user.id))
    p["balance"] = _market_balance(chat_id, uid, int(p.get("balance", 0) or 0))
    return p


def _market_price_for_day(chat_id: int, uid: str, day_key: str) -> int:
    # Игровая котировка не зависит от активности. Цены намеренно держатся
    # в компактном диапазоне, чтобы экономика не раздувалась до тысяч.
    from datetime import date, timedelta
    try: target = date.fromisoformat(day_key)
    except Exception: target = datetime.now().date()
    epoch = date(2026, 1, 1)
    days = max(0, (target - epoch).days)
    rng0 = random.Random(f"shinobi-base-v2:{chat_id}:{uid}")
    price = rng0.randint(35, 120)
    start = max(0, days - 365)
    if start:
        rngbase = random.Random(f"shinobi-anchor-v2:{chat_id}:{uid}:{start}")
        price = rngbase.randint(35, 120)
    for i in range(start, days + 1):
        d = (epoch + timedelta(days=i)).isoformat()
        r = random.Random(f"shinobi-market-v2:{chat_id}:{uid}:{d}")
        swing = r.uniform(-0.07, 0.07)
        if r.random() < 0.04:
            swing = r.uniform(-0.14, 0.14)
        # Мягкое возвращение к середине рынка не даёт цене навсегда
        # прилипнуть к минимуму/максимуму после длинной серии колебаний.
        swing += (80 - price) / 80 * 0.018
        price = max(20, min(250, int(round(price * (1 + swing)))))
    return price

def _market_delta(chat_id:int, uid:str):
    from datetime import timedelta
    today=datetime.now().date(); yesterday=today-timedelta(days=1)
    a=_market_price_for_day(chat_id,uid,yesterday.isoformat()); b=_market_price_for_day(chat_id,uid,today.isoformat())
    pct=((b-a)/a*100) if a else 0
    return a,b,pct

def _rarity(chat_id:int, uid:str):
    r=random.Random(f"rarity:{chat_id}:{uid}").randint(1,100); acc=0
    for name,w in RARITIES:
        acc+=w
        if r<=acc:return name
    return "Обычный"

def _reply_target(message: Message):
    u=getattr(getattr(message,'reply_to_message',None),'from_user',None)
    if u and not u.is_bot:
        return str(u.id), _game_player(message.chat.id,u)
    return None,None

def _target_from_message(message: Message, token=None):
    uid,p=_reply_target(message)
    if p:return uid,p
    if token:return _find_game_player(message.chat.id,token)
    return None,None


def _market_price(chat_id: int, uid: str) -> int:
    return _market_price_for_day(chat_id, uid, datetime.now().date().isoformat())


def _find_game_player(chat_id: int, token: str):
    token = (token or "").strip().lstrip("@").casefold()
    if not token: return None, None
    players = _game_chat(chat_id)["players"]
    if token.isdigit() and token in players: return token, players[token]
    for uid, p in players.items():
        if (p.get("username") or "").casefold() == token: return uid, p
    return None, None


def _player_label(p: dict) -> str:
    return "@" + p["username"] if p.get("username") else p.get("name", "Игрок")


def _market_list_label(p: dict) -> str:
    # Для массовых рейтингов не используем @username: список не должен
    # тегать участников и присылать им уведомления.
    name = str(p.get("name") or p.get("username") or "Игрок").strip()
    name = name.replace("@", "＠")
    return html.escape(name)


def _market_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⛏ Фарм", callback_data="game:farm"), InlineKeyboardButton(text="💼 Портфель", callback_data="game:portfolio")],
        [InlineKeyboardButton(text="📈 График рынка", callback_data="game:chart"), InlineKeyboardButton(text="🎲 Факт", callback_data="game:fact")],
        [InlineKeyboardButton(text="💸 Дивиденды", callback_data="game:dividends"), InlineKeyboardButton(text="🏆 Топ инвесторов", callback_data="game:investors")],
        [InlineKeyboardButton(text="🥷 Помощь", callback_data="game:help")],
    ])


def _hub_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📈 Рынок", callback_data="game:market"), InlineKeyboardButton(text="⛏ Фарм", callback_data="game:farm")],
        [InlineKeyboardButton(text="📁 Досье", callback_data="game:dossier"), InlineKeyboardButton(text="💘 MLBB Match", callback_data="game:match")],
        [InlineKeyboardButton(text="🎲 Факт MLBB", callback_data="game:mlbbfact")],
    ])


def _sparkline(values):
    bars = "▁▂▃▄▅▆▇█"
    lo, hi = min(values), max(values)
    if hi == lo: return bars[3] * len(values)
    return "".join(bars[round((v-lo)/(hi-lo)*(len(bars)-1))] for v in values)


def _market_chart_text(chat_id: int) -> str:
    players = _game_chat(chat_id)["players"]
    if not players: return "📈 Рынок пока пуст."
    top = sorted(players.items(), key=lambda x: _market_price(chat_id, x[0]), reverse=True)[:5]
    now = time.time()
    day_keys = [datetime.fromtimestamp(now - 86400*i).date().isoformat() for i in range(6, -1, -1)]
    labels = [datetime.fromtimestamp(now - 86400*i).strftime("%d.%m") for i in range(6, -1, -1)]
    lines = ["📈 <b>ГРАФИК РЫНКА · 7 ДНЕЙ</b>", "<code>" + "  ".join(labels) + "</code>", ""]
    for uid,p in top:
        vals=[_market_price_for_day(chat_id,uid,d) for d in day_keys]
        delta=vals[-1]-vals[0]; sign="+" if delta>=0 else ""
        lines.append(f"<b>{_player_label(p)}</b>  {_sparkline(vals)}  {vals[-1]} 💎 ({sign}{delta})")
    lines += ["", random.choice(MARKET_FACTS)]
    return "\n".join(lines)


def _tinder_score(a: dict, b: dict) -> tuple[int, list[str]]:
    score, reasons = 35, []
    if a.get("server") and a.get("server", "").casefold() == b.get("server", "").casefold(): score += 25; reasons.append("один сервер")
    ar={x.strip().casefold() for x in a.get("roles","").split(",") if x.strip()}; br={x.strip().casefold() for x in b.get("roles","").split(",") if x.strip()}
    if ar and br and ar.isdisjoint(br): score += 15; reasons.append("роли дополняют друг друга")
    if a.get("time") and a.get("time","").casefold()==b.get("time","").casefold(): score += 10; reasons.append("совпадает время игры")
    if a.get("mic") and a.get("mic","").casefold()==b.get("mic","").casefold(): score += 8; reasons.append("одинаково по микрофону")
    if a.get("style") and a.get("style","").casefold()==b.get("style","").casefold(): score += 7; reasons.append("совпадает стиль")
    return min(score,99), reasons


async def _track_game_activity(message: Message):
    if message.chat.type not in {"group","supergroup"} or not message.from_user or message.from_user.is_bot: return
    p=_game_player(message.chat.id,message.from_user)
    p["messages"]=int(p.get("messages",0))+1
    if message.reply_to_message: p["replies"]=int(p.get("replies",0))+1

    # Досье собирается автоматически в течение 7 дней только из активности в чате.
    now = datetime.now()
    try:
        started = datetime.fromisoformat(p.get("dossier_started_at") or now.isoformat())
    except Exception:
        started = now; p["dossier_started_at"] = now.isoformat()
    # После каждой завершённой недели начинаем новый цикл наблюдения.
    if (now - started).total_seconds() >= 14 * 86400:
        p["dossier_started_at"] = now.isoformat(); p["dossier_week_messages"] = 0
        p["dossier_night_games"] = 0; p["dossier_words"] = {}; p["dossier_commands"] = {}
        started = now
    p["dossier_week_messages"] = int(p.get("dossier_week_messages",0)) + 1
    if now.hour >= 23 or now.hour < 6:
        p["dossier_night_games"] = int(p.get("dossier_night_games",0)) + 1
    txt=(message.text or message.caption or "").strip()
    if txt.startswith("."):
        cmd=txt.split()[0].casefold()[:32]; d=p.setdefault("dossier_commands",{}); d[cmd]=int(d.get(cmd,0))+1
    # Собираем частые короткие слова/фразы, исключая команды и ссылки.
    if txt and not txt.startswith(".") and "http" not in txt.casefold():
        cleaned=" ".join(txt.replace("\n"," ").split())
        if 2 <= len(cleaned) <= 42:
            d=p.setdefault("dossier_words",{}); d[cleaned]=int(d.get(cleaned,0))+1
            if len(d)>40:
                for k,_ in sorted(d.items(), key=lambda kv: kv[1])[:-30]: d.pop(k,None)
    save_data()


class GameActivityMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, Message): await _track_game_activity(event)
        return await handler(event, middleware_data)


dp.message.outer_middleware(GameActivityMiddleware())


async def _do_farm(chat_id: int, user, reply_func):
    p=_game_player(chat_id,user); now=time.time(); last=float(p.get("last_farm",0) or 0)
    left=FARM_COOLDOWN-(now-last)
    if left>0:
        mins=max(1,int((left+59)//60))
        await reply_func(f"⏳ Фарм уже использован. Следующая попытка примерно через <b>{mins} мин.</b>\n\n{random.choice(MARKET_FACTS)}", parse_mode="HTML", reply_markup=_market_keyboard())
        return
    reward=random.randint(1,100); event=random.choice(FARM_EVENTS)
    # Каждый уникальный актив даёт +2% к фарму, максимум +20%.
    owned=len([q for q in p.get("portfolio",{}).values() if int(q)>0]); invest_bonus=int(round(reward*min(0.20,owned*0.02)))
    reward += invest_bonus
    # Небольшие редкие события делают фарм менее однообразным, но базовый ролл всегда 1–100.
    bonus=0; bonus_text=""
    roll=random.random()
    if roll < .03: bonus=100; bonus_text="\n🔥 <b>ДЖЕКПОТ:</b> +100 💎 бонусом!"
    elif roll < .10: bonus=25; bonus_text="\n✨ <b>Удача:</b> +25 💎 бонусом!"
    if random.random()<0.015:
        p["insurance"]=int(p.get("insurance",0))+1; bonus_text += "\n🛡 <b>Редкая находка:</b> страховка портфеля ×1!"
    if invest_bonus: bonus_text += f"\n📦 Бонус портфеля: <b>+{invest_bonus} 💎</b>"
    balance=_market_balance(chat_id,str(user.id),0)+reward+bonus
    _set_market_balance(chat_id,str(user.id),balance)
    p["balance"]=balance; p["last_farm"]=now; p["farm_count"]=int(p.get("farm_count",0))+1; p["farm_total"]=int(p.get("farm_total",0))+reward+bonus
    save_data()
    reactions=["Неплохо!", "Рынок сегодня добр 😎", "Копилка Shinobi пополнена.", "Ещё шаг к биржевой империи.", "Секретная операция завершена."]
    text=(f"⛏ <b>ФАРМ SHINOBI</b>\n\n{_player_label(p)} {event} и получил <b>{reward} 💎</b>.{bonus_text}\n"
          f"💰 Баланс: <b>{balance} 💎</b>\n📦 Всего нафармлено: <b>{p['farm_total']} 💎</b>\n\n{random.choice(reactions)}")
    await reply_func(text,parse_mode="HTML",reply_markup=_market_keyboard())


@dp.message(F.text.casefold().startswith(".шиноби"))
async def shinobi_game_help(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    arg=(message.text or "")[len(".шиноби"):].strip().casefold()
    if arg in {"рынок","черный рынок","чёрный рынок"}:
        text=("📈 <b>ЧЁРНЫЙ РЫНОК SHINOBI</b>\n\nВиртуальная биржа участников чата. Стоимость меняется случайно и <b>не зависит от активности</b>. "
              "Стартовый баланс у каждого — <b>0 💎</b>, валюту можно фармить.\n\n<b>Команды:</b>\n<code>.рынок</code> — топ рынка\n"
              "<code>.рынок график</code> — график за 7 дней\n<code>.акции @user</code> — цена игрока\n<code>.купить @user 1</code> — купить акции\n"
              "<code>.портфель</code> — баланс и активы\n<code>.фарм</code> — получить 1–100 💎 (раз в час)\n\n"+random.choice(MARKET_FACTS))
    elif arg in {"досье","досье бот"}:
        text=("📁 <b>ДОСЬЕ SHINOBI</b>\n\nШуточное досье строится только по активности внутри чата и игровых функций. "
              "Никаких личных данных извне бот не ищет.\n\n<code>.досье</code> — своё\n<code>.досье @user</code> — другого участника\n\n"+random.choice(DOSSIER_NOTES))
    elif arg in {"матч","mlbb","тиндер","tinder"}:
        text=("💘 <b>MLBB MATCH SHINOBI</b>\n\nПоиск тиммейтов по серверу, ролям, времени, микрофону и стилю игры.\n\n"
              "<code>.анкета</code> — формат анкеты\n<code>.матч</code> / <code>.mlbb</code> — найти тиммейта\n\n🎮 "+mlbb_random_fact())
    elif arg in {"фарм","валюта"}:
        text=("⛏ <b>ФАРМ ВАЛЮТЫ</b>\n\nРаз в час используй <code>.фарм</code>. Базовая награда случайная: <b>1–100 💎</b>. "
              "Иногда выпадает дополнительный бонус. Начальный баланс — 0 💎.\n\nМонеты используются на Чёрном рынке.")
    else:
        text=("🥷 <b>SHINOBI — игровые функции</b>\n\n<code>.шиноби рынок</code> — 📈 рынок и акции\n<code>.шиноби фарм</code> — ⛏ валюта\n"
              "<code>.шиноби досье</code> — 📁 досье\n<code>.шиноби матч</code> — 💘 тиммейты\n\n🎮 Эти игровые команды разрешены в общем чате.")
    await message.reply(text,parse_mode="HTML",reply_markup=_hub_keyboard())


@dp.message(F.text.casefold().in_({".рынок график", ".график рынка"}))
async def game_market_chart(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    await message.reply(_market_chart_text(message.chat.id),parse_mode="HTML",reply_markup=_market_keyboard())


MARKET_TOP_LIMIT = 15

def _market_top_text(chat_id: int) -> str:
    players = _game_chat(chat_id)["players"]
    rows = []
    for uid, player in players.items():
        if not isinstance(player, dict):
            continue
        try:
            rows.append((_market_price(chat_id, str(uid)), str(uid), player))
        except Exception as exc:
            print(f"[MARKET PLAYER SKIP] uid={uid}: {type(exc).__name__}: {exc}")
    rows.sort(key=lambda row: row[0], reverse=True)
    top = rows[:MARKET_TOP_LIMIT]
    if not top:
        return "📈 Рынок пока пуст. Игроки появятся после активности в чате."
    lines = ["📈 <b>ЧЁРНЫЙ РЫНОК SHINOBI · ТОП 15</b>", ""]
    for i, (price, uid, p) in enumerate(top, 1):
        # @username остаётся настоящим упоминанием, как запросил владелец.
        label = html.escape(_player_label(p))
        lines.append(f"{i}. {label} — <b>{price} 💎</b>")
    return "\n".join(lines)

@dp.message(F.text.casefold() == ".рынок")
async def game_market(message: Message):
    if message.chat.type not in {"group", "supergroup"}: return
    await message.reply(_market_top_text(message.chat.id), parse_mode="HTML", reply_markup=_market_keyboard())

@dp.message(F.text.casefold() == ".фарм")
async def game_farm(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    await _do_farm(message.chat.id,message.from_user,message.reply)


@dp.message(F.text.casefold().startswith(".акции"))
async def game_stock(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    parts=(message.text or "").split(maxsplit=1)
    uid,p=_target_from_message(message, parts[1] if len(parts)>1 else None)
    if not p: await message.reply("Ответь командой <code>.акции</code> на сообщение игрока или используй <code>.акции @username</code>.",parse_mode="HTML",reply_markup=_market_keyboard()); return
    old,price,pct=_market_delta(message.chat.id,uid); arrow="📈" if pct>=0 else "📉"
    await message.reply(f"📊 <b>{_player_label(p)}</b> · {_rarity(message.chat.id,uid)}\nЦена: <b>{price} 💎</b>\nЗа сутки: {arrow} <b>{pct:+.1f}%</b>\n\n{random.choice(MARKET_FACTS)}",parse_mode="HTML",reply_markup=_market_keyboard())

@dp.message(F.text.casefold().startswith(".купить"))
async def game_buy(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    parts=(message.text or "").split(); reply_uid,reply_p=_reply_target(message)
    if reply_p:
        if len(parts)!=2 or not parts[1].isdigit(): await message.reply("Reply-формат: <code>.купить 1</code>",parse_mode="HTML"); return
        uid,target,qty=reply_uid,reply_p,int(parts[1])
    else:
        if len(parts)!=3 or not parts[2].isdigit(): await message.reply("<code>.купить @user 1</code> или reply → <code>.купить 1</code>",parse_mode="HTML"); return
        uid,target=_find_game_player(message.chat.id,parts[1]); qty=int(parts[2])
    if qty<1 or not target: await message.reply("❌ Игрок/количество не найдены."); return
    if uid==str(message.from_user.id): await message.reply("❌ Свои акции покупать нельзя."); return
    buyer=_game_player(message.chat.id,message.from_user); price=_market_price(message.chat.id,uid); total=price*qty; balance=_market_balance(message.chat.id,str(message.from_user.id),0)
    if balance<total: await message.reply(f"❌ Нужно {total} 💎, у тебя {balance} 💎.\n⛏ <code>.фарм</code>",parse_mode="HTML",reply_markup=_market_keyboard()); return
    _set_market_balance(message.chat.id,str(message.from_user.id),balance-total); buyer["balance"]=balance-total
    pf=buyer.setdefault("portfolio",{}); pf[uid]=int(pf.get(uid,0))+qty
    costs=buyer.setdefault("portfolio_cost",{}); costs[uid]=int(costs.get(uid,0))+total; save_data()
    await message.reply(f"💹 <b>СДЕЛКА ЗАКРЫТА</b>\nКуплено: <b>{qty}</b> акц. {_player_label(target)}\nЦена сделки: <b>{total} 💎</b>\nОстаток: <b>{balance-total} 💎</b>\nРедкость актива: <b>{_rarity(message.chat.id,uid)}</b>",parse_mode="HTML",reply_markup=_market_keyboard())

@dp.message(F.text.casefold().startswith(".продать"))
async def game_sell(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    parts=(message.text or "").split(); reply_uid,reply_p=_reply_target(message)
    if reply_p:
        if len(parts)!=2 or not parts[1].isdigit(): await message.reply("Reply: <code>.продать 1</code>",parse_mode="HTML"); return
        uid,target,qty=reply_uid,reply_p,int(parts[1])
    else:
        if len(parts)!=3 or not parts[2].isdigit(): await message.reply("<code>.продать @user 1</code> или reply → <code>.продать 1</code>",parse_mode="HTML"); return
        uid,target=_find_game_player(message.chat.id,parts[1]); qty=int(parts[2])
    me=_game_player(message.chat.id,message.from_user); have=int(me.setdefault("portfolio",{}).get(uid,0)) if uid else 0
    if not target or qty<1 or have<qty: await message.reply(f"❌ Недостаточно акций. У тебя: {have}."); return
    price=_market_price(message.chat.id,uid); revenue=price*qty; oldcost=int(me.setdefault("portfolio_cost",{}).get(uid,0)); avg=(oldcost/have) if have else 0; basis=int(round(avg*qty)); profit=revenue-basis
    me["portfolio"][uid]=have-qty
    if me["portfolio"][uid]<=0: me["portfolio"].pop(uid,None); me["portfolio_cost"].pop(uid,None)
    else: me["portfolio_cost"][uid]=max(0,oldcost-basis)
    bal=_market_balance(message.chat.id,str(message.from_user.id),0)+revenue; _set_market_balance(message.chat.id,str(message.from_user.id),bal); save_data()
    await message.reply(f"💸 Продано <b>{qty}</b> акц. {_player_label(target)} за <b>{revenue} 💎</b>.\nРезультат сделки: <b>{profit:+} 💎</b>\nБаланс: <b>{bal} 💎</b>",parse_mode="HTML",reply_markup=_market_keyboard())

@dp.message(F.text.casefold().startswith(".передать"))
async def game_transfer_money(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    parts=(message.text or "").split(); reply_uid,reply_p=_reply_target(message)
    if reply_p:
        if len(parts)!=2 or not parts[1].isdigit(): await message.reply("Reply: <code>.передать 100</code>",parse_mode="HTML"); return
        uid,target,amount=reply_uid,reply_p,int(parts[1])
    else:
        if len(parts)!=3 or not parts[2].isdigit(): await message.reply("<code>.передать @user 100</code> или reply → <code>.передать 100</code>",parse_mode="HTML"); return
        uid,target=_find_game_player(message.chat.id,parts[1]); amount=int(parts[2])
    if not target or amount<1 or uid==str(message.from_user.id): await message.reply("❌ Некорректная передача."); return
    sid=str(message.from_user.id); bal=_market_balance(message.chat.id,sid,0)
    if bal<amount: await message.reply(f"❌ У тебя только {bal} 💎."); return
    _set_market_balance(message.chat.id,sid,bal-amount); _set_market_balance(message.chat.id,uid,_market_balance(message.chat.id,uid,0)+amount)
    await message.reply(f"💎 {_player_label(_game_player(message.chat.id,message.from_user))} передал {_player_label(target)} <b>{amount} 💎</b>.",parse_mode="HTML",reply_markup=_market_keyboard())

@dp.message(F.text.casefold().startswith(".подарить"))
async def game_gift_share(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    # Формат: .подарить @актив @получатель 1; либо reply получателю: .подарить @актив 1
    parts=(message.text or "").split(); rid,rp=_reply_target(message)
    if rp:
        if len(parts)!=3 or not parts[2].isdigit(): await message.reply("Reply получателю: <code>.подарить @актив 1</code>",parse_mode="HTML"); return
        asset_uid,asset=_find_game_player(message.chat.id,parts[1]); to_uid,to_p=rid,rp; qty=int(parts[2])
    else:
        if len(parts)!=4 or not parts[3].isdigit(): await message.reply("<code>.подарить @актив @получатель 1</code>",parse_mode="HTML"); return
        asset_uid,asset=_find_game_player(message.chat.id,parts[1]); to_uid,to_p=_find_game_player(message.chat.id,parts[2]); qty=int(parts[3])
    me=_game_player(message.chat.id,message.from_user); have=int(me.setdefault("portfolio",{}).get(asset_uid,0)) if asset_uid else 0
    if not asset or not to_p or qty<1 or have<qty or to_uid==str(message.from_user.id): await message.reply("❌ Не удалось подарить акции."); return
    recipient=_game_chat(message.chat.id)["players"][to_uid]; me["portfolio"][asset_uid]=have-qty
    if me["portfolio"][asset_uid]<=0: me["portfolio"].pop(asset_uid,None)
    recipient.setdefault("portfolio",{})[asset_uid]=int(recipient.setdefault("portfolio",{}).get(asset_uid,0))+qty; save_data()
    await message.reply(f"🎁 {_player_label(to_p)} получил <b>{qty}</b> акц. {_player_label(asset)}.",parse_mode="HTML",reply_markup=_market_keyboard())

def _portfolio_text(chat_id:int, uid:str, p:dict):
    players=_game_chat(chat_id)["players"]; bal=_market_balance(chat_id,uid,0); lines=[f"💼 <b>Портфель {_player_label(p)}</b>",f"Баланс: <b>{bal} 💎</b>",""]; total=0; invested=0
    for aid,qty in p.get("portfolio",{}).items():
        if aid in players and int(qty)>0:
            val=_market_price(chat_id,aid)*int(qty); cost=int(p.get("portfolio_cost",{}).get(aid,0)); total+=val; invested+=cost
            lines.append(f"• {_player_label(players[aid])}: {qty} шт. → {val} 💎 ({val-cost:+} 💎)")
    if not p.get("portfolio"): lines.append("Портфель пуст. Начни с <code>.фарм</code> и покупки акций.")
    distinct=len([q for q in p.get("portfolio",{}).values() if int(q)>0]); badge="🥉 Коллекционер" if distinct>=3 else ""; badge="🥈 Инвестор" if distinct>=5 else badge; badge="🥇 Магнат" if distinct>=10 else badge
    lines += [f"\nАктивы: <b>{total} 💎</b>",f"Общий результат: <b>{total-invested:+} 💎</b>"]
    if badge: lines.append(f"Достижение: <b>{badge}</b>")
    if int(p.get("insurance",0)): lines.append(f"🛡 Страховок: <b>{int(p.get('insurance',0))}</b>")
    return "\n".join(lines)

@dp.message(F.text.casefold().startswith(".портфель"))
async def game_portfolio(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    parts=(message.text or "").split(maxsplit=1); uid,p=_target_from_message(message,parts[1] if len(parts)>1 else None)
    if not p: uid=str(message.from_user.id); p=_game_player(message.chat.id,message.from_user)
    await message.reply(_portfolio_text(message.chat.id,uid,p),parse_mode="HTML",reply_markup=_market_keyboard())

async def _claim_dividends(chat_id:int,user,reply_func):
    p=_game_player(chat_id,user); now=time.time(); last=float(p.get("last_dividend",0) or 0)
    if now-last<86400:
        hrs=max(1,int((86400-(now-last)+3599)//3600)); await reply_func(f"⏳ Следующие дивиденды примерно через <b>{hrs} ч.</b>",parse_mode="HTML",reply_markup=_market_keyboard()); return
    players=_game_chat(chat_id)["players"]; value=sum(_market_price(chat_id,uid)*int(q) for uid,q in p.get("portfolio",{}).items() if uid in players)
    if value<=0: await reply_func("💸 Для дивидендов сначала нужны акции в портфеле.",reply_markup=_market_keyboard()); return
    rate=random.uniform(.01,.05); reward=max(1,int(value*rate)); jackpot=0
    # Серия удачи: если 3+ активов сегодня выросли, +10% к дивидендам.
    rising=sum(1 for uid,q in p.get("portfolio",{}).items() if int(q)>0 and uid in players and _market_delta(chat_id,uid)[2]>0)
    streak_bonus=int(reward*.10) if rising>=3 else 0; reward+=streak_bonus
    if random.random()<.02: jackpot=random.randint(500,1500)
    bal=_market_balance(chat_id,str(user.id),0)+reward+jackpot; _set_market_balance(chat_id,str(user.id),bal); p["last_dividend"]=now; save_data()
    extra=f"\n🎰 <b>ДЖЕКПОТ +{jackpot} 💎!</b>" if jackpot else ""
    if streak_bonus: extra += f"\n🔥 Серия роста: <b>+{streak_bonus} 💎</b>"
    await reply_func(f"💸 <b>ДИВИДЕНДЫ</b>\nСтоимость портфеля: {value} 💎\nСтавка дня: {rate*100:.1f}%\nПолучено: <b>+{reward} 💎</b>{extra}\nБаланс: <b>{bal} 💎</b>",parse_mode="HTML",reply_markup=_market_keyboard())

@dp.message(F.text.casefold()==".дивиденды")
async def game_dividends(message:Message):
    if message.chat.type in {"group","supergroup"}: await _claim_dividends(message.chat.id,message.from_user,message.reply)

@dp.message(F.text.casefold()==".рынок топ")
async def game_investor_top(message:Message):
    if message.chat.type not in {"group","supergroup"}:return
    players=_game_chat(message.chat.id)["players"]; rows=[]
    for uid,p in players.items():
        assets=sum(_market_price(message.chat.id,a)*int(q) for a,q in p.get("portfolio",{}).items() if a in players); rows.append((_market_balance(message.chat.id,uid,0)+assets,uid,p,assets))
    rows.sort(reverse=True,key=lambda x:x[0]); lines=["🏆 <b>ТОП ИНВЕСТОРОВ</b>",""]+[f"{i}. {_player_label(p)} — <b>{net} 💎</b> (активы {assets})" for i,(net,uid,p,assets) in enumerate(rows[:10],1)]
    await message.reply("\n".join(lines),parse_mode="HTML",reply_markup=_market_keyboard())

@dp.message(F.text.casefold()==".рынок события")
async def game_market_events(message:Message):
    if message.chat.type not in {"group","supergroup"}:return
    players=list(_game_chat(message.chat.id)["players"].items())
    if not players: await message.reply("Событий пока нет."); return
    sample=random.Random(f"events:{message.chat.id}:{datetime.now().date().isoformat()}").sample(players,min(5,len(players))); lines=["📰 <b>СОБЫТИЯ РЫНКА</b>",""]
    for uid,p in sample:
        _,_,pct=_market_delta(message.chat.id,uid); lines.append(f"• {_player_label(p)} {random.choice(MARKET_EVENT_TEXTS)} — <b>{pct:+.1f}%</b>")
    await message.reply("\n".join(lines),parse_mode="HTML",reply_markup=_market_keyboard())

def _dossier_view(chat_id: int, uid: str, p: dict) -> str:
    now=datetime.now()
    try: started=datetime.fromisoformat(p.get("dossier_started_at") or now.isoformat())
    except Exception: started=now
    age=max(0,(now-started).total_seconds())
    days=min(7,int(age//86400))
    dossier_no=(int(uid)%9000)+100
    if age < 7*86400:
        left=max(1,7-days)
        return (f"📁 <b>ДОСЬЕ №{dossier_no}</b>\n{_player_label(p)}\n\n"
                f"🔎 Сбор информации: <b>{days}/7 дней</b>\n"
                f"Сообщений изучено: <b>{int(p.get('dossier_week_messages',0))}</b>\n"
                f"До полного досье: <b>≈ {left} дн.</b>\n\n"
                "Бот автоматически наблюдает только за активностью внутри чата. Полное досье откроется через неделю.")
    prof=p.get("profile",{})
    roles=prof.get("roles") or "неизвестно"
    preferred=(roles.split(",")[0].strip() if roles else "неизвестно")
    heroes=prof.get("heroes") or "не установлен"
    hero=(heroes.split(",")[0].strip() if heroes else "не установлен")
    words=p.get("dossier_words",{}) or {}
    phrase=max(words,key=words.get) if words else random.choice(["рандомы","ещё одну и спать","го катку","ну почти"])
    night=int(p.get("dossier_night_games",0))
    msgs=int(p.get("dossier_week_messages",0))
    threat="🟢" if msgs<25 else "🟡" if msgs<80 else "🟠" if msgs<180 else "🔴"
    return (f"📁 <b>ДОСЬЕ №{dossier_no}</b>\n{_player_label(p)}\n\n"
            f"Предпочитает: <b>{preferred}</b>\n"
            f"Главный подозреваемый: <b>{hero}</b>\n"
            f"Был замечен в <b>{night}</b> ночных активностях\n"
            f"Любимая фраза: «<b>{phrase}</b>»\n\n"
            f"Уровень угрозы: {threat}\n\n"
            "<i>Сформировано автоматически по активности за 7 дней. Никакие данные вне чата не используются.</i>")


@dp.message(F.text.casefold().startswith(".досье"))
async def game_dossier(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    parts=(message.text or "").split(maxsplit=1)
    ruid,rp=_reply_target(message)
    if rp: uid,p=ruid,rp
    elif len(parts)>1: uid,p=_find_game_player(message.chat.id,parts[1])
    else: uid=str(message.from_user.id); p=_game_player(message.chat.id,message.from_user)
    if not p: await message.reply("❌ На этого участника досье ещё не начало собираться."); return
    await message.reply(_dossier_view(message.chat.id,uid,p),parse_mode="HTML",reply_markup=_hub_keyboard())


@dp.message(F.text.casefold().startswith(".анкета"))
async def game_profile(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    raw=(message.text or "")[len(".анкета"):].strip()
    if not raw: await message.reply("💘 <b>MLBB Match</b>\nЗаполни одной строкой:\n<code>.анкета Ранг | Сервер | Роли | Любимые герои | Время игры | микро/без микро | пот/chill</code>\n\n🎮 "+mlbb_random_fact(),parse_mode="HTML",reply_markup=_hub_keyboard()); return
    vals=[x.strip() for x in raw.split("|")]
    if len(vals)!=7: await message.reply("❌ Нужно 7 полей через символ <b>|</b>. Напиши <code>.анкета</code> для примера.",parse_mode="HTML"); return
    p=_game_player(message.chat.id,message.from_user); p["profile"]={k:v for k,v in zip(("rank","server","roles","heroes","time","mic","style"),vals)}; save_data()
    await message.reply("✅ MLBB-анкета сохранена. Теперь нажми кнопку или напиши <code>.матч</code>.\n\n🎮 "+mlbb_random_fact(),parse_mode="HTML",reply_markup=_hub_keyboard())


@dp.message(F.text.casefold().in_({".матч",".mlbb"}))
async def game_match(message: Message):
    if message.chat.type not in {"group","supergroup"}: return
    me=_game_player(message.chat.id,message.from_user); profile=me.get("profile",{})
    if not profile: await message.reply("Сначала создай анкету командой <code>.анкета</code>.",parse_mode="HTML",reply_markup=_hub_keyboard()); return
    candidates=[]
    for uid,p in _game_chat(message.chat.id)["players"].items():
        if uid==str(message.from_user.id) or not p.get("profile"): continue
        score,reasons=_tinder_score(profile,p["profile"]); candidates.append((score,uid,p,reasons))
    if not candidates: await message.reply("💘 Пока нет других заполненных анкет. Нужен ещё хотя бы один игрок.",reply_markup=_hub_keyboard()); return
    # При одинаковом лучшем проценте выбираем случайного кандидата — повторный поиск не всегда одинаковый.
    best=max(x[0] for x in candidates); pool=[x for x in candidates if x[0]==best]; score,uid,p,reasons=random.choice(pool); q=p["profile"]
    why=", ".join(reasons) if reasons else "есть базовая совместимость"
    await message.reply(f"💘 <b>MATCH {score}%</b>\n{_player_label(me)} + {_player_label(p)}\n\nРанг: {q.get('rank')} · Сервер: {q.get('server')}\nРоли: {q.get('roles')}\nВремя: {q.get('time')} · {q.get('mic')} · {q.get('style')}\n\nПочему: {why}.\n🎮 {mlbb_random_fact()}\n\nЕсть один способ проверить совместимость — сыграть.",parse_mode="HTML",reply_markup=_hub_keyboard())


@dp.callback_query(F.data == "game:market")
async def game_market_callback(callback: CallbackQuery):
    # Отдельный точный handler рынка: не зависит от общего game:* роутера.
    if not callback.message or callback.message.chat.type not in {"group", "supergroup"}:
        await callback.answer()
        return
    try:
        text = _market_top_text(callback.message.chat.id)
        await callback.message.answer(text, parse_mode="HTML", reply_markup=_market_keyboard())
        await callback.answer()
    except Exception as exc:
        print(f"[MARKET CALLBACK ERROR] {type(exc).__name__}: {exc}")
        try:
            await callback.answer("Ошибка рынка. Смотри лог бота.", show_alert=True)
        except Exception:
            pass


@dp.callback_query(F.data.startswith("game:"))
async def game_hub_callback(callback: CallbackQuery):
    if not callback.message or callback.message.chat.type not in {"group","supergroup"}: return
    action=(callback.data or "").split(":",1)[1]
    await callback.answer()
    chat_id=callback.message.chat.id
    if action=="market":
        return
    elif action=="farm":
        await _do_farm(chat_id,callback.from_user,callback.message.answer)
    elif action=="portfolio":
        me=_game_player(chat_id,callback.from_user); players=_game_chat(chat_id)["players"]; me["balance"]=_market_balance(chat_id,str(callback.from_user.id),0)
        total=sum(_market_price(chat_id,uid)*int(qty) for uid,qty in me.get("portfolio",{}).items() if uid in players)
        await callback.message.answer(f"💼 <b>{_player_label(me)}</b>\nБаланс: <b>{me['balance']} 💎</b>\nАктивы: <b>{total} 💎</b>\nНафармлено: <b>{int(me.get('farm_total',0))} 💎</b>",parse_mode="HTML",reply_markup=_market_keyboard())
    elif action=="chart": await callback.message.answer(_market_chart_text(chat_id),parse_mode="HTML",reply_markup=_market_keyboard())
    elif action=="dividends": await _claim_dividends(chat_id,callback.from_user,callback.message.answer)
    elif action=="investors":
        players=_game_chat(chat_id)["players"]; rows=[]
        for uid,p in players.items():
            assets=sum(_market_price(chat_id,a)*int(q) for a,q in p.get("portfolio",{}).items() if a in players); rows.append((_market_balance(chat_id,uid,0)+assets,p))
        rows.sort(reverse=True,key=lambda x:x[0]); await callback.message.answer("🏆 <b>ТОП ИНВЕСТОРОВ</b>\n\n"+"\n".join(f"{i}. {_player_label(p)} — <b>{net} 💎</b>" for i,(net,p) in enumerate(rows[:10],1)),parse_mode="HTML",reply_markup=_market_keyboard())
    elif action=="fact": await callback.message.answer(random.choice(MARKET_FACTS),reply_markup=_market_keyboard())
    elif action=="mlbbfact": await callback.message.answer("🎮 <b>Факт MLBB</b>\n\n"+mlbb_random_fact(),parse_mode="HTML",reply_markup=_hub_keyboard())
    elif action=="dossier":
        p=_game_player(chat_id,callback.from_user); await callback.message.answer(_dossier_view(chat_id,str(callback.from_user.id),p),parse_mode="HTML",reply_markup=_hub_keyboard())
    elif action=="match":
        await callback.message.answer("💘 Для полного поиска напиши <code>.матч</code>. Если анкеты ещё нет — <code>.анкета</code>.",parse_mode="HTML",reply_markup=_hub_keyboard())
    elif action=="help":
        await callback.message.answer("🥷 <b>SHINOBI</b>\n<code>.шиноби рынок</code> · <code>.шиноби фарм</code> · <code>.шиноби досье</code> · <code>.шиноби матч</code>",parse_mode="HTML",reply_markup=_hub_keyboard())


_ad_info.setdefault("join_request", True)

# Динамические администраторы сохраняются в data.json. Владелец всегда имеет доступ.
for _uid in data.get("bot_admins", []):
    try:
        _uid = int(_uid)
        if _uid not in ADMIN_IDS:
            ADMIN_IDS.append(_uid)
    except (TypeError, ValueError):
        pass
init_schedule_system(ADMIN_IDS)


def _sync_excel_backup_safe():
    """Обновляет постоянный Excel-снимок базы после каждого изменения data.json."""
    try:
        creator = globals().get("create_excel")
        if callable(creator):
            creator()
    except Exception as error:
        print("Не удалось обновить Excel-резерв:", error)


def save_data():

    # Не даём любому старому обработчику перезаписать переключатель, изменённый на сайте.
    try:
        from db_storage import get_setting
        _toggle = get_setting("server_tags_enabled", None)
        if isinstance(_toggle, bool):
            data["server_tags_enabled"] = _toggle
    except Exception:
        pass

    payload = json.dumps(
        data,
        ensure_ascii=False,
        indent=4
    )

    temp_file = os.path.join(
        BASE_DIR,
        f"data.{os.getpid()}.{int(time.time() * 1000)}.tmp"
    )

    try:

        with open(temp_file, "w", encoding="utf-8") as file:

            file.write(payload)
            file.flush()
            os.fsync(file.fileno())

        os.replace(temp_file, DATA_FILE)
        _sync_excel_backup_safe()
        return

    except OSError as error:

        print(
            "Не удалось атомарно сохранить data.json:",
            error
        )

    finally:

        if os.path.exists(temp_file):

            try:
                os.remove(temp_file)
            except OSError:
                pass

    try:

        with open(DATA_FILE, "w", encoding="utf-8") as file:

            file.write(payload)
            file.flush()

        _sync_excel_backup_safe()

    except OSError as error:

        print(
            "Не удалось сохранить data.json:",
            error
        )


# =========================================================
# ПОЛЬЗОВАТЕЛИ
# =========================================================

def get_user(user_id: int):

    uid = str(user_id)

    if uid not in data["users"]:

        data["users"][uid] = {
            "id": user_id,
            "name": "",
            "username": "",
            "links": [],
            "total_invited": 0,
            "invited_users": []
        }

    user = data["users"][uid]

    # Совместимость со старым data.json
    user.setdefault("id", user_id)
    user.setdefault("name", "")
    user.setdefault("username", "")
    user.setdefault("links", [])
    user.setdefault("total_invited", 0)
    user.setdefault("invited_users", [])

    return user


def update_user(tg_user):

    user = get_user(tg_user.id)

    user["id"] = tg_user.id
    user["name"] = tg_user.full_name
    user["username"] = tg_user.username or ""

    save_data()

    return user


def display_user(user):

    username = user.get("username", "")

    if username:
        return f"@{username}"

    return user.get("name", "Пользователь")


def is_admin(user_id: int):

    return user_id in ADMIN_IDS


def _staff_random_admin_allowed(message: Message) -> bool:
    return bool(message.from_user and message.from_user.id in ADMIN_IDS)

def _random_command_allowed(message: Message) -> bool:
    return message.chat.type in ("group", "supergroup") and _staff_random_admin_allowed(message)

# =========================================================
# SHINOBI MODERATION: PRED / WARN / MUTE / BAN
# =========================================================
MOD_PRED_TTL = 7 * 86400
MOD_WARN_TTL = 7 * 86400
MOD_WARN_LIMIT = 3
MOD_WARN_BAN_SECONDS = 2 * 86400

def _mod_root():
    root = data.setdefault("shinobi_moderation", {})
    root.setdefault("records", [])
    root.setdefault("next_id", 1)
    return root

def _mod_active_records(chat_id, user_id, kind=None):
    now=int(time.time()); out=[]
    for r in _mod_root()["records"]:
        if int(r.get("chat_id",0))!=int(chat_id) or int(r.get("user_id",0))!=int(user_id): continue
        if kind and r.get("kind")!=kind: continue
        if r.get("revoked"): continue
        exp=int(r.get("expires_at",0) or 0)
        if exp and exp<=now: continue
        out.append(r)
    return out

def _mod_rule_from_text(text):
    m=re.search(r"правил[оа]?\s*(\d+)\s*[,;:]?\s*(?:пункт|п\.?)?\s*(\d+)?", text or "", re.I)
    return (int(m.group(1)), int(m.group(2) or 1)) if m else (0,1)

def _mod_add_record(chat_id,user_id,kind,reason,rule=0,subpoint=1,actor_id=0,actor_name="AI",source="admin",duration=0,message_id=0):
    root=_mod_root(); now=int(time.time()); rid=int(root.get("next_id",1)); root["next_id"]=rid+1
    rec={"id":rid,"ts":now,"chat_id":int(chat_id),"user_id":int(user_id),"kind":kind,"reason":(reason or "Без причины")[:500],"rule":int(rule or 0),"subpoint":int(subpoint or 1),"actor_id":int(actor_id or 0),"actor_name":actor_name or "AI","source":source,"duration":int(duration or 0),"expires_at":now+int(duration) if duration else 0,"revoked":False,"message_id":int(message_id or 0)}
    root["records"].append(rec); del root["records"][:-2000]; save_data(); return rec

def _mod_duration(text):
    t=(text or "").casefold().strip()
    if t in {"навсегда","perm","permanent"}: return 0
    m=re.fullmatch(r"(\d+)\s*(м|мин|m|ч|h|д|d)",t)
    if not m: return None
    n=int(m.group(1)); unit=m.group(2)
    return n*(60 if unit in {"м","мин","m"} else 3600 if unit in {"ч","h"} else 86400)

def _mod_fmt_seconds(sec):
    if not sec: return "навсегда"
    if sec%86400==0: return f"{sec//86400} дн."
    if sec%3600==0: return f"{sec//3600} ч."
    return f"{max(1,sec//60)} мин."

async def _mod_target(message, tokens):
    if message.reply_to_message and message.reply_to_message.from_user:
        return message.reply_to_message.from_user, tokens
    if not tokens: return None,tokens
    raw=tokens[0].lstrip("@").casefold()
    uid=None
    if raw.isdigit(): uid=int(raw)
    else:
        # 1) Seen-members cache for the CURRENT chat. This is the most reliable
        # username -> Telegram ID source available to a bot without a full member-list API.
        seen = data.get("group_seen_members", {}).get(str(message.chat.id), {})
        iterable = seen.values() if isinstance(seen, dict) else seen
        for u in iterable:
            if str(u.get("username", "")).lstrip("@").casefold() == raw:
                uid = int(u.get("id", 0) or 0)
                if uid:
                    break

        # 2) Global user cache used by the rest of Shinobi Bot.
        if not uid:
            users = data.get("users", {})
            iterable = users.values() if isinstance(users, dict) else users
            for u in iterable:
                if str(u.get("username", "")).lstrip("@").casefold() == raw:
                    uid = int(u.get("id", 0) or 0)
                    if uid:
                        break

        # 3) Moderation history fallback (useful after migrations/older data.json).
        if not uid:
            for r in reversed(_mod_root().get("records", [])):
                if int(r.get("chat_id", 0) or 0) != int(message.chat.id):
                    continue
                if str(r.get("username", "")).lstrip("@").casefold() == raw:
                    uid = int(r.get("user_id", 0) or 0)
                    if uid:
                        break
    if not uid: return None,tokens
    try: member=await bot.get_chat_member(message.chat.id,uid); return member.user,tokens[1:]
    except Exception: return None,tokens

async def _mod_staff_allowed(message):
    return bool(message.from_user and await _is_admin_staff_member(message.from_user.id))

async def _mod_staff_log(rec,user,action_text):
    try:
        who=f"@{user.username}" if user.username else f"ID <code>{user.id}</code>"
        rule=f"Правило {rec['rule']}, пункт {rec['subpoint']}" if rec.get('rule') else "не указано"
        text=(f"🛡 <b>SHINOBI MODERATION</b>\n\n"
              f"👤 <b>{html.escape(user.full_name)}</b> ({who})\n"
              f"⚖️ Действие: <b>{html.escape(action_text)}</b>\n"
              f"📜 Правило: <b>{rule}</b>\n"
              f"🔎 Причина: {html.escape(rec.get('reason',''))}\n"
              f"👮 Выдал: <b>{html.escape(rec.get('actor_name',''))}</b>\n"
              f"🆔 Дело: <code>#{rec['id']}</code>")
        await bot.send_message(ADMIN_STAFF_CHAT_ID,text,parse_mode="HTML")
    except Exception as exc: print(f"[MOD STAFF LOG] {exc}")

async def _mod_issue_pred(chat_id,user,reason,rule,subpoint,actor_id=0,actor_name="AI",source="ai",message_id=0):
    rec=_mod_add_record(chat_id,user.id,"pred",reason,rule,subpoint,actor_id,actor_name,source,MOD_PRED_TTL,message_id)
    rec.update({"user_name":user.full_name,"username":user.username or ""}); save_data()
    await _mod_staff_log(rec,user,"⚠️ ПРЕД • 7 дней"); return rec

async def _mod_issue_warn(chat_id,user,reason,rule,subpoint,actor_id=0,actor_name="AI",source="ai",message_id=0):
    rec=_mod_add_record(chat_id,user.id,"warn",reason,rule,subpoint,actor_id,actor_name,source,MOD_WARN_TTL,message_id)
    rec.update({"user_name":user.full_name,"username":user.username or ""}); save_data()
    count=len(_mod_active_records(chat_id,user.id,"warn")); banned=False
    if count>=MOD_WARN_LIMIT:
        try:
            until=datetime.now(ZoneInfo("UTC"))+timedelta(seconds=MOD_WARN_BAN_SECONDS)
            await bot.ban_chat_member(chat_id,user.id,until_date=until)
            banrec=_mod_add_record(chat_id,user.id,"ban","3/3 активных WARN",rule,subpoint,0,"SYSTEM","system",MOD_WARN_BAN_SECONDS,message_id)
            banrec.update({"user_name":user.full_name,"username":user.username or ""})
            for x in _mod_active_records(chat_id,user.id,"warn"): x["revoked"]=True; x["revoked_reason"]="3/3 → бан 2 дня"
            save_data(); await _mod_staff_log(banrec,user,"⛔ Автобан • 2 дня (3/3 WARN)"); banned=True
        except Exception as exc: print(f"[MOD AUTO BAN] {exc}")
    await _mod_staff_log(rec,user,f"🟨 WARN {min(count,3)}/3 • 7 дней"); return rec,count,banned

async def _mod_issue_mute(chat_id,user,seconds,reason,rule,subpoint,actor_id=0,actor_name="AI",source="ai",message_id=0):
    until=datetime.now(ZoneInfo("UTC"))+timedelta(seconds=seconds)
    await bot.restrict_chat_member(chat_id,user.id,ChatPermissions(can_send_messages=False),until_date=until)
    rec=_mod_add_record(chat_id,user.id,"mute",reason,rule,subpoint,actor_id,actor_name,source,seconds,message_id)
    rec.update({"user_name":user.full_name,"username":user.username or ""}); save_data()
    await _mod_staff_log(rec,user,f"🔇 Мут • {_mod_fmt_seconds(seconds)}"); return rec

async def _mod_issue_ban(chat_id,user,seconds,reason,rule,subpoint,actor_id=0,actor_name="AI",source="ai",message_id=0):
    until=datetime.now(ZoneInfo("UTC"))+timedelta(seconds=seconds) if seconds else None
    if until: await bot.ban_chat_member(chat_id,user.id,until_date=until)
    else: await bot.ban_chat_member(chat_id,user.id)
    rec=_mod_add_record(chat_id,user.id,"ban",reason,rule,subpoint,actor_id,actor_name,source,seconds,message_id)
    rec.update({"user_name":user.full_name,"username":user.username or ""}); save_data()
    await _mod_staff_log(rec,user,f"⛔ Бан • {_mod_fmt_seconds(seconds)}"); return rec

async def _manual_mod_common(message, command):
    if message.chat.type not in {"group","supergroup"} or not await _mod_staff_allowed(message): return
    text=(message.text or "").strip(); rest=text[len(command):].strip(); tokens=rest.split()
    if command==".наказание":
        rid=int(tokens[0].lstrip("#")) if tokens and tokens[0].lstrip("#").isdigit() else 0
        rec=next((r for r in _mod_root()["records"] if r.get("id")==rid),None)
        return await message.reply((f"🆔 #{rid}\nТип: {rec.get('kind')}\nПричина: {rec.get('reason')}\nПравило {rec.get('rule')}, пункт {rec.get('subpoint',1)}\nВыдал: {rec.get('actor_name')}" if rec else "❌ Дело не найдено."))
    user,tokens=await _mod_target(message,tokens)
    if command in {".преды",".варны",".нарушения",".модлог"}:
        if not user: return await message.reply("Используй reply или укажи @username / ID пользователя.")
        now=int(time.time()); rows=[r for r in _mod_root()["records"] if int(r.get("chat_id",0))==message.chat.id and int(r.get("user_id",0))==user.id]
        if command==".преды": rows=[r for r in rows if r.get("kind")=="pred" and not r.get("revoked") and (not r.get("expires_at") or r.get("expires_at")>now)]
        elif command==".варны": rows=[r for r in rows if r.get("kind")=="warn" and not r.get("revoked") and (not r.get("expires_at") or r.get("expires_at")>now)]
        rows=rows[-10:]
        if not rows: return await message.reply("✅ Записей не найдено.")
        lines=[]
        for r in reversed(rows):
            rule=f"Правило {r.get('rule')}, пункт {r.get('subpoint',1)}" if r.get('rule') else "без пункта"; lines.append(f"#{r['id']} • {r['kind'].upper()} • {rule} • {r.get('reason','')}")
        return await message.reply("📋 <b>История модерации</b>\n\n"+"\n".join(html.escape(x) for x in lines),parse_mode="HTML")
    if not user: return await message.reply("Используй reply или укажи @username / ID пользователя.")
    actor=message.from_user; actor_name="@"+actor.username if actor.username else actor.full_name
    if user.id==bot.id: return
    reason=" ".join(tokens).strip() or "Нарушение правил"; rule,sub=_mod_rule_from_text(reason)
    if command in {".пред",".варн"}:
        if command==".пред": rec=await _mod_issue_pred(message.chat.id,user,reason,rule,sub,actor.id,actor_name,"admin",message.message_id); act="⚠️ ПРЕД на 7 дней"
        else:
            rec,count,banned=await _mod_issue_warn(message.chat.id,user,reason,rule,sub,actor.id,actor_name,"admin",message.message_id); act=f"🟨 WARN {min(count,3)}/3"+(" → ⛔ бан 2 дня" if banned else "")
        return await message.reply(f"{act}\n👤 {html.escape(user.full_name)}\n📜 Правило {rule}, пункт {sub}\n🆔 #{rec['id']}",parse_mode="HTML")
    if command in {".снятьпред",".снятьварн"}:
        kind="pred" if command==".снятьпред" else "warn"; rows=_mod_active_records(message.chat.id,user.id,kind)
        if not rows: return await message.reply("Активных записей нет.")
        rows[-1]["revoked"]=True; rows[-1]["revoked_by"]=actor.id; save_data(); return await message.reply("✅ Последняя запись снята.")
    if command==".размут":
        await bot.restrict_chat_member(message.chat.id,user.id,ChatPermissions(can_send_messages=True,can_send_audios=True,can_send_documents=True,can_send_photos=True,can_send_videos=True,can_send_video_notes=True,can_send_voice_notes=True,can_send_polls=True,can_send_other_messages=True,can_add_web_page_previews=True,can_invite_users=True)); return await message.reply("🔊 Мут снят.")
    if command==".разбан": await bot.unban_chat_member(message.chat.id,user.id,only_if_banned=True); return await message.reply("✅ Пользователь разбанен.")
    if command in {".мут",".бан"}:
        if not tokens: return await message.reply(f"Пример: {command} @user 2ч Правило 1, пункт 2")
        seconds=_mod_duration(tokens[0]);
        if seconds is None: return await message.reply("❌ Срок: 30м, 2ч, 1д или навсегда (для бана).")
        tokens=tokens[1:]; reason=" ".join(tokens).strip() or "Нарушение правил"; rule,sub=_mod_rule_from_text(reason)
        if command==".мут":
            if not seconds: return await message.reply("❌ Мут должен иметь срок.")
            rec=await _mod_issue_mute(message.chat.id,user,seconds,reason,rule,sub,actor.id,actor_name,"admin",message.message_id); act=f"🔇 Мут {_mod_fmt_seconds(seconds)}"
        else: rec=await _mod_issue_ban(message.chat.id,user,seconds,reason,rule,sub,actor.id,actor_name,"admin",message.message_id); act=f"⛔ Бан {_mod_fmt_seconds(seconds)}"
        return await message.reply(f"{act}\n👤 {html.escape(user.full_name)}\n📜 Правило {rule}, пункт {sub}\n🆔 #{rec['id']}",parse_mode="HTML")

@dp.message(F.text.func(lambda t: bool(t) and _is_moderation_command(t)))
async def shinobi_moderation_commands(message: Message):
    command=" ".join((message.text or "").casefold().split()).split()[0]
    await _manual_mod_common(message,command)

@dp.message(F.text.casefold() == ".мод ии вкл")
async def ai_mod_enable(message: Message):
    if not message.from_user or message.chat.type not in ("group", "supergroup"):
        return
    if not await _is_admin_staff_member(message.from_user.id):
        return await message.reply("⛔ Команда доступна только участникам Admin Staff.")
    _ai_mod_chat(message.chat.id)["enabled"] = True
    save_data()
    await message.reply("🛡 <b>AI-модерация включена.</b>", parse_mode="HTML")

@dp.message(F.text.casefold() == ".мод ии выкл")
async def ai_mod_disable(message: Message):
    if not message.from_user or message.chat.type not in ("group", "supergroup"):
        return
    if not await _is_admin_staff_member(message.from_user.id):
        return await message.reply("⛔ Команда доступна только участникам Admin Staff.")
    _ai_mod_chat(message.chat.id)["enabled"] = False
    save_data()
    await message.reply("🛡 <b>AI-модерация выключена для этого чата.</b>", parse_mode="HTML")

@dp.message(F.text.casefold() == ".мод ии статус")
async def ai_mod_status(message: Message):
    if not message.from_user:
        return
    if not await _is_admin_staff_member(message.from_user.id):
        return await message.reply("⛔ Команда доступна только участникам Admin Staff.")
    enabled = _ai_mod_enabled(message.chat.id) if message.chat.type in ("group", "supergroup") else _ai_mod_root().get("enabled", True)
    await message.reply(
        "🛡 <b>AI-модерация</b>\n\n"
        f"Статус: <b>{'включена ✅' if enabled else 'выключена ❌'}</b>\n"
        f"Groq API: <b>{'настроен ✅' if GROQ_API_KEY else 'ключ не задан ⚠️'}</b>\n"
        f"Gemini резерв: <b>{'настроен ✅' if GEMINI_API_KEY else 'ключ не задан ⚠️'}</b>\n"
        f"Модель: <code>{html.escape(AI_MODERATION_MODEL)}</code>\n"
        f"Порог уверенности: <b>{AI_MODERATION_MIN_CONFIDENCE:.0%}</b>\n"
        "Контекст: последние сообщения чата + история нарушений.",
        parse_mode="HTML")

@dp.message(F.text.casefold() == ".мод ии тест")
async def ai_mod_test_empty(message: Message):
    if not message.from_user:
        return
    if not await _is_admin_staff_member(message.from_user.id):
        return await message.reply("⛔ Команда доступна только участникам Admin Staff.")
    await message.reply("🧪 Укажи текст после команды.\nПример: <code>.мод ии тест я ебал вашу маму</code>", parse_mode="HTML")

@dp.message(F.text.casefold().startswith(".мод ии тест "))
async def ai_mod_test(message: Message):
    if not message.from_user:
        return
    if not await _is_admin_staff_member(message.from_user.id):
        return await message.reply("⛔ Команда доступна только участникам Admin Staff.")
    test_text = (message.text or "")[len(".мод ии тест "):].strip()
    if not test_text:
        return await message.reply("🧪 Укажи текст для проверки.")
    try:
        verdict = await _ai_mod_classify(
            message,
            _ai_mod_history(_ai_mod_user_state(message.chat.id, message.from_user.id)),
            text_override=test_text,
        )
    except Exception as exc:
        print(f"[AI MOD TEST] {type(exc).__name__}: {exc}")
        return await message.reply(f"⚠️ Ошибка теста AI: <code>{html.escape(type(exc).__name__)}</code>. Подробности в логах Railway.", parse_mode="HTML")
    if not verdict:
        return await message.reply("⚠️ Groq и Gemini не смогли вернуть корректное решение. Наказание не применено. Подробности — в логах Railway.")
    await message.reply("🧪 <b>Тест AI-модерации — без наказания</b>\n\n<pre>" + html.escape(json.dumps(verdict, ensure_ascii=False, indent=2)) + "</pre>", parse_mode="HTML")

@dp.message(F.text.casefold() == ".рандом вкл")
async def staff_random_enable(message: Message):
    if not _random_command_allowed(message):
        return
    state = _random_chat_state(message.chat.id)
    state["enabled"] = True
    state["title"] = message.chat.title or str(message.chat.id)
    _schedule_next_random(state)
    save_data()
    await message.reply("✅ Рандомные ответы в этом чате включены.")

@dp.message(F.text.casefold() == ".рандом выкл")
async def staff_random_disable(message: Message):
    if not _random_command_allowed(message):
        return
    state = _random_chat_state(message.chat.id)
    state["enabled"] = False
    save_data()
    await message.reply("⛔ Рандомные ответы в этом чате выключены.")

@dp.message(F.text.casefold() == ".рандом статус")
async def staff_random_status(message: Message):
    if not _random_command_allowed(message):
        return
    state = _random_chat_state(message.chat.id)
    enabled = "ВКЛ ✅" if state.get("enabled", True) else "ВЫКЛ ⛔"
    lo = int(state.get("min_minutes", RANDOM_MIN_MINUTES))
    hi = int(state.get("max_minutes", RANDOM_MAX_MINUTES))
    status_text = (
        f"🎲 <b>Рандом — {message.chat.title or 'этот чат'}</b>\n\n"
        f"Статус: <b>{enabled}</b>\n"
        f"Интервал: <b>{lo}–{hi} мин.</b>\n"
        f"Фраз в текущем цикле: <b>{len(state.get('messages', []))}/150</b>\n"
        f"Целей для ответа: <b>{len(state.get('targets', []))}</b>\n"
        f"ИИ-ответы: <b>{'ВКЛ ✅' if state.get('ai_enabled', True) else 'ВЫКЛ ⛔'}</b>\n"
        f"Gemini API: <b>{'настроен ✅' if GEMINI_API_KEY else 'ключ не задан ⚠️'}</b>\n"
        f"Groq резерв: <b>{'настроен ✅' if GROQ_API_KEY else 'ключ не задан ⚠️'}</b>\n\n"
        "При достижении 150 фраз база автоматически очищается."
    )
    await message.reply(status_text, parse_mode="HTML")

@dp.message(F.text.casefold().startswith(".рандом интервал"))
async def staff_random_interval(message: Message):
    if not _random_command_allowed(message):
        return
    parts = (message.text or "").split()
    if len(parts) != 4:
        await message.reply("Использование: <code>.рандом интервал 10 60</code>", parse_mode="HTML")
        return
    try:
        lo, hi = int(parts[2]), int(parts[3])
    except ValueError:
        await message.reply("❌ Интервал нужно указать числами в минутах.")
        return
    if lo < 1 or hi < lo or hi > 1440:
        await message.reply("❌ Допустимо: от 1 до 1440 минут, максимум должен быть ≥ минимума.")
        return
    state = _random_chat_state(message.chat.id)
    state["min_minutes"] = lo
    state["max_minutes"] = hi
    _schedule_next_random(state)
    save_data()
    await message.reply(f"⏱ Интервал этого чата изменён: <b>{lo}–{hi} мин.</b>", parse_mode="HTML")

@dp.message(F.text.casefold() == ".рандом сейчас")
async def staff_random_now(message: Message):
    if not _random_command_allowed(message):
        return
    ok, info = await _send_random_reply(message.chat.id)
    if not ok:
        await message.reply(f"❌ {info}")


@dp.message(F.text.casefold() == ".ии вкл")
async def ai_replies_enable(message: Message):
    if not _random_command_allowed(message):
        return
    state = _random_chat_state(message.chat.id)
    state["ai_enabled"] = True
    save_data()
    if GEMINI_API_KEY or GROQ_API_KEY:
        await message.reply("🤖 ИИ-ответы в этом чате включены.")
    else:
        await message.reply("⚠️ ИИ включён, но GEMINI_API_KEY и GROQ_API_KEY не заданы. Пока будет работать старый резервный режим.")

@dp.message(F.text.casefold() == ".ии выкл")
async def ai_replies_disable(message: Message):
    if not _random_command_allowed(message):
        return
    state = _random_chat_state(message.chat.id)
    state["ai_enabled"] = False
    save_data()
    await message.reply("⛔ ИИ-ответы в этом чате выключены. Рандомные ответы продолжают работать в резервном режиме.")

@dp.message(F.text.casefold() == ".ии статус")
async def ai_replies_status(message: Message):
    if not _random_command_allowed(message):
        return
    state = _random_chat_state(message.chat.id)
    enabled = "ВКЛ ✅" if state.get("ai_enabled", True) else "ВЫКЛ ⛔"
    gemini_api = "настроен ✅" if GEMINI_API_KEY else "не настроен ⚠️"
    groq_api = "настроен ✅" if GROQ_API_KEY else "не настроен ⚠️"
    await message.reply(
        f"🤖 <b>ИИ-ответы</b>\n\nСтатус: <b>{enabled}</b>\nGemini API: <b>{gemini_api}</b>\nGroq резерв: <b>{groq_api}</b>\n"
        f"Модель Gemini: <code>{GEMINI_MODEL}</code>\nКонтекст: <b>{AI_CONTEXT_MESSAGES} сообщений</b>",
        parse_mode="HTML",
    )


# =========================================================
# КЛАВИАТУРА
# =========================================================

def quick_reply_keyboard(user_id: int):
    # Обычные пользователи видят все доступные им функции бота.
    rows = [
        [KeyboardButton(text="🔗 Создать ссылку"), KeyboardButton(text="📊 Моя статистика")],
        [KeyboardButton(text="🔗 Мои ссылки"), KeyboardButton(text="👥 Приглашённые")],
        [KeyboardButton(text="🏆 Топ"), KeyboardButton(text="📋 Мой график")],
        [KeyboardButton(text="🏖 Запросить выходной"), KeyboardButton(text="🗓 График / заполнить")],
    ]

    # Только владелец видит управление техническими работами.
    if user_id in OWNER_IDS:
        rows.append([KeyboardButton(text="🛠 Тех. работы")])

    # У администраторов дополнительно отдельный блок административных кнопок.
    if is_admin(user_id):
        rows.extend([
            [KeyboardButton(text="⚙️ Админ-панель"), KeyboardButton(text="🗓 Управление графиком")],
            [KeyboardButton(text="📥 Скачать Excel"), KeyboardButton(text="📊 Общая статистика")],
            [KeyboardButton(text="📅 Excel с графиками"), KeyboardButton(text="🏖 Excel с выходными")],
        ])

    return ReplyKeyboardMarkup(
        keyboard=rows,
        resize_keyboard=True,
        is_persistent=True,
        input_field_placeholder="Выбери действие или напиши сообщение…",
    )

def main_keyboard(user_id: int):

    buttons = [
        [
            InlineKeyboardButton(
                text="🔗 Создать ссылку",
                callback_data="create_link"
            )
        ],
        [
            InlineKeyboardButton(
                text="📊 Моя статистика",
                callback_data="my_stats"
            ),
            InlineKeyboardButton(
                text="🔗 Мои ссылки",
                callback_data="my_links"
            )
        ],
        [
            InlineKeyboardButton(
                text="👥 Приглашённые",
                callback_data="my_refs"
            ),
            InlineKeyboardButton(
                text="🏆 Топ",
                callback_data="top"
            )
        ]
    ]

    buttons.append([
        InlineKeyboardButton(text="📋 Мой график", callback_data="sch_mine"),
        InlineKeyboardButton(text="🏖 Запросить выходной", callback_data="sch_dayoff")
    ])
    buttons.append([InlineKeyboardButton(text="🗓 График / заполнить", callback_data="sch_menu")])

    if user_id in OWNER_IDS:
        maintenance_on = bool(data.get("maintenance_mode", False))
        buttons.append([
            InlineKeyboardButton(
                text=("🟢 Завершить тех. работы" if maintenance_on else "🛠 Тех. работы"),
                callback_data="maintenance_menu"
            )
        ])

    if is_admin(user_id):

        buttons.append([
            InlineKeyboardButton(
                text="⚙️ Админ-панель",
                callback_data="admin_panel"
            )
        ])

    return InlineKeyboardMarkup(
        inline_keyboard=buttons
    )


def admin_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📥 Скачать Excel",
                    callback_data="export_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📊 Общая статистика",
                    callback_data="admin_stats"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🔴 Shinobi Stats — создать отчёт",
                    callback_data="shinobi_stats_report"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏆 Топ",
                    callback_data="top"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🗓 Управление графиком",
                    callback_data="admin_schedule"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📅 Excel с графиками",
                    callback_data="export_schedule_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏖 Excel с выходными",
                    callback_data="export_dayoff_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="👥 Пользователи бота",
                    callback_data="admin_users_0"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏘 Группы и доступ к командам",
                    callback_data="admin_groups_0"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📩 Заявки на вступление",
                    callback_data="admin_join_requests"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📣 Ссылка для рекламы",
                    callback_data="advertising_link"
                )
            ]
            ,*([[InlineKeyboardButton(text="🌐 Открыть Web App", web_app=WebAppInfo(url=WEBAPP_URL))]] if WEBAPP_URL else [])
        ]
    )


# =========================================================
# /start
# =========================================================

@dp.message(Command("start"))
async def start_handler(message: Message, state: FSMContext):

    # Персональные ссылки создаём только через ЛС
    if message.chat.type != "private":
        return

    update_user(message.from_user)

    payload = ""
    if message.text:
        parts = message.text.split(maxsplit=1)
        if len(parts) > 1:
            payload = parts[1].strip()

    if payload == "schedule":
        await show_schedule_menu(message, state)
        return

    await message.answer(
        "🥷 <b>SHINOBI TEAM BOT</b>\n\n"
        f"👤 {message.from_user.full_name}\n"
        f"🆔 ID: <code>{message.from_user.id}</code>\n\n"
        "Ты успешно авторизован в системе.\n\n"
        "Используй автоматизированную систему Shinobi Team для управления графиком, персональными ссылками и рабочими инструментами 👇",
        reply_markup=main_keyboard(message.from_user.id),
        parse_mode="HTML"
    )
    await message.answer("⚡ Быстрые действия закреплены под строкой ввода.", reply_markup=quick_reply_keyboard(message.from_user.id))

@dp.message(F.text == "📋 Мой график")
async def quick_my_schedule(message: Message):
    # Переиспользуем текстовую команду, которую обрабатывает модуль графика
    from schedule_system import personal_week_schedule
    await personal_week_schedule(message)

@dp.message(F.text == "🏖 Запросить выходной")
async def quick_dayoff(message: Message):
    if message.chat.type != "private":
        info=await message.bot.get_me()
        await message.answer("🏖 Запрос выходного оформляется в ЛС с ботом.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Открыть бота",url=f"https://t.me/{info.username}?start=schedule")]]))
        return
    await message.answer("🏖 Открой меню графика и нажми «Запросить выходной».", reply_markup=main_keyboard(message.from_user.id))

@dp.message(Command("menu"))
async def quick_menu(message: Message):
    await message.answer("⚡ Быстрые действия:", reply_markup=quick_reply_keyboard(message.from_user.id))

@dp.message(Command("link"))
async def link_command(message: Message):

    # Если команду написали в группе
    if message.chat.type in {"group", "supergroup"}:

        if message.chat.id != CHAT_ID:
            return

        bot_info = await bot.get_me()

        bot_url = f"https://t.me/{bot_info.username}?start=link"

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="🔗 Получить свою ссылку",
                        url=bot_url
                    )
                ]
            ]
        )

        await message.reply(
            "🥷 <b>SHINOBI TEAM</b>\n\n"
            f"👤 {message.from_user.full_name}\n\n"
            "Чтобы получить персональную пригласительную ссылку, "
            "перейди в ЛС бота 👇",
            reply_markup=keyboard,
            parse_mode="HTML"
        )

        return

    # Если /link написали сразу в ЛС
    if message.chat.type == "private":

        await create_ref_link(
            message.from_user,
            message
        )

@dp.message(F.text == "..ссылка")
async def ref_link_chat(message: Message):

    # Команда работает только в группе
    if message.chat.type not in {"group", "supergroup"}:
        return

    # Только в нужной группе
    if message.chat.id != CHAT_ID:
        return

    bot_info = await bot.get_me()

    bot_url = f"https://t.me/{bot_info.username}?start=link"

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🔗 Получить ссылку",
                    url=bot_url
                )
            ]
        ]
    )

    await message.reply(
        f"👤 <b>{message.from_user.full_name}</b>\n\n"
        "Для получения персональной реферальной ссылки "
        "перейди в ЛС бота и нажми <b>START</b> 👇",
        reply_markup=keyboard,
        parse_mode="HTML"
    )

# =========================================================
# /menu
# =========================================================

@dp.message(Command("menu"))
async def menu_handler(message: Message):

    update_user(message.from_user)

    await message.answer(
        "🥷 <b>SHINOBI TEAM BOT</b>\n\n"
        "Выбери нужное действие:",
        reply_markup=main_keyboard(message.from_user.id),
        parse_mode="HTML"
    )


# =========================================================
# СОЗДАНИЕ РЕФЕРАЛЬНОЙ ССЫЛКИ
# =========================================================

async def _send_created_link(tg_user, target, creates_join_request: bool = True):
    user = update_user(tg_user)
    link_number = len(user["links"]) + 1

    try:
        invite = await bot.create_chat_invite_link(
            chat_id=CHAT_ID,
            name=f"REF_{tg_user.id}_{link_number}",
            creates_join_request=True
        )
    except Exception as error:
        text = (
            "❌ <b>Не удалось создать ссылку.</b>\n\n"
            "Проверь, что бот является администратором группы и имеет право создавать "
            "пригласительные ссылки.\n\n"
            f"<code>{error}</code>"
        )
        if isinstance(target, Message):
            await target.answer(text, parse_mode="HTML")
        else:
            await target.message.answer(text, parse_mode="HTML")
        return

    link = invite.invite_link
    data["links"][link] = {
        "owner_id": tg_user.id,
        "created": datetime.now().isoformat(),
        "invited": 0,
        "users": [],
        "join_request": True
    }
    user["links"].append(link)
    save_data()

    mode = "📩 По заявке"
    text = (
        "🔗 <b>НОВАЯ РЕФЕРАЛЬНАЯ ССЫЛКА</b>\n\n"
        f"{link}\n\n"
        f"👤 Создатель: {tg_user.full_name}\n"
        f"🆔 ID: <code>{tg_user.id}</code>\n"
        f"📎 Ссылка №{link_number}\n"
        f"🚪 Режим: <b>{mode}</b>\n"
        "👥 Приглашено: <b>0</b>\n\n"
        + ("Заявки на вход будут видны администраторам бота и в Admin Staff."
           if creates_join_request else "Пользователь сможет войти сразу без подтверждения.")
    )
    if isinstance(target, Message):
        await target.answer(text, parse_mode="HTML", disable_web_page_preview=True)
    else:
        await target.message.answer(text, parse_mode="HTML", disable_web_page_preview=True)


async def create_ref_link(tg_user, target):
    # Все персональные реферальные ссылки автоматически создаются
    # только в режиме заявки на вступление.
    await _send_created_link(tg_user, target, True)


@dp.callback_query(F.data.in_({"create_link_direct", "create_link_request"}))
async def create_link_mode_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return
    await callback.answer("Создаю ссылку...")
    await _send_created_link(callback.from_user, callback, True)


# =========================================================
# /link
# =========================================================

@dp.message(F.text == "📋 Мой график")
async def quick_my_schedule(message: Message):
    # Переиспользуем текстовую команду, которую обрабатывает модуль графика
    from schedule_system import personal_week_schedule
    await personal_week_schedule(message)

@dp.message(F.text == "🏖 Запросить выходной")
async def quick_dayoff(message: Message):
    if message.chat.type != "private":
        info=await message.bot.get_me()
        await message.answer("🏖 Запрос выходного оформляется в ЛС с ботом.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Открыть бота",url=f"https://t.me/{info.username}?start=schedule")]]))
        return
    await message.answer("🏖 Открой меню графика и нажми «Запросить выходной».", reply_markup=main_keyboard(message.from_user.id))

@dp.message(Command("menu"))
async def quick_menu(message: Message):
    await message.answer("⚡ Быстрые действия:", reply_markup=quick_reply_keyboard(message.from_user.id))

@dp.message(Command("link"))
async def link_command(message: Message):

    await create_ref_link(
        message.from_user,
        message
    )


# =========================================================
# КНОПКА СОЗДАТЬ ССЫЛКУ
# =========================================================

@dp.callback_query(F.data == "create_link")
async def link_callback(callback: CallbackQuery):

    await callback.answer(
        "Создаю ссылку..."
    )

    await create_ref_link(
        callback.from_user,
        callback
    )


# =========================================================
# ЕДИНАЯ ССЫЛКА ДЛЯ РЕКЛАМЫ
# =========================================================

async def _rotate_advertising_link():
    """Create a new link without revoking unrelated staff links; keep old stats."""
    global ADVERTISING_LINK
    invite = await bot.create_chat_invite_link(chat_id=CHAT_ID, name="Shinobi advertising", creates_join_request=True)
    new_link = invite.invite_link
    old_link = data.get("advertising_link") or ADVERTISING_LINK
    old_info = data.setdefault("links", {}).get(old_link, {})
    data["links"][new_link] = {
        "owner_id": None, "type": "advertising", "label": "Рекламная ссылка",
        "created": datetime.now().isoformat(), "invited": 0, "users": [],
        "join_request": True, "previous_link": old_link,
    }
    # Do not revoke or erase the old link: its historical counts remain intact.
    data["advertising_link"] = new_link
    ADVERTISING_LINK = new_link
    save_data()
    return new_link, old_link, int(old_info.get("invited", 0) or 0)

@dp.callback_query(F.data == "ad_link_rotate")
async def advertising_link_rotate(callback: CallbackQuery):
    if callback.from_user.id not in OWNER_IDS:
        await callback.answer("Только владельцу", show_alert=True)
        return
    await callback.answer("Создаю новую ссылку...")
    try:
        new_link, old_link, old_count = await _rotate_advertising_link()
    except Exception as exc:
        await callback.message.answer("❌ Не удалось создать ссылку. Проверь право бота приглашать пользователей.\n" + html.escape(str(exc)[:500]), parse_mode="HTML")
        return
    await callback.message.answer(
        "✅ <b>Новая рекламная ссылка создана</b>\n" + html.escape(new_link) +
        "\n\nСтарая ссылка и её статистика сохранены (" + str(old_count) + " приглашений).",
        parse_mode="HTML", disable_web_page_preview=True)

@dp.callback_query(F.data == "advertising_link")
async def advertising_link_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return

    info = data.setdefault("links", {}).setdefault(ADVERTISING_LINK, {
        "owner_id": None, "type": "advertising", "label": "Рекламная ссылка",
        "created": datetime.now().isoformat(), "invited": 0, "users": [], "join_request": True
    })
    data["advertising_link"] = ADVERTISING_LINK
    save_data()
    invited = max(0, int(info.get("invited", 0) or 0))
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Добавить", callback_data="ad_inv_add"),
         InlineKeyboardButton(text="➖ Отнять", callback_data="ad_inv_sub")],
        [InlineKeyboardButton(text="✏️ Установить число", callback_data="ad_inv_set")],
        [InlineKeyboardButton(text="🔄 Создать новую рабочую ссылку", callback_data="ad_link_rotate")],
        [InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")]
    ])
    await callback.answer()
    await callback.message.answer(
        "📣 <b>РЕКЛАМНАЯ ССЫЛКА</b>\n\n"
        f"🔗 {ADVERTISING_LINK}\n"
        f"👥 Приглашено: <b>{invited}</b>\n\n"
        "Эта ссылка учитывается отдельно от персональных ссылок сотрудников.",
        parse_mode="HTML", disable_web_page_preview=True, reply_markup=kb
    )


def invite_source_text(invite_url: str) -> str:
    if not invite_url:
        return "Ссылка не относится к боту"
    info = data.get("links", {}).get(invite_url)
    if not info:
        return "Ссылка не относится к боту"
    if info.get("type") == "advertising":
        return "Ссылка для рекламы"
    owner_id = info.get("owner_id")
    owner = get_user(owner_id) if owner_id else None
    return owner.get("name", "Ссылка не относится к боту") if owner else "Ссылка не относится к боту"


# =========================================================
# ЗАЯВКИ НА ВСТУПЛЕНИЕ
# =========================================================

def _join_request_key(chat_id: int, user_id: int) -> str:
    return f"{chat_id}:{user_id}"


# Лок на одну заявку не даёт двум почти одновременным callback обработать её дважды.
_join_request_locks = {}


def _join_request_keyboard(chat_id: int, user_id: int):
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Принять", callback_data=f"jr_accept:{chat_id}:{user_id}"),
        InlineKeyboardButton(text="❌ Отклонить", callback_data=f"jr_decline:{chat_id}:{user_id}")
    ]])


def _join_request_text(req: dict) -> str:
    username = f"@{html.escape(str(req['username']))}" if req.get("username") else "—"
    source = req.get("link_source") or invite_source_text(req.get("invite_link", ""))
    return (
        "📩 <b>НОВАЯ ЗАЯВКА НА ВСТУПЛЕНИЕ</b>\n\n"
        f"👤 {html.escape(str(req.get('name', 'Без имени')))}\n"
        f"🔹 {username}\n"
        f"🆔 <code>{req.get('user_id')}</code>\n"
        f"🏘 Группа: <b>{html.escape(str(req.get('chat_title', 'Shinobi Team')))}</b>\n"
        f"🔗 Источник: <b>{html.escape(str(source))}</b>\n\n"
        "Решение синхронизируется для всех администраторов."
    )


def _join_request_result_text(req: dict) -> str:
    status = req.get("status")
    result = "✅ ПРИНЯТА" if status == "approved" else "❌ ОТКЛОНЕНА"
    who = req.get("decided_by_name") or str(req.get("decided_by") or "Администратор")
    where = req.get("decided_in") or "Telegram"
    return _join_request_text(req) + f"\n\n<b>{result}</b>\nРешение: {html.escape(str(who))}\nГде принято: {html.escape(str(where))}"


async def _sync_join_request_messages(req: dict):
    """Редактирует ВСЕ копии карточки: ЛС админов и Staff."""
    text = _join_request_result_text(req)
    for item in list(req.get("messages") or []):
        try:
            await bot.edit_message_text(
                chat_id=int(item["chat_id"]), message_id=int(item["message_id"]),
                text=text, parse_mode="HTML", reply_markup=None
            )
        except Exception:
            # Сообщение могло быть удалено вручную — это не мешает синхронизации остальных.
            pass


async def _auto_approve_join_request(chat_id: int, user_id: int, delay: int = 10 * 60):
    """Автоматически принимает заявку, если администрация не решила её за delay секунд."""
    await asyncio.sleep(delay)
    key = _join_request_key(chat_id, user_id)
    lock = _join_request_locks.setdefault(key, asyncio.Lock())
    async with lock:
        req = data.setdefault("pending_join_requests", {}).get(key)
        if not req or req.get("status") != "pending":
            return
        try:
            await bot.approve_chat_join_request(chat_id, user_id)
        except Exception as exc:
            req["auto_approve_error"] = str(exc)
            req["auto_approve_error_at"] = datetime.now().isoformat()
            save_data()
            return
        req["status"] = "approved"
        req["decided_by"] = 0
        req["decided_by_name"] = "🤖 Бот (автоприём через 10 минут)"
        req["decided_at"] = datetime.now().isoformat()
        req["decided_in"] = "Автоприём"
        req["auto_approved"] = True
        save_data()
        await _sync_join_request_messages(req)


@dp.chat_join_request()
async def join_request_handler(event: ChatJoinRequest):
    logging.getLogger("shinobi.join").warning(
        "[JOIN REQUEST RECEIVED] chat=%s user=%s link=%s",
        event.chat.id, event.from_user.id,
        (event.invite_link.invite_link if event.invite_link else "none")
    )
    if event.chat.id != CHAT_ID:
        logging.getLogger("shinobi.join").warning(
            "[JOIN REQUEST IGNORED] chat=%s expected=%s", event.chat.id, CHAT_ID
        )
        return
    key = _join_request_key(event.chat.id, event.from_user.id)
    existing = data.setdefault("pending_join_requests", {}).get(key)
    # Telegram/несколько update-путей могут прислать одно событие повторно.
    # Пока заявка активна, второй комплект карточек НЕ создаём.
    if existing and existing.get("status") == "pending" and any(m.get("place") == "staff" for m in (existing.get("messages") or [])):
        logging.getLogger("shinobi.join").warning("[JOIN REQUEST DUPLICATE] user=%s already delivered", event.from_user.id)
        return
    # Если старая заявка была сохранена, но уведомление не дошло — отправляем повторно.

    invite_url = event.invite_link.invite_link if event.invite_link else ""
    link_info = data.get("links", {}).get(invite_url, {})
    req = {
        "chat_id": event.chat.id,
        "chat_title": event.chat.title or "Shinobi Team",
        "user_id": event.from_user.id,
        "name": event.from_user.full_name,
        "username": event.from_user.username or "",
        "invite_link": invite_url,
        "link_owner_id": link_info.get("owner_id"),
        "link_source": invite_source_text(invite_url),
        "created": datetime.now().isoformat(),
        "status": "pending",
        "messages": []
    }
    # Проверка ЧС/бана перед обработкой заявки. Telegram не хранит публичную
    # текстовую причину бана, поэтому фиксируем доступную техническую причину.
    try:
        member = await bot.get_chat_member(event.chat.id, event.from_user.id)
        if getattr(member, "status", "") == "kicked":
            req["status"] = "declined"
            req["entry_block_reason"] = "Пользователь находится в ЧС (заблокирован в чате)."
            req["decided_by_name"] = "🤖 Бот — проверка входа"
            req["decided_at"] = datetime.now().isoformat()
            req["decided_in"] = "Проверка ЧС"
            data["pending_join_requests"][key] = req
            save_data()
            try:
                await bot.decline_chat_join_request(event.chat.id, event.from_user.id)
            except Exception:
                pass
            return
    except Exception as exc:
        req["entry_check_note"] = f"Не удалось проверить статус до решения: {exc}"

    data["pending_join_requests"][key] = req
    save_data()
    # Независимый таймер: через 10 минут примет только всё ещё pending-заявку.
    asyncio.create_task(_auto_approve_join_request(event.chat.id, event.from_user.id))
    text = _join_request_text(req)
    kb = _join_request_keyboard(event.chat.id, event.from_user.id)

    # В первую очередь Staff: не позволяем ошибкам личных сообщений мешать заявкам.
    async def _deliver(chat_id, place):
        try:
            msg = await bot.send_message(chat_id, text, reply_markup=kb, parse_mode="HTML")
            req["messages"].append({"chat_id": chat_id, "message_id": msg.message_id, "place": place})
            logging.getLogger("shinobi.join").warning("[JOIN REQUEST DELIVERED] place=%s chat=%s user=%s", place, chat_id, event.from_user.id)
            return
        except Exception as exc:
            import logging
            logging.getLogger("shinobi.delivery").exception("Ошибка отправки заявки в %s (%s)", place, chat_id)
            req.setdefault("delivery_errors", []).append(f"{place} {chat_id}: {str(exc)[:160]}")
        # Резервный вариант: убираем HTML, но сохраняем кнопки управления.
        try:
            import re
            plain = re.sub(r"<[^>]*>", "", text)
            msg = await bot.send_message(chat_id, html.unescape(plain), reply_markup=kb)
            req["messages"].append({"chat_id": chat_id, "message_id": msg.message_id, "place": place})
        except Exception as exc:
            import logging
            logging.getLogger("shinobi.delivery").exception("Не удалось доставить даже обычный текст заявки в %s (%s)", place, chat_id)
            req.setdefault("delivery_errors", []).append(f"fallback {place} {chat_id}: {str(exc)[:160]}")

    await _deliver(ADMIN_STAFF_CHAT_ID, "staff")
    for admin_id in sorted(set(ADMIN_IDS)):
        await _deliver(admin_id, "dm")
    save_data()


@dp.callback_query(F.data.startswith("jr_accept:") | F.data.startswith("jr_decline:"))
async def join_request_decision(callback: CallbackQuery):
    can_decide = is_admin(callback.from_user.id) or await _is_admin_staff_member(callback.from_user.id)
    if not can_decide:
        await callback.answer("⛔ Принимать заявки могут только участники Admin Staff и администрация.", show_alert=True)
        return
    action, chat_id_s, user_id_s = callback.data.split(":", 2)
    chat_id, user_id = int(chat_id_s), int(user_id_s)
    key = _join_request_key(chat_id, user_id)
    lock = _join_request_locks.setdefault(key, asyncio.Lock())
    async with lock:
        req = data.setdefault("pending_join_requests", {}).get(key)
        if not req or req.get("status") != "pending":
            await callback.answer("Эта заявка уже обработана.", show_alert=True)
            # На старой/лишней карточке тоже убираем кнопки.
            try:
                await callback.message.edit_reply_markup(reply_markup=None)
            except Exception:
                pass
            return
        try:
            if action == "jr_accept":
                await bot.approve_chat_join_request(chat_id, user_id)
                req["status"] = "approved"
            else:
                await bot.decline_chat_join_request(chat_id, user_id)
                req["status"] = "declined"
        except Exception as error:
            await callback.answer(f"Ошибка Telegram: {error}", show_alert=True)
            return
        req["decided_by"] = callback.from_user.id
        req["decided_by_name"] = callback.from_user.full_name
        req["decided_at"] = datetime.now().isoformat()
        req["decided_in"] = "Admin Staff" if callback.message and callback.message.chat.id == ADMIN_STAFF_CHAT_ID else "ЛС"
        save_data()
        await callback.answer("Готово")
        # Важное: меняем не только нажатое сообщение, а все сохранённые копии.
        await _sync_join_request_messages(req)


@dp.callback_query(F.data == "admin_join_requests")
async def admin_join_requests_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return
    pending = [r for r in data.get("pending_join_requests", {}).values() if r.get("status") == "pending"]
    if not pending:
        await callback.message.edit_text(
            "📩 <b>ЗАЯВКИ НА ВСТУПЛЕНИЕ</b>\n\nСейчас активных заявок нет.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")]]),
            parse_mode="HTML"
        )
        await callback.answer()
        return
    rows=[]
    for req in pending[-30:][::-1]:
        rows.append([InlineKeyboardButton(
            text=f"📩 {req.get('name','Без имени')}"[:55],
            callback_data=f"admin_jr:{req['chat_id']}:{req['user_id']}"
        )])
    rows.append([InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")])
    await callback.message.edit_text(
        f"📩 <b>ЗАЯВКИ НА ВСТУПЛЕНИЕ</b>\n\nОжидают решения: <b>{len(pending)}</b>",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("admin_jr:"))
async def admin_join_request_open(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return
    _, chat_id_s, user_id_s = callback.data.split(":", 2)
    req = data.get("pending_join_requests", {}).get(_join_request_key(int(chat_id_s), int(user_id_s)))
    if not req or req.get("status") != "pending":
        await callback.answer("Заявка уже обработана.", show_alert=True)
        return
    await callback.message.edit_text(_join_request_text(req), reply_markup=_join_request_keyboard(int(chat_id_s), int(user_id_s)), parse_mode="HTML")
    await callback.answer()


# =========================================================
# ВЫДАЧА СЕРВЕРНОГО ТЕГА ПО ОТВЕТУ НОВОГО УЧАСТНИКА
# =========================================================

@dp.message(F.reply_to_message)
async def server_tag_reply_handler(message: Message):
    if not _server_tags_enabled():
        return
    if message.chat.id != CHAT_ID or not message.from_user or message.from_user.is_bot:
        return
    key = f"{message.chat.id}:{message.from_user.id}"
    pending = _server_pending_root().get(key)
    if not pending:
        return
    if not message.reply_to_message or message.reply_to_message.message_id != int(pending.get("prompt_message_id", 0)):
        return
    if int(pending.get("expires_at", 0)) < int(time.time()):
        _server_pending_root().pop(key, None)
        save_data()
        return

    tag = _parse_server_tag(message.text or "")
    canonical_input, was_corrected = _canonical_server_input(message.text or "")
    if not tag:
        err = await message.reply(
            "❌ Не узнал такой сервер. Напиши, например: <b>Мск</b>, <b>Ру</b>, <b>Ру/Еу</b> или комбинацию из списка серверов.",
            parse_mode="HTML",
        )
        asyncio.create_task(_delete_message_later(message.chat.id, err.message_id, 30))
        return

    # Если сервер распознан, но написан неканонично (например, "мск", "москва", "мскк"),
    # отвечаем исправленным вариантом и продолжаем выдачу тега.
    if canonical_input and (was_corrected or (message.text or "").strip() != canonical_input):
        fixed = await message.reply(f"✏️ Исправлено: <b>{canonical_input}</b>", parse_mode="HTML")
        asyncio.create_task(_delete_message_later(message.chat.id, fixed.message_id, 30))

    try:
        assigned = await _telegram_set_member_tag(message.chat.id, message.from_user.id, tag)
        if not assigned:
            return
    except Exception as exc:
        text = str(exc)
        if "manage" in text.casefold() or "right" in text.casefold() or "administrator" in text.casefold():
            text = "у бота нет права Telegram «Управление тегами / can_manage_tags»"
        err = await message.reply(f"❌ Не удалось выдать тег: {text}")
        asyncio.create_task(_delete_message_later(message.chat.id, err.message_id, 60))
        return

    _server_pending_root().pop(key, None)
    _server_assignments_root()[f"{message.chat.id}:{message.from_user.id}"] = {
        "chat_id": message.chat.id,
        "user_id": message.from_user.id,
        "name": message.from_user.full_name,
        "username": message.from_user.username or "",
        "tag": tag,
        "assigned_at": datetime.now().isoformat(),
    }
    save_data()
    mention = f'<a href="tg://user?id={message.from_user.id}">{message.from_user.full_name}</a>'
    confirm = await message.answer(f"✅ {mention}, твой серверный тег выдан: <b>{tag}</b>", parse_mode="HTML")
    # Вопрос и ответ убираем из чата; подтверждение с тегом живёт 20 минут.
    asyncio.create_task(_delete_message_later(message.chat.id, int(pending["prompt_message_id"]), 5))
    asyncio.create_task(_delete_message_later(message.chat.id, message.message_id, 5))
    asyncio.create_task(_delete_message_later(message.chat.id, confirm.message_id, SERVER_TAG_CONFIRM_DELETE_SECONDS))


# =========================================================
# СПАМ-БЛОК: РЕАКЦИЯ НА ЗАПРОС СЕРВЕРНОГО ТЕГА
# =========================================================

@dp.message_reaction()
async def server_tag_spamblock_reaction(event: MessageReactionUpdated):
    if not _server_tags_enabled():
        return
    if event.chat.id != CHAT_ID or not event.user or event.user.is_bot:
        return
    key = f"{event.chat.id}:{event.user.id}"
    pending = _server_pending_root().get(key)
    if not pending:
        return
    # Реакция должна быть именно на персональный запрос этого пользователя.
    if event.message_id != int(pending.get("prompt_message_id", 0)):
        return
    if int(pending.get("expires_at", 0)) < int(time.time()):
        return
    # Снятие реакции не считается подтверждением.
    if not event.new_reaction:
        return

    # Сразу закрываем ожидание, чтобы повторные реакции не создавали дубли в Staff.
    _server_pending_root().pop(key, None)
    save_data()

    mention = f'<a href="tg://user?id={event.user.id}">{event.user.full_name}</a>'
    username = f"@{event.user.username}" if event.user.username else "нет username"
    try:
        await bot.send_message(
            ADMIN_STAFF_CHAT_ID,
            "🚫 <b>СПАМ-БЛОК — НУЖНА РУЧНАЯ ВЫДАЧА ТЕГА</b>\n\n"
            f"👤 {mention}\n"
            f"🔹 {username}\n"
            f"🆔 <code>{event.user.id}</code>\n\n"
            "Пользователь отреагировал на персональный запрос серверного тега и, вероятно, не может писать в группе.\n"
            "📩 Администратору нужно написать ему в ЛС и уточнить сервер для ручной выдачи тега.",
            parse_mode="HTML",
        )
    except Exception as exc:
        print(f"[SERVER TAG][SPAM BLOCK] Staff notify error: {exc}")

    asyncio.create_task(_delete_message_later(event.chat.id, int(pending["prompt_message_id"]), 5))


# =========================================================
# ОТСЛЕЖИВАНИЕ ВСТУПЛЕНИЙ
# =========================================================

@dp.chat_member()
async def member_join_handler(event: ChatMemberUpdated):

    # Только нужная группа
    if event.chat.id != CHAT_ID:
        return

    old_status = event.old_chat_member.status
    new_status = event.new_chat_member.status

    joined = (
        old_status in {"left", "kicked"}
        and new_status in {
            "member",
            "administrator",
            "restricted"
        }
    )

    if not joined:
        return

    joined_user = event.new_chat_member.user

    # История всех реальных входов в основной чат для веб-аналитики.
    # Сохраняем даже входы по обычным/неизвестным ссылкам; ботов не учитываем.
    if not joined_user.is_bot:
        used_join_link = event.invite_link.invite_link if event.invite_link else ""
        data.setdefault("join_history", []).append({
            "user_id": joined_user.id,
            "name": joined_user.full_name,
            "username": joined_user.username or "",
            "joined_at": datetime.now().isoformat(),
            "invite_link": used_join_link,
            "source": invite_source_text(used_join_link) if used_join_link else "Без отслеживаемой ссылки",
        })
        # Ограничение защищает data.json от бесконечного роста, сохраняя длинную историю.
        data["join_history"] = data["join_history"][-20000:]
        save_data()

    # Ботов не спрашиваем. Обычному участнику сразу предлагаем выбрать сервер.
    if not joined_user.is_bot and _server_tags_enabled():
        try:
            asyncio.create_task(_ask_server_tag_delayed(event.chat.id, joined_user, 15))
        except Exception as exc:
            print(f"[SERVER TAG] Не удалось отправить вопрос: {exc}")

    # Дальше идёт независимая реферальная логика. Если ссылки нет, тег всё равно уже запрошен.
    if not event.invite_link:
        return

    used_link = event.invite_link.invite_link

    # Ботов не считаем
    if joined_user.is_bot:
        return

    # Чужая/ручная ссылка не относится к реферальной системе бота.
    if used_link not in data.get("links", {}):
        return

    link_info = data["links"][used_link]

    # Рекламная ссылка общая и не принадлежит конкретному пользователю.
    if link_info.get("type") == "advertising":
        link_info.setdefault("invited", 0)
        link_info.setdefault("users", [])
        if any(x.get("id") == joined_user.id for x in link_info["users"]):
            return
        link_info["invited"] += 1
        link_info["users"].append({
            "id": joined_user.id,
            "name": joined_user.full_name,
            "username": joined_user.username or "",
            "joined": datetime.now().isoformat(),
            "source": "Ссылка для рекламы"
        })
        save_data()
        return

    owner_id = str(link_info.get("owner_id"))
    if owner_id not in data["users"]:
        return
    owner = data["users"][owner_id]
    owner.setdefault("invited_users", [])
    owner.setdefault("total_invited", 0)

    link_info.setdefault(
        "invited",
        0
    )

    link_info.setdefault(
        "users",
        []
    )

    joined_id = joined_user.id

    # =====================================================
    # АНТИНАКРУТКА
    # =====================================================
    #
    # Если человек уже был засчитан этому владельцу,
    # повторно его не считаем.
    #
    # =====================================================

    if joined_id in owner["invited_users"]:
        return

    # Записываем приглашённого
    owner["invited_users"].append(
        joined_id
    )

    owner["total_invited"] += 1

    link_info["invited"] += 1

    link_info["users"].append({
        "id": joined_user.id,
        "name": joined_user.full_name,
        "username": joined_user.username or "",
        "joined": datetime.now().isoformat()
    })

    save_data()

    # =====================================================
    # УВЕДОМЛЕНИЕ ВЛАДЕЛЬЦУ ССЫЛКИ
    # =====================================================

    try:

        username_text = ""

        if joined_user.username:

            username_text = (
                f"\n🔹 @{joined_user.username}"
            )

        await bot.send_message(
            int(owner_id),
            "🎉 <b>НОВЫЙ ПРИГЛАШЁННЫЙ!</b>\n\n"
            f"👤 {joined_user.full_name}"
            f"{username_text}\n"
            f"🆔 <code>{joined_user.id}</code>\n\n"
            f"📊 Всего приглашено: "
            f"<b>{owner['total_invited']}</b>",
            parse_mode="HTML"
        )

    except Exception:
        pass


# =========================================================
# МОЯ СТАТИСТИКА
# =========================================================

async def send_stats(user_id: int, target):

    user = get_user(user_id)

    links = user.get(
        "links",
        []
    )

    text = (
        "📊 <b>ТВОЯ СТАТИСТИКА</b>\n\n"
        f"🆔 ID: <code>{user_id}</code>\n"
        f"👥 Всего приглашено: "
        f"<b>{user.get('total_invited', 0)}</b>\n"
        f"🔗 Создано ссылок: "
        f"<b>{len(links)}</b>\n"
    )

    if links:

        text += (
            "\n<b>Статистика ссылок:</b>\n\n"
        )

        for number, link in enumerate(
            links,
            start=1
        ):

            info = data["links"].get(
                link
            )

            if not info:
                continue

            text += (
                f"🔗 Ссылка #{number}\n"
                f"👥 Приглашено: "
                f"<b>{info.get('invited', 0)}</b>\n"
                f"{link}\n\n"
            )

            if len(text) > 3500:

                if isinstance(target, Message):

                    await target.answer(
                        text,
                        parse_mode="HTML",
                        disable_web_page_preview=True
                    )

                else:

                    await target.message.answer(
                        text,
                        parse_mode="HTML",
                        disable_web_page_preview=True
                    )

                text = ""

    if text:

        if isinstance(target, Message):

            await target.answer(
                text,
                parse_mode="HTML",
                disable_web_page_preview=True
            )

        else:

            await target.message.answer(
                text,
                parse_mode="HTML",
                disable_web_page_preview=True
            )


# =========================================================
# /stats
# =========================================================

@dp.message(Command("stats"))
async def stats_command(message: Message):

    await send_stats(
        message.from_user.id,
        message
    )


# =========================================================
# КНОПКА СТАТИСТИКА
# =========================================================

@dp.callback_query(F.data == "my_stats")
async def stats_callback(callback: CallbackQuery):

    await callback.answer()

    await send_stats(
        callback.from_user.id,
        callback
    )


# =========================================================
# МОИ ССЫЛКИ
# =========================================================

async def send_links(user_id: int, target):

    user = get_user(user_id)

    links = user.get(
        "links",
        []
    )

    if not links:

        text = (
            "❌ У тебя пока нет реферальных ссылок.\n\n"
            "Нажми «🔗 Создать ссылку»."
        )

    else:

        text = (
            "🔗 <b>ТВОИ ССЫЛКИ</b>\n\n"
        )

        for number, link in enumerate(
            links,
            start=1
        ):

            info = data["links"].get(
                link,
                {}
            )

            text += (
                f"<b>#{number}</b>\n"
                f"{link}\n"
                f"👥 Приглашено: "
                f"<b>{info.get('invited', 0)}</b>\n\n"
            )

    if isinstance(target, Message):

        await target.answer(
            text,
            parse_mode="HTML",
            disable_web_page_preview=True
        )

    else:

        await target.message.answer(
            text,
            parse_mode="HTML",
            disable_web_page_preview=True
        )


# =========================================================
# /links
# =========================================================

@dp.message(Command("links"))
async def links_command(message: Message):

    await send_links(
        message.from_user.id,
        message
    )


# =========================================================
# КНОПКА МОИ ССЫЛКИ
# =========================================================

@dp.callback_query(F.data == "my_links")
async def links_callback(callback: CallbackQuery):

    await callback.answer()

    await send_links(
        callback.from_user.id,
        callback
    )


# =========================================================
# ПРИГЛАШЁННЫЕ
# =========================================================

async def send_refs(user_id: int, target):

    user = get_user(user_id)

    refs = []

    for link in user.get(
        "links",
        []
    ):

        link_info = data["links"].get(
            link
        )

        if not link_info:
            continue

        for ref in link_info.get(
            "users",
            []
        ):

            refs.append(ref)

    if not refs:

        text = (
            "👥 Ты пока никого не пригласил."
        )

        if isinstance(target, Message):

            await target.answer(
                text
            )

        else:

            await target.message.answer(
                text
            )

        return

    text = (
        "👥 <b>ТВОИ ПРИГЛАШЁННЫЕ</b>\n\n"
    )

    for number, ref in enumerate(
        refs,
        start=1
    ):

        username = ref.get(
            "username",
            ""
        )

        if username:

            name = f"@{username}"

        else:

            name = ref.get(
                "name",
                "Пользователь"
            )

        line = (
            f"{number}. {name}\n"
            f"🆔 <code>{ref.get('id')}</code>\n\n"
        )

        if len(text) + len(line) > 3500:

            if isinstance(target, Message):

                await target.answer(
                    text,
                    parse_mode="HTML"
                )

            else:

                await target.message.answer(
                    text,
                    parse_mode="HTML"
                )

            text = ""

        text += line

    if text:

        if isinstance(target, Message):

            await target.answer(
                text,
                parse_mode="HTML"
            )

        else:

            await target.message.answer(
                text,
                parse_mode="HTML"
            )


# =========================================================
# /refs
# =========================================================

@dp.message(Command("refs"))
async def refs_command(message: Message):

    await send_refs(
        message.from_user.id,
        message
    )


# =========================================================
# КНОПКА ПРИГЛАШЁННЫЕ
# =========================================================

@dp.callback_query(F.data == "my_refs")
async def refs_callback(callback: CallbackQuery):

    await callback.answer()

    await send_refs(
        callback.from_user.id,
        callback
    )


# =========================================================
# ТОП
# =========================================================

async def send_top(target):

    users = list(
        data["users"].items()
    )

    users.sort(
        key=lambda item: item[1].get(
            "total_invited",
            0
        ),
        reverse=True
    )

    text = (
        "🏆 <b>ТОП SHINOBI TEAM</b>\n\n"
    )

    medals = [
        "🥇",
        "🥈",
        "🥉"
    ]

    position = 0

    for uid, user in users:

        count = user.get(
            "total_invited",
            0
        )

        if count <= 0:
            continue

        position += 1

        if position <= 3:

            icon = medals[
                position - 1
            ]

        else:

            icon = f"{position}."

        text += (
            f"{icon} "
            f"{display_user(user)} — "
            f"<b>{count}</b>\n"
        )

        if position >= 20:
            break

    if position == 0:

        text += (
            "Пока приглашений нет."
        )

    if isinstance(target, Message):

        await target.answer(
            text,
            parse_mode="HTML"
        )

    else:

        await target.message.answer(
            text,
            parse_mode="HTML"
        )


# =========================================================
# /top
# =========================================================

@dp.message(Command("top"))
async def top_command(message: Message):

    await send_top(message)


# =========================================================
# КНОПКА ТОП
# =========================================================

@dp.callback_query(F.data == "top")
async def top_callback(callback: CallbackQuery):

    await callback.answer()

    await send_top(callback)


# =========================================================
# СОЗДАНИЕ EXCEL
# =========================================================

def create_excel():
    """Полный постоянный Excel-снимок реферальной базы Shinobi Team."""
    workbook = Workbook()

    # 1. Пользователи
    users_sheet = workbook.active
    users_sheet.title = "Пользователи"
    users_sheet.append(["ID", "Username", "Имя", "Всего приглашено", "Количество ссылок", "Ссылки"])
    for uid, user in sorted(data.get("users", {}).items(), key=lambda x: (x[1].get("name") or "").casefold()):
        links = user.get("links", [])
        users_sheet.append([
            int(user.get("id") or uid),
            ("@" + user.get("username", "")) if user.get("username") else "",
            user.get("name", ""),
            int(user.get("total_invited", 0) or 0),
            len(links),
            "\n".join(links),
        ])

    # 2. Все реферальные ссылки
    links_sheet = workbook.create_sheet("Реферальные ссылки")
    links_sheet.append([
        "№", "Ссылка", "Тип", "ID создателя", "Username создателя", "Имя создателя",
        "Приглашено по ссылке", "Дата создания", "Вход по заявке"
    ])
    for number, (link, info) in enumerate(data.get("links", {}).items(), 1):
        owner_id = info.get("owner_id")
        owner = data.get("users", {}).get(str(owner_id), {}) if owner_id is not None else {}
        links_sheet.append([
            number, link, info.get("type", "referral"), owner_id if owner_id is not None else "",
            ("@" + owner.get("username", "")) if owner.get("username") else "",
            owner.get("name", info.get("label", "")), int(info.get("invited", 0) or 0),
            info.get("created", ""), "Да" if info.get("join_request") else "Нет"
        ])

    # 3. Реально вступившие люди
    refs_sheet = workbook.create_sheet("Приглашённые")
    refs_sheet.append([
        "№", "Ссылка", "ID создателя", "Username создателя", "Имя создателя",
        "ID приглашённого", "Username приглашённого", "Имя приглашённого", "Дата вступления", "Источник"
    ])
    ref_number = 1
    for link, info in data.get("links", {}).items():
        owner_id = info.get("owner_id")
        owner = data.get("users", {}).get(str(owner_id), {}) if owner_id is not None else {}
        for ref in info.get("users", []):
            refs_sheet.append([
                ref_number, link, owner_id if owner_id is not None else "",
                ("@" + owner.get("username", "")) if owner.get("username") else "",
                owner.get("name", info.get("label", "")), ref.get("id", ""),
                ("@" + ref.get("username", "")) if ref.get("username") else "",
                ref.get("name", ""), ref.get("joined", ""), ref.get("source", "")
            ])
            ref_number += 1

    # 4. История ручных корректировок администратором
    adj_sheet = workbook.create_sheet("Ручные изменения")
    adj_sheet.append([
        "Дата", "Администратор ID", "Администратор", "Пользователь ID", "Username пользователя",
        "Имя пользователя", "Ссылка", "Действие", "Значение", "Было", "Стало"
    ])
    for item in data.get("invite_adjustments", []):
        adj_sheet.append([
            item.get("created", ""), item.get("admin_id", ""), item.get("admin_name", ""),
            item.get("user_id", ""), item.get("username", ""), item.get("name", ""),
            item.get("link", ""), item.get("action", ""), item.get("value", ""),
            item.get("old", ""), item.get("new", "")
        ])

    # 5. Системный снимок — помогает понять состояние резервной копии.
    sys_sheet = workbook.create_sheet("Система")
    sys_sheet.append(["Параметр", "Значение"])
    sys_sheet.append(["Обновлено", datetime.now().isoformat(timespec="seconds")])
    sys_sheet.append(["Пользователей", len(data.get("users", {}))])
    sys_sheet.append(["Ссылок", len(data.get("links", {}))])
    sys_sheet.append(["Ручных изменений", len(data.get("invite_adjustments", []))])
    sys_sheet.append(["Рекламная ссылка", data.get("advertising_link", "")])

    # 6. Скрытая полная копия data.json. Она нужна для точного восстановления
    # пользователей, ссылок, статистики, заявок и всех служебных полей.
    backup_sheet = workbook.create_sheet("Backup JSON")
    backup_sheet.append(["JSON backup chunks"])
    backup_payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    chunk_size = 30000
    for pos in range(0, len(backup_payload), chunk_size):
        backup_sheet.append([backup_payload[pos:pos + chunk_size]])
    backup_sheet.sheet_state = "hidden"

    # Оформление всех листов.
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col in sheet.columns:
            letter = col[0].column_letter
            max_len = max((len(str(c.value or "")) for c in col[:100]), default=10)
            sheet.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)

    workbook.save(EXCEL_FILE)
    return EXCEL_FILE


# =========================================================
# ПОСТОЯННЫЕ НИЖНИЕ КНОПКИ
# =========================================================

@dp.message(F.text == "🔗 Создать ссылку")
async def reply_create_link(message: Message):
    await link_command(message)

@dp.message(F.text == "📊 Моя статистика")
async def reply_my_stats(message: Message):
    await stats_command(message)

@dp.message(F.text == "🔗 Мои ссылки")
async def reply_my_links(message: Message):
    await links_command(message)

@dp.message(F.text == "👥 Приглашённые")
async def reply_my_refs(message: Message):
    await refs_command(message)

@dp.message(F.text == "🏆 Топ")
async def reply_top(message: Message):
    await top_command(message)

@dp.message(F.text == "🗓 График / заполнить")
async def reply_schedule_menu(message: Message, state: FSMContext):
    if message.chat.type != "private":
        await message.answer("🗓 Заполнение графика доступно в личных сообщениях с ботом.")
        return
    await show_schedule_menu(message, state)

@dp.message(F.text == "⚙️ Админ-панель")
async def reply_admin_panel(message: Message):
    await admin_command(message)

async def _admin_shortcut(message: Message, callback_data: str, title: str):
    if not is_admin(message.from_user.id):
        await message.answer("❌ У тебя нет доступа к этой функции.")
        return
    await message.answer(
        title,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="Открыть", callback_data=callback_data)
        ]])
    )

@dp.message(F.text == "🗓 Управление графиком")
async def reply_admin_schedule(message: Message):
    await _admin_shortcut(message, "admin_schedule", "🗓 Управление графиком")

@dp.message(F.text == "📥 Скачать Excel")
async def reply_export_excel(message: Message):
    await _admin_shortcut(message, "export_excel", "📥 Excel со ссылками")

@dp.message(F.text == "📊 Общая статистика")
async def reply_admin_stats(message: Message):
    await _admin_shortcut(message, "admin_stats", "📊 Общая статистика")

@dp.message(F.text == "📅 Excel с графиками")
async def reply_schedule_excel(message: Message):
    await _admin_shortcut(message, "export_schedule_excel", "📅 Excel с графиками")

@dp.message(F.text == "🏖 Excel с выходными")
async def reply_dayoff_excel(message: Message):
    await _admin_shortcut(message, "export_dayoff_excel", "🏖 Excel с запросами выходных")


# =========================================================
# ТЕХНИЧЕСКИЕ РАБОТЫ — ТОЛЬКО ВЛАДЕЛЕЦ
# =========================================================

def maintenance_keyboard():
    enabled = bool(data.get("maintenance_mode", False))
    if enabled:
        rows = [[InlineKeyboardButton(text="✅ Завершить тех. работы", callback_data="maintenance_finish")]]
    else:
        rows = [[InlineKeyboardButton(text="🛠 Начать тех. работы", callback_data="maintenance_start")]]
    rows.append([InlineKeyboardButton(text="❌ Закрыть", callback_data="maintenance_close")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def maintenance_menu_text():
    enabled = bool(data.get("maintenance_mode", False))
    status = "🟠 <b>Технические работы идут</b>" if enabled else "🟢 <b>Бот работает в штатном режиме</b>"
    return (
        "🛠 <b>ТЕХНИЧЕСКИЕ РАБОТЫ</b>\n\n"
        f"{status}\n\n"
        "Отсюда владелец может отправить служебное уведомление "
        "всем администраторам в ЛС и в группу <b>Admin Staff</b>."
    )


async def send_maintenance_notice(text: str):
    """Deliver to all configured admins and Staff; return truthful delivery summary."""
    import logging
    logger = logging.getLogger("shinobi.delivery")
    successes, failures = [], []
    for admin_id in sorted(set(ADMIN_IDS)):
        try:
            await bot.send_message(admin_id, text, parse_mode="HTML")
            successes.append(f"ЛС {admin_id}")
        except Exception as exc:
            logger.error("Не доставлено в ЛС %s: %s", admin_id, exc)
            failures.append(f"ЛС {admin_id}: {str(exc)[:140]}")
    try:
        await bot.send_message(ADMIN_STAFF_CHAT_ID, text, parse_mode="HTML")
        successes.append("Admin Staff")
    except Exception as exc:
        logger.error("Не доставлено в Staff %s: %s", ADMIN_STAFF_CHAT_ID, exc)
        failures.append(f"Staff: {str(exc)[:140]}")
    return successes, failures


@dp.message(Command("delivery_test"))
async def delivery_test(message: Message):
    if message.from_user.id not in OWNER_IDS or message.chat.type != "private":
        return
    ok, errors = await send_maintenance_notice("🧪 <b>Проверка доставки Shinobi Team</b>")
    await message.answer("📨 Проверка уведомлений\nУспешно: " + (", ".join(ok) or "нет") +
                         "\nОшибки: " + ("; ".join(errors) or "нет"))


@dp.message(F.text == "🛠 Тех. работы")
async def maintenance_reply_button(message: Message):
    if message.from_user.id not in OWNER_IDS:
        return
    await message.answer(
        maintenance_menu_text(),
        reply_markup=maintenance_keyboard(),
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "maintenance_menu")
async def maintenance_menu_callback(callback: CallbackQuery):
    if callback.from_user.id not in OWNER_IDS:
        await callback.answer("⛔ Доступно только владельцу.", show_alert=True)
        return
    await callback.answer()
    await callback.message.answer(
        maintenance_menu_text(),
        reply_markup=maintenance_keyboard(),
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "maintenance_start")
async def maintenance_start_callback(callback: CallbackQuery):
    if callback.from_user.id not in OWNER_IDS:
        await callback.answer("⛔ Доступно только владельцу.", show_alert=True)
        return
    if data.get("maintenance_mode", False):
        await callback.answer("Технические работы уже запущены.", show_alert=True)
        return

    confirm = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚠️ Да, начать", callback_data="maintenance_start_confirm")],
        [InlineKeyboardButton(text="❌ Отмена", callback_data="maintenance_menu")]
    ])
    await callback.answer()
    await callback.message.edit_text(
        "⚠️ <b>ПОДТВЕРЖДЕНИЕ</b>\n\n"
        "Начать технические работы?\n\n"
        "После подтверждения уведомление получат все администраторы в ЛС, "
        "а также группа <b>Admin Staff</b>.",
        reply_markup=confirm,
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "maintenance_start_confirm")
async def maintenance_start_confirm_callback(callback: CallbackQuery):
    if callback.from_user.id not in OWNER_IDS:
        await callback.answer("⛔ Доступно только владельцу.", show_alert=True)
        return

    data["maintenance_mode"] = True
    save_data()

    notice = (
        "🛠 <b>SHINOBI TEAM BOT · ТЕХНИЧЕСКИЕ РАБОТЫ</b>\n\n"
        "На боте начались технические работы.\n\n"
        "⚙️ Сейчас могут проводиться обновления, исправления ошибок "
        "и изменение функционала.\n\n"
        "❗ Некоторые функции бота могут временно работать нестабильно "
        "или быть недоступны.\n\n"
        "Просьба не отправлять одну и ту же команду несколько раз, "
        "если бот отвечает с задержкой.\n\n"
        "🥷 После завершения работ администрация получит отдельное уведомление."
    )
    delivered, failed = await send_maintenance_notice(notice)

    await callback.answer("Технические работы запущены.")
    await callback.message.edit_text(
        "🛠 <b>Технические работы запущены.</b>\n\n"
        f"📨 Доставлено: {len(delivered)}. Ошибок: {len(failed)}.\n" +
        ("\n".join(html.escape(x) for x in failed) if failed else "Все доставки успешны."),
        reply_markup=maintenance_keyboard(),
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "maintenance_finish")
async def maintenance_finish_callback(callback: CallbackQuery):
    if callback.from_user.id not in OWNER_IDS:
        await callback.answer("⛔ Доступно только владельцу.", show_alert=True)
        return
    if not data.get("maintenance_mode", False):
        await callback.answer("Технические работы уже завершены.", show_alert=True)
        return

    data["maintenance_mode"] = False
    save_data()

    notice = (
        "✅ <b>SHINOBI TEAM BOT · ТЕХНИЧЕСКИЕ РАБОТЫ ЗАВЕРШЕНЫ</b>\n\n"
        "Обновление завершено. Бот снова работает в штатном режиме.\n\n"
        "⚙️ Основные функции доступны для использования.\n"
        "🥷 Спасибо за ожидание."
    )
    delivered, failed = await send_maintenance_notice(notice)

    await callback.answer("Технические работы завершены.")
    await callback.message.edit_text(
        "✅ <b>Технические работы завершены.</b>\n\n"
        f"📨 Доставлено: {len(delivered)}. Ошибок: {len(failed)}.\n" +
        ("\n".join(html.escape(x) for x in failed) if failed else "Все доставки успешны."),
        reply_markup=maintenance_keyboard(),
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "maintenance_close")
async def maintenance_close_callback(callback: CallbackQuery):
    if callback.from_user.id not in OWNER_IDS:
        await callback.answer("⛔ Доступно только владельцу.", show_alert=True)
        return
    await callback.answer()
    try:
        await callback.message.delete()
    except Exception:
        pass


# =========================================================
# /admin
# =========================================================

@dp.message(Command("admin"))
async def admin_command(message: Message):

    if not is_admin(
        message.from_user.id
    ):

        await message.answer(
            "❌ У тебя нет доступа "
            "к админ-панели."
        )

        return

    await message.answer(
        "⚙️ <b>АДМИН-ПАНЕЛЬ</b>\n\n"
        "🥷 Shinobi Team · Automation Bot\n\n"
        "Выбери действие:",
        reply_markup=admin_keyboard(),
        parse_mode="HTML"
    )


# =========================================================
# КНОПКА АДМИН-ПАНЕЛЬ
# =========================================================

@dp.callback_query(
    F.data == "admin_panel"
)
async def admin_panel_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer()

    await callback.message.answer(
        "⚙️ <b>АДМИН-ПАНЕЛЬ</b>\n\n"
        "🥷 Shinobi Team · Automation Bot\n\n"
        "Выбери действие:",
        reply_markup=admin_keyboard(),
        parse_mode="HTML"
    )



# =========================================================
# ГРУППЫ БОТА / ДОСТУП К КОМАНДАМ
# =========================================================

def admin_groups_keyboard(page: int = 0):
    groups = []
    installed = {str(x) for x in data.get("installed_groups", [])}
    for cid, g in data.get("known_groups", {}).items():
        if str(cid) in installed and g.get("bot_status") in {"administrator", "creator"}:
            groups.append((cid, g))
    groups.sort(key=lambda item: (item[1].get("title") or item[0]).casefold())
    per_page = 8
    pages = max(1, (len(groups) + per_page - 1) // per_page)
    page = max(0, min(page, pages - 1))
    rows = []
    for cid, g in groups[page * per_page:(page + 1) * per_page]:
        denied = data.get("group_command_access", {}).get(str(cid), "allow") == "deny"
        badge = "🔴" if denied else "🟢"
        title = g.get("title") or f"Группа {cid}"
        rows.append([InlineKeyboardButton(text=f"{badge} {title}"[:60], callback_data=f"admin_group_{cid}_{page}")])
    if not groups:
        rows.append([InlineKeyboardButton(text="Группы пока не обнаружены", callback_data="noop")])
    nav=[]
    if page > 0: nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"admin_groups_{page-1}"))
    nav.append(InlineKeyboardButton(text=f"{page+1}/{pages}", callback_data="noop"))
    if page + 1 < pages: nav.append(InlineKeyboardButton(text="➡️", callback_data=f"admin_groups_{page+1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def admin_group_actions_keyboard(chat_id: int, page: int = 0):
    denied = data.get("group_command_access", {}).get(str(chat_id), "allow") == "deny"
    action_text = "✅ Разрешить команды" if denied else "⛔ Запретить команды"
    action = "allow" if denied else "deny"
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=action_text, callback_data=f"admin_groupaccess_{action}_{chat_id}_{page}")],
        [InlineKeyboardButton(text="⬅️ К группам", callback_data=f"admin_groups_{page}")]
    ])


CALL_EMOJIS = ["🥷", "⚔️", "🔥", "🎮", "💎", "👀", "🌀", "🌙", "⭐", "🖤", "🐉", "🎯", "🗡️", "🦊", "👾", "🚀", "🎲", "🏆"]

@dp.message(F.text.regexp(r"(?i)^\.вызов(?:\s|$)"))
async def staff_call_members(message: Message):
    """Staff-only зазыв участников текущей группы скрытыми упоминаниями по emoji."""
    if message.chat.type not in {"group", "supergroup"} or not message.from_user:
        return

    # .вызов разрешён только в специально назначенной группе.
    if message.chat.id != CALL_ALLOWED_CHAT_ID:
        return

    # Доступ только действующим участникам Admin Staff (-1004462686385).
    if not await _is_admin_staff_member(message.from_user.id):
        await message.reply("⛔ Команда <code>.вызов</code> доступна только участникам <b>Admin Staff</b>.", parse_mode="HTML")
        return

    custom_text = re.sub(r"(?i)^\.вызов\s*", "", message.text or "", count=1).strip()
    known = data.get("group_seen_members", {}).get(str(message.chat.id), {})
    if not known:
        await message.reply("⚠️ Пока нет сохранённых участников этой группы для вызова.")
        return

    mentions = []
    for raw_uid, info in list(known.items()):
        try:
            uid = int(raw_uid)
            if uid == message.from_user.id:
                continue
            member = await bot.get_chat_member(message.chat.id, uid)
            status = getattr(member.status, "value", member.status)
            if status in {"left", "kicked"}:
                continue
            emoji = random.choice(CALL_EMOJIS)
            mentions.append(f'<a href="tg://user?id={uid}">{emoji}</a>')
        except Exception:
            continue

    if not mentions:
        await message.reply("⚠️ Не нашёл доступных участников для вызова.")
        return

    random.shuffle(mentions)
    footer = f"\n\n💬 <b>{html.escape(custom_text)}</b>" if custom_text else ""
    header = "📣 <b>ВЫЗОВ УЧАСТНИКОВ</b>\n\n"
    # Telegram ограничивает одно сообщение примерно 4096 символами.
    chunks, current = [], header
    for mention in mentions:
        addition = mention + " "
        if len(current) + len(addition) + len(footer) > 3900:
            chunks.append(current.rstrip())
            current = ""
        current += addition
    if current.strip():
        chunks.append(current.rstrip())

    for index, chunk in enumerate(chunks):
        tail = footer if index == len(chunks) - 1 else ""
        await message.answer(chunk + tail, parse_mode="HTML", disable_web_page_preview=True)
        if index + 1 < len(chunks):
            await asyncio.sleep(0.8)

@dp.message(F.text.casefold().in_({".установить", "/установить", "/install"}))
async def install_group_command(message: Message):
    if message.chat.type not in {"group", "supergroup"}:
        await message.answer("Эта команда используется непосредственно в группе.")
        return

    # Устанавливать Shinobi Team Bot может только администратор/владелец группы.
    try:
        member = await bot.get_chat_member(message.chat.id, message.from_user.id)
        if member.status not in {"administrator", "creator"}:
            await message.reply("❌ Установить бота может только администратор группы.")
            return
        me = await bot.get_me()
        bot_member = await bot.get_chat_member(message.chat.id, me.id)
        if bot_member.status not in {"administrator", "creator"}:
            await message.reply("❌ Сначала выдайте боту права администратора, затем повторите <code>.установить</code>.", parse_mode="HTML")
            return
    except Exception:
        await message.reply("❌ Не удалось проверить права администратора. Проверьте права бота и повторите команду.")
        return

    cid = str(message.chat.id)
    installed = [str(x) for x in data.setdefault("installed_groups", [])]
    already = cid in installed
    if not already:
        data["installed_groups"].append(cid)

    data.setdefault("known_groups", {})[cid] = {
        "id": message.chat.id,
        "title": message.chat.title or f"Группа {message.chat.id}",
        "bot_status": bot_member.status,
        "installed_by": message.from_user.id,
        "installed_at": datetime.now().isoformat(),
    }
    data.setdefault("group_command_access", {}).setdefault(cid, "allow")
    save_data()

    if already:
        await message.reply("✅ <b>Shinobi Team Bot уже установлен в этой группе.</b>\n\nКоманды работают согласно настройкам доступа владельца.", parse_mode="HTML")
    else:
        await message.reply(
            "🥷 <b>SHINOBI TEAM BOT · УСТАНОВКА</b>\n\n"
            "✅ Группа зарегистрирована в системе.\n"
            f"🆔 <code>{message.chat.id}</code>\n"
            "🟢 Команды бота разрешены.\n\n"
            "Теперь группа отображается у владельца в разделе «🏘 Группы и доступ к командам».",
            parse_mode="HTML"
        )


@dp.callback_query(F.data.startswith("admin_groups_"))
async def admin_groups_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    page = int(callback.data.rsplit("_", 1)[1])
    await callback.answer()
    installed = {str(x) for x in data.get("installed_groups", [])}
    count = sum(1 for cid, g in data.get("known_groups", {}).items() if str(cid) in installed and g.get("bot_status") in {"administrator", "creator"})
    await callback.message.answer(
        f"🏘 <b>ГРУППЫ SHINOBI TEAM</b>\n\nГрупп, где бот администратор: <b>{count}</b>\n"
        "🟢 команды разрешены · 🔴 команды запрещены\n\nВыбери группу:",
        reply_markup=admin_groups_keyboard(page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_group_") & ~F.data.startswith("admin_groups_") & ~F.data.startswith("admin_groupaccess_"))
async def admin_group_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    rest = callback.data[len("admin_group_"):]
    chat_s, page_s = rest.rsplit("_", 1)
    chat_id, page = int(chat_s), int(page_s)
    g = data.get("known_groups", {}).get(str(chat_id))
    if not g:
        await callback.answer("Группа не найдена", show_alert=True); return
    denied = data.get("group_command_access", {}).get(str(chat_id), "allow") == "deny"
    await callback.answer()
    await callback.message.answer(
        f"🏘 <b>{g.get('title') or 'Группа'}</b>\n🆔 <code>{chat_id}</code>\n"
        f"🤖 Статус бота: <b>{g.get('bot_status', 'administrator')}</b>\n"
        f"⌨️ Команды: <b>{'ЗАПРЕЩЕНЫ' if denied else 'РАЗРЕШЕНЫ'}</b>\n\n"
        "При запрете любые команды бота в этой группе не выполняются. Ответ error 404 показывается одному пользователю не чаще одного раза в 4 часа.",
        reply_markup=admin_group_actions_keyboard(chat_id, page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_groupaccess_"))
async def admin_group_access_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    rest = callback.data[len("admin_groupaccess_"):]
    action, rest = rest.split("_", 1)
    chat_s, page_s = rest.rsplit("_", 1)
    chat_id, page = int(chat_s), int(page_s)
    data.setdefault("group_command_access", {})[str(chat_id)] = "deny" if action == "deny" else "allow"
    save_data()
    await callback.answer("Команды запрещены" if action == "deny" else "Команды разрешены")
    await callback.message.answer(
        "🔴 Команды бота в этой группе теперь запрещены." if action == "deny" else "🟢 Команды бота в этой группе снова работают в штатном режиме.",
        reply_markup=admin_group_actions_keyboard(chat_id, page))


@dp.my_chat_member()
async def remember_bot_group_status(event: ChatMemberUpdated):
    if event.chat.type not in {"group", "supergroup"}:
        return
    status = event.new_chat_member.status
    # Сам факт добавления бота в группу больше не считается установкой.
    # Обновляем статус только для групп, где уже выполняли .установить.
    if _group_is_installed(event.chat.id):
        data.setdefault("known_groups", {}).setdefault(str(event.chat.id), {
            "id": event.chat.id,
            "title": event.chat.title or f"Группа {event.chat.id}",
        })
        data["known_groups"][str(event.chat.id)]["title"] = event.chat.title or f"Группа {event.chat.id}"
        data["known_groups"][str(event.chat.id)]["bot_status"] = status
        save_data()


# =========================================================
# ПОЛЬЗОВАТЕЛИ БОТА / ПРАВА АДМИНИСТРАТОРА
# =========================================================

def admin_users_keyboard(page: int = 0):
    users = list(data.get("users", {}).values())
    users.sort(key=lambda u: (u.get("name") or u.get("username") or str(u.get("id"))).casefold())
    per_page = 8
    pages = max(1, (len(users) + per_page - 1) // per_page)
    page = max(0, min(page, pages - 1))
    rows = []
    for u in users[page * per_page:(page + 1) * per_page]:
        uid = int(u.get("id"))
        badge = "👑" if uid in OWNER_IDS else ("🛡" if is_admin(uid) else "👤")
        label = u.get("name") or ("@" + u.get("username", "") if u.get("username") else str(uid))
        rows.append([InlineKeyboardButton(text=f"{badge} {label}"[:55], callback_data=f"admin_user_{uid}_{page}")])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"admin_users_{page-1}"))
    nav.append(InlineKeyboardButton(text=f"{page+1}/{pages}", callback_data="noop"))
    if page + 1 < pages:
        nav.append(InlineKeyboardButton(text="➡️", callback_data=f"admin_users_{page+1}"))
    if nav: rows.append(nav)
    rows.append([InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


class InviteCountEdit(StatesGroup):
    waiting_value = State()


class AdvertisingInviteCountEdit(StatesGroup):
    waiting_value = State()


def admin_user_actions_keyboard(uid: int, page: int = 0):
    rows = [
        [InlineKeyboardButton(text="➕ Добавить", callback_data=f"admin_inv_add_{uid}_{page}"),
         InlineKeyboardButton(text="➖ Отнять", callback_data=f"admin_inv_sub_{uid}_{page}")],
        [InlineKeyboardButton(text="✏️ Установить число", callback_data=f"admin_inv_set_{uid}_{page}")]
    ]
    if uid not in OWNER_IDS:
        if is_admin(uid):
            rows.append([InlineKeyboardButton(text="🛡 Снять права админа", callback_data=f"admin_demote_{uid}_{page}")])
        else:
            rows.append([InlineKeyboardButton(text="🛡 Сделать админом бота", callback_data=f"admin_promote_{uid}_{page}")])
        rows.append([InlineKeyboardButton(text="🗑 Удалить пользователя", callback_data=f"admin_userdel_{uid}_{page}")])
    rows.append([InlineKeyboardButton(text="⬅️ К пользователям", callback_data=f"admin_users_{page}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@dp.callback_query(F.data.startswith("admin_users_"))
async def admin_users_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    page = int(callback.data.rsplit("_", 1)[1])
    await callback.answer()
    await callback.message.answer(
        f"👥 <b>ПОЛЬЗОВАТЕЛИ БОТА</b>\n\nЗарегистрировано: <b>{len(data.get('users', {}))}</b>\n👑 — владелец · 🛡 — админ · 👤 — пользователь\n\nВыбери пользователя:",
        reply_markup=admin_users_keyboard(page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_user_"))
async def admin_user_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    parts = callback.data.split("_")
    uid, page = int(parts[2]), int(parts[3])
    u = data.get("users", {}).get(str(uid))
    if not u:
        await callback.answer("Пользователь не найден", show_alert=True); return
    role = "Владелец" if uid in OWNER_IDS else ("Администратор бота" if is_admin(uid) else "Пользователь")
    username = f"@{u.get('username')}" if u.get('username') else "—"
    await callback.answer()
    await callback.message.answer(
        f"👤 <b>{u.get('name') or 'Без имени'}</b>\n🆔 <code>{uid}</code>\n🔗 {username}\n🔐 Роль: <b>{role}</b>\n👥 Приглашено: <b>{u.get('total_invited', 0)}</b>",
        reply_markup=admin_user_actions_keyboard(uid, page), parse_mode="HTML")


@dp.callback_query(F.data.in_({"ad_inv_add", "ad_inv_sub", "ad_inv_set"}))
async def advertising_invite_edit_start(callback: CallbackQuery, state: FSMContext):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True)
        return
    action = callback.data.rsplit("_", 1)[-1]
    current = max(0, int(data.get("links", {}).get(ADVERTISING_LINK, {}).get("invited", 0) or 0))
    labels = {"add": "добавить", "sub": "отнять", "set": "установить"}
    await state.set_state(AdvertisingInviteCountEdit.waiting_value)
    await state.update_data(ad_invite_action=action)
    await callback.answer()
    await callback.message.answer(
        f"📣 <b>Рекламная ссылка</b>\n{ADVERTISING_LINK}\n\n"
        f"Сейчас приглашено: <b>{current}</b>\n"
        f"Введи целое число, которое нужно <b>{labels[action]}</b>.",
        parse_mode="HTML", disable_web_page_preview=True
    )


@dp.message(AdvertisingInviteCountEdit.waiting_value)
async def advertising_invite_edit_value(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    raw = (message.text or "").strip()
    if not raw.isdigit():
        await message.answer("❌ Введи целое неотрицательное число, например <code>5</code>.", parse_mode="HTML")
        return
    value = int(raw)
    ctx = await state.get_data()
    action = ctx.get("ad_invite_action")
    info = data.setdefault("links", {}).setdefault(ADVERTISING_LINK, {
        "owner_id": None, "type": "advertising", "label": "Рекламная ссылка",
        "created": datetime.now().isoformat(), "invited": 0, "users": [], "join_request": True
    })
    old = max(0, int(info.get("invited", 0) or 0))
    if action == "add": new = old + value
    elif action == "sub": new = max(0, old - value)
    elif action == "set": new = value
    else:
        await state.clear(); await message.answer("❌ Неизвестное действие."); return
    info["invited"] = new
    data["advertising_link"] = ADVERTISING_LINK
    action_names = {"add": "Добавлено", "sub": "Отнято", "set": "Установлено"}
    data.setdefault("invite_adjustments", []).append({
        "created": datetime.now().isoformat(timespec="seconds"),
        "admin_id": message.from_user.id, "admin_name": message.from_user.full_name,
        "user_id": "ADVERTISING", "username": "", "name": "Рекламная ссылка",
        "link": ADVERTISING_LINK, "action": action_names.get(action, action),
        "value": value, "old": old, "new": new
    })
    save_data()
    await state.clear()
    await message.answer(
        "✅ <b>Рекламная ссылка обновлена</b>\n\n"
        f"🔗 {ADVERTISING_LINK}\n"
        f"Было: <b>{old}</b>\nСтало: <b>{new}</b>\n\n"
        "💾 Изменение сохранено и попадёт в Excel / Shinobi Stats.",
        parse_mode="HTML", disable_web_page_preview=True
    )


@dp.callback_query(F.data.startswith("admin_inv_"))
async def admin_invite_edit_start(callback: CallbackQuery, state: FSMContext):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True)
        return
    parts = callback.data.split("_")
    action, uid, page = parts[2], int(parts[3]), int(parts[4])
    if str(uid) not in data.get("users", {}):
        await callback.answer("Пользователь не найден", show_alert=True)
        return
    await state.set_state(InviteCountEdit.waiting_value)
    await state.update_data(invite_action=action, invite_uid=uid, invite_page=page)
    labels = {"add": "добавить", "sub": "отнять", "set": "установить"}
    await callback.answer()
    await callback.message.answer(
        f"👥 Сейчас приглашено: <b>{data['users'][str(uid)].get('total_invited', 0)}</b>\n\n"
        f"Введи целое число, которое нужно <b>{labels[action]}</b>.\nНапример: <code>5</code>",
        parse_mode="HTML"
    )


@dp.message(InviteCountEdit.waiting_value)
async def admin_invite_edit_value(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    raw = (message.text or "").strip()
    if not raw.isdigit():
        await message.answer("❌ Введи целое неотрицательное число, например <code>5</code>.", parse_mode="HTML")
        return
    value = int(raw)
    ctx = await state.get_data()
    action, uid, page = ctx.get("invite_action"), int(ctx.get("invite_uid")), int(ctx.get("invite_page", 0))
    user = data.get("users", {}).get(str(uid))
    if not user:
        await state.clear(); await message.answer("❌ Пользователь больше не найден."); return
    old = max(0, int(user.get("total_invited", 0) or 0))
    if action == "add": new = old + value
    elif action == "sub": new = max(0, old - value)
    elif action == "set": new = value
    else:
        await state.clear(); await message.answer("❌ Неизвестное действие."); return
    user["total_invited"] = new

    # Ручную корректировку привязываем к основной ссылке пользователя, если она есть.
    user_links = [link for link in user.get("links", []) if link in data.get("links", {})]
    primary_link = user_links[0] if user_links else ""
    delta = new - old
    if primary_link:
        link_info = data["links"][primary_link]
        link_info["invited"] = max(0, int(link_info.get("invited", 0) or 0) + delta)

    action_names = {"add": "Добавлено", "sub": "Отнято", "set": "Установлено"}
    data.setdefault("invite_adjustments", []).append({
        "created": datetime.now().isoformat(timespec="seconds"),
        "admin_id": message.from_user.id,
        "admin_name": message.from_user.full_name,
        "user_id": uid,
        "username": ("@" + user.get("username", "")) if user.get("username") else "",
        "name": user.get("name", ""),
        "link": primary_link,
        "action": action_names.get(action, action),
        "value": value,
        "old": old,
        "new": new
    })
    save_data()
    await state.clear()
    link_text = primary_link if primary_link else "Нет сохранённой ссылки"
    await message.answer(
        f"✅ <b>Количество приглашённых изменено</b>\n\n"
        f"👤 {display_user(user)}\n"
        f"🔗 {link_text}\n"
        f"Было: <b>{old}</b>\nСтало: <b>{new}</b>\n\n"
        "💾 Изменение записано в data.json и Excel.",
        parse_mode="HTML", reply_markup=admin_user_actions_keyboard(uid, page),
        disable_web_page_preview=True
    )


@dp.callback_query(F.data.startswith("admin_promote_"))
async def admin_promote_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid not in ADMIN_IDS: ADMIN_IDS.append(uid)
    admins = {int(x) for x in data.get("bot_admins", []) if str(x).isdigit()}; admins.add(uid)
    data["bot_admins"] = sorted(admins); save_data(); init_schedule_system(ADMIN_IDS)
    await callback.answer("Администратор добавлен")
    await callback.message.answer("✅ Пользователю выданы права администратора бота. Ему доступна админ-панель.", reply_markup=admin_user_actions_keyboard(uid, page))


@dp.callback_query(F.data.startswith("admin_demote_"))
async def admin_demote_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid in OWNER_IDS: await callback.answer("Владельца снять нельзя", show_alert=True); return
    if uid in ADMIN_IDS: ADMIN_IDS.remove(uid)
    data["bot_admins"] = [int(x) for x in data.get("bot_admins", []) if int(x) != uid]
    save_data(); init_schedule_system(ADMIN_IDS)
    await callback.answer("Права сняты")
    await callback.message.answer("✅ Права администратора сняты.", reply_markup=admin_user_actions_keyboard(uid, page))


@dp.callback_query(F.data.startswith("admin_userdel_"))
async def admin_user_delete_confirm(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid in OWNER_IDS: await callback.answer("Владельца удалить нельзя", show_alert=True); return
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="✅ Да, удалить", callback_data=f"admin_userdelok_{uid}_{page}"), InlineKeyboardButton(text="❌ Отмена", callback_data=f"admin_user_{uid}_{page}")]])
    await callback.answer(); await callback.message.answer("⚠️ Удалить пользователя из базы бота и его персональные ссылки?", reply_markup=kb)


@dp.callback_query(F.data.startswith("admin_userdelok_"))
async def admin_user_delete(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid in OWNER_IDS: await callback.answer("Владельца удалить нельзя", show_alert=True); return
    user = data.get("users", {}).pop(str(uid), None)
    # Удаляем созданные им ссылки из общей базы.
    for key in list(data.get("links", {})):
        rec = data["links"][key]
        creator = rec.get("creator_id", rec.get("owner_id", rec.get("user_id"))) if isinstance(rec, dict) else None
        if str(creator) == str(uid): data["links"].pop(key, None)
    if uid in ADMIN_IDS: ADMIN_IDS.remove(uid)
    data["bot_admins"] = [int(x) for x in data.get("bot_admins", []) if int(x) != uid]
    save_data(); init_schedule_system(ADMIN_IDS)
    await callback.answer("Пользователь удалён")
    await callback.message.answer("🗑 Пользователь удалён из базы бота. При следующем /start он зарегистрируется заново.", reply_markup=admin_users_keyboard(page))


@dp.callback_query(F.data == "noop")
async def noop_callback(callback: CallbackQuery):
    await callback.answer()

# =========================================================
# ОБЩАЯ СТАТИСТИКА
# =========================================================

@dp.callback_query(
    F.data == "admin_stats"
)
async def admin_stats_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer()

    total_users = len(
        data["users"]
    )

    total_links = len(
        data["links"]
    )

    total_invited = sum(
        user.get(
            "total_invited",
            0
        )
        for user in data[
            "users"
        ].values()
    )

    await callback.message.answer(
        "📊 <b>ОБЩАЯ СТАТИСТИКА</b>\n\n"
        f"👤 Пользователей бота: "
        f"<b>{total_users}</b>\n"
        f"🔗 Создано ссылок: "
        f"<b>{total_links}</b>\n"
        f"👥 Приглашено: "
        f"<b>{total_invited}</b>",
        parse_mode="HTML"
    )


# =========================================================
# EXCEL
# =========================================================

@dp.callback_query(
    F.data == "export_excel"
)
async def export_excel_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer(
        "📊 Создаю Excel..."
    )

    try:

        filename = create_excel()

        document = FSInputFile(
            filename
        )

        total_invited = sum(
            user.get(
                "total_invited",
                0
            )
            for user in data[
                "users"
            ].values()
        )

        await callback.message.answer_document(
            document=document,
            caption=(
                "📊 <b>SHINOBI TEAM</b>\n"
                "<b>Реферальный отчёт</b>\n\n"
                f"👤 Пользователей: "
                f"<b>{len(data['users'])}</b>\n"
                f"🔗 Ссылок: "
                f"<b>{len(data['links'])}</b>\n"
                f"👥 Приглашено: "
                f"<b>{total_invited}</b>\n\n"
                "📄 Пользователи и их итоговая статистика\n"
                "📄 Все реферальные ссылки\n"
                "📄 Реально приглашённые пользователи\n"
                "📄 История ручных изменений\n"
                "📄 Системный снимок"
            ),
            parse_mode="HTML"
        )

    except Exception as error:

        await callback.message.answer(
            "❌ Ошибка создания Excel:\n\n"
            f"<code>{error}</code>",
            parse_mode="HTML"
        )

    finally:
        # Excel теперь является постоянной резервной копией и не удаляется после отправки.
        pass



# =========================================================
# SHINOBI STATS
# =========================================================

@dp.callback_query(F.data == "shinobi_stats_report")
async def shinobi_stats_report_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа", show_alert=True)
        return
    await callback.answer()
    try:
        # Обновляем главный реферальный снимок перед архивированием.
        create_excel()
        report = create_report_snapshot(requested_by=callback.from_user.id)
        await callback.message.answer(
            "🔴 <b>SHINOBI STATS</b>\n\n"
            "Отчёт успешно сформирован и сохранён в архиве.\n\n"
            f"🆔 <code>{report['id']}</code>\n"
            f"🕒 {report['created_at'].replace('T', ' ')}\n"
            f"📁 Разделов: <b>{len(report['files'])}</b>\n\n"
            "Запустите Shinobi Stats на ПК и нажмите «Синхронизировать».",
            parse_mode="HTML"
        )
    except Exception as error:
        await callback.message.answer(f"❌ Не удалось создать отчёт:\n<code>{error}</code>", parse_mode="HTML")


async def october_content_reminder_loop():
    """One private reminder per October plan date, only to the bot creator."""
    plan_path = os.path.join(BASE_DIR, "october_content_plan.json")
    tz = ZoneInfo("Europe/Warsaw")
    while True:
        try:
            now = datetime.now(tz)
            if now.year == 2026 and now.month == 10 and now.hour >= 10:
                try:
                    with open(plan_path, "r", encoding="utf-8") as f: plan = json.load(f)
                except Exception:
                    plan = {"items": []}
                today = now.strftime("%Y-%m-%d"); changed = False
                for item in plan.get("items", []):
                    if item.get("date") != today or item.get("done") or item.get("reminder_sent"): continue
                    text = (
                        "💗 <b>Люблю Сашу</b>\n\n"
                        "📅 <b>Напоминание по контент-плану на сегодня</b>\n\n"
                        f"🔥 <b>{item.get('title','')}</b>\n"
                        f"🎮 Формат: <b>{item.get('format','')}</b>\n\n"
                        f"📝 Нужно выставить пост:\n{item.get('post_text','')}\n\n"
                        "После публикации отметь задачу выполненной в плане октября. ✅"
                    )
                    for owner_id in OWNER_IDS:
                        try: await bot.send_message(owner_id, text, parse_mode="HTML")
                        except Exception as exc: print(f"[OCTOBER REMINDER] {owner_id}: {exc}")
                    item["reminder_sent"] = True; changed = True
                if changed:
                    tmp = plan_path + ".tmp"
                    with open(tmp, "w", encoding="utf-8") as f: json.dump(plan, f, ensure_ascii=False, indent=2)
                    os.replace(tmp, plan_path)
        except Exception as exc:
            print(f"[OCTOBER REMINDER] loop: {exc}")
        await asyncio.sleep(60)

# =========================================================
# ЗАПУСК
# =========================================================

async def chat_summary_broadcast_loop():
    # First summary is sent 6 hours after startup, then every 6 hours.
    await asyncio.sleep(6 * 60 * 60)
    while True:
        try:
            stats_by_id = {int(x["chat_id"]): x for x in chat_stats_snapshot()}
            for raw_chat_id in data.get("installed_groups", []):
                chat_id = int(raw_chat_id)
                # Do not auto-post chat summaries to Staff or game chats.
                excluded_summary_chats = {ADMIN_STAFF_CHAT_ID, CHAT_ID, GAME_COMMAND_ALLOWED_CHAT_ID}
                if chat_id in excluded_summary_chats:
                    continue
                try:
                    started = time.perf_counter()
                    members = await bot.get_chat_member_count(chat_id)
                    api_ping = round((time.perf_counter() - started) * 1000)
                    st = stats_by_id.get(chat_id, {})
                    active = int(st.get("active_15m", 0))
                    today = int(st.get("messages_today", 0))
                    peak_count = int(st.get("peak_messages", 0))
                    peak_hour = st.get("peak_hour", "—")
                    text = (
                        "📊 <b>Сводка чата</b>\n\n"
                        f"🏓 Пинг Telegram API: <b>{api_ping} мс</b>\n"
                        f"👥 Участников: <b>{members}</b>\n"
                        f"🟢 Активны за 15 мин: <b>{active}</b>\n"
                        f"💬 Сообщений за день: <b>{today}</b>\n"
                        f"🔥 Максимальный пик: <b>{peak_count}</b> сообщений/час\n"
                        f"🕒 Время пика: <b>{peak_hour}</b>\n\n"
                        f"💡 <b>Интересный факт MLBB:</b> {mlbb_random_fact()}"
                    )
                    await bot.send_message(chat_id, text, parse_mode="HTML")
                except Exception as exc:
                    print(f"[CHAT SUMMARY] send {chat_id}: {exc}")
        except Exception as exc:
            print(f"[CHAT SUMMARY] loop: {exc}")
        await asyncio.sleep(6 * 60 * 60)

async def main():
    print("Удаляю старый webhook...")

    # Удаляем webhook перед запуском polling
    await bot.delete_webhook(drop_pending_updates=False)

    print("Webhook удалён; ожидающие обновления Telegram сохранены.")
    print("=" * 50)
    print(" SHINOBI TEAM REFERRAL BOT")
    print(" БОТ ЗАПУЩЕН")
    print("=" * 50)

    await start_stats_api()

    # Каждые несколько секунд зеркалим JSON, Excel и архив Shinobi Stats в PostgreSQL.
    asyncio.create_task(persistence_loop())
    asyncio.create_task(staff_random_reply_loop())
    asyncio.create_task(server_tag_cleanup_loop())
    asyncio.create_task(chat_summary_broadcast_loop())
    asyncio.create_task(october_content_reminder_loop())

    await dp.start_polling(
        bot,
        allowed_updates=[
            "message",
            "callback_query",
            "chat_member",
            "my_chat_member",
            "chat_join_request",
            "message_reaction"
        ]
    )


if __name__ == "__main__":
    asyncio.run(main())
