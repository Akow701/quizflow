import os
import json
import sqlite3
import logging
import tempfile
from pathlib import Path

from telegram import (
    Update, Poll, InlineKeyboardButton, InlineKeyboardMarkup
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    PollAnswerHandler, CallbackQueryHandler,
    ContextTypes, filters
)

# File parsers
import pdfplumber
import docx
import openpyxl

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ["BOT_TOKEN"]
CLAUDE_API_KEY = os.environ["CLAUDE_API_KEY"]

# ─── DATABASE ────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS scores (
            chat_id     INTEGER,
            user_id     INTEGER,
            username    TEXT,
            first_name  TEXT,
            correct     INTEGER DEFAULT 0,
            total       INTEGER DEFAULT 0,
            PRIMARY KEY (chat_id, user_id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS active_polls (
            poll_id      TEXT PRIMARY KEY,
            chat_id      INTEGER,
            correct_idx  INTEGER,
            question     TEXT
        )
    """)
    conn.commit()
    conn.close()

def upsert_score(chat_id, user_id, username, first_name, correct: bool):
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("""
        INSERT INTO scores (chat_id, user_id, username, first_name, correct, total)
        VALUES (?, ?, ?, ?, ?, 1)
        ON CONFLICT(chat_id, user_id) DO UPDATE SET
            username   = excluded.username,
            first_name = excluded.first_name,
            correct    = correct + excluded.correct,
            total      = total + 1
    """, (chat_id, user_id, username, first_name, 1 if correct else 0))
    conn.commit()
    conn.close()

def get_leaderboard(chat_id, limit=10):
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("""
        SELECT first_name, username, correct, total
        FROM scores
        WHERE chat_id = ?
        ORDER BY correct DESC, total ASC
        LIMIT ?
    """, (chat_id, limit))
    rows = c.fetchall()
    conn.close()
    return rows

def save_poll(poll_id, chat_id, correct_idx, question):
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("""
        INSERT OR REPLACE INTO active_polls (poll_id, chat_id, correct_idx, question)
        VALUES (?, ?, ?, ?)
    """, (poll_id, chat_id, correct_idx, question))
    conn.commit()
    conn.close()

def get_poll(poll_id):
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("SELECT chat_id, correct_idx, question FROM active_polls WHERE poll_id=?", (poll_id,))
    row = c.fetchone()
    conn.close()
    return row

# ─── FILE READERS ─────────────────────────────────────────────────────────────

def read_txt(path):
    with open(path, encoding="utf-8", errors="ignore") as f:
        return f.read()

def read_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text

def read_docx(path):
    doc = docx.Document(path)
    return "\n".join(p.text for p in doc.paragraphs)

def read_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    lines = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            line = "\t".join(str(c) if c is not None else "" for c in row)
            if line.strip():
                lines.append(line)
    return "\n".join(lines)

def extract_text(file_path: str) -> str:
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return read_pdf(file_path)
    elif ext in (".docx", ".doc"):
        return read_docx(file_path)
    elif ext in (".xlsx", ".xls"):
        return read_xlsx(file_path)
    else:
        return read_txt(file_path)

# ─── CLAUDE API ──────────────────────────────────────────────────────────────

import anthropic

claude = anthropic.Anthropic(api_key=CLAUDE_API_KEY)

def generate_quiz(raw_text: str) -> list[dict]:
    """
    Returns list of:
      { question, options: [str], correct_index: int, explanation: str }
    """
    prompt = f"""Quyidagi matndan Telegram quiz uchun savollar yarat.

MATN:
{raw_text[:12000]}

VAZIFANG:
1. Matndagi barcha savol-javoblarni top
2. To'g'ri javobni aniqla (matnda * = + / ' kabi belgilar bilan ko'rsatilgan bo'lishi mumkin)
3. Quyidagi JSON formatda qaytар

MUHIM QOIDALAR:
- Faqat JSON qaytар, hech qanday izoh yoki markdown yo'q
- correct_index: 0=birinchi variant, 1=ikkinchi, 2=uchinchi, 3=to'rtinchi
- options: kamida 2 ta, ko'pi bilan 4 ta
- Agar matnda to'g'ri javob belgisi bo'lsa (*, =, +, /, to'g'ri javob: ...) — undan foydalan
- Agar belgisi yo'q bo'lsa — kontekstga qarab o'zing aniqlа

FORMAT (faqat shu JSON):
{{
  "questions": [
    {{
      "question": "Savol matni",
      "options": ["A variant", "B variant", "C variant", "D variant"],
      "correct_index": 0,
      "explanation": "Nima uchun to'g'ri ekanini 1 jumlada tushuntir"
    }}
  ]
}}"""

    response = claude.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    # Strip possible markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    data = json.loads(raw)
    return data.get("questions", [])

# ─── HANDLERS ────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Salom! Men *QuizBot* man.\n\n"
        "📎 PDF, DOCX, TXT yoki XLSX fayl yuboring — men undan quiz yasayman!\n\n"
        "📊 /top — leaderboard ko'rish\n"
        "❓ /help — yordam",
        parse_mode="Markdown"
    )

async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *Yordam*\n\n"
        "1️⃣ Fayl yuboring (PDF / DOCX / TXT / XLSX)\n"
        "2️⃣ Bot faylni o'qiydi va savollar yaratadi\n"
        "3️⃣ Quiz boshlanadi — hamma javob beradi\n"
        "4️⃣ /top — kimning bali ko'p ekanini ko'ring\n\n"
        "⚠️ Guruhda ham, lichkada ham ishlaydi\n"
        "🏆 Leaderboard har guruh/chat uchun alohida",
        parse_mode="Markdown"
    )

async def cmd_top(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    rows = get_leaderboard(chat_id)
    if not rows:
        await update.message.reply_text("📊 Hali hech kim quiz yechmagan.")
        return

    medals = ["🥇", "🥈", "🥉"]
    lines = ["🏆 *Leaderboard*\n"]
    for i, (first_name, username, correct, total) in enumerate(rows):
        medal = medals[i] if i < 3 else f"{i+1}."
        pct = round(correct / total * 100) if total else 0
        name = f"@{username}" if username else first_name
        lines.append(f"{medal} {name} — {correct}/{total} ({pct}%)")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_reset(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Only group admins can reset leaderboard"""
    chat = update.effective_chat
    user = update.effective_user
    if chat.type in ("group", "supergroup"):
        admins = await ctx.bot.get_chat_administrators(chat.id)
        admin_ids = [a.user.id for a in admins]
        if user.id not in admin_ids:
            await update.message.reply_text("❌ Faqat adminlar reset qila oladi.")
            return
    conn = sqlite3.connect("quiz.db")
    conn.execute("DELETE FROM scores WHERE chat_id=?", (chat.id,))
    conn.commit()
    conn.close()
    await update.message.reply_text("✅ Leaderboard tozalandi.")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    fname = doc.file_name or ""
    ext = Path(fname).suffix.lower()

    if ext not in (".pdf", ".docx", ".doc", ".txt", ".xlsx", ".xls"):
        await update.message.reply_text(
            "❌ Faqat PDF, DOCX, TXT yoki XLSX fayllar qabul qilinadi."
        )
        return

    msg = await update.message.reply_text("⏳ Fayl o'qilmoqda...")

    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp_path = tmp.name

    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(tmp_path)
        raw_text = extract_text(tmp_path)
    except Exception as e:
        await msg.edit_text(f"❌ Faylni o'qishda xato: {e}")
        return
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    if not raw_text.strip():
        await msg.edit_text("❌ Fayl bo'sh yoki o'qib bo'lmadi.")
        return

    await msg.edit_text("🤖 Claude savollar yaratmoqda...")

    try:
        questions = generate_quiz(raw_text)
    except Exception as e:
        await msg.edit_text(f"❌ Quiz yaratishda xato: {e}")
        return

    if not questions:
        await msg.edit_text("❌ Savollar topilmadi. Faylda savol-javob borligini tekshiring.")
        return

    await msg.edit_text(f"✅ {len(questions)} ta savol topildi! Quiz boshlanmoqda...\n\n"
                        f"━━━━━━━━━━━━━━━━")

    # Store questions in context for sequential sending
    ctx.chat_data["pending_questions"] = questions
    ctx.chat_data["q_index"] = 0
    await send_next_question(update.effective_chat.id, ctx)

async def send_next_question(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE):
    questions = ctx.chat_data.get("pending_questions", [])
    idx = ctx.chat_data.get("q_index", 0)

    if idx >= len(questions):
        await ctx.bot.send_message(chat_id, "🎉 Quiz tugadi! /top — natijalarni ko'ring")
        ctx.chat_data["pending_questions"] = []
        return

    q = questions[idx]
    question_text = q.get("question", "Savol?")
    options = q.get("options", [])
    correct_idx = q.get("correct_index", 0)

    # Validate
    if len(options) < 2:
        ctx.chat_data["q_index"] = idx + 1
        await send_next_question(chat_id, ctx)
        return

    correct_idx = max(0, min(correct_idx, len(options) - 1))
    explanation = q.get("explanation", "")

    # Telegram poll: max option length 100 chars
    options_clean = [str(o)[:100] for o in options]
    question_clean = f"❓ {question_text}"[:300]

    try:
        poll_msg = await ctx.bot.send_poll(
            chat_id=chat_id,
            question=question_clean,
            options=options_clean,
            type=Poll.QUIZ,
            correct_option_id=correct_idx,
            explanation=explanation[:200] if explanation else None,
            is_anonymous=False,
            open_period=30,
        )
        save_poll(poll_msg.poll.id, chat_id, correct_idx, question_text)
        ctx.chat_data["q_index"] = idx + 1

        # Schedule next question after 35 seconds
        ctx.job_queue.run_once(
            next_question_job,
            35,
            chat_id=chat_id,
            data={"chat_id": chat_id},
            name=f"quiz_{chat_id}_{idx}"
        )
    except Exception as e:
        logger.error(f"Poll yuborishda xato: {e}")
        ctx.chat_data["q_index"] = idx + 1
        await send_next_question(chat_id, ctx)

async def next_question_job(ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = ctx.job.chat_id
    await send_next_question(chat_id, ctx)

async def handle_poll_answer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    answer = update.poll_answer
    poll_id = answer.poll_id
    user = answer.user

    row = get_poll(poll_id)
    if not row:
        return

    chat_id, correct_idx, question = row
    chosen = answer.option_ids

    if not chosen:
        return

    is_correct = chosen[0] == correct_idx
    upsert_score(
        chat_id=chat_id,
        user_id=user.id,
        username=user.username or "",
        first_name=user.first_name or "Nomsiz",
        correct=is_correct
    )

# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("top", cmd_top))
    app.add_handler(CommandHandler("reset", cmd_reset))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(PollAnswerHandler(handle_poll_answer))

    logger.info("Bot ishga tushdi ✅")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
