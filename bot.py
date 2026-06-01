import os
import re
import sqlite3
import logging
import tempfile
from pathlib import Path

from telegram import Update, Poll
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    PollAnswerHandler, ContextTypes, filters
)

import pdfplumber
import docx
import openpyxl

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ["BOT_TOKEN"]

# ─── DATABASE ────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS scores (
            chat_id    INTEGER,
            user_id    INTEGER,
            username   TEXT,
            first_name TEXT,
            correct    INTEGER DEFAULT 0,
            total      INTEGER DEFAULT 0,
            PRIMARY KEY (chat_id, user_id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS active_polls (
            poll_id     TEXT PRIMARY KEY,
            chat_id     INTEGER,
            correct_idx INTEGER
        )
    """)
    conn.commit()
    conn.close()

def upsert_score(chat_id, user_id, username, first_name, is_correct):
    conn = sqlite3.connect("quiz.db")
    conn.execute("""
        INSERT INTO scores (chat_id, user_id, username, first_name, correct, total)
        VALUES (?, ?, ?, ?, ?, 1)
        ON CONFLICT(chat_id, user_id) DO UPDATE SET
            username=excluded.username, first_name=excluded.first_name,
            correct=correct+excluded.correct, total=total+1
    """, (chat_id, user_id, username, first_name, 1 if is_correct else 0))
    conn.commit()
    conn.close()

def get_leaderboard(chat_id):
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("""
        SELECT first_name, username, correct, total FROM scores
        WHERE chat_id=? ORDER BY correct DESC, total ASC LIMIT 10
    """, (chat_id,))
    rows = c.fetchall()
    conn.close()
    return rows

def save_poll(poll_id, chat_id, correct_idx):
    conn = sqlite3.connect("quiz.db")
    conn.execute("INSERT OR REPLACE INTO active_polls VALUES (?,?,?)",
                 (poll_id, chat_id, correct_idx))
    conn.commit()
    conn.close()

def get_poll(poll_id):
    conn = sqlite3.connect("quiz.db")
    c = conn.cursor()
    c.execute("SELECT chat_id, correct_idx FROM active_polls WHERE poll_id=?", (poll_id,))
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

def extract_text(file_path):
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":       return read_pdf(file_path)
    elif ext in (".docx",): return read_docx(file_path)
    elif ext in (".xlsx",): return read_xlsx(file_path)
    else:                   return read_txt(file_path)

# ─── PARSER ──────────────────────────────────────────────────────────────────
# Format:
#   1. Savol matni
#   a) ...  b) ...  c) ...  d) ...
#   to'g'ri: b
#
# to'g'ri javob ko'rsatkichlari: to'g'ri/togri/javob/answer + : + harf
# YOKI variantdan keyin * + = belgisi

ANSWER_MAP = {"a": 0, "b": 1, "c": 2, "d": 3}

def parse_questions(text):
    questions = []

    # Savollarni bloklarga ajratamiz — raqam + nuqta/qavs bilan boshlanadi
    blocks = re.split(r"\n(?=\s*\d+[\.\)]\s)", text)

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        lines = [l.strip() for l in block.splitlines() if l.strip()]
        if not lines:
            continue

        # Birinchi qator — savol (raqamni olib tashlaymiz)
        question_line = re.sub(r"^\d+[\.\)]\s*", "", lines[0]).strip()
        if not question_line:
            continue

        options = []
        correct_idx = None

        for line in lines[1:]:
            # to'g'ri javob qatori: "to'g'ri: b" / "togri: b" / "javob: b" / "answer: b"
            m = re.match(r"(?:to['\u2019]?g['\u2019]?ri|togri|javob|answer)\s*[:\-]\s*([a-dA-D])", line, re.IGNORECASE)
            if m:
                correct_idx = ANSWER_MAP.get(m.group(1).lower())
                continue

            # Variant qatori: a) ... yoki a. ... yoki A) ...
            m = re.match(r"^([a-dA-D])\s*[\.\)]\s*(.+)", line)
            if m:
                letter = m.group(1).lower()
                opt_text = m.group(2).strip()

                # Variantdan keyin * + = belgisi bo'lsa — u to'g'ri javob
                if re.search(r"[*+=]", opt_text):
                    correct_idx = ANSWER_MAP.get(letter)
                    opt_text = re.sub(r"\s*[*+=]\s*$", "", opt_text).strip()

                options.append(opt_text)

        # Kamida 2 variant va to'g'ri javob bo'lishi shart
        if len(options) >= 2 and correct_idx is not None:
            correct_idx = min(correct_idx, len(options) - 1)
            questions.append({
                "question": question_line,
                "options": options,
                "correct_index": correct_idx
            })

    return questions

# ─── HANDLERS ────────────────────────────────────────────────────────────────

FORMAT_MSG = (
    "📋 *Fayl formati:*\n\n"
    "```\n"
    "1. Savol matni\n"
    "a) Birinchi javob\n"
    "b) Ikkinchi javob\n"
    "c) To'g'ri javob\n"
    "d) To'rtinchi javob\n"
    "to'g'ri: c\n"
    "\n"
    "2. Keyingi savol...\n"
    "```\n\n"
    "✅ `to'g'ri: c` o'rniga variant yoniga `*` qo'ysa ham bo'ladi:\n"
    "`c) To'g'ri javob *`"
)

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Salom! Men *QuizBot* man.\n\n"
        "📎 Fayl yuboring — men quizga aylantirib beray!\n\n"
        "📌 /format — fayl qanday bo'lishi kerak\n"
        "🏆 /top — leaderboard\n"
        "❓ /help — yordam",
        parse_mode="Markdown"
    )

async def cmd_format(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(FORMAT_MSG, parse_mode="Markdown")

async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *Yordam*\n\n"
        "1️⃣ /format buyrug'i bilan fayl formatini ko'ring\n"
        "2️⃣ PDF / DOCX / TXT / XLSX fayl yuboring\n"
        "3️⃣ Quiz boshlanadi, 30 soniya vaqt\n"
        "4️⃣ /top — kim oldinroq ekanini ko'ring\n\n"
        "👥 Guruhda ham, lichkada ham ishlaydi\n"
        "🏆 Har guruh/chat uchun alohida leaderboard",
        parse_mode="Markdown"
    )

async def cmd_top(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rows = get_leaderboard(update.effective_chat.id)
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
    chat = update.effective_chat
    user = update.effective_user
    if chat.type in ("group", "supergroup"):
        admins = await ctx.bot.get_chat_administrators(chat.id)
        if user.id not in [a.user.id for a in admins]:
            await update.message.reply_text("❌ Faqat adminlar reset qila oladi.")
            return
    conn = sqlite3.connect("quiz.db")
    conn.execute("DELETE FROM scores WHERE chat_id=?", (chat.id,))
    conn.commit()
    conn.close()
    await update.message.reply_text("✅ Leaderboard tozalandi.")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    ext = Path(doc.file_name or "").suffix.lower()

    if ext not in (".pdf", ".docx", ".txt", ".xlsx"):
        await update.message.reply_text(
            "❌ Faqat PDF, DOCX, TXT yoki XLSX fayllar.\n\n"
            "📌 /format — fayl qanday bo'lishi kerak"
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

    questions = parse_questions(raw_text)

    if not questions:
        await msg.edit_text(
            "❌ Savollar topilmadi.\n\n"
            "📌 /format — fayl qanday bo'lishi kerakligini ko'ring"
        )
        return

    await msg.edit_text(f"✅ {len(questions)} ta savol topildi! Quiz boshlanmoqda...\n━━━━━━━━━━━━")

    ctx.chat_data["questions"] = questions
    ctx.chat_data["q_index"] = 0
    await send_next(update.effective_chat.id, ctx)

async def send_next(chat_id, ctx):
    questions = ctx.chat_data.get("questions", [])
    idx = ctx.chat_data.get("q_index", 0)

    if idx >= len(questions):
        await ctx.bot.send_message(chat_id, "🎉 Quiz tugadi!\n🏆 /top — natijalarni ko'ring")
        return

    q = questions[idx]
    options = [str(o)[:100] for o in q["options"]]
    correct_idx = min(q["correct_index"], len(options) - 1)
    question_text = f"❓ {q['question']}"[:300]

    try:
        poll_msg = await ctx.bot.send_poll(
            chat_id=chat_id,
            question=question_text,
            options=options,
            type=Poll.QUIZ,
            correct_option_id=correct_idx,
            is_anonymous=False,
            open_period=30,
        )
        save_poll(poll_msg.poll.id, chat_id, correct_idx)
        ctx.chat_data["q_index"] = idx + 1

        ctx.job_queue.run_once(
            lambda c: c.application.create_task(send_next(chat_id, c)),
            35,
            name=f"q_{chat_id}_{idx}"
        )
    except Exception as e:
        logger.error(f"Poll xatosi: {e}")
        ctx.chat_data["q_index"] = idx + 1
        await send_next(chat_id, ctx)

async def handle_poll_answer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    answer = update.poll_answer
    row = get_poll(answer.poll_id)
    if not row or not answer.option_ids:
        return
    chat_id, correct_idx = row
    user = answer.user
    upsert_score(chat_id, user.id, user.username or "", user.first_name or "Nomsiz",
                 answer.option_ids[0] == correct_idx)

# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("format", cmd_format))
    app.add_handler(CommandHandler("top", cmd_top))
    app.add_handler(CommandHandler("reset", cmd_reset))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(PollAnswerHandler(handle_poll_answer))
    logger.info("Bot ishga tushdi ✅")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
