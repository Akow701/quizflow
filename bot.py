import os, re, sqlite3, logging, tempfile, uuid, json, asyncio
from pathlib import Path
from telegram import Update, Poll, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    PollAnswerHandler, CallbackQueryHandler,
    ContextTypes, filters
)
import pdfplumber, docx, openpyxl

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)
BOT_TOKEN = os.environ["BOT_TOKEN"]

# ══════════════════════════════════════════════════════════
#  DATABASE
# ══════════════════════════════════════════════════════════
def init_db():
    c = sqlite3.connect("quiz.db")
    c.executescript("""
        CREATE TABLE IF NOT EXISTS quizzes (
            quiz_id TEXT PRIMARY KEY, owner_id INTEGER,
            title TEXT, questions TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS sessions (
            session_id TEXT PRIMARY KEY, chat_id INTEGER,
            quiz_id TEXT, q_index INTEGER DEFAULT 0, active INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS scores (
            chat_id INTEGER, user_id INTEGER, username TEXT,
            first_name TEXT, correct INTEGER DEFAULT 0, total INTEGER DEFAULT 0,
            PRIMARY KEY (chat_id, user_id)
        );
        CREATE TABLE IF NOT EXISTS active_polls (
            poll_id TEXT PRIMARY KEY, chat_id INTEGER,
            correct_idx INTEGER, session_id TEXT
        );
    """)
    c.commit(); c.close()

def db():
    return sqlite3.connect("quiz.db")

def save_quiz(qid, owner, title, questions):
    c = db()
    c.execute("INSERT OR REPLACE INTO quizzes VALUES (?,?,?,?,CURRENT_TIMESTAMP)",
              (qid, owner, title, json.dumps(questions, ensure_ascii=False)))
    c.commit(); c.close()

def get_quiz(qid):
    c = db(); r = c.execute("SELECT quiz_id,owner_id,title,questions FROM quizzes WHERE quiz_id=?", (qid,)).fetchone(); c.close()
    return {"quiz_id":r[0],"owner_id":r[1],"title":r[2],"questions":json.loads(r[3])} if r else None

def get_user_quizzes(owner_id):
    c = db(); rows = c.execute("SELECT quiz_id,title,questions FROM quizzes WHERE owner_id=? ORDER BY created_at DESC LIMIT 20",(owner_id,)).fetchall(); c.close()
    return [{"quiz_id":r[0],"title":r[1],"count":len(json.loads(r[2]))} for r in rows]

def delete_quiz(qid):
    c = db(); c.execute("DELETE FROM quizzes WHERE quiz_id=?",(qid,)); c.commit(); c.close()

def create_session(sid, chat_id, qid):
    c = db(); c.execute("INSERT OR REPLACE INTO sessions VALUES (?,?,?,0,1)",(sid,chat_id,qid)); c.commit(); c.close()

def get_session(sid):
    c = db(); r = c.execute("SELECT session_id,chat_id,quiz_id,q_index,active FROM sessions WHERE session_id=?",(sid,)).fetchone(); c.close(); return r

def update_session_index(sid, idx):
    c = db(); c.execute("UPDATE sessions SET q_index=? WHERE session_id=?",(idx,sid)); c.commit(); c.close()

def close_session(sid):
    c = db(); c.execute("UPDATE sessions SET active=0 WHERE session_id=?",(sid,)); c.commit(); c.close()

def close_all_sessions(chat_id):
    c = db(); c.execute("UPDATE sessions SET active=0 WHERE chat_id=?",(chat_id,)); c.commit(); c.close()

def save_poll(poll_id, chat_id, correct_idx, sid):
    c = db(); c.execute("INSERT OR REPLACE INTO active_polls VALUES (?,?,?,?)",(poll_id,chat_id,correct_idx,sid)); c.commit(); c.close()

def get_poll(poll_id):
    c = db(); r = c.execute("SELECT chat_id,correct_idx,session_id FROM active_polls WHERE poll_id=?",(poll_id,)).fetchone(); c.close(); return r

def upsert_score(chat_id, user_id, username, first_name, is_correct):
    c = db()
    c.execute("""INSERT INTO scores(chat_id,user_id,username,first_name,correct,total) VALUES(?,?,?,?,?,1)
        ON CONFLICT(chat_id,user_id) DO UPDATE SET
        username=excluded.username, first_name=excluded.first_name,
        correct=correct+excluded.correct, total=total+1""",
        (chat_id,user_id,username,first_name,1 if is_correct else 0))
    c.commit(); c.close()

def get_leaderboard(chat_id, limit=10):
    c = db(); rows = c.execute("SELECT first_name,username,correct,total FROM scores WHERE chat_id=? ORDER BY correct DESC,total ASC LIMIT ?",(chat_id,limit)).fetchall(); c.close(); return rows

def get_quiz_stats(qid):
    c = db(); n = c.execute("SELECT COUNT(*) FROM sessions WHERE quiz_id=?",(qid,)).fetchone()[0]; c.close(); return n

# ══════════════════════════════════════════════════════════
#  FILE READERS
# ══════════════════════════════════════════════════════════
def read_txt(p):
    return open(p, encoding="utf-8", errors="ignore").read()

def read_pdf(p):
    t = ""
    with pdfplumber.open(p) as pdf:
        for page in pdf.pages:
            x = page.extract_text()
            if x: t += x + "\n"
    return t

def read_docx(p):
    return "\n".join(x.text for x in docx.Document(p).paragraphs)

def read_xlsx(p):
    wb = openpyxl.load_workbook(p, data_only=True)
    lines = []
    for sh in wb.worksheets:
        for row in sh.iter_rows(values_only=True):
            line = "\t".join(str(c) if c is not None else "" for c in row)
            if line.strip(): lines.append(line)
    return "\n".join(lines)

def extract_text(path):
    ext = Path(path).suffix.lower()
    if ext == ".pdf":  return read_pdf(path)
    if ext == ".docx": return read_docx(path)
    if ext == ".xlsx": return read_xlsx(path)
    return read_txt(path)

# ══════════════════════════════════════════════════════════
#  PARSER  (API siz — regex bilan)
# ══════════════════════════════════════════════════════════
AMAP = {"a":0,"b":1,"c":2,"d":3}

def parse_questions(text):
    qs = []
    for block in re.split(r"\n(?=\s*\d+[\.\)]\s)", text):
        block = block.strip()
        if not block: continue
        lines = [l.strip() for l in block.splitlines() if l.strip()]
        if not lines: continue
        question = re.sub(r"^\d+[\.\)]\s*", "", lines[0]).strip()
        if not question: continue
        options, correct_idx = [], None
        for line in lines[1:]:
            m = re.match(r"(?:to['\u2019]?g['\u2019]?ri|togri|javob|answer)\s*[:\-]\s*([a-dA-D])", line, re.I)
            if m: correct_idx = AMAP.get(m.group(1).lower()); continue
            m = re.match(r"^([a-dA-D])\s*[\.\)]\s*(.+)", line)
            if m:
                letter, opt = m.group(1).lower(), m.group(2).strip()
                if re.search(r"[*+=]", opt):
                    correct_idx = AMAP.get(letter)
                    opt = re.sub(r"\s*[*+=]\s*$", "", opt).strip()
                options.append(opt)
        if len(options) >= 2 and correct_idx is not None:
            qs.append({"question": question, "options": options,
                       "correct_index": min(correct_idx, len(options)-1)})
    return qs

# ══════════════════════════════════════════════════════════
#  QUIZ MENU
# ══════════════════════════════════════════════════════════
def quiz_menu(quiz, bot_username="QuizBot"):
    qid = quiz["quiz_id"]
    count = len(quiz["questions"])
    link = f"https://t.me/{bot_username}?start=q_{qid}"
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("▶️ Boshlash",      callback_data=f"start:{qid}"),
         InlineKeyboardButton("👥 Guruhda",        callback_data=f"group:{qid}")],
        [InlineKeyboardButton("📤 Ulashish",       callback_data=f"share:{qid}"),
         InlineKeyboardButton("📊 Statistika",     callback_data=f"stats:{qid}")],
        [InlineKeyboardButton("✏️ Nom",            callback_data=f"edittitle:{qid}"),
         InlineKeyboardButton("🗑 O'chirish",      callback_data=f"delete:{qid}")],
    ])
    text = (f"📋 *{quiz['title']}*\n"
            f"❓ {count} ta savol · ⏱ 30 soniya\n\n"
            f"🔗 `{link}`")
    return text, kb

# ══════════════════════════════════════════════════════════
#  QUIZ ENGINE
# ══════════════════════════════════════════════════════════
async def send_next_question(chat_id, session_id, app):
    session = get_session(session_id)
    if not session or not session[4]: return          # session yo'q yoki yopiq

    quiz = get_quiz(session[2])
    if not quiz: return

    q_index   = session[3]
    questions = quiz["questions"]

    # ── Quiz tugadi ──
    if q_index >= len(questions):
        close_session(session_id)
        rows = get_leaderboard(chat_id, 5)
        medals = ["🥇","🥈","🥉","4.","5."]
        lines = ["🎉 *Quiz tugadi!*\n\n🏆 *Top natijalar:*\n"]
        if rows:
            for i,(fn,un,cor,tot) in enumerate(rows):
                pct = round(cor/tot*100) if tot else 0
                name = f"@{un}" if un else fn
                lines.append(f"{medals[i]} {name} — {cor}/{tot} ({pct}%)")
        else:
            lines.append("Hech kim javob bermadi.")
        await app.bot.send_message(chat_id, "\n".join(lines), parse_mode="Markdown")
        return

    q = questions[q_index]
    options     = [str(o)[:100] for o in q["options"]]
    correct_idx = min(q["correct_index"], len(options)-1)
    q_text      = f"[{q_index+1}/{len(questions)}] {q['question']}"[:300]

    try:
        msg = await app.bot.send_poll(
            chat_id=chat_id,
            question=q_text,
            options=options,
            type=Poll.QUIZ,
            correct_option_id=correct_idx,
            is_anonymous=False,
            open_period=30,
        )
        save_poll(msg.poll.id, chat_id, correct_idx, session_id)
        update_session_index(session_id, q_index + 1)

        async def _next():
            await asyncio.sleep(35)
            await send_next_question(chat_id, session_id, app)

        asyncio.ensure_future(_next())

    except Exception as e:
        logger.error(f"Poll xatosi: {e}")
        update_session_index(session_id, q_index + 1)
        asyncio.ensure_future(send_next_question(chat_id, session_id, app))

# ══════════════════════════════════════════════════════════
#  HANDLERS
# ══════════════════════════════════════════════════════════
FORMAT_MSG = (
    "📋 *Fayl formati:*\n\n"
    "```\n"
    "1. Savol matni\n"
    "a) Birinchi\n"
    "b) Ikkinchi\n"
    "c) To'g'ri javob\n"
    "d) To'rtinchi\n"
    "to'g'ri: c\n\n"
    "2. Keyingi savol...\n"
    "```\n\n"
    "Yoki to'g'ri variant yoniga `*` qo'ying:\n"
    "`c) To'g'ri javob *`"
)

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ctx.args and ctx.args[0].startswith("q_"):
        qid = ctx.args[0][2:]
        quiz = get_quiz(qid)
        if not quiz:
            await update.message.reply_text("❌ Quiz topilmadi.")
            return
        me = await ctx.bot.get_me()
        text, kb = quiz_menu(quiz, me.username)
        await update.message.reply_text(text, reply_markup=kb, parse_mode="Markdown")
        return
    await update.message.reply_text(
        "👋 Salom! Men *QuizBot* man.\n\n"
        "📎 PDF · DOCX · TXT · XLSX fayl yuboring\n"
        "Men undan quiz yasab beraman!\n\n"
        "📚 /quizzes — quizlarim\n"
        "🏆 /top — leaderboard\n"
        "⛔️ /stop — quizni to'xtatish\n"
        "📌 /format — fayl formati",
        parse_mode="Markdown"
    )

async def cmd_format(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(FORMAT_MSG, parse_mode="Markdown")

async def cmd_quizzes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    quizzes = get_user_quizzes(update.effective_user.id)
    if not quizzes:
        await update.message.reply_text("📚 Quizlaringiz yo'q.\n\n📎 Fayl yuboring — quiz yasayman!")
        return
    kb = [[InlineKeyboardButton(f"📋 {q['title']} ({q['count']} savol)", callback_data=f"view:{q['quiz_id']}")] for q in quizzes]
    await update.message.reply_text("📚 *Quizlaringiz:*", reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

async def cmd_top(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rows = get_leaderboard(update.effective_chat.id)
    if not rows:
        await update.message.reply_text("📊 Hali hech kim quiz yechmagan.")
        return
    medals = ["🥇","🥈","🥉"]
    lines = ["🏆 *Leaderboard*\n"]
    for i,(fn,un,cor,tot) in enumerate(rows):
        medal = medals[i] if i < 3 else f"{i+1}."
        pct = round(cor/tot*100) if tot else 0
        name = f"@{un}" if un else fn
        lines.append(f"{medal} {name} — {cor}/{tot} ({pct}%)")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_stop(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    close_all_sessions(update.effective_chat.id)
    await update.message.reply_text("⛔️ Quiz to'xtatildi.")

async def cmd_reset(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    user = update.effective_user
    if chat.type in ("group","supergroup"):
        admins = await ctx.bot.get_chat_administrators(chat.id)
        if user.id not in [a.user.id for a in admins]:
            await update.message.reply_text("❌ Faqat adminlar reset qila oladi.")
            return
    c = db(); c.execute("DELETE FROM scores WHERE chat_id=?",(chat.id,)); c.commit(); c.close()
    await update.message.reply_text("✅ Leaderboard tozalandi.")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    ext = Path(doc.file_name or "").suffix.lower()
    if ext not in (".pdf",".docx",".txt",".xlsx"):
        await update.message.reply_text("❌ Faqat PDF, DOCX, TXT yoki XLSX.\n📌 /format")
        return

    msg = await update.message.reply_text("⏳ Fayl o'qilmoqda...")
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp_path = tmp.name
    try:
        await (await doc.get_file()).download_to_drive(tmp_path)
        raw = extract_text(tmp_path)
    except Exception as e:
        await msg.edit_text(f"❌ O'qishda xato: {e}"); return
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    if not raw.strip():
        await msg.edit_text("❌ Fayl bo'sh yoki o'qib bo'lmadi."); return

    questions = parse_questions(raw)
    if not questions:
        await msg.edit_text("❌ Savollar topilmadi.\n📌 /format — fayl formatini ko'ring"); return

    qid   = uuid.uuid4().hex[:8].upper()
    title = Path(doc.file_name).stem[:40]
    save_quiz(qid, update.effective_user.id, title, questions)

    me = await ctx.bot.get_me()
    text, kb = quiz_menu(get_quiz(qid), me.username)
    await msg.edit_text(f"✅ *{len(questions)} ta savol topildi!*\n\n{text}", reply_markup=kb, parse_mode="Markdown")

async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q    = update.callback_query
    data = q.data
    chat_id = q.message.chat_id

    async def safe_answer(text="", alert=False):
        try: await q.answer(text, show_alert=alert)
        except: pass

    async def safe_edit(text, kb=None):
        try: await q.edit_message_text(text, reply_markup=kb, parse_mode="Markdown")
        except: pass

    try:
        if data.startswith("view:"):
            qid  = data[5:]
            quiz = get_quiz(qid)
            if not quiz: await safe_answer("❌ Topilmadi", True); return
            me = await ctx.bot.get_me()
            text, kb = quiz_menu(quiz, me.username)
            await safe_edit(text, kb)
            await safe_answer()

        elif data.startswith("start:"):
            qid  = data[6:]
            quiz = get_quiz(qid)
            if not quiz: await safe_answer("❌ Topilmadi", True); return
            close_all_sessions(chat_id)
            sid = uuid.uuid4().hex[:8].upper()
            create_session(sid, chat_id, qid)
            await safe_edit(
                f"▶️ *{quiz['title']}* boshlanmoqda!\n"
                f"❓ {len(quiz['questions'])} ta savol · ⏱ 30 soniya"
            )
            await safe_answer()
            asyncio.ensure_future(send_next_question(chat_id, sid, ctx.application))

        elif data.startswith("group:"):
            qid = data[6:]
            me  = await ctx.bot.get_me()
            link = f"https://t.me/{me.username}?start=q_{qid}"
            await safe_answer(f"Havolani guruhga yuboring:\n{link}", True)

        elif data.startswith("share:"):
            qid = data[6:]
            me  = await ctx.bot.get_me()
            link = f"https://t.me/{me.username}?start=q_{qid}"
            await safe_answer(f"🔗 Havola:\n{link}", True)

        elif data.startswith("stats:"):
            qid  = data[6:]
            quiz = get_quiz(qid)
            if not quiz: await safe_answer("❌ Topilmadi", True); return
            n = get_quiz_stats(qid)
            await safe_answer(f"📊 {quiz['title']}\n❓ {len(quiz['questions'])} savol\n▶️ {n} marta o'ynaldi", True)

        elif data.startswith("delete:"):
            qid = data[7:]
            kb  = InlineKeyboardMarkup([[
                InlineKeyboardButton("✅ Ha", callback_data=f"confirmdelete:{qid}"),
                InlineKeyboardButton("❌ Yo'q", callback_data=f"view:{qid}")
            ]])
            await safe_edit("🗑 Rostdan ham o'chirmoqchimisiz?", kb)
            await safe_answer()

        elif data.startswith("confirmdelete:"):
            qid = data[14:]
            delete_quiz(qid)
            await safe_edit("✅ Quiz o'chirildi.")
            await safe_answer()

        elif data.startswith("edittitle:"):
            qid = data[10:]
            ctx.user_data["editing_title"] = qid
            kb  = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Bekor", callback_data=f"view:{qid}")]])
            await safe_edit("✏️ Yangi nom yozing:", kb)
            await safe_answer()

        else:
            await safe_answer()

    except Exception as e:
        logger.error(f"Callback xatosi [{data}]: {e}")
        await safe_answer("❌ Xato yuz berdi, qayta urinib ko'ring.", True)

async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if "editing_title" not in ctx.user_data: return
    qid = ctx.user_data.pop("editing_title")
    new_title = update.message.text.strip()[:40]
    c = db(); c.execute("UPDATE quizzes SET title=? WHERE quiz_id=?",(new_title,qid)); c.commit(); c.close()
    quiz = get_quiz(qid)
    if not quiz: await update.message.reply_text("❌ Quiz topilmadi."); return
    me = await ctx.bot.get_me()
    text, kb = quiz_menu(quiz, me.username)
    await update.message.reply_text(f"✅ Nom o'zgartirildi!\n\n{text}", reply_markup=kb, parse_mode="Markdown")

async def handle_poll_answer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ans = update.poll_answer
    row = get_poll(ans.poll_id)
    if not row or not ans.option_ids: return
    chat_id, correct_idx, _ = row
    u = ans.user
    upsert_score(chat_id, u.id, u.username or "", u.first_name or "Nomsiz", ans.option_ids[0] == correct_idx)

# ══════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════
def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("format",  cmd_format))
    app.add_handler(CommandHandler("quizzes", cmd_quizzes))
    app.add_handler(CommandHandler("top",     cmd_top))
    app.add_handler(CommandHandler("stop",    cmd_stop))
    app.add_handler(CommandHandler("reset",   cmd_reset))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(PollAnswerHandler(handle_poll_answer))
    logger.info("Bot ishga tushdi ✅")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
