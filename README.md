# 🤖 QuizBot — Telegram Quiz Bot

PDF, DOCX, TXT, XLSX fayllardan avtomatik quiz yaratuvchi Telegram bot.

---

## ✨ Imkoniyatlar

- 📎 **Fayl qabul qilish**: PDF, DOCX, TXT, XLSX
- 🤖 **AI quiz yaratish**: Claude orqali (API key kerak)
- 👥 **Guruh + lichka**: har birida alohida leaderboard
- 🏆 **Leaderboard**: /top buyrug'i bilan ko'rish
- ⏱️ **30 soniya timer**: har bir savol uchun
- ✅ **To'g'ri javob belgilash**: `* = + /` kabi belgilarni Claude o'zi tushunadi

---

## 🚀 Railway orqali ishga tushirish

### 1. Bot yaratish
1. Telegramda [@BotFather](https://t.me/BotFather) ga boring
2. `/newbot` yozing
3. Nom va username bering
4. **Token**ni saqlang

### 2. Claude API key olish
1. [console.anthropic.com](https://console.anthropic.com) ga kiring
2. **API Keys** → **Create Key**
3. Keyni saqlang

### 3. Railway deploy
1. [railway.app](https://railway.app) ga kiring
2. **New Project** → **Deploy from GitHub repo**
3. Bu papkani GitHub'ga push qiling
4. Railway'da **Variables** bo'limiga qo'shing:
   ```
   BOT_TOKEN = your_bot_token
   CLAUDE_API_KEY = your_claude_api_key
   ```
5. **Deploy** tugmasini bosing ✅

---

## 📱 Bot buyruqlari

| Buyruq | Vazifasi |
|--------|---------|
| `/start` | Botni ishga tushirish |
| `/help` | Yordam |
| `/top` | Leaderboard ko'rish |
| `/reset` | Leaderboardni tozalash (faqat admin) |

---

## 📂 Fayl formatlari

| Format | Kutubxona |
|--------|-----------|
| `.pdf` | pdfplumber |
| `.docx` | python-docx |
| `.txt` | built-in |
| `.xlsx` | openpyxl |

---

## 🗂️ Loyiha tuzilmasi

```
quizbot/
├── bot.py           ← Asosiy bot kodi
├── requirements.txt ← Python kutubxonalar
├── Procfile         ← Railway uchun
├── .env.example     ← Environment variables namunasi
└── README.md
```

---

## ❓ Ko'p uchraydigan muammolar

**Bot ishlamayapti?**
→ `BOT_TOKEN` va `CLAUDE_API_KEY` to'g'ri kiritilganini tekshiring

**Fayl o'qilmayapti?**
→ Fayl 20MB dan kichik bo'lishi kerak (Telegram cheklovi)

**Savollar chiqmayapti?**
→ Faylingizda savol-javob formatidagi matn borligiga ishonch hosil qiling
