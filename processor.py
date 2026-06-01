import os
import json
import pandas as pd
import mammoth
import google.generativeai as genai
from PyPDF2 import PdfReader
from config import GEMINI_API_KEY

genai.configure(api_key=GEMINI_API_KEY)

SYSTEM_PROMPT = """
Sen fayldan test savollarini ajratuvchi robotsan.
Savollarni top va JSON formatida qaytar:
[
  {"question":"...","options":["A","B","C","D"],"correct":0}
]
Agar javob aniqlanmasa correct indeksini 0 qo‘y.
Hech qanday izoh yozma, faqat JSON qaytar.
"""

model = genai.GenerativeModel("gemini-pro")

def extract_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".txt":
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read()
        elif ext == ".docx":
            with open(file_path, "rb") as docx_file:
                result = mammoth.extract_raw_text(docx_file)
                return result.value
        elif ext == ".pdf":
            reader = PdfReader(file_path)
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        elif ext in [".xlsx", ".csv"]:
            df = pd.read_excel(file_path) if ext == ".xlsx" else pd.read_csv(file_path)
            return df.to_string()
    except Exception as e:
        print(f"Extract error: {e}")
    return None

def generate_quiz(text):
    try:
        response = model.generate_content(f"{SYSTEM_PROMPT}\n\n{text}")
        return json.loads(response.text.strip())
    except Exception as e:
        print(f"Gemini error: {e}")
        return []
