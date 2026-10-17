"""AI via Groq – free developer plan (openai/gpt-oss-20b)."""
import os
from typing import Optional

from groq import Groq

DEFAULT_MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-20b")

SYSTEM_PROMPT = (
    "Ban la chuyen gia phan tich ngan hang VCB (MBNT, TTQT/TTTM). "
    "Viet bao cao tieng Viet: Tom tat -> Phan tich -> Tiem nang -> Khuyen nghi. "
    "Dung dung so lieu, khong bia."
)


def get_client(api_key: Optional[str] = None):
    key = api_key or os.getenv("GROQ_API_KEY")
    return Groq(api_key=key) if key else None


def generate_professional_report(
    context: str,
    user_question: str = "",
    api_key: Optional[str] = None,
    model: str = None,
) -> str:
    client = get_client(api_key)
    if not client:
        return "Chua set GROQ_API_KEY tren Render Environment."
    model = model or DEFAULT_MODEL
    prompt = f"Du lieu:\n{context}\n\n"
    prompt += ("Yeu cau: " + user_question) if user_question else "Viet bao cao phan tich day du."
    try:
        r = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,
            max_tokens=3000,
        )
        return r.choices[0].message.content or "..."
    except Exception as e:
        return f"Loi Groq ({model}): {e}"


def chat_with_data(
    messages: list,
    context: str,
    api_key: Optional[str] = None,
    model: str = None,
) -> str:
    client = get_client(api_key)
    if not client:
        return "Chua set GROQ_API_KEY tren Render Environment."
    model = model or DEFAULT_MODEL
    full = [
        {"role": "system", "content": SYSTEM_PROMPT + "\n\nDATA:\n" + context}
    ] + messages
    try:
        r = client.chat.completions.create(
            model=model,
            messages=full,
            temperature=0.4,
            max_tokens=1500,
        )
        return r.choices[0].message.content or "..."
    except Exception as e:
        return f"Loi Groq ({model}): {e}"
