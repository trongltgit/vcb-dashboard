"""AI via Groq – free-tier safe models."""
import os
from typing import Optional

from groq import Groq

# Production free models (Oct 2026). Override with GROQ_MODEL env if needed.
DEFAULT_MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-20b")
FALLBACKS = [
    "openai/gpt-oss-20b",
    "openai/gpt-oss-120b",
    "llama-3.1-8b-instant",
    "llama-3.3-70b-versatile",
]

SYSTEM_PROMPT = """Bạn là chuyên gia phân tích ngân hàng tại Vietcombank (VCB), chuyên MBNT và TTQT/TTTM.
Viết báo cáo tiếng Việt: Tóm tắt điều hành → Phân tích → Tiềm năng → Khuyến nghị.
Dùng đúng số liệu được cung cấp, không bịa. Giọng trang trọng, nội bộ ngân hàng."""


def get_client(api_key: Optional[str] = None) -> Optional[Groq]:
    key = api_key or os.getenv("GROQ_API_KEY")
    if not key:
        return None
    return Groq(api_key=key)


def _call(client, model, messages, temperature=0.3, max_tokens=4096):
    return client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature,
        max_tokens=max_tokens,
    )


def _try_models(client, messages, temperature=0.3, max_tokens=4096):
    models = [DEFAULT_MODEL] + [m for m in FALLBACKS if m != DEFAULT_MODEL]
    last_err = None
    for m in models:
        try:
            chat = _call(client, m, messages, temperature, max_tokens)
            return chat.choices[0].message.content or "..."
        except Exception as e:
            last_err = e
            if "model_not_found" not in str(e).lower() and "does not exist" not in str(e).lower():
                return f"Loi Groq API: {e}"
            continue
    return f"Loi Groq API (het model): {last_err}"


def generate_professional_report(context: str, user_question: str = "", api_key: Optional[str] = None, model: str = None) -> str:
    client = get_client(api_key)
    if client is None:
        return "Chua set GROQ_API_KEY tren Render Environment."
    prompt = f"Du lieu:\n\n{context}\n\n"
    prompt += ("Yeu cau: " + user_question) if user_question else "Viet bao cao phan tich chuyen sau day du."
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    if model:
        try:
            chat = _call(client, model, messages)
            return chat.choices[0].message.content or "..."
        except Exception as e:
            return f"Loi: {e}"
    return _try_models(client, messages)


def chat_with_data(messages: list, context: str, api_key: Optional[str] = None, model: str = None) -> str:
    client = get_client(api_key)
    if client is None:
        return "Chua set GROQ_API_KEY tren Render Environment."
    system = SYSTEM_PROMPT + f"\n\n=== DATA ===\n{context}\n========"
    full = [{"role": "system", "content": system}] + messages
    if model:
        try:
            chat = _call(client, model, full, temperature=0.4, max_tokens=2048)
            return chat.choices[0].message.content or "..."
        except Exception as e:
            return f"Loi: {e}"
    return _try_models(client, full, temperature=0.4, max_tokens=2048)
