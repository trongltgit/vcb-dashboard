"""
AI Report Generator using Groq API.
"""
import os
from typing import Optional

from groq import Groq

DEFAULT_MODEL = os.environ.get("GROQ_MODEL", "llama-3.1-8b-instant")

SYSTEM_PROMPT = """Bạn là chuyên gia phân tích ngân hàng tại Vietcombank (VCB), chuyên về doanh số ngoại tệ (MBNT) và thanh toán quốc tế (TTQT/TTTM).
Nhiệm vụ: Viết báo cáo phân tích chuyên sâu, súc tích, chuyên nghiệp bằng tiếng Việt.
- Dùng số liệu chính xác từ dữ liệu được cung cấp.
- Đưa ra nhận định, khuyến nghị hành động cụ thể cho phòng/ban.
- Cấu trúc rõ ràng: Tóm tắt điều hành → Phân tích chi tiết → Tiềm năng → Khuyến nghị.
- Giọng văn trang trọng, khách quan, phù hợp báo cáo nội bộ ngân hàng.
- Không bịa số liệu. Nếu thiếu dữ liệu hãy nêu rõ.
"""


def get_client(api_key: Optional[str] = None) -> Optional[Groq]:
    key = api_key or os.getenv("GROQ_API_KEY")
    if not key:
        return None
    return Groq(api_key=key)


def generate_professional_report(
    context: str,
    user_question: str = "",
    api_key: Optional[str] = None,
    model: str = None,
) -> str:
    client = get_client(api_key)
    if client is None:
        return "Chua cau hinh GROQ_API_KEY tren server (Render env)."

    model = model or DEFAULT_MODEL
    prompt = f"""Du lieu phan tich hien tai:

{context}

{"Cau hoi / yeu cau bo sung: " + user_question if user_question else "Hay viet bao cao phan tich chuyen sau day du dua tren du lieu tren."}
"""
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    try:
        chat = client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=0.3,
            max_tokens=4096,
        )
        return chat.choices[0].message.content or "Khong nhan duoc phan hoi tu AI."
    except Exception as e:
        if "model_not_found" in str(e).lower() and model != "llama-3.1-8b-instant":
            try:
                chat = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=messages,
                    temperature=0.3,
                    max_tokens=4096,
                )
                return chat.choices[0].message.content or "..."
            except Exception as e2:
                return f"Loi goi Groq API: {e2}"
        return f"Loi goi Groq API: {e}"


def chat_with_data(
    messages: list,
    context: str,
    api_key: Optional[str] = None,
    model: str = None,
) -> str:
    client = get_client(api_key)
    if client is None:
        return "Chua cau hinh GROQ_API_KEY tren server (Render env)."

    model = model or DEFAULT_MODEL
    system = SYSTEM_PROMPT + f"\n\n=== DU LIEU HIEN TAI ===\n{context}\n========================"
    full_messages = [{"role": "system", "content": system}] + messages

    try:
        chat = client.chat.completions.create(
            model=model,
            messages=full_messages,
            temperature=0.4,
            max_tokens=2048,
        )
        return chat.choices[0].message.content or "..."
    except Exception as e:
        if "model_not_found" in str(e).lower() and model != "llama-3.1-8b-instant":
            try:
                chat = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=full_messages,
                    temperature=0.4,
                    max_tokens=2048,
                )
                return chat.choices[0].message.content or "..."
            except Exception as e2:
                return f"Loi: {e2}"
        return f"Loi: {e}"
