"""
AI Report Generator using Groq API.
Enhances professional analysis reports with LLM insights.
"""
import os
from typing import Optional

from groq import Groq


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
    model: str = "llama-3.3-70b-versatile",
) -> str:
    """
    Generate an enhanced professional report from analysis context.
    """
    client = get_client(api_key)
    if client is None:
        return "⚠️ Chưa cấu hình GROQ_API_KEY. Vui lòng nhập key ở sidebar hoặc set biến môi trường."

    prompt = f"""Dữ liệu phân tích hiện tại:

{context}

{"Câu hỏi / yêu cầu bổ sung của người dùng: " + user_question if user_question else "Hãy viết báo cáo phân tích chuyên sâu đầy đủ dựa trên dữ liệu trên."}
"""

    try:
        chat = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,
            max_tokens=4096,
        )
        return chat.choices[0].message.content or "Không nhận được phản hồi từ AI."
    except Exception as e:
        return f"❌ Lỗi gọi Groq API: {str(e)}"


def chat_with_data(
    messages: list,
    context: str,
    api_key: Optional[str] = None,
    model: str = "llama-3.3-70b-versatile",
) -> str:
    """Multi-turn chat about the current filtered data."""
    client = get_client(api_key)
    if client is None:
        return "⚠️ Chưa cấu hình GROQ_API_KEY."

    system = SYSTEM_PROMPT + f"\n\n=== DỮ LIỆU HIỆN TẠI ===\n{context}\n========================"

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
        return f"❌ Lỗi: {str(e)}"
