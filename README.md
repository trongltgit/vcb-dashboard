# VCB Dashboard – MBNT & TTQT/TTTM Analytics (Flask)

Web application thay thế hoàn toàn các Python formula trong Excel (Dashboard, CUSTOMER_RELATIONSHIP, KPI-*, PHAN_TICH_CIF…).

**Đã chuyển từ Streamlit → Flask** để deploy Render ổn định hơn, không phụ thuộc Streamlit.

## Tính năng chính

- **Upload** 2 file raw: `Data_Raw_MBNT` + `Data_Raw_TTQT_TTTM` (Excel hoặc CSV)
- **Bộ lọc** giống Excel: Phòng/Chi nhánh + khoảng ngày
- **Chuẩn hóa tên khách hàng** + group theo CIF (logic union-find giống Z1)
- **Loại khách cá nhân** tự động
- **KPI TOP 20% / BOT 10%** Doanh số & Lợi nhuận MBNT, TTQT, tách XK/NK
- **Biểu đồ** giữ nguyên style matplotlib gốc (bar ngang + scatter)
- **Báo cáo quan hệ TTQT ↔ MBNT** (tỷ lệ chuyển đổi, tiềm năng, nhóm nguyên nhân)
- **AI Bot (Groq)**: Sinh báo cáo chuyên sâu + chat với dữ liệu
- **Upload báo cáo external** (PDF/DOCX/TXT) để AI tham chiếu
- **Export** CSV / text

## Chạy local

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# (Tùy chọn)
echo "GROQ_API_KEY=gsk_xxx" > .env
echo "SECRET_KEY=doi-key-nay" >> .env

python app.py
# hoặc: gunicorn app:app --bind 0.0.0.0:5000
```

Mở http://localhost:5000

## Deploy lên Render

1. Push repo lên GitHub (thư mục gốc chứa `app.py`, `requirements.txt`, `utils/`, `templates/`).
2. Trên [Render](https://render.com) → **New → Web Service** (hoặc Blueprint với `render.yaml`).
3. Connect repo.
4. Settings:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120`
5. Environment Variables:
   - `GROQ_API_KEY` = key của bạn (hoặc nhập trên UI)
   - `SECRET_KEY` = chuỗi ngẫu nhiên (Render có thể generate)
6. Deploy.

Free tier Render: instance sleep khi không dùng; file upload lớn có thể cần vài chục giây xử lý.

## Cấu trúc project

```
vcb_dashboard_app/
├── app.py                    # Flask entry point
├── requirements.txt
├── Procfile
├── render.yaml
├── .env.example
├── templates/
│   ├── base.html
│   ├── index.html
│   └── dashboard.html
├── utils/
│   ├── processing.py         # Load, group, filter, KPI (port từ Z1/Z2/Z3)
│   ├── charts.py             # Matplotlib charts giữ style gốc
│   └── ai_bot.py             # Groq integration
└── README.md
```

## Lưu ý dữ liệu

- Cột MBNT cần có (tối thiểu): Ngày giao dịch, Ma_CIF, Ten_Khach_Hang, Phong_Ban, Chieu_NH, DS_Quy_USD, Loi_Nhuan_VND, Nguon_Mua_KH, Muc_Dich_KH.
- Cột TTQT: Ngay_GD, Ma_CIF, Ten_Khach_Hang, Phong_Ban, DS_Quy_USD (XK), (NK), (XNK).
- Logic “G có số → chỉ dùng G (XNK)” được giữ nguyên.

## License

Nội bộ sử dụng VCB / team phân tích.
