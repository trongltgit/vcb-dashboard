"""
VCB Dashboard – Flask (Render free-tier safe)
- No Groq key on frontend (uses GROQ_API_KEY env only)
- Plan-completion 3x6 pie gauges (exact Excel KE_HOACH logic)
- PDF export (Unicode-safe)
- Optional bar/scatter charts via ENABLE_CHARTS=1
"""
import base64
import gc
import io
import os
import re
import unicodedata as ud
from datetime import date, datetime

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.utils import secure_filename

from utils.processing import (
    prepare_all,
    build_customer_df,
    filt,
    build,
    kpi_pair,
)
from utils.ai_bot import generate_professional_report, chat_with_data

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "vcb-dashboard-dev-key-change-me")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80 MB

# Set ENABLE_CHARTS=1 on Render if you have enough RAM for bar/scatter charts
ENABLE_CHARTS = os.environ.get("ENABLE_CHARTS", "0") == "1"
PLAN_YEAR = int(os.environ.get("PLAN_YEAR", "2026"))

_STORE = {}


def _sid():
    if "sid" not in session:
        session["sid"] = os.urandom(8).hex()
    return session["sid"]


# ───────────────────────── IO helpers ─────────────────────────

def read_df(file_storage):
    if file_storage is None or not file_storage.filename:
        return None

    filename = secure_filename(file_storage.filename) or "uploaded_file"
    name = filename.lower()
    stream = file_storage.stream
    stream.seek(0)

    try:
        if name.endswith(".csv"):
            try:
                df = pd.read_csv(stream, encoding="utf-8-sig", low_memory=False)
            except UnicodeDecodeError:
                stream.seek(0)
                df = pd.read_csv(stream, encoding="cp1258", low_memory=False)
        elif name.endswith((".xlsx", ".xlsm")):
            df = pd.read_excel(stream, engine="openpyxl")
        elif name.endswith(".xls"):
            df = pd.read_excel(stream, engine="xlrd")
        else:
            raise ValueError(
                f"Dinh dang file khong ho tro: {filename}. "
                "Vui long dung CSV, XLS, XLSX hoac XLSM."
            )

        if df.empty and len(df.columns) == 0:
            raise ValueError(f"File {filename} khong co bang du lieu.")

        df.columns = [str(col).strip() for col in df.columns]
        return df
    except Exception as exc:
        raise ValueError(f"Khong doc duoc file '{filename}': {exc}") from exc


def read_external_text(file_storage) -> str:
    if file_storage is None or file_storage.filename == "":
        return ""
    name = file_storage.filename.lower()
    data = file_storage.read()
    if name.endswith((".txt", ".md")):
        return data.decode("utf-8", errors="ignore")
    if name.endswith(".pdf"):
        try:
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(data))
            return "\n".join(p.extract_text() or "" for p in reader.pages)
        except Exception:
            return "[Khong doc duoc PDF]"
    if name.endswith(".docx"):
        try:
            from docx import Document
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            return "[Khong doc duoc DOCX]"
    return ""


def usd(v):
    a = abs(v)
    if a >= 1e9:
        return f"{a/1e9:,.2f} ty USD"
    if a >= 1e6:
        return f"{a/1e6:,.1f} trieu USD"
    if a >= 1e3:
        return f"{a/1e3:,.0f} nghin USD"
    return f"{a:,.0f} USD"


def pc(a, b):
    return f"{a/b*100:.0f}%" if b > 0 else "0%"


def safe_html(df, max_rows: int = 30) -> str:
    if df is None or df.empty:
        return ""
    try:
        d = df.drop(columns=["Key"], errors="ignore").head(max_rows).copy()
        d = d.astype(str)
        return d.to_html(classes="table table-sm", index=False, escape=True, border=0)
    except Exception as e:
        return f"<p class='text-danger'>Khong render duoc bang: {e}</p>"


# ───────────────────────── Plan (exact Excel KE_HOACH) ─────────────────────────

def build_plan_grid(plan_raw: pd.DataFrame, phong: str, d1, d2, year: int = PLAN_YEAR):
    """
    Exact port of Excel Python formula for sheet KE_HOACH!A1:O230.
    Returns: grid (3 metrics x 6 cells of (th,kh,pct)), metrics, q_label, y_label
    or (None, None, None, None) if plan missing.
    """
    if plan_raw is None or plan_raw.empty:
        return None, None, None, None

    def norm(s):
        return " ".join(str(s).replace("\xa0", " ").split()).lower()

    def code_of(s):
        s = ud.normalize("NFD", str(s).replace("đ", "d").replace("Đ", "D"))
        s = "".join(ch for ch in s if ud.category(ch) != "Mn")
        return "".join(w if w.isdigit() else w[0] for w in re.findall(r"[A-Za-z]+|\d+", s)).lower()

    def num(x):
        try:
            v = float(x)
            return 0.0 if pd.isna(v) else v
        except Exception:
            return 0.0

    plan = plan_raw.copy()
    while plan.shape[1] < 15:
        plan[plan.shape[1]] = 0
    if len(plan) < 186:
        extra = pd.DataFrame([[None] * plan.shape[1]] * (186 - len(plan)))
        plan = pd.concat([plan, extra], ignore_index=True)

    def V(r, col):
        try:
            return num(plan.iloc[r - 1, col])
        except Exception:
            return 0.0

    rooms_n = [norm(plan.iloc[6 - 1 + 3 * i, 0]) for i in range(15)]

    d1 = pd.Timestamp(d1)
    d2 = pd.Timestamp(d2)
    m1, m2 = d1.month, d2.month
    if d1.year < year:
        m1 = 1
    if d2.year > year:
        m2 = 12
    QN = (d2.month - 1) // 3 + 1
    Q_COL = 1 + d2.month
    Y_COL = 14
    Q_LABEL = f"Q{QN}/{d2.year}"
    Y_LABEL = f"{d2.year}"

    METRICS = [
        dict(name="DS MBNT", unit="tr.USD"),
        dict(name="LN MBNT", unit="ty VND"),
        dict(name="DS TTQT-TTTM", unit="tr.USD"),
    ]
    ROWS = [
        dict(k0=6, cn=54, tsc=60, tot=63),       # DS MBNT
        dict(k0=69, cn=117, tsc=123, tot=126),   # LN MBNT
        dict(k0=132, cn=177, tsc=180, tot=183),  # DS TTQT-TTTM
    ]

    def is_cum(k):
        R = ROWS[k]
        for r in (R["cn"] + 1, R["tot"] + 1):
            xs = [V(r, 1 + m) for m in range(1, m2 + 1)]
            if m2 >= 3 and xs[-1] > 0:
                up = sum(1 for i in range(1, m2) if xs[i] > xs[i - 1])
                return (
                    all(xs[i] >= xs[i - 1] - 1e-9 for i in range(1, m2))
                    and up >= max(2, int(0.6 * (m2 - 1)))
                )
        return False

    CUM = [is_cum(k) for k in range(3)]

    def th_cum(k, r_th):
        if CUM[k]:
            return V(r_th, 1 + m2) - (V(r_th, m1) if m1 > 1 else 0.0)
        return float(sum(V(r_th, 1 + m) for m in range(m1, m2 + 1)))

    key = norm(phong or "")
    codes = [code_of(x) for x in rooms_n]
    if key and key in rooms_n:
        i_room = rooms_n.index(key)
    elif key and key.replace(" ", "") in codes:
        i_room = codes.index(key.replace(" ", ""))
    else:
        i_room = None
    phong_mode = i_room is not None

    def cells(k, kh_col):
        R = ROWS[k]
        r_phong = (R["k0"] + 3 * i_room) if phong_mode else R["tot"]
        out = []
        for r_kh in (r_phong, R["cn"], R["tsc"]):
            th = th_cum(k, r_kh + 1)
            kh = V(r_kh, kh_col)
            out.append((th, kh, th / kh if kh > 0 else 0.0))
        return out

    grid = [cells(k, Q_COL) + cells(k, Y_COL) for k in range(3)]
    return grid, METRICS, Q_LABEL, Y_LABEL


def render_plan_chart(plan_raw, phong, d_from, d_to):
    """Returns (base64_png, summary_lines)."""
    grid, metrics, q_label, y_label = build_plan_grid(plan_raw, phong, d_from, d_to)
    if grid is None:
        return "", []

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from utils.charts import chart_plan_completion

        fig = chart_plan_completion(
            grid,
            metrics,
            loc=phong or "VCB HCM",
            d1=pd.Timestamp(d_from),
            d2=pd.Timestamp(d_to),
            q_label=q_label,
            y_label=y_label,
            bg="#14532D",
        )
        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=100, facecolor="#14532D")
        plt.close(fig)
        buf.seek(0)
        b64 = base64.b64encode(buf.read()).decode("ascii")

        names = ["KH phong", "KH CN giao", "KH TSC giao"]
        summary = []
        for mi, m in enumerate(metrics):
            for ci, cname in enumerate(names):
                th, kh, pct = grid[mi][ci]
                summary.append(
                    f"{m['name']} | {cname} Q: TH={th:,.1f} KH={kh:,.1f} ({pct*100:.1f}%)"
                )
            for ci, cname in enumerate(names):
                th, kh, pct = grid[mi][ci + 3]
                summary.append(
                    f"{m['name']} | {cname} Y: TH={th:,.1f} KH={kh:,.1f} ({pct*100:.1f}%)"
                )
        return b64, summary
    except Exception as e:
        print(f"Plan chart failed: {e}")
        return "", []


# ───────────────────────── Optional bar/scatter charts ─────────────────────────

def generate_charts_safe(TDS, BDS, TLN, BLN, TQ, BQ, q_filt, df):
    charts = {
        "mbnt_top_ds": "", "mbnt_bot_ds": "", "mbnt_top_ln": "", "mbnt_bot_ln": "",
        "ttqt_top": "", "ttqt_bot": "", "ttqt_xk_top": "", "ttqt_xk_bot": "",
        "ttqt_nk_top": "", "ttqt_nk_bot": "", "scatter": "", "plan_pie": "",
    }
    if not ENABLE_CHARTS:
        return charts

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from utils.charts import chart_bar, scatter_ttqt_mbnt

        def fig_to_b64(fig):
            buf = io.BytesIO()
            fig.savefig(buf, format="png", bbox_inches="tight", dpi=72)
            plt.close(fig)
            buf.seek(0)
            return base64.b64encode(buf.read()).decode("ascii")

        charts["mbnt_top_ds"] = fig_to_b64(chart_bar(TDS, "TOP DS MBNT", "USD", "#0EA5E9", top=True))
        charts["mbnt_bot_ds"] = fig_to_b64(chart_bar(BDS, "BOT DS MBNT", "USD", "#F97316", top=False))
        charts["mbnt_top_ln"] = fig_to_b64(chart_bar(TLN, "TOP LN MBNT", "VND", "#22C55E", top=True))
        charts["mbnt_bot_ln"] = fig_to_b64(chart_bar(BLN, "BOT LN MBNT", "VND", "#EF4444", top=False))
        charts["ttqt_top"] = fig_to_b64(chart_bar(TQ, "TOP DS TTQT", "USD", "#8B5CF6", top=True))
        charts["ttqt_bot"] = fig_to_b64(chart_bar(BQ, "BOT DS TTQT", "USD", "#EC4899", top=False))

        cq_xk = build(q_filt, ["xk"]).rename(columns={"XK": "DS"})
        cq_nk = build(q_filt, ["nk"]).rename(columns={"NK": "DS"})
        TXK, BXK = kpi_pair(cq_xk, "DS")
        TNK, BNK = kpi_pair(cq_nk, "DS")
        charts["ttqt_xk_top"] = fig_to_b64(chart_bar(TXK, "TOP XK", "USD", "#06B6D4", top=True))
        charts["ttqt_xk_bot"] = fig_to_b64(chart_bar(BXK, "BOT XK", "USD", "#F59E0B", top=False))
        charts["ttqt_nk_top"] = fig_to_b64(chart_bar(TNK, "TOP NK", "USD", "#10B981", top=True))
        charts["ttqt_nk_bot"] = fig_to_b64(chart_bar(BNK, "BOT NK", "USD", "#F43F5E", top=False))
        charts["scatter"] = fig_to_b64(scatter_ttqt_mbnt(df))
        del TXK, BXK, TNK, BNK, cq_xk, cq_nk
        gc.collect()
    except Exception as e:
        print(f"Chart generation skipped/failed: {e}")
    return charts


# ───────────────────────── Report text ─────────────────────────

def build_report(df, dall, qall, mm, tt, phong, d_from, d_to):
    def phan_tich(X, O, Y, Yv, tx_, ty_):
        L = []
        d = df[df[X] > 0].copy()
        if d.empty:
            return [f"Khong co cong ty nao co {tx_} trong ky da chon."], {}
        TX, TY, TV = d[X].sum(), d[Y].sum(), d[Yv].sum()
        r = d[Y] / d[X]
        L.append(
            f"Co {len(d):,} cong ty co {tx_} voi tong {usd(TX)}. "
            f"Trong so do {int((d[Y]>0).sum()):,} cong ty co giao dich {ty_} voi VCB, tong {usd(TY)}."
        )
        L.append(f"Ty le chuyen doi chung: cu 100 USD {tx_} thi VCB nhan duoc {TY/TX*100:.0f} USD {ty_}.")
        lab = ["Khong giao dich", "Duoi 30%", "Tu 30% den 80%", "Tu 80% den 120%", "Tren 120%"]
        d["Nhom"] = np.select([d[Y] == 0, r < 0.3, r < 0.8, r <= 1.2], lab[:4], default=lab[4])
        L.append(f"Phan nhom theo ty le {ty_}/{tx_}:")
        for k in lab:
            s = d[d.Nhom == k]
            if len(s):
                L.append(f"   - {k}: {len(s):,} cong ty, chiem {pc(s[X].sum(), TX)} tong {tx_}.")
        lo = d[d.Nhom.isin(lab[:2])]
        L.append(
            f"Nhom duoi 30% (gom ca khong giao dich) co {len(lo):,} cong ty, "
            f"nam {pc(lo[X].sum(), TX)} tong {tx_} nhung chi dong gop {pc(lo[Y].sum(), TY)} tong {ty_}."
        )
        L.append(f"Giai ngan von vay ngoai te chiem {pc(TV, TY)} tong {ty_} ({usd(TV)}).")
        d["Net"] = (d[X] - d[O]).clip(lower=0)
        d["Tn"] = (d.Net - d[Y]).clip(lower=0)
        low = d[d[Y] < 0.5 * d[X]].copy()
        low["Nn"] = np.select([low[O] >= 0.7 * low[X], low[Y] == 0], ["A", "B"], default="C")
        tn = low[low.Nn.isin(["B", "C"])].Tn.sum()
        L.append(f"Tiem nang khai thac thuc (nhom B+C): khoang {usd(tn)}.")
        return L, {"conv": TY / TX if TX else 0, "tn": tn}

    LX, sx = phan_tich("XK", "NK", "BUY", "BUY_VAY", "XK", "BUY")
    LN, sn = phan_tich("NK", "XK", "SELL", "SELL_VAY", "NK", "SELL")

    R = [
        "BAO CAO PHAN TICH QUAN HE DS TTQT - DS MBNT",
        f"Phong/don vi: {phong or 'VCB HCM'}. Ky: {d_from:%d/%m/%Y} den {d_to:%d/%m/%Y}.",
        f"Du lieu: {len(dall):,} dong MBNT raw -> {len(mm):,} dung phan tich; "
        f"{len(qall):,} dong TTQT -> {len(tt):,} dung. {len(df):,} doanh nghiep (da loai ca nhan).",
        "",
        "PHAN 1. XUAT KHAU (XK) VA BUY",
    ] + LX + ["", "PHAN 2. NHAP KHAU (NK) VA SELL"] + LN + ["", "PHAN 3. KET LUAN"]

    if sx and sn:
        cx, cnv = sx["conv"] * 100, sn["conv"] * 100
        R.append(f"Ty le chuyen doi XK->BUY: {cx:.0f}%, NK->SELL: {cnv:.0f}%.")
        R.append(
            "Cong ty giu chan VCB tot hon o chieu mua ngoai te (thanh toan NK)."
            if cnv > cx
            else "Cong ty giu chan VCB tot hon o chieu ban ngoai te."
        )
        ln_ = "XK" if sx["tn"] >= sn["tn"] else "NK"
        R.append(f"Tiem nang khai thac lon hon nam o phia {ln_}.")

    return "\n".join(R)


# ───────────────────────── PDF ─────────────────────────

def _pdf_safe(s: str) -> str:
    if not s:
        return ""
    repl = {
        "–": "-", "—": "-", "−": "-",
        "“": '"', "”": '"', "‘": "'", "’": "'",
        "…": "...", "•": "-", "·": "-",
        "₫": "VND", "→": "->", "←": "<-",
        "≥": ">=", "≤": "<=", "×": "x",
        "đ": "d", "Đ": "D",
    }
    for a, b in repl.items():
        s = s.replace(a, b)
    return s.encode("latin-1", errors="replace").decode("latin-1")


def build_pdf(data: dict) -> bytes:
    from fpdf import FPDF

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 10)
            self.set_text_color(30, 58, 138)
            self.cell(0, 6, "VCB Analytics - MBNT & TTQT/TTTM", ln=True)
            self.set_draw_color(30, 58, 138)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(4)

        def footer(self):
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8, f"Page {self.page_no()}/{{nb}}", align="C")

    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, _pdf_safe("Bao cao phan tich MBNT & TTQT"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(71, 85, 105)
    pdf.cell(
        0, 6,
        _pdf_safe(
            f"Phong: {data.get('phong') or 'VCB HCM'}  |  {data.get('d_from')} -> {data.get('d_to')}"
        ),
        ln=True,
    )
    pdf.ln(4)

    m = data.get("metrics") or {}
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 7, "Tong quan so lieu", ln=True)
    pdf.set_font("Helvetica", "", 9)
    for L in [
        f"MBNT DS (USD): {m.get('mbnt_ds', 0):,.0f}",
        f"MBNT LN (VND): {m.get('mbnt_ln', 0):,.0f}",
        f"MBNT so GD: {m.get('mbnt_gd', 0):,}",
        f"MBNT CIF: {m.get('mbnt_cif', 0):,}",
        f"TTQT DS (USD): {m.get('ttqt_ds', 0):,.0f}",
        f"TTQT XK: {m.get('ttqt_xk', 0):,.0f}  |  NK: {m.get('ttqt_nk', 0):,.0f}",
    ]:
        pdf.cell(0, 5, _pdf_safe(L), ln=True)
    pdf.ln(3)

    plan_summary = data.get("plan_summary") or []
    if plan_summary:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, "% hoan thanh ke hoach", ln=True)
        pdf.set_font("Helvetica", "", 8)
        for line in plan_summary:
            pdf.cell(0, 5, _pdf_safe(line), ln=True)
        pdf.ln(3)

    report = data.get("report_text") or ""
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Bao cao phan tich", ln=True)
    pdf.set_font("Helvetica", "", 8)
    for line in _pdf_safe(report).splitlines():
        pdf.multi_cell(0, 4, line)
    pdf.ln(4)

    charts = data.get("charts") or {}
    chart_order = [
        ("plan_pie", "Hoan thanh ke hoach (pie)"),
        ("mbnt_top_ds", "TOP DS MBNT"),
        ("mbnt_bot_ds", "BOT DS MBNT"),
        ("mbnt_top_ln", "TOP LN MBNT"),
        ("mbnt_bot_ln", "BOT LN MBNT"),
        ("ttqt_top", "TOP DS TTQT"),
        ("ttqt_bot", "BOT DS TTQT"),
        ("ttqt_xk_top", "TOP XK"),
        ("ttqt_nk_top", "TOP NK"),
        ("scatter", "Scatter TTQT vs MBNT"),
    ]
    for key, title in chart_order:
        b64 = charts.get(key) or ""
        if not b64:
            continue
        try:
            img_bytes = base64.b64decode(b64)
            img_path = io.BytesIO(img_bytes)
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 11)
            pdf.cell(0, 8, _pdf_safe(title), ln=True)
            pdf.image(img_path, x=10, w=190)
        except Exception as e:
            print(f"PDF chart embed failed ({key}): {e}")

    out = io.BytesIO()
    pdf.output(out)
    return out.getvalue()


# ───────────────────────── routes ─────────────────────────

@app.errorhandler(413)
def request_entity_too_large(_error):
    flash("Tong dung luong file upload vuot gioi han 80 MB.", "error")
    return redirect(url_for("index")), 413


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        has_data=bool(_STORE.get(_sid())),
        today=date.today().isoformat(),
    )


@app.route("/run", methods=["POST"])
def run_analysis():
    mbnt_f = request.files.get("mbnt")
    ttqt_f = request.files.get("ttqt")
    plan_f = request.files.get("plan")
    ext_f = request.files.get("external")
    phong = (request.form.get("phong") or "").strip()
    d_from_s = request.form.get("d_from") or "2026-01-01"
    d_to_s = request.form.get("d_to") or "2026-12-31"

    if not mbnt_f or not ttqt_f or mbnt_f.filename == "" or ttqt_f.filename == "":
        flash("Can upload du 2 file: Data_Raw_MBNT va Data_Raw_TTQT_TTTM.", "error")
        return redirect(url_for("index"))

    try:
        d_from = datetime.strptime(d_from_s, "%Y-%m-%d").date()
        d_to = datetime.strptime(d_to_s, "%Y-%m-%d").date()
    except ValueError:
        flash("Dinh dang ngay khong hop le.", "error")
        return redirect(url_for("index"))

    if d_from > d_to:
        flash("Ngay bat dau phai nho hon hoac bang ngay ket thuc.", "error")
        return redirect(url_for("index"))

    try:
        mbnt_raw = read_df(mbnt_f)
        ttqt_raw = read_df(ttqt_f)
        plan_raw = read_df(plan_f) if plan_f and plan_f.filename else None

        if mbnt_raw is None or ttqt_raw is None:
            raise ValueError("Khong doc duoc noi dung tu file du lieu.")

        dall, qall, SA, SQ = prepare_all(mbnt_raw, ttqt_raw)
        del mbnt_raw, ttqt_raw
        gc.collect()

        df, mm, tt, note = build_customer_df(dall, qall, phong, d_from, d_to)
        ext_text = read_external_text(ext_f)
        report_text = build_report(df, dall, qall, mm, tt, phong, d_from, d_to)

        d_filt, _ = filt(dall, phong, d_from, d_to)
        q_filt, _ = filt(qall, phong, d_from, d_to)

        cif_mbnt = build(d_filt, ["ds", "ln"])
        TDS, BDS = kpi_pair(cif_mbnt, "DS")
        TLN, BLN = kpi_pair(cif_mbnt, "LN")
        cq = build(q_filt, ["xk", "nk", "xnk", "ds"])
        TQ, BQ = kpi_pair(cq, "DS")

        metrics = {
            "mbnt_ds": float(d_filt.ds.sum()) if len(d_filt) else 0,
            "mbnt_ln": float(d_filt.ln.sum()) if len(d_filt) else 0,
            "mbnt_gd": int(len(d_filt)),
            "mbnt_cif": int(d_filt.loc[d_filt.cif != "", "cif"].nunique()) if len(d_filt) else 0,
            "ttqt_ds": float(q_filt.ds.sum()) if len(q_filt) else 0,
            "ttqt_xk": float(q_filt.xk.sum()) if len(q_filt) else 0,
            "ttqt_nk": float(q_filt.nk.sum()) if len(q_filt) else 0,
        }

        # Plan chart – always (exact Excel logic)
        plan_pie_b64, plan_summary = "", []
        if plan_raw is not None:
            plan_pie_b64, plan_summary = render_plan_chart(plan_raw, phong, d_from, d_to)

        charts = generate_charts_safe(TDS, BDS, TLN, BLN, TQ, BQ, q_filt, df)
        charts["plan_pie"] = plan_pie_b64

        # Group summary
        groups = {}
        try:
            for side, X, Y in [("XK -> BUY", "XK", "BUY"), ("NK -> SELL", "NK", "SELL")]:
                if X not in df.columns or Y not in df.columns:
                    continue
                dside = df[pd.to_numeric(df[X], errors="coerce").fillna(0) > 0].copy()
                if dside.empty:
                    continue
                dside[X] = pd.to_numeric(dside[X], errors="coerce").fillna(0)
                dside[Y] = pd.to_numeric(dside[Y], errors="coerce").fillna(0)
                r = dside[Y] / dside[X].replace(0, np.nan)
                lab = ["1. Khong GD MBNT", "2. <30%", "3. 30-80%", "4. 80-120%", "5. >120%"]
                dside["Nhom"] = np.select(
                    [dside[Y] == 0, r < 0.3, r < 0.8, r <= 1.2],
                    lab[:4],
                    default=lab[4],
                )
                summary = (
                    dside.groupby("Nhom", observed=True)
                    .agg(So_cty=("Ten", "size"), Tong_X=(X, "sum"), Tong_Y=(Y, "sum"))
                    .reset_index()
                )
                summary["Y/X_%"] = np.where(
                    summary["Tong_X"] > 0,
                    (summary["Tong_Y"] / summary["Tong_X"] * 100).round(1),
                    0,
                )
                groups[side] = summary.to_dict(orient="records")
                del dside, summary
        except Exception as group_err:
            print(f"Group summary warning: {group_err}")
            groups = {}

        df_csv = df.to_csv(index=True).encode("utf-8-sig")
        tds_html = safe_html(TDS)
        bds_html = safe_html(BDS)

        _STORE[_sid()] = {
            "note": note,
            "report_text": report_text,
            "ext_text": ext_text,
            "charts": charts,
            "groups": groups,
            "metrics": metrics,
            "plan_summary": plan_summary,
            "plan_rows": [],
            "df_csv": df_csv,
            "tds_html": tds_html,
            "bds_html": bds_html,
            "phong": phong,
            "d_from": d_from.isoformat(),
            "d_to": d_to.isoformat(),
            "ai_report": None,
            "chat_history": [],
        }
        msg = note
        if plan_raw is None:
            msg += " (Chua upload KE_HOACH - khong co pie ke hoach)"
        if not ENABLE_CHARTS:
            msg += " (Bar charts tat - set ENABLE_CHARTS=1 neu can)"
        flash(msg, "success")

        del dall, qall, df, mm, tt, d_filt, q_filt
        del cif_mbnt, TDS, BDS, TLN, BLN, cq, TQ, BQ, SA, SQ
        if plan_raw is not None:
            del plan_raw
        gc.collect()

    except Exception as e:
        flash(f"Loi xu ly du lieu hoac vuot qua gioi han bo nho server: {e}", "error")
        gc.collect()
        return redirect(url_for("index"))

    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    data = _STORE.get(_sid())
    if not data:
        flash("Chua co du lieu. Hay upload va chay phan tich truoc.", "error")
        return redirect(url_for("index"))
    return render_template("dashboard.html", data=data)


@app.route("/export/csv")
def export_csv():
    data = _STORE.get(_sid())
    if not data:
        return redirect(url_for("index"))
    return send_file(
        io.BytesIO(data["df_csv"]),
        mimetype="text/csv",
        as_attachment=True,
        download_name="customer_analysis.csv",
    )


@app.route("/export/report")
def export_report():
    data = _STORE.get(_sid())
    if not data:
        return redirect(url_for("index"))
    return send_file(
        io.BytesIO(data["report_text"].encode("utf-8")),
        mimetype="text/plain",
        as_attachment=True,
        download_name="bao_cao_phan_tich.txt",
    )


@app.route("/export/pdf")
def export_pdf():
    data = _STORE.get(_sid())
    if not data:
        return redirect(url_for("index"))
    try:
        pdf_bytes = build_pdf(data)
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="vcb_bao_cao.pdf",
        )
    except Exception as e:
        flash(f"Khong tao duoc PDF: {e}", "error")
        return redirect(url_for("dashboard"))


@app.route("/ai/generate", methods=["POST"])
def ai_generate():
    data = _STORE.get(_sid())
    if not data:
        return jsonify({"error": "Chua co du lieu"}), 400
    context = data.get("report_text", "")
    if data.get("ext_text"):
        context += "\n\n=== BAO CAO EXTERNAL ===\n" + data["ext_text"][:8000]
    report = generate_professional_report(context, api_key=None)
    data["ai_report"] = report
    return jsonify({"report": report})


@app.route("/ai/chat", methods=["POST"])
def ai_chat():
    data = _STORE.get(_sid())
    if not data:
        return jsonify({"error": "Chua co du lieu"}), 400
    body = request.json or {}
    user_msg = body.get("message", "").strip()
    if not user_msg:
        return jsonify({"error": "Empty message"}), 400
    history = data.get("chat_history") or []
    history.append({"role": "user", "content": user_msg})
    context = data.get("report_text", "")
    if data.get("ext_text"):
        context += "\n\n=== BAO CAO EXTERNAL ===\n" + data["ext_text"][:8000]
    ans = chat_with_data(history, context, api_key=None)
    history.append({"role": "assistant", "content": ans})
    data["chat_history"] = history
    return jsonify({"answer": ans, "history": history})


@app.route("/reset", methods=["POST"])
def reset():
    sid = _sid()
    _STORE.pop(sid, None)
    gc.collect()
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
