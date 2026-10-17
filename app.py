"""
VCB Dashboard – Flask (memory-safe for Render free tier)
- No Groq key on frontend (uses GROQ_API_KEY env only)
- Plan-completion pie charts (MBNT / TTQT vs KE_HOACH)
- Export full report to PDF
"""
import base64
import gc
import io
import os
import re
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

_STORE = {}


def _sid():
    if "sid" not in session:
        session["sid"] = os.urandom(8).hex()
    return session["sid"]


# ───────────────────────── helpers ─────────────────────────

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
                f"Định dạng file không hỗ trợ: {filename}. "
                "Vui lòng dùng CSV, XLS, XLSX hoặc XLSM."
            )

        if df.empty and len(df.columns) == 0:
            raise ValueError(f"File {filename} không có bảng dữ liệu.")

        df.columns = [str(col).strip() for col in df.columns]
        return df
    except Exception as exc:
        raise ValueError(f"Không đọc được file '{filename}': {exc}") from exc


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
            return "[Không đọc được PDF]"
    if name.endswith(".docx"):
        try:
            from docx import Document
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            return "[Không đọc được DOCX]"
    return ""


def usd(v):
    a = abs(v)
    if a >= 1e9:
        return f"{a/1e9:,.2f} tỷ USD"
    if a >= 1e6:
        return f"{a/1e6:,.1f} triệu USD"
    if a >= 1e3:
        return f"{a/1e3:,.0f} nghìn USD"
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
        return f"<p class='text-danger'>Không render được bảng: {e}</p>"


def _num_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)
    return (
        pd.to_numeric(
            s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", "", regex=False),
            errors="coerce",
        ).fillna(0)
    )


def load_plan(plan_raw: pd.DataFrame | None) -> list[dict]:
    """
    Flexible KE_HOACH loader.
    Looks for columns like: Metric/Chỉ tiêu, KH_Quy, KH_Nam, TH (optional).
    Returns list of {name, th, kh_q, kh_y, pct_q, pct_y}.
    Actual TH values are filled later from metrics.
    """
    if plan_raw is None or plan_raw.empty:
        return []

    cols = {str(c).strip().lower(): c for c in plan_raw.columns}

    def find(*cands):
        for c in cands:
            for k, orig in cols.items():
                if c in k or k == c:
                    return orig
        return None

    name_col = find("metric", "chi tieu", "chi_tieu", "chỉ tiêu", "ten", "name", "hang muc", "hạng mục")
    kh_q_col = find("kh_quy", "kh quý", "kh quy", "ke hoach quy", "kế hoạch quý", "target_q", "khq")
    kh_y_col = find("kh_nam", "kh năm", "kh nam", "ke hoach nam", "kế hoạch năm", "target_y", "khn")
    th_col = find("thuc hien", "thực hiện", "th", "actual", "ds")

    if name_col is None:
        # first column as name
        name_col = plan_raw.columns[0]

    rows = []
    for _, r in plan_raw.iterrows():
        name = str(r.get(name_col, "")).strip()
        if not name or name.lower() in ("nan", "none", ""):
            continue
        kh_q = float(_num_series(pd.Series([r.get(kh_q_col, 0)])).iloc[0]) if kh_q_col else 0.0
        kh_y = float(_num_series(pd.Series([r.get(kh_y_col, 0)])).iloc[0]) if kh_y_col else 0.0
        th = float(_num_series(pd.Series([r.get(th_col, 0)])).iloc[0]) if th_col else 0.0
        rows.append({
            "name": name,
            "th": th,
            "kh_q": kh_q,
            "kh_y": kh_y,
            "pct_q": (th / kh_q) if kh_q > 0 else 0.0,
            "pct_y": (th / kh_y) if kh_y > 0 else 0.0,
        })
    return rows


def match_plan_to_metrics(plan_rows: list, metrics: dict) -> list:
    """
    Overlay actual performance onto plan rows by name keywords.
    MBNT DS / LN, TTQT DS / XK / NK.
    """
    if not plan_rows:
        # synthesise default plan rows from metrics if no plan file
        return [
            {"name": "DS MBNT (USD)", "th": metrics.get("mbnt_ds", 0), "kh_q": 0, "kh_y": 0, "pct_q": 0, "pct_y": 0},
            {"name": "LN MBNT (VND)", "th": metrics.get("mbnt_ln", 0), "kh_q": 0, "kh_y": 0, "pct_q": 0, "pct_y": 0},
            {"name": "DS TTQT (USD)", "th": metrics.get("ttqt_ds", 0), "kh_q": 0, "kh_y": 0, "pct_q": 0, "pct_y": 0},
            {"name": "XK TTQT (USD)", "th": metrics.get("ttqt_xk", 0), "kh_q": 0, "kh_y": 0, "pct_q": 0, "pct_y": 0},
            {"name": "NK TTQT (USD)", "th": metrics.get("ttqt_nk", 0), "kh_q": 0, "kh_y": 0, "pct_q": 0, "pct_y": 0},
        ]

    mapping = [
        (["mbnt", "ngoai te", "ngoại tệ", "forex", "ds mbnt"], "mbnt_ds"),
        (["ln", "loi nhuan", "lợi nhuận", "profit"], "mbnt_ln"),
        (["ttqt", "xnk", "thanh toan quoc te", "thanh toán quốc tế", "ds ttqt"], "ttqt_ds"),
        (["xk", "xuat khau", "xuất khẩu", "export"], "ttqt_xk"),
        (["nk", "nhap khau", "nhập khẩu", "import"], "ttqt_nk"),
    ]

    out = []
    for row in plan_rows:
        name_l = row["name"].lower()
        th = row["th"]
        for keys, mkey in mapping:
            if any(k in name_l for k in keys):
                th = float(metrics.get(mkey, th) or 0)
                break
        kh_q, kh_y = row["kh_q"], row["kh_y"]
        out.append({
            "name": row["name"],
            "th": th,
            "kh_q": kh_q,
            "kh_y": kh_y,
            "pct_q": (th / kh_q) if kh_q > 0 else 0.0,
            "pct_y": (th / kh_y) if kh_y > 0 else 0.0,
        })
    return out


def chart_plan_pies(plan_rows: list, phong: str, d_from, d_to) -> str:
    """
    Simple 2-row pie/donut gauges: % vs KH Quý and % vs KH Năm.
    Lightweight – always generated when plan_rows exist (even if ENABLE_CHARTS=0).
    """
    if not plan_rows:
        return ""

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        n = min(len(plan_rows), 5)
        rows = plan_rows[:n]
        fig, axes = plt.subplots(2, n, figsize=(2.2 * n, 4.6), dpi=80)
        if n == 1:
            axes = np.array([[axes[0]], [axes[1]]])

        colors = ["#0EA5E9", "#22C55E", "#8B5CF6", "#F59E0B", "#EF4444"]
        track = "#E5E7EB"

        for i, r in enumerate(rows):
            for row_i, (pct, label) in enumerate([
                (r["pct_q"], "Quý"),
                (r["pct_y"], "Năm"),
            ]):
                ax = axes[row_i, i]
                fill = max(0.0, min(1.0, pct))
                if r["kh_q" if row_i == 0 else "kh_y"] <= 0:
                    ax.pie([1], colors=[track], startangle=90,
                           wedgeprops={"width": 0.35, "linewidth": 0})
                    ax.text(0, 0, "N/A", ha="center", va="center", fontsize=8, color="#9CA3AF")
                else:
                    vals = [fill, 1 - fill] if fill < 1 else [1]
                    cols = [colors[i % len(colors)], track] if fill < 1 else [colors[i % len(colors)]]
                    ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                           wedgeprops={"width": 0.35, "linewidth": 0})
                    ax.text(0, 0.05, f"{pct*100:.1f}%", ha="center", va="center",
                            fontsize=9, fontweight="bold", color="#111827")
                    ax.text(0, -0.28, label, ha="center", va="center", fontsize=6, color="#6B7280")
                ax.set_aspect("equal")
                if row_i == 0:
                    title = r["name"][:22] + ("…" if len(r["name"]) > 22 else "")
                    ax.set_title(title, fontsize=7, pad=4)

        fig.suptitle(
            f"% hoàn thành kế hoạch – {phong or 'VCB HCM'}  |  {d_from:%d/%m/%Y} → {d_to:%d/%m/%Y}",
            fontsize=9, fontweight="bold", y=0.98,
        )
        fig.subplots_adjust(top=0.88, bottom=0.05, left=0.02, right=0.98, wspace=0.15, hspace=0.25)

        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=80)
        plt.close(fig)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode("ascii")
    except Exception as e:
        print(f"Plan pie chart failed: {e}")
        return ""


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


def build_report(df, dall, qall, mm, tt, phong, d_from, d_to):
    def phan_tich(X, O, Y, Yv, tx_, ty_):
        L = []
        d = df[df[X] > 0].copy()
        if d.empty:
            return [f"Không có công ty nào có {tx_} trong kỳ đã chọn."], {}
        TX, TY, TV = d[X].sum(), d[Y].sum(), d[Yv].sum()
        r = d[Y] / d[X]
        L.append(
            f"Có {len(d):,} công ty có {tx_} với tổng {usd(TX)}. "
            f"Trong số đó {int((d[Y]>0).sum()):,} công ty có giao dịch {ty_} với VCB, tổng {usd(TY)}."
        )
        L.append(f"Tỷ lệ chuyển đổi chung: cứ 100 USD {tx_} thì VCB nhận được {TY/TX*100:.0f} USD {ty_}.")
        lab = ["Không giao dịch", "Dưới 30%", "Từ 30% đến 80%", "Từ 80% đến 120%", "Trên 120%"]
        d["Nhom"] = np.select([d[Y] == 0, r < 0.3, r < 0.8, r <= 1.2], lab[:4], default=lab[4])
        L.append(f"Phân nhóm theo tỷ lệ {ty_}/{tx_}:")
        for k in lab:
            s = d[d.Nhom == k]
            if len(s):
                L.append(f"   - {k}: {len(s):,} công ty, chiếm {pc(s[X].sum(), TX)} tổng {tx_}.")
        lo = d[d.Nhom.isin(lab[:2])]
        L.append(
            f"Nhóm dưới 30% (gồm cả không giao dịch) có {len(lo):,} công ty, "
            f"nắm {pc(lo[X].sum(), TX)} tổng {tx_} nhưng chỉ đóng góp {pc(lo[Y].sum(), TY)} tổng {ty_}."
        )
        L.append(f"Giải ngân vốn vay ngoại tệ chiếm {pc(TV, TY)} tổng {ty_} ({usd(TV)}).")
        d["Net"] = (d[X] - d[O]).clip(lower=0)
        d["Tn"] = (d.Net - d[Y]).clip(lower=0)
        low = d[d[Y] < 0.5 * d[X]].copy()
        low["Nn"] = np.select([low[O] >= 0.7 * low[X], low[Y] == 0], ["A", "B"], default="C")
        tn = low[low.Nn.isin(["B", "C"])].Tn.sum()
        L.append(f"Tiềm năng khai thác thực (nhóm B+C): khoảng {usd(tn)}.")
        return L, {"conv": TY / TX if TX else 0, "tn": tn}

    LX, sx = phan_tich("XK", "NK", "BUY", "BUY_VAY", "XK", "BUY")
    LN, sn = phan_tich("NK", "XK", "SELL", "SELL_VAY", "NK", "SELL")

    R = [
        "BÁO CÁO PHÂN TÍCH QUAN HỆ DS TTQT – DS MBNT",
        f"Phòng/đơn vị: {phong or 'VCB HCM'}. Kỳ: {d_from:%d/%m/%Y} đến {d_to:%d/%m/%Y}.",
        f"Dữ liệu: {len(dall):,} dòng MBNT raw → {len(mm):,} dùng phân tích; "
        f"{len(qall):,} dòng TTQT → {len(tt):,} dùng. {len(df):,} doanh nghiệp (đã loại cá nhân).",
        "",
        "PHẦN 1. XUẤT KHẨU (XK) VÀ BUY",
    ] + LX + ["", "PHẦN 2. NHẬP KHẨU (NK) VÀ SELL"] + LN + ["", "PHẦN 3. KẾT LUẬN"]

    if sx and sn:
        cx, cnv = sx["conv"] * 100, sn["conv"] * 100
        R.append(f"Tỷ lệ chuyển đổi XK→BUY: {cx:.0f}%, NK→SELL: {cnv:.0f}%.")
        R.append(
            "Công ty giữ chân VCB tốt hơn ở chiều mua ngoại tệ (thanh toán NK)."
            if cnv > cx
            else "Công ty giữ chân VCB tốt hơn ở chiều bán ngoại tệ."
        )
        ln_ = "XK" if sx["tn"] >= sn["tn"] else "NK"
        R.append(f"Tiềm năng khai thác lớn hơn nằm ở phía {ln_}.")

    return "\n".join(R)


def build_pdf(data: dict) -> bytes:
    """Build a multi-page PDF: cover metrics + report text + available charts."""
    from fpdf import FPDF

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 10)
            self.set_text_color(30, 58, 138)
            self.cell(0, 6, "VCB Analytics – MBNT & TTQT/TTTM", ln=True)
            self.set_draw_color(30, 58, 138)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(4)

        def footer(self):
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8, f"Trang {self.page_no()}/{{nb}}", align="C")

    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Title
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, "Bao cao phan tich MBNT & TTQT", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(71, 85, 105)
    pdf.cell(0, 6, f"Phong: {data.get('phong') or 'VCB HCM'}  |  {data.get('d_from')} -> {data.get('d_to')}", ln=True)
    pdf.ln(4)

    # Metrics box
    m = data.get("metrics") or {}
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 7, "Tong quan so lieu", ln=True)
    pdf.set_font("Helvetica", "", 9)
    lines = [
        f"MBNT DS (USD): {m.get('mbnt_ds', 0):,.0f}",
        f"MBNT LN (VND): {m.get('mbnt_ln', 0):,.0f}",
        f"MBNT so GD: {m.get('mbnt_gd', 0):,}",
        f"MBNT CIF: {m.get('mbnt_cif', 0):,}",
        f"TTQT DS (USD): {m.get('ttqt_ds', 0):,.0f}",
        f"TTQT XK: {m.get('ttqt_xk', 0):,.0f}  |  NK: {m.get('ttqt_nk', 0):,.0f}",
    ]
    for L in lines:
        pdf.cell(0, 5, L, ln=True)
    pdf.ln(3)

    # Plan rows
    plan_rows = data.get("plan_rows") or []
    if plan_rows:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, "% hoan thanh ke hoach", ln=True)
        pdf.set_font("Helvetica", "", 8)
        for r in plan_rows:
            pdf.cell(
                0, 5,
                f"{r['name'][:40]}: TH={r['th']:,.1f} | KH_Q={r['kh_q']:,.1f} ({r['pct_q']*100:.1f}%) | KH_Y={r['kh_y']:,.1f} ({r['pct_y']*100:.1f}%)",
                ln=True,
            )
        pdf.ln(3)

    # Report text (strip non-latin for core fonts; still readable)
    report = data.get("report_text") or ""
    # FPDF core fonts are Latin-1; keep ASCII-ish fallback by encoding safely
    safe_report = report.encode("latin-1", errors="replace").decode("latin-1")
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Bao cao phan tich", ln=True)
    pdf.set_font("Helvetica", "", 8)
    for line in safe_report.splitlines():
        pdf.multi_cell(0, 4, line)
    pdf.ln(4)

    # Charts as images
    charts = data.get("charts") or {}
    chart_order = [
        ("plan_pie", "Ke hoach % hoan thanh"),
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
            pdf.cell(0, 8, title, ln=True)
            pdf.image(img_path, x=10, w=190)
        except Exception as e:
            print(f"PDF chart embed failed ({key}): {e}")

    out = io.BytesIO()
    pdf.output(out)
    return out.getvalue()


# ───────────────────────── routes ─────────────────────────

@app.errorhandler(413)
def request_entity_too_large(_error):
    flash("Tổng dung lượng file upload vượt giới hạn 80 MB.", "error")
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
        flash("Cần upload đủ 2 file: Data_Raw_MBNT và Data_Raw_TTQT_TTTM.", "error")
        return redirect(url_for("index"))

    try:
        d_from = datetime.strptime(d_from_s, "%Y-%m-%d").date()
        d_to = datetime.strptime(d_to_s, "%Y-%m-%d").date()
    except ValueError:
        flash("Định dạng ngày không hợp lệ.", "error")
        return redirect(url_for("index"))

    if d_from > d_to:
        flash("Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.", "error")
        return redirect(url_for("index"))

    try:
        mbnt_raw = read_df(mbnt_f)
        ttqt_raw = read_df(ttqt_f)
        plan_raw = read_df(plan_f) if plan_f and plan_f.filename else None

        if mbnt_raw is None or ttqt_raw is None:
            raise ValueError("Không đọc được nội dung từ file dữ liệu.")

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

        # Plan
        plan_rows = match_plan_to_metrics(load_plan(plan_raw), metrics)
        del plan_raw

        # Charts (bar/scatter optional; plan pie always attempted if plan data exists)
        charts = generate_charts_safe(TDS, BDS, TLN, BLN, TQ, BQ, q_filt, df)
        charts["plan_pie"] = chart_plan_pies(plan_rows, phong, d_from, d_to)

        # Group summary
        groups = {}
        try:
            for side, X, Y in [("XK → BUY", "XK", "BUY"), ("NK → SELL", "NK", "SELL")]:
                if X not in df.columns or Y not in df.columns:
                    continue
                dside = df[pd.to_numeric(df[X], errors="coerce").fillna(0) > 0].copy()
                if dside.empty:
                    continue
                dside[X] = pd.to_numeric(dside[X], errors="coerce").fillna(0)
                dside[Y] = pd.to_numeric(dside[Y], errors="coerce").fillna(0)
                r = dside[Y] / dside[X].replace(0, np.nan)
                lab = ["1. Không GD MBNT", "2. <30%", "3. 30-80%", "4. 80-120%", "5. >120%"]
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
            "plan_rows": plan_rows,
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
        if not ENABLE_CHARTS:
            msg += " (Bar charts tắt – set ENABLE_CHARTS=1 nếu cần)"
        flash(msg, "success")

        del dall, qall, df, mm, tt, d_filt, q_filt
        del cif_mbnt, TDS, BDS, TLN, BLN, cq, TQ, BQ, SA, SQ
        gc.collect()

    except Exception as e:
        flash(f"Lỗi xử lý dữ liệu hoặc vượt quá giới hạn bộ nhớ server: {e}", "error")
        gc.collect()
        return redirect(url_for("index"))

    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    data = _STORE.get(_sid())
    if not data:
        flash("Chưa có dữ liệu. Hãy upload và chạy phân tích trước.", "error")
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
        flash(f"Không tạo được PDF: {e}", "error")
        return redirect(url_for("dashboard"))


@app.route("/ai/generate", methods=["POST"])
def ai_generate():
    data = _STORE.get(_sid())
    if not data:
        return jsonify({"error": "Chưa có dữ liệu"}), 400
    # Always use server env key – never accept from client
    context = data.get("report_text", "")
    if data.get("ext_text"):
        context += "\n\n=== BÁO CÁO EXTERNAL ===\n" + data["ext_text"][:8000]
    report = generate_professional_report(context, api_key=None)
    data["ai_report"] = report
    return jsonify({"report": report})


@app.route("/ai/chat", methods=["POST"])
def ai_chat():
    data = _STORE.get(_sid())
    if not data:
        return jsonify({"error": "Chưa có dữ liệu"}), 400
    body = request.json or {}
    user_msg = body.get("message", "").strip()
    if not user_msg:
        return jsonify({"error": "Empty message"}), 400
    history = data.get("chat_history") or []
    history.append({"role": "user", "content": user_msg})
    context = data.get("report_text", "")
    if data.get("ext_text"):
        context += "\n\n=== BÁO CÁO EXTERNAL ===\n" + data["ext_text"][:8000]
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
