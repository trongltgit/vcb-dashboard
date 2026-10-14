"""
VCB Dashboard – Flask version (thay Streamlit)
Deploy Render: gunicorn app:app
"""
import base64
import io
import os
from datetime import date, datetime
from functools import wraps

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
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
from utils.charts import chart_bar, scatter_ttqt_mbnt
from utils.ai_bot import generate_professional_report, chat_with_data

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "vcb-dashboard-dev-key-change-me")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80 MB

# In-memory store (single-process / one user at a time friendly on free Render)
# Keyed by simple session id
_STORE = {}


def _sid():
    if "sid" not in session:
        session["sid"] = os.urandom(8).hex()
    return session["sid"]


def fig_to_base64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def read_df(file_storage):
    if file_storage is None or file_storage.filename == "":
        return None
    name = file_storage.filename.lower()
    data = file_storage.read()
    bio = io.BytesIO(data)
    if name.endswith(".csv"):
        return pd.read_csv(bio)
    return pd.read_excel(bio)


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


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        has_data=bool(_STORE.get(_sid())),
        groq_key=os.getenv("GROQ_API_KEY", ""),
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
    groq_key = (request.form.get("groq_key") or "").strip()
    if groq_key:
        os.environ["GROQ_API_KEY"] = groq_key

    if not mbnt_f or not ttqt_f or mbnt_f.filename == "" or ttqt_f.filename == "":
        flash("Cần upload đủ 2 file: Data_Raw_MBNT và Data_Raw_TTQT_TTTM.", "error")
        return redirect(url_for("index"))

    try:
        d_from = datetime.strptime(d_from_s, "%Y-%m-%d").date()
        d_to = datetime.strptime(d_to_s, "%Y-%m-%d").date()
    except ValueError:
        flash("Định dạng ngày không hợp lệ.", "error")
        return redirect(url_for("index"))

    try:
        mbnt_raw = read_df(mbnt_f)
        ttqt_raw = read_df(ttqt_f)
        dall, qall, SA, SQ = prepare_all(mbnt_raw, ttqt_raw)
        df, mm, tt, note = build_customer_df(dall, qall, phong, d_from, d_to)
        ext_text = read_external_text(ext_f)
        report_text = build_report(df, dall, qall, mm, tt, phong, d_from, d_to)

        # Pre-generate charts as base64
        charts = {}
        d_filt, _ = filt(dall, phong, d_from, d_to)
        cif_mbnt = build(d_filt, ["ds", "ln"])
        TDS, BDS = kpi_pair(cif_mbnt, "DS")
        TLN, BLN = kpi_pair(cif_mbnt, "LN")
        charts["mbnt_top_ds"] = fig_to_base64(chart_bar(TDS, "TOP 20% – DOANH SỐ MBNT", "USD", "#2563EB", True, note))
        charts["mbnt_bot_ds"] = fig_to_base64(chart_bar(BDS, "BOT 10% – DOANH SỐ MBNT", "USD", "#F97316", False, note))
        charts["mbnt_top_ln"] = fig_to_base64(chart_bar(TLN, "TOP 20% – LỢI NHUẬN MBNT", "VND", "#16A34A", True, note))
        charts["mbnt_bot_ln"] = fig_to_base64(chart_bar(BLN, "BOT 10% – LỢI NHUẬN MBNT", "VND", "#7C3AED", False, note))

        q_filt, _ = filt(qall, phong, d_from, d_to)
        cq = build(q_filt, ["xk", "nk", "xnk", "ds"])
        TQ, BQ = kpi_pair(cq, "DS")
        charts["ttqt_top"] = fig_to_base64(chart_bar(TQ, "TOP 20% – DOANH SỐ TTQT", "USD", "#0EA5E9", True, note))
        charts["ttqt_bot"] = fig_to_base64(chart_bar(BQ, "BOT 10% – DOANH SỐ TTQT", "USD", "#EF4444", False, note))

        if "XK" in cq.columns:
            TX, BX = kpi_pair(cq[cq.XK > 0] if len(cq) else cq, "XK")
            if not TX.empty:
                charts["ttqt_xk_top"] = fig_to_base64(chart_bar(TX, "TOP 20% – TTQT XK", "USD", "#0EA5E9", True))
            if not BX.empty:
                charts["ttqt_xk_bot"] = fig_to_base64(chart_bar(BX, "BOT 10% – TTQT XK", "USD", "#EF4444", False))
        if "NK" in cq.columns:
            TN, BN = kpi_pair(cq[cq.NK > 0] if len(cq) else cq, "NK")
            if not TN.empty:
                charts["ttqt_nk_top"] = fig_to_base64(chart_bar(TN, "TOP 20% – TTQT NK", "USD", "#14B8A6", True))
            if not BN.empty:
                charts["ttqt_nk_bot"] = fig_to_base64(chart_bar(BN, "BOT 10% – TTQT NK", "USD", "#A855F7", False))

        if len(df):
            charts["scatter"] = fig_to_base64(scatter_ttqt_mbnt(df))

        # Group summary for relationship
        groups = {}
        for side, X, Y in [("XK → BUY", "XK", "BUY"), ("NK → SELL", "NK", "SELL")]:
            dside = df[df[X] > 0].copy()
            if dside.empty:
                continue
            r = dside[Y] / dside[X]
            lab = ["1. Không GD MBNT", "2. <30%", "3. 30-80%", "4. 80-120%", "5. >120%"]
            dside["Nhom"] = np.select([dside[Y] == 0, r < 0.3, r < 0.8, r <= 1.2], lab[:4], default=lab[4])
            summary = (
                dside.groupby("Nhom")
                .agg(So_cty=("Ten", "size"), Tong_X=(X, "sum"), Tong_Y=(Y, "sum"))
                .reset_index()
            )
            summary["Y/X_%"] = (summary.Tong_Y / summary.Tong_X * 100).round(1)
            groups[side] = summary.to_dict(orient="records")

        metrics = {
            "mbnt_ds": float(d_filt.ds.sum()) if len(d_filt) else 0,
            "mbnt_ln": float(d_filt.ln.sum()) if len(d_filt) else 0,
            "mbnt_gd": int(len(d_filt)),
            "mbnt_cif": int(d_filt.loc[d_filt.cif != "", "cif"].nunique()) if len(d_filt) else 0,
            "ttqt_ds": float(q_filt.ds.sum()) if len(q_filt) else 0,
            "ttqt_xk": float(q_filt.xk.sum()) if len(q_filt) else 0,
            "ttqt_nk": float(q_filt.nk.sum()) if len(q_filt) else 0,
        }

        # Store tables as CSV strings for download
        df_csv = df.to_csv(index=True).encode("utf-8-sig")
        tds_html = TDS.drop(columns=["Key"], errors="ignore").to_html(classes="table", index=False) if not TDS.empty else ""
        bds_html = BDS.drop(columns=["Key"], errors="ignore").to_html(classes="table", index=False) if not BDS.empty else ""

        _STORE[_sid()] = {
            "note": note,
            "report_text": report_text,
            "ext_text": ext_text,
            "charts": charts,
            "groups": groups,
            "metrics": metrics,
            "df_csv": df_csv,
            "tds_html": tds_html,
            "bds_html": bds_html,
            "phong": phong,
            "d_from": d_from.isoformat(),
            "d_to": d_to.isoformat(),
            "ai_report": None,
            "chat_history": [],
        }
        flash(note, "success")
    except Exception as e:
        flash(f"Lỗi xử lý dữ liệu: {e}", "error")
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


@app.route("/ai/generate", methods=["POST"])
def ai_generate():
    data = _STORE.get(_sid())
    if not data:
        return jsonify({"error": "Chưa có dữ liệu"}), 400
    key = request.json.get("api_key") or os.getenv("GROQ_API_KEY")
    context = data.get("report_text", "")
    if data.get("ext_text"):
        context += "\n\n=== BÁO CÁO EXTERNAL ===\n" + data["ext_text"][:8000]
    report = generate_professional_report(context, api_key=key or None)
    data["ai_report"] = report
    return jsonify({"report": report})


@app.route("/ai/chat", methods=["POST"])
def ai_chat():
    data = _STORE.get(_sid())
    if not data:
        return jsonify({"error": "Chưa có dữ liệu"}), 400
    body = request.json or {}
    key = body.get("api_key") or os.getenv("GROQ_API_KEY")
    user_msg = body.get("message", "").strip()
    if not user_msg:
        return jsonify({"error": "Empty message"}), 400
    history = data.get("chat_history") or []
    history.append({"role": "user", "content": user_msg})
    context = data.get("report_text", "")
    if data.get("ext_text"):
        context += "\n\n=== BÁO CÁO EXTERNAL ===\n" + data["ext_text"][:8000]
    ans = chat_with_data(history, context, api_key=key or None)
    history.append({"role": "assistant", "content": ans})
    data["chat_history"] = history
    return jsonify({"answer": ans, "history": history})


@app.route("/reset", methods=["POST"])
def reset():
    sid = _sid()
    _STORE.pop(sid, None)
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
