"""
VCB Dashboard – Phân tích MBNT & TTQT/TTTM
Web app thay thế Excel Python, hỗ trợ Excel 2016/365 qua upload + AI bot (Groq).
Deploy: Render / Railway / Streamlit Cloud.
"""
import io
import os
from datetime import date, datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from utils.processing import (
    prepare_all,
    build_customer_df,
    filt,
    build,
    kpi_pair,
    TOP_PCT,
    BOT_PCT,
)
from utils.charts import chart_bar, scatter_ttqt_mbnt
from utils.ai_bot import generate_professional_report, chat_with_data

load_dotenv()

st.set_page_config(
    page_title="VCB Dashboard MBNT & TTQT",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============== Sidebar ==============
with st.sidebar:
    st.title("🏦 VCB Analytics")
    st.markdown("**Bộ lọc & Cấu hình**")

    # Groq API
    groq_key = st.text_input(
        "Groq API Key",
        value=os.getenv("GROQ_API_KEY", ""),
        type="password",
        help="Dùng để AI viết báo cáo chuyên sâu. Lấy tại console.groq.com",
    )
    if groq_key:
        os.environ["GROQ_API_KEY"] = groq_key

    st.divider()
    st.subheader("1. Upload dữ liệu raw")

    mbnt_file = st.file_uploader(
        "Data_Raw_MBNT (Excel/CSV)",
        type=["xlsx", "xls", "csv"],
        key="mbnt",
    )
    ttqt_file = st.file_uploader(
        "Data_Raw_TTQT_TTTM (Excel/CSV)",
        type=["xlsx", "xls", "csv"],
        key="ttqt",
    )
    plan_file = st.file_uploader(
        "KE_HOACH (tùy chọn – cho pie hoàn thành KH)",
        type=["xlsx", "xls", "csv"],
        key="plan",
    )
    external_report = st.file_uploader(
        "Upload báo cáo external (PDF/DOCX/TXT) – AI sẽ tham chiếu",
        type=["pdf", "docx", "txt", "md"],
        key="ext",
    )

    st.divider()
    st.subheader("2. Bộ lọc")
    phong = st.text_input("Phòng / Chi nhánh (để trống = VCB HCM)", value="")
    col1, col2 = st.columns(2)
    with col1:
        d_from = st.date_input("Từ ngày", value=date(2026, 1, 1))
    with col2:
        d_to = st.date_input("Đến ngày", value=date(2026, 12, 31))

    run_btn = st.button("🚀 Chạy phân tích", type="primary", use_container_width=True)

# ============== Helpers ==============
@st.cache_data(show_spinner="Đang đọc & chuẩn hóa dữ liệu…")
def load_raw(mbnt_bytes, ttqt_bytes, mbnt_name, ttqt_name):
    def _read(b, name):
        if b is None:
            return None
        bio = io.BytesIO(b)
        if name.lower().endswith(".csv"):
            return pd.read_csv(bio)
        return pd.read_excel(bio)

    mbnt = _read(mbnt_bytes, mbnt_name or "mbnt.xlsx")
    ttqt = _read(ttqt_bytes, ttqt_name or "ttqt.xlsx")
    return mbnt, ttqt


def read_external(file) -> str:
    if file is None:
        return ""
    name = file.name.lower()
    data = file.read()
    if name.endswith(".txt") or name.endswith(".md"):
        return data.decode("utf-8", errors="ignore")
    if name.endswith(".pdf"):
        try:
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(data))
            return "\n".join(p.extract_text() or "" for p in reader.pages)
        except Exception:
            return "[Không đọc được PDF – cài thêm pypdf nếu cần]"
    if name.endswith(".docx"):
        try:
            from docx import Document
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            return "[Không đọc được DOCX]"
    return ""


# ============== Main ==============
st.title("📊 Dashboard Phân tích MBNT & TTQT/TTTM")
st.caption("Thay thế hoàn toàn Python in Excel – chạy ổn định, có AI bot viết báo cáo chuyên nghiệp.")

if not run_btn and "dall" not in st.session_state:
    st.info(
        "👈 Upload **Data_Raw_MBNT** và **Data_Raw_TTQT_TTTM**, chọn bộ lọc, rồi bấm **Chạy phân tích**.\n\n"
        "Bạn có thể nhập Groq API Key để kích hoạt AI bot viết báo cáo sâu hơn."
    )
    st.stop()

# Load & process
if run_btn or "dall" not in st.session_state:
    if mbnt_file is None or ttqt_file is None:
        st.error("Cần upload đủ 2 file raw (MBNT + TTQT).")
        st.stop()

    with st.spinner("Đang xử lý dữ liệu (group khách hàng, chuẩn hóa tên, lọc cá nhân)…"):
        mbnt_raw, ttqt_raw = load_raw(
            mbnt_file.getvalue(), ttqt_file.getvalue(), mbnt_file.name, ttqt_file.name
        )
        dall, qall, SA, SQ = prepare_all(mbnt_raw, ttqt_raw)
        df, mm, tt, note = build_customer_df(dall, qall, phong, d_from, d_to)

        st.session_state.dall = dall
        st.session_state.qall = qall
        st.session_state.SA = SA
        st.session_state.SQ = SQ
        st.session_state.df = df
        st.session_state.mm = mm
        st.session_state.tt = tt
        st.session_state.note = note
        st.session_state.phong = phong
        st.session_state.d_from = d_from
        st.session_state.d_to = d_to
        st.session_state.ext_text = read_external(external_report)

dall = st.session_state.dall
qall = st.session_state.qall
df = st.session_state.df
mm = st.session_state.mm
tt = st.session_state.tt
note = st.session_state.note

st.success(note)

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 KPI & Charts MBNT",
    "🌍 KPI & Charts TTQT",
    "🔗 Quan hệ TTQT–MBNT",
    "📋 Báo cáo chuyên sâu",
    "🤖 AI Bot",
    "📤 Export",
])

# ---------- Tab 1: MBNT ----------
with tab1:
    st.subheader("Tổng quan MBNT trong kỳ lọc")
    d_filt, _ = filt(dall, phong, d_from, d_to)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng DS (USD)", f"{d_filt.ds.sum():,.0f}")
    c2.metric("Tổng LN (VND)", f"{d_filt.ln.sum():,.0f}")
    c3.metric("Số giao dịch", f"{len(d_filt):,}")
    c4.metric("Số CIF unique", f"{d_filt.loc[d_filt.cif != '', 'cif'].nunique():,}")

    cif = build(d_filt, ["ds", "ln"])
    TDS, BDS = kpi_pair(cif, "DS")
    TLN, BLN = kpi_pair(cif, "LN")

    col_a, col_b = st.columns(2)
    with col_a:
        fig1 = chart_bar(TDS, "TOP 20% – DOANH SỐ MBNT", "USD", "#2563EB", True, note)
        st.pyplot(fig1, use_container_width=True)
        fig2 = chart_bar(TLN, "TOP 20% – LỢI NHUẬN MBNT", "VND", "#16A34A", True, note)
        st.pyplot(fig2, use_container_width=True)
    with col_b:
        fig3 = chart_bar(BDS, "BOT 10% – DOANH SỐ MBNT", "USD", "#F97316", False, note)
        st.pyplot(fig3, use_container_width=True)
        fig4 = chart_bar(BLN, "BOT 10% – LỢI NHUẬN MBNT", "VND", "#7C3AED", False, note)
        st.pyplot(fig4, use_container_width=True)

    with st.expander("Bảng TOP / BOT chi tiết"):
        st.write("**TOP 20% DS**")
        st.dataframe(TDS.drop(columns=["Key"], errors="ignore"), use_container_width=True)
        st.write("**BOT 10% DS**")
        st.dataframe(BDS.drop(columns=["Key"], errors="ignore"), use_container_width=True)

# ---------- Tab 2: TTQT ----------
with tab2:
    st.subheader("Tổng quan TTQT/TTTM")
    q_filt, _ = filt(qall, phong, d_from, d_to)
    cq = build(q_filt, ["xk", "nk", "xnk", "ds"])
    TQ, BQ = kpi_pair(cq, "DS")

    c1, c2, c3 = st.columns(3)
    c1.metric("Tổng DS TTQT (USD)", f"{q_filt.ds.sum():,.0f}")
    c2.metric("XK", f"{q_filt.xk.sum():,.0f}")
    c3.metric("NK", f"{q_filt.nk.sum():,.0f}")

    col1, col2 = st.columns(2)
    with col1:
        st.pyplot(chart_bar(TQ, "TOP 20% – DOANH SỐ TTQT", "USD", "#0EA5E9", True, note), use_container_width=True)
    with col2:
        st.pyplot(chart_bar(BQ, "BOT 10% – DOANH SỐ TTQT", "USD", "#EF4444", False, note), use_container_width=True)

    # XK / NK separate
    TX, BX = kpi_pair(cq[cq.XK > 0] if "XK" in cq.columns else cq, "XK") if "XK" in cq.columns else (pd.DataFrame(), pd.DataFrame())
    TN, BN = kpi_pair(cq[cq.NK > 0] if "NK" in cq.columns else cq, "NK") if "NK" in cq.columns else (pd.DataFrame(), pd.DataFrame())

    st.markdown("#### Tách chiều XK / NK")
    c1, c2 = st.columns(2)
    with c1:
        if not TX.empty:
            st.pyplot(chart_bar(TX, "TOP 20% – TTQT XK", "USD", "#0EA5E9", True), use_container_width=True)
        if not BX.empty:
            st.pyplot(chart_bar(BX, "BOT 10% – TTQT XK", "USD", "#EF4444", False), use_container_width=True)
    with c2:
        if not TN.empty:
            st.pyplot(chart_bar(TN, "TOP 20% – TTQT NK", "USD", "#14B8A6", True), use_container_width=True)
        if not BN.empty:
            st.pyplot(chart_bar(BN, "BOT 10% – TTQT NK", "USD", "#A855F7", False), use_container_width=True)

# ---------- Tab 3: Relationship ----------
with tab3:
    st.subheader("Quan hệ DS TTQT ↔ DS MBNT")
    if len(df) == 0:
        st.warning("Không có dữ liệu sau lọc.")
    else:
        st.pyplot(scatter_ttqt_mbnt(df), use_container_width=True)

        # Simple group summary
        st.markdown("#### Phân nhóm chuyển đổi")
        for side, X, Y in [("XK → BUY", "XK", "BUY"), ("NK → SELL", "NK", "SELL")]:
            dside = df[df[X] > 0].copy()
            if dside.empty:
                continue
            r = dside[Y] / dside[X]
            lab = ["1. Không GD MBNT", "2. <30%", "3. 30-80%", "4. 80-120%", "5. >120%"]
            dside["Nhom"] = np.select([dside[Y] == 0, r < 0.3, r < 0.8, r <= 1.2], lab[:4], default=lab[4])
            summary = dside.groupby("Nhom").agg(So_cty=("Ten", "size"), Tong_X=(X, "sum"), Tong_Y=(Y, "sum")).reset_index()
            summary["Y/X_%"] = (summary.Tong_Y / summary.Tong_X * 100).round(1)
            st.write(f"**{side}**")
            st.dataframe(summary, use_container_width=True)

# ---------- Tab 4: Deep report (text like A30) ----------
with tab4:
    st.subheader("Báo cáo phân tích chuyên sâu (tự động)")

    def usd(v):
        a = abs(v)
        return f"{a/1e9:,.2f} tỷ USD" if a >= 1e9 else f"{a/1e6:,.1f} triệu USD" if a >= 1e6 else f"{a/1e3:,.0f} nghìn USD" if a >= 1e3 else f"{a:,.0f} USD"

    def pc(a, b):
        return f"{a/b*100:.0f}%" if b > 0 else "0%"

    def phan_tich(X, O, Y, Yv, tx_, ty_):
        L = []
        d = df[df[X] > 0].copy()
        if d.empty:
            return [f"Không có công ty nào có {tx_} trong kỳ đã chọn."], {}
        TX, TY, TV = d[X].sum(), d[Y].sum(), d[Yv].sum()
        r = d[Y] / d[X]
        L.append(f"Có {len(d):,} công ty có {tx_} với tổng {usd(TX)}. Trong số đó {int((d[Y]>0).sum()):,} công ty có giao dịch {ty_} với VCB, tổng {usd(TY)}.")
        L.append(f"Tỷ lệ chuyển đổi chung: cứ 100 USD {tx_} thì VCB nhận được {TY/TX*100:.0f} USD {ty_}.")
        lab = ["Không giao dịch", "Dưới 30%", "Từ 30% đến 80%", "Từ 80% đến 120%", "Trên 120%"]
        d["Nhom"] = np.select([d[Y] == 0, r < 0.3, r < 0.8, r <= 1.2], lab[:4], default=lab[4])
        L.append(f"Phân nhóm theo tỷ lệ {ty_}/{tx_}:")
        for k in lab:
            s = d[d.Nhom == k]
            if len(s):
                L.append(f"   - {k}: {len(s):,} công ty, chiếm {pc(s[X].sum(), TX)} tổng {tx_}.")
        lo = d[d.Nhom.isin(lab[:2])]
        L.append(f"Nhóm dưới 30% (gồm cả không giao dịch) có {len(lo):,} công ty, nắm {pc(lo[X].sum(), TX)} tổng {tx_} nhưng chỉ đóng góp {pc(lo[Y].sum(), TY)} tổng {ty_}.")
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
        f"Dữ liệu: {len(dall):,} dòng MBNT raw → {len(mm):,} dùng phân tích; {len(qall):,} dòng TTQT → {len(tt):,} dùng. {len(df):,} doanh nghiệp (đã loại cá nhân).",
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

    report_text = "\n".join(R)
    st.text_area("Báo cáo tự động", report_text, height=500)

    st.session_state.report_text = report_text

# ---------- Tab 5: AI Bot ----------
with tab5:
    st.subheader("🤖 AI Bot – Viết báo cáo chuyên nghiệp (Groq)")
    st.caption("Sử dụng Llama-3.3-70B (hoặc model khác) để nâng cấp báo cáo, đưa khuyến nghị hành động.")

    context = st.session_state.get("report_text", "")
    if st.session_state.get("ext_text"):
        context += "\n\n=== BÁO CÁO EXTERNAL ĐÃ UPLOAD ===\n" + st.session_state.ext_text[:8000]

    # Quick generate
    if st.button("✨ Sinh báo cáo AI chuyên sâu từ dữ liệu hiện tại", type="primary"):
        with st.spinner("AI đang viết báo cáo…"):
            ai_report = generate_professional_report(context, api_key=groq_key or None)
            st.session_state.ai_report = ai_report

    if "ai_report" in st.session_state:
        st.markdown(st.session_state.ai_report)
        st.download_button(
            "⬇️ Tải báo cáo AI (.md)",
            st.session_state.ai_report,
            file_name=f"Bao_cao_AI_{datetime.now():%Y%m%d_%H%M}.md",
            mime="text/markdown",
        )

    st.divider()
    st.markdown("#### Chat với dữ liệu")
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_q = st.chat_input("Hỏi AI về dữ liệu / yêu cầu viết phần nào đó…")
    if user_q:
        st.session_state.chat_history.append({"role": "user", "content": user_q})
        with st.chat_message("user"):
            st.markdown(user_q)
        with st.chat_message("assistant"):
            with st.spinner("…"):
                ans = chat_with_data(
                    [{"role": m["role"], "content": m["content"]} for m in st.session_state.chat_history],
                    context,
                    api_key=groq_key or None,
                )
                st.markdown(ans)
        st.session_state.chat_history.append({"role": "assistant", "content": ans})

# ---------- Tab 6: Export ----------
with tab6:
    st.subheader("Export kết quả")
    st.download_button(
        "⬇️ Tải bảng doanh nghiệp (df) CSV",
        df.to_csv(index=True).encode("utf-8-sig"),
        file_name="customer_analysis.csv",
        mime="text/csv",
    )
    if "report_text" in st.session_state:
        st.download_button(
            "⬇️ Tải báo cáo text",
            st.session_state.report_text,
            file_name="bao_cao_phan_tich.txt",
            mime="text/plain",
        )

st.markdown("---")
st.caption("VCB Dashboard · Ported from Excel Python · Ready for Render deploy · AI powered by Groq")
