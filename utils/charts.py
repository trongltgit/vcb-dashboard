"""
Chart helpers matching original matplotlib style (horizontal bars + pie gauges).
"""
import textwrap
from typing import Optional

import matplotlib.pyplot as plt
import matplotlib.ticker as mt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D

from .processing import SHOW_N, BOT_LOWEST, WRAP, MAXLINES


def fmt(v, u):
    s = "-" if v < 0 else ""
    a = abs(v)
    if u == "USD":
        return f"{s}{a/1e6:,.2f}M" if a >= 1e6 else f"{s}{a/1e3:,.1f}K" if a >= 1e3 else f"{s}{a:,.0f}"
    return f"{s}{a/1e9:,.2f} tỷ" if a >= 1e9 else f"{s}{a/1e6:,.1f} tr" if a >= 1e6 else f"{s}{a:,.0f}"


def wrap(s):
    p = textwrap.wrap(str(s), WRAP, break_long_words=True)
    if len(p) > MAXLINES:
        p = p[: MAXLINES - 1] + [p[MAXLINES - 1][: WRAP - 1] + "…"]
    return "\n".join(p)


def chart_bar(z: pd.DataFrame, title: str, u: str, col: str, top: bool, note: str = ""):
    """Horizontal bar chart; avoid tight_layout because it can be expensive on Render."""
    fig, ax = plt.subplots(figsize=(5.4, 2.85), dpi=90)

    if z is None or z.empty:
        ax.text(
            0.5, 0.5, note or "Không có dữ liệu",
            ha="center", va="center", fontsize=8, color="#6B7280",
            transform=ax.transAxes,
        )
        ax.axis("off")
        fig.subplots_adjust(left=0.08, right=0.98, top=0.86, bottom=0.14)
        return fig

    # Only the chart is capped. The source DataFrame and KPI/ranking tables are untouched.
    show_n = max(1, int(SHOW_N))
    zz = (
        z.sort_values("Gia_Tri", ascending=(not top and BOT_LOWEST))
        .head(show_n)
        .sort_values("Gia_Tri")
        .copy()
    )
    v = pd.to_numeric(zz["Gia_Tri"], errors="coerce").fillna(0).to_numpy(dtype=float)
    y = list(range(len(v)))
    bars = ax.barh(y, v, color=col, height=0.55)

    ax.set_yticks(y)
    ax.set_yticklabels(zz["Ten_Khach_Hang"].map(wrap).tolist())

    lo, hi = min(0.0, float(v.min())), max(0.0, float(v.max()))
    sp = (hi - lo) or max(abs(hi), abs(lo), 1.0)

    shares = pd.to_numeric(zz["Dong_Gop"], errors="coerce").fillna(0).to_numpy(dtype=float)
    for b, x, sh in zip(bars, v, shares):
        offset = sp * 0.015
        ax.text(
            x + (offset if x >= 0 else -offset),
            b.get_y() + b.get_height() / 2,
            f"{fmt(x, u)} ({sh * 100:.1f}%)",
            va="center",
            ha="left" if x >= 0 else "right",
            fontsize=5.5,
            color="#374151",
            clip_on=True,
        )

    ax.set_xlim(
        lo - (sp * 0.3 if lo < 0 else 0),
        hi + sp * 0.34 if hi >= 0 else hi + sp * 0.05,
    )
    ax.set_title(title, loc="left", fontsize=8, fontweight="bold", color=col, pad=14)
    ax.text(
        0,
        1.03,
        f"Nhóm {len(z):,} công ty = {pd.to_numeric(z['Dong_Gop'], errors='coerce').fillna(0).sum() * 100:.1f}% tổng · hiển thị {len(zz)}",
        transform=ax.transAxes,
        fontsize=5.5,
        color="#6B7280",
        va="bottom",
    )
    ax.tick_params(axis="y", labelsize=4.8, length=0, pad=2)
    ax.tick_params(axis="x", labelsize=5.5, colors="#6B7280")
    ax.xaxis.set_major_locator(mt.MaxNLocator(4))
    ax.xaxis.set_major_formatter(mt.FuncFormatter(lambda x, _: fmt(x, u)))
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color("#D1D5DB")
    ax.grid(axis="x", alpha=0.25)
    ax.set_axisbelow(True)

    # Fixed margins avoid Matplotlib's expensive tight_layout label-measurement pass.
    ax.margins(y=0.08)
    fig.subplots_adjust(left=0.30, right=0.98, top=0.80, bottom=0.18)
    return fig


def chart_plan_completion(
    grid,
    metrics,
    loc: str,
    d1,
    d2,
    q_label: str,
    y_label: str,
    bg: str = "#14532D",
):
    """
    3x6 pie gauge chart matching original K12 style.
    grid: list of 3 metrics, each with 6 (th, kh, pct) tuples
    """
    COL_PHONG, COL_CN, COL_TSC = "#38BDF8", "#4ADE80", "#FB7C3C"
    TRACK = "#2F6B47"
    SUB = [("KH phòng", COL_PHONG), ("KH CN giao", COL_CN), ("KH TSC giao", COL_TSC)]
    COLS = SUB + SUB

    title = f"HOÀN THÀNH KẾ HOẠCH – {loc or 'VCB HCM'}\n{d1:%d/%m/%Y} → {d2:%d/%m/%Y}"
    LEFT, RIGHT = 0.05, 0.995
    fig, ax = plt.subplots(3, 6, figsize=(12, 6.0), facecolor=bg)
    fig.suptitle(title, color="white", fontsize=11, fontweight="bold", y=0.985)

    for k, m in enumerate(metrics):
        for c in range(6):
            a = ax[k, c]
            a.set_facecolor(bg)
            cname, color = COLS[c]
            th, kh, pct = grid[k][c]
            if kh <= 0:
                a.pie([1], colors=[TRACK], startangle=90, wedgeprops={"width": 0.28, "linewidth": 0})
                a.text(0, 0.08, "N/A", ha="center", va="center", color="#A7C4B2", fontsize=9, fontweight="bold")
                a.text(0, -0.25, "chưa có KH", ha="center", va="center", color="#A7C4B2", fontsize=5)
            else:
                fill = max(0.0, min(1.0, pct))
                vals, cols = ([1], [color]) if fill >= 1 else ([fill, 1 - fill], [color, TRACK])
                a.pie(vals, colors=cols, startangle=90, counterclock=False, wedgeprops={"width": 0.28, "linewidth": 0})
                a.text(0, 0.10, f"{pct * 100:.2f}%", ha="center", va="center", color="white", fontsize=9.5, fontweight="bold")
                a.text(0, -0.25, f"TH {th:,.1f}\nKH {kh:,.1f}", ha="center", va="center", color="#D1FAE5", fontsize=5, linespacing=1.2)
            if k == 0:
                a.set_title(cname, color=color, fontsize=7.5, fontweight="bold", pad=2)
            if c == 0:
                a.text(-1.6, 0, f'{m["name"]} ({m["unit"]})', rotation=90, ha="center", va="center", color="white", fontsize=8, fontweight="bold")
            a.set_aspect("equal")

    W_ = RIGHT - LEFT
    fig.text(LEFT + 0.25 * W_ + 0.01, 0.915, f"SO VỚI KH QUÝ ({q_label})", color="white", fontsize=9, fontweight="bold", ha="center")
    fig.text(LEFT + 0.75 * W_ + 0.01, 0.915, f"SO VỚI KH NĂM ({y_label})", color="white", fontsize=9, fontweight="bold", ha="center")
    fig.add_artist(Line2D([LEFT + 0.5 * W_ + 0.01] * 2, [0.02, 0.90], color="#86EFAC", linewidth=0.8, alpha=0.6))
    fig.subplots_adjust(top=0.86, bottom=0.015, left=LEFT, right=RIGHT, wspace=0.05, hspace=0.12)
    return fig


def scatter_ttqt_mbnt(df: pd.DataFrame):
    """Scatter XK vs BUY and NK vs SELL."""
    fig, axs = plt.subplots(1, 2, figsize=(9, 3.6), dpi=90)

    def fm(v):
        a = abs(v)
        return f"{a / 1e6:,.1f}M" if a >= 1e6 else f"{a / 1e3:,.0f}K" if a >= 1e3 else f"{a:,.0f}"

    for ax, (X, Y, c) in zip(axs, [("XK", "BUY", "#0EA5E9"), ("NK", "SELL", "#14B8A6")]):
        pp = df[(df[X] > 0) & (df[Y] > 0)]
        if len(pp) > 2:
            u, w = np.log10(pp[X]), np.log10(pp[Y])
            ax.scatter(u, w, s=8, alpha=0.5, color=c)
            xs = np.linspace(u.min(), u.max(), 20)
            ax.plot(xs, xs, "--", color="#9CA3AF", lw=0.8)
            ff = mt.FuncFormatter(lambda v, _: fm(10 ** v))
            ax.xaxis.set_major_formatter(ff)
            ax.yaxis.set_major_formatter(ff)
        ax.set_title(f"{X} TTQT vs MBNT {Y} (n={len(pp)}; nét đứt: y = x)", fontsize=7, loc="left")
        ax.tick_params(labelsize=6)
        ax.spines[["top", "right"]].set_visible(False)

    fig.subplots_adjust(left=0.08, right=0.98, top=0.82, bottom=0.18, wspace=0.25)
    return fig
