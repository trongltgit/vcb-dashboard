"""
Core processing – port from Excel Python (Z1 / Dashboard / CUSTOMER_*).
Dynamic rows: no fixed 150000 limit; uses full uploaded DataFrame.
TTQT: if XNK (col G) != 0 -> use XNK only; else XK + NK.
"""
import re
import unicodedata as ud
from typing import Tuple

import numpy as np
import pandas as pd

TOP_PCT = 0.20
BOT_PCT = 0.10
SHOW_N = 10
BOT_LOWEST = False
WRAP = 38
MAXLINES = 2

LG = r"\b(cong ty|cty|tnhh|co phan|cp|ctcp|mtv|mot thanh vien|1 thanh vien|jsc|ltd|limited|company|corp|inc|llc)\b"
KW = r"\b(cong ty|ctcp|tnhh|cp|co phan|tap doan|tct|dntn|doanh nghiep|htx|ngan hang|chi nhanh|trung tam|nha may|ltd|corp|inc|llc|jsc|company|group|limited|viet nam|vietnam|thuong mai|san xuat|dau tu|xay dung|van tai|xuat nhap khau|hkd|ho kinh doanh)\b"
KEEP = set()


def N(s):
    return " ".join(str(s).replace("\xa0", " ").split()).lower()


def KB(x):
    x = ud.normalize("NFD", str(x).replace("đ", "d").replace("Đ", "D"))
    return " ".join("".join(c for c in x if ud.category(c) != "Mn").lower().split())


def VN(s: pd.Series) -> pd.Series:
    s = s.astype(object).where(s.notna(), "").astype(str)
    u = {x: KB(x) for x in s.unique()}
    return s.map(u)


def NKY(x):
    return " ".join(re.sub(LG, " ", re.sub(r"[^a-z0-9 ]", " ", KB(x))).split())


def C(s):
    s = ud.normalize("NFD", str(s).replace("đ", "d").replace("Đ", "D"))
    s = "".join(c for c in s if ud.category(c) != "Mn")
    return "".join(w if w.isdigit() else w[0] for w in re.findall(r"[A-Za-z]+|\d+", s)).lower()


def DT(x):
    try:
        if x is None or str(x).strip() in ["", "None", "nan", "NaT"]:
            return None
        if isinstance(x, (int, float)):
            v = float(x)
            if 19000101 < v < 21001231:
                return pd.to_datetime(str(int(v)), format="%Y%m%d")
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=v)
        v = pd.to_datetime(x, dayfirst=True)
        return None if pd.isna(v) else v.normalize()
    except Exception:
        return None


def dates(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    n = pd.to_numeric(s, errors="coerce")
    r = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    m1_ = n.between(19000101, 21001231)
    if m1_.any():
        r[m1_] = pd.to_datetime(n[m1_].astype("int64").astype(str), format="%Y%m%d", errors="coerce")
    m2_ = n.between(20000, 80000)
    if m2_.any():
        r[m2_] = pd.Timestamp("1899-12-30") + pd.to_timedelta(n[m2_], unit="D")
    bad = n.isna() & s.notna()
    if bad.any():
        r[bad] = pd.to_datetime(s[bad].astype(str), dayfirst=True, errors="coerce")
    return r


def num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)
    return (
        pd.to_numeric(
            s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", "", regex=False),
            errors="coerce",
        ).fillna(0)
    )


def T(s: pd.Series) -> pd.Series:
    return (
        s.astype(object)
        .where(s.notna(), "")
        .astype(str)
        .str.strip()
        .replace({"nan": "", "None": "", "NaT": ""})
    )


def VAY(s: pd.Series) -> np.ndarray:
    v = VN(s)
    return (v.str.contains("giai ngan") & v.str.contains("vay")).to_numpy()


def group_keys(d: pd.DataFrame) -> pd.DataFrame:
    d = d.copy()
    d["_n"] = d.ten.map(N)
    P = {}

    def F(a):
        P.setdefault(a, a)
        while P[a] != a:
            P[a] = P[P[a]]
            a = P[a]
        return a

    pr = d[["cif", "_n"]].drop_duplicates()
    for c, n in zip(pr.cif, pr._n):
        if c and n:
            P[F("c:" + c)] = F("n:" + n)
    pr = pr.assign(
        g=[F("c:" + c) if c else (F("n:" + n) if n else "(trống)") for c, n in zip(pr.cif, pr._n)]
    )
    return d.merge(pr, on=["cif", "_n"], how="left")


def LOC(d: pd.DataFrame, p) -> Tuple[pd.DataFrame, str]:
    p = "" if p is None else str(p).strip()
    k = N(p)
    if k in ["", "none", "nan", "vcb hcm"]:
        return d, ""
    for c in ("ph", "cn"):
        if c not in d.columns:
            continue
        ns = [n for n in d[c].unique() if n]
        h = next((n for n in ns if N(n) == k), None)
        if h is None and c == "ph":
            h = next((n for n in ns if C(n) == k.replace(" ", "")), None)
        if h is not None:
            return d[d[c] == h], ""
    return d.iloc[0:0], f"Khong tim thay phong/CN '{p}'"


def filt(src: pd.DataFrame, p, a, b) -> Tuple[pd.DataFrame, str]:
    d, note = LOC(src, p)
    d1, d2 = DT(a), DT(b)
    if d1 is not None:
        d = d[d.ngay >= d1]
    if d2 is not None:
        d = d[d.ngay < d2 + pd.Timedelta(days=1)]
    return d, note


def CN(s: pd.Series) -> np.ndarray:
    u = {
        x: (
            x not in KEEP
            and re.search(KW, x) is None
            and re.fullmatch(r"[a-z]+( [a-z]+){1,4}", x) is not None
        )
        for x in s.unique()
    }
    return s.map(u).to_numpy()


def _col(df, names, idx):
    for n in names:
        if n in df.columns:
            return df[n]
    if idx < len(df.columns):
        return df.iloc[:, idx]
    return pd.Series([None] * len(df))


def load_mbnt(df: pd.DataFrame) -> pd.DataFrame:
    """
    Columns (name or 0-index fallback):
    1 Ngay GD, 3 Ma_CIF, 4 Ten, 5 Phong, 6 Ma CN,
    12 Chieu_NH, 19 DS_Quy_USD, 20 Loi_Nhuan_VND, 22 Nguon, 23 Muc_Dich
    """
    ngay = dates(_col(df, ["Ngày giao dịch", "Ngay_GD", "ngay"], 1))
    cif = T(_col(df, ["Ma_CIF", "CIF"], 3)).str.replace(r"\.0$", "", regex=True)
    ten = T(_col(df, ["Ten_Khach_Hang", "Tên khách hàng"], 4))
    ph = T(_col(df, ["Phong_Ban", "Phòng ban"], 5))
    cn = T(_col(df, ["Mã CN ghi nhận", "Ma_CN"], 6))
    side = T(_col(df, ["Chieu_NH", "Chiều NH"], 12)).str.upper()
    nguon = T(_col(df, ["Nguon_Mua_KH", "Nguồn mua KH"], 22))
    mucdich = T(_col(df, ["Muc_Dich_KH", "Mục đích KH"], 23))
    ds = num(_col(df, ["DS_Quy_USD", "DS"], 19))
    ln = num(_col(df, ["Loi_Nhuan_VND", "LN"], 20))

    d = pd.DataFrame({
        "ngay": ngay, "cif": cif, "ten": ten, "ph": ph, "cn": cn,
        "side": side, "nguon": nguon, "mucdich": mucdich, "ds": ds, "ln": ln,
    })
    d = d[~(d.ngay.isna() & (d.cif == "") & (d.ten == "") & (d.ds == 0) & (d.ln == 0))].reset_index(drop=True)
    d["vw"] = VAY(d.nguon)
    d["vx"] = VAY(d.mucdich)
    d["vay"] = d.vw | d.vx
    return d


def load_ttqt(df: pd.DataFrame) -> pd.DataFrame:
    """
    A Ngay | B CIF | C Ten | D Phong | E XK | F NK | G XNK
    Rule: if G (XNK) != 0 -> use XNK only (xk=0, nk=0, xnk=G)
          else -> xk=E, nk=F, xnk=0
    """
    ngay = dates(_col(df, ["Ngay_GD", "Ngày GD", "ngay"], 0))
    cif = T(_col(df, ["Ma_CIF", "CIF"], 1)).str.replace(r"\.0$", "", regex=True)
    ten = T(_col(df, ["Ten_Khach_Hang", "Tên"], 2))
    ph = T(_col(df, ["Phong_Ban", "Phòng"], 3))
    e1 = num(_col(df, ["DS_Quy_USD (XK)", "XK"], 4))
    f1 = num(_col(df, ["DS_Quy_USD (NK)", "NK"], 5))
    g1 = num(_col(df, ["DS_Quy_USD (XNK)", "XNK"], 6))
    hg = g1 != 0

    q = pd.DataFrame({
        "ngay": ngay, "cif": cif, "ten": ten, "ph": ph,
        "xk": e1.where(~hg, 0),
        "nk": f1.where(~hg, 0),
        "xnk": g1.where(hg, 0),
    })
    q["ds"] = q.xk + q.nk + q.xnk
    q = q[~(q.ngay.isna() & (q.cif == "") & (q.ten == "") & (q.ds == 0))].reset_index(drop=True)
    return q


def build_customer_df(dall, qall, phong, d1, d2):
    mf, note = filt(dall, phong, d1, d2)
    tf, _ = filt(qall, phong, d1, d2)

    mm = mf[mf.gj != "(trống)"].copy()
    tt = tf[tf.gj != "(trống)"].copy()
    mm["kb"] = VN(mm.ten)
    tt["kb"] = VN(tt.ten)

    FL = set(dall.gj[VN(dall.ph).str.contains("the nhan").to_numpy()]) - {"(trống)"}
    xm = mm.gj.isin(FL).to_numpy() | CN(mm.kb)
    xt = tt.gj.isin(FL).to_numpy() | CN(tt.kb)
    nex = len(set(mm.gj[xm]) | set(tt.gj[xt]))
    mm = mm[~xm]
    tt = tt[~xt]

    df = pd.concat([
        tt.groupby("gj")[["xk", "nk", "xnk"]].sum().rename(columns=str.upper),
        mm[mm.side == "BUY"].groupby("gj").ds.sum().rename("BUY"),
        mm[mm.side == "SELL"].groupby("gj").ds.sum().rename("SELL"),
        mm[(mm.side == "BUY") & mm.vay].groupby("gj").ds.sum().rename("BUY_VAY"),
        mm[(mm.side == "SELL") & mm.vay].groupby("gj").ds.sum().rename("SELL_VAY"),
    ], axis=1).fillna(0)

    nmx = pd.concat([mm[["gj", "ten", "cif"]], tt[["gj", "ten", "cif"]]])
    df["Ten"] = df.index.map(
        nmx[nmx.ten != ""].drop_duplicates("gj").set_index("gj").ten
    ).fillna("(khong ten)")
    df["CIF"] = df.index.map(
        nmx[nmx.cif != ""].drop_duplicates("gj").set_index("gj").cif
    ).fillna("")
    df["MB"] = df.BUY + df.SELL
    df["TQ"] = df.XK + df.NK + df.XNK
    df["QH"] = np.select(
        [(df.MB > 0) & (df.TQ > 0), (df.MB > 0) & (df.TQ <= 0), (df.MB <= 0) & (df.TQ > 0)],
        ["MBNT+TTQT", "Chi MBNT", "Chi TTQT"],
        default="Khac",
    )
    return df, mm, tt, note or f"OK: {len(df):,} DN (loai {nex:,} ca nhan)"


def build(d: pd.DataFrame, cols: list) -> pd.DataFrame:
    x = d.assign(
        _c=d.cif.replace("", np.nan),
        _t=d._n.replace("", np.nan) if "_n" in d.columns else d.ten.map(N),
    )
    cif = (
        x.groupby("g")
        .agg(
            Ma_CIF=("_c", "first"),
            So_GD=("g", "size"),
            So_CIF=("_c", "nunique"),
            So_Ten=("_t", "nunique"),
            **{k.upper(): (k, "sum") for k in cols},
        )
        .reset_index()
        .rename(columns={"g": "Key"})
    )
    cif["Ma_CIF"] = cif.Ma_CIF.fillna("")
    nm = (
        d[d.ten != ""]
        .groupby(["g", "ten"])
        .size()
        .reset_index(name="n")
        .sort_values(["g", "n"], ascending=[True, False])
        .drop_duplicates("g")
        .set_index("g")
        .ten
    )
    cif["Ten_Khach_Hang"] = cif.Key.map(nm).fillna("CIF " + cif.Ma_CIF)
    return cif


def pick(z: pd.DataFrame, t: float, p: float) -> pd.DataFrame:
    z = z.copy()
    z["Dong_Gop"] = z.Gia_Tri / t
    z["Luy_Ke"] = z.Dong_Gop.cumsum()
    h = (z.Luy_Ke >= p - 1e-12).to_numpy()
    return z.iloc[: (int(h.argmax()) + 1 if h.any() else len(z))].copy()


def kpi_pair(cif: pd.DataFrame, c: str):
    e = cif.iloc[0:0].copy()
    if cif.empty or cif[c].sum() <= 0:
        return e, e
    t = float(cif[c].sum())
    b = (
        cif.assign(Gia_Tri=cif[c])
        .sort_values(["Gia_Tri", "Key"], ascending=[False, True], kind="mergesort")
        .reset_index(drop=True)
    )
    top = pick(b, t, TOP_PCT)
    rs = b[~b.Key.isin(top.Key)]
    rs = rs.sort_values(["Gia_Tri", "Key"], kind="mergesort").reset_index(drop=True)
    return top, (pick(rs, t, BOT_PCT) if len(rs) else e)


def prepare_all(mbnt_raw: pd.DataFrame, ttqt_raw: pd.DataFrame):
    d = load_mbnt(mbnt_raw)
    q = load_ttqt(ttqt_raw)
    dall = group_keys(d)
    qall = group_keys(q)

    z = pd.concat([dall[["cif", "ten"]], qall[["cif", "ten"]]], ignore_index=True)
    uu = {x: NKY(x) for x in z.ten.unique()}
    z["ten"] = z.ten.map(uu)
    z["cif"] = z.cif.str.lstrip("0")
    z = group_keys(z)
    dall["gj"] = z.g.values[: len(dall)]
    qall["gj"] = z.g.values[len(dall) :]

    SA = dall.groupby("g")[["ds", "ln"]].sum()
    SQ = qall.groupby("g")[["xk", "nk", "xnk", "ds"]].sum()
    return dall, qall, SA, SQ
