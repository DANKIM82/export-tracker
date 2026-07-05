# -*- coding: utf-8 -*-
"""반도체 수출 트래커 데이터 빌더: cache CSV -> semi_export_data.js (SEMI_EXPORT_DATA).

semicon_tracker.py 의 run() 끝에서 build() 가 자동 호출되어, 트래커 한 번 실행으로
캐시 갱신 + 대시보드 JS 생성이 한 단계에 끝난다. 단독 실행도 가능:  python build_dashboard_data.py
"""
import json, os
from datetime import date
import pandas as pd
import numpy as np

_HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(_HERE, "data")
# industry-dashboard/data/semi_export_data.js (상대경로 우선, 없으면 절대경로 폴백)
_REL = os.path.normpath(os.path.join(_HERE, "..", "..", "..",
        "Project_industry_analyzer", "industry-dashboard", "data"))
_ABS = r"c:\!Workspace\Project\Project_industry_analyzer\industry-dashboard\data"
OUT_DIR = _REL if os.path.isdir(_REL) else _ABS
OUT = os.path.join(OUT_DIR, "semi_export_data.js")

WINDOW_START = "202201"            # 차트/표 표시 시작월
HBM = "8542323000"
KEY_COUNTRIES = ["TW", "US", "CN", "HK", "VN", "SG", "JP"]
CN_NAME = {"TW": "대만", "US": "미국", "CN": "중국", "HK": "홍콩",
           "VN": "베트남", "SG": "싱가포르", "JP": "일본"}


def jnull(x):
    if x is None:
        return None
    if isinstance(x, float) and (np.isnan(x) or np.isinf(x)):
        return None
    return x


def r2(x, n=2):
    x = jnull(x)
    return None if x is None else round(float(x), n)


def mom_yoy(series_by_month):
    """dict {yyyymm: value} -> dict {yyyymm: (mom%, yoy%)}."""
    ks = sorted(series_by_month)
    out = {}
    for i, k in enumerate(ks):
        v = series_by_month[k]
        prev = series_by_month[ks[i - 1]] if i >= 1 else None
        ky = f"{int(k[:4]) - 1}{k[4:]}"
        py = series_by_month.get(ky)
        mom = (v / prev - 1) * 100 if (v is not None and prev not in (None, 0)) else None
        yoy = (v / py - 1) * 100 if (v is not None and py not in (None, 0)) else None
        out[k] = (mom, yoy)
    return out


def build(verbose=True):
    """캐시 CSV를 읽어 대시보드용 JS 파일을 생성한다."""
    def _log(*a):
        if verbose:
            print(*a)

    cur_ym = date.today().strftime("%Y%m")  # 당월(미집계) 제외

    # ---------- load memory cache ----------
    mem = pd.read_csv(os.path.join(SRC, "cache_memory.csv"), dtype=str)
    for c in ["exp_usd", "exp_wgt", "imp_usd", "imp_wgt"]:
        mem[c] = pd.to_numeric(mem[c], errors="coerce")
    mem = mem[mem["yyyymm"] < cur_ym]
    mem = mem[mem["yyyymm"] >= "201901"]

    months_all = sorted(mem["yyyymm"].unique())
    months = [m for m in months_all if m >= WINDOW_START]
    labels = [f"{m[:4]}-{m[4:6]}" for m in months]
    latest = months[-1]

    def series_full(df, valcol):
        g = df.groupby("yyyymm")[valcol].sum()
        return {k: (None if pd.isna(v) else float(v)) for k, v in g.items()}

    # ===== 1) 전체 메모리 수출 =====
    tot_usd = series_full(mem, "exp_usd")
    tot_my = mom_yoy(tot_usd)
    total_series = [r2(tot_usd.get(m, 0) / 1e6, 1) for m in months]
    total_yoy = [r2(tot_my.get(m, (None, None))[1], 1) for m in months]
    total_mom = [r2(tot_my.get(m, (None, None))[0], 1) for m in months]

    # ===== 2) 제품 믹스 =====
    prod = mem.groupby(["yyyymm", "category"])[["exp_usd", "exp_wgt"]].sum().reset_index()
    cats = ["DRAM", "NAND Flash", "HBM/Advanced (복합구조칩)", "SRAM",
            "Other Memory", "Hybrid IC", "MCOs (복합부품)"]
    CAT_LABEL = {"DRAM": "DRAM", "NAND Flash": "NAND", "HBM/Advanced (복합구조칩)": "HBM/첨단패키징",
                 "SRAM": "SRAM", "Other Memory": "기타메모리", "Hybrid IC": "하이브리드IC",
                 "MCOs (복합부품)": "MCOs"}
    product = {}
    for cat in cats:
        sub = prod[prod["category"] == cat]
        by = {r["yyyymm"]: r["exp_usd"] for _, r in sub.iterrows()}
        bywt = {r["yyyymm"]: r["exp_wgt"] for _, r in sub.iterrows()}
        my = mom_yoy({k: (None if pd.isna(v) else v) for k, v in by.items()})
        product[cat] = {
            "label": CAT_LABEL[cat],
            "musd": [r2((by.get(m) or 0) / 1e6, 1) for m in months],
            "yoy": [r2(my.get(m, (None, None))[1], 1) for m in months],
            "mom": [r2(my.get(m, (None, None))[0], 1) for m in months],
            "asp": [(lambda u, w: r2(u / w, 0) if (u and w) else None)(by.get(m), bywt.get(m)) for m in months],
        }

    # ===== 3) ASP 지수 =====
    asp_index = {}
    for cat in ["DRAM", "NAND Flash", "HBM/Advanced (복합구조칩)"]:
        raw = product[cat]["asp"]
        base = next((v for v in raw if v), None)
        asp_index[cat] = {"label": CAT_LABEL[cat],
                          "idx": [(r2(v / base * 100, 1) if (v and base) else None) for v in raw],
                          "asp": raw}

    # ===== 4) 국가별 =====
    ctry = mem.groupby(["yyyymm", "country_code", "country_name"])[["exp_usd", "exp_wgt"]].sum().reset_index()
    tot_by_month = mem.groupby("yyyymm")["exp_usd"].sum().to_dict()
    snap = ctry[ctry["yyyymm"] == latest].copy()
    snap["share"] = snap["exp_usd"] / tot_by_month[latest] * 100

    def cseries(cc):
        s = ctry[ctry["country_code"] == cc]
        return {r["yyyymm"]: r["exp_usd"] for _, r in s.iterrows()}

    country_rows = []
    snap = snap.sort_values("exp_usd", ascending=False)
    for _, r in snap.iterrows():
        cc = r["country_code"]
        sm = cseries(cc)
        my = mom_yoy({k: (None if pd.isna(v) else v) for k, v in sm.items()})
        country_rows.append({
            "code": cc, "name": r["country_name"],
            "musd": r2(r["exp_usd"] / 1e6, 1), "share": r2(r["share"], 1),
            "asp": (r2(r["exp_usd"] / r["exp_wgt"], 0) if r["exp_wgt"] else None),
            "mom": r2(my.get(latest, (None, None))[0], 1), "yoy": r2(my.get(latest, (None, None))[1], 1),
            "key": cc in KEY_COUNTRIES,
        })
    country_rows.sort(key=lambda x: (KEY_COUNTRIES.index(x["code"]) if x["code"] in KEY_COUNTRIES else 99,
                                     -(x["musd"] or 0)))
    country_rows = country_rows[:14]

    key_country_ts = {}
    for cc in KEY_COUNTRIES:
        sm = cseries(cc)
        key_country_ts[cc] = {"name": CN_NAME[cc], "musd": [r2((sm.get(m) or 0) / 1e6, 1) for m in months]}
    cn_hk = [r2(((cseries("CN").get(m) or 0) + (cseries("HK").get(m) or 0)) / 1e6, 1) for m in months]

    # ===== 5) HBM 프록시 =====
    hbm = mem[mem["hs10"] == HBM]
    hbm_tot = series_full(hbm, "exp_usd")
    hbm_my = mom_yoy(hbm_tot)
    hbm_series = [r2((hbm_tot.get(m) or 0) / 1e6, 1) for m in months]
    hbm_yoy = [r2(hbm_my.get(m, (None, None))[1], 1) for m in months]
    hsnap = hbm[hbm["yyyymm"] == latest].groupby(["country_code", "country_name"])[["exp_usd", "exp_wgt"]].sum().reset_index()
    hbm_tot_latest = hsnap["exp_usd"].sum()
    hsnap = hsnap.sort_values("exp_usd", ascending=False).head(8)
    hbm_country = []
    for _, r in hsnap.iterrows():
        hbm_country.append({"name": r["country_name"], "musd": r2(r["exp_usd"] / 1e6, 2),
                            "share": r2(r["exp_usd"] / hbm_tot_latest * 100, 1) if hbm_tot_latest else None,
                            "asp": (r2(r["exp_usd"] / r["exp_wgt"], 0) if r["exp_wgt"] else None)})
    hbm_ctry_ts = {}
    for cc in ["US", "TW", "HK", "CN", "VN"]:
        s = hbm[hbm["country_code"] == cc].groupby("yyyymm")["exp_usd"].sum()
        sm = {k: v for k, v in s.items()}
        hbm_ctry_ts[cc] = {"name": CN_NAME[cc], "musd": [r2((sm.get(m) or 0) / 1e6, 2) for m in months]}

    # ===== 6) ICT =====
    ict = pd.read_csv(os.path.join(SRC, "cache_ict.csv"), dtype=str)
    ict["exp_usd"] = pd.to_numeric(ict["exp_usd"], errors="coerce")
    ict = ict[ict["yyyymm"] < cur_ym]
    ict_out = {}
    for item in ["반도체", "디스플레이", "휴대폰"]:
        sub = ict[ict["item"] == item]
        by = {r["yyyymm"]: r["exp_usd"] for _, r in sub.iterrows()}
        my = mom_yoy({k: (None if pd.isna(v) else v) for k, v in by.items()})
        ict_out[item] = {"musd": [(r2((by.get(m)) / 1e6, 0) if by.get(m) and not pd.isna(by.get(m)) else None) for m in months],
                         "yoy": [r2(my.get(m, (None, None))[1], 1) for m in months]}

    # ===== 7) 장비 capex =====
    eq = pd.read_csv(os.path.join(SRC, "cache_equip.csv"), dtype=str)
    eq["imp_usd"] = pd.to_numeric(eq["imp_usd"], errors="coerce")
    eq = eq[eq["yyyymm"] < cur_ym]
    equip = {}
    for cat in ["반도체장비(전공정)", "장비부품·소모성"]:
        sub = eq[eq["category"] == cat]
        by = {r["yyyymm"]: r["imp_usd"] for _, r in sub.iterrows()}
        my = mom_yoy({k: (None if pd.isna(v) else v) for k, v in by.items()})
        ser = [(r2((by.get(m)) / 1e6, 0) if by.get(m) and not pd.isna(by.get(m)) else None) for m in months]
        ma3 = []
        for i in range(len(ser)):
            win = [v for v in ser[max(0, i - 2):i + 1] if v is not None]
            ma3.append(r2(sum(win) / len(win), 0) if win else None)
        equip[cat] = {"musd": ser, "ma3": ma3, "yoy": [r2(my.get(m, (None, None))[1], 1) for m in months]}

    eqc = pd.read_csv(os.path.join(SRC, "cache_equip_country.csv"), dtype=str)
    eqc["imp_usd"] = pd.to_numeric(eqc["imp_usd"], errors="coerce")
    eqc = eqc[eqc["yyyymm"] < cur_ym]
    eq_latest = sorted(eqc["yyyymm"].unique())[-1]
    ec = eqc[eqc["yyyymm"] == eq_latest].groupby(["country_code", "country_name"])["imp_usd"].sum().reset_index()
    ec_tot = ec["imp_usd"].sum()
    ec = ec.sort_values("imp_usd", ascending=False).head(8)
    equip_country = [{"name": r["country_name"], "musd": r2(r["imp_usd"] / 1e6, 1),
                      "share": r2(r["imp_usd"] / ec_tot * 100, 1) if ec_tot else None} for _, r in ec.iterrows()]

    # ===== 8) FRED PPI =====
    fred = pd.read_csv(os.path.join(SRC, "fred_PCU334413334413.csv"), dtype=str)
    fred["value"] = pd.to_numeric(fred["value"], errors="coerce")
    fred = fred.dropna(subset=["value"])
    fred = fred[fred["yyyymm"] >= "202201"]
    fred_labels = [f"{r['yyyymm'][:4]}-{r['yyyymm'][4:6]}" for _, r in fred.iterrows()]
    fred_val = [r2(v, 1) for v in fred["value"].tolist()]
    fred_yoy = [r2(v, 1) for v in pd.to_numeric(fred["yoy_pct"], errors="coerce").tolist()]

    DATA = {
        "meta": {"latest_month": f"{latest[:4]}-{latest[4:6]}", "generated": date.today().isoformat(),
                 "eq_latest": f"{eq_latest[:4]}-{eq_latest[4:6]}"},
        "months": months, "labels": labels,
        "total": {"musd": total_series, "yoy": total_yoy, "mom": total_mom},
        "product": product,
        "asp_index": asp_index,
        "country_rows": country_rows, "key_country_ts": key_country_ts, "cn_hk": cn_hk,
        "hbm": {"series": hbm_series, "yoy": hbm_yoy, "country": hbm_country, "country_ts": hbm_ctry_ts},
        "ict": ict_out,
        "equip": equip, "equip_country": equip_country,
        "fred": {"labels": fred_labels, "value": fred_val, "yoy": fred_yoy},
    }

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("// 자동생성 — semicon_tracker 캐시 기반 반도체 수출 트래커 데이터\n")
        f.write("window.SEMI_EXPORT_DATA = ")
        json.dump(DATA, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")
    _log(f"대시보드 JS 생성: {OUT}  (기준월 {latest[:4]}-{latest[4:6]}, {len(months)}개월)")
    return OUT


if __name__ == "__main__":
    build()
