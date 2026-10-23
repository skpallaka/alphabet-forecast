"""Day 5: forward forecasts, scenario numbers, charts and the 4-slide leadership deck.

Scenario maths here mirrors the Excel formulas (verified to ~1e-11 against the workbook, see README).
"""
import sqlite3
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt

from models import ets, sarima

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output"
CH = OUT / "charts"
CH.mkdir(parents=True, exist_ok=True)

SEGS = {"rev_search": "Search & other", "rev_youtube": "YouTube ads", "rev_network": "Network",
        "rev_subscriptions": "Subscriptions, platforms & devices", "rev_cloud": "Google Cloud", "rev_other_bets": "Other Bets"}
COL = dict(zip(SEGS, ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"]))
INK, MUTED, GRID = "#0b0b0b", "#52514e", "#e6e5e1"
FADE = 0.9
DELTA = {"Bear": {"rev_search": -.02, "rev_youtube": -.02, "rev_cloud": -.05},
         "Base": {}, "Bull": {"rev_search": .02, "rev_youtube": .02, "rev_cloud": .05}}
N_FC = 6

plt.rcParams.update({"font.family": "sans-serif", "axes.edgecolor": GRID, "axes.labelcolor": MUTED,
                     "xtick.color": MUTED, "ytick.color": MUTED, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.grid": True, "grid.color": GRID, "grid.linewidth": .8,
                     "axes.axisbelow": True, "font.size": 10})


def load():
    con = sqlite3.connect(ROOT / "data" / "alphabet.db")
    return pd.read_sql("select * from v_pnl where quarter_end>='2019-03-31'", con, parse_dates=["quarter_end"]).set_index("quarter_end")


def scenario(df, sc, base_fixed=None):
    """Returns DataFrame of forecast quarters: segment revenue, total, costs, operating income."""
    idx = pd.date_range(df.index[-1], periods=N_FC + 1, freq="QE")[1:]
    seg = {}
    for s in SEGS:
        h = list(df[s]); g = h[-1] / h[-5] - 1
        for k in range(1, N_FC + 1):
            h.append(h[-4] * (1 + g * FADE ** k + DELTA[sc].get(s, 0)))
        seg[s] = h[len(df):]
    fc = pd.DataFrame(seg, index=idx)
    fc["revenue"] = fc.sum(axis=1)
    rev_all = list(df.revenue) + list(fc.revenue)
    for key, col in (("cor", "cost_of_revenues"), ("rd", "rd"), ("sm", "sales_marketing"), ("ga", "g_and_a")):
        r = list((df[col] / df.revenue))
        for i in range(N_FC):
            r.append(np.mean(r[-4:]))
        fc[key] = [rev_all[len(df) + i] * r[len(df) + i] for i in range(N_FC)]
    if base_fixed is not None:
        fc["rd"], fc["ga"] = base_fixed["rd"].values, base_fixed["ga"].values
    fc["operating_income"] = fc.revenue - fc.cor - fc.rd - fc.sm - fc.ga
    return fc


def fy(df, fc, year, col):
    hist = df[df.index.year == year][col] if col in df else pd.Series(dtype=float)
    f = fc[fc.index.year == year][col] if col in fc else pd.Series(dtype=float)
    return hist.sum() + f.sum()


def savefig(name):
    plt.savefig(CH / name, dpi=200, bbox_inches="tight", facecolor="#fcfcfb")
    plt.close()


def main():
    df = load()
    base = scenario(df, "Base")
    scen = {"Base": base, "Bull": scenario(df, "Bull", base), "Bear": scenario(df, "Bear", base)}

    # ---- forward statistical forecasts vs Excel growth model -----------------------------------
    rows = []
    for s, name in SEGS.items():
        if s == "rev_other_bets":
            continue
        y = df[s].values
        for mname, f in (("ETS", ets), ("SARIMA", sarima)):
            for q, v in zip(base.index, f(y, N_FC)):
                rows.append((name, mname, q.date(), v))
        for q, v in zip(base.index, base[s]):
            rows.append((name, "Excel growth model (Base)", q.date(), v))
    fwd = pd.DataFrame(rows, columns=["segment", "model", "quarter", "forecast_musd"])
    fwd.to_csv(OUT / "forward_forecasts.csv", index=False)
    cmp27 = fwd[pd.to_datetime(fwd.quarter).dt.year == 2027].groupby(["segment", "model"]).forecast_musd.sum().unstack().round(0)
    cmp27.to_csv(OUT / "fy27_model_comparison.csv")
    print("FY27E revenue by segment and model ($M):\n", cmp27.to_string())

    # ---- scenario table --------------------------------------------------------------------------
    tab = []
    for sc, fc in scen.items():
        for year in (2026, 2027):
            rev, oi, cloud = fy(df, fc, year, "revenue"), fy(df, fc, year, "operating_income"), fy(df, fc, year, "rev_cloud")
            tab.append(dict(case=sc, fy=year, cloud=cloud, revenue=rev, operating_income=oi, op_margin=oi / rev))
    tab = pd.DataFrame(tab)
    tab.round(3).to_csv(OUT / "scenarios.csv", index=False)
    print(tab.round(3).to_string())
    fy25_rev, fy25_oi = df[df.index.year == 2025].revenue.sum(), df[df.index.year == 2025].operating_income.sum()

    # ---- charts ----------------------------------------------------------------------------------
    bt = pd.read_csv(OUT / "mape_by_series_model.csv", index_col=0)
    qx = list(df.index) + list(base.index)
    labels = [f"Q{(d.month - 1) // 3 + 1}'{str(d.year)[2:]}" for d in qx]

    # 1) revenue mix, stacked columns, actual + base forecast
    fig, ax = plt.subplots(figsize=(11, 4.6))
    bottom = np.zeros(len(qx))
    for s, name in SEGS.items():
        if s == "rev_other_bets":
            continue
        vals = np.array(list(df[s]) + list(base[s])) / 1000
        ax.bar(range(len(qx)), vals, bottom=bottom, color=COL[s], width=.8, edgecolor="#fcfcfb", linewidth=1, label=name)
        bottom += vals
    ax.axvspan(len(df) - .5, len(qx) - .5, color="#f0efec", alpha=.7, zorder=0)
    ax.text(len(df) + N_FC / 2 - .5, bottom.max() * 1.02, "Base-case forecast", ha="center", color=MUTED, fontsize=9)
    ax.set_xticks(range(0, len(qx), 2)); ax.set_xticklabels(labels[::2], rotation=0, fontsize=8)
    ax.set_ylabel("Revenue, $B per quarter"); ax.grid(axis="x", visible=False)
    ax.legend(ncol=5, loc="upper left", frameon=False, fontsize=8, bbox_to_anchor=(0, 1.12))
    savefig("01_revenue_mix.png")

    # 2) Cloud: actual + three models + bull/bear
    fig, ax = plt.subplots(figsize=(7.2, 4.4))
    cl = df.rev_cloud / 1000
    ax.plot(cl.index, cl, color=INK, lw=2, label="Actual")
    last = (cl.index[-1], cl.iloc[-1])          # anchor forecast lines to the last actual so there is no visual gap
    ax.fill_between(base.index, scen["Bear"].rev_cloud / 1000, scen["Bull"].rev_cloud / 1000, color=COL["rev_cloud"], alpha=.18, lw=0)
    ax.plot([last[0], *base.index], [last[1], *(base.rev_cloud / 1000)], color=COL["rev_cloud"], lw=2, label="Excel growth model: base (bear–bull band)")
    for m, c in (("ETS", "#6250d6"), ("SARIMA", "#eb6834")):
        d = fwd[(fwd.segment == "Google Cloud") & (fwd.model == m)]
        ax.plot([last[0], *pd.to_datetime(d.quarter)], [last[1], *(d.forecast_musd / 1000)], color=c, lw=1.6, ls="--", label=m)
    ax.set_ylabel("Google Cloud revenue, $B per quarter"); ax.legend(frameon=False, fontsize=8, loc="upper left")
    savefig("02_cloud_forecast.png")

    # 3) MAPE heatmap
    show = ["Naive", "Seasonal naive", "YoY growth (last)", "YoY growth (faded)", "ETS", "SARIMA"]
    m = bt[show]
    fig, ax = plt.subplots(figsize=(7.4, 3.6))
    im = ax.imshow(m.values, cmap=matplotlib.colors.LinearSegmentedColormap.from_list("b", ["#eaf2fc", "#2a78d6", "#12386b"]), aspect="auto", vmin=0, vmax=26)
    ax.set_xticks(range(len(show))); ax.set_xticklabels([s.replace(" (", "\n(") .replace("Seasonal naive", "Seasonal\nnaive") for s in show], fontsize=8)
    ax.set_yticks(range(len(m))); ax.set_yticklabels([i.replace("Subscriptions, platforms & devices", "Subscriptions & devices") for i in m.index], fontsize=9)
    ax.grid(False); ax.xaxis.tick_top()
    for i in range(m.shape[0]):
        for j in range(m.shape[1]):
            v = m.values[i, j]
            ax.text(j, i, f"{v:.1f}", ha="center", va="center", fontsize=9, color="white" if v > 11 else INK,
                    fontweight="bold" if v == m.values[i].min() else "normal")
    for sp in ax.spines.values():
        sp.set_visible(False)
    savefig("03_backtest_mape.png")

    # 4) scenarios: operating income FY25A / FY26E / FY27E
    fig, axes = plt.subplots(1, 2, figsize=(8.6, 3.9))
    for ax, col, ttl, scale in ((axes[0], "operating_income", "Operating income, $B", 1000), (axes[1], "cloud", "Google Cloud revenue, $B", 1000)):
        xs = np.arange(2)
        for k, (sc, c) in enumerate((("Bear", "#eb6834"), ("Base", "#2a78d6"), ("Bull", "#1baf7a"))):
            vals = [tab[(tab.case == sc) & (tab.fy == y)][col].iloc[0] / scale for y in (2026, 2027)]
            bars = ax.bar(xs + (k - 1) * .27, vals, .25, color=c, label=sc)
            for b, v in zip(bars, vals):
                ax.text(b.get_x() + b.get_width() / 2, v + max(vals) * .01, f"{v:,.0f}", ha="center", fontsize=8, color=INK)
        ax.set_xticks(xs); ax.set_xticklabels(["FY2026E", "FY2027E"]); ax.set_title(ttl, loc="left", fontsize=10, color=INK)
        ax.grid(axis="x", visible=False)
    axes[0].legend(frameon=False, fontsize=8, loc="upper left")
    plt.tight_layout()
    savefig("04_scenarios.png")

    # ---- deck ---------------------------------------------------------------------------------------
    b26 = tab[(tab.case == "Base") & (tab.fy == 2026)].iloc[0]
    b27 = tab[(tab.case == "Base") & (tab.fy == 2027)].iloc[0]
    bear27 = tab[(tab.case == "Bear") & (tab.fy == 2027)].iloc[0]
    bull27 = tab[(tab.case == "Bull") & (tab.fy == 2027)].iloc[0]
    q_cloud_share = df.rev_cloud.iloc[-1] / df.revenue.iloc[-1]
    prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def slide(title, bullets, img, img_w, notes=None):
        s = prs.slides.add_slide(blank)
        tb = s.shapes.add_textbox(Inches(.6), Inches(.4), Inches(12.1), Inches(1.1)); tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = RGBColor(0x0b, 0x0b, 0x0b)
        left = Inches(.6)
        if bullets:
            bb = s.shapes.add_textbox(left, Inches(1.8), Inches(4.2), Inches(5)); bf = bb.text_frame; bf.word_wrap = True
            for i, t in enumerate(bullets):
                p = bf.paragraphs[0] if i == 0 else bf.add_paragraph()
                p.text = t; p.font.size = Pt(15); p.font.color.rgb = RGBColor(0x2b, 0x2b, 0x2b); p.space_after = Pt(12)
        if img:
            x = Inches(5.0) if bullets else Inches(.6)
            s.shapes.add_picture(str(CH / img), x, Inches(1.7), width=Inches(img_w))
        ft = s.shapes.add_textbox(Inches(.6), Inches(7.0), Inches(12), Inches(.3)); fp = ft.text_frame.paragraphs[0]
        fp.text = "Source: SEC EDGAR XBRL filings (10-K/10-Q), author's models. Forecasts are illustrative, not investment advice."
        fp.font.size = Pt(10); fp.font.color.rgb = RGBColor(0x75, 0x75, 0x70)
        return s

    slide(f"Alphabet is forecast to reach ~${b26.revenue / 1000:,.0f}B revenue in FY26 and ~${b27.revenue / 1000:,.0f}B in FY27, led by Cloud",
          [f"Q2'26 revenue ${df.revenue.iloc[-1] / 1000:,.1f}B, +{df.revenue.iloc[-1] / df.revenue.iloc[-5] - 1:.0%} YoY; Cloud now {q_cloud_share:.0%} of revenue (Q2'25: {df.rev_cloud.iloc[-5] / df.revenue.iloc[-5]:.0%}).",
           f"Base case: FY26E revenue +{b26.revenue / fy25_rev - 1:.0%} to ${b26.revenue / 1000:,.0f}B; operating income ${b26.operating_income / 1000:,.0f}B ({b26.op_margin:.0%} margin).",
           "Q4 is the seasonal peak every year, so each forecast quarter grows off the same quarter a year earlier.",
           "Search & other is still ~half of revenue; Cloud adds the most incremental dollars."],
          "01_revenue_mix.png", 7.9)
    slide("Simple YoY growth rules match or beat ETS and SARIMA at 1–4 quarter horizons",
          ["Rolling-origin backtest: 13 origins per segment (train through t, forecast t+1…t+4), 4 horizons each.",
           "Cloud: best model is the plain YoY rule (5.1% MAPE); seasonal naive is worst (24%) because Cloud grows ~30–80% a year.",
           "Search: SARIMA, ETS and the YoY rule are tied at ~4.4% MAPE; no model clearly wins.",
           "Subscriptions & devices is among the hardest for the better models (6–7% MAPE): launch- and promo-driven.",
           "Caveat: only 30 quarters of consistent segment data, so differences under ~1pt are noise."],
          "03_backtest_mape.png", 7.9)
    slide("Cloud is the swing factor: ±5pp of growth moves FY27E operating income by about ±" + f"{(bull27.operating_income / b27.operating_income - 1):.0%}",
          [f"Bear/Bull = Base growth −/+ 5pp on Cloud and −/+ 2pp on Search and YouTube.",
           f"FY27E operating income: Bear ${bear27.operating_income / 1000:,.0f}B / Base ${b27.operating_income / 1000:,.0f}B / Bull ${bull27.operating_income / 1000:,.0f}B.",
           f"FY27E Cloud revenue ranges ${bear27.cloud / 1000:,.0f}B–${bull27.cloud / 1000:,.0f}B.",
           "R&D and G&A are held at Base dollars in Bear/Bull, so the downside shows real operating deleverage."],
          "04_scenarios.png", 7.9)
    e27 = cmp27.loc["Google Cloud"]
    slide(f"Cloud forecasts diverge widely: FY27E ranges from ${e27['Excel growth model (Base)'] / 1000:,.0f}B (growth model) to ${e27['ETS'] / 1000:,.0f}B (ETS)",
          [f"FY27E Cloud revenue: Excel base ${e27['Excel growth model (Base)'] / 1000:,.0f}B, SARIMA ${e27['SARIMA'] / 1000:,.0f}B, ETS ${e27['ETS'] / 1000:,.0f}B.",
           "ETS extrapolates the recent acceleration (+82% YoY in Q2'26); the Excel base fades growth 10% per quarter, so it is the conservative end.",
           f"Even the Bull case (${bull27.cloud / 1000:,.0f}B) sits below ETS. In the backtest, fading growth cost accuracy on Cloud.",
           "Recommendation: plan off the transparent growth-rate model, treat ETS/SARIMA as upside cross-checks, and re-run the backtest each quarter."],
          "02_cloud_forecast.png", 7.2)
    prs.save(ROOT / "output" / "alphabet_forecast_summary.pptx")
    print("deck saved")


if __name__ == "__main__":
    main()
