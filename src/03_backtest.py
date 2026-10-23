"""Day 3 + 4: rolling-origin backtest of every model on every revenue series.

For each origin t (training data through quarter t) forecast t+1..t+4, compare with the actuals,
then step the origin forward one quarter. MAPE = mean(|actual - forecast| / actual).
"""
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

from models import MODELS

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)
MIN_TRAIN = 14      # quarters in the first training window
HORIZON = 4

SERIES = {
    "Search & other": "rev_search",
    "YouTube ads": "rev_youtube",
    "Network": "rev_network",
    "Subscriptions, platforms & devices": "rev_subscriptions",
    "Cloud": "rev_cloud",
    "Total revenue": "revenue",
}


def load():
    con = sqlite3.connect(ROOT / "data" / "alphabet.db")
    df = pd.read_sql("select * from v_pnl", con, parse_dates=["quarter_end"]).set_index("quarter_end")
    return df


def backtest(df):
    rows = []
    for name, col in SERIES.items():
        s = df[col].dropna()           # segments: 2019Q1+ (30 qtrs); total revenue: 2016Q1+ (42 qtrs)
        y, idx = s.values.astype(float), s.index
        for end in range(MIN_TRAIN, len(y) - HORIZON + 1):      # end = number of training quarters
            train = y[:end]
            for mname, f in MODELS.items():
                try:
                    fc = f(train, HORIZON)
                except Exception as e:  # a failed fit is recorded, not hidden
                    print("fit failed", name, mname, idx[end - 1].date(), e)
                    fc = np.full(HORIZON, np.nan)
                for h in range(1, HORIZON + 1):
                    a, p = y[end - 1 + h], fc[h - 1]
                    rows.append(dict(series=name, model=mname, origin=idx[end - 1], horizon=h,
                                     target=idx[end - 1 + h], actual=a, forecast=p, ape=abs(a - p) / a))
    return pd.DataFrame(rows)


if __name__ == "__main__":
    bt = backtest(load())
    bt.to_csv(OUT / "backtest_forecasts.csv", index=False)
    order = list(MODELS)
    mape = bt.groupby(["series", "model"]).ape.mean().unstack("model")[order] * 100
    mape = mape.loc[list(SERIES)]
    mape.round(1).to_csv(OUT / "mape_by_series_model.csv")
    by_h = bt.groupby(["model", "horizon"]).ape.mean().unstack("horizon") * 100
    by_h.loc[order].round(1).to_csv(OUT / "mape_by_horizon.csv")
    n = bt.groupby("series").origin.nunique()
    print("origins per series:", n.to_dict())
    print("\nMAPE (%) by series x model, all horizons 1-4 pooled:")
    print(mape.round(1).to_string())
    print("\nBest model per series:", mape.idxmin(axis=1).to_dict())
    print("\nMAPE (%) by horizon (pooled across series):")
    print(by_h.loc[order].round(1).to_string())
    # last 8 origins only: does the ranking change in the recent regime (post-2024)?
    recent = bt[bt.origin >= bt.origin.max() - pd.offsets.QuarterEnd(7)]
    print("\nMAPE (%) over the most recent 8 origins:")
    print((recent.groupby(["series", "model"]).ape.mean().unstack("model")[order].loc[list(SERIES)] * 100).round(1).to_string())
