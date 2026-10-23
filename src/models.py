"""Forecasting models shared by the backtest and the forward forecast.

Every model has the signature  f(y: np.ndarray, h: int) -> np.ndarray  where y is the history of
quarterly values (oldest first) and the return is the next h quarters.
"""
import warnings

import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.statespace.sarimax import SARIMAX

warnings.filterwarnings("ignore")


def naive(y, h):
    return np.repeat(y[-1], h)


def seasonal_naive(y, h):
    return np.array([y[len(y) - 4 + (k % 4)] for k in range(h)])


def _yoy_forecast(y, h, g):
    """Same quarter last year x (1+g), compounding once per year of horizon."""
    return np.array([y[len(y) - 4 + (k % 4)] * (1 + g) ** (k // 4 + 1) for k in range(h)])


def yoy_last(y, h):
    """Growth-rate model used in the Excel P&L: carry the latest YoY growth rate forward."""
    return _yoy_forecast(y, h, y[-1] / y[-5] - 1)


def yoy_avg4(y, h):
    """Smoother variant: average of the last four YoY growth rates."""
    g = np.mean([y[-i] / y[-i - 4] - 1 for i in range(1, 5)])
    return _yoy_forecast(y, h, g)


def yoy_faded(y, h, fade=0.9):
    """Excel-model rule: latest YoY growth, decaying by `fade` each quarter ahead (k=1 -> g*fade, k=2 -> g*fade^2 ...).
    Each forecast quarter = same quarter a year earlier (actual or already-forecast) x (1 + faded growth).
    `fade` is fixed a priori (not tuned on the backtest)."""
    g = y[-1] / y[-5] - 1
    f = []
    for k in range(h):
        base = y[len(y) - 4 + k] if k < 4 else f[k - 4]
        f.append(base * (1 + g * fade ** (k + 1)))
    return np.array(f)


def ets(y, h):
    """Holt-Winters on log revenue: damped additive trend + additive seasonality (=multiplicative in levels)."""
    fit = ExponentialSmoothing(np.log(y), trend="add", damped_trend=True, seasonal="add",
                               seasonal_periods=4, initialization_method="estimated").fit()
    return np.exp(fit.forecast(h))


def sarima(y, h):
    """'Airline' SARIMA(0,1,1)(0,1,1)4 on log revenue - the standard choice for short seasonal growth series."""
    fit = SARIMAX(np.log(y), order=(0, 1, 1), seasonal_order=(0, 1, 1, 4),
                  enforce_stationarity=False, enforce_invertibility=False).fit(disp=False, maxiter=200)
    return np.exp(fit.forecast(h))


MODELS = {
    "Naive": naive,
    "Seasonal naive": seasonal_naive,
    "YoY growth (last)": yoy_last,
    "YoY growth (faded)": yoy_faded,
    "YoY growth (avg 4Q)": yoy_avg4,
    "ETS": ets,
    "SARIMA": sarima,
}
