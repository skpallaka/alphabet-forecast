# Forecasting Alphabet's P&L: segment revenue, backtested models and scenarios

**Question:** How well can we forecast Alphabet's quarterly revenue by segment and operating income, and which method is worth using for planning?

**Short answer:** A transparent year-over-year growth rule is as accurate as ETS and SARIMA over 1–4 quarter horizons (about 4–7% MAPE per segment). The forecasts only disagree widely on Google Cloud, and Cloud growth is the largest swing factor in operating income.

Everything is built from SEC EDGAR filings and reproducible with the scripts in `src/`.

## Method

| Step | What | Where |
|---|---|---|
| 1. Data | Consolidated line items (revenue, cost of revenues, R&D, S&M, G&A, operating income) from the EDGAR `companyfacts` API. Segment revenue (Search, YouTube ads, Network, Subscriptions/platforms/devices, Cloud, Other Bets) from each 10-Q/10-K's XBRL instance document, because `companyfacts` strips dimensional facts. Q4 = FY − (Q1+Q2+Q3). Loaded into SQLite. | `src/01_fetch.py`, `src/02_build_db.py`, `data/alphabet.db` |
| 2. Excel model | Income statement with live formulas. Segment revenue = same quarter last year × (1 + growth); growth = latest YoY × fade^k + scenario delta. Cost of revenues, S&M, R&D, G&A = % of revenue at the trailing 4-quarter average. Separate Base / Bull / Bear sheets. | `excel/alphabet_pnl_model.xlsx`, `src/04_build_excel.py` |
| 3. Statistical models | ETS (damped trend, additive seasonality) and SARIMA(0,1,1)(0,1,1)₄, both on log revenue, one per segment. | `src/models.py` |
| 4. Backtest | Rolling origin: train through quarter t, forecast t+1…t+4, compare with actuals, step forward one quarter. 13 origins per segment (first training window = 14 quarters), 25 for total revenue. | `src/03_backtest.py` |
| 5. Scenarios + deck | Bear/Bull = Base growth −/+ 5pp on Cloud, −/+ 2pp on Search and YouTube. R&D and G&A held at Base dollars in Bear/Bull. 4-slide summary. | `src/05_report.py`, `output/alphabet_forecast_summary.pptx` |

## Backtest results (MAPE %, horizons 1–4 pooled; lower is better)

| Segment | Naive | Seasonal naive | YoY (last) | YoY (faded 0.9) | YoY (avg 4Q) | ETS | SARIMA |
|---|---|---|---|---|---|---|---|
| Search & other | 7.8 | 10.0 | **4.4** | 5.0 | 6.8 | **4.4** | **4.4** |
| YouTube ads | 9.9 | 10.4 | **5.6** | **5.6** | 7.8 | 5.8 | 6.6 |
| Network | 4.9 | 3.4 | 3.6 | **3.1** | 6.1 | 7.3 | 4.6 |
| Subscriptions, platforms & devices | 13.7 | 14.9 | 7.0 | 7.0 | **6.0** | 7.3 | **6.0** |
| Cloud | 16.3 | 24.3 | **5.1** | 7.0 | 6.1 | 6.2 | 6.0 |
| Total revenue* | 10.8 | 14.0 | 9.4 | 8.7 | 10.9 | **7.8** | 9.0 |

\*Total revenue uses 2016+ history, so its origins include the 2020 COVID shock; it is not comparable to the segment rows.
Full tables: `output/mape_by_series_model.csv`, `output/mape_by_horizon.csv`, all individual forecasts in `output/backtest_forecasts.csv`.

### What the backtest found
- **No model dominates.** On Search, ETS, SARIMA and the plain YoY rule tie at 4.4%. With only 30 quarters of consistent segment data, differences under about 1 point are noise.
- **Naive and seasonal-naive are clearly worst** on fast growers. Seasonal naive on Cloud is 24% MAPE, because it ignores growth.
- **Error grows with horizon** for every model: roughly 4% at 1 quarter to 8% at 4 quarters (`output/mape_by_horizon.csv`).
- **Subscriptions, platforms & devices is among the hardest segments** for the better models (6–7%), consistent with launch- and promotion-driven revenue.
- **Fading growth is a judgment call that the data did not reward on Cloud** (5.1% → 7.0%), because Cloud kept accelerating. It is neutral to slightly helpful on Network and total revenue. The 0.9 fade was fixed in advance, not tuned.
- **For Cloud the models diverge sharply going forward**, even though they scored similarly in backtest (FY27E Cloud revenue: Excel base $151B, SARIMA $161B, ETS $235B). ETS extrapolates the recent acceleration (+82% YoY in Q2'26), and a backtest on a decelerating-then-accelerating history cannot tell us which is right.

## Scenarios ($B)

| | FY25A | FY26E Bear / Base / Bull | FY27E Bear / Base / Bull |
|---|---|---|---|
| Cloud revenue | 58.7 | 98.8 / 100.5 / 102.1 | 143.9 / 151.2 / 158.7 |
| Total revenue | 402.8 | 487.4 / 491.8 / 496.3 | 562.0 / 578.5 / 595.2 |
| Operating income | 129.0 | 165.6 / 168.0 / 170.4 | 187.0 / 195.9 / 205.0 |
| Operating margin | 32.0% | 34.0 / 34.2 / 34.3% | 33.3 / 33.9 / 34.4% |

FY26 moves little because two of its four quarters are already reported. In FY27 the Bull/Bear spread is about ±4.6% of operating income.

## Key takeaway
For planning, use the transparent growth-rate model, since it is as accurate as ETS/SARIMA at 1–4 quarters and every number is auditable. Treat ETS/SARIMA as cross-checks, and put the real uncertainty where the models disagree: Cloud.

## Data notes and limitations
- **Segment history starts in 2019.** Alphabet only disclosed this product-line split (and restated 2019 comparatives) from its 2020 filings, so the segment series has 30 quarters, not 40. Consolidated P&L goes back to Q1 2016 (42 quarters, in the database).
- **No hand-filled data.** Every segment value came from the filings' XBRL instance documents, so nothing had to be typed in from 10-Q text. Segments plus hedging gains/losses sum exactly to total revenue by construction (hedging is the derived residual, between −$207M and +$669M per quarter, consistent with the filings).
- **European Commission fines** (Q2'17 $2.7B, Q2'18 $5.1B, Q1'19 $1.7B) are a separate opex line not covered by the tags used. They are backed out as the exact residual so operating income ties to reported figures in all 42 quarters.
- "Other revenues" was renamed "Subscriptions, platforms and devices"; the loader maps both names. Where a period appears in several filings, the latest filing's (recast) value is used.
- Forecasts are a modelling exercise on public data, not investment advice. Forecast hedging gains and EC fines are set to zero.

## Reproduce
```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
export SEC_USER_AGENT="Your Name your@email.com"   # SEC requires a contact in the User-Agent
.venv/bin/python src/01_fetch.py      # ~45 filings -> data/raw (cached, ~120MB, git-ignored)
.venv/bin/python src/02_build_db.py   # -> data/alphabet.db, data/quarterly_pnl.csv
.venv/bin/python src/03_backtest.py   # -> output/*.csv
.venv/bin/python src/04_build_excel.py
.venv/bin/python src/05_report.py     # -> charts + 4-slide deck
.venv/bin/python src/verify_excel.py  # recalcs the workbook and checks it against an independent implementation
```
Example SQL: `select fiscal_year, sum(revenue), sum(operating_income) from v_pnl group by 1;`

## Excel model guide
Open `excel/alphabet_pnl_model.xlsx`. **Assumptions** holds the only inputs (blue on blue fill). **Model_Base/Bull/Bear** have blue hard-coded SEC history and black formulas for everything else; the forecast columns are shaded yellow. **Scenarios** compares cases with formulas into the three model sheets. **Seasonality** shows each quarter's share of the year. The workbook calculates when opened in Excel, Numbers or Google Sheets.
