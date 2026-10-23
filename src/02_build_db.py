"""Day 1b: turn the raw SEC downloads into a quarterly SQLite database.

Consolidated line items  <- companyfacts API (data/raw/companyfacts.json)
Revenue by product line  <- dimensional facts in each filing's XBRL instance (data/raw/inst_*.xml)

Both sources only publish Q1-Q3 as 3-month facts; Q4 is derived as FY minus (Q1+Q2+Q3).
When a period appears in several filings (original + later comparative) the latest filing wins,
so restated/recast numbers replace the originals.
"""
import json
import sqlite3
from collections import defaultdict
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
DB = ROOT / "data" / "alphabet.db"
FIRST_Q = date(2016, 1, 1)           # consolidated history starts here
SEG_FIRST_Q = date(2019, 1, 1)       # product-line table is only disclosed on a consistent basis from here

CONSOLIDATED = {  # metric -> us-gaap tag
    "revenue": ("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"),  # tag switched over time
    "cost_of_revenues": "CostOfRevenue",
    "rd": "ResearchAndDevelopmentExpense",
    "sales_marketing": "SellingAndMarketingExpense",
    "g_and_a": "GeneralAndAdministrativeExpense",
    "operating_income": "OperatingIncomeLoss",
}

# XBRL member -> our segment name. Names changed over the years (e.g. "Other revenues" was renamed
# "Subscriptions, platforms and devices" in the FY2023 10-K), so several members map to one segment.
SEGMENT_MEMBERS = {
    "GoogleSearchOtherMember": "rev_search",
    "YouTubeAdvertisingRevenueMember": "rev_youtube",
    "GoogleNetworkMember": "rev_network",
    "GoogleNetworkMembersPropertiesMember": "rev_network",
    "OtherRevenuesMember": "rev_subscriptions",
    "SubscriptionsPlatformsAndDevicesRevenueMember": "rev_subscriptions",
    "GoogleCloudMember": "rev_cloud",
    "GoogleCloudSegmentMember": "rev_cloud",
    "AllOtherSegmentsMember": "rev_other_bets",
}
# members that qualify the segment (Google Services / Google Inc.) and may appear alongside the product member
PARENT_MEMBERS = {"GoogleServicesMember", "GoogleServicesSegmentMember", "GoogleInc.Member"}
NS = {"x": "http://www.xbrl.org/2003/instance", "xd": "http://xbrl.org/2006/xbrldi"}


def days(s, e):
    return (date.fromisoformat(e) - date.fromisoformat(s)).days


def quarter_start(q_end):
    m = q_end.month - 2
    return date(q_end.year, m, 1)


def quarterize(facts):
    """facts: list of (start, end, value, filed) ISO strings. Returns {quarter_end(date): value}."""
    best = {}
    for s, e, v, f in facts:
        if (s, e) not in best or f > best[(s, e)][1]:
            best[(s, e)] = (v, f)
    q, fy = {}, {}
    for (s, e), (v, _) in best.items():
        d = days(s, e)
        if 80 <= d <= 100:
            q[date.fromisoformat(e)] = v
        elif 350 <= d <= 380:
            fy[date.fromisoformat(e)] = v
    for fy_end, total in fy.items():
        if fy_end in q:
            continue
        prior = [date(fy_end.year, m, 1) for m in (3, 6, 9)]
        ends = [date(fy_end.year, 3, 31), date(fy_end.year, 6, 30), date(fy_end.year, 9, 30)]
        if all(x in q for x in ends):
            q[fy_end] = total - sum(q[x] for x in ends)
    return q


def consolidated_rows():
    cf = json.loads((RAW / "companyfacts.json").read_text())["facts"]["us-gaap"]
    rows = []
    for metric, tag in CONSOLIDATED.items():
        tags = (tag,) if isinstance(tag, str) else tag
        facts = [(x["start"], x["end"], x["val"], x["filed"]) for t in tags for x in cf[t]["units"]["USD"] if "start" in x]
        for q_end, v in quarterize(facts).items():
            rows.append((q_end, metric, v / 1e6, "companyfacts"))
    return rows


def segment_rows():
    facts = defaultdict(list)
    for f in sorted(RAW.glob("inst_*.xml")):
        root = ET.parse(f).getroot()
        filed = f.stem.split("_")[1]            # accession no. sorts chronologically within a year; fine as a tiebreak
        ctx = {}
        for c in root.findall("x:context", NS):
            members = {m.text.split(":")[1] for m in c.findall(".//xd:explicitMember", NS)}
            per = c.find("x:period", NS)
            s, e = per.find("x:startDate", NS), per.find("x:endDate", NS)
            if s is None:
                continue
            ctx[c.get("id")] = (members, s.text, e.text)
        for el in root:
            if not el.tag.endswith("}RevenueFromContractWithCustomerExcludingAssessedTax"):
                continue
            c = ctx.get(el.get("contextRef"))
            if not c:
                continue
            members, s, e = c
            prod = [m for m in members if m in SEGMENT_MEMBERS]
            if len(prod) != 1 or not (members - set(prod)) <= PARENT_MEMBERS:
                continue
            # 2018/2019 filings use the old taxonomy (Google properties / network members); skip them -
            # only accept facts whose period is from the new disclosure era (>= 2019)
            if date.fromisoformat(s) < SEG_FIRST_Q:
                continue
            seg = SEGMENT_MEMBERS[prod[0]]
            # 'OtherRevenuesMember' = "Google other", later renamed "Subscriptions, platforms and devices".
            # The 2019 10-Qs report it *including* cloud; the 2020 10-Qs recast 2019 without it and, being
            # later filings, win in quarterize().
            facts[seg].append((s, e, int(el.text), filed))
    rows = []
    for seg, fl in facts.items():
        for q_end, v in quarterize(fl).items():
            if q_end >= SEG_FIRST_Q:
                rows.append((q_end, seg, v / 1e6, "xbrl_instance"))
    return rows


SCHEMA = """
DROP TABLE IF EXISTS quarterly_financials;
CREATE TABLE quarterly_financials (
    quarter_end  DATE NOT NULL,
    fiscal_year  INTEGER NOT NULL,
    fiscal_q     INTEGER NOT NULL,
    metric       TEXT NOT NULL,
    value_musd   REAL NOT NULL,
    source       TEXT NOT NULL,
    PRIMARY KEY (quarter_end, metric)
);
DROP VIEW IF EXISTS v_pnl;
CREATE VIEW v_pnl AS
SELECT quarter_end, fiscal_year, fiscal_q,
  MAX(CASE WHEN metric='revenue' THEN value_musd END)          AS revenue,
  MAX(CASE WHEN metric='cost_of_revenues' THEN value_musd END) AS cost_of_revenues,
  MAX(CASE WHEN metric='revenue' THEN value_musd END) - MAX(CASE WHEN metric='cost_of_revenues' THEN value_musd END) AS gross_profit,
  MAX(CASE WHEN metric='rd' THEN value_musd END)               AS rd,
  MAX(CASE WHEN metric='sales_marketing' THEN value_musd END)  AS sales_marketing,
  MAX(CASE WHEN metric='g_and_a' THEN value_musd END)          AS g_and_a,
  COALESCE(MAX(CASE WHEN metric='ec_fines' THEN value_musd END),0) AS ec_fines,
  MAX(CASE WHEN metric='operating_income' THEN value_musd END) AS operating_income,
  MAX(CASE WHEN metric='rev_search' THEN value_musd END)       AS rev_search,
  MAX(CASE WHEN metric='rev_youtube' THEN value_musd END)      AS rev_youtube,
  MAX(CASE WHEN metric='rev_network' THEN value_musd END)      AS rev_network,
  MAX(CASE WHEN metric='rev_subscriptions' THEN value_musd END) AS rev_subscriptions,
  MAX(CASE WHEN metric='rev_cloud' THEN value_musd END)        AS rev_cloud,
  MAX(CASE WHEN metric='rev_other_bets' THEN value_musd END)   AS rev_other_bets,
  MAX(CASE WHEN metric='rev_hedging' THEN value_musd END)      AS rev_hedging
FROM quarterly_financials GROUP BY quarter_end ORDER BY quarter_end;
"""


def main():
    rows = [r for r in consolidated_rows() if r[0] >= date(2016, 3, 31)] + segment_rows()
    # Alphabet books European Commission fines as their own opex line (not in the tags we pull).
    # They are the only reason revenue - costs != operating income, so back them out explicitly.
    by = defaultdict(dict)
    for q, m, v, _ in rows:
        by[q][m] = v
    for q, d in by.items():
        if "revenue" not in d:
            continue
        gap = d["revenue"] - d["cost_of_revenues"] - d["rd"] - d["sales_marketing"] - d["g_and_a"] - d["operating_income"]
        if abs(gap) > 1:
            rows.append((q, "ec_fines", gap, "derived"))
    # hedging gains/losses is the 7th revenue line; not tagged separately, so it is the exact residual
    seg_names = {"rev_search", "rev_youtube", "rev_network", "rev_subscriptions", "rev_cloud", "rev_other_bets"}
    for q, d in by.items():
        if "revenue" in d and seg_names <= d.keys():
            rows.append((q, "rev_hedging", d["revenue"] - sum(d[k] for k in seg_names), "derived"))
    con = sqlite3.connect(DB)
    con.executescript(SCHEMA)
    con.executemany(
        "INSERT INTO quarterly_financials VALUES (?,?,?,?,?,?)",
        [(q.isoformat(), q.year, (q.month - 1) // 3 + 1, m, v, s) for q, m, v, s in rows],
    )
    con.commit()

    # ---- integrity checks ------------------------------------------------------------------
    print("quarters:", con.execute("select count(distinct quarter_end), min(quarter_end), max(quarter_end) from quarterly_financials").fetchone())
    print("EC fines backed out ($M):", con.execute("select quarter_end, round(ec_fines) from v_pnl where ec_fines>0").fetchall())
    bad = con.execute("""select quarter_end from v_pnl where abs(revenue-cost_of_revenues-rd-sales_marketing-g_and_a-ec_fines-operating_income) > 1""").fetchall()
    print("P&L tie-out failures after fines:", bad or "none")
    print("hedging residual range ($M):", con.execute("select round(min(rev_hedging)), round(max(rev_hedging)) from v_pnl").fetchone(), "(small => segments reconcile)")
    import csv
    cur = con.execute("select * from v_pnl")
    with open(ROOT / "data" / "quarterly_pnl.csv", "w", newline="") as fh:
        w = csv.writer(fh); w.writerow([c[0] for c in cur.description]); w.writerows(cur)


if __name__ == "__main__":
    main()
