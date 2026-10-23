"""Day 1a: pull raw data from SEC EDGAR.

1. companyfacts API  -> consolidated line items (revenue, costs, R&D, S&M, G&A, op income)
2. per-filing XBRL instance documents -> revenue by product/service (a dimensional fact,
   which the companyfacts API strips out)

Raw downloads are cached in data/raw so reruns are fast and polite to SEC (limit 10 req/s).
Set SEC_USER_AGENT to "Your Name your@email" - SEC requires a contact in the header.
"""
import json
import os
import re
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
RAW.mkdir(parents=True, exist_ok=True)
CIK = "0001652044"
UA = os.environ.get("SEC_USER_AGENT", "alphabet-forecast-project research (set-your-email@example.com)")
S = requests.Session()
S.headers["User-Agent"] = UA


def get(url, cache_name, as_json=False):
    p = RAW / cache_name
    if not p.exists():
        r = S.get(url, timeout=60)
        r.raise_for_status()
        p.write_bytes(r.content)
        time.sleep(0.15)
    return json.loads(p.read_text()) if as_json else p.read_text()


def list_filings():
    sub = get(f"https://data.sec.gov/submissions/CIK{CIK}.json", "submissions.json", True)
    blocks = [sub["filings"]["recent"]]
    for f in sub["filings"].get("files", []):
        blocks.append(get(f"https://data.sec.gov/submissions/{f['name']}", f["name"], True))
    out = []
    for b in blocks:
        for form, rep, acc in zip(b["form"], b["reportDate"], b["accessionNumber"]):
            if form in ("10-K", "10-Q") and rep >= "2015-12-31":
                out.append((form, rep, acc))
    return sorted(set(out), key=lambda x: x[1])


def instance_url(acc):
    nodash = acc.replace("-", "")
    idx = get(f"https://www.sec.gov/Archives/edgar/data/1652044/{nodash}/index.json",
              f"idx_{nodash}.json", True)
    names = [i["name"] for i in idx["directory"]["item"]]
    # the extracted instance is goog-YYYYMMDD_htm.xml (inline XBRL, 2020+) or goog<doc>_htm.xml (2019)
    cand = [n for n in names if re.fullmatch(r"goog.*_htm\.xml", n)] or \
           [n for n in names if re.fullmatch(r"goog-\d{8}\.xml", n)]
    if not cand:
        raise RuntimeError(f"no instance doc in {acc}: {names[:15]}")
    return f"https://www.sec.gov/Archives/edgar/data/1652044/{nodash}/{cand[0]}", f"inst_{nodash}.xml"


if __name__ == "__main__":
    get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json", "companyfacts.json")
    filings = list_filings()
    print(len(filings), "10-K/10-Q filings from", filings[0][1], "to", filings[-1][1])
    for form, rep, acc in filings:
        try:
            url, name = instance_url(acc)
            get(url, name)
            print("ok ", form, rep, acc)
        except Exception as e:  # keep going; gaps are reported by the parse step
            print("FAIL", form, rep, acc, e)
    json.dump(filings, open(RAW / "filings.json", "w"))
