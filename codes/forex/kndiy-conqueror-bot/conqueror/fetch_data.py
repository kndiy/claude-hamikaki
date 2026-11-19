"""Download daily OHLC from Yahoo Finance's unofficial chart endpoint into the CSV warehouse.

Usage: python3 -m conqueror.fetch_data EURUSD [GBPUSD ...]
Yahoo FX history is indicative, not tradable broker data, and its day boundary drifts with daylight saving (bars are
labelled Sun-Thu for part of the year and Mon-Fri for the rest; still one bar per trading day). Use for first passes only.
"""
import json
import sys
import urllib.request
from datetime import datetime, timezone

from . import warehouse

SOURCE = "yahoo"


def fetch_rows(symbol: str) -> list:
    # range=max silently downgrades to monthly bars, so ask for explicit bounds.
    p1, p2 = 946684800, int(datetime.now(timezone.utc).timestamp())  # 2000-01-01 .. now
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}=X?period1={p1}&period2={p2}&interval=1d"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        res = json.load(r)["chart"]["result"][0]
    if res["meta"].get("dataGranularity") != "1d":
        raise RuntimeError(f"expected daily data, got {res['meta'].get('dataGranularity')}")
    q = res["indicators"]["quote"][0]
    rows = []
    for t, o, h, l, c in zip(res["timestamp"], q["open"], q["high"], q["low"], q["close"]):
        if None in (o, h, l, c) or not (l <= min(o, c) and h >= max(o, c)):
            continue  # skip missing or inconsistent bars
        rows.append(dict(date=datetime.fromtimestamp(t, timezone.utc).strftime("%Y-%m-%d"), open=o, high=h, low=l, close=c))
    return rows


if __name__ == "__main__":
    for sym in (sys.argv[1:] or ["EURUSD"]):
        added, issues = warehouse.update(sym, fetch_rows, SOURCE)
        print(f"{sym}: +{added} rows, {len(issues)} validation issue(s)")
        for i in issues[:5]:
            print("   ", i)
