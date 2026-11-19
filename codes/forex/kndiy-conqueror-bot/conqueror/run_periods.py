"""Compare the book's methods in chosen windows, with full-history warm-up.

Usage: python3 -m conqueror.run_periods [--spread-pips 1.0] [--risk 0.005]
A trade belongs to a window by its ENTRY date. Window return = sum(trade pnl) / equity at the window start (approximate:
sizing compounds through the full run). Pairs: USD-quote pairs only.
"""
import argparse

from .channel import ChannelParams, channel_backtest
from .run_backtest import load
from .strategy import Params, backtest
from .trend import TrendParams, trend_backtest

PAIRS = ["EURUSD", "GBPUSD", "AUDUSD", "NZDUSD"]
WINDOWS = [("2017-2019", "2017-01-01", "2019-12-31"), ("2023-now", "2023-01-01", "2099-12-31"),
           ("full", "0000-01-01", "2099-12-31")]


def runners(pip, risk, spread):
    return {
        "Conqueror": lambda d, o, h, l, c: backtest(d, o, h, l, c, Params(risk_pct=risk, spread=spread * pip)),
        "Conqueror+Bishop": lambda d, o, h, l, c: backtest(d, o, h, l, c, Params(risk_pct=risk, spread=spread * pip, bishop=True)),
        "Channel (book)": lambda d, o, h, l, c: channel_backtest(d, o, h, l, c, ChannelParams(risk_pct=risk, spread=spread * pip, buffer=3 * pip)),
        "Trend (book)": lambda d, o, h, l, c: trend_backtest(d, o, h, l, c, TrendParams(risk_pct=risk, spread=spread * pip, buffer=3 * pip)),
    }


def window_stats(res, start, end, start_equity=10_000.0):
    tr = [t for t in res.trades if start <= t.entry_date <= end]
    eq0 = start_equity
    for dt, e in res.equity_curve:
        if dt < start:
            eq0 = e
        else:
            break
    wins = sum(t.pnl for t in tr if t.pnl > 0)
    losses = -sum(t.pnl for t in tr if t.pnl <= 0)
    return dict(n=len(tr), wins=wins, losses=losses, pnl=sum(t.pnl for t in tr), eq0=eq0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spread-pips", type=float, default=1.0)
    ap.add_argument("--risk", type=float, default=0.005)
    a = ap.parse_args()
    pip = 0.0001
    out = {}  # (method, window, pair) -> stats
    for pair in PAIRS:
        d, o, h, l, c = load(pair, "2003-01-01")
        for m, fn in runners(pip, a.risk, a.spread_pips).items():
            res = fn(d, o, h, l, c)
            for w, s, e in WINDOWS:
                out[(m, w, pair)] = window_stats(res, s, e)
    for w, s, e in WINDOWS:
        print(f"\n=== window {w}  (spread {a.spread_pips} pip, risk {a.risk:.2%}, pairs {', '.join(PAIRS)})")
        print(f"{'method':<18}{'trades':>7}{'PF pooled':>11}{'avg ret/pair':>14}   per-pair PF " + " ".join(f"{p[:3]}{p[3:]:<3}" for p in PAIRS))
        for m in runners(pip, a.risk, a.spread_pips):
            st = [out[(m, w, p)] for p in PAIRS]
            n = sum(x["n"] for x in st)
            W, L = sum(x["wins"] for x in st), sum(x["losses"] for x in st)
            pf = W / L if L > 0 else float("inf")
            ret = sum(x["pnl"] / x["eq0"] for x in st) / len(st)
            per = " ".join(f"{(x['wins'] / x['losses'] if x['losses'] > 0 else float('inf')):>6.2f}" for x in st)
            print(f"{m:<18}{n:>7}{pf:>11.2f}{ret:>+14.1%}   {per}")


if __name__ == "__main__":
    main()
