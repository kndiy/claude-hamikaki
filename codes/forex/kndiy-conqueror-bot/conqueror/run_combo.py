"""Usage: python3 -m conqueror.run_combo [--spread-pips 1.0] [--risk 0.005]   (Conqueror runs with the Bishop, owner decision)"""
import argparse

from .channel import ChannelParams, channel_backtest
from .combine import combine, independent
from .run_backtest import load
from .run_periods import PAIRS, WINDOWS, window_stats
from .strategy import Params, backtest
from .trend import TrendParams, trend_backtest


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spread-pips", type=float, default=1.0)
    ap.add_argument("--risk", type=float, default=0.005)
    a = ap.parse_args()
    pip, sp, rk = 0.0001, a.spread_pips * 0.0001, a.risk
    out = {}
    for pair in PAIRS:
        d, o, h, l, c = load(pair, "2003-01-01")
        res = {
            "Conqueror+Bishop": backtest(d, o, h, l, c, Params(risk_pct=rk, spread=sp, bishop=True)),
            "Channel (book)": channel_backtest(d, o, h, l, c, ChannelParams(risk_pct=rk, spread=sp, buffer=3 * pip)),
            "Trend (book)": trend_backtest(d, o, h, l, c, TrendParams(risk_pct=rk, spread=sp, buffer=3 * pip)),
        }
        tb = {k: v.trades for k, v in res.items()}
        res["A independent (sum)"] = independent(res)
        res["B all agree"] = combine(tb, "all", d, risk_pct=rk)
        res["C majority (2 of 3)"] = combine(tb, "majority", d, risk_pct=rk)
        for m, r in res.items():
            for w, s, e in WINDOWS:
                out[(m, w, pair)] = window_stats(r, s, e)
    methods = list(next(iter(out.values())) and dict.fromkeys(k[0] for k in out))
    for w, s, e in WINDOWS:
        print(f"\n=== window {w} (spread {a.spread_pips} pip, risk {rk:.2%} per method/position; trades pooled over {', '.join(PAIRS)})")
        print(f"{'method':<22}{'trades':>7}{'PF pooled':>11}{'avg ret/pair':>14}   per-pair PF EUR GBP AUD NZD")
        for m in methods:
            st = [out[(m, w, p)] for p in PAIRS]
            n = sum(x["n"] for x in st)
            W, L = sum(x["wins"] for x in st), sum(x["losses"] for x in st)
            pf = W / L if L > 0 else float("inf")
            ret = sum(x["pnl"] / x["eq0"] for x in st) / len(st)
            per = " ".join(f"{(x['wins'] / x['losses'] if x['losses'] > 0 else float('inf')):>5.2f}" for x in st)
            print(f"{m:<22}{n:>7}{pf:>11.2f}{ret:>+14.1%}   {per}")


if __name__ == "__main__":
    main()
