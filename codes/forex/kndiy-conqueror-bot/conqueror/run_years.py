"""Per-calendar-year breakdown for one pair. Usage: python3 -m conqueror.run_years [SYMBOL] [--spread-pips 1.0] [--risk 0.005]
A trade belongs to the year of its ENTRY. Return = sum(trade pnl) / equity at the start of that year (approximate: sizing
compounds through the whole run). 'move' is the pair's close-to-close change for the year, for context only."""
import argparse

from .channel import ChannelParams, channel_backtest
from .combine import combine
from .run_backtest import load
from .strategy import Params, backtest
from .trend import TrendParams, trend_backtest


def by_year(res, years):
    out = {}
    for y in years:
        tr = [t for t in res.trades if t.entry_date[:4] == y]
        eq0 = 10_000.0
        for dt, e in res.equity_curve:
            if dt[:4] < y:
                eq0 = e
            else:
                break
        W = sum(t.pnl for t in tr if t.pnl > 0)
        L = -sum(t.pnl for t in tr if t.pnl <= 0)
        out[y] = (len(tr), (W / L if L > 0 else float("inf")), sum(t.pnl for t in tr) / eq0)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbol", nargs="?", default="EURUSD")
    ap.add_argument("--spread-pips", type=float, default=1.0)
    ap.add_argument("--risk", type=float, default=0.005)
    a = ap.parse_args()
    pip, sp, rk = 0.0001, a.spread_pips * 0.0001, a.risk
    d, o, h, l, c = load(a.symbol, "2003-01-01")
    res = {
        "Conq": backtest(d, o, h, l, c, Params(risk_pct=rk, spread=sp)),
        "Conq+B": backtest(d, o, h, l, c, Params(risk_pct=rk, spread=sp, bishop=True)),
        "Chan": channel_backtest(d, o, h, l, c, ChannelParams(risk_pct=rk, spread=sp, buffer=3 * pip)),
        "Trend": trend_backtest(d, o, h, l, c, TrendParams(risk_pct=rk, spread=sp, buffer=3 * pip)),
    }
    res["AllAgr"] = combine({k: res[k].trades for k in ("Conq+B", "Chan", "Trend")}, "all", d, risk_pct=rk)
    years = sorted({x[:4] for x in d})
    tab = {m: by_year(r, years) for m, r in res.items()}
    first, last = {}, {}
    for x, cl in zip(d, c):
        first.setdefault(x[:4], cl); last[x[:4]] = cl
    print(f"{a.symbol} per-year, spread {a.spread_pips} pip, risk {rk:.2%}.  cell = trades / PF / return")
    print(f"{'year':<6}{'move':>7} | " + " | ".join(f"{m:^18}" for m in res))
    for y in years:
        mv = last[y] / first[y] - 1
        cells = []
        for m in res:
            n, pf, r = tab[m][y]
            cells.append(f"{n:>3} {('%.2f' % pf) if pf != float('inf') else ' inf':>5} {r:>+6.1%}")
        print(f"{y:<6}{mv:>+7.1%} | " + " | ".join(f"{x:^18}" for x in cells))
    print("\nyears with PF > 1:", {m: sum(1 for y in years if tab[m][y][0] and tab[m][y][1] > 1) for m in res}, "of", len(years))


if __name__ == "__main__":
    main()
