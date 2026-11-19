"""Usage: python3 -m conqueror.run_backtest [SYMBOL] [--spread-pips 1.0] [--risk 0.01] [--from 2005-01-01]"""
import argparse
import csv
from pathlib import Path

from .strategy import Params, backtest, stats

from . import registry, warehouse

DATA_DIR = warehouse.OHLC_DIR


def load(symbol, start):
    d, o, h, l, c = [], [], [], [], []
    with (DATA_DIR / f"{symbol}.csv").open() as f:
        for r in csv.DictReader(f):
            if r["date"] >= start:
                d.append(r["date"]); o.append(float(r["open"])); h.append(float(r["high"]))
                l.append(float(r["low"])); c.append(float(r["close"]))
    return d, o, h, l, c


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbol", nargs="?", default="EURUSD")
    ap.add_argument("--spread-pips", type=float, default=1.0)
    ap.add_argument("--risk", type=float, default=0.005, help="fraction of equity per trade; book starts at 0.5 percent")
    ap.add_argument("--from", dest="start", default="2003-01-01")
    ap.add_argument("--trades", action="store_true")
    ap.add_argument("--bishop", action="store_true", help="apply the Bishop (ADX>40 then down) exit")
    ap.add_argument("--adx-len", type=int, default=14)
    ap.add_argument("--tighten", default="cumulative", choices=["cumulative", "current", "none", "exit2", "exit2_events"])
    a = ap.parse_args()
    d, o, h, l, c = load(a.symbol, a.start)
    pip = 0.01 if a.symbol.endswith("JPY") else 0.0001
    print(f"{a.symbol} {d[0]}..{d[-1]} ({len(d)} bars)")
    for sp in sorted({0.0, a.spread_pips, a.spread_pips * 3}):
        p = Params(risk_pct=a.risk, spread=sp * pip, tighten=a.tighten, bishop=a.bishop, adx_len=a.adx_len)
        r = backtest(d, o, h, l, c, p)
        s = stats(r, p)
        registry.log_run("conqueror", f"{a.tighten}{'+bishop' if a.bishop else ''}", a.symbol, f"{d[0]}..{d[-1]}", d[-1], vars(p), s)
        print(f"spread {sp:>4.1f} pip | trades {s['trades']:>3} win {s['win_rate']:.0%} "
              f"avgW {s['avg_win']:.0f} avgL {s['avg_loss']:.0f} PF {s['profit_factor']:.2f} | "
              f"ret {s['total_return']:+.0%} CAGR {s['cagr']:+.1%} maxDD {s['max_drawdown']:.0%}")
        if a.trades and sp == a.spread_pips:
            for t in r.trades:
                print(f"  {t.entry_date} {'L' if t.side == 1 else 'S'} {t.entry:.5f} -> {t.exit_date} {t.exit:.5f} {t.pnl:+.0f} {t.reason}")


if __name__ == "__main__":
    main()
