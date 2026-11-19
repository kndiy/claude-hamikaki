"""The owner-agreed ACCEPT RULE (2026-11-16), applied identically to every method. Written before any new method was tested.

A method passes only if, with 1 pip spread, 0.5% risk, fixed textbook parameters (no tuning):
  (1) pooled profit factor over EURUSD, GBPUSD, AUDUSD, NZDUSD is above 1.20,
  (2) pooled net P&L is positive in BOTH halves of the history (split at the midpoint date of the data, by trade ENTRY date),
  (3) profit factor is above 1.0 on at least 3 of the 4 pairs.
A failure is a rejection, not a reason to re-tune. Every (method, pair) run is appended to warehouse/results/runs.csv.
All data up to 2026-11-16 counts as already seen: passing here is a necessary condition, not proof; the real test is forward data.

Usage: python3 -m conqueror.accept [method ...]    (no names = every registered method)
"""
import sys
from datetime import date

from . import registry
from .channel import ChannelParams, channel_backtest
from .run_backtest import load
from .strategy import Params, backtest, stats
from .trend import TrendParams, trend_backtest

PAIRS = ["EURUSD", "GBPUSD", "AUDUSD", "NZDUSD"]
PF_MIN, PAIRS_MIN = 1.20, 3
START = "2003-01-01"
PIP = 0.0001  # all four pairs are USD-quote, 4-decimal


def _pf(wins, losses):
    return wins / losses if losses > 0 else float("inf")


def _split_date(dates_by_pair):
    first = min(d[0] for d in dates_by_pair.values())
    last = max(d[-1] for d in dates_by_pair.values())
    # calendar midpoint of the data; deterministic, not chosen by looking at results
    a, b = date.fromisoformat(first), date.fromisoformat(last)
    return (a + (b - a) / 2).isoformat()


def evaluate(name, factory, pairs=PAIRS, spread_pips=1.0, risk=0.005, log=True, note=""):
    """factory(pip, risk, spread_pips) -> (backtest_fn(d,o,h,l,c) -> Result, params_object)."""
    data = {p: load(p, START) for p in pairs}
    split = _split_date({p: v[0] for p, v in data.items()})
    runs = {}
    for p, (d, o, h, l, c) in data.items():
        fn, params = factory(PIP, risk, spread_pips)
        res = fn(d, o, h, l, c)
        runs[p] = res
        if log:
            registry.log_run(name, "accept", p, f"{d[0]}..{d[-1]}", d[-1], vars(params), stats(res, params), note)
    W = sum(t.pnl for r in runs.values() for t in r.trades if t.pnl > 0)
    L = -sum(t.pnl for r in runs.values() for t in r.trades if t.pnl <= 0)
    n = sum(len(r.trades) for r in runs.values())
    halves = [sum(t.pnl for r in runs.values() for t in r.trades if (t.entry_date < split) == first)
              for first in (True, False)]
    per = {p: _pf(sum(t.pnl for t in r.trades if t.pnl > 0), -sum(t.pnl for t in r.trades if t.pnl <= 0))
           for p, r in runs.items()}
    pooled = _pf(W, L)
    checks = {
        f"pooled PF > {PF_MIN}": pooled > PF_MIN,
        "positive in both halves": halves[0] > 0 and halves[1] > 0,
        f"PF > 1 on >= {PAIRS_MIN} of {len(pairs)} pairs": sum(1 for v in per.values() if v > 1.0) >= PAIRS_MIN,
    }
    return dict(name=name, trades=n, pooled_pf=pooled, halves=halves, split=split, per_pair=per, checks=checks,
                passed=all(checks.values()))


def show(r):
    print(f"\n{r['name']}: trades {r['trades']}, pooled PF {r['pooled_pf']:.2f}, half P&L (before/after {r['split']}) "
          f"{r['halves'][0]:+.0f} / {r['halves'][1]:+.0f}")
    print("  per pair PF: " + "  ".join(f"{p} {v:.2f}" for p, v in r["per_pair"].items()))
    for k, ok in r["checks"].items():
        print(f"  [{'PASS' if ok else 'fail'}] {k}")
    print(f"  => {'ACCEPTED (necessary, not sufficient)' if r['passed'] else 'REJECTED'}")


def methods():
    """Registered methods. Add new ones here with their fixed textbook parameters BEFORE running."""
    from .stoch import StochParams, stoch_backtest

    def conq(pip, risk, sp):
        p = Params(risk_pct=risk, spread=sp * pip, bishop=True)
        return (lambda d, o, h, l, c: backtest(d, o, h, l, c, p)), p

    def chan(pip, risk, sp):
        p = ChannelParams(risk_pct=risk, spread=sp * pip, buffer=3 * pip)
        return (lambda d, o, h, l, c: channel_backtest(d, o, h, l, c, p)), p

    def trend(pip, risk, sp):
        p = TrendParams(risk_pct=risk, spread=sp * pip, buffer=3 * pip)
        return (lambda d, o, h, l, c: trend_backtest(d, o, h, l, c, p)), p

    def stoch(pip, risk, sp):
        p = StochParams(risk_pct=risk, spread=sp * pip)
        return (lambda d, o, h, l, c: stoch_backtest(d, o, h, l, c, p)), p

    return {"conqueror+bishop": conq, "channel (book)": chan, "trend (book)": trend, "stochastic 50-cross (book)": stoch}


def main():
    reg = methods()
    names = sys.argv[1:] or list(reg)
    for nm in names:
        show(evaluate(nm, reg[nm]))
    print(f"\nTotal logged runs so far: {registry.count()}")


if __name__ == "__main__":
    main()
