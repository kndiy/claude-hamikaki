"""Usage: python3 -m conqueror.run_channel [SYMBOL] [--spread-pips 1.0] [--risk 0.005] [--from 2003-01-01] [--trades]"""
import argparse

from .channel import ChannelParams, channel_backtest
from . import registry
from .run_backtest import load
from .strategy import stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbol", nargs="?", default="EURUSD")
    ap.add_argument("--spread-pips", type=float, default=1.0)
    ap.add_argument("--risk", type=float, default=0.005)
    ap.add_argument("--from", dest="start", default="2003-01-01")
    ap.add_argument("--trades", action="store_true")
    a = ap.parse_args()
    d, o, h, l, c = load(a.symbol, a.start)
    pip = 0.01 if a.symbol.endswith("JPY") else 0.0001
    base = dict(risk_pct=a.risk, spread=a.spread_pips * pip, buffer=3 * pip)
    off = dict(adx_filter=False, bishop=False, rejection=False, last_bar=False)
    variants = [
        ("plain 55/20", dict(off)),
        ("+ ADX filter", {**off, "adx_filter": True}),
        ("+ Bishop", {**off, "bishop": True}),
        ("+ Rejection rule", {**off, "rejection": True}),
        ("+ Last bar stop", {**off, "last_bar": True}),
        ("all four (book)", dict()),
    ]
    print(f"{a.symbol} {d[0]}..{d[-1]} ({len(d)} bars), spread {a.spread_pips} pip, risk {a.risk:.2%}")
    for name, kw in variants:
        p = ChannelParams(**base, **kw)
        r = channel_backtest(d, o, h, l, c, p)
        s = stats(r, p)
        registry.log_run("channel", name, a.symbol, f"{d[0]}..{d[-1]}", d[-1], vars(p), s)
        print(f"{name:<18}| trades {s['trades']:>3} win {s['win_rate']:.0%} avgW {s['avg_win']:.0f} avgL {s['avg_loss']:.0f} "
              f"PF {s['profit_factor']:.2f} | ret {s['total_return']:+.0%} CAGR {s['cagr']:+.1%} maxDD {s['max_drawdown']:.0%}")
        if a.trades and name.startswith("all"):
            for t in r.trades:
                print(f"  {t.entry_date} {'L' if t.side == 1 else 'S'} {t.entry:.5f} -> {t.exit_date} {t.exit:.5f} {t.pnl:+.0f} {t.reason}")
    p = ChannelParams(**{**base, "spread": 0.0})
    s0 = stats(channel_backtest(d, o, h, l, c, p), p)
    p = ChannelParams(**{**base, "spread": 3 * a.spread_pips * pip})
    s3 = stats(channel_backtest(d, o, h, l, c, p), p)
    print(f"all four, spread 0 pip: PF {s0['profit_factor']:.2f} ret {s0['total_return']:+.0%} | spread {3*a.spread_pips:.0f} pip: PF {s3['profit_factor']:.2f} ret {s3['total_return']:+.0%}")


if __name__ == "__main__":
    main()
