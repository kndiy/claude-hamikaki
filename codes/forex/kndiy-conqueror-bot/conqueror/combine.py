"""Chapter 10 combination (Courtney Smith, pp. 157-160) built from the three methods' trade lists.

  A  independent: each method keeps its own position and 0.5% risk; P&L adds up ("diversifying through time").
  B  all agree : one position; enter when all methods are on the same side, exit when any one exits
                 (the book's small-account variant, p. 159-160).
  C  majority  : one position; long/short while at least 2 of 3 methods are on that side ("weight of evidence", p. 159).
Sizing for B and C: risk% x equity / the TIGHTEST initial stop distance among the methods on that side (the book says to
use the tightest stop); capped by max leverage. Prices are taken from the methods' own fills (costs already included).
Assumptions: K1. B/C enter at the price of the method that completes the agreement; exit at the price of the method that
breaks it. K2. A's methods are each sized on their own equity path (approximation of one shared account).
K3. B/C equity curve is realised-only (no mark-to-market), so drawdown is understated.
"""
from .strategy import Result, Trade


def independent(results: dict, start_equity=10_000.0) -> Result:
    out = Result()
    for r in results.values():
        out.trades.extend(r.trades)
    out.trades.sort(key=lambda t: t.entry_date)
    curves = [r.equity_curve for r in results.values()]
    for k in range(len(curves[0])):
        out.equity_curve.append((curves[0][k][0], start_equity + sum(c[k][1] - start_equity for c in curves)))
    return out


def _events(trades_by_method):
    ev = []
    for m, trs in trades_by_method.items():
        for t in trs:
            ev.append((t.entry_date, 1, m, "in", t))
            ev.append((t.exit_date, 2 if t.exit_date == t.entry_date else 0, m, "out", t))
    ev.sort(key=lambda e: (e[0], e[1]))
    return ev


def combine(trades_by_method: dict, mode: str, dates, risk_pct=0.005, max_leverage=10.0, start_equity=10_000.0) -> Result:
    assert mode in ("all", "majority")
    names = list(trades_by_method)
    need = len(names) if mode == "all" else len(names) // 2 + 1
    open_t = {m: None for m in names}
    equity, pos, res = start_equity, None, Result()

    def state():
        longs = sum(1 for t in open_t.values() if t is not None and t.side == 1)
        shorts = sum(1 for t in open_t.values() if t is not None and t.side == -1)
        return 1 if longs >= need else (-1 if shorts >= need else 0)

    for date, _, m, kind, t in _events(trades_by_method):
        if kind == "in":
            open_t[m] = t
            price = t.entry
        else:
            if open_t[m] is t:
                open_t[m] = None
            price = t.exit
        st = state()
        if pos is not None and st != pos["side"]:
            c = pos["trade"]
            c.exit_date, c.exit, c.reason = date, price, "combo"
            c.pnl = pos["side"] * (price - c.entry) * c.units
            equity += c.pnl
            res.trades.append(c)
            pos = None
        if pos is None and st != 0:
            dists = [abs(x.entry - x.stop0) for x in open_t.values() if x is not None and x.side == st and x.stop0]
            dist = min(dists) if dists else 0.0
            if dist > 0:
                units = min(risk_pct * equity / dist, max_leverage * equity / price)
                pos = dict(side=st, trade=Trade(st, date, price, units))
    if pos is not None:  # all component trades are closed by the backtests, so this is only a safety net
        c = pos["trade"]
        c.exit_date, c.exit, c.reason = dates[-1], c.entry, "open"
        res.trades.append(c)

    # realised equity curve on the data's dates
    idx, e, k = {}, start_equity, 0
    for c in sorted(res.trades, key=lambda x: x.exit_date):
        idx.setdefault(c.exit_date, []).append(c.pnl)
    for d in dates:
        e += sum(idx.get(d, []))
        res.equity_curve.append((d, e))
    return res
