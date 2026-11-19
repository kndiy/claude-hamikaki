"""Trend Analysis (Courtney Smith, ch. 2, pp. 14-30), daily bars, ADX filter and Bishop switchable.

Rules (see notes/indicators.md sections 1-3):
  Swing high: bar whose high is above the previous bar's and the next bar's. Rank = number of consecutive lower highs to its left.
              Only rank >= 3 counts as significant. Swing low is the mirror. A swing is confirmed one bar later.
  Bull = last two significant highs rising AND last two lows rising; bear = both falling; otherwise neutral.
  Orders (placed at each close, from swings confirmed so far):
      buy  stop above the latest significant high   if the last two lows are rising (L2 > L1)
      sell stop below the latest significant low    if the last two highs are falling (H2 < H1)
      -> pennant (H falling, L rising) gives both; megaphone (H rising, L falling) gives neither.
  Stop: shorts just above the latest significant high, longs just below the latest significant low; only tightens.
  Exit when the market turns neutral: a long on a newly confirmed LOWER significant high; a short on a newly confirmed
      HIGHER significant low. Exit at the next open.
  ADX filter: take an entry only if ADX rose vs the prior day. Bishop: ADX > 40 then ticks down -> exit at next open.
Assumptions beyond the book:
  T1. The buffer (3 pips) is used both for entry orders and for stops (book gives no stop buffer).
  T2. Position size = risk% * equity / |entry - initial stop|; capped by max leverage.
  T3. A trade whose initial stop is not on the safe side of the entry is skipped.
  T4. If both orders are touched on one bar, the level nearer the open is taken.
  T5. ADX period 14 (Wilder standard; the book gives none).
"""
from dataclasses import dataclass

from .strategy import Result, Trade, adx


@dataclass
class TrendParams:
    min_rank: int = 3
    buffer: float = 0.0003
    spread: float = 0.0001
    slippage: float = 0.0
    risk_pct: float = 0.005
    max_leverage: float = 10.0
    start_equity: float = 10_000.0
    adx_filter: bool = True
    bishop: bool = True
    adx_len: int = 14
    bishop_level: float = 40.0


def trend_backtest(dates, op, hi, lo, cl, p: TrendParams = TrendParams()) -> Result:
    n = len(cl)
    adx_v = adx(hi, lo, cl, p.adx_len) if (p.adx_filter or p.bishop) else None
    half = p.spread / 2 + p.slippage
    equity, res, pos = p.start_equity, Result(), None
    highs, lows = [], []  # confirmed significant swings: (bar index, price)

    def close_pos(i, px, reason):
        nonlocal equity, pos
        t = pos["trade"]
        t.exit_date, t.exit, t.reason = dates[i], px, reason
        t.pnl = pos["side"] * (px - t.entry) * t.units
        equity += t.pnl
        res.trades.append(t)
        pos = None

    for i in range(n):
        if pos is not None and pos.get("exit_open"):
            close_pos(i, op[i] - pos["side"] * half, pos["exit_open"])

        # entry from orders placed at yesterday's close
        if pos is None and i >= 2:
            ok = True
            if p.adx_filter:
                a1, a2 = adx_v[i - 1], adx_v[i - 2]
                ok = a1 is not None and a2 is not None and a1 > a2
            cands = []
            if ok and len(lows) >= 2 and highs and lows[-1][1] > lows[-2][1]:
                cands.append((1, highs[-1][1] + p.buffer, lows[-1][1] - p.buffer))      # side, trigger, stop
            if ok and len(highs) >= 2 and lows and highs[-1][1] < highs[-2][1]:
                cands.append((-1, lows[-1][1] - p.buffer, highs[-1][1] + p.buffer))
            hit = [c for c in cands if (hi[i] >= c[1] if c[0] == 1 else lo[i] <= c[1])]
            if hit:
                side, trig, stop = min(hit, key=lambda c: abs(op[i] - c[1]))            # T4
                entry = (max(op[i], trig) if side == 1 else min(op[i], trig)) + side * half
                if (side == 1 and stop < entry) or (side == -1 and stop > entry):       # T3
                    dist = abs(entry - stop)
                    units = min(p.risk_pct * equity / dist, p.max_leverage * equity / entry)
                    pos = dict(side=side, stop=stop, units=units, trade=Trade(side, dates[i], entry, units))
                    pos["trade"].stop0 = stop

        # intraday stop
        if pos is not None:
            s = pos["side"]
            if (s == 1 and lo[i] <= pos["stop"]) or (s == -1 and hi[i] >= pos["stop"]):
                gapped = (op[i] <= pos["stop"]) if s == 1 else (op[i] >= pos["stop"])
                close_pos(i, (op[i] if gapped else pos["stop"]) - s * half, "stop")

        # end of day: confirm the swing at bar i-1 (needs bar i)
        j = i - 1
        new_high = new_low = False
        if j >= p.min_rank + 1:
            if hi[j] > hi[j - 1] and hi[j] > hi[i]:
                rank = 0
                k = j - 1
                while k >= 0 and hi[k] < hi[j]:
                    rank += 1
                    k -= 1
                if rank >= p.min_rank:
                    highs.append((j, hi[j])); new_high = True
            if lo[j] < lo[j - 1] and lo[j] < lo[i]:
                rank = 0
                k = j - 1
                while k >= 0 and lo[k] > lo[j]:
                    rank += 1
                    k -= 1
                if rank >= p.min_rank:
                    lows.append((j, lo[j])); new_low = True

        if pos is not None:
            s = pos["side"]
            if s == 1:
                if lows:
                    pos["stop"] = max(pos["stop"], lows[-1][1] - p.buffer)
                if new_high and len(highs) >= 2 and highs[-1][1] < highs[-2][1]:
                    pos["exit_open"] = "neutral"
            else:
                if highs:
                    pos["stop"] = min(pos["stop"], highs[-1][1] + p.buffer)
                if new_low and len(lows) >= 2 and lows[-1][1] > lows[-2][1]:
                    pos["exit_open"] = "neutral"
            if (p.bishop and i > 0 and adx_v[i] is not None and adx_v[i - 1] is not None
                    and adx_v[i - 1] > p.bishop_level and adx_v[i] < adx_v[i - 1]):
                pos["exit_open"] = "bishop"

        mtm = equity + (pos["side"] * (cl[i] - pos["trade"].entry) * pos["units"] if pos else 0.0)
        res.equity_curve.append((dates[i], mtm))

    if pos is not None:
        close_pos(n - 1, cl[-1] - pos["side"] * half, "open")
    return res
