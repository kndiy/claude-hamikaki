"""Stochastic 50-line trend system (Courtney Smith, ch. 5 p. 81, "Go with the trend"), daily bars.

Book rules (notes/indicators.md section 8):
  raw K = (C - L) / (H - L) * 100 over the last 14 days (H, L = highest high / lowest low of those days);
  %K = 3-day MA of raw K; %D = 3-day MA of %K (%D is not needed for this rule).
  Buy when %K crosses above 50, sell when %K crosses below 50; always in the market.
Not in the book (the author names no stop for this system) -> assumptions, labelled:
  S1. A disaster stop is added for sizing and protection: 2 x ATR(40) from entry (same distance as the Conqueror I stop),
      fixed at entry (not trailed). Without a stop the 0.5% risk rule cannot size a position.
  S2. After a stop-out the system stays flat until the next 50-line cross (it does not re-enter on the same trend).
  S3. The signal is read at the close and executed at the next open (as for the other methods).
Not implemented (discretionary or undefined in the book): divergences, left/right-handed crossovers, failures, the
  "crossover in the direction of the trend after a retracement" entry (needs a trend definition the chapter does not give).
"""
from dataclasses import dataclass

from .strategy import Result, Trade, sma, true_ranges


@dataclass
class StochParams:
    length: int = 14
    k_smooth: int = 3
    level: float = 50.0
    atr_len: int = 40
    stop_mult: float = 2.0
    spread: float = 0.0001
    slippage: float = 0.0
    risk_pct: float = 0.005
    max_leverage: float = 10.0
    start_equity: float = 10_000.0


def stoch_k(hi, lo, cl, length=14, smooth=3):
    """%K: smooth-day MA of raw K. None until defined. A zero-range window gives raw K = 50 (undefined otherwise)."""
    n = len(cl)
    raw = [None] * n
    for i in range(length - 1, n):
        h, l = max(hi[i - length + 1:i + 1]), min(lo[i - length + 1:i + 1])
        raw[i] = 50.0 if h == l else (cl[i] - l) / (h - l) * 100.0
    out = [None] * n
    for i in range(length - 1 + smooth - 1, n):
        out[i] = sum(raw[i - smooth + 1:i + 1]) / smooth
    return out


def stoch_backtest(dates, op, hi, lo, cl, p: StochParams = StochParams()) -> Result:
    n = len(cl)
    k = stoch_k(hi, lo, cl, p.length, p.k_smooth)
    atr = sma(true_ranges(hi, lo, cl), p.atr_len)
    half = p.spread / 2 + p.slippage
    equity, res, pos, pending = p.start_equity, Result(), None, 0

    def close_pos(i, px, reason):
        nonlocal equity, pos
        t = pos["trade"]
        t.exit_date, t.exit, t.reason = dates[i], px, reason
        t.pnl = pos["side"] * (px - t.entry) * t.units
        equity += t.pnl
        res.trades.append(t)
        pos = None

    for i in range(n):
        if pending:
            if pos is not None:
                close_pos(i, op[i] - pos["side"] * half, "reverse")
            if atr[i - 1] is not None:
                side = pending
                entry = op[i] + side * half
                dist = p.stop_mult * atr[i - 1]                                    # S1
                units = min(p.risk_pct * equity / dist, p.max_leverage * equity / entry)
                tr = Trade(side, dates[i], entry, units)
                tr.stop0 = entry - side * dist
                pos = dict(side=side, stop=tr.stop0, units=units, trade=tr)
            pending = 0

        if pos is not None:
            s = pos["side"]
            if (s == 1 and lo[i] <= pos["stop"]) or (s == -1 and hi[i] >= pos["stop"]):
                gapped = (op[i] <= pos["stop"]) if s == 1 else (op[i] >= pos["stop"])
                close_pos(i, (op[i] if gapped else pos["stop"]) - s * half, "stop")

        if i > 0 and k[i] is not None and k[i - 1] is not None:
            side = 1 if (k[i - 1] <= p.level < k[i]) else -1 if (k[i - 1] >= p.level > k[i]) else 0
            if side and (pos is None or pos["side"] != side):
                pending = side

        mtm = equity + (pos["side"] * (cl[i] - pos["trade"].entry) * pos["units"] if pos else 0.0)
        res.equity_curve.append((dates[i], mtm))

    if pos is not None:
        close_pos(n - 1, cl[-1] - pos["side"] * half, "open")
    return res
