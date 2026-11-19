"""55/20 channel breakout (Courtney Smith, ch. 3), daily bars, with switchable add-ons.

Rules (see notes/indicators.md section 4):
  Entry : buy stop 3 pips above the 55-day high / sell stop 3 pips below the 55-day low (levels as of the previous close).
  Exit  : stop at the opposite 20-day channel; flat afterwards (no reversal).
  Add-ons (each switchable): ADX filter (take a signal only if ADX rose vs the prior day), Bishop (ADX>40 then down -> exit),
          Rejection Rule (5-day flat channel before the breakout; exit on the close if the close is back through the
          breakout level on the breakout day or the next day), Last Bar stop (stop at the breakout bar's low/high,
          switching to the 20-day channel once that is more favourable).
Assumptions beyond the book:
  C1. Bishop exit executes at the next open (signal known at the close).
  C2. ADX filter reads ADX at the previous close when the order is placed.
  C3. Position size uses an estimated stop = max(previous bar's low, 20-day low) for longs (mirror for shorts),
      because the true last-bar stop is unknown until the breakout bar closes. Realised loss can exceed the risk %.
  C4. If both entry levels are touched on one bar, the level nearer the open is taken.
  C5. ADX period 14 (Wilder standard; the book gives none).
"""
from dataclasses import dataclass

from .strategy import Result, Trade, adx


@dataclass
class ChannelParams:
    entry_len: int = 55
    exit_len: int = 20
    buffer: float = 0.0003         # 3 pips beyond the channel (4-decimal pairs)
    spread: float = 0.0001
    slippage: float = 0.0
    risk_pct: float = 0.005
    max_leverage: float = 10.0
    start_equity: float = 10_000.0
    adx_filter: bool = True
    bishop: bool = True
    rejection: bool = True
    last_bar: bool = True
    adx_len: int = 14
    bishop_level: float = 40.0


def _roll(vals, n, fn):
    return [fn(vals[i - n + 1:i + 1]) if i >= n - 1 else None for i in range(len(vals))]


def channel_backtest(dates, op, hi, lo, cl, p: ChannelParams = ChannelParams()) -> Result:
    n = len(cl)
    H_in, L_in = _roll(hi, p.entry_len, max), _roll(lo, p.entry_len, min)
    H_out, L_out = _roll(hi, p.exit_len, max), _roll(lo, p.exit_len, min)
    adx_v = adx(hi, lo, cl, p.adx_len) if (p.adx_filter or p.bishop) else None
    half = p.spread / 2 + p.slippage
    equity, res, pos = p.start_equity, Result(), None
    start = p.entry_len  # first bar with a full channel at i-1

    def close_pos(i, px, reason):
        nonlocal equity, pos
        t = pos["trade"]
        t.exit_date, t.exit, t.reason = dates[i], px, reason
        t.pnl = pos["side"] * (px - t.entry) * t.units
        equity += t.pnl
        res.trades.append(t)
        pos = None

    for i in range(n):
        # pending Bishop exit from yesterday's close
        if pos is not None and pos.get("exit_open"):
            close_pos(i, op[i] - pos["side"] * half, "bishop")

        # entry (orders were placed at yesterday's close)
        if pos is None and i >= start + 1 and H_in[i - 1] is not None:
            ok = True
            if p.adx_filter:
                a1, a2 = adx_v[i - 1], adx_v[i - 2]
                ok = a1 is not None and a2 is not None and a1 > a2
            if ok:
                buy, sell = H_in[i - 1] + p.buffer, L_in[i - 1] - p.buffer
                lh, sh = hi[i] >= buy, lo[i] <= sell
                side = 0
                if lh and sh:
                    side = 1 if abs(op[i] - buy) <= abs(op[i] - sell) else -1  # C4
                elif lh:
                    side = 1
                elif sh:
                    side = -1
                if side:
                    level = H_in[i - 1] if side == 1 else L_in[i - 1]
                    entry = (max(op[i], buy) if side == 1 else min(op[i], sell)) + side * half
                    est = max(lo[i - 1], L_out[i - 1]) if side == 1 else min(hi[i - 1], H_out[i - 1])
                    dist = max(abs(entry - est), 1e-9)                                   # C3
                    units = min(p.risk_pct * equity / dist, p.max_leverage * equity / entry)
                    # five-day condition: entry channel flat/declining (longs) or flat/rising (shorts) for 5 days
                    five = all((H_in[k] <= H_in[k - 1]) if side == 1 else (L_in[k] >= L_in[k - 1])
                               for k in range(i - 5, i) if H_in[k - 1] is not None)
                    if side == 1:
                        lastbar = lo[i] if lo[i] < level else lo[i - 1]
                    else:
                        lastbar = hi[i] if hi[i] > level else hi[i - 1]
                    pos = dict(side=side, level=level, five=five, lastbar=lastbar, idx=i, units=units,
                               trade=Trade(side, dates[i], entry, units))
                    pos["trade"].stop0 = est

        # intraday stop
        if pos is not None:
            s = pos["side"]
            chan = L_out[i - 1] if s == 1 else H_out[i - 1]
            if p.last_bar and pos["idx"] < i:
                stop = max(chan, pos["lastbar"]) if s == 1 else min(chan, pos["lastbar"])
            else:
                stop = chan
            hit = (lo[i] <= stop) if s == 1 else (hi[i] >= stop)
            if hit:
                gapped = (op[i] <= stop) if s == 1 else (op[i] >= stop)
                close_pos(i, (op[i] if gapped else stop) - s * half, "stop")

        # close-based exits
        if pos is not None:
            s = pos["side"]
            if p.rejection and pos["five"] and i - pos["idx"] <= 1:
                if (s == 1 and cl[i] < pos["level"]) or (s == -1 and cl[i] > pos["level"]):
                    close_pos(i, cl[i] - s * half, "reject")
        if pos is not None and p.bishop and i > 0 and adx_v[i] is not None and adx_v[i - 1] is not None:
            if adx_v[i - 1] > p.bishop_level and adx_v[i] < adx_v[i - 1]:
                pos["exit_open"] = True

        mtm = equity + (pos["side"] * (cl[i] - pos["trade"].entry) * pos["units"] if pos else 0.0)
        res.equity_curve.append((dates[i], mtm))

    if pos is not None:
        close_pos(n - 1, cl[-1] - pos["side"] * half, "open")
    return res
