"""Conqueror (Courtney Smith, 'How to Make a Living Trading Foreign Exchange', ch. 4), daily bars.

Pure functions plus a bar-by-bar backtest. Standard library only.
Quote-currency must be USD (EURUSD, GBPUSD, AUDUSD ...) so P&L is in account currency.

Assumptions (the chapter is ambiguous; see CLAUDE.md):
  A1. ATR(40) = simple average of the last 40 true ranges.
  A2. A "flip" = any of the 3 conditions changing sign vs the previous bar while in a trade.
      Each flip cuts the ATR multiple by 1/3 of the ORIGINAL (2 -> 1.33 -> 0.67 -> 0), kept for the trade.
  A3. Signal on close of day t; entry at open of day t+1 (conservative).
  A4. Stops checked intraday on the next bar (fill at stop, or at the open if gapped through),
      then updated from that bar's close. Stop never moves against the trade.
  A5. Re-entry after any exit needs the 3-way alignment in that direction to break and re-form.
  A7. Bishop (owner decision): ADX(14) > 40 and falling on a close -> exit at next open, both directions.
  A6. Position size = risk% * equity / initial stop distance, capped by max leverage.
"""
from dataclasses import dataclass, field

MA_LEN, MA_LAG, MOMENTUM_LEN, ATR_LEN = 10, 10, 40, 40


def true_ranges(high, low, close):
    tr = [high[0] - low[0]]
    for i in range(1, len(close)):
        th = max(high[i], close[i - 1])
        tl = min(low[i], close[i - 1])
        tr.append(th - tl)
    return tr


def sma(values, n):
    out = [None] * len(values)
    s = 0.0
    for i, v in enumerate(values):
        s += v
        if i >= n:
            s -= values[i - n]
        if i >= n - 1:
            out[i] = s / n
    return out


def adx(high, low, close, n=14):
    """Wilder's ADX. Returns list aligned to bars; None until 2n bars. (Book gives no period: n=14 is an assumption.)"""
    L = len(close)
    out = [None] * L
    if L < 2 * n + 1:
        return out
    tr, pdm, mdm = [0.0] * L, [0.0] * L, [0.0] * L
    for i in range(1, L):
        up, dn = high[i] - high[i - 1], low[i - 1] - low[i]
        pdm[i] = up if (up > dn and up > 0) else 0.0
        mdm[i] = dn if (dn > up and dn > 0) else 0.0
        tr[i] = max(high[i] - low[i], abs(high[i] - close[i - 1]), abs(low[i] - close[i - 1]))
    str_, sp, sm = sum(tr[1:n + 1]), sum(pdm[1:n + 1]), sum(mdm[1:n + 1])
    dx = [None] * L
    for i in range(n, L):
        if i > n:
            str_ = str_ - str_ / n + tr[i]
            sp = sp - sp / n + pdm[i]
            sm = sm - sm / n + mdm[i]
        pdi, mdi = (100 * sp / str_, 100 * sm / str_) if str_ else (0.0, 0.0)
        dx[i] = 100 * abs(pdi - mdi) / (pdi + mdi) if (pdi + mdi) else 0.0
    a = sum(dx[n:2 * n]) / n
    out[2 * n - 1] = a
    for i in range(2 * n, L):
        a = (a * (n - 1) + dx[i]) / n
        out[i] = a
    return out


def conditions(close):
    """Return list of (c1, c2, c3) per bar; None until enough history. Each is a float, >0 bullish."""
    ma = sma(close, MA_LEN)
    out = []
    for i in range(len(close)):
        if i < MOMENTUM_LEN or ma[i - MA_LAG] is None:
            out.append(None)
            continue
        out.append((close[i] - ma[i], ma[i] - ma[i - MA_LAG], close[i] - close[i - MOMENTUM_LEN]))
    return out


def alignment(c):
    """+1 all positive, -1 all negative, 0 otherwise."""
    if c is None:
        return 0
    if all(x > 0 for x in c):
        return 1
    if all(x < 0 for x in c):
        return -1
    return 0


@dataclass
class Params:
    atr_mult: float = 2.0
    risk_pct: float = 0.01
    max_leverage: float = 10.0
    spread: float = 0.0001        # full spread in price units (1 pip = 0.0001 for EURUSD)
    slippage: float = 0.0         # per side, price units
    start_equity: float = 10_000.0
    # How the "reduce the stop by a third when a condition changes sign" rule is read (A2):
    #   "cumulative": every sign change since entry cuts 1/3 of the original multiple, kept for the trade
    #   "current":    multiple = original * (1 - k/3), k = conditions currently against the position
    #   "none":       no tightening (plain 2 x ATR trail), for comparison only
    #   OWNER VARIANT (not in the book): trailing stop stays at the full multiple, plus a discretionary-style cut:
    #   "exit2":        exit at next open when >=2 conditions are currently against the position
    #   "exit2_events": exit at next open after >=2 sign changes since entry
    tighten: str = "cumulative"
    # Bishop (ch. 2, pp. 28-30): ADX above `bishop_level` then ticks down -> exit all positions.
    # Owner decision 2026-11-16: apply it to the Conqueror (the book does not say). Exit at next open (A7).
    bishop: bool = False
    adx_len: int = 14            # GAP: book gives no ADX period
    bishop_level: float = 40.0


@dataclass
class Trade:
    side: int
    entry_date: str
    entry: float
    units: float
    exit_date: str = ""
    exit: float = 0.0
    pnl: float = 0.0
    reason: str = ""
    stop0: float = 0.0   # initial stop price at entry (used to size combined positions)


@dataclass
class Result:
    trades: list = field(default_factory=list)
    equity_curve: list = field(default_factory=list)  # (date, equity) marked at close


def backtest(dates, op, hi, lo, cl, p: Params = Params()) -> Result:
    n = len(cl)
    tr = true_ranges(hi, lo, cl)
    atr = sma(tr, ATR_LEN)
    conds = conditions(cl)
    adx_v = adx(hi, lo, cl, p.adx_len) if p.bishop else None
    half_cost = p.spread / 2 + p.slippage

    equity = p.start_equity
    res = Result()
    pos = None            # dict: side, units, stop, extreme, flips, prev_signs, trade
    pending = 0           # signal from previous close to enter at this open
    armed = {1: True, -1: True}

    for i in range(n):
        # 0) owner-variant cut decided at yesterday's close: exit at today's open
        if pos is not None and pos.get("exit_open"):
            s0 = pos["side"]
            px = op[i] - s0 * half_cost
            t = pos["trade"]
            t.exit_date, t.exit, t.reason = dates[i], px, pos.get("exit_reason", "cut")
            t.pnl = s0 * (px - t.entry) * t.units
            equity += t.pnl
            res.trades.append(t)
            pos = None

        # 1) enter at today's open from yesterday's signal
        if pos is None and pending != 0 and i > 0 and atr[i - 1] is not None:
            side = pending
            entry = op[i] + side * half_cost
            dist = p.atr_mult * atr[i - 1]
            ref = cl[i - 1]
            stop = ref - side * dist
            units = min(p.risk_pct * equity / dist, p.max_leverage * equity / entry)
            t = Trade(side, dates[i], entry, units)
            t.stop0 = stop
            pos = dict(side=side, units=units, stop=stop, extreme=ref, flips=0,
                       signs=[x > 0 for x in conds[i - 1]], trade=t, atr_mult=p.atr_mult)
            armed[side] = False
        pending = 0

        # 2) stop check during today's bar
        if pos is not None:
            s = pos["side"]
            hit = (lo[i] <= pos["stop"]) if s == 1 else (hi[i] >= pos["stop"])
            if hit:
                gapped = (op[i] <= pos["stop"]) if s == 1 else (op[i] >= pos["stop"])
                px = (op[i] if gapped else pos["stop"]) - s * half_cost
                t = pos["trade"]
                t.exit_date, t.exit, t.reason = dates[i], px, "stop"
                t.pnl = s * (px - t.entry) * t.units
                equity += t.pnl
                res.trades.append(t)
                pos = None

        # 3) end-of-day: update stop / look for signal
        if pos is not None and conds[i] is not None and atr[i] is not None:
            s = pos["side"]
            signs = [x > 0 for x in conds[i]]
            pos["flips"] += sum(a != b for a, b in zip(signs, pos["signs"]))
            pos["signs"] = signs
            if p.tighten == "current":
                against = sum((x < 0) if s == 1 else (x > 0) for x in conds[i])
                k = against
            elif p.tighten in ("none", "exit2", "exit2_events"):
                k = 0
                against = sum((x < 0) if s == 1 else (x > 0) for x in conds[i])
                if (p.tighten == "exit2" and against >= 2) or (p.tighten == "exit2_events" and pos["flips"] >= 2):
                    pos["exit_open"] = True
            else:
                k = pos["flips"]
            pos["extreme"] = max(pos["extreme"], cl[i]) if s == 1 else min(pos["extreme"], cl[i])
            mult = pos["atr_mult"] * max(0.0, 1 - k / 3)
            new = pos["extreme"] - s * mult * atr[i]
            pos["stop"] = max(pos["stop"], new) if s == 1 else min(pos["stop"], new)

        if pos is not None and p.bishop and i > 0 and adx_v[i] is not None and adx_v[i - 1] is not None:
            if adx_v[i - 1] > p.bishop_level and adx_v[i] < adx_v[i - 1]:
                pos["exit_open"] = True
                pos["exit_reason"] = "bishop"

        a = alignment(conds[i])
        for d in (1, -1):
            if a != d:
                armed[d] = True
        if pos is None and a != 0 and armed[a]:
            pending = a

        mtm = equity
        if pos is not None:
            mtm += pos["side"] * (cl[i] - pos["trade"].entry) * pos["units"]
        res.equity_curve.append((dates[i], mtm))

    if pos is not None:  # close open trade at last close for reporting
        t = pos["trade"]
        px = cl[-1] - pos["side"] * half_cost
        t.exit_date, t.exit, t.reason = dates[-1], px, "open"
        t.pnl = pos["side"] * (px - t.entry) * t.units
        res.trades.append(t)
    return res


def stats(res: Result, p: Params):
    tr = res.trades
    wins = [t.pnl for t in tr if t.pnl > 0]
    losses = [-t.pnl for t in tr if t.pnl <= 0]
    final = res.equity_curve[-1][1] if res.equity_curve else p.start_equity
    peak, mdd = p.start_equity, 0.0
    for _, e in res.equity_curve:
        peak = max(peak, e)
        mdd = max(mdd, (peak - e) / peak)
    years = max(len(res.equity_curve) / 260, 1e-9)
    return {
        "trades": len(tr),
        "win_rate": len(wins) / len(tr) if tr else 0.0,
        "avg_win": sum(wins) / len(wins) if wins else 0.0,
        "avg_loss": sum(losses) / len(losses) if losses else 0.0,
        "profit_factor": (sum(wins) / sum(losses)) if losses and sum(losses) > 0 else float("inf"),
        "final_equity": final,
        "total_return": final / p.start_equity - 1,
        "cagr": (final / p.start_equity) ** (1 / years) - 1 if final > 0 else -1.0,
        "max_drawdown": mdd,
    }
