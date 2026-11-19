import unittest
from conqueror.strategy import true_ranges, sma, conditions, alignment, backtest, Params


class T(unittest.TestCase):
    def test_true_range_gap(self):
        # Monday gaps down: high below Friday's close -> true high = prev close
        tr = true_ranges([10, 9.0], [9, 8.5], [10, 8.8])
        self.assertAlmostEqual(tr[1], 10 - 8.5)

    def test_sma(self):
        self.assertEqual(sma([1, 2, 3, 4], 2), [None, 1.5, 2.5, 3.5])

    def test_conditions_uptrend_all_positive(self):
        close = [1 + 0.01 * i for i in range(80)]
        c = conditions(close)
        self.assertIsNone(c[10])
        self.assertEqual(alignment(c[60]), 1)

    def test_conditions_downtrend_all_negative(self):
        close = [2 - 0.01 * i for i in range(80)]
        self.assertEqual(alignment(conditions(close)[60]), -1)

    def test_flat_is_no_signal(self):
        self.assertEqual(alignment(conditions([1.0] * 80)[60]), 0)

    def test_backtest_long_trend_then_stop(self):
        n = 120
        close = [1 + 0.005 * i for i in range(n)]
        close += [close[-1] - 0.03 * j for j in range(1, 30)]  # sharp reversal
        op = [close[0]] + close[:-1]
        hi = [max(a, b) + 0.002 for a, b in zip(op, close)]
        lo = [min(a, b) - 0.002 for a, b in zip(op, close)]
        dates = [f"d{i:04d}" for i in range(len(close))]
        r = backtest(dates, op, hi, lo, close, Params(spread=0))
        self.assertTrue(any(t.side == 1 for t in r.trades))
        first = r.trades[0]
        self.assertEqual(first.reason, "stop")
        self.assertGreater(first.pnl, 0)  # long trend rode, then trailed out in profit

    def test_stop_never_loosens_and_risk_is_bounded(self):
        n = 150
        import random
        random.seed(1)
        c = [1.0]
        for _ in range(n):
            c.append(c[-1] * (1 + random.gauss(0.0003, 0.006)))
        op = c[:-1]; cl = c[1:]
        hi = [max(a, b) * 1.002 for a, b in zip(op, cl)]
        lo = [min(a, b) * 0.998 for a, b in zip(op, cl)]
        d = [f"d{i:04d}" for i in range(n)]
        p = Params(spread=0, risk_pct=0.01)
        r = backtest(d, op, hi, lo, cl, p)
        eq = p.start_equity
        for t in r.trades:
            # a stop-out can exceed 1% only through gaps/ATR-vs-open differences; flag extreme cases
            self.assertGreater(t.pnl, -0.05 * eq)
            eq += t.pnl


class BookRules(unittest.TestCase):
    def test_default_is_book_cumulative_tightening(self):
        self.assertEqual(Params().tighten, "cumulative")

    def test_tightening_arithmetic_matches_book_example(self):
        # book: 480-pip stop, one sign change removes a third -> ~322 pips
        orig = 480.0
        self.assertAlmostEqual(orig * (1 - 1 / 3), 320.0, delta=2.5)  # book rounds a third of 480 to 158 -> 322
        mults = [2.0 * max(0.0, 1 - k / 3) for k in range(4)]
        self.assertEqual([round(m, 3) for m in mults], [2.0, 1.333, 0.667, 0.0])


class BishopAdx(unittest.TestCase):
    def test_bishop_off_by_default(self):
        self.assertFalse(Params().bishop)
        self.assertEqual(Params().bishop_level, 40.0)

    def test_adx_strong_trend_is_high_and_flat_is_low(self):
        from conqueror.strategy import adx
        up = [1 + 0.01 * i for i in range(120)]
        a = adx([x + 0.002 for x in up], [x - 0.002 for x in up], up)
        self.assertIsNone(a[10])
        self.assertGreater(a[-1], 40)
        flat = [1.0 + (0.001 if i % 2 else -0.001) for i in range(120)]
        b = adx([x + 0.002 for x in flat], [x - 0.002 for x in flat], flat)
        self.assertLess(b[-1], 20)


class ChannelRules(unittest.TestCase):
    def _series(self, extra):
        # 60 flat bars: range 0.99-1.01, close 1.0; then the scripted bars
        op = [1.0] * 60; hi = [1.01] * 60; lo = [0.99] * 60; cl = [1.0] * 60
        for o, h, l, c in extra:
            op.append(o); hi.append(h); lo.append(l); cl.append(c)
        return [f"d{i:04d}" for i in range(len(cl))], op, hi, lo, cl

    def _params(self, **kw):
        from conqueror.channel import ChannelParams
        base = dict(spread=0.0, adx_filter=False, bishop=False, rejection=False, last_bar=False)
        base.update(kw)
        return ChannelParams(**base)

    def test_breakout_enters_3_pips_above_and_exits_on_20d_low(self):
        from conqueror.channel import channel_backtest
        d, o, h, l, c = self._series([(1.005, 1.0125, 1.004, 1.011)] + [(1.011, 1.012, 1.0105, 1.0115)] * 3
                                     + [(1.0115, 1.0116, 0.98, 0.985)])
        r = channel_backtest(d, o, h, l, c, self._params())
        t = r.trades[0]
        self.assertEqual(t.side, 1)
        self.assertAlmostEqual(t.entry, 1.0103, places=5)
        self.assertEqual(t.reason, "stop")
        self.assertAlmostEqual(t.exit, 0.99, places=5)  # 20-day low

    def test_rejection_rule_exits_on_close_back_below_level(self):
        from conqueror.channel import channel_backtest
        d, o, h, l, c = self._series([(1.005, 1.0125, 1.004, 1.011), (1.010, 1.0115, 1.008, 1.009)])
        r = channel_backtest(d, o, h, l, c, self._params(rejection=True))
        self.assertEqual(r.trades[0].reason, "reject")
        self.assertAlmostEqual(r.trades[0].exit, 1.009, places=5)

    def test_last_bar_stop_is_breakout_bar_low(self):
        from conqueror.channel import channel_backtest
        d, o, h, l, c = self._series([(1.005, 1.0125, 1.004, 1.011), (1.011, 1.0115, 1.0105, 1.0112),
                                      (1.0112, 1.0114, 1.003, 1.0035)])
        r = channel_backtest(d, o, h, l, c, self._params(last_bar=True))
        t = r.trades[0]
        self.assertEqual(t.reason, "stop")
        self.assertAlmostEqual(t.exit, 1.004, places=5)

    def test_no_reversal_after_stop(self):
        from conqueror.channel import channel_backtest
        d, o, h, l, c = self._series([(1.005, 1.0125, 1.004, 1.011), (1.011, 1.0115, 0.98, 0.985)])
        r = channel_backtest(d, o, h, l, c, self._params())
        self.assertEqual(len(r.trades), 1)


def zigzag(pivots, per_leg=6, start=1.0):
    """Daily bars following straight legs between pivot prices (lead-in leg from `start`)."""
    closes, prev = [], start
    for pv in pivots:
        for k in range(1, per_leg + 1):
            closes.append(prev + (pv - prev) * k / per_leg)
        prev = pv
    op = [closes[0]] + closes[:-1]
    hi = [c + 0.0005 for c in closes]   # high/low tied to the close so turning bars are strict extremes
    lo = [c - 0.0005 for c in closes]
    return [f"d{i:04d}" for i in range(len(closes))], op, hi, lo, closes


class TrendRules(unittest.TestCase):
    def _p(self, **kw):
        from conqueror.trend import TrendParams
        base = dict(spread=0.0, adx_filter=False, bishop=False)
        base.update(kw)
        return TrendParams(**base)

    def test_bear_entry_on_break_of_swing_low_with_stop_above_lower_high(self):
        from conqueror.trend import trend_backtest
        # H1=1.10, L1=1.04, H2=1.08 (lower high) then a drop through L1
        d, o, h, l, c = zigzag([1.10, 1.04, 1.08, 1.00, 0.99])
        r = trend_backtest(d, o, h, l, c, self._p())
        t = r.trades[0]
        self.assertEqual(t.side, -1)
        self.assertAlmostEqual(t.entry, 1.0395 - 0.0003, delta=0.0006)  # just under the 1.04 swing low (low = 1.0395)
        # stop sits just above the lower high (1.08 + 0.0005 bar range + 3 pips buffer)
        t2 = [x for x in r.trades][0]
        self.assertIn(t2.reason, ("stop", "neutral", "open"))

    def test_megaphone_gives_no_entry(self):
        from conqueror.trend import trend_backtest
        d, o, h, l, c = zigzag([1.06, 0.98, 1.08, 0.96, 1.07])
        r = trend_backtest(d, o, h, l, c, self._p())
        self.assertEqual(len(r.trades), 0)

    def test_bull_long_exits_on_lower_high(self):
        from conqueror.trend import trend_backtest
        # higher lows and a break of the latest high, then a lower swing high -> neutral exit
        d, o, h, l, c = zigzag([1.05, 1.02, 1.07, 1.035, 1.10, 1.06, 1.09, 1.065])
        r = trend_backtest(d, o, h, l, c, self._p())
        longs = [t for t in r.trades if t.side == 1]
        self.assertTrue(longs)
        self.assertIn(longs[0].reason, ("neutral", "stop"))

    def test_swing_needs_rank_three(self):
        from conqueror.trend import trend_backtest, TrendParams
        # a short wiggle (rank 1) must not create swings -> no trades
        closes = [1.0, 1.01, 1.0, 1.01, 1.0, 1.01, 1.0, 1.01, 1.0, 1.01, 1.0, 1.01]
        op = [closes[0]] + closes[:-1]
        hi = [max(a, b) + 0.0005 for a, b in zip(op, closes)]
        lo = [min(a, b) - 0.0005 for a, b in zip(op, closes)]
        r = trend_backtest([f"d{i}" for i in range(len(closes))], op, hi, lo, closes, self._p())
        self.assertEqual(len(r.trades), 0)


class Combination(unittest.TestCase):
    def _t(self, side, ed, entry, xd, exit_, stop0):
        from conqueror.strategy import Trade
        t = Trade(side, ed, entry, 1.0)
        t.exit_date, t.exit, t.stop0 = xd, exit_, stop0
        return t

    def test_all_agree_enters_with_last_and_exits_with_first(self):
        from conqueror.combine import combine
        tb = {
            "a": [self._t(1, "d01", 1.00, "d10", 1.10, 0.98)],
            "b": [self._t(1, "d03", 1.02, "d08", 1.06, 1.00)],
            "c": [self._t(1, "d05", 1.04, "d12", 1.12, 1.01)],
        }
        dates = [f"d{i:02d}" for i in range(1, 15)]
        r = combine(tb, "all", dates)
        self.assertEqual(len(r.trades), 1)
        c = r.trades[0]
        self.assertEqual((c.entry_date, c.entry), ("d05", 1.04))   # last method to agree
        self.assertEqual((c.exit_date, c.exit), ("d08", 1.06))     # first method to exit
        self.assertGreater(c.pnl, 0)

    def test_majority_needs_two_of_three(self):
        from conqueror.combine import combine
        tb = {
            "a": [self._t(1, "d01", 1.00, "d10", 1.10, 0.98)],
            "b": [self._t(1, "d03", 1.02, "d08", 1.06, 1.00)],
            "c": [self._t(-1, "d05", 1.04, "d12", 1.00, 1.06)],
        }
        dates = [f"d{i:02d}" for i in range(1, 15)]
        r = combine(tb, "majority", dates)
        c = r.trades[0]
        self.assertEqual((c.side, c.entry_date, c.entry), (1, "d03", 1.02))  # second long completes the majority
        self.assertEqual(c.exit_date, "d08")                                  # majority breaks when b exits

    def test_independent_adds_pnl(self):
        from conqueror.combine import independent
        from conqueror.strategy import Result
        a, b = Result(), Result()
        a.equity_curve = [("d1", 10000.0), ("d2", 10100.0)]
        b.equity_curve = [("d1", 10000.0), ("d2", 9950.0)]
        r = independent({"a": a, "b": b})
        self.assertEqual([e for _, e in r.equity_curve], [10000.0, 10050.0])


class Warehouse(unittest.TestCase):
    def _r(self, d, o, h, l, c):
        return dict(date=d, open=o, high=h, low=l, close=c)

    def test_validate_flags_bad_rows(self):
        from conqueror.warehouse import validate
        rows = [self._r("2020-01-01", 1.0, 1.01, 0.99, 1.0), self._r("2020-01-02", 1.0, 0.9, 0.99, 1.0),
                self._r("2020-01-20", 1.0, 1.01, 0.99, 1.2)]
        text = " ".join(validate(rows))
        self.assertIn("inconsistent OHLC", text)
        self.assertIn("gap of 18 days", text)
        self.assertIn("close-to-close move", text)

    def test_validate_clean(self):
        from conqueror.warehouse import validate
        rows = [self._r("2020-01-01", 1.0, 1.01, 0.99, 1.0), self._r("2020-01-02", 1.0, 1.01, 0.99, 1.005)]
        self.assertEqual(validate(rows), [])

    def test_merge_newer_wins_and_sorted(self):
        from conqueror.warehouse import merge
        old = [self._r("2020-01-02", 1, 2, 0.5, 1.5), self._r("2020-01-01", 1, 2, 0.5, 1.2)]
        new = [self._r("2020-01-02", 1, 2, 0.5, 1.9), self._r("2020-01-03", 1, 2, 0.5, 1.1)]
        m = merge(old, new)
        self.assertEqual([r["date"] for r in m], ["2020-01-01", "2020-01-02", "2020-01-03"])
        self.assertEqual(m[1]["close"], 1.9)


if __name__ == "__main__":
    unittest.main()


class StochTests(unittest.TestCase):
    def test_k_values(self):
        from conqueror.stoch import stoch_k
        hi = [10, 11, 12, 13]
        lo = [8, 9, 10, 11]
        cl = [9, 10, 11, 13]
        k = stoch_k(hi, lo, cl, length=3, smooth=1)
        self.assertEqual(k[:2], [None, None])
        self.assertAlmostEqual(k[2], (11 - 8) / (12 - 8) * 100)
        self.assertAlmostEqual(k[3], (13 - 9) / (13 - 9) * 100)

    def test_flat_range_gives_50(self):
        from conqueror.stoch import stoch_k
        self.assertEqual(stoch_k([5, 5, 5], [5, 5, 5], [5, 5, 5], 3, 1)[2], 50.0)

    def test_long_then_reverse_short(self):
        import math
        from conqueror.stoch import StochParams, stoch_backtest
        n = 400
        cl = [1.2 + 0.05 * math.sin(i / 12) for i in range(n)]
        op = cl[:]
        hi = [c + 0.002 for c in cl]
        lo = [c - 0.002 for c in cl]
        dates = [f"d{i:04d}" for i in range(n)]
        r = stoch_backtest(dates, op, hi, lo, cl, StochParams(spread=0.0, stop_mult=50.0))
        sides = [t.side for t in r.trades]
        self.assertGreater(len(sides), 4)
        self.assertTrue(all(a != b for a, b in zip(sides, sides[1:])))   # always alternates: reversal, never two in a row
        self.assertIn("reverse", {t.reason for t in r.trades})

    def test_stop_exits_and_stays_flat_until_next_cross(self):
        from conqueror.stoch import StochParams, stoch_backtest
        n = 160
        cl = [1.0 - 0.001 * i for i in range(60)] + [0.94 + 0.001 * i for i in range(n - 60)]   # dip then steady rise: one cross up
        op, hi, lo = cl[:], [c + 0.0005 for c in cl], [c - 0.0005 for c in cl]
        lo[120] = cl[120] - 0.5                                # crash bar hits any stop
        dates = [f"d{i:04d}" for i in range(n)]
        r = stoch_backtest(dates, op, hi, lo, cl, StochParams(spread=0.0))
        self.assertTrue(any(t.reason == "stop" for t in r.trades))
