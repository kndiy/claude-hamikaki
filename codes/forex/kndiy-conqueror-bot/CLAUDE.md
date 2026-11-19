# kndiy-conqueror-bot

Created 2026-11-16 by the launcher. Fill in each section during the first session.

## Purpose and goals
Personal MetaTrader 5 trading bot (Expert Advisor) for the owner's own use. Demo account only for now.

## Tech stack and key dependencies
- MQL5 Expert Advisor (`.mq5`), compiled with MetaEditor (Windows only) and tested in the MT5 Strategy Tester.
- Files are authored in WSL; compile and backtest happen on the Windows MT5 terminal (Claude can't do this from WSL).

## Structure and key files
- `conqueror/strategy.py`: indicators, signal alignment, bar-by-bar backtest, stats. Standard library only. Assumptions A1-A6 are in its docstring.
- `conqueror/fetch_data.py`: downloads daily OHLC from Yahoo's unofficial chart endpoint to `data/<SYMBOL>.csv`.
- `conqueror/run_backtest.py`: CLI; prints results at 0, 1 and 3 pip spread.
- `tests/test_strategy.py`: unit tests (7).
- `conqueror/warehouse.py`: CSV data warehouse (`warehouse/ohlc/1d/<SYMBOL>.csv`, `warehouse/catalog.csv` with source, range, sha256, validation issues). One source per symbol; `fetch_data` writes through it. `data/` is the old copy (legacy, can be deleted after the owner confirms).
- `conqueror/registry.py`: append-only `warehouse/results/runs.csv` logging every run of the single-method runners (params, code hash, data end). Counts trials to judge results against how many things were tried. Not yet wired into run_periods/run_combo/run_years.
- `conqueror/accept.py`: the agreed accept rule (pooled PF > 1.2, positive in both halves, PF > 1 on >= 3 of 4 pairs) applied to every registered method; `conqueror/stoch.py`: ch. 5 stochastic 50-line system.
- `notes/indicators.md`: the book's indicators and rules compiled (ch. 2, 3, 4 done; 5-10 pending), with BOOK vs GAP marked. Source of truth for what the book says.
- Later, optional: the `.mq5` EA for the MT5 terminal (path TBD).

## Commands (build / run / test / lint)
- Data: `python3 -m conqueror.fetch_data EURUSD`
- Backtest: `python3 -m conqueror.run_backtest EURUSD [--spread-pips 1.0] [--risk 0.01] [--from 2003-01-01] [--tighten cumulative|current|none] [--trades]`
- Tests: `python3 -m unittest discover -s tests -t .`
- Run from the project root. No lint tool configured. MQL5 compile/backtest (if the EA is built) is Windows only.

## Conventions, constraints, decisions
- 2026-11-16: Chose a native MQL5 EA over Python `MetaTrader5` or a hybrid (Strategy Tester, no external process).
- Demo account only. Before any live use, add max daily loss, a kill switch, and spread and slippage guards.
- Market: a single forex pair (exact symbol and timeframe TBD).
- Planned safety guards from the start: spread check, one position per magic number, handling of rejected orders.
- Strategy: the Conqueror, from "How to Make a Living Trading Foreign Exchange" (Courtney Smith), Chapter 4, pp. 55-65. Book PDF: `D:\Documents\Learning Forex\` (WSL: `/mnt/d/Documents/Learning Forex/`). Daily bars.
  - Entry: long if all 3 are positive, short if all 3 are negative, else flat. (1) close minus 10-day MA of close; (2) 10-day MA minus the 10-day MA 10 days ago; (3) close minus close 40 days ago. Enter at the signal close or the next open.
  - Stop (Conqueror I): 2 x ATR(40) from the lowest close (shorts) / highest close (longs) since entry; moves only in the trade's favour; ATR re-read daily. ATR uses true range (true high = max(high, prev close), true low = min(low, prev close)).
  - Tightening: each time a condition changes sign, the ATR multiple shrinks by a third of the original (480 -> 322 in the book's example). This is the owner's/our reading of an ambiguous example; confirm.
  - Disaster stop: the book uses $2,000 on a standard contract; to be replaced by a percent-of-equity risk.
  - Re-entry after a stop-out: wait until a condition flips against the trade and flips back.
  - Conqueror II (12 x ATR) is out of scope.
- The book's table 4.1 is Freeburg's Conquistador (1972-98 futures), not the Conqueror; the chapter has no cost (spread, slippage, financing) model. Treat the rules as a hypothesis to backtest.
- Owner's context: previous losses came from revenge trading, not the strategy (self-reported). The bot must enforce: one position, fixed percent risk, no size increase after a loss, daily/weekly loss limit, re-entry rule in code.
- Owner lives in Vietnam, where retail forex is not permitted (from memory, unverified). IG access for Vietnam residents is unconfirmed; IG does offer MT5 (verified by search 2026-11-16, region-dependent, fewer markets than its main platform). No account yet; building and backtesting need none. Do not help misstate residency or address to a broker.
- Owner's trading history (his account, 2026-11-16): the 100 -> 500 USD gain came from trend trading around 2018 (he withdrew, restarted at 100, lost it); he stopped, resumed in 2023 using the Conqueror and lost the 100 USD invested. Earlier he attributed losses to revenge trading; both explanations are possible. Owner (2026-11-16): the 2018 gain was entering blind with no risk control and was luck. Anecdotal (one small account); the backtests, not this, decide.
- The owner built a Conqueror indicator before that "seems wrong somewhere"; code not yet seen.

## Backtest findings (2026-11-16)
- Yahoo `range=max` silently returns MONTHLY bars; the fetcher now uses explicit period1/period2 and raises if the granularity is not 1d.
- EURUSD daily 2003-12..2026-11 (5,827 bars), 1% risk, 10,000 start, entry at next open:
  - tighten=cumulative: 463 trades, win 36%, PF 0.88 at 0 spread, 0.86 at 1 pip, -18% total.
  - tighten=current: 370 trades, win 32%, PF 0.91 / 0.90.
  - tighten=none: 281 trades, win 32%, PF 0.87 / 0.86.
  - All three readings lose on EURUSD before costs. Book's own (futures, 1972-98) figures: 46% wins, win/loss ratio 3.3.
- WITHDRAWN (owner decided 2026-11-16: follow the book all the way; the book's cumulative tightening is the default; `exit2*` options stay in the code for reference only). Owner variant (from memory, he traded the 3-condition table with manual cuts when 2 conditions turned): `exit2` (exit at next open when >=2 conditions are against; 2xATR trail kept) 336 trades, PF 0.89 / 0.88 at 0 / 1 pip; `exit2_events` (2 sign changes since entry) 475 trades, PF 0.89 / 0.87. Same range as the book's rule; neither shows an edge on EURUSD daily.
- Chapter 7 sizing applied: default risk now 0.5% per trade (book's starting figure), was 1%. EURUSD, 1 pip spread: Conqueror alone PF 0.87, -9%, maxDD 12%. With the Bishop (owner decision, ADX(14) > 40 and falling -> exit next open): PF 0.91, -6%, maxDD 10%. ADX period: book gives none; fixed at 14 (Wilder standard and the book's default length for oscillators, ch. 5 p. 69); sensitivity 10 -> PF 0.96, 20 -> PF 0.89 recorded only, not used to choose. Bishop helps a little; still below 1.
- Not conclusive: one pair, Yahoo indicative data, our reading of ambiguous rules. 20 trades a year is more than the chapter's "rarely trades" suggests, so a mismatch with the book's intent is possible.

## Channel breakout findings (2026-11-16)
`conqueror/channel.py`, `conqueror/run_channel.py` (55/20, buffer 3 pips, add-ons switchable; assumptions C1-C5 in the module docstring). Daily, 1 pip spread, 0.5% risk, Dec 2003/2006 - Nov 2026, profit factor (PF):
- EURUSD: plain 1.15, +ADX filter 1.11, +Bishop 0.88, +Rejection 0.86, +Last bar 1.22, all four (book) 0.73.
- GBPUSD: plain 0.89, +Last bar 0.82, all four 0.66. AUDUSD: plain 0.72, +Last bar 0.67, all four 0.74. NZDUSD: plain 0.66, +Last bar 0.74, all four 0.69.
- The book-faithful "all four" loses on all four pairs (PF 0.66-0.74, win rate 12-14%, many small losers). Not tuned; owner decision 2026-11-16: always follow the book for now, improve later with more data.
- Worth inspecting before trusting: the Rejection Rule raised trade counts (110 -> 173 on EURUSD) while the book says it removes about half the stop-outs; my reading of re-entry after a rejection exit, the Bishop timing (next open), and the sizing estimate (C3) could differ from the author's intent.

## Trend Analysis findings (2026-11-16)
`conqueror/trend.py`, `conqueror/run_trend.py` (swings rank >= 3, HH/HL logic, stop at the latest swing, neutral exit, ADX filter, Bishop; assumptions T1-T5 in the docstring; 19 tests pass). Daily, 1 pip spread, 0.5% risk, profit factor, plain -> book (ADX filter + Bishop):
- EURUSD 0.84 -> 0.86, GBPUSD 0.96 -> 1.03, AUDUSD 0.66 -> 0.75, NZDUSD 0.82 -> 0.97. Book version roughly break-even to losing; the ADX filter and Bishop improved all four pairs.
- No worked example in ch. 2 is dated precisely enough to check against (Japanese Yen chart); swing logic only unit-tested on synthetic data.

## Window check on the owner's periods (2026-11-16)
`python3 -m conqueror.run_periods` (4 USD-quote pairs, 1 pip, 0.5% risk, full-history warm-up, trades attributed by entry date; window return approximate). Pooled profit factor:
- 2017-2019: Conqueror 0.97, Conqueror+Bishop 1.02, Channel (book) 0.41, Trend (book) 0.93.
- 2023-now: Conqueror 0.74, Conqueror+Bishop 0.76, Channel 0.22, Trend 0.45.
- Full: 0.88 / 0.89 / 0.71 / 0.91.
- The owner's 2018 trend gain is not reproduced (Trend 0.93 pooled; 84 trades is too few to separate 0.93 from 1.1); the 2023 Conqueror loss is consistent (0.74). USD-base pairs (USDJPY, USDCHF, USDCAD) are not supported yet (P&L conversion), so the owner's 2018 pair may be outside the test.

## Chapter 10 combination findings (2026-11-16)
`conqueror/combine.py`, `conqueror/run_combo.py` (A independent sum, B all agree, C majority 2 of 3; sized from the tightest initial stop; assumptions K1-K3 in the docstring; 22 tests pass). Pooled PF over EUR/GBP/AUD/NZD USD, 1 pip, 0.5% risk:
- Full: Conqueror+Bishop 0.89, Channel 0.71, Trend 0.91 | A 0.82, B 0.73, C 0.77.
- 2017-2019: 1.02, 0.41, 0.93 | A 0.73, B 0.82, C 0.62. 2023-now: 0.76, 0.22, 0.45 | A 0.46, B 0.31, C 0.33.
- The combination does not beat the best single method and is worse than Conqueror+Bishop or Trend alone; combining three methods without an edge gives no edge. B/C drawdown is understated (realised equity only).
- Owner: EURUSD is the pair he traded mainly (2018 and 2023). EURUSD per-pair PF 2017-2019: Conqueror+Bishop 0.62, Channel 0.40, Trend 0.70; 2023-now: 0.85, 0.31, 0.54. His 2018 gain is not reproduced by these rules on EURUSD.

## EURUSD per-year breakdown (2026-11-16)
`python3 -m conqueror.run_years EURUSD` (1 pip, 0.5% risk, trades by entry year, return approximate). Years with PF > 1 out of 24: Trend 11, Channel (book) 9, Conqueror+Bishop 8, all-agree 7. Yearly returns swing about +-5%; no method is consistent.
- 2018: Trend PF 3.57 (+0.8%, only 4 trades), Conqueror+Bishop 1.16 (+0.4%), all-agree 9.29 (3 trades): same direction as the owner's 2018 gain, much smaller.
- 2023: Conqueror 0.93 (-0.2%), +Bishop 1.00 (0.0%); 2024: 0.81 / 0.71 (-0.6% / -0.9%). At 0.5% risk the rules lose under 1% a year, so losing a whole 100 USD account cannot come from the signals alone; position size/leverage (and manual cuts) must explain most of it.
- 2017 (EURUSD +13%) all book methods lost. Cause checked: the Rejection Rule and Last Bar stop chopped the entry of the year's big trend (2017-04-24 breakout: plain 55/20 held to 2017-09-25 for +586; book version was rejected and stopped out 7 times, then caught only a late part). The book accepts that ('premature breakouts'); it is behaviour, not a bug. Candidate for the later improvement phase, not applied (owner: follow the book for now).

## Implementation check against the book's worked examples (2026-11-16)
- Conqueror, British Pound Oct 2008 (book pp. 61-65), Yahoo GBPUSD: all-three-negative short signal 2008-10-06 close 1.7455 (book: short near 1.7400); ATR(40) 0.0244 (book: about 0.0240); lowest close 1.5539 on 2008-10-27 (book: 1.5537); next-day stop-out above 1.6037 (book: about 1.6057), about 1,265 pips (book: over 1,200). Conditions, ATR and stop anchor match.
- Channel breakout, same pair: short entry 2008-10-06 at 1.7356 (book: early October near 1.7400); exit 2008-11-20 at 1.4982 by the Bishop (book: exit near 1.4750 via the Rejection Rule near the low). Close, not identical.
- Book's July 2007 Rejection Rule example (55-day high just above 2.0000) could NOT be matched: Yahoo GBPUSD in July 2007 trades at 2.01-2.06 with a 55-day high near 2.02-2.04, so the book's date or data differs. That example is unverified.
- Rejection Rule exits are by construction small losses (77 of 245 on EURUSD, 0 winners); Bishop exits were the winners (27 of 28). No bug found; the book-faithful result stands.

## Data warehouse and testing protocol (2026-11-16)
- Validation of the Yahoo data found: real 2008 crisis moves (>5%/day, flagged for checking), a missing week in EURUSD 2008-07-30..08-07 and 7-day gaps in NZDUSD, 17 zero-range (stale) GBPUSD bars, and a drifting day boundary (bars labelled Sun-Thu part of the year, Mon-Fri otherwise; still one bar per trading day, ~4.8-4.96 bars a week, so lookbacks count trading days correctly). All earlier backtests ran on this data.
- Candidate better source: Dukascopy export (hourly/tick CSV, history from about 2003 per forum and third-party docs; licence/terms NOT verified) to build New York 5pm-close dailies. Needs the owner's go-ahead and a terms check.
- Owner wants to implement all indicators (book's remaining + outside ones) and test on a demo account. Overfitting protocol proposed (not yet agreed): fixed canonical parameters before testing, every run logged in the registry, accept rule written before looking (candidate: pooled PF > 1.2 after 1 pip, positive in both halves of the history and on at least 3 of 4 pairs), and ALL history to 2026-11-16 counts as already seen, so the real out-of-sample is data after that date (forward paper test from the warehouse, then demo).
- Demo account: broker/platform undecided; Vietnam residency question open (see above).
- **AGREED by the owner 2026-11-16 ("ok, all 3"):**
  1. Dukascopy becomes the main data source. The OWNER reads and accepts Dukascopy's terms before any bulk download; code only imports CSV files the owner exports (no scraping/downloader written until the owner confirms the terms).
  2. Testing protocol (binding): (a) parameters fixed to the book / textbook values BEFORE testing, no tuning; (b) every run logged in `warehouse/results/runs.csv`; (c) ACCEPT RULE, written before looking: pooled PF > 1.2 after 1 pip spread, positive net result in BOTH halves of the history, and PF > 1 on at least 3 of the 4 pairs. Failing it = rejected, not re-tuned; (d) all history up to 2026-11-16 counts as already seen; real out-of-sample = data after that date (forward paper test, then demo).
  3. Indicator order: the book's remaining methods first (ch. 5 stochastics, ch. 6 inside days/patterns, ch. 8 Slingshot), then time-series momentum, MA crossover, RSI(2), Bollinger Bands. Intraday-hold methods (inside days, Slingshot) need intraday data, so they wait for the Dukascopy import.
  - Trial count matters: the more variants logged, the higher the bar a "good" result must clear (a good-looking PF among many trials is expected by chance).

## Accept-rule results (2026-11-16, first run of the agreed protocol)
`python3 -m conqueror.accept [method ...]` (`conqueror/accept.py`; new methods are registered in `methods()` with fixed parameters BEFORE running). 1 pip, 0.5% risk, 4 USD-quote pairs, split at the data's midpoint (2015-05-25). Pooled PF / P&L before and after the split:
- Conqueror+Bishop 0.89 (-945 / -1786), Channel (book) 0.71 (+487 / -8016), Trend (book) 0.91 (-169 / -1005), Stochastic 50-cross (book ch. 5, new) 0.94 (+632 / -2804). All REJECTED (fail all three tests). Numbers match the earlier runs, so the gate is wired correctly. 19 runs logged.
- Stochastic system assumptions S1-S3 are in `conqueror/stoch.py`; divergence/failure/trend-crossover uses of the stochastic are discretionary or undefined in the book and not implemented.
- 2015-2026 is worse than 2003-2015 for every method (the "after" half loses for all four): consistent with the owner's recent losses; not tuned for.

## Open questions and next steps
Next (agreed order): ch. 6 inside days/patterns and ch. 8 Slingshot (compile rules first; need intraday data -> Dukascopy CSV export by the owner), then time-series momentum, MA crossover, RSI(2), Bollinger Bands, each registered with fixed parameters in `accept.methods()` before running. Daily forward paper-test log after that.
Decided: Python backtester first (2026-11-16). EURUSD daily, 1% risk as defaults.
Open (waiting on the owner):
1. The old indicator's code (owner can't find it; no Conqueror/.mq5 file found under ~/projects or D:\ for it) and what looked wrong.
2. Confirm the tightening reading and the ATR(40) definition.
3. Whether IG accepts Vietnam residents (ask IG support).

Next: finish `notes/indicators.md` (ch. 5, 6, 8, 9; ch. 7 and 10 done), then decide how to combine in ch. 10 terms (independent methods vs weight of evidence), then re-run backtests per method. Pause further Conqueror-only backtests until then. Old plan, kept for reference: test other USD-quote pairs and sub-periods (report all, do not cherry-pick), re-check the rules against the chapter for a misread, then the live risk layer (daily/weekly loss limit, kill switch) only if a rule set shows an edge.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:linux:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
<!-- launcher:tools:linux:end -->

## Status at pause (2026-11-16)
- Owner's view: "any static method will be obsolete over time". Assistant's position: not proven. The first half (2003-2015) was also about break-even (pooled PF about 0.9), so there is no evidence of an edge that faded; the four methods are all daily trend-followers on the same 4 USD pairs (not independent); alternative explanations: book evidence is futures 1972-98, regime (range-bound EURUSD 2015-2026), swap/interest not modelled.
- Pending decision: (a) rolling 5-year PF per method (smooth decline vs one bad regime; feeds the adaptive/walk-forward question), or (b) continue with time-series momentum, MA crossover, RSI(2), Bollinger Bands under `accept.py`. Adaptive/walk-forward methods would count as trials too.
- Owner asked what is the most popular method now. Answered from memory (unverified): retail = price action/support-resistance, Smart Money Concepts/ICT (mostly subjective, not backtestable), classic indicators; professionals = systematic trend following, carry, momentum. Popularity is not evidence of profit. Optional web search on this offered, not done. Carry needs swap/interest data (not in the warehouse).
- Waiting on the owner: Dukascopy CSV export (after reading their terms) for ch. 6/8 intraday methods; demo-platform choice; Vietnam eligibility; deletion of legacy `data/`.
- Recommendation unchanged: no live trading; forward paper test is the next real evidence.

## Change log
- 2026-11-16: Saved progress and the discussion on static vs adaptive methods (see Status at pause).
- 2026-11-16: Owner agreed all three proposals ("ok, all 3"): Dukascopy as source (owner checks terms), the testing protocol and accept rule, the indicator order. Added `accept.py` (the gate) and `stoch.py` (ch. 5 50-line system), compiled ch. 5 into the notes; all four methods tested so far are rejected. 29 tests pass.
- 2026-11-16: Project created.
- 2026-11-16: Recorded the first decisions (MQL5 EA, demo only, single pair) and the open questions. No code yet.
- 2026-11-16: Owner asked to implement all indicators, build a CSV data warehouse and test on demo. Built the warehouse (catalog, validation, merge/update), migrated the 4 pairs, added the run registry; validation found data gaps/stale bars (see protocol section). 25 tests pass.
- 2026-11-16: Per-year EURUSD breakdown (`run_years.py`); checked why 2017 lost (add-ons chop big trends); 2018 direction matches the owner's gain, 2023 loss is far larger than the rules explain.
- 2026-11-16: Built the Chapter 10 combination (independent / all agree / majority) and ran it: no better than the best single method. Owner trades EURUSD mainly.
- 2026-11-16: Added `run_periods.py` and ran the owner's windows (2017-2019, 2023-now): no method shows an edge; 2023-now is worse for all.
- 2026-11-16: Built Trend Analysis (`trend.py`, `run_trend.py`) and ran four pairs: book version PF 0.75-1.03. Recorded the owner's trading history (trend gain around 2018, Conqueror loss from 2023).
- 2026-11-16: Verified the Conqueror and channel code against the book's Oct 2008 GBP examples (matches within data differences); July 2007 Rejection Rule example not matched. No bug found.
- 2026-11-16: Built the 55/20 channel breakout with ADX filter, Bishop, Rejection Rule, Last Bar stop (15 tests pass); ran EURUSD, GBPUSD, AUDUSD, NZDUSD. Owner: always follow the book for now, improve later with more data. Book-faithful version loses on all four pairs; see Channel breakout findings.
- 2026-11-16: Read ch. 7, rest of ch. 10 and Epilogue into `notes/indicators.md`. Added ADX and the Bishop exit to the backtester (`--bishop`, `--adx-len`; owner decided it applies to the Conqueror), default risk 0.5%. 11 tests pass.
- 2026-11-16: Owner: follow the book all the way; his cut/flip/lookback memory is consistent with the book. Book rules are the spec; variants withdrawn. Added a unit test pinning the tightening arithmetic.
- 2026-11-16: Owner said he traded the table with manual cuts and wants an automatic bot. Added `--tighten exit2|exit2_events` variants and backtested them (no better than the book's rule). Recommended not going live until a rule set shows an edge.
- 2026-11-16: Compiled ch. 2-4 indicators into `notes/indicators.md` (swing/trend analysis, ADX filter, Bishop, channel breakout 55/20 with rejection rule and last bar, Conqueror). Owner asked to slow down: compile first, combine later.
- 2026-11-16: Added the Python backtester (strategy, fetcher, CLI, 7 tests, all pass). First result on EURUSD daily: profit factor below 1 under all three readings of the tightening rule. See Backtest findings.
- 2026-11-16: Added the Conqueror rules from Chapter 4, the risk-layer requirements, the Vietnam/IG constraint, and updated the open questions. No code yet.
