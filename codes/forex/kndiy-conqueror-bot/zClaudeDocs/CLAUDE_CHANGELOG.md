# CLAUDE.md change log archive

Older change-log entries moved out of `../CLAUDE.md` on 2026-11-26, verbatim and in their original order. The newest entries stay in `CLAUDE.md`; when an entry leaves it, append it here.

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
