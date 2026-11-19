"""CSV data warehouse for price history.

Layout (all plain CSV, human-readable):
  warehouse/ohlc/1d/<SYMBOL>.csv   date,open,high,low,close   (one file per symbol and timeframe)
  warehouse/catalog.csv            one row per file: source, date range, rows, sha256, last update, validation issues
  warehouse/results/runs.csv       append-only log of backtest runs (see registry.py)
Rules: a symbol's file has ONE source (a different source needs force=True), rows are unique and sorted by date, every
update re-validates and refreshes the catalog. Raw downloads are never edited by hand.
"""
import csv
import hashlib
from datetime import date, datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WH = ROOT / "warehouse"
OHLC_DIR = WH / "ohlc" / "1d"
CATALOG = WH / "catalog.csv"
FIELDS = ["date", "open", "high", "low", "close"]
CAT_FIELDS = ["symbol", "timeframe", "source", "first_date", "last_date", "rows", "sha256", "updated_at_utc", "issues"]


def path(symbol: str) -> Path:
    return OHLC_DIR / f"{symbol}.csv"


def read(symbol: str) -> list:
    p = path(symbol)
    if not p.exists():
        return []
    with p.open() as f:
        return [dict(r) for r in csv.DictReader(f)]


def write(symbol: str, rows: list) -> None:
    OHLC_DIR.mkdir(parents=True, exist_ok=True)
    rows = sorted({r["date"]: r for r in rows}.values(), key=lambda r: r["date"])
    with path(symbol).open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r[k] for k in FIELDS})


def validate(rows: list) -> list:
    """Return a list of human-readable issues (empty list = clean)."""
    issues, prev, zero = [], None, 0
    for r in rows:
        o, h, l, c = (float(r[k]) for k in ("open", "high", "low", "close"))
        if not (l <= min(o, c) and h >= max(o, c)):
            issues.append(f"{r['date']}: inconsistent OHLC")
        if h == l:
            zero += 1
        if prev is not None:
            if r["date"] <= prev["date"]:
                issues.append(f"{r['date']}: date not increasing")
            gap = (date.fromisoformat(r["date"]) - date.fromisoformat(prev["date"])).days
            if gap > 6:
                issues.append(f"{prev['date']}..{r['date']}: gap of {gap} days")
            move = abs(c / float(prev["close"]) - 1)
            if move > 0.05:
                issues.append(f"{r['date']}: {move:.1%} close-to-close move (check)")
        prev = r
    if zero:
        issues.append(f"{zero} zero-range bars (stale quotes?)")
    return issues


def merge(old: list, new: list) -> list:
    """Upsert by date; the newer download wins on overlapping dates (the latest bar is often revised)."""
    d = {r["date"]: r for r in old}
    d.update({r["date"]: r for r in new})
    return sorted(d.values(), key=lambda r: r["date"])


def _sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()[:16]


def _read_catalog() -> dict:
    if not CATALOG.exists():
        return {}
    with CATALOG.open() as f:
        return {(r["symbol"], r["timeframe"]): r for r in csv.DictReader(f)}


def refresh_catalog(symbol: str, source: str, rows: list, issues: list) -> None:
    cat = _read_catalog()
    cat[(symbol, "1d")] = dict(
        symbol=symbol, timeframe="1d", source=source, first_date=rows[0]["date"], last_date=rows[-1]["date"],
        rows=len(rows), sha256=_sha(path(symbol)), updated_at_utc=datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        issues=" | ".join(issues[:5]) + (f" | (+{len(issues) - 5} more)" if len(issues) > 5 else ""))
    WH.mkdir(parents=True, exist_ok=True)
    with CATALOG.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAT_FIELDS)
        w.writeheader()
        for k in sorted(cat):
            w.writerow(cat[k])


def update(symbol: str, fetch_rows, source: str, force: bool = False):
    """fetch_rows(symbol) -> list of dicts with FIELDS. Returns (rows_added, issues)."""
    cat = _read_catalog().get((symbol, "1d"))
    if cat and cat["source"] != source and not force:
        raise RuntimeError(f"{symbol} is stored from '{cat['source']}', refusing to mix with '{source}' (force=True to replace)")
    old = [] if (cat and cat["source"] != source) else read(symbol)
    new = fetch_rows(symbol)
    merged = merge(old, new)
    write(symbol, merged)
    issues = validate(merged)
    refresh_catalog(symbol, source, merged, issues)
    return len(merged) - len(old), issues
