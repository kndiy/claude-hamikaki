"""Append-only log of every backtest run (warehouse/results/runs.csv).

Why: testing many indicators on the same history finds "winners" by chance. The log counts the trials so far and records the
exact parameters and code version of each, so a result can be judged against how many things were tried.
"""
import csv
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path

from .warehouse import ROOT, WH

RUNS = WH / "results" / "runs.csv"
FIELDS = ["run_id", "utc", "method", "variant", "pair", "window", "data_end", "params", "trades", "win_rate", "pf",
          "total_return", "max_drawdown", "code_rev", "note"]


def code_rev() -> str:
    h = hashlib.sha256()
    for p in sorted((ROOT / "conqueror").glob("*.py")):
        h.update(p.read_bytes())
    return h.hexdigest()[:10]


def count() -> int:
    if not RUNS.exists():
        return 0
    with RUNS.open() as f:
        return sum(1 for _ in csv.DictReader(f))


def log_run(method, variant, pair, window, data_end, params, s: dict, note="") -> None:
    RUNS.parent.mkdir(parents=True, exist_ok=True)
    new = not RUNS.exists()
    with RUNS.open("a", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        if new:
            w.writeheader()
        now = datetime.now(timezone.utc)
        w.writerow(dict(
            run_id=f"{now:%Y%m%d%H%M%S}-{count() + 1}", utc=now.strftime("%Y-%m-%dT%H:%M:%SZ"), method=method, variant=variant,
            pair=pair, window=window, data_end=data_end, params=json.dumps(params, sort_keys=True, default=str),
            trades=s["trades"], win_rate=f"{s['win_rate']:.4f}", pf=f"{s['profit_factor']:.4f}",
            total_return=f"{s['total_return']:.4f}", max_drawdown=f"{s['max_drawdown']:.4f}", code_rev=code_rev(), note=note))
