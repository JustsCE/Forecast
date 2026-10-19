"""NORDE forecast CLI — JSON in (stdin), JSON out (stdout), files in NORDE_DIR (default ./forecast).

python run.py dax [since]            the DAX to run with powerbi_run_dax (default since = Jan t-1)
python run.py actuals < result.json  refresh actuals.csv from the powerbi_run_dax result (months >= earliest pulled replaced)
python run.py submit  < rows.json    {"note": "...", "rows": [{entity|division|TOTAL, line, month, amount|yoy}]} -> forecast.csv, then build
python run.py build                  write NORDE-forecast.xlsx
"""
import csv
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import actuals as A
import forecast as F
import workbook as W

DIR = Path(os.environ.get("NORDE_DIR", "forecast"))
ACTUALS, FORECAST, XLSX = "actuals.csv", "forecast.csv", "NORDE-forecast.xlsx"


def rows(name):
    p = DIR / name
    return list(csv.DictReader(p.open(encoding="utf-8"))) if p.exists() else []


def table(name, data, fields):
    DIR.mkdir(parents=True, exist_ok=True)
    with (DIR / name).open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fields, lineterminator="\n")
        w.writeheader()
        w.writerows(data)


def load():
    """actuals {(entity, line, month): amount}, subs {(entity, line, month): amount} (latest submission wins)."""
    actuals = {(r["entity"], r["line"], r["month"]): float(r["amount"]) for r in rows(ACTUALS)}
    subs = {(r["entity"], r["line"], r["month"]): float(r["amount"]) for r in rows(FORECAST)}
    return actuals, subs


def build():
    actuals, subs = load()
    (DIR / XLSX).write_bytes(W.build(actuals, subs))
    current = F.current_month(actuals)
    return {"file": str(DIR / XLSX), "current": current, "horizon": [current, F.horizon(current)[-1]], "submissions": len(subs)}


def main(cmd, data=None):
    if cmd == "dax":
        return {"dax": A.dax(data or f"{datetime.now(timezone.utc).year - 1}-01")}
    if cmd == "actuals":
        out = A.refresh(rows(ACTUALS), A.parse(data))
        table(ACTUALS, out, ["line", "entity", "month", "amount"])
        return {"rows": len(out), "months": [min(r["month"] for r in out), max(r["month"] for r in out)]}
    if cmd == "submit":
        actuals, subs = load()
        try:
            out = F.submit(data["rows"], actuals, subs)
        except ValueError as e:
            return {"error": str(e)}
        ts, note = datetime.now(timezone.utc).isoformat(timespec="seconds"), data.get("note", "")
        log = rows(FORECAST) + [dict(entity=e, line=l, month=m, amount=round(a, 2), ts=ts, note=note) for e, l, m, a in out]
        table(FORECAST, log, ["entity", "line", "month", "amount", "ts", "note"])
        return {"submitted": log[-len(out):], "build": build()}
    if cmd == "build":
        return build()
    return {"error": f"unknown command {cmd!r}"}


if __name__ == "__main__":
    cmd = sys.argv[1]
    data = json.load(sys.stdin) if cmd in ("actuals", "submit") else (sys.argv[2] if len(sys.argv) > 2 else None)
    print(json.dumps(main(cmd, data), ensure_ascii=False, indent=1))
