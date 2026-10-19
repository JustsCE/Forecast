"""Forecast — per (entity, line, month): the submitted amount, else the YoY baseline.

current  = max month in actuals (partial month) — the first forecast month
horizon  = current .. Dec (t+1), t = year of current
trend    = months up to current-2: current-1 is posted but may still be filling (salaries, utilities post with a lag)
yoy      = per line (all entities summed): 0.4·YTD + 0.3·M + 0.2·M-1 + 0.1·M-2 over the trend months,
           each term cur/prev-1 (0 when prev = 0); line level because entity splits get reclassed year to year
baseline = (1 + yoy) × entity LY(month); LY of a t+1 month is the t value (actual or forecast)
"""
from collections import defaultdict

from actuals import DIVISIONS, ENTITIES

WEIGHTS = (0.4, 0.3, 0.2, 0.1)


def shift(month, k):
    y, m = divmod(int(month[:4]) * 12 + int(month[5:]) - 1 + k, 12)
    return f"{y}-{m + 1:02d}"


def months(year):
    return [f"{year}-{m:02d}" for m in range(1, 13)]


def growth(cur, prev):
    return cur / prev - 1 if prev else 0.0


def current_month(actuals):
    return max(m for _, _, m in actuals)


def horizon(current):
    t = int(current[:4])
    return [m for m in months(t) + months(t + 1) if m >= current]


def yoy(series, current):
    """Blended YoY of one series {month: amount} over the trend months (up to current-2)."""
    end = shift(current, -1)
    ytd = [m for m in months(int(current[:4])) if m < end]
    terms = [growth(sum(series.get(m, 0) for m in ytd), sum(series.get(shift(m, -12), 0) for m in ytd))]
    terms += [growth(series.get(shift(end, -k), 0), series.get(shift(end, -k - 12), 0)) for k in (1, 2, 3)]
    return sum(w * g for w, g in zip(WEIGHTS, terms))


def forecast(actuals, subs):
    """{(entity, line, month): amount} for every horizon month of every (entity, line) in actuals or subs."""
    current = current_month(actuals)
    series, totals = defaultdict(dict), defaultdict(lambda: defaultdict(float))
    for (e, l, m), a in actuals.items():
        series[(e, l)][m] = a
        totals[l][m] += a
    for e, l, _ in subs:
        series.setdefault((e, l), {})
    g = {l: yoy(s, current) for l, s in totals.items()}
    out = {}
    for (e, l), s in series.items():
        for m in horizon(current):
            s[m] = subs.get((e, l, m), (1 + g.get(l, 0)) * s.get(shift(m, -12), 0))
            out[(e, l, m)] = s[m]
    return out


def members(target, line, fc):
    """Entity codes a submission targets: one entity, a division's entities, or all (TOTAL)."""
    if target in ENTITIES:
        return [target]
    carried = sorted({e for e, l, _ in fc if l == line})
    if target == "TOTAL":
        return carried
    if target in DIVISIONS:
        return [e for e in carried if ENTITIES[e][1] == target]
    raise ValueError(f"unknown entity/division {target!r}")


def submit(rows, actuals, subs):
    """Chat rows [{entity|division|TOTAL, line, month, amount|yoy}] -> entity-grain [(entity, line, month, amount)].

    A division/TOTAL amount is split over its entities in proportion to their current forecast.
    """
    fc = forecast(actuals, subs)
    current = current_month(actuals)
    lines = {l.split(" ")[0]: l for _, l, _ in actuals}
    out = []
    for r in rows:
        line = lines.get(str(r["line"]).split(" ")[0])
        if line is None:
            raise ValueError(f"unknown line {r['line']!r}")
        if r["month"] not in horizon(current):
            raise ValueError(f"month {r['month']} outside horizon {current}..{horizon(current)[-1]}")
        ents = members(r["entity"], line, fc)
        if not ents:
            raise ValueError(f"no entity carries {line!r} in {r['entity']!r}")
        cur = {e: fc.get((e, line, r["month"]), 0) for e in ents}
        ly = sum(actuals.get((e, line, shift(r["month"], -12)), fc.get((e, line, shift(r["month"], -12)), 0)) for e in ents)
        amount = float(r["amount"]) if "amount" in r else (1 + float(r["yoy"])) * ly
        total = sum(cur.values())
        shares = {e: v / total if total else 1 / len(ents) for e, v in cur.items()}
        out += [(e, line, r["month"], amount * shares[e]) for e in ents if shares[e]]
    return out
