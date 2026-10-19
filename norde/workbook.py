"""Workbook — NORDE-forecast[-<scope>].xlsx, scope = TOTAL | division | entity.

Summary : the Horizon P&L matrix of the scope — growth metrics, Jan t .. Dec t+1, FY columns.
Per line: Calculations block (the line's metrics, all entities — the growth actually applied), then months down and
          column blocks across — TOTAL | TOTAL division | entity … — each Amount | LY | YoY % | F-YoY %.
Blue cells = forecast months. Totals are formulas; YoY % on actual rows, F-YoY % on forecast rows, both with red/green rules.
"""
import re
from collections import defaultdict
from io import BytesIO

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

from actuals import DIVISIONS, ENTITIES
from forecast import current_month, forecast, labels, line_series, metrics, months, shift, trend

BLUE = PatternFill("solid", fgColor="DCE6F1")
GREY = PatternFill("solid", fgColor="F2F2F2")
RED, RED_FONT = PatternFill("solid", fgColor="FFC7CE"), Font(color="9C0006")
GREEN, GREEN_FONT = PatternFill("solid", fgColor="C6EFCE"), Font(color="006100")
THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP, BOTTOM = Border(top=THIN), Border(bottom=THIN)
BOLD = Font(bold=True)
INT, PCT = "#,##0", "0.0%"
TOTALS = [("1.99999 KOPĀ IEŅĒMUMI", "1."), ("2.99999 KOPĀ IZDEVUMI", "2."), ("3.99999 PEĻŅA", None)]
CALC_ROWS, LABEL_ROW = 10, 14  # Calculations block rows 3..12; table label row 14, header 15, months from 16


def put(ws, r, c, v, fmt=None, bold=False, fill=None, border=None):
    cell = ws.cell(row=r, column=c, value=v)
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = BOLD
    if fill:
        cell.fill = fill
    if border:
        cell.border = border
    return cell


def plus(cols, r):
    return "=" + "+".join(f"{col(c)}{r}" for c in cols)


def ratio(c, r):
    """cur/prev-1 of column c over c+1, 0 when prev is 0."""
    return f"=IF({col(c + 1)}{r}=0,0,{col(c)}{r}/{col(c + 1)}{r}-1)"


def rule(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["-0.1"], fill=RED, font=RED_FONT))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.1"], fill=GREEN, font=GREEN_FONT))


def build(actuals, subs, scope="TOTAL"):
    current = current_month(actuals)
    fc = forecast(actuals, subs)
    value = lambda e, l, m: actuals.get((e, l, m), 0) if m < current else fc.get((e, l, m), 0)
    carried = defaultdict(set)
    for e, l, _ in list(actuals) + list(subs):
        if scope == "TOTAL" or scope == e or ENTITIES[e][1] == scope:
            carried[l].add(e)
    calc = {l: metrics(s, current) for l, s in line_series(actuals).items()}

    wb = Workbook()
    summary(wb.active, scope, sorted(carried), carried, calc, value, current)
    for line in sorted(carried):
        sheet(wb.create_sheet(re.sub(r"[:\\/?*\[\]]", " ", line)[:31]), scope, line, sorted(carried[line]), calc[line], value, current)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def summary(ws, scope, lines, carried, calc, value, current):
    """A line | B FY t-2 | C FY t-1 | D FY t-1 % | E YTD t | F YTD t-1 | G YTD % | H F-YoY % | I..T Jan..Dec t | U FY t | V FY t % |
    W..AH Jan..Dec t+1 | AI FY t+1 | AJ FY t+1 %"""
    ws.title = "Summary"
    t, (ytd, _) = int(current[:4]), trend(current)
    keys = labels(current)
    heads = keys[:6] + [keys[-1]] + months(t) + [f"FY {t}", f"FY {t} %"] + months(t + 1) + [f"FY {t + 1}", f"FY {t + 1} %"]
    put(ws, 1, 1, f"NORDE forecast · {scope} · actuals through {current} (MTD) · forecast from {current}", bold=True)
    put(ws, 2, 1, "P&L line", bold=True, border=BOTTOM)
    for c, h in enumerate(heads, 2):
        put(ws, 2, c, h, bold=True, border=BOTTOM)
    month_cols = list(zip(months(t), range(9, 21))) + list(zip(months(t + 1), range(23, 35)))
    first = 3
    series = {}
    for r, line in enumerate(lines, first):
        series[line] = {m: sum(value(e, line, m) for e in carried[line]) for y in (t - 2, t - 1, t, t + 1) for m in months(y)}
        put(ws, r, 1, line)
        for m, c in month_cols:
            fc = m >= current
            put(ws, r, c, series[line][m], INT, fill=BLUE if fc else None, border=BOX if fc else None)
    last = first + len(lines) - 1
    rows = {}
    for k, (name, prefix) in enumerate(TOTALS):
        r = last + 1 + k
        rows[name] = [i for i, l in enumerate(lines, first) if l.startswith(prefix)] if prefix else [last + 1, last + 2]
        series[name] = {m: sum(series[lines[i - first]][m] for i in rows[name]) for m in series[lines[0]]} if prefix else \
            {m: series[TOTALS[0][0]][m] + series[TOTALS[1][0]][m] for m in series[lines[0]]}
        put(ws, r, 1, name, bold=True, border=TOP)
        for c in [c for _, c in month_cols]:
            put(ws, r, c, "=" + "+".join(f"{col(c)}{i}" for i in rows[name]), INT, bold=True, border=TOP)
    for r, name in enumerate(lines + [n for n, _ in TOTALS], first):
        m = metrics(series[name], current)
        bold = name not in lines
        put(ws, r, 2, m[keys[0]], INT, bold=bold)
        put(ws, r, 3, m[keys[1]], INT, bold=bold)
        put(ws, r, 4, f"=IF(B{r}=0,0,C{r}/B{r}-1)", PCT, bold=bold)
        put(ws, r, 5, f"=SUM(I{r}:{col(8 + len(ytd))}{r})" if ytd else 0, INT, bold=bold)
        put(ws, r, 6, m[keys[4]], INT, bold=bold)
        put(ws, r, 7, f"=IF(F{r}=0,0,E{r}/F{r}-1)", PCT, bold=bold)
        put(ws, r, 8, m["F-YoY %"], PCT, bold=bold)
        put(ws, r, 21, f"=SUM(I{r}:T{r})", INT, bold=True)
        put(ws, r, 22, f"=IF(C{r}=0,0,U{r}/C{r}-1)", PCT, bold=bold)
        put(ws, r, 35, f"=SUM(W{r}:AH{r})", INT, bold=True)
        put(ws, r, 36, f"=IF(U{r}=0,0,AI{r}/U{r}-1)", PCT, bold=bold)
    for c in ("D", "G", "H", "V", "AJ"):
        rule(ws, f"{c}{first}:{c}{last + 3}")
    ws.column_dimensions["A"].width = 48
    for c in range(2, 37):
        ws.column_dimensions[col(c)].width = 12
    ws.freeze_panes = "B3"
    ws.sheet_view.zoomScale = 70


def layout(scope, entities):
    """Column blocks [(label, col, kind)] and {division col: [entity cols]}, 4 columns each.
    TOTAL scope: TOTAL | divisions with their entities; division scope: that division with its entities; entity scope: the entity."""
    if scope in ENTITIES:
        return [(f"{scope} {ENTITIES[scope][0]}", 2, scope)], {}
    blocks, groups, c = ([("TOTAL", 2, "company")], {}, 7) if scope == "TOTAL" else ([], {}, 2)
    for d in DIVISIONS:
        ents = [e for e in entities if ENTITIES[e][1] == d]
        if not ents:
            continue
        blocks.append((f"TOTAL {d}", c, "division"))
        groups[c] = [c + 4 * (i + 1) for i in range(len(ents))]
        blocks += [(f"{e} {ENTITIES[e][0]}", ec, e) for e, ec in zip(ents, groups[c])]
        c = groups[c][-1] + 5
    return blocks, groups


def sheet(ws, scope, line, entities, calc, value, current):
    """Rows 1-12 Calculations; row 14 block labels, 15 header, 16..39 = Jan t .. Dec t+1, 40 TOTAL t, 41 TOTAL t+1."""
    t = int(current[:4])
    span = months(t) + months(t + 1)
    first, last = LABEL_ROW + 2, LABEL_ROW + 1 + len(span)
    blocks, groups = layout(scope, entities)

    put(ws, 1, 1, line, bold=True)
    put(ws, 2, 1, "Calculations (line total, all entities)", bold=True, border=BOTTOM)
    put(ws, 2, 2, "Value", bold=True, border=BOTTOM)
    for r, (k, v) in enumerate(calc.items(), 3):
        put(ws, r, 1, k, bold=k == "F-YoY %")
        put(ws, r, 2, v, PCT if k.endswith("%") else INT, bold=k == "F-YoY %")
    put(ws, LABEL_ROW + 1, 1, "Month", bold=True, border=BOTTOM)
    for r, m in enumerate(span, first):
        put(ws, r, 1, m, bold=True)
    put(ws, last + 1, 1, f"TOTAL {t}", bold=True, border=TOP)
    put(ws, last + 2, 1, f"TOTAL {t + 1}", bold=True, border=TOP)
    for g in [c - 1 for c in groups if c > 2]:
        for r in range(LABEL_ROW, last + 3):
            put(ws, r, g, None, fill=GREY)
        ws.column_dimensions[col(g)].width = 3

    for label, c, kind in blocks:
        ws.merge_cells(start_row=LABEL_ROW, start_column=c, end_row=LABEL_ROW, end_column=c + 3)
        put(ws, LABEL_ROW, c, label, bold=True).alignment = Alignment(horizontal="center")
        for i, h in enumerate(("Amount", "LY", "YoY %", "F-YoY %")):
            put(ws, LABEL_ROW + 1, c + i, h, bold=True, border=BOTTOM)
            ws.column_dimensions[col(c + i)].width = 12
        for r, m in enumerate(span, first):
            if kind == "company":
                a, ly = plus(list(groups), r), plus([d + 1 for d in groups], r)
            elif kind == "division":
                a, ly = plus(groups[c], r), plus([e + 1 for e in groups[c]], r)
            else:
                a, ly = value(kind, line, m), value(kind, line, shift(m, -12))
            fc = m >= current
            put(ws, r, c, a, INT, fill=BLUE if fc else None, border=BOX if fc else None)
            put(ws, r, c + 1, ly, INT)
            put(ws, r, c + 3 if fc else c + 2, ratio(c, r), PCT)
        for r, (lo, hi) in ((last + 1, (first, first + 11)), (last + 2, (first + 12, last))):
            put(ws, r, c, f"=SUM({col(c)}{lo}:{col(c)}{hi})", INT, bold=True, border=TOP)
            put(ws, r, c + 1, f"=SUM({col(c + 1)}{lo}:{col(c + 1)}{hi})", INT, bold=True, border=TOP)
            put(ws, r, c + 2, ratio(c, r), PCT, bold=True, border=TOP)
            put(ws, r, c + 3, None, border=TOP)
        rule(ws, f"{col(c + 2)}{first}:{col(c + 3)}{last + 2}")
    for ents in groups.values():
        for e in ents:
            for i in range(4):
                ws.column_dimensions[col(e + i)].outlineLevel = 1
                ws.column_dimensions[col(e + i)].hidden = True
    ws.column_dimensions["A"].width = 18
    ws.freeze_panes = f"B{first}"
    ws.sheet_view.zoomScale = 70
