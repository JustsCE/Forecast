"""Sheet builder — emits the ordered Google Sheets/Drive MCP calls that build the `NORDE forecast` workbook.

python sheet.py <spreadsheet_id>            prints build.json: [{"tool": ..., "args": {...}}, ...]
python sheet.py <spreadsheet_id> --step N   prints one call

A Claude Code session relays the calls in order through the Google Sheets connector; nothing runs locally.
Tabs: Summary (0) · one per line (100+) · inputs (4) · pull (5) · actuals (6) · data (7) · grid (8) · model (3) ·
lines (1) · entities (2) · Data Sources (10). The forecast logic is in the formulas below; see TASK.md for the chat procedure.
"""
import json
import sys

from taxonomy import DIVISIONS, ENTITIES, LINES, TOTALS

SUMMARY, LINES_T, ENT, MODEL, INPUTS, PULL, ACTUALS, DATA, GRID, TEMPLATE, SOURCES = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
AMT, PCT, MONTH, TEXT = "#,##0;-#,##0;", "0.0%;-0.0%;", "yyyy-mm", "@"
BLUE, RED, RED_FG, GREEN, GREEN_FG = "DCE6F1", "FFC7CE", "9C0006", "C6EFCE", "006100"
N_LINES = len(LINES)                      # 20 → Summary rows 3..22, totals 23..25
REV = [i for i, l in enumerate(LINES) if l[2] == "revenue"]
ACTIVE = [e for e in ENTITIES if e[3]]
sid = None
calls = []


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def call(tool, **args):
    calls.append({"tool": tool, "args": {"spreadsheetId": sid, **args}})


def values(rng, rows):
    call("update_values", range=rng, values=rows)


def formulas(rng, rows):
    call("update_formulas", range=rng, formulas=rows)


def batch(requests):
    call("update_spreadsheet", requests=requests)


def rgb(hexcolor):
    return {"red": int(hexcolor[0:2], 16) / 255, "green": int(hexcolor[2:4], 16) / 255, "blue": int(hexcolor[4:6], 16) / 255}


def grid_range(sheet, a1):
    """A1 (e.g. 'B3:D10', 'C:C', 'A2:A') -> GridRange."""
    start, _, end = a1.partition(":")
    end = end or start

    def parse(ref):
        letters = "".join(ch for ch in ref if ch.isalpha())
        digits = "".join(ch for ch in ref if ch.isdigit())
        c = 0
        for ch in letters:
            c = c * 26 + ord(ch) - 64
        return c, int(digits) if digits else None

    c1, r1 = parse(start)
    c2, r2 = parse(end)
    out = {"sheetId": sheet}
    if c1:
        out["startColumnIndex"], out["endColumnIndex"] = c1 - 1, c2
    if r1:
        out["startRowIndex"] = r1 - 1
    if r2:
        out["endRowIndex"] = r2
    return out


def fmt(sheet, a1, number=None, bold=None, bg=None, fg=None, align=None, border=None):
    cell, fields = {}, []
    uf = cell.setdefault("userEnteredFormat", {})
    if number:
        uf["numberFormat"] = {"type": "NUMBER" if number not in (TEXT, MONTH) else ("TEXT" if number == TEXT else "DATE"), "pattern": number}
        fields.append("userEnteredFormat.numberFormat")
    if bold is not None:
        uf["textFormat"] = {"bold": bold}
        fields.append("userEnteredFormat.textFormat.bold")
    if bg:
        uf["backgroundColorStyle"] = {"rgbColor": rgb(bg)}
        fields.append("userEnteredFormat.backgroundColorStyle")
    if fg:
        uf.setdefault("textFormat", {})["foregroundColorStyle"] = {"rgbColor": rgb(fg)}
        fields.append("userEnteredFormat.textFormat.foregroundColorStyle")
    if align:
        uf["horizontalAlignment"] = align
        fields.append("userEnteredFormat.horizontalAlignment")
    if border:
        uf["borders"] = {border: {"style": "SOLID", "colorStyle": {"rgbColor": rgb("000000")}}}
        fields.append("userEnteredFormat.borders")
    return {"repeatCell": {"range": grid_range(sheet, a1), "cell": cell, "fields": ",".join(fields)}}


def width(sheet, c1, c2, px):
    return {"updateDimensionProperties": {"range": {"sheetId": sheet, "dimension": "COLUMNS", "startIndex": c1 - 1, "endIndex": c2},
                                          "properties": {"pixelSize": px}, "fields": "pixelSize"}}


def rule(sheet, ranges, formula, bg, fg=None):
    fmt_ = {"backgroundColorStyle": {"rgbColor": rgb(bg)}}
    if fg:
        fmt_["textFormat"] = {"foregroundColorStyle": {"rgbColor": rgb(fg)}}
    return {"addConditionalFormatRule": {"rule": {"ranges": [grid_range(sheet, r) for r in ranges],
                                                  "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]}, "format": fmt_}},
                                         "index": 0}}


def sheet_props(sheet, title=None, frozen_rows=None, frozen_cols=None, hide_gridlines=True):
    props, fields = {"sheetId": sheet, "gridProperties": {"hideGridlines": hide_gridlines}}, ["gridProperties.hideGridlines"]
    if title:
        props["title"], fields = title, fields + ["title"]
    if frozen_rows is not None:
        props["gridProperties"]["frozenRowCount"], fields = frozen_rows, fields + ["gridProperties.frozenRowCount"]
    if frozen_cols is not None:
        props["gridProperties"]["frozenColumnCount"], fields = frozen_cols, fields + ["gridProperties.frozenColumnCount"]
    return {"updateSheetProperties": {"properties": props, "fields": ",".join(fields)}}


# ---------------------------------------------------------------- 1. structure
def structure():
    tabs = [(LINES_T, "lines", 30, 5), (ENT, "entities", 40, 6), (MODEL, "model", 40, 14), (INPUTS, "inputs", 1000, 10),
            (PULL, "pull", 2000, 5), (ACTUALS, "actuals", 6000, 5), (DATA, "data", 8000, 10), (GRID, "grid", 1000, 60),
            (TEMPLATE, "_line", 45, 100), (SOURCES, "Data Sources", 100, 11)]
    reqs = [{"addSheet": {"properties": {"sheetId": s, "title": t, "gridProperties": {"rowCount": r, "columnCount": c}}}} for s, t, r, c in tabs]
    reqs.append(sheet_props(SUMMARY, title="Summary", frozen_rows=2, frozen_cols=1))
    reqs += [sheet_props(s, frozen_rows=1) for s in (LINES_T, ENT, INPUTS, PULL, ACTUALS, DATA, GRID, SOURCES)]
    reqs += [sheet_props(MODEL), sheet_props(TEMPLATE, frozen_rows=15, frozen_cols=1)]
    for s in (LINES_T, ENT):
        reqs.append(fmt(s, "A2:A", number=TEXT))
    for s in (INPUTS, PULL, ACTUALS):                 # inputs from row 1: appended rows inherit the format of the row above
        r = 1 if s == INPUTS else 2
        reqs += [fmt(s, f"A{r}:B", number=TEXT), fmt(s, f"C{r}:C", number=MONTH), fmt(s, f"D{r}:E" if s == INPUTS else f"D{r}:D", number=AMT)]
    reqs += [fmt(DATA, "C2:C", number=MONTH), fmt(DATA, "D2:D", number=AMT), fmt(GRID, "E1:AZ1", number=MONTH), fmt(TEMPLATE, "B14:CV14", number=TEXT)]
    batch(reqs)


# ---------------------------------------------------------------- 2. typed values
def taxonomy_values():
    values("lines!A1:C%d" % (N_LINES + 1), [["code", "line", "kind"]] + [list(l) for l in LINES])
    values("entities!A1:D%d" % (len(ENTITIES) + 1), [["code", "name", "division", "active"]] + [list(e) for e in ENTITIES])
    for t in ("inputs", "pull", "actuals"):
        values(f"{t}!A1:{'G' if t == 'inputs' else 'D'}1",
               [["target", "line", "month", "amount", "yoy", "note", "ts"] if t == "inputs" else ["line", "entity", "month", "amount"]])
    values("data!A1:I1", [["line", "entity", "month", "amount", "pair", "pair|month", "line|month", "line|year", "line|year|ytd"]])
    values("'Data Sources'!A1:K1", [["Written to", "Source", "Object", "Obtained via", "Query or request", "Parameters", "Run at",
                                     "Requested by", "Row count", "Changes after retrieval", "Notes"]])


# ---------------------------------------------------------------- 3. engine formulas
def engine():
    # data = pull months + actuals before the pull
    formulas("data!A2", [["=IF(COUNTA(pull!A2:A)=0, FILTER(actuals!A2:D, actuals!A2:A<>\"\"), "
                          "{FILTER(pull!A2:D, pull!A2:A<>\"\"); FILTER(actuals!A2:D, actuals!A2:A<>\"\", actuals!C2:C<MIN(pull!C2:C))})"]])
    key = lambda expr: f"=ARRAYFORMULA(IF(A2:A=\"\",, {expr}))"
    formulas("data!E2:I2", [[key("A2:A&\"|\"&B2:B"), key("A2:A&\"|\"&B2:B&\"|\"&TEXT(C2:C,\"yyyy-mm\")"), key("A2:A&\"|\"&TEXT(C2:C,\"yyyy-mm\")"),
                             key("A2:A&\"|\"&YEAR(C2:C)"), key("IF(MONTH(C2:C)<=MONTH(model!$B$4), A2:A&\"|\"&YEAR(C2:C)&\"|ytd\", ))")]])
    # inputs key: target|line|month; anchored in row 1 because append_values inserts rows; the last row per key wins (XLOOKUP from the bottom)
    formulas("inputs!H1", [["=ARRAYFORMULA(IF(ROW(A1:A)=1, \"key\", IF(A1:A=\"\",, A1:A&\"|\"&B1:B&\"|\"&TEXT(C1:C,\"yyyy-mm\"))))"]])
    # model: settings + per-line metrics + scope list
    lo, hi = 13, 12 + N_LINES
    by = lambda c, k: f"SUMIF(data!${c}:${c}, A{lo}:A{hi}&\"|\"&{k}, data!$D:$D)"
    ratio = lambda c, p: f"IF({p}=0, 0, {c}/{p}-1)"
    m = lambda k: f"=ARRAYFORMULA(LET(p, {by('G', f'TEXT(EDATE($B$2,{k - 12}),\"yyyy-mm\")')}, {ratio(by('G', f'TEXT(EDATE($B$2,{k}),\"yyyy-mm\")'), 'p')}))"
    formulas("model!A1:B10", [["Setting", "Value"], ["current", "=MAX(data!C2:C)"], ["t", "=YEAR(B2)"], ["trend end", "=EDATE(B2,-2)"],
                              ["YTD from", "=DATE(B3,1,1)"], ["w YTD", 0.4], ["w M", 0.3], ["w M-1", 0.2], ["w M-2", 0.1],
                              ["YTD label", "=\"YTD \"&B3&\" (\"&TEXT(B5,\"mm\")&\"–\"&TEXT(B4,\"mm\")&\")\""]])
    formulas("model!A12:K12", [["line", "=\"FY \"&(B3-2)", "=\"FY \"&(B3-1)", "=\"FY \"&(B3-1)&\" %\"", "=B10", "=\"YTD \"&(B3-1)", "YTD %",
                                "=TEXT(EDATE(B2,-2),\"yyyy-mm\")&\" %\"", "=TEXT(EDATE(B2,-3),\"yyyy-mm\")&\" %\"", "=TEXT(EDATE(B2,-4),\"yyyy-mm\")&\" %\"", "F-YoY %"]])
    formulas(f"model!A{lo}:K{lo}", [[
        f"=ARRAYFORMULA(lines!A2:A{N_LINES + 1})",
        f"=ARRAYFORMULA({by('H', '($B$3-2)')})",
        f"=ARRAYFORMULA({by('H', '($B$3-1)')})",
        f"=ARRAYFORMULA({ratio(f'C{lo}:C{hi}', f'B{lo}:B{hi}')})",
        f"=ARRAYFORMULA({by('I', '$B$3&\"|ytd\"')})",
        f"=ARRAYFORMULA({by('I', '($B$3-1)&\"|ytd\"')})",
        f"=ARRAYFORMULA({ratio(f'E{lo}:E{hi}', f'F{lo}:F{hi}')})",
        m(-2), m(-3), m(-4),
        f"=ARRAYFORMULA($B$6*G{lo}:G{hi}+$B$7*H{lo}:H{hi}+$B$8*I{lo}:I{hi}+$B$9*J{lo}:J{hi})"]])
    formulas("model!M1:M2", [["scope"], ["={\"TOTAL\"; UNIQUE(entities!C2:C%d); FILTER(entities!A2:A%d, entities!D2:D%d=TRUE)}" % ((len(ENTITIES) + 1,) * 3)]])
    # grid: pairs × 48 months (Jan t-2 .. Dec t+1) + row sums
    values("grid!A1:D1", [["line", "entity", "division", "in_scope"]])
    formulas("grid!E1:AZ1", [["=DATE(model!$B$3-2,1,1)"] + [f"=EDATE({col(4 + i)}1,1)" for i in range(1, 48)]])
    values("grid!BA1:BF1", [["FY t-2", "FY t-1", "YTD t", "YTD t-1", "line|division", "scoped line"]])
    formulas("grid!A2", [["=ARRAYFORMULA(LET(k, UNIQUE(FILTER(data!E2:E, data!E2:E<>\"\")), {REGEXEXTRACT(k, \"^[^|]+\"), REGEXEXTRACT(k, \"[^|]+$\")}))"]])
    formulas("grid!C2:D2", [["=ARRAYFORMULA(IF(A2:A=\"\",, IFERROR(VLOOKUP(B2:B, entities!$A$2:$C$%d, 3, FALSE), \"?\")))" % (len(ENTITIES) + 1),
                             "=ARRAYFORMULA(IF(A2:A=\"\",, (Summary!$B$1=\"TOTAL\")+(C2:C=Summary!$B$1)+(B2:B=Summary!$B$1)>0))"]])
    formulas("grid!BE2:BF2", [["=ARRAYFORMULA(IF(A2:A=\"\",, A2:A&\"|\"&C2:C))", "=ARRAYFORMULA(IF(A2:A=\"\",, IF(D2:D, A2:A, )))"]])
    actual = lambda x: f"SUMIF(data!$F:$F, $A2:$A&\"|\"&$B2:$B&\"|\"&TEXT({x}$1,\"yyyy-mm\"), data!$D:$D)"
    last = lambda k, c: f"XLOOKUP({k}, inputs!$H:$H, inputs!${c}:${c}, \"\", 0, -1)"   # the last input row for the key, "" if none
    row = []
    for i in range(48):
        x = col(5 + i)
        if i < 12:                                     # year t-2: never a forecast month
            row.append(f"=ARRAYFORMULA(IF($A2:$A=\"\",, {actual(x)}))")
            continue
        ly = col(5 + i - 12)                           # LY = the same row, 12 columns to the left
        fc = (f"LET(m, TEXT({x}$1,\"yyyy-mm\"), ke, $B2:$B&\"|\"&$A2:$A&\"|\"&m, kd, $C2:$C&\"|\"&$A2:$A&\"|\"&m, kt, \"TOTAL|\"&$A2:$A&\"|\"&m, "
              f"ly, {ly}2:{ly}, lyd, SUMIF($BE$2:$BE, $A2:$A&\"|\"&$C2:$C, {ly}2:{ly}), lyt, SUMIF($A$2:$A, $A2:$A, {ly}2:{ly}), "
              f"g, 1+IFERROR(VLOOKUP($A2:$A, model!$A${lo}:$K${hi}, 11, FALSE), 0), "
              f"ae, {last('ke', 'D')}, ye, {last('ke', 'E')}, ad, {last('kd', 'D')}, yd, {last('kd', 'E')}, at, {last('kt', 'D')}, yt, {last('kt', 'E')}, "
              f"IF(ae<>\"\", ae, IF(ye<>\"\", (1+ye)*ly, IF(ad<>\"\", IF(lyd=0, 0, ad*ly/lyd), IF(yd<>\"\", (1+yd)*ly, "
              f"IF(at<>\"\", IF(lyt=0, 0, at*ly/lyt), IF(yt<>\"\", (1+yt)*ly, g*ly)))))))")
        row.append(f"=ARRAYFORMULA(IF($A2:$A=\"\",, IF({x}$1<model!$B$2, {actual(x)}, {fc})))")
    for k in range(4):                                 # four calls of twelve months each
        formulas(f"grid!{col(5 + 12 * k)}2:{col(16 + 12 * k)}2", [row[12 * k:12 * k + 12]])
    ones = "SEQUENCE(12,1,1,0)"
    formulas("grid!BA2:BD2", [[f"=ARRAYFORMULA(IF($A2:$A=\"\",, MMULT(E2:P*1, {ones})))",
                               f"=ARRAYFORMULA(IF($A2:$A=\"\",, MMULT(Q2:AB*1, {ones})))",
                               f"=ARRAYFORMULA(IF($A2:$A=\"\",, MMULT((AC2:AN*1)*(AC$1:AN$1<=model!$B$4), {ones})))",
                               f"=ARRAYFORMULA(IF($A2:$A=\"\",, MMULT((Q2:AB*1)*(Q$1:AB$1<=EDATE(model!$B$4,-12)), {ones})))"]])


# ---------------------------------------------------------------- 4. Summary
def summary():
    first, last = 3, 2 + N_LINES
    tot = last + 1                                     # 23 KOPĀ IEŅĒMUMI, 24 KOPĀ IZDEVUMI, 25 PEĻŅA
    rev = (first, first + len(REV) - 1)
    exp = (first + len(REV), last)
    lines_a = f"lines!$A$2:$A${N_LINES + 1}"
    formulas("Summary!A1:D1", [["Scope", "TOTAL", None,
                                "=\"Actuals through \"&TEXT(model!B2,\"yyyy-mm\")&\" (MTD) · forecast from \"&TEXT(model!B2,\"yyyy-mm\")&\" · trend to \"&TEXT(model!B4,\"yyyy-mm\")"]])
    head = ["P&L line", "=model!B12", "=model!C12", "=model!D12", "=model!E12", "=model!F12", "YTD %", "F-YoY %", "=DATE(model!$B$3,1,1)"]
    head += [f"=EDATE({col(8 + i)}2,1)" for i in range(1, 12)] + ["=\"FY \"&model!$B$3", "=\"FY \"&model!$B$3&\" %\"", "=EDATE(T2,1)"]
    head += [f"=EDATE({col(22 + i)}2,1)" for i in range(1, 12)] + ["=\"FY \"&(model!$B$3+1)", "=\"FY \"&(model!$B$3+1)&\" %\""]
    formulas("Summary!A2:AJ2", [head])
    sumif = lambda c: f"=ARRAYFORMULA(SUMIF(grid!$BF:$BF, {lines_a}, grid!${c}:${c}))"
    month = lambda x: f"=ARRAYFORMULA(SUMIF(grid!$BF:$BF, {lines_a}, INDEX(grid!$E:$AZ, 0, MATCH({x}$2, grid!$E$1:$AZ$1, 0))))"
    r = lambda c, p: f"=ARRAYFORMULA(IF({p}{first}:{p}{last}=0, 0, {c}{first}:{c}{last}/{p}{first}:{p}{last}-1))"
    row = [f"=ARRAYFORMULA(lines!A2:A{N_LINES + 1}&\" \"&lines!B2:B{N_LINES + 1})", sumif("BA"), sumif("BB"), r("C", "B"), sumif("BC"), sumif("BD"), r("E", "F"),
           f"=ARRAYFORMULA(IFERROR(VLOOKUP({lines_a}, model!$A$13:$K${12 + N_LINES}, 11, FALSE), 0))"]
    row += [month(col(9 + i)) for i in range(12)] + [None, None] + [month(col(23 + i)) for i in range(12)]
    formulas(f"Summary!A{first}:AH{first}", [row])
    formulas(f"Summary!U{first}:V{last}", [[f"=SUM(I{i}:T{i})", f"=IF(C{i}=0, 0, U{i}/C{i}-1)"] for i in range(first, last + 1)])
    formulas(f"Summary!AI{first}:AJ{last}", [[f"=SUM(W{i}:AH{i})", f"=IF(U{i}=0, 0, AI{i}/U{i}-1)"] for i in range(first, last + 1)])
    sum_cols = [2, 3, 5, 6] + list(range(9, 21)) + list(range(23, 35))
    rows = []
    for k, (code, name) in enumerate(TOTALS):
        i = tot + k
        lo, hi = rev if k == 0 else exp
        cells = [f"{code} {name}"]
        for c in range(2, 37):
            L = col(c)
            if c in sum_cols:
                cells.append(f"=SUM({L}{lo}:{L}{hi})" if k < 2 else f"={L}{tot}+{L}{tot + 1}")
            elif c == 4:
                cells.append(f"=IF(B{i}=0, 0, C{i}/B{i}-1)")
            elif c == 7:
                cells.append(f"=IF(F{i}=0, 0, E{i}/F{i}-1)")
            elif c == 8:
                cells.append(f"=IF(C{i}=0, 0, SUMPRODUCT(H{lo}:H{hi}, C{lo}:C{hi})/C{i})" if k < 2 else f"=IF(C{i}=0, 0, (H{tot}*C{tot}+H{tot + 1}*C{tot + 1})/C{i})")
            elif c == 21:
                cells.append(f"=SUM(I{i}:T{i})")
            elif c == 22:
                cells.append(f"=IF(C{i}=0, 0, U{i}/C{i}-1)")
            elif c == 35:
                cells.append(f"=SUM(W{i}:AH{i})")
            elif c == 36:
                cells.append(f"=IF(U{i}=0, 0, AI{i}/U{i}-1)")
        rows.append(cells)
    formulas(f"Summary!A{tot}:AJ{tot + 2}", rows)
    c0 = tot + 5                                       # Checks block
    checks = [["Checks", "data", "Summary (scope TOTAL)", "ok"],
              ["=\"FY \"&(model!B3-2)&\" revenue\"", "=SUMIFS(data!D:D, data!A:A, \"1.*\", data!C:C, \">=\"&DATE(model!B3-2,1,1), data!C:C, \"<\"&DATE(model!B3-1,1,1))", f"=B{tot}", f"=ROUND(B{c0 + 1}-C{c0 + 1})=0"],
              ["=\"FY \"&(model!B3-2)&\" expenses\"", "=SUMIFS(data!D:D, data!A:A, \"2.*\", data!C:C, \">=\"&DATE(model!B3-2,1,1), data!C:C, \"<\"&DATE(model!B3-1,1,1))", f"=B{tot + 1}", f"=ROUND(B{c0 + 2}-C{c0 + 2})=0"],
              ["=\"FY \"&(model!B3-1)&\" revenue\"", "=SUMIFS(data!D:D, data!A:A, \"1.*\", data!C:C, \">=\"&DATE(model!B3-1,1,1), data!C:C, \"<\"&DATE(model!B3,1,1))", f"=C{tot}", f"=ROUND(B{c0 + 3}-C{c0 + 3})=0"],
              ["=\"FY \"&(model!B3-1)&\" expenses\"", "=SUMIFS(data!D:D, data!A:A, \"2.*\", data!C:C, \">=\"&DATE(model!B3-1,1,1), data!C:C, \"<\"&DATE(model!B3,1,1))", f"=C{tot + 1}", f"=ROUND(B{c0 + 4}-C{c0 + 4})=0"],
              ["=\"FY \"&model!B3&\" revenue, closed months\"", "=SUMIFS(data!D:D, data!A:A, \"1.*\", data!C:C, \">=\"&DATE(model!B3,1,1), data!C:C, \"<\"&model!B2)", f"=SUMPRODUCT((I$2:T$2<model!B2)*I{tot}:T{tot})", f"=ROUND(B{c0 + 5}-C{c0 + 5})=0"],
              ["data rows", "=COUNTA(data!A2:A)", None, None], ["grid pairs", "=COUNTA(grid!A2:A)", None, None],
              ["current month", "=TEXT(model!B2,\"yyyy-mm\")", None, None],
              ["pull months", "=IF(COUNT(pull!C2:C)=0, \"none\", TEXT(MIN(pull!C2:C),\"yyyy-mm\")&\"–\"&TEXT(MAX(pull!C2:C),\"yyyy-mm\"))", None, None],
              ["inputs", "=COUNTA(inputs!A2:A)", None, None]]
    formulas(f"Summary!A{c0}:D{c0 + len(checks) - 1}", checks)
    reqs = [fmt(SUMMARY, "A1:AJ2", bold=True), fmt(SUMMARY, "A2:AJ2", border="bottom"), fmt(SUMMARY, f"A{tot}:AJ{tot + 2}", bold=True, border="top"),
            fmt(SUMMARY, f"B{first}:C{tot + 2}", number=AMT), fmt(SUMMARY, f"E{first}:F{tot + 2}", number=AMT), fmt(SUMMARY, f"I{first}:U{tot + 2}", number=AMT),
            fmt(SUMMARY, f"W{first}:AI{tot + 2}", number=AMT), fmt(SUMMARY, f"D{first}:D{tot + 2}", number=PCT), fmt(SUMMARY, f"G{first}:H{tot + 2}", number=PCT),
            fmt(SUMMARY, f"V{first}:V{tot + 2}", number=PCT), fmt(SUMMARY, f"AJ{first}:AJ{tot + 2}", number=PCT), fmt(SUMMARY, "I2:T2", number=MONTH), fmt(SUMMARY, "W2:AH2", number=MONTH),
            fmt(SUMMARY, f"A{c0}:D{c0}", bold=True, border="bottom"), fmt(SUMMARY, f"B{c0 + 1}:C{c0 + 5}", number=AMT), fmt(SUMMARY, "B1:B1", number=TEXT, bg="FFF2CC"),
            rule(SUMMARY, [f"I{first}:T{tot + 2}", f"W{first}:AH{tot + 2}"], f"=I$2>=INDIRECT(\"model!$B$2\")", BLUE),
            rule(SUMMARY, [f"D{first}:D{tot + 2}", f"G{first}:H{tot + 2}", f"V{first}:V{tot + 2}", f"AJ{first}:AJ{tot + 2}"], f"=D{first}<-0.1", RED, RED_FG),
            rule(SUMMARY, [f"D{first}:D{tot + 2}", f"G{first}:H{tot + 2}", f"V{first}:V{tot + 2}", f"AJ{first}:AJ{tot + 2}"], f"=D{first}>0.1", GREEN, GREEN_FG),
            {"setDataValidation": {"range": grid_range(SUMMARY, "B1:B1"), "rule": {"condition": {"type": "ONE_OF_RANGE", "values": [{"userEnteredValue": "=model!$M$2:$M$40"}]}, "showCustomUi": True, "strict": True}}},
            width(SUMMARY, 1, 1, 330), width(SUMMARY, 2, 36, 92)]
    batch(reqs)


# ---------------------------------------------------------------- 5. line template, duplicated per line
def layout():
    """[(label, code, col, kind)] blocks of 4 columns; a gap column before every division but the first."""
    blocks, c = [("TOTAL", "TOTAL", 2, "total")], 7
    for d in DIVISIONS:
        ents = [e for e in ACTIVE if e[2] == d]
        blocks.append((f"TOTAL {d}", d, c, "division"))
        for e in ents:
            c += 4
            blocks.append((f"{e[0]} {e[1]}", e[0], c, "entity"))
        c += 5
    return blocks


def template():
    blocks = layout()
    first, last = 16, 39
    formulas("_line!A1:B2", [["", "=IFERROR(VLOOKUP(A1, lines!A:B, 2, FALSE), \"\")"], ["Calculations (line total, all entities)", "Value"]])
    formulas("_line!A3:B12", [[f"=model!{col(1 + n)}12", f"=IFERROR(VLOOKUP($A$1, model!$A$13:$K${12 + N_LINES}, {n + 1}, FALSE), 0)"] for n in range(1, 11)])
    formulas("_line!A15:A41", [["Month"]] + [["=DATE(model!$B$3,1,1)"]] + [[f"=EDATE(A{first + i - 1},1)"] for i in range(1, 24)]
             + [["=\"TOTAL \"&model!$B$3"], ["=\"TOTAL \"&(model!$B$3+1)"]])
    head13, head14, head15 = [], [], []
    for label, code, c, kind in blocks:
        while len(head13) < c - 2:
            head13.append(None); head14.append(None); head15.append(None)
        head13 += [label, None, None, None]
        head14 += [code, None, None, None]
        head15 += ["Amount", "LY", "YoY %", "F-YoY %"]
    values("_line!B13:%s15" % col(1 + len(head13)), [head13, head14, head15])
    zeros = "SEQUENCE(24,1,0,0)"
    endc = blocks[-1][2] + 3
    row16, row40, row41 = [None] * (endc - 1), [None] * (endc - 1), [None] * (endc - 1)
    for label, code, c, kind in blocks:
        a, l = col(c), col(c + 1)
        if kind == "entity":
            amt = f"=IFERROR(TRANSPOSE(FILTER(grid!$AC$2:$AZ, grid!$A$2:$A=$A$1, grid!$B$2:$B={a}$14)), {zeros})"
            ly = f"=IFERROR(TRANSPOSE(FILTER(grid!$Q$2:$AN, grid!$A$2:$A=$A$1, grid!$B$2:$B={a}$14)), {zeros})"
        else:
            cond = f"grid!$C$2:$C={a}$14" if kind == "division" else "grid!$A$2:$A=$A$1"
            n = f"COUNTIFS(grid!$A:$A, $A$1, grid!$C:$C, {a}$14)" if kind == "division" else "COUNTIF(grid!$A:$A, $A$1)"
            amt = f"=IFERROR(MMULT(TRANSPOSE(FILTER(grid!$AC$2:$AZ, grid!$A$2:$A=$A$1, {cond})), SEQUENCE({n},1,1,0)), {zeros})"
            ly = f"=IFERROR(MMULT(TRANSPOSE(FILTER(grid!$Q$2:$AN, grid!$A$2:$A=$A$1, {cond})), SEQUENCE({n},1,1,0)), {zeros})"
        yoy = f"=ARRAYFORMULA(IF($A{first}:$A{last}<model!$B$2, IF({l}{first}:{l}{last}=0, 0, {a}{first}:{a}{last}/{l}{first}:{l}{last}-1), ))"
        fyoy = f"=ARRAYFORMULA(IF($A{first}:$A{last}>=model!$B$2, IF({l}{first}:{l}{last}=0, 0, {a}{first}:{a}{last}/{l}{first}:{l}{last}-1), ))"
        row16[c - 2:c + 2] = [amt, ly, yoy, fyoy]
        row40[c - 2:c + 1] = [f"=SUM({a}{first}:{a}{first + 11})", f"=SUM({l}{first}:{l}{first + 11})", f"=IF({l}{last + 1}=0, 0, {a}{last + 1}/{l}{last + 1}-1)"]
        row41[c - 2:c + 1] = [f"=SUM({a}{first + 12}:{a}{last})", f"=SUM({l}{first + 12}:{l}{last})", f"=IF({l}{last + 2}=0, 0, {a}{last + 2}/{l}{last + 2}-1)"]
    formulas(f"_line!B{first}:{col(endc)}{first}", [row16])
    formulas(f"_line!B{last + 1}:{col(endc)}{last + 2}", [row40, row41])
    reqs = [fmt(TEMPLATE, "A1:B2", bold=True), fmt(TEMPLATE, "A2:B2", border="bottom"), fmt(TEMPLATE, "A12:B12", bold=True),
            fmt(TEMPLATE, "B3:B4", number=AMT), fmt(TEMPLATE, "B6:B7", number=AMT), fmt(TEMPLATE, "B5:B5", number=PCT), fmt(TEMPLATE, "B8:B12", number=PCT),
            fmt(TEMPLATE, f"A13:{col(endc)}15", bold=True), fmt(TEMPLATE, f"A15:{col(endc)}15", border="bottom"), fmt(TEMPLATE, f"A13:{col(endc)}14", align="CENTER"),
            fmt(TEMPLATE, f"A{first}:A{last}", number=MONTH, bold=True), fmt(TEMPLATE, f"A{last + 1}:{col(endc)}{last + 2}", bold=True, border="top"),
            width(TEMPLATE, 1, 1, 96), width(TEMPLATE, 2, endc, 84)]
    amt_cols, pct_cols = [], []
    for label, code, c, kind in blocks:
        reqs.append({"mergeCells": {"range": grid_range(TEMPLATE, f"{col(c)}13:{col(c + 3)}13"), "mergeType": "MERGE_ALL"}})
        reqs += [fmt(TEMPLATE, f"{col(c)}{first}:{col(c + 1)}{last + 2}", number=AMT), fmt(TEMPLATE, f"{col(c + 2)}{first}:{col(c + 3)}{last + 2}", number=PCT)]
        amt_cols.append(f"{col(c)}{first}:{col(c)}{last}")
        pct_cols.append(f"{col(c + 2)}{first}:{col(c + 3)}{last + 2}")
    gaps = [b[2] - 1 for b in blocks if b[3] == "division"][1:]
    for g in gaps + [6]:
        reqs += [fmt(TEMPLATE, f"{col(g)}13:{col(g)}{last + 2}", bg="F2F2F2"), width(TEMPLATE, g, g, 14)]
    reqs += [rule(TEMPLATE, amt_cols, f"=$A{first}>=INDIRECT(\"model!$B$2\")", BLUE),
             rule(TEMPLATE, pct_cols, f"={col(4)}{first}<-0.1", RED, RED_FG), rule(TEMPLATE, pct_cols, f"={col(4)}{first}>0.1", GREEN, GREEN_FG)]
    for d in DIVISIONS:                                # group entity columns under their division
        ents = [b for b in blocks if b[3] == "entity" and dict((e[0], e[2]) for e in ENTITIES)[b[1]] == d]
        if ents:
            reqs.append({"addDimensionGroup": {"range": {"sheetId": TEMPLATE, "dimension": "COLUMNS", "startIndex": ents[0][2] - 1, "endIndex": ents[-1][2] + 3}}})
    batch(reqs)


def duplicate():
    reqs = []
    for i, (code, name, _) in enumerate(LINES):
        reqs.append({"duplicateSheet": {"sourceSheetId": TEMPLATE, "insertSheetIndex": 1 + i, "newSheetId": 100 + i, "newSheetName": f"{code} {name}"[:100]}})
        reqs.append({"updateCells": {"start": {"sheetId": 100 + i, "rowIndex": 0, "columnIndex": 0}, "rows": [{"values": [{"userEnteredValue": {"stringValue": code}}]}], "fields": "userEnteredValue"}})
    reqs.append({"deleteSheet": {"sheetId": TEMPLATE}})
    batch(reqs)


def build(spreadsheet_id):
    global sid
    sid, calls[:] = spreadsheet_id, []
    structure()
    taxonomy_values()
    engine()
    summary()
    template()
    duplicate()
    return calls


if __name__ == "__main__":
    out = build(sys.argv[1])
    if "--step" in sys.argv:
        print(json.dumps(out[int(sys.argv[sys.argv.index("--step") + 1])], ensure_ascii=False, indent=1))
    else:
        print(json.dumps(out, ensure_ascii=False))
