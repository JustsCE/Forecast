"""Synthetic ledger: 3 entities, 2 lines, 2025-01..2026-10; 2026 runs +10 % on 2025, so every growth term is 0.1."""
from io import BytesIO

import pytest
from openpyxl import load_workbook

import actuals as A
import forecast as F
import workbook as W

REV, EXP = "1.1 Auto pārdošanas ieņēmumi", "2.1 Auto iegādes izdevumi"
ENTS = ["02.1", "02.2", "03.2"]


@pytest.fixture
def actuals():
    out = {}
    for y, base in ((2025, 100.0), (2026, 110.0)):
        for m in F.months(y):
            if m > "2026-10":
                break
            for e in ENTS:
                out[(e, REV, m)] = base
                out[(e, EXP, m)] = -base / 2
    return out


def test_dax_parse_refresh():
    assert '>= "2025-01"' in A.dax("2025-01")
    result = {"results": [{"tables": [{"rows": [{"[l]": REV, "[e]": "02.1", "[m]": "2026-09", "[a]": 5.0},
                                                {"[l]": REV, "[e]": "02.1", "[m]": "2026-10", "[a]": None}]}]}]}
    incoming = A.parse(result)
    assert incoming[1]["amount"] == 0
    stored = [{"line": REV, "entity": "02.1", "month": m, "amount": 1.0} for m in ("2026-08", "2026-09", "2026-10")]
    assert [r["month"] for r in A.refresh(stored, incoming)] == ["2026-08", "2026-09", "2026-10"]
    assert A.refresh(stored, incoming)[1]["amount"] == 5.0


def test_metrics(actuals):
    assert F.current_month(actuals) == "2026-10"
    assert F.horizon("2026-10")[0] == "2026-10" and F.horizon("2026-10")[-1] == "2027-12"
    assert F.trend("2026-10") == (F.months(2026)[:8], ["2026-08", "2026-07", "2026-06"])
    m = F.metrics(F.line_series(actuals)[REV], "2026-10")
    assert list(m) == ["FY 2024", "FY 2025", "FY 2025 %", "YTD 2026 (01–08)", "YTD 2025", "YTD %", "2026-08 %", "2026-07 %", "2026-06 %", "F-YoY %"]
    assert m["FY 2024"] == 0 and m["FY 2025"] == 3600 and m["YTD 2026 (01–08)"] == 2640 and m["YTD 2025"] == 2400
    assert m["F-YoY %"] == pytest.approx(0.1) and m["YTD %"] == pytest.approx(0.1)


def test_baseline(actuals):
    actuals[("02.1", REV, "2026-09")] = 1.0        # last posted month still filling: skipped by the trend
    fc = F.forecast(actuals, {})
    assert fc[("02.1", REV, "2026-11")] == pytest.approx(110)
    assert fc[("02.1", REV, "2027-03")] == pytest.approx(121)
    assert fc[("02.1", EXP, "2027-03")] == pytest.approx(-60.5)
    actuals[("03.2", REV, "2025-03")], actuals[("03.2", REV, "2026-03")] = 0.001, 50.0   # entity base ×50,000
    g = F.metrics(F.line_series(actuals)[REV], "2026-10")["F-YoY %"]
    assert abs(g - 0.1) < 0.05 and F.forecast(actuals, {})[("03.2", REV, "2027-03")] == pytest.approx((1 + g) * 50)


def test_submit(actuals):
    subs = {}
    out = F.submit([{"entity": "Pārdošana", "line": "1.1", "month": "2026-11", "amount": 600}], actuals, subs)
    assert sorted(out) == [("02.1", REV, "2026-11", 300.0), ("02.2", REV, "2026-11", 300.0)]
    out = F.submit([{"entity": "02.1", "line": REV, "month": "2026-12", "yoy": 0.5}], actuals, subs)
    assert out == [("02.1", REV, "2026-12", 150.0)]
    subs[("02.1", REV, "2026-12")] = 150.0
    assert F.forecast(actuals, subs)[("02.1", REV, "2027-12")] == pytest.approx(165)
    out = F.submit([{"entity": "TOTAL", "line": "1.1", "month": "2027-01", "amount": 363}], actuals, subs)
    assert [round(a) for *_, a in out] == [121, 121, 121]
    for bad in ({"entity": "X", "line": "1.1", "month": "2026-11", "amount": 1},
                {"entity": "02.1", "line": "9.9", "month": "2026-11", "amount": 1},
                {"entity": "02.1", "line": "1.1", "month": "2026-09", "amount": 1}):
        with pytest.raises(ValueError):
            F.submit([bad], actuals, subs)


def test_run(tmp_path, monkeypatch, actuals):
    import run
    monkeypatch.setattr(run, "DIR", tmp_path)
    result = {"results": [{"tables": [{"rows": [{"[l]": l, "[e]": e, "[m]": m, "[a]": a} for (e, l, m), a in actuals.items()]}]}]}
    assert run.main("actuals", result)["months"] == ["2025-01", "2026-10"]
    out = run.main("submit", {"note": "t", "rows": [{"entity": "02.1", "line": "1.1", "month": "2026-11", "amount": 5}]})
    assert out["submitted"][0]["amount"] == 5 and out["build"]["submissions"] == 1 and (tmp_path / "NORDE-forecast.xlsx").exists()
    assert run.main("build", "Serviss")["file"].endswith("NORDE-forecast-Serviss.xlsx")
    assert run.main("submit", {"rows": [{"entity": "X", "line": "1.1", "month": "2026-11", "amount": 5}]}) == {"error": "unknown entity/division 'X'"}
    assert '>= "' in run.main("dax")["dax"]


def test_workbook(actuals):
    subs = {("02.1", REV, "2026-11"): 300.0, ("02.2", REV, "2026-11"): 300.0}
    wb = load_workbook(BytesIO(W.build(actuals, subs)))
    s = wb["Summary"]
    assert [c.value for c in s["A"][2:7]] == [REV, EXP, "1.99999 KOPĀ IEŅĒMUMI", "2.99999 KOPĀ IZDEVUMI", "3.99999 PEĻŅA"]
    assert [c.value for c in s[2][:9]] == ["P&L line", "FY 2024", "FY 2025", "FY 2025 %", "YTD 2026 (01–08)", "YTD 2025", "YTD %", "F-YoY %", "2026-01"]
    assert s["C3"].value == 3600 and s["E3"].value == "=SUM(I3:P3)" and s["F3"].value == 2400 and s["H3"].value == pytest.approx(0.1)
    assert s["Q3"].value == 330 and s["S3"].value == 710 and s["U3"].value == "=SUM(I3:T3)" and s["AJ3"].value == "=IF(U3=0,0,AI3/U3-1)"
    assert s["R3"].fill.fgColor.rgb.endswith("DCE6F1") and not s["Q3"].fill.fgColor.rgb.endswith("DCE6F1")
    assert s["I5"].value == "=I3" and s["I7"].value == "=I5+I6" and s["C7"].value == 1800 and s["H7"].value == pytest.approx(0.1)

    ws = wb[REV[:31]]
    assert ws["A1"].value == REV and ws["A3"].value == "FY 2024" and ws["A12"].value == "F-YoY %" and ws["B12"].value == pytest.approx(0.1)
    assert ws["B14"].value == "TOTAL" and ws["G14"].value == "TOTAL Pārdošana" and ws["K14"].value == "02.1 Pārdošanas jaunas auto" and ws["T14"].value == "TOTAL Serviss"
    assert [c.value for c in ws[15][1:6]] == ["Amount", "LY", "YoY %", "F-YoY %", None]
    assert ws["A16"].value == "2026-01" and ws["A39"].value == "2027-12" and ws["A40"].value == "TOTAL 2026"
    assert ws["K24"].value == 110 and ws["M24"].value == "=IF(L24=0,0,K24/L24-1)" and ws["N24"].value is None     # 2026-09 actual
    assert ws["K26"].value == 300 and ws["L26"].value == 100 and ws["M26"].value is None and ws["N26"].value == "=IF(L26=0,0,K26/L26-1)"
    assert ws["G26"].value == "=K26+O26" and ws["B26"].value == "=G26+T26" and ws["K40"].value == "=SUM(K16:K27)" and ws["K41"].value == "=SUM(K28:K39)"
    assert ws["K26"].fill.fgColor.rgb.endswith("DCE6F1") and not ws["K24"].fill.fgColor.rgb.endswith("DCE6F1")
    assert ws.column_dimensions["K"].outlineLevel == 1 and ws.column_dimensions["F"].width == 3 and ws.freeze_panes == "B16"

    d = load_workbook(BytesIO(W.build(actuals, subs, "Pārdošana")))[REV[:31]]
    assert d["B14"].value == "TOTAL Pārdošana" and d["F14"].value == "02.1 Pārdošanas jaunas auto" and d["B26"].value == "=F26+J26"
    e = load_workbook(BytesIO(W.build(actuals, subs, "02.1")))
    assert e["Summary"]["S3"].value == 300 and e[REV[:31]]["B14"].value == "02.1 Pārdošanas jaunas auto" and e[REV[:31]]["B26"].value == 300
    assert e[REV[:31]].max_column == 5
