"""Actuals — the Horizon P&L grid: line × entity × month, year-end close excluded.

Input : the `powerbi_run_dax` result of DAX (GoodHealth Power BI tools run it).
Output: rows {line, entity, month, amount} at the forecast grain.
"""

# dim Strukturvienibas: ACCOUNT_NR -> (ACCOUNT_DESCRIPTION, ACCOUNT_DIVISION)
ENTITIES = {
    "02": ("Pārdošana", "Pārdošana"),
    "02.1": ("Pārdošanas jaunas auto", "Pārdošana"),
    "02.2": ("Pārdošanas lietoti auto", "Pārdošana"),
    "02.3": ("Pirmspārdošanas serviss", "Pārdošana"),
    "02.4": ("Lietotie auto", "Pārdošana"),
    "AS": ("Autosalons", "Pārdošana"),
    "AS2": ("Autosalons2", "Pārdošana"),
    "LA": ("Lietotie auto", "Pārdošana"),
    "NCS_R - Jaunu a/m tirdzniecība": ("Pārdošanas jaunas auto", "Pārdošana"),
    "03": ("Pēcpārdošanas serviss", "Serviss"),
    "03.2": ("Serviss", "Serviss"),
    "03.2.1": ("Rezerves daļas", "Serviss"),
    "03.2.2": ("Garantijas apkalpošana", "Serviss"),
    "PRT_R - Rezerves daļas": ("Serviss", "Serviss"),
    "S": ("Serviss", "Serviss"),
    "SRV_R - Serviss": ("Serviss", "Serviss"),
    "03.1": ("Virsbūves serviss", "Virsbūves"),
    "BOD_R - Virsbūvju serviss": ("Virsbūves serviss", "Virsbūves"),
    "KR": ("Auto krāsošana", "Virsbūves"),
    "06": ("Auto noma", "Auto noma"),
    "01": ("Administrācija", "Administrācija"),
    "A": ("Administrācija", "Administrācija"),
    "04": ("Saimniecība", "Kopējie"),
    "05": ("Kopējie", "Kopējie"),
    "07": ("Auto mazgātuve", "Kopējie"),
    "V": ("Neattiecināts", "Kopējie"),
}
DIVISIONS = ["Pārdošana", "Serviss", "Virsbūves", "Auto noma", "Administrācija", "Kopējie"]

DAX = """EVALUATE
SELECTCOLUMNS(
    SUMMARIZECOLUMNS(
        'HOR Financials'[PNL_LINE],
        'dim Strukturvienibas'[ACCOUNT_NR],
        'dim Datumi'[yyyy-mm],
        FILTER ( ALL ( 'dim Datumi'[yyyy-mm] ), 'dim Datumi'[yyyy-mm] >= "@since" ),
        FILTER ( ALL ( 'HOR Financials'[PNL_LINE_Name] ), 'HOR Financials'[PNL_LINE_Name] <> "na" ),
        FILTER ( ALL ( 'HOR Financials'[PNL_LINE] ), NOT CONTAINSSTRING ( 'HOR Financials'[PNL_LINE], "99999" ) ),
        FILTER ( ALL ( 'HOR Financials'[transaction_type] ), 'HOR Financials'[transaction_type] <> "Kontu reformācija" ),
        "a", [amount]
    ),
    "l", 'HOR Financials'[PNL_LINE],
    "e", 'dim Strukturvienibas'[ACCOUNT_NR],
    "m", 'dim Datumi'[yyyy-mm],
    "a", ROUND ( [a], 0 )
)"""


def dax(since):
    """The query to run; `since` = first month to pull, 'YYYY-MM'."""
    return DAX.replace("@since", since)


def parse(result):
    """powerbi_run_dax result -> [{line, entity, month, amount}]."""
    rows = result["results"][0]["tables"][0]["rows"]
    return [{"line": r["[l]"], "entity": r["[e]"], "month": r["[m]"], "amount": float(r["[a]"] or 0)} for r in rows]


def refresh(stored, incoming):
    """Replace every stored month >= the earliest incoming month with the incoming rows."""
    since = min(r["month"] for r in incoming)
    return [r for r in stored if r["month"] < since] + incoming
